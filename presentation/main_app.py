import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import platform
from presentation_detector import PowerPointTracker

class PPTTrackingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 PowerPoint Tracker Pro")
        self.root.geometry("520x420")
        self.root.minsize(500, 400)
        self.root.resizable(True, True)

        # Set window icon and modern styling
        try:
            # Try to set a modern icon if available
            self.root.iconname("PowerPoint Tracker")
        except:
            pass

        # Set Material Design styling
        self.setup_material_design()

        self.ppt_tracker = PowerPointTracker(auto_detect=True)  # Always enable auto-detection
        self.current_file_path = None
        self.auto_sync_var = None
        self.status_label = None
        self.detect_button = None
        self.auto_detection_active = True
        self.detection_timer = None

        self.create_widgets()
        self.start_auto_detection()  # Start auto-detection on startup

        # Handle window closing
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_material_design(self):
        """Configure Material Design appearance"""
        style = ttk.Style()

        # Use modern theme
        try:
            style.theme_use('clam')
        except:
            style.theme_use('default')

        # Enhanced Material Design color palette
        self.colors = {
            'primary': '#1976D2',      # Blue 700
            'primary_light': '#42A5F5', # Blue 400
            'primary_dark': '#1565C0', # Blue 800
            'accent': '#FF5722',       # Deep Orange 500
            'accent_light': '#FF8A65', # Deep Orange 300
            'background': '#F5F5F5',   # Gray 100
            'surface': '#FFFFFF',      # White
            'surface_variant': '#F8F9FA', # Light gray
            'error': '#F44336',        # Red 500
            'success': '#4CAF50',      # Green 500
            'warning': '#FF9800',      # Orange 500
            'info': '#2196F3',         # Blue 500
            'text_primary': '#212121', # Gray 900
            'text_secondary': '#757575', # Gray 600
            'text_hint': '#9E9E9E',    # Gray 500
            'divider': '#E0E0E0',      # Gray 300
            'elevation_1': '#FFFFFF',  # Card elevation
            'elevation_2': '#F8F9FA'   # Higher elevation
        }

        # Configure enhanced Material Design styles
        style.configure('MaterialCard.TFrame',
                       background=self.colors['surface'],
                       relief='flat', borderwidth=1,
                       lightcolor=self.colors['divider'])
        style.configure('MaterialTitle.TLabel',
                       font=('Segoe UI', 18, 'bold'),
                       background=self.colors['surface'],
                       foreground=self.colors['text_primary'])
        style.configure('MaterialSubtitle.TLabel',
                       font=('Segoe UI', 11),
                       background=self.colors['surface'],
                       foreground=self.colors['text_secondary'])
        style.configure('MaterialBody.TLabel',
                       font=('Segoe UI', 10),
                       background=self.colors['surface'],
                       foreground=self.colors['text_primary'])
        style.configure('MaterialCaption.TLabel',
                       font=('Segoe UI', 9),
                       background=self.colors['surface'],
                       foreground=self.colors['text_hint'])
        style.configure('MaterialPrimary.TButton',
                       font=('Segoe UI', 9, 'bold'),
                       padding=(12, 8))
        style.configure('MaterialSecondary.TButton',
                       font=('Segoe UI', 9),
                       padding=(10, 6))
        style.configure('MaterialStatus.TLabel',
                       font=('Segoe UI', 10),
                       background=self.colors['background'])

        # Set root background
        self.root.configure(bg=self.colors['background'])

    def create_widgets(self):
        # Main container with enhanced Material Design styling
        main_container = tk.Frame(self.root, bg=self.colors['background'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Create enhanced Material Design cards with better spacing
        self.create_header_card(main_container)
        self.create_live_tracking_card(main_container)
        self.create_controls_card(main_container)
        self.create_status_footer(main_container)

    def create_header_card(self, parent):
        """Create compact Material Design header card"""
        header_card = tk.Frame(parent, bg=self.colors['surface'], relief='solid', bd=0)
        header_card.pack(fill=tk.X, pady=(0, 8))

        # Add elevation shadow effect
        shadow_frame = tk.Frame(parent, height=2, bg='#E0E0E0')
        shadow_frame.pack(fill=tk.X)

        # Card content with padding
        content = tk.Frame(header_card, bg=self.colors['surface'], padx=16, pady=12)
        content.pack(fill=tk.X)

        # Title and detection button in same row
        title_row = tk.Frame(content, bg=self.colors['surface'])
        title_row.pack(fill=tk.X)

        # App title
        title_label = tk.Label(title_row, text="📊 PowerPoint Tracker",
                              font=('Roboto', 14, 'bold'),
                              bg=self.colors['surface'],
                              fg=self.colors['text_primary'])
        title_label.pack(side=tk.LEFT)

        # Detection button
        self.detect_button = tk.Button(title_row, text="🔍 Scanning...",
                                      command=self.manual_detect_powerpoint,
                                      font=('Roboto', 8),
                                      bg=self.colors['primary'],
                                      fg='white',
                                      relief='flat',
                                      padx=12, pady=4)
        self.detect_button.pack(side=tk.RIGHT)

    def create_live_tracking_card(self, parent):
        """Create Material Design live tracking card with current slide info"""
        tracking_card = tk.Frame(parent, bg=self.colors['surface'], relief='solid', bd=0)
        tracking_card.pack(fill=tk.X, pady=(0, 8))

        # Add elevation shadow effect
        shadow_frame = tk.Frame(parent, height=1, bg=self.colors['divider'])
        shadow_frame.pack(fill=tk.X)

        # Card content
        content = tk.Frame(tracking_card, bg=self.colors['surface'], padx=16, pady=16)
        content.pack(fill=tk.X)

        # Current presentation info
        pres_row = tk.Frame(content, bg=self.colors['surface'])
        pres_row.pack(fill=tk.X, pady=(0, 12))

        tk.Label(pres_row, text="📄", font=('Roboto', 14),
                bg=self.colors['surface'], fg=self.colors['accent']).pack(side=tk.LEFT)

        self.file_label = tk.Label(pres_row, text="No presentation detected",
                                  font=('Roboto', 11),
                                  bg=self.colors['surface'],
                                  fg=self.colors['text_secondary'])
        self.file_label.pack(side=tk.LEFT, padx=(10, 0))

        # Current slide indicator (large and prominent)
        slide_section = tk.Frame(content, bg=self.colors['surface'])
        slide_section.pack(fill=tk.X, pady=(0, 12))

        slide_header = tk.Frame(slide_section, bg=self.colors['surface'])
        slide_header.pack(fill=tk.X, pady=(0, 8))

        tk.Label(slide_header, text="Current Slide",
                font=('Roboto', 11, 'bold'),
                bg=self.colors['surface'],
                fg=self.colors['text_primary']).pack(side=tk.LEFT)

        # Auto-sync toggle
        self.auto_sync_var = tk.BooleanVar()
        self.auto_sync_var.set(True)
        self.auto_sync_checkbox = tk.Checkbutton(slide_header, text="🔄 Auto-Sync",
                                                variable=self.auto_sync_var,
                                                command=self.toggle_auto_sync,
                                                font=('Roboto', 9),
                                                bg=self.colors['surface'],
                                                fg=self.colors['text_primary'],
                                                selectcolor=self.colors['primary'],
                                                relief='flat')
        self.auto_sync_checkbox.pack(side=tk.RIGHT)

        # Large slide display
        slide_display = tk.Frame(slide_section, bg=self.colors['elevation_1'], relief='solid', bd=1, padx=20, pady=16)
        slide_display.pack(fill=tk.X)

        self.slide_info_label = tk.Label(slide_display, text="- / -",
                                        font=('Roboto', 28, 'bold'),
                                        bg=self.colors['elevation_1'],
                                        fg=self.colors['primary'])
        self.slide_info_label.pack()

    def create_status_footer(self, parent):
        """Create Material Design status footer"""
        footer_card = tk.Frame(parent, bg=self.colors['surface'], relief='solid', bd=0)
        footer_card.pack(fill=tk.X, side=tk.BOTTOM)

        # Add top border for separation
        border_frame = tk.Frame(footer_card, height=1, bg=self.colors['divider'])
        border_frame.pack(fill=tk.X)

        # Footer content
        content = tk.Frame(footer_card, bg=self.colors['surface'], padx=16, pady=10)
        content.pack(fill=tk.X)

        # Status indicator with icon
        self.status_label = tk.Label(content, text="🚀 Ready to track PowerPoint presentations",
                                    font=('Roboto', 9),
                                    bg=self.colors['surface'],
                                    fg=self.colors['text_secondary'],
                                    anchor='w')
        self.status_label.pack(fill=tk.X)


    def create_controls_card(self, parent):
        """Create enhanced Material Design controls card"""
        controls_card = tk.Frame(parent, bg=self.colors['surface'], relief='solid', bd=0)
        controls_card.pack(fill=tk.X, pady=(0, 8))

        # Add elevation shadow effect
        shadow_frame = tk.Frame(parent, height=1, bg=self.colors['divider'])
        shadow_frame.pack(fill=tk.X)

        # Card content
        content = tk.Frame(controls_card, bg=self.colors['surface'], padx=16, pady=14)
        content.pack(fill=tk.X)

        # Controls header
        controls_header = tk.Frame(content, bg=self.colors['surface'])
        controls_header.pack(fill=tk.X, pady=(0, 10))

        tk.Label(controls_header, text="🎮 Controls",
                font=('Roboto', 11, 'bold'),
                bg=self.colors['surface'],
                fg=self.colors['text_primary']).pack(side=tk.LEFT)

        # Navigation buttons row
        nav_row = tk.Frame(content, bg=self.colors['surface'])
        nav_row.pack(fill=tk.X, pady=(0, 10))

        # Previous button with enhanced styling
        self.prev_btn = tk.Button(nav_row, text="◀ Prev",
                                 command=self.previous_slide,
                                 font=('Roboto', 9, 'bold'),
                                 bg=self.colors['primary'],
                                 fg='white',
                                 relief='flat',
                                 padx=14, pady=8)
        self.prev_btn.pack(side=tk.LEFT, padx=(0, 6))

        # Next button with enhanced styling
        self.next_btn = tk.Button(nav_row, text="Next ▶",
                                 command=self.next_slide,
                                 font=('Roboto', 9, 'bold'),
                                 bg=self.colors['primary'],
                                 fg='white',
                                 relief='flat',
                                 padx=14, pady=8)
        self.next_btn.pack(side=tk.LEFT, padx=(0, 20))

        # Jump to slide section
        jump_frame = tk.Frame(nav_row, bg=self.colors['elevation_2'], relief='solid', bd=1, padx=10, pady=6)
        jump_frame.pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(jump_frame, text="Jump to:",
                font=('Roboto', 8),
                bg=self.colors['elevation_2'],
                fg=self.colors['text_secondary']).pack(side=tk.LEFT)

        self.slide_entry = tk.Entry(jump_frame, width=4,
                                   font=('Roboto', 9),
                                   relief='solid', bd=1,
                                   justify='center')
        self.slide_entry.pack(side=tk.LEFT, padx=(6, 4))

        tk.Button(jump_frame, text="Go",
                 command=self.go_to_slide,
                 font=('Roboto', 8, 'bold'),
                 bg=self.colors['accent'],
                 fg='white',
                 relief='flat',
                 padx=8, pady=2).pack(side=tk.LEFT)

        # Manual sync button
        tk.Button(nav_row, text="🔄 Sync Now",
                 command=self.sync_now,
                 font=('Roboto', 9),
                 bg=self.colors['success'],
                 fg='white',
                 relief='flat',
                 padx=12, pady=8).pack(side=tk.RIGHT)

        # Action buttons row
        action_row = tk.Frame(content, bg=self.colors['surface'])
        action_row.pack(fill=tk.X)

        tk.Button(action_row, text="📁 Browse File",
                 command=self.browse_file,
                 font=('Roboto', 9),
                 bg=self.colors['surface'],
                 fg=self.colors['text_primary'],
                 relief='solid', bd=1,
                 padx=14, pady=6).pack(side=tk.LEFT)



    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PowerPoint File",
            filetypes=[("PowerPoint files", "*.pptx *.ppt"), ("All files", "*.*")]
        )

        if file_path:
            try:
                self.ppt_tracker = PowerPointTracker(file_path)
                self.current_file_path = file_path
                self.file_label.config(text=os.path.basename(file_path))
                self.update_slide_info()
                self.display_slide_content()
                messagebox.showinfo("Success", f"Loaded presentation with {self.ppt_tracker.get_total_slides()} slides")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")

    def previous_slide(self):
        if self.ppt_tracker and self.ppt_tracker.previous_slide():
            self.update_slide_info()
            self.display_slide_content()

    def next_slide(self):
        if self.ppt_tracker and self.ppt_tracker.next_slide():
            self.update_slide_info()
            self.display_slide_content()

    def go_to_slide(self):
        if not self.ppt_tracker:
            return

        try:
            slide_num = int(self.slide_entry.get())
            if self.ppt_tracker.go_to_slide(slide_num):
                self.update_slide_info()
                self.display_slide_content()
            else:
                messagebox.showerror("Error", f"Invalid slide number. Please enter a number between 1 and {self.ppt_tracker.get_total_slides()}")
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid slide number")

    def update_slide_info(self):
        """Update slide information with enhanced visual feedback"""
        if self.ppt_tracker:
            current = self.ppt_tracker.get_current_slide_number()
            total = self.ppt_tracker.get_total_slides()

            # Update the large slide indicator
            self.slide_info_label.config(text=f"{current} / {total}")

            # Add visual feedback for slide changes
            self.slide_info_label.config(fg=self.colors['accent'])
            self.root.after(500, lambda: self.slide_info_label.config(fg=self.colors['primary']))
        else:
            self.slide_info_label.config(text="- / -")

    def display_slide_content(self):
        """Update the UI when slide content changes (simplified for compact UI)"""
        if not self.ppt_tracker:
            return
        # For compact UI, we just update slide info - no text display needed
        pass


    def start_auto_detection(self):
        """Start continuous auto-detection in the background"""
        self.update_status("Starting auto-detection...", "blue")
        self.check_powerpoint_automatically()

    def check_powerpoint_automatically(self):
        """Enhanced automatic PowerPoint detection with current slide tracking"""
        try:
            window_info = self.ppt_tracker.get_window_info()

            if window_info:
                slide_info = window_info['slide_info']

                # Update detection button
                self.detect_button.config(text="✅ Connected",
                                         bg=self.colors['success'])

                # Auto-load presentation if not already loaded
                if not self.current_file_path and slide_info.get('presentation_name'):
                    presentation_name = slide_info.get('presentation_name', 'Unknown')
                    presentation_mode = slide_info.get('mode', 'normal')

                    if self.ppt_tracker.auto_load_presentation():
                        self.current_file_path = self.ppt_tracker.ppt_path
                        self.file_label.config(text=presentation_name)
                        self.update_slide_info()

                        # Enable auto-sync automatically
                        self.ppt_tracker.enable_auto_sync()
                        self.auto_sync_var.set(True)

                        # Check if it's a .ppt file for appropriate messaging
                        file_extension = os.path.splitext(self.ppt_tracker.ppt_path)[1].lower()
                        if file_extension == '.ppt':
                            self.update_status(f"📋 Loaded .ppt file with live tracking: {presentation_name}", "green")
                        else:
                            self.update_status(f"📊 Loaded .pptx: {presentation_name}", "green")
                    else:
                        # Handle different cases for why loading failed
                        if presentation_mode in ['compatibility', 'limited']:
                            self.file_label.config(text=f"{presentation_name} (tracking via window)")
                            self.update_status(f"📋 Window-based tracking: {presentation_name}", "blue")
                        else:
                            self.file_label.config(text=f"{presentation_name} (file not found)")
                            self.update_status(f"🔍 Found: {presentation_name}", "orange")
                else:
                    # Update current slide if auto-sync is enabled
                    if self.auto_sync_var.get():
                        if self.ppt_tracker.sync_with_powerpoint_window():
                            self.update_slide_info()
                            current = self.ppt_tracker.get_current_slide_number()
                            total = self.ppt_tracker.get_total_slides()
                            self.update_status(f"🎯 Live tracking: {current}/{total}", "green")
                        else:
                            # Still show current slide from PowerPoint window
                            if slide_info.get('current_slide'):
                                current_slide = slide_info['current_slide']
                                total_slides = slide_info.get('total_slides', '?')
                                self.slide_info_label.config(text=f"{current_slide} / {total_slides}")

                                presentation_mode = slide_info.get('mode', 'normal')
                                if presentation_mode in ['compatibility', 'limited']:
                                    self.update_status(f"📋 Shows: {current_slide}/{total_slides} (limited tracking)", "orange")
                                else:
                                    self.update_status(f"📍 PowerPoint: {current_slide}/{total_slides}", "blue")

            else:
                # No PowerPoint found
                self.detect_button.config(text="🔍 Scanning...",
                                         bg=self.colors['primary'])
                if not self.current_file_path:
                    self.update_status("🔍 Scanning for PowerPoint...", "gray")
                    self.slide_info_label.config(text="- / -")

        except Exception as e:
            self.detect_button.config(text="❌ Error",
                                     bg=self.colors['error'])
            self.update_status("⚠️ Detection error", "red")

        # Schedule next check (faster for better slide tracking)
        if self.auto_detection_active:
            self.detection_timer = self.root.after(1500, self.check_powerpoint_automatically)  # Check every 1.5 seconds

    def stop_auto_detection(self):
        """Stop the automatic detection"""
        self.auto_detection_active = False
        if self.detection_timer:
            self.root.after_cancel(self.detection_timer)
            self.detection_timer = None

    def manual_detect_powerpoint(self):
        """Manual detection trigger - just forces an immediate check"""
        self.update_status("Manual detection triggered...", "blue")
        self.check_powerpoint_automatically()

    def auto_detect_powerpoint(self):
        """Legacy method - now just triggers manual detection"""
        self.manual_detect_powerpoint()

    def toggle_auto_sync(self):
        """Toggle auto-sync with PowerPoint window"""
        if not self.ppt_tracker or not getattr(self.ppt_tracker, 'auto_detect', False):
            messagebox.showwarning("Auto-Sync", "Please use Auto-Detect first to enable auto-sync.")
            self.auto_sync_var.set(False)
            return

        if self.auto_sync_var.get():
            if self.ppt_tracker.enable_auto_sync():
                self.update_status("Auto-sync enabled", "green")
                # Start a periodic sync check
                self.root.after(1000, self.periodic_sync_check)
            else:
                self.auto_sync_var.set(False)
                messagebox.showwarning("Auto-Sync", "Could not enable auto-sync. Please check PowerPoint connection.")
        else:
            self.ppt_tracker.disable_auto_sync()
            self.update_status("Auto-sync disabled", "gray")

    def sync_now(self):
        """Manually sync with PowerPoint window now"""
        if not self.ppt_tracker or not getattr(self.ppt_tracker, 'auto_detect', False):
            messagebox.showwarning("Sync", "Please use Auto-Detect first to enable sync.")
            return

        try:
            if self.ppt_tracker.sync_with_powerpoint_window():
                self.update_slide_info()
                self.display_slide_content()
                current = self.ppt_tracker.get_current_slide_number()
                self.update_status(f"Synced to slide {current}", "green")
            else:
                window_info = self.ppt_tracker.get_window_info()
                if window_info:
                    slide_info = window_info['slide_info']
                    if slide_info.get('current_slide'):
                        self.update_status(f"PowerPoint shows slide {slide_info['current_slide']} (no file loaded)", "orange")
                    else:
                        self.update_status("PowerPoint detected but no slide info", "orange")
                else:
                    self.update_status("No PowerPoint window found", "red")
        except Exception as e:
            messagebox.showerror("Sync Error", f"Failed to sync: {str(e)}")

    def periodic_sync_check(self):
        """Periodically check for sync when auto-sync is enabled"""
        if self.auto_sync_var.get() and self.ppt_tracker and getattr(self.ppt_tracker, 'auto_sync_enabled', False):
            try:
                if self.ppt_tracker.sync_with_powerpoint_window():
                    self.update_slide_info()
                    self.display_slide_content()
            except Exception:
                pass  # Silently handle errors in periodic sync

            # Schedule next check
            self.root.after(2000, self.periodic_sync_check)

    def update_status(self, message: str, color: str = "black"):
        """Update the status label with Material Design colors"""
        color_map = {
            'green': self.colors['success'],
            'blue': self.colors['primary'],
            'orange': self.colors['warning'],
            'red': self.colors['error'],
            'gray': self.colors['text_secondary'],
            'black': self.colors['text_primary']
        }
        actual_color = color_map.get(color, color)
        self.status_label.config(text=message, fg=actual_color)

    def on_closing(self):
        """Handle application closing"""
        self.stop_auto_detection()
        self.root.destroy()

def main():
    root = tk.Tk()
    app = PPTTrackingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
