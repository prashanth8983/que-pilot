"""
Synchronization service for real-time presentation tracking.
"""

import threading
import time
from typing import Optional, Dict, List, Callable
from dataclasses import dataclass


@dataclass
class SyncEvent:
    """Represents a synchronization event."""
    event_type: str
    timestamp: float
    data: Dict
    source: str


class SyncService:
    """Service for managing real-time synchronization between components."""

    def __init__(self):
        self.is_running = False
        self.sync_thread: Optional[threading.Thread] = None
        self.sync_interval = 1.0  # seconds

        # Event callbacks
        self.event_callbacks: List[Callable] = []
        self.slide_sync_callbacks: List[Callable] = []
        self.presentation_sync_callbacks: List[Callable] = []

        # State tracking
        self.last_slide_number = 0
        self.last_presentation_id = None
        self.sync_events: List[SyncEvent] = []
        self.max_events = 100  # Keep last 100 events

        # References to other services for actual sync
        self.presentation_service = None
    
    def start_sync(self, interval: float = 1.0):
        """Start the synchronization service."""
        if self.is_running:
            return
        
        self.sync_interval = interval
        self.is_running = True
        
        self.sync_thread = threading.Thread(target=self._sync_loop, daemon=True)
        self.sync_thread.start()
        
        print(f"🔄 Started sync service (interval: {interval}s)")
    
    def stop_sync(self):
        """Stop the synchronization service."""
        self.is_running = False
        if self.sync_thread:
            self.sync_thread.join(timeout=2.0)
        
        print("Stopped sync service")
    
    def _sync_loop(self):
        """Main synchronization loop."""
        while self.is_running:
            try:
                self._check_presentation_changes()
                self._check_slide_changes()
                time.sleep(self.sync_interval)
            except Exception as e:
                print(f"Sync error: {e}")
                time.sleep(self.sync_interval)
    
    def set_presentation_service(self, presentation_service):
        """Set reference to presentation service for syncing."""
        self.presentation_service = presentation_service

    def _check_presentation_changes(self):
        """Check for presentation changes."""
        if not self.presentation_service:
            return

        try:
            # Get current presentation summary
            summary = self.presentation_service.get_presentation_summary()
            if summary:
                current_id = summary.get('presentation_id')
                if current_id != self.last_presentation_id:
                    self.last_presentation_id = current_id
                    total_slides = summary.get('total_slides', 0)
                    self.on_presentation_load(current_id, total_slides)
        except Exception as e:
            print(f"Error checking presentation changes: {e}")

    def _check_slide_changes(self):
        """Check for slide changes."""
        if not self.presentation_service:
            return

        try:
            # Try to sync with PowerPoint and check for changes
            if hasattr(self.presentation_service, 'sync_with_powerpoint'):
                changed = self.presentation_service.sync_with_powerpoint()

                # Get current slide info
                summary = self.presentation_service.get_presentation_summary()
                if summary:
                    current_slide = summary.get('current_slide', 1)
                    total_slides = summary.get('total_slides', 1)

                    # Only emit event if slide actually changed
                    if current_slide != self.last_slide_number and changed:
                        slide_info = self.presentation_service.get_current_slide_info()
                        self.on_slide_change(current_slide, total_slides, slide_info or {})
        except Exception as e:
            print(f"Error checking slide changes: {e}")
    
    def emit_event(self, event_type: str, data: Dict, source: str = "sync_service"):
        """Emit a synchronization event."""
        event = SyncEvent(
            event_type=event_type,
            timestamp=time.time(),
            data=data,
            source=source
        )
        
        # Add to event history
        self.sync_events.append(event)
        if len(self.sync_events) > self.max_events:
            self.sync_events.pop(0)
        
        # Notify callbacks
        for callback in self.event_callbacks:
            callback(event)
        
        print(f"📡 Sync event: {event_type} from {source}")
    
    def on_slide_change(self, slide_number: int, total_slides: int, slide_info: Dict):
        """Handle slide change events."""
        if slide_number != self.last_slide_number:
            self.last_slide_number = slide_number
            
            event_data = {
                'slide_number': slide_number,
                'total_slides': total_slides,
                'slide_info': slide_info
            }
            
            self.emit_event('slide_changed', event_data, 'presentation_service')
            
            # Notify slide sync callbacks
            for callback in self.slide_sync_callbacks:
                callback(slide_number, total_slides, slide_info)
    
    def on_presentation_load(self, presentation_id: str, total_slides: int):
        """Handle presentation load events."""
        if presentation_id != self.last_presentation_id:
            self.last_presentation_id = presentation_id
            
            event_data = {
                'presentation_id': presentation_id,
                'total_slides': total_slides
            }
            
            self.emit_event('presentation_loaded', event_data, 'presentation_service')
            
            # Notify presentation sync callbacks
            for callback in self.presentation_sync_callbacks:
                callback(presentation_id, total_slides)
    
    def add_event_callback(self, callback: Callable):
        """Add a callback for all sync events."""
        self.event_callbacks.append(callback)
    
    def add_slide_sync_callback(self, callback: Callable):
        """Add a callback for slide change events."""
        self.slide_sync_callbacks.append(callback)
    
    def add_presentation_sync_callback(self, callback: Callable):
        """Add a callback for presentation load events."""
        self.presentation_sync_callbacks.append(callback)
    
    def get_sync_status(self) -> Dict:
        """Get current synchronization status."""
        return {
            'is_running': self.is_running,
            'sync_interval': self.sync_interval,
            'last_slide_number': self.last_slide_number,
            'last_presentation_id': self.last_presentation_id,
            'total_events': len(self.sync_events)
        }
    
    def get_recent_events(self, limit: int = 10) -> List[SyncEvent]:
        """Get recent synchronization events."""
        return self.sync_events[-limit:] if self.sync_events else []
    
    def clear_events(self):
        """Clear synchronization event history."""
        self.sync_events = []
        print("Cleared sync events")