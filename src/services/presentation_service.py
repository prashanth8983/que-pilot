"""
Presentation service for managing presentation state and operations.
"""

import os
from typing import Optional, Dict, List, Callable
from pathlib import Path

# Import core modules with fallback
try:
    from ..core.presentation import PresentationTracker, PowerPointWindowDetector, ContentProcessor
    from ..core.ai import VectorStore
except ImportError:
    # Fallback for when running from different contexts
    import sys
    from pathlib import Path
    core_path = Path(__file__).parent.parent / "core"
    sys.path.insert(0, str(core_path))
    from presentation import PresentationTracker, PowerPointWindowDetector, ContentProcessor
    from ai import VectorStore


class PresentationService:
    """Service for managing presentation operations and state."""
    
    def __init__(self):
        self.tracker: Optional[PresentationTracker] = None
        self.detector: Optional[PowerPointWindowDetector] = None
        self.processor = ContentProcessor()
        self.vector_store = VectorStore()
        
        self.current_presentation_id: Optional[str] = None
        self.is_presenting = False
        self.slide_change_callbacks: List[Callable] = []
        self.presentation_load_callbacks: List[Callable] = []
    
    def load_presentation(self, file_path: str, auto_detect: bool = True) -> bool:
        """Load a presentation file."""
        try:
            # Validate input parameters
            if file_path is None:
                print("Error: No file path provided")
                return False

            # Handle the case where a boolean might be passed accidentally
            if isinstance(file_path, bool):
                print(f"Error: file_path cannot be a boolean value ({file_path}). Expected string or PathLike object.")
                return False

            if not isinstance(file_path, (str, os.PathLike)):
                print(f"Error: file_path must be a string or PathLike object, got {type(file_path)} with value: {file_path}")
                return False

            file_path = str(file_path)  # Ensure it's a string

            if not os.path.exists(file_path):
                print(f"File not found: {file_path}")
                return False

            # Initialize tracker
            self.tracker = PresentationTracker(file_path, auto_detect)
            self.detector = PowerPointWindowDetector() if auto_detect else None
            
            # Generate presentation ID
            self.current_presentation_id = Path(file_path).stem
            
            # Process slides and add to vector store
            self._process_presentation_slides()
            
            # Notify callbacks
            for callback in self.presentation_load_callbacks:
                callback(self.current_presentation_id, self.tracker.total_slides)
            
            print(f"Loaded presentation: {os.path.basename(file_path)}")
            return True
            
        except Exception as e:
            print(f"Failed to load presentation: {e}")
            return False
    
    def _process_presentation_slides(self):
        """Process all slides and add to vector store."""
        if not self.tracker or not self.current_presentation_id:
            return
        
        slides_data = []
        
        for i in range(self.tracker.total_slides):
            slide_data = self.tracker.get_slide_info(i)
            slides_data.append(slide_data)
        
        # Add to vector store
        self.vector_store.add_presentation_slides(self.current_presentation_id, slides_data)
        print(f"Processed {len(slides_data)} slides for vector search")
    
    def start_presentation(self) -> bool:
        """Start presentation mode."""
        if not self.tracker:
            print("No presentation loaded")
            return False

        self.is_presenting = True

        # Enable auto-sync if detector is available
        if self.detector:
            self.tracker.enable_auto_sync()

        print("Presentation started")
        return True

    def sync_with_powerpoint(self) -> bool:
        """Manually sync with PowerPoint to get current slide info."""
        if not self.tracker or not self.detector:
            return False

        previous_slide = self.tracker.current_slide_index + 1

        # Try to sync with PowerPoint
        if self.tracker.sync_with_powerpoint_window():
            current_slide = self.tracker.current_slide_index + 1

            # Check if slide actually changed
            if current_slide != previous_slide:
                print(f"Slide changed: {previous_slide} → {current_slide}")

                # Get detailed slide info
                slide_info = self.tracker.get_slide_info()

                # Notify all callbacks about slide change
                for callback in self.slide_change_callbacks:
                    callback(current_slide, self.tracker.total_slides, slide_info)

                return True
        return False
    
    def stop_presentation(self):
        """Stop presentation mode."""
        self.is_presenting = False
        
        if self.tracker:
            self.tracker.disable_auto_sync()
        
        print("Presentation stopped")
    
    def _on_slide_change(self, tracker, window, slide_info):
        """Handle slide change events."""
        current_slide = slide_info.get('current_slide', 1)
        
        # Notify callbacks
        for callback in self.slide_change_callbacks:
            callback(current_slide, self.tracker.total_slides, slide_info)
    
    def get_current_slide_info(self) -> Optional[Dict]:
        """Get information about the current slide."""
        if not self.tracker:
            return None
        
        return self.tracker.get_slide_info()
    
    def get_slide_context(self, context_window: int = 2) -> str:
        """Get contextual information around the current slide."""
        if not self.current_presentation_id or not self.tracker:
            return ""
        
        current_slide = self.tracker.get_current_slide_number()
        return self.vector_store.get_slide_context(
            self.current_presentation_id, 
            current_slide, 
            context_window
        )
    
    def search_slides(self, query: str, top_k: int = 5) -> List[Dict]:
        """Search for slides containing the query."""
        if not self.current_presentation_id:
            return []
        
        results = self.vector_store.search(query, top_k)
        
        # Filter results to current presentation
        presentation_results = []
        for doc_id, similarity, metadata in results:
            if doc_id.startswith(f"{self.current_presentation_id}_slide_"):
                presentation_results.append({
                    'slide_number': metadata.get('slide_number', 0),
                    'similarity': similarity,
                    'text_content': metadata.get('text', ''),
                    'metadata': metadata
                })
        
        return presentation_results
    
    def navigate_to_slide(self, slide_number: int) -> bool:
        """Navigate to a specific slide."""
        if not self.tracker:
            return False
        
        return self.tracker.go_to_slide(slide_number)
    
    def next_slide(self) -> bool:
        """Go to the next slide."""
        if not self.tracker:
            return False
        
        return self.tracker.next_slide()
    
    def previous_slide(self) -> bool:
        """Go to the previous slide."""
        if not self.tracker:
            return False
        
        return self.tracker.previous_slide()
    
    def add_slide_change_callback(self, callback: Callable):
        """Add a callback for slide change events."""
        self.slide_change_callbacks.append(callback)
    
    def add_presentation_load_callback(self, callback: Callable):
        """Add a callback for presentation load events."""
        self.presentation_load_callbacks.append(callback)

    def auto_detect_presentation(self) -> bool:
        """Automatically detect and load a currently open presentation."""
        try:
            # Initialize detector if not already done
            if not self.detector:
                self.detector = PowerPointWindowDetector()

            # Try to auto-detect presentation
            self.tracker = PresentationTracker(None, auto_detect=True)

            if self.tracker.auto_load_presentation():
                self.current_presentation_id = os.path.basename(self.tracker.ppt_path).split('.')[0] if self.tracker.ppt_path else "detected_presentation"

                # Process slides and add to vector store
                self._process_presentation_slides()

                # Notify callbacks
                for callback in self.presentation_load_callbacks:
                    callback(self.current_presentation_id, self.tracker.total_slides)

                print(f"Auto-detected presentation: {self.current_presentation_id}")
                return True

            return False

        except Exception as e:
            print(f"Failed to auto-detect presentation: {e}")
            return False
    
    def get_presentation_summary(self) -> Dict:
        """Get a summary of the current presentation."""
        if not self.tracker:
            return {}
        
        return {
            'presentation_id': self.current_presentation_id,
            'total_slides': self.tracker.total_slides,
            'current_slide': self.tracker.get_current_slide_number(),
            'is_presenting': self.is_presenting,
            'auto_sync_enabled': self.tracker.auto_sync_enabled if self.tracker else False
        }
    
    def clear_presentation(self):
        """Clear the current presentation."""
        if self.current_presentation_id:
            self.vector_store.clear_presentation(self.current_presentation_id)
        
        self.tracker = None
        self.detector = None
        self.current_presentation_id = None
        self.is_presenting = False
        
        print("Cleared presentation")