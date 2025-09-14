"""
AI service for coordinating AI/ML functionality.
"""

import threading
import time
from typing import Optional, Dict, List, Callable, Any

# Import core modules with fallback
try:
    from ..core.ai import WhisperClient, VectorStore, LLMClient
except ImportError:
    # Fallback for when running from different contexts
    import sys
    from pathlib import Path
    core_path = Path(__file__).parent.parent / "core"
    sys.path.insert(0, str(core_path))
    from ai import WhisperClient, VectorStore, LLMClient


class AIService:
    """Service for coordinating AI/ML operations."""
    
    def __init__(self):
        self.whisper_client = WhisperClient()
        self.vector_store = VectorStore()
        self.llm_client = LLMClient()
        
        self.is_listening = False
        self.transcription_buffer = []
        self.last_transcription_time = 0
        self.silence_threshold = 2.0  # seconds
        
        # Callbacks
        self.transcription_callbacks: List[Callable] = []
        self.assistance_callbacks: List[Callable] = []
        self.analysis_callbacks: List[Callable] = []
        
        # Analysis data
        self.presentation_metrics = {
            'total_words': 0,
            'speaking_time': 0,
            'filler_words': 0,
            'questions_answered': 0,
            'confidence_scores': []
        }
    
    def start_listening(self) -> bool:
        """Start real-time speech transcription."""
        if self.is_listening:
            return True
        
        success = self.whisper_client.start_listening(self._on_transcription)
        if success:
            self.is_listening = True
            self.last_transcription_time = time.time()
            print("Started AI listening service")
        
        return success
    
    def stop_listening(self):
        """Stop speech transcription."""
        if self.is_listening:
            self.whisper_client.stop_listening()
            self.is_listening = False
            print("Stopped AI listening service")
    
    def _on_transcription(self, transcription: str):
        """Handle transcription results."""
        current_time = time.time()
        
        # Update metrics
        self.presentation_metrics['total_words'] += len(transcription.split())
        self.presentation_metrics['speaking_time'] += current_time - self.last_transcription_time
        
        # Add to buffer
        self.transcription_buffer.append({
            'text': transcription,
            'timestamp': current_time,
            'confidence': self.whisper_client.get_confidence_score(transcription)
        })
        
        # Keep only recent transcriptions
        cutoff_time = current_time - 30  # Keep last 30 seconds
        self.transcription_buffer = [
            t for t in self.transcription_buffer if t['timestamp'] > cutoff_time
        ]
        
        # Notify callbacks
        for callback in self.transcription_callbacks:
            callback(transcription, current_time)
        
        # Check for assistance triggers
        self._check_assistance_triggers(transcription, current_time)
        
        self.last_transcription_time = current_time
    
    def _check_assistance_triggers(self, transcription: str, timestamp: float):
        """Check if assistance should be triggered."""
        # Detect pauses (silence)
        if timestamp - self.last_transcription_time > self.silence_threshold:
            self._trigger_assistance("pause_detected", transcription)
        
        # Detect questions
        question_indicators = ['?', 'question', 'ask', 'wonder', 'curious']
        if any(indicator in transcription.lower() for indicator in question_indicators):
            self._trigger_assistance("question_detected", transcription)
        
        # Detect confusion or uncertainty
        confusion_indicators = ['confused', 'unclear', 'not sure', 'don\'t understand']
        if any(indicator in transcription.lower() for indicator in confusion_indicators):
            self._trigger_assistance("confusion_detected", transcription)
    
    def _trigger_assistance(self, trigger_type: str, context: str):
        """Trigger AI assistance based on context."""
        try:
            # Get current presentation context
            presentation_context = self._get_presentation_context()
            
            # Generate assistance
            assistance = self._generate_assistance(trigger_type, context, presentation_context)
            
            if assistance:
                # Notify callbacks
                for callback in self.assistance_callbacks:
                    callback(assistance, trigger_type, context)
                
                print(f"AI Assistance triggered ({trigger_type}): {assistance[:100]}...")
        
        except Exception as e:
            print(f"Failed to trigger assistance: {e}")
    
    def _get_presentation_context(self) -> str:
        """Get current presentation context for AI assistance."""
        # This would integrate with the presentation service
        # For now, return a placeholder
        return "Current presentation context would be provided here."
    
    def _generate_assistance(self, trigger_type: str, context: str, presentation_context: str) -> Optional[str]:
        """Generate AI assistance based on trigger and context."""
        try:
            if trigger_type == "pause_detected":
                prompt = f"""
                The presenter has paused. Based on the current presentation context:
                {presentation_context}
                
                And the recent speech:
                {context}
                
                Provide a helpful suggestion to continue the presentation smoothly.
                """
            
            elif trigger_type == "question_detected":
                prompt = f"""
                The presenter seems to be asking a question or addressing audience questions.
                Based on the presentation context:
                {presentation_context}
                
                Provide helpful guidance on how to handle this question effectively.
                """
            
            elif trigger_type == "confusion_detected":
                prompt = f"""
                The presenter seems confused or uncertain about something.
                Based on the presentation context:
                {presentation_context}
                
                Provide helpful clarification or guidance.
                """
            
            else:
                return None
            
            return self.llm_client.generate_response(prompt, presentation_context)
        
        except Exception as e:
            print(f"Failed to generate assistance: {e}")
            return None
    
    def generate_slide_notes(self, slide_content: str, slide_number: int) -> str:
        """Generate speaking notes for a slide."""
        return self.llm_client.generate_slide_notes(slide_content, slide_number)
    
    def analyze_presentation_performance(self) -> Dict[str, Any]:
        """Analyze overall presentation performance."""
        if not self.transcription_buffer:
            return {}
        
        # Combine recent transcriptions
        recent_text = ' '.join([t['text'] for t in self.transcription_buffer])
        total_duration = self.presentation_metrics['speaking_time']
        
        # Analyze pace
        pace_analysis = self.llm_client.analyze_presentation_pace(recent_text, total_duration)
        
        # Analyze filler words
        filler_analysis = self.llm_client.detect_filler_words(recent_text)
        
        # Calculate average confidence
        confidences = [t['confidence'] for t in self.transcription_buffer if 'confidence' in t]
        avg_confidence = sum(confidences) / len(confidences) if confidences else 0
        
        analysis = {
            'pace_analysis': pace_analysis,
            'filler_analysis': filler_analysis,
            'average_confidence': round(avg_confidence, 2),
            'total_speaking_time': round(total_duration, 1),
            'total_words': self.presentation_metrics['total_words'],
            'recent_transcriptions': len(self.transcription_buffer)
        }
        
        # Notify analysis callbacks
        for callback in self.analysis_callbacks:
            callback(analysis)
        
        return analysis
    
    def search_presentation_content(self, query: str, top_k: int = 5) -> List[Dict]:
        """Search presentation content using vector store."""
        return self.vector_store.search(query, top_k)
    
    def add_transcription_callback(self, callback: Callable):
        """Add a callback for transcription events."""
        self.transcription_callbacks.append(callback)
    
    def add_assistance_callback(self, callback: Callable):
        """Add a callback for AI assistance events."""
        self.assistance_callbacks.append(callback)
    
    def add_analysis_callback(self, callback: Callable):
        """Add a callback for analysis events."""
        self.analysis_callbacks.append(callback)
    
    def get_current_metrics(self) -> Dict[str, Any]:
        """Get current presentation metrics."""
        return {
            'is_listening': self.is_listening,
            'transcription_buffer_size': len(self.transcription_buffer),
            'metrics': self.presentation_metrics.copy()
        }
    
    def reset_metrics(self):
        """Reset presentation metrics."""
        self.presentation_metrics = {
            'total_words': 0,
            'speaking_time': 0,
            'filler_words': 0,
            'questions_answered': 0,
            'confidence_scores': []
        }
        self.transcription_buffer = []
        print("🔄 Reset AI service metrics")
