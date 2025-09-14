"""
Local LLM client for generating AI assistance and responses.
"""

import json
from typing import Optional, Dict, List, Any
import time


class LLMClient:
    """Client for local language model integration."""
    
    def __init__(self, model_name: str = "local", temperature: float = 0.7):
        self.model_name = model_name
        self.temperature = temperature
        self.model = None
        self.is_loaded = False
        
        self._load_model()
    
    def _load_model(self):
        """Load the local LLM model."""
        try:
            # TODO: Implement actual local LLM loading
            # This could integrate with Ollama, GPT4All, or other local LLM solutions
            # Example: self.model = ollama.load_model(self.model_name)
            print(f"Local LLM '{self.model_name}' loaded successfully")
            self.is_loaded = True
        except Exception as e:
            print(f"Failed to load local LLM: {e}")
            self.is_loaded = False
    
    def generate_response(self, prompt: str, context: Optional[str] = None) -> Optional[str]:
        """Generate a response using the local LLM."""
        if not self.is_loaded:
            return self._fallback_response(prompt)
        
        try:
            # TODO: Implement actual LLM generation
            # response = self.model.generate(prompt, temperature=self.temperature)
            # return response.strip()
            
            # Placeholder implementation
            return self._generate_placeholder_response(prompt, context)
        except Exception as e:
            print(f"LLM generation failed: {e}")
            return self._fallback_response(prompt)
    
    def _generate_placeholder_response(self, prompt: str, context: Optional[str] = None) -> str:
        """Generate a placeholder response for demonstration."""
        # Simple rule-based responses for common presentation scenarios
        prompt_lower = prompt.lower()
        
        if "pause" in prompt_lower or "silence" in prompt_lower:
            return "💡 Consider summarizing the key points you just covered, or ask if the audience has any questions about this topic."
        
        elif "question" in prompt_lower:
            return "❓ That's a great question! Let me think about the best way to address this..."
        
        elif "slide" in prompt_lower and "next" in prompt_lower:
            return "📄 Ready to move to the next slide. You might want to briefly recap the current slide's main points."
        
        elif "confused" in prompt_lower or "unclear" in prompt_lower:
            return "🔄 It sounds like you might need to clarify this point. Consider providing a concrete example or breaking it down into simpler terms."
        
        elif "time" in prompt_lower:
            return "⏰ You're doing well with timing. Consider checking if you need to adjust your pace for the remaining slides."
        
        elif "audience" in prompt_lower:
            return "👥 Great engagement! Keep making eye contact and asking for feedback to maintain audience attention."
        
        else:
            return "I'm here to help with your presentation. Feel free to ask me about slide transitions, timing, or any other presentation concerns."
    
    def _fallback_response(self, prompt: str) -> str:
        """Fallback response when LLM is not available."""
        return "AI assistance is currently unavailable. Please check your local LLM setup."
    
    def generate_slide_notes(self, slide_content: str, slide_number: int) -> str:
        """Generate speaking notes for a slide."""
        if not self.is_loaded:
            return f"📝 Speaking notes for slide {slide_number}: Focus on the key points and maintain good pacing."
        
        try:
            prompt = f"""
            Generate concise speaking notes for slide {slide_number} with the following content:
            
            {slide_content}
            
            Provide 2-3 key talking points that would help a presenter deliver this slide effectively.
            """
            
            response = self.generate_response(prompt)
            return f"📝 Slide {slide_number} Notes:\n{response}"
        except Exception as e:
            print(f"Failed to generate slide notes: {e}")
            return f"📝 Slide {slide_number}: Review the content and prepare your key talking points."
    
    def analyze_presentation_pace(self, transcript: str, duration: float) -> Dict[str, Any]:
        """Analyze presentation pacing and provide feedback."""
        word_count = len(transcript.split())
        wpm = (word_count / duration) * 60 if duration > 0 else 0
        
        analysis = {
            'words_per_minute': round(wpm, 1),
            'total_words': word_count,
            'duration_minutes': round(duration / 60, 1),
            'pace_feedback': self._get_pace_feedback(wpm),
            'recommendations': self._get_pace_recommendations(wpm)
        }
        
        return analysis
    
    def _get_pace_feedback(self, wpm: float) -> str:
        """Get feedback based on speaking pace."""
        if wpm < 120:
            return "Your pace is quite slow. Consider speaking a bit faster to maintain audience engagement."
        elif wpm > 180:
            return "Your pace is quite fast. Consider slowing down to help the audience follow along."
        else:
            return "Your speaking pace is good for presentations. Keep it up!"
    
    def _get_pace_recommendations(self, wpm: float) -> List[str]:
        """Get recommendations based on speaking pace."""
        recommendations = []
        
        if wpm < 120:
            recommendations.extend([
                "Practice speaking slightly faster",
                "Use more dynamic gestures to add energy",
                "Consider adding pauses for emphasis rather than slowing down"
            ])
        elif wpm > 180:
            recommendations.extend([
                "Practice taking deliberate pauses",
                "Focus on clear articulation",
                "Use visual aids to help audience follow along"
            ])
        else:
            recommendations.extend([
                "Maintain your current pace",
                "Continue using natural pauses",
                "Keep engaging with the audience"
            ])
        
        return recommendations
    
    def detect_filler_words(self, transcript: str) -> Dict[str, Any]:
        """Detect and analyze filler words in the transcript."""
        filler_words = ['um', 'uh', 'like', 'you know', 'so', 'well', 'actually', 'basically']
        
        words = transcript.lower().split()
        total_words = len(words)
        filler_count = sum(1 for word in words if word in filler_words)
        filler_percentage = (filler_count / total_words) * 100 if total_words > 0 else 0
        
        detected_fillers = {}
        for filler in filler_words:
            count = words.count(filler)
            if count > 0:
                detected_fillers[filler] = count
        
        return {
            'total_fillers': filler_count,
            'filler_percentage': round(filler_percentage, 1),
            'detected_fillers': detected_fillers,
            'feedback': self._get_filler_feedback(filler_percentage)
        }
    
    def _get_filler_feedback(self, percentage: float) -> str:
        """Get feedback based on filler word usage."""
        if percentage < 2:
            return "Excellent! You're using very few filler words."
        elif percentage < 5:
            return "Good job! Your filler word usage is within acceptable limits."
        elif percentage < 10:
            return "Consider reducing filler words. Practice pausing instead of using 'um' or 'uh'."
        else:
            return "High filler word usage detected. Focus on deliberate pauses and slower speech."
