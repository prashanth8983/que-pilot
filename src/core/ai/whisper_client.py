"""
Whisper speech-to-text client for real-time transcription.
"""

import threading
import queue
import time
import io
import wave
from typing import Optional, Callable, Dict, Any
import numpy as np

class WhisperClient:
    """Client for OpenAI Whisper speech-to-text functionality."""

    def __init__(self, model_size: str = "base", confidence_threshold: float = 0.75):
        self.model_size = model_size
        self.confidence_threshold = confidence_threshold
        self.is_listening = False
        self.audio_queue = queue.Queue()
        self.transcription_callback: Optional[Callable] = None
        self.error_callback: Optional[Callable] = None

        # Audio configuration
        self.sample_rate = 16000
        self.chunk_size = 1024
        self.audio_format = None  # Will be set when pyaudio is initialized

        # Initialize Whisper model and audio
        self.model = None
        self.audio_interface = None
        self.audio_stream = None
        self._load_dependencies()
        self._load_model()

    def _load_dependencies(self):
        """Load required dependencies."""
        try:
            global whisper, pyaudio
            import whisper
            import pyaudio
            self.audio_format = pyaudio.paFloat32
            print("✅ Audio dependencies loaded successfully")
        except ImportError as e:
            print(f"❌ Failed to load audio dependencies: {e}")
            print("Install with: pip install openai-whisper pyaudio")

    def _load_model(self):
        """Load the Whisper model."""
        try:
            if whisper is None:
                print("❌ Whisper not available")
                return

            print(f"🔄 Loading Whisper model '{self.model_size}'...")
            self.model = whisper.load_model(self.model_size)
            print(f"✅ Whisper model '{self.model_size}' loaded successfully")
        except Exception as e:
            print(f"❌ Failed to load Whisper model: {e}")
            self.model = None

    def _init_audio_interface(self):
        """Initialize PyAudio interface."""
        try:
            if pyaudio is None:
                return False

            self.audio_interface = pyaudio.PyAudio()
            return True
        except Exception as e:
            print(f"❌ Failed to initialize audio interface: {e}")
            return False

    def start_listening(self, callback: Optional[Callable] = None) -> bool:
        """Start real-time speech transcription."""
        if not self.model:
            print("❌ Whisper model not loaded")
            return False

        if not self._init_audio_interface():
            print("❌ Audio interface not available")
            return False

        try:
            # Open audio stream
            self.audio_stream = self.audio_interface.open(
                format=self.audio_format,
                channels=1,
                rate=self.sample_rate,
                input=True,
                frames_per_buffer=self.chunk_size,
                stream_callback=self._audio_callback
            )

            self.transcription_callback = callback
            self.is_listening = True

            # Start transcription thread
            self.transcription_thread = threading.Thread(target=self._transcription_loop)
            self.transcription_thread.daemon = True
            self.transcription_thread.start()

            # Start audio stream
            self.audio_stream.start_stream()

            print("🎤 Started listening for speech...")
            return True

        except Exception as e:
            print(f"❌ Failed to start listening: {e}")
            return False

    def stop_listening(self):
        """Stop speech transcription."""
        self.is_listening = False

        try:
            if self.audio_stream:
                self.audio_stream.stop_stream()
                self.audio_stream.close()

            if self.audio_interface:
                self.audio_interface.terminate()

            print("🛑 Stopped listening")
        except Exception as e:
            print(f"⚠️ Error stopping audio: {e}")

    def _audio_callback(self, in_data, frame_count, time_info, status):
        """PyAudio callback for incoming audio data."""
        if self.is_listening:
            # Convert audio data to numpy array
            audio_data = np.frombuffer(in_data, dtype=np.float32)
            self.audio_queue.put(audio_data)

        return (in_data, pyaudio.paContinue)

    def _transcription_loop(self):
        """Main transcription processing loop."""
        audio_buffer = []
        buffer_duration = 3.0  # seconds
        last_transcription_time = time.time()
        silence_threshold = 1.0  # seconds of silence before processing

        while self.is_listening:
            try:
                # Collect audio data with timeout
                try:
                    audio_chunk = self.audio_queue.get(timeout=0.1)
                    audio_buffer.append(audio_chunk)
                    last_transcription_time = time.time()
                except queue.Empty:
                    # Check if we should process accumulated audio due to silence
                    if (audio_buffer and
                        time.time() - last_transcription_time > silence_threshold):
                        self._process_audio_buffer(audio_buffer)
                        audio_buffer = []
                    continue

                # Process buffer when it reaches target duration
                buffer_length = len(audio_buffer) * (self.chunk_size / self.sample_rate)
                if buffer_length >= buffer_duration:
                    self._process_audio_buffer(audio_buffer)
                    audio_buffer = []

            except Exception as e:
                if self.error_callback:
                    self.error_callback(f"Transcription error: {e}")
                else:
                    print(f"❌ Transcription error: {e}")

    def _process_audio_buffer(self, audio_buffer):
        """Process accumulated audio buffer."""
        if not audio_buffer:
            return

        try:
            # Concatenate audio chunks
            full_audio = np.concatenate(audio_buffer)

            # Skip if audio is too quiet (likely silence)
            if np.max(np.abs(full_audio)) < 0.01:
                return

            # Transcribe audio
            transcription = self._transcribe_audio(full_audio)

            if transcription and transcription.strip() and self.transcription_callback:
                self.transcription_callback(transcription.strip())

        except Exception as e:
            print(f"❌ Error processing audio buffer: {e}")

    def _transcribe_audio(self, audio_data: np.ndarray) -> Optional[str]:
        """Transcribe audio data to text using Whisper."""
        try:
            # Ensure audio is in the correct format for Whisper
            if len(audio_data.shape) > 1:
                audio_data = np.mean(audio_data, axis=1)

            # Normalize audio
            if np.max(np.abs(audio_data)) > 0:
                audio_data = audio_data / np.max(np.abs(audio_data))

            # Transcribe with Whisper
            result = self.model.transcribe(
                audio_data,
                language=None,  # Auto-detect language
                task="transcribe",
                verbose=False
            )

            text = result["text"].strip()

            # Filter out very short or repetitive transcriptions
            if len(text) < 3 or text in ["", " ", "you", "Thank you."]:
                return None

            return text

        except Exception as e:
            print(f"❌ Transcription failed: {e}")
            return None

    def get_speaking_pace(self, text: str, duration_seconds: float) -> float:
        """Calculate speaking pace in words per minute."""
        if duration_seconds <= 0:
            return 0.0

        word_count = len(text.split())
        wpm = (word_count / duration_seconds) * 60
        return round(wpm, 1)

    def get_confidence_score(self, transcription: str) -> float:
        """Calculate confidence score for transcription."""
        # Simple heuristic-based confidence scoring
        # In a full implementation, this could use Whisper's attention weights

        if not transcription or len(transcription.strip()) < 3:
            return 0.3

        # Longer transcriptions tend to be more confident
        length_factor = min(len(transcription) / 50, 1.0)

        # Check for common filler words that indicate uncertainty
        filler_words = ["um", "uh", "er", "ah", "like", "you know"]
        word_count = len(transcription.split())
        filler_count = sum(1 for word in transcription.lower().split() if word in filler_words)
        filler_penalty = filler_count / word_count if word_count > 0 else 0

        # Calculate confidence (0.5-0.95 range)
        confidence = 0.5 + (length_factor * 0.4) - (filler_penalty * 0.2)
        return max(0.5, min(0.95, confidence))

    def transcribe_audio_file(self, file_path: str) -> Dict[str, Any]:
        """Transcribe an audio file (for batch processing)."""
        try:
            if not self.model:
                return {"error": "Whisper model not loaded"}

            result = self.model.transcribe(file_path)
            return {
                "text": result["text"],
                "segments": result["segments"],
                "language": result["language"]
            }
        except Exception as e:
            return {"error": str(e)}

    def get_available_models(self) -> list:
        """Get list of available Whisper models."""
        return ["tiny", "base", "small", "medium", "large"]

    def change_model(self, model_size: str) -> bool:
        """Change the Whisper model size."""
        if model_size not in self.get_available_models():
            print(f"❌ Invalid model size: {model_size}")
            return False

        try:
            print(f"🔄 Switching to model '{model_size}'...")
            self.model_size = model_size
            self._load_model()
            return True
        except Exception as e:
            print(f"❌ Failed to change model: {e}")
            return False