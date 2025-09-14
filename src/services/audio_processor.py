"""
Audio Processing and Transcription Service

Handles audio input capture, processing, and transcription for LLM integration.
"""

import logging
import threading
import time
import queue
from typing import Optional, Callable, Dict, Any
from dataclasses import dataclass
import numpy as np

try:
    import pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False
    logging.warning("PyAudio not available. Audio capture will be disabled.")

try:
    import whisper
    WHISPER_AVAILABLE = True
except ImportError:
    WHISPER_AVAILABLE = False
    logging.warning("OpenAI Whisper not available. Transcription will be disabled.")

@dataclass
class AudioConfig:
    """Audio capture configuration."""
    sample_rate: int = 16000
    chunk_size: int = 1024
    channels: int = 1
    format: int = pyaudio.paInt16 if PYAUDIO_AVAILABLE else None
    buffer_seconds: int = 30
    transcription_interval: float = 5.0  # Transcribe every N seconds

@dataclass
class TranscriptionResult:
    """Result of audio transcription."""
    text: str
    confidence: float
    timestamp: float
    duration: float
    language: str = "en"

class AudioProcessor:
    """Service for capturing and processing audio input."""

    def __init__(self, config: AudioConfig = None):
        self.config = config or AudioConfig()
        self.logger = logging.getLogger(__name__)

        # Audio capture
        self.audio = None
        self.stream = None
        self.is_recording = False
        self.audio_buffer = queue.Queue()

        # Transcription
        self.whisper_model = None
        self.transcription_callback: Optional[Callable[[TranscriptionResult], None]] = None

        # Threading
        self.capture_thread = None
        self.transcription_thread = None
        self.stop_event = threading.Event()

        self._initialize_audio()
        self._initialize_whisper()

    def _initialize_audio(self):
        """Initialize PyAudio for audio capture."""
        if not PYAUDIO_AVAILABLE:
            self.logger.warning("PyAudio not available. Audio capture disabled.")
            return

        try:
            self.audio = pyaudio.PyAudio()

            # Check for available input devices
            device_count = self.audio.get_device_count()
            self.logger.info(f"Found {device_count} audio devices")

            # Find default input device
            default_input = self.audio.get_default_input_device_info()
            self.logger.info(f"Default input device: {default_input['name']}")

        except Exception as e:
            self.logger.error(f"Failed to initialize audio: {e}")
            self.audio = None

    def _initialize_whisper(self):
        """Initialize Whisper model for transcription."""
        if not WHISPER_AVAILABLE:
            self.logger.warning("Whisper not available. Transcription disabled.")
            return

        try:
            self.logger.info("Loading Whisper model...")
            self.whisper_model = whisper.load_model("base")
            self.logger.info("Whisper model loaded successfully")
        except Exception as e:
            self.logger.error(f"Failed to load Whisper model: {e}")
            self.whisper_model = None

    def start_recording(self, transcription_callback: Optional[Callable[[TranscriptionResult], None]] = None):
        """Start audio recording and transcription."""
        if not self.audio or not self.whisper_model:
            self.logger.error("Audio or transcription not available")
            return False

        if self.is_recording:
            self.logger.warning("Already recording")
            return True

        try:
            self.transcription_callback = transcription_callback
            self.stop_event.clear()

            # Start audio stream
            self.stream = self.audio.open(
                format=self.config.format,
                channels=self.config.channels,
                rate=self.config.sample_rate,
                input=True,
                frames_per_buffer=self.config.chunk_size,
                stream_callback=self._audio_callback
            )

            self.stream.start_stream()

            # Start processing threads
            self.capture_thread = threading.Thread(target=self._audio_capture_loop)
            self.transcription_thread = threading.Thread(target=self._transcription_loop)

            self.capture_thread.start()
            self.transcription_thread.start()

            self.is_recording = True
            self.logger.info("Audio recording started")
            return True

        except Exception as e:
            self.logger.error(f"Failed to start recording: {e}")
            self.stop_recording()
            return False

    def stop_recording(self):
        """Stop audio recording and transcription."""
        if not self.is_recording:
            return

        self.logger.info("Stopping audio recording...")
        self.stop_event.set()
        self.is_recording = False

        # Stop audio stream
        if self.stream:
            try:
                self.stream.stop_stream()
                self.stream.close()
            except Exception as e:
                self.logger.error(f"Error stopping audio stream: {e}")
            finally:
                self.stream = None

        # Wait for threads to finish
        if self.capture_thread and self.capture_thread.is_alive():
            self.capture_thread.join(timeout=2.0)

        if self.transcription_thread and self.transcription_thread.is_alive():
            self.transcription_thread.join(timeout=2.0)

        self.logger.info("Audio recording stopped")

    def _audio_callback(self, in_data, frame_count, time_info, status):
        """PyAudio callback for audio data."""
        if status:
            self.logger.warning(f"Audio callback status: {status}")

        # Convert to numpy array
        audio_data = np.frombuffer(in_data, dtype=np.int16)

        # Add to buffer
        if not self.audio_buffer.full():
            self.audio_buffer.put((audio_data, time.time()))

        return (None, pyaudio.paContinue)

    def _audio_capture_loop(self):
        """Main audio capture processing loop."""
        self.logger.info("Audio capture loop started")

        while not self.stop_event.is_set():
            try:
                time.sleep(0.1)  # Small delay to prevent excessive CPU usage
            except Exception as e:
                self.logger.error(f"Error in audio capture loop: {e}")
                break

        self.logger.info("Audio capture loop stopped")

    def _transcription_loop(self):
        """Main transcription processing loop."""
        self.logger.info("Transcription loop started")

        audio_chunks = []
        last_transcription = time.time()

        while not self.stop_event.is_set():
            try:
                # Collect audio data
                try:
                    audio_data, timestamp = self.audio_buffer.get(timeout=0.5)
                    audio_chunks.append(audio_data)
                except queue.Empty:
                    continue

                # Check if it's time to transcribe
                current_time = time.time()
                if (current_time - last_transcription) >= self.config.transcription_interval:
                    if audio_chunks:
                        self._transcribe_audio_chunks(audio_chunks, current_time)
                        audio_chunks = []
                        last_transcription = current_time

            except Exception as e:
                self.logger.error(f"Error in transcription loop: {e}")
                continue

        # Final transcription of remaining chunks
        if audio_chunks:
            self._transcribe_audio_chunks(audio_chunks, time.time())

        self.logger.info("Transcription loop stopped")

    def _transcribe_audio_chunks(self, audio_chunks: list, timestamp: float):
        """Transcribe collected audio chunks."""
        if not audio_chunks or not self.whisper_model:
            return

        try:
            # Combine audio chunks
            combined_audio = np.concatenate(audio_chunks)

            # Convert to float32 and normalize
            audio_float = combined_audio.astype(np.float32) / 32768.0

            # Transcribe with Whisper
            result = self.whisper_model.transcribe(audio_float)

            if result and result.get("text", "").strip():
                transcription_result = TranscriptionResult(
                    text=result["text"].strip(),
                    confidence=1.0,  # Whisper doesn't provide confidence scores
                    timestamp=timestamp,
                    duration=len(combined_audio) / self.config.sample_rate,
                    language=result.get("language", "en")
                )

                self.logger.info(f"Transcription: {transcription_result.text[:100]}...")

                # Call callback if provided
                if self.transcription_callback:
                    try:
                        self.transcription_callback(transcription_result)
                    except Exception as e:
                        self.logger.error(f"Error in transcription callback: {e}")

        except Exception as e:
            self.logger.error(f"Transcription failed: {e}")

    def is_audio_available(self) -> bool:
        """Check if audio capture is available."""
        return PYAUDIO_AVAILABLE and self.audio is not None

    def is_transcription_available(self) -> bool:
        """Check if transcription is available."""
        return WHISPER_AVAILABLE and self.whisper_model is not None

    def get_audio_devices(self) -> Dict[int, Dict[str, Any]]:
        """Get available audio input devices."""
        devices = {}
        if not self.audio:
            return devices

        try:
            for i in range(self.audio.get_device_count()):
                device_info = self.audio.get_device_info_by_index(i)
                if device_info.get('maxInputChannels', 0) > 0:
                    devices[i] = {
                        'name': device_info['name'],
                        'channels': device_info['maxInputChannels'],
                        'sample_rate': device_info['defaultSampleRate']
                    }
        except Exception as e:
            self.logger.error(f"Error getting audio devices: {e}")

        return devices

    def __del__(self):
        """Cleanup on object destruction."""
        self.stop_recording()
        if self.audio:
            self.audio.terminate()

# Service instance
audio_processor = AudioProcessor()