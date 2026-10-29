import os
import wave
import pyaudio
from gemini_functions import gemini_transcript

def capture_and_transcribe_audio(duration_seconds=10, output_file="debug_audio.wav"):
    """
    Captures microphone audio for a specified duration and transcribes it using Google Gemini API.
    """
    chunk = 1024
    format = pyaudio.paInt16
    channels = 1
    rate = 16000

    p = pyaudio.PyAudio()

    print(f"[Audio] Recording {duration_seconds} seconds of audio from microphone...")
    stream = p.open(format=format,
                    channels=channels,
                    rate=rate,
                    input=True,
                    frames_per_buffer=chunk)

    frames = []
    for _ in range(0, int(rate / chunk * duration_seconds)):
        data = stream.read(chunk)
        frames.append(data)

    print("[Audio] Recording complete. Saving to file...")
    stream.stop_stream()
    stream.close()
    p.terminate()

    with wave.open(output_file, 'wb') as wf:
        wf.setnchannels(channels)
        wf.setsampwidth(p.get_sample_size(format))
        wf.setframerate(rate)
        wf.writeframes(b''.join(frames))

    print(f"[Audio] Audio saved to '{output_file}'. Transcribing with Google Gemini...")
    transcript = gemini_transcript(output_file)
    print(f"[Audio] Gemini Transcript:\n{transcript}")
    return transcript

if __name__ == "__main__":
    capture_and_transcribe_audio(duration_seconds=5)
