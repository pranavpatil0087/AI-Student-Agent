import os
import sys
import time
import wave
import signal
from collections import deque
import requests
from dotenv import load_dotenv

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import sounddevice as sd
from vosk import Model, KaldiRecognizer
from notion_integration import notes_to_notion
from gemini_functions import check_and_respond_with_gemini, gemini_make_notes, gemini_transcript

load_dotenv()

STUDENT_NAME = os.getenv("STUDENT_NAME", "Ganesh Jadhav")
NOTION_API_TOKEN = os.getenv("NOTION_API_TOKEN", "")
NOTION_PAGE_ID = os.getenv("NOTION_PAGE_ID", "")
trigger_str = os.getenv("TRIGGER_KEYWORDS", "ganesh,jadhav,attendance,present")
TRIGGER_KEYWORDS = [k.strip().lower() for k in trigger_str.split(",") if k.strip()]

print("=" * 65)
print(" AI Student Agent - Active Meeting Listener (Option B)")
print(f" Student: {STUDENT_NAME}")
print(f" Notion: {'Connected' if (NOTION_API_TOKEN and NOTION_PAGE_ID) else 'Not configured'}")
print(" Mode: Listening to Host/Class audio via Stereo Mix (Speakers)")
print("=" * 65)

# Load VOSK offline model
if not os.path.exists("model"):
    raise ValueError("model folder not found!")
vosk_model = Model("model")
recognizer = KaldiRecognizer(vosk_model, 16000)

is_running = True
recording_frames = []
audio_buffer = b""
buffer_threshold = 16000 * 2  # ~1 second buffer
rolling_history = deque(maxlen=50)

# Select best working audio input device
devices = sd.query_devices()
selected_idx = None
sample_rate = 16000
channels = 1

# Try to find a working device
candidate_indices = []
for idx, dev in enumerate(devices):
    if dev.get("max_input_channels", 0) > 0:
        if "stereo mix" in dev["name"].lower() and dev.get("hostapi") != 3:  # skip WDM-KS exclusive
            candidate_indices.insert(0, idx)
        elif idx == sd.default.device[0]:
            candidate_indices.append(idx)

if sd.default.device[0] not in candidate_indices and sd.default.device[0] >= 0:
    candidate_indices.append(sd.default.device[0])

for idx in candidate_indices:
    dev = devices[idx]
    for ch in [1, min(2, dev.get("max_input_channels", 1))]:
        for sr in [16000, int(dev.get("default_samplerate", 16000))]:
            try:
                test_stream = sd.InputStream(device=idx, channels=ch, samplerate=sr, dtype="int16")
                test_stream.close()
                selected_idx = idx
                channels = ch
                sample_rate = sr
                break
            except Exception:
                continue
        if selected_idx is not None:
            break
    if selected_idx is not None:
        break

if selected_idx is None:
    selected_idx = sd.default.device[0]
    dev = devices[selected_idx]
    sample_rate = int(dev.get("default_samplerate", 16000))
    channels = 1

device_name = devices[selected_idx]["name"]
print(f"[Audio] Capturing audio from: {device_name} (Index: {selected_idx}, Channels: {channels}, SR: {sample_rate}Hz)", flush=True)

def audio_callback(indata, frames, time_info, status):
    global recording_frames, is_running, audio_buffer, rolling_history
    if not is_running:
        return

    # Convert to 16kHz mono bytes
    if sample_rate == 48000:
        if channels >= 2:
            mono_16k = indata[::3, 0].tobytes()
        else:
            mono_16k = indata[::3].tobytes()
    else:
        if channels >= 2:
            mono_16k = indata[:, 0].tobytes()
        else:
            mono_16k = indata.tobytes()

    recording_frames.append(mono_16k)
    audio_buffer += mono_16k

    if len(audio_buffer) >= buffer_threshold:
        chunk = audio_buffer
        audio_buffer = b""

        if recognizer.AcceptWaveform(chunk):
            res = eval(recognizer.Result())
            text = res.get("text", "").strip()
        else:
            partial = eval(recognizer.PartialResult())
            text = partial.get("partial", "").strip()

        if text:
            print(f"[Class Audio Detected]: {text}", flush=True)
            words = text.split()
            rolling_history.extend(words)
            text_lower = text.lower()

            # Check if name or attendance is mentioned
            if any(kw in text_lower for kw in TRIGGER_KEYWORDS):
                context = " ".join(list(rolling_history)[-20:])
                print(f"\n🔔 [ATTENDANCE/QUESTION DETECTED]: '{text}'", flush=True)
                answer = check_and_respond_with_gemini(context)
                print(f"🤖 [Gemini Suggested Response]: {answer}\n", flush=True)

            with open("transcript.txt", "a", encoding="utf-8") as f:
                f.write(text + "\n")

def handle_exit(signum, frame):
    global is_running
    is_running = False

signal.signal(signal.SIGINT, handle_exit)

print("\n>>> LISTENING TO CLASS IN REAL-TIME <<<")
print(">>> Speak in the meeting or let the teacher talk.")
print(">>> Press ENTER anytime when the class finishes to generate Notion notes.\n", flush=True)

stream = sd.InputStream(
    device=selected_idx,
    channels=channels,
    samplerate=sample_rate,
    dtype="int16",
    callback=audio_callback,
)

if os.path.exists("stop_meeting.flag"):
    try:
        os.remove("stop_meeting.flag")
    except Exception:
        pass

with stream:
    try:
        if sys.stdin and sys.stdin.isatty():
            input("Press ENTER when class finishes to generate Notion notes...")
        else:
            while is_running:
                if os.path.exists("stop_meeting.flag"):
                    print("\n[Stop] Stop signal received. Generating notes...")
                    try:
                        os.remove("stop_meeting.flag")
                    except Exception:
                        pass
                    break
                time.sleep(1)
    except (EOFError, KeyboardInterrupt):
        while is_running:
            if os.path.exists("stop_meeting.flag"):
                try:
                    os.remove("stop_meeting.flag")
                except Exception:
                    pass
                break
            time.sleep(1)

is_running = False
print("\n[Notes] Processing class audio and generating notes...", flush=True)

# Save WAV
print("[Audio] Saving session to 'debug_audio.wav'...", flush=True)
with wave.open("debug_audio.wav", "wb") as wf:
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(16000)
    wf.writeframes(b"".join(recording_frames))
print("[Audio] Audio saved successfully.", flush=True)

# Transcribe with Google Gemini
print("\n[AI] Generating high-accuracy transcript with Google Gemini...", flush=True)
transcript = gemini_transcript("debug_audio.wav")
if not transcript and os.path.exists("transcript.txt"):
    with open("transcript.txt", "r", encoding="utf-8") as f:
        transcript = f.read()

if transcript and transcript.strip():
    print("\n" + "=" * 50)
    print(" CLASS TRANSCRIPT:")
    print(transcript.strip())
    print("=" * 50)

    print("\n[AI] Formatting comprehensive notes with Google Gemini...", flush=True)
    notes = gemini_make_notes(transcript)
    print("\n" + "=" * 50)
    print(" GENERATED NOTES:")
    print(notes)
    print("=" * 50)

    # Push to Notion
    if NOTION_API_TOKEN and NOTION_PAGE_ID:
        print("\n[Notion] Uploading notes to Notion...", flush=True)
        notes_to_notion(
            NOTION_API_TOKEN,
            NOTION_PAGE_ID,
            notes,
            title=f"Class Notes - {STUDENT_NAME}",
        )
else:
    print("[Notes] No speech was detected during the session.")

print("\nAgent finished successfully.")
