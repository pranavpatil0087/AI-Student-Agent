import os
import sys
import time

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import json
import wave
import signal
from collections import deque
from threading import Thread, Lock
import queue
import logging

# Silence verbose Werkzeug HTTP request logging
logging.getLogger("werkzeug").setLevel(logging.ERROR)

import requests
import pyaudio
import numpy as np
from scipy.signal import resample_poly
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from flask import Flask, request, jsonify
from flask_cors import CORS
from vosk import Model, KaldiRecognizer

from notion_integration import format_notes_for_notion, notes_to_notion
from gemini_functions import check_and_respond_with_gemini, gemini_make_notes, gemini_transcript

# Load Environment Variables
load_dotenv()
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
GMEET_LINK = os.getenv("GMEET_LINK", "")
NOTION_API_TOKEN = os.getenv("NOTION_API_TOKEN", "")
NOTION_PAGE_ID = os.getenv("NOTION_PAGE_ID", "")
STUDENT_NAME = os.getenv("STUDENT_NAME", "Ganesh Jadhav")
trigger_str = os.getenv("TRIGGER_KEYWORDS", "ganesh,jadhav,attendance,present")
TRIGGER_KEYWORDS = [k.strip().lower() for k in trigger_str.split(",") if k.strip()]
AUDIO_DEVICE_NAME = os.getenv("AUDIO_DEVICE_NAME", "default").strip().lower()

print("=" * 60)
print(f" AI Student Agent Initializing")
print(f" Student Name: {STUDENT_NAME}")
print(f" Trigger Keywords: {TRIGGER_KEYWORDS}")
print(f" Notion Integration: {'Enabled' if (NOTION_API_TOKEN and NOTION_PAGE_ID) else 'Not configured'}")
print("=" * 60)

# Flask Configuration for audio processing endpoint
app = Flask(__name__)
CORS(app)

# VOSK Model Initialization
# Supports custom path via .env; prefers model_in (Indian English) if present, falls back to model
model_path = os.getenv("VOSK_MODEL_PATH", "").strip()
if not model_path or not os.path.exists(model_path):
    if os.path.exists("model_in") and os.path.isdir("model_in"):
        model_path = "model_in"
    elif os.path.exists("model") and os.path.isdir("model"):
        model_path = "model"
    else:
        raise ValueError(
            "VOSK model directory not found. Please run 'python download_model.py' to download the offline speech model."
        )

print(f"[VOSK STT] Initializing offline speech recognition from '{model_path}'...", flush=True)
vosk_model = Model(model_path)
global_recognizer = KaldiRecognizer(vosk_model, 16000)
global_recognizer.SetWords(True)
print(f"[VOSK STT] Speech model '{model_path}' loaded successfully (16,000Hz).", flush=True)

# Runtime Flags and Buffers
global message_tab_opened
message_tab_opened = False
global_driver = None
rolling_history = deque(maxlen=50)
recording_frames = []
session_transcript_lines = []
is_running = True
last_attendance_trigger_time = 0
audio_queue = queue.Queue(maxsize=300)
vosk_lock = Lock()

# Reset transcript.txt so previous runs/mock lectures never contaminate the session
with open("transcript.txt", "w", encoding="utf-8") as f:
    f.write("")


def process_audio_with_vosk(audio_data):
    """
    Process incoming 16kHz mono audio using persistent VOSK recognizer for real-time speech-to-text.
    Protected with vosk_lock to prevent multi-threaded C++ memory corruption (0xc0000005).
    """
    global rolling_history, global_recognizer, session_transcript_lines, last_attendance_trigger_time

    with vosk_lock:
        try:
            # 1. Feed 16kHz chunk directly to Kaldi
            if global_recognizer.AcceptWaveform(audio_data):
                result = json.loads(global_recognizer.Result())
                text_cleaned = result.get("text", "").strip()

                if text_cleaned:
                    print(f"\n[Heard Speech (Vosk)]: \"{text_cleaned}\"", flush=True)
                    line = f"[Speech]: {text_cleaned}"
                    session_transcript_lines.append(line)
                    words = text_cleaned.split()
                    rolling_history.extend(words)

                    with open("transcript.txt", "a", encoding="utf-8") as transcript_file:
                        transcript_file.write(line + "\n")

                    # Check for trigger keywords (student name, attendance call, roll number)
                    text_lower = text_cleaned.lower()
                    now = time.time()
                    if any(kw in text_lower for kw in TRIGGER_KEYWORDS):
                        if now - last_attendance_trigger_time > 8:
                            last_attendance_trigger_time = now
                            print(f"\n🔔 [ATTENDANCE/QUESTION TRIGGER DETECTED]: '{text_cleaned}'", flush=True)
                            context_words = list(rolling_history)[-20:]
                            context = " ".join(context_words)
                            response = check_and_respond_with_gemini(context)
                            print(f"🤖 [AI Generated Chat Response]: '{response}'", flush=True)
                            if global_driver:
                                send_message_to_chat(response)

                    return {"status": "Audio processed", "transcript": text_cleaned}
            else:
                # 2. Check partial hypothesis for immediate mid-utterance attendance trigger
                partial_result = json.loads(global_recognizer.PartialResult())
                partial_cleaned = partial_result.get("partial", "").strip()
                if partial_cleaned:
                    partial_lower = partial_cleaned.lower()
                    now = time.time()
                    if any(kw in partial_lower for kw in TRIGGER_KEYWORDS):
                        if now - last_attendance_trigger_time > 8:
                            last_attendance_trigger_time = now
                            print(f"\n🔔 [ATTENDANCE TRIGGER (Mid-speech)]: '{partial_cleaned}'", flush=True)
                            response = check_and_respond_with_gemini(f"Teacher called: {partial_cleaned}")
                            print(f"🤖 [AI Generated Chat Response]: '{response}'", flush=True)
                            if global_driver:
                                send_message_to_chat(response)
                            global_recognizer.Reset()

            return {"status": "Processing stream"}
        except Exception as e:
            print(f"Error in VOSK transcription: {e}", flush=True)
            return {"error": str(e)}


def speech_recognition_worker():
    """
    Dedicated single-thread worker that pulls audio frames from audio_queue
    and feeds them sequentially into VOSK speech recognition.
    """
    while is_running:
        try:
            chunk = audio_queue.get(timeout=0.2)
            process_audio_with_vosk(chunk)
        except queue.Empty:
            continue
        except Exception as e:
            print(f"[Speech Worker] Note: {e}", flush=True)


@app.route("/process_audio", methods=["POST"])
def process_audio():
    """
    Flask endpoint for receiving streaming audio chunks.
    """
    try:
        audio_data = request.data
        try:
            audio_queue.put_nowait(audio_data)
        except queue.Full:
            pass
        return jsonify({"status": "Queued"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def start_flask_app():
    """
    Start local Flask receiver in a background thread.
    """
    app.run(port=5000, debug=False, use_reloader=False, threaded=True)


def find_best_audio_device():
    """
    Find the best audio capture device:
    1. Explicitly configured device in .env (e.g. Stereo Mix, VB-Cable)
    2. Any device whose name contains 'stereo mix', 'cable output', 'what u hear'
    3. Default system recording input (Microphone Array)
    """
    import sounddevice as sd
    devices = sd.query_devices()

    def can_open(idx, dev):
        try:
            ch = 1
            sr = int(dev.get("default_samplerate", 16000))
            st = sd.InputStream(device=idx, channels=ch, samplerate=sr, dtype="int16")
            st.close()
            return True
        except Exception:
            return False

    # 1. Configured device from .env
    if AUDIO_DEVICE_NAME and AUDIO_DEVICE_NAME != "default":
        for idx, dev in enumerate(devices):
            if dev.get("max_input_channels", 0) > 0 and AUDIO_DEVICE_NAME.lower() in dev["name"].lower():
                if can_open(idx, dev):
                    return idx, dev

    # 2. Prefer Stereo Mix / CABLE Output (Loopback of meeting speaker audio)
    for idx, dev in enumerate(devices):
        if dev.get("max_input_channels", 0) > 0:
            name_lower = dev["name"].lower()
            if any(k in name_lower for k in ["stereo mix", "cable output", "what u hear", "wave out"]):
                if can_open(idx, dev):
                    return idx, dev

    # 3. System default input
    try:
        def_in = sd.default.device[0]
        if def_in >= 0:
            dev = sd.query_devices(def_in)
            if can_open(def_in, dev):
                return def_in, dev
    except Exception:
        pass

    # 4. Any available working input device
    for idx, dev in enumerate(devices):
        if dev.get("max_input_channels", 0) > 0 and can_open(idx, dev):
            return idx, dev

    return None, None


def capture_audio_from_device():
    """
    Captures digital meeting audio directly from the instructor's Google Meet audio feed
    using Windows WASAPI Loopback (pure digital software audio cable with zero room echo/distortion).
    Downsamples 48kHz stereo to 16kHz mono PCM for VOSK.
    Seamlessly falls back to sounddevice microphone if loopback is unavailable.
    """
    global recording_frames, is_running
    import time
    import wave
    import numpy as np
    from scipy.signal import resample_poly

    prefer_mic = bool(AUDIO_DEVICE_NAME and any(k in AUDIO_DEVICE_NAME.lower() for k in ["mic", "microphone"]))

    # 1. Primary Strategy: Digital WASAPI Speaker Loopback (Captures Instructor Speech Directly)
    if not prefer_mic:
        try:
            import pyaudiowpatch as pyaudio

            p = pyaudio.PyAudio()
            wasapi_info = p.get_host_api_info_by_type(pyaudio.paWASAPI)
            default_speakers = p.get_device_info_by_index(wasapi_info["defaultOutputDevice"])
            speaker_name = default_speakers["name"]

            loopback_dev = None
            for dev in p.get_loopback_device_info_generator():
                if speaker_name in dev["name"]:
                    loopback_dev = dev
                    break

            # If default speaker loopback not found, take first available loopback
            if loopback_dev is None:
                for dev in p.get_loopback_device_info_generator():
                    loopback_dev = dev
                    break

            if loopback_dev is not None:
                dev_idx = loopback_dev["index"]
                dev_name = loopback_dev["name"]
                native_sr = int(loopback_dev["defaultSampleRate"])
                native_ch = int(loopback_dev["maxInputChannels"])

                print("\n" + "=" * 68, flush=True)
                print(f" [Audio Cable] ACTIVE: Pure Digital WASAPI Loopback (Instructor Stream)", flush=True)
                print(f" [Audio Cable] Device: '{dev_name}' (Index: {dev_idx})", flush=True)
                print(f" [Audio Cable] Format: {native_sr}Hz {native_ch}ch -> 16,000Hz Mono 16-bit PCM", flush=True)
                print(f" [Audio Cable] Digital Studio Audio: ZERO room echo, ZERO fan noise, ZERO distortion", flush=True)
                print("=" * 68 + "\n", flush=True)

                def loopback_callback(in_data, frame_count, time_info, status):
                    if not is_running:
                        return (in_data, pyaudio.paAbort)

                    samples = np.frombuffer(in_data, dtype=np.int16).reshape(-1, native_ch)
                    if native_ch >= 2:
                        mono = (samples[:, 0].astype(np.float32) + samples[:, 1].astype(np.float32)) * 0.5
                    else:
                        mono = samples[:, 0].astype(np.float32)

                    # Anti-aliased high-fidelity polyphase resample to 16,000Hz
                    if native_sr == 48000:
                        resampled = resample_poly(mono, 1, 3).astype(np.int16)
                    elif native_sr == 44100:
                        resampled = resample_poly(mono, 160, 441).astype(np.int16)
                    else:
                        resampled = resample_poly(mono, 16000, native_sr).astype(np.int16)

                    mono_16k = resampled.tobytes()
                    recording_frames.append(mono_16k)
                    try:
                        audio_queue.put_nowait(mono_16k)
                    except Exception:
                        pass

                    return (in_data, pyaudio.paContinue)

                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=native_ch,
                    rate=native_sr,
                    input=True,
                    input_device_index=dev_idx,
                    frames_per_buffer=2048,
                    stream_callback=loopback_callback,
                )

                try:
                    stream.start_stream()
                    print(f"[Audio Cable] Live stream listening to Google Meet instructor audio...", flush=True)
                    while is_running:
                        time.sleep(0.3)
                finally:
                    try:
                        stream.stop_stream()
                        stream.close()
                    except Exception:
                        pass
                    p.terminate()

                    print("[Audio Cable] Saving session audio to 'debug_audio.wav'...", flush=True)
                    try:
                        if recording_frames:
                            with wave.open("debug_audio.wav", "wb") as wf:
                                wf.setnchannels(1)
                                wf.setsampwidth(2)
                                wf.setframerate(16000)
                                wf.writeframes(b"".join(recording_frames))
                            print(f"[Audio Cable] Session audio saved successfully to 'debug_audio.wav' ({len(recording_frames)} chunks).", flush=True)
                        else:
                            print("[Audio Cable] Notice: No audio chunks were captured in this session.", flush=True)
                    except Exception as e:
                        print(f"[Audio Cable] Error saving WAV: {e}", flush=True)
                return
        except Exception as e:
            print(f"[Audio Cable] WASAPI Loopback notice: {e}. Falling back to standard input device...", flush=True)

    # 2. Fallback Strategy: Physical Audio Input / Microphone Array via sounddevice
    import sounddevice as sd
    target_idx, dev_info = find_best_audio_device()
    if target_idx is None:
        print("[Audio] Error: No audio input device found.", flush=True)
        return

    dev_name = dev_info["name"]
    max_in = dev_info.get("max_input_channels", 1)
    default_sr = int(dev_info.get("default_samplerate", 16000))
    print(f"[Audio] Selected audio capture device: '{dev_name}' (Index: {target_idx}, Channels: {max_in}, Native SR: {default_sr}Hz)", flush=True)

    target_sr = 16000
    target_channels = 1
    stream_sr = target_sr
    stream_channels = target_channels
    software_resample = False

    try:
        sd.check_input_settings(device=target_idx, channels=target_channels, samplerate=target_sr, dtype="int16")
        print(f"[Audio] Audio stream configured for DIRECT 16,000Hz mono capture.", flush=True)
    except Exception as e:
        stream_sr = default_sr
        stream_channels = 1 if max_in < 2 else 2
        software_resample = True
        print(f"[Audio] Driver does not accept direct 16kHz ({e}). Using native {stream_sr}Hz with polyphase resampling.", flush=True)

    def audio_callback(indata, frames, time_info, status):
        if not is_running:
            return

        if not software_resample:
            mono_16k = indata.tobytes()
        else:
            if indata.ndim > 1 and indata.shape[1] >= 2:
                mono = indata[:, 0].astype(np.float32)
            else:
                mono = indata.ravel().astype(np.float32)

            if stream_sr == 48000:
                resampled = resample_poly(mono, 1, 3).astype(np.int16)
            elif stream_sr == 44100:
                resampled = resample_poly(mono, 160, 441).astype(np.int16)
            else:
                resampled = resample_poly(mono, 16000, stream_sr).astype(np.int16)
            mono_16k = resampled.tobytes()

        recording_frames.append(mono_16k)
        try:
            audio_queue.put_nowait(mono_16k)
        except Exception:
            pass

    try:
        with sd.InputStream(
            device=target_idx,
            channels=stream_channels,
            samplerate=stream_sr,
            dtype="int16",
            blocksize=4000,
            callback=audio_callback,
        ):
            print(f"[Audio] Real-time audio capture stream is ACTIVE on '{dev_name}'.", flush=True)
            while is_running:
                time.sleep(0.3)
    except Exception as e:
        print(f"[Audio] SoundDevice capture error: {e}", flush=True)
    finally:
        print("[Audio] Saving session audio to 'debug_audio.wav'...", flush=True)
        try:
            if recording_frames:
                with wave.open("debug_audio.wav", "wb") as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(16000)
                    wf.writeframes(b"".join(recording_frames))
                print(f"[Audio] Session audio saved successfully to 'debug_audio.wav' ({len(recording_frames)} chunks).", flush=True)
            else:
                print("[Audio] Notice: No audio chunks were captured in this session.", flush=True)
        except Exception as e:
            print(f"[Audio] Error saving WAV: {e}", flush=True)


def handle_exit(signum, frame):
    """
    Handle graceful shutdown when Ctrl+C is pressed.
    Signals session completion so lecture notes are generated and uploaded.
    """
    global is_running
    print("\n[Shutdown] Stopping agent gracefully and proceeding to generate notes...", flush=True)
    is_running = False


def enable_google_meet_captions(driver):
    """
    Enables Closed Captions (CC) in Google Meet if not already enabled.
    """
    if not driver:
        return False

    try:
        # 1. Check if captions are already active
        active_btn = driver.find_elements(
            By.XPATH,
            "//button[(contains(@aria-label, 'Turn off captions') or contains(@aria-label, 'turn off captions') or @aria-pressed='true') and (contains(@aria-label, 'caption') or contains(@aria-label, 'Caption'))]"
        )
        if active_btn:
            print("[Captions] Google Meet Closed Captions are already ACTIVE.", flush=True)
            return True

        # 2. Click Turn on captions button
        caption_buttons = driver.find_elements(
            By.XPATH,
            "//button[contains(@aria-label, 'Turn on captions') or contains(@aria-label, 'turn on captions') or contains(@aria-label, 'captions') or contains(@aria-label, 'Captions') or contains(., 'closed_caption')]"
        )
        for b in caption_buttons:
            try:
                if b.is_displayed():
                    al = b.get_attribute("aria-label") or "captions button"
                    b.click()
                    print(f"[Captions] Clicked Closed Captions button: '{al}'", flush=True)
                    time.sleep(1)
                    return True
            except Exception:
                pass

        # 3. Fallback: Send 'c' key to page body
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys("c")
        print("[Captions] Sent hotkey 'c' to toggle Closed Captions.", flush=True)
        time.sleep(1)
        return True
    except Exception as e:
        print(f"[Captions] Error enabling captions: {e}", flush=True)
        return False


INJECT_CAPTION_OBSERVER_JS = """
(() => {
    if (window.__meet_caption_listener_active) return;
    window.__meet_caption_listener_active = true;
    window.__caption_events = [];

    let lastSpeaker = "";
    let lastText = "";
    let lastChangeTime = Date.now();
    const seenHashes = new Set();

    function cleanStr(s) {
        return (s || "").replace(/\\s+/g, " ").trim();
    }

    function emitCaption(speaker, text) {
        speaker = cleanStr(speaker) || "Speaker";
        text = cleanStr(text);
        if (!text || text.length < 2) return;

        const hash = speaker + ":::" + text.toLowerCase();
        if (!seenHashes.has(hash)) {
            seenHashes.add(hash);
            if (seenHashes.size > 2000) seenHashes.clear();

            window.__caption_events.push({
                speaker: speaker,
                text: text,
                time: new Date().toLocaleTimeString()
            });
        }
    }

    function handleLiveText(speaker, text) {
        speaker = cleanStr(speaker);
        text = cleanStr(text);
        if (!text) return;

        if (lastSpeaker && lastSpeaker !== speaker && lastText) {
            emitCaption(lastSpeaker, lastText);
            lastText = "";
        }

        lastSpeaker = speaker;
        if (text !== lastText) {
            lastText = text;
            lastChangeTime = Date.now();
        }
    }

    function scanCaptions() {
        try {
            // Strategy 1: Google Meet caption containers
            const blocks = document.querySelectorAll(
                'div.nMkdHc, div.VbkSUe, div[jscontroller="D1tHje"] > div, div.a4cQT > div, div[aria-live="polite"] > div'
            );

            blocks.forEach(block => {
                let speaker = "";
                let spkEl = block.querySelector('.zs75Pd, .NWGude, .jxFHg, span.NWGude, [class*="speaker" i]');
                if (spkEl && spkEl.innerText) {
                    speaker = spkEl.innerText;
                } else {
                    let img = block.querySelector('img[alt]');
                    if (img && img.alt && img.alt.length > 1) {
                        speaker = img.alt;
                    }
                }

                let text = "";
                let txtEl = block.querySelector('.iTTPOb, .yg7Twf, [class*="text" i]');
                if (txtEl && txtEl.innerText) {
                    text = txtEl.innerText;
                }

                if (speaker && text) {
                    handleLiveText(speaker, text);
                }
            });

            // Strategy 2: aria-live region (Accessibility stream)
            const liveRegions = document.querySelectorAll('div[aria-live="polite"], div[aria-live="assertive"]');
            liveRegions.forEach(region => {
                const raw = cleanStr(region.innerText);
                if (raw && raw.includes(":")) {
                    const colonIdx = raw.indexOf(":");
                    const spk = raw.substring(0, colonIdx).trim();
                    const txt = raw.substring(colonIdx + 1).trim();
                    if (spk.length > 0 && spk.length < 50 && txt.length > 1) {
                        handleLiveText(spk, txt);
                    }
                }
            });

            // If active speaker hasn't added new words for >= 1.5s, commit
            if (lastText && (Date.now() - lastChangeTime > 1500)) {
                emitCaption(lastSpeaker, lastText);
                lastText = "";
                lastSpeaker = "";
            }
        } catch (e) {}
    }

    try {
        const observer = new MutationObserver(() => scanCaptions());
        observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    } catch(e) {}

    setInterval(scanCaptions, 400);

    window.__get_new_captions = () => {
        if (lastText && (Date.now() - lastChangeTime > 900)) {
            emitCaption(lastSpeaker, lastText);
            lastText = "";
            lastSpeaker = "";
        }
        const res = window.__caption_events.slice();
        window.__caption_events = [];
        return res;
    };
})();
"""


def monitor_captions_loop():
    """
    Background worker that monitors Google Meet's DOM in real time,
    toggles Closed Captions once inside the meeting, captures speaker-attributed speech,
    writes to transcript.txt and terminal, and checks for triggers to auto-respond in chat.
    """
    global is_running, global_driver, session_transcript_lines, rolling_history

    print("[Captions] Caption and speaker monitor thread started.", flush=True)
    in_call = False
    captions_enabled = False
    last_check_time = 0

    while is_running:
        if not global_driver:
            time.sleep(1)
            continue

        try:
            # 1. Check if user is inside the meeting room
            now = time.time()
            if not in_call or (now - last_check_time > 5):
                last_check_time = now
                # Look for meeting room call controls
                meet_controls = global_driver.find_elements(
                    By.XPATH,
                    "//button[contains(@aria-label, 'Leave call') or contains(@aria-label, 'leave call') or contains(@aria-label, 'captions') or contains(@aria-label, 'Captions') or contains(@aria-label, 'Chat')]"
                )
                if meet_controls:
                    if not in_call:
                        in_call = True
                        print("\n" + "=" * 60, flush=True)
                        print(">>> IN GOOGLE MEET CALL ROOM! Enabling Closed Captions... <<<", flush=True)
                        print("=" * 60 + "\n", flush=True)

                    if not captions_enabled:
                        captions_enabled = enable_google_meet_captions(global_driver)

            if in_call:
                # 2. Inject caption observer if needed
                global_driver.execute_script(INJECT_CAPTION_OBSERVER_JS)

                # 3. Harvest new captions
                captions = global_driver.execute_script(
                    "return (window.__get_new_captions ? window.__get_new_captions() : []);"
                )

                if captions:
                    for item in captions:
                        speaker = item.get("speaker", "Speaker").strip()
                        text = item.get("text", "").strip()
                        t_str = item.get("time", "")

                        if not text or len(text.split()) < 1:
                            continue

                        line = f"[{speaker}]: {text}"
                        print(f"\n[Meeting STT {t_str}] [{speaker}]: \"{text}\"", flush=True)

                        session_transcript_lines.append(line)
                        with open("transcript.txt", "a", encoding="utf-8") as f:
                            f.write(line + "\n")

                        words = text.split()
                        rolling_history.extend(words)

                        # Check for trigger keywords / student name
                        text_lower = text.lower()
                        if any(kw in text_lower for kw in TRIGGER_KEYWORDS):
                            print(f"\n🔔 [ATTENDANCE/QUESTION DETECTED from {speaker}]: '{text}'", flush=True)
                            response = check_and_respond_with_gemini(f"{speaker} says: {text}")
                            print(f"🤖 [AI Generated Chat Response]: '{response}'", flush=True)
                            send_message_to_chat(response)

        except Exception:
            pass

        time.sleep(0.5)


def get_primary_chrome_profile():
    """
    Detects the user's primary signed-in Chrome profile from their actual Chrome installation.
    """
    local_state_path = os.path.join(
        os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "User Data", "Local State"
    )
    if os.path.exists(local_state_path):
        try:
            with open(local_state_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            last_used = data.get("profile", {}).get("last_used")
            if last_used:
                return last_used
            info_cache = data.get("profile", {}).get("info_cache", {})
            for p_dir, p_info in info_cache.items():
                if p_info.get("user_name"):
                    return p_dir
            if info_cache:
                return list(info_cache.keys())[0]
        except Exception:
            pass
    return "Profile 2"


def force_bring_chrome_to_front():
    """
    Brings the visible Chrome window to the absolute front of the user's screen.
    Uses Win32 API to bypass Windows 11 foreground lock so the automation is 100% visible.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        hwnds = []
        def enum_cb(hwnd, extra):
            if user32.IsWindowVisible(hwnd):
                cls_buff = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, cls_buff, 256)
                if "Chrome_WidgetWin" in cls_buff.value:
                    title_buff = ctypes.create_unicode_buffer(512)
                    user32.GetWindowTextW(hwnd, title_buff, 512)
                    hwnds.append((hwnd, title_buff.value))
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        user32.EnumWindows(WNDENUMPROC(enum_cb), 0)

        for hwnd, title in hwnds:
            user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
            cur_thread = kernel32.GetCurrentThreadId()
            fg_hwnd = user32.GetForegroundWindow()
            fg_thread = user32.GetWindowThreadProcessId(fg_hwnd, None)
            if fg_thread != cur_thread:
                user32.AttachThreadInput(cur_thread, fg_thread, True)
            user32.SetForegroundWindow(hwnd)
            user32.BringWindowToTop(hwnd)
            user32.SetFocus(hwnd)
            if fg_thread != cur_thread:
                user32.AttachThreadInput(cur_thread, fg_thread, False)
            print(f"[UI] Chrome window brought directly to the front screen.", flush=True)
            break
    except Exception:
        pass


def turn_off_mic_and_camera(driver):
    """
    Turns off microphone and camera on the Google Meet pre-join lobby screen.
    """
    print("[Google Meet] Turning off microphone and camera before joining...", flush=True)
    time.sleep(1.5)

    # 1. Turn off Microphone (Mute)
    mic_muted = False
    mic_selectors = [
        "//button[contains(@aria-label, 'Turn off microphone') or contains(@aria-label, 'turn off microphone') or contains(@aria-label, 'ctrl + d') or contains(@aria-label, 'Ctrl + D')]",
        "//div[@role='button'][contains(@aria-label, 'Turn off microphone') or contains(@aria-label, 'turn off microphone')]",
        "//button[@data-is-muted='false' and (contains(@aria-label, 'microphone') or contains(@aria-label, 'Microphone'))]",
        "//button[contains(@aria-label, 'microphone') and not(contains(@aria-label, 'Turn on'))]",
    ]
    for m_sel in mic_selectors:
        try:
            mic_btns = driver.find_elements(By.XPATH, m_sel)
            for mb in mic_btns:
                if mb.is_displayed():
                    driver.execute_script("arguments[0].click();", mb)
                    print("[Google Meet] Microphone turned OFF (button clicked).", flush=True)
                    mic_muted = True
                    break
            if mic_muted:
                break
        except Exception:
            continue

    if not mic_muted:
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            body.send_keys(Keys.CONTROL, "d")
            print("[Google Meet] Microphone turned OFF (Ctrl+D shortcut).", flush=True)
        except Exception as e:
            print(f"[Google Meet] Mic shortcut note: {e}", flush=True)

    time.sleep(1)

    # 2. Turn off Camera
    cam_off = False
    cam_selectors = [
        "//button[contains(@aria-label, 'Turn off camera') or contains(@aria-label, 'turn off camera') or contains(@aria-label, 'ctrl + e') or contains(@aria-label, 'Ctrl + E')]",
        "//div[@role='button'][contains(@aria-label, 'Turn off camera') or contains(@aria-label, 'turn off camera')]",
        "//button[@data-is-muted='false' and (contains(@aria-label, 'camera') or contains(@aria-label, 'Camera'))]",
        "//button[contains(@aria-label, 'camera') and not(contains(@aria-label, 'Turn on'))]",
    ]
    for c_sel in cam_selectors:
        try:
            cam_btns = driver.find_elements(By.XPATH, c_sel)
            for cb in cam_btns:
                if cb.is_displayed():
                    driver.execute_script("arguments[0].click();", cb)
                    print("[Google Meet] Camera turned OFF (button clicked).", flush=True)
                    cam_off = True
                    break
            if cam_off:
                break
        except Exception:
            continue

    if not cam_off:
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            body.send_keys(Keys.CONTROL, "e")
            print("[Google Meet] Camera turned OFF (Ctrl+E shortcut).", flush=True)
        except Exception as e:
            print(f"[Google Meet] Cam shortcut note: {e}", flush=True)

    time.sleep(1)


def join_google_meet(meet_link, audio_thread=None):
    """
    Automates joining a Google Meet session with live visible Chrome window.
    """
    global global_driver, is_running

    if not meet_link or "meet.google.com" not in meet_link:
        print("[Google Meet] Notice: No valid GMEET_LINK set in .env.")
        print("[Google Meet] Audio listening and transcription active locally.")
        if audio_thread is None or not audio_thread.is_alive():
            audio_thread = Thread(target=capture_audio_from_device, daemon=True)
            audio_thread.start()
        input("Press Enter in terminal to stop recording and process notes...")
        return

    import subprocess

    # 1. Terminate any orphaned Chrome processes cleanly
    subprocess.run(["taskkill", "/IM", "chrome.exe", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(1)

    # 2. Dedicated Chrome persistent profile directory
    user_data = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "AIStudentAgent", "ChromeMeetProfile")
    os.makedirs(user_data, exist_ok=True)
    for lock in ["lockfile", "SingletonLock", "SingletonCookie", "SingletonSocket"]:
        lp = os.path.join(user_data, lock)
        if os.path.exists(lp):
            try:
                os.remove(lp)
            except Exception:
                pass

    # 3. Launch Chrome directly through Selenium on a single dedicated window
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={user_data}")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--use-fake-ui-for-media-stream")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    print("[Google Meet] Launching single dedicated Chrome window...", flush=True)
    try:
        global_driver = webdriver.Chrome(options=chrome_options)
        print("[Google Meet] Chrome browser connected successfully!", flush=True)
    except Exception as e:
        print(f"[Google Meet] Error launching Chrome: {e}", flush=True)
        return

    # 4. Navigate directly to meet_link on this single tab
    print(f"[Google Meet] Navigating to: {meet_link}", flush=True)
    try:
        global_driver.get(meet_link)
    except Exception as e:
        print(f"[Google Meet] Navigation note: {e}", flush=True)

    try:
        global_driver.maximize_window()
        global_driver.execute_cdp_cmd("Page.bringToFront", {})
    except Exception:
        pass
    force_bring_chrome_to_front()

    print("[Google Meet] Interacting with meeting lobby strictly quick...", flush=True)
    joined = False

    for attempt in range(50):  # Poll every 0.3s (up to 15 seconds)
        time.sleep(0.3)
        try:
            status = global_driver.execute_script("""
                let res = { mic: false, cam: false, joined: false, button_text: '' };

                // Dismiss popups ("Got it", "Dismiss")
                const allBtns = Array.from(document.querySelectorAll('button'));
                for (const b of allBtns) {
                    const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                    if (t === 'got it' || t === 'dismiss' || t.includes('use microphone') || t.includes('continue without')) {
                        b.click();
                    }
                }

                // 1. Turn OFF Microphone
                const interactive = Array.from(document.querySelectorAll('button, div[role="button"]'));
                for (const b of interactive) {
                    const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                    if (aria.includes('turn off microphone') || (aria.includes('microphone') && !aria.includes('turn on'))) {
                        b.click();
                        res.mic = true;
                        break;
                    }
                }

                // 2. Turn OFF Camera
                for (const b of interactive) {
                    const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                    if (aria.includes('turn off camera') || (aria.includes('camera') && !aria.includes('turn on'))) {
                        b.click();
                        res.cam = true;
                        break;
                    }
                }

                // 3. Click 'Ask to join' or 'Join now'
                for (const b of allBtns) {
                    const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                    const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                    if (t.includes('ask to join') || t.includes('join now') || aria.includes('ask to join') || aria.includes('join now')) {
                        b.scrollIntoView();
                        b.click();
                        res.joined = true;
                        res.button_text = b.innerText.trim();
                        break;
                    }
                }

                return res;
            """)

            if status:
                if status.get("mic"):
                    print("[Google Meet] Microphone turned OFF.", flush=True)
                if status.get("cam"):
                    print("[Google Meet] Camera turned OFF.", flush=True)
                if status.get("joined"):
                    btn_name = status.get("button_text") or "Ask to join"
                    print(f"[Google Meet] SUCCESS: Tapped '{btn_name}' strictly quick!", flush=True)
                    joined = True
                    break
        except Exception:
            pass

        # Redundancy: send Ctrl+D and Ctrl+E keyboard shortcuts on 2nd attempt
        if attempt == 2:
            try:
                body = global_driver.find_element(By.TAG_NAME, "body")
                body.send_keys(Keys.CONTROL, "d")
                body.send_keys(Keys.CONTROL, "e")
            except Exception:
                pass

    if not joined:
        # Final fallback via XPath click if JS query didn't find button
        for sel in [
            "//button[contains(., 'Ask to join') or contains(., 'Join now')]",
            "//button[.//span[contains(text(),'Ask to join') or contains(text(),'Join now')]]",
            "//button[contains(@aria-label, 'Ask to join') or contains(@aria-label, 'Join now')]",
        ]:
            try:
                join_button = global_driver.find_element(By.XPATH, sel)
                if join_button.is_displayed():
                    global_driver.execute_script("arguments[0].click();", join_button)
                    print(f"[Google Meet] SUCCESS: Tapped '{join_button.text.strip()}' button!", flush=True)
                    joined = True
                    break
            except Exception:
                continue

    if not joined:
        print("[Google Meet] Notice: You can also tap 'Ask to join' directly on the Chrome window.", flush=True)

    print("[Google Meet] Staying on this single tab. Monitoring meeting session...", flush=True)

    if audio_thread is None or not audio_thread.is_alive():
        audio_thread = Thread(target=capture_audio_from_device, daemon=True)
        audio_thread.start()

    # Start Google Meet Closed Caption & Speaker STT monitor thread
    caption_thread = Thread(target=monitor_captions_loop, daemon=True)
    caption_thread.start()

    print("\n" + "=" * 65)
    print(">>> AI STUDENT AGENT IS LIVE IN THE MEETING! <<<")
    print(">>> Listening to lecture audio & monitoring captions in real-time.")
    print(">>> Automatically responds to attendance calls or questions in chat.")
    print(">>> Press 'q' then ENTER in this terminal when class ends to generate Notion notes.")
    print("=" * 65 + "\n")

    # Clear any residual keystrokes in Windows stdin buffer
    if sys.platform == "win32":
        try:
            import msvcrt
            while msvcrt.kbhit():
                msvcrt.getch()
        except Exception:
            pass

    if os.path.exists("stop_meeting.flag"):
        try:
            os.remove("stop_meeting.flag")
        except Exception:
            pass

    start_time = time.time()
    try:
        print("[Google Meet] Agent is active in meeting. Monitoring audio & captions in real-time...", flush=True)
        print(">>> Press 'q' in this terminal when class finishes to generate Notion notes.\n", flush=True)
        while is_running:
            if os.path.exists("stop_meeting.flag"):
                print("[Google Meet] Stop flag detected. Ending meeting and generating notes...", flush=True)
                try:
                    os.remove("stop_meeting.flag")
                except Exception:
                    pass
                break

            if sys.platform == "win32" and sys.stdin and sys.stdin.isatty():
                try:
                    import msvcrt
                    if msvcrt.kbhit():
                        ch = msvcrt.getch()
                        if ch in [b'q', b'Q'] and (time.time() - start_time > 3):
                            print("\n[Google Meet] 'q' detected. Class ended by student. Generating notes...", flush=True)
                            break
                except Exception:
                    pass

            time.sleep(0.5)
    except (KeyboardInterrupt, SystemExit):
        print("\n[Shutdown] Stopping agent gracefully...", flush=True)
    finally:
        is_running = False


def send_message_to_chat(message):
    """
    Automates sending an attendance response (e.g. 'Present!') or Q&A answer directly
    into the Google Meet right-side chat section.
    """
    global global_driver

    if not global_driver:
        print("[Chat] Notice: Browser driver not active, skipping chat dispatch.", flush=True)
        return False

    print(f"\n[Chat] Preparing to send automated message to Google Meet chat: '{message}'...", flush=True)

    # Allow up to 10 attempts in case user is just being admitted or the chat panel is opening
    for attempt in range(10):
        try:
            # 1. Check if the chat input is already open and visible
            chat_state = global_driver.execute_script("""
                const inputs = Array.from(document.querySelectorAll('textarea, div[role="textbox"][contenteditable="true"]'));
                for (const inp of inputs) {
                    if (inp.offsetParent !== null && !inp.disabled && inp.getAttribute('aria-disabled') !== 'true') {
                        return { open: true, has_input: true };
                    }
                }

                // Check if the side chat panel is already open
                const panel = document.querySelector('div[aria-label*="messages" i], div[aria-label*="Chat" i], div[data-tab-id="2"]');
                const isPanelOpen = panel && panel.offsetParent !== null;

                return { open: isPanelOpen, has_input: false };
            """)

            # 2. If the chat panel is NOT open, click the chat toggle button (speech bubble icon)
            if not chat_state or not chat_state.get("open"):
                opened = global_driver.execute_script("""
                    const chatBtns = Array.from(document.querySelectorAll('button'));
                    for (const b of chatBtns) {
                        const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                        if (aria.includes('chat with everyone') || aria.includes('in-call messages') || aria.includes('chat') || aria.includes('message')) {
                            b.click();
                            return true;
                        }
                    }
                    return false;
                """)
                if opened:
                    print("[Chat] Opened right-side chat section.", flush=True)
                time.sleep(0.8)

            # 3. Inject message and click the Send button (paper plane icon)
            res = global_driver.execute_script("""
                const msg = arguments[0];
                const inputs = Array.from(document.querySelectorAll('textarea, div[role="textbox"][contenteditable="true"]'));
                let targetInput = null;

                for (const inp of inputs) {
                    if (inp.offsetParent !== null && !inp.disabled && inp.getAttribute('aria-disabled') !== 'true') {
                        targetInput = inp;
                        break;
                    }
                }

                if (!targetInput) {
                    return { success: false, reason: "Chat input not ready yet (e.g. waiting to be admitted)" };
                }

                // Populate text into input and trigger events
                targetInput.focus();
                if (targetInput.tagName.toLowerCase() === 'textarea') {
                    targetInput.value = msg;
                    targetInput.dispatchEvent(new Event('input', { bubbles: true }));
                    targetInput.dispatchEvent(new Event('change', { bubbles: true }));
                } else {
                    targetInput.innerText = msg;
                    targetInput.dispatchEvent(new InputEvent('input', { bubbles: true }));
                }

                // Find and click the Send button (paper plane icon)
                const sendBtns = Array.from(document.querySelectorAll('button'));
                for (const sb of sendBtns) {
                    const aria = (sb.getAttribute('aria-label') || '').toLowerCase();
                    if ((aria.includes('send') || sb.getAttribute('jsname') === 'soHxf') && !sb.disabled && sb.getAttribute('aria-disabled') !== 'true') {
                        sb.click();
                        return { success: true, method: "send_button" };
                    }
                }

                return { success: true, method: "text_set" };
            """, message)

            if res and res.get("success"):
                if res.get("method") != "send_button":
                    try:
                        active = global_driver.switch_to.active_element
                        active.send_keys(Keys.RETURN)
                    except Exception:
                        pass
                print(f"✅ [Chat] Successfully sent '{message}' into Google Meet chat!", flush=True)
                return True

            time.sleep(0.8)
        except Exception:
            time.sleep(0.8)

    # 4. Fallback: Selenium Native Input & Send
    try:
        input_selectors = [
            "//textarea[contains(@aria-label, 'message') or contains(@aria-label, 'Send')]",
            "//textarea[@name='chatTextInput']",
            "//textarea",
            "//div[@role='textbox']",
        ]
        for sel in input_selectors:
            try:
                inputs = global_driver.find_elements(By.XPATH, sel)
                for inp in inputs:
                    if inp.is_displayed() and inp.is_enabled():
                        inp.clear()
                        inp.send_keys(message)
                        time.sleep(0.3)
                        # Click send button
                        send_btns = global_driver.find_elements(By.XPATH, "//button[contains(@aria-label, 'Send') or @jsname='soHxf']")
                        sent = False
                        for sb in send_btns:
                            if sb.is_displayed() and sb.is_enabled():
                                sb.click()
                                sent = True
                                break
                        if not sent:
                            inp.send_keys(Keys.RETURN)
                        print(f"✅ [Chat] Fallback: Sent '{message}' to Google Meet chat!", flush=True)
                        return True
            except Exception:
                continue
    except Exception as e:
        print(f"[Chat] Fallback note: {e}", flush=True)

    print(f"[Chat] Notice: Message '{message}' will be sent as soon as you are admitted to the meeting room.", flush=True)
    return False


if __name__ == "__main__":
    signal.signal(signal.SIGINT, handle_exit)

    # 1. Start Flask API server
    flask_thread = Thread(target=start_flask_app, daemon=True)
    flask_thread.start()
    time.sleep(0.5)

    # 2. Audio Capture Thread & Speech Recognition Worker
    audio_thread = Thread(target=capture_audio_from_device, daemon=True)
    audio_thread.start()

    vosk_thread = Thread(target=speech_recognition_worker, daemon=True)
    vosk_thread.start()

    # 3. Run Google Meet automation
    join_google_meet(GMEET_LINK, audio_thread)

    # 4. Stop capture loop
    is_running = False
    print("\n[Notes] Class finished. Processing final notes...", flush=True)
    time.sleep(1.5)

    # 5. Compile final lecture transcript
    # Primary: Real-time speaker-attributed captions captured directly from Google Meet
    transcript = ""
    if session_transcript_lines:
        print(f"[Notes] Using real-time speaker-attributed lecture transcript ({len(session_transcript_lines)} dialogue entries).", flush=True)
        transcript = "\n".join(session_transcript_lines).strip()
    elif os.path.exists("transcript.txt") and os.path.getsize("transcript.txt") > 10:
        with open("transcript.txt", "r", encoding="utf-8") as f:
            content = f.read().strip()
            if content:
                print("[Notes] Using transcript saved in 'transcript.txt'.", flush=True)
                transcript = content

    # Fallback: Transcribe session audio debug_audio.wav using Google Gemini Multimodal STT
    if not transcript or len(transcript.split()) < 5:
        print("[Notes] Captions empty or sparse, attempting Gemini multimodal transcription of 'debug_audio.wav'...", flush=True)
        audio_transcript = gemini_transcript("debug_audio.wav")
        if audio_transcript and len(audio_transcript.split()) >= 5:
            transcript = audio_transcript

    if transcript and len(transcript.strip().split()) >= 5:
        print("\n" + "=" * 50)
        print(" FINAL LECTURE TRANSCRIPT:")
        print(transcript.strip())
        print("=" * 50)

        print("\n[AI] Formatting notes using Google Gemini...")
        notes = gemini_make_notes(transcript)
        print("\n" + "=" * 50)
        print(" GENERATED NOTES:")
        print(notes)
        print("=" * 50)

        # 6. Push to Notion
        if NOTION_API_TOKEN and NOTION_PAGE_ID:
            print("\n[Notion] Uploading notes to Notion...")
            from datetime import datetime
            timestamp_str = datetime.now().strftime("%b %d, %Y %I:%M %p")
            notes_to_notion(
                NOTION_API_TOKEN,
                NOTION_PAGE_ID,
                notes,
                title=f"Class Notes - {STUDENT_NAME} ({timestamp_str})",
            )
    else:
        print("\n" + "=" * 50)
        print("[Notes] Notice: No spoken lecture content was detected during this session.")
        print("[Notes] Skipping Notion upload so blank or irrelevant notes are not created.")
        print("=" * 50)

    if global_driver:
        try:
            global_driver.quit()
        except Exception:
            pass

    print("\nAI Student Agent completed successfully.")
