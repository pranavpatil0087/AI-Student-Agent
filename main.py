import os
import sys
import time
import re
import json
import wave
import signal
import queue
import logging
from collections import deque
from datetime import datetime
from threading import Thread, Lock

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

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

from flask import Flask, request, jsonify, render_template, Response
from flask_cors import CORS
from vosk import Model, KaldiRecognizer

from notion_integration import format_notes_for_notion, notes_to_notion
from gemini_functions import check_and_respond_with_gemini, gemini_make_notes, gemini_transcript, gemini_generate_quiz, gemini_evaluate_quiz

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

# Flask Configuration
app = Flask(__name__, template_folder="templates", static_folder="static")
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
is_running = False
last_attendance_trigger_time = 0
audio_queue = queue.Queue(maxsize=300)
vosk_lock = Lock()
session_worker_thread = None

# Reset transcript.txt so previous runs never contaminate the session
with open("transcript.txt", "w", encoding="utf-8") as f:
    f.write("")

# Dynamic Session State
session_state = {
    "status": "idle",  # "idle", "launching", "connecting", "in_meeting", "generating_notes", "completed", "error"
    "message": "Ready to join Google Meet session.",
    "meet_link": "",
    "student_name": STUDENT_NAME,
    "trigger_keywords": TRIGGER_KEYWORDS,
    "start_time": None,
    "end_time": None,
    "captions": [],
    "activity_logs": [],
    "notes": "",
    "transcript": "",
    "notion_uploaded": False,
    "notion_title": "",
    "notion_error": None,
    "error": None,
    "quiz": [],
    "quiz_result": None,
}
session_lock = Lock()


def add_caption_to_session(speaker, text, t_str=None):
    """Add dialogue entry to in-memory session buffer and transcript.txt."""
    if not t_str:
        t_str = datetime.now().strftime("%I:%M:%S %p")
    clean_line = f"[{speaker}]: {text}"
    session_transcript_lines.append(clean_line)

    with open("transcript.txt", "a", encoding="utf-8") as f:
        f.write(clean_line + "\n")

    with session_lock:
        session_state["captions"].append({
            "speaker": speaker,
            "text": text,
            "time": t_str
        })
        if len(session_state["captions"]) > 250:
            session_state["captions"] = session_state["captions"][-250:]


def add_activity_log(log_type, message):
    """Add event to agent decision & attendance log."""
    t_str = datetime.now().strftime("%I:%M:%S %p")
    print(f"[{log_type.upper()}] {message}", flush=True)
    with session_lock:
        session_state["activity_logs"].append({
            "type": log_type,
            "message": message,
            "time": t_str
        })
        if len(session_state["activity_logs"]) > 100:
            session_state["activity_logs"] = session_state["activity_logs"][-100:]


def update_session_status(status, message=None, error=None):
    """Safely update top-level session status."""
    with session_lock:
        session_state["status"] = status
        if message:
            session_state["message"] = message
        if error is not None:
            session_state["error"] = error
    if message:
        print(f"[State: {status.upper()}] {message}", flush=True)


def normalize_meet_link(raw):
    """
    Validates and normalizes Google Meet URLs or meeting room codes.
    Accepts:
      - https://meet.google.com/xxx-yyyy-zzz
      - meet.google.com/xxx-yyyy-zzz
      - xxx-yyyy-zzz
    """
    raw = (raw or "").strip()
    if not raw:
        return None, "Please enter a Google Meet link or room code."

    # Pattern for room code xxx-yyyy-zzz
    code_match = re.match(r"^[a-zA-Z]{3}-[a-zA-Z]{4}-[a-zA-Z]{3}$", raw)
    if code_match:
        return f"https://meet.google.com/{raw.lower()}", None

    if raw.startswith("meet.google.com/"):
        raw = "https://" + raw

    meet_pattern = r"https?://meet\.google\.com/[a-zA-Z]{3}-[a-zA-Z]{4}-[a-zA-Z]{3}"
    if re.search(meet_pattern, raw):
        return raw, None

    return None, "Invalid Google Meet link. Format should be: https://meet.google.com/xxx-yyyy-zzz or xxx-yyyy-zzz"


def process_audio_with_vosk(audio_data):
    """
    Process incoming 16kHz mono audio using persistent VOSK recognizer for real-time speech-to-text.
    Protected with vosk_lock to prevent multi-threaded C++ memory corruption.
    """
    global rolling_history, global_recognizer, last_attendance_trigger_time

    with vosk_lock:
        try:
            if global_recognizer.AcceptWaveform(audio_data):
                result = json.loads(global_recognizer.Result())
                text_cleaned = result.get("text", "").strip()

                if text_cleaned:
                    add_caption_to_session("Audio Feed (Vosk)", text_cleaned)
                    words = text_cleaned.split()
                    rolling_history.extend(words)

                    # Check for trigger keywords (student name, attendance call, roll number)
                    text_lower = text_cleaned.lower()
                    now = time.time()
                    if any(kw in text_lower for kw in TRIGGER_KEYWORDS):
                        if now - last_attendance_trigger_time > 8:
                            last_attendance_trigger_time = now
                            add_activity_log("trigger", f"Attendance trigger detected in audio: '{text_cleaned}'")
                            context_words = list(rolling_history)[-20:]
                            context = " ".join(context_words)
                            response = check_and_respond_with_gemini(context)
                            add_activity_log("response", f"AI generated response: '{response}'")
                            if global_driver:
                                send_message_to_chat(response)

                    return {"status": "Audio processed", "transcript": text_cleaned}
            else:
                # Check partial hypothesis for immediate mid-utterance attendance trigger
                partial_result = json.loads(global_recognizer.PartialResult())
                partial_cleaned = partial_result.get("partial", "").strip()
                if partial_cleaned:
                    partial_lower = partial_cleaned.lower()
                    now = time.time()
                    if any(kw in partial_lower for kw in TRIGGER_KEYWORDS):
                        if now - last_attendance_trigger_time > 8:
                            last_attendance_trigger_time = now
                            add_activity_log("trigger", f"Attendance trigger (mid-speech): '{partial_cleaned}'")
                            response = check_and_respond_with_gemini(f"Teacher called: {partial_cleaned}")
                            add_activity_log("response", f"AI generated response: '{response}'")
                            if global_driver:
                                send_message_to_chat(response)
                            global_recognizer.Reset()

            return {"status": "Processing stream"}
        except Exception as e:
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
            pass


def find_best_audio_device():
    """Find the best audio capture device."""
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
    using Windows WASAPI Loopback. Seamlessly falls back to sounddevice microphone if unavailable.
    """
    global recording_frames, is_running

    prefer_mic = bool(AUDIO_DEVICE_NAME and any(k in AUDIO_DEVICE_NAME.lower() for k in ["mic", "microphone"]))

    # 1. Primary Strategy: Digital WASAPI Speaker Loopback
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

            if loopback_dev is None:
                for dev in p.get_loopback_device_info_generator():
                    loopback_dev = dev
                    break

            if loopback_dev is not None:
                dev_idx = loopback_dev["index"]
                dev_name = loopback_dev["name"]
                native_sr = int(loopback_dev["defaultSampleRate"])
                native_ch = int(loopback_dev["maxInputChannels"])

                print(f"[Audio Cable] ACTIVE: WASAPI Loopback on '{dev_name}' ({native_sr}Hz -> 16kHz mono)", flush=True)

                def loopback_callback(in_data, frame_count, time_info, status):
                    if not is_running:
                        return (in_data, pyaudio.paAbort)

                    samples = np.frombuffer(in_data, dtype=np.int16).reshape(-1, native_ch)
                    if native_ch >= 2:
                        mono = (samples[:, 0].astype(np.float32) + samples[:, 1].astype(np.float32)) * 0.5
                    else:
                        mono = samples[:, 0].astype(np.float32)

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
                    while is_running:
                        time.sleep(0.3)
                finally:
                    try:
                        stream.stop_stream()
                        stream.close()
                    except Exception:
                        pass
                    p.terminate()

                    try:
                        if recording_frames:
                            with wave.open("debug_audio.wav", "wb") as wf:
                                wf.setnchannels(1)
                                wf.setsampwidth(2)
                                wf.setframerate(16000)
                                wf.writeframes(b"".join(recording_frames))
                    except Exception as e:
                        print(f"[Audio Cable] Error saving WAV: {e}", flush=True)
                return
        except Exception as e:
            print(f"[Audio Cable] WASAPI notice: {e}. Falling back to sounddevice...", flush=True)

    # 2. Fallback Strategy: Physical Audio Input / Microphone Array via sounddevice
    import sounddevice as sd
    target_idx, dev_info = find_best_audio_device()
    if target_idx is None:
        print("[Audio] Error: No audio input device found.", flush=True)
        return

    dev_name = dev_info["name"]
    max_in = dev_info.get("max_input_channels", 1)
    default_sr = int(dev_info.get("default_samplerate", 16000))

    target_sr = 16000
    target_channels = 1
    stream_sr = target_sr
    stream_channels = target_channels
    software_resample = False

    try:
        sd.check_input_settings(device=target_idx, channels=target_channels, samplerate=target_sr, dtype="int16")
    except Exception:
        stream_sr = default_sr
        stream_channels = 1 if max_in < 2 else 2
        software_resample = True

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
            while is_running:
                time.sleep(0.3)
    except Exception as e:
        print(f"[Audio] SoundDevice capture error: {e}", flush=True)
    finally:
        try:
            if recording_frames:
                with wave.open("debug_audio.wav", "wb") as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(16000)
                    wf.writeframes(b"".join(recording_frames))
        except Exception:
            pass


def handle_exit(signum, frame):
    """Handle graceful shutdown."""
    global is_running
    print("\n[Shutdown] Stopping agent gracefully...", flush=True)
    is_running = False


def enable_google_meet_captions(driver):
    """Enables Closed Captions (CC) in Google Meet if not already enabled."""
    if not driver:
        return False

    try:
        active_btn = driver.find_elements(
            By.XPATH,
            "//button[(contains(@aria-label, 'Turn off captions') or contains(@aria-label, 'turn off captions') or @aria-pressed='true') and (contains(@aria-label, 'caption') or contains(@aria-label, 'Caption'))]"
        )
        if active_btn:
            print("[Captions] Google Meet Closed Captions are ACTIVE.", flush=True)
            return True

        caption_buttons = driver.find_elements(
            By.XPATH,
            "//button[contains(@aria-label, 'Turn on captions') or contains(@aria-label, 'turn on captions') or contains(@aria-label, 'captions') or contains(@aria-label, 'Captions') or contains(., 'closed_caption')]"
        )
        for b in caption_buttons:
            try:
                if b.is_displayed():
                    b.click()
                    print(f"[Captions] Clicked Closed Captions toggle.", flush=True)
                    time.sleep(1)
                    return True
            except Exception:
                pass

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
    toggles Closed Captions, captures speaker speech, and responds to triggers.
    """
    global is_running, global_driver, session_transcript_lines, rolling_history

    in_call = False
    captions_enabled = False
    last_check_time = 0

    while is_running:
        if not global_driver:
            time.sleep(1)
            continue

        try:
            now = time.time()
            if not in_call or (now - last_check_time > 5):
                last_check_time = now
                meet_controls = global_driver.find_elements(
                    By.XPATH,
                    "//button[contains(@aria-label, 'Leave call') or contains(@aria-label, 'leave call') or contains(@aria-label, 'captions') or contains(@aria-label, 'Captions') or contains(@aria-label, 'Chat')]"
                )
                if meet_controls:
                    if not in_call:
                        in_call = True
                        print("\n>>> ACTIVE IN GOOGLE MEET CALL ROOM! Enabling Closed Captions... <<<", flush=True)
                        update_session_status("in_meeting", "Active in meeting room. Monitoring audio & captions.")

                    if not captions_enabled:
                        captions_enabled = enable_google_meet_captions(global_driver)

            if in_call:
                global_driver.execute_script(INJECT_CAPTION_OBSERVER_JS)
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

                        add_caption_to_session(speaker, text, t_str)
                        words = text.split()
                        rolling_history.extend(words)

                        text_lower = text.lower()
                        if any(kw in text_lower for kw in TRIGGER_KEYWORDS):
                            add_activity_log("trigger", f"Attendance trigger from {speaker}: '{text}'")
                            response = check_and_respond_with_gemini(f"{speaker} says: {text}")
                            add_activity_log("response", f"AI generated response: '{response}'")
                            send_message_to_chat(response)

        except Exception:
            pass

        time.sleep(0.5)


def force_bring_chrome_to_front():
    """Brings the visible Chrome window to the absolute front."""
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
            break
    except Exception:
        pass


def join_google_meet(meet_link, audio_thread=None):
    """
    Automates joining a Google Meet session with live visible Chrome window.
    Proceeds directly with the provided meet_link from the front end.
    """
    global global_driver, is_running

    import subprocess

    update_session_status("launching", "Terminating background Chrome instances...")
    subprocess.run(["taskkill", "/IM", "chrome.exe", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(1)

    user_data = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "AIStudentAgent", "ChromeMeetProfile")
    os.makedirs(user_data, exist_ok=True)
    for lock in ["lockfile", "SingletonLock", "SingletonCookie", "SingletonSocket"]:
        lp = os.path.join(user_data, lock)
        if os.path.exists(lp):
            try:
                os.remove(lp)
            except Exception:
                pass

    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={user_data}")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--use-fake-ui-for-media-stream")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    update_session_status("launching", "Opening dedicated Chrome browser window...")
    try:
        global_driver = webdriver.Chrome(options=chrome_options)
        add_activity_log("info", "Chrome browser connected successfully.")
    except Exception as e:
        update_session_status("error", f"Error launching Chrome: {e}", error=str(e))
        return

    update_session_status("connecting", f"Navigating to: {meet_link}")
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

    update_session_status("connecting", "Handling lobby: muting mic & turning off camera...")
    joined = False

    for attempt in range(50):
        if not is_running:
            break
        time.sleep(0.3)
        try:
            status = global_driver.execute_script("""
                let res = { mic: false, cam: false, joined: false, button_text: '' };

                const allBtns = Array.from(document.querySelectorAll('button'));
                for (const b of allBtns) {
                    const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                    if (t === 'got it' || t === 'dismiss' || t.includes('use microphone') || t.includes('continue without')) {
                        b.click();
                    }
                }

                const interactive = Array.from(document.querySelectorAll('button, div[role="button"]'));
                for (const b of interactive) {
                    const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                    if (aria.includes('turn off microphone') || (aria.includes('microphone') && !aria.includes('turn on'))) {
                        b.click();
                        res.mic = true;
                        break;
                    }
                }

                for (const b of interactive) {
                    const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                    if (aria.includes('turn off camera') || (aria.includes('camera') && !aria.includes('turn on'))) {
                        b.click();
                        res.cam = true;
                        break;
                    }
                }

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
                if status.get("joined"):
                    btn_name = status.get("button_text") or "Ask to join"
                    add_activity_log("info", f"Tapped '{btn_name}' to join Google Meet lobby.")
                    joined = True
                    break
        except Exception:
            pass

        if attempt == 2:
            try:
                body = global_driver.find_element(By.TAG_NAME, "body")
                body.send_keys(Keys.CONTROL, "d")
                body.send_keys(Keys.CONTROL, "e")
            except Exception:
                pass

    if not joined:
        for sel in [
            "//button[contains(., 'Ask to join') or contains(., 'Join now')]",
            "//button[.//span[contains(text(),'Ask to join') or contains(text(),'Join now')]]",
            "//button[contains(@aria-label, 'Ask to join') or contains(@aria-label, 'Join now')]",
        ]:
            try:
                join_button = global_driver.find_element(By.XPATH, sel)
                if join_button.is_displayed():
                    global_driver.execute_script("arguments[0].click();", join_button)
                    add_activity_log("info", f"Tapped '{join_button.text.strip()}' button.")
                    joined = True
                    break
            except Exception:
                continue

    update_session_status("in_meeting", "Live in Google Meet room! Listening to lecture & monitoring attendance.")

    # Start Google Meet Closed Caption & Speaker STT monitor thread
    caption_thread = Thread(target=monitor_captions_loop, daemon=True)
    caption_thread.start()

    start_time = time.time()
    try:
        while is_running:
            if os.path.exists("stop_meeting.flag"):
                print("[Google Meet] Stop flag detected. Ending meeting...", flush=True)
                try:
                    os.remove("stop_meeting.flag")
                except Exception:
                    pass
                break

            # If user closed the Chrome window directly
            if global_driver:
                try:
                    _ = global_driver.title
                except Exception:
                    print("[Google Meet] Chrome window was closed. Ending meeting...", flush=True)
                    break

            # Support console 'q' key
            if sys.platform == "win32" and sys.stdin and sys.stdin.isatty():
                try:
                    import msvcrt
                    if msvcrt.kbhit():
                        ch = msvcrt.getch()
                        if ch in [b'q', b'Q'] and (time.time() - start_time > 3):
                            print("\n[Google Meet] 'q' detected. Class ended by student.", flush=True)
                            break
                except Exception:
                    pass

            time.sleep(0.5)
    except (KeyboardInterrupt, SystemExit):
        print("\n[Shutdown] Stopping agent gracefully...", flush=True)
    finally:
        is_running = False


def send_message_to_chat(message):
    """Automates sending an attendance response or Q&A answer into Google Meet chat."""
    global global_driver

    if not global_driver:
        return False

    for attempt in range(10):
        try:
            chat_state = global_driver.execute_script("""
                const inputs = Array.from(document.querySelectorAll('textarea, div[role="textbox"][contenteditable="true"]'));
                for (const inp of inputs) {
                    if (inp.offsetParent !== null && !inp.disabled && inp.getAttribute('aria-disabled') !== 'true') {
                        return { open: true, has_input: true };
                    }
                }

                const panel = document.querySelector('div[aria-label*="messages" i], div[aria-label*="Chat" i], div[data-tab-id="2"]');
                const isPanelOpen = panel && panel.offsetParent !== null;

                return { open: isPanelOpen, has_input: false };
            """)

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
                time.sleep(0.8)

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
                    return { success: false, reason: "Chat input not ready yet" };
                }

                targetInput.focus();
                if (targetInput.tagName.toLowerCase() === 'textarea') {
                    targetInput.value = msg;
                    targetInput.dispatchEvent(new Event('input', { bubbles: true }));
                    targetInput.dispatchEvent(new Event('change', { bubbles: true }));
                } else {
                    targetInput.innerText = msg;
                    targetInput.dispatchEvent(new InputEvent('input', { bubbles: true }));
                }

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
                print(f"✅ [Chat] Sent '{message}' into Google Meet chat!", flush=True)
                return True

            time.sleep(0.8)
        except Exception:
            time.sleep(0.8)

    return False


def run_meeting_session(meet_link):
    """
    Dedicated worker thread that runs the full end-to-end meeting session:
    1. Starts audio capture and Vosk speech recognition.
    2. Automates Chrome to join the user-provided Google Meet link.
    3. Listens to lecture and handles attendance in real-time.
    4. Compiles final notes with Gemini AI and syncs to Notion upon completion.
    """
    global is_running, global_driver, session_transcript_lines, recording_frames

    print(f"\n[Session] Starting background runner for: {meet_link}", flush=True)

    audio_thread = Thread(target=capture_audio_from_device, daemon=True)
    audio_thread.start()

    vosk_thread = Thread(target=speech_recognition_worker, daemon=True)
    vosk_thread.start()

    try:
        join_google_meet(meet_link, audio_thread)
    except Exception as e:
        print(f"[Session] Exception in join_google_meet: {e}", flush=True)
        update_session_status("error", f"Meeting error: {e}", error=str(e))
        add_activity_log("error", f"Meeting error: {e}")
        is_running = False
        return

    # Post-meeting processing
    is_running = False
    with session_lock:
        session_state["end_time"] = time.time()
    update_session_status("generating_notes", "Class finished. Compiling lecture transcript...")
    add_activity_log("info", "Meeting ended. Compiling lecture transcript and generating AI notes...")
    time.sleep(1.5)

    # 1. Compile final lecture transcript
    transcript = ""
    if session_transcript_lines:
        print(f"[Notes] Using real-time transcript ({len(session_transcript_lines)} dialogue entries).", flush=True)
        transcript = "\n".join(session_transcript_lines).strip()
    elif os.path.exists("transcript.txt") and os.path.getsize("transcript.txt") > 10:
        with open("transcript.txt", "r", encoding="utf-8") as f:
            transcript = f.read().strip()

    # 2. Multimodal STT Fallback if needed
    if not transcript or len(transcript.split()) < 5:
        update_session_status("generating_notes", "Analyzing session audio with Gemini multimodal...")
        audio_transcript = gemini_transcript("debug_audio.wav")
        if audio_transcript and len(audio_transcript.split()) >= 5:
            transcript = audio_transcript

    with session_lock:
        session_state["transcript"] = transcript

    notes = ""
    if transcript and len(transcript.strip().split()) >= 5:
        update_session_status("generating_notes", "Formatting structured notes with Google Gemini 3.8 Flash...")
        notes = gemini_make_notes(transcript)
        with session_lock:
            session_state["notes"] = notes
        add_activity_log("info", "Lecture notes generated successfully by Gemini.")

        # 3. Notion Integration
        if NOTION_API_TOKEN and NOTION_PAGE_ID:
            update_session_status("generating_notes", "Uploading notes to Notion...")
            try:
                from datetime import datetime
                timestamp_str = datetime.now().strftime("%b %d, %Y %I:%M %p")
                page_title = f"Class Notes - {STUDENT_NAME} ({timestamp_str})"
                notes_to_notion(
                    NOTION_API_TOKEN,
                    NOTION_PAGE_ID,
                    notes,
                    title=page_title,
                )
                with session_lock:
                    session_state["notion_uploaded"] = True
                    session_state["notion_title"] = page_title
                add_activity_log("info", f"Successfully synced notes to Notion: '{page_title}'")
            except Exception as ne:
                with session_lock:
                    session_state["notion_error"] = str(ne)
                add_activity_log("error", f"Notion upload note: {ne}")

        # 4. Generate Interactive Self-Evaluation Quiz from Notes
        if notes and len(notes.strip()) > 30:
            update_session_status("generating_notes", "Taking lecture notes to Gemini LLM to generate self-evaluation quiz...")
            add_activity_log("info", "Generating self-evaluation quiz (5-10 questions) from notes...")
            try:
                quiz_data = gemini_generate_quiz(notes, num_questions=5)
                with session_lock:
                    session_state["quiz"] = quiz_data
                if quiz_data:
                    add_activity_log("info", f"Self-evaluation quiz generated successfully ({len(quiz_data)} questions).")
            except Exception as qe:
                print(f"[Quiz Generation Error]: {qe}", flush=True)
    else:
        notes = "### Notice: No Spoken Lecture Detected\n\nNo spoken dialogue was detected during this session.\n\nPlease ensure your audio capture device or stereo mix is enabled while the instructor is speaking."
        with session_lock:
            session_state["notes"] = notes
        add_activity_log("info", "No spoken dialogue detected. Skipping Notion upload.")

    update_session_status("completed", "Lecture notes generated and ready!")

    if global_driver:
        try:
            global_driver.quit()
        except Exception:
            pass
        global_driver = None


# ==========================================
# FLASK REST API & WEB FRONT-END ROUTES
# ==========================================

@app.route("/")
def index():
    """Serves the front-end dashboard."""
    return render_template("index.html")


@app.route("/api/config")
def get_config():
    """Returns default student settings and Notion status."""
    return jsonify({
        "student_name": STUDENT_NAME,
        "trigger_keywords": ", ".join(TRIGGER_KEYWORDS),
        "notion_configured": bool(NOTION_API_TOKEN and NOTION_PAGE_ID),
        "default_meet_link": GMEET_LINK,
        "audio_device": AUDIO_DEVICE_NAME,
    })


@app.route("/api/start", methods=["POST"])
def api_start():
    """
    Receives the Google Meet link from the front-end user and proceeds with automation.
    """
    global is_running, session_worker_thread, STUDENT_NAME, TRIGGER_KEYWORDS

    if is_running:
        return jsonify({"error": "A meeting session is already in progress."}), 400

    data = request.get_json(silent=True) or {}
    raw_link = data.get("meet_link", "").strip()
    normalized_link, err = normalize_meet_link(raw_link)
    if err:
        return jsonify({"error": err}), 400

    student_name = data.get("student_name", "").strip() or STUDENT_NAME
    STUDENT_NAME = student_name
    trigger_kw_str = data.get("trigger_keywords", "").strip()
    if trigger_kw_str:
        TRIGGER_KEYWORDS = [k.strip().lower() for k in trigger_kw_str.split(",") if k.strip()]

    # Reset state and buffers for new session
    with session_lock:
        session_state["status"] = "launching"
        session_state["message"] = f"Launching agent for: {normalized_link}"
        session_state["meet_link"] = normalized_link
        session_state["student_name"] = STUDENT_NAME
        session_state["trigger_keywords"] = TRIGGER_KEYWORDS
        session_state["start_time"] = time.time()
        session_state["end_time"] = None
        session_state["captions"] = []
        session_state["activity_logs"] = []
        session_state["notes"] = ""
        session_state["transcript"] = ""
        session_state["notion_uploaded"] = False
        session_state["notion_title"] = ""
        session_state["notion_error"] = None
        session_state["error"] = None
        session_state["quiz"] = []
        session_state["quiz_result"] = None

    session_transcript_lines.clear()
    recording_frames.clear()
    rolling_history.clear()
    while not audio_queue.empty():
        try:
            audio_queue.get_nowait()
        except Exception:
            break

    with open("transcript.txt", "w", encoding="utf-8") as f:
        f.write("")

    if os.path.exists("stop_meeting.flag"):
        try:
            os.remove("stop_meeting.flag")
        except Exception:
            pass

    is_running = True
    add_activity_log("info", f"Agent initialized for student '{STUDENT_NAME}'")
    add_activity_log("info", f"Google Meet target link: {normalized_link}")

    session_worker_thread = Thread(target=run_meeting_session, args=(normalized_link,), daemon=True)
    session_worker_thread.start()

    return jsonify({"status": "started", "meet_link": normalized_link})


@app.route("/api/stop", methods=["POST"])
def api_stop():
    """Signals the meeting to stop and proceed with note generation."""
    global is_running

    if not is_running:
        return jsonify({"status": "already_stopped"}), 200

    is_running = False
    try:
        with open("stop_meeting.flag", "w") as f:
            f.write("stop")
    except Exception:
        pass

    update_session_status("generating_notes", "Class stopped. Compiling lecture notes with Gemini AI...")
    add_activity_log("info", "User requested session stop. Generating notes...")
    return jsonify({"status": "stopping", "message": "Stopping session and generating notes..."})


@app.route("/api/reset", methods=["POST"])
def api_reset():
    """Resets session state back to idle."""
    global is_running
    if is_running:
        return jsonify({"error": "Cannot reset while meeting is in progress."}), 400

    with session_lock:
        session_state["status"] = "idle"
        session_state["message"] = "Ready to join Google Meet session."
        session_state["meet_link"] = ""
        session_state["start_time"] = None
        session_state["captions"] = []
        session_state["activity_logs"] = []
        session_state["notes"] = ""
        session_state["transcript"] = ""
        session_state["notion_uploaded"] = False
        session_state["error"] = None
        session_state["quiz"] = []
        session_state["quiz_result"] = None

    return jsonify({"status": "reset"})


@app.route("/api/status")
def api_status():
    """Returns real-time session state, timer, captions, logs, and generated notes."""
    with session_lock:
        elapsed = 0
        if session_state["start_time"]:
            if session_state["status"] in ["launching", "connecting", "in_meeting"]:
                elapsed = int(time.time() - session_state["start_time"])
            elif session_state["end_time"]:
                elapsed = int(session_state["end_time"] - session_state["start_time"])

        return jsonify({
            "status": session_state["status"],
            "message": session_state["message"],
            "meet_link": session_state["meet_link"],
            "student_name": session_state["student_name"],
            "elapsed_seconds": elapsed,
            "captions_count": len(session_state["captions"]),
            "recent_captions": session_state["captions"][-40:],
            "activity_logs": session_state["activity_logs"][-30:],
            "notes": session_state["notes"],
            "notion_uploaded": session_state["notion_uploaded"],
            "notion_title": session_state["notion_title"],
            "notion_error": session_state["notion_error"],
            "error": session_state["error"],
            "quiz": session_state.get("quiz", []),
            "quiz_count": len(session_state.get("quiz", [])),
            "quiz_result": session_state.get("quiz_result")
        })




@app.route("/api/quiz")
def get_quiz():
    """Returns currently generated quiz questions and quiz result."""
    with session_lock:
        quiz = session_state.get("quiz", [])
        notes = session_state.get("notes", "")
        return jsonify({
            "quiz": quiz,
            "total": len(quiz),
            "has_notes": bool(notes),
            "quiz_result": session_state.get("quiz_result")
        })


@app.route("/api/quiz/generate", methods=["POST"])
def api_generate_quiz():
    """Generates or regenerates a 5-10 question quiz from lecture notes."""
    data = request.get_json(silent=True) or {}
    num_q = int(data.get("num_questions", 5))
    num_q = max(5, min(10, num_q))

    notes = session_state.get("notes", "")
    if not notes:
        if session_state.get("transcript"):
            notes = gemini_make_notes(session_state["transcript"])
            with session_lock:
                session_state["notes"] = notes
        else:
            return jsonify({"error": "No lecture notes found to generate a quiz. Please complete a class session or generate notes first."}), 400

    update_session_status(session_state["status"], f"Generating {num_q}-question quiz with Gemini LLM...")
    add_activity_log("info", f"Generating {num_q} quiz questions from lecture notes with Gemini...")
    quiz_data = gemini_generate_quiz(notes, num_questions=num_q)
    with session_lock:
        session_state["quiz"] = quiz_data
        session_state["quiz_result"] = None

    return jsonify({
        "status": "success",
        "quiz": quiz_data,
        "total": len(quiz_data)
    })


@app.route("/api/quiz/submit", methods=["POST"])
def api_submit_quiz():
    """Evaluates student quiz answers and provides diagnostic feedback."""
    data = request.get_json(silent=True) or {}
    answers = data.get("answers", {})
    quiz = session_state.get("quiz", [])
    notes = session_state.get("notes", "")

    if not quiz:
        return jsonify({"error": "No active quiz to evaluate."}), 400

    eval_result = gemini_evaluate_quiz(notes, quiz, answers)
    with session_lock:
        session_state["quiz_result"] = eval_result
    add_activity_log("info", f"Self-evaluation submitted: {eval_result['score']}/{eval_result['total']} ({eval_result['percentage']}%) - {eval_result['mastery_level']}")
    return jsonify(eval_result)

@app.route("/api/download/notes")
def download_notes():
    """Downloads generated notes as Markdown file."""
    notes = session_state.get("notes", "")
    filename = f"Class_Notes_{time.strftime('%Y%m%d_%H%M%S')}.md"
    return Response(
        notes,
        mimetype="text/markdown",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )


@app.route("/api/download/transcript")
def download_transcript():
    """Downloads lecture transcript as plain text file."""
    transcript = session_state.get("transcript", "")
    if not transcript and os.path.exists("transcript.txt"):
        with open("transcript.txt", "r", encoding="utf-8") as f:
            transcript = f.read()
    filename = f"Lecture_Transcript_{time.strftime('%Y%m%d_%H%M%S')}.txt"
    return Response(
        transcript,
        mimetype="text/plain",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )


@app.route("/process_audio", methods=["POST"])
def process_audio():
    """Flask endpoint for receiving streaming audio chunks."""
    try:
        audio_data = request.data
        try:
            audio_queue.put_nowait(audio_data)
        except queue.Full:
            pass
        return jsonify({"status": "Queued"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    signal.signal(signal.SIGINT, handle_exit)

    print("\n" + "=" * 65)
    print(" 🚀 AI STUDENT AGENT - WEB INTERFACE READY")
    print("=" * 65)
    print("  Student Name:       ", STUDENT_NAME)
    print("  Attendance Keywords:", TRIGGER_KEYWORDS)
    print("  Notion Integration: ", "Enabled" if (NOTION_API_TOKEN and NOTION_PAGE_ID) else "Not configured in .env")
    print("  Dashboard URL:       http://localhost:5000")
    print("=" * 65)
    print(" Opening web dashboard in your browser...")
    print(" Enter your Google Meet link on the dashboard to proceed.\n")

    def open_browser():
        time.sleep(1.2)
        try:
            import webbrowser
            webbrowser.open("http://localhost:5000")
        except Exception:
            pass

    Thread(target=open_browser, daemon=True).start()

    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False, threaded=True)
