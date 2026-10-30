import os
import sys
import time
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

# Google Gemini Configuration
GEMINI_API_KEY = (os.getenv("GEMINI_API_KEY") or "").strip()

if not GEMINI_API_KEY:
    raise ValueError("GEMINI_API_KEY is not set in .env! Please provide your Gemini API key.")

# Initialize Google GenAI client
client = genai.Client(api_key=GEMINI_API_KEY)
print("[AI Provider] Google Gemini API initialized successfully.")

# High-availability fast models for attendance & real-time answers
QA_MODELS = ["gemini-3.5-flash-lite", "gemini-3.5-flash", "gemini-3.8-flash", "gemini-flash-latest"]

# Multimodal models for audio transcription and comprehensive lecture notes
CONTENT_MODELS = ["gemini-3.5-flash-lite", "gemini-3.5-flash", "gemini-3.8-flash", "gemini-flash-latest"]


def _call_gemini(models, contents, config=None, max_retries_per_model=2):
    """
    Executes a Gemini generation call across a list of models with automatic fallback.
    """
    last_error = None
    for model_name in models:
        for attempt in range(max_retries_per_model):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=config,
                )
                if response and response.text:
                    return response.text.strip()
                return ""
            except Exception as e:
                last_error = e
                err_str = str(e).lower()
                # If rate-limited or temporarily unavailable, pause and retry or fallback
                if "503" in err_str or "unavailable" in err_str or "429" in err_str:
                    time.sleep(1.0 * (attempt + 1))
                    continue
                else:
                    # Model not available or non-retriable, fall through to next model
                    break

    print(f"[Gemini Error] Generation failed across models {models}. Error: {last_error}")
    return None


def check_and_respond_with_gemini(context):
    """
    Uses Google Gemini to analyze meeting audio context in real-time.
    If attendance is called or a question is asked, generates a direct, accurate response.
    """
    if not context or not context.strip():
        return "Present!"

    print(f"[Gemini Q&A] Analyzing context: '{context}'", flush=True)

    prompt = f"""You are a helpful university student assistant analyzing an uncleaned, real-time speech-to-text transcript snippet from an ongoing live college class:
"{context}"

Instructions:
1. If the professor or host is calling for attendance or calling the student's name, respond with "Present, sir!" or "Present!".
2. If the context contains or resembles an academic question asked to the class or student (e.g. "What is X?", "Explain Y"), answer the question directly, concisely, and accurately in 1 to 2 clear sentences. Do NOT echo or repeat the question.
3. If you cannot detect any clear question or attendance call, respond with "Present!".

Answer:"""

    try:
        config = types.GenerateContentConfig(temperature=0.2, max_output_tokens=300)
        res = _call_gemini(QA_MODELS, prompt, config=config)
        if res:
            return res
    except Exception as e:
        print(f"[Gemini Error] Attendance/Q&A error: {e}")

    return "Present!"


def clean_hallucinations(text):
    """
    Strips known subtitle artifacts, silence loops, or noise fragments.
    """
    if not text:
        return ""

    hallucination_phrases = [
        "amara.org",
        "thank you for watching",
        "thank you for listening",
        "thanks for watching",
        "please subscribe",
        "subtitles by",
        "zero zero zero zero",
        "[no_speech]",
        "no speech",
        "ご視聴ありがとうございました",
        "もう、終わりだった",
        "kanskje vi kan lukka",
    ]

    lines = []
    for line in text.split("\n"):
        line_lower = line.lower().strip()
        if any(h in line_lower for h in hallucination_phrases):
            continue
        words = line.split()
        if len(words) > 3 and len(set(words)) <= 2:
            continue
        lines.append(line)

    cleaned = "\n".join(lines).strip()
    # If the resulting text is just single repeated words or very short noise
    unique_words = set(cleaned.lower().split())
    if len(unique_words) <= 2 and len(cleaned.split()) > 3:
        return ""

    return cleaned


def gemini_make_notes(transcript):
    """
    Uses Google Gemini to transform a raw lecture transcript into comprehensive,
    beautifully structured academic notes formatted in Markdown.
    """
    if not transcript or not transcript.strip():
        return "# Class Notes - Session Summary\n\nNo speech was recorded during this session."

    cleaned_transcript = clean_hallucinations(transcript)

    meaningful_words = [w for w in cleaned_transcript.split() if len(w) > 2]
    if len(meaningful_words) < 5:
        print("[Gemini Notes] Note: Not enough speech detected to generate full lecture notes.")
        return (
            f"# Class Notes - Session Summary\n\n"
            f"- **Status:** Meeting audio recorded.\n"
            f"- **Detected Speech:** \"{cleaned_transcript or 'No clear speech detected'}\"\n"
            f"- **Note:** The recorded audio did not contain an extended lecture discussion."
        )

    prompt = f"""You are an expert academic notetaker for university lectures.
Analyze the following transcript from a class lecture:
\"\"\"{cleaned_transcript}\"\"\"

Generate comprehensive, well-structured lecture notes in Markdown format:
- # Title / Topic of the Lecture
- ## Overview & Key Themes
- ## Detailed Discussion Points (clear bullet points)
- ## Key Definitions & Concepts
- ## Action Items / Deadlines (if any mentioned)

Strictly summarize only the educational and academic content present in the transcript. Do not invent topics not discussed. Output valid Markdown only."""

    try:
        config = types.GenerateContentConfig(temperature=0.2, max_output_tokens=3500)
        res = _call_gemini(CONTENT_MODELS, prompt, config=config)
        if res:
            return res
    except Exception as e:
        print(f"[Gemini Error] Notes formatting error: {e}")

    return f"# Class Notes\n\n## Transcript\n\n{cleaned_transcript}"


def gemini_transcript(audio_file_path="debug_audio.wav"):
    """
    Transcribes meeting audio recording using Google Gemini multimodal audio processing.
    Direct byte streaming for files <= 15MB; Gemini File API for files > 15MB.
    """
    if not os.path.exists(audio_file_path):
        print(f"[Audio Error] Audio file '{audio_file_path}' not found.")
        return None

    file_size_mb = os.path.getsize(audio_file_path) / (1024 * 1024)
    print(f"[Audio STT] Transcribing '{audio_file_path}' ({file_size_mb:.2f} MB) using Google Gemini...")

    prompt = (
        "Please provide an accurate transcription of all spoken words in this audio recording. "
        "Whenever possible, identify and label different speakers (e.g. Speaker 1:, Speaker 2:, or their names if mentioned). "
        "Output only the verbatim spoken dialogue with speaker labels without introductory commentary. "
        "If there is only silence, background noise, or no recognizable speech, reply with an empty string."
    )

    try:
        config = types.GenerateContentConfig(temperature=0.0)

        if file_size_mb <= 15:
            with open(audio_file_path, "rb") as f:
                audio_bytes = f.read()
            part = types.Part.from_bytes(data=audio_bytes, mime_type="audio/wav")
            res = _call_gemini(CONTENT_MODELS, [part, prompt], config=config)
        else:
            # For files > 15MB, upload using Gemini Files API
            uploaded_file = client.files.upload(file=audio_file_path)
            try:
                res = _call_gemini(CONTENT_MODELS, [uploaded_file, prompt], config=config)
            finally:
                try:
                    client.files.delete(name=uploaded_file.name)
                except Exception:
                    pass

        if res is not None:
            cleaned_text = clean_hallucinations(res)
            word_count = len(cleaned_text.split())
            if word_count > 0:
                print(f"[Audio STT] Gemini transcription completed successfully ({word_count} words).", flush=True)
                return cleaned_text
            else:
                print("[Audio STT] Gemini detected no intelligible speech in recorded audio.", flush=True)
                return ""
    except Exception as e:
        print(f"[Gemini Error] Audio transcription error: {e}")

    return None


# Backward-compatibility aliases so existing modules continue to work seamlessly
check_and_respond_with_openai = check_and_respond_with_gemini
openai_make_notes = gemini_make_notes
chatgpt_transcript = gemini_transcript


def gemini_generate_quiz(notes, num_questions=5):
    """
    Uses Google Gemini to generate a high-quality 5-10 question multiple-choice quiz
    derived directly from the lecture notes.
    Returns a list of structured question dictionaries.
    """
    import json

    if not notes or not notes.strip():
        return []

    num_q = max(5, min(10, int(num_questions or 5)))

    prompt = f"""You are an expert university professor creating a high-quality self-evaluation quiz for students based on these lecture notes:
\"\"\"{notes}\"\"\"

Generate a JSON array of {num_q} multiple-choice questions (MCQs) designed to evaluate student understanding of the key concepts, formulas, definitions, and discussion points from the lecture.

Requirements:
1. Generate exactly {num_q} questions.
2. The questions should test comprehension, application, and critical concepts, not just trivial trivia.
3. Each question must have:
   - "id": integer starting from 1 to {num_q}
   - "question": string, clear academic question
   - "options": array of exactly 4 distinct string choices (clean text without 'A)' prefixes)
   - "correct_index": integer (0, 1, 2, or 3) pointing to the correct choice in options
   - "explanation": string, clear explanation of why this answer is correct and why the concept matters based on the lecture
   - "topic": string, 2-3 word concept topic tag (e.g. "Neural Networks", "Regularization", "Cost Function")

Output strictly a valid JSON array."""

    try:
        config = types.GenerateContentConfig(
            temperature=0.3,
            response_mime_type="application/json",
            max_output_tokens=3500,
        )
        res = _call_gemini(CONTENT_MODELS, prompt, config=config)
        if res:
            clean_res = res.strip()
            if clean_res.startswith("```json"):
                clean_res = clean_res[7:]
            if clean_res.startswith("```"):
                clean_res = clean_res[3:]
            if clean_res.endswith("```"):
                clean_res = clean_res[:-3]
            clean_res = clean_res.strip()

            questions = json.loads(clean_res)
            if isinstance(questions, dict) and "questions" in questions:
                questions = questions["questions"]

            if isinstance(questions, list):
                valid_questions = []
                for idx, q in enumerate(questions, start=1):
                    if isinstance(q, dict) and "question" in q and "options" in q:
                        opts = [str(o) for o in q.get("options", [])]
                        if len(opts) >= 2:
                            c_idx = int(q.get("correct_index", 0))
                            if c_idx < 0 or c_idx >= len(opts):
                                c_idx = 0
                            valid_questions.append({
                                "id": idx,
                                "question": str(q["question"]),
                                "options": opts[:4],
                                "correct_index": c_idx,
                                "explanation": str(q.get("explanation", "Correct based on the lecture notes.")),
                                "topic": str(q.get("topic", "Lecture Concept")),
                            })
                if valid_questions:
                    print(f"[Gemini Quiz] Successfully generated {len(valid_questions)} quiz questions from lecture notes.", flush=True)
                    return valid_questions
    except Exception as e:
        print(f"[Gemini Error] Quiz generation error: {e}", flush=True)

    return []


def gemini_evaluate_quiz(notes, questions, student_answers):
    """
    Evaluates student quiz answers and provides a personalized diagnostic evaluation
    with strengths, weaknesses, and targeted study advice based on lecture notes.
    """
    total = len(questions)
    correct_count = 0
    breakdown = []

    for q in questions:
        q_id = str(q.get("id"))
        user_choice = student_answers.get(q_id)
        if user_choice is None:
            user_choice = student_answers.get(int(q_id))

        is_correct = False
        if user_choice is not None and int(user_choice) == int(q.get("correct_index", 0)):
            is_correct = True
            correct_count += 1

        breakdown.append({
            "id": q.get("id"),
            "question": q.get("question"),
            "topic": q.get("topic", "Lecture Concept"),
            "options": q.get("options", []),
            "selected_index": int(user_choice) if user_choice is not None else None,
            "correct_index": int(q.get("correct_index", 0)),
            "is_correct": is_correct,
            "explanation": q.get("explanation", "")
        })

    percentage = round((correct_count / total * 100), 1) if total > 0 else 0.0

    if percentage >= 85:
        mastery = "Mastery"
    elif percentage >= 65:
        mastery = "Proficient"
    elif percentage >= 45:
        mastery = "Developing"
    else:
        mastery = "Needs Review"

    advice = "Great effort! Review the explanations for questions you missed to reinforce your understanding."
    try:
        incorrect_topics = [b["topic"] for b in breakdown if not b["is_correct"]]
        if incorrect_topics:
            topics_str = ", ".join(set(incorrect_topics))
            prompt = f"""A student scored {correct_count}/{total} ({percentage}%) on their lecture self-evaluation quiz.
They struggled with these topics: {topics_str}.
Based on these notes:
\"\"\"{notes[:1500]}\"\"\"

Provide warm, encouraging, concise 2-sentence personalized feedback to guide their study focus before exams."""
            config = types.GenerateContentConfig(temperature=0.3, max_output_tokens=250)
            ai_fb = _call_gemini(QA_MODELS, prompt, config=config)
            if ai_fb:
                advice = ai_fb.strip()
        else:
            advice = f"Outstanding performance ({percentage}%)! You demonstrated a complete mastery of all concepts covered in this lecture."
    except Exception as e:
        print(f"[Gemini Feedback Error]: {e}")

    return {
        "score": correct_count,
        "total": total,
        "percentage": percentage,
        "mastery_level": mastery,
        "feedback": advice,
        "breakdown": breakdown
    }
