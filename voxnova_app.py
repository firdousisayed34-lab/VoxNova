# voxnova_app.py
import streamlit as st
import pyttsx3
import tempfile
import os
import time
from typing import List, Tuple

st.set_page_config(page_title="VoxNova Studio", layout="centered")
st.title("🎙️ VoxNova Studio (Streamlit)")
st.write("Offline TTS powered by `pyttsx3`. Note: pyttsx3 may not work in some cloud environments (see note below).")

# ---- Helpers ----
@st.cache_resource
def init_engine():
    """Initialize pyttsx3 engine once (cached). May fail in headless/cloud."""
    try:
        engine = pyttsx3.init()
        return engine
    except Exception as e:
        return e  # return exception so caller can detect failure

@st.cache_data
def get_voices_list() -> List[Tuple[str, str]]:
    """Return list of (display_label, voice_id). Cached to avoid repeated engine init."""
    engine_or_exc = init_engine()
    if isinstance(engine_or_exc, Exception):
        return []
    engine = engine_or_exc
    try:
        voices = engine.getProperty("voices") or []
        vals = []
        for v in voices:
            vid = getattr(v, "id", "")
            name = getattr(v, "name", "") or vid
            lang = ""
            try:
                if hasattr(v, "languages") and v.languages:
                    lv = v.languages[0]
                    if isinstance(lv, bytes):
                        lv = lv.decode("utf-8", errors="ignore")
                    lang = f" ({lv})"
            except Exception:
                lang = ""
            # try detect gender token (best-effort)
            gender = getattr(v, "gender", "") or ""
            if isinstance(gender, str) and gender.strip():
                gender = gender.capitalize()
            else:
                gender = "Unknown"
            label = f"{name}{lang} — {gender} — {short_voice_id(vid)}"
            vals.append((label, vid))
        return vals
    except Exception:
        return []

def short_voice_id(vid: str) -> str:
    if not vid:
        return "default"
    sep = ":" if ":" in vid else "."
    return vid.split(sep)[-1]

def synthesize_to_wav(text: str, voice_id: str, rate: int, volume: float, timeout_sec=30) -> str:
    """
    Synthesize text to a temporary WAV file using pyttsx3 and return file path.
    Raises Exception on failure.
    """
    engine_or_exc = init_engine()
    if isinstance(engine_or_exc, Exception):
        raise RuntimeError(f"pyttsx3 init failed: {engine_or_exc}")
    engine = engine_or_exc

    # Create a temp file
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".wav")
    tmp.close()
    out_path = tmp.name

    # Configure engine
    try:
        engine.setProperty("rate", int(rate))
    except Exception:
        pass
    try:
        engine.setProperty("volume", max(0.0, min(1.0, float(volume))))
    except Exception:
        pass
    if voice_id:
        try:
            engine.setProperty("voice", voice_id)
        except Exception:
            pass

    # Save to file
    try:
        engine.save_to_file(text, out_path)
        # pyttsx3 runAndWait call is blocking until synthesis finished
        engine.runAndWait()
    except Exception as e:
        # remove partial file if any and re-raise
        try:
            if os.path.exists(out_path):
                os.remove(out_path)
        except Exception:
            pass
        raise RuntimeError(f"Failed to synthesize: {e}")

    if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
        raise RuntimeError("Synthesis produced empty file. Backend may be missing on this platform.")
    return out_path

# ---- UI ----
with st.sidebar:
    st.header("Settings")
    voices = get_voices_list()
    if voices:
        voice_labels = [v[0] for v in voices]
        default_idx = 0
        voice_choice = st.selectbox("Voice", voice_labels, index=default_idx)
        selected_voice_id = dict(voices).get(voice_choice, "")
    else:
        st.warning("No pyttsx3 voices found or pyttsx3 failed to initialize.")
        selected_voice_id = ""
    rate = st.slider("Rate (words per minute approx)", 80, 300, 160)
    volume = st.slider("Volume", 0.0, 1.0, 1.0)
    st.markdown("---")
    st.markdown("**Note:** If this is running on Streamlit Cloud, `pyttsx3` may not work (missing system TTS). If you see errors, run locally or use a cloud TTS API.")

st.header("Text to speak")
text = st.text_area("Enter text here", height=250, value="Type or paste your text here...")

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("🎙️ Preview (Generate & Play)"):
        if not text.strip():
            st.warning("Please enter some text to synthesize.")
        else:
            with st.spinner("Synthesizing audio..."):
                try:
                    wav_path = synthesize_to_wav(text, selected_voice_id, rate, volume)
                    # Read file bytes and present as audio + download
                    with open(wav_path, "rb") as f:
                        data = f.read()
                    st.audio(data, format="audio/wav")
                    st.success("Synthesis complete — playable below.")
                    # Offer download
                    fname = f"VoxNova_{int(time.time())}.wav"
                    st.download_button("💾 Download WAV", data=data, file_name=fname, mime="audio/wav")
                except Exception as e:
                    st.error(f"Synthesis error: {e}")
with col2:
    if st.button("💾 Save to server & download"):
        if not text.strip():
            st.warning("Please enter some text.")
        else:
            with st.spinner("Generating file..."):
                try:
                    wav_path = synthesize_to_wav(text, selected_voice_id, rate, volume)
                    with open(wav_path, "rb") as f:
                        data = f.read()
                    fname = f"VoxNova_{int(time.time())}.wav"
                    st.success("File generated — click download below.")
                    st.download_button("💾 Download WAV", data=data, file_name=fname, mime="audio/wav")
                    # cleanup
                    try:
                        os.remove(wav_path)
                    except Exception:
                        pass
                except Exception as e:
                    st.error(f"Synthesis error: {e}")

st.markdown("---")
st.caption("If voices list is empty or synthesis fails in cloud: run this app locally (on your machine) where system TTS backends are available, or use a cloud TTS provider.")

# Optional diagnostics
if st.checkbox("Show debug info"):
    st.subheader("Debug Info")
    st.write("pyttsx3 init result:", type(init_engine()).__name__)
    st.write("Available voices count:", len(voices))
    if voices:
        st.write([v[0] for v in voices])
