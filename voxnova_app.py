import streamlit as st
import pyttsx3
import time
import os

# ------------------- Streamlit Config -------------------
st.set_page_config(
    page_title="VoxNova Studio — Text to Speech",
    layout="centered",
    page_icon="🎙️",
)

# ------------------- Engine Setup -------------------
@st.cache_resource
def init_engine():
    try:
        engine = pyttsx3.init('sapi5') if os.name == "nt" else pyttsx3.init()
        return engine
    except Exception as e:
        st.error(f"Error initializing voice engine: {e}")
        return None

engine = init_engine()
if engine is None:
    st.stop()

voices = engine.getProperty("voices")
voice_names = [v.name for v in voices]

# ------------------- Sidebar -------------------
st.sidebar.title("🎛️ Control Panel")
selected_voice_name = st.sidebar.selectbox("🎤 Select Voice", voice_names)
rate = st.sidebar.slider("🕒 Speech Rate (words/min)", 80, 220, 160)
volume = st.sidebar.slider("🔊 Volume", 0.0, 1.0, 1.0)
theme_toggle = st.sidebar.toggle("🌙 Dark Mode", value=False)

# ------------------- Header -------------------
st.markdown(
    """
    <div style="text-align:center; padding:15px; background-color:#1E3A8A; color:white; border-radius:10px;">
        <h1>🎙️ VoxNova Studio</h1>
        <p>Your personal AI Text-to-Speech Generator</p>
    </div>
    """,
    unsafe_allow_html=True
)

# ------------------- Text Input -------------------
text = st.text_area(
    "📝 Enter text below:",
    height=250,
    placeholder="Type or paste your text here to convert it into speech...",
)

col1, col2, col3 = st.columns(3)
speak = col1.button("▶️ Speak")
save = col2.button("💾 Save Audio")
clear = col3.button("🧹 Clear")

if clear:
    st.experimental_rerun()

# ------------------- Voice Configuration -------------------
for v in voices:
    if v.name == selected_voice_name:
        engine.setProperty("voice", v.id)
        break

engine.setProperty("rate", rate)
engine.setProperty("volume", volume)

# ------------------- Text Statistics -------------------
if text:
    words = len(text.split())
    chars = len(text)
    est_time = (words / max(1, rate)) * 60
    st.caption(f"🧾 **Words:** {words} | **Characters:** {chars} | ⏱️ Estimated Duration: {est_time:.1f} sec")

# ------------------- Actions -------------------
if speak:
    if not text.strip():
        st.warning("⚠️ Please type some text to speak.")
    else:
        with st.spinner("🎧 Speaking..."):
            engine.say(text)
            engine.runAndWait()
        st.success("✅ Done speaking!")

if save:
    if not text.strip():
        st.warning("⚠️ Please enter text before saving.")
    else:
        filename = f"VoxNova_{int(time.time())}.wav"
        with st.spinner(f"💾 Saving {filename}..."):
            engine.save_to_file(text, filename)
            engine.runAndWait()
        st.success(f"✅ Audio saved as `{filename}`")
        with open(filename, "rb") as f:
            st.download_button("⬇️ Download Audio File", f, file_name=filename, mime="audio/wav")

# ------------------- Theme -------------------
if theme_toggle:
    st.markdown("""
        <style>
        .stApp {
            background-color: #0E1117;
            color: #E5E7EB;
        }
        textarea, input, select, button {
            background-color: #1A1C23 !important;
            color: #E5E7EB !important;
        }
        </style>
    """, unsafe_allow_html=True)
