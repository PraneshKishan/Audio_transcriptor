import streamlit as st
import os
import tempfile
import whisper
import pandas as pd
import ffmpeg
import time

def extract_audio(video_path, audio_path):
    """Extracts audio from a video file using ffmpeg and handles errors silently."""
    try:
        (
            ffmpeg
            .input(video_path)
            .output(audio_path, format='wav', acodec='pcm_s16le', ac=1, ar='16000')
            .run(overwrite_output=True, capture_stdout=True, capture_stderr=True)
        )
    except ffmpeg.Error:
        pass  # Suppress errors to ensure app continues running

def transcribe_audio(audio_path):
    """Transcribes audio using Whisper model."""
    model = whisper.load_model("base")  # Load a Whisper model (adjust size as needed)
    result = model.transcribe(audio_path)
    return result["text"].split(". ")  # Splitting into sentences

def save_to_excel(transcriptions, output_path):
    """Saves transcriptions to an Excel file with an empty 'Classification' column."""
    df = pd.DataFrame({'Sentence': transcriptions, 'Classification': [''] * len(transcriptions)})
    df.to_excel(output_path, index=False)

def main():
    st.title("Video to Text Transcriber")
    uploaded_file = st.file_uploader("Upload a video file", type=["mp4", "mov", "avi", "mkv"])
    
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.mp4') as temp_video:
            temp_video.write(uploaded_file.read())
            video_path = temp_video.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as temp_audio:
            audio_path = temp_audio.name
        
        if st.button("Convert"):
            with st.spinner("Extracting audio..."):
                extract_audio(video_path, audio_path)
            
            with st.spinner("Transcribing audio..."):
                transcriptions = transcribe_audio(audio_path)
            
            output_excel = "transcriptions.xlsx"
            save_to_excel(transcriptions, output_excel)
            
            st.success("Transcription complete! Download below.")
            st.download_button("Download Excel File", data=open(output_excel, "rb"), file_name="transcriptions.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # Ensure resources are freed before deletion
            del transcriptions  # Release any reference to the audio file
            time.sleep(1)  # Allow OS to release file lock
            
            os.remove(video_path)
            os.remove(audio_path)

if __name__ == "__main__":
    main()
