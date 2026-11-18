"""Opt-in microphone transcription and local Windows speech output."""
import threading


def listen(language='en-US'):
    import speech_recognition as sr
    recognizer = sr.Recognizer()
    recognizer.operation_timeout = 15
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        audio = recognizer.listen(source, timeout=5, phrase_time_limit=15)
    return recognizer.recognize_google(audio, language=language)


_speech_lock = threading.Lock()

def speak(text):
    import pythoncom
    import win32com.client
    if not _speech_lock.acquire(blocking=False):
        return
    try:
        pythoncom.CoInitialize()
        try:
            voice = win32com.client.Dispatch('SAPI.SpVoice')
            voice.Speak(str(text)[:1200])
        finally:
            pythoncom.CoUninitialize()
    finally:
        _speech_lock.release()
