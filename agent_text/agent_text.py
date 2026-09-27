import sounddevice as sd
import numpy as np
from faster_whisper import WhisperModel
from core.llm import BrainLLM
import queue
import sys

print("⏳ Cargando STT Whisper (Modelo base)...")
model_stt = WhisperModel("base", device="cpu", compute_type="int8")
brain = BrainLLM()
audio_queue = queue.Queue()
FS = 16000
SILENCE_THRESHOLD = 0.01
SILENCE_DURATION = 1.5

def audio_callback(indata, frames, time, status):
    if status: print(status, file=sys.stderr)
    audio_queue.put(indata.copy())

print("\n🚀 BMO ESCUCHANDO (Modo Texto)")
print("================================")
print("Habla libremente. BMO detectará cuando dejes de hablar.")

def start_listening():
    with sd.InputStream(samplerate=FS, channels=1, callback=audio_callback):
        while True:
            audio_buffer = []
            print("\n🎤 Escuchando...")
            recording = False
            silent_chunks = 0
            while True:
                chunk = audio_queue.get()
                volume = np.linalg.norm(chunk) / np.sqrt(len(chunk))
                if volume > SILENCE_THRESHOLD:
                    if not recording:
                        print("🎙️ Voz detectada...")
                        recording = True
                    audio_buffer.append(chunk)
                    silent_chunks = 0
                elif recording:
                    audio_buffer.append(chunk)
                    silent_chunks += 1
                    if silent_chunks > int(SILENCE_DURATION * (FS / frames)):
                        break
            if audio_buffer:
                print("🧠 Transcribiendo...")
                audio_data = np.concatenate(audio_buffer).flatten()
                segments, _ = model_stt.transcribe(audio_data, language="es")
                text = " ".join([seg.text for seg in segments]).strip()
                if text:
                    print(f"👤 Tú: {text}")
                    print("🤖 Pensando...")
                    respuesta = brain.request_response(text)
                    print(f"🤖 BMO: {respuesta}")
                else:
                    print("❓ No logré entender lo que dijiste.")

try:
    start_listening()
except KeyboardInterrupt:
    print("\n👋 ¡Adiós!")
