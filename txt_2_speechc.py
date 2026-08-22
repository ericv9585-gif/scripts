import pyttsx3
import tkinter as tk


# -----------------------------
# Get text from clipboard
# -----------------------------

root = tk.Tk()
root.withdraw()

try:
    text = root.clipboard_get()
except tk.TclError:
    print("The clipboard does not contain readable text.")
    input("Press Enter to exit...")
    raise SystemExit

root.destroy()

if not text.strip():
    print("Clipboard is empty.")
    input("Press Enter to exit...")
    raise SystemExit


# -----------------------------
# Speed presets
# -----------------------------

speeds = {
    "1": 150,
    "2": 175,
    "3": 200,
    "4": 225,
    "5": 250,
    "6": 300
}

print("\nClipboard Text-to-Speech")
print("------------------------")
print("Choose a speaking speed:")
print("1 - Slow")
print("2 - Normal")
print("3 - Fast")
print("4 - Very Fast")
print("5 - Extremely Fast")
print("6 - Maximum")

while True:
    choice = input("\nEnter speed (1-6): ").strip()

    if choice in speeds:
        rate = speeds[choice]
        break

    print("Please enter a number from 1 to 6.")


# -----------------------------
# Initialize text-to-speech
# -----------------------------

engine = pyttsx3.init()


# Try to select an English voice
for voice in engine.getProperty("voices"):
    if "english" in voice.name.lower():
        engine.setProperty("voice", voice.id)
        break


# Apply settings
engine.setProperty("rate", rate)
engine.setProperty("volume", 1.0)


# -----------------------------
# Speak clipboard
# -----------------------------

print(f"\nSpeaking clipboard at speed preset {choice} ({rate})...")
print("\nClipboard contents:")
print("-------------------")
print(text)
print("-------------------\n")

engine.say(text)
engine.runAndWait()

print("Finished.")
input("Press Enter to exit...")
