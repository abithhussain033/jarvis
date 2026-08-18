import sounddevice as sd
import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import subprocess
import ollama
import json
import ctypes
import threading
import queue
import pythoncom
import customtkinter as ctk

from urllib.parse import quote


# ==================================================
# JARVIS V6
# SAFE GUI + FIXED SAPI VOICE
# ==================================================


PERSONAL_MEMORY_FILE = "jarvis_personal_memory.json"
CHAT_MEMORY_FILE = "jarvis_chat_memory.json"


# ==================================================
# GLOBAL CONTROL
# ==================================================

jarvis_running = True

speech_queue = queue.Queue()

personal_memory = []

chat_memory = []


# ==================================================
# GUI
# ==================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


app = ctk.CTk()

app.title("JARVIS AI")

app.geometry("1000x650")

app.resizable(False, False)


# ==================================================
# COLORS
# ==================================================

BG_COLOR = "#050914"
PANEL_COLOR = "#0B1220"
CYAN = "#00E5FF"
TEXT_COLOR = "#D9F9FF"


app.configure(

    fg_color=BG_COLOR

)


# ==================================================
# SAFE GUI UPDATES
# ==================================================

def safe_gui(function, *args):

    try:

        app.after(

            0,

            lambda: function(*args)

        )

    except Exception:

        pass


def update_status(text):

    status_label.configure(

        text="● " + text

    )


def update_command(text):

    command_label.configure(

        text=text

    )


def update_ai(text):

    ai_label.configure(

        text=text

    )


def update_memory_status(text):

    memory_label.configure(

        text=text

    )


# ==================================================
# HEADER
# ==================================================

title_label = ctk.CTkLabel(

    app,

    text="J A R V I S",

    font=("Arial", 32, "bold"),

    text_color=CYAN

)

title_label.pack(

    pady=(25, 0)

)


subtitle_label = ctk.CTkLabel(

    app,

    text="PERSONAL ARTIFICIAL INTELLIGENCE SYSTEM",

    font=("Arial", 12),

    text_color=TEXT_COLOR

)

subtitle_label.pack(

    pady=(0, 15)

)


# ==================================================
# LEFT CORE PANEL
# ==================================================

core_frame = ctk.CTkFrame(

    app,

    width=450,

    height=450,

    fg_color=PANEL_COLOR,

    corner_radius=25

)

core_frame.pack(

    side="left",

    padx=30,

    pady=25

)

core_frame.pack_propagate(False)


# ==================================================
# JARVIS V5 - ANIMATED AI FACE
# ==================================================

face_frame = ctk.CTkFrame(
    core_frame,
    width=360,
    height=360,
    fg_color="#06101C",
    corner_radius=180,
    border_width=3,
    border_color=CYAN
)

face_frame.place(
    relx=0.5,
    rely=0.45,
    anchor="center"
)

face_frame.pack_propagate(False)


# ==================================================
# FACE CANVAS
# ==================================================

face_canvas = ctk.CTkCanvas(
    face_frame,
    width=360,
    height=360,
    bg="#06101C",
    highlightthickness=0
)

face_canvas.pack()


# ==================================================
# JARVIS FACE
# ==================================================

face_canvas.create_oval(
    35,
    35,
    325,
    325,
    outline="#008CFF",
    width=3
)


face_canvas.create_oval(
    55,
    55,
    305,
    305,
    outline="#00E5FF",
    width=2
)


# ==================================================
# EYES
# ==================================================

left_eye = face_canvas.create_oval(
    90,
    130,
    150,
    175,
    fill=CYAN,
    outline=CYAN
)


right_eye = face_canvas.create_oval(
    210,
    130,
    270,
    175,
    fill=CYAN,
    outline=CYAN
)


# ==================================================
# EYE GLOW
# ==================================================

left_glow = face_canvas.create_oval(
    80,
    120,
    160,
    185,
    outline="#008CFF",
    width=2
)


right_glow = face_canvas.create_oval(
    200,
    120,
    280,
    185,
    outline="#008CFF",
    width=2
)


# ==================================================
# NOSE
# ==================================================

face_canvas.create_line(
    180,
    170,
    165,
    220,
    195,
    220,
    fill=CYAN,
    width=2
)


# ==================================================
# MOUTH
# ==================================================

mouth = face_canvas.create_arc(
    125,
    210,
    235,
    275,
    start=200,
    extent=140,
    style="arc",
    outline=CYAN,
    width=3
)


# ==================================================
# FACE ANIMATION
# ==================================================

face_pulse = 0
face_pulse_direction = 1
eye_open = True


def animate_face():

    global face_pulse
    global face_pulse_direction

    face_pulse += face_pulse_direction

    if face_pulse >= 8:

        face_pulse_direction = -1

    if face_pulse <= 0:

        face_pulse_direction = 1


    # Eye glow animation

    glow_width = 2 + face_pulse // 4


    face_canvas.itemconfigure(

        left_glow,

        width=glow_width

    )


    face_canvas.itemconfigure(

        right_glow,

        width=glow_width

    )


    # Face border animation

    face_canvas.itemconfigure(

        left_eye,

        fill=CYAN

    )


    face_canvas.itemconfigure(

        right_eye,

        fill=CYAN

    )


    face_frame.configure(

        border_width=2 + face_pulse // 4

    )


    face_frame.after(

        60,

        animate_face

    )


# Start face animation

animate_face()


# ==================================================
# BLINKING
# ==================================================

def blink():

    global eye_open

    eye_open = not eye_open


    if eye_open:

        face_canvas.coords(

            left_eye,

            90,

            130,

            150,

            175

        )


        face_canvas.coords(

            right_eye,

            210,

            130,

            270,

            175

        )


    else:

        face_canvas.coords(

            left_eye,

            90,

            150,

            150,

            155

        )


        face_canvas.coords(

            right_eye,

            210,

            150,

            270,

            155

        )


    face_frame.after(

        3500,

        blink

    )


blink()


# ==================================================
# STATUS
# ==================================================

status_label = ctk.CTkLabel(

    core_frame,

    text="● ONLINE",

    font=("Arial", 16, "bold"),

    text_color=CYAN

)

status_label.pack(

    side="bottom",

    pady=20

)


# ==================================================
# RIGHT PANEL
# ==================================================

right_frame = ctk.CTkFrame(

    app,

    width=450,

    height=450,

    fg_color=PANEL_COLOR,

    corner_radius=20

)

right_frame.pack(

    side="right",

    padx=30,

    pady=25

)

right_frame.pack_propagate(False)


# ==================================================
# COMMAND PANEL
# ==================================================

command_title = ctk.CTkLabel(

    right_frame,

    text="VOICE COMMAND",

    font=("Arial", 15, "bold"),

    text_color=CYAN

)

command_title.pack(

    pady=(25, 5)

)


command_label = ctk.CTkLabel(

    right_frame,

    text="Waiting for command...",

    font=("Arial", 16),

    text_color=TEXT_COLOR,

    wraplength=350

)

command_label.pack(

    pady=15

)


# ==================================================
# AI RESPONSE
# ==================================================

ai_title = ctk.CTkLabel(

    right_frame,

    text="JARVIS RESPONSE",

    font=("Arial", 15, "bold"),

    text_color=CYAN

)

ai_title.pack(

    pady=(35, 5)

)


ai_label = ctk.CTkLabel(

    right_frame,

    text="System ready.",

    font=("Arial", 16),

    text_color=TEXT_COLOR,

    wraplength=350

)

ai_label.pack(

    pady=15

)


# ==================================================
# MEMORY STATUS
# ==================================================

memory_label = ctk.CTkLabel(

    right_frame,

    text="PERSONAL MEMORY: ACTIVE",

    font=("Arial", 13, "bold"),

    text_color=CYAN

)

memory_label.pack(

    pady=(40, 0)

)


# ==================================================
# INITIALIZATION
# ==================================================

recognizer = sr.Recognizer()

recognizer.energy_threshold = 300

recognizer.dynamic_energy_threshold = True

recognizer.pause_threshold = 0.8


# ==================================================
# MEMORY
# ==================================================

def load_memory():

    global personal_memory


    try:

        with open(

            PERSONAL_MEMORY_FILE,

            "r",

            encoding="utf-8"

        ) as file:

            data = json.load(file)


            if isinstance(data, list):

                personal_memory = data

            else:

                personal_memory = []


    except Exception:

        personal_memory = []


    safe_gui(

        update_memory_status,

        "PERSONAL MEMORY: ACTIVE"

    )


def save_memory():

    try:

        with open(

            PERSONAL_MEMORY_FILE,

            "w",

            encoding="utf-8"

        ) as file:

            json.dump(

                personal_memory,

                file,

                indent=4

            )


    except Exception as error:

        print("Memory Save Error:", error)


load_memory()


# ==================================================
# SPEECH WORKER
# ==================================================

def speech_worker():

    pythoncom.CoInitialize()


    try:

        voice = win32com.client.Dispatch(

            "SAPI.SpVoice"

        )


        while True:

            text = speech_queue.get()


            if text is None:

                break


            try:

                print("Jarvis:", text)


                safe_gui(

                    update_ai,

                    text

                )


                voice.Speak(text)


            except Exception as error:

                print(

                    "Voice Error:",

                    error

                )


            speech_queue.task_done()


    finally:

        pythoncom.CoUninitialize()


speech_thread = threading.Thread(

    target=speech_worker,

    daemon=True

)

speech_thread.start()


# ==================================================
# SPEAK
# ==================================================

def speak(text):

    speech_queue.put(text)


# ==================================================
# LISTEN
# ==================================================

def listen():

    safe_gui(

        update_status,

        "LISTENING"

    )


    safe_gui(

        update_command,

        "Listening..."

    )


    print("Listening...")


    sample_rate = 44100

    max_duration = 5


    try:

        audio_data = sd.rec(

            int(max_duration * sample_rate),

            samplerate=sample_rate,

            channels=1,

            dtype="int16"

        )


        sd.wait()


        audio = sr.AudioData(

            audio_data.tobytes(),

            sample_rate,

            2

        )


        command = recognizer.recognize_google(

            audio,

            language="en-IN"

        )


        command = command.lower().strip()


        print(

            "You said:",

            command

        )


        safe_gui(

            update_command,

            command

        )


        safe_gui(

            update_status,

            "ONLINE"

        )


        return command


    except sr.UnknownValueError:

        print("Could not understand")


        safe_gui(

            update_status,

            "ONLINE"

        )


        return ""


    except sr.RequestError:

        print("Speech recognition error")


        safe_gui(

            update_status,

            "ONLINE"

        )


        return ""


    except Exception as error:

        print(

            "Microphone Error:",

            error

        )


        safe_gui(

            update_status,

            "ONLINE"

        )


        return ""


# ==================================================
# SAVE FACT
# ==================================================

def save_fact(fact_type, value):

    value = str(value).strip().lower()


    if not value:

        return


    # Remove old special memory

    if fact_type != "general":

        for item in personal_memory[:]:

            if (

                isinstance(item, dict)

                and item.get("type") == fact_type

            ):

                personal_memory.remove(item)


    # Prevent duplicate general memory

    if fact_type == "general":

        for item in personal_memory:

            if (

                isinstance(item, dict)

                and item.get("type") == "general"

                and item.get("value") == value

            ):

                return


    personal_memory.append({

        "type": fact_type,

        "value": value

    })


    save_memory()


# ==================================================
# REMEMBER FACT
# ==================================================

def remember_fact(command):

    command = command.lower().strip()


    # ==================================================
    # NAME
    # ==================================================

    if "my name is" in command:

        name = command.split(

            "my name is",

            1

        )[1].strip()


        if not name:

            return "What is your name boss?"


        save_fact(

            "name",

            name

        )


        return "I will remember your name, boss."


    # ==================================================
    # FAVOURITE COLOUR
    # ==================================================

    colour_phrases = [

        "my favourite colour is",

        "my favorite color is"

    ]


    for phrase in colour_phrases:

        if command.startswith(phrase):

            colour = command.replace(

                phrase,

                "",

                1

            ).strip()


            if not colour:

                return "What is your favourite colour boss?"


            save_fact(

                "colour",

                colour

            )


            return (

                "I will remember your favourite "

                "colour is "

                + colour

                + ", boss."

            )


    # ==================================================
    # LEARNING
    # ==================================================

    learning_phrases = [

        "i am learning",

        "i'm learning",

        "i am studying",

        "i'm studying",

        "i learn"

    ]


    for phrase in learning_phrases:

        if command.startswith(phrase):

            subject = command.replace(

                phrase,

                "",

                1

            ).strip()


            if not subject:

                return "What are you learning boss?"


            save_fact(

                "learning",

                subject

            )


            return (

                "I will remember that you are "

                "learning "

                + subject

                + ", boss."

            )


    # ==================================================
    # GOAL
    # ==================================================

    if (

        "i want to become an ai engineer"

        in command

        or "my goal is ai engineer"

        in command

        or "my goal is to become an ai engineer"

        in command

    ):

        save_fact(

            "goal",

            "become an ai engineer"

        )


        return "I will remember your goal, boss."


    # ==================================================
    # JARVIS PROJECT
    # ==================================================

    if (

        "i am building jarvis"

        in command

        or "i am making jarvis"

        in command

        or "i am creating jarvis"

        in command

    ):

        save_fact(

            "project",

            "building a personal ai assistant called jarvis"

        )


        return "I will remember your Jarvis project, boss."


    # ==================================================
    # GENERAL MEMORY
    # ==================================================

    remember_phrases = [

        "remember that",

        "remember my",

        "remember i"

    ]


    for phrase in remember_phrases:

        if command.startswith(phrase):

            fact = command.replace(

                phrase,

                "",

                1

            ).strip()


            if not fact:

                return "What should I remember boss?"


            save_fact(

                "general",

                fact

            )


            return "I will remember that boss."


    # ==================================================
    # AUTO MEMORY
    # ==================================================

    auto_phrases = [

        "i like",

        "i love",

        "i hate",

        "i prefer",

        "i enjoy"

    ]


    for phrase in auto_phrases:

        if command.startswith(phrase):

            save_fact(

                "general",

                command

            )


            return "I will remember that boss."


    return None


# ==================================================
# FORGET SPECIFIC MEMORY
# ==================================================

def forget_specific_memory(command):

    command = command.lower().strip()


    forget_phrases = [

        "forget that i like",

        "forget that i love",

        "forget that i hate",

        "forget that i prefer",

        "forget that i enjoy",

        "forget i like",

        "forget i love",

        "forget i hate",

        "forget i prefer",

        "forget i enjoy"

    ]


    target = None


    for phrase in forget_phrases:

        if command.startswith(phrase):

            target = command.replace(

                phrase,

                "",

                1

            ).strip()


            break


    if not target:

        return None


    if not target:

        return "What memory should I forget boss?"


    deleted = False


    for item in personal_memory[:]:

        if not isinstance(item, dict):

            continue


        if item.get("type") != "general":

            continue


        value = str(

            item.get("value", "")

        ).lower().strip()


        if value == target:

            personal_memory.remove(item)

            deleted = True

            continue


        memory_phrases = [

            "i like ",

            "i love ",

            "i hate ",

            "i prefer ",

            "i enjoy "

        ]


        for phrase in memory_phrases:

            if value.startswith(phrase):

                subject = value[len(phrase):].strip()


                if subject == target:

                    personal_memory.remove(item)

                    deleted = True

                    break


    if deleted:

        save_memory()

        return "I forgot that memory, boss."


    return "I could not find that memory, boss."


# ==================================================
# FORGET MEMORY
# ==================================================

def forget_memory(command):

    command = command.lower().strip()


    # ==================================================
    # FORGET ALL
    # ==================================================

    if (

        "forget all memory" in command

        or "delete all memory" in command

        or "clear my memory" in command

    ):

        personal_memory.clear()

        save_memory()


        return (

            "I forgot all your personal memory, boss."

        )


    # ==================================================
    # FORGET COLOUR
    # ==================================================

    if (

        "forget my favourite colour"

        in command

        or "forget my favorite color"

        in command

    ):

        for item in personal_memory[:]:

            if (

                isinstance(item, dict)

                and item.get("type") == "colour"

            ):

                personal_memory.remove(item)


        save_memory()


        return (

            "I forgot your favourite colour, boss."

        )


    # ==================================================
    # FORGET LEARNING
    # ==================================================

    if (

        "forget what i am learning"

        in command

        or "forget what i'm learning"

        in command

    ):

        for item in personal_memory[:]:

            if (

                isinstance(item, dict)

                and item.get("type") == "learning"

            ):

                personal_memory.remove(item)


        save_memory()


        return (

            "I forgot what you are learning, boss."

        )


    # ==================================================
    # FORGET NAME
    # ==================================================

    if "forget my name" in command:

        for item in personal_memory[:]:

            if (

                isinstance(item, dict)

                and item.get("type") == "name"

            ):

                personal_memory.remove(item)


        save_memory()


        return "I forgot your name, boss."


    return None


# ==================================================
# RECALL MEMORY
# ==================================================

def recall_memory(command):

    command = command.lower().strip()


    # ==================================================
    # NAME
    # ==================================================

    if (

        "what is my name" in command

        or "what's my name" in command

        or "who am i" in command

    ):

        for item in personal_memory:

            if (

                isinstance(item, dict)

                and item.get("type") == "name"

            ):

                return (

                    "Your name is "

                    + item.get("value")

                    + ", boss."

                )


        return "I don't remember your name yet, boss."


    # ==================================================
    # COLOUR
    # ==================================================

    if (

        "what is my favourite colour"

        in command

        or "what's my favourite colour"

        in command

        or "what is my favorite color"

        in command

        or "what's my favorite color"

        in command

        or "what colour do i like"

        in command

        or "what color do i like"

        in command

    ):

        for item in personal_memory:

            if (

                isinstance(item, dict)

                and item.get("type") == "colour"

            ):

                return (

                    "Your favourite colour is "

                    + item.get("value")

                    + ", boss."

                )


        return (

            "I don't remember your favourite "

            "colour yet, boss."

        )


    # ==================================================
    # LEARNING
    # ==================================================

    if (

        "what am i learning" in command

        or "what am i studying" in command

        or "what do i learn" in command

    ):

        learning = []


        for item in personal_memory:

            if (

                isinstance(item, dict)

                and item.get("type") == "learning"

            ):

                learning.append(

                    item.get("value")

                )


        if not learning:

            return (

                "I don't remember what you are "

                "learning yet, boss."

            )


        return (

            "You are learning "

            + ", ".join(learning)

            + ", boss."

        )


    # ==================================================
    # GOAL
    # ==================================================

    if (

        "what is my goal" in command

        or "what is my dream" in command

        or "what do i want to become" in command

    ):

        for item in personal_memory:

            if (

                isinstance(item, dict)

                and item.get("type") == "goal"

            ):

                return (

                    "Your goal is to "

                    + item.get("value")

                    + ", boss."

                )


        return "I don't remember your goal yet, boss."


    # ==================================================
    # PROJECT
    # ==================================================

    if (

        "what am i building" in command

        or "what project am i building" in command

    ):

        for item in personal_memory:

            if (

                isinstance(item, dict)

                and item.get("type") == "project"

            ):

                return (

                    "You are "

                    + item.get("value")

                    + ", boss."

                )


        return (

            "I don't remember that project yet, boss."

        )


    # ==================================================
    # ALL MEMORY
    # ==================================================

    if (

        "what do you remember"

        in command

        or "recall memory"

        in command

        or "show my memory"

        in command

    ):

        if not personal_memory:

            return (

                "I don't remember anything about you "

                "yet, boss."

            )


        facts = []


        for item in personal_memory:

            if not isinstance(item, dict):

                continue


            fact_type = item.get("type")

            value = item.get("value")


            if not value:

                continue


            if fact_type == "name":

                facts.append(

                    "your name is " + value

                )


            elif fact_type == "colour":

                facts.append(

                    "your favourite colour is "

                    + value

                )


            elif fact_type == "learning":

                facts.append(

                    "you are learning "

                    + value

                )


            elif fact_type == "goal":

                facts.append(

                    "your goal is "

                    + value

                )


            elif fact_type == "project":

                facts.append(

                    "you are " + value

                )


            elif fact_type == "general":

                if value.startswith("i like "):

                    facts.append(

                        "you like " + value[7:]

                    )


                elif value.startswith("i love "):

                    facts.append(

                        "you love " + value[7:]

                    )


                elif value.startswith("i hate "):

                    facts.append(

                        "you hate " + value[7:]

                    )


                elif value.startswith("i prefer "):

                    facts.append(

                        "you prefer " + value[9:]

                    )


                elif value.startswith("i enjoy "):

                    facts.append(

                        "you enjoy " + value[8:]

                    )


                else:

                    facts.append(value)


        if not facts:

            return (

                "I don't remember anything about you "

                "yet, boss."

            )


        return (

            "I remember that "

            + ", and ".join(facts)

            + ", boss."

        )


    return None


# ==================================================
# CALCULATOR
# ==================================================

def calculate(command):

    try:

        parts = command.split()


        if len(parts) < 4:

            return None


        num1 = float(parts[1])

        operator = parts[2]

        num2 = float(parts[3])


        if operator in [

            "plus",

            "+"

        ]:

            return num1 + num2


        if operator in [

            "minus",

            "-"

        ]:

            return num1 - num2


        if operator in [

            "multiply",

            "*",

            "x"

        ]:

            return num1 * num2


        if operator in [

            "divide",

            "/"

        ]:

            if num2 == 0:

                return None


            return num1 / num2


        return None


    except Exception:

        return None


# ==================================================
# YOUTUBE SEARCH
# ==================================================

def youtube_search(command):

    command = command.lower().strip()


    search_text = ""


    if command.startswith("search youtube"):

        search_text = command.replace(

            "search youtube",

            "",

            1

        ).strip()


    elif command.startswith("youtube search"):

        search_text = command.replace(

            "youtube search",

            "",

            1

        ).strip()


    elif command.startswith("youtube videos"):

        search_text = command.replace(

            "youtube videos",

            "",

            1

        ).strip()


    elif "on youtube" in command:

        search_text = command.replace(

            "on youtube",

            "",

            1

        ).strip()


        prefixes = [

            "search",

            "find",

            "watch",

            "show me",

            "play"

        ]


        for prefix in prefixes:

            if search_text.startswith(prefix):

                search_text = search_text[

                    len(prefix):

                ].strip()


                break


    else:

        return False


    if not search_text:

        speak(

            "What should I search on YouTube boss?"

        )


        return True


    speak(

        "Searching YouTube for "

        + search_text

    )


    url = (

        "https://www.youtube.com/results?search_query="

        + quote(search_text)

    )


    webbrowser.open(url)


    return True


# ==================================================
# SMART SEARCH
# ==================================================

def smart_search(command):

    command = command.lower().strip()


    if command in [

        "open youtube",

        "launch youtube",

        "start youtube",

        "youtube"

    ]:

        speak("Opening YouTube")


        webbrowser.open(

            "https://www.youtube.com"

        )


        return True


    if youtube_search(command):

        return True


    if command.startswith("search"):

        search_text = command[

            len("search"):

        ].strip()


        if not search_text:

            speak(

                "What should I search boss?"

            )


            return True


        speak(

            "Searching Google for "

            + search_text

        )


        webbrowser.open(

            "https://www.google.com/search?q="

            + quote(search_text)

        )


        return True


    return False


# ==================================================
# APP CONTROL
# ==================================================

def handle_command(command):

    command = command.lower().strip()


    if (

        command == "calculator"

        or "open calculator" in command

    ):

        speak("Opening Calculator")


        subprocess.Popen("calc.exe")


        return True


    if (

        command == "notepad"

        or "open notepad" in command

        or "take a note" in command

    ):

        speak("Opening Notepad")


        subprocess.Popen("notepad.exe")


        return True


    if (

        "open vs code" in command

        or "visual studio code" in command

        or "code editor" in command

    ):

        speak(

            "Opening Visual Studio Code"

        )


        try:

            subprocess.Popen("code")


        except Exception:

            speak(

                "VS Code is not available boss"

            )


        return True


    if command.startswith("open "):

        website = command[5:].strip()


        websites = {

            "instagram":

            "https://www.instagram.com",

            "youtube":

            "https://www.youtube.com",

            "google":

            "https://www.google.com",

            "github":

            "https://github.com",

            "chatgpt":

            "https://chatgpt.com",

            "netflix":

            "https://www.netflix.com",

            "amazon":

            "https://www.amazon.in",

            "spotify":

            "https://open.spotify.com",

            "reddit":

            "https://www.reddit.com"

        }


        if website in websites:

            speak(

                "Opening " + website

            )


            webbrowser.open(

                websites[website]

            )


            return True


        speak(

            "I don't know that website boss"

        )


        return True


    return False


# ==================================================
# WINDOWS CONTROL
# ==================================================

def windows_control(command):

    command = command.lower().strip()


    if "lock computer" in command:

        speak(

            "Locking your computer boss"

        )


        ctypes.windll.user32.LockWorkStation()


        return True


    if "shutdown computer" in command:

        speak(

            "Shutting down your computer boss"

        )


        subprocess.Popen(

            "shutdown /s /t 5",

            shell=True

        )


        return True


    if "restart computer" in command:

        speak(

            "Restarting your computer boss"

        )


        subprocess.Popen(

            "shutdown /r /t 5",

            shell=True

        )


        return True


    if "mute volume" in command:

        speak(

            "Muting volume boss"

        )


        subprocess.Popen(

            "powershell -command "

            "\"(New-Object -ComObject WScript.Shell)"

            ".SendKeys([char]173)\"",

            shell=True

        )


        return True


    return False


# ==================================================
# TIME AND DATE
# ==================================================

def time_and_date(command):

    command = command.lower().strip()


    if (

        "what time" in command

        or command == "time"

    ):

        current_time = datetime.datetime.now().strftime(

            "%I:%M %p"

        )


        speak(

            "The time is "

            + current_time

        )


        return True


    if (

        "what date" in command

        or command == "date"

    ):

        current_date = datetime.datetime.now().strftime(

            "%d %B %Y"

        )


        speak(

            "Today is "

            + current_date

        )


        return True


    return False


# ==================================================
# LOCAL AI
# ==================================================

def ai_brain(command):

    safe_gui(

        update_status,

        "THINKING"

    )


    try:

        response = ollama.chat(

            model="gemma3:270m",

            messages=[

                {

                    "role": "system",

                    "content": """

You are Jarvis, a personal AI assistant created by Abith Hussain.

Your name is Jarvis.

The user's name is Abith Hussain.

Answer in simple English.

Maximum 2 sentences.

Call the user boss when appropriate.

Never mention your model name.

Never say you have no memory.

Do not answer personal memory questions.

The Python memory system handles personal memory.

"""

                },

                {

                    "role": "user",

                    "content": command

                }

            ]

        )


        answer = response[

            "message"

        ][

            "content"

        ].strip()


        safe_gui(

            update_status,

            "ONLINE"

        )


        return answer


    except Exception as error:

        print(

            "AI Error:",

            error

        )


        safe_gui(

            update_status,

            "ONLINE"

        )


        return (

            "Sorry boss, my local AI "

            "is not responding."

        )


# ==================================================
# JARVIS LOOP
# ==================================================

def jarvis_loop():

    global jarvis_running


    speak(

        "Hello Abith boss, I am ready"

    )


    while jarvis_running:

        command = listen()


        if not command:

            continue


        if "jarvis" not in command:

            print(

                "Waiting for wake word..."

            )


            continue


        command = command.replace(

            "jarvis",

            "",

            1

        ).strip()


        # ==================================================
        # BARE JARVIS
        # ==================================================

        if command == "":

            speak(

                "Yes Abith boss"

            )


            command = listen()


            if not command:

                continue


        # ==================================================
        # STOP
        # ==================================================

        if command in [

            "stop",

            "exit",

            "shutdown jarvis"

        ]:

            speak(

                "Goodbye Abith boss"

            )


            jarvis_running = False


            app.after(

                1000,

                app.destroy

            )


            break


        handled = False


        # ==================================================
        # FORGET SPECIFIC MEMORY
        # ==================================================

        response = forget_specific_memory(command)


        if response:

            speak(response)

            handled = True


        # ==================================================
        # FORGET MEMORY
        # ==================================================

        if not handled:

            response = forget_memory(command)


            if response:

                speak(response)

                handled = True


        # ==================================================
        # RECALL MEMORY
        # ==================================================

        if not handled:

            response = recall_memory(command)


            if response:

                speak(response)

                handled = True


        # ==================================================
        # REMEMBER
        # ==================================================

        if not handled:

            response = remember_fact(command)


            if response:

                speak(response)

                handled = True


        # ==================================================
        # SEARCH
        # ==================================================

        if not handled:

            handled = smart_search(command)


        # ==================================================
        # APPS
        # ==================================================

        if not handled:

            handled = handle_command(command)


        # ==================================================
        # WINDOWS
        # ==================================================

        if not handled:

            handled = windows_control(command)


        # ==================================================
        # TIME AND DATE
        # ==================================================

        if not handled:

            handled = time_and_date(command)


        # ==================================================
        # CALCULATOR
        # ==================================================

        if (

            not handled

            and command.startswith("calculate")

        ):

            answer = calculate(command)


            if answer is not None:

                speak(

                    "The answer is "

                    + str(answer)

                )


            else:

                speak(

                    "Sorry boss, I could not calculate that"

                )


            handled = True


        # ==================================================
        # AI
        # ==================================================

        if not handled:

            answer = ai_brain(command)

            speak(answer)


# ==================================================
# CLOSE APP
# ==================================================

def close_app():

    global jarvis_running


    jarvis_running = False


    speech_queue.put(None)


    app.destroy()


app.protocol(

    "WM_DELETE_WINDOW",

    close_app

)


# ==================================================
# START JARVIS THREAD
# ==================================================

jarvis_thread = threading.Thread(

    target=jarvis_loop,

    daemon=True

)


jarvis_thread.start()


# ==================================================
# START GUI
# ==================================================

app.mainloop()