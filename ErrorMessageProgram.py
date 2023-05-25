import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
from ttkthemes import ThemedStyle
import datetime
import openai
import win32evtlog
import asyncio
import httpx
from pywintypes import Time as PyTime

# Set up OpenAI API credentials
# NOTE replace this API key with a personally generated one at https://platform.openai.com/account/api-keys
openai.api_key = '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'

# Global variables
previous_message = ""
processed_errors = set()
output_box = None
processing_frame = None
pending_errors = []
background_color = "#000000"
button_color = "#2c5aa3"
hover_color = "#3a7bd5"
foreground_color = "white"
font_family = "Helvetica"


async def query_chatgpt(prompt):
    async with httpx.AsyncClient() as client:
        response = await client.post(
            "https://api.openai.com/v1/engines/text-davinci-002/completions",
            json={
                "prompt": prompt,
                "max_tokens": 1024,
                "n": 1,
                "stop": None,
                "temperature": 1.0,
            },
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {openai.api_key}",
            },
            timeout=60.0,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["text"]

async def retrieve_error_messages(time_threshold=None):
    print("retrieve_error_messages started")
    server = 'localhost'
    log_type = 'System'
    h = win32evtlog.OpenEventLog(server, log_type)
    flags = win32evtlog.EVENTLOG_FORWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    total = win32evtlog.GetNumberOfEventLogRecords(h)
    num_records_to_read = 30000
    records_read = 0

    global processed_errors
    global output_box
    global pending_errors

    while records_read < num_records_to_read:
        events = win32evtlog.ReadEventLog(h, flags, 0)
        if not events:
            break

        for event in events:
            if time_threshold is not None:
                event_time = pytime_to_datetime(event.TimeGenerated)
                if event_time < time_threshold:
                    continue

            print("Event: EventType:", event.EventType, "SourceName:", event.SourceName, "EventID:", event.EventID)
            if event.EventType == win32evtlog.EVENTLOG_ERROR_TYPE:
                event_id = event.EventID
                source_name = event.SourceName

                error_message = f"Source: {source_name}, EventID: {event_id}"
                
                print("New error found:", error_message)
                if error_message not in processed_errors:
                    print("Processing error:", error_message)
                    processed_errors.add(error_message)

                    formatted_message = f"I'm experiencing an error with the following details: {error_message}. Can you help me fix it?"
                    explanation = await query_chatgpt(formatted_message)  
                    print("explanation received:", explanation.strip())  
                    
                    if output_box is None:
                        pending_errors.append((error_message, explanation))
                        print(f"Appending to pending_errors: {error_message}")
                    else:
                        print(f"Calling root.after for error_message: {error_message}")
                        root.after(0, update_output_box, error_message, explanation)

                else:
                    print(f"Error already processed: {error_message}")

            records_read += 1
            if records_read >= num_records_to_read:
                break

    if not processed_errors:
        hide_processing_message()
        show_no_errors_found()
    else:
        hide_processing_message()
        show_chat()

def process_pending_errors():
    global pending_errors

    if output_box is not None:
        for error_message, explanation in pending_errors:
            print(f"Processing pending error: {error_message}")
            root.after(0, update_output_box, error_message, explanation)

        pending_errors = []

def clear_processed_errors():
    global processed_errors
    processed_errors.clear()

def hide_home():
    global home_frame
    home_frame.pack_forget()

def show_home(hide_buttons=False):
    global home_frame
    
    home_frame = tk.Frame(root, bg=background_color)
    home_frame.pack(fill="both", expand=True)

    title_label = tk.Label(home_frame, text="ChatGPT Windows Event Log Helper", font=("Helvetica", 16, "bold"))
    title_label.pack(pady=20)

    event_log_button = tk.Button(home_frame, text="Check Windows Event Log", command=check_log_button_click)
    event_log_button.pack(pady=10)

    chat_button = tk.Button(home_frame, text="Chat with AI", command=lambda: show_chat(greeting=True))
    chat_button.pack(pady=10)

    if hide_buttons:
        event_log_button.pack_forget()
        chat_button.pack_forget()

async def check_log_button_click_async():
    hide_home()
    show_date_options()

def check_log_button_click():
    asyncio.create_task(check_log_button_click_async())

def show_date_options():
    global date_options_frame
    date_options_frame = tk.Frame(root, bg=background_color)
    date_options_frame.pack(fill="both", expand=True)

    title_label = tk.Label(date_options_frame, text="Select a Time Range", font=("Helvetica", 14, "bold"))
    title_label.pack(pady=20)

    label = tk.Label(date_options_frame, text="Select a time range to scan for errors:")
    label.pack(pady=10)

    one_day_button = ttk.Button(date_options_frame, text="Past 24 hours", command=lambda: asyncio.create_task(retrieve_error_messages_with_time_threshold(1)))
    one_day_button.pack(pady=5)

    three_days_button = ttk.Button(date_options_frame, text="Past 3 days", command=lambda: asyncio.create_task(retrieve_error_messages_with_time_threshold(3)))
    three_days_button.pack(pady=5)

    one_week_button = ttk.Button(date_options_frame, text="Past week", command=lambda: asyncio.create_task(retrieve_error_messages_with_time_threshold(7)))
    one_week_button.pack(pady=5)

    one_month_button = ttk.Button(date_options_frame, text="Past month", command=lambda: asyncio.create_task(retrieve_error_messages_with_time_threshold(30)))
    one_month_button.pack(pady=5)

    all_time_button = ttk.Button(date_options_frame, text="All time", command=lambda: asyncio.create_task(retrieve_error_messages()))
    all_time_button.pack(pady=5)

    back_button = ttk.Button(date_options_frame, text="Back", command=lambda: [hide_date_options(), show_home()])
    back_button.pack(pady=10)

def hide_date_options():
    global date_options_frame
    date_options_frame.pack_forget()

async def retrieve_error_messages_with_time_threshold(days):
    hide_date_options()
    show_processing_message()
    time_threshold = datetime.datetime.utcnow() - datetime.timedelta(days=days)
    await retrieve_error_messages(time_threshold)
    hide_processing_message()

def hide_processing_message():
    global processing_frame
    processing_frame.pack_forget()

def show_no_errors_found():
    hide_home()
    global home_frame

    home_frame = tk.Frame(root, bg=background_color)
    home_frame.pack(fill="both", expand=True)

    title_label = tk.Label(home_frame, text="No Errors Found", font=("Helvetica", 14, "bold"))
    title_label.pack(pady=20)

    label = tk.Label(home_frame, text="No error messages found in the Windows Event Log.")
    label.pack(pady=10)

    button = tk.Button(home_frame, text="Back", command=lambda: [hide_home(), show_home()])
    button.pack(pady=10)

def show_chat(greeting=False):
    global output_box
    hide_home()

    global chat_frame
    chat_frame = tk.Frame(root, bg=background_color)
    chat_frame.pack(fill="both", expand=True)

    global output_box
    output_box = tk.Text(chat_frame, wrap="word", width=80, height=20, state='disabled', bg="white", fg="black")
    output_box.pack(pady=10, padx=10, expand=True, fill="both")

    if greeting:
        greeting_message = "Hey there! How can I help you with any problems on your Windows device?"
    else:
        greeting_message = ""
    output_box.config(state='normal')
    output_box.insert("end", "ChatGPT: " + greeting_message + "\n\n")
    output_box.config(state='disabled')

    input_box = tk.Entry(chat_frame, width=80)
    input_box.pack(pady=10)
    input_box.bind("<Return>", chat_input_enter)

    send_button = ttk.Button(chat_frame, text="Send", command=lambda: asyncio.create_task(chat_input(input_box)))
    send_button.pack(pady=10)
    send_button.bind("<Enter>", on_enter)
    send_button.bind("<Leave>", on_leave)

    back_button = ttk.Button(chat_frame, text="Back", command=lambda: [hide_chat(), clear_processed_errors(), show_home()])
    back_button.pack(pady=10)
    back_button.bind("<Enter>", on_enter)
    back_button.bind("<Leave>", on_leave)

    if processed_errors:
        for error in processed_errors:
            output_box.config(state='normal')
            output_box.insert("end", "\nError: " + error + "\n")
            output_box.config(state='disabled')
    process_pending_errors()

# Button hover effects
def on_enter(e):
    e.widget.configure(style="Hover.TButton")

def on_leave(e):
    e.widget.configure(style="TButton")


def update_output_box(error_message, explanation):
    global output_box
    if output_box is not None:
        output_box.config(state='normal')
        output_box.insert("end", "\nError: " + error_message + "\n")
        output_box.insert("end", "\nChatGPT: " + explanation.strip() + "\n")  # Strip any extra whitespace
        output_box.insert("end", "\n" + "-"*80 + "\n")  # Add a separator line
        output_box.see("end")
        output_box.config(state='disabled')
        root.update_idletasks()  # Force the application to update the screen
    else:
        print("output_box is None")


def hide_chat():
    global chat_frame
    chat_frame.pack_forget()

def chat_input_enter(event):
    asyncio.create_task(chat_input(event.widget))

async def chat_input(input_box):
    global output_box
    global previous_message

    user_message = input_box.get()

    if user_message:
        output_box.config(state='normal')
        output_box.insert("end", "\nYou: " + user_message + "\n")
        input_box.delete(0, "end")
        output_box.see("end")
        output_box.config(state='disabled')

        prompt = f"{previous_message}\nYou: {user_message}\nChatGPT:"
        previous_message = prompt

        response = await query_chatgpt(prompt)

        output_box.config(state='normal')
        output_box.insert("end", "\nChatGPT: " + response + "\n")
        output_box.see("end")
        output_box.config(state='disabled')

def calculate_time_threshold(option):
    now = datetime.datetime.now()
    if option == "day":
        return now - datetime.timedelta(days=1)
    elif option == "3_days":
        return now - datetime.timedelta(days=3)
    elif option == "week":
        return now - datetime.timedelta(weeks=1)
    elif option == "month":
        return now - datetime.timedelta(days=30)
    else:
        return None

def show_processing_message():
    hide_home()

    global processing_frame
    processing_frame = tk.Frame(root, bg=background_color)
    processing_frame.pack(fill="both", expand=True)

    label = ttk.Label(processing_frame, text="Processing, please wait...")
    label.pack(pady=10)


async def run_tkinter_loop_async():
    while True:
        root.update()
        await asyncio.sleep(0.01)

async def main_async():
    await run_tkinter_loop_async()
    
def pytime_to_datetime(pytime):
    return datetime.datetime(pytime.year, pytime.month, pytime.day, pytime.hour, pytime.minute, pytime.second)

if __name__ == "__main__":
    root = ThemedTk(theme="breeze")
    root.title("ChatGPT Windows Event Log Helper")
    root.geometry("800x600")
    
    style = ttk.Style()
    style.configure("TLabel", background=background_color, foreground="white")

    style.configure("TButton", background="white", foreground="black", font=("Helvetica", 12), borderwidth=0, focuscolor=style.configure(".")["background"])
    style.map("TButton", background=[("active", "white"), ("pressed", "white")], foreground=[("active", "black"), ("pressed", "black")])
    style.configure("Hover.TButton", background="#f0f0f0", foreground="black")
    style.configure("TEntry", background="white", foreground="black")

    show_home()

    main_loop = asyncio.new_event_loop()
    asyncio.set_event_loop(main_loop)

    main_loop.run_until_complete(main_async())

