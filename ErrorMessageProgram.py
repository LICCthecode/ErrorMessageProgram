import tkinter as tk
import openai
import win32evtlog
import win32evtlogutil
import win32event
import win32security
import winerror

# Set up OpenAI API credentials
openai.api_key = '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'

# Define a function to retrieve and process error messages using OpenAI's GPT
def process_error_messages(chatbox):
    # Connect to the System event log
    server = 'localhost'
    log_type = 'System'
    h = win32evtlog.OpenEventLog(server, log_type)
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    total = win32evtlog.GetNumberOfEventLogRecords(h)
    events = win32evtlog.ReadEventLog(h, flags, 0)

    # Loop through the events and look for error messages
    for event in events:
        if event.EventType == win32evtlog.EVENTLOG_ERROR_TYPE:
            message = event.StringInserts
            if message:
                # Join the message strings into a single message
                error_message = " ".join(message)
                # Use OpenAI's GPT to generate a response to the error message
                response = openai.Completion.create(
                    engine="text-davinci-002",
                    prompt=error_message,
                    max_tokens=1024,
                    n=1,
                    stop=None,
                    temperature=0.5,
                )
                explanation = response.choices[0].text
                # Display the response in the output box
                output_box.insert("end", "\nChatGPT: " + explanation + "\n")
                # Scroll down to show the latest message
                output_box.see("end")



# Create the main window
root = tk.Tk()
root.geometry("500x500")
root.title("Error Message Solution")

# Create the chatbox
chatbox = tk.Text(root, width=60, height=10, font=("Arial", 12))
chatbox.pack(pady=10)

# Create the output box
output_box = tk.Text(root, width=60, height=30, font=("Arial", 12))
output_box.pack(pady=10)

# Create the send button
def send_message():
    # Get the user's message from the chatbox
    message = chatbox.get("1.0", "end-1c")
    # Send the user's message to OpenAI's GPT
    response = openai.Completion.create(
        engine="text-davinci-002",
        prompt=message,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.5,
    )
    # Display the response in the output box
    output_box.insert("end", "\nUser: " + message)
    output_box.insert("end", "\nChatGPT: " + response.choices[0].text + "\n")
    # Clear the chatbox for the next message
    chatbox.delete("1.0", "end")
    # Give the chatbox focus
    chatbox.focus()
    # Scroll down to the bottom of the output box
    output_box.see("end")




    
# Bind the Enter key to the send button
root.bind('<Return>', lambda event=None: send_message())

send_button = tk.Button(root, text="Send", font=("Arial", 12), command=send_message)
send_button.pack(pady=10)

print("Starting program...")
# Call the process_error_messages function to retrieve and process error messages
process_error_messages(chatbox)

# Start the main event loop
root.mainloop()

