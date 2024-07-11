import tkinter as tk
from tkinter import messagebox
import subprocess

def run_command():
    name = entry.get()
    if name:
        result = subprocess.run(['python', 'commandline_app.py', name], capture_output=True, text=True)
    else:
        result = subprocess.run(['python', 'commandline_app.py'], capture_output=True, text=True)
    
    messagebox.showinfo("Output", result.stdout)

# Create the main window
root = tk.Tk()
root.title("Command Line App Frontend")

# Create and place the widgets
label = tk.Label(root, text="Enter your name:")
label.pack(pady=10)

entry = tk.Entry(root)
entry.pack(pady=5)

button = tk.Button(root, text="Run Command", command=run_command)
button.pack(pady=20)

# Run the application
root.mainloop()