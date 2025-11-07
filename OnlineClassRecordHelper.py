#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Online Class Record Helper (macOS version)
------------------------------------------
Paste grades copied from Excel → Auto-type them into a browser or online class record.
"""

import tkinter as tk
from tkinter import messagebox, scrolledtext
import pyautogui, time, threading, keyboard, sys

pyautogui.FAILSAFE = True
running = False


def type_data():
    """Perform the actual typing loop."""
    global running
    running = True
    raw = text_box.get("1.0", tk.END).strip()
    if not raw:
        messagebox.showwarning("No Data", "Please paste data from Excel first.")
        return

    data = raw.splitlines()
    messagebox.showinfo("Ready", "Click the FIRST web input cell within 5 seconds.")
    time.sleep(5)

    for i, line in enumerate(data, start=1):
        if not running:
            break
        value = line.strip()
        if value:
            pyautogui.typewrite(value)
            pyautogui.press("enter")
            time.sleep(0.15)  # slower pace for Safari/Chrome inputs
    messagebox.showinfo("Done", "Typing finished.")
    running = False


def start_thread():
    if running:
        messagebox.showwarning("Running", "Typing already in progress.")
        return
    t = threading.Thread(target=type_data, daemon=True)
    t.start()


def stop():
    global running
    running = False
    messagebox.showinfo("Stopped", "Typing stopped manually.")


def on_escape(event=None):
    stop()


# --- GUI SETUP ---
root = tk.Tk()
root.title("Excel → Web Grade Typer (macOS)")
root.geometry("460x380+600+200")
root.configure(bg="#f6f6f6")
root.attributes("-topmost", True)

tk.Label(
    root,
    text="1. Copy grades from Excel\n"
         "2. Paste below\n"
         "3. Click first web cell\n"
         "4. Click 'Start'",
    bg="#f6f6f6",
    font=("Helvetica", 11)
).pack(pady=10)

text_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=12, font=("Helvetica", 11))
text_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

frame = tk.Frame(root, bg="#f6f6f6")
frame.pack(pady=10)

tk.Button(frame, text="▶ Start", width=10, bg="#4CAF50", fg="white",
          font=("Helvetica", 11, "bold"), command=start_thread).grid(row=0, column=0, padx=8)
tk.Button(frame, text="■ Stop", width=10, bg="#f44336", fg="white",
          font=("Helvetica", 11, "bold"), command=stop).grid(row=0, column=1, padx=8)

root.bind("<Escape>", on_escape)
root.mainloop()
