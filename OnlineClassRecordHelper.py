# Cross-platform "Excel → Web Grade Typer"
# GUI: PySimpleGUI | Typing: pyautogui
# Stop anytime: move mouse to a screen corner (pyautogui FAILSAFE) or press Esc (global).
import threading, time, sys
import pyautogui as pag
import PySimpleGUI as sg
from pynput import keyboard

pag.FAILSAFE = True   # move mouse to a corner to abort

running = False

def parse_excel_block(txt, skip_blank=True):
    rows = [r.replace("\r","") for r in txt.strip("\r\n").split("\n")]
    vals = []
    for r in rows:
        cells = r.split("\t")
        for c in cells:
            v = c.strip()
            if v == "" and skip_blank:
                vals.append("")
            else:
                vals.append(v)
    return vals

def send_advance(key):
    if key == "Enter": pag.press("enter")
    elif key == "Tab": pag.press("tab")
    elif key == "Down": pag.press("down")
    elif key == "Right": pag.press("right")
    else: pag.press("enter")

def worker(values, init_delay, key_delay, adv_key, skip_blank, test_mode):
    global running
    time.sleep(max(0, init_delay))
    for v in values:
        if not running:
            break
        if v == "" and skip_blank:
            if not test_mode:
                send_advance(adv_key)
                time.sleep(key_delay/1000.0)
            continue
        if test_mode:
            # brief on-screen indicator without typing
            sg.popup_no_titlebar(f"Would type: {v}", keep_on_top=True, auto_close=True, auto_close_duration=0.3)
        else:
            pag.typewrite(v)
            time.sleep(key_delay/1000.0)
            send_advance(adv_key)
            time.sleep(key_delay/1000.0)

    # Final confirm so the last entry is saved (if advance wasn't Enter)
    if running and not test_mode and adv_key != "Enter":
        pag.press("enter")
        time.sleep(key_delay/1000.0)

    running = False
    sg.popup_no_titlebar("Done.", keep_on_top=True, auto_close=True, auto_close_duration=0.7)

def on_key(key):
    """Global Esc to stop."""
    global running
    try:
        if key == keyboard.Key.esc:
            running = False
            return False  # stop listener
    except Exception:
        pass

def main():
    global running

    sg.theme("SystemDefault")
    layout = [
        [sg.Text("1) Copy cells from Excel  2) Paste below  3) Click first web cell  4) Start")],
        [sg.Multiline(key="-PASTE-", size=(80,15))],
        [sg.Button("From Clipboard"), sg.Text("Initial delay (s):"),
         sg.Input("3", key="-INIT-", size=(5,1)),
         sg.Text("Key delay (ms):"), sg.Input("40", key="-KDELAY-", size=(6,1))],
        [sg.Text("Advance key:"), sg.Combo(["Enter","Tab","Down","Right"], default_value="Enter", key="-ADV-", size=(10,1)),
         sg.CB("Skip blank cells", key="-SKIP-", default=True),
         sg.CB("Test mode (don't type)", key="-TEST-", default=False)],
        [sg.Button("Start", bind_return_key=True), sg.Button("Stop"), sg.Button("Exit")]
    ]
    win = sg.Window("Excel → Web Grade Typer (mac/win)", layout, keep_on_top=True, finalize=True)

    listener = None
    while True:
        ev, vals = win.read(timeout=100)
        if ev in (sg.WINDOW_CLOSED, "Exit"):
            running = False
            break
        if ev == "From Clipboard":
            try:
                win["-PASTE-"].update(sg.clipboard_get())
            except Exception as e:
                sg.popup_error(f"Clipboard error: {e}")
        if ev == "Start":
            txt = vals["-PASTE-"]
            if not txt.strip():
                sg.popup("Paste some Excel cells first.", keep_on_top=True)
                continue
            if running:
                continue
            try:
                init_delay = float(vals["-INIT-"])
                key_delay = float(vals["-KDELAY-"])
            except ValueError:
                sg.popup_error("Delays must be numbers.")
                continue
            adv_key = vals["-ADV-"]
            skip_blank = vals["-SKIP-"]
            test_mode = vals["-TEST-"]

            data = parse_excel_block(txt, skip_blank)
            running = True
            # global Esc listener
            listener = keyboard.Listener(on_press=on_key)
            listener.start()
            threading.Thread(target=worker, args=(data, init_delay, key_delay, adv_key, skip_blank, test_mode), daemon=True).start()
            sg.popup_no_titlebar(f"Starting in {init_delay}s.\nClick the FIRST web cell now.", keep_on_top=True, auto_close=True, auto_close_duration=1.5)

        if ev == "Stop":
            running = False
            if listener:
                listener.stop(); listener = None

    if listener:
        listener.stop()

if __name__ == "__main__":
    main()
