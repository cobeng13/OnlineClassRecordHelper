#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

; ------------------ Globals ------------------
global running := false
global mainGui, pasteEdit, btnClip, initDelayEdit, keyDelayEdit
global advanceDDL, skipBlankCB, testModeCB, btnStart, btnStop

; ------------------ GUI ------------------
mainGui := Gui("+AlwaysOnTop", "Excel → Web Grade Typer")
mainGui.SetFont("s10", "Segoe UI")
mainGui.AddText(, "1) Copy cells from Excel  2) Paste below  3) Click first web cell  4) Start")

pasteEdit := mainGui.AddEdit("w650 r14")     ; where you paste Excel cells

btnClip := mainGui.AddButton("w130", "From Clipboard")
btnClip.OnEvent("Click", (*) => pasteEdit.Value := A_Clipboard)

mainGui.AddText("x+10 yp+2", "Initial delay (s):")
initDelayEdit := mainGui.AddEdit("w60"), initDelayEdit.Value := 3

mainGui.AddText("x+10 yp", "Key delay (ms):")
keyDelayEdit := mainGui.AddEdit("w60"), keyDelayEdit.Value := 40

mainGui.AddText("xm y+8", "Advance key (per value):")
advanceDDL := mainGui.AddDropDownList("w150", ["Enter","Tab","Down","Right"])
advanceDDL.Choose(1)

skipBlankCB := mainGui.AddCheckbox("xm", "Skip blank cells"),           skipBlankCB.Value := 1
testModeCB := mainGui.AddCheckbox("x+10", "Test mode (don't type)"),    testModeCB.Value := 0

btnStart := mainGui.AddButton("xm y+8 w160 Default", "Start (F9)")
btnStop  := mainGui.AddButton("x+10 w160", "Stop (F10)")
btnStart.OnEvent("Click", Start)
btnStop.OnEvent("Click", Stop)

mainGui.Show()

; Hotkeys
Hotkey "F9", Start
Hotkey "F10", Stop
Hotkey "Esc", Stop

; ------------------ Logic ------------------
Start(*) {
    global running, pasteEdit, initDelayEdit, keyDelayEdit, advanceDDL, skipBlankCB, testModeCB
    if running
        return

    txt := pasteEdit.Value
    if (txt = "") {
        MsgBox "Paste some Excel cells first.", "GradeTyper", "Icon!"
        return
    }

    running := true
    initDelay := initDelayEdit.Value + 0
    keyDelay  := keyDelayEdit.Value + 0
    advKey    := advanceDDL.Text
    skipBlank := !!skipBlankCB.Value
    tmode     := !!testModeCB.Value

    values := ParseExcelBlock(txt, skipBlank)

    ToolTip "Starting in " initDelay " s..." . "`nClick the FIRST cell in the website."
    Sleep initDelay * 1000
    ToolTip

    for v in values {
        if !running
            break

        if (v = "" && skipBlank) {
            SendAdvance(advKey, keyDelay, tmode)
            continue
        }

        if tmode {
            ToolTip "Would type: " v
            Sleep 350
            continue
        }

        SendText v
        Sleep keyDelay
        SendAdvance(advKey, keyDelay, false)
    }

    ; --- NEW: final confirm press so the last entry is saved ---
    if !tmode && (advKey != "Enter") {   ; avoid double-Enter if Enter was your advance key
        Send "{Enter}"
        Sleep keyDelay
    }

    running := false
    SoundBeep 1000, 120
    ToolTip "Done."
    SetTimer (() => ToolTip()), -800
}

Stop(*) {
    global running
    running := false
    ToolTip "Stopped."
    SetTimer (() => ToolTip()), -600
}

; Convert a tab/newline block (Excel copy) into a flat array of values
ParseExcelBlock(txt, skipBlank := true) {
    arr := []
    rows := StrSplit(RTrim(txt, "`r`n"), "`n")
    for , row in rows {
        row := StrReplace(row, "`r")             ; normalize CRLF -> LF
        cells := StrSplit(row, A_Tab)
        for , cell in cells {
            val := Trim(cell)
            if (val = "" && skipBlank)
                arr.Push("")                      ; keep position; will just advance
            else
                arr.Push(val)
        }
    }
    return arr
}

SendAdvance(key, dly, tmode := false) {
    if tmode
        return
    switch key {
        case "Enter": Send "{Enter}"
        case "Tab":   Send "{Tab}"
        case "Down":  Send "{Down}"
        case "Right": Send "{Right}"
        default:      Send "{Enter}"
    }
    Sleep dly
}
