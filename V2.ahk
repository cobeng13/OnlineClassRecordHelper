#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

global running := false
global mainGui, pasteEdit, btnClip, initDelayEdit, keyDelayEdit
global advanceDDL, skipBlankCB, testModeCB, btnStart, btnStop
global colModeCB, colCountDDL

; ---------------- GUI ----------------
mainGui := Gui("+AlwaysOnTop", "Excel → Web Grade Typer")
mainGui.SetFont("s10", "Segoe UI")
mainGui.AddText(, "1) Copy Excel cells  2) Paste below  3) Click first web cell  4) Start")

pasteEdit := mainGui.AddEdit("w650 r14")
btnClip := mainGui.AddButton("w130", "From Clipboard")
btnClip.OnEvent("Click", (*) => pasteEdit.Value := A_Clipboard)

mainGui.AddText("x+10 yp+2", "Initial delay (s):")
initDelayEdit := mainGui.AddEdit("w60"), initDelayEdit.Value := 3

mainGui.AddText("x+10 yp", "Key delay (ms):")
keyDelayEdit := mainGui.AddEdit("w60"), keyDelayEdit.Value := 250

mainGui.AddText("xm y+8", "Advance key (single-column):")
advanceDDL := mainGui.AddDropDownList("w150", ["Enter","Tab","Down","Right"])
advanceDDL.Choose(1)

colModeCB := mainGui.AddCheckbox("xm y+8", "Multi-column mode")
colModeCB.Value := 0

mainGui.AddText("x+10 yp+2", "Columns:")
colCountDDL := mainGui.AddDropDownList("w60", ["1","2","3","4","5"])
colCountDDL.Choose(1)

skipBlankCB := mainGui.AddCheckbox("xm y+8", "Skip blanks"), skipBlankCB.Value := 1
testModeCB := mainGui.AddCheckbox("x+10", "Test mode"), testModeCB.Value := 0

btnStart := mainGui.AddButton("xm y+8 w160 Default", "Start (F9)")
btnStop  := mainGui.AddButton("x+10 w160", "Stop (F10)")
btnStart.OnEvent("Click", Start)
btnStop.OnEvent("Click", Stop)

mainGui.Show()

Hotkey "F9", Start
Hotkey "F10", Stop
Hotkey "Esc", Stop

; ==========================================================
;                        START
; ==========================================================
Start(*) {
    global running, pasteEdit, initDelayEdit, keyDelayEdit
    global advanceDDL, skipBlankCB, testModeCB
    global colModeCB, colCountDDL

    if running
        return

    txt := pasteEdit.Value
    if (txt = "") {
        MsgBox "No Excel data pasted.", "Error", "Icon!"
        return
    }

    running   := true
    initDelay := initDelayEdit.Value + 0
    keyDelay  := keyDelayEdit.Value + 0
    skipBlank := !!skipBlankCB.Value
    tmode     := !!testModeCB.Value
    multiCol  := !!colModeCB.Value
    colCount  := colCountDDL.Text + 0

    ToolTip "Starting in " initDelay " seconds..."
    Sleep initDelay * 1000
    ToolTip

    if !multiCol {
        RunSingleColumn(txt, advanceDDL.Text, keyDelay, skipBlank, tmode)
        return
    }

    RunMultiColumn(txt, colCount, keyDelay, skipBlank, tmode)
}

; ==========================================================
;                 SINGLE-COLUMN MODE
; ==========================================================
RunSingleColumn(txt, advKey, keyDelay, skipBlank, tmode) {
    values := ParseExcelBlockFlat(txt, skipBlank)

    for v in values {
        if !running
            break

        if (v = "" && skipBlank) {
            SendAdvance(advKey, keyDelay, tmode)
            continue
        }

        if tmode {
            ToolTip "Would type: " v
            Sleep 300
        } else {
            SendText v
            Sleep keyDelay
            SendAdvance(advKey, keyDelay)
        }
    }

    Finish()
}

; ==========================================================
;                 MULTI-COLUMN MODE (NEW LOGIC)
; ==========================================================
RunMultiColumn(txt, colCount, keyDelay, skipBlank, tmode) {
    matrix := ParseExcelBlock2D(txt)
    rowCount := matrix.Length

    ; Adjust column count to Excel data
    maxCols := 0
    for row in matrix
        if row.Length > maxCols
            maxCols := row.Length
    if colCount > maxCols
        colCount := maxCols

    ; -------------- COLUMN 1: DOWN --------------
    Loop rowCount {
        if !running
            break

        r := A_Index
        val := GetCell(matrix, r, 1)
        TypeCell(val, skipBlank, keyDelay, tmode)
    }

    ; If only one column, done
    if colCount = 1 {
        Finish()
        return
    }

    ; ============== COLUMNS 2–N (NEW PATTERN) ==============
    Loop (colCount - 1) {
        col := A_Index + 1

        ; STEP 1 → Move RIGHT to next column
        Send "{Right}"
        Sleep keyDelay

        ; STEP 2 → Go UP (rowCount - 1 times) to reach top
        Loop (rowCount - 1) {
            Send "{Up}"
            Sleep keyDelay
        }

        ; STEP 3 → Type DOWNWARD using Enter
        Loop rowCount {
            if !running
                break

            r := A_Index
            val := GetCell(matrix, r, col)
            TypeCell(val, skipBlank, keyDelay, tmode)
        }
    }

    Finish()
}

; ==========================================================
;                  HELPER FUNCTIONS
; ==========================================================
TypeCell(val, skipBlank, keyDelay, tmode) {
    if (val = "" && skipBlank) {
        if tmode {
            ToolTip "Skip blank"
            Sleep 200
        } else {
            Send "{Enter}"
            Sleep keyDelay
        }
        return
    }

    if tmode {
        ToolTip "Would type: " val
        Sleep 300
    } else {
        SendText val
        Sleep keyDelay
        Send "{Enter}"
        Sleep keyDelay
    }
}

GetCell(matrix, r, c) {
    return (c <= matrix[r].Length) ? Trim(matrix[r][c]) : ""
}

Finish() {
    global running
    running := false
    SoundBeep 1000, 120
    ToolTip "Done!"
    SetTimer(() => ToolTip(), -700)
}

Stop(*) {
    global running
    running := false
    ToolTip "Stopped."
    SetTimer(() => ToolTip(), -600)
}

ParseExcelBlockFlat(txt, skipBlank := true) {
    arr := []
    rows := StrSplit(RTrim(txt, "`r`n"), "`n")
    for row in rows {
        row := StrReplace(row, "`r")
        cells := StrSplit(row, A_Tab)
        for cell in cells {
            val := Trim(cell)
            arr.Push(val)
        }
    }
    return arr
}

ParseExcelBlock2D(txt) {
    mat := []
    rows := StrSplit(RTrim(txt, "`r`n"), "`n")
    for rowText in rows {
        rowText := StrReplace(rowText, "`r")
        cells := StrSplit(rowText, A_Tab)
        for i, cell in cells
            cells[i] := Trim(cell)
        mat.Push(cells)
    }
    return mat
}

SendAdvance(key, dly, tmode := false) {
    if tmode
        return
    switch key {
        case "Enter": Send "{Enter}"
        case "Tab":   Send "{Tab}"
        case "Down":  Send "{Down}"
        case "Right": Send "{Right}"
    }
    Sleep dly
}
