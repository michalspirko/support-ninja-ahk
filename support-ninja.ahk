; AutoHotkey v2 Script for Canned Responses with Dynamic Input and Shortcut Overlay
; Reads configuration from responses.conf and shortcuts.conf

#SingleInstance Force

; Initialize variables
Responses := Map()
ResponseOrder := []  ; NEW: Track order of responses
Shortcuts := Map()
ConfigFile := "responses.conf"
ShortcutsFile := "shortcuts.conf"
TypingBuffer := ""
OverlayGui := ""

; Read configuration on startup
ReadConfig()
ReadShortcuts()

; Hotkeys
^+r::ShowResponseMenu()
^+l::ShowResponseList()
^+s::ShowShortcutSelector()  ; New shortcut for shortcuts overlay
^+F12::ReloadConfig()

; Monitor typing for canned responses
~a::HandleKey("a")
~b::HandleKey("b")
~c::HandleKey("c")
~d::HandleKey("d")
~e::HandleKey("e")
~f::HandleKey("f")
~g::HandleKey("g")
~h::HandleKey("h")
~i::HandleKey("i")
~j::HandleKey("j")
~k::HandleKey("k")
~l::HandleKey("l")
~m::HandleKey("m")
~n::HandleKey("n")
~o::HandleKey("o")
~p::HandleKey("p")
~q::HandleKey("q")
~r::HandleKey("r")
~s::HandleKey("s")
~t::HandleKey("t")
~u::HandleKey("u")
~v::HandleKey("v")
~w::HandleKey("w")
~x::HandleKey("x")
~y::HandleKey("y")
~z::HandleKey("z")
~0::HandleKey("0")
~1::HandleKey("1")
~2::HandleKey("2")
~3::HandleKey("3")
~4::HandleKey("4")
~5::HandleKey("5")
~6::HandleKey("6")
~7::HandleKey("7")
~8::HandleKey("8")
~9::HandleKey("9")
~Space::HandleKey(" ")
~+::HandleKey("+")
~Enter::HandleKey("Enter")
~Backspace::HandleKey("Backspace")
~Tab::HandleKey("Tab")
~Delete::HandleKey("Delete")
~Escape::HandleKey("Escape")

; Read shortcuts configuration
ReadShortcuts() {
    global Shortcuts, ShortcutsFile
    
    Shortcuts.Clear()
    
    if (!FileExist(ShortcutsFile)) {
        MsgBox("Shortcuts file '" . ShortcutsFile . "' not found!", "Error", 0x30)
        return
    }
    
    FileContent := FileRead(ShortcutsFile)
    Lines := StrSplit(FileContent, "`n")
    CurrentSection := ""
    SectionOrder := []  ; Track section order
    
    for Line in Lines {
        Line := Trim(Line, " `t`r")
        
        ; Skip comments and empty lines
        if (SubStr(Line, 1, 1) = ";" || Line = "") {
            continue
        }
        
        ; Section headers
        if (RegExMatch(Line, "^\[(.+)\]$", &Match)) {
            CurrentSection := Trim(Match[1])
            if (!Shortcuts.Has(CurrentSection)) {
                Shortcuts[CurrentSection] := []
                SectionOrder.Push(CurrentSection)  ; Track order
            }
            continue
        }
        
        ; Shortcut entries
        if (CurrentSection != "" && InStr(Line, "=")) {
            Parts := StrSplit(Line, "=", " `t", 2)
            if (Parts.Length >= 2) {
                ShortcutKey := Trim(Parts[1])
                Description := Trim(Parts[2])
                Shortcuts[CurrentSection].Push({Key: ShortcutKey, Desc: Description})
            }
        }
    }
    
    ; Store section order
    Shortcuts["_SectionOrder"] := SectionOrder
}

; Show shortcut category selector
ShowShortcutSelector() {
    global Shortcuts
    
    if (Shortcuts.Count = 0) {
        MsgBox("No shortcuts loaded!", "Shortcut Selector", 0x40)
        return
    }
    
    ; Create GUI for category selection with slightly larger width
    SelectorGui := Gui("+LastFound", "Select Shortcut Category")
    SelectorGui.MarginX := 15
    SelectorGui.MarginY := 15
    
    ; Add instruction text
    SelectorGui.Add("Text", "x15 y15 w320", "Choose a shortcut category to display:")
    
    ; Create ListView for categories (increased width)
    LV := SelectorGui.Add("ListView", "x15 y45 w320 h200 -Multi", ["Category", "Shortcuts"])
    
    ; MODIFIED: Populate with categories in file order (not alphabetical)
    if (Shortcuts.Has("_SectionOrder")) {
        SectionOrder := Shortcuts["_SectionOrder"]
        for Category in SectionOrder {
            if (Shortcuts.Has(Category)) {
                ShortcutList := Shortcuts[Category]
                LV.Add("", Category, ShortcutList.Length . " shortcuts")
            }
        }
    } else {
        ; Fallback to Map iteration if order tracking fails
        for Category, ShortcutList in Shortcuts {
            if (Category != "_SectionOrder") {
                LV.Add("", Category, ShortcutList.Length . " shortcuts")
            }
        }
    }
    
    ; Auto-size columns
    LV.ModifyCol(1, "AutoHdr")
    LV.ModifyCol(2, "AutoHdr")
    
    ; Select first item
    if (LV.GetCount() > 0) {
        LV.Modify(1, "Select Focus")
    }
    
    ; Add buttons (adjusted for new width)
    ShowBtn := SelectorGui.Add("Button", "x15 y260 w80 h30 Default", "Show")
    CancelBtn := SelectorGui.Add("Button", "x255 y260 w80 h30", "Cancel")
    
    ; Button events
    ShowBtn.OnEvent("Click", ShowSelectedCategory)
    CancelBtn.OnEvent("Click", (*) => SelectorGui.Destroy())
    SelectorGui.OnEvent("Close", (*) => SelectorGui.Destroy())
    
    ; Double-click to show
    LV.OnEvent("DoubleClick", ShowSelectedCategory)
    
    ; Handle Enter and Escape keys
    HotIfWinActive("Select Shortcut Category")
    Hotkey("Enter", ShowSelectedCategory)
    Hotkey("Escape", (*) => SelectorGui.Destroy())
    HotIf()
    
    ; Show category function
    ShowSelectedCategory(*) {
        RowNumber := LV.GetNext()
        if (RowNumber = 0) {
            MsgBox("Please select a category first.", "No Selection", 0x40)
            return
        }
        
        CategoryName := LV.GetText(RowNumber, 1)
        SelectorGui.Destroy()
        Sleep(100)
        ShowShortcutOverlay(CategoryName)
    }
    
    ; Show the selector GUI (increased width)
    SelectorGui.Show("w350 h305")
}

; Show shortcut overlay
ShowShortcutOverlay(CategoryName) {
    global Shortcuts, OverlayGui
    
    ; Close existing overlay if any
    if (OverlayGui) {
        try OverlayGui.Destroy()
    }
    
    if (!Shortcuts.Has(CategoryName)) {
        MsgBox("Category '" . CategoryName . "' not found!", "Error", 0x30)
        return
    }
    
    ShortcutList := Shortcuts[CategoryName]
    
    ; Calculate optimal layout
    TotalShortcuts := ShortcutList.Length
    MaxColumns := 3  ; Maximum number of columns
    MaxRowsPerColumn := 15  ; Maximum rows per column for readability
    
    ; Determine number of columns needed
    NumColumns := Min(MaxColumns, Ceil(TotalShortcuts / MaxRowsPerColumn))
    if (NumColumns = 0) NumColumns := 1
    
    ; Calculate rows per column
    RowsPerColumn := Ceil(TotalShortcuts / NumColumns)
    
    ; Fixed dimensions for better consistency
    KeyColWidth := 140
    DescColWidth := 200
    ColumnWidth := KeyColWidth + DescColWidth + 10  ; 10px gap between key and desc
    Margin := 15
    ItemHeight := 22
    HeaderHeight := 50
    FooterHeight := 25
    
    ; Calculate total dimensions
    TotalWidth := (ColumnWidth * NumColumns) + (Margin * (NumColumns + 1))
    TotalHeight := HeaderHeight + (RowsPerColumn * ItemHeight) + FooterHeight + (Margin * 2)
    
    ; Ensure minimum and maximum sizes
    TotalWidth := Max(400, Min(TotalWidth, A_ScreenWidth * 0.9))
    TotalHeight := Max(300, Min(TotalHeight, A_ScreenHeight * 0.8))
    
    ; Create overlay GUI
    OverlayGui := Gui("+LastFound +AlwaysOnTop +ToolWindow -MaximizeBox -MinimizeBox", CategoryName . " - Shortcuts")
    OverlayGui.BackColor := "0x2D2D30"  ; Dark background
    OverlayGui.MarginX := 0
    OverlayGui.MarginY := 0
    
    ; Add semi-transparent background
    OverlayGui.Add("Text", "x0 y0 w" . TotalWidth . " h" . TotalHeight . " Background0x2D2D30")
    
    ; Title
    TitleText := OverlayGui.Add("Text", "x" . Margin . " y15 w" . (TotalWidth - Margin*2) . " h30 Center c0xFFFFFF BackgroundTrans", CategoryName)
    TitleText.SetFont("s12 Bold", "Segoe UI")
    
    ; Header line
    OverlayGui.Add("Text", "x" . Margin . " y45 w" . (TotalWidth - Margin*2) . " h1 Background0x555555")
    
    ; Start position for content
    ContentStartY := 55
    
    ; Add shortcuts in columns
    CurrentColumn := 0
    CurrentRow := 0
    
    for Index, Shortcut in ShortcutList {
        ; Calculate position
        ColumnX := Margin + (CurrentColumn * (ColumnWidth + Margin))
        ItemY := ContentStartY + (CurrentRow * ItemHeight)
        
        ; Alternate row colors within each column
        RowColor := (Mod(CurrentRow, 2) = 0) ? "0x3E3E42" : "0x2D2D30"
        
        ; Row background for this column section
        OverlayGui.Add("Text", "x" . ColumnX . " y" . ItemY . " w" . ColumnWidth . " h" . ItemHeight . " Background" . RowColor)
        
        ; Shortcut key (with better truncation)
        ShortcutKey := Shortcut.Key
        if (StrLen(ShortcutKey) > 25) {
            ShortcutKey := SubStr(ShortcutKey, 1, 22) . "..."
        }
        KeyText := OverlayGui.Add("Text", "x" . (ColumnX + 5) . " y" . (ItemY + 3) . " w" . (KeyColWidth - 10) . " h18 c0xFFD700 BackgroundTrans", ShortcutKey)
        KeyText.SetFont("s9 Bold", "Consolas")
        
        ; Description (with better truncation)
        Description := Shortcut.Desc
        if (StrLen(Description) > 45) {
            Description := SubStr(Description, 1, 42) . "..."
        }
        DescText := OverlayGui.Add("Text", "x" . (ColumnX + KeyColWidth + 10) . " y" . (ItemY + 3) . " w" . (DescColWidth - 10) . " h18 c0xF0F0F0 BackgroundTrans", Description)
        DescText.SetFont("s9", "Segoe UI")
        
        ; Move to next position
        CurrentRow++
        if (CurrentRow >= RowsPerColumn && CurrentColumn < NumColumns - 1) {
            CurrentColumn++
            CurrentRow := 0
        }
    }
    
    ; Add column separators
    if (NumColumns > 1) {
        Loop NumColumns - 1 {
            SeparatorX := Margin + (A_Index * (ColumnWidth + Margin)) - (Margin / 2)
            OverlayGui.Add("Text", "x" . SeparatorX . " y" . ContentStartY . " w1 h" . (RowsPerColumn * ItemHeight) . " Background0x555555")
        }
    }
    
    ; Footer separator
    FooterY := ContentStartY + (RowsPerColumn * ItemHeight) + 10
    OverlayGui.Add("Text", "x" . Margin . " y" . FooterY . " w" . (TotalWidth - Margin*2) . " h1 Background0x555555")
    
    ; Footer text
    FooterText := OverlayGui.Add("Text", "x" . Margin . " y" . (FooterY + 5) . " w" . (TotalWidth - Margin*2) . " h15 Center c0x888888 BackgroundTrans", "Press ESC to close • Click outside to close")
    FooterText.SetFont("s8", "Segoe UI")
    
    ; Recalculate actual height based on content
    ActualHeight := FooterY + 25
    
    ; Position overlay in center of screen
    X := Integer((A_ScreenWidth - TotalWidth) / 2)
    Y := Integer((A_ScreenHeight - ActualHeight) / 3)  ; Slightly above center
    
    ; Event handlers
    OverlayGui.OnEvent("Close", CloseOverlay)
    OverlayGui.OnEvent("Escape", CloseOverlay)
    
    ; Global hotkey for ESC when overlay is active
    HotIfWinActive(CategoryName . " - Shortcuts")
    Hotkey("Escape", CloseOverlay)
    HotIf()
    
    ; Function to close overlay
    CloseOverlay(*) {
        global OverlayGui
        if (OverlayGui) {
            try OverlayGui.Destroy()
            OverlayGui := ""
        }
    }
    
    ; Show overlay and make it active
    OverlayGui.Show("x" . X . " y" . Y . " w" . TotalWidth . " h" . ActualHeight)
    
    ; Activate the window so ESC works immediately
    WinActivate(CategoryName . " - Shortcuts")
    
    ; Set up click-outside-to-close functionality
    SetTimer(CheckClickOutside, 100)
    
    CheckClickOutside() {
        global OverlayGui
        if (!OverlayGui) {
            SetTimer(CheckClickOutside, 0)
            return
        }
        
        ; Check if mouse is clicked and overlay exists
        if (GetKeyState("LButton", "P")) {
            ; Get mouse position
            MouseGetPos(&MouseX, &MouseY, &WinUnderMouse)
            
            ; Get overlay window info
            try {
                OverlayGui.GetPos(&OverlayX, &OverlayY, &OverlayW, &OverlayH)
                
                ; Check if click is outside overlay
                if (MouseX < OverlayX || MouseX > OverlayX + OverlayW || 
                    MouseY < OverlayY || MouseY > OverlayY + OverlayH) {
                    ; Wait for button release
                    while (GetKeyState("LButton", "P")) {
                        Sleep(10)
                    }
                    CloseOverlay()
                }
            } catch {
                ; Overlay was destroyed, stop timer
                SetTimer(CheckClickOutside, 0)
            }
        }
    }
}

; Key handling function (unchanged)
HandleKey(Key) {
    global TypingBuffer, Responses
    
    ; Clear buffer on special keys or space
    if (Key = "Enter" || Key = "Tab" || Key = "Delete" || Key = "Escape" || Key = " ") {
        TypingBuffer := ""
        return
    }
    
    ; Handle backspace
    if (Key = "Backspace") {
        if (StrLen(TypingBuffer) > 0)
            TypingBuffer := SubStr(TypingBuffer, 1, -1)
        return
    }
    
    ; Add character to buffer
    TypingBuffer .= Key
    
    ; Limit buffer size
    if (StrLen(TypingBuffer) > 50)
        TypingBuffer := SubStr(TypingBuffer, -49)
    
    ; Check for +keyword pattern and trigger response
    if (RegExMatch(TypingBuffer, "\+([a-zA-Z0-9]+)$", &Match)) {
        Keyword := StrLower(Match[1])
        if (Responses.Has(Keyword)) {
            ; Clear the keyword first with sufficient delay
            KeywordLength := StrLen(Match[0])
            SendInput("{Backspace " . KeywordLength . "}")
            
            ; Wait for backspaces to complete before inserting response
            Sleep(150)  ; Increased delay to ensure deletion completes
            
            InsertResponse(Keyword)
            TypingBuffer := ""
        }
    }
}

; MODIFIED: Configuration file reader with order tracking
ReadConfig() {
    global Responses, ResponseOrder
    
    Responses.Clear()
    ResponseOrder := []  ; Reset order tracking
    
    if (!FileExist(ConfigFile)) {
        MsgBox("Configuration file '" . ConfigFile . "' not found!", "Error", 0x30)
        return
    }
    
    FileContent := FileRead(ConfigFile)
    Lines := StrSplit(FileContent, "`n")
    CurrentSection := ""
    i := 1
    
    while (i <= Lines.Length) {
        Line := RTrim(Lines[i], "`r")
        
        ; Skip comments
        if (SubStr(Line, 1, 1) = ";") {
            i++
            continue
        }
        
        ; Section headers
        if (RegExMatch(Line, "^\[(.+)\]$", &Match)) {
            CurrentSection := Match[1]
            i++
            continue
        }
        
        ; Key=value pairs
        if (RegExMatch(Line, "^([^=]+)=(.+)$", &Match)) {
            Key := Trim(Match[1])
            Value := Trim(Match[2])
            
            ; Handle multi-line triple-quoted values
            if (SubStr(Value, 1, 3) = '"""') {
                Value := SubStr(Value, 4)  ; Remove opening quotes
                MultiLineContent := ""
                i++
                
                while (i <= Lines.Length) {
                    NextLine := RTrim(Lines[i], "`r")
                    if (InStr(NextLine, '"""')) {
                        ; Add content before closing quotes
                        ClosePos := InStr(NextLine, '"""')
                        if (ClosePos > 1) {
                            FinalLine := SubStr(NextLine, 1, ClosePos - 1)
                            MultiLineContent .= (MultiLineContent = "" ? "" : "`n") . FinalLine
                        }
                        break
                    } else {
                        MultiLineContent .= (MultiLineContent = "" ? "" : "`n") . NextLine
                    }
                    i++
                }
                Value := MultiLineContent
            }
            
            ; Store responses only and track order
            if (CurrentSection = "Responses") {
                LowerKey := StrLower(Key)
                Responses[LowerKey] := Value
                ResponseOrder.Push(LowerKey)  ; Track order
            }
        }
        i++
    }
}

; Parse template placeholders from response text (unchanged)
ParsePlaceholders(ResponseText) {
    Placeholders := []
    Pos := 1
    
    while (Pos <= StrLen(ResponseText)) {
        ; Look for {{placeholder:label}} or {{placeholder}} pattern
        if (RegExMatch(ResponseText, "{{([^}:]+)(?::([^}]+))?}}", &Match, Pos)) {
            PlaceholderName := Trim(Match[1])
            PlaceholderLabel := Match[2] ? Trim(Match[2]) : PlaceholderName
            
            ; Check if we already have this placeholder
            Found := false
            for Item in Placeholders {
                if (Item.Name = PlaceholderName) {
                    Found := true
                    break
                }
            }
            
            ; Add unique placeholder
            if (!Found) {
                Placeholders.Push({Name: PlaceholderName, Label: PlaceholderLabel, Value: ""})
            }
            
            Pos := Match.Pos + Match.Len
        } else {
            break
        }
    }
    
    return Placeholders
}

; Show input dialog for dynamic placeholders (unchanged)
ShowInputDialog(Placeholders) {
    if (Placeholders.Length = 0)
        return true
    
    ; Create GUI with result tracking
    DialogResult := ""
    Controls := []
    
    ; Create GUI
    InputGui := Gui("+LastFound", "Enter Information")
    InputGui.MarginX := 15
    InputGui.MarginY := 15
    
    ; Add controls for each placeholder
    YPos := 15
    
    for Index, Placeholder in Placeholders {
        ; Label
        InputGui.Add("Text", "x15 y" . YPos . " w260", Placeholder.Label . ":")
        YPos += 25
        
        ; Input field
        EditControl := InputGui.Add("Edit", "x15 y" . YPos . " w260 vField" . Index)
        Controls.Push({Control: EditControl, Placeholder: Placeholder})
        YPos += 35
    }
    
    ; OK Button
    YPos += 10
    OKButton := InputGui.Add("Button", "x15 y" . YPos . " w80 h30 Default", "OK")
    OKButton.OnEvent("Click", OKClicked)
    
    ; Cancel Button
    CancelButton := InputGui.Add("Button", "x105 y" . YPos . " w80 h30", "Cancel")
    CancelButton.OnEvent("Click", CancelClicked)
    
    ; Handle window close (X button)
    InputGui.OnEvent("Close", CancelClicked)
    
    ; Set GUI size
    InputGui.Move(,, 300, YPos + 60)
    
    ; Show GUI and wait for result
    InputGui.Show()
    
    ; Event handlers
    OKClicked(*) {
        ; Get values from controls
        for Index, Item in Controls {
            Item.Placeholder.Value := Item.Control.Text
        }
        DialogResult := "OK"
        InputGui.Destroy()
    }
    
    CancelClicked(*) {
        DialogResult := "Cancel"
        InputGui.Destroy()
    }
    
    ; Add Escape key support for the input dialog
    HotIfWinActive("Enter Information")
    Hotkey("Escape", (*) => CancelClicked())
    HotIf()  ; Reset hotkey context
    
    ; Wait for dialog to complete
    while (DialogResult = "") {
        Sleep(50)
    }
    
    ; Return result
    return (DialogResult = "OK")
}

; Function to detect if current window supports HTML (unchanged)
IsHTMLCapableWindow() {
    ; Get current window title and process name
    WinTitle := WinGetTitle("A")
    ProcessName := WinGetProcessName("A")
    
    ; Check for Teams (supports HTML)
    if (InStr(ProcessName, "Teams") || InStr(WinTitle, "Teams") || InStr(WinTitle, "Microsoft Teams")) {
        return true
    }
    
    ; Check for Outlook (supports HTML)
    if (InStr(ProcessName, "OUTLOOK") || InStr(WinTitle, "Outlook")) {
        return true
    }
    
    ; Check for rich text editors that support HTML
    if (InStr(ProcessName, "WINWORD") || InStr(WinTitle, "Word")) {
        return true
    }
    
    ; Default to plain text for everything else (ServiceNow, web forms, etc.)
    return false
}

; Insert canned response with dynamic input support (unchanged)
InsertResponse(ResponseKey) {
    global Responses
    
    Response := Responses[ResponseKey]
    
    ; Parse placeholders
    Placeholders := ParsePlaceholders(Response)
    
    ; Show input dialog if placeholders exist
    if (Placeholders.Length > 0) {
        if (!ShowInputDialog(Placeholders)) {
            ; User cancelled
            return
        }
        
        ; Replace placeholders with user input
        for Placeholder in Placeholders {
            ; Replace both {{name:label}} and {{name}} patterns
            Response := StrReplace(Response, "{{" . Placeholder.Name . ":" . Placeholder.Label . "}}", Placeholder.Value)
            Response := StrReplace(Response, "{{" . Placeholder.Name . "}}", Placeholder.Value)
        }
    }
    
    ; Process standard replacements
    Response := StrReplace(Response, "[User]", "{User}")
    Response := StrReplace(Response, "`r`n", "`n")
    
    ; Detect if current window supports HTML
    UseHTML := IsHTMLCapableWindow()
    
    if (UseHTML) {
        ; HTML format for Teams, Outlook, etc.
        ; Convert [text](url) to HTML links
        Response := RegExReplace(Response, "\[([^\]]+)\]\(([^)]+)\)", '<a href="$2">$1</a>')
        Response := StrReplace(Response, "`n", "<br>`n")
        
        ; Set HTML content in clipboard
        SetHtmlClipboard(Response)
    } else {
        ; Plain text format for web forms, ServiceNow, etc.
        ; Convert [text](url) to plain text format: "text (url)"
        Response := RegExReplace(Response, "\[([^\]]+)\]\(([^)]+)\)", '$1 ($2)')
        ; Keep line breaks as-is (don't convert to <br>)
        
        ; Set plain text in clipboard
        A_Clipboard := Response
    }
    
    ; Ensure the target window is active and ready
    Sleep(100)
    
    ; Wait for clipboard to be set
    Sleep(300)
    
    ; Send paste command
    SendInput("^v")
    
    ; Wait for paste to complete
    Sleep(400)
}

; Set HTML format in clipboard for Teams (unchanged)
SetHtmlClipboard(HtmlText) {
    ; Create HTML document and plain text version
    HtmlDocument := "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>" . HtmlText . "</body></html>"
    PlainText := RegExReplace(StrReplace(HtmlText, "<br>", "`n"), "<[^>]*>", "")
    
    ; Backup clipboard
    ClipboardBackup := ClipboardAll()
    
    ; Clear and set clipboard data
    DllCall("OpenClipboard", "Ptr", 0)
    DllCall("EmptyClipboard")
    
    ; Set plain text
    hGlobalText := DllCall("GlobalAlloc", "UInt", 0x2000, "Ptr", (StrLen(PlainText) + 1) * 2, "Ptr")
    pGlobalText := DllCall("GlobalLock", "Ptr", hGlobalText, "Ptr")
    StrPut(PlainText, pGlobalText, "UTF-16")
    DllCall("GlobalUnlock", "Ptr", hGlobalText)
    DllCall("SetClipboardData", "UInt", 13, "Ptr", hGlobalText)
    
    ; Set HTML format
    HtmlFormat := "Version:0.9`r`nStartHTML:0000000000`r`nEndHTML:0000000000`r`nStartFragment:0000000000`r`nEndFragment:0000000000`r`n<html><body><!--StartFragment-->" . HtmlText . "<!--EndFragment--></body></html>"
    
    ; Calculate and update offsets
    StartHTML := InStr(HtmlFormat, "<html>") - 1
    EndHTML := StrLen(HtmlFormat)
    StartFragment := InStr(HtmlFormat, "<!--StartFragment-->") + 20 - 1
    EndFragment := InStr(HtmlFormat, "<!--EndFragment-->") - 1
    
    HtmlFormat := StrReplace(HtmlFormat, "StartHTML:0000000000", "StartHTML:" . Format("{:010d}", StartHTML))
    HtmlFormat := StrReplace(HtmlFormat, "EndHTML:0000000000", "EndHTML:" . Format("{:010d}", EndHTML))
    HtmlFormat := StrReplace(HtmlFormat, "StartFragment:0000000000", "StartFragment:" . Format("{:010d}", StartFragment))
    HtmlFormat := StrReplace(HtmlFormat, "EndFragment:0000000000", "EndFragment:" . Format("{:010d}", EndFragment))
    
    ; Register and set HTML format
    CF_HTML := DllCall("RegisterClipboardFormat", "Str", "HTML Format", "UInt")
    hGlobalHtml := DllCall("GlobalAlloc", "UInt", 0x2000, "Ptr", StrLen(HtmlFormat) + 1, "Ptr")
    pGlobalHtml := DllCall("GlobalLock", "Ptr", hGlobalHtml, "Ptr")
    StrPut(HtmlFormat, pGlobalHtml, "CP0")
    DllCall("GlobalUnlock", "Ptr", hGlobalHtml)
    DllCall("SetClipboardData", "UInt", CF_HTML, "Ptr", hGlobalHtml)
    
    DllCall("CloseClipboard")
    
    ; Restore clipboard after a delay
    SetTimer(() => (A_Clipboard := ClipboardBackup), -500)
}

; MODIFIED: Show response list for quick reference with file order
ShowResponseList() {
    global Responses, ResponseOrder
    
    if (Responses.Count = 0) {
        MsgBox("No responses loaded!", "Response List", 0x40)
        return
    }
    
    ; Create GUI for response list
    ListGui := Gui("+Resize +MinSize400x300", "Select Canned Response")
    ListGui.MarginX := 10
    ListGui.MarginY := 10
    
    ; Add instructions
    ListGui.Add("Text", "x10 y10 w380", "Available canned responses (Use arrows to navigate, Enter to use):")
    
    ; Create ListView
    LV := ListGui.Add("ListView", "x10 y35 w380 h200 VScroll", ["Keyword", "Preview"])
    
    ; Store full responses for tooltips
    FullResponses := Map()
    
    ; MODIFIED: Populate ListView using file order instead of alphabetical
    for ResponseKey in ResponseOrder {
        if (Responses.Has(ResponseKey)) {
            Value := Responses[ResponseKey]
            
            ; Get first line of response for preview
            FirstLine := StrSplit(Value, "`n")[1]
            if (StrLen(FirstLine) > 50)
                FirstLine := SubStr(FirstLine, 1, 47) . "..."
            
            ; Clean up HTML tags and placeholders for preview
            Preview := RegExReplace(FirstLine, "<[^>]*>", "")  ; Remove HTML
            Preview := RegExReplace(Preview, "{{[^}]+}}", "[?]")  ; Replace placeholders with [?]
            Preview := StrReplace(Preview, "[User]", "[User]")
            
            LV.Add("", "+" . ResponseKey, Preview)
            
            ; Store full response for tooltip (clean it up for display)
            CleanResponse := StrReplace(Value, "`r`n", "`n")
            CleanResponse := RegExReplace(CleanResponse, "<br[^>]*>", "`n")  ; Convert <br> to newlines
            CleanResponse := RegExReplace(CleanResponse, "<[^>]*>", "")       ; Remove other HTML tags
            CleanResponse := RegExReplace(CleanResponse, "{{([^}:]+)(?::([^}]+))?}}", "[$2]")  ; Show placeholder labels
            CleanResponse := StrReplace(CleanResponse, "[]", "[?]")          ; Fix empty labels
            FullResponses["+" . ResponseKey] := CleanResponse
        }
    }
    
    ; Auto-size columns
    LV.ModifyCol(1, "AutoHdr")
    LV.ModifyCol(2, "AutoHdr")
    
    ; Select the first item by default
    if (LV.GetCount() > 0) {
        LV.Modify(1, "Select Focus")
        ; Show initial tooltip
        ShowTooltipForSelected()
    }
    
    ; Handle selection changes
    LV.OnEvent("ItemSelect", (*) => ShowTooltipForSelected())
    LV.OnEvent("ItemFocus", (*) => ShowTooltipForSelected())
    
    ; Function to show tooltip for currently selected item
    ShowTooltipForSelected() {
        SelectedRow := LV.GetNext()
        if (SelectedRow > 0) {
            Keyword := LV.GetText(SelectedRow, 1)
            if (FullResponses.Has(Keyword)) {
                ResponseText := FullResponses[Keyword]
                
                ; Limit tooltip length
                if (StrLen(ResponseText) > 500) {
                    ResponseText := SubStr(ResponseText, 1, 497) . "..."
                }
                
                ; Position tooltip near the ListView
                ControlGetPos(&LvX, &LvY, &LvW, &LvH, LV)
                ToolTip(ResponseText, LvX + LvW + 10, LvY + 10)
            }
        } else {
            ToolTip()
        }
    }
    
    ; Add buttons
    ButtonY := 245
    UseBtn := ListGui.Add("Button", "x10 y" . ButtonY . " w80 h30", "Use")
    CloseBtn := ListGui.Add("Button", "x310 y" . ButtonY . " w80 h30", "Close")
    
    ; Button events
    UseBtn.OnEvent("Click", (*) => UseSelectedResponse())
    CloseBtn.OnEvent("Click", (*) => (ToolTip(), ListGui.Close()))
    ListGui.OnEvent("Close", (*) => (ToolTip(), ListGui.Destroy()))
    
    ; Double-click to use response
    LV.OnEvent("DoubleClick", (*) => UseSelectedResponse())
    
    ; Handle Enter key with hotkey when ListView is focused
    HotIfWinActive("Select Canned Response")
    Hotkey("Enter", (*) => UseSelectedResponse())
    Hotkey("Escape", (*) => (ToolTip(), ListGui.Destroy()))
    HotIf()  ; Reset hotkey context
    
    ; Function for button actions
    UseSelectedResponse() {
        RowNumber := LV.GetNext()
        if (RowNumber = 0) {
            MsgBox("Please select a response first.", "No Selection", 0x40)
            return
        }
        
        Keyword := LV.GetText(RowNumber, 1)  ; Get keyword with +
        ResponseKey := SubStr(Keyword, 2)    ; Remove the + sign
        
        ToolTip()  ; Clear tooltip
        ListGui.Destroy()
        Sleep(100)  ; Brief pause to let GUI close
        InsertResponse(ResponseKey)
    }
    
    ; Resize handler
    ListGui.OnEvent("Size", ResizeList)
    ResizeList(GuiObj, MinMax, Width, Height) {
        if (MinMax = -1)  ; Minimized
            return
        
        ; Resize ListView and reposition buttons
        LV.Move(,, Width - 20, Height - 85)
        ButtonY := Height - 45
        UseBtn.Move(, ButtonY)
        CloseBtn.Move(Width - 90, ButtonY)
        
        ; Refresh tooltip position after resize
        ShowTooltipForSelected()
    }
    
    ; Show the GUI
    ListGui.Show("w400 h285")
}

; MODIFIED: Show response menu with file order
ShowResponseMenu() {
    global Responses, ResponseOrder
    
    if (Responses.Count = 0)
        return
    
    ResponseMenu := Menu()
    ; Use file order instead of Map iteration
    for ResponseKey in ResponseOrder {
        if (Responses.Has(ResponseKey)) {
            ResponseMenu.Add(ResponseKey, (*) => InsertResponse(ResponseKey))
        }
    }
    ResponseMenu.Show()
}

; Reload configuration
ReloadConfig() {
    ReadConfig()
    ReadShortcuts()
    TrayTip("Configuration reloaded successfully!", "Configs Reloaded", 0x1)
}

; Edit configuration file function
EditConfig() {
    global ConfigFile
    
    ; Check if config file exists
    if (!FileExist(ConfigFile)) {
        MsgBox("Configuration file '" . ConfigFile . "' not found!", "Error", 0x30)
        return
    }
    
    ; Get full path to config file
    FullPath := A_WorkingDir . "\" . ConfigFile
    
    ; Try to open with default text editor (notepad as fallback)
    try {
        Run('notepad.exe "' . FullPath . '"')
    } catch {
        MsgBox("Could not open configuration file for editing.", "Error", 0x30)
    }
}

; Edit shortcuts file function
EditShortcuts() {
    global ShortcutsFile
    
    ; Check if shortcuts file exists
    if (!FileExist(ShortcutsFile)) {
        MsgBox("Shortcuts file '" . ShortcutsFile . "' not found!", "Error", 0x30)
        return
    }
    
    ; Get full path to shortcuts file
    FullPath := A_WorkingDir . "\" . ShortcutsFile
    
    ; Try to open with default text editor (notepad as fallback)
    try {
        Run('notepad.exe "' . FullPath . '"')
    } catch {
        MsgBox("Could not open shortcuts file for editing.", "Error", 0x30)
    }
}

; Tray menu setup
A_TrayMenu.Delete()  ; Clear all default menu items
A_TrayMenu.Add("Response List", (*) => ShowResponseList())
A_TrayMenu.Add("Shortcuts List", (*) => ShowShortcutSelector())
A_TrayMenu.Add()  ; Separator
A_TrayMenu.Add("Edit Responses", (*) => EditConfig())
A_TrayMenu.Add("Edit Shortcuts", (*) => EditShortcuts())
A_TrayMenu.Add("Reload Configs", (*) => ReloadConfig())
A_TrayMenu.Add("Reload Script", (*) => Reload())
A_TrayMenu.Add()  ; Separator
A_TrayMenu.Add("Exit", (*) => ExitApp())
A_TrayMenu.Default := "Response List"
A_IconTip := "Support Ninja AHK"

; Startup notification
TrayTip("Ctrl+Shift+S for shortcuts list`nCtrl+Shift+L for response list`nCtrl+Shift+F12 to reload configs", "Support Ninja AHK", 0x1)