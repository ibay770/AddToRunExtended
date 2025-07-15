#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapp.icoer_GUI ****
#AutoIt3Wrapp.icoer_Icon=app.ico
#AutoIt3Wrapp.icoer_Outfile=AddToRun.exe
#AutoIt3Wrapp.icoer_Compression=4
#AutoIt3Wrapp.icoer_UseUpx=y
#AutoIt3Wrapp.icoer_UseX64=n
#AutoIt3Wrapp.icoer_Res_Comment=Program Alias Manager
#AutoIt3Wrapp.icoer_Res_Description=AddToRun - Create shortcuts to run programs from Start > Run
#AutoIt3Wrapp.icoer_Res_Fileversion=1.2.0.0
#AutoIt3Wrapp.icoer_Res_ProductName=AddToRun
#AutoIt3Wrapp.icoer_Res_ProductVersion=1.2
#AutoIt3Wrapp.icoer_Res_CompanyName=Halftime Software
#AutoIt3Wrapp.icoer_Res_LegalCopyright=Copyleft © 2024
#EndRegion ;**** Directives created by AutoIt3Wrapp.icoer_GUI ****

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>

Opt("GUIOnEventMode", 1)

; Global variables for drag and drop
Global $hDLL = DllOpen("shell32.dll")

; Create registry entry for context menu
CreateContextMenu()

#Region ### START Koda GUI section ### 
$main = GUICreate("AddToRun v1.1 - Program Alias Manager", 520, 450, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_MINIMIZEBOX, $WS_SIZEBOX))
GUISetBkColor(0xF5F5F5)
GUISetOnEvent($GUI_EVENT_CLOSE, "closew")

; Enable drag and drop for the main window
DllCall("shell32.dll", "none", "DragAcceptFiles", "hwnd", $main, "int", 1)

; Register for WM_DROPFILES message
GUIRegisterMsg($WM_DROPFILES, "WM_DROPFILES")

; Header section with icon and title
$headerGroup = GUICtrlCreateGroup("", 10, 10, 500, 80)
GUICtrlSetBkColor($headerGroup, 0xFFFFFF)
$Icon1 = GUICtrlCreateIcon("C:\WINDOWS\system32\shell32.dll", -131, 25, 25, 48, 48, BitOR($SS_NOTIFY, $WS_GROUP))
$Label2 = GUICtrlCreateLabel("AddToRun", 85, 20, 200, 35)
GUICtrlSetFont(-1, 18, 800, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x2C3E50)
$Label3 = GUICtrlCreateLabel("Program Alias Manager v1.1", 85, 50, 180, 18)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x7F8C8D)
$Label4 = GUICtrlCreateLabel("Create shortcuts to run programs from Start > Run", 85, 65, 300, 18)
GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x95A5A6)
GUICtrlCreateGroup("", -99, -99, 1, 1)

; Instructions section
$instructionGroup = GUICtrlCreateGroup(" Instructions ", 10, 100, 500, 60)
GUICtrlSetFont($instructionGroup, 9, 600, 0, "Segoe UI")
$instruction1 = GUICtrlCreateLabel("• Select a program executable file or drag & drop it into the program field", 25, 120, 460, 18)
GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
$instruction2 = GUICtrlCreateLabel("• Enter a short alias name (e.g., 'np' for Notepad, 'chrome' for Chrome)", 25, 135, 460, 18)
GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
GUICtrlCreateGroup("", -99, -99, 1, 1)

; Program selection section
$Group1 = GUICtrlCreateGroup(" 1. Select Program ", 10, 170, 500, 80)
GUICtrlSetFont($Group1, 9, 600, 0, "Segoe UI")
$Label5 = GUICtrlCreateLabel("Program Path:", 25, 195, 80, 18)
GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
$filename = GUICtrlCreateEdit("", 25, 210, 350, 25, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetFont($filename, 8, 400, 0, "Segoe UI")
$slect = GUICtrlCreateButton("Browse...", 385, 210, 80, 25, $WS_GROUP)
GUICtrlSetFont($slect, 8, 400, 0, "Segoe UI")
GUICtrlSetOnEvent($slect, "slectf")
$dragDropLabel = GUICtrlCreateLabel("(You can also drag & drop files here)", 25, 230, 200, 15)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x7F8C8D)
GUICtrlCreateGroup("", -99, -99, 1, 1)

; Alias section
$Group2 = GUICtrlCreateGroup(" 2. Set Alias Name ", 10, 260, 500, 80)
GUICtrlSetFont($Group2, 9, 600, 0, "Segoe UI")
$Label1 = GUICtrlCreateLabel("Alias Name:", 25, 285, 70, 18)
GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
$alias = GUICtrlCreateEdit("", 25, 300, 150, 25, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL))
GUICtrlSetFont($alias, 8, 400, 0, "Segoe UI")
$add = GUICtrlCreateButton("Add Alias", 190, 300, 80, 25, $WS_GROUP)
GUICtrlSetFont($add, 8, 600, 0, "Segoe UI")
GUICtrlSetBkColor($add, 0x27AE60)
GUICtrlSetColor($add, 0xFFFFFF)
GUICtrlSetOnEvent($add, "add")
$remove1 = GUICtrlCreateButton("Remove Alias", 280, 300, 80, 25, $WS_GROUP)
GUICtrlSetFont($remove1, 8, 600, 0, "Segoe UI")
GUICtrlSetBkColor($remove1, 0xE74C3C)
GUICtrlSetColor($remove1, 0xFFFFFF)
GUICtrlSetOnEvent($remove1, "remove")
$testButton = GUICtrlCreateButton("Test Alias", 370, 300, 80, 25, $WS_GROUP)
GUICtrlSetFont($testButton, 8, 400, 0, "Segoe UI")
GUICtrlSetOnEvent($testButton, "testAlias")
$aliasHint = GUICtrlCreateLabel("(Enter a short name like 'np', 'chrome', 'calc', etc.)", 25, 320, 250, 15)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x7F8C8D)
GUICtrlCreateGroup("", -99, -99, 1, 1)

; Context menu section
$Group3 = GUICtrlCreateGroup(" 3. Context Menu Integration ", 10, 350, 500, 60)
GUICtrlSetFont($Group3, 9, 600, 0, "Segoe UI")
$contextLabel = GUICtrlCreateLabel("Right-click context menu status:", 25, 375, 150, 18)
GUICtrlSetFont(-1, 8, 400, 0, "Segoe UI")
$contextStatus = GUICtrlCreateLabel("", 180, 375, 80, 18)
GUICtrlSetFont(-1, 8, 600, 0, "Segoe UI")
$addContextMenu = GUICtrlCreateButton("Add to Context Menu", 270, 372, 110, 25, $WS_GROUP)
GUICtrlSetFont($addContextMenu, 8, 400, 0, "Segoe UI")
GUICtrlSetOnEvent($addContextMenu, "addContextMenu")
$removeContextMenu = GUICtrlCreateButton("Remove from Context Menu", 390, 372, 110, 25, $WS_GROUP)
GUICtrlSetFont($removeContextMenu, 8, 400, 0, "Segoe UI")
GUICtrlSetOnEvent($removeContextMenu, "removeContextMenu")
GUICtrlCreateGroup("", -99, -99, 1, 1)

; Status bar
$statusBar = GUICtrlCreateLabel("Ready - Select a program and enter an alias name", 10, 420, 500, 20, $SS_SUNKEN)
GUICtrlSetFont($statusBar, 8, 400, 0, "Segoe UI")
GUICtrlSetBkColor($statusBar, 0xECF0F1)

GUISetState(@SW_SHOW)

; Update context menu status on startup
UpdateContextMenuStatus()
#EndRegion ### END Koda GUI section ###

; Check for command line parameters (context menu usage)
If $CmdLine[0] > 0 Then
    ProcessContextMenuCall()
EndIf

While 1
    Sleep(1000)
WEnd

Func closew()
    ; Clean up
    DllCall("shell32.dll", "none", "DragAcceptFiles", "hwnd", $main, "int", 0)
    DllClose($hDLL)
    Exit
EndFunc

Func slectf()
    $filen = FileOpenDialog("Select the program to create an alias for", @ProgramFilesDir, "Executable Files (*.exe)|All Files (*.*)", 1 + 2)
    If Not @error Then
        GUICtrlSetData($filename, $filen)
        ; Auto-generate alias from filename
        $basename = StringRegExpReplace($filen, ".*\\(.*)\..*", "$1")
        GUICtrlSetData($alias, StringLower($basename))
        UpdateStatusBar("Program selected: " & $basename)
    EndIf
EndFunc

Func add()
    ; Check if parameters are filled
    $filen = GUICtrlRead($filename)
    $a = GUICtrlRead($alias)
    
    If $filen == "" Then
        MsgBox(48, "Missing Information", "Please select a program first!" & @CRLF & @CRLF & "Use the Browse button or drag & drop a file.")
        Return
    ElseIf $a == "" Then
        MsgBox(48, "Missing Information", "Please enter an alias name!" & @CRLF & @CRLF & "Example: 'np' for Notepad, 'chrome' for Chrome")
        Return
    EndIf
    
    ; Check if file exists
    If Not FileExists($filen) Then
        MsgBox(16, "File Not Found", "The selected file does not exist!" & @CRLF & @CRLF & "Path: " & $filen)
        Return
    EndIf
    
    ; Check if alias already exists
    RegEnumKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\app.ico Paths\" & $a & ".exe", 1)
    If @error <> 1 Then
        $result = MsgBox(4 + 32, "Alias Already Exists", "The alias '" & $a & "' already exists." & @CRLF & @CRLF & "Do you want to overwrite it?")
        If $result = 7 Then Return ; No selected
    EndIf
    
    ; Add registry entry
    RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\app.ico Paths\" & $a & ".exe", "", "REG_SZ", $filen)
    If @error Then
        MsgBox(16, "Access Denied", "Failed to add alias. Please run as Administrator." & @CRLF & @CRLF & "Right-click the program and select 'Run as Administrator'")
        UpdateStatusBar("Error: Administrator privileges required")
    Else
        MsgBox(64, "Success", "Alias '" & $a & "' created successfully!" & @CRLF & @CRLF & "You can now run the program by typing '" & $a & "' in Start > Run")
        ; Clear fields after successful addition
        GUICtrlSetData($filename, "")
        GUICtrlSetData($alias, "")
        UpdateStatusBar("Alias '" & $a & "' added successfully")
    EndIf
EndFunc

Func remove()
    $a = GUICtrlRead($alias)
    If $a == "" Then
        MsgBox(48, "Missing Information", "Please enter the alias name to remove!")
        Return
    EndIf
    
    RegEnumKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\app.ico Paths\" & $a & ".exe", 1)
    If @error == 1 Then
        MsgBox(48, "Alias Not Found", "The alias '" & $a & "' does not exist!")
        UpdateStatusBar("Alias '" & $a & "' not found")
    Else
        $result = MsgBox(4 + 32, "Confirm Removal", "Are you sure you want to remove the alias '" & $a & "'?")
        If $result = 6 Then ; Yes selected
            RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\app.ico Paths\" & $a & ".exe")
            If @error Then
                MsgBox(16, "Access Denied", "Failed to remove alias. Please run as Administrator.")
                UpdateStatusBar("Error: Administrator privileges required")
            Else
                MsgBox(64, "Success", "Alias '" & $a & "' removed successfully!")
                GUICtrlSetData($alias, "")
                UpdateStatusBar("Alias '" & $a & "' removed successfully")
            EndIf
        EndIf
    EndIf
EndFunc

Func testAlias()
    $a = GUICtrlRead($alias)
    If $a == "" Then
        MsgBox(48, "Missing Information", "Please enter an alias name to test!")
        Return
    EndIf
    
    ; Check if alias exists
    Local $aliasPath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\app.ico Paths\" & $a & ".exe", "")
    If @error Then
        MsgBox(48, "Alias Not Found", "The alias '" & $a & "' does not exist in the registry.")
        UpdateStatusBar("Alias '" & $a & "' not found")
    Else
        If FileExists($aliasPath) Then
            MsgBox(64, "Alias Test", "Alias '" & $a & "' is working correctly!" & @CRLF & @CRLF & "Points to: " & $aliasPath)
            UpdateStatusBar("Alias '" & $a & "' is working correctly")
        Else
            MsgBox(48, "Broken Alias", "Alias '" & $a & "' exists but points to a missing file:" & @CRLF & @CRLF & $aliasPath)
            UpdateStatusBar("Alias '" & $a & "' points to missing file")
        EndIf
    EndIf
EndFunc

Func UpdateStatusBar($message)
    GUICtrlSetData($statusBar, $message)
EndFunc

; Drag and drop message handler
Func WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
    Local $nSize, $pFileName
    Local $nAmt = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", 0xFFFFFFFF, "ptr", 0, "int", 255)
    
    If $nAmt[0] > 0 Then
        $nSize = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", 0, "ptr", 0, "int", 0)
        $nSize = $nSize[0] + 1
        $pFileName = DllStructCreate("char[" & $nSize & "]")
        DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", 0, "ptr", DllStructGetPtr($pFileName), "int", $nSize)
        Local $sFileName = DllStructGetData($pFileName, 1)
        
        ; Set the dropped file in the filename field
        GUICtrlSetData($filename, $sFileName)
        
        ; Auto-generate alias from filename
        Local $basename = StringRegExpReplace($sFileName, ".*\\(.*)\..*", "$1")
        GUICtrlSetData($alias, StringLower($basename))
        
        UpdateStatusBar("File dropped: " & $basename)
    EndIf
    
    DllCall("shell32.dll", "none", "DragFinish", "hwnd", $wParam)
    Return 0
EndFunc

; Get the icon path for context menu
Func GetIconPath()
    Local $iconPath = ""
    
    If @Compiled Then
        ; When compiled, use the executable itself as the icon source
        $iconPath = @ScriptFullPath & ",0"
    Else
        ; When not compiled, try to find the icon file in the script directory
        Local $scriptDir = @ScriptDir
        Local $iconFile = $scriptDir & "\app.ico"
        
        If FileExists($iconFile) Then
            $iconPath = $iconFile
        Else
            ; Try other common icon names
            Local $iconLocations[] = [ _
                @ScriptDir & "\icon.ico", _
                @ScriptDir & "\app.ico", _
                @ScriptDir & "\AddToRun.ico", _
                @ScriptDir & "\Add2Run.ico" _
            ]
            
            For $i = 0 To UBound($iconLocations) - 1
                If FileExists($iconLocations[$i]) Then
                    $iconPath = $iconLocations[$i]
                    ExitLoop
                EndIf
            Next
            
            ; Ultimate fallback to a system icon (different index)
            If $iconPath = "" Then
                $iconPath = @SystemDir & "\shell32.dll,-131"
            EndIf
        EndIf
    EndIf
    
    Return $iconPath
EndFunc

; Create context menu registry entry
Func CreateContextMenu()
    Local $exePath = @ScriptFullPath
    Local $iconPath = GetIconPath()
    
    ; Build the command line properly
    If @Compiled Then
        $exePath = '"' & @ScriptFullPath & '"'
    Else
        $exePath = '"' & @AutoItExe & '" "' & @ScriptFullPath & '"'
    EndIf
    
    ; Debug: Show what icon path we're using (uncomment to debug)
    ; MsgBox(0, "Debug", "Icon path: " & $iconPath)
    
    ; Create context menu entry for executable files
    RegWrite("HKEY_CLASSES_ROOT\exefile\shell\AddToRun", "", "REG_SZ", "Create/Set Alias with AddToRun")
    RegWrite("HKEY_CLASSES_ROOT\exefile\shell\AddToRun", "Icon", "REG_SZ", $iconPath)
    RegWrite("HKEY_CLASSES_ROOT\exefile\shell\AddToRun\command", "", "REG_SZ", $exePath & ' "%1"')
    
    ; Create context menu entry for all files (optional)
    RegWrite("HKEY_CLASSES_ROOT\*\shell\AddToRun", "", "REG_SZ", "Create/Set Alias with AddToRun")
    RegWrite("HKEY_CLASSES_ROOT\*\shell\AddToRun", "Icon", "REG_SZ", $iconPath)
    RegWrite("HKEY_CLASSES_ROOT\*\shell\AddToRun\command", "", "REG_SZ", $exePath & ' "%1"')
EndFunc

; Add context menu function
Func addContextMenu()
    CreateContextMenu()
    If @error Then
        MsgBox(16, "Error", "Failed to add context menu. Please run as Administrator.")
        UpdateStatusBar("Error: Administrator privileges required for context menu")
    Else
        MsgBox(64, "Success", "Context menu added successfully!" & @CRLF & @CRLF & "Right-click any executable file and select 'Create/Set Alias with AddToRun'")
        UpdateContextMenuStatus()
        UpdateStatusBar("Context menu added successfully")
    EndIf
EndFunc

; Remove context menu function
Func removeContextMenu()
    $result = MsgBox(4 + 32, "Confirm Removal", "Are you sure you want to remove the context menu integration?")
    If $result = 6 Then ; Yes selected
        CleanupContextMenu()
        If @error Then
            MsgBox(16, "Error", "Failed to remove context menu. Please run as Administrator.")
            UpdateStatusBar("Error: Administrator privileges required")
        Else
            MsgBox(64, "Success", "Context menu removed successfully!")
            UpdateContextMenuStatus()
            UpdateStatusBar("Context menu removed successfully")
        EndIf
    EndIf
EndFunc

; Update context menu status
Func UpdateContextMenuStatus()
    ; Check if context menu exists
    Local $keyExists = RegRead("HKEY_CLASSES_ROOT\exefile\shell\AddToRun", "")
    If @error Then
        GUICtrlSetData($contextStatus, "Not Added")
        GUICtrlSetColor($contextStatus, 0xE74C3C)
    Else
        GUICtrlSetData($contextStatus, "Added")
        GUICtrlSetColor($contextStatus, 0x27AE60)
    EndIf
EndFunc

; Process context menu call
Func ProcessContextMenuCall()
    If $CmdLine[0] > 0 Then
        Local $contextFile = $CmdLine[1]
        If FileExists($contextFile) Then
            GUICtrlSetData($filename, $contextFile)
            ; Auto-generate alias from filename
            Local $basename = StringRegExpReplace($contextFile, ".*\\(.*)\..*", "$1")
            GUICtrlSetData($alias, StringLower($basename))
            UpdateStatusBar("Loaded from context menu: " & $basename)
        EndIf
    EndIf
EndFunc

; Clean up context menu
Func CleanupContextMenu()
    RegDelete("HKEY_CLASSES_ROOT\exefile\shell\AddToRun")
    RegDelete("HKEY_CLASSES_ROOT\*\shell\AddToRun")
EndFunc