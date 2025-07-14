#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=Add2Run_32512.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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

#Region ### START Koda GUI section ### Form=c:\autoit3\koda\add2run.kxf
$main = GUICreate("AddToRun v1.1", 381, 305, -1, -1)
GUISetBkColor(0xFFFBF0)
GUISetOnEvent($GUI_EVENT_CLOSE, "closew")

; Enable drag and drop for the main window
DllCall("shell32.dll", "none", "DragAcceptFiles", "hwnd", $main, "int", 1)

; Register for WM_DROPFILES message
GUIRegisterMsg($WM_DROPFILES, "WM_DROPFILES")

$Group1 = GUICtrlCreateGroup(" Select The Program ", 16, 99, 353, 65)
$filename = GUICtrlCreateEdit("", 32, 123, 233, 25, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY))
$slect = GUICtrlCreateButton("Select", 280, 123, 75, 25, $WS_GROUP)
GUICtrlSetOnEvent($slect, "slectf")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group2 = GUICtrlCreateGroup(" Input The Alias ", 16, 178, 353, 65)
$alias = GUICtrlCreateEdit("", 64, 202, 105, 25, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL))
$add = GUICtrlCreateButton("Add", 192, 202, 75, 25, $WS_GROUP)
GUICtrlSetOnEvent($add, "add")
$Label1 = GUICtrlCreateLabel("Alias", 28, 207, 28, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group3 = GUICtrlCreateGroup(" Context Menu ", 16, 250, 353, 45)
$addContextMenu = GUICtrlCreateButton("Add to Context Menu", 32, 268, 120, 20, $WS_GROUP)
GUICtrlSetOnEvent($addContextMenu, "addContextMenu")
$removeContextMenu = GUICtrlCreateButton("Remove from Context Menu", 160, 268, 120, 20, $WS_GROUP)
GUICtrlSetOnEvent($removeContextMenu, "removeContextMenu")
$contextStatus = GUICtrlCreateLabel("", 290, 271, 70, 17)
GUICtrlSetFont(-1, 7, 400, 0, "MS Sans Serif")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Icon1 = GUICtrlCreateIcon("C:\WINDOWS\system32\shell32.dll", -131, 21, 27, 48, 48, BitOR($SS_NOTIFY, $WS_GROUP))
$Label2 = GUICtrlCreateLabel("AddToRun", 89, 6, 178, 49)
GUICtrlSetFont(-1, 24, 800, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x800000)
$Label3 = GUICtrlCreateLabel("AddToRun  v1.1", 89, 52, 100, 20)
GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
$Label4 = GUICtrlCreateLabel("Copyright 2009 pctip#163.com", 89, 68, 212, 20)
GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")

; Remove function
$remove1 = GUICtrlCreateButton("Remove", 280, 202, 75, 25, $WS_GROUP)
GUICtrlSetOnEvent($remove1, "remove")
GUICtrlSetFont(-1, 7, 400, 0, "MS Sans Serif")

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
    $filen = FileOpenDialog("Select the program which need to set up an alias", @ProgramFilesDir, "All TYPE(*.*)", 1 + 2)
    If Not @error Then
        GUICtrlSetData($filename, $filen)
        ; Auto-generate alias from filename
        $basename = StringRegExpReplace($filen, ".*\\(.*)\..*", "$1")
        GUICtrlSetData($alias, $basename)
    EndIf
EndFunc

Func add()
    ; Check if parameters are filled
    $filen = GUICtrlRead($filename)
    $a = GUICtrlRead($alias)
    
    If $filen == "" Then
        MsgBox(0, "Error", "Please select a program first!")
        Return
    ElseIf $a == "" Then
        MsgBox(0, "Error", "Please input an alias!")
        Return
    EndIf
    
    ; Check if file exists
    If Not FileExists($filen) Then
        MsgBox(0, "Error", "Selected file does not exist!")
        Return
    EndIf
    
    ; Check if alias already exists
    RegEnumKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & $a & ".exe", 1)
    If @error <> 1 Then
        $result = MsgBox(4, "Alias Exists", "This alias already exists. Do you want to overwrite it?")
        If $result = 7 Then Return ; No selected
    EndIf
    
    ; Add registry entry
    RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & $a & ".exe", "", "REG_SZ", $filen)
    If @error Then
        MsgBox(0, "Error", "Failed to add alias. Make sure you're running as administrator.")
    Else
        ; Clear fields after successful addition
        GUICtrlSetData($filename, "")
        GUICtrlSetData($alias, "")
    EndIf
EndFunc

Func remove()
    $a = GUICtrlRead($alias)
    If $a == "" Then
        MsgBox(0, "Error", "Please input an alias to remove!")
        Return
    EndIf
    
    RegEnumKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & $a & ".exe", 1)
    If @error == 1 Then
        MsgBox(0, "Error", "This alias does not exist!")
    Else
        RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & $a & ".exe")
        If @error Then
            MsgBox(0, "Error", "Failed to remove alias. Make sure you're running as administrator.")
        Else
            GUICtrlSetData($alias, "")
        EndIf
    EndIf
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
        GUICtrlSetData($alias, $basename)
    EndIf
    
    DllCall("shell32.dll", "none", "DragFinish", "hwnd", $wParam)
    Return 0
EndFunc

; Create context menu registry entry
Func CreateContextMenu()
    Local $exePath = @ScriptFullPath
    If @Compiled Then
        $exePath = @ScriptFullPath
    Else
        $exePath = @AutoItExe & ' "' & @ScriptFullPath & '"'
    EndIf
    
    ; Create context menu entry for executable files
    RegWrite("HKEY_CLASSES_ROOT\exefile\shell\AddToRun", "", "REG_SZ", "Set Alias")
    RegWrite("HKEY_CLASSES_ROOT\exefile\shell\AddToRun\command", "", "REG_SZ", $exePath & ' "%1"')
    
    ; Create context menu entry for all files (optional)
    RegWrite("HKEY_CLASSES_ROOT\*\shell\AddToRun", "", "REG_SZ", "Set Alias")
    RegWrite("HKEY_CLASSES_ROOT\*\shell\AddToRun\command", "", "REG_SZ", $exePath & ' "%1"')
EndFunc

; Add context menu function
Func addContextMenu()
    CreateContextMenu()
    If @error Then
        MsgBox(0, "Error", "Failed to add context menu. Make sure you're running as administrator.")
    Else
        UpdateContextMenuStatus()
    EndIf
EndFunc

; Remove context menu function
Func removeContextMenu()
    CleanupContextMenu()
    If @error Then
        MsgBox(0, "Error", "Failed to remove context menu. Make sure you're running as administrator.")
    Else
        UpdateContextMenuStatus()
    EndIf
EndFunc

; Update context menu status
Func UpdateContextMenuStatus()
    ; Check if context menu exists
    Local $keyExists = RegRead("HKEY_CLASSES_ROOT\exefile\shell\AddToRun", "")
    If @error Then
        GUICtrlSetData($contextStatus, "Not Added")
        GUICtrlSetColor($contextStatus, 0xFF0000)
    Else
        GUICtrlSetData($contextStatus, "Added")
        GUICtrlSetColor($contextStatus, 0x008000)
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
            GUICtrlSetData($alias, $basename)
        EndIf
    EndIf
EndFunc

; Clean up context menu
Func CleanupContextMenu()
    RegDelete("HKEY_CLASSES_ROOT\exefile\shell\AddToRun")
    RegDelete("HKEY_CLASSES_ROOT\*\shell\AddToRun")
EndFunc