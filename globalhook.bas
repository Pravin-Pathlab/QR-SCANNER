Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const WH_KEYBOARD_LL = 13
Public Const WM_KEYDOWN = &H100
Public hHook As Long
Private sBuf As String ' Internal scan buffer

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Function LowLevelKeyboardProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim kbd As KBDLLHOOKSTRUCT
    If ncode = 0 And wParam = WM_KEYDOWN Then
        CopyMemory kbd, ByVal lParam, Len(kbd)
        
        ' If it's a bracket or brace (start of JSON), start collecting [cite: 32, 34]
        If kbd.vkCode = 219 Or sBuf <> "" Then
            If kbd.vkCode = 13 Then ' ENTER key = Scan Done
                Form2.ProcessGlobalScan sBuf
                sBuf = ""
                LowLevelKeyboardProc = 1 ' Swallow the Enter key
                Exit Function
            Else
                ' Translate and add to buffer [cite: 33, 34]
                sBuf = sBuf & GetScannerChar(kbd.vkCode)
                LowLevelKeyboardProc = 1 ' SWALLOW: Don't let it type in textbox!
                Exit Function
            End If
        End If
    End If
    LowLevelKeyboardProc = CallNextHookEx(hHook, ncode, wParam, ByVal lParam)
End Function

' Simplified translator specifically for JSON QR symbols [cite: 34, 35]
Private Function GetScannerChar(vk As Long) As String
    Select Case vk
        Case 48 To 57: GetScannerChar = Chr(vk)
        Case 65 To 90: GetScannerChar = LCase(Chr(vk))
        Case 186: GetScannerChar = ":"
        Case 188: GetScannerChar = ","
        Case 219: GetScannerChar = "{"
        Case 221: GetScannerChar = "}"
        Case 222: GetScannerChar = """"
        Case 32:  GetScannerChar = " "
        Case 190: GetScannerChar = "."
    End Select
End Function
