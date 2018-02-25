Attribute VB_Name = "Lock"
Option Explicit


Private Declare Function CallNextHookEx Lib "user32" _
(ByVal hHook As Long, ByVal nCode As Long, ByVal wParam _
As Long, lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32" _
Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal _
lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) _
As Long
Private Declare Function UnhookWindowsHookEx Lib _
"user32" (ByVal hHook As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" (Destination As Any, Source As _
Any, ByVal Length As Long)
Private Type PKBDLLHOOKSTRUCT

    vkCode As Long

    scanCode As Long

    flags As Long

    time As Long

    dwExtraInfo As Long

End Type
Private Const WM_KEYDOWN = &H100
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYUP = &H105
Private Const VK_LWIN = &H5B
Private Const VK_RWIN = &H5C
Private Const HC_ACTION = 0
Private Const WH_KEYBOARD_LL = 13
Private lngHook As Long


'使用底层KeyboardHook拦截按键消息

Public Function LowLevelKeyboardProc(ByVal nCode _
As Long, ByVal wParam As Long, ByVal lParam As Long) _
As Long

    Dim blnHook As Boolean

    Dim p As PKBDLLHOOKSTRUCT

    

    If nCode = HC_ACTION Then

If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN _
 Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            CopyMemory p, ByVal lParam, Len(p)
  
blnHook = ((p.vkCode = VK_ESCAPE) And _
((GetKeyState(VK_CONTROL) And &H8000) <> 0)) Or _
((p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
(p.vkCode = VK_LWIN) Or (p.vkCode = VK_RWIN) Or _
((p.vkCode = vbKeyF4) And _
((GetKeyState(vbKeyMenu) And &H8000) <> 0))


'((p.vkCode = VK_ESCAPE) And  _((GetKeyState(VK_CONTROL) _
And &H8000) <> 0))  这个是屏蔽ESC+Ctrl
            
'((p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> _
0))  这个是屏蔽Tab+Alt
            
'(p.vkCode = VK_LWIN) or (p.vkCode = VK_RWIN)   这个是屏蔽左右win
            
 '((p.vkCode = vbKeyF4) And ((GetKeyState(vbKeyMenu) _
And &H8000) <> 0))  这个是屏蔽Alt+F4
            
        End If


    End If

    

    If blnHook Then

        LowLevelKeyboardProc = 1

    Else

        Call CallNextHookEx(WH_KEYBOARD_LL, nCode, wParam, lParam)

    End If

End Function


Public Sub HooK()

    lngHook = SetWindowsHookEx(WH_KEYBOARD_LL, _
    AddressOf LowLevelKeyboardProc, App.hInstance, 0)

End Sub


Public Sub UnHooK()

    Call UnhookWindowsHookEx(lngHook)

End Sub

