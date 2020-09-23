Attribute VB_Name = "modaero"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
Private Const HWND_TOPMOST = -1

Option Explicit
Dim bTrans As Byte ' The level of transparency (0 - 255)
Dim lOldStyle As Long

Function SetFormOpacity(HostForm As Form, OpacityValue As Byte)

lOldStyle = SetWindowLong(HostForm.hwnd, GWL_EXSTYLE, WS_EX_LAYERED)

bTrans = OpacityValue + 5

If bTrans < 50 Then
Exit Function
End If

If bTrans >= 255 Then
Exit Function
End If

SetLayeredWindowAttributes HostForm.hwnd, 0, bTrans, LWA_ALPHA

End Function

Function CloseFormButton(HostForm As Form, ButtonState As Boolean)
    If ButtonState = False Then
    Dim hMenu As Long
    hMenu = GetSystemMenu(HostForm.hwnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION
    End If
End Function

Function MakeFormOnTop(HostForm As Form, FormState As Boolean)
    If FormState = True Then
    SetWindowPos HostForm.hwnd, HWND_TOPMOST, HostForm.Width, HostForm.Height, 320, 265, 128
    End If
End Function

