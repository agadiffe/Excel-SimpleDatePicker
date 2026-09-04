Attribute VB_Name = "DP_modUserFormAPI"
Option Explicit

'----------------------------------------
' Windows API declarations
'----------------------------------------
'
' The API is only implemented for VBA7 on Windows.
'
' VBA6:
'   The module still compiles, but title-bar removal is not available.
'
' VBA7:
'   The Windows API is declared using LongPtr for 32-bit/64-bit compatibility.
'
' Mac:
'   The Windows API is not compiled.

#If VBA7 And Not Mac Then

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal WindowClassName As String, _
    ByVal WindowCaption As String) As LongPtr

Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
    ByVal FormWindowHandle As LongPtr, _
    ByVal WindowStyleIndex As Long) As LongPtr

Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
    ByVal FormWindowHandle As LongPtr, _
    ByVal WindowStyleIndex As Long, _
    ByVal WindowStyle As LongPtr) As LongPtr

Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal FormWindowHandle As LongPtr) As Long

Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000

#End If


'----------------------------------------
' Remove UserForm title bar
'----------------------------------------

Public Sub RemoveUserFormTitleBar(ByVal FormCaption As String)

#If VBA7 And Not Mac Then

    Dim FormWindowHandle As LongPtr
    Dim WindowStyle As LongPtr

    FormWindowHandle = FindWindow("ThunderDFrame", FormCaption)

    If FormWindowHandle = 0 Then Exit Sub

    WindowStyle = GetWindowLongPtr(FormWindowHandle, GWL_STYLE)
    WindowStyle = WindowStyle And Not WS_CAPTION

    SetWindowLongPtr FormWindowHandle, GWL_STYLE, WindowStyle

    DrawMenuBar FormWindowHandle

#End If

End Sub

