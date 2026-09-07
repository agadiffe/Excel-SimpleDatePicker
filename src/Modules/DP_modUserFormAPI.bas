Attribute VB_Name = "DP_modUserFormAPI"
Option Explicit

'----------------------------------------
' Windows API declarations
'----------------------------------------
'
' Windows only.
'
' VBA6:
'   Uses Long for window handles and styles.
'
' VBA7:
'   Uses LongPtr for 32-bit/64-bit compatibility.
'
' Mac:
'   Windows API declarations are not compiled.

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

Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal FormWindowHandle As LongPtr, _
    ByVal InsertAfter As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal Flags As Long) As Long

#ElseIf Not Mac Then

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal WindowClassName As String, _
    ByVal WindowCaption As String) As Long

Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
    ByVal FormWindowHandle As Long, _
    ByVal WindowStyleIndex As Long) As Long

Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
    ByVal FormWindowHandle As Long, _
    ByVal WindowStyleIndex As Long, _
    ByVal WindowStyle As Long) As Long

Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal FormWindowHandle As Long, _
    ByVal InsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal Flags As Long) As Long

#End If


'----------------------------------------
' Window styles
'----------------------------------------

#If Not Mac Then

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20

Private Const WS_CAPTION As Long = &HC00000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_SYSMENU As Long = &H80000

Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_CLIENTEDGE As Long = &H200

Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

#End If


'----------------------------------------
' Remove UserForm title bar and frame
'----------------------------------------

Public Sub DP_RemoveUserFormTitleBar(ByVal FormCaption As String)

#If Mac Then

    Exit Sub

#ElseIf VBA7 Then

    Dim FormWindowHandle As LongPtr
    Dim WindowStyle As LongPtr
    Dim ExtendedStyle As LongPtr

#Else

    Dim FormWindowHandle As Long
    Dim WindowStyle As Long
    Dim ExtendedStyle As Long

#End If


#If Not Mac Then

    FormWindowHandle = FindWindow("ThunderDFrame", FormCaption)

    If FormWindowHandle = 0 Then Exit Sub

    '----------------------------------------
    ' Remove standard window frame
    '----------------------------------------

    WindowStyle = GetWindowLongPtr( _
                    FormWindowHandle, _
                    GWL_STYLE)

    WindowStyle = WindowStyle And _
                  Not (WS_CAPTION Or _
                       WS_DLGFRAME Or _
                       WS_THICKFRAME Or _
                       WS_MINIMIZEBOX Or _
                       WS_MAXIMIZEBOX Or _
                       WS_SYSMENU)

    SetWindowLongPtr _
        FormWindowHandle, _
        GWL_STYLE, _
        WindowStyle

    '----------------------------------------
    ' Remove extended MSForms frame
    '----------------------------------------

    ExtendedStyle = GetWindowLongPtr( _
                        FormWindowHandle, _
                        GWL_EXSTYLE)

    ExtendedStyle = ExtendedStyle And _
                    Not (WS_EX_DLGMODALFRAME Or _
                         WS_EX_CLIENTEDGE)

    SetWindowLongPtr _
        FormWindowHandle, _
        GWL_EXSTYLE, _
        ExtendedStyle

    '----------------------------------------
    ' Force Windows to recalculate frame
    '----------------------------------------

    SetWindowPos _
        FormWindowHandle, _
        0, _
        0, 0, 0, 0, _
        SWP_NOMOVE Or _
        SWP_NOSIZE Or _
        SWP_NOZORDER Or _
        SWP_FRAMECHANGED

#End If

End Sub

