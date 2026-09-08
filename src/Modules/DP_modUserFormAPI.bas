Attribute VB_Name = "DP_modUserFormAPI"
Option Explicit

'============================================================
' Rounded UserForm / VB6 / VBA6 / VBA7
'============================================================

'============================================================
' Windows API declarations
'============================================================

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

Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal FormWindowHandle As LongPtr, _
    ByRef WindowRectangle As RECT) As Long

Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal WidthEllipse As Long, _
    ByVal HeightEllipse As Long) As LongPtr

Private Declare PtrSafe Function SetWindowRgn Lib "user32" ( _
    ByVal FormWindowHandle As LongPtr, _
    ByVal RegionHandle As LongPtr, _
    ByVal Redraw As Boolean) As Long

'============================================================
' DWM
'============================================================

Private Declare PtrSafe Function DwmSetWindowAttribute Lib "dwmapi" ( _
    ByVal FormWindowHandle As LongPtr, _
    ByVal DwmAttribute As Long, _
    ByRef DwmValue As Long, _
    ByVal DwmValueSize As Long) As Long

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

Private Declare Function GetWindowRect Lib "user32" ( _
    ByVal FormWindowHandle As Long, _
    ByRef WindowRectangle As RECT) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal WidthEllipse As Long, _
    ByVal HeightEllipse As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal FormWindowHandle As Long, _
    ByVal RegionHandle As Long, _
    ByVal Redraw As Boolean) As Long

'============================================================
' DWM
'============================================================

Private Declare Function DwmSetWindowAttribute Lib "dwmapi" ( _
    ByVal FormWindowHandle As Long, _
    ByVal DwmAttribute As Long, _
    ByRef DwmValue As Long, _
    ByVal DwmValueSize As Long) As Long

#End If


'============================================================
' Structures
'============================================================

#If Not Mac Then

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#End If


'============================================================
' Window styles
'============================================================

#If Not Mac Then

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20

'------------------------------------------------------------
' Standard window styles
'------------------------------------------------------------

Private Const WS_BORDER As Long = &H800000

Private Const WS_CAPTION As Long = &HC00000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_SYSMENU As Long = &H80000

'------------------------------------------------------------
' Extended window styles
'------------------------------------------------------------

Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_CLIENTEDGE As Long = &H200

'------------------------------------------------------------
' SetWindowPos flags
'------------------------------------------------------------

Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

'------------------------------------------------------------
' DWM attributes
'------------------------------------------------------------

Private Const DWMWA_WINDOW_CORNER_PREFERENCE As Long = 33
Private Const DWMWA_BORDER_COLOR As Long = 34

'------------------------------------------------------------
' DWM corner preferences
'------------------------------------------------------------

Private Const DWMWCP_DEFAULT As Long = 0
Private Const DWMWCP_DONOTROUND As Long = 1
Private Const DWMWCP_ROUND As Long = 2
Private Const DWMWCP_ROUNDSMALL As Long = 3

'------------------------------------------------------------
' Rounded corner radius
'
' Used only by the CreateRoundRectRgn fallback.
'
' DWM uses its own Windows 11 corner geometry.
'------------------------------------------------------------

Private Const DP_CORNER_RADIUS As Long = 12

#End If


'============================================================
' Remove UserForm title bar and frame
'
' Removes:
'
'   - title bar
'   - dialog frame
'   - resize frame
'   - minimize button
'   - maximize button
'   - system menu
'   - MSForms client edge
'
' Keeps:
'
'   - simple WS_BORDER
'
' This avoids the double-border problem caused by
' WS_EX_CLIENTEDGE.
'============================================================

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

    '========================================================
    ' Find the UserForm window
    '========================================================

    FormWindowHandle = FindWindow( _
        "ThunderDFrame", _
        FormCaption)

    If FormWindowHandle = 0 Then Exit Sub

    '========================================================
    ' Get current standard window style
    '========================================================

    WindowStyle = GetWindowLongPtr( _
        FormWindowHandle, _
        GWL_STYLE)

    '========================================================
    ' Remove all standard frame elements
    '========================================================

    WindowStyle = WindowStyle And _
                  Not (WS_CAPTION Or _
                       WS_DLGFRAME Or _
                       WS_THICKFRAME Or _
                       WS_MINIMIZEBOX Or _
                       WS_MAXIMIZEBOX Or _
                       WS_SYSMENU)

    '========================================================
    ' Keep a simple 1px border
    '========================================================

    WindowStyle = WindowStyle Or WS_BORDER

    SetWindowLongPtr _
        FormWindowHandle, _
        GWL_STYLE, _
        WindowStyle

    '========================================================
    ' Remove MSForms extended frame
    '========================================================

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

    '========================================================
    ' Tell Windows that the frame changed
    '========================================================

    SetWindowPos _
        FormWindowHandle, _
        0, _
        0, _
        0, _
        0, _
        0, _
        SWP_NOMOVE Or _
        SWP_NOSIZE Or _
        SWP_NOZORDER Or _
        SWP_FRAMECHANGED

#End If

End Sub


'============================================================
' Apply rounded corners
'
' Windows 11:
'
'   DWM handles the rounded corners.
'   This is preferred because the DWM border then follows
'   the rounded corners.
'
' Older Windows:
'
'   Falls back to CreateRoundRectRgn.
'============================================================

Public Sub DP_ApplyRoundedCorners(ByVal PickerForm As Object)

#If Mac Then

    Exit Sub

#ElseIf VBA7 Then

    Dim FormWindowHandle As LongPtr
    Dim RegionHandle As LongPtr

#Else

    Dim FormWindowHandle As Long
    Dim RegionHandle As Long

#End If

#If Not Mac Then

    Dim CornerPreference As Long

    '========================================================
    ' Find actual UserForm window
    '========================================================

    FormWindowHandle = FindWindow( _
        "ThunderDFrame", _
        PickerForm.Caption)

    If FormWindowHandle = 0 Then Exit Sub

    '========================================================
    ' Try DWM rounded corners first.
    '
    ' Windows 11 supports this attribute.
    '
    ' If the call succeeds, do NOT apply SetWindowRgn.
    '
    ' This is important because SetWindowRgn would clip the
    ' DWM border away from the rounded corners.
    '========================================================

    CornerPreference = DWMWCP_ROUND

    If DwmSetWindowAttribute( _
        FormWindowHandle, _
        DWMWA_WINDOW_CORNER_PREFERENCE, _
        CornerPreference, _
        4) = 0 Then

        Exit Sub

    End If

    '========================================================
    ' Older Windows fallback
    '========================================================

    Dim WindowRectangle As RECT
    Dim WindowWidth As Long
    Dim WindowHeight As Long

    If GetWindowRect( _
        FormWindowHandle, _
        WindowRectangle) = 0 Then Exit Sub

    WindowWidth = _
        WindowRectangle.Right - WindowRectangle.Left

    WindowHeight = _
        WindowRectangle.Bottom - WindowRectangle.Top

    If WindowWidth <= 0 Then Exit Sub
    If WindowHeight <= 0 Then Exit Sub

    '========================================================
    ' Create rounded region
    '========================================================

    RegionHandle = CreateRoundRectRgn( _
        0, _
        0, _
        WindowWidth + 1, _
        WindowHeight + 1, _
        DP_CORNER_RADIUS * 2, _
        DP_CORNER_RADIUS * 2)

    If RegionHandle = 0 Then Exit Sub

    '========================================================
    ' Apply rounded region
    '========================================================

    SetWindowRgn _
        FormWindowHandle, _
        RegionHandle, _
        True

#End If

End Sub


'============================================================
' Apply border color
'
' The UserForm itself remains DP_ColorBg().
'
' The border uses DP_ColorBorder().
'
' Windows 11:
'   DWM border color is used.
'
' Older Windows:
'   WS_BORDER remains as the fallback.
'============================================================

Public Sub DP_ApplyBorderColor(ByVal PickerForm As Object)

#If Mac Then

    Exit Sub

#ElseIf VBA7 Then

    Dim FormWindowHandle As LongPtr

#Else

    Dim FormWindowHandle As Long

#End If

#If Not Mac Then

    Dim BorderColor As Long

    '========================================================
    ' Find actual UserForm window
    '========================================================

    FormWindowHandle = FindWindow( _
        "ThunderDFrame", _
        PickerForm.Caption)

    If FormWindowHandle = 0 Then Exit Sub

    '========================================================
    ' Get project border color
    '========================================================

    BorderColor = DP_ColorBorder()

    '========================================================
    ' Tell DWM to use the project border color.
    '
    ' On Windows versions without this DWM attribute,
    ' the call simply fails and WS_BORDER remains active.
    '========================================================

    DwmSetWindowAttribute _
        FormWindowHandle, _
        DWMWA_BORDER_COLOR, _
        BorderColor, _
        4

#End If

End Sub

