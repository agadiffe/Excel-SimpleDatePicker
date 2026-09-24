Attribute VB_Name = "DP_modAPIRoundedForm"
Option Explicit

'============================================================
' Rounded UserForm / VB6 / VBA6 / VBA7
'
' Features:
'
'   - VBA6 / VBA7 compatible
'   - 32-bit / 64-bit Office compatible
'   - Mac-safe compilation
'   - UserForm HWND resolved directly when needed
'   - DWM rounded corners
'   - DWM border color
'   - CreateRoundRectRgn fallback for older Windows
'   - Fallback region recreated when rounded corners are
'     reapplied
'   - Correct HRGN ownership handling
'============================================================


'============================================================
' Windows API declarations
'============================================================

#If VBA7 And Not Mac Then

'------------------------------------------------------------
' Window handling
'------------------------------------------------------------

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

'------------------------------------------------------------
' GDI regions
'------------------------------------------------------------

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

Private Declare PtrSafe Function DeleteObject Lib "gdi32" ( _
    ByVal ObjectHandle As LongPtr) As Long

'------------------------------------------------------------
' DWM
'------------------------------------------------------------

Private Declare PtrSafe Function DwmSetWindowAttribute Lib "dwmapi" ( _
    ByVal FormWindowHandle As LongPtr, _
    ByVal DwmAttribute As Long, _
    ByRef DwmValue As Long, _
    ByVal DwmValueSize As Long) As Long


#ElseIf Not Mac Then

'------------------------------------------------------------
' Window handling
'------------------------------------------------------------

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

'------------------------------------------------------------
' GDI regions
'------------------------------------------------------------

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

Private Declare Function DeleteObject Lib "gdi32" ( _
    ByVal ObjectHandle As Long) As Long

'------------------------------------------------------------
' DWM
'------------------------------------------------------------

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
' Fallback rounded corner radius
'
' CreateRoundRectRgn expects ellipse dimensions, so the
' actual value passed is DP_CORNER_RADIUS * 2.
'------------------------------------------------------------

Private Const DP_CORNER_RADIUS As Single = 9.3

#End If


'============================================================
' Get the UserForm HWND
'
' The HWND is resolved directly from the UserForm caption
' whenever an API operation needs it.
'
' No HWND is cached.
'
' This avoids retaining a handle across UserForm destruction
' and recreation.
'============================================================

#If VBA7 And Not Mac Then

Private Function GetPickerHwnd( _
    ByVal PickerForm As Object) As LongPtr

    If PickerForm Is Nothing Then Exit Function

    GetPickerHwnd = FindWindow("ThunderDFrame", PickerForm.Caption)

End Function


#ElseIf Not Mac Then

Private Function GetPickerHwnd( _
    ByVal PickerForm As Object) As Long

    If PickerForm Is Nothing Then Exit Function

    GetPickerHwnd = FindWindow("ThunderDFrame", PickerForm.Caption)

End Function

#End If


'============================================================
' Remove UserForm title bar and frame
'
' Removes:
'   - title bar
'   - dialog frame
'   - resize frame
'   - minimize button
'   - maximize button
'   - system menu
'   - simple WS_BORDER
'   - MSForms client edge
'
' Avoids the double-border caused by WS_EX_CLIENTEDGE.
'============================================================

Public Sub DP_RemoveUserFormTitleBar( _
    ByVal PickerForm As Object)

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

    '--------------------------------------------------------
    ' Get current HWND
    '--------------------------------------------------------

    FormWindowHandle = GetPickerHwnd(PickerForm)

    If FormWindowHandle = 0 Then Exit Sub

    '--------------------------------------------------------
    ' Get current standard window style
    '--------------------------------------------------------

    WindowStyle = GetWindowLongPtr(FormWindowHandle, GWL_STYLE)

    '--------------------------------------------------------
    ' Remove standard frame elements
    '--------------------------------------------------------

    WindowStyle = WindowStyle And _
                  Not (WS_BORDER Or _
                       WS_CAPTION Or _
                       WS_DLGFRAME Or _
                       WS_THICKFRAME Or _
                       WS_MINIMIZEBOX Or _
                       WS_MAXIMIZEBOX Or _
                       WS_SYSMENU)

    SetWindowLongPtr FormWindowHandle, GWL_STYLE, WindowStyle

    '--------------------------------------------------------
    ' Remove MSForms extended frame
    '--------------------------------------------------------

    ExtendedStyle = GetWindowLongPtr(FormWindowHandle, GWL_EXSTYLE)

    ExtendedStyle = ExtendedStyle And _
                    Not (WS_EX_DLGMODALFRAME Or _
                         WS_EX_CLIENTEDGE)

    SetWindowLongPtr FormWindowHandle, GWL_EXSTYLE, ExtendedStyle

    '--------------------------------------------------------
    ' Tell Windows that the non-client frame changed
    '--------------------------------------------------------

    SetWindowPos FormWindowHandle, 0, _
                 0, 0, 0, 0, _
                 SWP_NOMOVE Or _
                   SWP_NOSIZE Or _
                   SWP_NOZORDER Or _
                   SWP_FRAMECHANGED

End Sub


'============================================================
' Apply rounded corners
'
' Windows 11: DWM rounded corners are requested.
' Older Windows: CreateRoundRectRgn is used.
'
' Return value:
'   True: Rounded corners were successfully applied.
'   False: Rounded corners could not be applied.
'
' IMPORTANT:
'   The HRGN is NOT cached.
'   SetWindowRgn transfers ownership of a successfully
'   applied HRGN to Windows.
'============================================================

Public Function DP_ApplyRoundedCorners( _
    ByVal PickerForm As Object, _
    Optional ByVal CornerRadius As Single = DP_CORNER_RADIUS) As Boolean

#If Mac Then

    DP_ApplyRoundedCorners = False
    Exit Function

#ElseIf VBA7 Then

    Dim FormWindowHandle As LongPtr
    Dim RegionHandle As LongPtr

#Else

    Dim FormWindowHandle As Long
    Dim RegionHandle As Long

#End If

    Dim WindowRectangle As RECT
    Dim WindowWidth As Long
    Dim WindowHeight As Long

    '--------------------------------------------------------
    ' Get current HWND
    '--------------------------------------------------------

    FormWindowHandle = GetPickerHwnd(PickerForm)

    If FormWindowHandle = 0 Then Exit Function

    '--------------------------------------------------------
    ' Try DWM.
    '
    ' If DWM accepts the rounded-corner preference, there is
    ' no need to create a GDI region.
    '
    ' It is safe to call again when the UserForm changes
    ' size or view.
    '--------------------------------------------------------

    If DwmSetWindowAttribute(FormWindowHandle, _
                             DWMWA_WINDOW_CORNER_PREFERENCE, _
                             DWMWCP_ROUND, 4) = 0 Then

        DP_ApplyRoundedCorners = True

        Exit Function

    End If

    '--------------------------------------------------------
    ' GDI FALLBACK
    '
    ' The region is recreated each time this fallback is
    ' applied because its geometry is based on the current
    ' window dimensions.
    '--------------------------------------------------------

    If GetWindowRect(FormWindowHandle, WindowRectangle) = 0 Then Exit Function

    WindowWidth = WindowRectangle.Right - WindowRectangle.Left
    WindowHeight = WindowRectangle.Bottom - WindowRectangle.Top

    If WindowWidth <= 0 Then Exit Function
    If WindowHeight <= 0 Then Exit Function

    '--------------------------------------------------------
    ' Create rounded region
    '--------------------------------------------------------

    RegionHandle = CreateRoundRectRgn( _
                       0, 0, _
                       WindowWidth + 1, WindowHeight + 1, _
                       CornerRadius * 2, CornerRadius * 2)

    If RegionHandle = 0 Then Exit Function

    '--------------------------------------------------------
    ' Apply region
    '--------------------------------------------------------

    If SetWindowRgn(FormWindowHandle, RegionHandle, True) <> 0 Then

        ' Windows now owns RegionHandle.
        ' DO NOT call DeleteObject here.
        '----------------------------------------------------
        DP_ApplyRoundedCorners = True

    Else

        ' SetWindowRgn failed.
        ' Windows did NOT take ownership.
        ' Clean up the GDI region.
        '----------------------------------------------------
        DeleteObject RegionHandle

    End If

End Function


'============================================================
' Apply border color
'
' Windows 11: DWM border color.
' Older Windows: No border.
'============================================================

Public Function DP_ApplyBorderColor( _
    ByVal PickerForm As Object, _
    ByVal BorderColor As Long) As Boolean

#If Mac Then

    DP_ApplyBorderColor = False
    Exit Function

#ElseIf VBA7 Then

    Dim FormWindowHandle As LongPtr

#Else

    Dim FormWindowHandle As Long

#End If

    '--------------------------------------------------------
    ' Get current HWND
    '--------------------------------------------------------

    FormWindowHandle = GetPickerHwnd(PickerForm)

    If FormWindowHandle = 0 Then Exit Function

    '--------------------------------------------------------
    ' Apply DWM border color.
    '
    ' Windows versions that do not support this attribute
    ' reject the request.
    '--------------------------------------------------------

    If DwmSetWindowAttribute(FormWindowHandle, _
                             DWMWA_BORDER_COLOR, _
                             BorderColor, 4) = 0 Then

        DP_ApplyBorderColor = True

    End If

End Function

