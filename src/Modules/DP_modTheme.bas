Attribute VB_Name = "DP_modTheme"
Option Explicit

'----------------------------------------
' Font Size
'----------------------------------------

Public Const DP_FONT_SIZE_HEADER As Long = 10
Public Const DP_FONT_SIZE_ARROW As Long = 10
Public Const DP_FONT_SIZE_DAY As Long = 9
Public Const DP_FONT_SIZE_WEEKDAY As Long = 8
Public Const DP_FONT_SIZE_PERIOD As Long = 9
Public Const DP_FONT_SIZE_ACTION_BUTTON As Long = 9


'----------------------------------------
' Theme
'----------------------------------------

Private Function IsDarkMode() As Boolean

    If IsWindows() Then
        IsDarkMode = IsExcelDarkThemeActive()
    Else
        IsDarkMode = IsMacDarkThemeActive()
    End If

End Function


Private Function IsWindows() As Boolean

#If Mac Then
    IsWindows = False
#Else
    IsWindows = True
#End If

End Function


' Windows
'--------------------

Private Function IsExcelDarkThemeActive() As Boolean

    Const OfficeThemeKey As String = _
        "HKCU\Software\Microsoft\Office\16.0\Common\UI Theme"

    Const WindowsThemeKey As String = _
        "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme"

    Dim officeTheme As Long
    Dim windowsLightTheme As Long

    With CreateObject("WScript.Shell")

        On Error Resume Next

        ' Read Excel / Office theme
        officeTheme = .RegRead(OfficeThemeKey)

        ' If Windows key is missing, assume Light Mode
        windowsLightTheme = 1
        windowsLightTheme = .RegRead(WindowsThemeKey)

        On Error GoTo 0

    End With

    ' Office Theme:
    ' 3 = Gray
    ' 4 = Black
    ' 5 = White
    ' 6 = Use System setting
    ' 7 = Color

    Select Case officeTheme
        Case 3, 4
            ' Office explicitly uses a dark theme
            IsExcelDarkThemeActive = True
        Case 6
            ' Office follows the Windows system theme
            IsExcelDarkThemeActive = (windowsLightTheme = 0)
        Case Else
            ' White, Color, or unknown
            IsExcelDarkThemeActive = False
    End Select

End Function


' Mac
'--------------------

' Mac requires an external AppleScript because VBA does not provide
' a native way to read the current macOS appearance.
'
' Save the following as:
'
'   DatePickerTheme.applescript
'
' in:
'
'   ~/Library/Application Scripts/com.microsoft.Excel/
'
' The script is called through AppleScriptTask and returns True when
' macOS is currently using Dark Mode:
'
'   on GetDarkMode(paramString)
'       tell application "System Events"
'           tell appearance preferences
'               return dark mode
'           end tell
'       end tell
'   end GetDarkMode
'
' AppleScriptTask and this Excel-specific folder are documented by
' Microsoft for Office on Mac.

Private Function IsMacDarkThemeActive() As Boolean

#If VBA7 Then

    Dim result As String

    On Error GoTo LightMode

    result = AppleScriptTask( _
        "DatePickerTheme.applescript", _
        "GetDarkMode", _
        "")

    IsMacDarkThemeActive = _
        (LCase$(Trim$(result)) = "true")

    Exit Function

LightMode:

    ' Safe fallback if AppleScript fails
    IsMacDarkThemeActive = False

#Else

    ' AppleScriptTask is not available in VBA6.
    ' Fall back to Light Mode.
    IsMacDarkThemeActive = False

#End If

End Function


'----------------------------------------
' Color
'----------------------------------------

' Background
'--------------------

Public Function DP_ColorBg() As Long

    If IsDarkMode() Then
        DP_ColorBg = RGB(43, 44, 47)
    Else
        DP_ColorBg = RGB(248, 249, 251)
    End If

End Function


Public Function DP_ColorBgToday() As Long

    If IsDarkMode() Then
        DP_ColorBgToday = RGB(48, 72, 96)
    Else
        DP_ColorBgToday = RGB(210, 232, 255)
    End If

End Function


Public Function DP_ColorBgSelected() As Long

    If IsDarkMode() Then
        DP_ColorBgSelected = RGB(65, 135, 205)
    Else
        DP_ColorBgSelected = RGB(70, 145, 220)
    End If

End Function


Public Function DP_ColorBgHover() As Long

    If IsDarkMode() Then
        DP_ColorBgHover = RGB(70, 72, 77)
    Else
        DP_ColorBgHover = RGB(225, 230, 237)
    End If

End Function


Public Function DP_ColorBgSelectedHover() As Long

    If IsDarkMode() Then
        DP_ColorBgSelectedHover = RGB(85, 155, 220)
    Else
        DP_ColorBgSelectedHover = RGB(95, 165, 230)
    End If

End Function


' Text
'--------------------

Public Function DP_ColorText() As Long

    If IsDarkMode() Then
        DP_ColorText = RGB(228, 228, 231)
    Else
        DP_ColorText = RGB(35, 40, 48)
    End If

End Function


Public Function DP_ColorTextSecondary() As Long

    If IsDarkMode() Then
        DP_ColorTextSecondary = RGB(148, 148, 155)
    Else
        DP_ColorTextSecondary = RGB(155, 160, 168)
    End If

End Function


Public Function DP_ColorTextHeader() As Long

    If IsDarkMode() Then
        DP_ColorTextHeader = RGB(220, 220, 223)
    Else
        DP_ColorTextHeader = RGB(30, 50, 75)
    End If

End Function


Public Function DP_ColorTextToday() As Long

    If IsDarkMode() Then
        DP_ColorTextToday = RGB(100, 180, 240)
    Else
        DP_ColorTextToday = RGB(0, 90, 170)
    End If

End Function


Public Function DP_ColorTextSelected() As Long

    DP_ColorTextSelected = RGB(255, 255, 255)

End Function

