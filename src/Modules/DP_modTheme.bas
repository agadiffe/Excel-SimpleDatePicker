Attribute VB_Name = "DP_modTheme"
Option Explicit

'----------------------------------------
' Font Size
'----------------------------------------

Public Const FONT_SIZE_HEADER As Long = 10
Public Const FONT_SIZE_ARROW As Long = 10
Public Const FONT_SIZE_DAY As Long = 9
Public Const FONT_SIZE_WEEKDAY As Long = 8
Public Const FONT_SIZE_PERIOD As Long = 9
Public Const FONT_SIZE_ACTION_BUTTON As Long = 9


'----------------------------------------
' Theme
'----------------------------------------

Public Function IsDarkMode() As Boolean

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

Public Function ColorBg() As Long

    If IsDarkMode() Then
        ColorBg = RGB(43, 44, 47)
    Else
        ColorBg = RGB(248, 249, 251)
    End If

End Function


Public Function ColorBgToday() As Long

    If IsDarkMode() Then
        ColorBgToday = RGB(48, 72, 96)
    Else
        ColorBgToday = RGB(210, 232, 255)
    End If

End Function


Public Function ColorBgSelected() As Long

    If IsDarkMode() Then
        ColorBgSelected = RGB(65, 135, 205)
    Else
        ColorBgSelected = RGB(70, 145, 220)
    End If

End Function


Public Function ColorBgHover() As Long

    If IsDarkMode() Then
        ColorBgHover = RGB(70, 72, 77)
    Else
        ColorBgHover = RGB(225, 230, 237)
    End If

End Function


Public Function ColorBgSelectedHover() As Long

    If IsDarkMode() Then
        ColorBgSelectedHover = RGB(85, 155, 220)
    Else
        ColorBgSelectedHover = RGB(95, 165, 230)
    End If

End Function


' Text
'--------------------

Public Function ColorText() As Long

    If IsDarkMode() Then
        ColorText = RGB(228, 228, 231)
    Else
        ColorText = RGB(35, 40, 48)
    End If

End Function


Public Function ColorTextSecondary() As Long

    If IsDarkMode() Then
        ColorTextSecondary = RGB(148, 148, 155)
    Else
        ColorTextSecondary = RGB(155, 160, 168)
    End If

End Function


Public Function ColorTextHeader() As Long

    If IsDarkMode() Then
        ColorTextHeader = RGB(220, 220, 223)
    Else
        ColorTextHeader = RGB(30, 50, 75)
    End If

End Function


Public Function ColorTextToday() As Long

    If IsDarkMode() Then
        ColorTextToday = RGB(100, 180, 240)
    Else
        ColorTextToday = RGB(0, 90, 170)
    End If

End Function


Public Function ColorTextSelected() As Long

    ColorTextSelected = RGB(255, 255, 255)

End Function

