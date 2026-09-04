Attribute VB_Name = "DP_modLocalization"
Option Explicit

' Primary Language IDs
Private Const LANG_ARABIC As Long = &H1
Private Const LANG_CHINESE_SIMPLIFIED As Long = &H4
Private Const LANG_GERMAN As Long = &H7
Private Const LANG_SPANISH As Long = &HA
Private Const LANG_FRENCH As Long = &HC
Private Const LANG_ITALIAN As Long = &H10
Private Const LANG_JAPANESE As Long = &H11
Private Const LANG_KOREAN As Long = &H12
Private Const LANG_RUSSIAN As Long = &H19
Private Const LANG_VIETNAMESE As Long = &H2A
Private Const LANG_HINDI As Long = &H39


' GetDatePickerActionText
Public Function GetDatePickerActionText(ByVal ActionKey As String) As String

    Dim LanguageID As Long
    Dim PrimaryLanguageID As Long

    LanguageID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    PrimaryLanguageID = LanguageID And &H3FF

    Select Case ActionKey

        '--------------------------------------------------
        ' Today
        '--------------------------------------------------
        Case "GO_TO_TODAY"

            Select Case PrimaryLanguageID
                Case LANG_FRENCH
                    GetDatePickerActionText = "Aujourd'hui"
                Case LANG_GERMAN
                    GetDatePickerActionText = "Heute"
                Case LANG_SPANISH
                    GetDatePickerActionText = "Hoy"
                Case LANG_ITALIAN
                    GetDatePickerActionText = "Oggi"
                Case LANG_JAPANESE
                    GetDatePickerActionText = _
                        ChrW(&H4ECA) & ChrW(&H65E5)
                Case LANG_KOREAN
                    GetDatePickerActionText = _
                        ChrW(&HC624) & ChrW(&HB298)
                Case LANG_CHINESE_SIMPLIFIED
                    GetDatePickerActionText = _
                        ChrW(&H4ECA) & ChrW(&H5929)
                Case LANG_ARABIC
                    GetDatePickerActionText = _
                        ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & _
                        ChrW(&H648) & ChrW(&H645)
                Case LANG_HINDI
                    GetDatePickerActionText = _
                        ChrW(&H906) & ChrW(&H91C)
                Case LANG_RUSSIAN
                    GetDatePickerActionText = _
                        ChrW(&H421) & ChrW(&H435) & ChrW(&H433) & _
                        ChrW(&H43E) & ChrW(&H434) & ChrW(&H43D) & _
                        ChrW(&H44F)
                Case LANG_VIETNAMESE
                    GetDatePickerActionText = _
                        ChrW(&H48) & ChrW(&HF4) & ChrW(&H6D) & _
                        " " & _
                        ChrW(&H6E) & ChrW(&H61) & ChrW(&H79)
                Case Else
                    GetDatePickerActionText = "Today"
            End Select

        '--------------------------------------------------
        ' This Month
        '--------------------------------------------------
        Case "GO_TO_CURRENT_MONTH"

            Select Case PrimaryLanguageID
                Case LANG_FRENCH
                    GetDatePickerActionText = "Ce mois"
                Case LANG_GERMAN
                    GetDatePickerActionText = "Dieser Monat"
                Case LANG_SPANISH
                    GetDatePickerActionText = "Este mes"
                Case LANG_ITALIAN
                    GetDatePickerActionText = "Questo mese"
                Case LANG_JAPANESE
                    GetDatePickerActionText = _
                        ChrW(&H4ECA) & ChrW(&H6708)
                Case LANG_KOREAN
                    GetDatePickerActionText = _
                        ChrW(&HC774) & ChrW(&HBC88) & _
                        " " & _
                        ChrW(&HB2EC)
                Case LANG_CHINESE_SIMPLIFIED
                    GetDatePickerActionText = _
                        ChrW(&H672C) & ChrW(&H6708)
                Case LANG_ARABIC
                    GetDatePickerActionText = _
                        ChrW(&H647) & ChrW(&H630) & ChrW(&H627) & _
                        " " & _
                        ChrW(&H627) & ChrW(&H644) & ChrW(&H634) & _
                        ChrW(&H647) & ChrW(&H631)
                Case LANG_HINDI
                    GetDatePickerActionText = _
                        ChrW(&H907) & ChrW(&H938) & _
                        " " & _
                        ChrW(&H92E) & ChrW(&H939) & ChrW(&H940) & _
                        ChrW(&H928) & ChrW(&H947)
                Case LANG_RUSSIAN
                    GetDatePickerActionText = _
                        ChrW(&H42D) & ChrW(&H442) & ChrW(&H43E) & _
                        ChrW(&H442) & " " & _
                        ChrW(&H43C) & ChrW(&H435) & ChrW(&H441) & _
                        ChrW(&H44F) & ChrW(&H446)
                Case LANG_VIETNAMESE
                    GetDatePickerActionText = _
                        ChrW(&H54) & ChrW(&H68) & ChrW(&HE1) & _
                        ChrW(&H6E) & ChrW(&H67) & " " & _
                        ChrW(&H6E) & ChrW(&HE0) & ChrW(&H79)
                Case Else
                    GetDatePickerActionText = "This Month"
            End Select

        '--------------------------------------------------
        ' This Year
        '--------------------------------------------------
        Case "GO_TO_CURRENT_YEAR"

            Select Case PrimaryLanguageID
                Case LANG_FRENCH
                    GetDatePickerActionText = "Cette année"
                Case LANG_GERMAN
                    GetDatePickerActionText = "Dieses Jahr"
                Case LANG_SPANISH
                    GetDatePickerActionText = "Este año"
                Case LANG_ITALIAN
                    GetDatePickerActionText = "Quest'anno"
                Case LANG_JAPANESE
                    GetDatePickerActionText = _
                        ChrW(&H4ECA) & ChrW(&H5E74)
                Case LANG_KOREAN
                    GetDatePickerActionText = _
                        ChrW(&HC62C) & ChrW(&HD574)
                Case LANG_CHINESE_SIMPLIFIED
                    GetDatePickerActionText = _
                        ChrW(&H4ECA) & ChrW(&H5E74)
                Case LANG_ARABIC
                    GetDatePickerActionText = _
                        ChrW(&H647) & ChrW(&H630) & ChrW(&H647) & _
                        " " & _
                        ChrW(&H627) & ChrW(&H644) & ChrW(&H633) & _
                        ChrW(&H646) & ChrW(&H629)
                Case LANG_HINDI
                    GetDatePickerActionText = _
                        ChrW(&H907) & ChrW(&H938) & _
                        " " & _
                        ChrW(&H938) & ChrW(&H93E) & ChrW(&H932)
                Case LANG_RUSSIAN
                    GetDatePickerActionText = _
                        ChrW(&H42D) & ChrW(&H442) & ChrW(&H43E) & _
                        ChrW(&H442) & " " & _
                        ChrW(&H433) & ChrW(&H43E) & ChrW(&H434)
                Case LANG_VIETNAMESE
                    GetDatePickerActionText = _
                        ChrW(&H4E) & ChrW(&H103) & ChrW(&H6D) & _
                        " " & _
                        ChrW(&H6E) & ChrW(&H61) & ChrW(&H79)
                Case Else
                    GetDatePickerActionText = "This Year"
            End Select

    End Select

End Function

