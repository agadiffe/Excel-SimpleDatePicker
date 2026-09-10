Attribute VB_Name = "DP_modForm"
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Public Enum DP_PickerSize
    DP_SIZE_DATE = 0
    DP_SIZE_MONTH = 1
    DP_SIZE_YEAR = 2
End Enum

Private PickerFormWidth(DP_SIZE_DATE To DP_SIZE_YEAR) As Single
Private PickerFormHeight(DP_SIZE_DATE To DP_SIZE_YEAR) As Single


'----------------------------------------
' Initialization
'----------------------------------------

Public Sub DP_InitializePicker(ByVal PickerForm As Object, _
                               ByVal DesiredInsideWidth As Single, _
                               ByVal DesiredInsideHeight As Single)

    'Get the HWND for THIS UserForm instance.
    DP_CachePickerWindow PickerForm

    If Not DP_SHOW_TITLEBAR Then
        DP_RemoveUserFormTitleBar PickerForm.Caption
    End If

    With PickerForm
        .Width = DesiredInsideWidth
        .Height = DesiredInsideHeight
    End With

    'Apply each desired inside size and cache the resulting form dimensions.
    CachePickerSizes PickerForm

    DP_SetPickerSize PickerForm, DP_SIZE_DATE

    If Not DP_SHOW_TITLEBAR Then
        DP_ApplyRoundedCorners PickerForm
        DP_ApplyBorderColor PickerForm
    End If

End Sub


'----------------------------------------
' Size
'----------------------------------------

Public Sub DP_SetPickerSize(ByVal PickerForm As Object, _
                            ByVal PickerSize As DP_PickerSize)

    Dim SizeChanged As Boolean

    With PickerForm

        If .Width <> PickerFormWidth(PickerSize) Then
            .Width = PickerFormWidth(PickerSize)
            SizeChanged = True
        End If

        If .Height <> PickerFormHeight(PickerSize) Then
            .Height = PickerFormHeight(PickerSize)
            SizeChanged = True
        End If

    End With

    If SizeChanged Then
        If Not DP_IsUsingDwmCorners() Then
            DP_ApplyRoundedCorners PickerForm
        End If
    End If

End Sub


'----------------------------------------
' Cache
'----------------------------------------

'Cache the adjusted form dimensions to avoid flickering caused by resizing
'the form twice when adjusting Width/Height to match InsideWidth/InsideHeight.

Private Sub CachePickerSizes(ByVal PickerForm As Object)

    CacheSize PickerForm, DP_SIZE_DATE, _
              DP_DATEPICKER_WIDTH, DP_DATEPICKER_HEIGHT

    CacheSize PickerForm, DP_SIZE_MONTH, _
              DP_MONTHPICKER_WIDTH, DP_MONTHPICKER_HEIGHT

    CacheSize PickerForm, DP_SIZE_YEAR, _
              DP_YEARPICKER_WIDTH, DP_YEARPICKER_HEIGHT

End Sub


Private Sub CacheSize(ByVal PickerForm As Object, _
                      ByVal PickerSize As DP_PickerSize, _
                      ByVal DesiredInsideWidth As Single, _
                      ByVal DesiredInsideHeight As Single)

    With PickerForm

        .Width = DesiredInsideWidth
        .Height = DesiredInsideHeight

        'Convert desired inside dimensions to actual UserForm dimensions.
        .Width = .Width + DesiredInsideWidth - .InsideWidth
        .Height = .Height + DesiredInsideHeight - .InsideHeight

        'Cache the resulting outer dimensions.
        PickerFormWidth(PickerSize) = .Width
        PickerFormHeight(PickerSize) = .Height

    End With

End Sub

