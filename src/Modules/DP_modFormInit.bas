Attribute VB_Name = "DP_modFormInit"
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

Public Sub DP_InitializePicker(ByVal PickerForm As DP_frmDatePicker)

    If Not DP_SHOW_TITLEBAR Then
        DP_RemoveUserFormTitleBar PickerForm
    End If

    'Apply each desired inside size and cache the resulting form dimensions.
    CachePickerSizes PickerForm

    ' If DP_SIZE_DATE is cached in last, DP_SetPickerSize SizeChanged will be False
    With PickerForm
        .Width = 0
        .Height = 0
    End With

End Sub


Public Function DP_CreateWindowCaption() As String

    DP_CreateWindowCaption = "SimpleDatePicker_" & _
                             Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                             Format$(CLng((Timer - Int(Timer)) * 1000), "000")

End Function


'----------------------------------------
' Size
'----------------------------------------

Public Sub DP_SetPickerSize(ByVal PickerForm As DP_frmDatePicker, _
                            ByVal PickerSize As DP_PickerSize)

    Dim SizeChanged As Boolean

    With PickerForm

        ' Appears to prevent an occasional white flash
        ' on the clicked date-picker control while resizing
        .BackColor = DP_ColorBg()

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
        DP_ApplyRoundedCorners PickerForm
        DP_ApplyPickerBorder PickerForm, PickerSize
    End If

End Sub


Public Sub DP_ApplyPickerBorder(ByVal PickerForm As DP_frmDatePicker, _
                                ByVal PickerSize As DP_PickerSize)

    #If Mac Then

        ' Mac: DWM is unavailable; no border picture is required.
        PickerForm.Picture = Nothing

    #Else

        Dim PickerType As DP_PictureType

        If DP_ApplyBorderColor(PickerForm, DP_ColorBorder()) Then

            ' DWM provides the border.
            PickerForm.Picture = Nothing

        Else
            ' DWM border unavailable; use the 1px border picture.

            Select Case PickerSize
                Case DP_SIZE_DATE
                    PickerType = DP_PICTURE_TYPE_FORM_DATE
                Case DP_SIZE_MONTH
                    PickerType = DP_PICTURE_TYPE_FORM_MONTH
                Case DP_SIZE_YEAR
                    PickerType = DP_PICTURE_TYPE_FORM_YEAR
            End Select

            PickerForm.PictureSizeMode = fmPictureSizeModeStretch
            PickerForm.Picture = _
                DP_GetCachedPicture(PickerType, _
                                    DP_ColorBg(), _
                                    DP_ColorBg(), _
                                    BorderColor:=DP_ColorBorder(), _
                                    RenderScale:=4)
        End If

    #End If

End Sub


'----------------------------------------
' Size cache
'----------------------------------------

' Cache the adjusted form dimensions to prevent an occasional white flash
' on the clicked date-picker control caused by resizing the form twice
' when adjusting Width/Height to match InsideWidth/InsideHeight.

Private Sub CachePickerSizes(ByVal PickerForm As DP_frmDatePicker)

    CacheSize PickerForm, DP_SIZE_DATE, _
              DP_DATEPICKER_WIDTH, DP_DATEPICKER_HEIGHT

    CacheSize PickerForm, DP_SIZE_MONTH, _
              DP_MONTHPICKER_WIDTH, DP_MONTHPICKER_HEIGHT

    CacheSize PickerForm, DP_SIZE_YEAR, _
              DP_YEARPICKER_WIDTH, DP_YEARPICKER_HEIGHT

End Sub


Private Sub CacheSize(ByVal PickerForm As DP_frmDatePicker, _
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

