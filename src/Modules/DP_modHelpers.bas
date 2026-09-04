Attribute VB_Name = "DP_modHelpers"
Option Explicit

'----------------------------------------
' Date Picker
'----------------------------------------

Public Sub ClearPickerControls(ByVal PickerForm As Object)

    Dim ControlIndex As Long

    For ControlIndex = PickerForm.Controls.Count - 1 To 0 Step -1
        PickerForm.Controls.Remove PickerForm.Controls(ControlIndex).Name
    Next ControlIndex

End Sub


Public Sub InitializePicker(ByVal PickerForm As Object, _
                            ByVal FormWidth As Single, _
                            ByVal FormHeight As Single)

    With PickerForm
        .Width = FormWidth
        .Height = FormHeight
        .BackColor = ColorBg()
    End With

End Sub


'----------------------------------------
' Layout
'----------------------------------------

Public Function VerticalTextTop(ByVal ContainerTop As Single, _
                                ByVal ContainerHeight As Single, _
                                ByVal TextHeight As Single) As Single

    Const TEXT_VERTICAL_CORRECTION As Single = 0.033

    VerticalTextTop = ContainerTop + _
                      (ContainerHeight - TextHeight) / 2 - _
                      ContainerHeight * TEXT_VERTICAL_CORRECTION

End Function


Public Function GetFontScale(ByVal CurrentWidth As Single) As Single

    GetFontScale = CurrentWidth / DATEPICKER_BASE_WIDTH

End Function


Public Function GridColumn(ByVal CellIndex As Long, _
                           ByVal GridColumns As Long) As Long

    GridColumn = CellIndex Mod GridColumns

End Function


Public Function GridRow(ByVal CellIndex As Long, _
                        ByVal GridColumns As Long) As Long

    GridRow = CellIndex \ GridColumns

End Function


Public Function GridLeft(ByVal ColumnIndex As Long, _
                         ByVal GridLeftPosition As Single, _
                         ByVal GridCellWidth As Single) As Single

    GridLeft = GridLeftPosition + ColumnIndex * GridCellWidth

End Function


Public Function GridTop(ByVal RowIndex As Long, _
                        ByVal GridTopPosition As Single, _
                        ByVal GridCellHeight As Single) As Single

    GridTop = GridTopPosition + RowIndex * GridCellHeight

End Function


'----------------------------------------
' Reset hover
'----------------------------------------

Public Sub ResetHoverCollection(ByVal Handlers As Collection)

    Dim Handler As Object

    For Each Handler In Handlers
        Handler.ResetHover
    Next Handler

End Sub


Public Sub ResetHoverCollectionExcept(ByVal Handlers As Collection, _
                                      ByVal KeepHandler As Object)

    Dim Handler As Object

    For Each Handler In Handlers
        If Not Handler Is KeepHandler Then
            Handler.ResetHover
        End If
    Next Handler

End Sub

