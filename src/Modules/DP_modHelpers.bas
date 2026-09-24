Attribute VB_Name = "DP_modHelpers"
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private PictureCache As Collection


'----------------------------------------
' Types
'----------------------------------------

Public Enum DP_NavigationDirection
    DP_NAV_PREVIOUS = -1
    DP_NAV_NEXT = 1
End Enum


Public Enum DP_PictureType
    DP_PICTURE_TYPE_FORM_DATE = 0
    DP_PICTURE_TYPE_FORM_MONTH = 1
    DP_PICTURE_TYPE_FORM_YEAR = 2
    DP_PICTURE_TYPE_DAY = 3
    DP_PICTURE_TYPE_MONTH = 4
    DP_PICTURE_TYPE_YEAR = 5
    DP_PICTURE_TYPE_HEADER = 6
    DP_PICTURE_TYPE_ARROW = 7
    DP_PICTURE_TYPE_SHORTCUT = 8
End Enum


'----------------------------------------
' DatePicker form
'----------------------------------------

Public Function DP_GetWindowKey(ByVal TargetWindow As Excel.Window) As String

    DP_GetWindowKey = TargetWindow.Parent.FullName & "|" & _
                      CStr(TargetWindow.WindowNumber)

End Function


'----------------------------------------
' Date validation
'----------------------------------------

Public Function DP_MinDate() As Date

    DP_MinDate = DateSerial(1901, 1, 1)

End Function


Public Function DP_MaxDate() As Date

    DP_MaxDate = DateSerial(Year(Date) + 100, 12, 31)

End Function


Public Function DP_IsDateOutsideRange(ByVal CellDate As Date) As Boolean

    ' strip the time
    CellDate = DateValue(CellDate)

    DP_IsDateOutsideRange = CellDate < DP_MinDate() Or CellDate > DP_MaxDate()

End Function


'----------------------------------------
' Layout
'----------------------------------------

Public Function DP_VerticalTextTop(ByVal ContainerTop As Single, _
                                   ByVal ContainerHeight As Single, _
                                   ByVal TextHeight As Single) As Single

    Const TEXT_VERTICAL_CORRECTION As Single = 0.03

    DP_VerticalTextTop = ContainerTop + _
                         (ContainerHeight - TextHeight) / 2 - _
                         ContainerHeight * TEXT_VERTICAL_CORRECTION

End Function


Public Function DP_GetFontScale(ByVal CurrentWidth As Single) As Single

    DP_GetFontScale = CurrentWidth / DP_DATEPICKER_BASE_WIDTH

End Function


Public Function DP_GridColumn(ByVal CellIndex As Long, _
                              ByVal GridColumns As Long) As Long

    DP_GridColumn = CellIndex Mod GridColumns

End Function


Public Function DP_GridRow(ByVal CellIndex As Long, _
                           ByVal GridColumns As Long) As Long

    DP_GridRow = CellIndex \ GridColumns

End Function


Private Function RoundPosition(ByVal Value As Single) As Single

    'Use conventional rounding for stable UI positioning.
    RoundPosition = Int(Value + 0.5)

End Function


Public Function DP_GridLeft(ByVal ColumnIndex As Long, _
                            ByVal GridLeftPosition As Single, _
                            ByVal GridCellWidth As Single, _
                            ByVal CellWidth As Single) As Single

    DP_GridLeft = RoundPosition(GridLeftPosition + _
                                ColumnIndex * GridCellWidth + _
                                (GridCellWidth - CellWidth) / 2)

End Function


Public Function DP_GridTop(ByVal RowIndex As Long, _
                           ByVal GridTopPosition As Single, _
                           ByVal GridCellHeight As Single, _
                           ByVal CellHeight As Single) As Single

    DP_GridTop = RoundPosition(GridTopPosition + _
                               RowIndex * GridCellHeight + _
                               (GridCellHeight - CellHeight) / 2)

End Function


'----------------------------------------
' UI helpers
'----------------------------------------

Public Sub DP_SetVisibleCollection(ByVal Handlers As Collection, _
                                   ByVal IsVisible As Boolean)

    Dim Handler As Object

    For Each Handler In Handlers
        Handler.SetVisible IsVisible
    Next Handler

End Sub


Public Sub DP_UpdateHandlerCollection(ByVal Handlers As Collection)

    Dim Handler As Object

    For Each Handler In Handlers
        Handler.UpdateHandler
    Next Handler

End Sub


'----------------------------------------
' Button pictures
'----------------------------------------

Public Sub DP_ClearPictureCache()

    Set PictureCache = Nothing

End Sub


Public Function DP_GetCachedPicture(ByVal PictureType As DP_PictureType, _
                                    ByVal FillColor As Long, _
                                    ByVal BackgroundColor As Long, _
                                    Optional ByVal ButtonWidth As Single = 0, _
                                    Optional ByVal BorderColor As Long = DP_NO_BORDER_COLOR, _
                                    Optional ByVal RenderScale As Long = DP_RENDER_SCALE) As StdPicture

    Dim Width As Single
    Dim Height As Single
    Dim Shape As DP_ButtonShape

    Dim CacheKey As String
    Dim Picture As StdPicture

    #If Mac Then
        Exit Function
    #End If

    If PictureCache Is Nothing Then
        Set PictureCache = New Collection
    End If

    Select Case PictureType
        Case DP_PICTURE_TYPE_FORM_DATE
            Width = DP_DATEPICKER_WIDTH
            Height = DP_DATEPICKER_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_PICTURE_TYPE_FORM_MONTH
            Width = DP_MONTHPICKER_WIDTH
            Height = DP_MONTHPICKER_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_PICTURE_TYPE_FORM_YEAR
            Width = DP_YEARPICKER_WIDTH
            Height = DP_YEARPICKER_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_PICTURE_TYPE_DAY
            Width = DP_CALENDAR_CELL_WIDTH
            Height = DP_CALENDAR_CELL_HEIGHT
            Shape = DP_DAY_BUTTON_SHAPE
        Case DP_PICTURE_TYPE_MONTH
            Width = DP_MONTH_CELL_WIDTH
            Height = DP_PERIOD_CELL_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_PICTURE_TYPE_YEAR
            Width = DP_YEAR_CELL_WIDTH
            Height = DP_PERIOD_CELL_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_PICTURE_TYPE_HEADER
            Width = DP_HEADER_WIDTH
            Height = DP_HEADER_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_PICTURE_TYPE_ARROW
            Width = DP_ARROW_WIDTH
            Height = DP_ARROW_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_PICTURE_TYPE_SHORTCUT
            Width = ButtonWidth
            Height = DP_ACTION_BUTTON_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case Else
            Exit Function
    End Select

    CacheKey = CStr(Shape) & "|" & _
               CStr(Width) & "|" & _
               CStr(Height) & "|" & _
               CStr(FillColor) & "|" & _
               CStr(BackgroundColor) & "|" & _
               CStr(BorderColor) & "|" & _
               CStr(RenderScale)

    On Error Resume Next
    Set Picture = PictureCache.Item(CacheKey)
    On Error GoTo 0

    If Picture Is Nothing Then

        Set Picture = DP_GetPicture(Width, Height, _
                                    FillColor, BackgroundColor, _
                                    Shape, BorderColor, RenderScale)

        If Picture Is Nothing Then Exit Function

        PictureCache.Add Picture, CacheKey

    End If

    Set DP_GetCachedPicture = Picture

End Function

