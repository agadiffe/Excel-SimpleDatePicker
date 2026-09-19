Attribute VB_Name = "DP_modHelpers"
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private ButtonPictureCache As Collection


'----------------------------------------
' Types
'----------------------------------------

Public Enum DP_NavigationDirection
    DP_NAV_PREVIOUS = -1
    DP_NAV_NEXT = 1
End Enum


Public Enum DP_ButtonType
    DP_BUTTON_TYPE_DAY = 0
    DP_BUTTON_TYPE_MONTH = 1
    DP_BUTTON_TYPE_YEAR = 2
    DP_BUTTON_TYPE_HEADER = 3
    DP_BUTTON_TYPE_ARROW = 4
    DP_BUTTON_TYPE_SHORTCUT = 5
End Enum


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


Public Sub DP_RefreshThemeCollection(ByVal Handlers As Collection)

    Dim Handler As Object

    For Each Handler In Handlers
        Handler.RefreshTheme
    Next Handler

End Sub


Public Sub DP_UpdateStateCollection(ByVal Handlers As Collection)

    Dim Handler As Object

    For Each Handler In Handlers
        Handler.UpdateState
    Next Handler

End Sub


'----------------------------------------
' Button pictures
'----------------------------------------

Public Sub DP_ClearButtonPictureCache()

    Set ButtonPictureCache = Nothing

End Sub


Public Function DP_GetCachedButtonPicture(ByVal ButtonType As DP_ButtonType, _
                                          ByVal FillColor As Long, _
                                          ByVal BackgroundColor As Long, _
                                          Optional ByVal ButtonWidth As Single = 0) As StdPicture

    Dim Width As Single
    Dim Height As Single
    Dim Shape As DP_ButtonShape

    Dim CacheKey As String
    Dim Picture As StdPicture

    #If Mac Then
        Exit Function
    #End If

    If ButtonPictureCache Is Nothing Then
        Set ButtonPictureCache = New Collection
    End If

    Select Case ButtonType
        Case DP_BUTTON_TYPE_DAY
            Width = DP_CALENDAR_CELL_WIDTH
            Height = DP_CALENDAR_CELL_HEIGHT
            Shape = DP_DAY_BUTTON_SHAPE
        Case DP_BUTTON_TYPE_MONTH
            Width = DP_MONTH_CELL_WIDTH
            Height = DP_PERIOD_CELL_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_BUTTON_TYPE_YEAR
            Width = DP_YEAR_CELL_WIDTH
            Height = DP_PERIOD_CELL_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_BUTTON_TYPE_HEADER
            Width = DP_HEADER_WIDTH
            Height = DP_HEADER_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_BUTTON_TYPE_ARROW
            Width = DP_ARROW_WIDTH
            Height = DP_ARROW_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case DP_BUTTON_TYPE_SHORTCUT
            Width = ButtonWidth
            Height = DP_ACTION_BUTTON_HEIGHT
            Shape = DP_BUTTON_SHAPE_RECTANGLE
        Case Else
            Exit Function
    End Select

    CacheKey = CStr(Shape) & "|" & _
               CStr(Width) & "|" & CStr(Height) & "|" & _
               CStr(FillColor) & "|" & CStr(BackgroundColor)

    On Error Resume Next
    Set Picture = ButtonPictureCache.Item(CacheKey)
    On Error GoTo 0

    If Picture Is Nothing Then

        Set Picture = DP_GetButtonPicture(Width, Height, _
                                          FillColor, BackgroundColor, _
                                          Shape)

        If Picture Is Nothing Then Exit Function

        ButtonPictureCache.Add Picture, CacheKey

    End If

    Set DP_GetCachedButtonPicture = Picture

End Function

