Attribute VB_Name = "DP_modPopup"
Option Explicit

' Opens and positions a popup form next to the specified Excel cell,
' automatically adjusting its position to keep it within the Excel window.

'----------------------------------------
' Declarations
'----------------------------------------

Private Const POPUP_GAP As Single = 0
Private Const POPUP_MARGIN As Single = 10
Private Const CALIBRATION_POINTS_BASE As Double = 7200#


'----------------------------------------
' Display
'----------------------------------------

Public Sub DP_ShowPopupNextToCell(ByVal Popup As Object, _
                                  ByVal Target As Range, _
                                  ByVal TargetWindow As Excel.Window)

    Dim Position As Variant

    Position = GetPopupPosition(Target, TargetWindow)

    With Popup
        .StartUpPosition = 0
        .Left = Position(0)
        .Top = Position(1)
    End With

    AdjustPopupPosition Popup, Target, TargetWindow
    Popup.Show

End Sub


'----------------------------------------
' Positioning
'----------------------------------------

Private Function GetPopupPosition(ByVal Target As Range, _
                                  ByVal TargetWindow As Excel.Window) As Variant

    Dim CellRight As Double
    Dim CellTop As Double

    With TargetWindow.ActivePane
        CellRight = .PointsToScreenPixelsX(Target.Left + Target.Width) * GetPointsPerPixelX(TargetWindow)
        CellTop = .PointsToScreenPixelsY(Target.Top) * GetPointsPerPixelY(TargetWindow)
    End With

    GetPopupPosition = Array(CellRight + POPUP_GAP, CellTop)

End Function


Private Sub AdjustPopupPosition(ByVal Popup As Object, _
                                ByVal Target As Range, _
                                ByVal TargetWindow As Excel.Window)

    Dim ExcelRight As Double
    Dim ExcelBottom As Double
    Dim CellLeft As Double
    Dim CellTop As Double

    ExcelRight = TargetWindow.Left + TargetWindow.Width - POPUP_MARGIN
    ExcelBottom = TargetWindow.Top + TargetWindow.Height - POPUP_MARGIN

    CellLeft = GetCellScreenLeft(Target, TargetWindow)
    CellTop = GetCellScreenTop(Target, TargetWindow)

    ' Horizontal
    '--------------------

    If Popup.Left + Popup.Width > ExcelRight Then
        Popup.Left = CellLeft - Popup.Width - POPUP_GAP
    End If

    If Popup.Left < TargetWindow.Left + POPUP_MARGIN Then
        Popup.Left = TargetWindow.Left + POPUP_MARGIN
    End If

    ' Vertical
    '--------------------

    If Popup.Top + Popup.Height > ExcelBottom Then
        Popup.Top = ExcelBottom - Popup.Height
    End If

    If Popup.Top < TargetWindow.Top + POPUP_MARGIN Then
        Popup.Top = TargetWindow.Top + POPUP_MARGIN
    End If

End Sub


'----------------------------------------
' Screen coordinates
'----------------------------------------

Private Function GetCellScreenLeft(ByVal Target As Range, _
                                   ByVal TargetWindow As Excel.Window) As Double

    GetCellScreenLeft = TargetWindow.ActivePane.PointsToScreenPixelsX(Target.Left) * _
                        GetPointsPerPixelX(TargetWindow)

End Function


Private Function GetCellScreenTop(ByVal Target As Range, _
                                  ByVal TargetWindow As Excel.Window) As Double

    GetCellScreenTop = TargetWindow.ActivePane.PointsToScreenPixelsY(Target.Top) * _
                       GetPointsPerPixelY(TargetWindow)

End Function


Private Function GetZoomAdjustedCalibrationPoints(ByVal TargetWindow As Excel.Window) As Double

    GetZoomAdjustedCalibrationPoints = CALIBRATION_POINTS_BASE / (TargetWindow.Zoom / 100#)

End Function


Private Function GetPointsPerPixelX(ByVal TargetWindow As Excel.Window) As Double

    Dim CalibrationPoints As Double

    CalibrationPoints = GetZoomAdjustedCalibrationPoints(TargetWindow)

    GetPointsPerPixelX = _
        1# / _
        ( _
            (TargetWindow.ActivePane.PointsToScreenPixelsX(CalibrationPoints) - _
              TargetWindow.ActivePane.PointsToScreenPixelsX(0)) / _
            CALIBRATION_POINTS_BASE _
        )

End Function


Private Function GetPointsPerPixelY(ByVal TargetWindow As Excel.Window) As Double

    Dim CalibrationPoints As Double

    CalibrationPoints = GetZoomAdjustedCalibrationPoints(TargetWindow)

    GetPointsPerPixelY = _
        1# / _
        ( _
            (TargetWindow.ActivePane.PointsToScreenPixelsY(CalibrationPoints) - _
                TargetWindow.ActivePane.PointsToScreenPixelsY(0)) / _
            CALIBRATION_POINTS_BASE _
        )

End Function

