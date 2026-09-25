Attribute VB_Name = "DP_modAPIPicture"
Option Explicit

'============================================================
' DP_modButtonPictureAPI
'
' Creates opaque anti-aliased button pictures as StdPicture
' objects suitable for an MSForms.Image.Picture property.
'
' Supported shapes:
'   DP_BUTTON_SHAPE_RECTANGLE: Rectangle or rounded rectangle.
'   DP_BUTTON_SHAPE_CIRCLE: Circle.
'
' A circle always uses the smaller of the requested width and
' height and is centered inside the bitmap.
'
' Rectangle:
'   CornerRadius = 0: Normal rectangle.
'   CornerRadius > 0: Rounded rectangle.
'
' The bitmap is rendered using RenderScale.
' A RenderScale of 1 can be used for larger pictures where
' supersampling is unnecessary.
'
' Optional BorderColor:
'   -1: No border.
'   Any other color: 1 logical pixel border.
'
' The border is created by drawing the outer shape in
' BorderColor and then drawing the shape inset by exactly
' one logical pixel in FillColor.
'
'
' Windows only.
' VBA6 / VBA7.
' 32-bit / 64-bit.
'============================================================


'============================================================
' Public API
'============================================================

Public Enum DP_ButtonShape
    DP_BUTTON_SHAPE_RECTANGLE = 0
    DP_BUTTON_SHAPE_CIRCLE = 1
End Enum


'============================================================
' GDI+ startup
'============================================================

Private Type GdiplusStartupInput

    GdiplusVersion As Long

#If VBA7 Then
    DebugEventCallback As LongPtr
#Else
    DebugEventCallback As Long
#End If

    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long

End Type


'============================================================
' Picture descriptor
'============================================================

#If VBA7 Then

Private Type PICTDESC
    Size As Long
    Type As Long
    hPic As LongPtr
    hPal As LongPtr
End Type

#Else

Private Type PICTDESC
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

#End If


'============================================================
' GUID
'============================================================

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


'============================================================
' Constants
'============================================================

Private Const GDIP_OK As Long = 0

Private Const PICTYPE_BITMAP As Long = 1

Private Const PIXEL_FORMAT_32BPP_ARGB As Long = &H26200A

' SmoothingModeHighQuality
Private Const SMOOTHING_HIGH_QUALITY As Long = 2

' PixelOffsetModeHighQuality
Private Const PIXEL_OFFSET_HIGH_QUALITY As Long = 2

' FillModeWinding
Private Const FILL_MODE_WINDING As Long = 1

' Default supersampling factor.
Public Const DP_RENDER_SCALE As Long = 4

Public Const DP_NO_BORDER_COLOR As Long = -1

Private Const CORNER_RADIUS As Single = 6

'------------------------------------------------------------
' Border width in logical pixels.
'
' The border is intentionally fixed at exactly 1 logical
' pixel. RenderScale determines how many bitmap pixels are
' used internally to represent that pixel.
'------------------------------------------------------------

Private Const BORDER_WIDTH As Long = 1


'============================================================
' GDI+ declarations
'============================================================

#If VBA7 And Not Mac Then

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
    ByRef Token As LongPtr, _
    ByRef StartupInput As GdiplusStartupInput, _
    ByVal Output As LongPtr) As Long

Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" ( _
    ByVal Token As LongPtr)

Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" ( _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal Stride As Long, _
    ByVal PixelFormat As Long, _
    ByVal Scan0 As LongPtr, _
    ByRef Bitmap As LongPtr) As Long

Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" ( _
    ByVal Image As LongPtr, _
    ByRef Graphics As LongPtr) As Long

Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" ( _
    ByVal Graphics As LongPtr) As Long

Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" ( _
    ByVal Image As LongPtr) As Long

Private Declare PtrSafe Function GdipSetSmoothingMode Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, _
    ByVal SmoothingMode As Long) As Long

Private Declare PtrSafe Function GdipSetPixelOffsetMode Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, _
    ByVal PixelOffsetMode As Long) As Long

Private Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" ( _
    ByVal Color As Long, _
    ByRef Brush As LongPtr) As Long

Private Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" ( _
    ByVal Brush As LongPtr) As Long

Private Declare PtrSafe Function GdipFillRectangleI Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, _
    ByVal Brush As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As Long

'------------------------------------------------------------
' Native ellipse
'------------------------------------------------------------

Private Declare PtrSafe Function GdipFillEllipseI Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, _
    ByVal Brush As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As Long

'------------------------------------------------------------
' GraphicsPath
'------------------------------------------------------------

Private Declare PtrSafe Function GdipCreatePath Lib "gdiplus" ( _
    ByVal BrushMode As Long, _
    ByRef Path As LongPtr) As Long

Private Declare PtrSafe Function GdipDeletePath Lib "gdiplus" ( _
    ByVal Path As LongPtr) As Long

Private Declare PtrSafe Function GdipAddPathArcI Lib "gdiplus" ( _
    ByVal Path As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal StartAngle As Single, _
    ByVal SweepAngle As Single) As Long

Private Declare PtrSafe Function GdipClosePathFigure Lib "gdiplus" ( _
    ByVal Path As LongPtr) As Long

Private Declare PtrSafe Function GdipFillPath Lib "gdiplus" ( _
    ByVal Graphics As LongPtr, _
    ByVal Brush As LongPtr, _
    ByVal Path As LongPtr) As Long

'------------------------------------------------------------
' Bitmap -> HBITMAP
'------------------------------------------------------------

Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
    ByVal Bitmap As LongPtr, _
    ByRef HBitmap As LongPtr, _
    ByVal Background As Long) As Long

'------------------------------------------------------------
' StdPicture
'------------------------------------------------------------

Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef PictureDescription As PICTDESC, _
    ByRef ReferenceIID As GUID, _
    ByVal PictureOwnsHandle As Long, _
    ByRef Picture As IPicture) As Long

'------------------------------------------------------------
' GDI
'------------------------------------------------------------

Private Declare PtrSafe Function DeleteObject Lib "gdi32" ( _
    ByVal ObjectHandle As LongPtr) As Long


#ElseIf Not Mac Then

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
    ByRef Token As Long, _
    ByRef StartupInput As GdiplusStartupInput, _
    ByVal Output As Long) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus" ( _
    ByVal Token As Long)

Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" ( _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal Stride As Long, _
    ByVal PixelFormat As Long, _
    ByVal Scan0 As Long, _
    ByRef Bitmap As Long) As Long

Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" ( _
    ByVal Image As Long, _
    ByRef Graphics As Long) As Long

Private Declare Function GdipDeleteGraphics Lib "gdiplus" ( _
    ByVal Graphics As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
    ByVal Image As Long) As Long

Private Declare Function GdipSetSmoothingMode Lib "gdiplus" ( _
    ByVal Graphics As Long, _
    ByVal SmoothingMode As Long) As Long

Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" ( _
    ByVal Graphics As Long, _
    ByVal PixelOffsetMode As Long) As Long

Private Declare Function GdipCreateSolidFill Lib "gdiplus" ( _
    ByVal Color As Long, _
    ByRef Brush As Long) As Long

Private Declare Function GdipDeleteBrush Lib "gdiplus" ( _
    ByVal Brush As Long) As Long

Private Declare Function GdipFillRectangleI Lib "gdiplus" ( _
    ByVal Graphics As Long, _
    ByVal Brush As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As Long

'------------------------------------------------------------
' Native ellipse
'------------------------------------------------------------

Private Declare Function GdipFillEllipseI Lib "gdiplus" ( _
    ByVal Graphics As Long, _
    ByVal Brush As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As Long

'------------------------------------------------------------
' GraphicsPath
'------------------------------------------------------------

Private Declare Function GdipCreatePath Lib "gdiplus" ( _
    ByVal BrushMode As Long, _
    ByRef Path As Long) As Long

Private Declare Function GdipDeletePath Lib "gdiplus" ( _
    ByVal Path As Long) As Long

Private Declare Function GdipAddPathArcI Lib "gdiplus" ( _
    ByVal Path As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal StartAngle As Single, _
    ByVal SweepAngle As Single) As Long

Private Declare Function GdipClosePathFigure Lib "gdiplus" ( _
    ByVal Path As Long) As Long

Private Declare Function GdipFillPath Lib "gdiplus" ( _
    ByVal Graphics As Long, _
    ByVal Brush As Long, _
    ByVal Path As Long) As Long

'------------------------------------------------------------
' Bitmap -> HBITMAP
'------------------------------------------------------------

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
    ByVal Bitmap As Long, _
    ByRef HBitmap As Long, _
    ByVal Background As Long) As Long

'------------------------------------------------------------
' StdPicture
'------------------------------------------------------------

Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef PictureDescription As PICTDESC, _
    ByRef ReferenceIID As GUID, _
    ByVal PictureOwnsHandle As Long, _
    ByRef Picture As IPicture) As Long

'------------------------------------------------------------
' GDI
'------------------------------------------------------------

Private Declare Function DeleteObject Lib "gdi32" ( _
    ByVal ObjectHandle As Long) As Long

#End If


'============================================================
' Public function
'
' Creates a picture for any UI element.
'
' Shape:
'   DP_BUTTON_SHAPE_RECTANGLE: Rectangle or rounded rectangle.
'   DP_BUTTON_SHAPE_CIRCLE: Circle.
'
' CornerRadius:
'   Used only for DP_BUTTON_SHAPE_RECTANGLE.
'
'   Values below 6 may produce sharper corners,
'   appearing more like diagonal cuts than rounded corners.
'
' BorderColor:
'   -1: No border.
'   Any other color: 1 logical pixel border.
'
' RenderScale:
'   Default = DP_RENDER_SCALE.
'
'   Use 1 for larger pictures where supersampling is not
'   necessary.
'============================================================

Public Function DP_GetPicture( _
    ByVal CellWidth As Single, _
    ByVal CellHeight As Single, _
    ByVal FillColor As Long, _
    ByVal BackgroundColor As Long, _
    Optional ByVal Shape As DP_ButtonShape = DP_BUTTON_SHAPE_RECTANGLE, _
    Optional ByVal BorderColor As Long = DP_NO_BORDER_COLOR, _
    Optional ByVal RenderScale As Long = DP_RENDER_SCALE, _
    Optional ByVal CornerRadius As Single = CORNER_RADIUS) As StdPicture

#If Mac Then

    Exit Function

#ElseIf VBA7 Then

    Dim GdiToken As LongPtr
    Dim Bitmap As LongPtr
    Dim Graphics As LongPtr

    Dim OuterPath As LongPtr
    Dim InnerPath As LongPtr

    Dim FillBrush As LongPtr
    Dim BackgroundBrush As LongPtr
    Dim BorderBrush As LongPtr

    Dim HBitmap As LongPtr

#Else

    Dim GdiToken As Long
    Dim Bitmap As Long
    Dim Graphics As Long

    Dim OuterPath As Long
    Dim InnerPath As Long

    Dim FillBrush As Long
    Dim BackgroundBrush As Long
    Dim BorderBrush As Long

    Dim HBitmap As Long

#End If

    Dim StartupInput As GdiplusStartupInput

    Dim PictureDescription As PICTDESC
    Dim PictureIID As GUID
    Dim Picture As IPicture

    Dim BaseWidth As Long
    Dim BaseHeight As Long

    Dim Width As Long
    Dim Height As Long

    Dim ShapeX As Long
    Dim ShapeY As Long
    Dim ShapeWidth As Long
    Dim ShapeHeight As Long

    Dim InnerX As Long
    Dim InnerY As Long
    Dim InnerWidth As Long
    Dim InnerHeight As Long

    Dim Radius As Long
    Dim InnerRadius As Long

    Dim Diameter As Long
    Dim InnerDiameter As Long

    Dim BorderSize As Long

    Dim Status As Long

    On Error GoTo CleanFail

    '--------------------------------------------------------
    ' Validate requested dimensions
    '--------------------------------------------------------

    BaseWidth = CLng(CellWidth)
    BaseHeight = CLng(CellHeight)

    If BaseWidth <= 0 Then Exit Function
    If BaseHeight <= 0 Then Exit Function

    '--------------------------------------------------------
    ' Validate render scale
    '--------------------------------------------------------

    If RenderScale < 1 Then
        RenderScale = 1
    End If

    '--------------------------------------------------------
    ' Validate shape
    '--------------------------------------------------------

    Select Case Shape
        Case DP_BUTTON_SHAPE_RECTANGLE
            ' Valid.
        Case DP_BUTTON_SHAPE_CIRCLE
            ' Valid.
        Case Else
            Shape = DP_BUTTON_SHAPE_RECTANGLE
    End Select

    '--------------------------------------------------------
    ' Supersampling
    '--------------------------------------------------------

    Width = BaseWidth * RenderScale
    Height = BaseHeight * RenderScale

    '--------------------------------------------------------
    ' Determine outer shape geometry
    '--------------------------------------------------------

    Select Case Shape

        Case DP_BUTTON_SHAPE_RECTANGLE

            ShapeX = 0
            ShapeY = 0
            ShapeWidth = Width
            ShapeHeight = Height

            ' Clamp radius to half of the smallest dimension.
            '------------------------------------------------

            Radius = CornerRadius

            If Radius < 0 Then
                Radius = 0
            End If

            If Radius > BaseWidth \ 2 Then
                Radius = BaseWidth \ 2
            End If

            If Radius > BaseHeight \ 2 Then
                Radius = BaseHeight \ 2
            End If

            Radius = Radius * RenderScale
            Diameter = Radius * 2

        Case DP_BUTTON_SHAPE_CIRCLE

            ' Use the smaller dimension for a true circle.
            '------------------------------------------------

            Diameter = Width

            If Height < Diameter Then
                Diameter = Height
            End If

            ShapeWidth = Diameter
            ShapeHeight = Diameter

            ShapeX = (Width - Diameter) \ 2
            ShapeY = (Height - Diameter) \ 2

    End Select

    '--------------------------------------------------------
    ' Border geometry
    '
    ' One logical pixel becomes RenderScale bitmap pixels
    '--------------------------------------------------------

    If BorderColor <> -1 Then

        BorderSize = BORDER_WIDTH * RenderScale

        InnerX = ShapeX + BorderSize
        InnerY = ShapeY + BorderSize

        InnerWidth = ShapeWidth - (BorderSize * 2)
        InnerHeight = ShapeHeight - (BorderSize * 2)

        If InnerWidth <= 0 Then
            InnerWidth = 0
        End If

        If InnerHeight <= 0 Then
            InnerHeight = 0
        End If

        ' Rounded rectangle inner radius.
        '
        ' Reducing the radius by the border width keeps the
        ' outer contour and inner contour geometrically
        ' consistent.
        '----------------------------------------------------

        InnerRadius = Radius - BorderSize

        If InnerRadius < 0 Then
            InnerRadius = 0
        End If

        ' Circle inner diameter.
        '----------------------------------------------------

        InnerDiameter = Diameter - (BorderSize * 2)

        If InnerDiameter < 0 Then
            InnerDiameter = 0
        End If

    End If

    '--------------------------------------------------------
    ' Start GDI+
    '--------------------------------------------------------

    StartupInput.GdiplusVersion = 1

    Status = GdiplusStartup(GdiToken, StartupInput, 0)

    If Status <> GDIP_OK Then
        Exit Function
    End If

    '--------------------------------------------------------
    ' Create bitmap
    '--------------------------------------------------------

    Status = GdipCreateBitmapFromScan0(Width, Height, _
                                       0, PIXEL_FORMAT_32BPP_ARGB, _
                                       0, Bitmap)

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    '--------------------------------------------------------
    ' Create graphics
    '--------------------------------------------------------

    Status = GdipGetImageGraphicsContext(Bitmap, Graphics)

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    '--------------------------------------------------------
    ' Anti-aliasing
    '--------------------------------------------------------

    Status = GdipSetSmoothingMode(Graphics, SMOOTHING_HIGH_QUALITY)

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    Status = GdipSetPixelOffsetMode(Graphics, PIXEL_OFFSET_HIGH_QUALITY)

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    '--------------------------------------------------------
    ' Paint complete bitmap background
    '--------------------------------------------------------

    Status = GdipCreateSolidFill(ColorToGdiPlus(BackgroundColor), _
                                 BackgroundBrush)

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    Status = GdipFillRectangleI(Graphics, BackgroundBrush, _
                                0, 0, Width, Height)

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    GdipDeleteBrush BackgroundBrush
    BackgroundBrush = 0

    '--------------------------------------------------------
    ' Create fill brush
    '--------------------------------------------------------

    Status = GdipCreateSolidFill(ColorToGdiPlus(FillColor), _
                                 FillBrush)

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    '--------------------------------------------------------
    ' Draw shape
    '
    ' No border:
    '   Existing rendering path is used.
    '
    ' Border:
    '   Outer shape is BorderColor.
    '   Inner shape is FillColor and inset by exactly
    '   one logical pixel.
    '--------------------------------------------------------

    If BorderColor = -1 Then

        '----------------------------------------------------
        ' NO BORDER
        '----------------------------------------------------

        Select Case Shape

            Case DP_BUTTON_SHAPE_RECTANGLE

                If Radius = 0 Then

                    Status = GdipFillRectangleI(Graphics, FillBrush, _
                                                ShapeX, ShapeY, _
                                                ShapeWidth, ShapeHeight)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                Else

                    Status = GdipCreatePath(FILL_MODE_WINDING, OuterPath)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Top-right
                    Status = GdipAddPathArcI(OuterPath, _
                                             Width - Diameter, 0, _
                                             Diameter, Diameter, _
                                             270!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Bottom-right
                    Status = GdipAddPathArcI(OuterPath, _
                                             Width - Diameter, Height - Diameter, _
                                             Diameter, Diameter, _
                                             0!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Bottom-left
                    Status = GdipAddPathArcI(OuterPath, _
                                             0, Height - Diameter, _
                                             Diameter, Diameter, _
                                             90!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Top-left
                    Status = GdipAddPathArcI(OuterPath, _
                                             0, 0, _
                                             Diameter, Diameter, _
                                             180!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    Status = GdipClosePathFigure(OuterPath)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    Status = GdipFillPath(Graphics, FillBrush, OuterPath)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                End If

            Case DP_BUTTON_SHAPE_CIRCLE

                Status = GdipFillEllipseI(Graphics, FillBrush, _
                                          ShapeX, ShapeY, _
                                          ShapeWidth, ShapeHeight)

                If Status <> GDIP_OK Then
                    GoTo CleanFail
                End If

        End Select

    Else

        '----------------------------------------------------
        ' BORDER
        '----------------------------------------------------

        ' Create border brush
        '----------------------------------------------------

        Status = GdipCreateSolidFill(ColorToGdiPlus(BorderColor), _
                                     BorderBrush)

        If Status <> GDIP_OK Then
            GoTo CleanFail
        End If

        ' Draw outer border shape
        '----------------------------------------------------

        Select Case Shape

            Case DP_BUTTON_SHAPE_RECTANGLE

                If Radius = 0 Then

                    Status = GdipFillRectangleI(Graphics, BorderBrush, _
                                                ShapeX, ShapeY, _
                                                ShapeWidth, ShapeHeight)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                Else

                    Status = GdipCreatePath(FILL_MODE_WINDING, OuterPath)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Top-right
                    Status = GdipAddPathArcI(OuterPath, _
                                             Width - Diameter, 0, _
                                             Diameter, Diameter, _
                                             270!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Bottom-right
                    Status = GdipAddPathArcI(OuterPath, _
                                             Width - Diameter, Height - Diameter, _
                                             Diameter, Diameter, _
                                             0!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Bottom-left
                    Status = GdipAddPathArcI(OuterPath, _
                                             0, Height - Diameter, _
                                             Diameter, Diameter, _
                                             90!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    ' Top-left
                    Status = GdipAddPathArcI(OuterPath, _
                                             0, 0, _
                                             Diameter, Diameter, _
                                             180!, 90!)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    Status = GdipClosePathFigure(OuterPath)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                    Status = GdipFillPath(Graphics, BorderBrush, OuterPath)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                End If

            Case DP_BUTTON_SHAPE_CIRCLE

                Status = GdipFillEllipseI(Graphics, BorderBrush, _
                                          ShapeX, ShapeY, _
                                          ShapeWidth, ShapeHeight)

                If Status <> GDIP_OK Then
                    GoTo CleanFail
                End If

        End Select

        ' Border brush no longer needed.
        '----------------------------------------------------

        GdipDeleteBrush BorderBrush
        BorderBrush = 0

        ' Draw inner FillColor shape.
        '
        ' This creates a true 1 logical pixel border.
        '----------------------------------------------------

        Select Case Shape

            Case DP_BUTTON_SHAPE_RECTANGLE

                If InnerWidth > 0 And InnerHeight > 0 Then

                    If InnerRadius = 0 Then

                        Status = GdipFillRectangleI(Graphics, FillBrush, _
                                                    InnerX, InnerY, _
                                                    InnerWidth, InnerHeight)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                    Else

                        Status = GdipCreatePath(FILL_MODE_WINDING, InnerPath)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                        InnerDiameter = InnerRadius * 2

                        ' Top-right
                        Status = GdipAddPathArcI(InnerPath, _
                                                 InnerX + InnerWidth - InnerDiameter, _
                                                 InnerY, _
                                                 InnerDiameter, InnerDiameter, _
                                                 270!, 90!)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                        ' Bottom-right
                        Status = GdipAddPathArcI(InnerPath, _
                                                 InnerX + InnerWidth - InnerDiameter, _
                                                 InnerY + InnerHeight - InnerDiameter, _
                                                 InnerDiameter, InnerDiameter, _
                                                 0!, 90!)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                        ' Bottom-left
                        Status = GdipAddPathArcI(InnerPath, _
                                                 InnerX, _
                                                 InnerY + InnerHeight - InnerDiameter, _
                                                 InnerDiameter, InnerDiameter, _
                                                 90!, 90!)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                        ' Top-left
                        Status = GdipAddPathArcI(InnerPath, _
                                                 InnerX, _
                                                 InnerY, _
                                                 InnerDiameter, InnerDiameter, _
                                                 180!, 90!)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                        Status = GdipClosePathFigure(InnerPath)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                        Status = GdipFillPath(Graphics, FillBrush, InnerPath)

                        If Status <> GDIP_OK Then
                            GoTo CleanFail
                        End If

                    End If

                End If

            Case DP_BUTTON_SHAPE_CIRCLE

                If InnerDiameter > 0 Then

                    Status = GdipFillEllipseI(Graphics, FillBrush, _
                                              ShapeX + BorderSize, ShapeY + BorderSize, _
                                              InnerDiameter, InnerDiameter)

                    If Status <> GDIP_OK Then
                        GoTo CleanFail
                    End If

                End If

        End Select

    End If

    '--------------------------------------------------------
    ' Release drawing objects
    '--------------------------------------------------------

    If FillBrush <> 0 Then
        GdipDeleteBrush FillBrush
        FillBrush = 0
    End If

    If BorderBrush <> 0 Then
        GdipDeleteBrush BorderBrush
        BorderBrush = 0
    End If

    If OuterPath <> 0 Then
        GdipDeletePath OuterPath
        OuterPath = 0
    End If

    If InnerPath <> 0 Then
        GdipDeletePath InnerPath
        InnerPath = 0
    End If

    If Graphics <> 0 Then
        GdipDeleteGraphics Graphics
        Graphics = 0
    End If

    '--------------------------------------------------------
    ' Convert GDI+ bitmap to HBITMAP
    '--------------------------------------------------------

    Status = GdipCreateHBITMAPFromBitmap(Bitmap, HBitmap, _
                                         ColorToGdiPlus(BackgroundColor))

    If Status <> GDIP_OK Then
        GoTo CleanFail
    End If

    '--------------------------------------------------------
    ' Dispose GDI+ bitmap
    '--------------------------------------------------------

    GdipDisposeImage Bitmap
    Bitmap = 0

    '--------------------------------------------------------
    ' Shutdown GDI+
    '--------------------------------------------------------

    GdiplusShutdown GdiToken
    GdiToken = 0

    '--------------------------------------------------------
    ' Build StdPicture descriptor
    '--------------------------------------------------------

    With PictureDescription
        .Size = LenB(PictureDescription)
        .Type = PICTYPE_BITMAP
        .hPic = HBitmap
        .hPal = 0
    End With

    '--------------------------------------------------------
    ' IID_IPicture
    '
    ' {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    '--------------------------------------------------------

    With PictureIID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    '--------------------------------------------------------
    ' Transfer HBITMAP ownership to StdPicture
    '--------------------------------------------------------

    Status = OleCreatePictureIndirect(PictureDescription, _
                                      PictureIID, _
                                      1, _
                                      Picture)

    If Status <> GDIP_OK Then
        DeleteObject HBitmap
        HBitmap = 0
        Exit Function
    End If

    Set DP_GetPicture = Picture

    ' StdPicture now owns the HBITMAP.
    HBitmap = 0

    Exit Function

'------------------------------------------------------------
' Cleanup
'------------------------------------------------------------

CleanFail:

    On Error Resume Next

    If FillBrush <> 0 Then
        GdipDeleteBrush FillBrush
        FillBrush = 0
    End If

    If BorderBrush <> 0 Then
        GdipDeleteBrush BorderBrush
        BorderBrush = 0
    End If

    If BackgroundBrush <> 0 Then
        GdipDeleteBrush BackgroundBrush
        BackgroundBrush = 0
    End If

    If OuterPath <> 0 Then
        GdipDeletePath OuterPath
        OuterPath = 0
    End If

    If InnerPath <> 0 Then
        GdipDeletePath InnerPath
        InnerPath = 0
    End If

    If Graphics <> 0 Then
        GdipDeleteGraphics Graphics
        Graphics = 0
    End If

    If Bitmap <> 0 Then
        GdipDisposeImage Bitmap
        Bitmap = 0
    End If

    If HBitmap <> 0 Then
        DeleteObject HBitmap
        HBitmap = 0
    End If

    If GdiToken <> 0 Then
        GdiplusShutdown GdiToken
        GdiToken = 0
    End If

End Function


'============================================================
' VBA COLORREF -> GDI+ ARGB
'
' VBA RGB: &H00BBGGRR
' GDI+ ARGB: &HAARRGGBB
'
' Alpha is FF because the bitmap is intentionally opaque.
'============================================================

Private Function ColorToGdiPlus( _
    ByVal Color As Long) As Long

    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long

    Red = Color And &HFF&
    Green = (Color \ &H100&) And &HFF&
    Blue = (Color \ &H10000) And &HFF&

    ColorToGdiPlus = &HFF000000 Or _
                     (Red * &H10000) Or _
                     (Green * &H100&) Or _
                     Blue

End Function

