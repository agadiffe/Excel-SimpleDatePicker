Attribute VB_Name = "DP_modButtonPictureAPI"
Option Explicit

'============================================================
' DP_modButtonPictureAPI
'
' Creates opaque anti-aliased button pictures as StdPicture
' objects suitable for an MSForms.Image.Picture property.
'
' Supported shapes:
'
'   DP_BUTTON_SHAPE_RECTANGLE
'       Rounded rectangle.
'
'   DP_BUTTON_SHAPE_CIRCLE
'       True circle fitted inside the requested dimensions.
'
' A circle always uses the smaller of the requested width and
' height and is centered inside the bitmap.
'
' Rectangle:
'
'   CornerRadius = 0
'       Normal rectangle.
'
'   CornerRadius > 0
'       Rounded rectangle.
'
' The bitmap is rendered at 4x the requested dimensions.
'
' The bitmap is intentionally opaque:
'
'   BackgroundColor -> complete bitmap background
'   FillColor       -> button shape
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

Private Type DP_GdiplusStartupInput

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

Private Type DP_PICTDESC
    Size As Long
    Type As Long
    hPic As LongPtr
    hPal As LongPtr
End Type

#Else

Private Type DP_PICTDESC
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

#End If


'============================================================
' GUID
'============================================================

Private Type DP_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


'============================================================
' Constants
'============================================================

Private Const DP_GDIP_OK As Long = 0

Private Const DP_PICTYPE_BITMAP As Long = 1

Private Const DP_PIXEL_FORMAT_32BPP_ARGB As Long = &H26200A

' SmoothingModeHighQuality
Private Const DP_SMOOTHING_HIGH_QUALITY As Long = 2

' PixelOffsetModeHighQuality
Private Const DP_PIXEL_OFFSET_HIGH_QUALITY As Long = 2

' FillModeWinding
Private Const DP_FILL_MODE_WINDING As Long = 1

' Supersampling factor.
Private Const DP_RENDER_SCALE As Long = 4


'============================================================
' GDI+ declarations
'============================================================

#If VBA7 And Not Mac Then

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
    ByRef Token As LongPtr, _
    ByRef StartupInput As DP_GdiplusStartupInput, _
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
    ByRef PictureDescription As DP_PICTDESC, _
    ByRef ReferenceIID As DP_GUID, _
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
    ByRef StartupInput As DP_GdiplusStartupInput, _
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
    ByRef PictureDescription As DP_PICTDESC, _
    ByRef ReferenceIID As DP_GUID, _
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
' Creates a picture for any DatePicker button.
'
' Shape:
'
'   DP_BUTTON_SHAPE_RECTANGLE
'       Rounded rectangle.
'
'   DP_BUTTON_SHAPE_CIRCLE
'       True circle fitted inside the requested dimensions.
'
' CornerRadius:
'
'   Used only for DP_BUTTON_SHAPE_RECTANGLE.
'
'   Values below 6 may produce sharper corners,
'   appearing more like diagonal cuts than rounded corners.
'============================================================

Public Function DP_GetButtonPicture( _
    ByVal CellWidth As Single, _
    ByVal CellHeight As Single, _
    ByVal FillColor As Long, _
    ByVal BackgroundColor As Long, _
    Optional ByVal Shape As DP_ButtonShape = DP_BUTTON_SHAPE_RECTANGLE, _
    Optional ByVal CornerRadius As Long = 6) As StdPicture

#If Mac Then

    Exit Function

#ElseIf VBA7 Then

    Dim GdiToken As LongPtr
    Dim Bitmap As LongPtr
    Dim Graphics As LongPtr
    Dim Path As LongPtr
    Dim FillBrush As LongPtr
    Dim BackgroundBrush As LongPtr
    Dim HBitmap As LongPtr

#Else

    Dim GdiToken As Long
    Dim Bitmap As Long
    Dim Graphics As Long
    Dim Path As Long
    Dim FillBrush As Long
    Dim BackgroundBrush As Long
    Dim HBitmap As Long

#End If

    Dim StartupInput As DP_GdiplusStartupInput

    Dim PictureDescription As DP_PICTDESC
    Dim PictureIID As DP_GUID
    Dim Picture As IPicture

    Dim BaseWidth As Long
    Dim BaseHeight As Long

    Dim Width As Long
    Dim Height As Long

    Dim ShapeX As Long
    Dim ShapeY As Long
    Dim ShapeWidth As Long
    Dim ShapeHeight As Long

    Dim Radius As Long
    Dim Diameter As Long

    Dim Status As Long

    On Error GoTo CleanFail

    '========================================================
    ' Validate requested dimensions
    '========================================================

    BaseWidth = CLng(CellWidth)
    BaseHeight = CLng(CellHeight)

    If BaseWidth <= 0 Then Exit Function
    If BaseHeight <= 0 Then Exit Function

    '========================================================
    ' Validate shape
    '========================================================

    Select Case Shape
        Case DP_BUTTON_SHAPE_RECTANGLE
            ' Valid.
        Case DP_BUTTON_SHAPE_CIRCLE
            ' Valid.
        Case Else
            Shape = DP_BUTTON_SHAPE_RECTANGLE
    End Select

    '========================================================
    ' Supersampling
    '========================================================

    Width = BaseWidth * DP_RENDER_SCALE
    Height = BaseHeight * DP_RENDER_SCALE

    '========================================================
    ' Determine shape geometry
    '========================================================

    Select Case Shape

        Case DP_BUTTON_SHAPE_RECTANGLE

            ShapeX = 0
            ShapeY = 0
            ShapeWidth = Width
            ShapeHeight = Height

            '------------------------------------------------
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

            Radius = Radius * DP_RENDER_SCALE
            Diameter = Radius * 2

        Case DP_BUTTON_SHAPE_CIRCLE

            '------------------------------------------------
            ' Use the smaller dimension so the result is
            ' always a true circle.
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

    '========================================================
    ' Start GDI+
    '========================================================

    StartupInput.GdiplusVersion = 1

    Status = GdiplusStartup( _
        GdiToken, _
        StartupInput, _
        0)

    If Status <> DP_GDIP_OK Then
        Exit Function
    End If

    '========================================================
    ' Create bitmap
    '========================================================

    Status = GdipCreateBitmapFromScan0( _
        Width, _
        Height, _
        0, _
        DP_PIXEL_FORMAT_32BPP_ARGB, _
        0, _
        Bitmap)

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    '========================================================
    ' Create graphics
    '========================================================

    Status = GdipGetImageGraphicsContext( _
        Bitmap, _
        Graphics)

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    '========================================================
    ' Anti-aliasing
    '========================================================

    Status = GdipSetSmoothingMode( _
        Graphics, _
        DP_SMOOTHING_HIGH_QUALITY)

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    Status = GdipSetPixelOffsetMode( _
        Graphics, _
        DP_PIXEL_OFFSET_HIGH_QUALITY)

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    '========================================================
    ' Paint complete bitmap background
    '========================================================

    Status = GdipCreateSolidFill( _
        DP_ColorToGdiPlus(BackgroundColor), _
        BackgroundBrush)

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    Status = GdipFillRectangleI( _
        Graphics, _
        BackgroundBrush, _
        0, _
        0, _
        Width, _
        Height)

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    GdipDeleteBrush BackgroundBrush
    BackgroundBrush = 0

    '========================================================
    ' Create foreground brush
    '========================================================

    Status = GdipCreateSolidFill( _
        DP_ColorToGdiPlus(FillColor), _
        FillBrush)

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    '========================================================
    ' Draw requested shape
    '========================================================

    Select Case Shape

        Case DP_BUTTON_SHAPE_RECTANGLE

            '------------------------------------------------
            ' Radius = 0 -> normal rectangle.
            '------------------------------------------------

            If Radius = 0 Then

                Status = GdipFillRectangleI( _
                    Graphics, _
                    FillBrush, _
                    ShapeX, _
                    ShapeY, _
                    ShapeWidth, _
                    ShapeHeight)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

            Else

                '------------------------------------------------
                ' Rounded rectangle.
                '
                ' The complete boundary is made from four
                ' connected corner arcs. There are no separate
                ' line segments between the arcs, avoiding
                ' rasterization seams at the joins.
                '------------------------------------------------

                Status = GdipCreatePath( _
                    DP_FILL_MODE_WINDING, _
                    Path)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

                '------------------------------------------------
                ' Top-right corner (270 -> 360)
                '------------------------------------------------

                Status = GdipAddPathArcI( _
                    Path, _
                    Width - Diameter, _
                    0, _
                    Diameter, _
                    Diameter, _
                    270!, _
                    90!)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

                '------------------------------------------------
                ' Bottom-right corner (0 -> 90)
                '------------------------------------------------

                Status = GdipAddPathArcI( _
                    Path, _
                    Width - Diameter, _
                    Height - Diameter, _
                    Diameter, _
                    Diameter, _
                    0!, _
                    90!)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

                '------------------------------------------------
                ' Bottom-left corner (90 -> 180)
                '------------------------------------------------

                Status = GdipAddPathArcI( _
                    Path, _
                    0, _
                    Height - Diameter, _
                    Diameter, _
                    Diameter, _
                    90!, _
                    90!)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

                '------------------------------------------------
                ' Top-left corner (180 -> 270)
                '------------------------------------------------

                Status = GdipAddPathArcI( _
                    Path, _
                    0, _
                    0, _
                    Diameter, _
                    Diameter, _
                    180!, _
                    90!)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

                '------------------------------------------------
                ' Close the complete rounded rectangle.
                '------------------------------------------------

                Status = GdipClosePathFigure(Path)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

                '------------------------------------------------
                ' Fill once.
                '------------------------------------------------

                Status = GdipFillPath( _
                    Graphics, _
                    FillBrush, _
                    Path)

                If Status <> DP_GDIP_OK Then
                    GoTo CleanFail
                End If

            End If

        Case DP_BUTTON_SHAPE_CIRCLE

            '------------------------------------------------
            ' Native GDI+ ellipse.
            '
            ' ShapeWidth and ShapeHeight are equal, so this
            ' produces a true circle.
            '------------------------------------------------

            Status = GdipFillEllipseI( _
                Graphics, _
                FillBrush, _
                ShapeX, _
                ShapeY, _
                ShapeWidth, _
                ShapeHeight)

            If Status <> DP_GDIP_OK Then
                GoTo CleanFail
            End If

    End Select

    '========================================================
    ' Release drawing objects
    '========================================================

    If FillBrush <> 0 Then
        GdipDeleteBrush FillBrush
        FillBrush = 0
    End If

    If Path <> 0 Then
        GdipDeletePath Path
        Path = 0
    End If

    If Graphics <> 0 Then
        GdipDeleteGraphics Graphics
        Graphics = 0
    End If

    '========================================================
    ' Convert GDI+ bitmap to HBITMAP
    '========================================================

    Status = GdipCreateHBITMAPFromBitmap( _
        Bitmap, _
        HBitmap, _
        DP_ColorToGdiPlus(BackgroundColor))

    If Status <> DP_GDIP_OK Then
        GoTo CleanFail
    End If

    '========================================================
    ' Dispose GDI+ bitmap
    '========================================================

    GdipDisposeImage Bitmap
    Bitmap = 0

    '========================================================
    ' Shutdown GDI+
    '========================================================

    GdiplusShutdown GdiToken
    GdiToken = 0

    '========================================================
    ' Build StdPicture descriptor
    '========================================================

    With PictureDescription
        .Size = LenB(PictureDescription)
        .Type = DP_PICTYPE_BITMAP
        .hPic = HBitmap
        .hPal = 0
    End With

    '========================================================
    ' IID_IPicture
    '
    ' {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    '========================================================

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

    '========================================================
    ' Transfer HBITMAP ownership to StdPicture
    '========================================================

    Status = OleCreatePictureIndirect( _
        PictureDescription, _
        PictureIID, _
        1, _
        Picture)

    If Status <> DP_GDIP_OK Then
        DeleteObject HBitmap
        HBitmap = 0
        Exit Function
    End If

    Set DP_GetButtonPicture = Picture

    ' StdPicture now owns the HBITMAP.
    HBitmap = 0

    Exit Function

'============================================================
' Cleanup
'============================================================

CleanFail:

    On Error Resume Next

    If FillBrush <> 0 Then
        GdipDeleteBrush FillBrush
        FillBrush = 0
    End If

    If BackgroundBrush <> 0 Then
        GdipDeleteBrush BackgroundBrush
        BackgroundBrush = 0
    End If

    If Path <> 0 Then
        GdipDeletePath Path
        Path = 0
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
' VBA RGB: &H00BBGGR
' GDI+ ARGB: &HAARRGGBB
'
' Alpha is FF because the bitmap is intentionally opaque.
'============================================================

Private Function DP_ColorToGdiPlus( _
    ByVal Color As Long) As Long

    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long

    Red = Color And &HFF&
    Green = (Color \ &H100&) And &HFF&
    Blue = (Color \ &H10000) And &HFF&

    DP_ColorToGdiPlus = _
        &HFF000000 Or _
        (Red * &H10000) Or _
        (Green * &H100&) Or _
        Blue

End Function

