Attribute VB_Name = "mGDIPlus"
Option Explicit

'========================================================================================
' AxTimeLineIC v2 - Motor Grafico GDI+ y Funciones de Renderizado Avanzado
' Proporciona: Doble bufer, anti-aliasing, sombras suaves, figuras vectoriales
' (incluyendo triangulos direccionales), gradientes y renderizado de imagenes PNG/JPG/BMP.
'========================================================================================

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type RECTL
    Left   As Long
    Top    As Long
    Width  As Long
    Height As Long
End Type

Public Type RECTS
    Left   As Single
    Top    As Single
    Width  As Single
    Height As Single
End Type

Public Type POINTL
    X As Long
    Y As Long
End Type

Public Type POINTS
    X As Single
    Y As Single
End Type

Public Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Private Type PicBmp
    Size     As Long
    Type     As Long
    hBmp     As Long
    hpal     As Long
    Reserved As Long
End Type

' GDI & User32 APIs
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

' Enums GDI+
Public Enum SmoothingMode
    SmoothingModeInvalid = -1
    SmoothingModeDefault = 0
    SmoothingModeHighSpeed = 1
    SmoothingModeHighQuality = 2
    SmoothingModeNone = 3
    SmoothingModeAntiAlias = 4
End Enum

Public Enum TextRenderingHint
    TextRenderingHintSystemDefault = 0
    TextRenderingHintSingleBitPerPixelGridFit = 1
    TextRenderingHintSingleBitPerPixel = 2
    TextRenderingHintAntiAliasGridFit = 3
    TextRenderingHintAntiAlias = 4
    TextRenderingHintClearTypeGridFit = 5
End Enum

Public Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Public Enum Unit
    UnitWorld = 0
    UnitDisplay = 1
    UnitPixel = 2
    UnitPoint = 3
    UnitInch = 4
    UnitDocument = 5
    UnitMillimeter = 6
End Enum

Public Enum StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
End Enum

Public Enum StringTrimming
    StringTrimmingNone = 0
    StringTrimmingCharacter = 1
    StringTrimmingWord = 2
    StringTrimmingEllipsisCharacter = 3
    StringTrimmingEllipsisWord = 4
    StringTrimmingEllipsisPath = 5
End Enum

Public Enum DashStyle
    DashStyleSolid = 0
    DashStyleDash = 1
    DashStyleDot = 2
    DashStyleDashDot = 3
    DashStyleDashDotDot = 4
    DashStyleCustom = 5
End Enum

' Startup / Shutdown
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

' Graphics
Public Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hdc As Long, ByRef graphics As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal graphics As Long) As Long
Public Declare Function GdipSetSmoothingMode Lib "gdiplus.dll" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As Long
Public Declare Function GdipSetTextRenderingHint Lib "gdiplus.dll" (ByVal graphics As Long, ByVal mode As TextRenderingHint) As Long

' Pens & Brushes
Public Declare Function GdipCreatePen1 Lib "gdiplus.dll" (ByVal color As Long, ByVal Width As Single, ByVal Unit As Unit, ByRef pen As Long) As Long
Public Declare Function GdipSetPenDashStyle Lib "gdiplus.dll" (ByVal pen As Long, ByVal DashStyle As DashStyle) As Long
Public Declare Function GdipDeletePen Lib "gdiplus.dll" (ByVal pen As Long) As Long
Public Declare Function GdipCreateSolidFill Lib "gdiplus.dll" (ByVal color As Long, ByRef brush As Long) As Long
Public Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "gdiplus.dll" (ByRef rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal angle As Single, ByVal isAngleScalable As Long, ByVal wrapMode As Long, ByRef lineGradient As Long) As Long
Public Declare Function GdipDeleteBrush Lib "gdiplus.dll" (ByVal brush As Long) As Long

' Paths
Public Declare Function GdipCreatePath Lib "gdiplus.dll" (ByVal brushMode As Long, ByRef path As Long) As Long
Public Declare Function GdipDeletePath Lib "gdiplus.dll" (ByVal path As Long) As Long
Public Declare Function GdipResetPath Lib "gdiplus.dll" (ByVal path As Long) As Long
Public Declare Function GdipAddPathLineI Lib "gdiplus.dll" (ByVal path As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GdipAddPathArcI Lib "gdiplus.dll" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Public Declare Function GdipAddPathPolygonI Lib "gdiplus.dll" (ByVal path As Long, ByRef points As Any, ByVal count As Long) As Long
Public Declare Function GdipAddPathEllipseI Lib "gdiplus.dll" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipAddPathString Lib "gdiplus.dll" (ByVal path As Long, ByVal stringPtr As Long, ByVal length As Long, ByVal family As Long, ByVal style As Long, ByVal emSize As Single, ByRef layoutRect As RECTS, ByVal format As Long) As Long
Public Declare Function GdipClosePathFigure Lib "gdiplus.dll" (ByVal path As Long) As Long
Public Declare Function GdipClosePathFigures Lib "gdiplus.dll" (ByVal path As Long) As Long
Public Declare Function GdipDrawPath Lib "gdiplus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal path As Long) As Long
Public Declare Function GdipFillPath Lib "gdiplus.dll" (ByVal graphics As Long, ByVal brush As Long, ByVal path As Long) As Long

' Primitives
Public Declare Function GdipDrawLineI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GdipDrawRectangleI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipFillRectangleI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipDrawEllipseI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipFillEllipseI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipDrawPolygonI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal pen As Long, ByRef points As Any, ByVal count As Long) As Long
Public Declare Function GdipFillPolygonI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal brush As Long, ByRef points As Any, ByVal count As Long, ByVal fillMode As Long) As Long

' Fonts & Text
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus.dll" (ByVal Name As Long, ByVal fontCollection As Long, ByRef fontFamily As Long) As Long
Public Declare Function GdipDeleteFontFamily Lib "gdiplus.dll" (ByVal fontFamily As Long) As Long
Public Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus.dll" (ByRef nativeFamily As Long) As Long
Public Declare Function GdipCreateStringFormat Lib "gdiplus.dll" (ByVal formatAttributes As Long, ByVal language As Integer, ByRef StringFormat As Long) As Long
Public Declare Function GdipDeleteStringFormat Lib "gdiplus.dll" (ByVal StringFormat As Long) As Long
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus.dll" (ByVal StringFormat As Long, ByVal align As StringAlignment) As Long
Public Declare Function GdipSetStringFormatLineAlign Lib "gdiplus.dll" (ByVal StringFormat As Long, ByVal align As StringAlignment) As Long
Public Declare Function GdipSetStringFormatTrimming Lib "gdiplus.dll" (ByVal StringFormat As Long, ByVal trimming As StringTrimming) As Long

' Images & Bitmaps
Public Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal filename As Long, ByRef image As Long) As Long
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hpal As Long, ByRef bitmap As Long) As Long
Public Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal image As Long, ByRef Width As Long) As Long
Public Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal image As Long, ByRef Height As Long) As Long
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal image As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Unit, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As Long) As Long
Public Declare Function GdipSetClipPath Lib "gdiplus.dll" (ByVal graphics As Long, ByVal path As Long, ByVal combineMode As Long) As Long
Public Declare Function GdipResetClip Lib "gdiplus.dll" (ByVal graphics As Long) As Long

Public Const IDC_HAND As Long = 32649
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90
Public Const CombineModeReplace As Long = 0

Private m_GdiPlusToken As Long
Private m_RefCount     As Long

' Inicializacion y Cierre de GDI+
Public Sub GDIP_Init()
    If m_RefCount = 0 Then
        Dim gsi As GDIPlusStartupInput
        gsi.GdiPlusVersion = 1
        Call GdiplusStartup(m_GdiPlusToken, gsi)
    End If
    m_RefCount = m_RefCount + 1
End Sub

Public Sub GDIP_Terminate()
    If m_RefCount > 0 Then
        m_RefCount = m_RefCount - 1
        If m_RefCount = 0 Then
            Call GdiplusShutdown(m_GdiPlusToken)
            m_GdiPlusToken = 0
        End If
    End If
End Sub

' Conversion de color OLE_COLOR a ARGB (GDI+)
Public Function ARGB(ByVal OleColor As OLE_COLOR, Optional ByVal Alpha As Byte = 255) As Long
    Dim rgbVal As Long
    If (OleColor And &H80000000) Then
        rgbVal = GetSysColor(OleColor And &HFF&)
    Else
        rgbVal = OleColor
    End If
    
    Dim R As Long, G As Long, B As Long
    R = rgbVal And &HFF&
    G = (rgbVal And &HFF00&) \ &H100&
    B = (rgbVal And &HFF0000) \ &H10000
    
    Dim aVal As Long
    aVal = Alpha
    
    If aVal >= 128 Then
        ARGB = ((aVal - 128) * &H1000000) Or (R * &H10000) Or (G * &H100&) Or B Or &H80000000
    Else
        ARGB = (aVal * &H1000000) Or (R * &H10000) Or (G * &H100&) Or B
    End If
End Function

' Mezclador de colores (Interpolacion para gradientes)
Public Function BlendColors(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Ratio As Single) As Long
    If Ratio <= 0 Then BlendColors = Color1: Exit Function
    If Ratio >= 1 Then BlendColors = Color2: Exit Function
    
    Dim R1 As Long, G1 As Long, B1 As Long, A1 As Long
    Dim R2 As Long, G2 As Long, B2 As Long, A2 As Long
    
    A1 = (Color1 And &HFF000000) \ &H1000000 And &HFF&
    R1 = (Color1 And &HFF0000) \ &H10000
    G1 = (Color1 And &HFF00&) \ &H100&
    B1 = Color1 And &HFF&
    
    A2 = (Color2 And &HFF000000) \ &H1000000 And &HFF&
    R2 = (Color2 And &HFF0000) \ &H10000
    G2 = (Color2 And &HFF00&) \ &H100&
    B2 = Color2 And &HFF&
    
    Dim R As Byte, G As Byte, B As Byte, A As Byte
    R = CByte(R1 + (R2 - R1) * Ratio)
    G = CByte(G1 + (G2 - G1) * Ratio)
    B = CByte(B1 + (B2 - B1) * Ratio)
    A = CByte(A1 + (A2 - A1) * Ratio)
    
    BlendColors = RGB(R, G, B)
End Function

' Factor de escala DPI
Public Function GetDPIScale() As Single
    Dim hdc As Long, lpx As Long
    hdc = GetDC(0)
    If hdc <> 0 Then
        lpx = GetDeviceCaps(hdc, LOGPIXELSX)
        Call ReleaseDC(0, hdc)
    End If
    If lpx <= 0 Then lpx = 96
    GetDPIScale = CSng(lpx) / 96!
End Function

' Camino para Rectangulo Redondeado
Public Function CreateRoundRectPath(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal Radius As Single) As Long
    Dim path As Long
    If GdipCreatePath(0, path) = 0 Then
        Dim maxR As Single
        maxR = IIf(Width < Height, Width / 2!, Height / 2!)
        If Radius > maxR Then Radius = maxR
        If Radius < 1! Then Radius = 1!
        
        Dim d As Single: d = Radius * 2!
        Call GdipAddPathArcI(path, X, Y, d, d, 180!, 90!)
        Call GdipAddPathArcI(path, X + Width - d, Y, d, d, 270!, 90!)
        Call GdipAddPathArcI(path, X + Width - d, Y + Height - d, d, d, 0!, 90!)
        Call GdipAddPathArcI(path, X, Y + Height - d, d, d, 90!, 90!)
        Call GdipClosePathFigure(path)
        CreateRoundRectPath = path
    End If
End Function

' Camino para Triangulo apuntando Hacia Arriba
Public Function CreateTriangleUpPath(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
    Dim path As Long
    If GdipCreatePath(0, path) = 0 Then
        Dim pts(0 To 2) As POINTL
        pts(0).X = X + (Width \ 2): pts(0).Y = Y
        pts(1).X = X + Width:       pts(1).Y = Y + Height
        pts(2).X = X:               pts(2).Y = Y + Height
        Call GdipAddPathPolygonI(path, pts(0), 3)
        Call GdipClosePathFigure(path)
        CreateTriangleUpPath = path
    End If
End Function

' Camino para Triangulo apuntando Hacia Abajo
Public Function CreateTriangleDownPath(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
    Dim path As Long
    If GdipCreatePath(0, path) = 0 Then
        Dim pts(0 To 2) As POINTL
        pts(0).X = X:               pts(0).Y = Y
        pts(1).X = X + Width:       pts(1).Y = Y
        pts(2).X = X + (Width \ 2): pts(2).Y = Y + Height
        Call GdipAddPathPolygonI(path, pts(0), 3)
        Call GdipClosePathFigure(path)
        CreateTriangleDownPath = path
    End If
End Function

' Camino para Diamante / Rombo
Public Function CreateDiamondPath(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
    Dim path As Long
    If GdipCreatePath(0, path) = 0 Then
        Dim pts(0 To 3) As POINTL
        pts(0).X = X + (Width \ 2): pts(0).Y = Y
        pts(1).X = X + Width:       pts(1).Y = Y + (Height \ 2)
        pts(2).X = X + (Width \ 2): pts(2).Y = Y + Height
        pts(3).X = X:               pts(3).Y = Y + (Height \ 2)
        Call GdipAddPathPolygonI(path, pts(0), 4)
        Call GdipClosePathFigure(path)
        CreateDiamondPath = path
    End If
End Function

' Sombra suave (Soft Drop Shadow) difusa para tarjetas y nodos
Public Sub DrawSoftShadow(ByVal graphics As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal Radius As Single, ByVal ShadowSize As Single, ByVal ShadowColor As Long, ByVal BaseAlpha As Byte)
    If ShadowSize <= 0 Then Exit Sub
    Dim i As Long, steps As Long: steps = CLng(ShadowSize)
    If steps < 2 Then steps = 2
    If steps > 8 Then steps = 8
    
    Dim curAlpha As Byte
    Dim hBrush As Long
    Dim hPath As Long
    Dim expand As Single, alphaStep As Single
    alphaStep = CSng(BaseAlpha) / CSng(steps * 1.5!)
    
    For i = steps To 1 Step -1
        expand = CSng(i) * 1.2!
        curAlpha = CByte(alphaStep * (1! - (CSng(i) / CSng(steps + 1))))
        If curAlpha > 0 Then
            hPath = CreateRoundRectPath(X - expand + 1!, Y - expand + 3!, Width + (expand * 2!), Height + (expand * 2!), Radius + expand)
            If hPath <> 0 Then
                Call GdipCreateSolidFill(ARGB(ShadowColor, curAlpha), hBrush)
                Call GdipFillPath(graphics, hBrush, hPath)
                Call GdipDeleteBrush(hBrush)
                Call GdipDeletePath(hPath)
            End If
        End If
    Next i
End Sub

' Renderizado de texto tipografico con anti-alias
Public Sub RenderText(ByVal graphics As Long, ByVal Text As String, ByVal fontObj As StdFont, ByRef rectBox As RECTS, ByVal TextColor As Long, Optional ByVal hAlign As StringAlignment = StringAlignmentNear, Optional ByVal vAlign As StringAlignment = StringAlignmentCenter, Optional ByVal Trimming As StringTrimming = StringTrimmingEllipsisWord)
    If Len(Text) = 0 Then Exit Sub
    
    Dim hFamily As Long, hFormat As Long, hBrush As Long
    Dim fStyle As Long, fSize As Single
    Dim hdcTmp As Long
    
    If fontObj.Bold Then fStyle = fStyle Or FontStyleBold
    If fontObj.Italic Then fStyle = fStyle Or FontStyleItalic
    If fontObj.Underline Then fStyle = fStyle Or FontStyleUnderline
    If fontObj.Strikethrough Then fStyle = fStyle Or FontStyleStrikeout
    
    hdcTmp = GetDC(0)
    fSize = CSng(MulDiv(fontObj.Size, GetDeviceCaps(hdcTmp, LOGPIXELSY), 72))
    Call ReleaseDC(0, hdcTmp)
    
    If GdipCreateFontFamilyFromName(StrPtr(fontObj.Name), 0, hFamily) <> 0 Then
        Call GdipGetGenericFontFamilySansSerif(hFamily)
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        Call GdipSetStringFormatAlign(hFormat, hAlign)
        Call GdipSetStringFormatLineAlign(hFormat, vAlign)
        Call GdipSetStringFormatTrimming(hFormat, Trimming)
    End If
    
    Dim hPath As Long
    If GdipCreatePath(0, hPath) = 0 Then
        Call GdipAddPathString(hPath, StrPtr(Text), -1, hFamily, fStyle, fSize, rectBox, hFormat)
        Call GdipCreateSolidFill(TextColor, hBrush)
        Call GdipFillPath(graphics, hBrush, hPath)
        
        Call GdipDeleteBrush(hBrush)
        Call GdipDeletePath(hPath)
    End If
    
    If hFormat <> 0 Then Call GdipDeleteStringFormat(hFormat)
    If hFamily <> 0 Then Call GdipDeleteFontFamily(hFamily)
End Sub

' Convierte codigo hexadecimal o texto a cadena UTF-16
Public Function IconCodeToGlyph(ByVal IconCode As Variant) As String
    Dim sCode As String, lCode As Long
    sCode = Trim$(CStr(IconCode))
    If Len(sCode) = 0 Then Exit Function
    
    sCode = UCase$(Replace(sCode, " ", vbNullString))
    sCode = UCase$(Replace(sCode, "U+", "&H"))
    If Left$(sCode, 2) <> "&H" And IsNumeric("&H" & sCode) Then
        sCode = "&H" & sCode
    End If
    
    On Error Resume Next
    lCode = CLng(sCode)
    If Err.Number = 0 And lCode > 0 Then
        Const POW10 As Long = 1024
        If lCode <= &HFFFF& Then
            IconCodeToGlyph = ChrW$(lCode)
        Else
            IconCodeToGlyph = ChrW$(&HD800& + (lCode And &HFFFF&) \ POW10) & _
                              ChrW$(&HDC00& + (lCode And (POW10 - 1)))
        End If
    Else
        IconCodeToGlyph = CStr(IconCode)
    End If
End Function

' Carga imagen GDI+ desde archivo o StdPicture
Public Function LoadGDIPlusImage(ByVal Source As Variant) As Long
    Dim hImg As Long
    On Error GoTo ErrH
    
    If VarType(Source) = vbString Then
        Dim sPath As String: sPath = CStr(Source)
        If Len(sPath) > 0 Then
            If GdipLoadImageFromFile(StrPtr(sPath), hImg) = 0 Then
                LoadGDIPlusImage = hImg
                Exit Function
            End If
        End If
    ElseIf VarType(Source) = vbObject Or VarType(Source) = vbDataObject Then
        If Not Source Is Nothing Then
            If TypeOf Source Is StdPicture Then
                Dim pic As StdPicture: Set pic = Source
                If pic.Type = vbPicTypeBitmap And pic.handle <> 0 Then
                    If GdipCreateBitmapFromHBITMAP(pic.handle, 0, hImg) = 0 Then
                        LoadGDIPlusImage = hImg
                        Exit Function
                    End If
                End If
            End If
        End If
    ElseIf IsNumeric(Source) Then
        Dim hBmp As Long: hBmp = CLng(Source)
        If hBmp <> 0 Then
            If GdipCreateBitmapFromHBITMAP(hBmp, 0, hImg) = 0 Then
                LoadGDIPlusImage = hImg
                Exit Function
            End If
        End If
    End If
ErrH:
    LoadGDIPlusImage = 0
End Function

' Renderiza una imagen recortada con clip circular (Avatar) o cuadrada
Public Sub DrawImageClipped(ByVal graphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstW As Long, ByVal dstH As Long, Optional ByVal CircularClip As Boolean = True)
    If hImage = 0 Then Exit Sub
    Dim imgW As Long, imgH As Long
    Call GdipGetImageWidth(hImage, imgW)
    Call GdipGetImageHeight(hImage, imgH)
    If imgW <= 0 Or imgH <= 0 Then Exit Sub
    
    If CircularClip Then
        Dim hPath As Long
        If GdipCreatePath(0, hPath) = 0 Then
            Call GdipAddPathEllipseI(hPath, dstX, dstY, dstW, dstH)
            Call GdipSetClipPath(graphics, hPath, CombineModeReplace)
            Call GdipDrawImageRectRectI(graphics, hImage, dstX, dstY, dstW, dstH, 0, 0, imgW, imgH, UnitPixel)
            Call GdipResetClip(graphics)
            Call GdipDeletePath(hPath)
        End If
    Else
        Call GdipDrawImageRectRectI(graphics, hImage, dstX, dstY, dstW, dstH, 0, 0, imgW, imgH, UnitPixel)
    End If
End Sub
