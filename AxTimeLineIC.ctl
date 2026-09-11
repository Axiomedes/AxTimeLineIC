VERSION 5.00
Begin VB.UserControl AxTimeLineIC 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   ToolboxBitmap   =   "AxTimeLineIC.ctx":0000
   Begin vb6projectAxTimeLineIC.ucScrollbar ucScroll 
      Height          =   3315
      Left            =   6270
      TabIndex        =   0
      Top             =   675
      Width           =   120
      _ExtentX        =   212
      _ExtentY        =   5847
   End
End
Attribute VB_Name = "AxTimeLineIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================================
' AxTimeLineIC (Evolucion de AxTimeLine) - Control Grafico de Linea de Tiempo para Visual Basic 6
' Caracteristicas:
'  - Renderizado vectorial suave con GDI+ (Anti-Aliasing).
'  - Doble bufer en memoria (DIB Section) - CERO PARPADEO (Flicker-Free).
'  - Soporte de Imagenes en Nodos (PNG con canal alfa, JPG, BMP, StdPicture y Avatares).
'  - Tarjetas de contenido (Cards) con Sombras Suaves (Soft Drop Shadows).
'  - Triangulos direccionales (Hacia Arriba / Hacia Abajo), Diamantes, Circulos y Pildoras.
'  - Modos: Vertical, Horizontal y Vertical Alternado (Zig-Zag / Split Timeline).
'  - Estados de progreso (Completed, In-Progress, Pending, Warning, Failed).
'  - Hit-Testing milimetrico y eventos ricos (ItemClick, ItemHover, ItemDblClick).
'========================================================================================

Public Enum tlOrientation
    tlVertical = 0
    tlHorizontal = 1
    tlVerticalAlternating = 2
End Enum

Public Enum tlLineStyle
    tlLineSolid = 0
    tlLineDashed = 1
    tlLineDotted = 2
    tlLineGradient = 3
End Enum

Public Enum tlTheme
    tlThemeCustom = 0
    tlThemeLight = 1
    tlThemeDark = 2
    tlThemeCyberpunk = 3
    tlThemeCorporate = 4
End Enum

Public Event Click()
Public Event DblClick()
Public Event ItemClick(ByVal Index As Long, ByVal Point As clsTimePoint)
Public Event ItemDblClick(ByVal Index As Long, ByVal Point As clsTimePoint)
Public Event ItemHover(ByVal Index As Long, ByVal Point As clsTimePoint)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Variables Privadas
Private m_Points            As clsTimePoints
Private m_ActiveSection     As Long
Private m_HoverIndex        As Long

' Layout & Orientacion
Private m_Orientation       As tlOrientation
Private m_SectionSpace      As Long
Private m_NodeSize          As Long
Private m_NodeShape         As tlNodeShape
Private m_CornerCurve       As Long

' Colores del Control
Private m_BackColor         As OLE_COLOR
Private m_Theme             As tlTheme
Private m_PointBackColor    As OLE_COLOR
Private m_BorderColor       As OLE_COLOR
Private m_BorderColorActive As OLE_COLOR
Private m_BorderWidth       As Long

' Linea Conectora
Private m_LineStyle         As tlLineStyle
Private m_LineWidth         As Long
Private m_LineColor         As OLE_COLOR
Private m_LineProgressColor As OLE_COLOR
Private m_LineDistance      As Long

' Tarjetas (Cards)
Private m_CardVisible       As Boolean
Private m_CardBackColor    As OLE_COLOR
Private m_CardBorderColor  As OLE_COLOR
Private m_CardBorderWidth  As Long
Private m_CardRadius       As Long
Private m_CardShadowVisible As Boolean
Private m_CardShadowColor  As OLE_COLOR
Private m_CardShadowBlur   As Long
Private m_CardShadowAlpha  As Byte
Private m_CalloutPointer    As Boolean

' Tipografias y Colores de Texto
Private m_Font1             As StdFont
Private m_Font2             As StdFont
Private m_FontDate          As StdFont
Private m_IconFont          As StdFont
Private m_ForeColor1        As OLE_COLOR
Private m_ForeColor2        As OLE_COLOR
Private m_IconForeColor     As OLE_COLOR
Private m_DateForeColor     As OLE_COLOR

Private m_TimeVisible       As Boolean
Private m_DateVisible       As Boolean

' Doble Bufer
Private m_hMemDC            As Long
Private m_hDIBBitmap        As Long
Private m_hOldBmp           As Long
Private m_BufWidth          As Long
Private m_BufHeight         As Long
Private m_nScale            As Single

Private m_hHandCursor       As Long

' Rectangulos calculados en tiempo de renderizado para Hit-Testing

Private Sub UserControl_Initialize()
    Call GDIP_Init
    m_nScale = GetDPIScale()
    m_hHandCursor = LoadCursor(0, IDC_HAND)
    m_HoverIndex = 0
    Set m_Points = New clsTimePoints
    
    Set m_Font1 = New StdFont: m_Font1.Name = "Segoe UI": m_Font1.Size = 10: m_Font1.Bold = True
    Set m_Font2 = New StdFont: m_Font2.Name = "Segoe UI": m_Font2.Size = 9
    Set m_FontDate = New StdFont: m_FontDate.Name = "Segoe UI": m_FontDate.Size = 8: m_FontDate.Italic = True
    Set m_IconFont = New StdFont: m_IconFont.Name = "Segoe MDL2 Assets": m_IconFont.Size = 12
    
    ' Valores por defecto modernos
    m_Orientation = tlVertical
    m_SectionSpace = 90
    m_NodeSize = 38
    m_NodeShape = tlShapeCircle
    m_CornerCurve = 10
    
    m_BackColor = &H1E1E1E
    m_PointBackColor = &H2D2D30
    m_BorderColor = &H555555
    m_BorderColorActive = &HFF9900
    m_BorderWidth = 2
    
    m_LineStyle = tlLineGradient
    m_LineWidth = 3
    m_LineColor = &H444444
    m_LineProgressColor = &HFF9900
    m_LineDistance = 10
    
    m_CardVisible = True
    m_CardBackColor = &H252526
    m_CardBorderColor = &H3E3E42
    m_CardBorderWidth = 1
    m_CardRadius = 8
    m_CardShadowVisible = True
    m_CardShadowColor = &H0
    m_CardShadowBlur = 5
    m_CardShadowAlpha = 90
    m_CalloutPointer = True
    
    m_ForeColor1 = &HFFFFFF
    m_ForeColor2 = &HCCCCCC
    m_IconForeColor = &HFFFFFF
    m_DateForeColor = &H999999
    
    m_DateVisible = True
    m_TimeVisible = True
    m_ActiveSection = 1
    m_Theme = tlThemeCustom
End Sub

Private Sub UserControl_InitProperties()
    ApplyTheme tlThemeLight
End Sub

Private Sub UserControl_Terminate()
    DestroyBackBuffer
    Set m_Points = Nothing
    Call GDIP_Terminate
End Sub

Private Sub CreateBackBuffer(ByVal w As Long, ByVal h As Long)
    If w <= 0 Or h <= 0 Then Exit Sub
    If m_hMemDC <> 0 And m_BufWidth = w And m_BufHeight = h Then Exit Sub
    
    DestroyBackBuffer
    
    Dim hdcScreen As Long
    hdcScreen = GetDC(0)
    m_hMemDC = CreateCompatibleDC(hdcScreen)
    
    Dim bi As BITMAPINFO
    With bi.bmiHeader
        .biSize = Len(bi.bmiHeader)
        .biWidth = w
        .biHeight = -h
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
    End With
    
    Dim pBits As Long
    m_hDIBBitmap = CreateDIBSection(m_hMemDC, bi, 0, pBits, 0, 0)
    m_hOldBmp = SelectObject(m_hMemDC, m_hDIBBitmap)
    m_BufWidth = w
    m_BufHeight = h
    
    Call ReleaseDC(0, hdcScreen)
End Sub

Private Sub DestroyBackBuffer()
    If m_hMemDC <> 0 Then
        If m_hOldBmp <> 0 Then Call SelectObject(m_hMemDC, m_hOldBmp): m_hOldBmp = 0
        If m_hDIBBitmap <> 0 Then Call DeleteObject(m_hDIBBitmap): m_hDIBBitmap = 0
        Call DeleteDC(m_hMemDC)
        m_hMemDC = 0
    End If
    m_BufWidth = 0: m_BufHeight = 0
End Sub

Public Sub Refresh()
    If UserControl.ScaleWidth <= 0 Or UserControl.ScaleHeight <= 0 Then Exit Sub
    
    Dim totalPoints As Long: totalPoints = m_Points.Count
    Dim totalLen As Long
    
    If m_Orientation = tlHorizontal Then
        totalLen = 60 + (m_SectionSpace * totalPoints) + 60
        With ucScroll
            .Orientation = sbHorizontal
            .Max = totalLen - UserControl.ScaleWidth
            .SmallChange = m_SectionSpace \ 2
            .LargeChange = m_SectionSpace
            If .Max > 0 Then
                .Visible = True
                .ZOrder 0
            Else
                .Visible = False
                .Value = 0
            End If
        End With
    Else
        totalLen = 40 + (m_SectionSpace * totalPoints) + 60
        With ucScroll
            .Orientation = sbVertical
            .Max = totalLen - UserControl.ScaleHeight
            .SmallChange = m_SectionSpace \ 2
            .LargeChange = m_SectionSpace
            If .Max > 0 Then
                .Visible = True
                .ZOrder 0
            Else
                .Visible = False
                .Value = 0
            End If
        End With
    End If
    
    DrawScene
End Sub

Private Sub DrawScene()
    Dim w As Long, h As Long
    w = UserControl.ScaleWidth
    h = UserControl.ScaleHeight
    If w <= 0 Or h <= 0 Then Exit Sub
    
    CreateBackBuffer w, h
    If m_hMemDC = 0 Then Exit Sub
    
    Dim G As Long
    If GdipCreateFromHDC(m_hMemDC, G) <> 0 Then Exit Sub
    
    Call GdipSetSmoothingMode(G, SmoothingModeAntiAlias)
    Call GdipSetTextRenderingHint(G, TextRenderingHintAntiAliasGridFit)
    
    ' 1. Fondo del Control
    Dim hBgBrush As Long
    Call GdipCreateSolidFill(ARGB(m_BackColor, 255), hBgBrush)
    Call GdipFillRectangleI(G, hBgBrush, 0, 0, w, h)
    Call GdipDeleteBrush(hBgBrush)
    
    ' 2. Renderizado de Elementos
    Dim scrollVal As Long: scrollVal = ucScroll.Value
    
    Select Case m_Orientation
        Case tlVertical
            DrawVerticalTimeline G, w, h, scrollVal, False
        Case tlHorizontal
            DrawHorizontalTimeline G, w, h, scrollVal
        Case tlVerticalAlternating
            DrawVerticalTimeline G, w, h, scrollVal, True
    End Select
    
    Call GdipDeleteGraphics(G)
    
    ' 3. Volcado al DC de pantalla sin parpadeo
    Call BitBlt(UserControl.hdc, 0, 0, w, h, m_hMemDC, 0, 0, &HCC0020)
    UserControl.Refresh
End Sub

Private Sub DrawVerticalTimeline(ByVal G As Long, ByVal viewW As Long, ByVal viewH As Long, ByVal scrollVal As Long, ByVal Alternating As Boolean)
    Dim Count As Long: Count = m_Points.Count
    If Count = 0 Then Exit Sub
    
    Dim i As Long, startY As Long, endY As Long
    Dim nodeCenterX As Long
    Dim dateColWidth As Long
    
    If m_DateVisible Or m_TimeVisible Then
        dateColWidth = 85 * m_nScale
    Else
        dateColWidth = 10 * m_nScale
    End If
    
    If Alternating Then
        nodeCenterX = viewW \ 2
    Else
        nodeCenterX = dateColWidth + (m_NodeSize \ 2) + 15
    End If
    
    startY = 40 - scrollVal
    endY = startY + ((Count - 1) * m_SectionSpace)
    
    DrawConnectorLine G, nodeCenterX, startY, nodeCenterX, endY, True
    
    Dim curY As Long
    For i = 1 To Count
        curY = startY + ((i - 1) * m_SectionSpace)
        If curY >= -m_SectionSpace And curY <= viewH + m_SectionSpace Then
            Dim pt As clsTimePoint
            Set pt = m_Points.Item(i)
            If pt.Visible Then
                Dim isLeft As Boolean
                If Alternating Then
                    isLeft = (i Mod 2 <> 0)
                Else
                    isLeft = False
                End If
                DrawVerticalNodeItem G, pt, i, nodeCenterX, curY, viewW, isLeft, Alternating
            End If
        End If
    Next i
End Sub

Private Sub DrawHorizontalTimeline(ByVal G As Long, ByVal viewW As Long, ByVal viewH As Long, ByVal scrollVal As Long)
    Dim Count As Long: Count = m_Points.Count
    If Count = 0 Then Exit Sub
    
    Dim i As Long, startX As Long, endX As Long
    Dim nodeCenterY As Long
    nodeCenterY = 60 * m_nScale
    
    startX = 60 - scrollVal
    endX = startX + ((Count - 1) * m_SectionSpace)
    
    DrawConnectorLine G, startX, nodeCenterY, endX, nodeCenterY, False
    
    Dim curX As Long
    For i = 1 To Count
        curX = startX + ((i - 1) * m_SectionSpace)
        If curX >= -m_SectionSpace And curX <= viewW + m_SectionSpace Then
            Dim pt As clsTimePoint
            Set pt = m_Points.Item(i)
            If pt.Visible Then
                DrawHorizontalNodeItem G, pt, i, curX, nodeCenterY, viewH
            End If
        End If
    Next i
End Sub

Private Sub DrawConnectorLine(ByVal G As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal IsVertical As Boolean)
    If m_Points.Count <= 1 Then Exit Sub
    
    Dim Count As Long: Count = m_Points.Count
    Dim i As Long
    Dim segX1 As Long, segY1 As Long, segX2 As Long, segY2 As Long
    Dim activeIdx As Long: activeIdx = m_ActiveSection
    
    For i = 1 To Count - 1
        If IsVertical Then
            segX1 = X1: segX2 = X2
            segY1 = Y1 + ((i - 1) * m_SectionSpace)
            segY2 = Y1 + (i * m_SectionSpace)
        Else
            segY1 = Y1: segY2 = Y2
            segX1 = X1 + ((i - 1) * m_SectionSpace)
            segX2 = X1 + (i * m_SectionSpace)
        End If
        
        Dim segColor As Long
        If i < activeIdx Then
            segColor = m_LineProgressColor
        ElseIf i = activeIdx Then
            segColor = m_LineProgressColor
        Else
            segColor = m_LineColor
        End If
        
        Dim hPen As Long
        Call GdipCreatePen1(ARGB(segColor, 255), m_LineWidth * m_nScale, UnitPixel, hPen)
        
        Select Case m_LineStyle
            Case tlLineDashed: Call GdipSetPenDashStyle(hPen, DashStyleDash)
            Case tlLineDotted: Call GdipSetPenDashStyle(hPen, DashStyleDot)
        End Select
        
        Call GdipDrawLineI(G, hPen, segX1, segY1, segX2, segY2)
        Call GdipDeletePen(hPen)
    Next i
End Sub

Private Sub DrawVerticalNodeItem(ByVal G As Long, ByVal pt As clsTimePoint, ByVal Index As Long, ByVal nodeCenterX As Long, ByVal nodeCenterY As Long, ByVal viewW As Long, ByVal isCardLeft As Boolean, ByVal isAlternating As Boolean)
    Dim nodeR As Long: nodeR = m_NodeSize \ 2
    Dim nodeX As Long: nodeX = nodeCenterX - nodeR
    Dim nodeY As Long: nodeY = nodeCenterY - nodeR
    
    Call pt.SetNodeRect(nodeX, nodeY, m_NodeSize, m_NodeSize)
    
    Dim isActive As Boolean: isActive = (Index = m_ActiveSection)
    Dim isHover As Boolean: isHover = (Index = m_HoverIndex)
    
    If m_CardShadowVisible Then
        DrawSoftShadow G, nodeX, nodeY, m_NodeSize, m_NodeSize, m_NodeSize \ 2, 3, 0, 80
    End If
    
    If isActive Then
        Dim hGlowPath As Long, hGlowBrush As Long
        hGlowPath = CreateRoundRectPath(nodeX - 4, nodeY - 4, m_NodeSize + 8, m_NodeSize + 8, (m_NodeSize + 8) \ 2)
        If hGlowPath <> 0 Then
            Call GdipCreateSolidFill(ARGB(m_BorderColorActive, 60), hGlowBrush)
            Call GdipFillPath(G, hGlowBrush, hGlowPath)
            Call GdipDeleteBrush(hGlowBrush)
            Call GdipDeletePath(hGlowPath)
        End If
    End If
    
    RenderNodeShape G, pt, nodeX, nodeY, m_NodeSize, m_NodeSize, isActive, isHover
    
    If Not isAlternating And (m_DateVisible Or m_TimeVisible) Then
        Dim rDateBox As RECTS
        rDateBox.Left = 5
        rDateBox.Top = nodeCenterY - 18
        rDateBox.Width = nodeX - 15
        rDateBox.Height = 36
        
        If m_DateVisible And Len(pt.DateText) > 0 Then
            Dim rD As RECTS: rD = rDateBox: rD.Height = 18
            RenderText G, pt.DateText, m_FontDate, rD, ARGB(m_DateForeColor, 220), StringAlignmentFar, StringAlignmentCenter
        End If
        If m_TimeVisible And Len(pt.Timestamp) > 0 Then
            Dim rT As RECTS: rT = rDateBox: rT.Top = rDateBox.Top + 16: rT.Height = 18
            RenderText G, pt.Timestamp, m_FontDate, rT, ARGB(m_DateForeColor, 160), StringAlignmentFar, StringAlignmentCenter
        End If
    End If
    
    If m_CardVisible Then
        Dim cardX As Long, cardY As Long, cardW As Long, cardH As Long
        cardH = m_SectionSpace - 16
        If cardH < 50 Then cardH = 50
        cardY = nodeCenterY - (cardH \ 2)
        
        If isCardLeft Then
            cardW = nodeX - 30
            cardX = 15
        Else
            cardX = nodeX + m_NodeSize + 18
            cardW = viewW - cardX - 25
        End If
        
        If cardW > 40 Then
            Call pt.SetCardRect(cardX, cardY, cardW, cardH)
            
            RenderCard G, pt, cardX, cardY, cardW, cardH, nodeCenterX, nodeCenterY, isCardLeft, isActive, isHover
        End If
    End If
End Sub

Private Sub DrawHorizontalNodeItem(ByVal G As Long, ByVal pt As clsTimePoint, ByVal Index As Long, ByVal nodeCenterX As Long, ByVal nodeCenterY As Long, ByVal viewH As Long)
    Dim nodeR As Long: nodeR = m_NodeSize \ 2
    Dim nodeX As Long: nodeX = nodeCenterX - nodeR
    Dim nodeY As Long: nodeY = nodeCenterY - nodeR
    
    Call pt.SetNodeRect(nodeX, nodeY, m_NodeSize, m_NodeSize)
    
    Dim isActive As Boolean: isActive = (Index = m_ActiveSection)
    Dim isHover As Boolean: isHover = (Index = m_HoverIndex)
    
    If m_CardShadowVisible Then
        DrawSoftShadow G, nodeX, nodeY, m_NodeSize, m_NodeSize, m_NodeSize \ 2, 3, 0, 80
    End If
    
    If isActive Then
        Dim hGlowPath As Long, hGlowBrush As Long
        hGlowPath = CreateRoundRectPath(nodeX - 4, nodeY - 4, m_NodeSize + 8, m_NodeSize + 8, (m_NodeSize + 8) \ 2)
        If hGlowPath <> 0 Then
            Call GdipCreateSolidFill(ARGB(m_BorderColorActive, 60), hGlowBrush)
            Call GdipFillPath(G, hGlowBrush, hGlowPath)
            Call GdipDeleteBrush(hGlowBrush)
            Call GdipDeletePath(hGlowPath)
        End If
    End If
    
    RenderNodeShape G, pt, nodeX, nodeY, m_NodeSize, m_NodeSize, isActive, isHover
    
    If m_CardVisible Then
        Dim cardX As Long, cardY As Long, cardW As Long, cardH As Long
        cardW = m_SectionSpace - 16
        If cardW < 70 Then cardW = 70
        cardX = nodeCenterX - (cardW \ 2)
        cardY = nodeY + m_NodeSize + 16
        cardH = viewH - cardY - 20
        If cardH > 40 Then
            Call pt.SetCardRect(cardX, cardY, cardW, cardH)
            
            RenderCard G, pt, cardX, cardY, cardW, cardH, nodeCenterX, nodeCenterY, False, isActive, isHover
        End If
    End If
End Sub

Private Sub RenderNodeShape(ByVal G As Long, ByVal pt As clsTimePoint, ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal h As Long, ByVal isActive As Boolean, ByVal isHover As Boolean)
    Dim effectiveShape As tlNodeShape
    effectiveShape = IIf(pt.NodeShape = tlNodeShape.tlShapeDefault, m_NodeShape, pt.NodeShape)
    
    Dim nodeBg As OLE_COLOR, nodeBorder As OLE_COLOR
    nodeBg = IIf(pt.NodeColor <> -1, pt.NodeColor, m_PointBackColor)
    
    If isActive Then
        nodeBorder = m_BorderColorActive
    ElseIf isHover Then
        nodeBorder = BlendColors(m_BorderColor, m_BorderColorActive, 0.6)
    Else
        nodeBorder = IIf(pt.NodeBorderColor <> -1, pt.NodeBorderColor, m_BorderColor)
    End If
    
    Select Case pt.Status
        Case tlStatusCompleted: nodeBorder = &H43A047: nodeBg = &H2E7D32
        Case tlStatusWarning:   nodeBorder = &H20B2AA: nodeBg = &H388E3C
        Case tlStatusFailed:    nodeBorder = &H2E2EFF: nodeBg = &H1E1EB8
    End Select
    
    Dim hPath As Long
    Select Case effectiveShape
        Case tlShapeCircle
            hPath = CreateRoundRectPath(X, Y, w, h, w \ 2)
        Case tlShapeRoundedRect
            hPath = CreateRoundRectPath(X, Y, w, h, m_CornerCurve)
        Case tlShapeSquare
            hPath = CreateRoundRectPath(X, Y, w, h, 0)
        Case tlShapePill
            hPath = CreateRoundRectPath(X, Y, w, h, h \ 2)
        Case tlShapeDiamond
            hPath = CreateDiamondPath(X, Y, w, h)
        Case tlShapeTriangleUp
            hPath = CreateTriangleUpPath(X, Y, w, h)
        Case tlShapeTriangleDown
            hPath = CreateTriangleDownPath(X, Y, w, h)
        Case Else
            hPath = CreateRoundRectPath(X, Y, w, h, w \ 2)
    End Select
    
    If hPath <> 0 Then
        Dim hBrush As Long, hPen As Long
        Call GdipCreateSolidFill(ARGB(nodeBg, 255), hBrush)
        Call GdipFillPath(G, hBrush, hPath)
        Call GdipDeleteBrush(hBrush)
        
        If pt.hGdiImage <> 0 Then
            Dim isCircleClip As Boolean
            isCircleClip = (effectiveShape = tlShapeCircle Or pt.ImageCircular)
            DrawImageClipped G, pt.hGdiImage, X + 3, Y + 3, w - 6, h - 6, isCircleClip
        ElseIf Len(pt.IconChar) > 0 Then
            Dim rIco As RECTS
            rIco.Left = X: rIco.Top = Y: rIco.Width = w: rIco.Height = h
            Dim glyphStr As String
            glyphStr = IconCodeToGlyph(pt.IconChar)
            Dim fIco As StdFont
            Set fIco = IIf(Not pt.IconFont Is Nothing, pt.IconFont, m_IconFont)
            RenderText G, glyphStr, fIco, rIco, ARGB(m_IconForeColor, 255), StringAlignmentCenter, StringAlignmentCenter
        Else
            If pt.Status = tlStatusCompleted Then
                Dim rChk As RECTS: rChk.Left = X: rChk.Top = Y: rChk.Width = w: rChk.Height = h
                RenderText G, ChrW$(&H2713), m_Font1, rChk, ARGB(&HFFFFFF, 255), StringAlignmentCenter, StringAlignmentCenter
            End If
        End If
        
        Call GdipCreatePen1(ARGB(nodeBorder, 255), m_BorderWidth * m_nScale, UnitPixel, hPen)
        Call GdipDrawPath(G, hPen, hPath)
        Call GdipDeletePen(hPen)
        
        Call GdipDeletePath(hPath)
    End If
End Sub

Private Sub RenderCard(ByVal G As Long, ByVal pt As clsTimePoint, ByVal cardX As Long, ByVal cardY As Long, ByVal cardW As Long, ByVal cardH As Long, ByVal nodeCenterX As Long, ByVal nodeCenterY As Long, ByVal isLeft As Boolean, ByVal isActive As Boolean, ByVal isHover As Boolean)
    Dim cardBg As OLE_COLOR, cardBorder As OLE_COLOR
    cardBg = IIf(pt.CardColor <> -1, pt.CardColor, m_CardBackColor)
    
    If isActive Then
        cardBorder = m_BorderColorActive
    ElseIf isHover Then
        cardBorder = BlendColors(m_CardBorderColor, m_BorderColorActive, 0.5)
    Else
        cardBorder = IIf(pt.CardBorderColor <> -1, pt.CardBorderColor, m_CardBorderColor)
    End If
    
    If m_CardShadowVisible Then
        DrawSoftShadow G, cardX, cardY, cardW, cardH, m_CardRadius, m_CardShadowBlur, m_CardShadowColor, m_CardShadowAlpha
    End If
    
    Dim hPath As Long
    hPath = CreateRoundRectPath(cardX, cardY, cardW, cardH, m_CardRadius)
    If hPath <> 0 Then
        Dim hBrush As Long, hPen As Long
        Call GdipCreateSolidFill(ARGB(cardBg, 255), hBrush)
        Call GdipFillPath(G, hBrush, hPath)
        Call GdipDeleteBrush(hBrush)
        
        Call GdipCreatePen1(ARGB(cardBorder, 255), m_CardBorderWidth * m_nScale, UnitPixel, hPen)
        Call GdipDrawPath(G, hPen, hPath)
        Call GdipDeletePen(hPen)
        Call GdipDeletePath(hPath)
    End If
    
    If m_CalloutPointer And pt.CalloutArrow Then
        Dim hTriPath As Long
        If GdipCreatePath(0, hTriPath) = 0 Then
            Dim triPts(0 To 2) As POINTL
            If isLeft Then
                triPts(0).X = cardX + cardW: triPts(0).Y = nodeCenterY - 6
                triPts(1).X = cardX + cardW + 8: triPts(1).Y = nodeCenterY
                triPts(2).X = cardX + cardW: triPts(2).Y = nodeCenterY + 6
            Else
                triPts(0).X = cardX: triPts(0).Y = nodeCenterY - 6
                triPts(1).X = cardX - 8: triPts(1).Y = nodeCenterY
                triPts(2).X = cardX: triPts(2).Y = nodeCenterY + 6
            End If
            Call GdipAddPathPolygonI(hTriPath, triPts(0), 3)
            
            Dim hTriBrush As Long
            Call GdipCreateSolidFill(ARGB(cardBg, 255), hTriBrush)
            Call GdipFillPath(G, hTriBrush, hTriPath)
            Call GdipDeleteBrush(hTriBrush)
            
            Call GdipDeletePath(hTriPath)
        End If
    End If
    
    Dim pad As Single: pad = 12 * m_nScale
    Dim rTitle As RECTS, rSub As RECTS, rBadge As RECTS
    
    If Len(pt.BadgeText) > 0 Then
        Dim bColor As Long: bColor = IIf(pt.BadgeColor <> -1, pt.BadgeColor, m_BorderColorActive)
        Dim bW As Long: bW = (Len(pt.BadgeText) * 7 + 14) * m_nScale
        Dim bH As Long: bH = 18 * m_nScale
        Dim bX As Long: bX = cardX + cardW - bW - pad
        Dim bY As Long: bY = cardY + pad
        
        Dim hBPath As Long
        hBPath = CreateRoundRectPath(bX, bY, bW, bH, bH \ 2)
        If hBPath <> 0 Then
            Dim hBBrush As Long
            Call GdipCreateSolidFill(ARGB(bColor, 200), hBBrush)
            Call GdipFillPath(G, hBBrush, hBPath)
            Call GdipDeleteBrush(hBBrush)
            Call GdipDeletePath(hBPath)
        End If
        
        rBadge.Left = bX: rBadge.Top = bY: rBadge.Width = bW: rBadge.Height = bH
        RenderText G, pt.BadgeText, m_FontDate, rBadge, ARGB(&HFFFFFF, 255), StringAlignmentCenter, StringAlignmentCenter
    End If
    
    rTitle.Left = cardX + pad
    rTitle.Top = cardY + pad - 2
    rTitle.Width = cardW - (pad * 2) - IIf(Len(pt.BadgeText) > 0, 70, 0)
    rTitle.Height = 22 * m_nScale
    
    Dim titleCol As Long
    titleCol = IIf(pt.TextColor1 <> -1, pt.TextColor1, m_ForeColor1)
    RenderText G, pt.Title, m_Font1, rTitle, ARGB(titleCol, 255), StringAlignmentNear, StringAlignmentCenter
    
    If Len(pt.Subtitle) > 0 Then
        rSub.Left = cardX + pad
        rSub.Top = cardY + pad + (20 * m_nScale)
        rSub.Width = cardW - (pad * 2)
        rSub.Height = cardH - (pad * 2) - (18 * m_nScale)
        
        Dim subCol As Long
        subCol = IIf(pt.TextColor2 <> -1, pt.TextColor2, m_ForeColor2)
        RenderText G, pt.Subtitle, m_Font2, rSub, ARGB(subCol, 200), StringAlignmentNear, StringAlignmentNear
    End If
End Sub

Private Function HitTestPoint(ByVal X As Single, ByVal Y As Single) As Long
    Dim i As Long
    For i = 1 To m_Points.Count
        Dim pt As clsTimePoint
        Set pt = m_Points.Item(i)
        If pt.Visible Then
            If pt.HitTestNode(X, Y) Then
                HitTestPoint = i
                Exit Function
            End If
            If m_CardVisible Then
                If pt.HitTestCard(X, Y) Then
                    HitTestPoint = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' Deteccion de fila/columna completa (Fallback)
    Dim scrollVal As Long: scrollVal = ucScroll.Value
    If m_Orientation = tlHorizontal Then
        Dim startX As Long: startX = 60 - scrollVal
        For i = 1 To m_Points.Count
            Dim curX As Long: curX = startX + ((i - 1) * m_SectionSpace)
            If X >= curX - (m_SectionSpace \ 2) And X <= curX + (m_SectionSpace \ 2) Then
                HitTestPoint = i
                Exit Function
            End If
        Next i
    Else
        Dim startY As Long: startY = 40 - scrollVal
        For i = 1 To m_Points.Count
            Dim curY As Long: curY = startY + ((i - 1) * m_SectionSpace)
            If Y >= curY - (m_SectionSpace \ 2) And Y <= curY + (m_SectionSpace \ 2) Then
                HitTestPoint = i
                Exit Function
            End If
        Next i
    End If
    HitTestPoint = 0
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim hitIdx As Long
    hitIdx = HitTestPoint(X, Y)
    
    If hitIdx <> m_HoverIndex Then
        m_HoverIndex = hitIdx
        If m_HoverIndex > 0 Then
            Call SetCursor(m_hHandCursor)
            RaiseEvent ItemHover(m_HoverIndex, m_Points.Item(m_HoverIndex))
        End If
        DrawScene
    ElseIf hitIdx > 0 Then
        Call SetCursor(m_hHandCursor)
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim hitIdx As Long
    hitIdx = HitTestPoint(X, Y)
    
    If hitIdx > 0 Then
        m_ActiveSection = hitIdx
        PropertyChanged "ActiveSection"
        DrawScene
        RaiseEvent ItemClick(hitIdx, m_Points.Item(hitIdx))
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If m_ActiveSection > 0 And m_ActiveSection <= m_Points.Count Then
        RaiseEvent ItemDblClick(m_ActiveSection, m_Points.Item(m_ActiveSection))
    End If
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ucScroll_Change()
    DrawScene
End Sub

Private Sub ucScroll_Scroll()
    DrawScene
End Sub

Private Sub UserControl_Paint()
    DrawScene
End Sub

Private Sub UserControl_Resize()
    If m_Orientation = tlHorizontal Then
        ucScroll.Move 0, UserControl.ScaleHeight - 14, UserControl.ScaleWidth, 14
        ucScroll.Orientation = sbHorizontal
    Else
        ucScroll.Move UserControl.ScaleWidth - 14, 0, 14, UserControl.ScaleHeight
        ucScroll.Orientation = sbVertical
    End If
    If ucScroll.Visible Then ucScroll.ZOrder 0
    Refresh
End Sub

Public Sub ApplyTheme(ByVal newTheme As tlTheme)
    m_Theme = newTheme
    Select Case newTheme
        Case tlThemeLight
            m_BackColor = &HF8F9FA
            m_PointBackColor = &HFFFFFF
            m_BorderColor = &HD1D5DB
            m_BorderColorActive = &HE65100
            m_CardBackColor = &HFFFFFF
            m_CardBorderColor = &HE5E7EB
            m_ForeColor1 = &H111827
            m_ForeColor2 = &H4B5563
            m_DateForeColor = &H6B7280
            m_LineColor = &HE5E7EB
            m_LineProgressColor = &HE65100
            m_IconForeColor = &H4B5563
            m_CardShadowColor = &H808080
            m_CardShadowAlpha = 40
            ucScroll.BackColor = &HF8F9FA
            ucScroll.TrackColor = &HEEEDED
            ucScroll.ThumbColor = &HC5C5C5
            ucScroll.ThumbHoverColor = &HA0A0A0
            ucScroll.ThumbDragColor = &H808080
            
        Case tlThemeDark
            m_BackColor = &H121212
            m_PointBackColor = &H1E1E1E
            m_BorderColor = &H383838
            m_BorderColorActive = &HFF9100
            m_CardBackColor = &H1E1E1E
            m_CardBorderColor = &H2C2C2C
            m_ForeColor1 = &HFFFFFF
            m_ForeColor2 = &HAAAAAA
            m_DateForeColor = &H777777
            m_LineColor = &H2C2C2C
            m_LineProgressColor = &HFF9100
            m_IconForeColor = &HFFFFFF
            m_CardShadowColor = &H0
            m_CardShadowAlpha = 110
            ucScroll.BackColor = &H121212
            ucScroll.TrackColor = &H1E1E1E
            ucScroll.ThumbColor = &H484848
            ucScroll.ThumbHoverColor = &H666666
            ucScroll.ThumbDragColor = &HFF9100
            
        Case tlThemeCyberpunk
            m_BackColor = &HB0813
            m_PointBackColor = &H16102B
            m_BorderColor = &HFF007F
            m_BorderColorActive = &HF5FF
            m_CardBackColor = &H150F28
            m_CardBorderColor = &HFF007F
            m_ForeColor1 = &HF5FF
            m_ForeColor2 = &HD1C4E9
            m_DateForeColor = &HFF007F
            m_LineColor = &H391E5A
            m_LineProgressColor = &HF5FF
            m_IconForeColor = &HF5FF
            m_CardShadowColor = &HF5FF
            m_CardShadowAlpha = 50
            ucScroll.BackColor = &HB0813
            ucScroll.TrackColor = &H16102B
            ucScroll.ThumbColor = &HFF007F
            ucScroll.ThumbHoverColor = &HF5FF
            ucScroll.ThumbDragColor = &HF5FF
            
        Case tlThemeCorporate
            m_BackColor = &HF0F4F8
            m_PointBackColor = &HFFFFFF
            m_BorderColor = &H90CAF9
            m_BorderColorActive = &HC62828
            m_CardBackColor = &HFFFFFF
            m_CardBorderColor = &HBBDEFB
            m_ForeColor1 = &HD47A1
            m_ForeColor2 = &H37474F
            m_DateForeColor = &H546E7A
            m_LineColor = &HCFD8DC
            m_LineProgressColor = &HC62828
            m_IconForeColor = &HD47A1
            m_CardShadowColor = &H90A4AE
            m_CardShadowAlpha = 60
            ucScroll.BackColor = &HF0F4F8
            ucScroll.TrackColor = &HE1E9F0
            ucScroll.ThumbColor = &H90CAF9
            ucScroll.ThumbHoverColor = &HC62828
            ucScroll.ThumbDragColor = &HC62828
    End Select
    Refresh
End Sub

Public Function AddTimePoint(ByVal eCaption1 As String, _
                             Optional ByVal eCaption2 As String = vbNullString, _
                             Optional ByVal eIconchar As Variant = vbNullString, _
                             Optional ByVal eTime As String = vbNullString, _
                             Optional ByVal eDate As String = vbNullString, _
                             Optional ByVal eVisible As Boolean = True, _
                             Optional ByVal Picture As Variant = Empty, _
                             Optional ByVal Status As tlItemStatus = tlStatusNone) As clsTimePoint
    
    Dim pt As clsTimePoint
    Set pt = m_Points.Add(eCaption1, eCaption2, eTime, eDate, eIconchar, Picture, Status)
    pt.Visible = eVisible
    Set AddTimePoint = pt
    Refresh
End Function

Public Function UpdateTimePoint(ByVal Index As Long, _
                                ByVal eVisible As Boolean, _
                                Optional ByVal eCaption1 As String = vbNullString, _
                                Optional ByVal eCaption2 As String = vbNullString, _
                                Optional ByVal eIconchar As Variant = vbNullString, _
                                Optional ByVal eTime As String = vbNullString, _
                                Optional ByVal eDate As String = vbNullString, _
                                Optional ByVal Picture As Variant = Empty, _
                                Optional ByVal Status As tlItemStatus = tlStatusNone) As Boolean
    
    If Index >= 1 And Index <= m_Points.Count Then
        Dim pt As clsTimePoint
        Set pt = m_Points.Item(Index)
        With pt
            .Visible = eVisible
            If Len(eCaption1) > 0 Then .Title = eCaption1
            If Len(eCaption2) > 0 Then .Subtitle = eCaption2
            If Len(CStr(eIconchar)) > 0 Then .IconChar = eIconchar
            If Len(eTime) > 0 Then .Timestamp = eTime
            If Len(eDate) > 0 Then .DateText = eDate
            If Not IsEmpty(Picture) Then .Picture = Picture
            If Status <> tlStatusNone Then .Status = Status
        End With
        UpdateTimePoint = True
        Refresh
    End If
End Function

Public Sub EnsureVisible(ByVal Index As Long)
    If Index < 1 Or Index > m_Points.Count Then Exit Sub
    Dim targetPos As Long
    targetPos = (Index - 1) * m_SectionSpace
    If m_Orientation = tlHorizontal Then
        If targetPos < ucScroll.Value Or targetPos > ucScroll.Value + UserControl.ScaleWidth - m_SectionSpace Then
            ucScroll.Value = targetPos
        End If
    Else
        If targetPos < ucScroll.Value Or targetPos > ucScroll.Value + UserControl.ScaleHeight - m_SectionSpace Then
            ucScroll.Value = targetPos
        End If
    End If
    DrawScene
End Sub

Public Property Get points() As clsTimePoints
    Set points = m_Points
End Property

Public Property Get ActiveSection() As Long
    ActiveSection = m_ActiveSection
End Property
Public Property Let ActiveSection(ByVal Value As Long)
    m_ActiveSection = Value
    PropertyChanged "ActiveSection"
    Refresh
End Property

Public Property Get Orientation() As tlOrientation
    Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal Value As tlOrientation)
    m_Orientation = Value
    PropertyChanged "Orientation"
    UserControl_Resize
End Property

Public Property Get style() As Long
    style = m_Orientation
End Property
Public Property Let style(ByVal Value As Long)
    m_Orientation = Value
    PropertyChanged "Style"
    UserControl_Resize
End Property

Public Property Get Theme() As tlTheme
    Theme = m_Theme
End Property
Public Property Let Theme(ByVal Value As tlTheme)
    ApplyTheme Value
    PropertyChanged "Theme"
End Property

Public Property Get NodeShape() As tlNodeShape
    NodeShape = m_NodeShape
End Property
Public Property Let NodeShape(ByVal Value As tlNodeShape)
    m_NodeShape = Value
    PropertyChanged "NodeShape"
    Refresh
End Property

Public Property Get NodeSize() As Long
    NodeSize = m_NodeSize
End Property
Public Property Let NodeSize(ByVal Value As Long)
    m_NodeSize = Value
    PropertyChanged "NodeSize"
    Refresh
End Property

Public Property Get SectionSpace() As Long
    SectionSpace = m_SectionSpace
End Property
Public Property Let SectionSpace(ByVal Value As Long)
    m_SectionSpace = Value
    PropertyChanged "SectionSpace"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get PointBackColor() As OLE_COLOR
    PointBackColor = m_PointBackColor
End Property
Public Property Let PointBackColor(ByVal Value As OLE_COLOR)
    m_PointBackColor = Value
    PropertyChanged "PointBackColor"
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    PropertyChanged "BorderColor"
    Refresh
End Property

Public Property Get BorderColorActive() As OLE_COLOR
    BorderColorActive = m_BorderColorActive
End Property
Public Property Let BorderColorActive(ByVal Value As OLE_COLOR)
    m_BorderColorActive = Value
    PropertyChanged "BorderColorActive"
    Refresh
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = m_BorderWidth
End Property
Public Property Let BorderWidth(ByVal Value As Long)
    m_BorderWidth = Value
    PropertyChanged "BorderWidth"
    Refresh
End Property

Public Property Get CornerCurve() As Long
    CornerCurve = m_CornerCurve
End Property
Public Property Let CornerCurve(ByVal Value As Long)
    m_CornerCurve = Value
    PropertyChanged "CornerCurve"
    Refresh
End Property

Public Property Get LineStyle() As tlLineStyle
    LineStyle = m_LineStyle
End Property
Public Property Let LineStyle(ByVal Value As tlLineStyle)
    m_LineStyle = Value
    PropertyChanged "LineStyle"
    Refresh
End Property

Public Property Get LineWidth() As Long
    LineWidth = m_LineWidth
End Property
Public Property Let LineWidth(ByVal Value As Long)
    m_LineWidth = Value
    PropertyChanged "LineWidth"
    Refresh
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = m_LineColor
End Property
Public Property Let LineColor(ByVal Value As OLE_COLOR)
    m_LineColor = Value
    PropertyChanged "LineColor"
    Refresh
End Property

Public Property Get LineProgressColor() As OLE_COLOR
    LineProgressColor = m_LineProgressColor
End Property
Public Property Let LineProgressColor(ByVal Value As OLE_COLOR)
    m_LineProgressColor = Value
    PropertyChanged "LineProgressColor"
    Refresh
End Property

Public Property Get CardVisible() As Boolean
    CardVisible = m_CardVisible
End Property
Public Property Let CardVisible(ByVal Value As Boolean)
    m_CardVisible = Value
    PropertyChanged "CardVisible"
    Refresh
End Property

Public Property Get CardBackColor() As OLE_COLOR
    CardBackColor = m_CardBackColor
End Property
Public Property Let CardBackColor(ByVal Value As OLE_COLOR)
    m_CardBackColor = Value
    PropertyChanged "CardBackColor"
    Refresh
End Property

Public Property Get CardBorderColor() As OLE_COLOR
    CardBorderColor = m_CardBorderColor
End Property
Public Property Let CardBorderColor(ByVal Value As OLE_COLOR)
    m_CardBorderColor = Value
    PropertyChanged "CardBorderColor"
    Refresh
End Property

Public Property Get CardRadius() As Long
    CardRadius = m_CardRadius
End Property
Public Property Let CardRadius(ByVal Value As Long)
    m_CardRadius = Value
    PropertyChanged "CardRadius"
    Refresh
End Property

Public Property Get CardShadowVisible() As Boolean
    CardShadowVisible = m_CardShadowVisible
End Property
Public Property Let CardShadowVisible(ByVal Value As Boolean)
    m_CardShadowVisible = Value
    PropertyChanged "CardShadowVisible"
    Refresh
End Property

Public Property Get CalloutPointer() As Boolean
    CalloutPointer = m_CalloutPointer
End Property
Public Property Let CalloutPointer(ByVal Value As Boolean)
    m_CalloutPointer = Value
    PropertyChanged "CalloutPointer"
    Refresh
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = m_Font1
End Property
Public Property Set TitleFont(ByVal Value As StdFont)
    Set m_Font1 = Value
    PropertyChanged "TitleFont"
    Refresh
End Property

Public Property Get SubtitleFont() As StdFont
    Set SubtitleFont = m_Font2
End Property
Public Property Set SubtitleFont(ByVal Value As StdFont)
    Set m_Font2 = Value
    PropertyChanged "SubtitleFont"
    Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property
Public Property Set IconFont(ByVal Value As StdFont)
    Set m_IconFont = Value
    PropertyChanged "IconFont"
    Refresh
End Property

Public Property Get DateFont() As StdFont
    Set DateFont = m_FontDate
End Property
Public Property Set DateFont(ByVal Value As StdFont)
    Set m_FontDate = Value
    PropertyChanged "DateFont"
    Refresh
End Property

Public Property Get Caption1Color() As OLE_COLOR
    Caption1Color = m_ForeColor1
End Property
Public Property Let Caption1Color(ByVal Value As OLE_COLOR)
    m_ForeColor1 = Value
    PropertyChanged "Caption1Color"
    Refresh
End Property

Public Property Get Caption2Color() As OLE_COLOR
    Caption2Color = m_ForeColor2
End Property
Public Property Let Caption2Color(ByVal Value As OLE_COLOR)
    m_ForeColor2 = Value
    PropertyChanged "Caption2Color"
    Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property
Public Property Let IconForeColor(ByVal Value As OLE_COLOR)
    m_IconForeColor = Value
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get DateForeColor() As OLE_COLOR
    DateForeColor = m_DateForeColor
End Property
Public Property Let DateForeColor(ByVal Value As OLE_COLOR)
    m_DateForeColor = Value
    PropertyChanged "DateForeColor"
    Refresh
End Property

Public Property Get DateVisible() As Boolean
    DateVisible = m_DateVisible
End Property
Public Property Let DateVisible(ByVal Value As Boolean)
    m_DateVisible = Value
    PropertyChanged "DateVisible"
    Refresh
End Property

Public Property Get TimeVisible() As Boolean
    TimeVisible = m_TimeVisible
End Property
Public Property Let TimeVisible(ByVal Value As Boolean)
    m_TimeVisible = Value
    PropertyChanged "TimeVisible"
    Refresh
End Property


Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get Version() As String
    Version = "2.0.0"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Orientation = .ReadProperty("Orientation", tlVertical)
        m_SectionSpace = .ReadProperty("SectionSpace", 90)
        m_NodeSize = .ReadProperty("NodeSize", 38)
        m_NodeShape = .ReadProperty("NodeShape", tlShapeCircle)
        m_CornerCurve = .ReadProperty("CornerCurve", 10)
        
        m_BackColor = .ReadProperty("BackColor", &HF8F9FA)
        m_PointBackColor = .ReadProperty("PointBackColor", &HFFFFFF)
        m_BorderColor = .ReadProperty("BorderColor", &HD1D5DB)
        m_BorderColorActive = .ReadProperty("BorderColorActive", &HE65100)
        m_BorderWidth = .ReadProperty("BorderWidth", 2)
        
        m_LineStyle = .ReadProperty("LineStyle", tlLineGradient)
        m_LineWidth = .ReadProperty("LineWidth", 3)
        m_LineColor = .ReadProperty("LineColor", &HE5E7EB)
        m_LineProgressColor = .ReadProperty("LineProgressColor", &HE65100)
        
        m_CardVisible = .ReadProperty("CardVisible", True)
        m_CardBackColor = .ReadProperty("CardBackColor", &HFFFFFF)
        m_CardBorderColor = .ReadProperty("CardBorderColor", &HE5E7EB)
        m_CardRadius = .ReadProperty("CardRadius", 8)
        m_CardShadowVisible = .ReadProperty("CardShadowVisible", True)
        m_CalloutPointer = .ReadProperty("CalloutPointer", True)
        
        Set m_Font1 = .ReadProperty("TitleFont", m_Font1)
        Set m_Font2 = .ReadProperty("SubtitleFont", m_Font2)
        Set m_FontDate = .ReadProperty("DateFont", m_FontDate)
        Set m_IconFont = .ReadProperty("IconFont", m_IconFont)
        
        m_ForeColor1 = .ReadProperty("Caption1Color", &H111827)
        m_ForeColor2 = .ReadProperty("Caption2Color", &H4B5563)
        m_IconForeColor = .ReadProperty("IconForeColor", &H4B5563)
        m_DateForeColor = .ReadProperty("DateForeColor", &H6B7280)
        
        m_DateVisible = .ReadProperty("DateVisible", True)
        m_TimeVisible = .ReadProperty("TimeVisible", True)
        m_ActiveSection = .ReadProperty("ActiveSection", 1)
        m_Theme = .ReadProperty("Theme", tlThemeLight)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Orientation", m_Orientation, tlVertical)
        Call .WriteProperty("SectionSpace", m_SectionSpace, 90)
        Call .WriteProperty("NodeSize", m_NodeSize, 38)
        Call .WriteProperty("NodeShape", m_NodeShape, tlShapeCircle)
        Call .WriteProperty("CornerCurve", m_CornerCurve, 10)
        
        Call .WriteProperty("BackColor", m_BackColor, &HF8F9FA)
        Call .WriteProperty("PointBackColor", m_PointBackColor, &HFFFFFF)
        Call .WriteProperty("BorderColor", m_BorderColor, &HD1D5DB)
        Call .WriteProperty("BorderColorActive", m_BorderColorActive, &HE65100)
        Call .WriteProperty("BorderWidth", m_BorderWidth, 2)
        
        Call .WriteProperty("LineStyle", m_LineStyle, tlLineGradient)
        Call .WriteProperty("LineWidth", m_LineWidth, 3)
        Call .WriteProperty("LineColor", m_LineColor, &HE5E7EB)
        Call .WriteProperty("LineProgressColor", m_LineProgressColor, &HE65100)
        
        Call .WriteProperty("CardVisible", m_CardVisible, True)
        Call .WriteProperty("CardBackColor", m_CardBackColor, &HFFFFFF)
        Call .WriteProperty("CardBorderColor", m_CardBorderColor, &HE5E7EB)
        Call .WriteProperty("CardRadius", m_CardRadius, 8)
        Call .WriteProperty("CardShadowVisible", m_CardShadowVisible, True)
        Call .WriteProperty("CalloutPointer", m_CalloutPointer, True)
        
        Call .WriteProperty("TitleFont", m_Font1)
        Call .WriteProperty("SubtitleFont", m_Font2)
        Call .WriteProperty("DateFont", m_FontDate)
        Call .WriteProperty("IconFont", m_IconFont)
        
        Call .WriteProperty("Caption1Color", m_ForeColor1, &H111827)
        Call .WriteProperty("Caption2Color", m_ForeColor2, &H4B5563)
        Call .WriteProperty("IconForeColor", m_IconForeColor, &H4B5563)
        Call .WriteProperty("DateForeColor", m_DateForeColor, &H6B7280)
        
        Call .WriteProperty("DateVisible", m_DateVisible, True)
        Call .WriteProperty("TimeVisible", m_TimeVisible, True)
        Call .WriteProperty("ActiveSection", m_ActiveSection, 1)
        Call .WriteProperty("Theme", m_Theme, tlThemeLight)
    End With
End Sub
