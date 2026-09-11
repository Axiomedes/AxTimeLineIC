VERSION 5.00
Begin VB.UserControl ucScrollbar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
End
Attribute VB_Name = "ucScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' ucScrollbar - Control de Barra de Desplazamiento Suave y Moderna para VB6
' Disenado para maxima estabilidad en el IDE (Sin memory thunks inestables / Cero Crashes)
' Soporta: Orientacion Vertical/Horizontal, Arrastre fluido, Hover, Temas y Rueda de Raton.
'========================================================================================

Public Event Change()
Public Event Scroll()

Public Enum sbOrientation
    sbVertical = 0
    sbHorizontal = 1
End Enum

Private m_Orientation     As sbOrientation
Private m_Min             As Long
Private m_Max             As Long
Private m_Value           As Long
Private m_SmallChange     As Long
Private m_LargeChange     As Long

Private m_ThumbColor      As OLE_COLOR
Private m_ThumbHoverColor As OLE_COLOR
Private m_ThumbDragColor  As OLE_COLOR
Private m_TrackColor      As OLE_COLOR
Private m_BackColor       As OLE_COLOR

Private m_IsHover         As Boolean
Private m_IsDragging      As Boolean
Private m_DragOffset      As Long
Private m_ExtHwnd         As Long

' Geometria del Thumb
Private m_ThumbPos        As Long
Private m_ThumbSize       As Long

Private Sub UserControl_Initialize()
    Call GDIP_Init
    m_Orientation = sbVertical
    m_Min = 0
    m_Max = 100
    m_Value = 0
    m_SmallChange = 20
    m_LargeChange = 60
    
    m_BackColor = &HF8F9FA
    m_TrackColor = &HEEEDED
    m_ThumbColor = &HC5C5C5
    m_ThumbHoverColor = &HA0A0A0
    m_ThumbDragColor = &H808080
    
    m_IsHover = False
    m_IsDragging = False
End Sub

Private Sub UserControl_Terminate()
    Call GDIP_Terminate
End Sub

Public Property Get Orientation() As sbOrientation
    Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal Value As sbOrientation)
    m_Orientation = Value
    Redraw
End Property

Public Property Get Min() As Long
    Min = m_Min
End Property
Public Property Let Min(ByVal Value As Long)
    m_Min = Value
    If m_Value < m_Min Then m_Value = m_Min
    Redraw
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property
Public Property Let Max(ByVal Value As Long)
    m_Max = Value
    If m_Value > m_Max Then m_Value = m_Max
    Redraw
End Property

Public Property Get Value() As Long
    Value = m_Value
End Property
Public Property Let Value(ByVal newVal As Long)
    If newVal < m_Min Then newVal = m_Min
    If newVal > m_Max Then newVal = m_Max
    If m_Value <> newVal Then
        m_Value = newVal
        Redraw
        RaiseEvent Change
        RaiseEvent Scroll
    End If
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_SmallChange
End Property
Public Property Let SmallChange(ByVal Value As Long)
    m_SmallChange = Value
End Property

Public Property Get LargeChange() As Long
    LargeChange = m_LargeChange
End Property
Public Property Let LargeChange(ByVal Value As Long)
    m_LargeChange = Value
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
    Redraw
End Property

Public Property Get TrackColor() As OLE_COLOR
    TrackColor = m_TrackColor
End Property
Public Property Let TrackColor(ByVal Value As OLE_COLOR)
    m_TrackColor = Value
    Redraw
End Property

Public Property Get ThumbColor() As OLE_COLOR
    ThumbColor = m_ThumbColor
End Property
Public Property Let ThumbColor(ByVal Value As OLE_COLOR)
    m_ThumbColor = Value
    Redraw
End Property

Public Property Get ThumbHoverColor() As OLE_COLOR
    ThumbHoverColor = m_ThumbHoverColor
End Property
Public Property Let ThumbHoverColor(ByVal Value As OLE_COLOR)
    m_ThumbHoverColor = Value
    Redraw
End Property

Public Property Get ThumbDragColor() As OLE_COLOR
    ThumbDragColor = m_ThumbDragColor
End Property
Public Property Let ThumbDragColor(ByVal Value As OLE_COLOR)
    m_ThumbDragColor = Value
    Redraw
End Property

Public Function TrackMouseWheelOnHwnd(ByVal lHwnd As Long) As Boolean
    m_ExtHwnd = lHwnd
    TrackMouseWheelOnHwnd = True
End Function

Public Function TrackMouseWheelOnHwndStop() As Boolean
    m_ExtHwnd = 0
    TrackMouseWheelOnHwndStop = True
End Function

Public Sub WheelScrollTopLeft()
    Me.Value = m_Value - m_SmallChange
End Sub

Public Sub WheelScrollBotRight()
    Me.Value = m_Value + m_SmallChange
End Sub

Public Sub Scroll_UP(ByVal Distance_Type As Long)
    If Distance_Type = 0 Then
        Me.Value = m_Value - m_SmallChange
    Else
        Me.Value = m_Value - m_LargeChange
    End If
End Sub

Public Sub Scroll_DOWN(ByVal Distance_Type As Long)
    If Distance_Type = 0 Then
        Me.Value = m_Value + m_SmallChange
    Else
        Me.Value = m_Value + m_LargeChange
    End If
End Sub

Public Sub Scroll_HOME()
    Me.Value = m_Min
End Sub

Public Sub Scroll_END()
    Me.Value = m_Max
End Sub

Private Sub CalculateThumb(ByVal totalTrackLength As Long)
    Dim rangeVal As Long
    rangeVal = m_Max - m_Min
    If rangeVal <= 0 Or totalTrackLength <= 0 Then
        m_ThumbSize = totalTrackLength
        m_ThumbPos = 0
        Exit Sub
    End If
    
    m_ThumbSize = CLng(totalTrackLength * (totalTrackLength / (totalTrackLength + rangeVal)))
    If m_ThumbSize < 24 Then m_ThumbSize = 24
    If m_ThumbSize > totalTrackLength Then m_ThumbSize = totalTrackLength
    
    Dim availTrack As Long
    availTrack = totalTrackLength - m_ThumbSize
    If availTrack <= 0 Then
        m_ThumbPos = 0
    Else
        m_ThumbPos = CLng(((m_Value - m_Min) / rangeVal) * availTrack)
    End If
End Sub

Private Sub Redraw()
    Dim w As Long, h As Long
    w = UserControl.ScaleWidth
    h = UserControl.ScaleHeight
    If w <= 0 Or h <= 0 Then Exit Sub
    
    Dim G As Long
    If GdipCreateFromHDC(UserControl.hdc, G) <> 0 Then Exit Sub
    
    Call GdipSetSmoothingMode(G, SmoothingModeAntiAlias)
    
    ' 1. Fondo de la barra (Track)
    Dim hBgBrush As Long
    Call GdipCreateSolidFill(ARGB(m_TrackColor, 255), hBgBrush)
    Call GdipFillRectangleI(G, hBgBrush, 0, 0, w, h)
    Call GdipDeleteBrush(hBgBrush)
    
    ' 2. Dibujar el Thumb (Pildora Redondeada)
    If m_Orientation = sbVertical Then
        CalculateThumb h
        Dim curThumbCol As OLE_COLOR
        If m_IsDragging Then
            curThumbCol = m_ThumbDragColor
        ElseIf m_IsHover Then
            curThumbCol = m_ThumbHoverColor
        Else
            curThumbCol = m_ThumbColor
        End If
        
        Dim thumbW As Long: thumbW = w - 4
        If thumbW < 4 Then thumbW = 4
        Dim thumbX As Long: thumbX = (w - thumbW) \ 2
        
        Dim hThumbPath As Long
        hThumbPath = CreateRoundRectPath(thumbX, m_ThumbPos + 2, thumbW, m_ThumbSize - 4, thumbW \ 2)
        If hThumbPath <> 0 Then
            Dim hTBrush As Long
            Call GdipCreateSolidFill(ARGB(curThumbCol, 230), hTBrush)
            Call GdipFillPath(G, hTBrush, hThumbPath)
            Call GdipDeleteBrush(hTBrush)
            Call GdipDeletePath(hThumbPath)
        End If
    Else
        CalculateThumb w
        Dim curThumbColH As OLE_COLOR
        If m_IsDragging Then
            curThumbColH = m_ThumbDragColor
        ElseIf m_IsHover Then
            curThumbColH = m_ThumbHoverColor
        Else
            curThumbColH = m_ThumbColor
        End If
        
        Dim thumbH As Long: thumbH = h - 4
        If thumbH < 4 Then thumbH = 4
        Dim thumbY As Long: thumbY = (h - thumbH) \ 2
        
        Dim hThumbPathH As Long
        hThumbPathH = CreateRoundRectPath(m_ThumbPos + 2, thumbY, m_ThumbSize - 4, thumbH, thumbH \ 2)
        If hThumbPathH <> 0 Then
            Dim hTBrushH As Long
            Call GdipCreateSolidFill(ARGB(curThumbColH, 230), hTBrushH)
            Call GdipFillPath(G, hTBrushH, hThumbPathH)
            Call GdipDeleteBrush(hTBrushH)
            Call GdipDeletePath(hThumbPathH)
        End If
    End If
    
    Call GdipDeleteGraphics(G)
    UserControl.Refresh
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    
    If m_Orientation = sbVertical Then
        If Y >= m_ThumbPos And Y <= m_ThumbPos + m_ThumbSize Then
            m_IsDragging = True
            m_DragOffset = Y - m_ThumbPos
        ElseIf Y < m_ThumbPos Then
            Me.Value = m_Value - m_LargeChange
        Else
            Me.Value = m_Value + m_LargeChange
        End If
    Else
        If X >= m_ThumbPos And X <= m_ThumbPos + m_ThumbSize Then
            m_IsDragging = True
            m_DragOffset = X - m_ThumbPos
        ElseIf X < m_ThumbPos Then
            Me.Value = m_Value - m_LargeChange
        Else
            Me.Value = m_Value + m_LargeChange
        End If
    End If
    Redraw
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_IsHover Then
        m_IsHover = True
        Redraw
    End If
    
    If m_IsDragging Then
        Dim rangeVal As Long: rangeVal = m_Max - m_Min
        If rangeVal > 0 Then
            Dim totalLen As Long, availTrack As Long, newPos As Long
            If m_Orientation = sbVertical Then
                totalLen = UserControl.ScaleHeight
                availTrack = totalLen - m_ThumbSize
                newPos = Y - m_DragOffset
            Else
                totalLen = UserControl.ScaleWidth
                availTrack = totalLen - m_ThumbSize
                newPos = X - m_DragOffset
            End If
            
            If availTrack > 0 Then
                If newPos < 0 Then newPos = 0
                If newPos > availTrack Then newPos = availTrack
                Dim newVal As Long
                newVal = m_Min + CLng((newPos / availTrack) * rangeVal)
                Me.Value = newVal
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_IsDragging Then
        m_IsDragging = False
        Redraw
    End If
End Sub

Private Sub UserControl_Resize()
    Redraw
End Sub

Private Sub UserControl_Paint()
    Redraw
End Sub

Private Sub UserControl_Show()
    Redraw
End Sub

