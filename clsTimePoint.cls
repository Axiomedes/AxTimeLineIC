VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================================
' AxTimeLineIC v2 - Modelo de Datos para Hito Individual (clsTimePoint)
' Disenado estrictamente con tipos primitivos compatibles con COM en VB6
'========================================================================================

Public Enum tlNodeShape
    tlShapeDefault = -1
    tlShapeCircle = 0
    tlShapeRoundedRect = 1
    tlShapeSquare = 2
    tlShapePill = 3
    tlShapeDiamond = 4
    tlShapeTriangleUp = 5
    tlShapeTriangleDown = 6
End Enum

Public Enum tlItemStatus
    tlStatusNone = 0
    tlStatusCompleted = 1
    tlStatusInProgress = 2
    tlStatusPending = 3
    tlStatusWarning = 4
    tlStatusFailed = 5
End Enum

Private m_Title            As String
Private m_Subtitle         As String
Private m_Timestamp        As String
Private m_DateText         As String
Private m_IconChar         As Variant
Private m_IconFont         As StdFont
Private m_Picture          As Variant
Private m_hGdiImage        As Long
Private m_ImageCircular    As Boolean
Private m_Status           As tlItemStatus
Private m_NodeShape        As tlNodeShape
Private m_NodeColor        As OLE_COLOR
Private m_NodeBorderColor  As OLE_COLOR
Private m_CardColor        As OLE_COLOR
Private m_CardBorderColor  As OLE_COLOR
Private m_TextColor1       As OLE_COLOR
Private m_TextColor2       As OLE_COLOR
Private m_BadgeText        As String
Private m_BadgeColor       As OLE_COLOR
Private m_Tag              As Variant
Private m_Key              As String
Private m_Visible          As Boolean
Private m_CalloutArrow     As Boolean
Private m_UserData         As Variant

' Coordenadas calculadas en tiempo de ejecucion (Tipos Long primitivos 100% compatibles con VB6)
Private m_NodeX            As Long
Private m_NodeY            As Long
Private m_NodeWidth        As Long
Private m_NodeHeight       As Long

Private m_CardX            As Long
Private m_CardY            As Long
Private m_CardWidth        As Long
Private m_CardHeight       As Long

Private Sub Class_Initialize()
    m_Visible = True
    m_NodeShape = tlShapeDefault
    m_Status = tlStatusNone
    m_ImageCircular = True
    m_NodeColor = -1
    m_NodeBorderColor = -1
    m_CardColor = -1
    m_CardBorderColor = -1
    m_TextColor1 = -1
    m_TextColor2 = -1
    m_BadgeColor = -1
    m_CalloutArrow = True
End Sub

Private Sub Class_Terminate()
    ClearImageHandle
End Sub

Public Sub ClearImageHandle()
    If m_hGdiImage <> 0 Then
        Call GdipDisposeImage(m_hGdiImage)
        m_hGdiImage = 0
    End If
End Sub

' Metodos para establecer dimensiones sin usar UDT
Public Sub SetNodeRect(ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal h As Long)
    m_NodeX = X
    m_NodeY = Y
    m_NodeWidth = w
    m_NodeHeight = h
End Sub

Public Sub SetCardRect(ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal h As Long)
    m_CardX = X
    m_CardY = Y
    m_CardWidth = w
    m_CardHeight = h
End Sub

Public Function HitTestNode(ByVal X As Single, ByVal Y As Single) As Boolean
    HitTestNode = (X >= m_NodeX And X <= m_NodeX + m_NodeWidth And Y >= m_NodeY And Y <= m_NodeY + m_NodeHeight)
End Function

Public Function HitTestCard(ByVal X As Single, ByVal Y As Single) As Boolean
    HitTestCard = (X >= m_CardX And X <= m_CardX + m_CardWidth And Y >= m_CardY And Y <= m_CardY + m_CardHeight)
End Function

' Region "Propiedades Publicas de Coordenadas"

Public Property Get NodeLeft() As Long
    NodeLeft = m_NodeX
End Property
Public Property Get NodeTop() As Long
    NodeTop = m_NodeY
End Property
Public Property Get NodeWidth() As Long
    NodeWidth = m_NodeWidth
End Property
Public Property Get NodeHeight() As Long
    NodeHeight = m_NodeHeight
End Property

Public Property Get CardLeft() As Long
    CardLeft = m_CardX
End Property
Public Property Get CardTop() As Long
    CardTop = m_CardY
End Property
Public Property Get CardWidth() As Long
    CardWidth = m_CardWidth
End Property
Public Property Get CardHeight() As Long
    CardHeight = m_CardHeight
End Property

' End Region

' Region "Propiedades Principales"

Public Property Get Title() As String
    Title = m_Title
End Property
Public Property Let Title(ByVal Value As String)
    m_Title = Value
End Property

Public Property Get Subtitle() As String
    Subtitle = m_Subtitle
End Property
Public Property Let Subtitle(ByVal Value As String)
    m_Subtitle = Value
End Property

Public Property Get Timestamp() As String
    Timestamp = m_Timestamp
End Property
Public Property Let Timestamp(ByVal Value As String)
    m_Timestamp = Value
End Property

Public Property Get DateText() As String
    DateText = m_DateText
End Property
Public Property Let DateText(ByVal Value As String)
    m_DateText = Value
End Property

Public Property Get IconChar() As Variant
    IconChar = m_IconChar
End Property
Public Property Let IconChar(ByVal Value As Variant)
    m_IconChar = Value
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property
Public Property Set IconFont(ByVal Value As StdFont)
    Set m_IconFont = Value
End Property

Public Property Get Picture() As Variant
    If IsObject(m_Picture) Then
        Set Picture = m_Picture
    Else
        Picture = m_Picture
    End If
End Property
Public Property Let Picture(ByVal Value As Variant)
    m_Picture = Value
    ClearImageHandle
    m_hGdiImage = LoadGDIPlusImage(m_Picture)
End Property
Public Property Set Picture(ByVal Value As Object)
    Set m_Picture = Value
    ClearImageHandle
    m_hGdiImage = LoadGDIPlusImage(m_Picture)
End Property

Public Property Get hGdiImage() As Long
    hGdiImage = m_hGdiImage
End Property

Public Property Get ImageCircular() As Boolean
    ImageCircular = m_ImageCircular
End Property
Public Property Let ImageCircular(ByVal Value As Boolean)
    m_ImageCircular = Value
End Property

Public Property Get Status() As tlItemStatus
    Status = m_Status
End Property
Public Property Let Status(ByVal Value As tlItemStatus)
    m_Status = Value
End Property

Public Property Get NodeShape() As tlNodeShape
    NodeShape = m_NodeShape
End Property
Public Property Let NodeShape(ByVal Value As tlNodeShape)
    m_NodeShape = Value
End Property

Public Property Get NodeColor() As OLE_COLOR
    NodeColor = m_NodeColor
End Property
Public Property Let NodeColor(ByVal Value As OLE_COLOR)
    m_NodeColor = Value
End Property

Public Property Get NodeBorderColor() As OLE_COLOR
    NodeBorderColor = m_NodeBorderColor
End Property
Public Property Let NodeBorderColor(ByVal Value As OLE_COLOR)
    m_NodeBorderColor = Value
End Property

Public Property Get CardColor() As OLE_COLOR
    CardColor = m_CardColor
End Property
Public Property Let CardColor(ByVal Value As OLE_COLOR)
    m_CardColor = Value
End Property

Public Property Get CardBorderColor() As OLE_COLOR
    CardBorderColor = m_CardBorderColor
End Property
Public Property Let CardBorderColor(ByVal Value As OLE_COLOR)
    m_CardBorderColor = Value
End Property

Public Property Get TextColor1() As OLE_COLOR
    TextColor1 = m_TextColor1
End Property
Public Property Let TextColor1(ByVal Value As OLE_COLOR)
    m_TextColor1 = Value
End Property

Public Property Get TextColor2() As OLE_COLOR
    TextColor2 = m_TextColor2
End Property
Public Property Let TextColor2(ByVal Value As OLE_COLOR)
    m_TextColor2 = Value
End Property

Public Property Get BadgeText() As String
    BadgeText = m_BadgeText
End Property
Public Property Let BadgeText(ByVal Value As String)
    m_BadgeText = Value
End Property

Public Property Get BadgeColor() As OLE_COLOR
    BadgeColor = m_BadgeColor
End Property
Public Property Let BadgeColor(ByVal Value As OLE_COLOR)
    m_BadgeColor = Value
End Property

Public Property Get CalloutArrow() As Boolean
    CalloutArrow = m_CalloutArrow
End Property
Public Property Let CalloutArrow(ByVal Value As Boolean)
    m_CalloutArrow = Value
End Property

Public Property Get Tag() As Variant
    Tag = m_Tag
End Property
Public Property Let Tag(ByVal Value As Variant)
    m_Tag = Value
End Property

Public Property Get Key() As String
    Key = m_Key
End Property
Public Property Let Key(ByVal Value As String)
    m_Key = Value
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property
Public Property Let Visible(ByVal Value As Boolean)
    m_Visible = Value
End Property

Public Property Get UserData() As Variant
    If IsObject(m_UserData) Then
        Set UserData = m_UserData
    Else
        UserData = m_UserData
    End If
End Property
Public Property Let UserData(ByVal Value As Variant)
    m_UserData = Value
End Property
Public Property Set UserData(ByVal Value As Object)
    Set m_UserData = Value
End Property

' End Region
