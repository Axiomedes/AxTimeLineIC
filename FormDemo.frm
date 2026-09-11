VERSION 5.00
Object = "*\AAxTimeLineIC.vbp"
Begin VB.Form FormDemo 
   BackColor       =   &H00F0F0F0&
   Caption         =   "AxTimeLineIC - Showcase & Demostracion de Funcionalidades"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "FormDemo"
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   880
   StartUpPosition =   2  'CenterScreen
   Begin vb6projectAxTimeLineIC.AxTimeLineIC AxTimeLine1 
      Height          =   7410
      Left            =   255
      TabIndex        =   17
      Top             =   240
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   13070
      NodeSize        =   50
      NodeShape       =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubtitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DateFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraControls 
      Caption         =   " Panel de Configuracion y Estilos "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtSectionSpace 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   390
         TabIndex        =   24
         Text            =   "90"
         Top             =   5010
         Width           =   390
      End
      Begin VB.TextBox txtNodeSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   390
         TabIndex        =   22
         Text            =   "38"
         Top             =   4680
         Width           =   390
      End
      Begin VB.TextBox txtCorner 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   390
         TabIndex        =   20
         Text            =   "10"
         Top             =   4350
         Width           =   390
      End
      Begin VB.TextBox txtRadius 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   390
         TabIndex        =   18
         Text            =   "8"
         Top             =   4020
         Width           =   390
      End
      Begin VB.CommandButton cmdAddTriangle 
         Caption         =   "+ Hito con Triangulo"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdAddCustom 
         Caption         =   "+ Agregar Hito"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdEnsureVis 
         Caption         =   "Ir a Seccion Activa"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   6240
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdvance 
         Caption         =   "Avanzar Paso >>"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   6240
         Width           =   1815
      End
      Begin VB.CheckBox chkCallout 
         Caption         =   "Flecha Puntero (Callout Arrow)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3630
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkShadows 
         Caption         =   "Sombras Suaves (Drop Shadows)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3345
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkCards 
         Caption         =   "Tarjetas de Contenido (Cards)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3075
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.ComboBox cboLineStyle 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2655
         Width           =   3495
      End
      Begin VB.ComboBox cboShape 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2010
         Width           =   3495
      End
      Begin VB.ComboBox cboOrientation 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   3495
      End
      Begin VB.ComboBox cboTheme 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   660
         Width           =   3495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Space"
         Height          =   195
         Left            =   870
         TabIndex        =   25
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Node Size"
         Height          =   195
         Left            =   870
         TabIndex        =   23
         Top             =   4710
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corner Curve"
         Height          =   195
         Left            =   870
         TabIndex        =   21
         Top             =   4380
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Radius"
         Height          =   195
         Left            =   870
         TabIndex        =   19
         Top             =   4050
         Width           =   870
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Haz clic en cualquier nodo para interactuar"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   6840
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Estilo de Linea Conectora:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Forma de Nodos (Shapes):"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1740
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Modo de Presentacion:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1065
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Temas Visuales:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   3495
      End
   End
End
Attribute VB_Name = "FormDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With cboTheme
        .AddItem "1. Modern Light (Por Defecto)"
        .AddItem "2. Dark Mode Studio"
        .AddItem "3. Cyberpunk Neon"
        .AddItem "4. Corporate Blue"
        .ListIndex = 0
    End With
    
    With cboOrientation
        .AddItem "Vertical Clasico"
        .AddItem "Horizontal (Wizard / Stepper)"
        .AddItem "Vertical Alternado (Zig-Zag / Split)"
        .ListIndex = 0
    End With
    
    With cboShape
        .AddItem "Circulo (Circle)"
        .AddItem "Rectangulo Redondeado (Rounded)"
        .AddItem "Cuadrado (Square)"
        .AddItem "Pildora (Pill)"
        .AddItem "Diamante / Rombo (Diamond)"
        .AddItem "Triangulo Hacia Arriba (Triangle Up)"
        .AddItem "Triangulo Hacia Abajo (Triangle Down)"
        .ListIndex = 0
    End With
    
    With cboLineStyle
        .AddItem "Solida con Degradado de Progreso"
        .AddItem "Discontinua (Dashed)"
        .AddItem "Punteada (Dotted)"
        .ListIndex = 0
    End With
    
    LoadDemoData
End Sub

Private Sub LoadDemoData()
    AxTimeLine1.Points.Clear
    
    Dim pt As clsTimePoint
    
    ' Hito 1: Completado
    Set pt = AxTimeLine1.AddTimePoint("Requerimientos y Alcance", _
                             "Especificaciones aprobadas por el comite directivo.", _
                             "E73E", "09:00", "01/10/2026", True, , tlStatusCompleted)
    pt.BadgeText = "Completado"
    pt.BadgeColor = &H43A047
    
    ' Hito 2: Completado
    Set pt = AxTimeLine1.AddTimePoint("Diseno UI/UX y Prototipos", _
                             "Mockups y flujos de usuario validados con clientes.", _
                             "E790", "11:30", "03/10/2026", True, , tlStatusCompleted)
    pt.BadgeText = "Completado"
    pt.BadgeColor = &H43A047
    
    ' Hito 3: Completado
    Set pt = AxTimeLine1.AddTimePoint("Arquitectura de Base de Datos", _
                             "Modelado de esquemas, indices y migraciones.", _
                             "E7B8", "14:15", "05/10/2026", True, , tlStatusCompleted)
    pt.BadgeText = "Completado"
    pt.BadgeColor = &H43A047
    
    ' Hito 4: En Progreso (Activo)
    Set pt = AxTimeLine1.AddTimePoint("Desarrollo de AxTimeLineIC v2", _
                             "Implementacion del motor GDI+, doble bufer, sombras y figuras.", _
                             "E7BE", "16:45", "08/10/2026", True, , tlStatusInProgress)
    pt.BadgeText = "En Progreso"
    pt.BadgeColor = &HE65100
    
    ' Hito 5: Con Forma Especial Triangulo
    Set pt = AxTimeLine1.AddTimePoint("Validacion de Rendimiento (QA)", _
                             "Pruebas de stress y renderizado sin fugas de memoria.", _
                             "E814", "10:00", "10/10/2026", True, , tlStatusPending)
    pt.NodeShape = tlShapeTriangleUp
    pt.BadgeText = "Importante"
    pt.BadgeColor = &H9C27B0
    
    ' Hito 6: Pruebas de Integracion
    Set pt = AxTimeLine1.AddTimePoint("Pruebas de Integracion y API", _
                             "Verificacion de endpoints y estabilidad de componentes.", _
                             "E770", "13:30", "12/10/2026", True, , tlStatusPending)
    pt.BadgeText = "Pendiente"
    
    ' Hito 7: Despliegue
    Set pt = AxTimeLine1.AddTimePoint("Despliegue a Produccion", _
                             "Lanzamiento de la nueva version del componente ActiveX.", _
                             "E753", "15:00", "15/10/2026", True, , tlStatusPending)
    pt.BadgeText = "Pendiente"
    
    ' Hito 8: Auditoria y Cierre
    Set pt = AxTimeLine1.AddTimePoint("Auditoria y Cierre de Proyecto", _
                             "Revision final y entrega de documentacion tecnica.", _
                             "E74E", "18:00", "18/10/2026", True, , tlStatusPending)
    pt.BadgeText = "Pendiente"
    
    AxTimeLine1.ActiveSection = 4
    AxTimeLine1.Refresh
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.ScaleWidth > 320 And Me.ScaleHeight > 200 Then
        fraControls.Left = Me.ScaleWidth - fraControls.Width - 10
        fraControls.Height = Me.ScaleHeight - 20
        'lblStatus.Top = fraControls.Height - lblStatus.Height - 15
        'cmdAdvance.Top = lblStatus.Top - cmdAdvance.Height - 10
        'cmdEnsureVis.Top = cmdAdvance.Top
        'cmdAddCustom.Top = cmdAdvance.Top - cmdAddCustom.Height - 10
        'cmdAddTriangle.Top = cmdAddCustom.Top
        
        AxTimeLine1.Move 10, 10, fraControls.Left - 20, Me.ScaleHeight - 20
    End If
End Sub

Private Sub cboTheme_Click()
    AxTimeLine1.Theme = cboTheme.ListIndex + 1
End Sub

Private Sub cboOrientation_Click()
    AxTimeLine1.Orientation = cboOrientation.ListIndex
End Sub

Private Sub cboShape_Click()
    AxTimeLine1.NodeShape = cboShape.ListIndex
End Sub

Private Sub cboLineStyle_Click()
    Select Case cboLineStyle.ListIndex
        Case 0: AxTimeLine1.LineStyle = tlLineGradient
        Case 1: AxTimeLine1.LineStyle = tlLineDashed
        Case 2: AxTimeLine1.LineStyle = tlLineDotted
    End Select
End Sub

Private Sub chkCards_Click()
    AxTimeLine1.CardVisible = (chkCards.Value = 1)
End Sub

Private Sub chkShadows_Click()
    AxTimeLine1.CardShadowVisible = (chkShadows.Value = 1)
End Sub

Private Sub chkCallout_Click()
    AxTimeLine1.CalloutPointer = (chkCallout.Value = 1)
End Sub

Private Sub cmdAdvance_Click()
    If AxTimeLine1.ActiveSection < AxTimeLine1.Points.Count Then
        Dim curPt As clsTimePoint
        Set curPt = AxTimeLine1.Points.Item(AxTimeLine1.ActiveSection)
        curPt.Status = tlStatusCompleted
        curPt.BadgeText = "Completado"
        curPt.BadgeColor = &H43A047
        
        AxTimeLine1.ActiveSection = AxTimeLine1.ActiveSection + 1
        Dim nextPt As clsTimePoint
        Set nextPt = AxTimeLine1.Points.Item(AxTimeLine1.ActiveSection)
        nextPt.Status = tlStatusInProgress
        nextPt.BadgeText = "En Progreso"
        nextPt.BadgeColor = &HE65100
        
        AxTimeLine1.EnsureVisible AxTimeLine1.ActiveSection
        AxTimeLine1.Refresh
        
        lblStatus.Caption = " Paso avanzado a: " & nextPt.Title
    End If
End Sub

Private Sub cmdAddCustom_Click()
    Dim idx As Long: idx = AxTimeLine1.Points.Count + 1
    Dim pt As clsTimePoint
    Set pt = AxTimeLine1.AddTimePoint("Hito Personalizado #" & idx, _
                             "Descripcion del nuevo paso creado dinamicamente en tiempo de ejecucion.", _
                             "E70F", Format$(Now, "hh:mm"), Date$, True, , tlStatusPending)
    pt.BadgeText = "Nuevo"
    pt.BadgeColor = &H2196F3
    AxTimeLine1.EnsureVisible idx
    lblStatus.Caption = " Nuevo hito #" & idx & " agregado."
End Sub

Private Sub cmdAddTriangle_Click()
    Dim idx As Long: idx = AxTimeLine1.Points.Count + 1
    Dim pt As clsTimePoint
    Set pt = AxTimeLine1.AddTimePoint("Hito con Triangulo Direccional", _
                             "Nodo configurado con forma de triangulo apuntando hacia arriba/abajo.", _
                             "E713", Format$(Now, "hh:mm"), Date$, True, , tlStatusWarning)
    pt.NodeShape = tlShapeTriangleUp
    pt.BadgeText = "Triangulo"
    pt.BadgeColor = &HE91E63
    AxTimeLine1.EnsureVisible idx
    lblStatus.Caption = " Hito con Triangulo #" & idx & " agregado."
End Sub

Private Sub cmdEnsureVis_Click()
    AxTimeLine1.EnsureVisible AxTimeLine1.ActiveSection
End Sub

Private Sub AxTimeLine1_ItemClick(ByVal Index As Long, ByVal Point As clsTimePoint)
    lblStatus.Caption = " ItemClick:" & vbCrLf & _
                        "  Indice: " & Index & vbCrLf & _
                        "  Titulo: " & Point.Title & vbCrLf & _
                        "  Estado: " & Point.BadgeText & vbCrLf & _
                        "  Fecha/Hora: " & Point.DateText & " " & Point.TimeStamp
End Sub

Private Sub AxTimeLine1_ItemDblClick(ByVal Index As Long, ByVal Point As clsTimePoint)
    MsgBox "Doble clic en: " & Point.Title & vbCrLf & Point.Subtitle, vbInformation, "Detalle del Hito"
End Sub

Private Sub AxTimeLine1_ItemHover(ByVal Index As Long, ByVal Point As clsTimePoint)
    ' Feedback en hover
End Sub

Private Sub txtCorner_Change()
AxTimeLine1.CornerCurve = Val(txtCorner.Text)
End Sub

Private Sub txtNodeSize_Change()
AxTimeLine1.NodeSize = Val(txtNodeSize.Text)
End Sub

Private Sub txtRadius_Change()
AxTimeLine1.CardRadius = Val(txtRadius.Text)
End Sub

Private Sub txtSectionSpace_Change()
AxTimeLine1.SectionSpace = Val(txtSectionSpace.Text)
End Sub
