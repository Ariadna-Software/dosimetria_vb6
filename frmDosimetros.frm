VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDosimetros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dos�metros"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7665
   Icon            =   "frmDosimetros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   7665
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   450
      Left            =   2655
      TabIndex        =   57
      Top             =   450
      Visible         =   0   'False
      Width           =   4950
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   13
         Left            =   1545
         MaxLength       =   5
         TabIndex        =   2
         Tag             =   "Tipo de Medicion|T|S|||dosimetros|tipo_medicion|||"
         Text            =   "Text1"
         Top             =   90
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   1980
         MaxLength       =   30
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   90
         Width           =   2895
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   6
         Left            =   1260
         MouseIcon       =   "frmDosimetros.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimetros.frx":0E1C
         ToolTipText     =   "Buscar tipo de medici�n"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Medici�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   105
         TabIndex        =   58
         Top             =   90
         Width           =   840
      End
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmDosimetros.frx":0F1E
      Left            =   1170
      List            =   "frmDosimetros.frx":0F20
      TabIndex        =   3
      Text            =   "Combo4"
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   52
      Top             =   1395
      Width           =   7425
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   2520
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Cristal c|N|N|0|999.999|dosimetros|cristal_c|||"
         Text            =   "Text1"
         Top             =   1080
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   4320
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Cristal d|N|N|0|999.999|dosimetros|cristal_d|||"
         Text            =   "Text1"
         Top             =   1080
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   4320
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Cristal b|N|N|0|999.999|dosimetros|cristal_b|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   2520
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Cristal a|N|N|0|999.999|dosimetros|cristal_a|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label8 
         Caption         =   "Valores C y D s�lo v�lidos en PANASONIC (para resto sin efecto)"
         Height          =   255
         Left            =   1800
         TabIndex        =   63
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   1800
         X2              =   7200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Cristal C"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   12
         Left            =   1800
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Cristal D"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   11
         Left            =   3600
         TabIndex        =   61
         Top             =   1080
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Cristal B"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   8
         Left            =   3600
         TabIndex        =   55
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Cristal A"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   7
         Left            =   1800
         TabIndex        =   54
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "FACTORES CORRECION"
         ForeColor       =   &H00000000&
         Height          =   660
         Index           =   6
         Left            =   180
         TabIndex        =   53
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmDosimetros.frx":0F22
      Left            =   1170
      List            =   "frmDosimetros.frx":0F24
      TabIndex        =   1
      Tag             =   "Tipo Dosimetro|N|N|||dosimetros|tipo_dosimetro||S|"
      Text            =   "Combo3"
      Top             =   525
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   8
      Left            =   3915
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "N.Dosimetro|T|N|||dosimetros|n_dosimetro||S|"
      Text            =   "Text1"
      Top             =   1050
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   7
      Left            =   6360
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "N.Registro|N|N|||dosimetros|n_reg_dosimetro||S|"
      Text            =   "Text1"
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Frame Frame7 
      Height          =   1065
      Left            =   120
      TabIndex        =   33
      Top             =   2865
      Width           =   7425
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   240
         Width           =   4050
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Codigo Empresa|T|N|||dosimetros|c_empresa|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3150
         MaxLength       =   40
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   600
         Width           =   4050
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "Codigo Instalaci�n|T|N|||dosimetros|c_instalacion|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   1305
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   2
         Left            =   1470
         MouseIcon       =   "frmDosimetros.frx":0F26
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimetros.frx":1078
         ToolTipText     =   "Buscar empresa"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   180
         TabIndex        =   48
         Top             =   240
         Width           =   915
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   4
         Left            =   1470
         MouseIcon       =   "frmDosimetros.frx":117A
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimetros.frx":12CC
         ToolTipText     =   "Buscar instalaci�n"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Instalaci�n"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   180
         TabIndex        =   44
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1665
      Left            =   120
      TabIndex        =   25
      Top             =   3945
      Width           =   7425
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   12
         Tag             =   "DNI|T|N|||dosimetros|dni_usuario|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1260
         Width           =   5340
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   930
         Width           =   5325
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   570
         Width           =   5325
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   5
         Left            =   1470
         MouseIcon       =   "frmDosimetros.frx":13CE
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimetros.frx":1520
         ToolTipText     =   "Buscar D.N.I."
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "D.N.I."
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   2910
         TabIndex        =   32
         Top             =   945
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   180
         TabIndex        =   31
         Top             =   1290
         Width           =   1290
      End
      Begin VB.Label Label9 
         Caption         =   "Primer Apellido:"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label17 
         Caption         =   "Segundo Apellido:"
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   960
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   120
      TabIndex        =   34
      Top             =   5625
      Width           =   7425
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5700
         TabIndex        =   15
         Text            =   "Combo2"
         Top             =   600
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   5700
         MaxLength       =   10
         TabIndex        =   46
         Tag             =   "Mes Par/Impar|T|N|||dosimetros|mes_p_i|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   5700
         MaxLength       =   40
         TabIndex        =   17
         Tag             =   "Fecha Retirada|F|S|||dosimetros|f_retirada|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   990
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   16
         Tag             =   "Fecha Asignacion|F|N|||dosimetros|f_asig_dosimetro|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   990
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1830
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Plantilla/Contrata|T|N|||dosimetros|plantilla_contrata|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3240
         MaxLength       =   40
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   240
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   18
         Tag             =   "Observaciones|T|S|||dosimetros|observaciones|||"
         Text            =   "Text1"
         Top             =   1350
         Width           =   5280
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Tipo de Trabajo|T|N|||dosimetros|c_tipo_trabajo|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1320
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmDosimetros.frx":1622
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimetros.frx":1774
         ToolTipText     =   "Seleccionar fecha"
         Top             =   990
         Width           =   240
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   1
         Left            =   5430
         MouseIcon       =   "frmDosimetros.frx":17FF
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimetros.frx":1951
         ToolTipText     =   "Seleccionar fecha"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Mes Par/Impar"
         Height          =   255
         Left            =   4200
         TabIndex        =   45
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Retirada"
         Height          =   255
         Left            =   4230
         TabIndex        =   42
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Asignaci�n"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   1020
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Plantilla/Contrata:"
         Height          =   255
         Left            =   180
         TabIndex        =   40
         Top             =   630
         Width           =   1905
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   3
         Left            =   1500
         MouseIcon       =   "frmDosimetros.frx":19DC
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimetros.frx":1B2E
         ToolTipText     =   "Buscar tipo de trabajo"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   180
         TabIndex        =   37
         Top             =   1380
         Width           =   1560
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Trabajo:"
         Height          =   255
         Left            =   180
         TabIndex        =   36
         Top             =   270
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6450
      TabIndex        =   21
      Top             =   7575
      Width           =   1110
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6465
      TabIndex        =   20
      Top             =   7575
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   22
      Top             =   7455
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   210
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5190
      TabIndex        =   19
      Top             =   7575
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   330
      Top             =   7575
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5220
         TabIndex        =   0
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   60
      Tag             =   "Sistema|T|N|||dosimetros|sistema||S|"
      Text            =   "Text3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   150
      X2              =   7545
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Sistema"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   9
      Left            =   210
      TabIndex        =   56
      Top             =   1050
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   5
      Left            =   210
      TabIndex        =   51
      Top             =   525
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Dos�metro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   4
      Left            =   2760
      TabIndex        =   50
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   2
      Left            =   5295
      TabIndex        =   49
      Top             =   1050
      Width           =   975
   End
End
Attribute VB_Name = "frmDosimetros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmEmp As frmEmpresas
Attribute frmEmp.VB_VarHelpID = -1
Private WithEvents frmOpe As frmOperarios
Attribute frmOpe.VB_VarHelpID = -1
Private WithEvents frmIns As frmInstalaciones
Attribute frmIns.VB_VarHelpID = -1
Private WithEvents frmTTr As frmTiposTrab
Attribute frmTTr.VB_VarHelpID = -1
Private WithEvents frmTEx As frmTiposExtremidades
Attribute frmTEx.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private ValorAnterior As String
Private IndiceErroneo As Byte
Private conexion As Integer
Private familia As Integer
Dim SumaLinea As Currency
Dim I As Integer
Dim Numlinea As Integer
Dim Aux As Currency
Dim PulsadoSalir As Boolean
Private ModificandoLineas As Byte
Dim AntiguoText1 As String
Dim ape1 As String
Dim ape2 As String
Dim nombre As String

' campo que indica si la familia es fitosanitaria
' si lo es: obligamos a introducir los campos de fitos.
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

Private Sub frmTEx_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(11).Text = RecuperaValor(CadenaSeleccion, 2)
        Text1_LostFocus 13
    End If
End Sub

Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
    AsignarTeclasFuncion KeyCode
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    Dim v_aux As Integer
    Dim sql As String
    
    
    Screen.MousePointer = vbHourglass
    If Combo4.ListIndex <> -1 Then
      Text3.Text = Left(Combo4.Text, 1)
      If Combo4.ListIndex <> 1 Then
        '-- VRS:1.3.5 Si no es Panasonic ponemos los fatores de correcci�n adicionales
        '   a 1
        Text1(14) = "1"
        Text1(15) = "1"
      End If
    Else
      Text3.Text = ""
    End If
      
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                 
                If SituarData1 Then
                    PonerModo 2
                    PonerCampos
                Else
                    PonerModo 0
                End If
                 
            End If
        End If
    
    Case 4
        'Modificar
        If DatosOk Then
            '-----------------------------------------
            'Hacemos modificar
          
            If ModificaDesdeFormulario(Me, 1) Then
                DesbloqueaRegistroForm1 Me
                If SituarData1 Then
                    PonerModo 2
                    PonerCampos
                Else

                    PonerModo 0
                End If
            End If
        End If

    Case 1
        HacerBusqueda
    End Select

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation, "�Error!"
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        
        Case 4
            'Modificar
            DesBloqueaRegistroForm Text1(0)
            PonerModo 2
            PonerCampos
   End Select
        
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
    LimpiarCampos
    'Obtenemos la siguiente numero de factura
    'A�adiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    '###A mano
    'Text1(7).Text = NumF
    Text1(6).Text = Format(Now, "dd/mm/yyyy")
    
    ' ### [DavidV] 10/04/2006: Valores por defecto de los factores de correcci�n.
    Text1(11).Text = "1"
    Text1(12).Text = "1"
    PonerFoco Combo3
End Sub

Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "Select Max(n_reg_dosimetro) from dosimetros where tipo_dosimetro = " & Combo3.ListIndex
    sql = sql & " and sistema = '" & Text3.Text & "'"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, , , adCmdText
    sql = "1"
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then
            sql = CStr(rs.Fields(0) + 1)
        End If
    End If
    rs.Close
    SugerirCodigoSiguiente = sql
End Function

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
'        lblIndicador.Caption = "BUSCAR"
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarillo
        
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Text1(kCampo).SetFocus
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
  If Not (Data1.Recordset.BOF Or Data1.Recordset.EOF) Then
    Select Case Index
        Case 0
          Data1.Recordset.MoveFirst
        Case 1
          Data1.Recordset.MovePrevious
          If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
        Case 2
          Data1.Recordset.MoveNext
          If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
        Case 3
          Data1.Recordset.MoveLast
    End Select
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Combo3_Change
  End If
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    If Not BloqueaRegistroForm(Me) Then Exit Sub
   
    PonerModo 4
    'nro de registro
    Text1(7).Locked = True
    Text1(7).BackColor = &H80000018
    Combo3.Locked = True
    Combo3.BackColor = &H80000018
    Combo4.Locked = True
    Combo4.BackColor = &H80000018
    DespalzamientoVisible False
    cmdCancelar.Caption = "&Cancelar"
    
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    '******* canviar el mensage i la cadena *********************
    Cad = "Seguro que desea eliminar el dosimetro:" & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "de " & Combo3.Text
    '**********************************************************
    I = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2, "�Atenci�n!")
    
   'Borramos
    If I <> vbYes Then
        DesbloqueaRegistroForm1 Me
        Exit Sub
    End If
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    If Not Eliminar Then Exit Sub
   
    NumRegElim = Data1.Recordset.AbsolutePosition
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        PonerModo 0
        Else
            If NumRegElim > Data1.Recordset.RecordCount Then
                Data1.Recordset.MoveLast
            Else
                Data1.Recordset.MoveFirst
                Data1.Recordset.Move NumRegElim - 1
            End If
            PonerCampos
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Dos�metro"
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

If Data1.Recordset.EOF Then
    MsgBox "Ning�n registro devuelto.", vbExclamation, "�Atenci�n!"
    Exit Sub
End If

Cad = ""
I = 0
Do
    J = I + 1
    I = InStr(J, DatosADevolverBusqueda, "|")
    If I > 0 Then
        Aux = Mid(DatosADevolverBusqueda, J, I - J)
        J = Val(Aux)
        Cad = Cad & Text1(J).Text & "|"
    End If
Loop Until I = 0
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub


Private Sub Combo1_LostFocus()
    If Combo1.ListIndex = 0 Then Text1(1).Text = "01"
    If Combo1.ListIndex = 1 Then Text1(1).Text = "02"
End Sub


Private Sub Combo2_LostFocus()
    If Combo2.ListIndex = 0 Then Text1(5).Text = "P"
    If Combo2.ListIndex = 1 Then Text1(5).Text = "I"
End Sub


Private Sub Combo3_LostFocus()
Dim NF As Long
    If Modo = 3 Then
        If Combo3.ListIndex <> -1 And Combo4.ListIndex <> -1 Then
            NF = SugerirCodigoSiguiente
            Text1(7).Text = NF
        End If
    End If
    Combo3_Change
End Sub

Private Sub Combo4_LostFocus()
  Combo4_Change
  Combo3_LostFocus
  '-- VRS:1.3.1
  If Combo4.ListIndex = 1 Then
        '-- Panasonic
        Label8.Visible = True
        Line2.Visible = True
        Text1(14).Visible = True
        Text1(15).Visible = True
        Label1(11).Visible = True
        Label1(12).Visible = True
  Else
        '-- No panasonic
        Label8.Visible = False
        Line2.Visible = False
        Text1(14).Visible = False
        Text1(15).Visible = False
        Label1(11).Visible = False
        Label1(12).Visible = False
  End If
End Sub

Private Sub Combo4_Change()
  If Combo4.ListIndex <> -1 Then
    Text3.Text = Left(Combo4.Text, 1)
  End If
End Sub

Private Sub Combo3_Change()
  If Combo3.ListIndex = -1 Then
    Label1(7).Caption = "Cristal A"
    Label1(8).Caption = "Cristal B"
    Frame4.Visible = False
    
  ElseIf Combo3.ListIndex = 1 Then 'Or Combo4.ListIndex = 1 Then
    Label1(7).Caption = "Cristal 1"
    Label1(8).Caption = "Cristal 2"
    Frame4.Visible = True
  Else
    Label1(7).Caption = "Cristal 2"
    Label1(8).Caption = "Cristal 3"
    Frame4.Visible = False
  End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer


'    Me.Top = 0
'    Me.Left = 0
    PulsadoSalir = False

      ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    

    LimpiarCampos
    
    ' Protegemos el factor de correcci�n de los cristales si no es admin.
    If vUsu.NivelUsu < 2 Then
      Text1(11).Enabled = False
      Text1(12).Enabled = False
    End If
        
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
    End If
    
        
    '***** canviar el nom de la taula i el ORDER BY ********
    NombreTabla = "dosimetros"
    Ordenacion = " ORDER BY n_reg_dosimetro"
    '******************************************************+
        
'    PonerOpcionesMenu
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
'    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    End If
    
    CargarCombo
        
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Combo3.Text = ""
    Combo4.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If

    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(7), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        'If CadB <> "" Then CadB = CadB & " AND "
        'CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    PulsadoSalir = True
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub frmEmp_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmIns_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(2).Text = RecuperaValor(CadenaSeleccion, 2)
'        Text2(2).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub frmOpe_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(9).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(3).Text = RecuperaValor(CadenaSeleccion, 2)
'        Text2(4).Text = RecuperaValor(CadenaSeleccion, 3)
'        Text2(5).Text = RecuperaValor(CadenaSeleccion, 4)
    End If
End Sub

Private Sub frmTTR_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(4).Text = RecuperaValor(CadenaSeleccion, 2)
'        Text2(0).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub ImgMed_Click(Index As Integer)

End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    If Modo = 0 Or Modo = 2 Then Exit Sub
    Select Case Index
       Case 0
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(6).Text <> "" Then
                If IsDate(Text1(6).Text) Then f = Text1(6).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(6).Text = frmC.fecha
                mTag.DarFormato Text1(6)
            End If
            Set frmC = Nothing
       Case 1
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(10).Text <> "" Then
                If IsDate(Text1(10).Text) Then f = Text1(10).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(10).Text = frmC.fecha
                mTag.DarFormato Text1(10)
            End If
            Set frmC = Nothing
        Case 2 ' codigo de empresa
            Set frmEmp = New frmEmpresas
            frmEmp.DatosADevolverBusqueda = "0|1|"
            frmEmp.Show
        Case 4 ' instalacion
            Set frmIns = New frmInstalaciones
            frmIns.DatosADevolverBusqueda = "0|13|1|"
            frmIns.Show
        Case 3 ' tipo de trabajo
            Set frmTTr = New frmTiposTrab
            frmTTr.DatosADevolverBusqueda = "0|1|2|3|4|"
            frmTTr.Show
        Case 5 ' operarios
            Set frmOpe = New frmOperarios
            frmOpe.DatosADevolverBusqueda = "9|13|10|5|"
            frmOpe.Show
        Case 6
          Set frmTEx = New frmTiposExtremidades
          frmTEx.DatosADevolverBusqueda = "0|1|"
          frmTEx.Show
   End Select
End Sub






'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim sql As String

    kCampo = Index
    
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
    Else
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
   Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim I As Integer
    Dim sql As String
    Dim mTag As CTag
    Dim valor As Currency
    ''Quitamos blancos por los lados
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Modo = 1 And ConCaracteresBusqueda(Text1(Index).Text) Then Exit Sub
    
    Select Case Index
        Case 0, 2, 4, 7, 8, 9, 13
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el car�cter ' en ese campo.", vbExclamation, "�Atenci�n!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            If Modo = 1 And Index <> 13 Then Exit Sub
            Select Case Index
                Case 0 'empresa
                    Text2(1).Text = ""
                    If Text1(Index).Text <> "" Then
                      Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
                      If Text2(1).Text = "" Then
                          MsgBox "El c�digo de empresa no existe. Reintroduzca.", vbExclamation, "�Error!"
                          Text1(Index).Text = ""
                          PonerFoco Text1(Index)
                      End If
                    End If
                Case 4 ' tipos de trabajo
                    Text2(0).Text = ""
                    If Text1(Index).Text <> "" Then
                        Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", Text1(Index).Text & "|", "T|", 1)
                        If Text2(0).Text = "" Then
                            MsgBox "El c�digo de tipo de trabajo no existe. Reintroduzca.", vbExclamation, "�Error!"
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                Case 2 ' instalacion
                    Text2(2).Text = ""
                    If Text1(Index).Text <> "" And Text1(0).Text <> "" Then
                        Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|f_alta|", Text1(0).Text & "|" & Text1(Index).Text & "|", "T|T|", 2)
                        If Text2(2).Text = "" Then
                            MsgBox "El c�digo de instalacion no existe. Reintroduzca.", vbExclamation, "�Error!"
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                Case 9 ' dni de operario
                    If Modo = 1 Then Exit Sub
                    
                    CargarDatosOperarios Text1(9).Text, ape1, ape2, nombre
                    Text2(3).Text = ape1
                    Text2(4).Text = ape2
                    Text2(5).Text = nombre
                Case 13
                    Text2(11).Text = ""
                    Text2(11).Text = DevuelveDesdeBD(1, "descripcion", "tipmedext", "c_tipo_med|", Text1(Index).Text & "|", "T|", 1)
                    If Text2(11).Text = "" And Text1(13).Text <> "" Then
                        MsgBox "El c�digo de tipo de medici�n no existe. Reintroduzca.", vbExclamation, "�Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                    
            End Select
        
        Case 11, 12
          ' ### [DavidV] 10/04/2006: Los factores de correcci�n de cada dos�metro.
          If EsNumerico(Text1(Index).Text) Then
            If InStr(1, Text1(Index).Text, ",") > 0 Then
              valor = ImporteFormateado(Text1(Index).Text)
            Else
              valor = CCur(TransformaPuntosComas(Text1(Index).Text))
            End If
        
            Text1(Index).Text = Format(valor, "##0.000")
          End If
        
        Case 6, 10
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation, "�Error!"
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
            End If
    End Select
    Text1(Index).Text = Format(Text1(Index).Text, ">")
    '---
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String

Combo1.Enabled = False
Combo2.Enabled = False
Combo4.Enabled = False

CadB = ObtenerBusqueda(Me)
Combo1.Enabled = True
Combo2.Enabled = True
Combo4.Enabled = True

If CadB = "" Then
    MsgBox vbCrLf & "  Debe introducir alguna condici�n de b�squeda. " & vbCrLf, vbExclamation, "�Atenci�n!"
    PonerModo 0
    Exit Sub
End If

If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Combo3_LostFocus
        End If
End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
Dim tabla As String
Dim Titulo As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(7), 12, "N.Registro")
        Cad = Cad & ParaGrid(Text1(8), 12, "Dosimetro")
        Cad = Cad & ParaGrid(Text1(0), 15, "Empresa")
        Cad = Cad & ParaGrid(Text1(2), 16, "Instalacion")
        Cad = Cad & ParaGrid(Text1(9), 15, "DNI Operario")
        Cad = Cad & ParaGrid(Text1(3), 30, "Observaciones")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|3|4|"
            frmB.vTitulo = "Dos�metros a Cuerpo"
            frmB.vSelElem = 0
            frmB.vConexionGrid = 1
            frmB.vCargaFrame = False
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                Text1(kCampo).SetFocus
            End If
        End If
        Screen.MousePointer = vbDefault

End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation, "�Atenci�n!"
    Screen.MousePointer = vbDefault
    PonerModo 0
    Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
End If

Screen.MousePointer = vbDefault
Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    Dim I As Integer
    Dim mTag As CTag
    Dim sql As String
    If Data1.Recordset.EOF Then Exit Sub
    
    If PonerCamposForma(Me, Data1) Then
        Combo1.ListIndex = CInt(Text1(1).Text) - 1
        Combo2.ListIndex = -1
        If Text1(5).Text = "P" Then Combo2.ListIndex = 0
        If Text1(5).Text = "I" Then Combo2.ListIndex = 1
    End If
    
    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", Text1(4).Text & "|", "T|", 1)
    Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Text1(0).Text & "|" & Text1(2).Text & "|", "T|T|", 2)
    Text2(11).Text = DevuelveDesdeBD(1, "descripcion", "tipmedext", "c_tipo_med|", Text1(13).Text & "|", "T|", 1)
    CargarDatosOperarios Text1(9).Text, ape1, ape2, nombre
    Text2(3).Text = ape1
    Text2(4).Text = ape2
    Text2(5).Text = nombre
'    If Combo4.Text = "H" Then
'      Combo4.ListIndex = 0
'    ElseIf Combo4.Text = "P" Then
'      Combo4.ListIndex = 1
'    End If
    Frame4.Visible = Combo3.ListIndex = 1
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    '-- VRS: 1.3.1
    Combo4_LostFocus
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Dim b As Boolean
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    PonerIndicador lblIndicador, Modo
    If Modo = 0 Then LimpiarCampos
    
    Text1(0).Enabled = True
    
    chkVistaPrevia.Visible = True
    
    
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For I = 0 To Text1.Count - 1
            Text1(I).BackColor = &H80000018
        Next I
    End If
    
    b = (Modo = 0) Or (Modo = 2)
    Toolbar1.Buttons(6).Enabled = b
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    Toolbar1.Buttons(11).Enabled = b
    
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    Else
        cmdRegresar.Visible = False
    End If
    
    'Modo insertar o modificar
    b = (Kmodo >= 3)  '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = b Or Modo = 1
    cmdCancelar.Visible = b Or Modo = 1
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(1).Enabled = Not b And Modo <> 1
    Toolbar1.Buttons(2).Enabled = Not b And Modo <> 1
    
    '### A mano
    'Aqui a�adiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    b = (Modo = 2) Or Modo = 0
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = b
        Text1(I).BackColor = vbWhite
    Next I
    
    For I = 0 To Imgppal.Count - 1
        Imgppal(I).Enabled = Not b
    Next I
    Combo1.Enabled = Not b
    Combo2.Enabled = Not b
    Combo3.Enabled = Not b
    Combo4.Enabled = Not b
    Frame4.Visible = Not b And Combo3.ListIndex = 1
    
'    If Combo4.Text = "H" Then
'      Combo4.ListIndex = 0
'    ElseIf Combo4.Text = "P" Then
'      Combo4.ListIndex = 1
'    End If
    
    PonerFoco chkVistaPrevia
End Sub

Private Function DatosOk() As Boolean
Dim rs As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim Datos As String
Dim Cad As String
    
    b = CompForm(Me)
    IndiceErroneo = 0
    If (b = True) And ((Modo = 3) Or (Modo = 4)) Then
'        For I = 0 To Text1.Count - 1
'             If InStr(1, Text1(I).Text, "'") > 0 Then
'                MsgBox "No puede introducir el car�cter ' en ese campo.", vbExclamation, "�Atenci�n!"
'                IndiceErroneo = I
'                DatosOk = False
'                Exit Function
'             End If
'        Next I

        ' comprobamos la integridad de la bd
        ' existe la instalacion
        Datos = ""
        Datos = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Trim(Text1(0).Text) & "|" & Trim(Text1(2).Text) & "|", "T|T|", 2)
        If Datos = "" Then
            If MsgBox("No existe la Instalaci�n para la Empresa. Desea continuar.", vbQuestion + vbYesNo + vbDefaultButton2, "�Atenci�n!") = vbNo Then
                DatosOk = False
                Exit Function
            End If
        End If
        ' exite el operario en la instalacion  introducida
        Datos = ""
        Datos = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", Text1(0).Text & "|" & Text1(2).Text & " |" & Text1(9).Text & "|", "T|T|T|", 3)
        If Datos = "" Then
            If MsgBox("No existe el operario en la instalaci�n introducida. Desea continuar.", vbQuestion + vbYesNo + vbDefaultButton1, "�Atenci�n!") = vbNo Then
                DatosOk = False
                Exit Function
            End If
        End If
        If Text1(6).Text <> "" And Text1(10).Text <> "" Then
            If CDate(Text1(6).Text) > CDate(Text1(10).Text) Then
                MsgBox "La Fecha de Retirada no puede ser inferior a la de Asignaci�n", vbExclamation, "�Atenci�n!"
                DatosOk = False
                Exit Function
            End If
        End If
        
        If Text1(8).Text <> "" And (Combo3.ListIndex <> -1) Then
            If DosimetroEnUso(Text1(8).Text, CLng(Text1(7).Text), Combo3.ListIndex, Text3.Text) Then
                MsgBox "Este dosimetro est� asignado. Revise.", vbExclamation, "�Atenci�n!"
                DatosOk = False
                Exit Function
            End If
        End If
        
    End If

If (b = True) And (Modo = 3) Then
    'Estamos insertando
    'a�o es com posar: select codvarie from svarie where codvarie = txtAux(0)
    'la N es pa dir que es numeric
     Datos = DevuelveDesdeBD(1, "n_reg_dosimetro", "dosimetros", "n_reg_dosimetro|tipo_dosimetro|sistema|", Text1(7).Text & "|" & Combo3.ListIndex & "|" & Text3.Text, "N|N|T|", 3)
     If Datos <> "" Then
        MsgBox "Ya existe el n�mero de registro de dosimetro : " & Text1(7).Text, vbExclamation, "�Atenci�n!"
        DatosOk = False
        IndiceErroneo = 7
        Exit Function
    End If
End If

DatosOk = b
End Function

Private Sub Text3_Change()
  If Text3.Text = "H" Then
    Combo4.ListIndex = 0
  ElseIf Text3.Text = "P" Then
    Combo4.ListIndex = 1
  Else
    Combo4.ListIndex = -1
  End If
End Sub

'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonBuscar
        Case 2
            BotonVerTodos
        Case 6
            BotonAnyadir
        Case 7
            BotonModificar
        Case 8
            BotonEliminar
            
        Case 12
            mnSalir_Click
        Case 14 To 17
            Desplazamiento (Button.Index - 14)
        Case 11
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 4 'Listado de dosimetros a cuerpo
            FrmListado.Show
        
        Case Else
    
    End Select
End Sub

Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub CargarCombo()

'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Si, 1-No
    Combo3.Clear
    Combo3.AddItem "Cuerpo"
    Combo3.ItemData(Combo3.NewIndex) = 0
    
    Combo3.AddItem "Organo"
    Combo3.ItemData(Combo3.NewIndex) = 1
    
    Combo3.AddItem "Area"
    Combo3.ItemData(Combo3.NewIndex) = 2
    
    Combo4.Clear
    Combo4.AddItem "Harshaw"
    Combo4.ItemData(Combo4.NewIndex) = 0
    Combo4.AddItem "Panasonic"
    Combo4.ItemData(Combo4.NewIndex) = 1
    
'    Combo3.AddItem "Fondo"
'    Combo3.ItemData(Combo3.NewIndex) = 3
'
'    Combo3.AddItem "Tr�nsito"
'    Combo3.ItemData(Combo3.NewIndex) = 4
'
'    Combo3.AddItem "Control"
'    Combo3.ItemData(Combo3.NewIndex) = 5
'
'    Combo3.AddItem "Libre"
'    Combo3.ItemData(Combo3.NewIndex) = 6
    
    Combo2.Clear
    Combo2.AddItem "Par"
    Combo2.ItemData(Combo2.NewIndex) = 0

    Combo2.AddItem "Impar"
    Combo2.ItemData(Combo2.NewIndex) = 1

    
    Combo1.Clear
    Combo1.AddItem "Plantilla"
    Combo1.ItemData(Combo1.NewIndex) = 0

    Combo1.AddItem "Contrata"
    Combo1.ItemData(Combo1.NewIndex) = 1

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu
        Case "Listado"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 2 'Listado de articulos
            FrmListado.Show
        Case "Etiquetas"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 17 'Listado de articulos
            FrmListado.Show
    End Select
End Sub


Private Function Eliminar() As Boolean
Dim I As Integer
Dim sql As String

        sql = " WHERE n_reg_dosimetro=" & Data1.Recordset!n_reg_dosimetro & " and tipo_dosimetro = " & Combo3.ListIndex
        
        conn.Execute "Delete  from dosimetros " & sql
       
        Eliminar = True
        
End Function

Private Sub PideCalculadora()
On Error GoTo EPideCalculadora
    Shell App.Path & "\arical.exe", vbNormalFocus
    Exit Sub
EPideCalculadora:
    Err.Clear
End Sub

Private Sub AsignarTeclasFuncion(key As Integer)

    If Modo = 2 Or Modo = 0 Then
        Select Case key
            Case vbESC '27
                If Modo = 0 Then
                    Toolbar1_ButtonClick Toolbar1.Buttons(12)
                Else
                    PonerModo 0
                End If
            Case vbAnterior '33
                If Modo = 2 Then Desplazamiento (1)
            Case vbSiguiente '34
                If Modo = 2 Then Desplazamiento (2)
            Case vbPrimero  ' 36 ' inicio
                If Modo = 2 Then Desplazamiento (0)
            Case vbUltimo '35 ' fin
                If Modo = 2 Then Desplazamiento (3)
           Case vbBuscar
                Toolbar1_ButtonClick Toolbar1.Buttons(1)
           Case vbVerTodos
                Toolbar1_ButtonClick Toolbar1.Buttons(2)
           Case vbA�adir
                Toolbar1_ButtonClick Toolbar1.Buttons(6)
           Case vbModificar
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(7)
            Case vbEliminar
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(8)
            Case vbLineas
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(10)
            Case vbImprimir
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(13)
            Case vbSalir
                  Toolbar1_ButtonClick Toolbar1.Buttons(12)
        End Select
   End If

End Sub

Private Function SituarData1() As Boolean
Dim clone_rs As ADODB.Recordset

    Dim sql As String
    On Error GoTo ESituarData1
        'Actualizamos el recordset
        Data1.Refresh
        
        ' ### [DavidV] 10/04/2006: Peque�a soluci�n improvisada, ya que el Find del recordset
        ' s�lo admite una columna como criterio de b�squeda, y aqu� necesitamos m�s para situar
        ' el cursor en el registro correcto.

        '********* canviar la clau primaria codsocio per la que siga *********
        'El sql para que se situe en el registro en especial es el siguiente
        sql = "n_reg_dosimetro = " & Text1(7).Text & " AND n_dosimetro = '" & Text1(8) & "'"
        sql = sql & " AND tipo_dosimetro = " & Combo3.ListIndex
        '*****************************************************************
        
        Set clone_rs = Data1.Recordset.Clone
        clone_rs.Filter = sql
        If clone_rs.BOF Or clone_rs.EOF Then
          Data1.Recordset.MoveLast
          Data1.Recordset.MoveNext
        Else
          Data1.Recordset.Bookmark = clone_rs.Bookmark
        End If
        'Data1.Recordset.Find SQL
        If Data1.Recordset.EOF Then GoTo ESituarData1
        SituarData1 = True
        clone_rs.Close
        Set clone_rs = Nothing
    Exit Function

ESituarData1:
    clone_rs.Close
    Set clone_rs = Nothing
    Err.Clear
    Limpiar Me
    PonerModo 0
    SituarData1 = False
End Function



