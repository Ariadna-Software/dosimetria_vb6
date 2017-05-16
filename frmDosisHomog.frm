VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDosisHomog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dosis Homogéneas (Cuerpo)"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9315
   Icon            =   "frmDosisHomog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   9315
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   7
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "N.Registro|N|N|||dosiscuerpo|n_registro||S|"
      Text            =   "Text1"
      Top             =   540
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   1500
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "Migrado|T|S|||dosiscuerpo|migrado|||"
      Text            =   "Text1"
      Top             =   5580
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Height          =   915
      Left            =   180
      TabIndex        =   52
      Top             =   810
      Width           =   9000
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   5520
         MaxLength       =   40
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   7380
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "N.Dosimetro|T|N|||dosiscuerpo|n_dosimetro|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   3700
         MaxLength       =   40
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   180
         MaxLength       =   40
         TabIndex        =   2
         Tag             =   "Codigo Empresa|T|N|||dosiscuerpo|n_reg_dosimetro|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1935
         MaxLength       =   30
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Par/Impar"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   8
         Left            =   5520
         TabIndex        =   60
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Retirada"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   7
         Left            =   3700
         TabIndex        =   58
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Asignación"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   6
         Left            =   1935
         TabIndex        =   57
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Dosímetro"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   7380
         TabIndex        =   56
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Registro Dosímetro"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   5
         Left            =   180
         TabIndex        =   55
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1500
      MaxLength       =   120
      TabIndex        =   15
      Tag             =   "Observaciones|T|S|||dosiscuerpo|observaciones|||"
      Text            =   "Text1"
      Top             =   5250
      Width           =   7680
   End
   Begin VB.Frame Frame2 
      Height          =   885
      Left            =   180
      TabIndex        =   40
      Top             =   3360
      Width           =   9000
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Fecha Migracion|F|N|||dosiscuerpo|f_migracion|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   180
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Fecha Dosis|F|N|||dosiscuerpo|f_dosis|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   3700
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Dosis Superficial|N|S|0.00|999.99|dosiscuerpo|dosis_superf|##0.00||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   5535
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Dosis Profunda|N|N|0,00|999,99|dosiscuerpo|dosis_profunda|##0.00||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1425
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7410
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   51
         Tag             =   "Plantilla/Contrata|T|N|||dosiscuerpo|plantilla_contrata|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1245
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   7
         Left            =   180
         MouseIcon       =   "frmDosisHomog.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":0E1C
         ToolTipText     =   "Seleccionar fecha"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   6
         Left            =   1935
         MouseIcon       =   "frmDosisHomog.frx":0EA7
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":0FF9
         ToolTipText     =   "Seleccionar fecha"
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Plantilla/Contrata:"
         Height          =   255
         Left            =   7410
         TabIndex        =   50
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label13 
         Caption         =   "Dosis Profunda:"
         Height          =   255
         Left            =   5535
         TabIndex        =   44
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha Dosis"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   525
         TabIndex        =   43
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "Fec.Migración"
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label8 
         Caption         =   "Dosis Superficial:"
         Height          =   255
         Left            =   3700
         TabIndex        =   41
         Top             =   210
         Width           =   1500
      End
   End
   Begin VB.Frame Frame7 
      Height          =   795
      Left            =   180
      TabIndex        =   30
      Top             =   1710
      Width           =   9000
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1515
         MaxLength       =   30
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   420
         Width           =   2970
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   180
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Codigo Empresa|T|N|||dosiscuerpo|c_empresa|||"
         Text            =   "Text1"
         Top             =   420
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   5850
         MaxLength       =   40
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   420
         Width           =   3030
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   4530
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "Codigo Instalación|T|N|||dosiscuerpo|c_instalacion|||"
         Text            =   "Text1"
         Top             =   420
         Width           =   1290
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   2
         Left            =   180
         MouseIcon       =   "frmDosisHomog.frx":1084
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":11D6
         ToolTipText     =   "Buscar empresa"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   38
         Top             =   180
         Width           =   915
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   4
         Left            =   4530
         MouseIcon       =   "frmDosisHomog.frx":12D8
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":142A
         ToolTipText     =   "Buscar instalación"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Instalación"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   4875
         TabIndex        =   36
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame Frame5 
      Height          =   825
      Left            =   180
      TabIndex        =   23
      Top             =   2520
      Width           =   9000
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   180
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "DNI|T|N|||dosiscuerpo|dni_usuario|||"
         Text            =   "Text1"
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   6240
         MaxLength       =   20
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   450
         Width           =   2640
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   3660
         MaxLength       =   20
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   450
         Width           =   2505
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   450
         Width           =   1935
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   5
         Left            =   180
         MouseIcon       =   "frmDosisHomog.frx":152C
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":167E
         ToolTipText     =   "Buscar D.N.I."
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "D.N.I."
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   570
         TabIndex        =   34
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label Label9 
         Caption         =   "Primer Apellido:"
         Height          =   195
         Left            =   1590
         TabIndex        =   28
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label17 
         Caption         =   "Segundo Apellido:"
         Height          =   255
         Left            =   3660
         TabIndex        =   27
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   180
      TabIndex        =   31
      Top             =   4260
      Width           =   9000
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   180
         MaxLength       =   5
         TabIndex        =   12
         Tag             =   "Rama Generica|T|N|||dosiscuerpo|rama_generica|||"
         Text            =   "Text1"
         Top             =   450
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   570
         MaxLength       =   30
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   450
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   13
         Tag             =   "Rama Específica|T|N|||dosiscuerpo|rama_especifica|||"
         Text            =   "Text1"
         Top             =   450
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3030
         MaxLength       =   30
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   450
         Width           =   2490
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5940
         MaxLength       =   40
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   450
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   5580
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Tipo de Trabajo|T|N|||dosiscuerpo|c_tipo_trabajo|||"
         Text            =   "Text1"
         Top             =   450
         Width           =   360
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   1
         Left            =   180
         MouseIcon       =   "frmDosisHomog.frx":1780
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":18D2
         ToolTipText     =   "Buscar rama genérica"
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Rama Genérica:"
         Height          =   255
         Left            =   495
         TabIndex        =   48
         Top             =   210
         Width           =   1155
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   0
         Left            =   2640
         MouseIcon       =   "frmDosisHomog.frx":19D4
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":1B26
         ToolTipText     =   "Buscar rama específica"
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Rama Específica:"
         Height          =   255
         Left            =   2970
         TabIndex        =   47
         Top             =   210
         Width           =   1335
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   3
         Left            =   5580
         MouseIcon       =   "frmDosisHomog.frx":1C28
         MousePointer    =   99  'Custom
         Picture         =   "frmDosisHomog.frx":1D7A
         ToolTipText     =   "Buscar tipo de trabajo"
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Trabajo:"
         Height          =   255
         Left            =   5880
         TabIndex        =   33
         Top             =   210
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7950
      TabIndex        =   19
      Top             =   6060
      Width           =   1110
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7995
      TabIndex        =   18
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   270
      TabIndex        =   20
      Top             =   6000
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6780
      TabIndex        =   17
      Top             =   6060
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   540
      Top             =   6060
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
      TabIndex        =   22
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
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
            Object.ToolTipText     =   "Último"
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
   Begin VB.Label Label10 
      Caption         =   "Migrado CSN"
      Height          =   255
      Left            =   330
      TabIndex        =   61
      Top             =   5610
      Width           =   1170
   End
   Begin VB.Label Label3 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   330
      TabIndex        =   49
      Top             =   5280
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Número Registro"
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
      Left            =   240
      TabIndex        =   39
      Top             =   540
      Width           =   1785
   End
End
Attribute VB_Name = "frmDosisHomog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
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
Private WithEvents frmRGe As frmRamasGener
Attribute frmRGe.VB_VarHelpID = -1
Private WithEvents frmREs As frmRamasEspe
Attribute frmREs.VB_VarHelpID = -1

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

Dim CadB As String
Dim ape1 As String
Dim ape2 As String
Dim nombre As String


' campo que indica si la familia es fitosanitaria
' si lo es: obligamos a introducir los campos de fitos.
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

Dim rs As ADODB.Recordset


Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
    AsignarTeclasFuncion KeyCode
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    Dim v_aux As Integer
    Dim sql As String
    
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                 PonerModo 0
            End If
        End If
    
    Case 4
        'Modificar
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If ModificaDesdeFormulario(Me, 1) Then
                DesbloqueaRegistroForm1 Me
                If SituarData1 Then
                    PonerModo 2
                    PonerCampos
                Else
                    LimpiarCampos
                    PonerModo 0
                End If
            End If
        End If

    Case 1
        HacerBusqueda
    End Select

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation, "¡Error!"
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
    NumF = SugerirCodigoSiguiente
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    '###A mano
    Text1(7).Text = NumF
    Text1(13).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(14)
End Sub

Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "Select Max(n_registro) from dosiscuerpo"
    
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
    DespalzamientoVisible False
    cmdCancelar.Caption = "&Cancelar"
    
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    '******* canviar el mensage i la cadena *********************
    Cad = "Seguro que desea eliminar la dosis homogénea:" & Data1.Recordset.Fields(0)
    '**********************************************************
    I = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!")
    
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
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Dosis de cuerpo"
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

If Data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation, "¡Atención!"
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

Private Sub Combo1_LostFocus()
    If Combo1.ListIndex = 0 Then Text1(1).Text = "01"
    If Combo1.ListIndex = 1 Then Text1(1).Text = "02"
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer


    Me.Top = 0
    Me.Left = 0
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
    
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
    End If
    '***** canviar el nom de la taula i el ORDER BY ********
    NombreTabla = "dosiscuerpo"
    Ordenacion = " ORDER BY n_registro"
    '******************************************************+
        
'    PonerOpcionesMenu
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
   ' Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        PonerFoco Text1(7)
        Text1(7).BackColor = vbYellow
    End If
    
    CargarCombo
    
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
    
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
'    Dim cadB As String
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
        Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
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

Private Sub frmREs_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(7).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Trim(Text1(5).Text) & "|", "T|", 1)
        Text1(6).Text = RecuperaValor(CadenaSeleccion, 2)
'        Text2(6).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub frmrge_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
    End If

End Sub

Private Sub frmTTR_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(7).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Trim(Text1(5).Text) & "|", "T|", 1)
        Text1(4).Text = RecuperaValor(CadenaSeleccion, 2)
        Text2(0).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    If Modo = 0 Or Modo = 2 Then Exit Sub
    Select Case Index
       Case 6 'fecha de migracion
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
       Case 7 ' fecha de dosis
            f = Now
            If Text1(13).Text <> "" Then
                If IsDate(Text1(13).Text) Then f = Text1(13).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(13).Text = frmC.fecha
                mTag.DarFormato Text1(13)
            End If
            Set frmC = Nothing
        Case 0 ' rama especifica
            Set frmREs = New frmRamasEspe
            frmREs.DatosADevolverBusqueda = "0|2|3|"
            frmREs.Show
        Case 1 ' rama generica
            Set frmRGe = New frmRamasGener
            frmRGe.DatosADevolverBusqueda = "0|1|"
            frmRGe.Show
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
    Dim Consulta As ADODB.Recordset
    ''Quitamos blancos por los lados
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" Then Exit Sub
    
    If Modo = 1 And ConCaracteresBusqueda(Text1(Index).Text) Then Exit Sub
    
    ' Para que aparezcan automáticamente los datos en la dosis.
    If Index = 8 Or Index = 14 Then
      sql = "select dni_usuario,c_empresa,c_instalacion,c_tipo_trabajo,plantilla_contrata from dosimetros "
      sql = sql & "where n_reg_dosimetro = " & Val(Text1(14).Text) & " and n_dosimetro = '"
      sql = sql & Text1(8).Text & "' and tipo_dosimetro = 0"
      Set Consulta = New ADODB.Recordset
      Consulta.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
      
      If Not Consulta.EOF Then
        Text1(0).Text = Consulta!c_empresa & ""
        Text1(2).Text = Consulta!c_instalacion & ""
        Text1(9).Text = Consulta!dni_usuario & ""
        Text1(4).Text = Consulta!c_tipo_trabajo & ""
        Combo1.ListIndex = Val(Consulta!plantilla_contrata & "") - 1
        Text1_LostFocus 0
        Text1_LostFocus 2
        Text1_LostFocus 9
        Text1_LostFocus 4
        Text1(10).SetFocus
      Else
        Text1(0).Text = ""
        Text1(2).Text = ""
        Text1(9).Text = ""
        Text1(4).Text = ""
        For I = 0 To 5
          Text2(I).Text = ""
        Next
        Combo1.ListIndex = -1
      End If
      Set Consulta = Nothing
    End If
    
    Select Case Index
        Case 7, 8, 11, 12, 14
            'valores numericos
             If Text1(Index).Text <> "" Then
                If EsNumerico(Text1(Index).Text) Then
                    If Modo = 1 Then Exit Sub
                    Select Case Index
                        Case 7
                            Text1(Index).Text = Format(Text1(Index).Text, "00000000")
                        Case 11, 12
                            If InStr(1, Text1(Index).Text, ",") > 0 Then
                                valor = ImporteFormateado(Text1(Index).Text)
                            Else
                                valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                            End If
                            
                            Text1(Index).Text = Format(valor, "##0.00")
                        Case 14
                            Text1(Index).Text = Format(Text1(Index).Text, "00000000")
                            CargarDatosDosimetros Text1(14).Text
                    End Select
                 End If
             End If
        Case 0, 2, 4, 5, 6, 9, 15
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            
            If Modo = 1 Then Exit Sub
            Select Case Index
                Case 0 'empresa
                    Text2(1).Text = ""
                    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(Index).Text & "|", "T|", 1)
                    If Text2(1).Text = "" Then
                        MsgBox "El código de empresa no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                Case 4 ' tipos de trabajo
                    Text2(0).Text = ""
                    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", Text1(Index).Text & "|", "T|", 1)
                    If Text2(0).Text = "" Then
                        MsgBox "El código de tipo de trabajo no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                Case 2 ' instalacion
                    If Text1(Index).Text <> "" And Text1(0).Text <> "" Then
                        Text2(2).Text = ""
                        Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Text1(0).Text & "|" & Text1(2).Text & "|", "T|T|", 2)
                        If Text2(2).Text = "" Then
                            MsgBox "El código de instalacion no existe. Reintroduzca.", vbExclamation, "¡Error!"
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                Case 5 ' rama generica
                    Text2(7).Text = ""
                    Text2(7).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(Index).Text & "|", "T|", 1)
                    If Text2(7).Text = "" Then
                        MsgBox "El código de rama generica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                Case 6 ' rama especifica
                    Text2(6).Text = ""
                    If Text1(5).Text <> "" Then
                        Text2(6).Text = DevuelveDesdeBD(1, "descripcion", "ramaespe", "cod_rama_gen|c_rama_especifica|", Text1(5).Text & "|" & Text1(Index).Text & "|", "T|T|", 2)
                        If Text2(6).Text = "" Then
                            MsgBox "El código de rama especifica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                
                Case 9 ' operario
                    CargarDatosOperarios Text1(9).Text, ape1, ape2, nombre
                    Text2(3).Text = ape1
                    Text2(4).Text = ape2
                    Text2(5).Text = nombre
                
            End Select
        Case 10, 13
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation, "¡Error!"
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
'Dim cadB As String
CadB = ObtenerBusqueda(Me)

If CadB = "" Then
    MsgBox vbCrLf & "  Debe introducir alguna condición de búsqueda. " & vbCrLf, vbExclamation, "¡Error!"
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
        Cad = Cad & ParaGrid(Text1(14), 12, "N.Reg.Dosimet.")
        Cad = Cad & ParaGrid(Text1(8), 12, "Dosimetro")
        Cad = Cad & ParaGrid(Text1(0), 15, "Empresa")
        Cad = Cad & ParaGrid(Text1(2), 16, "Instalacion")
        Cad = Cad & ParaGrid(Text1(9), 15, "DNI Operario")
        Cad = Cad & ParaGrid(Text1(3), 20, "Observaciones")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|3|4|"
            frmB.vTitulo = "Dosis Homogéneas (Cuerpo)"
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

'Set rs = New ADODB.Recordset
'
'Dim t1 As Single
't1 = Timer
'
'rs.Open CadenaConsulta, Conn, adOpenDynamic, adLockOptimistic
'
'Set Data1.Recordset = rs
'
'MsgBox Timer - t1
'rs.MoveLast
'rs.MoveFirst
'
'If rs.EOF Then
lblIndicador.Caption = "Leyendo campos BD"
lblIndicador.Refresh
DoEvents
Data1.RecordSource = CadenaConsulta
Data1.Refresh
lblIndicador.Caption = ""
lblIndicador.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation, "¡Atención!"
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
    End If
    
    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Trim(Text1(0).Text) & "|", "T|", 1)
    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", Trim(Text1(4).Text) & "|", "T|", 1)
    Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Trim(Text1(0).Text) & "|" & Trim(Text1(2).Text) & "|", "T|T|", 2)
    Text2(6).Text = DevuelveDesdeBD(1, "descripcion", "ramaespe", "cod_rama_gen|c_rama_especifica|", Trim(Text1(5).Text) & "|" & Trim(Text1(6).Text) & "|", "T|T|", 2)
    Text2(7).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Trim(Text1(5).Text) & "|", "T|", 1)
    
    CargarDatosOperarios Text1(9).Text, ape1, ape2, nombre
    Text2(3).Text = ape1
    Text2(4).Text = ape2
    Text2(5).Text = nombre
    CargarDatosDosimetros CLng(Data1.Recordset!n_reg_dosimetro)
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
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
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
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
'                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
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
            If MsgBox("No existe la Instalación para la Empresa. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") = vbNo Then
                DatosOk = False
                Exit Function
            End If
        End If
        ' exite el operario en la empresa introducida
        Datos = ""
        Datos = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", Text1(0).Text & "|" & Text1(2).Text & "|" & Text1(9).Text & "|", "T|T|T|", 3)
        If Datos = "" Then
            If MsgBox("No existe el operario en la empresa/instalacion introducida. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, "¡Atención!") = vbNo Then
                DatosOk = False
                Exit Function
            End If
        End If
        
        ' existe la rama especifica para la rama generica
        Datos = ""
        Datos = DevuelveDesdeBD(1, "descripcion", "ramaespe", "cod_rama_gen|c_rama_especifica|", Text1(5).Text & "|" & Text1(6).Text & "|", "T|T|", 2)
        If Datos = "" Then
            MsgBox "No existe la rama específica. Reintroduzca.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
        ' existe el tipo de trabajo para la rama generica
        Datos = ""
        Datos = DevuelveDesdeBD(1, "descripcion", "tipostrab", "cod_rama_gen|c_tipo_trabajo|", Trim(Text1(5).Text) & "|" & Trim(Text1(4).Text) & "|", "T|T|", 2)
        If Datos = "" Then
            MsgBox "No existe el tipo de trabajo. Reintroduzca.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
        
    End If

If (b = True) And (Modo = 3) Then
     Datos = DevuelveDesdeBD(1, "n_registro", "dosiscuerpo", "n_registro|", Text1(7).Text & "|", "N|", 1)
     If Datos <> "" Then
        MsgBox "Ya existe el número de registro de dosis homogénea : " & Text1(7).Text, vbExclamation, "¡Error!"
        DatosOk = False
        IndiceErroneo = 7
        Exit Function
    End If
End If

DatosOk = b
End Function

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
            printNou
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
    Combo1.Clear
    Combo1.AddItem "Plantilla"
    Combo1.ItemData(Combo1.NewIndex) = 0

    Combo1.AddItem "Contrata"
    Combo1.ItemData(Combo1.NewIndex) = 1

End Sub

Private Function Eliminar() As Boolean
Dim I As Integer
Dim sql As String

        sql = " WHERE n_registro=" & Data1.Recordset!N_registro
        
        conn.Execute "Delete  from dosiscuerpo " & sql
       
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
           Case vbAñadir
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
    Dim sql As String
    On Error GoTo ESituarData1
        'Actualizamos el recordset
        Data1.Refresh
        '********* canviar la clau primaria codsocio per la que siga *********
        'El sql para que se situe en el registro en especial es el siguiente
        sql = "n_registro = " & Text1(7).Text & ""
        '*****************************************************************
        Data1.Recordset.Find sql
        If Data1.Recordset.EOF Then GoTo ESituarData1
        SituarData1 = True
    Exit Function
ESituarData1:
    If Err.Number <> 0 Then Err.Clear
    Limpiar Me
    PonerModo 0
    SituarData1 = False
End Function

Private Sub CargarDatosDosimetros(nregistro As Long)
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "Select f_asig_dosimetro, f_retirada, mes_p_i from dosimetros "
    sql = sql & " where tipo_dosimetro = 0 and n_reg_dosimetro = " & Text1(14).Text & " and n_dosimetro = '" & Text1(8).Text & "'"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, , , adCmdText
    
    Text2(8).Text = ""
    Text2(9).Text = ""
    Text2(10).Text = ""
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0).Value) Then
            Text2(8).Text = rs.Fields(0).Value
        End If
        If Not IsNull(rs.Fields(1).Value) Then
            Text2(9).Text = rs.Fields(1).Value
        End If
        If rs.Fields(2).Value = "P" Then
            Text2(10).Text = "Par"
        Else
            Text2(10).Text = "Impar"
        End If
    End If
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "dosiscuerpo"
        .Informe2 = "DosisHomogeneas.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(Data1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{sdexpgrp.codempre} = " & codempre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={sdexpgrp.codsupdt}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False

        .Show vbModal
    End With
End Sub

