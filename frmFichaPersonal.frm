VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmFichaPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha Personal"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15060
   Icon            =   "frmFichaPersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   15060
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9960
      TabIndex        =   3
      Text            =   "Combo4"
      Top             =   615
      Width           =   1365
   End
   Begin VB.Frame Frame15 
      Caption         =   "INSTALACION DE TRABAJO"
      Height          =   2085
      Left            =   60
      TabIndex        =   93
      Top             =   5850
      Width           =   14925
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   43
         Left            =   4590
         MaxLength       =   40
         TabIndex        =   47
         Tag             =   "instalaciones|persona_contacto|S|"
         Text            =   "Text1"
         Top             =   1620
         Width           =   5400
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   44
         Left            =   10050
         MaxLength       =   40
         TabIndex        =   48
         Tag             =   "instalaciones|observaciones|S|"
         Text            =   "Text1"
         Top             =   1620
         Width           =   4710
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   40
         Left            =   12180
         MaxLength       =   10
         TabIndex        =   44
         Tag             =   "instalaciones|telefono|S|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1110
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   41
         Left            =   13350
         MaxLength       =   10
         TabIndex        =   45
         Tag             =   "instalaciones|fax|S|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   42
         Left            =   150
         MaxLength       =   50
         TabIndex        =   46
         Tag             =   "instalaciones|mail_internet|S|"
         Text            =   "Text1"
         Top             =   1620
         Width           =   4230
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   39
         Left            =   8460
         MaxLength       =   5
         TabIndex        =   43
         Tag             =   "instalaciones|distrito|S|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   36
         Left            =   150
         MaxLength       =   40
         TabIndex        =   40
         Tag             =   "instalaciones|direccion|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   4260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   38
         Left            =   8040
         MaxLength       =   5
         TabIndex        =   42
         Tag             =   "Instalaciones|Código Postal|N|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   37
         Left            =   4590
         MaxLength       =   30
         TabIndex        =   41
         Tag             =   "instalaciones|poblacion|S|"
         Text            =   "Text1"
         Top             =   1080
         Width           =   3420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   9375
         MaxLength       =   30
         TabIndex        =   99
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1080
         Width           =   2715
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   150
         MaxLength       =   40
         TabIndex        =   37
         Tag             =   "Instalaciones|Descripción|N|"
         Text            =   "Text1"
         Top             =   450
         Width           =   5880
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   34
         Left            =   6240
         MaxLength       =   5
         TabIndex        =   38
         Tag             =   "Instalaciones|Código de Rama Genérica|N|"
         Text            =   "Text1"
         Top             =   450
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   6630
         MaxLength       =   30
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   450
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   35
         Left            =   9960
         MaxLength       =   5
         TabIndex        =   39
         Tag             =   "Instalaciones|Código de Rama Especifica|N|"
         Text            =   "Text1"
         Top             =   450
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   10380
         MaxLength       =   30
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   450
         Width           =   4380
      End
      Begin VB.Label Label24 
         Caption         =   "Persona de Contacto:"
         Height          =   255
         Left            =   4620
         TabIndex        =   108
         Top             =   1410
         Width           =   1710
      End
      Begin VB.Label Label52 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   10050
         TabIndex        =   107
         Top             =   1410
         Width           =   1710
      End
      Begin VB.Label Label27 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   12180
         TabIndex        =   106
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label26 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   13380
         TabIndex        =   105
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Mail:"
         Height          =   255
         Left            =   210
         TabIndex        =   104
         Top             =   1410
         Width           =   615
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   12
         Left            =   8040
         MouseIcon       =   "frmFichaPersonal.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":0E1C
         ToolTipText     =   "Buscar código postal"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label51 
         Caption         =   "C.Postal:"
         Height          =   255
         Left            =   8340
         TabIndex        =   103
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label50 
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   150
         TabIndex        =   102
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label42 
         Caption         =   "Poblacion:"
         Height          =   255
         Left            =   4620
         TabIndex        =   101
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label37 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   9360
         TabIndex        =   100
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "Departamento:"
         Height          =   255
         Left            =   150
         TabIndex        =   98
         Top             =   200
         Width           =   1710
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   10
         Left            =   6240
         MouseIcon       =   "frmFichaPersonal.frx":0F1E
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":1070
         ToolTipText     =   "Buscar rama genérica"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Rama Genérica:"
         Height          =   255
         Left            =   6585
         TabIndex        =   97
         Top             =   200
         Width           =   1155
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   11
         Left            =   9945
         MouseIcon       =   "frmFichaPersonal.frx":1172
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":12C4
         ToolTipText     =   "Buscar rama específica"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Rama Específica:"
         Height          =   255
         Left            =   10260
         TabIndex        =   96
         Top             =   200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "EMPRESA"
      Height          =   3945
      Left            =   7560
      TabIndex        =   80
      Top             =   1890
      Width           =   7455
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   32
         Left            =   1500
         MaxLength       =   40
         TabIndex        =   35
         Tag             =   "Empresas|Fecha de Alta|N|"
         Text            =   "Text1"
         Top             =   3180
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Index           =   33
         Left            =   6150
         MaxLength       =   40
         TabIndex        =   36
         Tag             =   "empresas|f_baja|S|"
         Text            =   "Text1"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   31
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   34
         Tag             =   "empresas|mail_internet|S|"
         Text            =   "Text1"
         Top             =   2790
         Width           =   4035
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   30
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "empresas|fax|S|"
         Text            =   "Text1"
         Top             =   2790
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   150
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "empresas|tel_contacto|S|"
         Text            =   "Text1"
         Top             =   2790
         Width           =   1440
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   4710
         MaxLength       =   30
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2205
         Width           =   2580
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   150
         MaxLength       =   30
         TabIndex        =   29
         Tag             =   "empresas|poblacion|S|"
         Text            =   "Text1"
         Top             =   2175
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   30
         Tag             =   "Empresas|Código Postal|N|"
         Text            =   "Text1"
         Top             =   2190
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   150
         MaxLength       =   40
         TabIndex        =   28
         Tag             =   "empresas|direccion|S|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   7185
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   3990
         MaxLength       =   5
         TabIndex        =   31
         Tag             =   "empresas|distrito|S|"
         Text            =   "Text1"
         Top             =   2190
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   150
         MaxLength       =   40
         TabIndex        =   26
         Tag             =   "Empresas|Nombre Comercial|N|"
         Text            =   "Text1"
         Top             =   450
         Width           =   7140
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   150
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "empresas|cif_nif|S|"
         Text            =   "Text1"
         Top             =   1020
         Width           =   1425
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha Alta"
         Height          =   255
         Left            =   180
         TabIndex        =   92
         Top             =   3210
         Width           =   810
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Baja"
         Height          =   255
         Left            =   4890
         TabIndex        =   91
         Top             =   3210
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Imgppal 
         Enabled         =   0   'False
         Height          =   240
         Index           =   9
         Left            =   5895
         MouseIcon       =   "frmFichaPersonal.frx":13C6
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":1518
         ToolTipText     =   "Seleccionar fecha"
         Top             =   3195
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   8
         Left            =   1245
         MouseIcon       =   "frmFichaPersonal.frx":15A3
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":16F5
         ToolTipText     =   "Seleccionar fecha"
         Top             =   3195
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "E_mail:"
         Height          =   255
         Left            =   3240
         TabIndex        =   90
         Top             =   2550
         Width           =   570
      End
      Begin VB.Label Label8 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   1665
         TabIndex        =   89
         Top             =   2550
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   150
         TabIndex        =   88
         Top             =   2550
         Width           =   930
      End
      Begin VB.Label Label16 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   4725
         TabIndex        =   87
         Top             =   1950
         Width           =   930
      End
      Begin VB.Label Label15 
         Caption         =   "Poblacion:"
         Height          =   255
         Left            =   150
         TabIndex        =   86
         Top             =   1950
         Width           =   930
      End
      Begin VB.Label Label13 
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   150
         TabIndex        =   85
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "C.Postal:"
         Height          =   255
         Left            =   3915
         TabIndex        =   84
         Top             =   1950
         Width           =   675
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   7
         Left            =   3615
         MouseIcon       =   "frmFichaPersonal.frx":1780
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":18D2
         ToolTipText     =   "Buscar código postal"
         Top             =   1935
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre: "
         Height          =   195
         Left            =   150
         TabIndex        =   82
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "CIF o DNI:"
         Height          =   255
         Left            =   150
         TabIndex        =   81
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "USUARIOS"
      Height          =   3945
      Left            =   60
      TabIndex        =   57
      Top             =   1890
      Width           =   7485
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   20
         Left            =   120
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "Operarios|Cod.Rama Generica|N|"
         Text            =   "Text1"
         Top             =   3180
         Width           =   510
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   660
         MaxLength       =   40
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3180
         Width           =   3165
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4530
         MaxLength       =   40
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3180
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   25
         Tag             =   "operarios|profesion_catego|S|"
         Text            =   "Text1"
         Top             =   3570
         Width           =   5670
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   21
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Operarios|Cód.Tipo Trabajo|N|"
         Text            =   "Text1"
         Top             =   3180
         Width           =   540
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5310
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   2580
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   18
         Left            =   3690
         MaxLength       =   20
         TabIndex        =   21
         Tag             =   "operarios|n_seg_social|S|"
         Text            =   "Text1"
         Top             =   2580
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   16
         Left            =   150
         MaxLength       =   40
         TabIndex        =   19
         Tag             =   "operarios|f_nacimiento|S|"
         Text            =   "Text1"
         Top             =   2580
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmFichaPersonal.frx":19D4
         Left            =   2010
         List            =   "frmFichaPersonal.frx":19D6
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   2565
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   4110
         MaxLength       =   5
         TabIndex        =   18
         Tag             =   "Operarios|distrito|S|"
         Text            =   "Text1"
         Top             =   1950
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   960
         MaxLength       =   40
         TabIndex        =   15
         Tag             =   "operarios|direccion|S|"
         Text            =   "Text1"
         Top             =   1410
         Width           =   6390
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   3690
         MaxLength       =   5
         TabIndex        =   17
         Tag             =   "Operarios|Código postal|N|"
         Text            =   "Text1"
         Top             =   1950
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   150
         MaxLength       =   30
         TabIndex        =   16
         Tag             =   "operarios|poblacion|S|"
         Text            =   "Text1"
         Top             =   1950
         Width           =   3420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   4920
         MaxLength       =   30
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1950
         Width           =   2430
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   4650
         MaxLength       =   20
         TabIndex        =   14
         Tag             =   "operarios|Nombre|N|"
         Text            =   "Text1"
         Top             =   1020
         Width           =   2670
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   2490
         MaxLength       =   20
         TabIndex        =   13
         Tag             =   "Operarios|Apellido 2|N|"
         Text            =   "Text1"
         Top             =   1020
         Width           =   2085
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   150
         MaxLength       =   20
         TabIndex        =   12
         Tag             =   "Operarios|Apellido 1|N|"
         Text            =   "Text1"
         Top             =   1020
         Width           =   2265
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   45
         Left            =   150
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "Operarios|Dni|N|"
         Text            =   "Text1"
         Top             =   390
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   10
         Tag             =   "operarios|n_carnet_radiolog|S|"
         Text            =   "Text1"
         Top             =   450
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   5880
         MaxLength       =   40
         TabIndex        =   11
         Tag             =   "operarios|f_emi_carnet_rad|S|"
         Text            =   "Text1"
         Top             =   450
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   2010
         MaxLength       =   40
         TabIndex        =   73
         Tag             =   "Operarios|Sexo|N|"
         Text            =   "Text1"
         Top             =   2580
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   5340
         MaxLength       =   10
         TabIndex        =   74
         Tag             =   "Operarios|plantilla/contrata|N|"
         Text            =   "Text1"
         Top             =   2610
         Width           =   1425
      End
      Begin VB.Label Label49 
         Caption         =   "Rama Generica:"
         Height          =   255
         Left            =   495
         TabIndex        =   79
         Top             =   2940
         Width           =   1215
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   5
         Left            =   120
         MouseIcon       =   "frmFichaPersonal.frx":19D8
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":1B2A
         ToolTipText     =   "Buscar rama genérica"
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   6
         Left            =   3930
         MouseIcon       =   "frmFichaPersonal.frx":1C2C
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":1D7E
         ToolTipText     =   "Buscar tipo de trabajo"
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label Label48 
         Caption         =   "Profesión Categoria:"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   3600
         Width           =   1560
      End
      Begin VB.Label Label47 
         Caption         =   "Tipo de Trabajo:"
         Height          =   255
         Left            =   4300
         TabIndex        =   77
         Top             =   2940
         Width           =   1290
      End
      Begin VB.Label Label46 
         Caption         =   "Plantilla/Contrata:"
         Height          =   255
         Left            =   5340
         TabIndex        =   72
         Top             =   2340
         Width           =   1905
      End
      Begin VB.Label Label45 
         Caption         =   "N.Seguridad Social:"
         Height          =   195
         Left            =   3690
         TabIndex        =   71
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label33 
         Caption         =   "Sexo :"
         Height          =   195
         Left            =   2040
         TabIndex        =   70
         Top             =   2340
         Width           =   525
      End
      Begin VB.Label Label32 
         Caption         =   "Fec.Nacimiento"
         Height          =   255
         Left            =   465
         TabIndex        =   69
         Top             =   2340
         Width           =   1200
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   4
         Left            =   150
         MouseIcon       =   "frmFichaPersonal.frx":1E80
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":1FD2
         ToolTipText     =   "Seleccionar fecha"
         Top             =   2325
         Width           =   240
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   3
         Left            =   3690
         MouseIcon       =   "frmFichaPersonal.frx":205D
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":21AF
         ToolTipText     =   "Buscar código postal"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label41 
         Caption         =   "C.Postal:"
         Height          =   255
         Left            =   3990
         TabIndex        =   68
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label Label40 
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   150
         TabIndex        =   67
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "Poblacion:"
         Height          =   255
         Left            =   150
         TabIndex        =   66
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Label38 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   4920
         TabIndex        =   65
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Label36 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   4650
         TabIndex        =   63
         Top             =   780
         Width           =   1290
      End
      Begin VB.Label Label35 
         Caption         =   "Primer Apellido:"
         Height          =   195
         Left            =   150
         TabIndex        =   62
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label34 
         Caption         =   "Segundo Apellido:"
         Height          =   255
         Left            =   2520
         TabIndex        =   61
         Top             =   780
         Width           =   1365
      End
      Begin VB.Label Label53 
         Caption         =   "DNI:"
         Height          =   255
         Left            =   150
         TabIndex        =   60
         Top             =   180
         Width           =   1485
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   2
         Left            =   5880
         MouseIcon       =   "frmFichaPersonal.frx":22B1
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":2403
         ToolTipText     =   "Seleccionar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label54 
         Caption         =   "Carnet Radiológico:"
         Height          =   255
         Left            =   4320
         TabIndex        =   59
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label55 
         Caption         =   "Fecha Emisión "
         Height          =   255
         Left            =   6255
         TabIndex        =   58
         Top             =   210
         Width           =   1110
      End
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7485
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   615
      Width           =   1305
   End
   Begin VB.Frame Frame5 
      Caption         =   "DOSIMETROS"
      Height          =   825
      Left            =   60
      TabIndex        =   53
      Top             =   1050
      Width           =   14955
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   6225
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Dosimetro Impar|Número de Dosimetro|N|"
         Text            =   "Text1"
         Top             =   330
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   8670
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Dosimetro Impar|Fecha de Alta|S|"
         Text            =   "Text1"
         Top             =   330
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Dosimetros|Nro.Dosimetro Par|N|"
         Text            =   "Text1"
         Top             =   330
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   3750
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Dosimetro Par|Fecha Alta|N|"
         Text            =   "Text1"
         Top             =   330
         Width           =   1000
      End
      Begin VB.Frame FrameMed 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   9840
         TabIndex        =   112
         Top             =   270
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   46
            Left            =   1545
            MaxLength       =   5
            TabIndex        =   8
            Tag             =   "Dosimetros|Tipo Medición|N|"
            Text            =   "Text1"
            Top             =   60
            Width           =   375
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   113
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   60
            Width           =   2895
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo Medición"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   90
            Width           =   1050
         End
         Begin VB.Image Imgppal 
            Height          =   240
            Index           =   13
            Left            =   1275
            MouseIcon       =   "frmFichaPersonal.frx":248E
            MousePointer    =   99  'Custom
            Picture         =   "frmFichaPersonal.frx":25E0
            ToolTipText     =   "Buscar tipo de medición"
            Top             =   90
            Width           =   240
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Fecha Alta"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7560
         TabIndex        =   111
         Top             =   360
         Width           =   885
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   1
         Left            =   8430
         MouseIcon       =   "frmFichaPersonal.frx":26E2
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":2834
         ToolTipText     =   "Seleccionar fecha"
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Dosímetro Impar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4875
         TabIndex        =   110
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Dosímetro Par"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Alta"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2595
         TabIndex        =   54
         Top             =   360
         Width           =   885
      End
      Begin VB.Image Imgppal 
         Height          =   240
         Index           =   0
         Left            =   3510
         MouseIcon       =   "frmFichaPersonal.frx":28BF
         MousePointer    =   99  'Custom
         Picture         =   "frmFichaPersonal.frx":2A11
         ToolTipText     =   "Seleccionar fecha"
         Top             =   330
         Width           =   240
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   4065
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Instalacion|Código|N|"
      Text            =   "Text1"
      Top             =   615
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13740
      TabIndex        =   50
      Top             =   8100
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   0
      Tag             =   "Empresa|Código|N|"
      Text            =   "Text1"
      Top             =   615
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12540
      TabIndex        =   49
      Top             =   8100
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
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
      Left            =   9000
      TabIndex        =   115
      Top             =   645
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de dosimetría"
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
      Left            =   5595
      TabIndex        =   55
      Top             =   645
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
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
      Index           =   1
      Left            =   225
      TabIndex        =   52
      Top             =   645
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Instalación"
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
      Height          =   270
      Index           =   0
      Left            =   2925
      TabIndex        =   51
      Top             =   645
      Width           =   1155
   End
End
Attribute VB_Name = "frmFichaPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public quien As Integer
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmEmp As frmEmpresas
Attribute frmEmp.VB_VarHelpID = -1
Private WithEvents frmPro As frmProvincias
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmRGe As frmRamasGener
Attribute frmRGe.VB_VarHelpID = -1
Private WithEvents frmREs As frmRamasEspe
Attribute frmREs.VB_VarHelpID = -1
Private WithEvents frmTTr As frmTiposTrab
Attribute frmTTr.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmOpeIns As frmOperariosInstala
Attribute frmOpeIns.VB_VarHelpID = -1
Private WithEvents frmTEx As frmTiposExtremidades
Attribute frmTEx.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  estamos introduciendo datos
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

Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
   If Modo = 0 Then
        Select Case KeyCode
            Case vbESC '27
                PonerModo 0
           Case vbAñadir
                Toolbar1_ButtonClick Toolbar1.Buttons(1)
            Case vbSalir
                Toolbar1_ButtonClick Toolbar1.Buttons(3)
        End Select
   End If


End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    If Combo3.ListIndex = 2 Then ' si es area el usuario es automatico
        If Not DatosAutomaticosUsuario Then Exit Sub
    End If
    
    Cad = Text1(46).Tag
    If Not FrameMed.Visible Then
      Text1(46).Tag = ""
    End If
    
    If DatosOk Then
        '-----------------------------------------
        'Hacemos insertar
        If InsertarRegistros Then
            MsgBox "Inserción de registros realizada con éxito.", vbExclamation, "Ficha Personal."
            'Combo3.ListIndex = -1
            FrameMed.Visible = False
            Text1(46).Tag = Cad
            PonerModo 0
        End If
    End If
    
Error1:
    Text1(46).Tag = Cad
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation, "¡Error!"
End Sub

Private Sub cmdCancelar_Click()
    LimpiarCampos
    PonerModo 0
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 1
    
    '###A mano
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
    If Combo1.ListIndex = 0 Then Text1(19).Text = "01"
    If Combo1.ListIndex = 1 Then Text1(19).Text = "02"
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If

End Sub

Private Sub Combo2_LostFocus()
    If Combo2.ListIndex = 0 Then
        Text1(17).Text = "V"
    Else
        Text1(17).Text = "M"
    End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Combo3_LostFocus()
    If Combo3.ListIndex = 2 Then
        ' el dni tiene que ser 77777777
        Text1(45).Text = "777777777"
        Frame18.Enabled = False
        FrameMed.Visible = False
        Text1(46).Text = ""
        Text2(11).Text = ""
        If DatosAutomaticosUsuario Then
        
        End If
    Else
        FrameMed.Visible = Combo3.ListIndex = 1
        If Not FrameMed.Visible Then
          Text1(46).Text = ""
          Text2(11).Text = ""
        End If
        
        Frame18.Enabled = True
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
      
End Sub

Private Sub Form_Load()
Dim I As Integer
    
    Me.Top = 0
    Me.Left = 0

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(3).Image = 15
    End With

    LimpiarCampos
    
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(1).Visible = False
    End If
    
    CargarCombo
    Modo = 0
    PonerModo CInt(Modo)
    
End Sub

Private Sub LimpiarCampos()
    FrameMed.Visible = False
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Limpiar Me
End Sub


Private Sub frmTEx_DatoSeleccionado(CadenaSeleccion As String)
  If CadenaSeleccion <> "" Then
    Text1(46).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(11).Text = RecuperaValor(CadenaSeleccion, 2)
  End If
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Select Case quien
            Case 3
                Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
                Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
            Case 7
                Text1(27).Text = RecuperaValor(CadenaSeleccion, 1)
                Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
            Case 12
                Text1(38).Text = RecuperaValor(CadenaSeleccion, 1)
                Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
        End Select
    
    End If
End Sub

Private Sub frmTTR_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(20).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(5).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(20).Text & "|", "T|", 1)
        Text1(21).Text = RecuperaValor(CadenaSeleccion, 2)
        Text2(4).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub frmrge_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Select Case quien
            Case 5
                Text1(20).Text = RecuperaValor(CadenaSeleccion, 1)
                Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
            Case 10
                Text1(34).Text = RecuperaValor(CadenaSeleccion, 1)
                Text2(6).Text = RecuperaValor(CadenaSeleccion, 2)
        End Select
                
    End If
End Sub

Private Sub frmREs_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(34).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(6).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Trim(Text1(34).Text) & "|", "T|", 1)
        Text1(35).Text = RecuperaValor(CadenaSeleccion, 2)
        Text2(3).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
'    If Modo = 0 Or Modo = 2 Then Exit Sub
    Select Case Index
       Case 0 'fecha de alta dosimetro par
            f = Now
            If Text1(3).Text <> "" Then
                If IsDate(Text1(3).Text) Then f = Text1(3).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(3).Text = frmC.fecha
            Text1(3).Text = Format(Text1(3).Text, "dd/mm/yyyy")
            Set frmC = Nothing
       Case 1 'fecha de alta dosimetro impar
            f = Now
            If Text1(5).Text <> "" Then
                If IsDate(Text1(5).Text) Then f = Text1(5).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(5).Text = frmC.fecha
            Text1(5).Text = Format(Text1(5).Text, "dd/mm/yyyy")
            Set frmC = Nothing
       Case 2 'fecha de emision
            f = Now
            If Text1(7).Text <> "" Then
                If IsDate(Text1(7).Text) Then f = Text1(7).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(7).Text = frmC.fecha
            Text1(7).Text = Format(Text1(7).Text, "dd/mm/yyyy")
            Set frmC = Nothing
       Case 4 'fecha de nacimiento
            f = Now
            If Text1(16).Text <> "" Then
                If IsDate(Text1(16).Text) Then f = Text1(16).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(16).Text = frmC.fecha
            Text1(16).Text = Format(Text1(16).Text, "dd/mm/yyyy")
            Set frmC = Nothing
       Case 8 'fecha de alta de empresa
            f = Now
            If Text1(32).Text <> "" Then
                If IsDate(Text1(32).Text) Then f = Text1(32).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(32).Text = frmC.fecha
            Text1(32).Text = Format(Text1(32).Text, "dd/mm/yyyy")
            Set frmC = Nothing
        
        Case 3, 7, 12 ' codigo de provincia
            quien = Index
            Set frmPro = New frmProvincias
            frmPro.DatosADevolverBusqueda = "0|1|"
            frmPro.Show
        Case 6 ' tipo de trabajo
            Set frmTTr = New frmTiposTrab
            frmTTr.DatosADevolverBusqueda = "0|1|2|"
            frmTTr.Show
        Case 5, 10 ' rama generica
            quien = Index
            Set frmRGe = New frmRamasGener
            frmRGe.DatosADevolverBusqueda = "0|1|"
            frmRGe.Show
        Case 11 ' rama especifica
            Set frmREs = New frmRamasEspe
            frmREs.DatosADevolverBusqueda = "0|1|2|3|4|"
            frmREs.Show
        Case 13 ' tipo medición
            Set frmTEx = New frmTiposExtremidades
            frmTEx.DatosADevolverBusqueda = "0|1|"
            frmTEx.Show
   End Select
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub




'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
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
    Dim ano As Integer
    Dim Mes As Integer
    Dim fec As Date
    
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" And Index <> 46 Then Exit Sub
    
    Select Case Index
        Case 0, 1, 2, 4, 8, 14, 15, 17, 18, 19, 20, 21, 24, 27, 28, 29, 30, 31, 34, 35, 38, 39, 40, 41, 42, 44, 45, 46
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
        
            If Index = 0 Then ' comprobamos que la empresa no se encuentra ya en la BD
                sql = ""
                sql = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
                If sql <> "" Then
                    MsgBox "Esta empresa ya existe. Vaya a los mantenimientos correspondientes.", vbExclamation, "¡Error!"
                    Text1(0).Text = ""
                    PonerFoco Text1(0)
                    Exit Sub
                End If
            End If
            
            If Index = 46 Then
              If Text1(46).Text = "" Then
                Text2(11).Text = ""
              Else
                Text2(11).Text = DevuelveDesdeBD(1, "descripcion", "tipmedext", "c_tipo_med|", Text1(46).Text & "|", "N|", 1)
                If Text2(11).Text = "" Then
                  MsgBox "Tipo de medición no existe. Reintroduzca.", vbExclamation, "¡Error!"
                  PonerFoco Text1(11)
                End If
              End If

            End If
            
            If Index = 15 Or Index = 28 Or Index = 39 Then ' distrito
                If Text1(15).Text = "" Then Text1(15).Text = Text1(Index).Text
                If Text1(28).Text = "" Then Text1(28).Text = Text1(Index).Text
                If Text1(39).Text = "" Then Text1(39).Text = Text1(Index).Text
            End If
            If Index = 24 Or Index = 45 Then ' dnis
                If Text1(24).Text = "" Then Text1(24).Text = Text1(Index).Text
                If Text1(45).Text = "" Then Text1(45).Text = Text1(Index).Text
            End If
            If Index = 29 Or Index = 40 Then ' telefono
                If Text1(29).Text = "" Then Text1(29).Text = Text1(Index).Text
                If Text1(40).Text = "" Then Text1(40).Text = Text1(Index).Text
            End If
            If Index = 30 Or Index = 41 Then ' fax
                If Text1(30).Text = "" Then Text1(30).Text = Text1(Index).Text
                If Text1(41).Text = "" Then Text1(41).Text = Text1(Index).Text
            End If
            If Index = 31 Or Index = 42 Then ' email
                If Text1(31).Text = "" Then Text1(31).Text = Text1(Index).Text
                If Text1(42).Text = "" Then Text1(42).Text = Text1(Index).Text
            End If
            If Index = 14 Or Index = 27 Or Index = 38 Then  ' codigos postales
                If Text1(Index).Text <> "" Then
                    Select Case Index
                        Case 14
                            Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(14).Text & "|", "T|", 1)
                            If Text2(2).Text = "" Then
                                MsgBox "Código de provincia no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                            If Text1(27).Text = "" Then
                                Text1(27).Text = Text1(Index).Text
                                Text2(1).Text = Text2(2).Text
                            End If
                            If Text1(38).Text = "" Then
                                Text1(38).Text = Text1(Index).Text
                                Text2(7).Text = Text2(2).Text
                            End If
                            
                        
                        Case 27
                            Text2(1).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(27).Text & "|", "T|", 1)
                            If Text2(1).Text = "" Then
                                MsgBox "Código de provincia no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                            If Text1(14).Text = "" Then
                                Text1(14).Text = Text1(Index).Text
                                Text2(2).Text = Text2(1).Text
                            End If
                            If Text1(38).Text = "" Then
                                Text1(38).Text = Text1(Index).Text
                                Text2(7).Text = Text2(1).Text
                            End If
                        Case 38
                            Text2(7).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(38).Text & "|", "T|", 1)
                            If Text2(7).Text = "" Then
                                MsgBox "Código de provincia no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                            If Text1(14).Text = "" Then
                                Text1(14).Text = Text1(Index).Text
                                Text2(2).Text = Text2(7).Text
                            End If
                            If Text1(27).Text = "" Then
                                Text1(27).Text = Text1(Index).Text
                                Text2(1).Text = Text2(7).Text
                            End If
                    End Select
                End If
            End If
            
            If Index = 20 Or Index = 34 Then
                If Text1(Index).Text <> "" Then
                    Select Case Index
                        Case 20
                            Text2(5).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(Index).Text & "|", "T|", 1)
                            If Text2(5).Text = "" Then
                                MsgBox "El Código de rama genérica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                            If Text1(34).Text = "" Then
                                Text1(34).Text = Text1(Index).Text
                                Text2(6).Text = Text2(5).Text
                            End If
                            
                        
                        Case 34
                            Text2(6).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(Index).Text & "|", "T|", 1)
                            If Text2(6).Text = "" Then
                                MsgBox "El Código de rama genérica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                            If Text1(20).Text = "" Then
                                Text1(20).Text = Text1(Index).Text
                                Text2(5).Text = Text2(6).Text
                            End If
                    
                    End Select
                End If
            End If
            
            If Index = 35 Then
                If Text1(34).Text <> "" And Text1(35) <> "" Then
                    Text2(3).Text = DevuelveDesdeBD(1, "descripcion", "ramaespe", "cod_rama_gen|c_rama_especifica|", Text1(34).Text & "|" & Text1(35).Text & "|", "T|T|", 2)
                    If Text2(3).Text = "" Then
                        MsgBox "El código de rama específica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        Text1(34).Text = ""
                        PonerFoco Text1(34)
                    End If
                End If
            End If
            
            If Index = 21 Then
                If Text1(20).Text <> "" And Text1(21).Text <> "" Then
                    Text2(4).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "cod_rama_gen|c_tipo_trabajo|", Text1(20).Text & "|" & Text1(21).Text & "|", "T|T|", 2)
                    If Text2(4).Text = "" Then
                        MsgBox "Código de tipo de trabajo no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
        
        Case 12, 25, 36 ' domicilio
            If Text1(12).Text = "" Then Text1(12).Text = Text1(Index).Text
            If Text1(25).Text = "" Then Text1(25).Text = Text1(Index).Text
            If Text1(36).Text = "" Then Text1(36).Text = Text1(Index).Text
                       
        Case 13, 26, 37 ' poblacion
            If Text1(13).Text = "" Then Text1(13).Text = Text1(Index).Text
            If Text1(26).Text = "" Then Text1(26).Text = Text1(Index).Text
            If Text1(37).Text = "" Then Text1(37).Text = Text1(Index).Text
                      
           
        Case 3, 5, 7, 16, 32, 33 ' campos de fechas
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation, "¡Error!"
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
              
              If Index = 3 Then
                    If Text1(5).Text = "" Then
                        Text1(5).Text = Format(DateAdd("M", 1, CDate(Text1(3).Text)), "dd/mm/yyyy")
                        'Mes = Month(CDate(Text1(3).Text)) + 2
                        'If Mes >= 11 Then
                        '    ano = Year(CDate(Text1(3).Text)) + 1
                        '    Mes = Mes - 12
                        'Else
                        '    ano = Year(CDate(Text1(3).Text))
                        'End If
                        'fec = CDate("01-" & Format(Mes, "00") & "-" & Format(ano, "0000"))
                        'Text1(5).Text = Format(fec - 1, "dd/mm/yyyy")
                    End If
              End If
            End If
              
    End Select
    Text1(Index).Text = Format(Text1(Index).Text, ">")
End Sub


Private Function DatosOk() As Boolean
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim AdmiteNulos As String
Dim Campo As String
Dim tabla As String

    DatosOk = True
    
    If Combo3.ListIndex = -1 Then
      MsgBox "Ha de seleccionar un tipo de dosimetría válido.", vbExclamation, "¡Error!"
      Combo3.SetFocus
      DatosOk = False
      Exit Function
    End If
    
    If Combo4.ListIndex = -1 Then
      MsgBox "Ha de seleccionar un sistema válido.", vbExclamation, "¡Error!"
      Combo4.SetFocus
      DatosOk = False
      Exit Function
    End If
    
    For I = 0 To Text1.Count - 1
        If Text1(I).Text = "" Then
            AdmiteNulos = RecuperaValor(Text1(I).Tag, 3)
            If AdmiteNulos = "N" Then
                Campo = RecuperaValor(Text1(I).Tag, 2)
                tabla = RecuperaValor(Text1(I).Tag, 1)
                MsgBox "El valor de " & Campo & " de " & tabla & " no debe estar vacío.", vbExclamation, "¡Error!"
                PonerFoco Text1(I)
                DatosOk = False
                Exit Function
            End If
        End If
    Next I
    
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonAnyadir
        Case 3
            mnSalir_Click
        
        Case Else
    
    End Select
End Sub


Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ValoresPorDefecto()
'    Text1(3).Text = "46"
'    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
'    Text1(10).Text = "mail"
    Text1(32).Text = Format(Now, "dd/mm/yyyy")
    
End Sub


Private Sub PonerModo(Modo As Integer)
Dim b As Boolean
Dim I As Integer

    b = (Modo = 1)
        
    For I = 0 To Text1.Count - 1
        Text1(I).Enabled = b
    Next I
    Text1(33).Enabled = False ' no quiero mostrar la fecha de baja de empresa
    Combo1.Enabled = b
    Combo2.Enabled = b
    Combo3.Enabled = b
    Combo4.Enabled = b
    
    cmdAceptar.Enabled = b
    cmdAceptar.Visible = b
    cmdCancelar.Enabled = b
    cmdCancelar.Visible = b
    For I = 0 To Imgppal.Count - 1
        Imgppal(I).Enabled = b
    Next I
    Imgppal(9).Enabled = False ' fecha de baja siempre inactiva
    If b Then ValoresPorDefecto
    
    If Modo = 1 Then
         PonerFoco Text1(0)
    End If
    Toolbar1.Buttons(1).Enabled = (Modo = 0)
    Toolbar1.Buttons(3).Enabled = (Modo = 0)
    

End Sub

Private Sub CargarCombo()
'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Si, 1-No
    Combo2.Clear
    Combo2.AddItem "Varón"
    Combo2.ItemData(Combo2.NewIndex) = 0

    Combo2.AddItem "Mujer"
    Combo2.ItemData(Combo2.NewIndex) = 1

    
    Combo1.Clear
    Combo1.AddItem "Plantilla"
    Combo1.ItemData(Combo1.NewIndex) = 0

    Combo1.AddItem "Contrata"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
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

End Sub


Private Function InsertarRegistros() As Boolean
Dim sql As String
Dim sql1 As String
Dim NF As Currency

    On Error GoTo eInsertarRegistros

    Conn.BeginTrans

    InsertarRegistros = False
    ' TENEMOS TODOS LOS DATOS QUE SON NECESARIOS PARA HACER LAS INSERCIONES
    
    sql1 = ""
    sql1 = DevuelveDesdeBD(1, "c_empresa", "empresas", "c_empresa|", Trim(Text1(0).Text) & "|", "T|", 1)
    If sql1 = "" Then
        ' EMPRESA
        sql = "insert into empresas (c_empresa, f_alta, f_baja, cif_nif, "
        sql = sql & "nom_comercial, direccion, poblacion, c_postal, distrito, "
        sql = sql & "tel_contacto, fax, pers_contacto, migrado, mail_internet, c_tipo) "
        sql = sql & "VALUES ('"
        sql = sql & Trim(Text1(0).Text) & "',"
        sql = sql & "'" & Format(Text1(32).Text, FormatoFecha) & "',null," ' fecha de alta, baja null
        If Text1(24).Text <> "" Then
            sql = sql & "'" & Trim(Text1(24).Text) & "'," ' dni
        Else
            sql = sql & "null,"
        End If
        
        sql = sql & "'" & Trim(Text1(23).Text) & "'," ' nombre comercial
        
        If Text1(25).Text <> "" Then
            sql = sql & "'" & Trim(Text1(25).Text) & "'," ' direccion
        Else
            sql = sql & "null,"
        End If
        
        If Text1(26).Text <> "" Then
            sql = sql & "'" & Trim(Text1(26).Text) & "'," ' poblacion
        Else
            sql = sql & "null,"
        End If
        
        sql = sql & "'" & Trim(Text1(27).Text) & "'," ' codigo postal
        
        If Text1(28).Text <> "" Then
            sql = sql & "'" & Trim(Text1(28).Text) & "'," ' distrito
        Else
            sql = sql & "null,"
        End If
        
        If Text1(29).Text <> "" Then
            sql = sql & "'" & Trim(Text1(29).Text) & "',"  ' telefono de contacto
        Else
            sql = sql & "null,"
        End If
        
        If Text1(30).Text <> "" Then
            sql = sql & "'" & Trim(Text1(30).Text) & "'," ' fax
        Else
            sql = sql & "null,"
        End If
        
        If Text1(43).Text <> "" Then
            sql = sql & "'" & Trim(Text1(43).Text) & "',null," ' persona de contacto
        Else
            sql = sql & "null,null,"
        End If
        
        If Text1(31).Text <> "" Then
            sql = sql & "'" & Trim(Text1(31).Text) & "'," ' mail
        Else
            sql = sql & "null,"
        End If
        
        ' tipo de dosimetria
        If Combo3.ListIndex = 2 Then
            sql = sql & "1)"    'si el tipo es area la dosimetria es de area
        Else
            sql = sql & "0)"    'si el tipo es cuerpo u organo la dosimetria es personal
        End If
                
        Conn.Execute sql
        
    Else
        ' existe la empresa
        MsgBox "Esta empresa existe en la Base de Datos. Vaya a los mantenimientos correspondientes.", vbExclamation, "¡Error!"
        Conn.RollbackTrans
        Exit Function
    End If
    
    ' INSTALACION
    sql1 = ""
    sql1 = DevuelveDesdeBD(1, "c_instalacion", "instalaciones", "c_empresa|c_instalacion|", Trim(Text1(0).Text) & "|" & Trim(Text1(1).Text) & "|", "T|T|", 2)
    If sql1 = "" Then
        sql = "insert into instalaciones (c_empresa, c_instalacion, "
        sql = sql & "f_alta,  f_baja,  descripcion ,  direccion, "
        sql = sql & "poblacion, c_postal, distrito , telefono, "
        sql = sql & "fax, persona_contacto, migrado, rama_gen, "
        sql = sql & "rama_especifica, mail_internet, observaciones, c_tipo) "
        sql = sql & "VALUES ("
        sql = sql & "'" & Trim(Text1(0).Text) & "',"
        sql = sql & "'" & Trim(Text1(1).Text) & "',"
        sql = sql & "'" & Format(Text1(32).Text, FormatoFecha) & "',null,"
        sql = sql & "'" & Trim(Text1(6).Text) & "'," 'descripcion de la instalacion/departamento
        If Text1(36).Text <> "" Then 'direccion
            sql = sql & "'" & Trim(Text1(36).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        If Text1(37).Text <> "" Then 'poblacion
            sql = sql & "'" & Trim(Text1(37).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        sql = sql & "'" & Trim(Text1(38).Text) & "'," 'codigo postal
        If Text1(39).Text <> "" Then  ' distrito
            sql = sql & "'" & Trim(Text1(39).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        If Text1(40).Text <> "" Then    'telefono
            sql = sql & "'" & Trim(Text1(40).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        If Text1(41).Text <> "" Then  'fax
            sql = sql & "'" & Trim(Text1(41).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        If Text1(43).Text <> "" Then 'persona de contacto
            sql = sql & "'" & Trim(Text1(43).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        sql = sql & "null," 'migrado
        'rama generica
        sql = sql & "'" & Trim(Text1(34).Text) & "',"
        'rama especifica
        sql = sql & "'" & Trim(Text1(35).Text) & "',"
        If Text1(42).Text <> "" Then ' mail internet
            sql = sql & "'" & Trim(Text1(42).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        If Text1(44).Text <> "" Then ' observaciones
            sql = sql & "'" & Trim(Text1(44).Text) & "',"
        Else
            sql = sql & "null,"
        End If
            
        ' tipo de dosimetria
        If Combo3.ListIndex = 2 Then
            sql = sql & "1)"    'si el tipo es area la dosimetria es de area
        Else
            sql = sql & "0)"    'si el tipo es cuerpo u organo la dosimetria es personal
        End If
            
        Conn.Execute sql
    End If
        
    ' OPERARIO
    sql1 = ""
    sql1 = DevuelveDesdeBD(1, "dni", "operarios", "dni|", Text1(45).Text & "|", "T|", 1)
    If sql1 = "" Then
        sql = "insert into operarios (dni, n_seg_social, n_carnet_radiolog, "
        sql = sql & "f_emi_carnet_rad, apellido_1, apellido_2, "
        sql = sql & "nombre, direccion, poblacion, c_postal,"
        sql = sql & "distrito, c_tipo_trabajo, f_nacimiento,"
        sql = sql & "profesion_catego, sexo, plantilla_contrata,"
        sql = sql & "f_alta, f_baja, migrado, cod_rama_gen)  "
        sql = sql & "VALUES ("
        
        sql = sql & "'" & Trim(Text1(45).Text) & "'," 'dni
        If Text1(18).Text <> "" Then   'numero de ss
            sql = sql & "'" & Trim(Text1(18).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        'num carnet radiolog
        If Text1(8).Text <> "" Then
            sql = sql & "'" & Trim(Text1(8).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        'fecha emision de carnet
        If Text1(7).Text <> "" Then
            sql = sql & "'" & Format(Text1(7).Text, FormatoFecha) & "',"
        Else
            sql = sql & "null,"
        End If
        'apellido 1
        sql = sql & "'" & Trim(Text1(9).Text) & "',"
        'apellido 2
        sql = sql & "'" & Trim(Text1(10).Text) & "',"
        'nombre
        sql = sql & "'" & Trim(Text1(11).Text) & "',"
        'direccion
        If Text1(12).Text <> "" Then
            sql = sql & "'" & Trim(Text1(12).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        'poblacion
        If Text1(13).Text <> "" Then
            sql = sql & "'" & Trim(Text1(13).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        'cpostal
        sql = sql & "'" & Trim(Text1(14).Text) & "',"
        'distrito
        If Text1(15).Text <> "" Then
            sql = sql & "'" & Trim(Text1(15).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        ' tipo de trabajo
        sql = sql & "'" & Trim(Text1(21).Text) & "',"
        ' fecha de nacimiento
        sql = sql & "'" & Format(Text1(16).Text, FormatoFecha) & "',"
        'profesion categoria
        If Text1(22).Text <> "" Then
            sql = sql & "'" & Trim(Text1(22).Text) & "',"
        Else
            sql = sql & "null,"
        End If
        'sexo
        sql = sql & "'" & Trim(Text1(17).Text) & "',"
        'plantilla/contrata
        sql = sql & "'" & Trim(Text1(19).Text) & "',"
        'fecha alta , fecha baja, migrado
        sql = sql & "'" & Format(Text1(32).Text, FormatoFecha) & "',null,null,"
        'rama generica
        sql = sql & "'" & Trim(Text1(20).Text) & "')"
        
        Conn.Execute sql
    Else
        MsgBox "Existe el operario vaya al mantenimiento correspondiente", vbExclamation, "¡Error!"
        Conn.RollbackTrans
        Exit Function
        
    End If
        
    ' RELACION INSTALACION/OPERARIO
    sql1 = ""
    sql1 = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|f_alta|", Trim(Text1(0).Text) & "|" & Trim(Text1(1).Text) & "|" & Trim(Text1(45).Text) & "|" & Text1(32).Text & "|", "T|T|T|F|", 4)
    If sql1 = "" Then
        sql = "insert into operainstala (c_empresa, c_instalacion, "
        sql = sql & "dni, f_alta, f_baja, migrado) VALUES ("
        sql = sql & "'" & Trim(Text1(0).Text) & "',"
        sql = sql & "'" & Trim(Text1(1).Text) & "',"
        sql = sql & "'" & Trim(Text1(45).Text) & "',"
        sql = sql & "'" & Format(Text1(32).Text, FormatoFecha) & "',null,null)"
        
        Conn.Execute sql
    End If
    
    ' DOSIMETRO MES PAR
    'comprobamos que el dosimetro no este asignado
    If Combo3.ListIndex <> -1 Then
        If TieneFechaRetirada(Text1(2).Text, Combo3.ListIndex) Then
        'personal
'        If Combo3.ListIndex = 0 Or Combo3.ListIndex = 1 Then
            sql = "insert into dosimetros (n_reg_dosimetro, n_dosimetro,"
            sql = sql & "c_empresa, c_instalacion, dni_usuario, c_tipo_trabajo,"
            sql = sql & "plantilla_contrata, f_asig_dosimetro, f_retirada,"
            sql = sql & "mes_p_i, tipo_dosimetro, observaciones, sistema, tipo_medicion) VALUES ("
                
            NF = SugerirCodigoSiguiente(0)
            sql = sql & ImporteSinFormato(CStr(NF)) & "," ' numero de registro
            sql = sql & "'" & Trim(Text1(2).Text) & "'," 'n_dosimetro
            sql = sql & "'" & Trim(Text1(0).Text) & "'," 'empresa
            sql = sql & "'" & Trim(Text1(1).Text) & "'," 'instalacion
            sql = sql & "'" & Trim(Text1(45).Text) & "'," 'dni
            sql = sql & "'" & Trim(Text1(21).Text) & "'," 'tipo de trabajo
            sql = sql & "'" & Trim(Text1(19).Text) & "'," 'plantilla contrata
            sql = sql & "'" & Format(Text1(3).Text, FormatoFecha) & "'," 'fecha de asignacion
            sql = sql & "null,'P'," & Format(Combo3.ListIndex, "0") & ",null, '"
            sql = sql & IIf(Combo4.ListIndex = 0, "H", "P") & "', "
            sql = sql & IIf(Combo3.ListIndex = 1, "'" & Text1(46).Text & "'", "null") & ")"
            Conn.Execute sql
'        Else 'area
'            sql = "insert into dosimarea (n_reg_dosimetro, n_dosimetro,"
'            sql = sql & "c_empresa, c_instalacion, dni_usuario, c_tipo_trabajo,"
'            sql = sql & "plantilla_contrata, f_asig_dosimetro, f_retirada,"
'            sql = sql & "mes_p_i, observaciones) VALUES ("
'
'            NF = SugerirCodigoSiguiente(1)
'            sql = sql & ImporteSinFormato(CStr(NF)) & "," ' numero de registro
'            sql = sql & "'" & Trim(Text1(2).Text) & "'," 'n_dosimetro
'            sql = sql & "'" & Trim(Text1(0).Text) & "'," 'empresa
'            sql = sql & "'" & Trim(Text1(1).Text) & "'," 'instalacion
'            sql = sql & "'" & Trim(Text1(45).Text) & "'," 'dni
'            sql = sql & "'" & Trim(Text1(21).Text) & "'," 'tipo de trabajo
'            sql = sql & "'" & Trim(Text1(19).Text) & "'," 'plantilla contrata
'            sql = sql & "'" & Format(Text1(3).Text, FormatoFecha) & "'," 'fecha de asignacion
'            sql = sql & "null,'P',null)"
'
'            Conn.Execute sql
        Else
            MsgBox "El dosimetro del mes par esta asignado. Revise.", vbExclamation, "¡Error!"
        
        End If
    End If
    
    ' DOSIMETRO MES IMPAR
    If Combo3.ListIndex <> -1 Then
         If TieneFechaRetirada(Text1(4).Text, Combo3.ListIndex) Then
        'personal
'        If Combo3.ListIndex = 0 Or Combo3.ListIndex = 1 Then
            sql = "insert into dosimetros (n_reg_dosimetro, n_dosimetro,"
            sql = sql & "c_empresa, c_instalacion, dni_usuario, c_tipo_trabajo,"
            sql = sql & "plantilla_contrata, f_asig_dosimetro, f_retirada,"
            sql = sql & "mes_p_i, tipo_dosimetro, observaciones, sistema, tipo_medicion) VALUES ("
    
            NF = SugerirCodigoSiguiente(0)
            sql = sql & ImporteSinFormato(CStr(NF)) & "," ' numero de registro
            sql = sql & "'" & Trim(Text1(4).Text) & "'," 'n_dosimetro
            sql = sql & "'" & Trim(Text1(0).Text) & "'," 'empresa
            sql = sql & "'" & Trim(Text1(1).Text) & "'," 'instalacion
            sql = sql & "'" & Trim(Text1(45).Text) & "'," 'dni
            sql = sql & "'" & Trim(Text1(21).Text) & "'," 'tipo de trabajo
            sql = sql & "'" & Trim(Text1(19).Text) & "'," 'plantilla contrata
            sql = sql & "'" & Format(Text1(5).Text, FormatoFecha) & "'," 'fecha de asignacion
            sql = sql & "null,'I'," & Format(Combo3.ListIndex, "0") & ",null, '"
            sql = sql & IIf(Combo4.ListIndex = 0, "H", "P") & "', "
            sql = sql & IIf(Combo3.ListIndex = 1, "'" & Text1(46).Text & "'", "null") & ")"
            Conn.Execute sql
'        Else
'            sql = "insert into dosimarea (n_reg_dosimetro, n_dosimetro,"
'            sql = sql & "c_empresa, c_instalacion, dni_usuario, c_tipo_trabajo,"
'            sql = sql & "plantilla_contrata, f_asig_dosimetro, f_retirada,"
'            sql = sql & "mes_p_i, observaciones) VALUES ("
'
'            NF = SugerirCodigoSiguiente(1)
'            sql = sql & ImporteSinFormato(CStr(NF)) & "," ' numero de registro
'            sql = sql & "'" & Trim(Text1(4).Text) & "'," 'n_dosimetro
'            sql = sql & "'" & Trim(Text1(0).Text) & "'," 'empresa
'            sql = sql & "'" & Trim(Text1(1).Text) & "'," 'instalacion
'            sql = sql & "'" & Trim(Text1(45).Text) & "'," 'dni
'            sql = sql & "'" & Trim(Text1(21).Text) & "'," 'tipo de trabajo
'            sql = sql & "'" & Trim(Text1(19).Text) & "'," 'plantilla contrata
'            sql = sql & "'" & Format(Text1(5).Text, FormatoFecha) & "'," 'fecha de asignacion
'            sql = sql & "null,'I',null)"
'
'            Conn.Execute sql
'
        Else
            MsgBox "El dosimetro del mes impar esta asignado. Revise.", vbExclamation, "¡Error!"
            
        End If
    End If

eInsertarRegistros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la insercion de registros"
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        InsertarRegistros = True
    End If
End Function



Private Function SugerirCodigoSiguiente(Tipo As Byte) As String
Dim sql As String
Dim Rs As ADODB.Recordset
    
'    If tipo = 0 Then
        sql = "Select Max(n_reg_dosimetro) from dosimetros where tipo_dosimetro = " & Tipo
'    Else
'        sql = "Select Max(n_reg_dosimetro) from dosimarea "
'    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    sql = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            sql = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = sql
End Function


Private Function DatosAutomaticosUsuario() As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset
    
    Text1(45).Text = "777777777"
    DatosAutomaticosUsuario = True
    
    sql = "Select * from operarios where dni = '" & Trim(Text1(45).Text) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    If Not Rs.EOF Then
        Text1(7).Text = Rs!f_emi_carnet_rad & ""
        Text1(8).Text = Rs!n_carnet_radiolog & ""
        Text1(9).Text = Rs!apellido_1 & ""
        Text1(10).Text = Rs!apellido_2 & ""
        Text1(11).Text = Rs!nombre & ""
        Text1(12).Text = Rs!direccion & ""
        Text1(13).Text = Rs!poblacion & ""
        Text1(14).Text = Rs!c_postal & ""
        Text1(15).Text = Rs!distrito & ""
        Text1(16).Text = Rs!f_nacimiento & ""
        Text1(17).Text = Rs!sexo & ""
        Text1(18).Text = Rs!n_seg_social & ""
        Text1(19).Text = Rs!plantilla_contrata & ""
        Text1(20).Text = Rs!cod_rama_gen & ""
        Text1(21).Text = Rs!c_tipo_trabajo & ""
        Text1(22).Text = Rs!profesion_catego & ""
        'rama generica
        Text2(5).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(20).Text & "|", "T|", 1)
        'tipo de trabajo
        Text2(4).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "cod_rama_gen|c_tipo_trabajo|", Text1(20).Text & "|" & Text1(21).Text & "|", "T|T|", 2)
        'provincia
        Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(14).Text & "|", "T|", 1)
        'sexo
        Combo2.ListIndex = 0
        If (Rs!sexo & "") = "M" Then
           Combo2.ListIndex = 1
        End If
        
        'plantilla/contrata
        Combo1.ListIndex = 0
        If CInt(Rs!plantilla_contrata & "") = 2 Then
            Combo1.ListIndex = 1
        End If
    Else
        MsgBox "Debe introducir un usuario ficticio de dni = 777777777", vbExclamation, "¡Error!"
        DatosAutomaticosUsuario = False
    End If
    Rs.Close

End Function
