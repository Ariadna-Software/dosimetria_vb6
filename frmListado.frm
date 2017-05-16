VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   7350
   Begin VB.Frame FrameListDosisInstal 
      Height          =   5250
      Left            =   30
      TabIndex        =   133
      Top             =   30
      Width           =   5940
      Begin VB.Frame Frame7 
         Caption         =   "Tipo de dosimetría"
         ForeColor       =   &H8000000D&
         Height          =   585
         Left            =   240
         TabIndex        =   523
         Top             =   4575
         Visible         =   0   'False
         Width           =   2415
         Begin VB.OptionButton OptSist 
            Caption         =   "Harshaw"
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   155
            Top             =   285
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton OptSist 
            Caption         =   "Panasonic"
            Height          =   255
            Index           =   1
            Left            =   1230
            TabIndex        =   156
            Top             =   300
            Width           =   1050
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir sólo el fichero migrado"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   158
         Top             =   4215
         Width           =   2475
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Guardar el listado en formato .pdf  "
         Height          =   255
         Left            =   360
         TabIndex        =   157
         Top             =   4215
         Width           =   2700
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   9
         Left            =   3750
         TabIndex        =   147
         Top             =   945
         Width           =   1020
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   1620
         TabIndex        =   146
         Top             =   945
         Width           =   1095
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   3060
         TabIndex        =   138
         Text            =   "Text5"
         Top             =   1980
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   3060
         TabIndex        =   137
         Text            =   "Text5"
         Top             =   1530
         Width           =   2535
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   9
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   149
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   8
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   148
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton CmdAceptarLisDosisIns 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3540
         TabIndex        =   159
         Top             =   4635
         Width           =   975
      End
      Begin VB.CommandButton CmdCanListDosisIns 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4635
         TabIndex        =   160
         Top             =   4635
         Width           =   975
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3090
         TabIndex        =   136
         Text            =   "Text5"
         Top             =   3060
         Width           =   2535
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3090
         TabIndex        =   135
         Text            =   "Text5"
         Top             =   2610
         Width           =   2535
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   5
         Left            =   1620
         MaxLength       =   11
         TabIndex        =   151
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   4
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   150
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de dosimetría"
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   300
         TabIndex        =   134
         Top             =   3480
         Width           =   5280
         Begin VB.OptionButton OptIns 
            Caption         =   "Por Area"
            Height          =   255
            Index           =   4
            Left            =   3810
            TabIndex        =   154
            Top             =   210
            Width           =   1065
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "No Homogénea"
            Height          =   255
            Index           =   3
            Left            =   1860
            TabIndex        =   153
            Top             =   225
            Width           =   1485
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Homogénea"
            Height          =   255
            Index           =   2
            Left            =   210
            TabIndex        =   152
            Top             =   225
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin MSComctlLib.ProgressBar Pb5 
         Height          =   300
         Left            =   345
         TabIndex        =   420
         Top             =   4710
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   1000
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         Caption         =   "DIRECTORIO :  /temp"
         Height          =   255
         Index           =   56
         Left            =   630
         TabIndex        =   424
         Top             =   4485
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Dosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   12
         Left            =   270
         TabIndex        =   163
         Top             =   630
         Width           =   990
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   9
         Left            =   3510
         Picture         =   "frmListado.frx":0CCA
         Top             =   945
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   8
         Left            =   1350
         Picture         =   "frmListado.frx":0DCC
         Top             =   945
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   9
         Left            =   3060
         TabIndex        =   162
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   8
         Left            =   735
         TabIndex        =   161
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   145
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Dosis por Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   645
         TabIndex        =   144
         Top             =   225
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   17
         Left            =   720
         TabIndex        =   143
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   16
         Left            =   720
         TabIndex        =   142
         Top             =   1545
         Width           =   615
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   9
         Left            =   1320
         Picture         =   "frmListado.frx":0ECE
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   8
         Left            =   1320
         Picture         =   "frmListado.frx":0FD0
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   141
         Top             =   2340
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   15
         Left            =   750
         TabIndex        =   140
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   14
         Left            =   750
         TabIndex        =   139
         Top             =   2625
         Width           =   615
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   5
         Left            =   1350
         Picture         =   "frmListado.frx":10D2
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   4
         Left            =   1350
         Picture         =   "frmListado.frx":11D4
         Top             =   2580
         Width           =   240
      End
   End
   Begin VB.Frame FrameListDosisOpeAcum12 
      Height          =   5085
      Left            =   0
      TabIndex        =   335
      Top             =   30
      Width           =   5850
      Begin VB.Frame Frame23 
         Caption         =   "Tipo de dosimetría"
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   150
         TabIndex        =   524
         Top             =   4365
         Visible         =   0   'False
         Width           =   2610
         Begin VB.OptionButton OptSist 
            Caption         =   "Panasonic"
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   526
            Top             =   300
            Width           =   1050
         End
         Begin VB.OptionButton OptSist 
            Caption         =   "Harshaw"
            Height          =   255
            Index           =   2
            Left            =   195
            TabIndex        =   525
            Top             =   285
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame Frame15 
         Height          =   720
         Left            =   150
         TabIndex        =   357
         Top             =   3630
         Width           =   5535
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir sólo el fichero migrado"
            Height          =   435
            Index           =   1
            Left            =   2910
            TabIndex        =   352
            Top             =   180
            Width           =   2475
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Solo usuarios sin fecha de baja"
            Height          =   435
            Index           =   0
            Left            =   180
            TabIndex        =   351
            Top             =   180
            Width           =   2565
         End
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   1590
         TabIndex        =   340
         Top             =   1110
         Width           =   1095
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   3060
         TabIndex        =   339
         Text            =   "Text5"
         Top             =   3150
         Width           =   2535
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3060
         TabIndex        =   338
         Text            =   "Text5"
         Top             =   2700
         Width           =   2535
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   7
         Left            =   1620
         MaxLength       =   15
         TabIndex        =   344
         Top             =   3150
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   6
         Left            =   1620
         MaxLength       =   15
         TabIndex        =   343
         Top             =   2700
         Width           =   1335
      End
      Begin VB.CommandButton CmdAceptarListDosisOpeAcum12 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3645
         TabIndex        =   345
         Top             =   4530
         Width           =   975
      End
      Begin VB.CommandButton CmdCanListDosisOpeAcum12 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4740
         TabIndex        =   346
         Top             =   4530
         Width           =   975
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   3060
         TabIndex        =   337
         Text            =   "Text5"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   3060
         TabIndex        =   336
         Text            =   "Text5"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   13
         Left            =   1620
         MaxLength       =   11
         TabIndex        =   342
         Top             =   2130
         Width           =   1305
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   12
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   341
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año Dosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   34
         Left            =   240
         TabIndex        =   356
         Top             =   1110
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dni Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   33
         Left            =   240
         TabIndex        =   355
         Top             =   2430
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Dosis por Operario Año Oficial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   354
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   45
         Left            =   720
         TabIndex        =   353
         Top             =   3150
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   44
         Left            =   720
         TabIndex        =   350
         Top             =   2715
         Width           =   615
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   7
         Left            =   1350
         Picture         =   "frmListado.frx":12D6
         Top             =   3150
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   6
         Left            =   1350
         Picture         =   "frmListado.frx":13D8
         Top             =   2730
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   32
         Left            =   240
         TabIndex        =   349
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   43
         Left            =   720
         TabIndex        =   348
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   42
         Left            =   720
         TabIndex        =   347
         Top             =   1785
         Width           =   615
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   13
         Left            =   1350
         Picture         =   "frmListado.frx":14DA
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   12
         Left            =   1350
         Picture         =   "frmListado.frx":15DC
         Top             =   1740
         Width           =   240
      End
   End
   Begin VB.Frame FrameDosisColectiva 
      Height          =   4530
      Left            =   30
      TabIndex        =   164
      Top             =   30
      Width           =   6135
      Begin VB.Frame Frame17 
         Caption         =   "Tipo de Dosimetria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   705
         Left            =   330
         TabIndex        =   447
         Top             =   750
         Width           =   5310
         Begin VB.OptionButton OptIns 
            Caption         =   "Personal"
            Height          =   255
            Index           =   18
            Left            =   1890
            TabIndex        =   449
            Top             =   270
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Area"
            Height          =   255
            Index           =   17
            Left            =   3780
            TabIndex        =   448
            Top             =   270
            Width           =   1485
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Dosis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   705
         Left            =   330
         TabIndex        =   178
         Top             =   2280
         Width           =   5310
         Begin VB.OptionButton OptIns 
            Caption         =   "Profunda"
            Height          =   255
            Index           =   9
            Left            =   3780
            TabIndex        =   180
            Top             =   270
            Width           =   1485
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Superficial"
            Height          =   255
            Index           =   8
            Left            =   1890
            TabIndex        =   179
            Top             =   270
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   345
         Left            =   180
         TabIndex        =   177
         Top             =   3990
         Visible         =   0   'False
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   720
         Left            =   330
         TabIndex        =   169
         Top             =   1500
         Width           =   5280
         Begin VB.OptionButton OptIns 
            Caption         =   "Mensual"
            Height          =   255
            Index           =   7
            Left            =   210
            TabIndex        =   172
            Top             =   300
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Semestral"
            Height          =   255
            Index           =   6
            Left            =   1920
            TabIndex        =   171
            Top             =   300
            Width           =   1485
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Anual"
            Height          =   255
            Index           =   5
            Left            =   3780
            TabIndex        =   170
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.CommandButton CmdCanLisDosisCol 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   167
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListDosisColec 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   166
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   11
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   168
         Top             =   3405
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000014&
         Height          =   285
         Index           =   10
         Left            =   1650
         TabIndex        =   165
         Top             =   3390
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Dosis Colectiva"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   675
         TabIndex        =   176
         Top             =   225
         Width           =   4695
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   11
         Left            =   765
         TabIndex        =   175
         Top             =   3420
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   10
         Left            =   3090
         TabIndex        =   174
         Top             =   3420
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   10
         Left            =   1380
         Picture         =   "frmListado.frx":16DE
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Dosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   14
         Left            =   450
         TabIndex        =   173
         Top             =   3090
         Width           =   990
      End
   End
   Begin VB.Frame FrameListOperarios 
      Height          =   6810
      Left            =   30
      TabIndex        =   39
      Top             =   30
      Width           =   6540
      Begin VB.Frame Frame8 
         Caption         =   "Situación Actual de Operarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   390
         TabIndex        =   358
         Top             =   5400
         Width           =   2895
         Begin VB.OptionButton OptOpe 
            Caption         =   "Todos"
            Height          =   435
            Index           =   6
            Left            =   1920
            TabIndex        =   361
            Top             =   330
            Width           =   765
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "Alta "
            Height          =   435
            Index           =   4
            Left            =   180
            TabIndex        =   360
            Top             =   330
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "Baja"
            Height          =   435
            Index           =   5
            Left            =   1020
            TabIndex        =   359
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   2
         Left            =   2205
         MaxLength       =   11
         TabIndex        =   31
         Top             =   1965
         Width           =   1335
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   3
         Left            =   2205
         MaxLength       =   11
         TabIndex        =   32
         Top             =   2445
         Width           =   1335
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3600
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   1980
         Width           =   2535
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3600
         TabIndex        =   56
         Text            =   "Text5"
         Top             =   2430
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   915
         Left            =   390
         TabIndex        =   55
         Top             =   4440
         Width           =   2895
         Begin VB.OptionButton OptOpe 
            Caption         =   "Específico"
            Height          =   435
            Index           =   1
            Left            =   1500
            TabIndex        =   62
            Top             =   300
            Width           =   1245
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "General"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   61
            Top             =   390
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3600
         TabIndex        =   50
         Text            =   "Text5"
         Top             =   1455
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3600
         TabIndex        =   49
         Text            =   "Text5"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   5
         Left            =   2205
         MaxLength       =   11
         TabIndex        =   30
         Top             =   1455
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   4
         Left            =   2205
         MaxLength       =   11
         TabIndex        =   29
         Top             =   975
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   2250
         TabIndex        =   35
         Top             =   4065
         Width           =   1020
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   4170
         TabIndex        =   36
         Top             =   4065
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptarLisOpe 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4125
         TabIndex        =   37
         Top             =   6045
         Width           =   975
      End
      Begin VB.CommandButton CmdCanLisOpe 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   38
         Top             =   6045
         Width           =   975
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   3090
         Width           =   2535
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   40
         Text            =   "Text5"
         Top             =   3570
         Width           =   2535
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   0
         Left            =   2235
         MaxLength       =   15
         TabIndex        =   33
         Top             =   3090
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   1
         Left            =   2235
         MaxLength       =   15
         TabIndex        =   34
         Top             =   3570
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   510
         Left            =   390
         TabIndex        =   54
         Top             =   5505
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Frame Frame4 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   915
         Left            =   3660
         TabIndex        =   63
         Top             =   4440
         Width           =   2475
         Begin VB.OptionButton OptOpe 
            Caption         =   "Empresa"
            Height          =   525
            Index           =   3
            Left            =   1380
            TabIndex        =   65
            Top             =   270
            Width           =   915
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "DNI"
            Height          =   375
            Index           =   2
            Left            =   330
            TabIndex        =   64
            Top             =   330
            Value           =   -1  'True
            Width           =   765
         End
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   2
         Left            =   1965
         Picture         =   "frmListado.frx":17E0
         ToolTipText     =   "Buscar instalación"
         Top             =   1965
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   3
         Left            =   1965
         Picture         =   "frmListado.frx":18E2
         ToolTipText     =   "Buscar instalación"
         Top             =   2445
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   5
         Left            =   1365
         TabIndex        =   60
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   4
         Left            =   1365
         TabIndex        =   59
         Top             =   2445
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   570
         TabIndex        =   58
         Top             =   1770
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   13
         Left            =   570
         TabIndex        =   53
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   13
         Left            =   1365
         TabIndex        =   52
         Top             =   1455
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   12
         Left            =   1365
         TabIndex        =   51
         Top             =   1020
         Width           =   525
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   5
         Left            =   1965
         Picture         =   "frmListado.frx":19E4
         ToolTipText     =   "Buscar empresa"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   4
         Left            =   1950
         Picture         =   "frmListado.frx":1AE6
         ToolTipText     =   "Buscar empresa"
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   11
         Left            =   570
         TabIndex        =   48
         Top             =   3780
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Operarios en Instalaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   5
         Left            =   855
         TabIndex        =   47
         Top             =   390
         Width           =   4650
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   2010
         Picture         =   "frmListado.frx":1BE8
         ToolTipText     =   "Seleccionar fecha"
         Top             =   4065
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   3900
         Picture         =   "frmListado.frx":1CEA
         ToolTipText     =   "Seleccionar fecha"
         Top             =   4065
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   7
         Left            =   3390
         TabIndex        =   46
         Top             =   4110
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   4
         Left            =   1425
         TabIndex        =   45
         Top             =   4110
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DNI Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   555
         TabIndex        =   44
         Top             =   2820
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   9
         Left            =   1410
         TabIndex        =   43
         Top             =   3540
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   8
         Left            =   1410
         TabIndex        =   42
         Top             =   3105
         Width           =   495
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   0
         Left            =   1995
         Picture         =   "frmListado.frx":1DEC
         ToolTipText     =   "Buscar D.N.I."
         Top             =   3090
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   1
         Left            =   1995
         Picture         =   "frmListado.frx":1EEE
         ToolTipText     =   "Buscar D.N.I."
         Top             =   3540
         Width           =   240
      End
   End
   Begin VB.Frame FrameListInstalaciones 
      Height          =   6285
      Left            =   45
      TabIndex        =   16
      Top             =   30
      Width           =   6165
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   29
         Left            =   1590
         TabIndex        =   4
         Top             =   3165
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   30
         Left            =   3750
         TabIndex        =   5
         Top             =   3165
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   32
         Left            =   3765
         TabIndex        =   7
         Top             =   3810
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   31
         Left            =   1605
         TabIndex        =   6
         Top             =   3810
         Width           =   1095
      End
      Begin VB.Frame Frame21 
         Caption         =   "Situación Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   210
         TabIndex        =   479
         Top             =   4290
         Width           =   2895
         Begin VB.OptionButton OptIns 
            Caption         =   "Todos"
            Height          =   435
            Index           =   21
            Left            =   1920
            TabIndex        =   10
            Top             =   300
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Alta "
            Height          =   435
            Index           =   19
            Left            =   180
            TabIndex        =   8
            Top             =   270
            Width           =   765
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Baja"
            Height          =   435
            Index           =   20
            Left            =   1050
            TabIndex        =   9
            Top             =   300
            Width           =   765
         End
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmListado.frx":1FF0
         Left            =   1650
         List            =   "frmListado.frx":1FF2
         TabIndex        =   13
         Text            =   "Todas"
         Top             =   5430
         Width           =   1305
      End
      Begin VB.Frame Frame9 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   3180
         TabIndex        =   28
         Top             =   4290
         Width           =   2640
         Begin VB.OptionButton OptIns 
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton OptIns 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   1
            Left            =   1410
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   0
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   2
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   1
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   3
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3030
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3030
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CommandButton cmdCanLisIns 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4875
         TabIndex        =   15
         Top             =   5730
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarLisIns 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   14
         Top             =   5730
         Width           =   975
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   2
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   0
         Top             =   990
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   3
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3030
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3030
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   58
         Left            =   300
         TabIndex        =   517
         Top             =   2925
         Width           =   885
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   31
         Left            =   1350
         Picture         =   "frmListado.frx":1FF4
         Top             =   3165
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   30
         Left            =   3480
         Picture         =   "frmListado.frx":20F6
         Top             =   3165
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   35
         Left            =   2970
         TabIndex        =   516
         Top             =   3210
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   34
         Left            =   750
         TabIndex        =   515
         Top             =   3210
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   33
         Left            =   750
         TabIndex        =   514
         Top             =   3855
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   32
         Left            =   2985
         TabIndex        =   513
         Top             =   3855
         Width           =   495
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   240
         Index           =   29
         Left            =   3495
         Picture         =   "frmListado.frx":21F8
         Top             =   3810
         Width           =   240
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   240
         Index           =   11
         Left            =   1365
         Picture         =   "frmListado.frx":22FA
         Top             =   3810
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Baja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   57
         Left            =   300
         TabIndex        =   512
         Top             =   3555
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Dosimetria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   43
         Left            =   270
         TabIndex        =   423
         Top             =   5460
         Width           =   1335
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmListado.frx":23FC
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmListado.frx":24FE
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   61
         Left            =   750
         TabIndex        =   27
         Top             =   2085
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   60
         Left            =   750
         TabIndex        =   26
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   92
         Left            =   270
         TabIndex        =   25
         Top             =   1800
         Width           =   945
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado.frx":2600
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmListado.frx":2702
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   59
         Left            =   720
         TabIndex        =   22
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   58
         Left            =   720
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Instalaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   22
         Left            =   675
         TabIndex        =   20
         Top             =   225
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   91
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame FrameCartaSobredosis 
      Height          =   5490
      Left            =   30
      TabIndex        =   362
      Top             =   30
      Width           =   6135
      Begin VB.Frame Frame16 
         Height          =   885
         Left            =   120
         TabIndex        =   385
         Top             =   3600
         Width           =   5745
         Begin VB.CheckBox Check1 
            Caption         =   "Solo usuarios sin fecha de baja"
            Height          =   435
            Index           =   2
            Left            =   210
            TabIndex        =   387
            Top             =   240
            Width           =   2895
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir sólo el fichero migrado"
            Height          =   435
            Index           =   3
            Left            =   3090
            TabIndex        =   386
            Top             =   270
            Width           =   2475
         End
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   15
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   366
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   14
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   365
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   3060
         TabIndex        =   374
         Text            =   "Text5"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   3060
         TabIndex        =   373
         Text            =   "Text5"
         Top             =   1770
         Width           =   2535
      End
      Begin VB.CommandButton CmdCanCartaSobredosis 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4845
         TabIndex        =   370
         Top             =   4830
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarCartaSobredosis 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3765
         TabIndex        =   369
         Top             =   4830
         Width           =   975
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   9
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   368
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   8
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   367
         Top             =   2700
         Width           =   1335
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   3060
         TabIndex        =   372
         Text            =   "Text5"
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   3060
         TabIndex        =   371
         Text            =   "Text5"
         Top             =   2700
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   23
         Left            =   3870
         TabIndex        =   364
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   22
         Left            =   1650
         TabIndex        =   363
         Top             =   1185
         Width           =   1020
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   15
         Left            =   1320
         Picture         =   "frmListado.frx":2804
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   14
         Left            =   1320
         Picture         =   "frmListado.frx":2906
         Top             =   1770
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   49
         Left            =   720
         TabIndex        =   384
         Top             =   1785
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   48
         Left            =   720
         TabIndex        =   383
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   37
         Left            =   240
         TabIndex        =   382
         Top             =   1500
         Width           =   735
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   9
         Left            =   1320
         Picture         =   "frmListado.frx":2A08
         Top             =   3150
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   8
         Left            =   1320
         Picture         =   "frmListado.frx":2B0A
         Top             =   2700
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   47
         Left            =   720
         TabIndex        =   381
         Top             =   2715
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   46
         Left            =   720
         TabIndex        =   380
         Top             =   3150
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Carta CSN de Potencial de Sobredosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   379
         Top             =   240
         Width           =   5715
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dni Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   36
         Left            =   240
         TabIndex        =   378
         Top             =   2430
         Width           =   1050
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   25
         Left            =   735
         TabIndex        =   377
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   24
         Left            =   3060
         TabIndex        =   376
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   23
         Left            =   3570
         Picture         =   "frmListado.frx":2C0C
         Top             =   1215
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   22
         Left            =   1380
         Picture         =   "frmListado.frx":2D0E
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Dosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   35
         Left            =   270
         TabIndex        =   375
         Top             =   870
         Width           =   990
      End
   End
   Begin VB.Frame FrameCartaDosimNRec 
      Height          =   4890
      Left            =   30
      TabIndex        =   311
      Top             =   30
      Width           =   6135
      Begin VB.CheckBox chkEmail 
         Caption         =   "Guardar el listado en formato .pdf  "
         Height          =   255
         Left            =   1380
         TabIndex        =   334
         Top             =   3690
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   11
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   317
         Top             =   3210
         Width           =   1335
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   10
         Left            =   1620
         MaxLength       =   11
         TabIndex        =   316
         Top             =   2730
         Width           =   1305
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   3090
         TabIndex        =   323
         Text            =   "Text5"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   3090
         TabIndex        =   322
         Text            =   "Text5"
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton CmdCanCartaDosimNRec 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4635
         TabIndex        =   319
         Top             =   4290
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarCartaDosimNRec 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3540
         TabIndex        =   318
         Top             =   4260
         Width           =   975
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   11
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   315
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   10
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   314
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   3060
         TabIndex        =   321
         Text            =   "Text5"
         Top             =   2130
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   3060
         TabIndex        =   320
         Text            =   "Text5"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   19
         Left            =   3870
         TabIndex        =   313
         Top             =   1125
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   18
         Left            =   1620
         TabIndex        =   312
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "DIRECTORIO :  /temp"
         Height          =   255
         Index           =   57
         Left            =   1620
         TabIndex        =   425
         Top             =   3990
         Width           =   1875
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   11
         Left            =   1350
         Picture         =   "frmListado.frx":2E10
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   10
         Left            =   1350
         Picture         =   "frmListado.frx":2F12
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   41
         Left            =   750
         TabIndex        =   333
         Top             =   2805
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   40
         Left            =   750
         TabIndex        =   332
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   31
         Left            =   270
         TabIndex        =   331
         Top             =   2520
         Width           =   945
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   11
         Left            =   1350
         Picture         =   "frmListado.frx":3014
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   10
         Left            =   1350
         Picture         =   "frmListado.frx":3116
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   39
         Left            =   720
         TabIndex        =   330
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   38
         Left            =   720
         TabIndex        =   329
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cartas Oficiales Dosímetros NO Recepcionados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   13
         Left            =   150
         TabIndex        =   328
         Top             =   210
         Width           =   5865
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   30
         Left            =   240
         TabIndex        =   327
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   21
         Left            =   735
         TabIndex        =   326
         Top             =   1140
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   20
         Left            =   3060
         TabIndex        =   325
         Top             =   1140
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   19
         Left            =   3600
         Picture         =   "frmListado.frx":3218
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   18
         Left            =   1380
         Picture         =   "frmListado.frx":331A
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Dosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   29
         Left            =   270
         TabIndex        =   324
         Top             =   810
         Width           =   990
      End
   End
   Begin VB.Frame FrameListOperariosSobredosis 
      Height          =   4470
      Left            =   30
      TabIndex        =   427
      Top             =   30
      Width           =   6135
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   21
         Left            =   1650
         TabIndex        =   438
         Top             =   1185
         Width           =   1020
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   3060
         TabIndex        =   437
         Text            =   "Text5"
         Top             =   2730
         Width           =   2535
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   3060
         TabIndex        =   436
         Text            =   "Text5"
         Top             =   3150
         Width           =   2535
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   13
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   435
         Top             =   2730
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   12
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   434
         Top             =   3150
         Width           =   1335
      End
      Begin VB.CommandButton CmdAceptarListOperariosSobredosis 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3765
         TabIndex        =   433
         Top             =   3810
         Width           =   975
      End
      Begin VB.CommandButton CmdCanListOperariosSobredosis 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4845
         TabIndex        =   432
         Top             =   3810
         Width           =   975
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   3060
         TabIndex        =   431
         Text            =   "Text5"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   3060
         TabIndex        =   430
         Text            =   "Text5"
         Top             =   1770
         Width           =   2535
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   19
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   429
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   18
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   428
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   47
         Left            =   240
         TabIndex        =   446
         Top             =   1080
         Width           =   540
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   21
         Left            =   1380
         Picture         =   "frmListado.frx":341C
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dni Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   46
         Left            =   240
         TabIndex        =   445
         Top             =   2430
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Operarios con Sobredosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   444
         Top             =   360
         Width           =   5715
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   65
         Left            =   720
         TabIndex        =   443
         Top             =   3150
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   64
         Left            =   720
         TabIndex        =   442
         Top             =   2715
         Width           =   615
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   13
         Left            =   1320
         Picture         =   "frmListado.frx":351E
         Top             =   2730
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   12
         Left            =   1320
         Picture         =   "frmListado.frx":3620
         Top             =   3180
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   45
         Left            =   240
         TabIndex        =   441
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   63
         Left            =   720
         TabIndex        =   440
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   62
         Left            =   720
         TabIndex        =   439
         Top             =   1785
         Width           =   615
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   19
         Left            =   1320
         Picture         =   "frmListado.frx":3722
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   18
         Left            =   1320
         Picture         =   "frmListado.frx":3824
         Top             =   1770
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6900
      Top             =   4035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameListRecepDosimCuerpo 
      Height          =   6300
      Left            =   30
      TabIndex        =   388
      Top             =   30
      Width           =   6540
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         TabIndex        =   401
         Text            =   "Combo4"
         Top             =   5070
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ordenado por código de instalacion"
         Height          =   435
         Index           =   4
         Left            =   3030
         TabIndex        =   402
         Top             =   5070
         Width           =   3225
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   11
         Left            =   2055
         MaxLength       =   15
         TabIndex        =   398
         Top             =   3870
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   10
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   397
         Top             =   3390
         Width           =   1335
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   3390
         TabIndex        =   406
         Text            =   "Text5"
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   3390
         TabIndex        =   405
         Text            =   "Text5"
         Top             =   3390
         Width           =   2535
      End
      Begin VB.CommandButton CmdCanListRecepDosimCuerpo 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   404
         Top             =   5670
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListRecepDosimCuerpo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   403
         Top             =   5670
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   25
         Left            =   3930
         TabIndex        =   400
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   24
         Left            =   2040
         TabIndex        =   399
         Top             =   4425
         Width           =   1020
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   17
         Left            =   1995
         MaxLength       =   11
         TabIndex        =   394
         Top             =   1755
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   16
         Left            =   1995
         MaxLength       =   11
         TabIndex        =   393
         Top             =   1305
         Width           =   1335
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   3390
         TabIndex        =   392
         Text            =   "Text5"
         Top             =   1770
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   3390
         TabIndex        =   391
         Text            =   "Text5"
         Top             =   1305
         Width           =   2535
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   3390
         TabIndex        =   390
         Text            =   "Text5"
         Top             =   2730
         Width           =   2535
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   3390
         TabIndex        =   389
         Text            =   "Text5"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   13
         Left            =   1980
         MaxLength       =   11
         TabIndex        =   396
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   12
         Left            =   1995
         MaxLength       =   11
         TabIndex        =   395
         Top             =   2265
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Paridad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   44
         Left            =   330
         TabIndex        =   426
         Top             =   4830
         Width           =   645
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   11
         Left            =   1785
         Picture         =   "frmListado.frx":3926
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   10
         Left            =   1785
         Picture         =   "frmListado.frx":3A28
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   55
         Left            =   1200
         TabIndex        =   419
         Top             =   3405
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   54
         Left            =   1200
         TabIndex        =   418
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DNI Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   41
         Left            =   345
         TabIndex        =   417
         Top             =   3120
         Width           =   1080
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   27
         Left            =   1215
         TabIndex        =   416
         Top             =   4470
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   26
         Left            =   3180
         TabIndex        =   415
         Top             =   4470
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   25
         Left            =   3690
         Picture         =   "frmListado.frx":3B2A
         Top             =   4425
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   24
         Left            =   1800
         Picture         =   "frmListado.frx":3C2C
         Top             =   4425
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Informe de Recepción de Dosímetros a Cuerpo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   16
         Left            =   135
         TabIndex        =   414
         Top             =   405
         Width           =   6180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Creación Recepción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   40
         Left            =   330
         TabIndex        =   413
         Top             =   4140
         Width           =   2190
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   17
         Left            =   1740
         Picture         =   "frmListado.frx":3D2E
         Top             =   1755
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   16
         Left            =   1740
         Picture         =   "frmListado.frx":3E30
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   53
         Left            =   1155
         TabIndex        =   412
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   52
         Left            =   1155
         TabIndex        =   411
         Top             =   1755
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   39
         Left            =   360
         TabIndex        =   410
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   38
         Left            =   360
         TabIndex        =   409
         Top             =   2070
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   51
         Left            =   1155
         TabIndex        =   408
         Top             =   2745
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   50
         Left            =   1155
         TabIndex        =   407
         Top             =   2310
         Width           =   525
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   13
         Left            =   1755
         Picture         =   "frmListado.frx":3F32
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   12
         Left            =   1755
         Picture         =   "frmListado.frx":4034
         Top             =   2265
         Width           =   240
      End
   End
   Begin VB.Frame FrameListEmpresas 
      Height          =   3870
      Left            =   30
      TabIndex        =   66
      Top             =   30
      Width           =   6135
      Begin VB.Frame Frame19 
         Caption         =   "Situación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   180
         TabIndex        =   475
         Top             =   1890
         Width           =   2895
         Begin VB.OptionButton OptEmp 
            Caption         =   "Baja"
            Height          =   435
            Index           =   3
            Left            =   1050
            TabIndex        =   478
            Top             =   300
            Width           =   765
         End
         Begin VB.OptionButton OptEmp 
            Caption         =   "Alta "
            Height          =   435
            Index           =   2
            Left            =   180
            TabIndex        =   477
            Top             =   270
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton OptEmp 
            Caption         =   "Todos"
            Height          =   435
            Index           =   4
            Left            =   1920
            TabIndex        =   476
            Top             =   300
            Width           =   765
         End
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1590
         TabIndex        =   422
         Top             =   3000
         Width           =   1395
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   74
         Text            =   "Text5"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   1
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   71
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   70
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton CmdAceptarLisEmp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3795
         TabIndex        =   72
         Top             =   3090
         Width           =   975
      End
      Begin VB.CommandButton cmdCanLisEmp 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4875
         TabIndex        =   73
         Top             =   3090
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   840
         Left            =   3120
         TabIndex        =   67
         Top             =   1890
         Width           =   2760
         Begin VB.OptionButton OptEmp 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   1
            Left            =   1470
            TabIndex        =   69
            Top             =   360
            Width           =   1125
         End
         Begin VB.OptionButton OptEmp 
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   68
            Top             =   360
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Dosimetria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   42
         Left            =   210
         TabIndex        =   421
         Top             =   3030
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   79
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Empresas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   660
         TabIndex        =   78
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   77
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   10
         Left            =   720
         TabIndex        =   76
         Top             =   1005
         Width           =   615
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":4136
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado.frx":4238
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.Frame FrameListUsu 
      Height          =   4440
      Left            =   30
      TabIndex        =   451
      Top             =   30
      Width           =   6540
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   15
         Left            =   2145
         MaxLength       =   15
         TabIndex        =   466
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   14
         Left            =   2115
         MaxLength       =   15
         TabIndex        =   465
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   3480
         TabIndex        =   464
         Text            =   "Text5"
         Top             =   1530
         Width           =   2535
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   3480
         TabIndex        =   463
         Text            =   "Text5"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton CmdCanListUsu 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5070
         TabIndex        =   462
         Top             =   3735
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListUsu 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3990
         TabIndex        =   461
         Top             =   3735
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   26
         Left            =   4050
         TabIndex        =   460
         Top             =   2055
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   20
         Left            =   2130
         TabIndex        =   459
         Top             =   2055
         Width           =   1020
      End
      Begin VB.Frame Frame22 
         Caption         =   "Tipo de Informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   915
         Left            =   270
         TabIndex        =   456
         Top             =   2430
         Width           =   2895
         Begin VB.OptionButton OptOpe 
            Caption         =   "General"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   458
            Top             =   390
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "Específico"
            Height          =   435
            Index           =   16
            Left            =   1500
            TabIndex        =   457
            Top             =   300
            Width           =   1245
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Situación de Operarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   915
         Left            =   3330
         TabIndex        =   452
         Top             =   2430
         Width           =   2895
         Begin VB.OptionButton OptOpe 
            Caption         =   "Baja"
            Height          =   435
            Index           =   11
            Left            =   1020
            TabIndex        =   455
            Top             =   330
            Width           =   765
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "Alta "
            Height          =   435
            Index           =   10
            Left            =   180
            TabIndex        =   454
            Top             =   330
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "Todos"
            Height          =   435
            Index           =   12
            Left            =   1920
            TabIndex        =   453
            Top             =   330
            Width           =   765
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar5 
         Height          =   510
         Left            =   270
         TabIndex        =   467
         Top             =   3495
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   15
         Left            =   1875
         Picture         =   "frmListado.frx":433A
         Top             =   1530
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   14
         Left            =   1875
         Picture         =   "frmListado.frx":443C
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   75
         Left            =   1290
         TabIndex        =   474
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   74
         Left            =   1290
         TabIndex        =   473
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DNI Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   51
         Left            =   420
         TabIndex        =   472
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   23
         Left            =   1305
         TabIndex        =   471
         Top             =   2100
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   22
         Left            =   3270
         TabIndex        =   470
         Top             =   2100
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   26
         Left            =   3780
         Picture         =   "frmListado.frx":453E
         Top             =   2055
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   20
         Left            =   1890
         Picture         =   "frmListado.frx":4640
         Top             =   2055
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Operarios "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   18
         Left            =   855
         TabIndex        =   469
         Top             =   390
         Width           =   4650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   50
         Left            =   450
         TabIndex        =   468
         Top             =   1770
         Width           =   885
      End
   End
   Begin VB.Frame FrameListTiposTrabajo 
      Height          =   4590
      Left            =   30
      TabIndex        =   265
      Top             =   30
      Width           =   6135
      Begin VB.Frame Frame14 
         Height          =   600
         Left            =   1500
         TabIndex        =   273
         Top             =   3090
         Width           =   3660
         Begin VB.OptionButton optRGe 
            Caption         =   "Código"
            Height          =   255
            Index           =   5
            Left            =   405
            TabIndex        =   278
            Top             =   225
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optRGe 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   4
            Left            =   2070
            TabIndex        =   277
            Top             =   225
            Width           =   1245
         End
      End
      Begin VB.CommandButton CmdCanTiposTrab 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4830
         TabIndex        =   276
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListTipoTrab 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3750
         TabIndex        =   275
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtRGe 
         Height          =   285
         Index           =   5
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   271
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtRGe 
         Height          =   285
         Index           =   4
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   270
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox DtxtRGe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2670
         TabIndex        =   269
         Text            =   "Text5"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox DtxtRGe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2670
         TabIndex        =   268
         Text            =   "Text5"
         Top             =   1020
         Width           =   2535
      End
      Begin VB.TextBox txtTTr 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   274
         Text            =   "290"
         Top             =   2550
         Width           =   975
      End
      Begin VB.TextBox txtTTr 
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   272
         Text            =   "289"
         Top             =   2130
         Width           =   975
      End
      Begin VB.TextBox DtxtTTr 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2670
         TabIndex        =   267
         Text            =   "Text5"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox DtxtTTr 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2670
         TabIndex        =   266
         Text            =   "Text5"
         Top             =   2130
         Width           =   2535
      End
      Begin VB.Image ImgRGe 
         Height          =   240
         Index           =   5
         Left            =   1320
         Picture         =   "frmListado.frx":4742
         Top             =   1470
         Width           =   240
      End
      Begin VB.Image ImgRGe 
         Height          =   240
         Index           =   4
         Left            =   1320
         Picture         =   "frmListado.frx":4844
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   33
         Left            =   720
         TabIndex        =   285
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   32
         Left            =   720
         TabIndex        =   284
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Tipos de Trabajo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   11
         Left            =   600
         TabIndex        =   283
         Top             =   300
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código Rama Genérica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   282
         Top             =   720
         Width           =   1905
      End
      Begin VB.Image ImgTTr 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":4946
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image ImgTTr 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado.frx":4A48
         Top             =   2130
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   31
         Left            =   720
         TabIndex        =   281
         Top             =   2115
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   30
         Left            =   720
         TabIndex        =   280
         Top             =   2550
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   279
         Top             =   1830
         Width           =   570
      End
   End
   Begin VB.Frame FrameListTipoMedicion 
      Height          =   3570
      Left            =   30
      TabIndex        =   195
      Top             =   30
      Width           =   6135
      Begin VB.TextBox DtxtTMe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2670
         TabIndex        =   204
         Text            =   "Text5"
         Top             =   1470
         Width           =   2535
      End
      Begin VB.TextBox DtxtTMe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2670
         TabIndex        =   203
         Text            =   "Text5"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox txtTMe 
         Height          =   285
         Index           =   1
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   200
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtTMe 
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   199
         Top             =   990
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListTipoMedicion 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3780
         TabIndex        =   201
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton CmdCanListTipoMedicion 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   202
         Top             =   2880
         Width           =   975
      End
      Begin VB.Frame Frame11 
         Height          =   600
         Left            =   1500
         TabIndex        =   196
         Top             =   2040
         Width           =   3660
         Begin VB.OptionButton optTMe 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   1
            Left            =   2070
            TabIndex        =   198
            Top             =   225
            Width           =   1245
         End
         Begin VB.OptionButton optTMe 
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   405
            TabIndex        =   197
            Top             =   225
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Medición"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   208
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Tipos de Medición"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   207
         Top             =   300
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   21
         Left            =   720
         TabIndex        =   206
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   20
         Left            =   720
         TabIndex        =   205
         Top             =   1005
         Width           =   615
      End
      Begin VB.Image ImgTMe 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":4B4A
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image ImgTMe 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado.frx":4C4C
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.Frame FrameListRamasGenericas 
      Height          =   3570
      Left            =   30
      TabIndex        =   209
      Top             =   30
      Width           =   6135
      Begin VB.Frame Frame12 
         Height          =   600
         Left            =   1500
         TabIndex        =   216
         Top             =   2040
         Width           =   3660
         Begin VB.OptionButton optRGe 
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   405
            TabIndex        =   218
            Top             =   225
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optRGe 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   1
            Left            =   2070
            TabIndex        =   217
            Top             =   225
            Width           =   1245
         End
      End
      Begin VB.CommandButton CmdCanListRamasGenericas 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   215
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListRamasGenericas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3780
         TabIndex        =   214
         Top             =   2910
         Width           =   975
      End
      Begin VB.TextBox txtRGe 
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   212
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox txtRGe 
         Height          =   285
         Index           =   1
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   213
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox DtxtRGe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2670
         TabIndex        =   211
         Text            =   "Text5"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox DtxtRGe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2670
         TabIndex        =   210
         Text            =   "Text5"
         Top             =   1470
         Width           =   2535
      End
      Begin VB.Image ImgRGe 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado.frx":4D4E
         Top             =   990
         Width           =   240
      End
      Begin VB.Image ImgRGe 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":4E50
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   23
         Left            =   720
         TabIndex        =   222
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   22
         Left            =   720
         TabIndex        =   221
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Ramas Genericas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   220
         Top             =   300
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   219
         Top             =   720
         Width           =   570
      End
   End
   Begin VB.Frame FrameListRamasEspec 
      Height          =   4590
      Left            =   30
      TabIndex        =   244
      Top             =   30
      Width           =   6135
      Begin VB.TextBox DtxtREs 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2670
         TabIndex        =   261
         Text            =   "Text5"
         Top             =   2580
         Width           =   2535
      End
      Begin VB.TextBox DtxtREs 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2670
         TabIndex        =   260
         Text            =   "Text5"
         Top             =   2100
         Width           =   2535
      End
      Begin VB.TextBox txtREs 
         Height          =   285
         Index           =   1
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   251
         Top             =   2550
         Width           =   975
      End
      Begin VB.TextBox txtREs 
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   250
         Top             =   2100
         Width           =   975
      End
      Begin VB.TextBox DtxtRGe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2670
         TabIndex        =   254
         Text            =   "Text5"
         Top             =   1470
         Width           =   2535
      End
      Begin VB.TextBox DtxtRGe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2670
         TabIndex        =   252
         Text            =   "Text5"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox txtRGe 
         Height          =   285
         Index           =   3
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   249
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtRGe 
         Height          =   285
         Index           =   2
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   248
         Top             =   990
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListRamasEsp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3750
         TabIndex        =   253
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton CmdCanListRamasEsp 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4830
         TabIndex        =   255
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame Frame13 
         Height          =   600
         Left            =   1500
         TabIndex        =   245
         Top             =   3180
         Width           =   3660
         Begin VB.OptionButton optRGe 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   3
            Left            =   2070
            TabIndex        =   247
            Top             =   225
            Width           =   1245
         End
         Begin VB.OptionButton optRGe 
            Caption         =   "Código"
            Height          =   255
            Index           =   2
            Left            =   405
            TabIndex        =   246
            Top             =   225
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   264
         Top             =   1830
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   29
         Left            =   720
         TabIndex        =   263
         Top             =   2550
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   28
         Left            =   720
         TabIndex        =   262
         Top             =   2115
         Width           =   615
      End
      Begin VB.Image ImgREs 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":4F52
         Top             =   2550
         Width           =   240
      End
      Begin VB.Image ImgREs 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado.frx":5054
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código Rama Genérica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   259
         Top             =   720
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Ramas Específicas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   10
         Left            =   600
         TabIndex        =   258
         Top             =   300
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   27
         Left            =   720
         TabIndex        =   257
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   26
         Left            =   720
         TabIndex        =   256
         Top             =   1005
         Width           =   615
      End
      Begin VB.Image ImgRGe 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmListado.frx":5156
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image ImgRGe 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado.frx":5258
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.Frame FrameListProvincias 
      Height          =   3570
      Left            =   30
      TabIndex        =   181
      Top             =   60
      Width           =   6135
      Begin VB.Frame Frame10 
         Height          =   600
         Left            =   1500
         TabIndex        =   187
         Top             =   2040
         Width           =   3660
         Begin VB.OptionButton optPro 
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   405
            TabIndex        =   190
            Top             =   225
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optPro 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   1
            Left            =   2070
            TabIndex        =   189
            Top             =   225
            Width           =   1245
         End
      End
      Begin VB.CommandButton CmdCanListPro 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4875
         TabIndex        =   188
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListProvincias 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3795
         TabIndex        =   186
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtPro 
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   11
         TabIndex        =   184
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox txtPro 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   185
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox DtxtPro 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2670
         TabIndex        =   183
         Text            =   "Text5"
         Top             =   990
         Width           =   2535
      End
      Begin VB.TextBox DtxtPro 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   182
         Text            =   "Text5"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Image ImgPro 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmListado.frx":535A
         Top             =   990
         Width           =   240
      End
      Begin VB.Image ImgPro 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":545C
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   19
         Left            =   720
         TabIndex        =   194
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   18
         Left            =   720
         TabIndex        =   193
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Provincias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   192
         Top             =   300
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   191
         Top             =   720
         Width           =   780
      End
   End
   Begin VB.Frame FrameListFondos 
      Height          =   3615
      Left            =   30
      TabIndex        =   228
      Top             =   30
      Width           =   6360
      Begin VB.OptionButton Option2 
         Caption         =   "Extremidad"
         Height          =   375
         Index           =   2
         Left            =   3315
         TabIndex        =   488
         Tag             =   "E"
         Top             =   2370
         Width           =   1230
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Todos"
         Height          =   375
         Index           =   0
         Left            =   1260
         TabIndex        =   486
         Top             =   2370
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Solapa"
         Height          =   375
         Index           =   1
         Left            =   2265
         TabIndex        =   485
         Tag             =   "S"
         Top             =   2370
         Width           =   855
      End
      Begin VB.CommandButton CmdCanListFondos 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4875
         TabIndex        =   234
         Top             =   2955
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListFondos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3780
         TabIndex        =   233
         Top             =   2955
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   15
         Left            =   4095
         TabIndex        =   232
         Top             =   1905
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   14
         Left            =   2175
         TabIndex        =   231
         Top             =   1905
         Width           =   1020
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   13
         Left            =   4095
         TabIndex        =   230
         Top             =   1125
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   12
         Left            =   2175
         TabIndex        =   229
         Top             =   1140
         Width           =   1020
      End
      Begin MSComctlLib.ProgressBar ProgressBar4 
         Height          =   510
         Left            =   465
         TabIndex        =   235
         Top             =   2910
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   49
         Left            =   480
         TabIndex        =   487
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   15
         Left            =   1380
         TabIndex        =   242
         Top             =   1185
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   14
         Left            =   3345
         TabIndex        =   241
         Top             =   1185
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   15
         Left            =   3825
         Picture         =   "frmListado.frx":555E
         Top             =   1905
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   14
         Left            =   1935
         Picture         =   "frmListado.frx":5660
         Top             =   1905
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Fondos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   9
         Left            =   465
         TabIndex        =   240
         Top             =   405
         Width           =   5490
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   20
         Left            =   525
         TabIndex        =   239
         Top             =   855
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   13
         Left            =   1350
         TabIndex        =   238
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   12
         Left            =   3315
         TabIndex        =   237
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   13
         Left            =   3855
         Picture         =   "frmListado.frx":5762
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   12
         Left            =   1935
         Picture         =   "frmListado.frx":5864
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Finalización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   19
         Left            =   495
         TabIndex        =   236
         Top             =   1590
         Width           =   1515
      End
   End
   Begin VB.Frame FrameListFactCalib 
      Height          =   3660
      Left            =   30
      TabIndex        =   118
      Top             =   30
      Width           =   6360
      Begin VB.OptionButton Option1 
         Caption         =   "Pulsera"
         Height          =   375
         Index           =   3
         Left            =   4395
         TabIndex        =   484
         Tag             =   "P"
         Top             =   2385
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Anillo"
         Height          =   375
         Index           =   2
         Left            =   3465
         TabIndex        =   483
         Tag             =   "A"
         Top             =   2385
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Solapa"
         Height          =   375
         Index           =   1
         Left            =   2355
         TabIndex        =   482
         Tag             =   "S"
         Top             =   2385
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   375
         Index           =   0
         Left            =   1350
         TabIndex        =   481
         Top             =   2385
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   7
         Left            =   4170
         TabIndex        =   122
         Top             =   1890
         Width           =   1020
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   2220
         TabIndex        =   121
         Top             =   1875
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   4170
         TabIndex        =   120
         Top             =   1110
         Width           =   1020
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   2265
         TabIndex        =   119
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptarFactCalib 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3795
         TabIndex        =   124
         Top             =   3030
         Width           =   975
      End
      Begin VB.CommandButton CmdCanLisFactCalib 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4845
         TabIndex        =   126
         Top             =   3030
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   510
         Left            =   510
         TabIndex        =   123
         Top             =   2955
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   48
         Left            =   570
         TabIndex        =   480
         Top             =   2235
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Finalización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   555
         TabIndex        =   132
         Top             =   1605
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   7
         Left            =   3930
         Picture         =   "frmListado.frx":5966
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   6
         Left            =   1965
         Picture         =   "frmListado.frx":5A68
         Top             =   1890
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   6
         Left            =   3465
         TabIndex        =   131
         Top             =   1935
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   5
         Left            =   1395
         TabIndex        =   130
         Top             =   1905
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   585
         TabIndex        =   129
         Top             =   885
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Factores de Calibración"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   2
         Left            =   465
         TabIndex        =   128
         Top             =   405
         Width           =   5490
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   5
         Left            =   3915
         Picture         =   "frmListado.frx":5B6A
         Top             =   1125
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   1995
         Picture         =   "frmListado.frx":5C6C
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   127
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   125
         Top             =   1155
         Width           =   525
      End
   End
   Begin VB.Frame FrameListLotes 
      Height          =   3615
      Left            =   0
      TabIndex        =   489
      Top             =   0
      Width           =   6360
      Begin VB.TextBox TextLot 
         Height          =   285
         Index           =   0
         Left            =   2175
         TabIndex        =   495
         Top             =   1140
         Width           =   1020
      End
      Begin VB.TextBox TextLot 
         Height          =   285
         Index           =   1
         Left            =   4095
         TabIndex        =   496
         Top             =   1125
         Width           =   1095
      End
      Begin VB.TextBox TextLot 
         Height          =   285
         Index           =   2
         Left            =   2175
         TabIndex        =   497
         Top             =   1905
         Width           =   1020
      End
      Begin VB.TextBox TextLot 
         Height          =   285
         Index           =   3
         Left            =   4095
         TabIndex        =   498
         Top             =   1905
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptarLotes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3780
         TabIndex        =   494
         Top             =   2955
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarLotes 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4875
         TabIndex        =   493
         Top             =   2955
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Solapa"
         Height          =   375
         Index           =   5
         Left            =   2265
         TabIndex        =   492
         Tag             =   "S"
         Top             =   2370
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Todos"
         Height          =   375
         Index           =   4
         Left            =   1260
         TabIndex        =   491
         Top             =   2370
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Extremidad"
         Height          =   375
         Index           =   3
         Left            =   3315
         TabIndex        =   490
         Tag             =   "E"
         Top             =   2370
         Width           =   1230
      End
      Begin MSComctlLib.ProgressBar ProgressBar6 
         Height          =   510
         Left            =   465
         TabIndex        =   499
         Top             =   2910
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dosímetro Final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   54
         Left            =   495
         TabIndex        =   507
         Top             =   1590
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   31
         Left            =   3315
         TabIndex        =   506
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   30
         Left            =   1350
         TabIndex        =   505
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dosímetro Inicial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   53
         Left            =   525
         TabIndex        =   504
         Top             =   855
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Lotes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   19
         Left            =   465
         TabIndex        =   503
         Top             =   405
         Width           =   5490
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   29
         Left            =   3345
         TabIndex        =   502
         Top             =   1185
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   28
         Left            =   1380
         TabIndex        =   501
         Top             =   1185
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   52
         Left            =   480
         TabIndex        =   500
         Top             =   2220
         Width           =   360
      End
   End
   Begin VB.Frame FrameListDosimetros 
      Height          =   7500
      Left            =   75
      TabIndex        =   80
      Top             =   15
      Width           =   7230
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   2190
         TabIndex        =   91
         Top             =   4350
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   4350
         TabIndex        =   92
         Top             =   4350
         Width           =   1095
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2190
         TabIndex        =   99
         Tag             =   "Tipo Dosimetro|N|N|||dosimetros|tipo_dosimetro||N|"
         Text            =   "Combo2"
         Top             =   6210
         Width           =   1305
      End
      Begin VB.Frame Frame18 
         Caption         =   "Situación Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   4080
         TabIndex        =   450
         Top             =   5460
         Width           =   2895
         Begin VB.OptionButton OptOpe 
            Caption         =   "Baja"
            Height          =   315
            Index           =   9
            Left            =   1020
            TabIndex        =   97
            Top             =   240
            Width           =   765
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "Alta "
            Height          =   315
            Index           =   8
            Left            =   180
            TabIndex        =   96
            Top             =   240
            Width           =   765
         End
         Begin VB.OptionButton OptOpe 
            Caption         =   "Todos"
            Height          =   315
            Index           =   7
            Left            =   1920
            TabIndex        =   98
            Top             =   240
            Value           =   -1  'True
            Width           =   765
         End
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2190
         TabIndex        =   95
         Tag             =   "Tipo Dosimetro|N|N|||dosimetros|tipo_dosimetro||N|"
         Text            =   "Combo2"
         Top             =   5580
         Width           =   1305
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   3600
         TabIndex        =   224
         Text            =   "Text5"
         Top             =   2310
         Width           =   3100
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3600
         TabIndex        =   223
         Text            =   "Text5"
         Top             =   1980
         Width           =   3100
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   7
         Left            =   2190
         MaxLength       =   11
         TabIndex        =   86
         Top             =   2310
         Width           =   1335
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   6
         Left            =   2190
         MaxLength       =   11
         TabIndex        =   85
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   3
         Left            =   2190
         MaxLength       =   15
         TabIndex        =   88
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   2
         Left            =   2190
         MaxLength       =   15
         TabIndex        =   87
         Top             =   2910
         Width           =   1335
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3600
         TabIndex        =   103
         Text            =   "Text5"
         Top             =   3240
         Width           =   3100
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3600
         TabIndex        =   102
         Text            =   "Text5"
         Top             =   2910
         Width           =   3100
      End
      Begin VB.CommandButton CmdCanLisDosim 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5760
         TabIndex        =   101
         Top             =   6870
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarDosim 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4680
         TabIndex        =   100
         Top             =   6870
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   4350
         TabIndex        =   90
         Top             =   3795
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   2190
         TabIndex        =   89
         Top             =   3795
         Width           =   1095
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   7
         Left            =   2190
         MaxLength       =   11
         TabIndex        =   84
         Top             =   1335
         Width           =   1335
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   6
         Left            =   2190
         MaxLength       =   11
         TabIndex        =   83
         Top             =   1005
         Width           =   1335
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   3600
         TabIndex        =   82
         Text            =   "Text5"
         Top             =   1350
         Width           =   3100
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3600
         TabIndex        =   81
         Text            =   "Text5"
         Top             =   1005
         Width           =   3100
      End
      Begin VB.TextBox txtDos 
         Height          =   285
         Index           =   1
         Left            =   4350
         MaxLength       =   8
         TabIndex        =   94
         Top             =   4980
         Width           =   975
      End
      Begin VB.TextBox txtDos 
         Height          =   285
         Index           =   0
         Left            =   2190
         MaxLength       =   8
         TabIndex        =   93
         Top             =   4980
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   345
         TabIndex        =   104
         Top             =   6900
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   1000
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Baja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   56
         Left            =   480
         TabIndex        =   511
         Top             =   4155
         Width           =   915
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   240
         Index           =   27
         Left            =   1950
         Picture         =   "frmListado.frx":5D6E
         Top             =   4350
         Width           =   240
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   240
         Index           =   28
         Left            =   4080
         Picture         =   "frmListado.frx":5E70
         Top             =   4350
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   19
         Left            =   3570
         TabIndex        =   510
         Top             =   4395
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   18
         Left            =   1380
         TabIndex        =   509
         Top             =   4395
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   55
         Left            =   495
         TabIndex        =   508
         Top             =   5985
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Dosimetro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   21
         Left            =   480
         TabIndex        =   243
         Top             =   5340
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   18
         Left            =   480
         TabIndex        =   227
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   25
         Left            =   1320
         TabIndex        =   226
         Top             =   2310
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   24
         Left            =   1320
         TabIndex        =   225
         Top             =   1995
         Width           =   615
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   7
         Left            =   1950
         Picture         =   "frmListado.frx":5F72
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   6
         Left            =   1950
         Picture         =   "frmListado.frx":6074
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   3
         Left            =   1950
         Picture         =   "frmListado.frx":6176
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   2
         Left            =   1950
         Picture         =   "frmListado.frx":6278
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   7
         Left            =   1380
         TabIndex        =   117
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   6
         Left            =   1380
         TabIndex        =   116
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DNI Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   115
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   114
         Top             =   3840
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   0
         Left            =   3570
         TabIndex        =   113
         Top             =   3840
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   4080
         Picture         =   "frmListado.frx":637A
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   1950
         Picture         =   "frmListado.frx":647C
         Top             =   3795
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Dosímetros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   0
         Left            =   465
         TabIndex        =   112
         Top             =   405
         Width           =   5490
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   111
         Top             =   3600
         Width           =   885
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   7
         Left            =   1950
         Picture         =   "frmListado.frx":657E
         Top             =   1335
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   6
         Left            =   1950
         Picture         =   "frmListado.frx":6680
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   3
         Left            =   1365
         TabIndex        =   110
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   1365
         TabIndex        =   109
         Top             =   1335
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   108
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número Dosimetro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   107
         Top             =   4740
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   3570
         TabIndex        =   106
         Top             =   4980
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   1380
         TabIndex        =   105
         Top             =   4980
         Width           =   525
      End
   End
   Begin VB.Frame FrameListDosisNHomOpe 
      Height          =   5220
      Left            =   60
      TabIndex        =   286
      Top             =   30
      Width           =   6135
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   3060
         TabIndex        =   519
         Text            =   "Text5"
         Top             =   1620
         Width           =   2535
      End
      Begin VB.TextBox DtxtEmp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   3060
         TabIndex        =   518
         Text            =   "Text5"
         Top             =   2070
         Width           =   2535
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   21
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   289
         Top             =   1620
         Width           =   1305
      End
      Begin VB.TextBox txtEmp 
         Height          =   285
         Index           =   20
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   290
         Top             =   2055
         Width           =   1305
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   9
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   292
         Top             =   3030
         Width           =   1335
      End
      Begin VB.TextBox txtIns 
         Height          =   285
         Index           =   8
         Left            =   1590
         MaxLength       =   11
         TabIndex        =   291
         Top             =   2580
         Width           =   1305
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   3060
         TabIndex        =   300
         Text            =   "Text5"
         Top             =   3060
         Width           =   2535
      End
      Begin VB.TextBox DtxtIns 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   3060
         TabIndex        =   299
         Text            =   "Text5"
         Top             =   2610
         Width           =   2535
      End
      Begin VB.CommandButton CmdCanListDosisNHomOpe 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4665
         TabIndex        =   296
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarListDosisNHomOpe 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3585
         TabIndex        =   295
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   4
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   293
         Top             =   3570
         Width           =   1320
      End
      Begin VB.TextBox txtOpe 
         Height          =   285
         Index           =   5
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   294
         Top             =   4020
         Width           =   1320
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3060
         TabIndex        =   298
         Text            =   "Text5"
         Top             =   3570
         Width           =   2535
      End
      Begin VB.TextBox DtxtOpe 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3060
         TabIndex        =   297
         Text            =   "Text5"
         Top             =   4020
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   17
         Left            =   3930
         TabIndex        =   288
         Top             =   1005
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   16
         Left            =   1590
         TabIndex        =   287
         Top             =   1005
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   59
         Left            =   255
         TabIndex        =   522
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   67
         Left            =   765
         TabIndex        =   521
         Top             =   2055
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   66
         Left            =   765
         TabIndex        =   520
         Top             =   1635
         Width           =   525
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   21
         Left            =   1320
         Picture         =   "frmListado.frx":6782
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image ImgEmp 
         Height          =   240
         Index           =   20
         Left            =   1320
         Picture         =   "frmListado.frx":6884
         Top             =   2055
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   9
         Left            =   1320
         Picture         =   "frmListado.frx":6986
         Top             =   3030
         Width           =   240
      End
      Begin VB.Image ImgIns 
         Height          =   240
         Index           =   8
         Left            =   1320
         Picture         =   "frmListado.frx":6A88
         Top             =   2610
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   37
         Left            =   720
         TabIndex        =   310
         Top             =   2655
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   36
         Left            =   720
         TabIndex        =   309
         Top             =   3030
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   308
         Top             =   2370
         Width           =   945
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   4
         Left            =   1320
         Picture         =   "frmListado.frx":6B8A
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image ImgOpe 
         Height          =   240
         Index           =   5
         Left            =   1320
         Picture         =   "frmListado.frx":6C8C
         Top             =   4020
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   35
         Left            =   720
         TabIndex        =   307
         Top             =   3585
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   34
         Left            =   720
         TabIndex        =   306
         Top             =   4020
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado de Dosis no Homogenea por Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   305
         Top             =   240
         Width           =   5715
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dni Operario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   27
         Left            =   240
         TabIndex        =   304
         Top             =   3300
         Width           =   1050
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   17
         Left            =   735
         TabIndex        =   303
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   16
         Left            =   3060
         TabIndex        =   302
         Top             =   1020
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   17
         Left            =   3660
         Picture         =   "frmListado.frx":6D8E
         Top             =   1005
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   16
         Left            =   1350
         Picture         =   "frmListado.frx":6E90
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Dosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   26
         Left            =   270
         TabIndex        =   301
         Top             =   690
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
    '1 .- Listado Empresas
    '2 .- Listado de instalaciones
    '3 .- Listado de Operarios en instalaciones
    '4 .- Listado de Dosimetros  (todos)
    '5 .- Listado de Dosimetros organo (no se usa)
    '6 .- Listado de Dosimetros de Area (no se usa)
    '7 .- Listado de Factores de calibracion 4400
    '8 .- Listado de Factores de calibracion 6600
    '9 .- Listado de Dosis por instalacion
    '12.- Listado de Dosis CSN
    '13.- Listado de Provincias
    '14.- Listado de tipos de medicion
    '15.- Listado de ramas genericas
    '16.- Listado de ramas especificas
    '17.- Listado de tipos de trabajo
    '18.- Listado de Fondos 6600
    '19.- Listado de dosis no  homogeneas por operario
    '20.- Cartas de dosimetros no recibidos por instalaciones
    '21.- Listado de Dosis por operario (acumulado 12 meses)
    '22.- Carta de sobredosis al CSN
    '23.- Listado de Recepcion de dosimetros personal
    '24.- listado de Etiquetas de empresas
    '23.- Listado de Recepcion de dosimetros personal
    '24.- Listado de etiquetas de Empresas
    '25.- Listado de Instalaciones
    '26.- Listado de Operarios
    '27.- listado de dosimetros de area recepcionados
    '28.- Listado de Operarios con Sobredosis
    '29.- Listado de Lotes 6600
    '30.- Listado de Lotes Panasonic
    '31.- Listado de Factores de calibracion panasonic
    '32.- Listado de Fondos Panasonic
    'xx.- Proceso de Insercion Automatica de Dosis por Penalizacion
    
'******************
    
Public Event DatoSeleccionado(CadenaSeleccion As String)

Dim Tablas As String
    
Private WithEvents frmM As frmMensajes
Attribute frmM.VB_VarHelpID = -1

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmEmp As frmEmpresas
Attribute frmEmp.VB_VarHelpID = -1
Private WithEvents frmPro As frmProvincias
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmRGe As frmRamasGener
Attribute frmRGe.VB_VarHelpID = -1
Private WithEvents frmREs As frmRamasEspe
Attribute frmREs.VB_VarHelpID = -1
Private WithEvents frmTMe As frmTiposExtremidades
Attribute frmTMe.VB_VarHelpID = -1
Private WithEvents frmTTr As frmTiposTrab
Attribute frmTTr.VB_VarHelpID = -1
Private WithEvents frmIns As frmInstalaciones
Attribute frmIns.VB_VarHelpID = -1
Private WithEvents frmOpe As frmOperarios
Attribute frmOpe.VB_VarHelpID = -1

Private Empresa As String
Private instalacion As String
Private Operario As String

Dim sql As String
Dim RC As String
Dim rs As Recordset
Dim PrimeraVez As Boolean


Dim Cad As String
Dim cad1 As String
Dim cont As Long
Dim I As Integer


Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub


'Private Sub CmdAceptarListOpe_Click()
''Imprimir el listado, segun sea
'    If Not ComprobarOperarios(4, 5) Then Exit Sub
'
'    'Hacemos el select y si tiene resultados mostramos los valores
'    Cad = " SELECT operarios.* from operarios WHERE 1 = 1 "
'    If txtOpe(4).Text <> "" Then Cad = Cad & " AND dni >= '" & Trim(txtOpe(4).Text) & "'"
'    If txtOpe(5).Text <> "" Then Cad = Cad & " AND dni <= '" & Trim(txtOpe(5).Text) & "'"
'
'    Set RS = New Adodb.Recordset
'    RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'    If RS.EOF Then
'        'NO hay registros a mostrar
'        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
'    Else
'        'Mostramos el frame de resultados
'        SQL = ""
'        cad1 = "Operarios: "
'        If txtOpe(4).Text <> "" Then
'            SQL = SQL & "desde= " & txtOpe(4).Text & "|"
'            cad1 = cad1 & " desde " & txtOpe(4).Text & " " & DtxtOpe(4).Text
'        End If
'
'        If txtOpe(5).Text <> "" Then
'            SQL = SQL & "hasta= " & txtOpe(5).Text & "|"
'            cad1 = cad1 & "   hasta " & txtOpe(5).Text & " " & DtxtOpe(5).Text
'        End If
'
'        If txtOpe(4).Text <> "" Or txtOpe(5).Text <> "" Then SQL = SQL & "Operarios= """ & cad1 & """|"
'
'        If OptOpe(0).Value = True Then
'            SQL = SQL & "orden= ""Por Código""|"
'            frmImprimir.CampoOrden = 1
'        End If
'        If OptOpe(1).Value = True Then
'            SQL = SQL & "orden= ""Alfabético""|"
'            frmImprimir.CampoOrden = 2
'        End If
'
'        frmImprimir.Opcion = 3
'        frmImprimir.NumeroParametros = 3
'        frmImprimir.FormulaSeleccion = SQL
'        frmImprimir.OtrosParametros = SQL
'        frmImprimir.SoloImprimir = False
'        frmImprimir.email = False
'        frmImprimir.Show 'vbModal
'
'     End If
'
'
'End Sub



Private Sub CmdAceptarCartaDosimNRec_Click()
Dim Tipo As Boolean
Dim Mes As String
Dim Mail As String
Dim HayRegistros As Boolean
Dim sql1 As String
Dim sql2 As String
Dim rs As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim cad1 As String
Dim cad2 As String
Dim Conta As Integer

    Screen.MousePointer = vbHourglass

    If Text3(18).Text = "" Or Text3(19).Text = "" Then
        MsgBox "Es obligatorio introducir los campos fecha.", vbExclamation, "¡Error!"
        PonerFoco Text3(18)
        Exit Sub
    End If

    If Not ComprobarFechas(18, 19) Then Exit Sub
    If Not ComprobarEmpresas(10, 11) Then Exit Sub
    If Not ComprobarInstalaciones(10, 11) Then Exit Sub
    'Hacemos el select y si tiene resultados mostramos los valores

    
'    Tipo = ((Month(CDate(Text3(18).Text)) / 2) = Round2(Month(CDate(Text3(18).Text)) / 2, 1))
    
    If (Month(CDate(Text3(18).Text)) Mod 2) = 0 Then
        Mes = "P"
    Else
        Mes = "I"
    End If
    
    Cad = " SELECT distinct dosimetros.* from dosimetros,operarios WHERE 1 = 1 and tipo_dosimetro = 0 "
    Cad = Cad & " and operarios.semigracsn = 1 and operarios.dni = dosimetros.dni_usuario "
    Cad = Cad & " AND dosimetros.f_retirada is null and mes_p_i = '" & Mes & "'"

    If txtEmp(10).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa >= '" & Trim(txtEmp(10).Text) & "' "
    If txtEmp(11).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa <= '" & Trim(txtEmp(11).Text) & "' "
    If txtIns(10).Text <> "" Then Cad = Cad & " AND dosimetros.c_instalacion >= '" & Trim(txtIns(10).Text) & "' "
    If txtIns(11).Text <> "" Then Cad = Cad & " AND dosimetros.c_instalacion <= '" & Trim(txtIns(11).Text) & "' "
    

'    Cad = Cad & " AND dosiscuerpo.f_dosis >= '" & Format(Text3(18).Text, FormatoFecha) & "' "
'    Cad = Cad & " AND dosiscuerpo.f_dosis <= '" & Format(Text3(19).Text, FormatoFecha) & "' "



    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        'Mostramos el frame de resultados
        
        ' borramos la tabla temporal
        sql1 = "delete from zdosimnorec where codusu= " & vUsu.codigo
        conn.Execute sql1
        
        'cargamos la tabla temporal
        HayRegistros = False
        While Not rs.EOF
            sql1 = "select * from dosiscuerpo where n_dosimetro = '" & Trim(rs!n_dosimetro) & "' "
            sql1 = sql1 & " AND dosiscuerpo.f_dosis >= '" & Format(Text3(18).Text, FormatoFecha) & "' "
            sql1 = sql1 & " AND dosiscuerpo.f_dosis <= '" & Format(Text3(19).Text, FormatoFecha) & "' "
            
            Set RT = New ADODB.Recordset
            RT.Open sql1, conn, adOpenKeyset, adLockPessimistic, adCmdText
            If RT.EOF Then
                sql2 = "insert into zdosimnorec (codusu, c_empresa, c_instalacion, n_dosimetro, dni_usuario) VALUES ("
                sql2 = sql2 & vUsu.codigo & ",'" & Trim(rs!c_empresa) & "','" & Trim(rs!c_instalacion) & "','"
                sql2 = sql2 & Trim(rs!n_dosimetro) & "','" & Trim(rs!dni_usuario) & "')"
                
                conn.Execute sql2
                
                HayRegistros = True
            End If
            Set RT = Nothing
            
            rs.MoveNext
        Wend
        If Not HayRegistros Then
            'NO hay registros a mostrar
            MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            
            If chkEmail.Value = 0 Then
                sql = "usu= " & vUsu.codigo & "|"
                
                'empresas
                cad1 = ""
                If txtEmp(10).Text <> "" Then
                    sql = sql & "desemp= """ & Trim(txtEmp(10).Text) & """|"
                    cad1 = cad1 & " desde " & txtEmp(10).Text & " " & DtxtEmp(10).Text
                End If
        
                If txtEmp(11).Text <> "" Then
                    sql = sql & "hasemp= """ & Trim(txtEmp(11).Text) & """|"
                    cad1 = cad1 & "   hasta " & txtEmp(11).Text & " " & DtxtEmp(11).Text
                End If
        
                If txtEmp(10).Text <> "" Or txtEmp(11).Text <> "" Then sql = sql & "Empresas= """ & cad1 & """|"
        
                cad1 = ""  'instalaciones
                If txtIns(10).Text <> "" Then
                    sql = sql & "desins= """ & Trim(txtIns(10).Text) & """|"
                    cad1 = cad1 & " desde " & txtIns(10).Text & " " & DtxtIns(10).Text
                End If
        
                If txtIns(11).Text <> "" Then
                    sql = sql & "hasins= """ & Trim(txtIns(11).Text) & """|"
                    cad1 = cad1 & "   hasta " & txtIns(11).Text & " " & DtxtIns(11).Text
                End If
        
                If txtIns(10).Text <> "" Or txtIns(11).Text <> "" Then sql = sql & "Instalaciones= """ & cad1 & """|"
        
                cad1 = ""  'fecha dosis
                If Text3(18).Text <> "" Then
                    sql = sql & "desfec= """ & Format(Text3(18).Text, FormatoFecha) & """|"
                    cad1 = cad1 & Trim(Text3(18).Text) & " "
                End If
        
                If Text3(19).Text <> "" Then
                    sql = sql & "hasfec= """ & Format(Text3(19).Text, FormatoFecha) & """|"
                    cad1 = cad1 & "   hasta " & Trim(Text3(19).Text)
                End If
        
                If Text3(18).Text <> "" Or Text3(19).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"
        
                frmImprimir.NumeroParametros = 10
                frmImprimir.Opcion = 21
                frmImprimir.email = False
                frmImprimir.FormulaSeleccion = sql
                frmImprimir.OtrosParametros = sql
                frmImprimir.SoloImprimir = False
                frmImprimir.Show 'vbModal
            Else
'                ' Enviar por email. IREMOS UNO A UNO
'                ' fechaadq = codmacta
'                Screen.MousePointer = vbHourglass
'
'                Cad = "DELETE FROM ztempemail WHERE codusu =" & vUsu.codigo
'                Conn.Execute Cad
'
'
'                RS.Close
'
'                Cad = " SELECT * from zdosimnorec where codusu = " & vUsu.codigo
'                Cad = Cad & " group by c_instalacion"
'
'                RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'        '
'        '        Cad = "select fechaadq,maidatos,razosoci,nommacta FROM USUARIOS.zentrefechas,cuentas WHERE"
'        '        Cad = Cad & " fechaadq=codmacta AND    CodUsu = " & vUsu.codigo
'        '        Cad = Cad & " GROUP BY fechaadq ORDER BY maidatos"
'        '        RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'        '
'
'                cad1 = "usu= " & vUsu.codigo & "|"
'                'fecha dosis
'                cad2 = Trim(Text3(18).Text) & " "
'                cad2 = cad2 & "   hasta " & Trim(Text3(19).Text)
'                cad1 = cad1 & "FechaAlta= """ & cad2 & """|"
'
'                NumRegElim = 0
'                Cont = 0
'                frmPpal.Visible = False
'
'                While Not RS.EOF
'                    Me.Refresh
'                    espera 0.5
'
'                    Mail = ""
'                    Mail = DevuelveDesdeBD(1, "mail_internet", "instalaciones", "c_instalacion|", Trim(RS!c_instalacion) & "|", "T|", 1)
'                    If Mail = "" Then
'                        MsgBox "Sin mail para la instalación: " & Trim(RS!c_instalacion), vbExclamation
'                        SQL = "INSERT INTO ztempemail (codusu, c_instalacion, email, fichero) values (" & vUsu.codigo
'                        SQL = SQL & ",'" & Trim(RS!c_instalacion) & "',NULL,'')"
'
'                        'AL meter la cuenta con el importe a 0, entonces no la leera para enviarala
'                        'Pero despues si k podremos NO actualizar sus pagosya que no se han enviado nada
'                        Conn.Execute SQL
'                    Else
'                        Cad = "desins= """ & Trim(RS!c_instalacion) & """|"
'                        Cad = Cad & "hasins= """ & Trim(RS!c_instalacion) & """|"
'                        Cad = Cad & cad1
'
'                        With frmImprimir
'                            .OtrosParametros = Cad
'                            .NumeroParametros = 10
'        '                    sql = "{ado.codusu}=" & vUsu.codigo & " AND {ado.nif}= """ & RS.Fields(0) & """"
'                            .FormulaSeleccion = Cad
'                            .email = True
'                            CadenaDesdeOtroForm = "GENERANDO"
'                            .Opcion = 21
'                            .Show vbModal
'
'                            If CadenaDesdeOtroForm = "" Then
'                                Me.Refresh
'                                espera 0.5
'                                Cont = Cont + 1
'                                'Se ha generado bien el documento
'                                'Lo copiamos sobre app.path & \temp
'                                SQL = "A" & Trim(RS!c_instalacion) & ".pdf"
'
'                                If Dir(App.Path & "\temp", vbDirectory) = "" Then
'                                    MkDir App.Path & "\temp"
'                                End If
'
'                                FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & SQL
'
'
'                                'Insertamos en al temporal
'                                Sql1 = "INSERT INTO ztempemail (codusu, c_instalacion, email, fichero) values (" & vUsu.codigo
'                                Sql1 = Sql1 & ",'" & Trim(RS!c_instalacion) & "','" & Trim(Mail) & "','" & Trim(SQL) & "')"
'                                Conn.Execute Sql1
'
'
'                            End If
'
'                        End With
'                    End If
'                    RS.MoveNext
'                Wend
'                RS.Close
'
'                If Cont > 0 Then
'
'                     espera 0.5
'
'                     SQL = "Carta de reclamación de dosimetros no recibidos"
'                     frmEMail.Opcion = 3
'                     frmEMail.MisDatos = SQL
'                     frmEMail.Show vbModal
'
'                End If
'                Screen.MousePointer = vbDefault
'
'                Me.Hide
'                frmPpal.Visible = True
'                Me.Visible = True
'                Me.Refresh
    
    
'lo vamos a guardar en el directorio temp
'estoy aqui
                
                Cad = " SELECT * from zdosimnorec where codusu = " & vUsu.codigo
                Cad = Cad & " group by c_instalacion"
                rs.Close
                rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
                cad1 = "usu= " & vUsu.codigo & "|"
                
                'fecha dosis
                cad2 = Trim(Text3(18).Text) & " "
                cad2 = cad2 & "   hasta " & Trim(Text3(19).Text)
                cad1 = cad1 & "FechaAlta= """ & cad2 & """|"

                NumRegElim = 0
                Conta = 0
                rs.MoveFirst
                
                While Not rs.EOF
                    Me.Refresh
                    espera 0.5

                    Cad = "desins= """ & Trim(rs!c_instalacion) & """|"
                    Cad = Cad & "hasins= """ & Trim(rs!c_instalacion) & """|"
                    Cad = Cad & cad1
'desde
                    With frmVisReport
                        .Informe = App.Path & "\informes\CartaDosimNRec.rpt"
                        .OtrosParametros = Cad
                        .NumeroParametros = 10
        '                    sql = "{ado.codusu}=" & vUsu.codigo & " AND {ado.nif}= """ & RS.Fields(0) & """"
                        .FormulaSeleccion = Cad
                        .ExportarPDF = True
                        CadenaDesdeOtroForm = "GENERANDO"
                        Load frmVisReport
                        Unload frmVisReport
                        
                        If CadenaDesdeOtroForm = "OK" Then
                            Me.Refresh
                            espera 0.5
                            'Se ha generado bien el documento
                            'Lo copiamos sobre app.path & \temp
                            sql1 = "Carta-A-" & Trim(rs!c_instalacion) & ".pdf"
                            
                            If Dir(App.Path & "\temp", vbDirectory) = "" Then
                                MkDir App.Path & "\temp"
                            End If
                            
                            FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & sql1
            
                        End If
                    End With
                    
                    rs.MoveNext
            
'hasta

                Wend
                rs.Close
            End If
            End If
       End If

Screen.MousePointer = vbDefault

End Sub

Private Sub chkEmail_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CmdAceptarCartaSobredosis_Click()
Dim rs As ADODB.Recordset
Dim rL As ADODB.Recordset
Dim rf As ADODB.Recordset
Dim Mesa As Integer
Dim Anoa As Integer
Dim ano As Integer
Dim fec As String
Dim sql1 As String
Dim sql2 As String
Dim situ As Integer
Dim fecha As String
Dim Valor1 As Currency
Dim valor2 As Currency



    If Not ComprobarFechas(22, 23) Then Exit Sub
    If Not ComprobarEmpresas(14, 15) Then Exit Sub
    If Not ComprobarOperarios(8, 9) Then Exit Sub
    
    'Hacemos el select y si tiene resultados mostramos los valores

    If Check1(0).Value = 1 Then
        ' solo los que no tienen fecha de baja
        Cad = "SELECT dosimetros.c_empresa, dosimetros.c_instalacion, dosimetros.dni_usuario, "
        Cad = Cad & "dosimetros.n_dosimetro from dosimetros, operarios  "
        
        If Check1(1).Value = 1 Then
            Cad = Cad & ", tempnc where dosimetros.n_dosimetro = tempnc.n_dosimetro "
        Else
            Cad = Cad & " where 1=1 "
        End If
        
        Cad = Cad & " and operarios.f_baja is null "
        Cad = Cad & " and operarios.c_empresa = dosimetros.c_empresa "
        Cad = Cad & " and operarios.c_instalacion = dosimetros.c_instalacion "

    Else
        ' todos los usuarios tengan o no fecha de baja
    
        Cad = "SELECT  c_empresa, c_instalacion, dni_usuario, n_dosimetro from dosimetros "
        
        If Check1(1).Value = 1 Then
            Cad = Cad & ", tempnc where dosimetros.n_dosimetro = tempnc.n_dosimetro "
        Else
            Cad = Cad & " where 1=1 "
        End If

    End If
    ' solo seleccionamos los dosimetros de cuerpo
    Cad = Cad & " AND dosimetros.tipo_dosimetro = 0 and dosimetros.f_retirada is null "
    
    
    If txtEmp(14).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa >= '" & Trim(txtEmp(14).Text) & "' "
    If txtEmp(15).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa <= '" & Trim(txtEmp(15).Text) & "' "
    If txtOpe(8).Text <> "" Then Cad = Cad & " AND dosimetros.dni_usuario >= '" & Trim(txtOpe(8).Text) & "' "
    If txtOpe(9).Text <> "" Then Cad = Cad & " AND dosimetros.dni_usuario <= '" & Trim(txtOpe(9).Text) & "' "

    'Cad = Cad & " order by dosimetros.c_empresa, dosimetros.dni_usuario, dosimetros.c_instalacion "
    
    
    ano = Year(CDate(Text3(22).Text))

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = "delete from zdosisacum where codusu = " & vUsu.codigo
        conn.Execute sql
        
        rs.MoveFirst
        While Not rs.EOF
            For I = 1 To 12
                sql1 = "select sum(dosis_superf), sum(dosis_profunda) from dosiscuerpo "
                sql1 = sql1 & " where dni_usuario = '" & Trim(rs.Fields(2).Value) & "' and "
                sql1 = sql1 & " c_instalacion = '" & Trim(rs.Fields(1).Value) & "' and "
                sql1 = sql1 & " month(f_dosis) = " & I & " and year(f_dosis) = " & Format(ano, "0000")
                
                Set rL = New ADODB.Recordset
                rL.Open sql1, conn, adOpenKeyset, adLockPessimistic, adCmdText
                          
                ' el campo situ me indica: 0- situacion normal
                '                          1- sin alta en SDE
                '                          2- dosimetro no recibido
                situ = 0
                fecha = "01/" & Format(I, "00") & "/" & Format(ano, "0000")
                If I < 12 Then
                    Mesa = I + 1
                    Anoa = ano
                Else
                    Mesa = 1
                    Anoa = ano + 1
                End If
                
                'cursor para averiguar la maxima fecha de alta de ese usuario en la instalacion
                Set rf = New ADODB.Recordset
                sql = "select max(f_alta) from operainstala where c_empresa = '" & Trim(rs.Fields(0).Value) & "' "
                sql = sql & " and c_instalacion = '" & Trim(rs.Fields(1).Value) & "' and "
                sql = sql & " dni = '" & Trim(rs.Fields(2).Value) & "'"
                rf.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
                fec = "31/12/2999"
                If Not rf.EOF Then
                    If IsNull(rf.Fields(0).Value) Then
                        fec = "31/12/2999"
                    Else
                        fec = CStr(rf.Fields(0).Value)
                    End If
                End If
                
                If CDate(fec) >= CDate("01/" & Format(Mesa, "00") & "/" & Format(Anoa, "0000")) Then
                    situ = 1 ' sin alta en SDE
                End If
                Set rf = Nothing
                sql2 = "insert into zdosisacum (codusu, c_empresa, c_instalacion, "
                sql2 = sql2 & "dni_usuario, mes, ano, n_dosimetro, dosissuper, dosisprofu, situ) VALUES ("
                sql2 = sql2 & vUsu.codigo & ",'" & Trim(rs.Fields(0).Value) & "','"
                sql2 = sql2 & Trim(rs.Fields(1).Value) & "','" & Trim(rs.Fields(2).Value) & "',"
                sql2 = sql2 & I & "," & ano & ",'" & Trim(rs.Fields(3).Value) & "',"
                          
                          
                If Not rL.EOF Then
                    If IsNull(rL.Fields(0).Value) Then
                        Valor1 = 0
                    Else
                        Valor1 = rL.Fields(0).Value
                    End If
                    
                    If IsNull(rL.Fields(1).Value) Then
                        valor2 = 0
                    Else
                        valor2 = rL.Fields(1).Value
                    End If
                    
                
                    sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(Valor1))) & ","
                    sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(valor2))) & ","
                    sql2 = sql2 & Format(situ, "0") & ")"
                Else
                    sql2 = sql2 & "0.0, 0.0, 2)"
                End If
                conn.Execute sql2
            
                Set rL = Nothing
                
            Next I
            
            rs.MoveNext
        Wend
        
        rs.Close
        
        
        ' tenemos que eliminar aquellos registros de la temporal que no lleguen a sobredosis
        sql = "select c_empresa,dni_usuario, sum(dosissuper), sum(dosisprofu) from zdosisacum where codusu= " & vUsu.codigo
        sql = sql & " group by c_empresa, dni_usuario having sum(dosissuper) < 500 and sum(dosisprofu) < 20 "
        
        rs.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
        If Not rs.EOF Then
            rs.MoveFirst
            
            While Not rs.EOF
                sql1 = "delete from zdosisacum where codusu = " & vUsu.codigo & " and "
                sql1 = sql1 & "c_empresa = '" & Trim(rs.Fields(0).Value) & "' and dni_usuario = '"
                sql1 = sql1 & Trim(rs.Fields(1).Value) & "'"
                
                conn.Execute sql1
            
                rs.MoveNext
            Wend
        
        Else
            'NO hay registros a mostrar
            MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
            Exit Sub
        End If
        
        rs.Close
        rs.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
        If rs.EOF Then
            'NO hay registros a mostrar
            MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
            Exit Sub
        End If
       
        ' una vez cargada la temporal imprimimos el informe
        sql = "usu= " & vUsu.codigo & "|"
        sql = sql & "FechaAlta= ""Registros dosimétricos desde: " & Text3(22).Text & " hasta " & Text3(23).Text & """|"
        
        frmImprimir.Opcion = 23
        frmImprimir.NumeroParametros = 8
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If



End Sub

Private Sub CmdAceptarDosim_Click()
Dim Tipo As Byte
Dim Sistema As String
    
    If Not ComprobarEmpresas(6, 7) Then Exit Sub
    If Not ComprobarInstalaciones(6, 7) Then Exit Sub
    If Not ComprobarOperarios(2, 3) Then Exit Sub
    If Not ComprobarFechas(2, 3) Then Exit Sub
    If Not ComprobarFechas(27, 28) Then Exit Sub
    If Not ComprobarDosimetros(0, 1) Then Exit Sub

    'Hacemos el select y si tiene resultados mostramos los valores
    Tipo = Combo3.ListIndex
    If Combo5.ListIndex <> -1 Then
      Sistema = IIf(Combo4.ListIndex = 0, "H", "P")
    End If
    
'    Cad = "SELECT d.* FROM dosimetros d,recepdosim r WHERE d.n_reg_dosimetro = r.n_reg_dosimetro"
'    Cad = Cad & " AND d.n_dosimetro = r.n_dosimetro AND d.dni_usuario = r.dni_usuario AND "
'    Cad = Cad & "d.c_empresa = r.c_empresa AND d.c_instalacion = r.c_instalacion AND "
'    Cad = Cad & "d.mes_p_i = r.mes_p_i AND d.tipo_dosimetro = r.tipo_dosimetro"
    
    Cad = "SELECT d.* FROM dosimetros d WHERE 1=1 "
    
'    If Tipo <> 3 Then
'        Cad = " SELECT dosimetros.* from dosimetros,recepdosim WHERE  tipo_dosimetro = " & Format(Tipo, "0")
'    Else
'        Cad = " SELECT dosimetros.* from dosimetros,recepdosim WHERE  1=1"
'    End If
    Sistema = ""
    If Combo5.ListIndex <> -1 And Combo5.ListIndex <> 0 Then
      Sistema = IIf(Combo5.ListIndex = 1, "H", "P")
      Cad = Cad & " and d.sistema = '" & Sistema & "'"
    End If
    
    If Tipo <> 3 Then Cad = Cad & " AND d.tipo_dosimetro = " & Format(Tipo, "0")
    If txtEmp(6).Text <> "" Then Cad = Cad & " AND d.c_empresa >= '" & Trim(txtEmp(6).Text) & "' "
    If txtEmp(7).Text <> "" Then Cad = Cad & " AND d.c_empresa <= '" & Trim(txtEmp(7).Text) & "' "
    If txtIns(6).Text <> "" Then Cad = Cad & " AND d.c_instalacion >= '" & Trim(txtIns(6).Text) & "' "
    If txtIns(7).Text <> "" Then Cad = Cad & " AND d.c_instalacion <= '" & Trim(txtIns(7).Text) & "' "
    If txtOpe(2).Text <> "" Then Cad = Cad & " AND d.dni_usuario >= '" & Trim(txtOpe(2).Text) & "' "
    If txtOpe(3).Text <> "" Then Cad = Cad & " AND d.dni_usuario <= '" & Trim(txtOpe(3).Text) & "' "
    If Text3(2).Text <> "" Then Cad = Cad & " AND d.f_asig_dosimetro >= '" & Format(Text3(2).Text, FormatoFecha) & "' "
    If Text3(3).Text <> "" Then Cad = Cad & " AND d.f_asig_dosimetro <= '" & Format(Text3(3).Text, FormatoFecha) & "' "
    If Text3(27).Text <> "" Then Cad = Cad & " AND d.f_retirada >= '" & Format(Text3(27).Text, FormatoFecha) & "' "
    If Text3(28).Text <> "" Then Cad = Cad & " AND d.f_retirada <= '" & Format(Text3(28).Text, FormatoFecha) & "' "
    If txtDos(0).Text <> "" Then Cad = Cad & " AND d.n_dosimetro>= '" & Trim(txtDos(0).Text) & "' "
    If txtDos(1).Text <> "" Then Cad = Cad & " AND d.n_dosimetro<= '" & Trim(txtDos(1).Text) & "' "
    
    Cad = Cad & " ORDER BY d.c_instalacion,orden_recepcion"
    
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Empresas: "
        If txtEmp(6).Text <> "" Then
            sql = sql & "desemp= """ & Trim(txtEmp(6).Text) & """|"
            cad1 = cad1 & " desde " & txtEmp(6).Text & " " & DtxtEmp(6).Text
        End If

        If txtEmp(7).Text <> "" Then
            sql = sql & "hasemp= """ & Trim(txtEmp(7).Text) & """|"
            cad1 = cad1 & "   hasta " & txtEmp(7).Text & " " & DtxtEmp(7).Text
        End If

        If txtEmp(6).Text <> "" Or txtEmp(7).Text <> "" Then sql = sql & "Empresas= """ & cad1 & """|"

        'instalaciones
        cad1 = "Instalaciones: "
        If txtIns(6).Text <> "" Then
            sql = sql & "desins= """ & Trim(txtIns(6).Text) & """|"
            cad1 = cad1 & " desde " & txtIns(6).Text & " " & DtxtIns(6).Text
        End If

        If txtIns(7).Text <> "" Then
            sql = sql & "hasins= """ & Trim(txtIns(7).Text) & """|"
            cad1 = cad1 & "   hasta " & txtIns(7).Text & " " & DtxtIns(7).Text
        End If

        If txtIns(6).Text <> "" Or txtIns(7).Text <> "" Then sql = sql & "Instalaciones= """ & cad1 & """|"


        ' dnis de operarios
        cad1 = "Operarios: "
        If txtOpe(2).Text <> "" Then
            sql = sql & "desope= """ & Trim(txtOpe(2).Text) & """|"
            cad1 = cad1 & " desde " & txtOpe(2).Text & " " & DtxtIns(2).Text
        End If

        If txtOpe(3).Text <> "" Then
            sql = sql & "hasope= """ & Trim(txtOpe(3).Text) & """|"
            cad1 = cad1 & "   hasta " & txtOpe(3).Text & " " & DtxtOpe(3).Text
        End If

        If txtOpe(2).Text <> "" Or txtOpe(3).Text <> "" Then sql = sql & "DNIs= """ & cad1 & """|"

        ' fechas de alta
        cad1 = "Fechas de Alta: "
        If Text3(2).Text <> "" Then
            sql = sql & "desfec= """ & Format(Text3(2).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(2).Text
        End If

        If Text3(3).Text <> "" Then
            sql = sql & "hasfec= """ & Format(Text3(3).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(3).Text
        End If

        If Text3(2).Text <> "" Or Text3(3).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"

    ' fechas de baja
        cad1 = "Fechas de Baja: "
        If Text3(27).Text <> "" Then
            sql = sql & "desfec2= """ & Format(Text3(27).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(27).Text
        End If

        If Text3(28).Text <> "" Then
            sql = sql & "hasfec2= """ & Format(Text3(28).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(28).Text
        End If

        If Text3(27).Text <> "" Or Text3(28).Text <> "" Then sql = sql & "FechaBaja= """ & cad1 & """|"

        
        ' numero de dosimetros
        cad1 = "Dosimetros: "
        If txtDos(0).Text <> "" Then
            sql = sql & "desdos= """ & Trim(txtDos(0).Text) & """|"
            cad1 = cad1 & " desde " & txtDos(0).Text
        End If

        If txtDos(1).Text <> "" Then
            sql = sql & "hasdos= """ & Trim(txtDos(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtDos(1).Text
        End If

        If txtDos(0).Text <> "" Or txtDos(1).Text <> "" Then sql = sql & "Dosimetros= """ & cad1 & """|"

        sql = sql & "tipo= " & Format(Tipo, "0") & "|"
        
        '0=alta
        '1=baja
        '2=ambos
        If OptOpe(8).Value Then sql = sql & "alta= 0|"
        If OptOpe(9).Value Then sql = sql & "alta= 1|"
        If OptOpe(7).Value Then sql = sql & "alta= 2|"

'        If Tipo <> 2 Then
            frmImprimir.Opcion = 4
'        Else
'            frmImprimir.Opcion = 5
'        End If
        
        If Sistema <> "" Then
          sql = sql & "desdeSist= """ & Sistema & """|hastaSist= """ & Sistema & """|"
        End If
        
        frmImprimir.NumeroParametros = 14
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.CampoOrden = -5
        frmImprimir.Show 'vbModal
     End If
End Sub

Private Sub CmdAceptarFactCalib_Click()
Dim I As Integer
Dim tabla As String
Dim Tipo As String
    If Not ComprobarFechas(4, 5) Then Exit Sub
    If Not ComprobarFechas(6, 7) Then Exit Sub
    'Hacemos el select y si tiene resultados mostramos los valores

    If Opcion = 7 Then
      tabla = "factcali4400"
    ElseIf Opcion = 8 Then
      tabla = "factcali6600"
    Else
      tabla = "factcalipana"
    End If
    Cad = " SELECT " & tabla & ".* from " & tabla & " WHERE 1 = 1 "
    If Text3(4).Text <> "" Then Cad = Cad & " AND f_inicio >= '" & Format(Text3(4).Text, FormatoFecha) & "' "
    If Text3(5).Text <> "" Then Cad = Cad & " AND f_inicio <= '" & Format(Text3(5).Text, FormatoFecha) & "' "
    If Text3(6).Text <> "" Then Cad = Cad & " AND f_fin >= '" & Format(Text3(6).Text, FormatoFecha) & "' "
    If Text3(7).Text <> "" Then Cad = Cad & " AND f_fin <= '" & Format(Text3(7).Text, FormatoFecha) & "' "
    For I = 0 To Option1.Count - 1
      If Option1(I).Value = True Then Tipo = Option1(I).Tag
    Next I
    If Tipo <> "" Then
      Cad = Cad & "and tipo = '" & Tipo & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Fecha Inicio: "
        If Text3(4).Text <> "" Then
            sql = sql & "desfe1= """ & Format(Text3(4).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(4).Text & " "
        End If

        If Text3(5).Text <> "" Then
            sql = sql & "hasfe1= """ & Format(Text3(5).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(5).Text & " "
        End If
        If Text3(4).Text <> "" Or Text3(5).Text <> "" Then sql = sql & "FechaInicio= """ & cad1 & """|"


        cad1 = "Fecha Finalización: "
        If Text3(6).Text <> "" Then
            sql = sql & "desfe2= """ & Format(Text3(6).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(6).Text & " "
        End If

        If Text3(7).Text <> "" Then
            sql = sql & "hasfe2= """ & Format(Text3(7).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(7).Text & " "
        End If

        If Text3(6).Text <> "" Or Text3(7).Text <> "" Then sql = sql & "FechaFinal= """ & cad1 & """|"
        
        If Tipo <> "" Then sql = sql & "desdeTipo= """ & Tipo & """|" & "hastaTipo= """ & Tipo & """|"
        
        If Opcion <> 31 Then
          frmImprimir.Opcion = Opcion
        Else
          frmImprimir.Opcion = 36
        End If
        
        frmImprimir.NumeroParametros = 6
        frmImprimir.FormulaSeleccion = ""
        frmImprimir.OtrosParametros = sql
        frmImprimir.CampoOrden = 3
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If


End Sub

Private Sub CmdAceptarLisDosisIns_Click()
'Dim sql As String
'Dim sql1 As String
''Dim cad2 As String
''Dim Rs As ADODB.Recordset
''Dim Rs1 As ADODB.Recordset
Dim Contador As Currency

Dim tabla As String
Dim fechaini As Date
Dim fechafin As Date
Dim fechaact As Date
Dim Sist As String
Dim Cad As String
Dim cad2 As String
Dim cad3 As String

Dim I As Integer
Dim Maximo As String
Dim Minimo As String
Dim rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

    ' Es obligatorio especificar rango de fechas.
    If Not (Text3(8).Text <> "" And Text3(9).Text <> "") Then
      MsgBox "Debe especificar un periodo válido para continuar.", vbOKOnly + vbExclamation, "¡Atención!"
      Exit Sub
    End If
    
    ' Comprobaciones previas.
    Screen.MousePointer = vbHourglass
    If Not ComprobarFechas(8, 9) Then Exit Sub
    If Not ComprobarEmpresas(8, 9) Then Exit Sub
    If Not ComprobarInstalaciones(4, 5) Then Exit Sub
         
    ' Sistema y tabla implicados.
    Sist = IIf(OptSist(0).Value, "H", "P")
    tabla = IIf(OptIns(2).Value, "dosiscuerpo", IIf(OptIns(1).Value, "dosisnohomog", "dosisarea"))
    
    ' Montamos las 2 select (la segunda nos dirá la fecha máxima y mínima
    ' de dosis... es para mejorar rendimiento)
    cad3 = " FROM dosimetros dosim," & tabla & " dosis"
    Cad = "SELECT distinct dosim.c_empresa,dosim.c_instalacion,dosim.n_reg_dosimetro,"
    Cad = Cad & "dosim.n_dosimetro,dosim.n_reg_dosimetro,dosim.mes_p_i,"
    Cad = Cad & "dosim.f_asig_dosimetro,dosim.f_retirada,dosim.dni_usuario" & cad3
    cad2 = "SELECT min(dosis.f_dosis) as minimo, max(dosis.f_dosis) as maximo " & cad3
    If Check1(5).Value = 1 Then
      cad3 = " WHERE dosim.f_retirada is null"
      Cad = Cad & ",tempnc" & cad3 & " AND dosim.n_dosimetro = tempnc.n_dosimetro AND codusu = " & vUsu.codigo
      cad2 = cad2 & cad3
    Else
      Cad = Cad & " WHERE 1=1"
      cad2 = cad2 & " WHERE 1=1"
    End If
    'cad3 = " AND dosim.sistema='" & Sist & "' AND dosim.n_reg_dosimetro = dosis.n_reg_dosimetro"
    cad3 = " AND dosim.n_reg_dosimetro = dosis.n_reg_dosimetro" ' (rafa) vrs 1.3.7
    cad3 = cad3 & " AND dosim.tipo_dosimetro = " & IIf(OptIns(2).Value, 0, IIf(OptIns(3).Value, 1, 2))
    
    ' Ahora las condiciones del listado.
    If Text3(8).Text <> "" Then
      cad3 = cad3 & " AND dosis.f_dosis >= '" & Format(Text3(8).Text, FormatoFecha) & "'"
      fechaini = CDate(Text3(8).Text)
    End If
    If Text3(9).Text <> "" Then
      cad3 = cad3 & " AND dosis.f_dosis <= '" & Format(Text3(9).Text, FormatoFecha) & "'"
      fechafin = CDate(Text3(9).Text)
    End If
    If txtEmp(8).Text <> "" Then cad3 = cad3 & " AND dosis.c_empresa >= '" & Trim(txtEmp(8).Text) & "'"
    If txtEmp(9).Text <> "" Then cad3 = cad3 & " AND dosis.c_empresa <= '" & Trim(txtEmp(9).Text) & "'"
    If txtIns(4).Text <> "" Then cad3 = cad3 & " AND dosis.c_instalacion >= '" & Trim(txtIns(4).Text) & "'"
    If txtIns(5).Text <> "" Then cad3 = cad3 & " AND dosis.c_instalacion <= '" & Trim(txtIns(5).Text) & "'"
    Cad = Cad & cad3
    cad2 = cad2 & cad3
    
    ' Recorremos el recordset, si hay algo que recorrer.
    Contador = 0
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
    If Not rs.EOF Then
      
      ' Borramos temporales.
      Cad = "delete from zdosisacum where codusu = " & vUsu.codigo
      conn.Execute Cad
      Cad = "delete from zdosisacumtot where codusu = " & vUsu.codigo
      conn.Execute Cad
      
      ' De momento, evitaremos esa "optimización", porque se salta los
      ' dosímetros no recibidos, al no existir dosis para los mismos.
      'Set Rs1 = New ADODB.Recordset
      'Rs1.Open cad2, Conn, adOpenKeyset, adLockPessimistic, adCmdText
      'Minimo = Rs1!Minimo
      'Maximo = Rs1!Maximo
      'Rs1.Close
      'Set Rs1 = Nothing
      Minimo = Text3(8).Text
      Maximo = Text3(9).Text
      While Not rs.EOF
        
        ' Establecemos límites de fechas a los de "vigencia" del dosímetro,
        ' ya que no tiene sentido ver las dosis del dosímetro en su periodo
        ' de inactividad.
        fechaini = CDate(Minimo)
        fechafin = CDate(Maximo)
        If IsDate(rs!f_asig_dosimetro) Then
          If fechaini < rs!f_asig_dosimetro Then fechaini = CDate(rs!f_asig_dosimetro)
        End If
        If IsDate(rs!f_retirada) Then
          If fechafin > rs!f_retirada Then fechafin = CDate(rs!f_retirada)
        End If
        
        ' Iniciamos el bucle.
        fechaact = fechaini
        While fechaact <= fechafin
          
          If ((Month(fechaact) And 1) = 1 And rs!mes_p_i = "I") Or ((Month(fechaact) And 1) <> 1 And rs!mes_p_i = "P") Then
            ' Insertamos la dosis correspondiente en la tabla temporal.
            InsertaDosisMensual tabla, Trim(rs!n_dosimetro), Trim(rs!n_reg_dosimetro), Trim(rs!c_empresa), Trim(rs!c_instalacion), fechaact, rs!dni_usuario
            Contador = Contador + 1
            
          Else
            ' Insertamos la dosis mensual del dosímetro del mes opuesto,
            ' si estamos cruzando con la tabla del archivo migrado.
            If Check1(5).Value = 1 Then
              Cad = "SELECT n_dosimetro,n_reg_dosimetro FROM dosimetros WHERE "
              Cad = Cad & "c_empresa='" & Trim(rs!c_empresa) & "' AND "
              Cad = Cad & "c_instalacion='" & Trim(rs!c_instalacion) & "' AND "
              Cad = Cad & "dni_usuario='" & Trim(rs!dni_usuario) & "' AND "
              Cad = Cad & "mes_p_i='" & IIf(rs!mes_p_i = "I", "P", "I") & "' AND "
              Cad = Cad & "f_retirada is null"
              Set Rs1 = New ADODB.Recordset
              Rs1.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
              If Not Rs1.EOF Then
                InsertaDosisMensual tabla, Trim(Rs1!n_dosimetro), Trim(Rs1!n_reg_dosimetro), Trim(rs!c_empresa), Trim(rs!c_instalacion), fechaact, rs!dni_usuario
                Contador = Contador + 1
              End If
              Rs1.Close
              Set Rs1 = Nothing
            End If
         
          End If
          
          fechaact = DateAdd("m", 1, fechaact)
          
        Wend
        rs.MoveNext
      Wend
    
    Else
      MsgBox "No hay datos para esos criterios.", vbOKOnly + vbExclamation, "Realizar Listado"
      Screen.MousePointer = vbDefault
      Exit Sub
    End If

    Screen.MousePointer = vbDefault

'''    If OptIns(2).Value = True Then
'''        Cad = " SELECT distinct dosiscuerpo.c_instalacion from dosiscuerpo, dosimetros, operarios"
'''        ' ### [DavidV] 18/04/2006: Para que pueda imprimir sólo migrados.
'''        If Check1(5).Value = 1 Then
'''          Cad = Cad & ", tempnc where dosimetros.n_dosimetro = tempnc.n_dosimetro and codusu = " & vUsu.codigo
'''        Else
'''          Cad = Cad & " where 1=1 "
'''        End If
'''        ' ### [DavidV] Hasta aquí.
'''        Cad = Cad & " and dosimetros.n_reg_dosimetro = dosiscuerpo.n_reg_dosimetro "
'''        Cad = Cad & " and dosimetros.tipo_dosimetro = 0 "
'''        Cad = Cad & " and operarios.f_baja is null "
'''        Cad = Cad & " and dosiscuerpo.dni_usuario = operarios.dni "
'''
'''        If Text3(8).Text <> "" Then Cad = Cad & " AND dosiscuerpo.f_dosis >= '" & Format(Text3(8).Text, FormatoFecha) & "' "
'''        If Text3(9).Text <> "" Then Cad = Cad & " AND dosiscuerpo.f_dosis <= '" & Format(Text3(9).Text, FormatoFecha) & "' "
'''        If txtEmp(8).Text <> "" Then Cad = Cad & " AND dosiscuerpo.c_empresa >= '" & Trim(txtEmp(8).Text) & "' "
'''        If txtEmp(9).Text <> "" Then Cad = Cad & " AND dosiscuerpo.c_empresa <= '" & Trim(txtEmp(9).Text) & "' "
'''        If txtIns(4).Text <> "" Then Cad = Cad & " AND dosiscuerpo.c_instalacion >= '" & Trim(txtIns(4).Text) & "' "
'''        If txtIns(5).Text <> "" Then Cad = Cad & " AND dosiscuerpo.c_instalacion <= '" & Trim(txtIns(5).Text) & "' "
'''
'''    ElseIf OptIns(3).Value Then
'''        Cad = " SELECT distinct dosisnohomog.c_instalacion from dosisnohomog, dosimetros, operarios"
'''        ' ### [DavidV] 18/04/2006: Para que pueda imprimir sólo migrados.
'''        If Check1(5).Value = 1 Then
'''          Cad = Cad & ", tempnc where dosimetros.n_dosimetro = tempnc.n_dosimetro and codusu = " & vUsu.codigo
'''        Else
'''          Cad = Cad & " where 1=1 "
'''        End If
'''        ' ### [DavidV] Hasta aquí.
'''        Cad = Cad & " and dosimetros.n_reg_dosimetro = dosisnohomog.n_reg_dosimetro "
'''        Cad = Cad & " and dosimetros.tipo_dosimetro = 1 "
'''        Cad = Cad & " and operarios.f_baja is null "
'''        Cad = Cad & " and dosisnohomog.dni_usuario = operarios.dni "
'''
'''        If Text3(8).Text <> "" Then Cad = Cad & " AND dosisnohomog.f_dosis >= '" & Format(Text3(8).Text, FormatoFecha) & "' "
'''        If Text3(9).Text <> "" Then Cad = Cad & " AND dosisnohomog.f_dosis <= '" & Format(Text3(9).Text, FormatoFecha) & "' "
'''        If txtEmp(8).Text <> "" Then Cad = Cad & " AND dosisnohomog.c_empresa >= '" & Trim(txtEmp(8).Text) & "' "
'''        If txtEmp(9).Text <> "" Then Cad = Cad & " AND dosisnohomog.c_empresa <= '" & Trim(txtEmp(9).Text) & "' "
'''        If txtIns(4).Text <> "" Then Cad = Cad & " AND dosisnohomog.c_instalacion >= '" & Trim(txtIns(4).Text) & "' "
'''        If txtIns(5).Text <> "" Then Cad = Cad & " AND dosisnohomog.c_instalacion <= '" & Trim(txtIns(5).Text) & "' "
'''    ElseIf OptIns(4).Value Then
'''        Cad = " SELECT distinct dosisarea.c_instalacion from dosisarea, dosimetros, instalaciones"
'''        ' ### [DavidV] 18/04/2006: Para que pueda imprimir sólo migrados.
'''        If Check1(5).Value = 1 Then
'''          Cad = Cad & ", tempnc where dosimetros.n_dosimetro = dosisarea.n_dosimetro and dosimetros.n_dosimetro = tempnc.n_dosimetro and codusu = " & vUsu.codigo
'''        Else
'''          Cad = Cad & " where 1=1 "
'''        End If
'''        ' ### [DavidV] Hasta aquí.
'''        Cad = Cad & " and dosimetros.n_reg_dosimetro = dosisarea.n_reg_dosimetro "
'''        Cad = Cad & " and dosimetros.tipo_dosimetro = 2 "
'''        Cad = Cad & " and instalaciones.f_baja is null "
'''        Cad = Cad & " and dosisarea.c_instalacion = instalaciones.c_instalacion "
'''
'''        If Text3(8).Text <> "" Then Cad = Cad & " AND dosisarea.f_dosis >= '" & Format(Text3(8).Text, FormatoFecha) & "' "
'''        If Text3(9).Text <> "" Then Cad = Cad & " AND dosisarea.f_dosis <= '" & Format(Text3(9).Text, FormatoFecha) & "' "
'''        If txtEmp(8).Text <> "" Then Cad = Cad & " AND dosisarea.c_empresa >= '" & Trim(txtEmp(8).Text) & "' "
'''        If txtEmp(9).Text <> "" Then Cad = Cad & " AND dosisarea.c_empresa <= '" & Trim(txtEmp(9).Text) & "' "
'''        If txtIns(4).Text <> "" Then Cad = Cad & " AND dosisarea.c_instalacion >= '" & Trim(txtIns(4).Text) & "' "
'''        If txtIns(5).Text <> "" Then Cad = Cad & " AND dosisarea.c_instalacion <= '" & Trim(txtIns(5).Text) & "' "
'''
'''    End If
'''
'''    If Check2.Value = 1 Then
'''
'''        ' hacemos el count
'''        If OptIns(2).Value = True Then
'''            cad2 = " SELECT count(distinct dosiscuerpo.c_instalacion) from dosiscuerpo, dosimetros, operarios"
'''            ' ### [DavidV] 18/04/2006: Para que pueda imprimir sólo migrados.
'''            If Check1(5).Value = 1 Then
'''              cad2 = cad2 & ", tempnc where dosimetros.n_dosimetro = tempnc.n_dosimetro and codusu = " & vUsu.codigo
'''            Else
'''              cad2 = cad2 & " where 1=1 "
'''            End If
'''            ' ### [DavidV] Hasta aquí.
'''            cad2 = cad2 & " and dosimetros.n_reg_dosimetro = dosiscuerpo.n_reg_dosimetro "
'''            cad2 = cad2 & " and dosimetros.tipo_dosimetro = 0 "
'''            cad2 = cad2 & " and operarios.f_baja is null "
'''            cad2 = cad2 & " and dosiscuerpo.dni_usuario = operarios.dni "
'''
'''            If Text3(8).Text <> "" Then cad2 = cad2 & " AND dosiscuerpo.f_dosis >= '" & Format(Text3(8).Text, FormatoFecha) & "' "
'''            If Text3(9).Text <> "" Then cad2 = cad2 & " AND dosiscuerpo.f_dosis <= '" & Format(Text3(9).Text, FormatoFecha) & "' "
'''            If txtEmp(8).Text <> "" Then cad2 = cad2 & " AND dosiscuerpo.c_empresa >= '" & Trim(txtEmp(8).Text) & "' "
'''            If txtEmp(9).Text <> "" Then cad2 = cad2 & " AND dosiscuerpo.c_empresa <= '" & Trim(txtEmp(9).Text) & "' "
'''            If txtIns(4).Text <> "" Then cad2 = cad2 & " AND dosiscuerpo.c_instalacion >= '" & Trim(txtIns(4).Text) & "' "
'''            If txtIns(5).Text <> "" Then cad2 = cad2 & " AND dosiscuerpo.c_instalacion <= '" & Trim(txtIns(5).Text) & "' "
'''
'''        ElseIf OptIns(3).Value Then
'''            cad2 = " SELECT count(distinct dosisnohomog.c_instalacion) from dosisnohomog, dosimetros, operarios"
'''            ' ### [DavidV] 18/04/2006: Para que pueda imprimir sólo migrados.
'''            If Check1(5).Value = 1 Then
'''              cad2 = cad2 & ", tempnc where dosimetros.n_dosimetro = tempnc.n_dosimetro and codusu = " & vUsu.codigo
'''            Else
'''              cad2 = cad2 & " where 1=1 "
'''            End If
'''            ' ### [DavidV] Hasta aquí.
'''            cad2 = cad2 & " and dosimetros.n_reg_dosimetro = dosisnohomog.n_reg_dosimetro "
'''            cad2 = cad2 & " and dosimetros.tipo_dosimetro = 1 "
'''            cad2 = cad2 & " and operarios.f_baja is null "
'''            cad2 = cad2 & " and dosisnohomog.dni_usuario = operarios.dni "
'''
'''            If Text3(8).Text <> "" Then cad2 = cad2 & " AND dosisnohomog.f_dosis >= '" & Format(Text3(8).Text, FormatoFecha) & "' "
'''            If Text3(9).Text <> "" Then cad2 = cad2 & " AND dosisnohomog.f_dosis <= '" & Format(Text3(9).Text, FormatoFecha) & "' "
'''            If txtEmp(8).Text <> "" Then cad2 = cad2 & " AND dosisnohomog.c_empresa >= '" & Trim(txtEmp(8).Text) & "' "
'''            If txtEmp(9).Text <> "" Then cad2 = cad2 & " AND dosisnohomog.c_empresa <= '" & Trim(txtEmp(9).Text) & "' "
'''            If txtIns(4).Text <> "" Then cad2 = cad2 & " AND dosisnohomog.c_instalacion >= '" & Trim(txtIns(4).Text) & "' "
'''            If txtIns(5).Text <> "" Then cad2 = cad2 & " AND dosisnohomog.c_instalacion <= '" & Trim(txtIns(5).Text) & "' "
'''        ElseIf OptIns(4).Value Then
'''            cad2 = " SELECT count(distinct dosisarea.c_instalacion) from dosisarea, instalaciones, dosimetros"
'''            ' ### [DavidV] 18/04/2006: Para que pueda imprimir sólo migrados.
'''            If Check1(5).Value = 1 Then
'''              cad2 = cad2 & ", tempnc where dosimetros.n_dosimetro = dosisarea.n_dosimetro and dosimetros.n_dosimetro = tempnc.n_dosimetro and codusu = " & vUsu.codigo
'''            Else
'''              cad2 = cad2 & " where 1=1 "
'''            End If
'''            ' ### [DavidV] Hasta aquí.
'''            cad2 = cad2 & " and dosimetros.n_reg_dosimetro = dosisarea.n_reg_dosimetro "
'''            cad2 = cad2 & " and dosimetros.tipo_dosimetro = 2 "
'''            cad2 = cad2 & " and instalaciones.f_baja is null "
'''            cad2 = cad2 & " and dosisarea.c_instalacion = instalaciones.c_instalacion "
'''
'''            If Text3(8).Text <> "" Then cad2 = cad2 & " AND dosisarea.f_dosis >= '" & Format(Text3(8).Text, FormatoFecha) & "' "
'''            If Text3(9).Text <> "" Then cad2 = cad2 & " AND dosisarea.f_dosis <= '" & Format(Text3(9).Text, FormatoFecha) & "' "
'''            If txtEmp(8).Text <> "" Then cad2 = cad2 & " AND dosisarea.c_empresa >= '" & Trim(txtEmp(8).Text) & "' "
'''            If txtEmp(9).Text <> "" Then cad2 = cad2 & " AND dosisarea.c_empresa <= '" & Trim(txtEmp(9).Text) & "' "
'''            If txtIns(4).Text <> "" Then cad2 = cad2 & " AND dosisarea.c_instalacion >= '" & Trim(txtIns(4).Text) & "' "
'''            If txtIns(5).Text <> "" Then cad2 = cad2 & " AND dosisarea.c_instalacion <= '" & Trim(txtIns(5).Text) & "' "
'''
'''        End If
'''        cad2 = cad2 & " and dosimetros.c_instalacion in (select c_instalacion from instalaciones where f_baja is null)"
'''        Set Rs1 = New ADODB.Recordset
'''        Rs1.Open cad2, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'''
'''        Contador = 0
'''        If Not Rs1.EOF Then Contador = Rs1.Fields(0).Value
'''
'''    End If


'''    Set Rs = New ADODB.Recordset
'''    Rs.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'''    If Rs.EOF Then
'''        'NO hay registros a mostrar
'''        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
'''        Screen.MousePointer = vbDefault
'''    Else
    
    
    ' Miles y miles de cambios... 23/01/2006 [DV]
    If Check2.Value = 0 Then ' caso de no guardarlo en pdf

        'Mostramos el frame de resultados
        Cad = ""
        cad2 = "Empresas: "
        If txtEmp(8).Text <> "" Then
            Cad = Cad & "desemp= """ & Trim(txtEmp(8).Text) & """|"
            cad2 = cad2 & " desde " & txtEmp(8).Text & " " & DtxtEmp(8).Text
        End If

        If txtEmp(9).Text <> "" Then
            Cad = Cad & "hasemp= """ & Trim(txtEmp(9).Text) & """|"
            cad2 = cad2 & "   hasta " & txtEmp(9).Text & " " & DtxtEmp(9).Text
        End If

        If txtEmp(8).Text <> "" Or txtEmp(9).Text <> "" Then Cad = Cad & "Empresas= """ & cad2 & """|"

        cad2 = "Instalaciones: "
        If txtIns(4).Text <> "" Then
            Cad = Cad & "desins= """ & Trim(txtIns(4).Text) & """|"
            cad2 = cad2 & " desde " & txtIns(4).Text & " " & DtxtIns(4).Text
        End If

        If txtIns(5).Text <> "" Then
            Cad = Cad & "hasins= """ & Trim(txtIns(5).Text) & """|"
            cad2 = cad2 & "   hasta " & txtIns(5).Text & " " & DtxtIns(5).Text
        End If

        If txtIns(4).Text <> "" Or txtIns(5).Text <> "" Then Cad = Cad & "Instalaciones= """ & cad2 & """|"

        cad2 = "Fecha de Dosis: "
        If Text3(8).Text <> "" Then
            Cad = Cad & "desfec= """ & Format(Text3(8).Text, FormatoFecha) & """|"
            cad2 = cad2 & " desde " & Trim(Text3(8).Text) & " "
        End If

        If Text3(9).Text <> "" Then
            Cad = Cad & "hasfec= """ & Format(Text3(9).Text, FormatoFecha) & """|"
            cad2 = cad2 & "   hasta " & Trim(Text3(9).Text)
        End If

        If Text3(9).Text <> "" Or Text3(9).Text <> "" Then Cad = Cad & "FechaAlta= """ & cad2 & """|"

        If Check1(5).Value = 0 Then
          If OptIns(2).Value Then frmImprimir.Opcion = 9
          If OptIns(3).Value Then frmImprimir.Opcion = 10
          If OptIns(4).Value Then frmImprimir.Opcion = 11
          Cad = Cad & "codus= " & vUsu.codigo & "|"
          frmImprimir.NumeroParametros = 8
        Else
          If OptIns(2).Value Then frmImprimir.Opcion = 41
          If OptIns(3).Value Then frmImprimir.Opcion = 42
          If OptIns(4).Value Then frmImprimir.Opcion = 43
          Cad = Cad & "codus= " & vUsu.codigo & "|"
          Cad = Cad & "migrado= ""si""|"
          frmImprimir.NumeroParametros = 9
        End If
        
        frmImprimir.FormulaSeleccion = Cad
        frmImprimir.OtrosParametros = Cad
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
        
     Else ' lo guardamos en formato pdf
        
        Screen.MousePointer = vbHourglass
        
        Cad = ""
        cad2 = "Fecha de Dosis: "
        If Text3(8).Text <> "" Then
            Cad = Cad & "desfec= """ & Format(Text3(8).Text, FormatoFecha) & """|"
            cad2 = cad2 & " desde " & Trim(Text3(8).Text) & " "
        End If

        If Text3(9).Text <> "" Then
            Cad = Cad & "hasfec= """ & Format(Text3(9).Text, FormatoFecha) & """|"
            cad2 = cad2 & "   hasta " & Trim(Text3(9).Text)
        End If

        If Text3(9).Text <> "" Or Text3(9).Text <> "" Then Cad = Cad & "FechaAlta= """ & cad2 & """|"
                
        NumRegElim = 0
        
        Pb5.max = 100
        Pb5.Visible = True
        Pb5.Value = 0
        I = 0
        If Not rs Is Nothing Then rs.MoveFirst
        
        Do While Not rs.EOF
            Me.Refresh
            espera 0.5
            
            cad2 = "desins= """ & Trim(rs!c_instalacion) & """|"
            cad2 = cad2 & "hasins= """ & Trim(rs!c_instalacion) & """|"
            cad2 = cad2 & Cad
                    
            With frmVisReport
            
            
        If OptIns(2).Value Then .Informe = App.Path & "\informes\DosisHomoInstalacionNew.rpt"
        If OptIns(3).Value Then .Informe = App.Path & "\informes\DosisNoHomoInstalacion.rpt"
        If OptIns(4).Value Then .Informe = App.Path & "\informes\DosisAreaInstalacionNew.rpt"
            
            '    .Informe = App.Path & "\informes\DosisHomoInstalacion.rpt"
                .OtrosParametros = cad2
                .NumeroParametros = 4
'                    sql = "{ado.codusu}=" & vUsu.codigo & " AND {ado.nif}= """ & RS.Fields(0) & """"
                .FormulaSeleccion = cad2
                .ExportarPDF = True
                CadenaDesdeOtroForm = "GENERANDO"
                Load frmVisReport
                Unload frmVisReport
                
                If CadenaDesdeOtroForm = "OK" Then
                    Me.Refresh
                    espera 0.5
                    I = I + 1
                    Pb5.Value = CInt((100 * I) / Contador)
                    Pb5.Refresh
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    cad3 = "Dosis-" & Trim(rs!c_instalacion) & ".pdf"
                    
                    If Dir(App.Path & "\temp", vbDirectory) = "" Then
                        MkDir App.Path & "\temp"
                    End If
                    
                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Replace(cad3, "/", "-")
    
                Else
                   If MsgBox("¿Desea cancelar la exportación?", vbYesNo + vbQuestion, "Exportación.") = vbYes Then
                    Exit Do
                  End If
                End If
            End With
            
            rs.MoveNext
        Loop
        
        Pb5.Value = Pb5.max
        Pb5.Refresh
        
   
'        If Cont > 0 Then
'
'             espera 0.5
'
'             sql = "Carta de reclamación de dosimetros no recibidos"
'             frmEMail.Opcion = 3
'             frmEMail.MisDatos = sql
'             frmEMail.Show vbModal
'
'        End If
        Screen.MousePointer = vbDefault
                
        Me.Visible = True
        Me.Refresh
        If CadenaDesdeOtroForm <> "OK" Then
          MsgBox "Proceso cancelado por el usuario", vbExclamation, "¡Atención!"
        Else
          MsgBox "Proceso finalizado", vbExclamation, "¡Atención!"
        End If
    End If
        
' End If

    rs.Close
    Set rs = Nothing
    Pb5.Visible = False
    
End Sub

Private Sub InsertaDosisMensual(tabla As String, n_dos As String, n_reg As String, c_emp As String, c_ins As String, fechaact As Date, dni_usuario As String)
Dim Cad As String
Dim ano As Integer
Dim Mes As Integer
Dim super As Currency
Dim profun As Currency
Dim situ As Integer
Dim rs As ADODB.Recordset
  
  ano = Year(fechaact)
  Mes = Month(fechaact)
         
  Cad = "SELECT SUM(dosis_superf),SUM(dosis_profunda) FROM " & tabla & " WHERE "
  Cad = Cad & "n_dosimetro = '" & n_dos & "' AND n_reg_dosimetro = '" & n_reg & "' AND "
  Cad = Cad & "c_empresa = '" & c_emp & "' AND c_instalacion = '" & c_ins & "' AND "
  Cad = Cad & "month(f_dosis) = " & Mes & " AND year(f_dosis) = " & Format(ano, "0000")
  'Cad = Cad & " GROUP BY dni_usuario"
  Set rs = New ADODB.Recordset
  rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            
  ' Valores por defecto.
  super = 0
  profun = 0
  situ = 2
            
  If Not rs.EOF Then
    If Not IsNull(rs.Fields(0).Value) Or Not IsNull(rs.Fields(1).Value) Then
      super = Val(TransformaComasPuntos(rs.Fields(0).Value & ""))
      profun = Val(TransformaComasPuntos(rs.Fields(1).Value & ""))
      situ = 0
    End If
  End If
            
  ' Insertamos la línea.
  Cad = "INSERT INTO zdosisacum (codusu, c_empresa, c_instalacion, dni_usuario, "
  Cad = Cad & "mes, ano, n_dosimetro, dosissuper, dosisprofu, situ) VALUES ("
  Cad = Cad & vUsu.codigo & ",'" & c_emp & "','" & c_ins & "','" & dni_usuario & "',"
  Cad = Cad & Mes & "," & ano & ",'" & n_dos & "'," & TransformaComasPuntos(super & "") & "," & TransformaComasPuntos(profun & "")
  Cad = Cad & "," & Format(situ, "0") & ")"
  conn.Execute Cad
  rs.Close
  Set rs = Nothing
  
End Sub

Private Sub CmdAceptarLisEmp_Click()
    If Not ComprobarEmpresas(0, 1) Then Exit Sub
    'Hacemos el select y si tiene resultados mostramos los valores


    Cad = " SELECT empresas.* from empresas WHERE 1 = 1 "
    If Combo1.ListIndex = 0 Then Cad = Cad & " and c_tipo = 0 "
    If Combo1.ListIndex = 1 Then Cad = Cad & " and c_tipo = 1 "
    
    If txtEmp(0).Text <> "" Then Cad = Cad & " AND c_empresa >= '" & Trim(txtEmp(0).Text) & "' "
    If txtEmp(1).Text <> "" Then Cad = Cad & " AND c_empresa <= '" & Trim(txtEmp(1).Text) & "' "

    If OptEmp(2).Value Then Cad = Cad & " AND f_baja is null "
    If OptEmp(3).Value Then Cad = Cad & " AND not (f_baja is null) "
    
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Empresas: "
        If txtEmp(0).Text <> "" Then
            sql = sql & "desemp= """ & Trim(txtEmp(0).Text) & """|"
            cad1 = cad1 & " desde " & txtEmp(0).Text & " " & DtxtEmp(0).Text
        End If

        If txtEmp(1).Text <> "" Then
            sql = sql & "hasemp= """ & Trim(txtEmp(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtEmp(1).Text & " " & DtxtEmp(1).Text
        End If

        If txtEmp(0).Text <> "" Or txtEmp(1).Text <> "" Then sql = sql & "Empresas= """ & cad1 & """|"
        
        sql = sql & "Tipo= " & Combo1.ListIndex & "|"

        If OptEmp(2).Value Then sql = sql & "tipo1= 0|"
        If OptEmp(3).Value Then sql = sql & "tipo1= 1|"
        If OptEmp(4).Value Then sql = sql & "tipo1= 2|"

        If OptEmp(0).Value = True Then
            sql = sql & "orden= ""Por Código""|"
            frmImprimir.CampoOrden = 1
        End If
        
        If OptEmp(1).Value = True Then
            sql = sql & "orden= ""Alfabético""|"
            frmImprimir.CampoOrden = 5
        End If
        

        If Opcion = 1 Then
            frmImprimir.Opcion = Opcion ' listado de empresas
        Else
            frmImprimir.Opcion = 25 ' listado de etiquetas
        End If
        
        frmImprimir.NumeroParametros = 4
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarLisIns_Click()
    
    If Not ComprobarEmpresas(2, 3) Then Exit Sub
    If Not ComprobarInstalaciones(0, 1) Then Exit Sub
    If Not ComprobarFechas(29, 30) Then Exit Sub
    If Not ComprobarFechas(31, 32) Then Exit Sub
    
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT instalaciones.* from instalaciones WHERE 1 = 1 "
    If Combo2.ListIndex <> 2 Then Cad = Cad & " and c_tipo = " & Combo2.ListIndex
    If txtEmp(2).Text <> "" Then Cad = Cad & " AND c_empresa >= '" & Trim(txtEmp(2).Text) & "' "
    If txtEmp(3).Text <> "" Then Cad = Cad & " AND c_empresa <= '" & Trim(txtEmp(3).Text) & "' "
    If txtIns(0).Text <> "" Then Cad = Cad & " AND c_instalacion >= '" & Trim(txtIns(0).Text) & "' "
    If txtIns(1).Text <> "" Then Cad = Cad & " AND c_instalacion <= '" & Trim(txtIns(1).Text) & "' "
    If Text3(29).Text <> "" Then Cad = Cad & " AND f_alta >= '" & Format(Text3(29).Text, FormatoFecha) & "' "
    If Text3(30).Text <> "" Then Cad = Cad & " AND f_alta <= '" & Format(Text3(30).Text, FormatoFecha) & "' "
    If Text3(31).Text <> "" Then Cad = Cad & " AND f_baja >= '" & Format(Text3(31).Text, FormatoFecha) & "' "
    If Text3(32).Text <> "" Then Cad = Cad & " AND f_baja <= '" & Format(Text3(32).Text, FormatoFecha) & "' "
    If OptIns(19).Value Then Cad = Cad & " AND f_baja is null"
    If OptIns(20).Value Then Cad = Cad & " AND not (f_baja is null)"

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Empresas: "
        If txtEmp(2).Text <> "" Then
            sql = sql & "desemp= """ & Trim(txtEmp(2).Text) & """|"
            cad1 = cad1 & " desde " & txtEmp(2).Text & " " & DtxtEmp(2).Text
        End If

        If txtEmp(3).Text <> "" Then
            sql = sql & "hasemp= """ & Trim(txtEmp(3).Text) & """|"
            cad1 = cad1 & "   hasta " & txtEmp(3).Text & " " & DtxtEmp(3).Text
        End If

        If txtEmp(2).Text <> "" Or txtEmp(3).Text <> "" Then sql = sql & "Empresas= """ & cad1 & """|"

        cad1 = "Instalaciones: "
        If txtIns(0).Text <> "" Then
            sql = sql & "desins= """ & Trim(txtIns(0).Text) & """|"
            cad1 = cad1 & " desde " & txtIns(0).Text & " " & DtxtIns(0).Text
        End If

        If txtIns(1).Text <> "" Then
            sql = sql & "hasins= """ & Trim(txtIns(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtIns(1).Text & " " & DtxtIns(1).Text
        End If

        If txtIns(0).Text <> "" Or txtIns(1).Text <> "" Then sql = sql & "Instalaciones= """ & cad1 & """|"
        
        ' fechas de alta
        cad1 = "Fechas de Alta: "
        If Text3(29).Text <> "" Then
            sql = sql & "desfec= """ & Format(Text3(29).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(29).Text
        End If

        If Text3(30).Text <> "" Then
            sql = sql & "hasfec= """ & Format(Text3(30).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(30).Text
        End If

        If Text3(29).Text <> "" Or Text3(30).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"

        ' fechas de baja
        cad1 = "Fechas de Baja: "
        If Text3(31).Text <> "" Then
            sql = sql & "desfec2= """ & Format(Text3(31).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(31).Text
        End If

        If Text3(32).Text <> "" Then
            sql = sql & "hasfec2= """ & Format(Text3(32).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(32).Text
        End If

        If Text3(31).Text <> "" Or Text3(32).Text <> "" Then sql = sql & "FechaBaja= """ & cad1 & """|"
        
        sql = sql & "Tipo= " & Combo2.ListIndex & "|"
 
        'tipo
        If OptIns(19).Value Then sql = sql & "tipo1= 0|"
        If OptIns(20).Value Then sql = sql & "tipo1= 1|"
        If OptIns(21).Value Then sql = sql & "tipo1= 2|"

        If OptIns(0).Value = True Then
            sql = sql & "orden= ""Por Código""|"
            frmImprimir.CampoOrden = 2
        End If
        If OptIns(1).Value = True Then
            sql = sql & "orden= ""Alfabético""|"
            frmImprimir.CampoOrden = 5
        End If

        If Opcion = 2 Then
            frmImprimir.Opcion = 2
        Else
            frmImprimir.Opcion = 26
        End If
        
        frmImprimir.NumeroParametros = 8
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If


End Sub

Private Sub CmdAceptarLisOpe_Click()
    
    Screen.MousePointer = vbHourglass
    
    If Not ComprobarEmpresas(4, 5) Then Exit Sub
    If Not ComprobarInstalaciones(2, 3) Then Exit Sub
    If Not ComprobarOperarios(0, 1) Then Exit Sub
    If Not ComprobarFechas(0, 1) Then Exit Sub

    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT operainstala.* from operainstala WHERE 1 = 1 "
    If txtEmp(4).Text <> "" Then Cad = Cad & " AND c_empresa >= '" & Trim(txtEmp(4).Text) & "' "
    If txtEmp(5).Text <> "" Then Cad = Cad & " AND c_empresa <= '" & Trim(txtEmp(5).Text) & "' "
    If txtIns(2).Text <> "" Then Cad = Cad & " AND c_instalacion >= '" & Trim(txtIns(2).Text) & "' "
    If txtIns(3).Text <> "" Then Cad = Cad & " AND c_instalacion <= '" & Trim(txtIns(3).Text) & "' "
    If txtOpe(0).Text <> "" Then Cad = Cad & " AND dni >= '" & Trim(txtOpe(0).Text) & "' "
    If txtOpe(1).Text <> "" Then Cad = Cad & " AND dni <= '" & Trim(txtOpe(1).Text) & "' "
    If Text3(0).Text <> "" Then Cad = Cad & " AND f_alta >= '" & Format(Text3(0).Text, FormatoFecha) & "' "
    If Text3(1).Text <> "" Then Cad = Cad & " AND f_alta <= '" & Format(Text3(1).Text, FormatoFecha) & "' "
    If OptOpe(4).Value Then Cad = Cad & " AND f_baja IS NULL "
    If OptOpe(5).Value Then Cad = Cad & " AND f_baja IS NOT NULL "
    
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        Screen.MousePointer = vbDefault
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        CargarOperarios
        
        sql = ""
        sql = "usu= " & vUsu.codigo & "|"
        cad1 = "Empresas: "
        If txtEmp(4).Text <> "" Then
            sql = sql & "desemp= """ & Trim(txtEmp(4).Text) & """|"
            cad1 = cad1 & " desde " & txtEmp(4).Text & " " & DtxtEmp(4).Text
        End If

        If txtEmp(5).Text <> "" Then
            sql = sql & "hasemp= """ & Trim(txtEmp(5).Text) & """|"
            cad1 = cad1 & "   hasta " & txtEmp(5).Text & " " & DtxtEmp(5).Text
        End If

        If txtEmp(4).Text <> "" Or txtEmp(5).Text <> "" Then sql = sql & "Empresas= """ & cad1 & """|"

        ' codigo de instalaciones
        cad1 = "Instalaciones: "
        If txtIns(2).Text <> "" Then
            sql = sql & "desins= """ & Trim(txtIns(2).Text) & """|"
            cad1 = cad1 & " desde " & txtIns(2).Text & " " & DtxtIns(2).Text
        End If

        If txtIns(3).Text <> "" Then
            sql = sql & "hasins= """ & Trim(txtIns(3).Text) & """|"
            cad1 = cad1 & "   hasta " & txtIns(3).Text & " " & DtxtIns(3).Text
        End If

        If txtIns(2).Text <> "" Or txtIns(3).Text <> "" Then sql = sql & "Instalaciones= """ & cad1 & """|"

        ' dnis de operarios
        cad1 = "Operarios: "
        If txtOpe(0).Text <> "" Then
            sql = sql & "desope= """ & Trim(txtOpe(0).Text) & """|"
            cad1 = cad1 & " desde " & txtOpe(0).Text & " " & DtxtIns(0).Text
        End If

        If txtOpe(1).Text <> "" Then
            sql = sql & "hasope= """ & Trim(txtOpe(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtOpe(1).Text & " " & DtxtOpe(1).Text
        End If

        If txtOpe(0).Text <> "" Or txtOpe(1).Text <> "" Then sql = sql & "DNIs= """ & cad1 & """|"

        ' fechas de alta
        cad1 = "Fechas de Alta: "
        If Text3(0).Text <> "" Then
            sql = sql & "desfec= """ & Format(Text3(0).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(0).Text
        End If

        If Text3(1).Text <> "" Then
            sql = sql & "hasfec= """ & Format(Text3(1).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(1).Text
        End If

        If Text3(0).Text <> "" Or Text3(1).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"



        If OptOpe(2).Value = True Then
            sql = sql & "campogrupo= {operainstala.dni}|"
            frmImprimir.CampoOrden = -1
        End If
        If OptOpe(3).Value = True Then
            frmImprimir.CampoOrden = -2
            sql = sql & "campogrupo= {operainstala.c_empresa}|"
        End If


        If OptOpe(0).Value = True Then
            sql = sql & "tipo= 1|"
        End If
        If OptOpe(1).Value = True Then
            sql = sql & "tipo= 0|"
        End If

        If OptOpe(4).Value Then
            sql = sql & "tipo1= 0|"
        End If
        If OptOpe(5).Value Then
            sql = sql & "tipo1= 1|"
        End If
        If OptOpe(6).Value Then
            sql = sql & "tipo1= 2|"
        End If


        If Opcion = 3 Then frmImprimir.Opcion = 3
        If Opcion = 26 Then frmImprimir.Opcion = 27
'        If Opcion = 30 Then frmImprimir.Opcion = 33
        
        Screen.MousePointer = vbDefault
        
        frmImprimir.NumeroParametros = 20
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarListDosisColec_Click()
Dim Tipo As Byte
Dim Titol As String
Dim TipoDosimetria As Byte

On Error GoTo eListadoDosisColectiva

    If Text3(10).Text <> "" Then
        Text3(11).Text = CargarFechaHasta(CDate(Text3(10).Text), 1)
    Else
        MsgBox "Ha de introducir la fecha de inicio.", vbExclamation, "¡Atención!"
        Exit Sub
    End If
    
    If Not ComprobarFechas(10, 11) Then Exit Sub

    'Hacemos el select y si tiene resultados mostramos los valores
    If OptIns(18).Value Then
        TipoDosimetria = 0
    ElseIf OptIns(17).Value Then
            TipoDosimetria = 1
    End If

    Screen.MousePointer = vbHourglass

    Titol = "Dosis Colectivas "
    If OptIns(8).Value Then
        Tipo = 0
        Titol = Titol & "Superficial "
    ElseIf OptIns(9).Value Then
        Tipo = 1
        Titol = Titol & "Profunda "
    End If
    
    If OptIns(7).Value Then Titol = Titol & "Mensual"
    If OptIns(6).Value Then Titol = Titol & "Semestral"
    If OptIns(5).Value Then Titol = Titol & "Anual"
    
    If GenerarListadoCSN(vUsu.codigo, CDate(Text3(10).Text), CDate(Text3(11).Text), Tipo, TipoDosimetria) Then
        sql = "emp= """ & Trim(vParam.NombreEmpresa) & """|"
        sql = sql & "titol= """ & Trim(Titol) & """|"
        sql = sql & "usu= " & vUsu.codigo & "|"
        sql = sql & "tipo= " & Tipo & "|"
        sql = sql & "tipodosimetria= " & TipoDosimetria & "|"
        
        cad1 = "Rango de Fechas: "
        If Text3(10).Text <> "" Then
            sql = sql & "desfec= """ & Format(Text3(10).Text, "dd/mm/yyyy") & """|"
            cad1 = cad1 & " desde " & Text3(10).Text
        End If

        If Text3(11).Text <> "" Then
            sql = sql & "hasfec= """ & Format(Text3(11).Text, "dd/mm/yyyy") & """|"
            cad1 = cad1 & "   hasta " & Text3(11).Text
        End If

        If Text3(10).Text <> "" Or Text3(11).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"

        
        frmImprimir.Opcion = 12
        frmImprimir.Titulo = Titol
        frmImprimir.NumeroParametros = 5
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
    Else
        MsgBox "No existen datos entre esos límites. Reintroduzca.", vbExclamation, "¡Atención!"
        Exit Sub
    End If
    
    
eListadoDosisColectiva:
    If Err.Number <> 0 Then MuestraError Err.Number, "Error en la carga de datos"

    Screen.MousePointer = vbDefault

End Sub

Private Sub CmdAceptarListDosisNHomOpe_Click()
    
    Screen.MousePointer = vbHourglass
    
    
    If Not ComprobarFechas(16, 17) Then Exit Sub
    If Not ComprobarInstalaciones(8, 9) Then Exit Sub
    If Not ComprobarOperarios(4, 5) Then Exit Sub
    
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT dosisnohomog.* from dosisnohomog WHERE 1 = 1 "
    
    If Text3(16).Text <> "" Then Cad = Cad & " AND f_dosis >= '" & Format(Text3(16).Text, FormatoFecha) & "' "
    If Text3(17).Text <> "" Then Cad = Cad & " AND f_dosis <= '" & Format(Text3(17).Text, FormatoFecha) & "' "
    If txtEmp(21).Text <> "" Then Cad = Cad & " AND c_empresa >= '" & Trim(txtEmp(21).Text) & "' "
    If txtEmp(20).Text <> "" Then Cad = Cad & " AND c_empresa <= '" & Trim(txtEmp(20).Text) & "' "
    If txtIns(8).Text <> "" Then Cad = Cad & " AND c_instalacion >= '" & Trim(txtIns(8).Text) & "' "
    If txtIns(9).Text <> "" Then Cad = Cad & " AND c_instalacion <= '" & Trim(txtIns(9).Text) & "' "
    If txtOpe(4).Text <> "" Then Cad = Cad & " AND dni_usuario >= '" & Trim(txtOpe(4).Text) & "' "
    If txtOpe(5).Text <> "" Then Cad = Cad & " AND dni_usuario <= '" & Trim(txtOpe(5).Text) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
        Screen.MousePointer = vbDefault
    
    Else
        'Mostramos el frame de resultados
        
        CargarEmpresas
        CargarOperarios
        CargarInstalaciones
        
        
        sql = ""
        sql = "usu= " & vUsu.codigo & "|"
        cad1 = "Operarios: "
        If txtOpe(4).Text <> "" Then
            sql = sql & "desope= """ & Trim(txtOpe(4).Text) & """|"
            cad1 = cad1 & " desde " & txtOpe(4).Text & " " & DtxtOpe(4).Text
        End If

        If txtOpe(5).Text <> "" Then
            sql = sql & "hasope= """ & Trim(txtOpe(5).Text) & """|"
            cad1 = cad1 & "   hasta " & txtOpe(5).Text & " " & DtxtOpe(5).Text
        End If

        If txtOpe(4).Text <> "" Or txtOpe(5).Text <> "" Then sql = sql & "Operarios= """ & cad1 & """|"

        cad1 = "Empresas: "
        If txtEmp(21).Text <> "" Then
            sql = sql & "desemp= """ & Trim(txtEmp(21).Text) & """|"
            cad1 = cad1 & " desde " & txtEmp(21).Text & " " & DtxtEmp(21).Text
        End If

        If txtEmp(20).Text <> "" Then
            sql = sql & "hasemp= """ & Trim(txtEmp(20).Text) & """|"
            cad1 = cad1 & "   hasta " & txtEmp(20).Text & " " & DtxtEmp(20).Text
        End If

        If txtEmp(21).Text <> "" Or txtEmp(20).Text <> "" Then sql = sql & "Empresas= """ & cad1 & """|"

        cad1 = "Instalaciones: "
        If txtIns(8).Text <> "" Then
            sql = sql & "desins= """ & Trim(txtIns(8).Text) & """|"
            cad1 = cad1 & " desde " & txtIns(8).Text & " " & DtxtIns(8).Text
        End If

        If txtIns(9).Text <> "" Then
            sql = sql & "hasins= """ & Trim(txtIns(9).Text) & """|"
            cad1 = cad1 & "   hasta " & txtIns(9).Text & " " & DtxtIns(9).Text
        End If

        If txtIns(8).Text <> "" Or txtIns(9).Text <> "" Then sql = sql & "Instalaciones= """ & cad1 & """|"

        cad1 = "Fecha de Dosis: "
        If Text3(16).Text <> "" Then
            sql = sql & "desfec= """ & Format(Text3(16).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Trim(Text3(16).Text) & " "
        End If

        If Text3(17).Text <> "" Then
            sql = sql & "hasfec= """ & Format(Text3(17).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Trim(Text3(17).Text)
        End If

        If Text3(16).Text <> "" Or Text3(17).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"

        Screen.MousePointer = vbDefault

        frmImprimir.Opcion = 20
        frmImprimir.NumeroParametros = 9
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If


End Sub

Private Sub CmdAceptarListDosisOpeAcum12_Click()
Dim rs As ADODB.Recordset
Dim rL As ADODB.Recordset
Dim rf As ADODB.Recordset
Dim Rg As ADODB.Recordset

Dim Mesa As Integer
Dim Anoa As Integer
Dim ano As Integer
Dim fec As String
Dim sql1 As String
Dim sql2 As String
Dim situ As Integer
Dim fecha As String
Dim fdesde As String
Dim fdesde1 As Date
Dim fhasta As String
Dim fhasta1 As Date
Dim Encontrado As Boolean
Dim Valor1 As Currency
Dim valor2 As Currency
Dim sql3 As String
Dim f_mes As Integer
Dim f_ano As Integer
Dim Sist As String

  Screen.MousePointer = vbHourglass
  Sist = IIf(OptSist(2).Value, "H", "P")
  If IsNumeric(Text4(0).Text) Then
    ano = CInt(Text4(0).Text)
  Else
    ano = -1
  End If
    
  If Not (ano <= Year(Now) And ano >= 0) Then
    MsgBox "Año incorrecto. Reintroduzca.", vbExclamation, "¡Error!"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
        
  If Not ComprobarEmpresas(12, 13) Then Exit Sub
  If Not ComprobarOperarios(6, 7) Then Exit Sub
    
  'Hacemos el select y si tiene resultados mostramos los valores


    'DAVID GANDUL###
    'Si el dosimetro ha sido dado de alta este periodo SOLO apareceran los datos del mes
    'a partir del cual el dosimetro esta dado de alta
    ' 1er paso. Pedir FECHA ASIGNACION DOSIMETRO
    ' 2º paso. Variable para el IF  sea legible
    Dim AsignacionDosimetroFechaCorrecta As Boolean
    Dim FechaBaja As Date
    
  ' NUEVA HISTORIA para evitar los (campo1,campo2,campo3) IN (SELECT campo1,campo2...)
  ' Afortunadamente, no todo es malo... esto ha mejorado la velocidad de la consulta.
  Cad = "SELECT DISTINCT c1.c_empresa,c1.c_instalacion,c1.dni_usuario,c1.n_dosimetro,"
  Cad = Cad & "c1.mes_p_i,c1.f_asig_dosimetro,c1.f_retirada,c1.n_reg_dosimetro FROM "
  Cad = Cad & "(SELECT DISTINCT dosimetros.c_empresa,dosimetros.c_instalacion,dosimetros.dni_usuario,"
  Cad = Cad & "dosimetros.n_dosimetro,dosimetros.mes_p_i,f_asig_dosimetro,f_retirada,"
  Cad = Cad & "dosimetros.n_reg_dosimetro FROM operarios"
  
  ' Solo el fichero migrado
  If Check1(1).Value = 1 Then
    Cad = Cad & " INNER JOIN "
    Cad = Cad & "(SELECT DISTINCT dosimetros.dni_usuario dni FROM dosimetros,tempnc WHERE "
    Cad = Cad & "dosimetros.n_dosimetro=tempnc.n_dosimetro AND dosimetros.f_retirada IS NULL AND codusu="
    Cad = Cad & vUsu.codigo & ") t1 USING(dni),dosimetros INNER JOIN "
    Cad = Cad & "(SELECT DISTINCT dni_usuario,c_empresa,c_instalacion FROM dosimetros,tempnc "
    Cad = Cad & "WHERE dosimetros.n_dosimetro=tempnc.n_dosimetro AND dosimetros.f_retirada IS NULL AND codusu="
    Cad = Cad & vUsu.codigo & ") t2 USING(dni_usuario,c_empresa,c_instalacion)"
  Else
    Cad = Cad & ",dosimetros"
  End If
  
  Cad = Cad & " WHERE operarios.semigracsn=1 AND operarios.dni=dosimetros.dni_usuario AND "
  Cad = Cad & "dosimetros.tipo_dosimetro=0 AND dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%' "
  'Cad = Cad & "AND dosimetros.sistema='" & Sist & "'" ' (rafa VRS 1.3.6)
  If Check1(0).Value = 1 Then Cad = Cad & " AND operarios.f_baja IS NULL"
  If txtEmp(12).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa>='" & Trim(txtEmp(12).Text) & "'"
  If txtEmp(13).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa<='" & Trim(txtEmp(13).Text) & "'"
  If txtOpe(6).Text <> "" Then Cad = Cad & " AND dosimetros.dni_usuario>='" & Trim(txtOpe(6).Text) & "'"
  If txtOpe(7).Text <> "" Then Cad = Cad & " AND dosimetros.dni_usuario<='" & Trim(txtOpe(7).Text) & "'"
  Cad = Cad & ") c1 INNER JOIN (SELECT * FROM dosiscuerpo WHERE YEAR(f_dosis)=" & ano
  Cad = Cad & ") c2 USING(n_dosimetro,n_reg_dosimetro)"
  
' A ver si ahora..... >_>

'  If Check1(0).Value = 1 Then
'    ' solo los que no tienen fecha de baja
'    Cad = "SELECT distinct dosimetros.c_empresa, dosimetros.c_instalacion, dosimetros.dni_usuario, dosimetros.n_dosimetro, dosimetros.mes_p_i "
'
'    Cad = Cad & ",f_asig_dosimetro,f_retirada "
'
'    Cad = Cad & "from dosimetros, operarios where operarios.f_baja is null "
'    Cad = Cad & "and operarios.semigracsn = 1 and operarios.dni = dosimetros.dni_usuario "
'
'  Else
'    ' todos los usuarios tengan o no fecha de baja
'    Cad = "SELECT distinct dosimetros.c_empresa, dosimetros.c_instalacion, dosimetros.dni_usuario, dosimetros.n_dosimetro, dosimetros.mes_p_i "
'
'    Cad = Cad & ",f_asig_dosimetro,f_retirada "
'
'    Cad = Cad & "from dosimetros, operarios where operarios.semigracsn = 1 "
'    Cad = Cad & "and operarios.dni = dosimetros.dni_usuario "
'  End If
    
  ' solo seleccionamos los dosimetros de cuerpo
  ' ### [DavidV] 10/04/2006: Arreglos en la fórmula para solucionar un error de consulta.
  ' 23/06/2006: Ahora no, al parecer ahora tiene que ser como estaba antes...
  '
  ' Después de varios cambios aquí no registrados, se supone que si se marca
  ' "sólo fichero migrado" debe de salir aquellos con fecha de retirada NULL.
  ' si no, también aquellos que hayan sido dados de baja el año de la consulta.
  '
  ' ### [DavidV] 18/10/2006: Espero que esta sea la definitiva y última vez que me piden
  ' cambiar esto. Ya no sé ni cual era el código inicial.
''  Cad = Cad & " AND dosimetros.tipo_dosimetro = 0 and (dosimetros.f_retirada is null or "
''  Cad = Cad & "year(dosimetros.f_retirada) = " & ano & ")"
''  Cad = Cad & " AND dosimetros.tipo_dosimetro = 0 and (dosimetros.f_retirada is null"
''  If Check1(1).Value = 0 Then Cad = Cad & " or year(dosimetros.f_retirada)=" & ano
''  Cad = Cad & ") and dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%' "
'  Cad = Cad & " AND dosimetros.tipo_dosimetro = 0 and dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%' "
'  Cad = Cad & " and dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%' "

  ' ### [DavidV] 10/04/2006: Arreglos en la fórmula para ordenación por orden de recepción.
  ' ### [DavidV] 18/10/2006: Sin orden... no hace ninguna falta, porque no mostramos
  ' dosímetros, si no DOSIS basadas en estos, y agrupadas por meses.
  'Cad = Cad & " order by dosimetros.c_empresa, dosimetros.dni_usuario, dosimetros.c_instalacion "
  'Cad = Cad & " order by dosimetros.orden_recepcion, dosimetros.f_retirada"

  Set rs = New ADODB.Recordset
  rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
  If rs.EOF Then
    'NO hay registros a mostrar
    MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Screen.MousePointer = vbDefault
    
  Else
    CargarEmpresas
    CargarOperarios
    CargarInstalaciones
        
    'Mostramos el frame de resultados
    sql = "delete from zdosisacum where codusu = " & vUsu.codigo
    conn.Execute sql
        
    sql = "delete from zdosisacumtot where codusu = " & vUsu.codigo
    conn.Execute sql
        
    Set Rg = New ADODB.Recordset
        
    rs.MoveFirst
    While Not rs.EOF
      I = 1
      
      'La fecha retirada del dosimetro
      FechaBaja = CDate("31/12/9999")
      If Not IsNull(rs!f_retirada) Then FechaBaja = rs!f_retirada
            
      'If Check1(1).Value = 1 Then Cad = DevuelveDesdeBD(1, "n_dosimetro", "tempnc", "n_dosimetro|", Rs.Fields(3).Value & "|", "N|", 1)
      
      Do While I < 13 And Cad <> ""
            
        ' Evitamos que salga el mes actual si el listado es de este año.
        If ano = Year(Now) And I >= Month(Now) Then Exit Do
            
            
        'DAVID GANDUL###
        If Year(rs!f_asig_dosimetro) < ano Then
            'Año anterior a la fecha nforme
            AsignacionDosimetroFechaCorrecta = True
        Else
            If Year(rs!f_asig_dosimetro) > ano Then
                'Dado de alta en una año posterior al pedido en el informe
                AsignacionDosimetroFechaCorrecta = False
                
            Else
                'Esta dado de alta el mismo año
                If Month(rs!f_asig_dosimetro) > I Then
                    AsignacionDosimetroFechaCorrecta = False
                Else
                    AsignacionDosimetroFechaCorrecta = True
                End If
            End If
        End If
        
        'Contemplamos la fecha de retirada
        'Si se ha retirado en un determinado mes
        If AsignacionDosimetroFechaCorrecta Then
            If Year(FechaBaja) < ano Then
                'Año anterior a la fecha informe
                AsignacionDosimetroFechaCorrecta = False
            Else
                If Year(FechaBaja) > ano Then
                    'Dado de alta en una año posterior al pedido en el informe
                    AsignacionDosimetroFechaCorrecta = True
                    
                Else
                    'Esta dado de alta el mismo año
                    If Month(FechaBaja) > I Then
                        AsignacionDosimetroFechaCorrecta = True
                    Else
                        AsignacionDosimetroFechaCorrecta = False
                    End If
                End If
            End If
        End If

        If AsignacionDosimetroFechaCorrecta Then
                sql1 = "select sum(dosis_superf), sum(dosis_profunda) from dosiscuerpo "
                sql1 = sql1 & " where dni_usuario = '" & Trim(rs.Fields(2).Value) & "' and "
                sql1 = sql1 & " n_dosimetro = '" & Trim(rs.Fields(3).Value) & "' and "
                sql1 = sql1 & " c_instalacion = '" & Trim(rs.Fields(1).Value) & "' and "
                sql1 = sql1 & " month(f_dosis) = " & I & " and year(f_dosis) = " & Format(ano, "0000")
                
                Set rL = New ADODB.Recordset
                rL.Open sql1, conn, adOpenKeyset, adLockPessimistic, adCmdText
                
                ' el campo situ me indica: 0- situacion normal
                '                          1- sin alta en SDE
                '                          2- dosimetro no recibido
                fecha = CDate("01/" & Format(I, "00") & "/" & Format(ano, "0000"))
                
                'cursor para averiguar la situacion del operario en la instalacion
                Set rf = New ADODB.Recordset
                sql = "select f_alta, f_baja from operainstala where c_empresa = '" & Trim(rs.Fields(0).Value) & "' "
                sql = sql & " and c_instalacion = '" & Trim(rs.Fields(1).Value) & "' and "
                sql = sql & " dni = '" & Trim(rs.Fields(2).Value) & "'"
                rf.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
                
                situ = 1 ' sin alta en SDE
                If Not rf.EOF Then
                  rf.MoveFirst
                  Encontrado = False
                  While Not rf.EOF And Not Encontrado
                
                    f_mes = Month(rf!f_alta) + 1
                    f_ano = Year(rf!f_alta)
                    If f_mes > 12 Then
                      f_mes = 1
                      f_ano = f_ano + 1
                    End If
                                                                
                    fdesde = CDate("01/" & Format(f_mes, "00") & "/" & Format(f_ano, "0000"))
                                    
        '            fdesde = CDate("01/" & Format(Month(rf!f_alta), "00") & "/" & Format(Year(rf!f_alta), "0000"))
                    If IsNull(rf!f_baja) Then
                      fhasta = "31/12/9999"
                    Else
                      fhasta = "31/" & Format(Month(rf!f_baja), "00") & "/" & Format(Year(rf!f_baja), "0000")
                      If Not IsDate(fhasta) Then
                        fhasta = "30/" & Format(Month(rf!f_baja), "00") & "/" & Format(Year(rf!f_baja), "0000")
                      End If
                      If Not IsDate(fhasta) Then
                        fhasta = "29/" & Format(Month(rf!f_baja), "00") & "/" & Format(Year(rf!f_baja), "0000")
                      End If
                      If Not IsDate(fhasta) Then
                        fhasta = "28/" & Format(Month(rf!f_baja), "00") & "/" & Format(Year(rf!f_baja), "0000")
                      End If
                    End If
                    fhasta = CDate(fhasta)
                                    
                    If CDate(fecha) >= CDate(fdesde) And CDate(fecha) <= CDate(fhasta) Then
                      situ = 0
                      Encontrado = True
                    End If
                    rf.MoveNext
                  Wend
                End If
                Set rf = Nothing
                                                        
                If Not rL.EOF Then
                  If ((I = 1 Or I = 3 Or I = 5 Or I = 7 Or I = 9 Or I = 11) And rs.Fields(4).Value = "I") Or _
                     ((I = 2 Or I = 4 Or I = 6 Or I = 8 Or I = 10 Or I = 12) And rs.Fields(4).Value = "P") Then
                                        
                    If Not IsNull(rL.Fields(0).Value) Then
                      sql2 = "insert into zdosisacum (codusu, c_empresa, c_instalacion, "
                      sql2 = sql2 & "dni_usuario, mes, ano, n_dosimetro, dosissuper, dosisprofu, situ) VALUES ("
                      sql2 = sql2 & vUsu.codigo & ",'" & Trim(rs.Fields(0).Value) & "','"
                      sql2 = sql2 & Trim(rs.Fields(1).Value) & "','" & Trim(rs.Fields(2).Value) & "',"
                      sql2 = sql2 & I & "," & ano & ",'" & Trim(rs.Fields(3).Value) & "',"
                      
                      If IsNull(rL.Fields(0).Value) Then
                        Valor1 = 0
                      Else
                        Valor1 = rL.Fields(0).Value
                      End If
                                        
                      If IsNull(rL.Fields(1).Value) Then
                        valor2 = 0
                      Else
                        valor2 = rL.Fields(1).Value
                      End If
                                    
                      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(Valor1))) & ","
                      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(valor2))) & ","
                      sql2 = sql2 & Format(situ, "0") & ")"
                      
                      conn.Execute sql2
                    Else
                      'caso de dosimetros no recepcionados
                      If situ = 0 Then situ = 2
                      sql2 = "insert into zdosisacum (codusu, c_empresa, c_instalacion, "
                      sql2 = sql2 & "dni_usuario, mes, ano, n_dosimetro, dosissuper, dosisprofu, situ) VALUES ("
                      sql2 = sql2 & vUsu.codigo & ",'" & Trim(rs.Fields(0).Value) & "','"
                      sql2 = sql2 & Trim(rs.Fields(1).Value) & "','" & Trim(rs.Fields(2).Value) & "',"
                      sql2 = sql2 & I & "," & ano & ",'" & Trim(rs.Fields(3).Value) & "',"
                      sql2 = sql2 & "0.0, 0.0, " & Format(situ, "0") & ")"
                            
                      conn.Execute sql2
                    End If
                  End If
                End If
        
            
        End If   'f_asig_dosimetro
        Set rL = Nothing
        I = I + 1
      Loop        'Para cada mes
      rs.MoveNext
    Wend
        
    If CartaSobredosis Then
      rs.Close
             
      ' tenemos que eliminar aquellos registros de la temporal que no lleguen a sobredosis
      sql = "select c_empresa,dni_usuario, sum(dosissuper), sum(dosisprofu) from zdosisacum where codusu= " & vUsu.codigo
      sql = sql & " group by c_empresa, dni_usuario having sum(dosissuper) < 500 and sum(dosisprofu) < 20 "
      
      rs.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
      If Not rs.EOF Then
        rs.MoveFirst
                 
        While Not rs.EOF
          sql1 = "delete from zdosisacum where codusu = " & vUsu.codigo & " and "
          sql1 = sql1 & "c_empresa = '" & Trim(rs.Fields(0).Value) & "' and dni_usuario = '"
          sql1 = sql1 & Trim(rs.Fields(1).Value) & "'"
          conn.Execute sql1
          rs.MoveNext
        Wend
      Else
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
             
      rs.Close
      rs.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
      If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
            
      ' una vez cargada la temporal imprimimos el informe
      sql = "usu= " & vUsu.codigo & "|"
      sql = sql & "FechaAlta= ""Registros dosimétricos desde: " & Text3(22).Text & " hasta " & Text3(23).Text & """|"
      
      frmImprimir.Opcion = 23
      frmImprimir.NumeroParametros = 8
      frmImprimir.FormulaSeleccion = sql
      frmImprimir.OtrosParametros = sql
      frmImprimir.SoloImprimir = False
      frmImprimir.email = False
      frmImprimir.Show 'vbModal
    Else
      CargaAcumuladosQuinquenales
      
      Screen.MousePointer = vbDefault
    
      ' una vez cargada la temporal imprimimos el informe
      sql = "usu= " & vUsu.codigo & "|"
      sql = sql & "FechaAlta= ""Registros dosimétricos del SDP de Lainsa desde: " & "01/01/" & Format(Text4(0).Text, "0000") & " hasta " & "31/12/" & Format(Text4(0).Text, "0000") & """|"
      
      frmImprimir.Opcion = 22
      frmImprimir.NumeroParametros = 8
      frmImprimir.FormulaSeleccion = sql
      frmImprimir.OtrosParametros = sql
      frmImprimir.SoloImprimir = False
      frmImprimir.email = False
      frmImprimir.Show 'vbModal
        
    End If
  End If

End Sub

Private Sub CmdAceptarListFondos_Click()
Dim Tipo As String
Dim I As Integer
    If Not ComprobarFechas(12, 13) Then Exit Sub
    If Not ComprobarFechas(14, 15) Then Exit Sub
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT fondos.* from fondos WHERE 1 = 1 "
    If Text3(12).Text <> "" Then Cad = Cad & " AND f_inicio >= '" & Format(Text3(12).Text, FormatoFecha) & "' "
    If Text3(13).Text <> "" Then Cad = Cad & " AND f_inicio <= '" & Format(Text3(13).Text, FormatoFecha) & "' "
    If Text3(14).Text <> "" Then Cad = Cad & " AND f_fin >= '" & Format(Text3(14).Text, FormatoFecha) & "' "
    If Text3(15).Text <> "" Then Cad = Cad & " AND f_fin <= '" & Format(Text3(15).Text, FormatoFecha) & "' "
    For I = 0 To Option2.Count - 1
      If Option2(I).Value = True Then Tipo = Option2(I).Tag
    Next I
    If Tipo <> "" Then
      Cad = Cad & "and tipo = '" & Tipo & "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Fecha Inicio: "
        If Text3(12).Text <> "" Then
            sql = sql & "desfe1= """ & Format(Text3(4).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(4).Text & " "
        End If

        If Text3(13).Text <> "" Then
            sql = sql & "hasfe1= """ & Format(Text3(5).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(5).Text & " "
        End If

        If Text3(12).Text <> "" Or Text3(13).Text <> "" Then sql = sql & "FechaInicio= """ & cad1 & """|"


        cad1 = "Fecha Finalización: "
        If Text3(14).Text <> "" Then
            sql = sql & "desfe2= """ & Format(Text3(14).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(14).Text & " "
        End If

        If Text3(15).Text <> "" Then
            sql = sql & "hasfe2= """ & Format(Text3(15).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(15).Text & " "
        End If

        If Text3(14).Text <> "" Or Text3(15).Text <> "" Then sql = sql & "FechaFinal= """ & cad1 & """|"

        If Tipo <> "" Then sql = sql & "desdeTipo= """ & Tipo & """|" & "hastaTipo= """ & Tipo & """|"
        
        If Opcion = 18 Then frmImprimir.Opcion = 19 Else frmImprimir.Opcion = 36
        frmImprimir.NumeroParametros = 6
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.CampoOrden = 3
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If


End Sub

Private Sub CmdAceptarListProvincias_Click()
    If Not ComprobarProvincias(0, 1) Then Exit Sub
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT provincias.* from provincias WHERE 1 = 1 "
    If txtPro(0).Text <> "" Then Cad = Cad & " AND c_postal >= '" & Trim(txtPro(0).Text) & "' "
    If txtPro(1).Text <> "" Then Cad = Cad & " AND c_postal <= '" & Trim(txtPro(1).Text) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Provincias: "
        If txtPro(0).Text <> "" Then
            sql = sql & "despro= """ & Trim(txtPro(0).Text) & """|"
            cad1 = cad1 & " desde " & txtPro(0).Text & " " & DtxtPro(0).Text
        End If

        If txtPro(1).Text <> "" Then
            sql = sql & "haspro= """ & Trim(txtPro(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtPro(1).Text & " " & DtxtPro(1).Text
        End If

        If txtPro(0).Text <> "" Or txtPro(1).Text <> "" Then sql = sql & "Provincias= """ & cad1 & """|"

        If optPro(0).Value = True Then
            sql = sql & "orden= ""Por Código""|"
            frmImprimir.CampoOrden = 1
        End If
        If optPro(1).Value = True Then
            sql = sql & "orden= ""Alfabético""|"
            frmImprimir.CampoOrden = 2
        End If

        frmImprimir.Opcion = 13
        frmImprimir.NumeroParametros = 3
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarListRamasEsp_Click()
    If Not ComprobarRamasGenericas(2, 3) Then Exit Sub
    If Not ComprobarRamasEspecificas(0, 1) Then Exit Sub
    
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT ramaespe.* from ramaespe WHERE 1 = 1 "
    If txtRGe(2).Text <> "" Then Cad = Cad & " AND cod_rama_gen >= '" & Trim(txtRGe(2).Text) & "' "
    If txtRGe(3).Text <> "" Then Cad = Cad & " AND cod_rama_gen <= '" & Trim(txtRGe(3).Text) & "' "
    If txtREs(0).Text <> "" Then Cad = Cad & " AND c_rama_especifica >= '" & Trim(txtREs(0).Text) & "' "
    If txtREs(1).Text <> "" Then Cad = Cad & " AND c_rama_especifica <= '" & Trim(txtREs(1).Text) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Ramas Genéricas: "
        If txtRGe(2).Text <> "" Then
            sql = sql & "desrge= """ & Trim(txtRGe(2).Text) & """|"
            cad1 = cad1 & " desde " & txtRGe(2).Text & " " & DtxtRGe(2).Text
        End If

        If txtRGe(3).Text <> "" Then
            sql = sql & "hasrge= """ & Trim(txtRGe(3).Text) & """|"
            cad1 = cad1 & "   hasta " & txtRGe(3).Text & " " & DtxtRGe(3).Text
        End If

        If txtRGe(2).Text <> "" Or txtRGe(3).Text <> "" Then sql = sql & "RamaGenericas= """ & cad1 & """|"

        cad1 = "Ramas Específicas: "
        If txtREs(0).Text <> "" Then
            sql = sql & "desres= """ & Trim(txtREs(0).Text) & """|"
            cad1 = cad1 & " desde " & txtREs(0).Text & " " & DtxtREs(0).Text
        End If

        If txtREs(1).Text <> "" Then
            sql = sql & "hasres= """ & Trim(txtREs(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtREs(1).Text & " " & DtxtREs(1).Text
        End If

        If txtREs(0).Text <> "" Or txtREs(1).Text <> "" Then sql = sql & "RamaEspecificas= """ & cad1 & """|"



        If optRGe(2).Value = True Then
            sql = sql & "orden= ""Por Código""|"
            frmImprimir.CampoOrden = 1
        End If
        If optRGe(3).Value = True Then
            sql = sql & "orden= ""Alfabético""|"
            frmImprimir.CampoOrden = 2
        End If

        frmImprimir.Opcion = 16
        frmImprimir.NumeroParametros = 5
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarListRamasGenericas_Click()
    If Not ComprobarRamasGenericas(0, 1) Then Exit Sub
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT ramagene.* from ramagene WHERE 1 = 1 "
    If txtRGe(0).Text <> "" Then Cad = Cad & " AND cod_rama_gen >= '" & Trim(txtRGe(0).Text) & "' "
    If txtRGe(1).Text <> "" Then Cad = Cad & " AND cod_rama_gen <= '" & Trim(txtRGe(1).Text) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Ramas Genéricas: "
        If txtRGe(0).Text <> "" Then
            sql = sql & "desrge= """ & Trim(txtRGe(0).Text) & """|"
            cad1 = cad1 & " desde " & txtRGe(0).Text & " " & DtxtRGe(0).Text
        End If

        If txtRGe(1).Text <> "" Then
            sql = sql & "hasrge= """ & Trim(txtRGe(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtRGe(1).Text & " " & DtxtRGe(1).Text
        End If

        If txtRGe(0).Text <> "" Or txtRGe(1).Text <> "" Then sql = sql & "RamaGenericas= """ & cad1 & """|"

        If optRGe(0).Value = True Then
            sql = sql & "orden= ""Por Código""|"
            frmImprimir.CampoOrden = 1
        End If
        If optRGe(1).Value = True Then
            sql = sql & "orden= ""Alfabético""|"
            frmImprimir.CampoOrden = 2
        End If

        frmImprimir.Opcion = 15
        frmImprimir.NumeroParametros = 3
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarListRecepDosimCuerpo_Click()

    Screen.MousePointer = vbHourglass


    If Not ComprobarEmpresas(16, 17) Then Exit Sub
    If Not ComprobarInstalaciones(12, 13) Then Exit Sub
    If Not ComprobarOperarios(10, 11) Then Exit Sub
    If Not ComprobarFechas(24, 25) Then Exit Sub



    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT recepdosim.* from recepdosim WHERE 1 = 1 "
    If Opcion = 23 Then
        Cad = Cad & " and tipo_dosimetro <> 2 "
    Else
        Cad = Cad & " and tipo_dosimetro = 2 "
    End If
    If Combo4.ListIndex = 0 Then Cad = Cad & " and mes_p_i = 'P'"
    If Combo4.ListIndex = 1 Then Cad = Cad & " and mes_p_i = 'I'"
    
    If txtEmp(16).Text <> "" Then Cad = Cad & " AND c_empresa >= '" & Trim(txtEmp(16).Text) & "' "
    If txtEmp(17).Text <> "" Then Cad = Cad & " AND c_empresa <= '" & Trim(txtEmp(17).Text) & "' "
    If txtIns(12).Text <> "" Then Cad = Cad & " AND c_instalacion >= '" & Trim(txtIns(12).Text) & "' "
    If txtIns(13).Text <> "" Then Cad = Cad & " AND c_instalacion <= '" & Trim(txtIns(13).Text) & "' "
    If txtOpe(10).Text <> "" Then Cad = Cad & " AND dni_usuario >= '" & Trim(txtOpe(10).Text) & "' "
    If txtOpe(11).Text <> "" Then Cad = Cad & " AND dni_usuario <= '" & Trim(txtOpe(11).Text) & "' "
    If Text3(24).Text <> "" Then Cad = Cad & " AND f_creacion_recep >= '" & Format(Text3(24).Text, FormatoFecha) & "' "
    If Text3(25).Text <> "" Then Cad = Cad & " AND f_creacion_recep <= '" & Format(Text3(25).Text, FormatoFecha) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    
        Screen.MousePointer = vbDefault
    
    Else
        CargarEmpresas
        CargarOperarios
        CargarInstalaciones
        
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Empresas: "
        If txtEmp(16).Text <> "" Then
            sql = sql & "desemp= """ & Trim(txtEmp(16).Text) & """|"
            cad1 = cad1 & " desde " & txtEmp(16).Text & " " & DtxtEmp(16).Text
        End If

        If txtEmp(17).Text <> "" Then
            sql = sql & "hasemp= """ & Trim(txtEmp(17).Text) & """|"
            cad1 = cad1 & "   hasta " & txtEmp(17).Text & " " & DtxtEmp(17).Text
        End If

        If txtEmp(16).Text <> "" Or txtEmp(17).Text <> "" Then sql = sql & "Empresas= """ & cad1 & """|"

        ' codigo de instalaciones
        cad1 = "Instalaciones: "
        If txtIns(12).Text <> "" Then
            sql = sql & "desinst= """ & Trim(txtIns(12).Text) & """|"
            cad1 = cad1 & " desde " & txtIns(12).Text & " " & DtxtIns(12).Text
        End If

        If txtIns(13).Text <> "" Then
            sql = sql & "hasinst= """ & Trim(txtIns(13).Text) & """|"
            cad1 = cad1 & "   hasta " & txtIns(13).Text & " " & DtxtIns(13).Text
        End If

        If txtIns(12).Text <> "" Or txtIns(13).Text <> "" Then sql = sql & "Instalaciones= """ & cad1 & """|"

        ' dnis de operarios
        cad1 = "Operarios: "
        If txtOpe(10).Text <> "" Then
            sql = sql & "desdni= """ & Trim(txtOpe(10).Text) & """|"
            cad1 = cad1 & " desde " & txtOpe(10).Text & " " & DtxtIns(10).Text
        End If

        If txtOpe(11).Text <> "" Then
            sql = sql & "hasdni= """ & Trim(txtOpe(11).Text) & """|"
            cad1 = cad1 & "   hasta " & txtOpe(11).Text & " " & DtxtOpe(11).Text
        End If

        If txtOpe(10).Text <> "" Or txtOpe(11).Text <> "" Then sql = sql & "DNIs= """ & cad1 & """|"

        ' fechas de alta
        cad1 = "Fechas de Alta: "
        If Text3(24).Text <> "" Then
            sql = sql & "desfec= """ & Format(Text3(24).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(24).Text
        End If

        If Text3(25).Text <> "" Then
            sql = sql & "hasfec= """ & Format(Text3(25).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(25).Text
        End If

        If Text3(24).Text <> "" Or Text3(25).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"

        If Opcion = 23 Then  ' distinto de 2
            sql = sql & "Tipo= 0|"
        Else                 ' tipo 2 (solo de area)
            sql = sql & "Tipo= 1|"
        End If
        
        If Combo4.ListIndex = 0 Then
            sql = sql & "Par= 0|"
        Else
            If Combo4.ListIndex = 1 Then
                sql = sql & "Par= 1|"
            Else
                sql = sql & "Par= 2|"
            End If
        End If
        sql = sql & "usu= " & vUsu.codigo & "|"
        
        If Opcion = 23 Then  ' distinto de 2
            frmImprimir.Titulo = "Listado Recepción Dosímetros a Cuerpo"
        Else                 ' tipo 2 (solo de area)
            frmImprimir.Titulo = "Listado Recepción Dosímetros Area"
        End If
                
        If Check1(4).Value = 0 Then
            frmImprimir.NomDocu = "RecepDosimCuerpo.rpt"
        Else
            frmImprimir.NomDocu = "RecepDosimCuerpoInstala.rpt"
        End If
        
        Screen.MousePointer = vbDefault
        frmImprimir.Opcion = 24
        frmImprimir.NumeroParametros = 13
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show vbModal
     End If

End Sub

Private Sub CmdAceptarListTipoMedicion_Click()
    If Not ComprobarTiposMedicion(0, 1) Then Exit Sub
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT tipmedext.* from tipmedext WHERE 1 = 1 "
    If txtTMe(0).Text <> "" Then Cad = Cad & " AND c_tipo_med >= '" & Trim(txtTMe(0).Text) & "' "
    If txtTMe(1).Text <> "" Then Cad = Cad & " AND c_tipo_med <= '" & Trim(txtTMe(1).Text) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Tipos Medición: "
        If txtTMe(0).Text <> "" Then
            sql = sql & "desmed= """ & Trim(txtTMe(0).Text) & """|"
            cad1 = cad1 & " desde " & txtTMe(0).Text & " " & DtxtTMe(0).Text
        End If

        If txtTMe(1).Text <> "" Then
            sql = sql & "hasmed= """ & Trim(txtTMe(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtTMe(1).Text & " " & DtxtTMe(1).Text
        End If

        If txtTMe(0).Text <> "" Or txtTMe(1).Text <> "" Then sql = sql & "TiposMedicion= """ & cad1 & """|"

        If optTMe(0).Value = True Then
            sql = sql & "orden= ""Por Código""|"
            frmImprimir.CampoOrden = 1
        End If
        If optTMe(1).Value = True Then
            sql = sql & "orden= ""Alfabético""|"
            frmImprimir.CampoOrden = 2
        End If

        frmImprimir.Opcion = 14
        frmImprimir.NumeroParametros = 3
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarListTipoTrab_Click()
    If Not ComprobarRamasGenericas(4, 5) Then Exit Sub
    If Not ComprobarTipoTrab(0, 1) Then Exit Sub
    
    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT tipostrab.* from tipostrab WHERE 1 = 1 "
    If txtRGe(4).Text <> "" Then Cad = Cad & " AND cod_rama_gen >= '" & Trim(txtRGe(4).Text) & "' "
    If txtRGe(5).Text <> "" Then Cad = Cad & " AND cod_rama_gen <= '" & Trim(txtRGe(5).Text) & "' "
    If txtTTr(0).Text <> "" Then Cad = Cad & " AND c_tipo_trabajo >= '" & Trim(txtTTr(0).Text) & "' "
    If txtTTr(1).Text <> "" Then Cad = Cad & " AND c_tipo_trabajo <= '" & Trim(txtTTr(1).Text) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Ramas Genéricas: "
        If txtRGe(4).Text <> "" Then
            sql = sql & "desrge= """ & Trim(txtRGe(4).Text) & """|"
            cad1 = cad1 & " desde " & txtRGe(4).Text & " " & DtxtRGe(4).Text
        End If

        If txtRGe(5).Text <> "" Then
            sql = sql & "hasrge= """ & Trim(txtRGe(5).Text) & """|"
            cad1 = cad1 & "   hasta " & txtRGe(5).Text & " " & DtxtRGe(5).Text
        End If

        If txtRGe(4).Text <> "" Or txtRGe(5).Text <> "" Then sql = sql & "RamaGenericas= """ & cad1 & """|"

        cad1 = "Tipo Trabajo: "
        If txtTTr(0).Text <> "" Then
            sql = sql & "desttr= """ & Trim(txtTTr(0).Text) & """|"
            cad1 = cad1 & " desde " & txtTTr(0).Text & " " & DtxtTTr(0).Text
        End If

        If txtTTr(1).Text <> "" Then
            sql = sql & "hasttr= """ & Trim(txtTTr(1).Text) & """|"
            cad1 = cad1 & "   hasta " & txtTTr(1).Text & " " & DtxtTTr(1).Text
        End If

        If txtTTr(0).Text <> "" Or txtTTr(1).Text <> "" Then sql = sql & "TipoTrabajos= """ & cad1 & """|"



        If optRGe(4).Value = True Then
            sql = sql & "orden= ""Por Código""|"
            frmImprimir.CampoOrden = 1
        End If
        If optRGe(5).Value = True Then
            sql = sql & "orden= ""Alfabético""|"
            frmImprimir.CampoOrden = 2
        End If

        frmImprimir.Opcion = 17
        frmImprimir.NumeroParametros = 5
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarListOperariosSobredosis_Click()
Dim rs As ADODB.Recordset
Dim rL As ADODB.Recordset
Dim ano As Integer
Dim sql1 As String
Dim sql2 As String
Dim fecha As String
Dim Valor1 As Currency
Dim valor2 As Currency
Dim fecini As Date

    Screen.MousePointer = vbHourglass

    If Not ComprobarFechas(21, 21) Then Exit Sub
    If Not ComprobarEmpresas(18, 19) Then Exit Sub
    If Not ComprobarOperarios(12, 13) Then Exit Sub
    
    If Not IsDate(Text3(21).Text) Then
      MsgBox "Es necesario indicar una fecha válida.", vbExclamation, "¡Atención!"
      Screen.MousePointer = vbDefault
      Text3(21).SetFocus
      Exit Sub
    End If
    
    ano = Year(CDate(Text3(21).Text))
    
    fecini = CDate("01/01/" & Format(ano, "0000"))
    fecha = CStr(fecini)



    CargarOperarios
    CargarInstalaciones

    sql1 = "select dosiscuerpo.dni_usuario, sum(dosis_superf), sum(dosis_profunda) from dosiscuerpo, dosimetros, voperarios "
    sql1 = sql1 & " where 1 = 1 and "
    sql1 = sql1 & " voperarios.codusu = " & vUsu.codigo & " and "
    sql1 = sql1 & " dosiscuerpo.f_dosis <= '" & Format(Text3(21).Text, FormatoFecha) & "' and "
    sql1 = sql1 & " dosiscuerpo.f_dosis >= '" & Format(fecini, FormatoFecha) & "' "
    sql1 = sql1 & " and dosimetros.tipo_dosimetro = 0 and dosimetros.f_retirada is null "
    sql1 = sql1 & " and voperarios.semigracsn = 1 "
'    sql1 = sql1 & " and voperarios.f_baja is null and voperarios.semigracsn = 1 "
    
    If txtEmp(18).Text <> "" Then Cad = Cad & " AND dosiscuerpo.c_empresa >= '" & Trim(txtEmp(18).Text) & "' "
    If txtEmp(19).Text <> "" Then Cad = Cad & " AND dosiscuerpo.c_empresa <= '" & Trim(txtEmp(19).Text) & "' "
    If txtOpe(12).Text <> "" Then Cad = Cad & " AND dosiscuerpo.dni_usuario >= '" & Trim(txtOpe(12).Text) & "' "
    If txtOpe(13).Text <> "" Then Cad = Cad & " AND dosiscuerpo.dni_usuario <= '" & Trim(txtOpe(13).Text) & "' "
    
    sql1 = sql1 & " and dosiscuerpo.dni_usuario = voperarios.dni "
    sql1 = sql1 & " and dosiscuerpo.n_reg_dosimetro = dosimetros.n_reg_dosimetro "
    sql1 = sql1 & " group by dosiscuerpo.dni_usuario "
    sql1 = sql1 & " having sum(dosis_superf) >= 500 or sum(dosis_profunda) >= 20"
    
    
    Set rs = New ADODB.Recordset
    rs.Open sql1, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        Screen.MousePointer = vbDefault
        
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = "delete from zdosisacumtot where codusu = " & vUsu.codigo
        conn.Execute sql
        
        rs.MoveFirst
        While Not rs.EOF
            sql2 = "insert into zdosisacumtot (codusu, c_empresa, c_instalacion, "
            sql2 = sql2 & "dni_usuario,dosissuper, dosisprofu) VALUES ("
            sql2 = sql2 & vUsu.codigo & ","
            
            'la instalacion la sacamos de operainstala la que no tenga fecha de baja
            sql = "select c_empresa, c_instalacion from operainstala where dni = '" & Trim(rs.Fields(0).Value) & "' "
            sql = sql & " and f_baja is null "
            
            Set rL = New ADODB.Recordset
            rL.Open sql, conn, adOpenKeyset, adLockPessimistic, adCmdText
            If rL.EOF Then
                'NO hay registros a mostrar
                sql2 = sql2 & "'', '',"
            Else
                rL.MoveFirst
                sql2 = sql2 & "'" & Trim(rL.Fields(0).Value) & "','" & Trim(rL.Fields(1).Value) & "',"
            End If
            
            sql2 = sql2 & "'" & Trim(rs.Fields(0).Value) & "',"
                      
            If IsNull(rs.Fields(1).Value) Then
                Valor1 = 0
            Else
                Valor1 = rs.Fields(1).Value
            End If
            
            If IsNull(rs.Fields(2).Value) Then
                valor2 = 0
            Else
                valor2 = rs.Fields(2).Value
            End If
            
            sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(Valor1))) & ","
            sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(valor2))) & ")"
            
            conn.Execute sql2
           
            rs.MoveNext
        Wend
        
        rs.Close
        
        Screen.MousePointer = vbDefault
       
        ' una vez cargada la temporal imprimimos el informe
        sql = "usu= " & vUsu.codigo & "|"
        sql = sql & "FechaAlta= ""Registros dosimétricos desde: " & Format(fecini, "dd/mm/yyyy") & " hasta " & Text3(21).Text & """|"
        
        frmImprimir.Opcion = 31
        frmImprimir.NumeroParametros = 2
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If




End Sub

Private Sub CmdAceptarListUsu_Click()
    If Not ComprobarOperarios(0, 1) Then Exit Sub
    If Not ComprobarFechas(0, 1) Then Exit Sub

    'Hacemos el select y si tiene resultados mostramos los valores

    Cad = " SELECT operarios.* from operarios WHERE 1 = 1 "
    If txtOpe(14).Text <> "" Then Cad = Cad & " AND dni >= '" & Trim(txtOpe(14).Text) & "' "
    If txtOpe(15).Text <> "" Then Cad = Cad & " AND dni <= '" & Trim(txtOpe(15).Text) & "' "
    If Text3(20).Text <> "" Then Cad = Cad & " AND f_alta >= '" & Format(Text3(20).Text, FormatoFecha) & "' "
    If Text3(26).Text <> "" Then Cad = Cad & " AND f_alta <= '" & Format(Text3(26).Text, FormatoFecha) & "' "

    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        ' dnis de operarios
        cad1 = "Operarios: "
        If txtOpe(14).Text <> "" Then
            sql = sql & "desope= """ & Trim(txtOpe(14).Text) & """|"
            cad1 = cad1 & " desde " & txtOpe(14).Text & " " & DtxtOpe(14).Text
        End If

        If txtOpe(15).Text <> "" Then
            sql = sql & "hasope= """ & Trim(txtOpe(15).Text) & """|"
            cad1 = cad1 & "   hasta " & txtOpe(15).Text & " " & DtxtOpe(15).Text
        End If

        If txtOpe(14).Text <> "" Or txtOpe(15).Text <> "" Then sql = sql & "DNIs= """ & cad1 & """|"

        ' fechas de alta
        cad1 = "Fechas de Alta: "
        If Text3(20).Text <> "" Then
            sql = sql & "desfec= """ & Format(Text3(20).Text, FormatoFecha) & """|"
            cad1 = cad1 & " desde " & Text3(20).Text
        End If

        If Text3(26).Text <> "" Then
            sql = sql & "hasfec= """ & Format(Text3(26).Text, FormatoFecha) & """|"
            cad1 = cad1 & "   hasta " & Text3(26).Text
        End If

        If Text3(20).Text <> "" Or Text3(26).Text <> "" Then sql = sql & "FechaAlta= """ & cad1 & """|"



        If OptOpe(15).Value = True Then
            sql = sql & "tipo= 1|"
        End If
        If OptOpe(16).Value = True Then
            sql = sql & "tipo= 0|"
        End If

        If OptOpe(10).Value Then
            sql = sql & "tipo1= 0|"
        End If
        If OptOpe(11).Value Then
            sql = sql & "tipo1= 1|"
        End If
        If OptOpe(12).Value Then
            sql = sql & "tipo1= 2|"
        End If


        If Opcion = 30 Then frmImprimir.Opcion = 33
        
        frmImprimir.NumeroParametros = 10
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If

End Sub

Private Sub CmdAceptarLotes_Click()
Dim Tipo As String
Dim tabla As String
Dim I As Integer
Dim error As Integer
    
    error = 0
    If Not IsNumeric(TextLot(0).Text) And TextLot(0).Text <> "" Then error = 32
    If Not IsNumeric(TextLot(1).Text) And TextLot(1).Text <> "" Then error = 31
    If Not IsNumeric(TextLot(2).Text) And TextLot(2).Text <> "" Then error = 28
    If Not IsNumeric(TextLot(3).Text) And TextLot(3).Text <> "" Then error = 27
    If error <> 0 Then
      MsgBox "El campo ha de ser numérico.", vbOKOnly + vbExclamation, "¡Error!"
      Text3(error).SetFocus
      Exit Sub
    End If
    error = 0
    If Val(TextLot(0).Text) > Val(TextLot(1).Text) Then error = 31
    If Val(TextLot(2).Text) > Val(TextLot(3).Text) Then error = 27
    If error <> 0 Then
      MsgBox "El dosímetro final no puede menor que el inicial.", vbOKOnly + vbExclamation, "¡Error!"
      Text3(error).SetFocus
      Exit Sub
    End If
    
    'Hacemos el select y si tiene resultados mostramos los valores
    tabla = IIf(Opcion = 29, "lotes", "lotespana")
    Cad = " SELECT " & tabla & ".* from " & tabla & " WHERE 1 = 1 "
    If TextLot(0).Text <> "" Then Cad = Cad & " AND dosimetro_inicial >= " & TextLot(0).Text & " "
    If TextLot(1).Text <> "" Then Cad = Cad & " AND dosimetro_inicial <= " & TextLot(1).Text & " "
    If TextLot(2).Text <> "" Then Cad = Cad & " AND dosimetro_final >= " & TextLot(2).Text & " "
    If TextLot(3).Text <> "" Then Cad = Cad & " AND dosimetro_final <= " & TextLot(3).Text & " "
    For I = 0 To Option2.Count - 1
      If Option2(I).Value = True Then Tipo = Option2(I).Tag
    Next I
    If Tipo <> "" Then
      Cad = Cad & "and tipo = '" & Tipo & "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        'Mostramos el frame de resultados
        sql = ""
        cad1 = "Dosímetro Inicial: "
        If TextLot(0).Text <> "" Then
            sql = sql & "desdos= " & TextLot(0).Text & "|"
            cad1 = cad1 & " desde " & TextLot(0).Text & " "
        End If

        If TextLot(1).Text <> "" Then
            sql = sql & "hasdos= " & TextLot(1).Text & "|"
            cad1 = cad1 & "   hasta " & TextLot(1).Text & " "
        End If

        If TextLot(0).Text <> "" Or TextLot(1).Text <> "" Then sql = sql & "DosimetroInicial= """ & cad1 & """|"


        cad1 = "Dosímetro Final: "
        If TextLot(2).Text <> "" Then
            sql = sql & "desdos2= " & TextLot(2).Text & "|"
            cad1 = cad1 & " desde " & TextLot(2).Text & " "
        End If

        If TextLot(3).Text <> "" Then
            sql = sql & "hasdos2= " & TextLot(3).Text & "|"
            cad1 = cad1 & "   hasta " & TextLot(3).Text & " "
        End If

        If TextLot(2).Text <> "" Or TextLot(3).Text <> "" Then sql = sql & "DosimetroFinal= """ & cad1 & """|"

        If Tipo <> "" Then sql = sql & "desdeTipo= """ & Tipo & """|" & "hastaTipo= """ & Tipo & """|"
        
        frmImprimir.Opcion = Me.Opcion + 4
        frmImprimir.NumeroParametros = 6
        frmImprimir.FormulaSeleccion = sql
        frmImprimir.OtrosParametros = sql
        frmImprimir.CampoOrden = 3
        frmImprimir.SoloImprimir = False
        frmImprimir.email = False
        frmImprimir.Show 'vbModal
     End If


End Sub

Private Sub CmdCanCartaDosimNRec_Click()
    Unload Me
End Sub

Private Sub CmdCanCartaSobredosis_Click()
    Unload Me
End Sub

Private Sub cmdCancelarLotes_Click()
    Unload Me
End Sub

Private Sub CmdCanLisDosim_Click()
    Unload Me
End Sub

Private Sub CmdCanLisDosisCol_Click()
    Unload Me
End Sub

Private Sub CmdCanLisFactCalib_Click()
    Unload Me
End Sub

Private Sub cmdCanLisIns_Click()
    Unload Me
End Sub

Private Sub CmdCanLisOpe_Click()
    Unload Me
End Sub

Private Sub CmdCanListDosisIns_Click()
    Unload Me
End Sub

Private Sub CmdCanListDosisNHomOpe_Click()
    Unload Me
End Sub

Private Sub CmdCanListDosisOpeAcum12_Click()
    Unload Me
End Sub

Private Sub CmdCanListFondos_Click()
    Unload Me
End Sub

Private Sub CmdCanListOperariosSobredosis_Click()
    Unload Me
End Sub

Private Sub CmdCanListPro_Click()
    Unload Me
End Sub

Private Sub CmdCanListRamasEsp_Click()
    Unload Me
End Sub

Private Sub CmdCanListRamasGenericas_Click()
    Unload Me
End Sub

Private Sub CmdCanListRecepDosimCuerpo_Click()
    Unload Me
End Sub

Private Sub CmdCanListTipoMedicion_Click()
    Unload Me
End Sub


Private Sub CmdCanListUsu_Click()
    Unload Me
End Sub

Private Sub CmdCanTiposTrab_Click()
    Unload Me
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub frmREs_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtREs(Empresa).Text = RecuperaValor(CadenaSeleccion, 2)
    Me.DtxtREs(Empresa).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmrge_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtRGe(Empresa).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.DtxtRGe(Empresa).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTMe_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtTMe(Empresa).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.DtxtTMe(Empresa).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTTR_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtTTr(Empresa).Text = RecuperaValor(CadenaSeleccion, 2)
    Me.DtxtTTr(Empresa).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub ImgOpe_Click(Index As Integer)
    Operario = Index
    Set frmOpe = New frmOperarios
    frmOpe.DatosADevolverBusqueda = "9|"
    frmOpe.Show
    
End Sub

Private Sub ImgREs_Click(Index As Integer)
    Empresa = Index
    Set frmREs = New frmRamasEspe
    frmREs.DatosADevolverBusqueda = "2|3|"
    frmREs.Show
End Sub

Private Sub ImgRGe_Click(Index As Integer)
    Empresa = Index
    Set frmRGe = New frmRamasGener
    frmRGe.DatosADevolverBusqueda = "0|1|"
    frmRGe.Show
End Sub

Private Sub ImgTMe_Click(Index As Integer)
    Empresa = Index
    Set frmTMe = New frmTiposExtremidades
    frmTMe.DatosADevolverBusqueda = "0|1|"
    frmTMe.Show

End Sub

Private Sub ImgTTr_Click(Index As Integer)
    Empresa = Index
    Set frmTTr = New frmTiposTrab
    frmTTr.DatosADevolverBusqueda = "2|3|"
    frmTTr.Show
End Sub

Private Sub optIns_Click(Index As Integer)
  
  Select Case Index
    Case 19, 21
      Image2(11).Enabled = False
      Image2(29).Enabled = False
      Text3(31).Enabled = False
      Text3(32).Enabled = False
      Text3(31).BackColor = &H80000018
      Text3(32).BackColor = &H80000018
      Text3(31).Text = ""
      Text3(32).Text = ""
    Case 20
      Image2(11).Enabled = True
      Image2(29).Enabled = True
      Text3(31).Enabled = True
      Text3(32).Enabled = True
      Text3(31).BackColor = &H80000005
      Text3(32).BackColor = &H80000005
  End Select
  
  If Index <> 5 And Index <> 6 And Index <> 7 Then Exit Sub
  If Text3(10).Text <> "" Then
      Text3(11).Text = CargarFechaHasta(CDate(Text3(10).Text), 1)
  End If
End Sub

Private Sub OptOpe_Click(Index As Integer)
  
  Select Case Index
    Case 8, 7
      Image2(27).Enabled = False
      Image2(28).Enabled = False
      Text3(27).Enabled = False
      Text3(28).Enabled = False
      Text3(27).BackColor = &H80000018
      Text3(28).BackColor = &H80000018
      Text3(27).Text = ""
      Text3(28).Text = ""
    Case 9
      Image2(27).Enabled = True
      Image2(28).Enabled = True
      Text3(27).Enabled = True
      Text3(28).Enabled = True
      Text3(27).BackColor = &H80000005
      Text3(28).BackColor = &H80000005
  End Select

End Sub

Private Sub txtDos_GotFocus(Index As Integer)
    txtDos(Index).SelStart = 0
    txtDos(Index).SelLength = Len(txtDos(Index).Text)
End Sub

Private Sub txtDos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDos_LostFocus(Index As Integer)
    If txtDos(Index).Text <> "" Then
        If Not IsNumeric(txtDos(Index).Text) Then
            MsgBox "El número de dosímetro ha de ser numérico", vbExclamation, "¡Error!"
            Exit Sub
        End If
    End If

End Sub

Private Sub txtEmp_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtIns_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtIns_LostFocus(Index As Integer)
    txtIns(Index).Text = Trim(txtIns(Index).Text)
    If txtIns(Index).Text = "" Then
        DtxtIns(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtIns(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtIns(Index).Text = Replace(Format(txtIns(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtIns(Index)
        Exit Sub
    End If
    txtIns(Index).Text = Format(txtIns(Index).Text, ">")


    sql = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_instalacion|", txtIns(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        DtxtIns(Index).Text = sql
    Else
        DtxtIns(Index).Text = "No existe la instalación"
    End If

End Sub

Private Sub txtope_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'Private Sub CmdAceptarCuadre_Click()
'Dim TipoContado As String
'Dim NomDocu As String
'Dim indRPT As Integer
'Dim cadparam As String
'Dim numParam As Integer
'
'    Screen.MousePointer = vbHourglass
'
''    If Not ComprobarFechas(41, 42) Then Exit Sub
'
'    If Text3(2).Text = "" Then
'        Text3(2).Text = Format(Text3(2).Text, FormatoFecha)
'    End If
'
'    'VRS:1.0.5(9)
'    Text3(41).Text = Text3(2).Text
'    Text3(42).Text = Text3(2).Text
'    If GenerarControlCuadre(vUsu.codigo, Text3(41).Text, Text3(42).Text) Then
'        Sql = "usu= " & vUsu.codigo & "|"
'        'Sql = Sql & "forpa= " & txtFpa(6).Text & "|"
'
'        'Sql = Sql & "nomforpa= """ & DtxtFpa(6).Text & """|"
'
'        Cad1 = "Fecha Cobro: " & Text3(41).Text
'
'        If Text3(41).Text <> "" Or Text3(42).Text <> "" Then Sql = Sql & "Fechas= """ & Cad1 & """|"
'
'        If TxtImp(0).Text <> "" Then
'            Sql = Sql & "importe= " & TransformaComasPuntos(ImporteSinFormato(TxtImp(0).Text)) & "|"
'        Else
'            Sql = Sql & "importe= 0|"
'        End If
'
'        NomDocu = ""
'        cadparam = ""
'        numParam = 0
'        indRPT = 7
'
'        If Not PonerParamEmpresa(indRPT, cadparam, numParam, NomDocu) Then
'            Exit Sub
'        End If
'
'
'        frmImprimir.Opcion = 20
'        frmImprimir.NumeroParametros = 4
'        frmImprimir.NomDocu = NomDocu
'        frmImprimir.FormulaSeleccion = Sql
'        frmImprimir.OtrosParametros = Sql
'        frmImprimir.SoloImprimir = False
'        frmImprimir.Show vbModal
'
'        If vParam.HayContabilidad Then
'            Sql = vbCrLf & "Desea actualizar la caja y contabilizar.   " & vbCrLf
'        Else
'            Sql = vbCrLf & "Desea actualizar la caja    " & vbCrLf
'        End If
'
'        If MsgBox(Sql, vbQuestion + vbDefaultButton2 + vbYesNoCancel) = vbYes Then
'            If Text3(2).Text = "" Then Text3(2).Text = Format(Now, FormatoFecha)
'            If vParam.HayContabilidad Then
'                'la actualizacion incluye borrado de efectos dentro de la transaccion
'                ActualizarCobrosContados Text3(41).Text, Text3(42).Text, Text3(2).Text
'            Else
'                'hemos de borrar contados de cartera
'                BorradoCobrosContados Text3(41).Text, Text3(42).Text
'            End If
'        End If
'        Unload Me
'
'    End If
'
'    Screen.MousePointer = vbDefault
'
'End Sub

'Private Sub CmdAceptarFacturaAlb_Click()
'Dim mC As Contadores
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'Dim Sql As String
'Dim Sql1 As String
'Dim DesdeFactura As Long
'Dim HastaFactura As Long
'Dim AntSocio As Long
'Dim AntForpa As Integer
'Dim AntDtogen As Double
'Dim AntDtoppa As Double
'Dim v_aux As Long
'
'Dim NomDocu As String
'Dim indRPT As Integer
'Dim cadparam As String
'Dim numParam As Integer
'
'    On Error GoTo EFacturacionAlbaranes
'
'    If Not ComprobarAlbaranes(0, 1) Then Exit Sub
'
'    If Not ComprobarFechas(35, 36) Then Exit Sub
'
'    If Not ComprobarSocios(0, 1) Then Exit Sub
'
'    If Not ComprobarFormasPago(2, 3) Then Exit Sub
'
'    'VRS:1.0.1(9)
'    If Combo2(0).ListIndex = -1 Then
'        MsgBox "Debe de introducir el número de serie de la factura", vbExclamation
'        Exit Sub
'    End If
'
'    Cad = "select * from scaalb where tipofact = 1 "
'    If TxtAlb(0).Text <> "" Then Cad = Cad & " and numalbar >= " & TxtAlb(0).Text
'    If TxtAlb(1).Text <> "" Then Cad = Cad & " and numalbar <= " & TxtAlb(1).Text
'    If Text3(35).Text <> "" Then Cad = Cad & " and fecalbar >= '" & Format(Text3(35).Text, FormatoFecha) & "'"
'    If Text3(36).Text <> "" Then Cad = Cad & " and fecalbar <= '" & Format(Text3(36).Text, FormatoFecha) & "'"
'    If txtFpa(2).Text <> "" Then Cad = Cad & " and codforpa >= " & txtFpa(2).Text
'    If txtFpa(3).Text <> "" Then Cad = Cad & " and codforpa <= " & txtFpa(3).Text
'    If txtSoc(0).Text <> "" Then Cad = Cad & " and codsocio >= " & txtSoc(0).Text
'    If txtSoc(1).Text <> "" Then Cad = Cad & " and codsocio <= " & txtSoc(1).Text
'    Cad = Cad & " order by codsocio, codforpa, dtognral, dtoppago, numserie, numalbar "
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
'    If Rs.EOF Then
'        'NO hay registros a mostrar
'        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
'        'Cont = -1
'        Exit Sub
'    Else
'        'Mostramos el frame de resultados
'        Cont = 0
'        While Not Rs.EOF
'            Cont = Cont + 1
'            Rs.MoveNext
'        Wend
'        If Cont > 32000 Then Cont = 32000
'        pb2.Max = Cont + 1
'        pb2.Visible = True
'        pb2.Value = 0
'        Me.Refresh
'
'        pb2.Value = pb2.Value + 1
'        pb2.Refresh
'        Rs.MoveFirst
'
'        Conn.BeginTrans
'
'        Set mC = New Contadores
'        DesdeFactura = -1
'
'        AntSocio = -1
'        AntForpa = -1
'        AntDtogen = -1
'        AntDtoppa = -1
'
'        While Not Rs.EOF
'            ' cogemos el contador de factura, solo si toca
'            If Rs!codsocio <> AntSocio Or _
'               Rs!codforpa <> AntForpa Or _
'               Rs!dtognral <> AntDtogen Or _
'               Rs!DtoPPago <> AntDtoppa Then
'
'                'InsertarTotales
'                ' insertamos en la tabla de totales de factura
'                ' calculo de importes para insertar
'                ' una vez tenemos cargados los brutos de la factura calculamos bases e ivas
'                If DesdeFactura <> -1 Then
'                    InsertarTotales Combo2(0).Text, mC.Contador, Text3(9).Text
'                End If
'
'
'                mC.ConseguirContador Combo2(0).Text, 2, False
'
'                AntSocio = Rs!codsocio
'                AntForpa = Rs!codforpa
'                AntDtogen = Rs!dtognral
'                AntDtoppa = Rs!DtoPPago
'
'
'                If DesdeFactura = -1 Then DesdeFactura = mC.Contador
'
'                ' iniciamos variables de totales
'                Base1 = 0
'                Base2 = 0
'                Base3 = 0
'                bruto1 = 0
'                bruto2 = 0
'                bruto3 = 0
'                Tipo1 = 0
'                Tipo2 = 0
'                Tipo3 = 0
'                'facturas
'                ivainclu = vParam.IvaIncluidoFactu
'                DtoGeneral = Rs!dtognral
'                DtoPPago = Rs!DtoPPago
'
'            End If
'
'            ' insertamos cabecera de factura
'            Sql = "insert into scafac (tipofact, numserie, numfactu, "
'            Sql = Sql & "fecfactu, numseralb, numalbar, fecalbar, codsocio, codforpa, "
'            Sql = Sql & "dtognral, dtoppago, observac, codusu, contabilizado) VALUES ("
'            Sql = Sql & "2,'" & DevNombreSQL(Combo2(0).Text) & "'," & mC.Contador & ",'"
'            Sql = Sql & Format(Text3(9).Text, FormatoFecha) & "','" & DevNombreSQL(Rs!numserie) & "'," & Rs!NumAlbar & ",'"
'            Sql = Sql & Format(Rs!fecalbar, FormatoFecha) & "'," & Rs!codsocio & ","
'            Sql = Sql & Rs!codforpa & "," & TransformaComasPuntos(ImporteSinFormato(Rs!dtognral)) & ","
'            Sql = Sql & TransformaComasPuntos(ImporteSinFormato(Rs!DtoPPago)) & ",'"
'
'            If Not IsNull(Rs!observac) Then
'                Sql = Sql & DevNombreSQL(Rs!observac) & "'," & Rs!Codusu & ",1)"
'            Else
'                Sql = Sql & "'," & Rs!Codusu & ",1)"
'            End If
'            Conn.Execute Sql
'
'            ' insertamos linea de factura
'            Sql1 = "select * from slialb where tipofact = 1 and numserie = '" & DevNombreSQL(Rs!numserie) & "' and "
'            Sql1 = Sql1 & " numalbar = " & Rs!NumAlbar
'
'            Set RS1 = New ADODB.Recordset
'            RS1.Open Sql1, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'            If Not RS1.EOF Then RS1.MoveFirst
'
'            While Not RS1.EOF
'                Sql = "insert into slifac (tipofact, numserie, numfactu, fecfactu,"
'                Sql = Sql & " numseralb, numalbar, "
'                Sql = Sql & " numlinea, codartic, codigean, cantidad, precioar, dtoline1,"
'                Sql = Sql & " implinea, ampliaci) VALUES (2,'" & DevNombreSQL(Combo2(0).Text) & "',"
'                Sql = Sql & mC.Contador & ",'" & Format(Text3(9).Text, FormatoFecha) & "','" & DevNombreSQL(Rs!numserie) & "'," & Rs!NumAlbar & ","
'                Sql = Sql & RS1!Numlinea & ","
'                Sql = Sql & RS1!codArtic & ",'" & DevNombreSQL(RS1!codigean) & "'," & TransformaComasPuntos(ImporteSinFormato(RS1!cantidad)) & ","
'                Sql = Sql & TransformaComasPuntos(ImporteSinFormato(RS1!Precioar)) & ","
'                Sql = Sql & TransformaComasPuntos(ImporteSinFormato(RS1!dtoline1)) & ","
'                Sql = Sql & TransformaComasPuntos(ImporteSinFormato(RS1!implinea)) & ",'"
'                Sql = Sql & DevNombreSQL(RS1!ampliaci) & "')"
'
'                Conn.Execute Sql
'
'                'calculos de importe
'                TipoIva = 0
'                TipoIva = CCur(DevuelveDesdeBD(1, "codi_iva", "sartic", "codartic|", RS1!codArtic & "|", "N|", 1))
'                If TipoIva <> Tipo1 And TipoIva <> Tipo2 And TipoIva <> Tipo3 Then
'                    If Tipo1 = 0 Then
'                        Tipo1 = TipoIva
'
'                        If vParam.HayContabilidad Then
'                            PorcIva1 = CCur(DevuelveDesdeBD(2, "porceiva", "tiposiva", "codigiva|", Tipo1 & "|", "N|", 1))
'                        Else
'                            PorcIva1 = CCur(DevuelveDesdeBD(1, "porceiva", "tiposiva", "codigiva|", Tipo1 & "|", "N|", 1))
'                        End If
'                    Else
'                    If Tipo2 = 0 Then
'                        Tipo2 = TipoIva
'                        If vParam.HayContabilidad Then
'                            PorcIva2 = CCur(DevuelveDesdeBD(2, "porceiva", "tiposiva", "codigiva|", Tipo2 & "|", "N|", 1))
'                        Else
'                            PorcIva2 = CCur(DevuelveDesdeBD(1, "porceiva", "tiposiva", "codigiva|", Tipo2 & "|", "N|", 1))
'                        End If
'
'                    Else
'                    If Tipo3 = 0 Then
'                        Tipo3 = TipoIva
'                        If vParam.HayContabilidad Then
'                            PorcIva3 = CCur(DevuelveDesdeBD(2, "porceiva", "tiposiva", "codigiva|", Tipo3 & "|", "N|", 1))
'                        Else
'                            PorcIva3 = CCur(DevuelveDesdeBD(1, "porceiva", "tiposiva", "codigiva|", Tipo3 & "|", "N|", 1))
'                        End If
'
'                    End If
'                    End If
'                    End If
'                End If
'                Select Case TipoIva
'                    Case Tipo1
'                         bruto1 = bruto1 + RS1!implinea
'                    Case Tipo2
'                         bruto2 = bruto2 + RS1!implinea
'                    Case Tipo3
'                         bruto3 = bruto3 + RS1!implinea
'                End Select
'
'                ' insertamos sublineas
'                Sql = "select * from slialb1 where tipofact = 1 and numserie = '" & DevNombreSQL(Rs!numserie)
'                Sql = Sql & "' and numalbar = " & Rs!NumAlbar & " and "
'                Sql = Sql & " numlinea = " & RS1!Numlinea
'
'                Set Rs2 = New ADODB.Recordset
'                Rs2.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'                If Not Rs2.EOF Then Rs2.MoveFirst
'                While Not Rs2.EOF
'                    Sql = "insert into slifac1 (tipofact, numserie, numfactu, fecfactu, "
'                    Sql = Sql & " numseralb, numalbar, "
'                    Sql = Sql & " numlinea, sublinea, codartic, nrolotes, regfitosanitario,"
'                    Sql = Sql & "cantidad) VALUES (2,'" & DevNombreSQL(Combo2(0).Text) & "',"
'                    Sql = Sql & mC.Contador & ",'" & Format(Text3(9).Text, FormatoFecha) & "','" & DevNombreSQL(Rs!numserie) & "'," & Rs!NumAlbar & "," & Rs2!Numlinea & ","
'                    Sql = Sql & Rs2!Sublinea & "," & Rs2!codArtic & ",'" & DevNombreSQL(Rs2!nrolotes) & "','"
'                    Sql = Sql & DevNombreSQL(Rs2!Regfitosanitario) & "',"
'                    Sql = Sql & TransformaComasPuntos(ImporteSinFormato(Rs2!cantidad)) & ")"
'
'                    Conn.Execute Sql
'
'                    Rs2.MoveNext
'                Wend
'                Rs2.Close
'
'                ' borramos sublineas
'                Sql = "delete from slialb1 where tipofact = 1 and numserie = '" & DevNombreSQL(Rs!numserie)
'                Sql = Sql & "' and numalbar = " & Rs!NumAlbar & " and "
'                Sql = Sql & " numlinea = " & RS1!Numlinea
'                Conn.Execute Sql
'                RS1.MoveNext
'            Wend
'            RS1.Close
'
'            'añadido los calculos de totales
'
'
'            ' borramos lineas
'            Sql1 = "delete from slialb where tipofact = 1 and numserie = '" & DevNombreSQL(Rs!numserie) & "' and "
'            Sql1 = Sql1 & " numalbar = " & Rs!NumAlbar
'            Conn.Execute Sql1
'
'            ' borramos albaran
'            Sql1 = "delete from scaalb where tipofact = 1 and numserie = '" & DevNombreSQL(Rs!numserie) & "' and "
'            Sql1 = Sql1 & " numalbar = " & Rs!NumAlbar
'            Conn.Execute Sql1
'
'            'Progress
'            pb2.Value = pb2.Value + 1
'            pb2.Refresh
'            'Siguiente albaran
'            Rs.MoveNext
'        Wend
'
'        'InsertarTotales
'        ' insertamos en la tabla de totales de factura
'        ' calculo de importes para insertar
'        ' una vez tenemos cargados los brutos de la factura calculamos bases e ivas
'        InsertarTotales Combo2(0).Text, mC.Contador, Text3(9).Text
'
'
'
'        HastaFactura = mC.Contador
'
'        Set mC = Nothing
'
'        InsertarCabeceraFactura DesdeFactura, HastaFactura, Combo2(0).Text, CDate(Text3(9).Text), CDate(Text3(9).Text)
'
'
'    End If
'EFacturacionAlbaranes:
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " - " & Err.Description, vbExclamation
'
'        For v_aux = HastaFactura To DesdeFactura Step -1
'            mC.DevolverContador Combo2(0).Text, "2", v_aux
'        Next v_aux
'
'        Conn.RollbackTrans
'    Else
'        Conn.CommitTrans
'
'        NomDocu = ""
'        cadparam = ""
'        numParam = 0
'        indRPT = 3
'
'        If Not PonerParamEmpresa(indRPT, cadparam, numParam, NomDocu) Then
'            Exit Sub
'        End If
'
'        Sql = "desde= " & DesdeFactura & "|hasta= " & HastaFactura & "|"
'        Sql = Sql & "serie= """ & Combo2(0).Text & """|"
'        Sql = Sql & "desdefec= """ & Text3(9).Text & """|hastafec= """ & Text3(9).Text & """|"
'
'        frmImprimir.Opcion = 17
'        frmImprimir.NumeroParametros = 5
'        frmImprimir.NomDocu = NomDocu
'        frmImprimir.FormulaSeleccion = Sql
'        frmImprimir.OtrosParametros = Sql
'        frmImprimir.SoloImprimir = False
'        frmImprimir.Show 'vbModal
'    End If
'End Sub



'Private Sub FamiliasAceptar_Click()
''Imprimir el listado, segun sea
'    If Not ComprobarFamilias(0, 1) Then Exit Sub
'    'Hacemos el select y si tiene resultados mostramos los valores
'
'    Cad = " SELECT tmpsfamia.* from tmpsfamia WHERE codusu = " & vUsu.codigo
'    If txtFam(0).Text <> "" Then Cad = Cad & " AND codfamia >= " & txtFam(0).Text
'    If txtFam(1).Text <> "" Then Cad = Cad & " AND codfamia <= " & txtFam(1).Text
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'    If Rs.EOF Then
'        'NO hay registros a mostrar
'        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
'    Else
'        'Mostramos el frame de resultados
'        Sql = "usu= " & vUsu.codigo & "|"
'        Cad1 = "Familias: "
'        If txtFam(0).Text <> "" Then
'            Sql = Sql & "desfam= " & txtFam(0).Text & "|"
'            Cad1 = Cad1 & " desde " & txtFam(0).Text & " " & DtxtFam(0).Text
'        End If
'
'        If txtFam(1).Text <> "" Then
'            Sql = Sql & "hasfam= " & txtFam(1).Text & "|"
'            Cad1 = Cad1 & "   hasta " & txtFam(1).Text & " " & DtxtFam(1).Text
'        End If
'
'        If txtFam(0).Text <> "" Or txtFam(1).Text <> "" Then Sql = Sql & "Familias= """ & Cad1 & """|"
'
'        If optListFam(0).Value = True Then
'            Sql = Sql & "orden= ""Por Código""|"
'            frmImprimir.CampoOrden = 1
'        End If
'        If optListFam(1).Value = True Then
'            Sql = Sql & "orden= ""Alfabético""|"
'            frmImprimir.CampoOrden = 2
'        End If
'
'        frmImprimir.Opcion = 1
'        frmImprimir.NumeroParametros = 3
'        frmImprimir.FormulaSeleccion = Sql
'        frmImprimir.OtrosParametros = Sql
'        frmImprimir.SoloImprimir = False
'        frmImprimir.Show 'vbModal
'     End If
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)

    Select Case Opcion
        Case 9
'            If frmVisReport.EstaImpreso Then
'              If MsgBox("Impresión correcta para actualizar", vbYesNoCancel + vbDefaultButton2) = vbYes Then 'VRS:1.0.1(11)
'                  If txtFam(4) = "" Then txtFam(4).Text = 0
'                  If txtFam(5) = "" Then txtFam(5).Text = 999
'                  If txtArt(4) = "" Then txtArt(4).Text = 0
'                  If txtArt(5) = "" Then txtArt(5).Text = 999999
'
'                  If InsertarTomaInventario(CInt(txtFam(4).Text), CInt(txtFam(5).Text), CLng(txtArt(4).Text), CLng(txtArt(5).Text), Text3(5).Text) Then
'                      Sql = "Proceso finalizado. Puede proceder a realizar la entrada"
'                      Sql = Sql & " de existencia real."
'                      MsgBox Sql, vbExclamation
'                  End If
'              End If
'            End If
        Case 13
'            If BloqueoManual(False, "FACALB", "Factur") Then
'
'            End If
            
        Case 16
'            'VRS:1.0.5(9) desbloqueamos el cuadre de caja
'            If BloqueoManual(False, "CUADRE", "CAJA") Then
'
'            End If
'        Case 18
'            If BloqueoManual(False, "CAMPRE", "Cambio") Then
'
'            End If
    End Select
End Sub

Private Sub frmEmp_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtEmp(Empresa).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.DtxtEmp(Empresa).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtPro(Empresa).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.DtxtPro(Empresa).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ImgEmp_Click(Index As Integer)
    Empresa = Index
    Set frmEmp = New frmEmpresas
    frmEmp.DatosADevolverBusqueda = "0|1|"
    frmEmp.Show
End Sub

Private Sub Imgpro_Click(Index As Integer)
    Empresa = Index
    Set frmPro = New frmProvincias
    frmPro.DatosADevolverBusqueda = "0|1|"
    frmPro.Show
End Sub

Private Sub Txtope_GotFocus(Index As Integer)
    txtOpe(Index).SelStart = 0
    txtOpe(Index).SelLength = Len(txtOpe(Index).Text)
End Sub

Private Sub txtEmp_GotFocus(Index As Integer)
    txtEmp(Index).SelStart = 0
    txtEmp(Index).SelLength = Len(txtEmp(Index).Text)
End Sub

Private Sub txtOpe_LostFocus(Index As Integer)
Dim ape1 As String
Dim ape2 As String
Dim nombre As String

    txtOpe(Index).Text = Trim(txtOpe(Index).Text)
    If txtOpe(Index).Text = "" Then
        DtxtOpe(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtOpe(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtOpe(Index).Text = Replace(Format(txtOpe(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtOpe(Index)
        Exit Sub
    End If
    
    txtOpe(Index).Text = Format(txtOpe(Index).Text, ">")

    sql = DevuelveDesdeBD(1, "dni", "operarios", "dni|", txtOpe(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        CargarDatosOperarios Trim(txtOpe(Index).Text), ape1, ape2, nombre
        sql = Trim(ape1) & " " & Trim(ape2) & ", " & Trim(nombre)
        DtxtOpe(Index).Text = sql
    Else
        DtxtOpe(Index).Text = "No existe el DNI de Operario"
    End If

End Sub

Private Sub txtPro_GotFocus(Index As Integer)
    txtPro(Index).SelStart = 0
    txtPro(Index).SelLength = Len(txtPro(Index).Text)
End Sub

Private Sub txtIns_GotFocus(Index As Integer)
    txtIns(Index).SelStart = 0
    txtIns(Index).SelLength = Len(txtIns(Index).Text)
End Sub

Private Sub cmdCanLisEmp_Click()
    Unload Me
End Sub

'Private Sub ArticulosAceptar_Click()
'Dim NomDocu As String
'Dim indRPT As Integer
'Dim cadparam As String
'Dim numParam As Integer
'
''Imprimir el listado, segun sea
'
'    If Not ComprobarFamilias(2, 3) Then Exit Sub
'    If Not ComprobarArticulos(0, 1) Then Exit Sub
'
'    'Hacemos el select y si tiene resultados mostramos los valores
'    Cad = " SELECT sartic.* from sartic WHERE 1 = 1 "
'    If txtFam(2).Text <> "" Then Cad = Cad & " AND codfamia >= " & txtFam(2).Text
'    If txtFam(3).Text <> "" Then Cad = Cad & " AND codfamia <= " & txtFam(3).Text
'    If txtArt(0).Text <> "" Then Cad = Cad & " AND codartic >= " & txtArt(0).Text
'    If txtArt(1).Text <> "" Then Cad = Cad & " AND codartic <= " & txtArt(1).Text
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'    If Rs.EOF Then
'        'NO hay registros a mostrar
'        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
'    Else
'        'Mostramos el frame de resultados
'        Sql = ""
'        Cad1 = "Familias: "
'        If txtFam(2).Text <> "" Then
'            Sql = Sql & "desfam= " & txtFam(2).Text & "|"
'            Cad1 = Cad1 & " desde " & txtFam(2).Text & " " & DtxtFam(2).Text
'        End If
'
'        If txtFam(3).Text <> "" Then
'            Sql = Sql & "hasfam= " & txtFam(3).Text & "|"
'            Cad1 = Cad1 & "   hasta " & txtFam(3).Text & " " & DtxtFam(3).Text
'        End If
'
'        If txtFam(2).Text <> "" Or txtFam(3).Text <> "" Then Sql = Sql & "Familias= """ & Cad1 & """|"
'
'        Cad1 = "Articulos: "
'        If txtArt(0).Text <> "" Then
'            Sql = Sql & "desart= " & txtArt(0).Text & "|"
'            Cad1 = Cad1 & " desde " & txtArt(0).Text & " " & DtxtArt(0).Text
'        End If
'
'        If txtArt(1).Text <> "" Then
'            Sql = Sql & "hasart= " & txtArt(1).Text & "|"
'            Cad1 = Cad1 & "   hasta " & txtArt(1).Text & " " & DtxtArt(1).Text
'        End If
'
'        If txtArt(0).Text <> "" Or txtArt(1).Text <> "" Then Sql = Sql & "Articulos= """ & Cad1 & """|"
'
'
'        If optListArt(0).Value = True Then
'            Sql = Sql & "orden= ""Por Código""|"
'            frmImprimir.CampoOrden = 1
'        End If
'
'        If optListArt(1).Value = True Then
'            Sql = Sql & "orden= ""Alfabético""|"
'            frmImprimir.CampoOrden = 2
'        End If
'
'        If Me.Opcion = 2 Then
'            frmImprimir.NomDocu = "Articulos.rpt"
'        End If
'        If Me.Opcion = 17 Then
'            'VRS:1.0.4(2)
'            NomDocu = ""
'            cadparam = ""
'            numParam = 0
'            indRPT = 4
'
'            If Not PonerParamEmpresa(indRPT, cadparam, numParam, NomDocu) Then
'                Exit Sub
'            End If
'            frmImprimir.NomDocu = NomDocu
'        End If
'
'
'        frmImprimir.Opcion = 2
'        frmImprimir.NumeroParametros = 7
'        frmImprimir.FormulaSeleccion = Sql
'        frmImprimir.OtrosParametros = Sql
'        frmImprimir.SoloImprimir = False
'        frmImprimir.Show 'vbModal
'        Screen.MousePointer = vbDefault
'
'     End If
'End Sub

Private Sub frmIns_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtIns(instalacion).Text = RecuperaValor(CadenaSeleccion, 2)
    Me.DtxtIns(instalacion).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub ImgIns_Click(Index As Integer)
    instalacion = Index
    Set frmIns = New frmInstalaciones
    frmIns.DatosADevolverBusqueda = "0|13|1|"
    frmIns.Show
End Sub

Private Sub txtemp_LostFocus(Index As Integer)
Dim sql As String
    
    txtEmp(Index).Text = Trim(txtEmp(Index).Text)
    If txtEmp(Index).Text = "" Then
        DtxtEmp(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtEmp(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtEmp(Index).Text = Replace(Format(txtEmp(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtEmp(Index)
        Exit Sub
    End If
    txtEmp(Index).Text = Format(txtEmp(Index).Text, ">")


    sql = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", txtEmp(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        DtxtEmp(Index).Text = sql
    Else
        DtxtEmp(Index).Text = "No existe la empresa"
    End If

End Sub

Private Sub Image2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.fecha = Now
    If Text3(Index).Text <> "" Then frmC.fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub frmOpe_DatoSeleccionado(CadenaSeleccion As String)
Dim ape1 As String
Dim ape2 As String
Dim nombre As String

    ape1 = ""
    ape2 = ""
    nombre = ""
    Me.txtOpe(Operario).Text = RecuperaValor(CadenaSeleccion, 1)
    CargarDatosOperarios Me.txtOpe(Operario).Text, ape1, ape2, nombre
    Me.DtxtOpe(Operario).Text = Trim(ape1) & " " & Trim(ape2) & ", " & Trim(nombre)
End Sub


Private Sub Form_Activate()
Dim mesAnt As Integer
Dim anoAnt As Integer

    If PrimeraVez Then
        PrimeraVez = False
        'Ponemos el foco
        Select Case Opcion
        Case 1  'Listado de empresas
            txtEmp(0).SetFocus
        Case 2  ' Listado de instalaciones
            txtEmp(2).SetFocus
        Case 3  ' Listado de operarios
            txtEmp(4).SetFocus
        Case 13
            txtPro(0).SetFocus
        Case 14
            txtTMe(0).SetFocus
        Case 15
            txtRGe(0).SetFocus
        Case 20
        End Select
    End If
        Screen.MousePointer = vbDefault
End Sub
'
Private Sub Form_Load()
Dim H As Single
Dim W As Single
Dim Acabar As Boolean

    Me.Top = 0
    Me.Left = 0

    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    Limpiar Me
    NoExistenDatos = True
    FrameListEmpresas.Visible = False
    FrameListInstalaciones.Visible = False
    FrameListOperarios.Visible = False
    FrameListDosimetros.Visible = False
    FrameListFactCalib.Visible = False
    FrameDosisColectiva.Visible = False
    FrameListDosisInstal.Visible = False
    FrameListProvincias.Visible = False
    FrameListTipoMedicion.Visible = False
    FrameListRamasGenericas.Visible = False
    FrameListFondos.Visible = False
    FrameListRamasEspec.Visible = False
    FrameListTiposTrabajo.Visible = False
    FrameListDosisNHomOpe.Visible = False
    FrameCartaDosimNRec.Visible = False
    FrameListDosisOpeAcum12.Visible = False
    FrameCartaSobredosis.Visible = False
    FrameListRecepDosimCuerpo.Visible = False
    FrameListOperariosSobredosis.Visible = False
    FrameListUsu.Visible = False
    FrameListLotes.Visible = False
    
    Acabar = False
    Select Case Opcion
    Case 1
         ' listado de Empresas
        CargarCombo
        
        Me.FrameListEmpresas.Visible = True
        W = Me.FrameListEmpresas.Width
        H = Me.FrameListEmpresas.Height
         
    Case 2
        'Listado de Instalaciones
        CargarCombo
        
        Me.FrameListInstalaciones.Visible = True
        Me.Label2(22).Caption = "Listado de Instalaciones"
        W = Me.FrameListInstalaciones.Width
        H = Me.FrameListInstalaciones.Height
    
    Case 3
        'Listado de Operarios en instalaciones
        Me.FrameListOperarios.Visible = True
        W = Me.FrameListOperarios.Width
        H = Me.FrameListOperarios.Height
    
    Case 4
        'Listado de Dosimetros a cuerpo
        CargarCombo
        Combo3.ListIndex = 0
        Combo5.ListIndex = 0
        
        Me.FrameListDosimetros.Visible = True
        Me.Label2(0).Caption = "Listado de Dosímetros"
        W = Me.FrameListDosimetros.Width
        H = Me.FrameListDosimetros.Height
    
    Case 5
        'Listado de Dosimetros a organo
        Me.FrameListDosimetros.Visible = True
        Me.Label2(0).Caption = "Listado de Dosímetros a Órgano"
        W = Me.FrameListDosimetros.Width
        H = Me.FrameListDosimetros.Height
    
    Case 7
        'Listado de Factores de calibracion 4400
        Me.FrameListFactCalib.Visible = True
        Me.Label2(2).Caption = "Listado de Factores de Calibración 4400"
        W = Me.FrameListFactCalib.Width
        H = Me.FrameListFactCalib.Height
        
    Case 8
        'Listado de Factores de calibracion 6600
        Me.FrameListFactCalib.Visible = True
        Me.Label2(2).Caption = "Listado de Factores de Calibración 6600"
        W = Me.FrameListFactCalib.Width
        H = Me.FrameListFactCalib.Height
    
    Case 9
        'Listado de Dosis por instalacion
        Me.FrameListDosisInstal.Visible = True
        Me.Label2(2).Caption = "Listado de Dosis por Instalación"
        W = Me.FrameListDosisInstal.Width
        H = Me.FrameListDosisInstal.Height
    
    Case 12
        'Listado de Dosis CSN
        Me.FrameDosisColectiva.Visible = True
        Me.Label2(2).Caption = "Listado de Dosis Colectiva CSN"
        W = Me.FrameDosisColectiva.Width
        H = Me.FrameDosisColectiva.Height
    
    Case 13
        'Listado de provincias
        Me.FrameListProvincias.Visible = True
        W = Me.FrameListProvincias.Width
        H = Me.FrameListProvincias.Height
    
    Case 14
        'Listado de tipos de medicion
        Me.FrameListTipoMedicion.Visible = True
        W = Me.FrameListTipoMedicion.Width
        H = Me.FrameListTipoMedicion.Height
    
    Case 15
        'Listado de ramas genericas
        Me.FrameListRamasGenericas.Visible = True
        W = Me.FrameListRamasGenericas.Width
        H = Me.FrameListRamasGenericas.Height
    
    Case 16
        'Listado de ramas especificas
        Me.FrameListRamasEspec.Visible = True
        W = Me.FrameListRamasEspec.Width
        H = Me.FrameListRamasEspec.Height
    
    Case 17
        'Listado de tipos de trabajo
        Me.FrameListTiposTrabajo.Visible = True
        W = Me.FrameListTiposTrabajo.Width
        H = Me.FrameListTiposTrabajo.Height
    
    
    Case 18
        'Listado de fondos
        Me.Label2(9).Caption = "Listado de Fondos Harshaw 6600"
        Me.FrameListFondos.Visible = True
        W = Me.FrameListFondos.Width
        H = Me.FrameListFondos.Height
    
    Case 19
        'Listado de dosis no homogenea por operario
        Me.FrameListDosisNHomOpe.Visible = True
        W = Me.FrameListDosisNHomOpe.Width
        H = Me.FrameListDosisNHomOpe.Height
    
    Case 20
        ' cartas de dosimetros no recibidos por instalación
        Me.FrameCartaDosimNRec.Visible = True
        W = Me.FrameCartaDosimNRec.Width
        H = Me.FrameCartaDosimNRec.Height
    
        Dim mesAnt As Integer
        Dim anoAnt As Integer
    
        mesAnt = Month(Now) - 1
        If mesAnt < 1 Then
            mesAnt = 12
            anoAnt = Year(Now) - 1
        Else
            anoAnt = Year(Now)
        End If
        Text3(18).Text = "01/" & Format(mesAnt, "00") & "/" & Format(anoAnt, "0000")
        
        Text3(19).Text = Format(CDate("01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")) - 1, "dd/mm/yyyy")
    
    Case 21
        ' listado de dosis operarios acumulada 12 meses
        
        Me.FrameListDosisOpeAcum12.Visible = True
        CartaSobredosis = False
        Me.Label2(14).Caption = "Dosis por Operario Año Oficial"
        W = Me.FrameListDosisOpeAcum12.Width
        H = Me.FrameListDosisOpeAcum12.Height
    
        Text4(0).Text = Format(Year(Now), "0000")
        
    Case 22
        'carta de sobredosis al consejo

'        Me.FrameCartaSobredosis.Visible = True
'        W = Me.FrameCartaSobredosis.Width
'        H = Me.FrameCartaSobredosis.Height
'        Text3(22).Text = "01/01/" & Format(Year(Now), "00")
'        Text3(23).Text = "31/12/" & Format(Year(Now), "00")
        
        Me.FrameListDosisOpeAcum12.Visible = True
        Me.Label2(14).Caption = "Carta CSN de Potencial de Sobredosis"
        CartaSobredosis = True
        W = Me.FrameListDosisOpeAcum12.Width
        H = Me.FrameListDosisOpeAcum12.Height
    
        Text4(0).Text = Format(Year(Now), "0000")
        
        
        
    Case 23
        'listado de recepcion de dosimetros
        Me.Label2(16).Caption = "Informe de Recepción de Dosímetros a Cuerpo"
        Me.FrameListRecepDosimCuerpo.Visible = True
        W = Me.FrameListRecepDosimCuerpo.Width
        H = Me.FrameListRecepDosimCuerpo.Height
    
        Text3(24).Text = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "00")
        Text3(25).Text = Format(Now, "dd/mm/yyyy")
        
        CargarCombo
        Combo4.ListIndex = 2
        Check1(4).Value = 1
    Case 24
         ' listado de etiquetas de Empresas
        CargarCombo
        
        
        Me.FrameListEmpresas.Visible = True
        W = Me.FrameListEmpresas.Width
        H = Me.FrameListEmpresas.Height
        
    Case 25
        'Listado de Instalaciones
        CargarCombo
        
        Me.FrameListInstalaciones.Visible = True
        Me.Label2(22).Caption = "Listado de Instalaciones"
        W = Me.FrameListInstalaciones.Width
        H = Me.FrameListInstalaciones.Height
    
    Case 26
        'Listado de Operarios
        Me.FrameListOperarios.Visible = True
        W = Me.FrameListOperarios.Width
        H = Me.FrameListOperarios.Height
        
    Case 27
        ' listado de dosimetros de area recepcionados
        'listado de recepcion de dosimetros
        Me.Label2(16).Caption = "Informe de Recepción de Dosímetros Area"
        Me.FrameListRecepDosimCuerpo.Visible = True
        W = Me.FrameListRecepDosimCuerpo.Width
        H = Me.FrameListRecepDosimCuerpo.Height
    
        Text3(24).Text = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "00")
        Text3(25).Text = Format(Now, "dd/mm/yyyy")
        
        CargarCombo
        Combo4.ListIndex = 2
        Check1(4).Value = 1
    
    Case 28
        ' listado de operarios con sobredosis
        Me.FrameListOperariosSobredosis.Visible = True
        W = Me.FrameListOperariosSobredosis.Width
        H = Me.FrameListOperariosSobredosis.Height
        
        Text3(21).Text = Format(Now, "dd/mm/yyyy")
    
    Case 29
        ' lotes
        Me.Label2(19).Caption = "Listado de Lotes Harshaw 6600"
        Me.FrameListLotes.Visible = True
        W = Me.FrameListLotes.Width
        H = Me.FrameListLotes.Height
        
    Case 30
        ' lotes panasonic
        Me.Label2(19).Caption = "Listado de Lotes Panasonic"
        Me.FrameListLotes.Visible = True
        W = Me.FrameListLotes.Width
        H = Me.FrameListLotes.Height
    
    Case 31
        'Listado de Factores de calibracion panasonic
        Me.Label2(2).Caption = "Listado de Factores de Calibración Panasonic"
        Me.FrameListFactCalib.Visible = True
        W = Me.FrameListFactCalib.Width
        H = Me.FrameListFactCalib.Height
    
    Case 32
        'Listado de fondos
        Me.Label2(9).Caption = "Listado de Fondos Panasonic"
        Me.FrameListFondos.Visible = True
        W = Me.FrameListFondos.Width
        H = Me.FrameListFondos.Height
    
    End Select
    
    Me.Width = W + 240
    Me.Height = H + 400
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub
Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation, "¡Error!"
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    Else
        'si index = 10 caso del listado csn
        If Index = 10 Then
            Text3(11).Text = CargarFechaHasta(CDate(Text3(10).Text), 1)
        End If
        If Index = 18 Then
            Text3(19).Text = CargarFechaHasta(CDate(Text3(Index).Text), 2)
        End If
    End If
End Sub

Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
            MsgBox "Fecha 'desde' mayor que fecha 'hasta'.", vbExclamation, "¡Error!"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function

Private Function ComprobarOperarios(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarOperarios = False
    If txtOpe(Indice1).Text <> "" And txtOpe(Indice2).Text <> "" Then
        If Trim(txtOpe(Indice1).Text) > Trim(txtOpe(Indice2).Text) Then
            MsgBox "Operario 'desde' mayor que operario 'hasta'.", vbExclamation, "¡Error!"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    ComprobarOperarios = True
End Function

Private Function ComprobarDosimetros(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarDosimetros = False
    If txtDos(Indice1).Text <> "" And txtDos(Indice2).Text <> "" Then
        If Trim(txtDos(Indice1).Text) > Trim(txtDos(Indice2).Text) Then
            MsgBox "Dosímetro 'desde' mayor que dosímetro 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarDosimetros = True
End Function


Private Function ComprobarEmpresas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarEmpresas = False
    If txtEmp(Indice1).Text <> "" And txtEmp(Indice2).Text <> "" Then
        If Trim(txtEmp(Indice1).Text) > Trim(txtEmp(Indice2).Text) Then
            MsgBox "Empresa 'desde' mayor que empresa 'hasta'.", vbExclamation, "¡Error!"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    ComprobarEmpresas = True
End Function

Private Function ComprobarProvincias(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarProvincias = False
    If txtPro(Indice1).Text <> "" And txtPro(Indice2).Text <> "" Then
        If Trim(txtPro(Indice1).Text) > Trim(txtPro(Indice2).Text) Then
            MsgBox "Provincia 'desde' mayor que provincia 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarProvincias = True
End Function

Private Function ComprobarRamasGenericas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarRamasGenericas = False
    If txtRGe(Indice1).Text <> "" And txtRGe(Indice2).Text <> "" Then
        If Trim(txtRGe(Indice1).Text) > Trim(txtRGe(Indice2).Text) Then
            MsgBox "Ramas Genéricas 'desde' mayor que ramas genéricas 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarRamasGenericas = True
End Function

Private Function ComprobarRamasEspecificas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarRamasEspecificas = False
    If txtREs(Indice1).Text <> "" And txtREs(Indice2).Text <> "" Then
        If Trim(txtREs(Indice1).Text) > Trim(txtREs(Indice2).Text) Then
            MsgBox "Ramas específicas 'desde' mayor que ramas específicas 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarRamasEspecificas = True
End Function

Private Function ComprobarTipoTrab(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarTipoTrab = False
    If txtREs(Indice1).Text <> "" And txtREs(Indice2).Text <> "" Then
        If Trim(txtREs(Indice1).Text) > Trim(txtREs(Indice2).Text) Then
            MsgBox "Tipo de trabajo 'desde' mayor que tipo de trabajo 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarTipoTrab = True
End Function


Private Function ComprobarTiposMedicion(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarTiposMedicion = False
    If txtTMe(Indice1).Text <> "" And txtTMe(Indice2).Text <> "" Then
        If Trim(txtTMe(Indice1).Text) > Trim(txtTMe(Indice2).Text) Then
            MsgBox "Tipo de medición 'desde' mayor que tipo de medición 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarTiposMedicion = True
End Function



Private Function ComprobarInstalaciones(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarInstalaciones = False
    If txtIns(Indice1).Text <> "" And txtIns(Indice2).Text <> "" Then
        If Trim(txtIns(Indice1).Text) > Trim(txtIns(Indice2).Text) Then
            MsgBox "Instalación 'desde' mayor que instalación 'hasta'.", vbExclamation, "¡Error!"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    ComprobarInstalaciones = True
End Function


'Private Sub CargarCombo2(Index As Integer)
'Dim Rs As Recordset
'Dim Sql As String
''###
''Cargaremos el combo, o bien desde una tabla o con valores fijos o como
''se quiera, la cuestion es cargarlo
'' El estilo del combo debe de ser 2 - Dropdown List
'' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
'' o marcamos la opcion sorted del combo
''0-Si, 1-No
'
'    Combo2(Index).Clear
'
'    Set Rs = New ADODB.Recordset
'    Sql = "select numserie from contadores where tipserie = 2 "
'    Rs.Open Sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
'    Rs.MoveFirst
'    i = 0
'    While Not Rs.EOF
'        Combo2(Index).AddItem Rs.Fields!numserie
'        Combo2(Index).ItemData(Combo2(Index).NewIndex) = i
'        i = i + 1
'        Rs.MoveNext
'    Wend
'    Rs.Close
'
'End Sub
Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CargarFechaHasta(desde As Date, Opcion As Integer) As Date
Dim Dia As Integer
Dim Mes As Integer
Dim ano As Integer

    Select Case Opcion
        Case 1

            ano = Year(desde)
            Mes = Month(desde)
            If OptIns(7).Value Then
                Mes = Mes + 1
                If Mes > 12 Then
                    Mes = 1
                    ano = ano + 1
                End If
            ElseIf OptIns(6).Value Then
                Mes = Mes + 6
                If Mes > 12 Then
                    Mes = Mes - 12
                    ano = ano + 1
                End If
            ElseIf OptIns(5).Value Then
                ano = ano + 1
            End If
            CargarFechaHasta = (CDate(Day(desde) & "/" & Format(Mes, "00") & "/" & Format(ano, "0000")) - 1)
        Case 2
            ano = Year(desde)
            Mes = Month(desde)
            Mes = Mes + 1
            If Mes > 12 Then
                Mes = 1
                ano = ano + 1
            End If
            CargarFechaHasta = (CDate(Day(desde) & "/" & Format(Mes, "00") & "/" & Format(ano, "0000")) - 1)
    End Select
End Function

Private Sub txtPro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtPro_LostFocus(Index As Integer)
    txtPro(Index).Text = Trim(txtPro(Index).Text)
    If txtPro(Index).Text = "" Then
        DtxtPro(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtPro(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtPro(Index).Text = Replace(Format(txtPro(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtPro(Index)
        Exit Sub
    End If

    txtPro(Index).Text = Format(txtPro(Index).Text, ">")

    sql = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", txtPro(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        DtxtPro(Index).Text = sql
    Else
        DtxtPro(Index).Text = "No existe la provincia"
    End If

End Sub




Private Sub txtREs_GotFocus(Index As Integer)
    txtREs(Index).SelStart = 0
    txtREs(Index).SelLength = Len(txtREs(Index).Text)
End Sub

Private Sub txtREs_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtREs_LostFocus(Index As Integer)
    txtREs(Index).Text = Trim(txtREs(Index).Text)
    If txtREs(Index).Text = "" Then
        DtxtREs(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtREs(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtREs(Index).Text = Replace(Format(txtREs(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtREs(Index)
        Exit Sub
    End If
    txtREs(Index).Text = Format(txtREs(Index).Text, ">")

    sql = DevuelveDesdeBD(1, "descripcion", "ramaespe", "c_rama_especifica|", txtREs(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        DtxtREs(Index).Text = sql
    Else
        DtxtREs(Index).Text = "No existe la rama específica."
    End If


End Sub

Private Sub txtRGe_GotFocus(Index As Integer)
    txtRGe(Index).SelStart = 0
    txtRGe(Index).SelLength = Len(txtRGe(Index).Text)
End Sub

Private Sub txtRGe_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtRGe_LostFocus(Index As Integer)
    txtRGe(Index).Text = Trim(txtRGe(Index).Text)
    If txtRGe(Index).Text = "" Then
        DtxtRGe(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtRGe(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtRGe(Index).Text = Replace(Format(txtRGe(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtRGe(Index)
        Exit Sub
    End If
    txtRGe(Index).Text = Format(txtRGe(Index).Text, ">")

    sql = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", txtRGe(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        DtxtRGe(Index).Text = sql
    Else
        DtxtRGe(Index).Text = "No existe la rama genérica."
    End If

End Sub

Private Sub txtTMe_GotFocus(Index As Integer)
    txtTMe(Index).SelStart = 0
    txtTMe(Index).SelLength = Len(txtTMe(Index).Text)
End Sub

Private Sub txtTMe_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtTMe_LostFocus(Index As Integer)
    txtTMe(Index).Text = Trim(txtTMe(Index).Text)
    If txtTMe(Index).Text = "" Then
        DtxtTMe(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtTMe(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtTMe(Index).Text = Replace(Format(txtTMe(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtTMe(Index)
        Exit Sub
    End If

    txtTMe(Index).Text = Format(txtTMe(Index).Text, ">")
    
    sql = DevuelveDesdeBD(1, "descripcion", "tipmedext", "c_tipo_med|", txtTMe(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        DtxtTMe(Index).Text = sql
    Else
        DtxtTMe(Index).Text = "No existe el tipo de medición"
    End If

End Sub

Private Sub CargarCombo()
' cargamos el tipo de dosimetro
    
    Combo3.Clear
    Combo3.AddItem "Cuerpo"
    Combo3.ItemData(Combo3.NewIndex) = 0
    
    Combo3.AddItem "Organo"
    Combo3.ItemData(Combo3.NewIndex) = 1
    
    Combo3.AddItem "Area"
    Combo3.ItemData(Combo3.NewIndex) = 2
    
    Combo3.AddItem "Todos"
    Combo3.ItemData(Combo3.NewIndex) = 3
    

    Combo1.Clear
    Combo1.AddItem "Personal"
    Combo1.ItemData(Combo1.NewIndex) = 0

    Combo1.AddItem "Area"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "Todas"
    Combo1.ItemData(Combo1.NewIndex) = 2

    Combo2.Clear
    Combo2.AddItem "Personal"
    Combo2.ItemData(Combo2.NewIndex) = 0

    Combo2.AddItem "Area"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    Combo2.AddItem "Todas"
    Combo2.ItemData(Combo2.NewIndex) = 2
    Combo2.ListIndex = 2


    Combo4.Clear
    Combo4.AddItem "Par"
    Combo4.ItemData(Combo4.NewIndex) = 0

    Combo4.AddItem "Impar"
    Combo4.ItemData(Combo4.NewIndex) = 1
    
    Combo4.AddItem "Ambas"
    Combo4.ItemData(Combo4.NewIndex) = 2

    Combo5.Clear
    Combo5.AddItem "Todos"
    Combo5.AddItem "Harshaw"
    Combo5.AddItem "Panasonic"
    
End Sub

Private Sub txtTTr_GotFocus(Index As Integer)
    txtTTr(Index).SelStart = 0
    txtTTr(Index).SelLength = Len(txtTTr(Index).Text)
End Sub

Private Sub txtTTr_LostFocus(Index As Integer)
    txtTTr(Index).Text = Trim(txtTTr(Index).Text)
    If txtTTr(Index).Text = "" Then
        DtxtTTr(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtTTr(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        txtTTr(Index).Text = Replace(Format(txtTTr(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco txtTTr(Index)
        Exit Sub
    End If
    txtTTr(Index).Text = Format(txtTTr(Index).Text, ">")

    sql = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", txtTTr(Index).Text & "|", "T|", 1)
    If sql <> "" Then
        DtxtTTr(Index).Text = sql
    Else
        DtxtTTr(Index).Text = "No existe el tipo de trabajo."
    End If


End Sub


Private Sub CargaAcumuladosQuinquenales()
Dim anoini As Integer
Dim fecini As String
Dim rs As ADODB.Recordset
Dim rL As ADODB.Recordset
Dim sql1 As String
Dim sql2 As String
Dim Cad As String
Dim dosissuper As Currency
Dim dosisprofu As Currency
Dim ano As Integer
Dim Sist As String

  ano = CInt(Text4(0).Text)
    
  Sist = IIf(OptSist(2).Value, "H", "P")
  
  
    ' NUEVA HISTORIA para evitar los (campo1,campo2,campo3) IN (SELECT campo1,campo2...)
  Cad = "SELECT DISTINCT c1.c_empresa,c1.c_instalacion,c1.dni_usuario,c1.n_dosimetro,"
  Cad = Cad & "c1.mes_p_i,c1.f_asig_dosimetro,c1.f_retirada,c1.n_reg_dosimetro FROM "
  Cad = Cad & "(SELECT DISTINCT dosimetros.c_empresa,dosimetros.c_instalacion,dosimetros.dni_usuario,"
  Cad = Cad & "dosimetros.n_dosimetro,dosimetros.mes_p_i,f_asig_dosimetro,f_retirada,"
  Cad = Cad & "dosimetros.n_reg_dosimetro FROM operarios"
  
  ' Solo el fichero migrado
  If Check1(1).Value = 1 Then
    Cad = Cad & " INNER JOIN "
    Cad = Cad & "(SELECT DISTINCT dosimetros.dni_usuario dni FROM dosimetros,tempnc WHERE "
    Cad = Cad & "dosimetros.n_dosimetro=tempnc.n_dosimetro AND dosimetros.f_retirada IS NULL AND codusu="
    Cad = Cad & vUsu.codigo & ") t1 USING(dni),dosimetros INNER JOIN "
    Cad = Cad & "(SELECT DISTINCT dni_usuario,c_empresa,c_instalacion FROM dosimetros,tempnc "
    Cad = Cad & "WHERE dosimetros.n_dosimetro=tempnc.n_dosimetro AND dosimetros.f_retirada IS NULL AND codusu="
    Cad = Cad & vUsu.codigo & ") t2 USING(dni_usuario,c_empresa,c_instalacion)"
  Else
    Cad = Cad & ",dosimetros"
  End If
  
  Cad = Cad & " WHERE operarios.semigracsn=1 AND operarios.dni=dosimetros.dni_usuario AND "
  Cad = Cad & "dosimetros.tipo_dosimetro=0 AND dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%'"
  'Cad = Cad & " AND dosimetros.sistema='" & Sist & "'" (rafa) VRS 1.3.6
  If Check1(0).Value = 1 Then Cad = Cad & " AND operarios.f_baja IS NULL"
  If txtEmp(12).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa>='" & Trim(txtEmp(12).Text) & "'"
  If txtEmp(13).Text <> "" Then Cad = Cad & " AND dosimetros.c_empresa<='" & Trim(txtEmp(13).Text) & "'"
  If txtOpe(6).Text <> "" Then Cad = Cad & " AND dosimetros.dni_usuario>='" & Trim(txtOpe(6).Text) & "'"
  If txtOpe(7).Text <> "" Then Cad = Cad & " AND dosimetros.dni_usuario<='" & Trim(txtOpe(7).Text) & "'"
  Cad = Cad & ") c1 INNER JOIN (SELECT * FROM dosiscuerpo WHERE YEAR(f_dosis)=" & ano
  Cad = Cad & ") c2 USING(n_dosimetro,n_reg_dosimetro)"

  
'''  Cad = "select distinct c1.c_empresa, c1.c_instalacion, c1.dni_usuario, c1.n_dosimetro, "
'''  Cad = Cad & "c1.mes_p_i, c1.f_asig_dosimetro, c1.f_retirada, c1.n_reg_dosimetro from "
'''  Cad = Cad & "(SELECT distinct dosimetros.c_empresa, dosimetros.c_instalacion, dosimetros.dni_usuario, "
'''  Cad = Cad & "dosimetros.n_dosimetro, dosimetros.mes_p_i, f_asig_dosimetro, f_retirada, "
'''  Cad = Cad & "dosimetros.n_reg_dosimetro from dosimetros, operarios "
'''  Cad = Cad & "where operarios.semigracsn = 1 and operarios.dni = dosimetros.dni_usuario and "
'''  Cad = Cad & "dosimetros.tipo_dosimetro = 0 and dosimetros.n_dosimetro not like 'VIRTUAL%' and "
'''  Cad = Cad & "dosimetros.sistema = '" & Sist & "'"
'''  If Check1(0).Value = 1 Then Cad = Cad & " and operarios.f_baja is null"
  
' A ver si ahora..... >_>

'  If Check1(0).Value = 1 Then
'    ' solo los que no tienen fecha de baja
'    Cad = "SELECT distinct dosimetros.c_empresa, dosimetros.c_instalacion, dosimetros.dni_usuario, dosimetros.n_dosimetro, dosimetros.mes_p_i "
'
'    Cad = Cad & ",f_asig_dosimetro,f_retirada "
'
'    Cad = Cad & "from dosimetros, operarios where operarios.f_baja is null "
'    Cad = Cad & "and operarios.semigracsn = 1 and operarios.dni = dosimetros.dni_usuario "
'
'  Else
'    ' todos los usuarios tengan o no fecha de baja
'    Cad = "SELECT distinct dosimetros.c_empresa, dosimetros.c_instalacion, dosimetros.dni_usuario, dosimetros.n_dosimetro, dosimetros.mes_p_i "
'
'    Cad = Cad & ",f_asig_dosimetro,f_retirada "
'
'    Cad = Cad & "from dosimetros, operarios where operarios.semigracsn = 1 "
'    Cad = Cad & "and operarios.dni = dosimetros.dni_usuario "
'  End If
    
  ' solo seleccionamos los dosimetros de cuerpo
  ' ### [DavidV] 10/04/2006: Arreglos en la fórmula para solucionar un error de consulta.
  ' 23/06/2006: Ahora no, al parecer ahora tiene que ser como estaba antes...
  '
  ' Después de varios cambios aquí no registrados, se supone que si se marca
  ' "sólo fichero migrado" debe de salir aquellos con fecha de retirada NULL.
  ' si no, también aquellos que hayan sido dados de baja el año de la consulta.
  '
  ' ### [DavidV] 18/10/2006: Espero que esta sea la definitiva y última vez que me piden
  ' cambiar esto. Ya no sé ni cual era el código inicial.
''  Cad = Cad & " AND dosimetros.tipo_dosimetro = 0 and (dosimetros.f_retirada is null or "
''  Cad = Cad & "year(dosimetros.f_retirada) = " & ano & ")"
''  Cad = Cad & " AND dosimetros.tipo_dosimetro = 0 and (dosimetros.f_retirada is null"
''  If Check1(1).Value = 0 Then Cad = Cad & " or year(dosimetros.f_retirada)=" & ano
''  Cad = Cad & ") and dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%' "
'  Cad = Cad & " AND dosimetros.tipo_dosimetro = 0 and dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%' "
'  Cad = Cad & " and dosimetros.n_dosimetro NOT LIKE 'VIRTUAL%' "

'''  If txtEmp(12).Text <> "" Then Cad = Cad & " and dosimetros.c_empresa >= '" & Trim(txtEmp(12).Text) & "'"
'''  If txtEmp(13).Text <> "" Then Cad = Cad & " and dosimetros.c_empresa <= '" & Trim(txtEmp(13).Text) & "'"
'''  If txtOpe(6).Text <> "" Then Cad = Cad & " and dosimetros.dni_usuario >= '" & Trim(txtOpe(6).Text) & "'"
'''  If txtOpe(7).Text <> "" Then Cad = Cad & " and dosimetros.dni_usuario <= '" & Trim(txtOpe(7).Text) & "'"
'''
'''  ' Sólo fichero migrado.
'''  If Check1(1).Value = 1 Then
'''    Cad = Cad & " and operarios.dni in (select distinct dosimetros.dni_usuario from "
'''    Cad = Cad & "dosimetros, tempnc where dosimetros.n_dosimetro=tempnc.n_dosimetro "
'''    Cad = Cad & "and dosimetros.f_retirada is null and codusu = " & vUsu.codigo & ") "
'''
'''    Cad = Cad & "and (operarios.dni,dosimetros.c_empresa,dosimetros.c_instalacion) in "
'''    Cad = Cad & "(select distinct dni_usuario,c_empresa,c_instalacion from "
'''    Cad = Cad & "dosimetros, tempnc where dosimetros.n_dosimetro=tempnc.n_dosimetro and "
'''    Cad = Cad & "dosimetros.f_retirada is null and codusu = " & vUsu.codigo & ") "
'''  End If
'''
'''  Cad = Cad & ") c1 inner join (select * from dosiscuerpo where year(f_dosis) = " & ano
'''  Cad = Cad & ") c2 on c1.n_dosimetro = c2.n_dosimetro and c1.n_reg_dosimetro = c2.n_reg_dosimetro"
   
  ' ### [DavidV] 10/04/2006: Arreglos en la fórmula para ordenación por orden de recepción.
  ' ### [DavidV] 18/10/2006: Sin orden... no hace ninguna falta, porque no mostramos
  ' dosímetros, si no DOSIS basadas en estos, y agrupadas por meses.
  'Cad = Cad & " order by dosimetros.c_empresa, dosimetros.dni_usuario, dosimetros.c_instalacion "
  'Cad = Cad & " order by dosimetros.orden_recepcion, dosimetros.f_retirada"
    
    
    Set rs = New ADODB.Recordset
    rs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If rs.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation, "¡Atención!"
    Else
        sql = "delete from zdosisacumtot where codusu = " & vUsu.codigo
        conn.Execute sql
        
        anoini = CInt(Text4(0).Text) - 4
        If anoini < 2002 Then anoini = 2002
        fecini = Format(anoini, "0000") & "-01-01"
        
        rs.MoveFirst
        While Not rs.EOF
            
        '  If Check1(1).Value = 1 Then Cad = DevuelveDesdeBD(1, "n_dosimetro", "tempnc", "n_dosimetro|", Rs.Fields(3).Value & "|", "N|", 1)
          If Cad <> "" Then
            sql1 = "select sum(dosis_superf), sum(dosis_profunda) from dosiscuerpo "
            sql1 = sql1 & " where dni_usuario = '" & Trim(rs.Fields(2).Value) & "' and "
            sql1 = sql1 & "f_dosis <= '" & Format("31/12/" & Text4(0).Text, FormatoFecha) & "' and "
            sql1 = sql1 & "f_dosis >='" & Format(fecini, FormatoFecha) & "'"
            
            Set rL = New ADODB.Recordset
            rL.Open sql1, conn, adOpenKeyset, adLockPessimistic, adCmdText
            
            sql2 = "insert into zdosisacumtot (codusu, c_empresa, c_instalacion, dni_usuario,"
            sql2 = sql2 & "dosissuper, dosisprofu) values (" & vUsu.codigo & ",'" & Trim(rs.Fields(0).Value) & "','"
            sql2 = sql2 & Trim(rs.Fields(1).Value) & "','" & rs.Fields(2).Value & "',"
            
            dosissuper = 0
            dosisprofu = 0
            
            If Not rL.EOF Then
                If Not IsNull(rL.Fields(0).Value) Then dosissuper = rL.Fields(0).Value
                If Not IsNull(rL.Fields(1).Value) Then dosisprofu = rL.Fields(1).Value
                
                sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(dosissuper))) & ","
                sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(dosisprofu))) & ")"
            Else
                sql2 = sql2 & "0,0)"
            End If

            conn.Execute sql2
            
            rL.Close
            Set rL = Nothing
            
          End If
          rs.MoveNext
        Wend
    End If
End Sub
