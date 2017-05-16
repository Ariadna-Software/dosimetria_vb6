VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOperarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operarios"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8640
   Icon            =   "frmOperarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   8640
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2850
      MaxLength       =   40
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   600
      Width           =   4740
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   9
      Left            =   1410
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "DNI|T|N|||operarios|dni||S|"
      Text            =   "Text1"
      Top             =   630
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7410
      TabIndex        =   23
      Top             =   5580
      Width           =   1110
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4350
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   7673
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos &Personales"
      TabPicture(0)   =   "frmOperarios.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Datos Laborales"
      TabPicture(1)   =   "frmOperarios.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ImgPpal(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label20"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1(19)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Combo3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1710
         TabIndex        =   20
         Tag             =   "Se migra CSN|N|N|||operarios|semigracsn|||"
         Text            =   "Combo1"
         Top             =   3840
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   19
         Left            =   4410
         MaxLength       =   40
         TabIndex        =   21
         Tag             =   "Migrado|T|S|||operarios|migrado|||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   405
      End
      Begin VB.Frame Frame7 
         Height          =   705
         Left            =   270
         TabIndex        =   56
         Top             =   1620
         Width           =   7815
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   1950
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "Fecha Alta|F|N|||operarios|f_alta|dd/mm/yyyy|S|"
            Text            =   "Text1"
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   16
            Left            =   6030
            MaxLength       =   40
            TabIndex        =   16
            Tag             =   "Fecha Baja|F|S|||operarios|f_baja|dd/mm/yyyy||"
            Text            =   "Text1"
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha Alta"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   58
            Top             =   330
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Fecha Baja"
            Height          =   255
            Left            =   4590
            TabIndex        =   57
            Top             =   330
            Width           =   900
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   6
            Left            =   5730
            MouseIcon       =   "frmOperarios.frx":0D02
            MousePointer    =   99  'Custom
            Picture         =   "frmOperarios.frx":0E54
            ToolTipText     =   "Seleccionar fecha"
            Top             =   330
            Width           =   240
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   5
            Left            =   1650
            MouseIcon       =   "frmOperarios.frx":0EDF
            MousePointer    =   99  'Custom
            Picture         =   "frmOperarios.frx":1031
            ToolTipText     =   "Seleccionar fecha"
            Top             =   330
            Width           =   240
         End
      End
      Begin VB.Frame Frame6 
         Height          =   675
         Left            =   -74640
         TabIndex        =   49
         Top             =   3240
         Width           =   7605
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmOperarios.frx":10BC
            Left            =   4530
            List            =   "frmOperarios.frx":10BE
            TabIndex        =   10
            Text            =   "Combo1"
            Top             =   240
            Width           =   1425
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   9
            Tag             =   "Fecha Nacimiento|F|S|||operarios|f_nacimiento|dd/mm/yyyy||"
            Text            =   "Text1"
            Top             =   270
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   4530
            MaxLength       =   40
            TabIndex        =   24
            Tag             =   "Sexo|T|N|||operarios|sexo|||"
            Text            =   "Text1"
            Top             =   270
            Width           =   780
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   3
            Left            =   1860
            MouseIcon       =   "frmOperarios.frx":10C0
            MousePointer    =   99  'Custom
            Picture         =   "frmOperarios.frx":1212
            ToolTipText     =   "Seleccionar fecha"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label18 
            Caption         =   "Fecha Nacimiento"
            Height          =   255
            Left            =   390
            TabIndex        =   51
            Top             =   270
            Width           =   1410
         End
         Begin VB.Label Label21 
            Caption         =   "Sexo :"
            Height          =   195
            Left            =   3930
            TabIndex        =   50
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   270
         TabIndex        =   46
         Top             =   2400
         Width           =   7785
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   18
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "Tipo de Trabajo|T|N|||operarios|cod_rama_gen|||"
            Text            =   "Text1"
            Top             =   180
            Width           =   1320
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   3330
            MaxLength       =   40
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   180
            Width           =   4155
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3330
            MaxLength       =   40
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   510
            Width           =   4155
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   1980
            MaxLength       =   40
            TabIndex        =   19
            Tag             =   "Profesión/Categoria|T|S|||operarios|profesion_catego|||"
            Text            =   "Text1"
            Top             =   870
            Width           =   5520
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Tipo de Trabajo|T|N|||operarios|c_tipo_trabajo|||"
            Text            =   "Text1"
            Top             =   510
            Width           =   1320
         End
         Begin VB.Label Label16 
            Caption         =   "Rama Generica:"
            Height          =   255
            Left            =   330
            TabIndex        =   60
            Top             =   210
            Width           =   1290
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   7
            Left            =   1620
            MouseIcon       =   "frmOperarios.frx":129D
            MousePointer    =   99  'Custom
            Picture         =   "frmOperarios.frx":13EF
            ToolTipText     =   "Buscar socio"
            Top             =   210
            Width           =   240
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   1
            Left            =   1620
            MouseIcon       =   "frmOperarios.frx":14F1
            MousePointer    =   99  'Custom
            Picture         =   "frmOperarios.frx":1643
            ToolTipText     =   "Buscar socio"
            Top             =   540
            Width           =   240
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión Categoria:"
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   900
            Width           =   1560
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo de Trabajo:"
            Height          =   255
            Left            =   330
            TabIndex        =   47
            Top             =   540
            Width           =   1290
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1305
         Left            =   -74640
         TabIndex        =   41
         Top             =   480
         Width           =   7605
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   2
            Tag             =   "Primer Apellido|T|N|||operarios|apellido_1|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   5475
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   3
            Tag             =   "Segundo Apellido|T|N|||operarios|apellido_2|||"
            Text            =   "Text1"
            Top             =   570
            Width           =   5505
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   4
            Tag             =   "Nombre|T|N|||operarios|nombre|||"
            Text            =   "Text1"
            Top             =   930
            Width           =   5520
         End
         Begin VB.Label Label17 
            Caption         =   "Segundo Apellido:"
            Height          =   255
            Left            =   300
            TabIndex        =   45
            Top             =   630
            Width           =   1365
         End
         Begin VB.Label Label9 
            Caption         =   "Primer Apellido:"
            Height          =   195
            Left            =   300
            TabIndex        =   44
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label Label5 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   300
            TabIndex        =   43
            Top             =   960
            Width           =   1290
         End
         Begin VB.Label Label4 
            Caption         =   "Provincia:"
            Height          =   255
            Left            =   2910
            TabIndex        =   42
            Top             =   945
            Width           =   930
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1245
         Left            =   300
         TabIndex        =   37
         Top             =   390
         Width           =   7755
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   6000
            TabIndex        =   12
            Text            =   "Combo1"
            Top             =   300
            Width           =   1425
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   6000
            MaxLength       =   10
            TabIndex        =   55
            Tag             =   "Plantilla/Contrata|T|N|||operarios|plantilla_contrata|||"
            Text            =   "Text1"
            Top             =   330
            Width           =   1425
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   1890
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "Carner Radilogico|T|S|||operarios|n_carnet_radiolog|||"
            Text            =   "Text1"
            Top             =   810
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   1890
            MaxLength       =   20
            TabIndex        =   11
            Tag             =   "N.Seguridad Soc.|T|S|||operarios|n_seg_social|||"
            Text            =   "Text1"
            Top             =   270
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   5970
            MaxLength       =   40
            TabIndex        =   14
            Tag             =   "Fecha Emision|F|S|||operarios|f_emi_carnet_rad|dd/mm/yyyy||"
            Text            =   "Text1"
            Top             =   780
            Width           =   1425
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   4
            Left            =   5670
            MouseIcon       =   "frmOperarios.frx":1745
            MousePointer    =   99  'Custom
            Picture         =   "frmOperarios.frx":1897
            ToolTipText     =   "Seleccionar fecha"
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label14 
            Caption         =   "Plantilla/Contrata:"
            Height          =   255
            Left            =   4560
            TabIndex        =   52
            Top             =   330
            Width           =   1905
         End
         Begin VB.Label Label6 
            Caption         =   "N.Seguridad Social:"
            Height          =   195
            Left            =   330
            TabIndex        =   40
            Top             =   330
            Width           =   1515
         End
         Begin VB.Label Label8 
            Caption         =   "Carnet Radiológico:"
            Height          =   255
            Left            =   330
            TabIndex        =   39
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Emisión "
            Height          =   255
            Left            =   4560
            TabIndex        =   38
            Top             =   810
            Width           =   1110
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1305
         Left            =   -74640
         TabIndex        =   30
         Top             =   1860
         Width           =   7605
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   3810
            MaxLength       =   30
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   930
            Width           =   3420
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   3810
            MaxLength       =   30
            TabIndex        =   8
            Tag             =   "Poblacion|T|S|||operarios|poblacion|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   3420
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   1710
            MaxLength       =   5
            TabIndex        =   6
            Tag             =   "C.Postal|T|N|||operarios|c_postal|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   1710
            MaxLength       =   40
            TabIndex        =   5
            Tag             =   "Direccion|T|S|||operarios|direccion|||"
            Text            =   "Text1"
            Top             =   210
            Width           =   5520
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   7
            Tag             =   "Distrito|T|S|||operarios|distrito|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   645
         End
         Begin VB.Label Label13 
            Caption         =   "Provincia:"
            Height          =   255
            Left            =   2910
            TabIndex        =   35
            Top             =   945
            Width           =   930
         End
         Begin VB.Label Label12 
            Caption         =   "Poblacion:"
            Height          =   255
            Left            =   2910
            TabIndex        =   34
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label3 
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   390
            TabIndex        =   33
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Código Postal:"
            Height          =   255
            Left            =   390
            TabIndex        =   32
            Top             =   630
            Width           =   1035
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   2
            Left            =   1440
            MouseIcon       =   "frmOperarios.frx":1922
            MousePointer    =   99  'Custom
            Picture         =   "frmOperarios.frx":1A74
            ToolTipText     =   "Buscar código postal"
            Top             =   630
            Width           =   240
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   -2430
         Top             =   1665
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
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
         UserName        =   "root"
         Password        =   "aritel"
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
      Begin VB.Label Label22 
         Caption         =   "Migra al CSN:"
         Height          =   255
         Left            =   600
         TabIndex        =   62
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label Label20 
         Caption         =   "Migrado:"
         Height          =   255
         Left            =   3660
         TabIndex        =   61
         Top             =   3840
         Width           =   720
      End
      Begin VB.Image ImgPpal 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Index           =   0
         Left            =   7470
         Picture         =   "frmOperarios.frx":1B76
         Top             =   3750
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7470
      TabIndex        =   25
      Top             =   5550
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   315
      TabIndex        =   26
      Top             =   5490
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6150
      TabIndex        =   22
      Top             =   5580
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   540
      Top             =   5640
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
      TabIndex        =   28
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "L"
                  Object.Tag             =   "L"
                  Text            =   "Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "E"
                  Object.Tag             =   "E"
                  Text            =   "Etiquetas"
               EndProperty
            EndProperty
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
   Begin VB.Label Label1 
      Caption         =   "D.N.I."
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
      Left            =   510
      TabIndex        =   36
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmOperarios"
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
Private WithEvents frmPro As frmProvincias
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmIns As frmInstalaciones
Attribute frmIns.VB_VarHelpID = -1
Private WithEvents frmTTr As frmTiposTrab
Attribute frmTTr.VB_VarHelpID = -1
Private WithEvents frmOpeIns As frmOperariosInstala
Attribute frmOpeIns.VB_VarHelpID = -1
Private WithEvents frmRGe As frmRamasGener
Attribute frmRGe.VB_VarHelpID = -1


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

' campo que indica si la familia es fitosanitaria
' si lo es: obligamos a introducir los campos de fitos.
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

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
                 imgppal_Click (0)
'                 cargagrid ""
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
'            lblIndicador.Caption = ""
            DesBloqueaRegistroForm Text1(0)
            PonerModo 2
            PonerCampos
        
        
        End Select
        
'    SSTab1.Tab = 0
End Sub

Private Function SituarData1() As Boolean
    Dim sql As String
    Dim Empresa As String
    Dim dni As String
    Dim fecha As Date
    
    On Error GoTo ESituarData1
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            Empresa = Text1(0).Text
            dni = Text1(13).Text
            fecha = Format(Text1(11).Text, FormatoFecha)
            If (Trim(Data1.Recordset.Fields!c_empresa) = Trim(Empresa)) Then
               If Trim(Data1.Recordset.Fields!dni) = Trim(dni) Then
                    If Trim(Data1.Recordset.Fields!f_alta) = Trim(fecha) Then
                        SituarData1 = True
                        Exit Function
                    End If
               End If
            End If
            .MoveNext
        Wend
    End With
        
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    '###A mano
    'precios
    
    PonerFoco Text1(9)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
'        lblIndicador.Caption = "BUSCAR"
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarillo
         Select Case SSTab1.Tab
             Case 0
                 PonerFoco Text1(0)
                 Text1(0).BackColor = vbYellow
             Case 1
                 PonerFoco Text1(11)
                 Text1(11).BackColor = vbYellow
        End Select
        
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
    'empresa
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    'dni
    Text1(9).Locked = True
    Text1(9).BackColor = &H80000018
    'fecha alta
    Text1(11).Locked = True
    Text1(11).BackColor = &H80000018
    DespalzamientoVisible False
    cmdCancelar.Caption = "&Cancelar"
    Imgppal(4).Enabled = False
    
    Select Case SSTab1.Tab
        Case 0
            PonerFoco Text1(17)
        Case 1
            PonerFoco Text1(11)
   End Select

End Sub



Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    '******* canviar el mensage i la cadena *********************
    Cad = "Seguro que desea eliminar el operario:"
    Cad = Cad & vbCrLf & "DNI: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Fecha Alta: " & Data1.Recordset.Fields(16)
    '**********************************************************
    I = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") 'VRS:1.0.1(11)
    
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
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Operario"
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

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SSTab1.Tab = 1
        PonerFoco Text1(12)
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Combo1_LostFocus()
    If Combo1.ListIndex = 0 Then Text1(7).Text = "01"
    If Combo1.ListIndex = 1 Then Text1(7).Text = "02"
End Sub


Private Sub Combo2_LostFocus()
    If Combo2.ListIndex = 0 Then Text1(15).Text = "V"
    If Combo2.ListIndex = 1 Then Text1(15).Text = "M"
End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SSTab1.Tab = 1
        PonerFoco Text1(12)
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
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
    NombreTabla = "operarios"
    Ordenacion = " ORDER BY dni"
    '******************************************************+
        
'    PonerOpcionesMenu
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    
    Data1.ConnectionString = Conn
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
    
    SSTab1.Tab = 0
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
'    lblIndicador.Caption = ""
    
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
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
        Aux = ValorDevueltoFormGrid(Text1(9), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
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
        Text1(10).Text = RecuperaValor(CadenaSeleccion, 2)
'        Text2(2).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub frmrge_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(18).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End If

End Sub

Private Sub frmTTR_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(18).Text = RecuperaValor(CadenaSeleccion, 1)
        Text1(6).Text = RecuperaValor(CadenaSeleccion, 2)
'        Text2(3).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
'    If Modo = 0 Or Modo = 2 Then Exit Sub
    Select Case Index
       Case 0
            If Modo < 2 Then Exit Sub
            
            Set frmOpeIns = New frmOperariosInstala
            frmOpeIns.dni = Trim(Text1(9).Text)
            'frmOpeIns.fechaalta = Format(Text1(17).Text)
            frmOpeIns.Show
       Case 3 'fecha de nacimiento
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(14).Text <> "" Then
                If IsDate(Text1(14).Text) Then f = Text1(14).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(14).Text = frmC.fecha
                mTag.DarFormato Text1(14)
            End If
            Set frmC = Nothing
'
       Case 4
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(0).Text <> "" Then
                If IsDate(Text1(0).Text) Then f = Text1(0).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(0).Text = frmC.fecha
                mTag.DarFormato Text1(0)
            End If
            Set frmC = Nothing
        
       Case 5
            'fecha de alta
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(17).Text <> "" Then
                If IsDate(Text1(17).Text) Then f = Text1(17).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(17).Text = frmC.fecha
                mTag.DarFormato Text1(17)
            End If
            Set frmC = Nothing
       
       Case 6
            'fecha de baja
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(16).Text <> "" Then
                If IsDate(Text1(16).Text) Then f = Text1(16).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(16).Text = frmC.fecha
                mTag.DarFormato Text1(16)
            End If
            Set frmC = Nothing
        
        Case 2 ' codigo de provincia
            Set frmPro = New frmProvincias
            frmPro.DatosADevolverBusqueda = "0|1|"
            frmPro.Show
        Case 1 ' tipo de trabajo
            Set frmTTr = New frmTiposTrab
            frmTTr.DatosADevolverBusqueda = "2|3|"
            frmTTr.Show
        Case 7
            Set frmRGe = New frmRamasGener
            frmRGe.DatosADevolverBusqueda = "0|1|"
            frmRGe.Show
   End Select
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 18 Then Exit Sub  ' caso de pulsar ALT
    AsignarTeclasFuncion KeyCode
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If

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

    If Text1(Index).Text = "" Then Exit Sub
    
    If Modo = 1 And ConCaracteresBusqueda(Text1(Index).Text) Then Exit Sub
    
    Select Case Index
        Case 2, 4, 6, 9, 11, 12, 15, 18
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            If Modo = 1 Then Exit Sub
            Select Case Index
'                Case 0 'empresa
'                    Text2(1).Text = ""
'                    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
'                    If Text2(1).Text = "" Then
'                        MsgBox "El código de empresa no existe. Reintroduzca.", vbExclamation
'                        Text1(Index).Text = ""
'                        PonerFoco Text1(Index)
'                    End If
                Case 2 'codigo de provincia
                    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(2).Text & "|", "T|", 1)
                    If Text2(0).Text = "" Then
                        MsgBox "Código de provincia no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                Case 6 ' tipo de trabajo
                    If Text1(Index).Text <> "" Then
                        Text2(3).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", Text1(Index).Text & "|", "T|", 1)
                        If Text2(3).Text = "" Then
                            MsgBox "El código de tipo de trabajo no existe. Reintroduzca.", vbExclamation, "¡Error!"
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                Case 18 ' tipo de trabajo
                    If Text1(Index).Text <> "" Then
                        Text2(1).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(Index).Text & "|", "T|", 1)
                        If Text2(1).Text = "" Then
                            MsgBox "El código de rama genérica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
            End Select
        Case 0, 14, 16, 17
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
Dim CadB As String

Combo1.Enabled = False
Combo2.Enabled = False

CadB = ObtenerBusqueda(Me)

Combo1.Enabled = True
Combo2.Enabled = True

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
        Cad = Cad & ParaGrid(Text1(9), 15, "DNI")
        Cad = Cad & ParaGrid(Text1(13), 20, "Apellido 1")
        Cad = Cad & ParaGrid(Text1(10), 20, "Apellido 2")
        Cad = Cad & ParaGrid(Text1(5), 20, "Nombre")
        Cad = Cad & ParaGrid(Text1(8), 25, "Profesión")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|"
            frmB.vTitulo = "Operarios"
            frmB.vSelElem = 1
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
        Combo1.ListIndex = CInt(Text1(7).Text) - 1
        Combo2.ListIndex = -1
        If Text1(15).Text = "V" Then Combo2.ListIndex = 0
        If Text1(15).Text = "M" Then Combo2.ListIndex = 1
    End If
    
    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Trim(Text1(2).Text) & "|", "T|", 1)
    Text2(1).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(18).Text & "|", "T|", 1)
    Text2(4).Text = Trim(Text1(5).Text) & " " & Trim(Text1(13).Text) & " " & Trim(Text1(10).Text)
    Text2(3).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", Text1(6).Text & "|", "T|", 1)
'    Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Text1(0).Text & "|" & Text1(10).Text & "|", "T|T|", 2)
   
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
    
    Text1(1).Enabled = True
    Text1(2).Enabled = True
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    
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
    
    For I = 1 To Imgppal.Count - 1
        Imgppal(I).Enabled = Not b
    Next I
    Imgppal(0).Enabled = (Modo <> 0) And (Modo <> 3)
    Combo1.Enabled = Not b
    Combo2.Enabled = Not b
    Combo3.Enabled = Not b
    PonerFoco chkVistaPrevia
End Sub

Private Function DatosOk() As Boolean
Dim Rs As ADODB.Recordset
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
        ' debe existir el tipo de trabajo
        Datos = ""
        Datos = DevuelveDesdeBD(1, "descripcion", "tipostrab", "cod_rama_gen|c_tipo_trabajo|", Text1(18).Text & "|" & Text1(6).Text & "|", "T|T|", 2)
        If Datos = "" Then
            MsgBox "Debe existir el tipo de trabajo. Reintroduzca.", vbExclamation, "¡Error!"
            IndiceErroneo = I
            DatosOk = False
        End If
End If

If (b = True) And (Modo = 3) Then
     Datos = DevuelveDesdeBD(1, "dni", "operarios", "dni|f_alta|", Text1(9).Text & "|" & Text1(17).Text, "T|F|", 2)
     If Datos <> "" Then
        MsgBox "Ya existe el DNI de operario en esa fecha de alta: " & Text1(9).Text, vbExclamation, "¡Error!"
        DatosOk = False
        IndiceErroneo = 0
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
        Case 11
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 3 'Listado de operarios
            FrmListado.Show
        Case 12
            mnSalir_Click
        Case 14 To 17
            Desplazamiento (Button.Index - 14)
        Case 11
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
    Combo3.AddItem "No"
    Combo3.ItemData(Combo3.NewIndex) = 0

    Combo3.AddItem "Si"
    Combo3.ItemData(Combo3.NewIndex) = 1

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu
        Case "Listado"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 3 'Listado de operarios
            FrmListado.Show
        Case "Etiquetas"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 26 'Listado de etiquetas
            FrmListado.Show
    End Select
End Sub


Private Function Eliminar() As Boolean
Dim I As Integer
Dim sql As String

        sql = " WHERE dni='" & Data1.Recordset!dni & "' and f_alta = '" & Format(Data1.Recordset!f_alta, FormatoFecha) & "'"
        
        Conn.Execute "Delete  from operarios " & sql
       
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

