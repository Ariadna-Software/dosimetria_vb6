VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmPpal 
   BackColor       =   &H8000000C&
   Caption         =   "Cálculo y Gestión de Dosimetría Personal"
   ClientHeight    =   5025
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   8085
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "extremidades"
            Object.ToolTipText     =   "Extremidades"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ramasgen"
            Object.ToolTipText     =   "Ramas Generales"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ramasesp"
            Object.ToolTipText     =   "Ramas Específicas"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tipostrab"
            Object.ToolTipText     =   "Tipo de Trabajos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "empresas"
            Object.ToolTipText     =   "Empresas"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "instalaciones"
            Object.ToolTipText     =   "Instalaciones"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "operarios"
            Object.ToolTipText     =   "Operarios"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dosímetros"
            Object.ToolTipText     =   "Dosímetros"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dosiscuerpo"
            Object.ToolTipText     =   "Dosis a cuerpo"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dosisnohomo"
            Object.ToolTipText     =   "Dosis No Homogénea"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dosisárea"
            Object.ToolTipText     =   "Dosis por Area"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "1"
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fichapers"
            Object.ToolTipText     =   "Ficha de Personal"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sobredosis"
            Object.ToolTipText     =   "Operarios con Sobredosis"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpal.frx":139C
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6191
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:41"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":495E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4F92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":52AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":55C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":58E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5BFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":86C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DEB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1512C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":15B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":16550
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":16F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":17974
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":18386
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":18D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":197AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1A1BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1A4D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1AEE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1B8FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1C30C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListComun 
      Left            =   6000
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1CD1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1EA28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":24CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":256E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2AED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2D684
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2DF5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2E838
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2F112
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2F9EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":35B4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":35FA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":360BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":361CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":362DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":365F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3C21A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":41A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4241E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":486B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4DEAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5369C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":539B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":53CD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnDatos 
      Caption         =   "&Entrada de datos"
      Begin VB.Menu mnGenerales 
         Caption         =   "&Generales"
         Begin VB.Menu mnProvincias 
            Caption         =   "&Provincias"
         End
         Begin VB.Menu mnTiposExtremidades 
            Caption         =   "&Tipos de Medición de Extremidades"
         End
         Begin VB.Menu mnbarra1 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnRamasActividades 
            Caption         =   "&Ramas de Actividades"
            Begin VB.Menu mnGenerica 
               Caption         =   "&Genéricas"
            End
            Begin VB.Menu mnEspecificas 
               Caption         =   "&Específicas"
            End
            Begin VB.Menu mnTiposTrabajos 
               Caption         =   "&Tipos de Trabajos"
            End
         End
      End
      Begin VB.Menu mnbarra2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnEmpresas 
         Caption         =   "&Empresas"
      End
      Begin VB.Menu mnInstalaciones 
         Caption         =   "&Instalaciones"
      End
      Begin VB.Menu mnOperarios 
         Caption         =   "&Operarios"
      End
      Begin VB.Menu mnFichaPersonal 
         Caption         =   "&Ficha de Personal"
         HelpContextID   =   1
      End
      Begin VB.Menu mnOperariosInstalaciones 
         Caption         =   "&Relación Operarios/Instalaciones"
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnFactoresCalibracion 
         Caption         =   "&Factores de Calibración"
         Begin VB.Menu mnFactores4400 
            Caption         =   "Harshaw &4400"
         End
         Begin VB.Menu mnFactores6600 
            Caption         =   "Harshaw &6600"
         End
         Begin VB.Menu mnFactoresPanasonic 
            Caption         =   "&Panasonic"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnbarra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnFondos 
         Caption         =   "&Mantenimiento de Fondos"
         Begin VB.Menu mnFondos6600 
            Caption         =   "&Harshaw 6600"
         End
         Begin VB.Menu mnFondosPana 
            Caption         =   "&Panasonic"
         End
      End
      Begin VB.Menu mnbarra6 
         Caption         =   "-"
         HelpContextID   =   2
      End
      Begin VB.Menu mnConfiguracionAplicacion 
         Caption         =   "Confi&guracion"
         HelpContextID   =   2
         Begin VB.Menu mnParametros 
            Caption         =   "&Parametros"
            HelpContextID   =   2
         End
         Begin VB.Menu mnUsuarios 
            Caption         =   "&Mantenimiento de Usuarios"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu mnbarra8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSal 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnbarra99 
         Caption         =   "-"
         HelpContextID   =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnNoTocar 
         Caption         =   "&No tocar"
         HelpContextID   =   2
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnDosimetrosCuerpo 
      Caption         =   "&Dosímetros"
      Begin VB.Menu mnDosimetros 
         Caption         =   "&Dosímetros"
      End
      Begin VB.Menu mnLotes 
         Caption         =   "&Lotes de Dosímetros"
         Begin VB.Menu mnLotes6600 
            Caption         =   "&Harshaw 6600"
         End
         Begin VB.Menu mnLotesPana 
            Caption         =   "&Panasonic"
         End
      End
      Begin VB.Menu mnbarra41 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnRecepcionDosimetros 
         Caption         =   "&Recepción de Dosímetros"
         Begin VB.Menu mnGenerarRecepcion 
            Caption         =   "&Generar Recepción"
            HelpContextID   =   1
         End
         Begin VB.Menu mnControlDosimLapiz 
            Caption         =   "&Control de Dosímetros Lápiz Óptico"
            HelpContextID   =   1
         End
      End
   End
   Begin VB.Menu mnDosis 
      Caption         =   "&Dosis"
      Begin VB.Menu mnHomogeneas 
         Caption         =   "&Homogéneas"
      End
      Begin VB.Menu mnNoHomogeneas 
         Caption         =   "&No Homogéneas"
      End
      Begin VB.Menu mnDosisArea 
         Caption         =   "&Por Area"
      End
      Begin VB.Menu mnbarra42 
         Caption         =   "-"
         HelpContextID   =   1
      End
      Begin VB.Menu mnCargaAutomaticaDosis 
         Caption         =   "&Carga Automática de Dosis"
         HelpContextID   =   1
         Begin VB.Menu mnHarshaw4400 
            Caption         =   "Harshaw &4400"
            HelpContextID   =   1
         End
         Begin VB.Menu mnbarra43 
            Caption         =   "-"
            HelpContextID   =   1
         End
         Begin VB.Menu mnCargaAutom6600 
            Caption         =   "Harshaw &6600"
            HelpContextID   =   1
         End
         Begin VB.Menu mnCargaAutomPana 
            Caption         =   "&Panasonic"
            HelpContextID   =   1
         End
      End
      Begin VB.Menu mnCargaExtremidades 
         Caption         =   "Carga de &Extremidades"
         HelpContextID   =   1
         Begin VB.Menu mnCargaExtremidad6600 
            Caption         =   "Harshaw &6600"
            HelpContextID   =   1
         End
         Begin VB.Menu mnCargaExtremidadPanasonic 
            Caption         =   "&Panasonic"
            HelpContextID   =   1
         End
      End
      Begin VB.Menu mnErroresMigracion 
         Caption         =   "&Errores de Migración"
         Begin VB.Menu mnErroresMigraPersonal 
            Caption         =   "&Personal"
         End
         Begin VB.Menu mnErroresMigraArea 
            Caption         =   "&Area"
         End
      End
      Begin VB.Menu mnCancelacionMigracion 
         Caption         =   "&Cancelación Migración"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu mnInformesAdministrativos 
      Caption         =   "&Informes Administrativos"
      Begin VB.Menu mnInformeEmpresas 
         Caption         =   "&Empresas"
         Begin VB.Menu mnEtiquetasEmpresas 
            Caption         =   "&Etiquetas Adhesivas"
         End
         Begin VB.Menu mnListadoEmpresas 
            Caption         =   "&Listado"
         End
      End
      Begin VB.Menu mnInformeInstalaciones 
         Caption         =   "&Instalaciones"
         Begin VB.Menu mnEtiquetasInstalaciones 
            Caption         =   "&Etiquetas Adhesivas"
         End
         Begin VB.Menu mnListadoInstalaciones 
            Caption         =   "&Listado"
         End
      End
      Begin VB.Menu mInformeUsuarios 
         Caption         =   "&Usuarios"
         Begin VB.Menu mnEtiquetasUsuarios 
            Caption         =   "&Etiquetas Adhesivas"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mndListadoUsuarios 
            Caption         =   "&Listado"
         End
      End
      Begin VB.Menu mnbarra9 
         Caption         =   "-"
      End
      Begin VB.Menu mnListadoDosimetrosCuerpo 
         Caption         =   "&Dosímetros"
      End
      Begin VB.Menu mnListadoFactoresCalibracion 
         Caption         =   "&Factores de Calibración"
         Begin VB.Menu mnListadoCalibracion4400 
            Caption         =   "Factores de Calibración &4400"
         End
         Begin VB.Menu mnListadoCalibracion6600 
            Caption         =   "Factores de Calibración &6600"
         End
      End
      Begin VB.Menu mnbarra10 
         Caption         =   "-"
      End
      Begin VB.Menu mnInformesMensuales 
         Caption         =   "&Informes Mensuales"
         Begin VB.Menu mnPersonal 
            Caption         =   "&Personal"
            Begin VB.Menu mnListadoDosimetros 
               Caption         =   "&Dosímetros"
               Begin VB.Menu mnImpConRecDosCue 
                  Caption         =   "&Impreso de Control Recepción Dosimetros Cuerpo"
               End
            End
            Begin VB.Menu mnListadoDosis 
               Caption         =   "Do&sis"
               Begin VB.Menu mnDosisInstalacion 
                  Caption         =   "&Dosis por Instalación"
               End
               Begin VB.Menu mnDosisNOHomogOperario 
                  Caption         =   "Dosis &No homogénea por Operario"
               End
               Begin VB.Menu mnDosisOpeAcumulado 
                  Caption         =   "Dosis por Operario Año Oficial"
               End
               Begin VB.Menu mnCartaSobredosis 
                  Caption         =   "Carta al &CSN de Sobredosis"
               End
               Begin VB.Menu mnCartaDosimetrosNoRec 
                  Caption         =   "&Carta Dosímetros no Recepcionados"
               End
            End
         End
         Begin VB.Menu mnArea 
            Caption         =   "&Area"
            Begin VB.Menu mnListadoDosimetrosAre 
               Caption         =   "&Dosímetros"
               Begin VB.Menu mnImpConRecDosAre 
                  Caption         =   "&Impreso de Control Recepción Dosimetros Area"
               End
            End
            Begin VB.Menu mnListadoDosisAre 
               Caption         =   "Do&sis"
               Begin VB.Menu mnDosisInstalacionArea 
                  Caption         =   "&Dosis por Instalación"
               End
            End
         End
      End
      Begin VB.Menu mnbarra19 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnVentas 
         Caption         =   "Informes &CSN"
         Begin VB.Menu mnRangosCsn 
            Caption         =   "&Rangos CSN"
         End
         Begin VB.Menu mnProfunda 
            Caption         =   "&Dosis Colectiva"
         End
      End
   End
   Begin VB.Menu mnFuncAtipicas 
      Caption         =   "&Funciones Atípicas"
      Begin VB.Menu mnTraspaseCSN 
         Caption         =   "&Traspaso Automático al CSN"
         HelpContextID   =   1
      End
      Begin VB.Menu mnCancelacionTraspaso 
         Caption         =   "&Cancelación Traspaso Automático"
         HelpContextID   =   1
      End
      Begin VB.Menu mnComprobarCSN 
         Caption         =   "&Visualizar Fichero CSN"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuUtiUniDni 
         Caption         =   "&Unificación de operaciones a un DNI"
      End
      Begin VB.Menu mnModificacionDNI 
         Caption         =   "&Modificación Automática DNI de Usuario"
         HelpContextID   =   1
      End
      Begin VB.Menu mnModificacionCodigosInstalacion 
         Caption         =   "Modificación &Automática Códigos de Instalación"
         HelpContextID   =   1
      End
      Begin VB.Menu mnbarra20 
         Caption         =   "-"
         HelpContextID   =   1
      End
      Begin VB.Menu mnOperariosSobredosis 
         Caption         =   "&Operarios con Sobredosis"
      End
      Begin VB.Menu mnExportarAnual 
         Caption         =   "&Exportar Dosis por Año"
      End
      Begin VB.Menu mnExportarCodInstalacion 
         Caption         =   "Exportar &Códigos de Instalación"
      End
      Begin VB.Menu mnDosisPenalizacionAutomaticas 
         Caption         =   "&Dosis de Penalización Automáticas"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnUsuariosActivos 
         Caption         =   "&Usuarios Activos"
      End
      Begin VB.Menu mnCopiaSeguridad 
         Caption         =   "&Copia de Seguridad"
      End
      Begin VB.Menu mnRestaurarCopia 
         Caption         =   "&Restaurar Copia"
         Enabled         =   0   'False
         HelpContextID   =   2
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnVentanas 
      Caption         =   "&Ventanas"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrimeraVez As Boolean

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub MDIForm_Load()

    'Botones
    With Me.Toolbar1
        .ImageList = Me.ImageList1
        .Buttons("extremidades").Image = 16
        .Buttons("ramasgen").Image = 13
        .Buttons("ramasesp").Image = 19
        .Buttons("tipostrab").Image = 18
        .Buttons("empresas").Image = 15
        .Buttons("instalaciones").Image = 17
        .Buttons("operarios").Image = 12
        .Buttons("dosímetros").Image = 14
        .Buttons("dosiscuerpo").Image = 20
        .Buttons("dosisnohomo").Image = 10
        .Buttons("dosisárea").Image = 7
        .Buttons("fichapers").Image = 21
        .Buttons("sobredosis").Image = 24
        
    End With
    
    'Poner opciones de nivel de usuario
    PonerDatosFormulario
    
    FormatoFecha = "yyyy-mm-dd"
    
    Randomize
    Me.Picture = LoadPicture(App.Path & "\imagenes\img_atomo1.jpg")

End Sub



Private Sub mnArticulosPdtes_Click()
Dim frmL As FrmListado

    Set frmL = New FrmListado
    Screen.MousePointer = vbHourglass
    frmL.Opcion = 6 'Listado de articulos pdtes de recibir
    frmL.Show
End Sub

Private Sub mnCambioPreciosMasivo_Click()
    
    If Not BloqueoManual(True, "CAMPRE", "Cambio") Then
        MsgBox "Hay otro usuario realizando este proceso. Inténtelo más tarde.", "¡Error!"
    Else
        Screen.MousePointer = vbHourglass
        FrmListado.Opcion = 18
        FrmListado.Show
'        frmCambioPrecios.Show
    End If
        
End Sub

Private Sub mnCancelacionMigracion_Click()
    If Not BloqueoManual(True, "CANMIGRA", "CANMIGRA") Then
        MsgBox "Proceso realizandose por otro usuario. Espere.", vbExclamation, "¡Error!"
        Exit Sub
    End If

    frmCancelacionMigra.Show vbModal
End Sub

Private Sub mnCancelacionTraspaso_Click()
    
    If Not BloqueoManual(True, "TRASPASO", "TRASPASO") Then
        MsgBox "Proceso realizandose por otro usuario. Espere.", vbExclamation, "¡Error!"
        Exit Sub
    End If

    FrmCanTraspasoCSN.Show vbModal

End Sub

Private Sub mnCargaAutom6600_Click()
    
    Dim frmT As FrmHarshaw6600
    
    Directorio = ""
    Set frmT = New FrmHarshaw6600
    frmT.Show vbModal
    frmCalculoMsv.Sistema = "H"
    frmCalculoMsv.Caption = "Mantenimiento Temporal nC"
    frmCalculoMsv.Show vbModal

End Sub

Private Sub mnCargaAutomPana_Click()
    Dim frmT As frmPanasonic
    
    Directorio = ""
    Set frmT = New frmPanasonic
    frmT.Show vbModal
    frmCalculoMsv.Sistema = "P"
    frmCalculoMsv.Caption = "Mantenimiento Temporal mSv*"
    frmCalculoMsv.Show vbModal
End Sub

Private Sub mnCargaExtremidad6600_Click()
    Dim frmT As frmDosisExtremidades
    
    Set frmT = New frmDosisExtremidades
    frmT.Sistema = "H"
    frmT.Show

End Sub

Private Sub mnCargaExtremidadPanasonic_Click()
    Dim frmT As frmDosisExtremidades
    
    Set frmT = New frmDosisExtremidades
    frmT.Sistema = "P"
    frmT.Show

End Sub

Private Sub mnCartaDosimetrosNoRec_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    frmT.Opcion = 20
    frmT.Show
End Sub

Private Sub mnCartaSobredosis_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 22 'Carta al CSN de sobredosis
    frmT.Show

End Sub

Private Sub mnComprobarCSN_Click()
Dim frmT As frmCSNTest
  Set frmT = New frmCSNTest
  frmT.Show
End Sub

Private Sub mnControlDosimLapiz_Click()
Dim frmT As frmRecepDosim

    Set frmT = New frmRecepDosim
    frmT.Show

End Sub

Private Sub mnCopiaSeguridad_Click()
    frmBackUP.Show vbModal
End Sub

Private Sub mndListadoUsuarios_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 3 'Listado de operarios
    frmT.Show

End Sub

Private Sub mnDosimetros_Click()
Dim frmT As frmDosimetros

    Set frmT = New frmDosimetros
    frmT.Show

End Sub

'Private Sub mnDosimetrosPersonal_Click()
'Dim frmT As frmDosimetros
'
'    Set frmT = New frmDosimetros
'    frmT.Show
'
'End Sub
'
'Private Sub mnDosimetrosArea_Click()
'Dim frmT As frmDosimArea
'
'    Set frmT = New frmDosimArea
'
'    frmT.Show
'
'End Sub
'
'
'Private Sub mnDosimetrosCuerpo_Click()
'Dim frmT As frmDosimetros
'
'    Set frmT = New frmDosimetros
'    frmT.Show
'
'End Sub

Private Sub mnDosisArea_Click()
Dim frmT As frmDosisArea

    Set frmT = New frmDosisArea
    frmT.Show

End Sub

Private Sub mnDosisInstalacion_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 9 'Listado de dosis por instalacion
    frmT.Show

End Sub

Private Sub mnDosisInstalacionArea_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 9 'Listado de dosis por instalacion
    frmT.Show


End Sub

Private Sub mnDosisNOHomogOperario_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 19 'Listado de dosis no homogeneas por operario
    frmT.Show
    
End Sub

Private Sub mnDosisOpeAcumulado_Click()
Dim frmL As FrmListado
    
    Set frmL = New FrmListado
    frmL.Opcion = 21
    frmL.Show

End Sub


Private Sub mnDosisPenalizacionAutomaticas_Click()

    frmPenalizacionDosis.Show vbModal
    
End Sub

'Private Sub mnCopiaSeguridad_Click() 'VRS:1.0.4(0)
'    frmBackUP.Show vbModal
'End Sub


Private Sub mnEmpresas_Click()
Dim frmT As frmEmpresas

    Set frmT = New frmEmpresas
    frmT.Show
    
End Sub

Private Sub mnErroresMigraArea_Click()
    frmErroresMigraArea.Show vbModal
End Sub

'Private Sub mnErroresMigracion_Click()
'
'    frmErroresMigra.Show vbModal
'
'End Sub

Private Sub mnErroresMigraPersonal_Click()
    frmErroresMigra.Show vbModal
End Sub

Private Sub mnEspecificas_Click()
Dim frmT As frmRamasEspe

    Set frmT = New frmRamasEspe
    frmT.Show

End Sub

Private Sub mnEtiquetasEmpresas_Click()
Dim frmT As FrmListado
    Screen.MousePointer = vbHourglass
    Set frmT = New FrmListado
    frmT.Opcion = 24 'Listado de etiquetas
    frmT.Show

End Sub

Private Sub mnEtiquetasInstalaciones_Click()
Dim frmT As FrmListado

    Screen.MousePointer = vbHourglass
    Set frmT = New FrmListado
    frmT.Opcion = 25 'Listado de etiquetas de instalaciones
    frmT.Show

End Sub

Private Sub mnEtiquetasUsuarios_Click()
Dim frmT As FrmListado

    Screen.MousePointer = vbHourglass
    Set frmT = New FrmListado
    frmT.Opcion = 26 'Listado de etiquetas
    frmT.Show

End Sub

Private Sub mnExportarAnual_Click()
Dim frmE As frmExportarAnual

    Set frmE = New frmExportarAnual
    frmE.Show vbModal, Me
    
End Sub

Private Sub mnExportarCodInstalacion_Click()
Dim frmE As frmExportarInstalaciones

    Set frmE = New frmExportarInstalaciones
    frmE.Show vbModal, Me
    

End Sub

Private Sub mnFactores4400_Click()
Dim frmT As frmFactCali4400

    Set frmT = New frmFactCali4400
    frmT.Show

End Sub

Private Sub mnFactores6600_Click()
Dim frmT As frmFactCali6600

    Set frmT = New frmFactCali6600
    frmT.Show

End Sub

Private Sub mnFactoresPanasonic_Click()
Dim frmT As frmFactCaliPana

    Set frmT = New frmFactCaliPana
    frmT.Show
End Sub

Private Sub mnFichaPersonal_Click()
Dim frmT As frmFichaPersonal
    
    frmMensajes.Opcion = 0
    frmMensajes.Show vbModal
    
    Set frmT = New frmFichaPersonal
    frmT.Show
    
End Sub


Private Sub mnFondos6600_Click()
Dim frmT As frmFondos

    Set frmT = New frmFondos
    frmT.Show
    
End Sub

Private Sub mnFondosPana_Click()
Dim frmT As frmFondosPana

    Set frmT = New frmFondosPana
    frmT.Show
    
End Sub

Private Sub mnGenerarRecepcion_Click()
    
    FrmGenRecepDosim.Show vbModal

End Sub

Private Sub mnGenerica_Click()
Dim frmT As frmRamasGener

    Set frmT = New frmRamasGener
    frmT.Show
    
End Sub


Private Sub mnHomogeneas_Click()
    Dim frmT As frmDosisHomog
    
    Set frmT = New frmDosisHomog
    frmT.Show

End Sub

Private Sub mnImpConRecDosAre_Click()
' ******************aqui
    Dim frmT As FrmListado
    Set frmT = New FrmListado
    frmT.Opcion = 27 ' listado de recepcion de dosimetros de area (recepdosim)
    frmT.Show
End Sub

Private Sub mnImpConRecDosCue_Click()
    Dim frmT As FrmListado
    Set frmT = New FrmListado
    frmT.Opcion = 23 ' listado de recepcion de dosimetros (recepdosim)
    frmT.Show
End Sub

Private Sub mnInstalaciones_Click()
    Dim frmT As frmInstalaciones
    
    Set frmT = New frmInstalaciones
    frmT.Show
End Sub

Private Sub mnListadoCalibracion4400_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 7 'Listado de factores de calibracion 4400
    frmT.Show

End Sub

Private Sub mnListadoCalibracion6600_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 8 'Listado de factores de calibracion 6600
    frmT.Show

End Sub

Private Sub mnListadoCalibracionPana_Click()
Dim frmT As FrmListado

Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 31 'Listado de factores de calibracion Panasonic
    frmT.Show

End Sub

Private Sub mnListadoDosimetrosCuerpo_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 4 'Listado de dosimetros a cuerpo
    frmT.Show
    
End Sub


'Private Sub mnListadoDosimetrosOrgano_Click()
'Dim frmT As FrmListado
'
'    Set frmT = New FrmListado
'    Screen.MousePointer = vbHourglass
'    frmT.Opcion = 5 'Listado de dosimetros a órgano
'    frmT.Show
'
'End Sub

Private Sub mnListadoEmpresas_Click()
Dim frmT As FrmListado
    
    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 1 'Listado de empresas
    frmT.Show

End Sub

Private Sub mnListadoInstalaciones_Click()
Dim frmT As FrmListado
    
    Set frmT = New FrmListado
    Screen.MousePointer = vbHourglass
    frmT.Opcion = 2 'Listado de instalaciones
    frmT.Show

End Sub


Private Sub mnLotes6600_Click()
Dim frmT As frmLotes

    Set frmT = New frmLotes
    frmT.Show

End Sub

Private Sub mnLotesPana_Click()
Dim frmT As frmLotesPana

    Set frmT = New frmLotesPana
    frmT.Show

End Sub

'Private Sub mnMigracionMsv_Click()
'Dim frmT As frmCalculoMsv
'
'    Set frmT = New frmCalculoMsv
'
'    frmT.Show vbModal
'
'End Sub

Private Sub mnModificacionCodigosInstalacion_Click()
    
    If Not BloqueoManual(True, "CAMBIINS", "CAMBIINS") Then
        MsgBox "Proceso realizandose por otro usuario. Espere.", vbExclamation, "¡Error!"
        Exit Sub
    End If
    If DevuelveDesdeBD(1, "codusu", "zbloqueos", "tabla|clave|", "CAMBIDNI|CAMBIDNI|", "T|T|", 2) <> "" Then
        BloqueoManual False, "CAMBIINS", "CAMBIINS"
        MsgBox "Otro usuario se encuentra realizando un cambio de DNI en estos momentos. No es aconsejable hacer un cambio de código de instalación al mismo tiempo. Por favor, espere.", vbExclamation, "¡Error!"
        Exit Sub
    End If
    FrmCambioInstala.Show vbModal
    

End Sub

Private Sub mnModificacionDNI_Click()
    
    If Not BloqueoManual(True, "CAMBIDNI", "CAMBIDNI") Then
        MsgBox "Proceso realizandose por otro usuario. Espere.", vbExclamation, "¡Error!"
        Exit Sub
    End If
    If DevuelveDesdeBD(1, "codusu", "zbloqueos", "tabla|clave|", "CAMBIINS|CAMBIINS|", "T|T|", 2) <> "" Then
        BloqueoManual False, "CAMBIDNI", "CAMBIDNI"
        MsgBox "Otro usuario se encuentra realizando un cambio de código de instalación en estos momentos. No es aconsejable hacer un cambio de DNI al mismo tiempo. Por favor, espere.", vbExclamation, "¡Error!"
        Exit Sub
    End If
    FrmCambioDNIope.Show vbModal
    
End Sub

Private Sub mnNoHomogeneas_Click()
Dim frmT As frmDosisNoHomog

    Set frmT = New frmDosisNoHomog
    frmT.Show

End Sub

Private Sub mnNoTocar_Click()
    Cargatablas.Show vbModal
End Sub

Private Sub mnOperarios_Click()
Dim frmT As frmOperarios

    Set frmT = New frmOperarios
    frmT.Show

End Sub

Private Sub mnOperariosInstalaciones_Click()
Dim frmT As frmOperariosInstala

    Set frmT = New frmOperariosInstala
    frmT.Show
End Sub

Private Sub mnOperariosSobredosis_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    frmT.Opcion = 28
    frmT.Show
    
End Sub

Private Sub mnParametros_Click()
Dim frmT As frmParametros

    Set frmT = New frmParametros
    frmT.Show
    
End Sub

Private Sub mnProfunda_Click()
Dim frmT As FrmListado

    Set frmT = New FrmListado
    frmT.Opcion = 12
    frmT.Show

End Sub

'Private Sub mnProfundaMensual_Click()
'Dim frmT As FrmListado
'
'    Set frmT = New FrmListado
'    frmT.Opcion = 12
'    frmT.Show
'
'End Sub

Private Sub mnProvincias_Click()
Dim frmT As frmProvincias

    Set frmT = New frmProvincias
    frmT.Show
    
End Sub

Private Sub mnRangosCsn_Click()
    Dim frmT As frmRangosCsn
    
    Set frmT = New frmRangosCsn
    frmT.Show
    
End Sub

Private Sub mnRestaurarCopia_Click()
    frmAccionesBD2.Show vbModal
    
End Sub

Private Sub mnTiposExtremidades_Click()
    Dim frmT As frmTiposExtremidades
    
    Set frmT = New frmTiposExtremidades
    frmT.Show
End Sub

Private Sub mnTiposTrabajos_Click()
    Dim frmT As frmTiposTrab
    
    Set frmT = New frmTiposTrab
    frmT.Show

End Sub

Private Sub mnTraspaseCSN_Click()
    
    If Not BloqueoManual(True, "TRASPASO", "TRASPASO") Then
        MsgBox "Proceso realizandose por otro usuario. Espere.", vbExclamation, "¡Error!"
        Exit Sub
    End If
    
    FrmTraspasoCSN.Show vbModal
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  ' Elimino los posibles bloqueos que hayan podido quedarse colgados para
  ' este usuario.
   If conn Is Nothing Then
     AbrirConexion
   End If
   conn.Execute "delete from zbloqueos where codusu=" & vUsu.codigo
   Set conn = Nothing
End Sub
Private Sub mnuSal_Click()
    Unload Me
End Sub

Private Sub mnUsuarios_Click()
    Set conn = Nothing
    If AbrirConexionUsuarios Then
        frmMantenusu.Show vbModal
    End If
    If Not AbrirConexion Then
         MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical, "¡Error!"
         End
    End If
    
End Sub

Private Sub mnUsuariosActivos_Click()
Dim sql As String
Dim I As Integer
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    If CadenaDesdeOtroForm <> "" Then
        I = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vParam.NombreEmpresa & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            sql = RecuperaValor(CadenaDesdeOtroForm, I)
            If sql <> "" Then Me.Tag = Me.Tag & "    - " & sql & vbCrLf
            I = I + 1
        Loop Until sql = ""
        MsgBox Me.Tag, vbExclamation, "Usuarios activos."
    Else
        MsgBox "Ningun usuario, además de usted, conectado a: " & vParam.NombreEmpresa & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation, "Usuarios activos."
    End If
    CadenaDesdeOtroForm = ""
End Sub

Private Sub mnuUtiUniDni_Click()
    Dim respuesta As String
    respuesta = InputBox("Password:")
    If UCase(respuesta) = "JMCE" Then
        frmUniDni.Show vbModal
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    ' ### [DavidV] 10/04/2006: Cambio de Button.Index a Button.Key.
    Select Case Button.key
        Case "extremidades" 'Tipos de extremidades
            mnTiposExtremidades_Click
        Case "ramasgen" 'Ramas genericas
            mnGenerica_Click
        Case "ramasesp" 'Ramas especificas
            mnEspecificas_Click
        Case "tipostrab" 'Tipos de trabajos
            mnTiposTrabajos_Click
        Case "empresas" 'Empresas
            mnEmpresas_Click
        Case "instalaciones" 'Instalaciones
            mnInstalaciones_Click
        Case "operarios" 'Operarios
            mnOperarios_Click
        Case "dosímetros" 'Dosimetros
            mnDosimetros_Click
'        Case 12 'Dosimetros Area
'            mnDosimetrosArea_Click
        Case "dosiscuerpo" 'Dosis homogeneas (cuerpo)
            mnHomogeneas_Click
        Case "dosisnohomo" 'dosis no homogeneas
            mnNoHomogeneas_Click
        Case "dosisárea" 'dosis por area
            mnDosisArea_Click
        Case "fichapers" 'ficha de Personal
            mnFichaPersonal_Click
        Case "sobredosis" ' listado de usuarios con sobredosis
            mnOperariosSobredosis_Click
    End Select
End Sub

Private Sub PonerDatosVisiblesForm()
Dim Cad As String
    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    Cad = Cad & ", " & Format(Now, "d")
    Cad = Cad & " de " & Format(Now, "mmmm")
    Cad = Cad & " de " & Format(Now, "yyyy")
    Cad = "    " & Cad & "    "
    Me.StatusBar1.Panels(5).Text = Cad
    
'    If vEmpresa Is Nothing Then
'        Caption = "SUMINISTROS" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
'    Else

                                                                                                         ' antes vEmpresa.nomempre
 '   If vParam Is Nothing Then
        Caption = "Cálculo y Gestión de Dosimetría Personal" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "      Usuario: " & vUsu.nombre
 '   Else
'        Caption = "DOSIMETRIA" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vParam.NombreEmpresa & "  -    Usuario: " & vUsu.Nombre
'    End If
'
'    End If
    
    
End Sub

Private Sub PonerDatosFormulario()
Dim config As Boolean

    config = (vParam Is Nothing)

    If config Then
        HabilitarSoloPrametros True
    Else


        ' según el permiso de usuario podrá ver distintas opciones
        ' de menu
'        If vUsu.NivelSumi <= 2 Then
'            Me.mnCambioPreciosMasivo.Enabled = True
'            Me.mnCambioPreciosMasivo.Visible = True
'            Me.mnFacturacion.Enabled = True
'            Me.mnFacturacion.Visible = True
'        Else
'            Me.mnCambioPreciosMasivo.Enabled = False
'            Me.mnCambioPreciosMasivo.Visible = False
'            Me.mnFacturacion.Enabled = False
'            Me.mnFacturacion.Visible = False
'        End If

'añadido
    Dim Cad As String
    Dim NF As Integer
    Dim c1 As String

    Cad = App.Path & "\ultempre.dat"
    If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            If Cad <> "" Then
                'El primer pipe es el usuario. Como ya no lo necesito, no toco nada
                
                c1 = RecuperaValor(Cad, 2)

                If c1 = "Empresa Backup" Then
                    Me.mnCopiaSeguridad.Enabled = False
                    Me.mnCopiaSeguridad.Visible = False
                    Me.mnRestaurarCopia.Enabled = False
                    Me.mnRestaurarCopia.Visible = False
                Else
                    Me.mnCopiaSeguridad.Enabled = True
                    Me.mnCopiaSeguridad.Visible = True
                    Me.mnRestaurarCopia.Enabled = True
                    Me.mnRestaurarCopia.Visible = True
                End If
            End If
    End If
    End If

    'Poner datos visible del form
    PonerDatosVisiblesForm
    'Poner opciones de nivel de usuario
    PonerOpcionesMenuGeneral Me
    
    'Habilitar
    If config Then HabilitarSoloPrametros False
    'Panel con el nombre de la empresa
    If Not vParam Is Nothing Then
        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vParam.NombreEmpresa
    Else
        Me.StatusBar1.Panels(2).Text = "Falta configurar"
    End If
End Sub

Private Sub HabilitarSoloPrametros(Habilitar As Boolean)
Dim T As Control
Dim Cad As String

    
    For Each T In Me
        Cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            If LCase(Mid(T.Name, 1, 6)) <> "mnbarr" Then _
                T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar1.Enabled = Habilitar
    mnParametros.Enabled = True
    Me.mnParametros.Enabled = True
    Me.mnConfiguracionAplicacion.Enabled = True
    mnDatos.Enabled = True
    Me.mnuSal.Enabled = True
    
End Sub

Private Sub mnCambioUsuario_Click()

    
    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation, "¡Atención!"
        Exit Sub
    End If
    
    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.codigo

    
    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"

    Set conn = Nothing

    If Not AbrirConexionUsuarios Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical, "¡Error!"
        End
    End If
    frmLogin.Show vbModal
    
    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    
    conn.Close

    
    If AbrirConexion() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical, "¡Error!"
        End
    End If
    
    
'    Set vParam = Nothing
'    Set vEmpresa = Nothing
'    LeerEmpresaParametros
'
'    If Not vParam Is Nothing Then
'        If vParam.HayContabilidad Then
'            If AbrirConexionConta(vParam.NombreHost, vParam.NombreUsuario) = False Then
'                MsgBox "La aplicación no puede continuar sin acceso a la Contabilidad. ", vbCritical
'                End
'            Else
'                ' parametros de contabilidad
'                Set vParamC = Nothing
'                LeerEmpresaParametrosC
'                If vParamC Is Nothing Then
'                      End
'                End If
'            End If
'        End If
'
'        If AbrirConexionGestion(vParam.NombreHost, vParam.NombreUsuario) = False Then
'            MsgBox "La aplicación no puede continuar sin acceso a la Gestión Social. ", vbCritical
'            End
'        End If
'    End If

    PonerDatosFormulario
    
    'Ponemos primera vez a false
    PrimeraVez = True
    Me.SetFocus
    
    Screen.MousePointer = vbDefault
End Sub


'Private Sub mnUsuariosActivos_Click()
'Dim Sql As String
'Dim i As Integer
'    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
'    If CadenaDesdeOtroForm <> "" Then
'        i = 1
'        Me.Tag = "Los siguientes PC's están conectados a: " & vParam.NombreEmpresa & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
'        Do
'            Sql = RecuperaValor(CadenaDesdeOtroForm, i)
'            If Sql <> "" Then Me.Tag = Me.Tag & "    - " & Sql & vbCrLf
'            i = i + 1
'        Loop Until Sql = ""
'        MsgBox Me.Tag, vbExclamation
'    Else
'        MsgBox "Ningun usuario, además de usted, conectado a: " & vParam.NombreEmpresa & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
'    End If
'    CadenaDesdeOtroForm = ""
'End Sub

'Public Function LeerEmpresaParametros()
'        'Abrimos la empresa
''        Set vEmpresa = New Cempresa
''        If vEmpresa.Leer = 1 Then
''            MsgBox "No se han podido cargar datos empresa. Debe confgurar la aplicación.", vbExclamation
''            Set vEmpresa = Nothing
''        End If
'
'        Set vParam = New Cparametros
'        If vParam.Leer() = 1 Then
'            MsgBox "No se han podido cargar los parámetros. Debe configurar la aplicación.", vbExclamation
'            Set vParam = Nothing
'        End If
'End Function

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

