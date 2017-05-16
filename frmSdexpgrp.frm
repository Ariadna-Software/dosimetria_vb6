VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSdexpgrp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supl. / Dtos. Exp. Grupos"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmSdexpgrp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   6840
      MaxLength       =   16
      TabIndex        =   11
      Tag             =   "Importe|N|N|0|9999999999.99|sdexpgrp|impsupdt|#,###,###,##0.00||"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   8880
      MaxLength       =   6
      TabIndex        =   12
      Tag             =   "Porcentaje|N|N|0|100.00|sdexpgrp|porsupdto|##0.00||"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ComboBox cmbAux 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmSdexpgrp.frx":000C
      Left            =   8880
      List            =   "frmSdexpgrp.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Comisión|N|N|0|1|sdexpgrp|comision|0||"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   6840
      MaxLength       =   11
      TabIndex        =   7
      Tag             =   "Nº de Expediente|N|N|0|99999999999|sdexpgrp|numexped|00000000000||"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   4800
      MaskColor       =   &H00000000&
      TabIndex        =   5
      ToolTipText     =   "Buscar destino"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Código de Folleto|N|N|0|999999|sdexpgrp|codfovia|000000||"
      Text            =   "follet"
      Top             =   4920
      Width           =   555
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Código de Empresa|N|N|1|999|sdexpgrp|codempre|000|S|"
      Text            =   "codmepre"
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Height          =   285
      Index           =   7
      Left            =   8880
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Fecha Hasta|F|N|||sdexpgrp|hasfecha|dd/mm/yyyy||"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Height          =   285
      Index           =   6
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "Fecha Desde|F|N|||sdexpgrp|desfecha|dd/mm/yyyy||"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   6840
      MaxLength       =   3
      TabIndex        =   6
      Tag             =   "Código de Supl. / Dto.|N|N|0|999|sdexpgrp|codsuple|000||"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   7680
      TabIndex        =   14
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5040
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "Código de Destino|N|N|0|9999|sdexpgrp|coddesti|0000||"
      Text            =   "dest"
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   2400
      MaskColor       =   &H00000000&
      TabIndex        =   3
      ToolTipText     =   "Buscar folleto"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7860
      TabIndex        =   28
      Top             =   5340
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   29
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   960
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Código de Supl. / Dto. Exp. Grupos|N|N|1|999999|sdexpgrp|codsupdt|000000|S|"
      Text            =   "codsup"
      Top             =   4920
      Width           =   795
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSdexpgrp.frx":0010
      Height          =   4410
      Left            =   120
      TabIndex        =   17
      Top             =   540
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9120
      TabIndex        =   30
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   5175
      Width           =   2385
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   40
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   7
      Left            =   9480
      Picture         =   "frmSdexpgrp.frx":0025
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   6
      Left            =   7440
      Picture         =   "frmSdexpgrp.frx":00B0
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   27
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Porcentaje"
      Height          =   255
      Left            =   8880
      TabIndex        =   26
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Comisión"
      Height          =   255
      Index           =   1
      Left            =   8880
      TabIndex        =   24
      Top             =   1320
      Width           =   900
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   7800
      MousePointer    =   4  'Icon
      Picture         =   "frmSdexpgrp.frx":013B
      ToolTipText     =   "Buscar expediente"
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Supl. / Dto."
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   23
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   8880
      TabIndex        =   22
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   7800
      MousePointer    =   4  'Icon
      Picture         =   "frmSdexpgrp.frx":06C5
      ToolTipText     =   "Buscar supl. / dto."
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   255
      Index           =   9
      Left            =   6840
      TabIndex        =   21
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Expediente"
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   20
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmSdexpgrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: CÈSAR                                         +-+-
' +-+- Menú: Grupos-Suplem./Dtos.-Supl. / Dtos. Exp. Grupos +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFol As frmFollviaj2
Attribute frmFol.VB_VarHelpID = -1
Private WithEvents frmDes As frmDestinos
Attribute frmDes.VB_VarHelpID = -1
Private WithEvents frmSup As frmSupdtogr
Attribute frmSup.VB_VarHelpID = -1
Private WithEvents frmExp As frmExpgrupo2
Attribute frmExp.VB_VarHelpID = -1
' *****************************************************

Private CadenaConsulta As String
Private CadB As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------


Private Sub PonerModo(vModo)
Dim b As Boolean
Dim i As Integer
    
    Modo = vModo
    PonerIndicador lblIndicador, Modo
    
    b = (Modo = 2)
    ' **** posar tots els controls (botons inclosos) que siguen del Grid
    ' si n'hi ha codEmpre, posar a False *****
    txtAux(0).Visible = False 'codempre
    txtAux(1).Visible = Not b
    txtAux(2).Visible = Not b
    btnBuscar(0).Visible = Not b
    txtAux2(2).Visible = Not b
    txtAux(3).Visible = Not b
    btnBuscar(1).Visible = Not b
    txtAux2(3).Visible = Not b
    ' **************************************************
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    For i = 4 To 9
        BloquearTxt txtAux(i), b
    Next i
    'BloquearTxt txtAux2(4), b
    BloquearCmb cmbAux(0), b
    BloquearImgBuscar Me, Modo ' ** si n'hi han imagens de buscar codi fora del grid **
    BloquearImgFec Me, 6, Modo ' ** si n'hi han imagens de buscar data fora del grid **
    BloquearImgFec Me, 7, Modo
    ' ********************************************************

    cmdAceptar.Visible = Not b
    cmdCancelar.Visible = Not b
    DataGrid1.Enabled = b
    
    'Si es retornar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    End If
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari
    
    ' *** bloquejar tota la PK quan estem en modificar  ***
    BloquearTxt txtAux(0), (Modo = 4) 'codEmpre
    BloquearTxt txtAux(1), (Modo = 4)
    ' ******************************************************
    
    ' *** adrede, per a bloquejar importe/porc ***
    If (Modo = 3 Or Modo = 4) And (DataGrid1.Columns(6).Text <> "") Then _
        imp_porc (CInt(DataGrid1.Columns(6).Text))
    ' ********************************************
End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim b As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And Adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b

    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
    'Toolbar1.Buttons(11).Enabled = False
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim i As Integer
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        'NumF = SugerirCodigoSiguienteStr("follviaj", "codfovia")
        NumF = SugerirCodigoSiguienteStr("sdexpgrp", "codsupdt", "codempre=" & codempre)
        'NumF = ""
    End If
    '***************************************************************
    'Situem el grid al final
    AnyadirLinea DataGrid1, Adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = codempre
    txtAux(1).Text = NumF
    FormateaCampo txtAux(1)
    For i = 2 To 7
        txtAux(i).Text = ""
    Next i
    txtAux2(2).Text = ""
    txtAux2(3).Text = ""
    txtAux2(4).Text = ""
    cmbAux(0).ListIndex = 1 'per defecte a si
    txtAux(8).Text = 0
    txtAux(9).Text = 0
    FormateaCampo txtAux(8)
    FormateaCampo txtAux(9)
    ' **************************************************

    LLamaLineas anc, 3
       
    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(2) '**** 1r camp visible que NO siga PK ****
    Else
        PonerFoco txtAux(1) '**** 1r camp visible que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    PonerModo 2
    CadB = ""
End Sub

Private Sub BotonBuscar()
    Dim i As Integer
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    'CargaGrid "codsupdt = -1 AND codempre = " & codEmpre
    CargaGrid "codsupdt = -1"
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    ' repasar codEmpre ******
    txtAux(0).Text = codempre
    For i = 1 To 9
        txtAux(i).Text = ""
    Next i
    txtAux2(2).Text = ""
    txtAux2(3).Text = ""
    txtAux2(4).Text = ""
    cmbAux(0).ListIndex = -1
    ' ****************************************************

    LLamaLineas DataGrid1.Top + 206, 1
    
    ' *** posar el foco al 1r camp visible que siga PK ***
    PonerFoco txtAux(1)
    ' ***************************************************************
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux2(2).Text = DataGrid1.Columns(3).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux2(3).Text = DataGrid1.Columns(5).Text
    ' ********************************************************

    LLamaLineas anc, 4 'modo 4
   
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco txtAux(2)
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Integer
    DeseleccionaGrid
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    'Me.imgFec(2).Top = alto
    For i = 0 To 3
        txtAux(i).Top = alto
    Next i
    txtAux2(2).Top = alto
    txtAux2(3).Top = alto
    btnBuscar(0).Top = alto
    btnBuscar(1).Top = alto
    ' ***************************************************
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Certes comprovacions
    If Adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    '*** canviar la pregunta, els noms dels camps i el DELETE; repasar codEmpre ***
    SQL = "¿Seguro que desea eliminar la Ruta para recogida?"
    'SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    SQL = SQL & vbCrLf & "Código: " & Adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Folleto: " & Adodc1.Recordset.Fields(3)
    SQL = SQL & vbCrLf & "Destino: " & Adodc1.Recordset.Fields(5)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'N'hi ha que eliminar
        NumRegElim = Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from sdexpgrp where codsupdt = " & Adodc1.Recordset!codsupdt & " AND codempre = " & codempre
        Conn.Execute SQL
        'SQL = SQL & " AND codempre = " & codEmpre
    '******************************************************************************
        
        'Conn.Execute SQL
        If CadB <> "" Then
            CargaGrid CadB
            lblIndicador.Caption = "RESULTADO BUSQUEDA"
        Else
            CargaGrid ""
            lblIndicador.Caption = ""
        End If
        temp = SituarDataTrasEliminar(Adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        Adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
    Dim i As Integer
    
    ' *** repasar <=== si els camps es diuen txtAux o Text1 ***
    If Modo = 1 Then 'BUSQUEDA
        For i = 0 To txtAux.Count - 1 ' <===
            With txtAux(i) ' <===
                If .MaxLength <> 0 Then
                   .HelpContextID = .MaxLength
                    .MaxLength = (.HelpContextID * 2) + 1
                End If
            End With
        Next i
    Else
        For i = 0 To txtAux.Count - 1 ' <===
            With txtAux(i) ' <===
                If .HelpContextID <> 0 Then
                    .MaxLength = .HelpContextID
                    .HelpContextID = 0
                End If
            End With
        Next i
    End If
    ' ****************************************************
End Sub

Private Sub cmdAceptar_Click()
Dim i As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
                'If InsertarDesdeForm(Me) Then
                If InsertarDesdeForm2(Me, 0) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not Adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            Adodc1.Recordset.Filter = "codempre = " & txtAux(0).Text & " AND codsupdt = " & txtAux(1).Text
                            'adodc1.Recordset.Filter = "codfovia = " & txtAux(0).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    i = Adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    If CadB <> "" Then
                        CargaGrid CadB
                        lblIndicador.Caption = "RESULTADO BUSQUEDA"
                    Else
                        CargaGrid
                        lblIndicador.Caption = ""
                    End If
                    Adodc1.Recordset.Move i - 1
                End If
            End If
            
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
                lblIndicador.Caption = "RESULTADO BUSQUEDA"
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid CadB
    End Select
    
    If Not Adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
    If CadB <> "" Then
        lblIndicador.Caption = "RESULTADO BUSQUEDA"
    Else
        lblIndicador.Caption = ""
    End If
    
    DataGrid1.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim j As Integer
Dim Aux As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        j = i + 1
        i = InStr(j, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, i - j)
            j = Val(Aux)
            cad = cad & Adodc1.Recordset.Fields(j) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub Form_Load()

    '******* repasar si n'hi ha botó d'imprimir o no******
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Tots
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'Imprimir
        .Buttons(12).Image = 11  'Eixir
    End With
    '*****************************************************

    chkVistaPrevia.Value = CheckValueLeer(Name)
    ' *** SI N'HI HAN COMBOS ***
    CargaCombo 0
    ' **************************
    
    '****************** canviar la consulta *********************************+
    'CadenaConsulta = "SELECT rutasrec.codempre, rutasrec.codiruta, rutasrec.nomruta FROM rutasrec WHERE rutasrec.codempre = " & codEmpre
    CadenaConsulta = "SELECT sdexpgrp.codempre, sdexpgrp.codsupdt, sdexpgrp.codfovia, follviaj.desfovia, sdexpgrp.coddesti, destinos.nomdesti, sdexpgrp.codsuple, supdtogr.nomsuple, sdexpgrp.numexped, sdexpgrp.comision, sdexpgrp.desfecha, sdexpgrp.hasfecha, sdexpgrp.impsupdt, sdexpgrp.porsupdto"
    CadenaConsulta = CadenaConsulta & " FROM sdexpgrp LEFT OUTER JOIN follviaj ON (sdexpgrp.codfovia = follviaj.codfovia AND sdexpgrp.codempre = follviaj.codempre )"
    CadenaConsulta = CadenaConsulta & " LEFT OUTER JOIN destinos ON (sdexpgrp.coddesti = destinos.coddesti )"
    CadenaConsulta = CadenaConsulta & " LEFT OUTER JOIN supdtogr ON (sdexpgrp.codsuple = supdtogr.codsuple AND sdexpgrp.codempre = supdtogr.codempre )"
    CadenaConsulta = CadenaConsulta & " WHERE sdexpgrp.codempre = " & codempre
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    ' ****** Si n'hi han camps fora del grid ******
    'CargaForaGrid
    ' *********************************************
    
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario2(Me, Adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                BotonBuscar
        Case 3
                BotonVerTodos
        Case 6
                BotonAnyadir
        Case 7
                mnModificar_Click
        Case 8
                BotonEliminar
        Case 11 'Imprimir
                printNou
        Case 12 'Salir
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim i As Integer
    Dim SQL As String
    Dim tots As String
    
    Adodc1.ConnectionString = Conn
    ' *** si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
    ' `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " ORDER BY codempre, codsupdt"
    'SQL = SQL & " ORDER BY codfovia"
    '**************************************************************++
    
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    DataGrid1.ScrollBars = dbgNone
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1 ' per a que no ixca l'error de "la fila actual no está disponible"
       
    ' *** posar només els controls del grid ***
    tots = "N||||0|;S|txtAux(1)|T|Cód.|650|;S|txtAux(2)|T|Folleto|800|;S|btnBuscar(0)|B||195|;"
    tots = tots & "S|txtAux2(2)|T|Título Folleto|1957|;S|txtAux(3)|T|Dest.|650|;S|btnBuscar(1)|B||195|;"
    tots = tots & "S|txtAux2(3)|T|Descripción Destino|1957|;"
    For i = 1 To 8
        tots = tots & "N||||0|;"
    Next i
    arregla tots, DataGrid1, Me
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************
    
    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    DataGrid1.Columns(1).Alignment = dbgRight
    DataGrid1.Columns(2).Alignment = dbgRight
    DataGrid1.Columns(4).Alignment = dbgRight
    ' *****************************
    
    
    ' *** Si n'hi han camps fora del grid ***
    If Not Adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    ' **************************************
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
'    If txtAux(Index).Text = "" Then Exit Sub
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case Index
        Case 1
            PonerFormatoEntero txtAux(Index)
        Case 2 'folletos
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = DevuelveDesdeBDnew(1, "follviaj", "desfovia", "codfovia", txtAux(Index).Text, "N", "", "codempre", CStr(codempre), "N")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Folleto: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFol = New frmFollviaj2
                        frmFol.DatosADevolverBusqueda = "1|2|" 'no pose el 0 per a que no torne el codempre
                        frmFol.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmFol.Show vbModal
                        Set frmFol = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
        Case 3 'destinos
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "destinos", "nomdesti")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Destino: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmDes = New frmDestinos
                        frmDes.DatosADevolverBusqueda = "0|1|"
                        frmDes.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmDes.Show vbModal
                        Set frmDes = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
        Case 4 'suplementos
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = DevuelveDesdeBDnew(1, "supdtogr", "nomsuple", "codsuple", txtAux(Index).Text, "N", "", "codempre", CStr(codempre), "N")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Supl./Dto: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSup = New frmSupdtogr
                        frmSup.DatosADevolverBusqueda = "1|2|" 'no pose el 0 per a que no torne el codempre
                        frmSup.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmSup.Show vbModal
                        Set frmSup = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                ' *** adrede, si existix el codi, bloqueja import/porcentaje ***
                Else
                    imp_porc (txtAux(Index).Text)
                ' **************************************************************
                End If
            Else
                txtAux2(Index).Text = ""
            End If
        Case 5 'expediente
            If PonerFormatoEntero(txtAux(Index)) Then
                If PonerNombreDeCod(txtAux(Index), "expgrupo", "numexped") = "" Then
                    MsgBox "No existe el Nº de Expediente: " & txtAux(Index).Text
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
        Case 6, 7 'dates
            If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
        Case 8 'Importe
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1 'tipo1: Decimal(12,2)
        Case 9
            PonerFormatoDecimal txtAux(Index), 4 'tipo 4: Decimal(5,2)
    End Select
    '**************************************************************************
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
Dim i As Integer
' *** només per ad este manteniment ***
Dim RS As Recordset
Dim cad As String
Dim exped As String
' *************************************

    b = CompForm(Me)
    If Not b Then Exit Function


    If b And (Modo = 3) Then
        'Estem insertant
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numèric
         
        ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
'        Datos = DevuelveDesdeBD("codfovia", "follviaj", "codfovia", txtAux(0).Text, "N")
        Datos = DevuelveDesdeBDnew(1, "sdexpgrp", "codsupdt", "codsupdt", txtAux(1).Text, "N", "", "codempre", CStr(codempre), "N")
        
'        cad = "SELECT codrapel FROM rappelcl WHERE codempre = " & codEmpre
'        cad = cad & " AND codrapel = '" & txtAux(0).Text & "'"
'
'        Set RS = New ADODB.Recordset
'        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'        If Not RS.EOF Then
'            Datos = DBLet(RS.Fields(0))
'        Else
'            Datos = ""
'        End If
'        RS.Close
'        Set RS = Nothing
         
        If Datos <> "" Then
            MsgBox "Ya existe el Código de Supl. / Dto. Exp. Grupos: " & txtAux(1).Text, vbExclamation
            DatosOk = False
            PonerFoco txtAux(1) '*** posar el foco al 1r camp visible de la PK de la capçalera ***
            Exit Function
        End If
        '*************************************************************************************
    End If

    ' *** Si cal fer atres comprovacions ***
    If (Modo = 3) Or (Modo = 4) Then 'insertar o modificar
        If CDate(txtAux(7).Text) <= CDate(txtAux(6).Text) Then
            MsgBox "La fecha Hasta debe ser mayor que la fecha Desde", vbExclamation
            DatosOk = False
            Exit Function
        End If
        
        If ((txtAux(8).Text = 0) And (txtAux(9).Text = 0)) Or ((txtAux(8).Text <> 0) And (txtAux(9).Text <> 0)) Then
            MsgBox "Debe introducir el importe o el porcentaje", vbExclamation
            DatosOk = False
            Exit Function
        End If
        
        If (txtAux(5).Text = "") Then
            exped = "is null"
        Else
            exped = "= " & txtAux(5).Text
        End If
        
        cad = "SELECT * FROM sdexpgrp WHERE codempre = " & codempre
        cad = cad & " AND codfovia = " & txtAux(2).Text
        cad = cad & " AND coddesti = " & txtAux(3).Text
        cad = cad & " AND codsuple = " & txtAux(4).Text
        cad = cad & " AND numexped " & exped ' el = està a la variable
        cad = cad & " AND ((desfecha <= """ & Format(txtAux(6).Text, FormatoFecha) & """ AND """ & Format(txtAux(6).Text, FormatoFecha) & """ <= hasfecha)"
        cad = cad & " OR (desfecha <= """ & Format(txtAux(7).Text, FormatoFecha) & """ AND """ & Format(txtAux(7).Text, FormatoFecha) & """ <= hasfecha))"
        cad = cad & " AND codsupdt <> " & txtAux(1).Text

        Set RS = New ADODB.Recordset
        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

        If Not RS.EOF Then
            Datos = DBLet(RS.Fields(0))
        Else
            Datos = ""
        End If
        RS.Close
        Set RS = Nothing

        If Datos <> "" Then
            MsgBox "Ya existe un Suplemento/Descuento para ese Folleto, Destino, Supl./Dto., Expediente y cuyas fechas Desde-Hasta se solapan con las introducidas.", vbExclamation
            DatosOk = False
            PonerFoco txtAux(1)
            Exit Function
        End If
        
    End If
    ' *********************************************

    DatosOk = b
End Function

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    
    DataGrid1.SelStartCol = -1
    DataGrid1.SelEndCol = -1
    Exit Sub
    
EDeseleccionaGrid:
    Err.Clear
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Function SepuedeBorrar() As Boolean
Dim SQL As String
'    SepuedeBorrar = False
'    '**********************canviar parametres de la funcio ^******************
'    SQL = DevuelveDesdeBD("codprodu", "comisclt", "codprodu", adodc1.Recordset!codprodu, "N")
'    If SQL <> "" Then
'        MsgBox "Este Producto Propio está vinculado con registros de la tabla de Comisiones", vbExclamation
'        Exit Function
'    End If
'    '**************************************************************************+
    
    SepuedeBorrar = True
End Function

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 And Modo = 2 Then Unload Me  'ESC
    End If
End Sub

' *** si n'hi han combos ***
Private Sub CargaCombo(Index As Integer)
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    Select Case Index
        Case 0 'cuentas bancarias
            cmbAux(Index).AddItem "Si"
            cmbAux(Index).ItemData(cmbAux(i).NewIndex) = 1
            cmbAux(Index).AddItem "No"
            cmbAux(Index).ItemData(cmbAux(i).NewIndex) = 0
            
            cmbAux(Index).ListIndex = 0 'per defecte a si
    End Select
End Sub

Private Sub SelComboBool(valor As Integer, combo As ComboBox)
    Dim i As Integer
    Dim j As Integer

    i = valor
    For j = 0 To combo.ListCount - 1
        If combo.ItemData(j) = i Then
            combo.ListIndex = j
            Exit For
        End If
    Next j
End Sub
' ******************************


' ********** SI N'HI HAN CAMPS FORA DEL GRID ******************

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Modo <> 4 Then 'Modificar
        CargaForaGrid
    Else
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
End Sub

Private Sub CargaForaGrid()
Dim i As Integer
'Dim tipclien

    If DataGrid1.Columns.Count > 2 Then
        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
        txtAux(4) = DataGrid1.Columns(6).Text
        txtAux2(4) = DataGrid1.Columns(7).Text
        txtAux(5) = DataGrid1.Columns(8).Text
        If DataGrid1.Columns(9).Text <> "" Then _
            SelComboBool CInt(DataGrid1.Columns(9).Text), cmbAux(0)
        txtAux(6) = DataGrid1.Columns(10).Text
        txtAux(7) = DataGrid1.Columns(11).Text
        txtAux(8) = DataGrid1.Columns(12).Text
        txtAux(9) = DataGrid1.Columns(13).Text
        
        PonerFormatoEntero txtAux(4)
        PonerFormatoEntero txtAux(5)
        If txtAux(6).Text <> "" Then PonerFormatoFecha txtAux(6)
        If txtAux(7).Text <> "" Then PonerFormatoFecha txtAux(7)
        If txtAux(8).Text <> "" Then PonerFormatoDecimal txtAux(8), 1 'tipo1: Decimal(12,2)
        PonerFormatoDecimal txtAux(9), 4 'tipo 4: Decimal(5,2)
        ' ****************************************************************************

        ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
        'txtAux2(4).Text = PonerNombreDeCod(txtAux(4), "poblacio", "despobla", "codpobla", "N")
        If txtAux(4).Text <> "" Then _
            txtAux2(4).Text = DevuelveDesdeBDnew(1, "supdtogr", "nomsuple", "codsuple", txtAux(4).Text, "N", "", "codempre", CStr(codempre), "N")
        ' **********************************************************************
    End If
End Sub

Private Sub LimpiarCampos()
Dim i As Integer
On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    For i = 4 To 9
        txtAux(i).Text = ""
    Next i
    txtAux2(4).Text = ""
    cmbAux(0).ListIndex = -1
    ' ****************************************************

    If Err.Number <> 0 Then Err.Clear
End Sub
' ******************************************************************

' *** si n'hi ha buscar data, posar a les <=== el menor index de les imagens de buscar data ***
' NOTA: ha de coincidir l'index de la image en el del camp a on va a parar el valor
Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend

    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    ' *** elegir si es desplega cap a la dreta o cap a l'esquerra ***
    'DRETA
    'frmC.Left = esq + imgFec(Index).Parent.Left + 30
    'ESQUERRA
    frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
    ' ****************************************************************

    ' *** elegir si es desplega cap a la dreta o cap a l'esquerra ***
    'BAIX
    'frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    'DALT
    frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25
    ' ***************************************************************

    imgFec(6).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtAux(Index).Text <> "" Then frmC.NovaData = txtAux(Index).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtAux(CByte(imgFec(6).Tag)) '<===
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtAux(CByte(imgFec(6).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub
' *****************************************************

' *** si n'hi han botons de buscar codi al grid ***
Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Folletos
            Set frmFol = New frmFollviaj2
            frmFol.DatosADevolverBusqueda = "1|2|" 'no pose el 0 per a no retornar el codempre
            frmFol.Show vbModal
            Set frmFol = Nothing
            PonerFoco txtAux(2)
        Case 1 'Destinos
            Set frmDes = New frmDestinos
            frmDes.DatosADevolverBusqueda = "0|1|"
            frmDes.Show vbModal
            Set frmDes = Nothing
            PonerFoco txtAux(3)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adodc1, 1
End Sub

Private Sub frmFol_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(2)
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmDes_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(3).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(3)
    txtAux2(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub
' ************************************************

' *** si n'hi han botons de buscar codi fora del grid ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Suplementos
            Set frmSup = New frmSupdtogr
            frmSup.DatosADevolverBusqueda = "1|2|" 'no pose el 0 per a no retornar el codempre
            frmSup.Show vbModal
            Set frmSup = Nothing
            PonerFoco txtAux(4)
        Case 1 'Expediente
            Set frmExp = New frmExpgrupo2
            frmExp.DatosADevolverBusqueda = "0|"
            frmExp.Show vbModal
            Set frmExp = Nothing
            PonerFoco txtAux(5)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adodc1, 1
End Sub

Private Sub frmSup_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(4)
    txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 2)
    '*** adrede ***
    imp_porc (CInt(txtAux(4).Text))
    ' *************
End Sub

Private Sub frmExp_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(5)
End Sub
' ************************************************

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "sdexpgrp"
        .Informe2 = "rSdexpgrp.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(Adodc1, Me)
        ' *** repasar codEmpre ***
        '.cadTodosReg = ""
        .cadTodosReg = "{sdexpgrp.codempre} = " & codempre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={sdexpgrp.codsupdt}|"
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

' *** adrede ***
Private Sub imp_porc(cod As Integer)
    
    Dim tipo As Integer
    
    tipo = DevuelveDesdeBDnew(1, "supdtogr", "tipsuple", "codsuple", CStr(cod), "N", "", "codempre", CStr(codempre), "N")
    
    If (tipo = 0) Then 'suplemento
        BloquearTxt txtAux(8), False
        BloquearTxt txtAux(9), True
        txtAux(9).Text = 0
        FormateaCampo txtAux(9)
    ElseIf (tipo = 1) Then 'descuento
        BloquearTxt txtAux(8), True
        txtAux(8).Text = 0
        FormateaCampo txtAux(8)
        BloquearTxt txtAux(9), False
    End If
End Sub
' *************
