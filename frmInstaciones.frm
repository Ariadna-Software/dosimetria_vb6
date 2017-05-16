VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInstalaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instalaciones"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8520
   Icon            =   "frmInstaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   8520
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6930
      TabIndex        =   7
      Tag             =   "Tipo Dosimetria|N|N|||instalaciones|c_tipo||N|"
      Text            =   "Combo2"
      Top             =   1380
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "Migrado|T|S|||instalaciones|migrado|||"
      Text            =   "Text1"
      Top             =   1380
      Width           =   405
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   90
      TabIndex        =   28
      Top             =   4710
      Width           =   8205
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2610
         MaxLength       =   30
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   570
         Width           =   5340
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   15
         Left            =   2130
         MaxLength       =   5
         TabIndex        =   18
         Tag             =   "Rama Específica|T|N|||instalaciones|rama_especifica|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2610
         MaxLength       =   30
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   240
         Width           =   5340
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   14
         Left            =   2130
         MaxLength       =   5
         TabIndex        =   17
         Tag             =   "Rama Generica|T|N|||instalaciones|rama_gen|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Rama Específica:"
         Height          =   255
         Left            =   270
         TabIndex        =   45
         Top             =   570
         Width           =   1335
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   5
         Left            =   1830
         MouseIcon       =   "frmInstaciones.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmInstaciones.frx":0E1C
         ToolTipText     =   "Buscar rama específica"
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Rama Genérica:"
         Height          =   255
         Left            =   270
         TabIndex        =   43
         Top             =   240
         Width           =   1155
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   4
         Left            =   1830
         MouseIcon       =   "frmInstaciones.frx":0F1E
         MousePointer    =   99  'Custom
         Picture         =   "frmInstaciones.frx":1070
         ToolTipText     =   "Buscar rama genérica"
         Top             =   270
         Width           =   240
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1605
      Left            =   90
      TabIndex        =   34
      Top             =   3060
      Width           =   8205
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "Observaciones|T|S|||instalaciones|observaciones|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   6270
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1710
         MaxLength       =   100
         TabIndex        =   15
         Tag             =   "Persona Contacto|T|S|||instalaciones|persona_contacto|||"
         Text            =   "Text1"
         Top             =   885
         Width           =   6270
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "Mail|T|S|||instalaciones|mail_internet|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   6270
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   3810
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Fax|T|S|||instalaciones|fax|||"
         Text            =   "Text1"
         Top             =   210
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Telefono|T|S|||instalaciones|telefono|||"
         Text            =   "Text1"
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label Label6 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   255
         TabIndex        =   47
         Top             =   1215
         Width           =   1305
      End
      Begin VB.Label Label15 
         Caption         =   "Pers.Contacto:"
         Height          =   255
         Left            =   255
         TabIndex        =   38
         Top             =   885
         Width           =   1140
      End
      Begin VB.Label Label16 
         Caption         =   "Mail:"
         Height          =   255
         Left            =   255
         TabIndex        =   37
         Top             =   555
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   2910
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   255
         TabIndex        =   35
         Top             =   225
         Width           =   930
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   12
      Left            =   4170
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "Fecha Baja|F|S|||instalaciones|f_baja|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   1380
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   1305
      Left            =   90
      TabIndex        =   29
      Top             =   1710
      Width           =   8205
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "Distrito|T|S|||instalaciones|distrito|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "Direccion|T|S|||instalaciones|direccion|||"
         Text            =   "Text1"
         Top             =   210
         Width           =   6270
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   9
         Tag             =   "C.Postal|T|N|||instalaciones|c_postal|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   3810
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "Poblacion|T|S|||instalaciones|poblacion|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   4170
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3810
         MaxLength       =   30
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   930
         Width           =   4170
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   3
         Left            =   1410
         MouseIcon       =   "frmInstaciones.frx":1172
         MousePointer    =   99  'Custom
         Picture         =   "frmInstaciones.frx":12C4
         ToolTipText     =   "Buscar código postal"
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Código Postal:"
         Height          =   255
         Left            =   270
         TabIndex        =   33
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   270
         TabIndex        =   32
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Poblacion:"
         Height          =   255
         Left            =   2910
         TabIndex        =   31
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label13 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   2910
         TabIndex        =   30
         Top             =   945
         Width           =   930
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "Fecha Alta|F|N|||instalaciones|f_alta|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   1380
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3000
      MaxLength       =   30
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   570
      Width           =   4350
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   13
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Codigo Instalación|T|N|||instalaciones|c_instalacion||S|"
      Text            =   "Text1"
      Top             =   990
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "Descripción|T|N|||instalaciones|descripcion|||"
      Text            =   "Text1"
      Top             =   990
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7170
      TabIndex        =   21
      Top             =   5880
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   0
      Left            =   1590
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Codigo Empresa|T|N|||instalaciones|c_empresa||S|"
      Text            =   "Text1"
      Top             =   570
      Width           =   1305
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7170
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   90
      TabIndex        =   23
      Top             =   5820
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   210
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6015
      TabIndex        =   20
      Top             =   5880
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   360
      Top             =   5910
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
      TabIndex        =   26
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "L"
                  Object.Tag             =   "l"
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
         Left            =   5250
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Dosim.:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   5
      Left            =   7410
      TabIndex        =   48
      Top             =   1140
      Width           =   885
   End
   Begin VB.Image ImgPpal 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   6
      Left            =   7740
      MouseIcon       =   "frmInstaciones.frx":13C6
      MousePointer    =   99  'Custom
      Picture         =   "frmInstaciones.frx":16D0
      ToolTipText     =   "Instalaciones del Operario"
      Top             =   510
      Width           =   540
   End
   Begin VB.Image ImgPpal 
      Height          =   240
      Index           =   0
      Left            =   1320
      MouseIcon       =   "frmInstaciones.frx":3052
      MousePointer    =   99  'Custom
      Picture         =   "frmInstaciones.frx":31A4
      ToolTipText     =   "Seleccionar fecha"
      Top             =   1380
      Width           =   240
   End
   Begin VB.Image ImgPpal 
      Height          =   240
      Index           =   1
      Left            =   3900
      MouseIcon       =   "frmInstaciones.frx":322F
      MousePointer    =   99  'Custom
      Picture         =   "frmInstaciones.frx":3381
      ToolTipText     =   "Seleccionar fecha"
      Top             =   1380
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Migrado:"
      Height          =   255
      Left            =   5400
      TabIndex        =   46
      Top             =   1410
      Width           =   720
   End
   Begin VB.Label Label19 
      Caption         =   "Fecha Baja"
      Height          =   255
      Left            =   3030
      TabIndex        =   41
      Top             =   1410
      Width           =   900
   End
   Begin VB.Image ImgPpal 
      Height          =   240
      Index           =   2
      Left            =   1320
      MouseIcon       =   "frmInstaciones.frx":340C
      MousePointer    =   99  'Custom
      Picture         =   "frmInstaciones.frx":355E
      ToolTipText     =   "Buscar empresa"
      Top             =   600
      Width           =   240
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
      Left            =   210
      TabIndex        =   39
      Top             =   570
      Width           =   1275
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha Alta"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   210
      TabIndex        =   27
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
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
      Left            =   210
      TabIndex        =   25
      Top             =   990
      Width           =   735
   End
End
Attribute VB_Name = "frmInstalaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmOpeIns As frmOperariosInstala
Attribute frmOpeIns.VB_VarHelpID = -1

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

Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
   If Modo = 2 Or Modo = 0 Then
        Select Case KeyCode
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
            Case vbImprimir
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(11)
            Case vbSalir
                  Toolbar1_ButtonClick Toolbar1.Buttons(12)
        End Select
   End If


End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                'MsgBox "Registro insertado.", vbInformation
                PonerModo 2
'                If SituarData1 Then
'                    lblIndicador.Caption = ""
'                    PonerModo 2
'                Else
'                    LimpiarCampos
'                    PonerModo 0
'                End If
            End If
        End If
    Case 4
        'Modificar
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If ModificaDesdeFormulario(Me, 1) Then
                If SituarData1 Then
                    lblIndicador.Caption = ""
                    PonerModo 2
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
            PonerModo 2
            PonerCampos
        End Select
End Sub

' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
        
    Dim sql As String
    Dim Empresa As String
    Dim instalacion As String
    Dim fecha As Date
    
    On Error GoTo ESituarData1
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            Empresa = Text1(0).Text
            instalacion = Text1(13).Text
            fecha = Format(Text1(11).Text, FormatoFecha)
            If (Trim(Data1.Recordset.Fields!c_empresa) = Trim(Empresa)) Then
               If Trim(Data1.Recordset.Fields!c_instalacion) = Trim(instalacion) Then
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
    DespalzamientoVisible False
    '###A mano
    PonerFoco Text1(0)
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
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
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
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    Text1(13).Locked = True
    Text1(13).BackColor = &H80000018
    'Text1(11).Locked = True
    'Text1(11).BackColor = &H80000018
    'ImgPpal(0).Enabled = False
    ImgPpal(2).Enabled = False
    DespalzamientoVisible False
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    '******* canviar el mensage i la cadena *********************
    Cad = "Seguro que desea eliminar la instalación:"
    Cad = Cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0) & "-" & Data1.Recordset.Fields(1)
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(4)
    Cad = Cad & vbCrLf & " de fecha de alta : " & Data1.Recordset.Fields(2)
    '**********************************************************
    I = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!")
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        Cad = "delete from instalaciones where c_empresa = '"
        Cad = Cad & Data1.Recordset.Fields(0) & "' and c_instalacion = '"
        Cad = Cad & Data1.Recordset.Fields(1) & "' and f_alta = '"
        Cad = Cad & Format(Data1.Recordset.Fields(2), FormatoFecha) & "'"
        
        Conn.Execute Cad
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            PonerModo 0
            Else
                Data1.Recordset.MoveFirst
                NumRegElim = NumRegElim - 1
                If NumRegElim > 1 Then
                    For I = 1 To NumRegElim - 1
                        Data1.Recordset.MoveNext
                    Next I
                End If
                PonerCampos
        End If
    End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Instalaciones"
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
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        '.Buttons(10).Image = 10
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
    NombreTabla = "instalaciones"
    Ordenacion = " ORDER BY c_empresa,c_instalacion,f_alta"
    '******************************************************+
        
'    PonerOpcionesMenu
    
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    
    CargarCombo
    
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
'    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
    End If
    
End Sub

Private Sub LimpiarCampos()
    Limpiar Me
    Combo3.ListIndex = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        Aux = Aux & " and " & ValorDevueltoFormGrid(Text1(13), CadenaDevuelta, 2)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla
        If CadB <> "" Then
            CadenaConsulta = CadenaConsulta & " WHERE  " & CadB & " " & Ordenacion
        Else
            CadenaConsulta = CadenaConsulta & " " & Ordenacion
        End If
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmEmp_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmrge_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmREs_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Trim(Text1(14).Text) & "|", "T|", 1)
        Text1(15).Text = RecuperaValor(CadenaSeleccion, 2)
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
            f = Now
            If Text1(11).Text <> "" Then
                If IsDate(Text1(11).Text) Then f = Text1(11).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(11).Text = frmC.fecha
                mTag.DarFormato Text1(11)
            End If
            Set frmC = Nothing
       Case 1
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(12).Text <> "" Then
                If IsDate(Text1(12).Text) Then f = Text1(12).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(12).Text = frmC.fecha
                mTag.DarFormato Text1(12)
            End If
            Set frmC = Nothing
        Case 2 ' codigo de empresa
            Set frmEmp = New frmEmpresas
            frmEmp.DatosADevolverBusqueda = "0|1|"
            frmEmp.Show
        Case 3 ' codigo de provincia
            Set frmPro = New frmProvincias
            frmPro.DatosADevolverBusqueda = "0|1|"
            frmPro.Show
        Case 4 ' rama generica
            Set frmRGe = New frmRamasGener
            frmRGe.DatosADevolverBusqueda = "0|1|"
            frmRGe.Show
        Case 5 ' rama especifica
            Set frmREs = New frmRamasEspe
            frmREs.DatosADevolverBusqueda = "0|1|2|3|4|"
            frmREs.Show
        Case 6 ' operarios de la instalacion
            'solo tenemos el registro visualizado accederemos a los operarios instalaciones
            If Modo <> 2 Then Exit Sub
            
            Set frmOpeIns = New frmOperariosInstala
            frmOpeIns.Empresa = Trim(Text1(0).Text)
            frmOpeIns.instalacion = Trim(Text1(13).Text)
            frmOpeIns.Show 'vbModal
   End Select
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
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
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
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" Then Exit Sub
    
    If Modo = 1 And ConCaracteresBusqueda(Text1(Index).Text) Then Exit Sub
    
    Select Case Index
        Case 0, 1, 3, 5, 6, 7, 8, 10, 13, 14, 15
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            If Modo = 1 Then Exit Sub
            If Modo <> 1 And Text1(Index).Text = "" Then
                PonerFoco Text1(Index)
                MsgBox "Este campo requiere un valor", vbExclamation, "¡Error!"
            End If
            If Index = 3 Or Index = 0 Or Index = 14 Or Index = 15 Then
                If Text1(Index).Text <> "" And Modo <> 1 Then
                    Select Case Index
                        Case 3
                            Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
                            If Text2(0).Text = "" Then
                                MsgBox "Código de provincia no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                        Case 0
                            Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
                            If Text2(1).Text = "" Then
                                MsgBox "El código de empresa no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                        Case 14
                            Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(14).Text & "|", "T|", 1)
                            If Text2(2).Text = "" Then
                                MsgBox "El Código de rama genérica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                        Case 15
                            If Text1(14).Text <> "" Then
                                Text2(3).Text = DevuelveDesdeBD(1, "descripcion", "ramaespe", "cod_rama_gen|c_rama_especifica|", Text1(14).Text & "|" & Text1(15).Text & "|", "T|T|", 2)
                                If Text2(3).Text = "" Then
                                    MsgBox "El código de rama específica no existe. Reintroduzca.", vbExclamation, "¡Error!"
                                    Text1(Index).Text = ""
                                    Text1(14).Text = ""
                                    PonerFoco Text1(14)
                                End If
                            End If
                    End Select
                    
                End If
            End If
              
        Case 11, 12 ' campos de fechas
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
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String
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
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 20, "Empresa")
        Cad = Cad & ParaGrid(Text1(13), 15, "Código")
        Cad = Cad & ParaGrid(Text1(11), 15, "F.Alta")
        Cad = Cad & ParaGrid(Text1(1), 50, "Nombre")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|3|"
            frmB.vTitulo = "Instalaciones"
            frmB.vSelElem = 0
            frmB.vConexionGrid = 1
            'frmB.vBuscaPrevia = chkVistaPrevia
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
End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation, "¡Error!"
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
    
    PonerCamposForma Me, Data1
    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
    Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(14).Text & "|", "T|", 1)
    Text2(3).Text = DevuelveDesdeBD(1, "descripcion", "ramaespe", "cod_rama_gen|c_rama_especifica|", Text1(14).Text & "|" & Text1(15).Text & "|", "T|T|", 2)
    
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
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For I = 0 To Text1.Count - 1
            Text1(I).BackColor = vbWhite
        Next I
    End If
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    'Modificar
    Toolbar1.Buttons(7).Enabled = b 'And vUsu.NivelUsu <= 2
 '   mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b 'And vUsu.NivelUsu <= 2
    Toolbar1.Buttons(11).Enabled = (Modo = 2)

'    mnModificar.Enabled = b
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    Else
        cmdRegresar.Visible = False
    End If
    
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = b Or Modo = 1
    cmdCancelar.Visible = b Or Modo = 1
'    mnOpciones.Enabled = Not b
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(6).Enabled = Not b And Modo <> 1 'And vUsu.NivelUsu <= 2
    Toolbar1.Buttons(1).Enabled = Not b And Modo <> 1
    Toolbar1.Buttons(2).Enabled = Not b And Modo <> 1
    
    b = (Modo = 2) Or Modo = 0
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = b
        Text1(I).BackColor = vbWhite
    Next I
    Combo3.Enabled = Not b
    
    If Modo = 3 Then ValoresPorDefecto
    
    
    For I = 0 To ImgPpal.Count - 1
        ImgPpal(I).Enabled = Not b
    Next I
    
    ImgPpal(6).Enabled = (Modo <> 0)
    PonerFoco chkVistaPrevia
    
    
End Sub

Private Function DatosOk() As Boolean
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim Datos As String
Dim Mens As String

    b = CompForm(Me)
    
    If (b = True) And ((Modo = 3) Or (Modo = 4)) Then
'        For I = 0 To Text1.Count - 1
'             If InStr(1, Text1(I).Text, "'") > 0 Then
'                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
'                DatosOk = False
'                Exit Function
'             End If
'        Next I
        'provincia
        Datos = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
        If Datos = "" Then
            MsgBox "No existe la provincia.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
        'empresa
        Datos = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
        If Datos = "" Then
            MsgBox "No existe la empresa.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
        'rama generica
        Datos = DevuelveDesdeBD(1, "descripcion", "ramagene", "cod_rama_gen|", Text1(14).Text & "|", "T|", 1)
        If Datos = "" Then
            MsgBox "No existe la rama genérica.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
        'rama especifica
        Datos = DevuelveDesdeBD(1, "descripcion", "ramaespe", "cod_rama_gen|c_rama_especifica|", Text1(14).Text & "|" & Text1(15).Text & "|", "T|T|", 2)
        If Datos = "" Then
            MsgBox "No existe la rama específica.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
        If Modo = 3 Then
             Datos = DevuelveDesdeBD(1, "c_instalacion", "instalaciones", "c_empresa|c_instalacion|f_alta|", Text1(0).Text & "|" & Text1(13).Text & "|" & Text1(11).Text & "|", "T|T|F|", 3)
             If Datos <> "" Then
                MsgBox "Ya existe el código de instalación : " & Text1(0).Text & " - " & Text1(3).Text & " de fecha alta : " & Format(Text1(11).Text, FormatoFecha), vbExclamation, "¡Error!"
                DatosOk = False
                Exit Function
             End If
        End If
        
        ' el tipo de dosimetria ha de ser igual en instalacion y empresa
        If Modo = 3 Or Modo = 4 Then
            Datos = DevuelveDesdeBD(1, "c_tipo", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
            If Datos <> "" Then
                If CByte(Datos) <> CByte(Combo3.ListIndex) Then
                    If MsgBox("La instalación no pertenece al mismo tipo de dosimetria que la empresa a la que pertenece. Desea continuar.", vbQuestion + vbYesNo + vbDefaultButton2, "¡Error!") = vbNo Then
                        DatosOk = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    DatosOk = b
End Function


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

Private Sub ValoresPorDefecto()
    Text1(3).Text = "46"
    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
    Text1(10).Text = "mail"
    Text1(11).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu
        Case "Listado"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 2 'Listado de instalaciones
            FrmListado.Show
        Case "Etiquetas"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 25 'Listado de etiquetas de instalaciones
            FrmListado.Show
    End Select
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
    Combo3.AddItem "Personal"
    Combo3.ItemData(Combo3.NewIndex) = 0
    
    Combo3.AddItem "Area"
    Combo3.ItemData(Combo3.NewIndex) = 1
    
    Combo3.AddItem "Personal/Area"
    Combo3.ItemData(Combo3.NewIndex) = 2
    
End Sub

