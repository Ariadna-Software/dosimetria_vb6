VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDosimOrganos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dos�metros a Organo"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7980
   Icon            =   "frmDosimOrganos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   7980
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   8
      Left            =   6210
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "N.Dosimetro|T|N|||dosimorganos|n_dosimetro||N|"
      Text            =   "Text1"
      Top             =   480
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   7
      Left            =   2580
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "N.Registro|N|N|||dosimorganos|n_reg_dosimetro||S|"
      Text            =   "Text1"
      Top             =   480
      Width           =   1305
   End
   Begin VB.Frame Frame7 
      Height          =   1065
      Left            =   330
      TabIndex        =   26
      Top             =   750
      Width           =   7425
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   40
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
         TabIndex        =   3
         Tag             =   "Codigo Empresa|T|N|||dosimorganos|c_empresa|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3180
         MaxLength       =   40
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   600
         Width           =   4005
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "Codigo Instalaci�n|T|N|||dosimorganos|c_instalacion|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   1305
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   2
         Left            =   1470
         MouseIcon       =   "frmDosimOrganos.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimOrganos.frx":0E1C
         ToolTipText     =   "Buscar socio"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   180
         TabIndex        =   41
         Top             =   240
         Width           =   915
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   4
         Left            =   1470
         MouseIcon       =   "frmDosimOrganos.frx":0F1E
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimOrganos.frx":1070
         ToolTipText     =   "Buscar socio"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Instalaci�n"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   180
         TabIndex        =   37
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1665
      Left            =   330
      TabIndex        =   18
      Top             =   1830
      Width           =   7425
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "DNI|T|N|||dosimorganos|dni_usuario|||"
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
         TabIndex        =   21
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
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   900
         Width           =   5325
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   570
         Width           =   5325
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   5
         Left            =   1470
         MouseIcon       =   "frmDosimOrganos.frx":1172
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimOrganos.frx":12C4
         ToolTipText     =   "Buscar socio"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "D.N.I."
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   2910
         TabIndex        =   25
         Top             =   945
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   1290
         Width           =   1290
      End
      Begin VB.Label Label9 
         Caption         =   "Primer Apellido:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label17 
         Caption         =   "Segundo Apellido:"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   960
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   330
      TabIndex        =   27
      Top             =   3510
      Width           =   7425
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5700
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   570
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   5700
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Mes Par/Impar|T|N|||dosimorganos|mes_p_i|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   5700
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Fecha Retirada|F|S|||dosimorganos|f_retirada|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   930
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Fecha Asignacion|F|N|||dosimorganos|f_asig_dosimetro|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   990
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1830
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   570
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Plantilla/Contrata|T|N|||dosimorganos|plantilla_contrata|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3240
         MaxLength       =   40
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   210
         Width           =   3885
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   11
         Tag             =   "Observaciones|T|S|||dosimorganos|observaciones|||"
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
         TabIndex        =   6
         Tag             =   "Tipo de Trabajo|T|N|||dosimorganos|c_tipo_trabajo|||"
         Text            =   "Text1"
         Top             =   210
         Width           =   1320
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmDosimOrganos.frx":13C6
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimOrganos.frx":1518
         ToolTipText     =   "Seleccionar fecha"
         Top             =   990
         Width           =   240
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   1
         Left            =   5430
         MouseIcon       =   "frmDosimOrganos.frx":15A3
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimOrganos.frx":16F5
         ToolTipText     =   "Seleccionar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Mes Par/Impar"
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         Top             =   570
         Width           =   1500
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Retirada"
         Height          =   255
         Left            =   4230
         TabIndex        =   35
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Asignaci�n"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   34
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Plantilla/Contrata:"
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   600
         Width           =   1905
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   3
         Left            =   1500
         MouseIcon       =   "frmDosimOrganos.frx":1780
         MousePointer    =   99  'Custom
         Picture         =   "frmDosimOrganos.frx":18D2
         ToolTipText     =   "Buscar socio"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   1380
         Width           =   1560
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Trabajo:"
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6660
      TabIndex        =   14
      Top             =   5580
      Width           =   1110
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   5550
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   330
      TabIndex        =   15
      Top             =   5460
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
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
      TabIndex        =   17
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
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
      Left            =   4890
      TabIndex        =   43
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "N�mero de Registro"
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
      Left            =   450
      TabIndex        =   42
      Top             =   480
      Width           =   1995
   End
End
Attribute VB_Name = "frmDosimOrganos"
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
Dim i As Integer
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

Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
    AsignarTeclasFuncion KeyCode
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim i As Integer
    Dim v_aux As Integer
    Dim SQL As String
    
    
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
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
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
    'A�adiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    '###A mano
    Text1(7).Text = NumF
    Text1(6).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(8)
End Sub

Private Function SugerirCodigoSiguiente() As String
    Dim SQL As String
    Dim RS As Adodb.Recordset
    
    SQL = "Select Max(n_reg_dosimetro) from dosimorganos"
    
    Set RS = New Adodb.Recordset
    RS.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            SQL = CStr(RS.Fields(0) + 1)
        End If
    End If
    RS.Close
    SugerirCodigoSiguiente = SQL
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
    Dim i As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    '******* canviar el mensage i la cadena *********************
    Cad = "Seguro que desea eliminar el dosimetro:" & Data1.Recordset.Fields(0)
    '**********************************************************
    i = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2)
    
   'Borramos
    If i <> vbYes Then
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
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Dos�metro a �rganos"
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

If Data1.Recordset.EOF Then
    MsgBox "Ning�n registro devuelto.", vbExclamation
    Exit Sub
End If

Cad = ""
i = 0
Do
    J = i + 1
    i = InStr(J, DatosADevolverBusqueda, "|")
    If i > 0 Then
        Aux = Mid(DatosADevolverBusqueda, J, i - J)
        J = Val(Aux)
        Cad = Cad & Text1(J).Text & "|"
    End If
Loop Until i = 0
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

Private Sub Combo1_LostFocus()
    If Combo1.ListIndex = 0 Then Text1(1).Text = "01"
    If Combo1.ListIndex = 1 Then Text1(1).Text = "02"
End Sub


Private Sub Combo2_LostFocus()
    If Combo2.ListIndex = 0 Then Text1(5).Text = "P"
    If Combo2.ListIndex = 1 Then Text1(5).Text = "I"
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer


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
    
    
    '***** canviar el nom de la taula i el ORDER BY ********
    NombreTabla = "dosimorganos"
    Ordenacion = " ORDER BY n_reg_dosimetro"
    '******************************************************+
        
    PonerOpcionesMenu
    
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
    
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    
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
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmIns_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(2).Text = RecuperaValor(CadenaSeleccion, 2)
        Text2(2).Text = RecuperaValor(CadenaSeleccion, 3)
    End If
End Sub

Private Sub frmOpe_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(9).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(3).Text = RecuperaValor(CadenaSeleccion, 2)
        Text2(4).Text = RecuperaValor(CadenaSeleccion, 3)
        Text2(5).Text = RecuperaValor(CadenaSeleccion, 4)
    End If
End Sub

Private Sub frmTTR_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
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
       Case 0
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(6).Text <> "" Then
                If IsDate(Text1(6).Text) Then f = Text1(6).Text
            End If
            Set frmC = New frmCal
            frmC.Fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(6).Text = frmC.Fecha
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
            frmC.Fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(10).Text = frmC.Fecha
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
            frmOpe.DatosADevolverBusqueda = "13|17|14|6|"
            frmOpe.Show
   End Select
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim SQL As String

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
    Dim i As Integer
    Dim SQL As String
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
        Case 0, 2, 3, 4, 8, 9
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el car�cter ' en ning�n campo de texto", vbExclamation
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, ">")
            
            Select Case Index
                Case 0 'empresa
                    Text2(1).Text = ""
                    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
                    If Text2(1).Text = "" Then
                        MsgBox "El c�digo de empresa no existe. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                Case 4 ' tipos de trabajo
                    If Text1(Index).Text <> "" Then
                        Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "tipostrab", "c_tipo_trabajo|", Text1(Index).Text & "|", "T|", 1)
                        If Text2(0).Text = "" Then
                            MsgBox "El c�digo de tipo de trabajo no existe. Reintroduzca.", vbExclamation
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                Case 5 ' instalacion
                    If Text1(Index).Text <> "" And Text1(0).Text <> "" Then
                        Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|f_alta|", Text1(0).Text & "|" & Text1(Index).Text & "|", "T|T|", 2)
                        If Text2(2).Text = "" Then
                            MsgBox "El c�digo de instalacion no existe. Reintroduzca.", vbExclamation
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
            End Select
        Case 6, 10
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
            End If
    End Select
    
    '---
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String
CadB = ObtenerBusqueda(Me)

If CadB = "" Then
    MsgBox vbCrLf & "  Debe introducir alguna condici�n de b�squeda. " & vbCrLf, vbExclamation
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
Dim titulo As String
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
    MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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
    Dim i As Integer
    Dim mTag As CTag
    Dim SQL As String
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
    
    CargarDatosOperarios Text1(9).Text, ape1, ape2, nombre
    Text2(3).Text = ape1
    Text2(4).Text = ape2
    Text2(5).Text = nombre
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim i As Integer
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
        For i = 0 To Text1.Count - 1
            Text1(i).BackColor = &H80000018
        Next i
    End If
    
    b = (Modo = 0) Or (Modo = 2)
    Toolbar1.Buttons(6).Enabled = (b And vUsu.NivelSumi <= 2)
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    Toolbar1.Buttons(11).Enabled = b
    
    'Modificar
    Toolbar1.Buttons(7).Enabled = (b And vUsu.NivelSumi <= 2)
    'eliminar
    Toolbar1.Buttons(8).Enabled = (b And vUsu.NivelSumi <= 2)
    
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
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = b
        Text1(i).BackColor = vbWhite
    Next i
    
    For i = 0 To ImgPpal.Count - 1
        ImgPpal(i).Enabled = Not b
    Next i
    Combo1.Enabled = Not b
    Combo2.Enabled = Not b
    
    PonerFoco chkVistaPrevia
End Sub

Private Function DatosOk() As Boolean
Dim RS As Adodb.Recordset
Dim b As Boolean
Dim i As Integer
Dim Datos As String
Dim Cad As String
    
    b = CompForm(Me)
    IndiceErroneo = 0
    If (b = True) And ((Modo = 3) Or (Modo = 4)) Then
        For i = 0 To Text1.Count - 1
             If InStr(1, Text1(i).Text, "'") > 0 Then
                MsgBox "No puede introducir el car�cter ' en ning�n campo de texto", vbExclamation
                IndiceErroneo = i
                DatosOk = False
                Exit Function
             End If
        Next i

        ' comprobamos la integridad de la bd
        ' existe la instalacion
        Datos = ""
        Datos = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Trim(Text1(0).Text) & "|" & Trim(Text1(2).Text) & "|", "T|T|", 2)
        If Datos = "" Then
            If MsgBox("No existe la Instalaci�n para la Empresa. Desea continuar.", vbQuestion + vbYesNo + vbDefaultButton2) = False Then
                DatosOk = False
                Exit Function
            End If
        End If
        ' exite el operario en la empresa introducida
        Datos = ""
        Datos = DevuelveDesdeBD(1, "dni", "operarios", "c_empresa|dni|", Text1(0).Text & "|" & Text1(9).Text & "|", "T|T|", 2)
        If Datos = "" Then
            If MsgBox("No existe el operario en la empresa introducida. Desea continuar.", vbQuestion + vbYesNo + vbDefaultButton1) = False Then
                DatosOk = False
                Exit Function
            End If
        End If
        If Text1(6).Text <> "" And Text1(10).Text <> "" Then
            If CDate(Text1(6).Text) > CDate(Text1(10).Text) Then
                MsgBox "La Fecha de Retirada no puede ser inferior a la de Asignaci�n", vbExclamation
                DatosOk = False
                Exit Function
            End If
        End If
    End If

If (b = True) And (Modo = 3) Then
    'Estamos insertando
    'a�o es com posar: select codvarie from svarie where codvarie = txtAux(0)
    'la N es pa dir que es numeric
     Datos = DevuelveDesdeBD(1, "n_reg_dosimetro", "dosimorganos", "n_reg_dosimetro|", Text1(7).Text & "|", "N|", 1)
     If Datos <> "" Then
        MsgBox "Ya existe el n�mero de registro de dosimetro : " & Text1(7).Text, vbExclamation
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
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 5 'Listado de dosimetros de organos
            FrmListado.Show
        
        Case Else
    
    End Select
End Sub

Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = Bol
    Next i
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
Dim i As Integer
Dim SQL As String

        SQL = " WHERE n_reg_dosimetro=" & Data1.Recordset!n_reg_dosimetro
        
        Conn.Execute "Delete  from dosimorganos " & SQL
       
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
    Dim SQL As String
    On Error GoTo ESituarData1
        'Actualizamos el recordset
        Data1.Refresh
        '********* canviar la clau primaria codsocio per la que siga *********
        'El sql para que se situe en el registro en especial es el siguiente
        SQL = "n_reg_dosimetro = " & Text1(7).Text & ""
        '*****************************************************************
        Data1.Recordset.Find SQL
        If Data1.Recordset.EOF Then GoTo ESituarData1
        SituarData1 = True
    Exit Function
ESituarData1:
    If Err.Number <> 0 Then Err.Clear
    Limpiar Me
    PonerModo 0
    SituarData1 = False
End Function

