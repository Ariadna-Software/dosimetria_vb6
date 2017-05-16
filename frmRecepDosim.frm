VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRecepDosim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Dosímetros"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7980
   Icon            =   "frmRecepDosim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   7980
   Begin VB.Frame FrameRecepcion 
      Caption         =   "Número de Dosímetro a Recepcionar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2235
      Left            =   2220
      TabIndex        =   38
      Top             =   1590
      Visible         =   0   'False
      Width           =   3795
      Begin VB.Timer TimerRec 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   60
         Top             =   210
      End
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   780
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton CmdCan 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1350
         TabIndex        =   41
         Top             =   1530
         Width           =   1110
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0E0FF&
         Height          =   345
         Left            =   1050
         TabIndex        =   39
         Text            =   "Text3"
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label LabelRec 
         Caption         =   "Dosímetro recepcionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   570
         TabIndex        =   42
         Top             =   1110
         Visible         =   0   'False
         Width           =   2595
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1665
      Left            =   330
      TabIndex        =   16
      Top             =   1950
      Width           =   7425
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "DNI|T|N|||recepdosim|dni_usuario||S|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1260
         Width           =   5325
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   900
         Width           =   5325
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   540
         Width           =   5325
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   3
         Left            =   1470
         MouseIcon       =   "frmRecepDosim.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmRecepDosim.frx":0E1C
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
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   2910
         TabIndex        =   23
         Top             =   945
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   1290
         Width           =   1290
      End
      Begin VB.Label Label9 
         Caption         =   "Primer Apellido:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label17 
         Caption         =   "Segundo Apellido:"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   960
         Width           =   1365
      End
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1020
      TabIndex        =   1
      Tag             =   "Tipo Dosimetro|N|N|||recepdosim|tipo_dosimetro||S|"
      Text            =   "Combo3"
      Top             =   510
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   4
      Left            =   6450
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "N.Dosimetro|T|N|||recepdosim|n_dosimetro||S|"
      Text            =   "Text1"
      Top             =   540
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   7
      Left            =   3870
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "N.Registro|N|N|||recepdosim|n_reg_dosimetro||S|"
      Text            =   "Text1"
      Top             =   540
      Width           =   1305
   End
   Begin VB.Frame Frame7 
      Height          =   1065
      Left            =   330
      TabIndex        =   24
      Top             =   870
      Width           =   7425
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3150
         MaxLength       =   30
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   240
         Width           =   4005
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Codigo Empresa|T|N|||recepdosim|c_empresa|||"
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
         TabIndex        =   29
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
         TabIndex        =   5
         Tag             =   "Codigo Instalación|T|N|||recepdosim|c_instalacion|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   1305
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   2
         Left            =   1470
         MouseIcon       =   "frmRecepDosim.frx":0F1E
         MousePointer    =   99  'Custom
         Picture         =   "frmRecepDosim.frx":1070
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
         TabIndex        =   34
         Top             =   240
         Width           =   915
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   4
         Left            =   1470
         MouseIcon       =   "frmRecepDosim.frx":1172
         MousePointer    =   99  'Custom
         Picture         =   "frmRecepDosim.frx":12C4
         ToolTipText     =   "Buscar instalación"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Instalación"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   180
         TabIndex        =   30
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   330
      TabIndex        =   25
      Top             =   3630
      Width           =   7425
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5730
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   5745
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Mes Par/Impar|T|N|||recepdosim|mes_p_i||S|"
         Text            =   "Text1"
         Top             =   285
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   2340
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Fecha Creacion Recepción|F|N|||recepdosim|f_creacion_recep|dd/mm/yyyy|S|"
         Text            =   "Text1"
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   510
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Fecha Recepcion|F|S|||recepdosim|fecha_recepcion|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   0
         Left            =   510
         MouseIcon       =   "frmRecepDosim.frx":13C6
         MousePointer    =   99  'Custom
         Picture         =   "frmRecepDosim.frx":1518
         ToolTipText     =   "Seleccionar fecha"
         Top             =   330
         Width           =   240
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   1
         Left            =   2340
         MouseIcon       =   "frmRecepDosim.frx":15A3
         MousePointer    =   99  'Custom
         Picture         =   "frmRecepDosim.frx":16F5
         ToolTipText     =   "Seleccionar fecha"
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Mes Par/Impar"
         Height          =   255
         Left            =   4380
         TabIndex        =   31
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Creación"
         Height          =   255
         Left            =   2685
         TabIndex        =   28
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Recepción"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   330
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6660
      TabIndex        =   12
      Top             =   4860
      Width           =   1110
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   4860
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   330
      TabIndex        =   13
      Top             =   4770
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   180
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   4860
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   780
      Top             =   4890
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
      TabIndex        =   15
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
            Object.ToolTipText     =   "Recepción Lápiz Óptico"
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
      Left            =   330
      TabIndex        =   37
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Dosímetro"
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
      Left            =   5370
      TabIndex        =   36
      Top             =   540
      Width           =   1215
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
      Left            =   2970
      TabIndex        =   35
      Top             =   570
      Width           =   975
   End
End
Attribute VB_Name = "frmRecepDosim"
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
Dim fechactrl As String

' campo que indica si la familia es fitosanitaria
' si lo es: obligamos a introducir los campos de fitos.
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

' ### [DavidV] 07/04/2006: Para la ordenación según recepción.
' La fecha de recepción no nos vale, puesto que no es real. Ellos
' cambian la del sistema según les conviene.
Private Function SiguienteOrdenRecep() As Long
Dim sql As String
Dim Consulta As ADODB.Recordset
                
  sql = "select max(orden_recepcion) from dosimetros"
 
  Set Consulta = New ADODB.Recordset
  Consulta.Open sql, Conn, , , adCmdText
  SiguienteOrdenRecep = (Consulta.Fields(0) + 1)
  Set Consulta = Nothing
  
End Function

Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
    AsignarTeclasFuncion KeyCode
End Sub

Private Sub CmdAcep_Click()
Dim numV As Integer
Dim sql As String
Dim max As Long
Dim Consulta As ADODB.Recordset
On Error GoTo ECmdAcep_Click

    If Text3.Text <> "" Then
        numV = 0
        numV = VecesDosimetroNoRecepcionado(Trim(Text3.Text))
        Select Case numV
            Case 0
                MsgBox "Este Dosímetro no existe o ya está recepcionado. Reintroduzca.", vbExclamation, "¡Error!"
                PonerFoco Text3
            Case 1
                
                ' ### [DavidV] 07/04/2006: Cambiada la fórmula para que actualice
                ' el campo orden_recepcion.
                ' Actualizar la fecha de recepción del dosímetro en la tabla de
                ' recepción.
                sql = "update recepdosim set fecha_recepcion = '" & Format(Now, FormatoFecha)
                sql = sql & "' where n_dosimetro = '" & Trim(Text3.Text) & "' and fecha_recepcion is null "
                If ActualizarOrdenRecep Then
                  Conn.Execute sql
                  ' ### DavidV 27/03/2006 (mensaje de recepción correcta).
                  LabelRec.ForeColor = &HC000&
                  LabelRec.Caption = "Dosímetro recepcionado"
                  LabelRec.Visible = True
                  TimerRec.Enabled = True
                  
                Else
                  ' ### DavidV 11/04/2006 (mensaje de recepción incorrecta).
                  LabelRec.ForeColor = &H40C0&
                  LabelRec.Caption = "Error recepcionando"
                  LabelRec.Visible = True
                  TimerRec.Enabled = True
                End If
                
                Text3.Text = ""
                PonerFoco Text3
                
            Case Is > 1
                MsgBox " Existe este Dosimetro esta varias veces no recepcionado. Recepcionelo manualmente", vbExclamation, "¡Atención!"
                CmdCan_Click
        
        End Select
    Else
        MsgBox "Debe introducir el Número de Dosímetro a Recepcionar", vbExclamation, "¡Error!"
        PonerFoco Text3
    End If
    Exit Sub

ECmdAcep_Click:

  ' Liberamos memoria y propagamos el error.
  If Not Consulta Is Nothing Then
    Set Consulta = Nothing
  End If
  If Not Conn Is Nothing Then
    Conn.RollbackTrans
    Set Conn = Nothing
  End If
  Err.Raise Err.Number, Err.Source, Err.Description

End Sub

' ### [DavidV] 11/04/2006: Actualiza en orden de recepción del dosímetro en la
' tabla de dosímetros (para el tema de la ordenación por recepción en listados).
Private Function ActualizarOrdenRecep(Optional manual As Boolean) As Boolean
Dim sql As String
Dim max As Long
Dim Consulta As ADODB.Recordset
On Error GoTo EActualizarOrdenRecep

  ' Siguiente orden de recepción.
  max = SiguienteOrdenRecep
  
  ' Buscamos los datos del dosímetro correspondiente a la recepción que vamos a
  ' actualizar.
  If manual Then
    
    ' Fórmula de actualización del orden de recepción.
    sql = "update dosimetros set orden_recepcion = " & max & " where n_dosimetro = '"
    sql = sql & Trim(Text1(4).Text) & "' and tipo_dosimetro = " & IIf(Combo3.ListIndex = 0, 0, 2)
    sql = sql & " and n_reg_dosimetro = " & Trim(Text1(7).Text) & " and dni_usuario = '"
    sql = sql & Trim(Text1(3).Text) & "' and mes_p_i = '" & Trim(Text1(5).Text) & "' "
    sql = sql & "and c_empresa = '" & Trim(Text1(0).Text) & "' and c_instalacion = '"
    sql = sql & Trim(Text1(2).Text) & "'"
  
  Else
    ' Primero buscamos los datos del dosímetro correspondiente a la recepción que
    ' vamos a actualizar.
    sql = "select * from recepdosim where n_dosimetro = '" & Trim(Text3.Text)
    sql = sql & "' and fecha_recepcion is null"
    Set Consulta = New ADODB.Recordset
    Consulta.Open sql, Conn, , , adCmdText
        
    ' Fórmula de actualización del orden de recepción.
    sql = "update dosimetros set orden_recepcion = " & max & " where n_dosimetro = '"
    sql = sql & Trim(Text3.Text) & "' and tipo_dosimetro = " & Consulta!tipo_dosimetro
    sql = sql & " and n_reg_dosimetro = " & Consulta!n_reg_dosimetro & " and dni_usuario = '"
    sql = sql & Consulta!dni_usuario & "' and mes_p_i = '" & Consulta!mes_p_i & "' "
    sql = sql & "and c_empresa = '" & Consulta!c_empresa & "' and c_instalacion = '"
    sql = sql & Consulta!c_instalacion & "'"
 
  End If
  
  Conn.Execute sql
                
                
  ' Cerramos la consulta.
  Set Consulta = Nothing
  ActualizarOrdenRecep = True
  Exit Function
  
EActualizarOrdenRecep:
  
  Set Consulta = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Function


Private Sub cmdAceptar_Click()
Dim max As Long
Dim sql As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
              ' ### [DavidV] 11/04/2006: Para que al insertar también contemple el orden de
              ' recepción, en caso de introducirlo.
              If Text1(6).Text <> "" Then ActualizarOrdenRecep True
              PonerModo 0
            End If
        End If
    
    Case 4
        'Modificar
        If DatosOk Then
            
            '-----------------------------------------
            ' Hacemos modificar
            If ModificaDesdeFormulario(Me, 1) Then
              ' ### [DavidV] 11/04/2006: Añadido para que al recepcionar manualmente se
              ' actualice el campo orden_recepcion de dosímetros.
              If fechactrl <> Text1(6).Text Then ActualizarOrdenRecep True
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

Private Sub CmdCan_Click()
    TimerRec_Timer
    DesactivarFrameRecepcion
    FrameRecepcion.Visible = False
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
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(4)
End Sub

Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "Select Max(n_reg_dosimetro) from recepdosim"
    
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
    PonerFoco Text1(6)
    
    ' ### [DavidV] 07/04/2006: Control de cambio de la fecha de recepción.
    fechactrl = Text1(6).Text
    
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
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Dosímetro"
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


Private Sub Combo2_LostFocus()
    If Combo2.ListIndex = 0 Then Text1(5).Text = "P"
    If Combo2.ListIndex = 1 Then Text1(5).Text = "I"
    Text1_LostFocus 3
End Sub

Private Sub Combo3_LostFocus()
  Text1_LostFocus 4
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
        .Buttons(11).Image = 24
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
      Toolbar1.Buttons(11).Visible = False
    End If
    
    '***** canviar el nom de la taula i el ORDER BY ********
    NombreTabla = "recepdosim"
    Ordenacion = " ORDER BY f_creacion_recep ASC, n_reg_dosimetro ASC"
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
    
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    
    'Aqui va el especifico de cada form es
    '### a mano
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
        Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
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
            If Text1(1).Text <> "" Then
                If IsDate(Text1(1).Text) Then f = Text1(1).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(1).Text = frmC.fecha
                mTag.DarFormato Text1(1)
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
        Case 3 ' operarios
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
    Dim ape1 As String
    Dim ape2 As String
    Dim nombre As String
    Dim resp As String
    Dim Tipo As Integer
    
    ''Quitamos blancos por los lados
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" Then Exit Sub
    
    If Modo = 1 And ConCaracteresBusqueda(Text1(Index).Text) Then Exit Sub
    
    Select Case Index
        Case 0, 2, 3, 4, 7
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            
            Select Case Index
                Case 0 'empresa
                    Text2(1).Text = ""
                    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
                    If Text2(1).Text = "" Then
                        MsgBox "El código de empresa no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                Case 2 ' instalacion
                    If Text1(Index).Text <> "" And Text1(0).Text <> "" Then
                        Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|f_alta|", Text1(0).Text & "|" & Text1(Index).Text & "|", "T|T|", 2)
                        If Text2(2).Text = "" Then
                            MsgBox "El código de instalacion no existe. Reintroduzca.", vbExclamation, "¡Error!"
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                Case 3 ' dni de operario
                    If Modo = 1 Then Exit Sub
                    
                    CargarDatosOperarios Text1(3).Text, ape1, ape2, nombre
                    Text2(3).Text = ape1
                    Text2(4).Text = ape2
                    Text2(5).Text = nombre
                                  
               Case 4, 7
                    ' ### [DavidV] 23/08/2006: Quitar los ceros a la izquierda de la lectura del lapiz.
                    If Text1(4).Text <> "" Then Text1(4).Text = Trim(Str(Val(Text1(4).Text)))
                    ' ### [DavidV] 07/04/2006: Para que aparezcan los datos del dosímetro
                    ' automáticamente al dar de alta una recepción.
                    If Modo = 3 And Combo3.Text <> "" And Text1(7).Text <> "" And Text1(4).Text <> "" Then
                      Tipo = IIf(Combo3.ListIndex = 0, 0, 2)
                      resp = DevuelveDesdeBD(1, "c_empresa", "dosimetros", "n_reg_dosimetro|n_dosimetro|tipo_dosimetro|", Text1(7).Text & "|" & Text1(4).Text & "|" & Tipo & "|", "N|N|N|", 3)
                      If resp <> "" Then
                        Text1(0).Text = resp
                      Else
                        Text1(0).Text = ""
                        Text2(1).Text = ""
                      End If
                      Text1_LostFocus 0
                      
                      resp = DevuelveDesdeBD(1, "c_instalacion", "dosimetros", "n_reg_dosimetro|n_dosimetro|tipo_dosimetro|", Text1(7).Text & "|" & Text1(4).Text & "|" & Tipo & "|", "N|N|N|", 3)
                      If resp <> "" Then
                        Text1(2).Text = resp
                      Else
                        Text1(2).Text = ""
                        Text2(2).Text = ""
                      End If
                      Text1_LostFocus 2
                      
                      resp = DevuelveDesdeBD(1, "dni_usuario", "dosimetros", "n_reg_dosimetro|n_dosimetro|tipo_dosimetro|", Text1(7).Text & "|" & Text1(4).Text & "|" & Tipo & "|", "N|N|N|", 3)
                      If resp <> "" Then
                        Text1(3).Text = resp
                      Else
                        Text1(3).Text = ""
                        Text2(3).Text = ""
                        Text2(4).Text = ""
                        Text2(5).Text = ""
                      End If
                      Text1_LostFocus 3
                      
                      resp = DevuelveDesdeBD(1, "mes_p_i", "dosimetros", "n_reg_dosimetro|n_dosimetro|tipo_dosimetro|", Text1(7).Text & "|" & Text1(4).Text & "|" & Tipo & "|", "N|N|N|", 3)
                      If resp <> "" Then
                        Combo2.ListIndex = IIf(resp = "P", 0, 1)
                      Else
                        Combo2.ListIndex = -1
                      End If
                      Combo2_LostFocus
                    End If
            End Select
        Case 6, 1
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

Combo2.Enabled = False

CadB = ObtenerBusqueda(Me)
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
        Cad = Cad & ParaGrid(Text1(7), 12, "N.Registro")
        Cad = Cad & ParaGrid(Text1(4), 12, "Dosimetro")
        Cad = Cad & ParaGrid(Text1(0), 15, "Empresa")
        Cad = Cad & ParaGrid(Text1(2), 16, "Instalacion")
        Cad = Cad & ParaGrid(Text1(3), 15, "DNI Operario")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|3|4|"
            frmB.vTitulo = "Dosímetros a Cuerpo"
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
        Combo2.ListIndex = -1
        If Text1(5).Text = "P" Then Combo2.ListIndex = 0
        If Text1(5).Text = "I" Then Combo2.ListIndex = 1
    End If
    
    Text2(1).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(0).Text & "|", "T|", 1)
    Text2(2).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Text1(0).Text & "|" & Text1(2).Text & "|", "T|T|", 2)
    
    CargarDatosOperarios Text1(3).Text, ape1, ape2, nombre
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
    Toolbar1.Buttons(6).Enabled = b 'And vUsu.NivelUsu <= 2
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
'    Toolbar1.Buttons(11).Enabled = b
    
    'Modificar
    Toolbar1.Buttons(7).Enabled = b 'And vUsu.NivelUsu <= 2)
    'eliminar
    Toolbar1.Buttons(8).Enabled = b 'And vUsu.NivelUsu <= 2)
    
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
    b = (Modo = 2) Or Modo = 0 Or (Modo = 4)
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = b
        Text1(I).BackColor = vbWhite
    Next I
    
    For I = 0 To Imgppal.Count - 1
        Imgppal(I).Enabled = Not b
    Next I
    
    ' ### [DavidV] 07/04/2006: Faltaba habilitar la imagen para el calendario de
    ' fecha de recepción.
    ' 23/04/2006: Por petición expresa de Javi, se dejan sin bloquear TAMBIÉN los
    ' siguientes campos.
    If Modo = 4 Then
        Text1(0).Locked = False
        Text1(2).Locked = False
        Text1(3).Locked = False
        Text1(4).Locked = False
        Text1(6).Locked = False
        Imgppal(0).Enabled = True
    End If
    
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
        ' exite el operario en la instalacion  introducida
        Datos = ""
        Datos = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", Text1(0).Text & "|" & Text1(2).Text & " |" & Text1(3).Text & "|", "T|T|T|", 3)
        If Datos = "" Then
            If MsgBox("No existe el operario en la instalación introducida. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, "¡Atención!") = vbNo Then
                DatosOk = False
                Exit Function
            End If
        End If
        If Text1(6).Text <> "" And Text1(1).Text <> "" Then
            If CDate(Text1(6).Text) < CDate(Text1(1).Text) Then
                MsgBox "La Fecha de Recepción no puede ser inferior a la de Creación", vbExclamation, "¡Error!"
                DatosOk = False
                Exit Function
            End If
        End If
        Datos = ""
        ' ### 30/03/2006 DavidV: Cambio de la búsqueda (en este caso buscamos por n_dosimetro).
        'Datos = DevuelveDesdeBD(1, "n_reg_dosimetro", "dosimetros", "n_reg_dosimetro|tipo_dosimetro|", Text1(7).Text & "|" & Combo3.ListIndex & "|", "N|N|", 2)
        Datos = DevuelveDesdeBD(1, "n_reg_dosimetro", "dosimetros", "n_reg_dosimetro|n_dosimetro|", Text1(7).Text & "|" & Text1(4).Text & "|", "N|N|", 2)
        If Datos = "" Then
            MsgBox "Este dosimetro no existe en el mantenimiento de dosimetros. Insertelo previamente.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
    End If

'If (b = True) And (Modo = 3) Then
'    'Estamos insertando
'     Datos = DevuelveDesdeBD(1, "n_reg_dosimetro", "recepdosim", "n_reg_dosimetro|", Text1(7).Text & "|", "N|", 1)
'     If Datos <> "" Then
'        MsgBox "Ya existe el número de registro de dosimetro : " & Text1(7).Text, vbExclamation
'        DatosOk = False
'        IndiceErroneo = 7
'        Exit Function
'    End If
'End If

DatosOk = b
End Function

Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
   Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If

End Sub

Private Sub Text3_LostFocus()
    
    If Text3.Text = "" Then Exit Sub
    
    If InStr(1, Text3.Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
        Text3.Text = Replace(Format(Text3.Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco Text3
        Exit Sub
    End If
    Text3.Text = Format(Text3.Text, ">")
    
    Text3.Text = CStr(CInt(Text3.Text))
    CmdAcep_Click

End Sub

' ### DavidV 27/03/2006 (desactivar mensaje de recepción correcta)
Private Sub TimerRec_Timer()
  LabelRec.Visible = False
  TimerRec.Enabled = False
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
            FrameRecepcion.Visible = True
            Text3.Text = ""
            ActivarFrameRecepcion
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
    
'    Combo3.AddItem "Organo"
'    Combo3.ItemData(Combo3.NewIndex) = 1
    
    Combo3.AddItem "Area"
    Combo3.ItemData(Combo3.NewIndex) = 2
    
'    Combo3.AddItem "Fondo"
'    Combo3.ItemData(Combo3.NewIndex) = 3
'
'    Combo3.AddItem "Tránsito"
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

        sql = " WHERE n_reg_dosimetro=" & Data1.Recordset!n_reg_dosimetro & " and tipo_dosimetro = " & IIf(Combo3.ListIndex = 0, 0, 2)
        sql = sql & " and f_creacion_recep = '" & Format(Text1(1).Text, FormatoFecha) & "'"
        
        Conn.Execute "Delete  from recepdosim " & sql
       
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
        sql = "n_reg_dosimetro = " & Text1(7).Text & ""
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

Function VecesDosimetroNoRecepcionado(dosimetro As String) As Integer
Dim sql As String
Dim Rs As ADODB.Recordset

    VecesDosimetroNoRecepcionado = 0
    sql = "select count(*) from recepdosim where n_dosimetro = '" & Trim(dosimetro) & "'"
    sql = sql & " and fecha_recepcion is null"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , adLockOptimistic
    
    If Not Rs.EOF Then
        VecesDosimetroNoRecepcionado = Rs.Fields(0).Value
    End If
    
End Function


Private Sub ActivarFrameRecepcion()
Dim C As Control

    For Each C In Me.Controls
      ' Para que no intente activar el TimerRec.
      If C.Name <> "TimerRec" Then
        If C.Container.Name <> "FrameRecepcion" Then
          C.Enabled = False
        End If
      End If
    Next C
    FrameRecepcion.Enabled = True
End Sub

Private Sub DesactivarFrameRecepcion()
Dim C As Control

    For Each C In Me.Controls
      ' Para que no intente desactivar el TimerRec
      If C.Name <> "TimerRec" Then
        If C.Container.Name <> "FrameRecepcion" Then
             C.Enabled = True
        End If
      End If
    Next C
    FrameRecepcion.Enabled = False
End Sub
