VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmErroresMigraArea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errores Migración Area"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10500
   Icon            =   "frmErroresMigraArea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ImgPpal 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   5340
      Width           =   255
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Tag             =   "Descripcion|T|S|||erroresmigra|descripcion|||"
      Text            =   "Dato2"
      Top             =   5340
      Width           =   2865
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Tag             =   "Número Registro|N|N|||erroresmigra|n_registro|||"
      Text            =   "Dat"
      Top             =   5340
      Width           =   945
   End
   Begin VB.Frame FrameCancelacion 
      Enabled         =   0   'False
      Height          =   3930
      Left            =   2010
      TabIndex        =   11
      Top             =   1050
      Visible         =   0   'False
      Width           =   6540
      Begin VB.TextBox txtReg 
         Height          =   285
         Index           =   1
         Left            =   4410
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1470
         Width           =   1335
      End
      Begin VB.TextBox txtReg 
         Height          =   285
         Index           =   0
         Left            =   2085
         MaxLength       =   15
         TabIndex        =   12
         Top             =   1470
         Width           =   1335
      End
      Begin VB.CommandButton CmdCanCancelacion 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5220
         TabIndex        =   19
         Top             =   3285
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarCancelacion 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4140
         TabIndex        =   17
         Top             =   3285
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   2100
         MaxLength       =   15
         TabIndex        =   15
         Top             =   2445
         Width           =   1020
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   510
         Left            =   360
         TabIndex        =   14
         Top             =   3180
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   8
         Left            =   1260
         TabIndex        =   22
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   9
         Left            =   3630
         TabIndex        =   21
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Registro"
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
         Left            =   405
         TabIndex        =   20
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   1860
         Picture         =   "frmErroresMigraArea.frx":030A
         ToolTipText     =   "Seleccionar fecha"
         Top             =   2445
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cancelación Carga Automática"
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
         TabIndex        =   18
         Top             =   405
         Width           =   4650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Migración"
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
         Left            =   420
         TabIndex        =   16
         Top             =   2160
         Width           =   1620
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmErroresMigraArea.frx":040C
      Height          =   5025
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   8864
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelación Carga Automática"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
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
         Left            =   4560
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9210
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   975
      Left            =   1440
      Top             =   3480
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1720
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":0421
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":0533
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":0645
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":0757
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":0869
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":097B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":1255
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":1B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":2409
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":2CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":35BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":3A0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":3B21
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":3C33
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":3D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmErroresMigraArea.frx":43BF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmErroresMigraArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim fich As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim modo As Byte
Dim RC As String

Private ValorAnterior As String
'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim b As Boolean
Dim I As Integer

modo = vModo

b = (modo = 0)

For I = 0 To txtAux.Count - 1
    txtAux(I).Visible = Not b
Next I
For I = 0 To Imgppal.Count - 1
    Imgppal(I).Visible = Not b
Next I
If modo = 2 Then Imgppal(0).Visible = False
Toolbar1.Buttons(1).Enabled = b
Toolbar1.Buttons(2).Enabled = b
Toolbar1.Buttons(6).Enabled = b
Toolbar1.Buttons(7).Enabled = b
Toolbar1.Buttons(8).Enabled = b
cmdAceptar.Visible = Not b
cmdCancelar.Visible = Not b
DataGrid1.Enabled = b

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = b
End If
'Si estamo mod or insert
If modo = 2 Then
   txtAux(0).BackColor = &H80000018
Else
   txtAux(0).BackColor = &H80000005
End If
txtAux(0).Enabled = (modo <> 2)

End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    Dim I As Integer
    
    'Obtenemos la siguiente numero de factura
    'NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
          
    If DataGrid1.Row < 0 Then
        anc = 770
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 495 '545
    End If
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    LLamaLineas anc, 0
    
    'Ponemos el foco
    PonerFoco txtAux(0)
    
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
Dim I As Integer

'    CadenaConsulta = "Select erroresmigra.n_registro, erroresmigra.descripcion, dosiscuerpo.dni_usuario, operarios.apellido_1,"
'    CadenaConsulta = CadenaConsulta & "operarios.apellido_2 , operarios.nombre "
'    CadenaConsulta = CadenaConsulta & "from erroresmigra, dosisarea, operarios "
'    CadenaConsulta = CadenaConsulta & "where erroresmigra.c_tipo = 1 and erroresmigra.n_registro = dosisarea.n_registro and "
'    CadenaConsulta = CadenaConsulta & "dosisarea.dni_usuario = operarios.dni "
    CadenaConsulta = "Select erroresmigra.n_registro, erroresmigra.descripcion, dosisarea.dni_usuario, operarios.apellido_1, "
    CadenaConsulta = CadenaConsulta & "operarios.apellido_2 , operarios.nombre "
    CadenaConsulta = CadenaConsulta & "from erroresmigra left join dosisarea on erroresmigra.n_registro = dosisarea.n_registro "
    CadenaConsulta = CadenaConsulta & "left join operarios on dosisarea.dni_usuario = operarios.dni "
    CadenaConsulta = CadenaConsulta & "Where erroresmigra.c_tipo = 1 "

    CargaGrid ("erroresmigra.n_registro = -1")
    Me.lblIndicador.Caption = "BUSQUEDA"
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    LLamaLineas DataGrid1.Top + 206, 2
    
    PonerFoco txtAux(0)

End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 495 '545
    End If
    Cad = ""
    For I = 0 To 1
        Cad = Cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text

'    i = adodc1.Recordset!tipoconce
'    Combo1.ListIndex = i - 1
    LLamaLineas anc, 1
   
   'Como es modificar
   PonerFoco txtAux(0)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer

    PonerModo xModo + 1
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        txtAux(I).Top = alto
    Next I
    Imgppal(0).Top = alto
End Sub

Private Sub BotonEliminar()
Dim sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    sql = "Seguro que desea eliminar este registro:"
    sql = sql & vbCrLf & "N.Registro: " & adodc1.Recordset.Fields(0).Value
    sql = sql & vbCrLf & "Decripcion: " & adodc1.Recordset.Fields(1).Value
    If MsgBox(sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        sql = "Delete from erroresmigra where n_registro = " & adodc1.Recordset.Fields(0).Value & " and "
        sql = sql & " descripcion = '" & Trim(adodc1.Recordset.Fields(1).Value) & "' and c_tipo = 1"
        Conn.Execute sql
        CargaGrid ""
    End If

Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lectura Dosímetro"
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
Select Case modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me, 1) Then
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        Else
            MsgBox vbCrLf & "  Debe introducir alguna condición de búsqueda. " & vbCrLf, vbExclamation
            PonerModo 0
        End If
    End Select
    PonerFoco DataGrid1
    
End Sub

Private Sub CmdAceptarCancelacion_Click()
    If txtReg(0).Text <> "" And txtReg(1).Text <> "" Then
    
        BorradoRegistros
        
        DesactivarFrameCancelacion
        
        PonerModo 0
    
    Else
        MsgBox "Debe introducir los valores de número de registro desde y hasta", vbExclamation
        
        PonerFoco txtReg(0)
    
    End If
End Sub

Private Sub CmdCanCancelacion_Click()
    DesactivarFrameCancelacion
End Sub

Private Sub cmdCancelar_Click()
    Select Case modo
    Case 1
        DataGrid1.AllowAddNew = False
        'CargaGrid
        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        
    Case 3
        CargaGrid
    End Select
    PonerModo 0
    If Not adodc1.Recordset.EOF Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    Else
        lblIndicador.Caption = ""
    End If
    PonerFoco DataGrid1
End Sub


Private Sub ImgFec_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.fecha = Now
    If Text3(Index).Text <> "" Then frmC.fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing

End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    Select Case Index
       Case 0 'nro registro de dosiscuerpo
    
    End Select
End Sub

Private Sub cmdRegresar_Click()
    Dim Cad As String
    
    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation
        Exit Sub
    End If
    
    Cad = adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & adodc1.Recordset.Fields(2) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
           Case vbESC, vbSalir
                If modo = 0 Then
                    Unload Me
                Else
                    cmdCancelar_Click
                End If
           Case vbBuscar
                 Toolbar1_ButtonClick Toolbar1.Buttons(1)
           Case vbVerTodos
                 Toolbar1_ButtonClick Toolbar1.Buttons(2)
           Case vbAñadir
                 Toolbar1_ButtonClick Toolbar1.Buttons(6)
           Case vbModificar
                 Toolbar1_ButtonClick Toolbar1.Buttons(7)
           Case vbEliminar
                 Toolbar1_ButtonClick Toolbar1.Buttons(8)
           Case vbImprimir
                 Toolbar1_ButtonClick Toolbar1.Buttons(11)
        End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If modo = 0 Then lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
      
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
        .Buttons(10).Image = 21
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
      Toolbar1.Buttons(10).Visible = False
    End If
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    'Cadena consulta
    CadenaConsulta = "Select erroresmigra.n_registro, erroresmigra.descripcion, dosisarea.dni_usuario, operarios.apellido_1, "
    CadenaConsulta = CadenaConsulta & "operarios.apellido_2 , operarios.nombre "
    CadenaConsulta = CadenaConsulta & "from erroresmigra left join dosisarea on erroresmigra.n_registro = dosisarea.n_registro "
    CadenaConsulta = CadenaConsulta & "left join operarios on dosisarea.dni_usuario = operarios.dni "
    CadenaConsulta = CadenaConsulta & "Where erroresmigra.c_tipo = 1 "

    
'    CadenaConsulta = "Select erroresmigra.n_registro, erroresmigra.descripcion, dosisarea.dni_usuario, operarios.apellido_1,"
'    CadenaConsulta = CadenaConsulta & "operarios.apellido_2 , operarios.nombre "
'    CadenaConsulta = CadenaConsulta & "from erroresmigra, dosisarea, operarios "
'    CadenaConsulta = CadenaConsulta & "where erroresmigra.c_tipo = 1 and erroresmigra.n_registro = dosisarea.n_registro and "
'    CadenaConsulta = CadenaConsulta & "dosisarea.dni_usuario = operarios.dni "

    
    CargaGrid

End Sub

Private Sub Form_Unload(cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub
'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "Select Max(n_registro) from dosisarea "
    
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
    Case 10
           ' Cancelacion carga automatica
            ActivarFrameCancelacion
    Case 11
            'volvemos a mostrar el notepad del fichero que teniamos generado
            frmImprimir.Opcion = 28
            frmImprimir.email = False
            frmImprimir.Show vbModal
            
            
    Case 12
            Unload Me
    Case Else
    
    End Select
End Sub

Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub

Private Sub CargaGrid(Optional sql As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    DataGrid1.Enabled = False
    adodc1.ConnectionString = Conn
    If sql <> "" Then
        sql = CadenaConsulta & " AND " & sql
        Else
        sql = CadenaConsulta
    End If
    sql = sql & " ORDER BY n_registro"
    adodc1.RecordSource = sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    DataGrid1.Enabled = True
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
      
    
    'n.Registro
    I = 0
        DataGrid1.Columns(I).Caption = "N.Registro"
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'Observaciones
    I = 1
        DataGrid1.Columns(I).Caption = "Descripcion Error "
        DataGrid1.Columns(I).Width = 2700
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'Dni operario
    I = 2
        DataGrid1.Columns(I).Caption = "DNI Operario "
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'Apellido 1
    I = 3
        DataGrid1.Columns(I).Caption = "Apellido 1 "
        DataGrid1.Columns(I).Width = 1500
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'Apellido 2
    I = 4
        DataGrid1.Columns(I).Caption = "Apellido 2 "
        DataGrid1.Columns(I).Width = 1500
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
        
    'Nombre
    I = 5
        DataGrid1.Columns(I).Caption = "Nombre "
        DataGrid1.Columns(I).Width = 1500
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
        
        
    'Fijamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 55
        txtAux(1).Width = DataGrid1.Columns(1).Width - 55
        
        txtAux(0).Left = DataGrid1.Left + 340
        Imgppal(0).Left = txtAux(0).Left + txtAux(0).Width - Imgppal(0).Width
        txtAux(1).Left = Imgppal(0).Left + Imgppal(0).Width + 55
        
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF

   If Not adodc1.Recordset.EOF Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        lblIndicador.Caption = ""
   End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    With txtAux(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim sql As String
Dim valor As Currency

    ''Quitamos blancos por los lados
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    If txtAux(Index).BackColor = vbYellow Then
        txtAux(Index).BackColor = vbWhite
    End If
    
    If txtAux(Index) = "" Then Exit Sub
    
    If ValorAnterior = txtAux(Index).Text Then Exit Sub
    
    If modo = 3 And ConCaracteresBusqueda(txtAux(Index).Text) Then Exit Sub 'Busquedas
    
    Select Case Index
      Case 0 ' numericos
            If EsNumerico(txtAux(Index).Text) Then
                If InStr(1, txtAux(Index).Text, ",") > 0 Then
                    valor = ImporteFormateado(txtAux(Index).Text)
                Else
                    valor = CCur(TransformaPuntosComas(txtAux(Index).Text))
                End If
                
            End If
      Case 1 ' tipo string
      
        
      
      
    End Select
    
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
b = CompForm(Me)
If Not b Then Exit Function

If modo = 1 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD(1, "n_registro", "erroresmigra", "n_registro|", txtAux(0).Text & "|", "N|", 1)
     If Datos <> "" Then
        MsgBox "Ya existe el numero de registro. Reintroduzca.", vbExclamation
        b = False
    End If
End If
DatosOk = b
End Function

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ActivarFrameCancelacion()
Dim C As Control

    txtReg(0).Text = ""
    txtReg(1).Text = ""
    Text3(0).Text = ""

    For Each C In Me.Controls
        
        If Not TypeOf C Is ImageList Then
        
         If C.Container.Name <> "FrameCancelacion" Then
             C.Enabled = False
         End If
        End If
    Next C
    FrameCancelacion.Enabled = True
    FrameCancelacion.Visible = True
End Sub

Private Sub DesactivarFrameCancelacion()
Dim C As Control

    For Each C In Me.Controls
        If Not TypeOf C Is ImageList Then
         
         If C.Container.Name <> "FrameCancelacion" Then
             C.Enabled = True
         End If
        End If
        
    Next C
    FrameCancelacion.Enabled = False
    FrameCancelacion.Visible = False
End Sub

Private Sub BorradoRegistros()
Dim sql As String

    On Error GoTo eBorradoRegistros
    
    Conn.BeginTrans
    
    sql = "delete from dosiscuerpo where n_registro >=" & txtReg(0).Text & " and "
    sql = sql & "n_registro <= " & txtReg(1).Text
    
    Conn.Execute sql
    
eBorradoRegistros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en el borrado de Registros de Dosis Cuerpo"
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
    End If

End Sub


Private Sub txtreg_GotFocus(Index As Integer)
    With txtReg(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtreg_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub txtReg_LostFocus(Index As Integer)
Dim sql As String
Dim valor As Currency

    ''Quitamos blancos por los lados
    txtReg(Index).Text = Trim(txtReg(Index).Text)
    If txtReg(Index).Text = "" Then Exit Sub
    If txtReg(Index).BackColor = vbYellow Then
        txtReg(Index).BackColor = vbWhite
    End If
    
    If txtReg(Index) = "" Then Exit Sub
    
    If ValorAnterior = txtReg(Index).Text Then Exit Sub
    
    If modo = 3 And ConCaracteresBusqueda(txtReg(Index).Text) Then Exit Sub 'Busquedas
    
    If EsNumerico(txtReg(Index).Text) Then
        If InStr(1, txtReg(Index).Text, ",") > 0 Then
            valor = ImporteFormateado(txtReg(Index).Text)
        Else
            valor = CCur(TransformaPuntosComas(txtReg(Index).Text))
        End If
        
        'miramos que el registro existe en dosiscuerpo y que no está migrado a CSN
        Dim cad1 As String
        cad1 = "migrado"
        
        sql = ""
        sql = DevuelveDesdeBD(1, "n_registro", "dosiscuerpo", "n_registro|", txtReg(Index).Text & "|", "N|", 1, cad1)
        If sql = "" Then
            MsgBox "No existe este número registro. Reintroduzca", vbExclamation
            txtReg(Index).Text = ""
            PonerFoco txtReg(Index)
        Else
           If cad1 = "**" Then
                MsgBox "Este número de registro está migrado al CSN. Reintroduzca."
                txtReg(Index).Text = ""
                PonerFoco txtReg(Index)
           End If
        End If
    End If
End Sub
