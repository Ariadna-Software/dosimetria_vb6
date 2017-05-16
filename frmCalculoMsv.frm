VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCalculoMsv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Temporal nC"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8610
   Icon            =   "frmCalculoMsv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   6
      Left            =   7680
      TabIndex        =   6
      Tag             =   "Cristal 4|N|S|0|999.999|tempnc|cristal_4|||"
      Text            =   "Dato4"
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   5
      Left            =   6390
      TabIndex        =   5
      Tag             =   "Cristal 3|N|S|0|999.999|tempnc|cristal_3|||"
      Text            =   "Dato4"
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   4
      Left            =   5130
      TabIndex        =   4
      Tag             =   "Cristal 2|N|S|0|999.999|tempnc|cristal_2|||"
      Text            =   "Dato4"
      Top             =   5370
      Width           =   1215
   End
   Begin VB.CommandButton CmdFec 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   15
      Top             =   5340
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Tag             =   "Cristal 1|N|S|0|999.999|tempnc|cristal_1|||"
      Text            =   "Dato4"
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   2
      Left            =   2490
      TabIndex        =   2
      Tag             =   "N.Dosimetro|T|N|||tempnc|n_dosimetro|||"
      Text            =   "Dato3"
      Top             =   5340
      Width           =   1275
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8610
      _ExtentX        =   15187
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
            Object.ToolTipText     =   "Cálculo a mSv"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Personal"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Area"
               EndProperty
            EndProperty
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
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   5805
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7410
      TabIndex        =   8
      Top             =   5805
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Tag             =   "Hora Lectura|H|S|||tempnc|hora_lectura|hh:mm:ss||"
      Text            =   "Dato2"
      Top             =   5340
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Tag             =   "Fecha Lectura|F|N|||tempnc|fecha_lectura|dd/mm/yyyy||"
      Text            =   "Dat"
      Top             =   5340
      Width           =   945
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7425
      TabIndex        =   11
      Top             =   5805
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   10
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
            Picture         =   "frmCalculoMsv.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":1A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":22F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":2BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":38F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":3A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":3B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCalculoMsv.frx":42A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCalculoMsv.frx":43BA
      Height          =   5025
      Left            =   60
      TabIndex        =   12
      Top             =   480
      Width           =   8400
      _ExtentX        =   14817
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
End
Attribute VB_Name = "frmCalculoMsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Sistema As String
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim fich As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
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

Modo = vModo

b = (Modo = 0)

For I = 0 To txtAux.Count - 1
    txtAux(I).Visible = Not b
Next I
For I = 0 To CmdFec.Count - 1
    CmdFec(I).Visible = Not b
Next I
If Modo = 2 Then CmdFec(0).Visible = False
Toolbar1.Buttons(1).Enabled = b
Toolbar1.Buttons(2).Enabled = b
Toolbar1.Buttons(6).Enabled = b
Toolbar1.Buttons(7).Enabled = b
Toolbar1.Buttons(8).Enabled = b
CmdAceptar.Visible = Not b
CmdCancelar.Visible = Not b
DataGrid1.Enabled = b

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = b
End If
'Si estamo mod or insert
If Modo = 2 Then
   txtAux(2).BackColor = &H80000018
Else
   txtAux(2).BackColor = &H80000005
End If
txtAux(2).Enabled = (Modo <> 2)

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

    CadenaConsulta = "Select fecha_lectura, hora_lectura, n_dosimetro, cristal_1, cristal_2, "
    CadenaConsulta = CadenaConsulta & "cristal_3, cristal_4 from tempnc where codusu = " & vUsu.codigo
    If Sistema <> "" Then CadenaConsulta = CadenaConsulta & " and sistema = '" & Sistema & "'"
    CargaGrid ("tempnc.fecha_lectura = '9999-99-99'")
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
    Dim cad As String
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
    cad = ""
    For I = 0 To 3
        cad = cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux(3).Text = DataGrid1.Columns(3).Text
    
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
    CmdFec(0).Top = alto
End Sub

Private Sub BotonEliminar()
Dim sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    sql = "Seguro que desea eliminar la lectura de este registro:"
    sql = sql & vbCrLf & "Fecha Lectura: " & Format(adodc1.Recordset.Fields(0).Value, "dd/mm/yyyy")
    sql = sql & vbCrLf & "N.Dosimetro: " & adodc1.Recordset.Fields(2).Value
    If MsgBox(sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        sql = "Delete from tempnc where fecha_lectura='" & Format(adodc1.Recordset!fecha_lectura, FormatoFecha) & "' and "
        sql = sql & " n_dosimetro = '" & Trim(adodc1.Recordset!n_dosimetro) & "' and sistema = '"
        sql = Trim(adodc1.Recordset!Sistema)
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
Select Case Modo
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

Private Sub cmdCancelar_Click()
    Select Case Modo
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

Private Sub cmdFec_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    Select Case Index
       Case 0 'fecha de lectura
            f = Now
            If txtAux(0).Text <> "" Then
                If IsDate(txtAux(0).Text) Then f = txtAux(0).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            txtAux(0).Text = frmC.fecha
                mTag.DarFormato txtAux(0)
            Set frmC = Nothing
    End Select
End Sub

Private Sub cmdRegresar_Click()
    Dim cad As String
    
    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation
        Exit Sub
    End If
    
    cad = adodc1.Recordset.Fields(0) & "|"
    cad = cad & adodc1.Recordset.Fields(2) & "|"
    cad = cad & adodc1.Recordset.Fields(3) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
           Case vbESC, vbSalir
                If Modo = 0 Then
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
    If Modo = 0 Then lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
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
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    'Cadena consulta
    PonerOpcionesMenuGeneral Me
    
    CadenaConsulta = "Select fecha_lectura, hora_lectura, n_dosimetro, cristal_1, cristal_2, "
    CadenaConsulta = CadenaConsulta & "cristal_3, cristal_4 from tempnc where codusu = " & vUsu.codigo
    If Sistema <> "" Then CadenaConsulta = CadenaConsulta & " and sistema = '" & Sistema & "'"
    
    CargaGrid

End Sub

Private Sub Form_Unload(Cancel As Integer)
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
Private Function SugerirCodigoSiguiente(Tipo As Byte) As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    If Tipo = 0 Then
        sql = "Select Max(n_registro) from dosiscuerpo "
    Else
        sql = "Select Max(n_registro) from dosisarea "
    End If
    
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
           ' Calculo a Msv
'            Screen.MousePointer = vbHourglass
'            CalculoMsv
'            Screen.MousePointer = vbDefault
    Case 11
            'volvemos a mostrar el notepad del fichero que teniamos generado
            Shell "notepad " & Directorio & IIf(Right(Directorio, 1) <> "\", "\", "") & "Informe" & IIf(Sistema = "H", "6600.txt", "Pana.txt")
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
    sql = sql & " ORDER BY fecha_lectura"
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
      
    
    'fecha de lectura
    I = 0
        DataGrid1.Columns(I).Caption = "F.Lectura"
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'hora de lectura
    I = 1
        DataGrid1.Columns(I).Caption = "Hora "
        DataGrid1.Columns(I).Width = 700
        DataGrid1.Columns(I).NumberFormat = "hh:mm:ss"
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'numero de dosimetro
    I = 2
        DataGrid1.Columns(I).Caption = "N.Dosimetro"
        DataGrid1.Columns(I).Width = 1500
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width

    'cristal 1
    I = 3
        DataGrid1.Columns(I).Caption = "Cristal 1"
        DataGrid1.Columns(I).Width = 1000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
        'DataGrid1.Columns(I).Visible = Sistema <> "P"
        
    'cristal 2
    I = 4
        DataGrid1.Columns(I).Caption = "Cristal 2" ' IIf(Sistema <> "P", "Cristal 2", "Dosis Sup.")
        DataGrid1.Columns(I).Width = 1000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'cristal 3
    I = 5
        DataGrid1.Columns(I).Caption = "Cristal 3" ' IIf(Sistema <> "P", "Cristal 3", "Dosis Prof.")
        DataGrid1.Columns(I).Width = 1000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'cristal 4
    I = 6
        DataGrid1.Columns(I).Caption = "Cristal 4"
        DataGrid1.Columns(I).Width = 1000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
        'DataGrid1.Columns(I).Visible = Sistema <> "P"
        
    'Fijamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 55
        txtAux(1).Width = DataGrid1.Columns(1).Width - 55
        txtAux(2).Width = DataGrid1.Columns(2).Width - 55
        txtAux(3).Width = DataGrid1.Columns(3).Width - 55
        txtAux(4).Width = DataGrid1.Columns(4).Width - 55
        txtAux(5).Width = DataGrid1.Columns(5).Width - 55
        txtAux(6).Width = DataGrid1.Columns(6).Width - 55
        
        txtAux(0).Left = DataGrid1.Left + 340
        CmdFec(0).Left = txtAux(0).Left + txtAux(0).Width - CmdFec(0).Width
        txtAux(1).Left = CmdFec(0).Left + CmdFec(0).Width + 55
        txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 55
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 55
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 55
        txtAux(5).Left = txtAux(4).Left + txtAux(4).Width + 55
        txtAux(6).Left = txtAux(5).Left + txtAux(5).Width + 55
        
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

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim infor As Integer
    infor = IIf(Sistema = "H", 29, 37)
    If ButtonMenu = "Personal" Then
        If MsgBox("¿Está seguro de querer realizar la Migracion de Datos de Personal a mSv?", vbQuestion + vbYesNo + vbDefaultButton1, "¡Atención!") = vbYes Then
            Screen.MousePointer = vbHourglass
            CalculoMsv 0, Sistema 'personal
            Screen.MousePointer = vbDefault
            frmImprimir.OtrosParametros = "usu= " & vUsu.codigo & "|" & "tipo= 0|"
            frmImprimir.Opcion = infor
            frmImprimir.Show vbModal
        End If
    Else
        If MsgBox("¿Está seguro de querer realizar la Migracion de Datos de Area a mSv?", vbQuestion + vbYesNo + vbDefaultButton1, "¡Atención!") = vbYes Then
            Screen.MousePointer = vbHourglass
            CalculoMsv 2, Sistema 'area
            Screen.MousePointer = vbDefault
            frmImprimir.OtrosParametros = "usu= " & vUsu.codigo & "|" & "tipo= 2|"
            frmImprimir.Opcion = infor + 1
            frmImprimir.Show vbModal
        End If
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
    
    If Modo = 3 And ConCaracteresBusqueda(txtAux(Index).Text) Then Exit Sub 'Busquedas
    
    Select Case Index
      Case 3, 4, 5, 6 ' numericos
            If EsNumerico(txtAux(Index).Text) Then
                If InStr(1, txtAux(Index).Text, ",") > 0 Then
                    valor = ImporteFormateado(txtAux(Index).Text)
                Else
                    valor = CCur(TransformaPuntosComas(txtAux(Index).Text))
                End If
                
                txtAux(Index).Text = Format(valor, "##0.000")
            End If
        
      Case 0 ' fechas
            If txtAux(Index).Text <> "" Then
              If Not EsFechaOK(txtAux(Index)) Then
                    MsgBox "Fecha incorrecta: " & txtAux(Index).Text, vbExclamation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                    Exit Sub
              End If
              txtAux(Index).Text = Format(txtAux(Index).Text, "dd/mm/yyyy")
            End If
      
      Case 1 'campo hora
            If txtAux(Index).Text <> "" Then
                txtAux(Index).Text = Format(txtAux(Index).Text, "hh:mm:ss")
            End If
      
      
    End Select
    
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
b = CompForm(Me)
If Not b Then Exit Function

If Modo = 1 Then
    'Estamos insertando
     If Sistema <> "" Then
       Datos = DevuelveDesdeBD(1, "fecha_lectura", "tempnc", "fecha_lectura|sistema|", txtAux(0).Text & "|" & Sistema & "|", "F|", 2)
     Else
       Datos = DevuelveDesdeBD(1, "fecha_lectura", "tempnc", "fecha_lectura|", txtAux(0).Text & "|", "F|", 1)
     End If
     If Datos <> "" Then
        MsgBox "Ya existe el factor de calibración para esa fecha de inicio. Reintroduzca.", vbExclamation
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


'Private Sub CalculoMsv(Tipo As Byte)
'Dim Rs As ADODB.Recordset
'Dim rL As ADODB.Recordset
'Dim rf As ADODB.Recordset
'
'' Factores de fondo 2 y 3 para extremidad y solapa.
'Dim Fond2_ext As String
'Dim Fond3_ext As String
'Dim Fond2_sol As String
'Dim Fond3_sol As String
'
'' Factores de calibración 1 y 2 para anillo y pulsera, y 2 y 3 para solapa.
'Dim FCal2_sol As String
'Dim FCal3_sol As String
'Dim FCal1_ani As String
'Dim FCal2_ani As String
'Dim FCal1_pul As String
'Dim FCal2_pul As String
'
'' Varios para la fórmula.
'Dim Tipo_dos As String
'Dim Fondo1 As String
'Dim Fondo2 As String
'Dim Calib1 As String
'Dim Calib2 As String
'Dim Fact_dos As Currency
'Dim Fact_lot As Currency
'
'Dim sql As String
'Dim sql2 As String
'Dim sql1 As String
'Dim mSv2 As Currency
'Dim mSv3 As Currency
'Dim ErrorLectura As Boolean
'Dim DosisElevada As Boolean
'
'Dim Observaciones As String
'Dim NF As Currency
'
'Dim f_dosis As Date
'Dim f_migracion As Date
'Dim ndosi As String
'Dim punt_error As String
'Dim dni_usuario As String
'Dim c_empresa As String
'Dim c_instalacion As String
'Dim c_tipo_trabajo As String
'Dim plantilla_contrata As String
'Dim n_reg_dosimetro As String
'Dim rama_generica As String
'Dim rama_especifica As String
'Dim dato As String
'
'    On Error GoTo eCalculoMsv
'
'    Conn.BeginTrans
'
'    ' borramos la tabla auxiliar del listado
'    Conn.Execute "delete from zlistadomigracion where codusu = " & vUsu.codigo
'
'
'    sql = "select fecha_lectura, hora_lectura, n_dosimetro, cristal_2, cristal_3 from tempnc "
'    sql = sql & " where codusu = " & vUsu.codigo & " and sistema = '" & Sistema & "'"
'    sql = sql & " order by n_dosimetro"
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If Not Rs.EOF Then
'
'      ' Cargamos factores de fondo de Solapa.
'      Fond2_sol = ""
'      Fond3_sol = ""
'      If Not CargarFondos(Fond2_sol, Fond3_sol, "S") Then
'        MsgBox "No existe un registro de fondo de Solapa con fecha fin vacia. Revise"
'        Conn.RollbackTrans
'        Exit Sub
'      End If
'
'      ' Cargamos factores de fondo de Extremidad.
'      Fond2_ext = ""
'      Fond3_ext = ""
'      If Not CargarFondos(Fond2_ext, Fond3_ext, "E") Then
'        MsgBox "No existe un registro de fondo de Extremidad con fecha fin vacia. Revise"
'        Conn.RollbackTrans
'        Exit Sub
'      End If
'
'      ' Cargamos factores de calibración de Solapa.
'      FCal2_sol = ""
'      FCal3_sol = ""
'      If Not CargarFactores(FCal2_sol, FCal3_sol, "S") Then
'        MsgBox "No existe un registro de factores 6600 de Solapa con fecha fin vacia. Revise."
'        Conn.RollbackTrans
'        Exit Sub
'      End If
'
'      ' Cargamos factores de calibración de Anillo.
'      FCal1_ani = ""
'      FCal2_ani = ""
'      If Not CargarFactores(FCal1_ani, FCal2_ani, "A") Then
'        MsgBox "No existe un registro de factores 6600 de Anillo con fecha fin vacia. Revise."
'        Conn.RollbackTrans
'        Exit Sub
'      End If
'
'      ' Cargamos factores de calibración de Pulsera.
'      FCal1_pul = ""
'      FCal2_pul = ""
'      If Not CargarFactores(FCal1_pul, FCal2_pul, "P") Then
'        MsgBox "No existe un registro de factores 6600 de Pulsera con fecha fin vacia. Revise."
'        Conn.RollbackTrans
'        Exit Sub
'      End If
'      Rs.MoveFirst
'    End If
'
'    While Not Rs.EOF
'      ErrorLectura = False
'      DosisElevada = False
'      Observaciones = ""
'      punt_error = ""
'
'      ' ### [DavidV] 03/04/2006: Depende del tipo pasado como argumento,
'      ' es personal(0) o área(2).
'      sql1 = "select c_empresa, c_instalacion, dni_usuario, c_tipo_trabajo, "
'      sql1 = sql1 & "plantilla_contrata, n_reg_dosimetro from dosimetros "
'      sql1 = sql1 & " where n_dosimetro = '" & Trim(Rs.Fields(2).Value) & "' and "
'      sql1 = sql1 & " (f_retirada is null or f_retirada = '0000-00-00') and tipo_dosimetro = "
'      sql1 = sql1 & Tipo & " and sistema = '" & Sistema & "'"
'
'      Set rL = New ADODB.Recordset
'
'      rL.Open sql1, Conn, adOpenDynamic, adLockOptimistic
'
'      If Not rL.EOF Then
'        rL.MoveFirst
'        ndosi = rL.Fields(5).Value
'      Else
'        ErrorLectura = True
'        Observaciones = "DOSIMETRO NO ENCONTRADO"
'        ndosi = "-1"
'      End If
'
'      ' Depende del tipo de dosímetro, se usan unos fondos y factores distintos.
'      If Tipo = 0 Then
'        ' Cuerpo o Área, son de Solapa.
'        Fondo1 = Fond2_sol
'        Fondo2 = Fond3_sol
'        Calib1 = FCal2_sol
'        Calib2 = FCal3_sol
'      Else
'        ' Órgano.
'        Fondo1 = Fond2_ext
'        Fondo2 = Fond3_ext
'        Calib1 = "0"
'        Calib2 = "0"
'        Tipo_dos = DevuelveDesdeBD(1, "tipo_medicion", "dosisnohomog", "n_dosimetro|n_reg_dosimetro|", Rs.Fields(2) & "|" & ndosi & "|", "T|N|", 2)
'        Select Case Tipo_dos
'          Case "01", "05"
'            ' Pulsera.
'            Calib1 = FCal1_pul
'            Calib2 = FCal2_pul
'            Tipo_dos = "P"
'          Case "06", "07"
'            ' Anillo.
'            Calib1 = FCal1_ani
'            Calib2 = FCal2_ani
'            Tipo_dos = "A"
'          Case "08"
'            ' Abdomen (este es un caso raro de Solapa).
'            Fondo1 = Fond2_sol
'            Fondo2 = Fond3_sol
'            Calib1 = FCal2_sol
'            Calib2 = FCal3_sol
'          Case Else
'
'        End Select
'      End If
'
'      ' Calculando el cristal A.
'      mSv2 = 0
'      If Rs.Fields(3).Value <> "" Then
'        dato = DevuelveDesdeBD(1, "cristal_a", "dosimetros", "n_dosimetro|tipo_dosimetro|n_reg_dosimetro|sistema|", Rs.Fields(2) & "|" & Tipo & "|" & ndosi & "|" & Sistema & "|", "T|N|N|T|", 4)
'        If dato <> "" Then
'          Fact_dos = CCur(dato)
'        Else
'          Fact_dos = 1
'        End If
'        dato = DevuelveDesdeBD(1, "cristal_a", "lotes", "dosimetro_inicial|dosimetro_final|tipo|sistema|", "<=" & Rs.Fields(2) & "|>=" & Rs.Fields(2) & "|" & Tipo_dos & "|" & Sistema & "|", "N|N|T|T|", 4)
'        If dato <> "" Then
'          Fact_lot = CCur(dato)
'        Else
'          Fact_lot = 1
'        End If
'        mSv2 = Round2((Rs.Fields(3).Value - CCur(Fondo1)) * CCur(Calib1) * Fact_dos * Fact_lot, 2)
'      End If
'      mSv3 = 0
'
'      ' Calculando el cristal B.
'      If Rs.Fields(4).Value <> "" Then
'        dato = DevuelveDesdeBD(1, "cristal_b", "dosimetros", "n_dosimetro|tipo_dosimetro|n_reg_dosimetro|sistema|", Rs.Fields(2) & "|" & Tipo & "|" & ndosi & "|" & Sistema & "|", "T|N|N|T|", 4)
'        If dato <> "" Then
'          Fact_dos = CCur(dato)
'        Else
'          Fact_dos = 1
'        End If
'        dato = DevuelveDesdeBD(1, "cristal_b", "lotes", "dosimetro_inicial|dosimetro_final|sistema|", "<=" & Rs.Fields(2) & "|>=" & Rs.Fields(2) & "|" & Tipo_dos & "|" & Sistema & "|", "N|N|T|T|", 4)
'        If dato <> "" Then
'          Fact_lot = CCur(dato)
'        Else
'          Fact_lot = 1
'        End If
'        mSv3 = Round2((Rs.Fields(4).Value - CCur(Fondo2)) * CCur(Calib2) * Fact_dos * Fact_lot, 2)
'      End If
'
'      If mSv2 < 0.1 Then mSv2 = 0
'      If mSv3 < 0.1 Then mSv3 = 0
'
'      ' datos de la rama especifica y generica
'      If Not ErrorLectura Then
'
'        sql1 = "select rama_gen, rama_especifica from instalaciones where c_instalacion = '"
'        sql1 = sql1 & Trim(rL.Fields(1).Value) & "'"
'
'        Set rf = New ADODB.Recordset
'
'        rf.Open sql1, Conn, adOpenDynamic, adLockOptimistic
'
'        If Not rf.EOF Then rf.MoveFirst
'
'      End If
'
'      NF = SugerirCodigoSiguiente(Tipo)
'
'      f_dosis = Rs.Fields(0).Value - 30
'      f_migracion = Now
'
'
'      If mSv2 > 4 Or mSv3 > 4 Then
'        Observaciones = "DOSIS ELEVADA"
'        DosisElevada = True
'      End If
'
'      If ErrorLectura Or DosisElevada Then
'        sql2 = "insert into erroresmigra (n_registro, descripcion, c_tipo) VALUES ("
'        sql2 = sql2 & ImporteSinFormato(CStr(NF)) & ",'" & Trim(Observaciones) & "'," & Format(Tipo, "0") & ")"
'
'        Conn.Execute sql2
'      End If
'
'      If ErrorLectura Then
'        punt_error = "**"
'        dni_usuario = "999999999"
'        c_empresa = "DESCON"
'        c_instalacion = "DESCON"
'        c_tipo_trabajo = "99"
'        plantilla_contrata = "00"
'        n_reg_dosimetro = ""
'        rama_generica = "99"
'        rama_especifica = "99"
'      Else
'        If DosisElevada Then punt_error = "**"
'        dni_usuario = rL.Fields(2).Value
'        c_empresa = rL.Fields(0).Value
'        c_instalacion = rL.Fields(1).Value
'        c_tipo_trabajo = rL.Fields(3).Value
'        plantilla_contrata = rL.Fields(4).Value
'        n_reg_dosimetro = ndosi
'        rama_generica = rf.Fields(0).Value
'        rama_especifica = rf.Fields(1).Value
'      End If
'
'      If Tipo = 0 Then
'        'personal (homogéneas)
'        sql2 = "dosiscuerpo"
'      ElseIf Tipo = 1 Then
'        'extremidades (no homogéneas)
'        sql2 = "dosisnohomog"
'      Else
'        'área
'        sql2 = "dosisarea"
'      End If
'
'      sql2 = "insert into " & sql2 & " (n_registro, n_dosimetro, c_empresa, c_instalacion, "
'      sql2 = sql2 & " dni_usuario, f_dosis, f_migracion, dosis_superf, dosis_profunda, "
'      sql2 = sql2 & " plantilla_contrata, rama_generica, rama_especifica, c_tipo_trabajo, "
'      sql2 = sql2 & " observaciones, migrado, n_reg_dosimetro) values ("
'      sql2 = sql2 & ImporteSinFormato(CStr(NF)) & ",'" & Trim(Rs.Fields(2).Value) & "','"  'n_dosimetro
'      sql2 = sql2 & Trim(c_empresa) & "','"          ' empresa
'      sql2 = sql2 & Trim(c_instalacion) & "','"          ' instalacion
'      sql2 = sql2 & Trim(dni_usuario) & "','"          ' dni de usuario
'      sql2 = sql2 & Format(f_dosis, FormatoFecha) & "','"     'fecha de dosis
'      sql2 = sql2 & Format(f_migracion, FormatoFecha) & "',"  ' fecha de migracion
'      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv3))) & ","   'dosis superficial
'      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv2))) & ",'"  'dosis profunda
'      sql2 = sql2 & Trim(plantilla_contrata) & "','"  'plantilla contrata
'      sql2 = sql2 & Trim(rama_generica) & "','"  'rama generica
'      sql2 = sql2 & Trim(rama_especifica) & "','"  'rama especifica
'      sql2 = sql2 & Trim(c_tipo_trabajo) & "','"  ' tipo de trabajo
'      sql2 = sql2 & Trim(Observaciones) & "',null,'"  'observaciones, migrado
'      sql2 = sql2 & Trim(n_reg_dosimetro) & "')"  'n_reg_dosimetro
'      Conn.Execute sql2
'
'      ' tenemos que insertar en la tabla temporal para poder imprimir
'      sql2 = "insert into zlistadomigracion (codusu, n_registro, n_dosimetro, dni_usuario, cristal2,"
'      sql2 = sql2 & "cristal3, f_migracion, punt_error) values (" & vUsu.codigo & ","
'      sql2 = sql2 & ImporteSinFormato(CStr(NF)) & ",'"  'numero de registro
'      sql2 = sql2 & Trim(Rs.Fields(2).Value) & "','" & Trim(dni_usuario) & "',"
'      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv2))) & ","  'dosis profunda
'      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv3))) & ",'"   'dosis superficial
'      sql2 = sql2 & Format(f_migracion, FormatoFecha) & "','"  ' fecha de migracion
'      sql2 = sql2 & Trim(punt_error) & "')"
'      Conn.Execute sql2
'
'      Set rL = Nothing
'      Set rf = Nothing
'      Rs.MoveNext
'
'    Wend
'
'    Set Rs = Nothing
'
'eCalculoMsv:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Error en el cálculo de mSvs"
'        Conn.RollbackTrans
'    Else
'        Conn.CommitTrans
'    End If
'
'End Sub
'
'Private Function CargarFondos(ByRef Fondo1 As String, ByRef Fondo2 As String, ByVal Tipo As String) As Boolean
'Dim rf As ADODB.Recordset
'Dim sql As String
'Dim tabla As String
'
'    CargarFondos = False
'    tabla = IIf(Sistema = "H", "fondos", "fondospana")
'    sql = "select fondo_2, fondo_3 from " & tabla & " where f_fin is null and tipo = '" & Tipo & "'"
'    Set rf = New ADODB.Recordset
'
'    rf.Open sql, Conn, adOpenDynamic, adLockOptimistic
'    If Not rf.EOF Then
'        rf.MoveFirst
'        Fondo1 = rf.Fields(0).Value
'        Fondo2 = rf.Fields(1).Value
'        CargarFondos = True
'    End If
'    rf.Close
'    Set rf = Nothing
'End Function
'
'Private Function CargarFactores(ByRef Factor1 As String, ByRef Factor2 As String, ByVal Tipo As String) As Boolean
'Dim rf As ADODB.Recordset
'Dim sql As String
'Dim tabla As String
'
'    CargarFactores = False
'    tabla = IIf(Sistema = "H", "factcali6600", "factcalipana")
'    sql = "select cristal_a, cristal_b, f_inicio from factcali6600 where f_fin is null and tipo = '" & Tipo & "'"
'    Set rf = New ADODB.Recordset
'
'    rf.Open sql, Conn, adOpenDynamic, adLockOptimistic
'    If Not rf.EOF Then
'        rf.MoveFirst
'        Factor1 = rf.Fields(0).Value
'        Factor2 = rf.Fields(1).Value
'        CargarFactores = True
'    End If
'End Function
'
