VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFondos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fondos Harsaw 6600"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   5640
   Icon            =   "frmFondos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   4
      Left            =   5085
      TabIndex        =   14
      Tag             =   "Tipo|T|N|||fondos|tipo||S|"
      Text            =   "Tipo"
      Top             =   5505
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   5790
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4455
      TabIndex        =   12
      Top             =   5790
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton CmdFec 
      Height          =   255
      Index           =   1
      Left            =   4395
      TabIndex        =   9
      Top             =   5505
      Width           =   255
   End
   Begin VB.CommandButton CmdFec 
      Height          =   255
      Index           =   0
      Left            =   3045
      TabIndex        =   8
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   3
      Left            =   3345
      TabIndex        =   3
      Tag             =   "Fecha Finalización|F|S|||fondos|f_fin|dd/mm/yyyy||"
      Text            =   "Dato4"
      Top             =   5505
      Width           =   1275
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   2
      Left            =   1950
      TabIndex        =   2
      Tag             =   "Fecha de Inicio|F|N|||fondos|f_inicio|dd/mm/yyyy|S|"
      Text            =   "Dato3"
      Top             =   5520
      Width           =   1320
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
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
            Object.ToolTipText     =   "Carga Automática Fondos"
            ImageIndex      =   11
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
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   765
      TabIndex        =   1
      Tag             =   "Fondo 3|N|N|0|999.999|fondos|fondo_3|||"
      Text            =   "Dato2"
      Top             =   5520
      Width           =   840
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Tag             =   "Fondo 2|N|N|0|999.999|fondos|fondo_2|||"
      Text            =   "Dat"
      Top             =   5520
      Width           =   630
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
      Left            =   1080
      Top             =   3600
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
            Picture         =   "frmFondos.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":1A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":22F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":2BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":38F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":3A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":3B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFondos.frx":42A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFondos.frx":43BA
      Height          =   4365
      Left            =   150
      TabIndex        =   10
      Top             =   945
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7699
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   45
      TabIndex        =   11
      Top             =   480
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   8784
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Solapa"
      TabPicture(0)   =   "frmFondos.frx":43CF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Extremidad"
      TabPicture(1)   =   "frmFondos.frx":43EB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
End
Attribute VB_Name = "frmFondos"
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
Dim b2 As Boolean
Dim I As Integer

Modo = vModo

b = (Modo = 0)
If Not adodc1.Recordset Is Nothing Then
  b2 = adodc1.Recordset.EOF
Else
  b2 = False
End If

SSTab1.Enabled = b
For I = 0 To txtAux.Count - 2
    txtAux(I).Visible = Not b
Next I
For I = 0 To CmdFec.Count - 1
    CmdFec(I).Visible = Not b And Not (Modo = 2)
Next I

If Modo = 2 Then CmdFec(0).Visible = False
Toolbar1.Buttons(1).Enabled = b
Toolbar1.Buttons(2).Enabled = b
Toolbar1.Buttons(6).Enabled = b
Toolbar1.Buttons(7).Enabled = b And Not b2
Toolbar1.Buttons(8).Enabled = b And Not b2
Toolbar1.Buttons(10).Enabled = b
cmdAceptar.Visible = Not b
cmdCancelar.Visible = Not b
DataGrid1.Enabled = b And Not b2

'Si estamo mod or insert
  If Modo = 2 Then
    txtAux(2).BackColor = &H80000018
    txtAux(2).Enabled = False
    If txtAux(3).Text <> "" Then
      txtAux(3).BackColor = &H80000018
      txtAux(3).Enabled = False
    Else
      txtAux(3).BackColor = &H80000005
      txtAux(3).Enabled = True
    End If
  Else
    txtAux(2).Enabled = True
    txtAux(3).Enabled = True
    txtAux(2).BackColor = &H80000005
    txtAux(3).BackColor = &H80000005
  End If

  If SSTab1.Caption = "Extremidad" Then ' xqx
    txtAux(1).BackColor = &H80000018
    txtAux(1).Enabled = False
  Else
    txtAux(1).BackColor = &H80000005
    txtAux(1).Enabled = True
  End If

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
        anc = 1160
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 950.2362
    End If
    For I = 0 To txtAux.Count - 2
        txtAux(I).Text = ""
    Next I
    LLamaLineas anc, 0
    
    'Ponemos el foco
    PonerFoco txtAux(0)
    
End Sub

Private Sub BotonVerTodos()
    CargaGrid "tipo ='" & txtAux(4).Text & "'"
End Sub

Private Sub BotonBuscar()
Dim I As Integer

    CadenaConsulta = "Select * from fondos where 1=1 "
    CargaGrid ("fondos.f_inicio = '9999-99-99' and tipo ='" & txtAux(4).Text & "'")
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + 950.2362
    End If
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux(3).Text = DataGrid1.Columns(3).Text
    
    LLamaLineas anc, 1
   
   'Como es modificar
   PonerFoco txtAux(0)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    PonerModo xModo + 1
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    CmdFec(0).Top = alto
    CmdFec(1).Top = alto
End Sub

Private Sub BotonEliminar()
Dim sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    sql = "Seguro que desea eliminar el factor de calibración:"
    sql = sql & vbCrLf & "Fecha Inicio: " & adodc1.Recordset.Fields(2)
    If MsgBox(sql, vbQuestion + vbYesNoCancel, "¡Atención!") = vbYes Then
        'Hay que eliminar
        sql = "Delete from fondos where f_inicio='" & Format(adodc1.Recordset!f_inicio, FormatoFecha) & "' and tipo ='" & txtAux(4).Text & "'"
        Conn.Execute sql
        CargaGrid "tipo ='" & txtAux(4).Text & "'"
    End If

Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Fondo"
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
                CargaGrid "tipo = '" & txtAux(4).Text & "'"
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
                    CargaGrid "tipo = '" & txtAux(4).Text & "'"
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB & " and tipo ='" & txtAux(4).Text & "'"
        Else
            MsgBox vbCrLf & "  Debe introducir alguna condición de búsqueda. " & vbCrLf, vbExclamation, "¡Error!"
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
        CargaGrid "tipo ='" & txtAux(4).Text & "'"
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
       Case 0 'fecha de inicio
            f = Now
            If txtAux(2).Text <> "" Then
                If IsDate(txtAux(2).Text) Then f = txtAux(2).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            txtAux(2).Text = frmC.fecha
                mTag.DarFormato txtAux(2)
            Set frmC = Nothing
       Case 1 ' fecha finalizacion
            f = Now
            If txtAux(3).Text <> "" Then
                If IsDate(txtAux(3).Text) Then f = txtAux(3).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            txtAux(3).Text = frmC.fecha
            mTag.DarFormato txtAux(3)
            Set frmC = Nothing
    End Select
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
    
    If Modo = 0 Then
      If adodc1.Recordset.AbsolutePosition <> -1 Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
      Else
        lblIndicador.Caption = ""
      End If
    End If

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
        .Buttons(10).Image = 22
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    'Cadena consulta
'    PonerOpcionesMenuGeneral Me

    CadenaConsulta = "Select * from fondos where 1=1"
    SSTab1_Click 0
    
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
      Toolbar1.Buttons(10).Visible = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmrge_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
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
    
    sql = "Select Max(codcomar) from scomar where codprovi = '" & txtAux(0).Text
    sql = sql & "'"
    
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
Dim Rs As ADODB.Recordset
Dim sql As String


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
        ' carga automatica de fondos
        sql = "SELECT f_inicio FROM fondos WHERE f_fin IS NULL AND tipo = '" & txtAux(4).Text & "'"
        Set Rs = New ADODB.Recordset
        Rs.Open sql, Conn, , , adCmdText
        If Not Rs.EOF Then
          If Format(Rs!f_inicio, "yyyy-MM-dd") = Format(Now, "yyyy-MM-dd") Then
            MsgBox "Ya se han calculado los Fondos hoy. Elimine los datos de hoy si quiere volver a calcularlos.", vbOKOnly + vbExclamation, "¡Atención!"
            Exit Sub
          End If
        End If
        
        FrmHarshaw6600.Show vbModal
        If MsgBox("Desea continuar con el proceso de cálculo", vbQuestion + vbYesNo + vbDefaultButton1, "Cálculo de Fondos") = vbYes Then
            CalculaFondo
            BotonVerTodos
        End If
    Case 11
        'Imprimimos el listado
        Screen.MousePointer = vbHourglass
        FrmListado.Opcion = 18 'Listado de fondos
        FrmListado.Show
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
    Dim I As Integer
    Dim cristal1 As Boolean
    
    DataGrid1.Enabled = False
    adodc1.ConnectionString = Conn
    If sql <> "" Then
        sql = CadenaConsulta & " AND " & sql
        Else
        sql = CadenaConsulta
    End If
    sql = sql & " ORDER BY f_inicio"
    adodc1.RecordSource = sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    DataGrid1.Enabled = True
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    ' ### [DavidV] Añadido el cristal 1 (para los dosímetros de anillo y muñeca).
    
    'cristales a y b
    cristal1 = (txtAux(4).Text <> "S") ' xqx
    DataGrid1.Columns(0).Caption = IIf(cristal1, "Cristal 1 =", "Cristal 2")
    DataGrid1.Columns(1).Caption = IIf(cristal1, "Cristal 2", "Cristal 3")
    DataGrid1.Columns(0).Width = 900
    DataGrid1.Columns(1).Width = 900
    txtAux(0).Tag = DataGrid1.Columns(0).Caption & Mid(txtAux(0).Tag, InStr(1, txtAux(0).Tag, "|"))
    txtAux(1).Tag = DataGrid1.Columns(1).Caption & Mid(txtAux(1).Tag, InStr(1, txtAux(1).Tag, "|"))

    'fecha de inicio
    DataGrid1.Columns(2).Caption = "F.Inicio"
    DataGrid1.Columns(2).Width = 1400

    'fecha de finalizacion
    DataGrid1.Columns(3).Caption = "F.Finalización"
    DataGrid1.Columns(3).Width = 1400

    ' tipo de calibración invisible
    DataGrid1.Columns(4).Visible = False

        'Fijamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        txtAux(2).Width = DataGrid1.Columns(2).Width - 60
        txtAux(3).Width = DataGrid1.Columns(3).Width - 60
        txtAux(0).Left = DataGrid1.Left + 350
        txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 55
        txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 55
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 60
        CmdFec(0).Left = txtAux(3).Left - 55 - CmdFec(0).Width
        CmdFec(1).Left = txtAux(3).Left + txtAux(3).Width - CmdFec(1).Width
        
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
   DataGrid1.Enabled = (Modo = 0) And Not adodc1.Recordset.EOF
   
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
    
    If Modo = 3 And ConCaracteresBusqueda(txtAux(Index).Text) Then Exit Sub 'Busquedas
    
    Select Case Index
      Case 0, 1 ' numericos
            If EsNumerico(txtAux(Index).Text) Then
                If InStr(1, txtAux(Index).Text, ",") > 0 Then
                    valor = ImporteFormateado(txtAux(Index).Text)
                Else
                    valor = CCur(TransformaPuntosComas(txtAux(Index).Text))
                End If
                
                If SSTab1.Caption = "Extremidad" Then txtAux(1).Text = txtAux(0).Text ' xqx
                txtAux(Index).Text = Format(valor, "##0.000")
            End If
        
      Case 2, 3 ' fechas
            If txtAux(Index).Text <> "" Then
              If Not EsFechaOK(txtAux(Index)) Then
                    MsgBox "Fecha incorrecta: " & txtAux(Index).Text, vbExclamation, "¡Error!"
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                    Exit Sub
              End If
              txtAux(Index).Text = Format(txtAux(Index).Text, "dd/mm/yyyy")
            End If
      
    End Select
    
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim sql As String
Dim error As Boolean
Dim Consulta As ADODB.Recordset
Dim feci As Date
Dim fecf As Date
Dim fecaux As Date

  error = Not CompForm(Me)
    
  fecaux = CDate(IIf(txtAux(3).Text <> "", txtAux(3).Text, "9999-12-31"))
    
  If CDate(txtAux(2).Text) > fecaux Then
    MsgBox "La fecha final no puede ser menor que la inicial. Reintroduzca.", vbExclamation, "¡Error!"
    DatosOk = False
    Exit Function
  End If
   
  If Not error And Modo = 1 Then
  
    sql = "select * from fondos where tipo = '" & txtAux(4).Text & "'"
      
    Set Consulta = New ADODB.Recordset
    Consulta.Open sql, Conn, , , adCmdText
    error = False
    While Not Consulta.EOF And Not error
      feci = Consulta!f_inicio
      fecf = IIf(Not IsNull(Consulta!f_fin), Consulta!f_fin, "9999-12-31")
      ' ¿Alguno de los nº de dosímetro se encuentran en un rango de lote existente?
      If CDate(txtAux(2).Text) >= feci And CDate(txtAux(2).Text) <= fecf Then error = True
      If fecaux >= feci And fecaux <= fecf Then error = True
      ' Hemos comprobado si nuestro rango estaba incluido en alguno existente, ahora comprobamos
      ' si existe alguno incluido dentro de nuestro rango.
      If feci > CDate(txtAux(2).Text) And feci < fecaux Then error = True
      If fecf > CDate(txtAux(2).Text) And fecf < fecaux Then error = True
     
      Consulta.MoveNext
    Wend
    Set Consulta = Nothing
    
    If error Then
      MsgBox "El rango de fechas ya existe o forma parte de uno existente. Reintroduzca.", vbExclamation, "¡Error!"
    End If
  End If
  DatosOk = Not error
  
End Function

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

' funcion que calcula el fondo
Private Sub CalculaFondo()
Dim sql As String
Dim Rs As ADODB.Recordset
Dim v_totfon As Integer
Dim v_cris2 As Currency
Dim v_cris3 As Currency
Dim v_med2 As Currency
Dim v_med3 As Currency

On Error GoTo eCalculaFondo

    Conn.BeginTrans

    sql = "select count(*) from tempnc, dosimetros where (dosimetros.tipo_dosimetro = 0 or dosimetros.tipo_dosimetro = 2) "
    sql = sql & " and dosimetros.f_retirada is null and dosimetros.n_dosimetro = tempnc.n_dosimetro"
    sql = sql & " and tempnc.sistema = 'H' and tempnc.codusu = " & vUsu.codigo
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    v_totfon = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            v_totfon = Rs.Fields(0)
        End If
    End If
    Rs.Close
    
    sql = "select sum(cristal_2), sum(cristal_3) from tempnc, dosimetros where "
    sql = sql & "(dosimetros.tipo_dosimetro = 0 or dosimetros.tipo_dosimetro = 2) and dosimetros.f_retirada is null and "
    sql = sql & "dosimetros.n_dosimetro = tempnc.n_dosimetro and tempnc.sistema = 'H' and tempnc.codusu = " & vUsu.codigo
    Rs.Open sql, Conn, , , adCmdText
    v_cris2 = 0
    v_cris3 = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            v_cris2 = Rs.Fields(0)
        End If
        If Not IsNull(Rs.Fields(1)) Then
            v_cris3 = Rs.Fields(1)
        End If
    End If
    Rs.Close
    
    ' sacamos la media de los fondos para almacenarlos
    v_med2 = 0
    v_med3 = 0
    If v_totfon <> 0 Then
        v_med2 = Round2(v_cris2 / v_totfon, 3)
        v_med3 = Round2(v_cris3 / v_totfon, 3)
    End If
    
    ' actualizamos el ultimo registro
    sql = "update fondos set f_fin = '" & Format(DateAdd("d", -1, Now), FormatoFecha) & "' "
    sql = sql & "where f_fin is null and tipo = '" & txtAux(4).Text & "'"
    
    Conn.Execute sql
    ' aquii
    ' insertamos el ultimo registro de fondo
    sql = "insert into fondos (fondo_2, fondo_3, f_inicio, f_fin, tipo) values ("
    sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(v_med2))) & ","
    sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(v_med3))) & ","
    sql = sql & "'" & Format(Now, FormatoFecha) & "',null,'" & txtAux(4).Text & "')"

    Conn.Execute sql
    
eCalculaFondo:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en el calculo de fondos. Revise registros"
        Conn.RollbackTrans
    Else
        MsgBox "Proceso de cálculo realizado correctamente.", , "Proceso de cálculo."
        Conn.CommitTrans
        
    End If
    
End Sub

Private Sub SSTab1_Click(Index As Integer)
  
  txtAux(4).Text = Left(SSTab1.Caption, 1)
  Toolbar1.Buttons(10).Enabled = txtAux(4).Text = "S"
  CargaGrid "tipo = '" & txtAux(4).Text & "' "
  
End Sub

