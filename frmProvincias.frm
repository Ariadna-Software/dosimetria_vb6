VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProvincias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Provincias"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6600
   Icon            =   "frmProvincias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   6600
   Begin VB.TextBox Txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   2
      Left            =   4020
      TabIndex        =   2
      Tag             =   "Prefijo|T|S|||provincias|prefijo|||"
      Text            =   "Dat"
      Top             =   4320
      Width           =   800
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
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
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4050
      TabIndex        =   3
      Top             =   5715
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5370
      TabIndex        =   4
      Top             =   5715
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||provincias|descripcion|||"
      Text            =   "Dato2"
      Top             =   4320
      Width           =   2475
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "Código|T|N|||provincias|c_postal||S|"
      Text            =   "Dat"
      Top             =   4320
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5310
      TabIndex        =   7
      Top             =   5730
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   5
      Top             =   5610
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   975
      Left            =   2520
      Top             =   5400
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
            Picture         =   "frmProvincias.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":0DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":0EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":1112
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":1224
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":1AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":23D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":2CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":358C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":3E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":42B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":43CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":44DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":45EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProvincias.frx":4C68
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmProvincias.frx":4D7A
      Height          =   4995
      Left            =   60
      TabIndex        =   8
      Top             =   450
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   8811
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
Attribute VB_Name = "frmProvincias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Private ValorAnterior As String


'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim b As Boolean
Modo = vModo

b = (Modo = 0)

txtAux(0).Visible = Not b
txtAux(1).Visible = Not b
txtAux(2).Visible = Not b

'mnOpciones.Enabled = b
Toolbar1.Buttons(1).Enabled = b
Toolbar1.Buttons(2).Enabled = b
Toolbar1.Buttons(6).Enabled = b 'And vUsu.NivelUsu <= 2
Toolbar1.Buttons(7).Enabled = b 'And vUsu.NivelUsu <= 2
Toolbar1.Buttons(8).Enabled = b 'And vUsu.NivelUsu <= 2
cmdAceptar.Visible = Not b
cmdCancelar.Visible = Not b
DataGrid1.Enabled = b

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = b
End If
'Si estamo mod or insert
If Modo = 2 Then
   txtAux(0).BackColor = &H80000018
   Else
    txtAux(0).BackColor = &H80000005
End If
txtAux(0).Enabled = (Modo <> 2)

End Sub


Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
   ' NumF = SugerirCodigoSiguiente
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 690
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 470  '500
        
    End If
    txtAux(0).Text = NumF
    txtAux(0).Text = Format(txtAux(0).Text, "00")
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    LLamaLineas anc, 0
    lblIndicador.Caption = "INSERTANDO"
    
    'Ponemos el foco
    PonerFoco txtAux(0)
    
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()

    CadenaConsulta = "Select c_postal, descripcion, prefijo "
    CadenaConsulta = CadenaConsulta & " from provincias where 1 = 1 "
    CargaGrid ("c_postal = 'xx'")
    Me.lblIndicador.Caption = "BUSQUEDA"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""

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
        anc = DataGrid1.RowTop(DataGrid1.Row) + 475 '495 '45
    End If
    Cad = ""
    For I = 0 To 2
        Cad = Cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    
    LLamaLineas anc, 1
   
   'Como es modificar
   PonerFoco txtAux(1)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    PonerModo xModo + 1
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
End Sub

Private Sub BotonEliminar()
Dim sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    sql = "Seguro que desea eliminar la Provincia:"
    sql = sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    sql = sql & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    If MsgBox(sql, vbQuestion + vbYesNoCancel + vbDefaultButton2, "¡Atención!") = vbYes Then  'VRS:1.0.1(11)
        'Hay que eliminar
        sql = "Delete from provincias where c_postal = '" & adodc1.Recordset!c_postal & "'"
        Conn.Execute sql
        CargaGrid ""
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Provincias"
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

Private Sub cmdRegresar_Click()
Dim Cad As String
    
    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation, "¡Atención!"
        Exit Sub
    End If
    
    Cad = adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & adodc1.Recordset.Fields(1) & "|"
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
    
    ' actualizamos los valores de forma de pago alternativa
'   If Not Adodc1.Recordset.EOF Then FormaPagoAlternativa

   If Modo = 0 Then lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
    'adodc1.Password = vUsuario.Passwd
    
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

    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 2 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
    End If

    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
'    PonerOpcionesMenu

    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
   
    CadenaConsulta = "Select c_postal, descripcion, prefijo "
    CadenaConsulta = CadenaConsulta & " from provincias "
    
    CargaGrid

End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
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
Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "Select Max(codforpa) from sforpa"
    
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
    Case 11
        'Imprimimos el listado
        Screen.MousePointer = vbHourglass
        FrmListado.Opcion = 13
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
    Dim TotalAncho As Integer
    Dim I As Integer
    
    DataGrid1.Enabled = False
    
    adodc1.ConnectionString = Conn
    If sql <> "" Then
        sql = CadenaConsulta & " AND " & sql
        Else
        sql = CadenaConsulta
    End If
    sql = sql & " ORDER BY c_postal"
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
    
    I = 0
        DataGrid1.Columns(I).Caption = "Código"
        DataGrid1.Columns(I).Width = 700
'        DataGrid1.Columns(i).NumberFormat = "00"
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Nombre"
        DataGrid1.Columns(I).Width = 3000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    I = 2
        DataGrid1.Columns(I).Caption = "Prefijo Telefónico"
        DataGrid1.Columns(I).Width = 2000
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
        'añadido

        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        txtAux(2).Width = DataGrid1.Columns(2).Width - 60
        
        txtAux(0).Left = DataGrid1.Left + 340
        txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
        txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 55
        
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF 'And vUsu.NivelUsu <= 2
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF 'And vUsu.NivelUsu <= 2

    
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
    ValorAnterior = txtAux(Index).Text
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    ''Quitamos blancos por los lados
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    'If Txtaux(Index).Text = "" Then Exit Sub
    If txtAux(Index).BackColor = vbYellow Then
        txtAux(Index).BackColor = vbWhite
    End If
    
    If txtAux(Index).Text = "" Then Exit Sub
    
    'If ValorAnterior = txtAux(Index).Text Then Exit Sub
    
    If Modo = 3 And ConCaracteresBusqueda(txtAux(Index).Text) Then Exit Sub

    Select Case Index
    Case 0
        If InStr(1, txtAux(Index).Text, "'") > 0 Then
             MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
             txtAux(Index).Text = Replace(Format(txtAux(Index).Text, ">"), "'", "", , , vbTextCompare)
             PonerFoco txtAux(Index)
             Exit Sub
        End If
    End Select
    txtAux(Index).Text = Format(txtAux(Index).Text, ">")
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
b = CompForm(Me)
If Not b Then Exit Function

If Modo = 1 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD(1, "c_postal", "provincias", "c_postal|", txtAux(0).Text & "|", "T|", 1)
     If Datos <> "" Then
        MsgBox "Ya existe la Provincia : " & txtAux(0).Text, vbExclamation, "¡Error!"
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

