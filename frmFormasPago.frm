VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFormasPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de Pago"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7890
   Icon            =   "frmFormasPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   7890
   Begin VB.Frame Frame2 
      Caption         =   "Banco Propio Asociado"
      Height          =   855
      Left            =   60
      TabIndex        =   21
      Top             =   4740
      Width           =   7575
      Begin VB.TextBox Txtaux 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2115
         TabIndex        =   22
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox Txtaux 
         Height          =   285
         Index           =   9
         Left            =   1530
         TabIndex        =   6
         Tag             =   "Banco Propio|N|N|0|99|sforpa|codbanpr|00||"
         Top             =   360
         Width           =   495
      End
      Begin VB.Image ImgBPr 
         Height          =   240
         Left            =   1215
         MouseIcon       =   "frmFormasPago.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmFormasPago.frx":0E1C
         ToolTipText     =   "Buscar socio"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Código:"
         Height          =   255
         Left            =   405
         TabIndex        =   23
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.TextBox Txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   4
      Left            =   6480
      TabIndex        =   5
      Tag             =   "Resto.Vto|N|N|0|999|sforpa|restoven|||"
      Text            =   "Dat"
      Top             =   5535
      Width           =   800
   End
   Begin VB.TextBox Txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   3
      Left            =   5550
      TabIndex        =   4
      Tag             =   "Primer.Vto|N|N|0|999|sforpa|primerve|||"
      Text            =   "Dat"
      Top             =   5535
      Width           =   800
   End
   Begin VB.TextBox Txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   2
      Left            =   4680
      TabIndex        =   3
      Tag             =   "Nro.Vtos|N|N|0|9|sforpa|numerove|0||"
      Text            =   "Dat"
      Top             =   5535
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmFormasPago.frx":0F1E
      Left            =   3420
      List            =   "frmFormasPago.frx":0F20
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "T.Pago|N|N|||sforpa|tipopago|||"
      Top             =   5520
      Width           =   1200
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
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
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   5715
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6420
      TabIndex        =   10
      Top             =   5715
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   900
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "Denominación|T|N|||sforpa|nomforpa|||"
      Text            =   "Dato2"
      Top             =   5520
      Width           =   2475
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Tag             =   "Código|N|N|0|99|sforpa|codforpa|00|S|"
      Text            =   "Dat"
      Top             =   5520
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6390
      TabIndex        =   13
      Top             =   5730
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   11
      Top             =   5610
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   12
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
            Picture         =   "frmFormasPago.frx":0F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":1034
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":1146
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":1258
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":147C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":1D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":2630
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":2F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":37E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":40BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":4510
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":4622
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":4734
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":4846
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormasPago.frx":4EC0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFormasPago.frx":4FD2
      Height          =   4095
      Left            =   60
      TabIndex        =   14
      Top             =   450
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   7223
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
   Begin VB.Frame Frame3 
      Caption         =   "Forma de Pago Alternativa"
      Enabled         =   0   'False
      Height          =   855
      Left            =   60
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   7860
      Begin VB.TextBox Txtaux 
         Height          =   285
         Index           =   7
         Left            =   4140
         TabIndex        =   8
         Tag             =   "FP.Alter.|N|S|0|9|sforpa|forpaalt|0||"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Txtaux 
         Height          =   285
         Index           =   6
         Left            =   1350
         TabIndex        =   7
         Tag             =   "Importe Min.|N|S|||sforpa|impormin|###,##0.00||"
         Top             =   360
         Width           =   1305
      End
      Begin VB.TextBox Txtaux 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   4725
         TabIndex        =   18
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Importe mínimo:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   390
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "Código FP:"
         Height          =   255
         Left            =   3015
         TabIndex        =   19
         Top             =   360
         Width           =   795
      End
      Begin VB.Image ImgCtacom 
         Height          =   240
         Left            =   3825
         MouseIcon       =   "frmFormasPago.frx":4FE7
         MousePointer    =   99  'Custom
         Picture         =   "frmFormasPago.frx":5139
         ToolTipText     =   "Buscar socio"
         Top             =   360
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFormasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmBPr As frmBancosPropios
Attribute frmBPr.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim modo As Byte
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
modo = vModo

b = (modo = 0)

Txtaux(0).Visible = Not b
Txtaux(1).Visible = Not b
Txtaux(2).Visible = Not b
Txtaux(3).Visible = Not b
Txtaux(4).Visible = Not b
Txtaux(6).Enabled = Not b
Txtaux(7).Enabled = Not b
Txtaux(9).Enabled = Not b
ImgBPr.Enabled = Not b
Combo1.Visible = Not b

'mnOpciones.Enabled = b
Toolbar1.Buttons(1).Enabled = b
Toolbar1.Buttons(2).Enabled = b
Toolbar1.Buttons(6).Enabled = b And vUsu.NivelSumi <= 2
Toolbar1.Buttons(7).Enabled = b And vUsu.NivelSumi <= 2
Toolbar1.Buttons(8).Enabled = b And vUsu.NivelSumi <= 2
cmdAceptar.Visible = Not b
cmdCancelar.Visible = Not b
DataGrid1.Enabled = b

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = b
End If
'Si estamo mod or insert
If modo = 2 Then
   Txtaux(0).BackColor = &H80000018
   Else
    Txtaux(0).BackColor = &H80000005
End If
Txtaux(0).Enabled = (modo <> 2)

End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
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
    Txtaux(0).Text = NumF
    Txtaux(0).Text = Format(Txtaux(0).Text, "00")
    Txtaux(1).Text = ""
    Txtaux(2).Text = ""
    Txtaux(3).Text = ""
    Txtaux(4).Text = ""
    Txtaux(6).Text = ""
    Txtaux(7).Text = 0
    Txtaux(8).Text = ""
    
    Combo1.ListIndex = -1
    LLamaLineas anc, 0
    lblIndicador.Caption = "INSERTANDO"
    
    'Ponemos el foco
    PonerFoco Txtaux(0)
    
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()

    CadenaConsulta = "Select sforpa.codforpa, sforpa.nomforpa, tipoforp.descforp, sforpa.numerove, "
    CadenaConsulta = CadenaConsulta & "sforpa.primerve, sforpa.restoven, sforpa.impormin, "
    CadenaConsulta = CadenaConsulta & "sforpa.forpaalt, sforpa.codbanpr from sforpa,tipoforp "
    CadenaConsulta = CadenaConsulta & "where sforpa.tipopago = tipoforp.tipoforp"
    CargaGrid ("codforpa = 99")
    Me.lblIndicador.Caption = "BUSQUEDA"
    'Buscar
    Txtaux(0).Text = ""
    Txtaux(1).Text = ""
    Combo1.ListIndex = -1
    Txtaux(2).Text = ""
    Txtaux(3).Text = ""
    Txtaux(4).Text = ""
    Txtaux(6).Text = ""
    Txtaux(7).Text = ""
    Txtaux(8).Text = ""

    LLamaLineas DataGrid1.Top + 206, 2
    PonerFoco Txtaux(0)
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim i As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 475 '495 '45
    End If
    Cad = ""
    For i = 0 To 2
        Cad = Cad & DataGrid1.Columns(i).Text & "|"
    Next i
    'Llamamos al form
    Txtaux(0).Text = DataGrid1.Columns(0).Text
    Txtaux(1).Text = DataGrid1.Columns(1).Text
    Combo1.Text = DataGrid1.Columns(2).Text
    Txtaux(2).Text = DataGrid1.Columns(3).Text
    Txtaux(3).Text = DataGrid1.Columns(4).Text
    Txtaux(4).Text = DataGrid1.Columns(5).Text
    Txtaux(6).Text = DataGrid1.Columns(6).Text
    Txtaux(7).Text = DataGrid1.Columns(7).Text
    
    LLamaLineas anc, 1
   
   'Como es modificar
   PonerFoco Txtaux(1)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    PonerModo xModo + 1
    'Fijamos el ancho
    Txtaux(0).Top = alto
    Txtaux(1).Top = alto
    Combo1.Top = alto
    Txtaux(2).Top = alto
    Txtaux(3).Top = alto
    Txtaux(4).Top = alto
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    SQL = "Seguro que desea eliminar la forma de Pago:"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then  'VRS:1.0.1(11)
        'Hay que eliminar
        SQL = "Delete from sforpa where codforpa = " & adodc1.Recordset!codforpa
        Conn.Execute SQL
        CargaGrid ""
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Formas de Pago"
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
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
                i = adodc1.Recordset.Fields(0)
                PonerModo 0
                CargaGrid
                adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
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
'        FormaPagoAlternativa
    Else
        lblIndicador.Caption = ""
    End If
    PonerFoco DataGrid1
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
    
    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation
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
    
    ' actualizamos los valores de forma de pago alternativa
'   If Not Adodc1.Recordset.EOF Then FormaPagoAlternativa
    ' actualizamos los valores de banco propio
   If Not adodc1.Recordset.EOF Then BancoPropio

   If modo = 0 Then lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
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

   
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    PonerOpcionesMenu

    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
   
    CadenaConsulta = "Select sforpa.codforpa, sforpa.nomforpa, tipoforp.descforp, sforpa.numerove, "
    CadenaConsulta = CadenaConsulta & "sforpa.primerve, sforpa.restoven, sforpa.impormin,"
    CadenaConsulta = CadenaConsulta & " sforpa.forpaalt, sforpa.codbanpr from sforpa,tipoforp "
    CadenaConsulta = CadenaConsulta & "where sforpa.tipopago = tipoforp.tipoforp"
    
    CargarComboTipoPagos Combo1
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
    Dim SQL As String
    Dim RS As Adodb.Recordset
    
    SQL = "Select Max(codforpa) from sforpa"
    
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

Private Sub frmBPr_DatoSeleccionado(CadenaSeleccion As String)
    Txtaux(9).Text = RecuperaValor(CadenaSeleccion, 1)
    Txtaux(9).Text = Format(Txtaux(9).Text, "00")
    Txtaux(11).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ImgBpr_Click()
    Set frmBPr = New frmBancosPropios
    frmBPr.DatosADevolverBusqueda = "0|1|"
    frmBPr.Show
End Sub

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
        FrmListado.Opcion = 3 'Listado de formas de pago
        FrmListado.Show
    Case 12
        Unload Me
    Case Else

    End Select
End Sub


Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = Bol
    Next i
End Sub

Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codforpa"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
    Next i
    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "Cod."
        DataGrid1.Columns(i).Width = 700
        DataGrid1.Columns(i).NumberFormat = "00"
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Denominación"
        DataGrid1.Columns(i).Width = 3000
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
    
    i = 2
        DataGrid1.Columns(i).Caption = "Tipo"
        DataGrid1.Columns(i).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
        'añadido

    i = 3
        DataGrid1.Columns(i).Caption = "Vtos."
        DataGrid1.Columns(i).Width = 600
        DataGrid1.Columns(i).Alignment = dbgRight
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
        'añadido
        
    i = 4
        DataGrid1.Columns(i).Caption = "Primero"
        DataGrid1.Columns(i).Width = 800
        DataGrid1.Columns(i).Alignment = dbgRight
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
        'añadido
    
    i = 5
        DataGrid1.Columns(i).Caption = "Resto"
        DataGrid1.Columns(i).Width = 800
        DataGrid1.Columns(i).Alignment = dbgRight
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
        'añadido
        DataGrid1.Columns(6).Visible = False
        DataGrid1.Columns(7).Visible = False
        DataGrid1.Columns(8).Visible = False
        
'        ' importe minimo y forma de pago alternativa
'        If Not Adodc1.Recordset.EOF Then FormaPagoAlternativa
        If Not adodc1.Recordset.EOF Then BancoPropio
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        Txtaux(0).Width = DataGrid1.Columns(0).Width - 60
        Txtaux(1).Width = DataGrid1.Columns(1).Width - 60
        Combo1.Width = DataGrid1.Columns(2).Width - 40
        Txtaux(2).Width = DataGrid1.Columns(3).Width - 40
        Txtaux(3).Width = DataGrid1.Columns(4).Width - 40
        Txtaux(4).Width = DataGrid1.Columns(5).Width - 40
        
        Txtaux(0).Left = DataGrid1.Left + 340
        Txtaux(1).Left = Txtaux(0).Left + Txtaux(0).Width + 45
        Combo1.Left = Txtaux(1).Left + Txtaux(1).Width + 55
        Txtaux(2).Left = Combo1.Left + Combo1.Width + 55
        Txtaux(3).Left = Txtaux(2).Left + Txtaux(2).Width + 55
        Txtaux(4).Left = Txtaux(3).Left + Txtaux(3).Width + 55
        
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF And vUsu.NivelSumi <= 2
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF And vUsu.NivelSumi <= 2

    
    If Not adodc1.Recordset.EOF Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
'        FormaPagoAlternativa
        BancoPropio
    Else
        lblIndicador.Caption = ""
    End If
   
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    With Txtaux(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    ValorAnterior = Txtaux(Index).Text
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    ''Quitamos blancos por los lados
    Txtaux(Index).Text = Trim(Txtaux(Index).Text)
    'If Txtaux(Index).Text = "" Then Exit Sub
    If Txtaux(Index).BackColor = vbYellow Then
        Txtaux(Index).BackColor = vbWhite
    End If
    
    If Txtaux(Index).Text = "" Then Exit Sub
    
    If ValorAnterior = Txtaux(Index).Text Then Exit Sub
    
    If modo = 3 And ConCaracteresBusqueda(Txtaux(Index).Text) Then Exit Sub

    Select Case Index
    Case 0
        If Not IsNumeric(Txtaux(0).Text) Then
            MsgBox "Código del forma de pago tiene que ser numérico", vbExclamation
            Exit Sub
        End If
        Txtaux(0).Text = Format(Txtaux(0).Text, "00")
    Case 1
'        If InStr(1, txtAux(Index).Text, "'") > 0 Then
'             MsgBox "No puede introducir el carácter ' en ningún campo de texto", vbExclamation
'             Exit Sub
'        End If
        Txtaux(Index).Text = Format(Txtaux(Index).Text, ">")
   Case 2, 3, 4, 6, 7, 9
        If Not IsNumeric(Txtaux(Index).Text) Then
            MsgBox "El campo tiene que ser numérico", vbExclamation
            Exit Sub
        End If
        If Index = 2 Then
            Txtaux(Index).Text = Format(Txtaux(Index).Text, "0")
        Else
        If Index = 7 Then
            Txtaux(Index).Text = Format(Txtaux(Index).Text, "00")
            Txtaux(8).Text = DevuelveDesdeBD(1, "nomforpa", "sforpa", "codforpa|", Txtaux(7).Text & "|", "N|", 1)
        Else
        If Index = 6 Then
            Txtaux(6).Text = Format(Txtaux(6), "###,##0.00")
        End If
        End If
        End If
        If Index = 9 Then
            Txtaux(9).Text = Format(Txtaux(9).Text, "00")
            Txtaux(11).Text = ""
            Txtaux(11).Text = DevuelveDesdeBD(1, "nombanpr", "sbanpr", "codbanpr|", Txtaux(9).Text & "|", "N|", 1)
        End If
    End Select
    
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
b = CompForm(Me)
If Not b Then Exit Function

If modo = 1 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD(1, "codforpa", "sforpa", "codforpa|", Txtaux(0).Text & "|", "N|", 1)
     If Datos <> "" Then
        MsgBox "Ya existe la Forma de Pago : " & Txtaux(0).Text, vbExclamation
        b = False
    End If
End If
DatosOk = b
End Function

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Combo1.BackColor = vbWhite
      SendKeys "{tab}"
      KeyAscii = 0
    End If
End Sub

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub FormaPagoAlternativa()
        Txtaux(6).Text = adodc1.Recordset.Fields(6).Value
        Txtaux(6).Text = Format(Txtaux(6), "###,##0.00")
        Txtaux(7).Text = adodc1.Recordset.Fields(7).Value
        Txtaux(7).Text = Format(Txtaux(7).Text, "00")
        If Txtaux(7) <> "" Then
             Txtaux(8).Text = DevuelveDesdeBD(1, "nomforpa", "sforpa", "codforpa|", Txtaux(7).Text & "|", "N|", 1)
        End If
End Sub

Private Sub BancoPropio()
        Txtaux(9).Text = adodc1.Recordset.Fields(8).Value
        Txtaux(9).Text = Format(Txtaux(9), "00")
        Txtaux(11).Text = ""
        If Txtaux(9).Text <> "" Then
            Txtaux(11).Text = DevuelveDesdeBD(1, "nombanpr", "sbanpr", "codbanpr|", Txtaux(9).Text & "|", "N|", 1)
        End If
End Sub

