VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDosisExtremidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dosis Extremidades"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "frmDosisExtremidades.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   4770
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3765
      TabIndex        =   13
      Top             =   4815
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   3
      Left            =   3465
      TabIndex        =   12
      Tag             =   "Código Usuario|N|N|||tempnc|codusu|||"
      Text            =   "Código"
      Top             =   5130
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.CommandButton CmdFec 
      Height          =   255
      Index           =   0
      Left            =   3045
      TabIndex        =   6
      Top             =   5130
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Tag             =   "Nº Dosímetro|N|N|0||tempnc|n_dosimetro|||"
      Text            =   "Dat"
      Top             =   5130
      Width           =   630
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Tag             =   "Lectura|N|N|0||tempnc|cristal_2|||"
      Text            =   "Dato2"
      Top             =   5145
      Width           =   840
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   2
      Left            =   1950
      TabIndex        =   2
      Tag             =   "Fecha Lectura|F|N|||tempnc|fecha_lectura|dd/mm/yyyy||"
      Text            =   "Dato3"
      Top             =   5130
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   9
      Top             =   5250
      Width           =   2295
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   5415
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2505
      TabIndex        =   4
      Top             =   5415
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   4
      Left            =   4215
      TabIndex        =   3
      Tag             =   "Sistema|T|N|||tempnc|sistema|||"
      Text            =   "Sistema"
      Top             =   5145
      Visible         =   0   'False
      Width           =   390
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
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
            Object.ToolTipText     =   "Cálculo de mSv"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   975
      Left            =   555
      Top             =   2865
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
            Picture         =   "frmDosisExtremidades.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":1A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":22F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":2BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":38F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":3A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":3B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDosisExtremidades.frx":42A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDosisExtremidades.frx":43BA
      Height          =   4500
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   7938
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
Attribute VB_Name = "frmDosisExtremidades"
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
Public Sistema As String
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

For I = 0 To Txtaux.Count - 3
    Txtaux(I).Visible = Not b
Next I
For I = 0 To CmdFec.Count - 1
    CmdFec(I).Visible = Not b
Next I

If Modo = 2 Then CmdFec(0).Visible = False
Toolbar1.Buttons(1).Enabled = b
Toolbar1.Buttons(2).Enabled = b
Toolbar1.Buttons(6).Enabled = b
Toolbar1.Buttons(7).Enabled = b And Not b2
Toolbar1.Buttons(8).Enabled = b And Not b2
If Not adodc1.Recordset Is Nothing Then
  Toolbar1.Buttons(10).Enabled = b And Not adodc1.Recordset.EOF
Else
  Toolbar1.Buttons(10).Enabled = False
End If
cmdAceptar.Visible = Not b
cmdCancelar.Visible = Not b
DataGrid1.Enabled = b And Not b2

'Si estamo mod or insert
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
        anc = 700
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 500
    End If
    For I = 0 To Txtaux.Count - 3
        Txtaux(I).Text = ""
    Next I
    LLamaLineas anc, 0
    
    'Ponemos el foco
    Txtaux(2).Text = Text1.Text
    PonerFoco Txtaux(0)
    
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
        anc = 700
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 500
    End If
    
    'Llamamos al form
    Txtaux(0).Text = DataGrid1.Columns(0).Text
    Txtaux(1).Text = DataGrid1.Columns(1).Text
    Txtaux(2).Text = DataGrid1.Columns(2).Text
    
    LLamaLineas anc, 1
   
   'Como es modificar
   PonerFoco Txtaux(0)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    PonerModo xModo + 1
    'Fijamos el ancho
    Txtaux(0).Top = alto
    Txtaux(1).Top = alto
    Txtaux(2).Top = alto
    CmdFec(0).Top = alto
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
        adodc1.Recordset.Delete
        CargaGrid "codusu = " & vUsu.codigo
    End If

Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lectura"
End Sub

Private Sub cmdAceptar_Click()
Dim I As Variant
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid "codusu = " & vUsu.codigo
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos modificar
                If ModificaDesdeFormulario(Me, 1) Then
                    I = adodc1.Recordset.Bookmark
                    PonerModo 0
                    CargaGrid "codusu = " & vUsu.codigo
                    adodc1.Recordset.Bookmark = I
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB & " and codusu = " & vUsu.codigo
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
        CargaGrid "codusu = " & vUsu.codigo
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
            If Txtaux(2).Text <> "" Then
                If IsDate(Txtaux(2).Text) Then f = Txtaux(2).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Txtaux(2).Text = frmC.fecha
                mTag.DarFormato Txtaux(2)
            Set frmC = Nothing
       Case 1 ' fecha finalizacion
            f = Now
            If Txtaux(3).Text <> "" Then
                If IsDate(Txtaux(3).Text) Then f = Txtaux(3).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Txtaux(3).Text = frmC.fecha
            mTag.DarFormato Txtaux(3)
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
Dim sql As String

    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    Txtaux(4).Text = Sistema
    Txtaux(3).Text = vUsu.codigo
    Caption = "Dosis Extremidades " & IIf(Sistema = "H", "Harshaw 6600", "Panasonic")
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
    
        
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
      Toolbar1.Buttons(9).Visible = False
      Toolbar1.Buttons(10).Visible = False
    End If
    
    DesplazamientoVisible False
    PonerModo 0
    CadAncho = False
    'Cadena consulta
'    PonerOpcionesMenuGeneral Me
    sql = "delete from tempnc where codusu = " & vUsu.codigo
    Conn.Execute sql
    CadenaConsulta = "Select n_dosimetro, cristal_2, fecha_lectura "
    CadenaConsulta = CadenaConsulta & "from tempnc where codusu = " & vUsu.codigo
    If Txtaux(4).Text <> "" Then CadenaConsulta = CadenaConsulta & " and sistema = '" & Txtaux(4).Text & "'"
    Text1.Text = Format(Now, "dd/MM/yyyy")
    CargaGrid "codusu = " & vUsu.codigo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmrge_DatoSeleccionado(CadenaSeleccion As String)
    Txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
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

'----------------------------------------------------------------


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Rs As ADODB.Recordset
Dim sql As String


    Select Case Button.Index
           
    Case 6
        BotonAnyadir
    Case 7
        BotonModificar
    Case 8
        BotonEliminar
    Case 10
        ' calculo de mSv
        If MsgBox("Esta Seguro de Realizar la Migracion de Dosis No Homogéneas a mSv", vbQuestion + vbYesNo + vbDefaultButton1, "¡Atención!") = vbYes Then
          Screen.MousePointer = vbHourglass
          CalculoMsv 1, Sistema
          Screen.MousePointer = vbDefault
          frmImprimir.OtrosParametros = "usu= " & vUsu.codigo & "|"
          frmImprimir.Opcion = IIf(Sistema = "H", 39, 40)
          frmImprimir.Show vbModal
          ' Limpiamos la tabla.
          Conn.Execute "delete from tempnc where codusu = " & vUsu.codigo
          CargaGrid "codusu = " & vUsu.codigo
          PonerModo Modo
        End If
    Case 12
        Unload Me
    Case Else
    
    End Select
End Sub

Private Sub DesplazamientoVisible(Bol As Boolean)
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
    'sql = sql & " ORDER BY fecha_lectura"
    adodc1.RecordSource = sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    DataGrid1.Enabled = True
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    DataGrid1.Columns(0).Caption = "Nº Dosímetro"
    DataGrid1.Columns(0).Width = 1400
    DataGrid1.Columns(1).Caption = "Lectura (nC)"
    DataGrid1.Columns(1).Width = 1400
    DataGrid1.Columns(2).Caption = "F. Lectura"
    DataGrid1.Columns(2).Width = 1400
      
        'Fijamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        Txtaux(0).Width = DataGrid1.Columns(0).Width - 60
        Txtaux(1).Width = DataGrid1.Columns(1).Width - 60
        Txtaux(2).Width = DataGrid1.Columns(2).Width - 60
        Txtaux(0).Left = DataGrid1.Left + 350
        Txtaux(1).Left = Txtaux(0).Left + Txtaux(0).Width + 55
        Txtaux(2).Left = Txtaux(1).Left + Txtaux(1).Width + 55
        CmdFec(0).Left = Txtaux(2).Left + Txtaux(2).Width - CmdFec(0).Width
        
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
    With Txtaux(Index)
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
    Txtaux(Index).Text = Trim(Txtaux(Index).Text)
    If Txtaux(Index).Text = "" Then Exit Sub
    If Txtaux(Index).BackColor = vbYellow Then
        Txtaux(Index).BackColor = vbWhite
    End If
    
    If Txtaux(Index) = "" Then Exit Sub
    
    If ValorAnterior = Txtaux(Index).Text Then Exit Sub
    
    If Modo = 3 And ConCaracteresBusqueda(Txtaux(Index).Text) Then Exit Sub 'Busquedas
    
    Select Case Index
      Case 0 ' enteros
        If EsNumerico(Txtaux(Index).Text) Then
          If InStr(1, Txtaux(Index).Text, ",") Or InStr(1, Txtaux(Index).Text, ".") Or InStr(1, Txtaux(Index).Text, "-") Then
            MsgBox "El número de dosímetro ha de ser entero y positivo.", vbOKOnly + vbExclamation, "¡Error!"
            PonerFoco Txtaux(Index)
          End If
        Else
          PonerFoco Txtaux(Index)
        End If
      Case 1 ' numericos
            If EsNumerico(Txtaux(Index).Text) Then
                If InStr(1, Txtaux(Index).Text, ",") > 0 Then
                    valor = ImporteFormateado(Txtaux(Index).Text)
                Else
                    valor = CCur(TransformaPuntosComas(Txtaux(Index).Text))
                End If
                
                Txtaux(Index).Text = Format(valor, "##0.000")
            Else
              PonerFoco Txtaux(Index)
            End If
        
      Case 2 ' fechas
            If Txtaux(Index).Text <> "" Then
              If Not EsFechaOK(Txtaux(Index)) Then
                    MsgBox "Fecha incorrecta: " & Txtaux(Index).Text, vbExclamation, "¡Error!"
                    Txtaux(Index).Text = ""
                    PonerFoco Txtaux(Index)
                    Exit Sub
              End If
              Txtaux(Index).Text = Format(Txtaux(Index).Text, "dd/mm/yyyy")
              Text1.Text = Txtaux(Index).Text
            End If
      
    End Select
    
End Sub

Private Function DatosOk() As Boolean
  
  DatosOk = CompForm(Me)
    
End Function

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

