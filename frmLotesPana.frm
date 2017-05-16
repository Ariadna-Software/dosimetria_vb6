VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLotesPana 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factores de Corrección de Lotes Panasonic"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmLotesPana.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   4
      Left            =   4365
      TabIndex        =   6
      Tag             =   "Tipo|T|N|||lotespana|tipo||S|"
      Text            =   "Tipo"
      Top             =   5505
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Tag             =   "Cristal 1|N|N|0|999.999|lotespana|cristal_a|||"
      Text            =   "Dato1"
      Top             =   5520
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   3
      Left            =   3165
      TabIndex        =   3
      Tag             =   "Dosímetro Final|N|N|||lotes|dosimetro_final||N|"
      Text            =   "Dato4"
      Top             =   5460
      Width           =   810
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   2
      Left            =   2025
      TabIndex        =   2
      Tag             =   "Dosímetro Inicial|N|N|||lotes|dosimetro_inicial||S|"
      Text            =   "Dato3"
      Top             =   5460
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   1125
      TabIndex        =   1
      Tag             =   "Cristal 2|N|N|0|999.999|lotespana|cristal_b|||"
      Text            =   "Dato2"
      Top             =   5460
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLotesPana.frx":030A
      Height          =   4365
      Left            =   135
      TabIndex        =   5
      Top             =   945
      Width           =   5340
      _ExtentX        =   9419
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   5790
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4455
      TabIndex        =   0
      Top             =   5790
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
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
            Object.Tag             =   "2"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
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
         Left            =   4425
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   975
      Left            =   780
      Top             =   4020
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
      Left            =   3105
      Top             =   -15
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
            Picture         =   "frmLotesPana.frx":031F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":0431
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":0543
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":0655
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":0767
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":0879
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":1153
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":1A2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":2307
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":2BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":34BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":390D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":3A1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":3B31
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":3C43
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLotesPana.frx":42BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   45
      TabIndex        =   12
      Top             =   480
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   8784
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Solapa"
      TabPicture(0)   =   "frmLotesPana.frx":43CF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Extremidad"
      TabPicture(1)   =   "frmLotesPana.frx":43EB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
End
Attribute VB_Name = "frmLotesPana"
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

Toolbar1.Buttons(1).Enabled = b
Toolbar1.Buttons(2).Enabled = b
Toolbar1.Buttons(6).Enabled = b And vUsu.NivelUsu <= 2
Toolbar1.Buttons(7).Enabled = b And Not b2 And vUsu.NivelUsu <= 2
Toolbar1.Buttons(8).Enabled = b And Not b2 And vUsu.NivelUsu <= 2
cmdAceptar.Visible = Not b
cmdCancelar.Visible = Not b
DataGrid1.Enabled = b And Not b2

'Si estamo mod or insert
If Modo = 2 Then
   txtAux(2).BackColor = &H80000018
   txtAux(3).BackColor = &H80000018
Else
   txtAux(2).BackColor = &H80000005
   txtAux(3).BackColor = &H80000005
End If

If SSTab1.Caption = "Extremidad" Then
  txtAux(1).BackColor = &H80000018
  txtAux(1).Enabled = False
Else
  txtAux(1).BackColor = &H80000005
  txtAux(1).Enabled = False ' (VRS 1.2.2) Antes era true
End If

txtAux(2).Enabled = (Modo <> 2)
txtAux(3).Enabled = (Modo <> 2)

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

    CadenaConsulta = "Select * from lotespana where 1=1 "
    CargaGrid "where dosimetro_inicial = -1 and tipo ='" & txtAux(4).Text & "'"
    Me.lblIndicador.Caption = "BUSQUEDA"
    'Buscar
    For I = 0 To txtAux.Count - 2
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
    
'    i = adodc1.Recordset!tipoconce
'    Combo1.ListIndex = i - 1
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
End Sub

Private Sub BotonEliminar()
Dim sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    sql = "Seguro que desea eliminar el lote:"
    sql = sql & vbCrLf & "Dosímetros " & adodc1.Recordset.Fields(2) & " - " & adodc1.Recordset.Fields(3)
    If MsgBox(sql, vbQuestion + vbYesNoCancel, "¡Atención!") = vbYes Then
        'Hay que eliminar
        sql = "Delete from lotespana where dosimetro_inicial=" & adodc1.Recordset!dosimetro_inicial & " and tipo ='" & txtAux(4).Text & "'"
        conn.Execute sql
        CargaGrid "tipo ='" & txtAux(4).Text & "'"
    End If

Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar lote."
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
'-- Siempre igualamos Cristal a y b (VRS 1.2.2)--
txtAux(1) = txtAux(0)
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                CargaGrid "tipo ='" & txtAux(4).Text & "'"
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me, 1) Then
                    I = adodc1.Recordset.Fields(2)
                    PonerModo 0
                    CargaGrid "tipo ='" & txtAux(4).Text & "'"
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(2).Name & " =" & I)
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
        '.Buttons(10).Image = 10
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
    PonerOpcionesMenuGeneral Me

    CadenaConsulta = "Select * from lotespana where 1=1 "
    
    SSTab1_Click 0

    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
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
            FrmListado.Opcion = 30 'Listado factores de Lotes Panasonic
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
    adodc1.ConnectionString = conn
    If sql <> "" Then
        sql = CadenaConsulta & " AND " & sql
        Else
        sql = CadenaConsulta
    End If
    sql = sql & " ORDER BY dosimetro_inicial"
    adodc1.RecordSource = sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    DataGrid1.Enabled = True
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    ' ### [DavidV] Añadido el cristal 1 (para los dosímetros de anillo y muñeca).

    'cristales a y b
    cristal1 = (txtAux(4).Text <> "S")
'    DataGrid1.Columns(0).Caption = IIf(cristal1, "Cristal 1=", "Cristal 2")
'    DataGrid1.Columns(1).Caption = IIf(cristal1, "Cristal 2", "Cristal 3")
    '-- VRS:1.3.3
    DataGrid1.Columns(0).Caption = "----C----"
    DataGrid1.Columns(1).Caption = "---------"
    DataGrid1.Columns(0).Width = 850
    DataGrid1.Columns(1).Width = 850
    txtAux(0).Tag = DataGrid1.Columns(0).Caption & Mid(txtAux(0).Tag, InStr(1, txtAux(0).Tag, "|"))
    txtAux(1).Tag = DataGrid1.Columns(1).Caption & Mid(txtAux(1).Tag, InStr(1, txtAux(1).Tag, "|"))
    
    'fecha de inicio
    DataGrid1.Columns(2).Caption = "Dosímetro Inicial"
    DataGrid1.Columns(2).Width = 1550
    
    'fecha de finalizacion
    DataGrid1.Columns(3).Caption = "Dosímetro Final"
    DataGrid1.Columns(3).Width = 1450
        
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
        
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF And vUsu.NivelUsu <= 2
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF And vUsu.NivelUsu <= 2
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
Dim punto As Integer
Dim coma As Integer

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
               
                If SSTab1.Caption = "Extremidad" Then txtAux(1).Text = txtAux(0).Text
                txtAux(Index).Text = Format(valor, "##0.000")
            End If
        
      Case 2, 3 ' numéricos enteros
            If EsNumerico(txtAux(Index).Text) Then
                punto = InStr(1, txtAux(Index).Text, ".")
                coma = InStr(1, txtAux(Index).Text, ",")
                If punto > 0 Or coma > 0 Then
                    MsgBox "El campo no puede contener decimales.", vbExclamation, "¡Error!"
                  If punto > 0 Then
                    txtAux(Index).Text = Int(CCur(TransformaPuntosComas(txtAux(Index).Text)))
                  Else
                    txtAux(Index).Text = Int(txtAux(Index).Text)
                  End If
                End If
            End If
      
    End Select
    
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim sql As String
Dim error As Boolean
Dim Consulta As ADODB.Recordset
Dim dosi As Long
Dim dosf As Long

  error = Not CompForm(Me)
  If error Then Exit Function

  If Modo = 1 Then
    'Estamos insertando
    If Val(txtAux(2).Text) > Val(txtAux(3).Text) Then
      MsgBox "El dosímetro inicial no puede ser mayor que el final. Reintroduzca.", vbExclamation, "¡Error!"
      DatosOk = False
      Exit Function
    End If
    
    sql = "select * from lotespana where tipo = '" & txtAux(4).Text & "'"
    
    Set Consulta = New ADODB.Recordset
    Consulta.Open sql, conn, , , adCmdText
    error = False
    While Not Consulta.EOF And Not error
      dosi = Consulta!dosimetro_inicial
      dosf = Consulta!dosimetro_final
      ' ¿Alguno de los nº de dosímetro se encuentran en un rango de lote existente?
      If txtAux(2).Text >= dosi And txtAux(2).Text <= dosf Then error = True
      If txtAux(3).Text >= dosi And txtAux(3).Text <= dosf Then error = True
      ' ¿Alguno de los rangos de lotes existentes se encuentran en el rango que queremos
      ' introducir?
      If dosi > txtAux(2).Text And dosi < txtAux(3).Text Then error = True
      If dosf > txtAux(2).Text And dosf < txtAux(3).Text Then error = True
       
      Consulta.MoveNext
    Wend
    Set Consulta = Nothing
    
    If error Then
      MsgBox "El rango de dosímetros ya está contemplado en otro lote. Reintroduzca.", vbExclamation, "¡Error!"
    End If
  End If

  DatosOk = Not error

End Function

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub SSTab1_Click(Index As Integer)
  txtAux(4).Text = Left(SSTab1.Caption, 1)
  CargaGrid "tipo = '" & txtAux(4).Text & "' "
  
End Sub


