VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmSelecPenalizaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de penalizaciones a aplicar..."
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmSelecPenalizaciones.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Boton 
      Caption         =   "Continuar"
      Height          =   420
      Index           =   3
      Left            =   6000
      TabIndex        =   0
      ToolTipText     =   "Seleccionar Todo"
      Top             =   4845
      Width           =   1590
   End
   Begin VB.CommandButton Boton 
      Caption         =   "Deseleccionar Todo"
      Height          =   420
      Index           =   2
      Left            =   4230
      TabIndex        =   4
      ToolTipText     =   "Seleccionar Todo"
      Top             =   4860
      Width           =   1590
   End
   Begin VB.CommandButton Boton 
      Caption         =   "Seleccionar Todo"
      Height          =   420
      Index           =   1
      Left            =   2460
      TabIndex        =   3
      ToolTipText     =   "Seleccionar Todo"
      Top             =   4860
      Width           =   1590
   End
   Begin VB.CommandButton Boton 
      Caption         =   "Seleccionar/Deseleccionar"
      Height          =   420
      Index           =   0
      Left            =   150
      TabIndex        =   2
      ToolTipText     =   "Seleccionar/Deseleccionar"
      Top             =   4860
      Width           =   2130
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   4350
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7673
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   5
      Top             =   90
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   7620
      Picture         =   "frmSelecPenalizaciones.frx":030A
      Stretch         =   -1  'True
      Top             =   -30
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "frmSelecPenalizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Configuración previa del Grid.
Private Sub ConfigurarGrid()
Dim I As Integer
Dim ancho As Variant

  With Grid1
    .Cols = 8
    ancho = Array(1260, 1110, 885, 1065, 735, 735, 735, 600)
    For I = 0 To .Cols - 1
      .ColWidth(I) = ancho(I)
      .CellFontBold = True
    Next I
    .ColAlignment(7) = flexAlignCenterCenter
 End With
End Sub

' Carga del Grid.
Private Sub LlenarGrid()
Dim rs As ADODB.Recordset
Dim fila As Integer

  With Grid1
    .Clear
    .Rows = 2
    .Row = 0
    .RowSel = 0
    .Col = 0
    .ColSel = .Cols - 1
    .Clip = "Empresa" & vbTab & "Instalación" & vbTab & "DNI" & vbTab & "Dosímetro" & vbTab & "Dosis S" & vbTab & "Dosis P" & vbTab & "Nº Reg" & vbTab & "Selec"
    .Row = 1
    
    Set rs = New Recordset
    rs.Open "select * from zdosisacum where codusu = " & vUsu.codigo & " order by n_dosimetro", Conn, , , adCmdText
    If Not rs.EOF Then
      rs.MoveFirst
      .Enabled = True
    Else
      .Enabled = False
    End If
    
    While Not rs.EOF
      fila = .Rows - 1
      .TextMatrix(fila, 0) = rs!c_empresa
      .TextMatrix(fila, 1) = rs!c_instalacion
      .TextMatrix(fila, 2) = rs!dni_usuario
      .TextMatrix(fila, 3) = rs!n_dosimetro
      .TextMatrix(fila, 4) = rs!dosissuper
      .TextMatrix(fila, 5) = rs!dosisprofu
      .TextMatrix(fila, 6) = rs!n_reg_dosimetro
      .Rows = .Rows + 1
      rs.MoveNext
    Wend
    Set rs = Nothing
    .FixedRows = 1
    If .Enabled Then
      .Rows = .Rows - 1
      Boton(0).Enabled = True
      Boton(1).Enabled = True
      Boton(2).Enabled = True
    Else
      Boton(0).Enabled = False
      Boton(1).Enabled = False
      Boton(2).Enabled = False
    End If
  End With
End Sub

' Click en cualquiera de los botones.
Private Sub Boton_Click(Index As Integer)
  
  Select Case Index
    
    ' Marcar/Desmarcar registros seleccionados.
    Case 0
      Seleccionar False
    ' Marcar TODOS los registros.
    Case 1
      Seleccionar True, True
    ' Desmarcar TODOS los registros.
    Case 2
      Seleccionar True, False
    ' Aplicar selección y salir.
    Case 3
      If Grid1.Enabled Then FiltrarSeleccion
      Unload Me
  End Select
End Sub

' Selecciona/Deselecciona según le pasemos parámetros.
Private Sub Seleccionar(ByVal Todo As Boolean, Optional ByVal seleccion As Boolean)
Dim I As Integer
Dim desde As Integer
Dim hasta As Integer
Dim valor As String

  ' Todo: Se ha pulsado seleccionar o deseleccionar TODO.
  If Todo Then
    valor = IIf(seleccion, "X", "")
    For I = 1 To Grid1.Rows - 1
      Grid1.TextMatrix(I, 7) = valor
    Next I
  Else
      
  ' Se ha pulsado Seleccionar/Deseleccionar lo que hay enfocado.
    If Grid1.Row > Grid1.RowSel Then
      desde = Grid1.RowSel
      hasta = Grid1.Row
    Else
      hasta = Grid1.RowSel
      desde = Grid1.Row
    End If
    
    For I = desde To hasta
      Grid1.TextMatrix(I, 7) = IIf(Grid1.TextMatrix(I, 7) = "X", "", "X")
    Next I

  End If
End Sub

' Seleccionamos/Deseleccionamos. la fila sobre la que hacemos doble click.
Private Sub grid1_dblclick()
  Seleccionar False
End Sub

' Al inicio, cargamos el Grid.
Private Sub Form_Load()
  
  Screen.MousePointer = vbHourglass
  ConfigurarGrid
  LlenarGrid
  If Not Grid1.Enabled Then
    Label1.Caption = "Ningún dosímetro a penalizar."
    CadenaDevueltaFormHijo = Label1.Caption
  Else
    If Grid1.Rows = 2 Then
      Label1.Caption = "1 dosímetro a penalizar."
    Else
      Label1.Caption = Grid1.Rows - 1 & " dosímetros a penalizar."
    End If
  End If
  CadenaDevueltaFormHijo = Label1.Caption
  Screen.MousePointer = vbDefault
End Sub

' Elimina de la tabla temporal aquellos elementos que NO han sido seleccionados
' en el Grid.
Private Sub FiltrarSeleccion()
Dim query As String
Dim I As Integer
On Error GoTo EFiltrarSeleccion

  Screen.MousePointer = vbHourglass
  With Grid1
    query = "delete from zdosisacum where codusu = " & vUsu.codigo & " and  n_reg_dosimetro in ("
    For I = 1 To Grid1.Rows - 1
      If .TextMatrix(I, .Cols - 1) <> "X" Then
        query = query & .TextMatrix(I, 6) & ", "
      End If
    Next I
    query = Left(query, Len(query) - 2) & ")"
    If Right(query, 3) <> "in)" Then Conn.Execute query
  End With
  Screen.MousePointer = vbDefault
  Exit Sub
  
EFiltrarSeleccion:

  Screen.MousePointer = vbDefault
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Sub


