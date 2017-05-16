VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCSNTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizar Fichero CSN"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "frmCSNTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7185
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmCSNTest.frx":030A
      Left            =   1395
      List            =   "frmCSNTest.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   585
      Width           =   1305
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   165
      Left            =   195
      TabIndex        =   5
      Top             =   5925
      Visible         =   0   'False
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   180
      Top             =   6090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDialog 
      Height          =   330
      Left            =   6615
      Picture         =   "frmCSNTest.frx":032F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   165
      Width           =   360
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1020
      Width           =   6780
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1395
      TabIndex        =   0
      Top             =   165
      Width           =   5190
   End
   Begin VB.CommandButton cmdGoGoGo 
      Caption         =   "Comenzar"
      Height          =   450
      Left            =   4725
      TabIndex        =   3
      Top             =   6120
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      Height          =   450
      Left            =   5880
      TabIndex        =   4
      Top             =   6120
      Width           =   1110
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   645
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero CSN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   225
      Width           =   1200
   End
End
Attribute VB_Name = "frmCSNTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Conexiones necesarias.
Dim InitDir As String

Private Sub cmdDialog_Click()
  If InitDir = "" Then InitDir = App.Path & "\TraspasoCSN"
  cd1.InitDir = InitDir
  cd1.Filter = "Todos los archivos (*.*)|*.*"
  cd1.ShowOpen
  If cd1.FileName <> "" Then Text1.Text = cd1.FileName
End Sub

' Salimos.
Private Sub cmdCancelar_Click()
  Unload Me
End Sub

' Recorremos cada elemento de la lista y ejecutamos la función correspondiente.
Private Sub cmdGoGoGo_Click()
Dim fich As CFichero
Dim cache As Dictionary
Dim dni As String, sql As String
Dim Index As Integer, cont As Integer, total As Integer, rscont As Integer
Dim Rs As ADODB.Recordset
Dim min As Currency, max As Currency, valor As Currency, suma As Currency
On Error GoTo TratarError
  
  If Text1.Text = "" Then
    MsgBox "Debe especificar un nombre de fichero.", vbInformation + vbOKOnly, "¡Atención!"
    Exit Sub
  End If
  
  If Dir(Text1.Text) = "" Then
    MsgBox "El fichero o ruta especificado no existe.", vbInformation + vbOKOnly, "¡Atención!"
    Exit Sub
  End If
  
  ' Preliminares.
  Me.Enabled = False
  Set fich = New CFichero
  Set cache = New Dictionary
  fich.abrir Text1.Text
  PB1.Value = 0
  PB1.max = 100
  PB1.Visible = True
  
  Text2.Text = "[" & Text1.Text & "]" & vbCrLf & String(90, "-") & vbCrLf
  Text2.Text = Text2.Text & "Obteniendo información del fichero... "
  
  ' Procesamos el fichero.
  If Combo1.ListIndex = 0 Then
    cont = 33
    rscont = 9
  Else
    cont = 25
    rscont = 8
  End If
  While fich.leerLinea
  
    If Left(fich.linea, 2) = "05" Then
      valor = Val(Mid(fich.linea, cont, rscont)) / 100
      dni = Trim(Mid(fich.linea, 3, 11))
      cache(dni) = CCur(cache(dni)) + valor
    End If
    PB1.Value = fich.porcentaje
    DoEvents
  
  Wend
  
  ' Cerramos el fichero.
  fich.cerrar
  Set fich = Nothing
  
  ' Cargamos los rangos desde la base de datos.
  Text2.Text = Text2.Text & "Ok." & vbCrLf & "Cargando rangos CSN de la base de datos..."
  sql = "SELECT * FROM rangoscsn WHERE tipo=" & Combo1.ListIndex & " ORDER BY orden"
  Set Rs = New ADODB.Recordset
  Rs.Open sql, Conn, , , adCmdText
  
  ' Abierto el recordset, continuamos...
  Text2.Text = Text2.Text & "Ok." & vbCrLf & "Procesando información obtenida: " & vbCrLf & vbCrLf
  Text2.Text = Text2.Text & "Rango CSN" & vbTab & "   mSv" & vbTab & vbTab & "nº usuarios" & vbCrLf
  Text2.Text = Text2.Text & String(31, "-") & vbTab & "   " & String(20, "-") & vbTab & String(35, "-") & vbCrLf
  
  ' Procesamos la información adquirida.
  PB1.Value = 0
  rscont = 0
  total = cache.Count
  While Not Rs.EOF
    
    ' Obtenemos el mínimo y el máximo.
    If Not IsNull(Rs!desde) Then min = CCur(Rs!desde) Else min = -1
    If Not IsNull(Rs!hasta) Then max = CCur(Rs!hasta) Else max = 999999999
    
    ' Inicializamos contadores y comenzamos bucle.
    Index = 0
    cont = 0
    suma = 0
    While cache.Count > Index
      dni = cache.Keys(Index)
      
      ' Si la cantidad de mSv está en el intervalo actual, lo añadimos y eliminamos
      ' el elemento del diccionario. Si no, pasamos al siguiente elemento.
      If cache.Item(dni) > CCur(min) And cache.Item(dni) <= CCur(max) Then
        suma = suma + CCur(cache(dni))
        cache.Remove dni
        cont = cont + 1
        PB1.Value = CInt((total - cache.Count) * (100 / total))
        DoEvents
      Else
        Index = Index + 1
      End If
      
    Wend
    rscont = rscont + 1
    
    ' Mostramos la información.
    If rscont = 1 Then
      Text2.Text = Text2.Text & "Dosis < NR"
    ElseIf rscont = 2 Then
      Text2.Text = Text2.Text & "NR <= Dosis <= " & max
    ElseIf IsNull(Rs!hasta) Then
      Text2.Text = Text2.Text & min & " < Dosis"
    Else
      Text2.Text = Text2.Text & min & " < Dosis <= " & max
    End If
    
    Text2.Text = Text2.Text & vbTab & "   " & FormatNumber(suma, 2) & IIf(Len(FormatNumber(suma, 2)) > 7, "", vbTab) & vbTab & cont & vbCrLf
    Rs.MoveNext
    
  Wend
  
  ' Acciones finales.
  PB1.Visible = False
  Rs.Close
  Set Rs = Nothing
  cache.RemoveAll
  Set cache = Nothing
  Me.Enabled = True
  Exit Sub
  
TratarError:
  If Not Rs Is Nothing Then
    MsgBox "Ha ocurrido un error accediendo a la base de datos. Si el problema persiste, contacte con el servicio técnico.", vbCritical + vbOKOnly, "¡Error!"
    If Rs.State <> adStateClosed Then Rs.Close
    Set Rs = Nothing
  Else
    MsgBox "Ha ocurrido un error durante la carga del fichero " & Text1.Text & ". Es posible que se haya equivocado al indicar el nombre o ruta del mismo.", vbCritical + vbOKOnly, "¡Error!"
    fich.cerrar
    Set fich = Nothing
  End If
  If Not cache Is Nothing Then
    cache.RemoveAll
    Set cache = Nothing
  End If
  Err.Clear
  Me.Enabled = True
End Sub

Private Sub Form_Load()
  Combo1.ListIndex = 0
End Sub
