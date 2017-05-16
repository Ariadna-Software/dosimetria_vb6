VERSION 5.00
Begin VB.Form frmExportarAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar fichero anual de dosis"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmExportarAnual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListDosisOpeAcum12 
      Height          =   2460
      Left            =   60
      TabIndex        =   6
      Top             =   -15
      Width           =   6120
      Begin VB.Frame Frame15 
         Caption         =   "Tipo de dosimetría "
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
         Height          =   645
         Left            =   225
         TabIndex        =   7
         Top             =   1155
         Width           =   5670
         Begin VB.OptionButton Option1 
            Caption         =   "Ambas"
            Height          =   225
            Index           =   2
            Left            =   4080
            TabIndex        =   3
            Top             =   270
            Value           =   -1  'True
            Width           =   1260
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Área"
            Height          =   225
            Index           =   1
            Left            =   2370
            TabIndex        =   2
            Top             =   270
            Width           =   1260
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cuerpo"
            Height          =   225
            Index           =   0
            Left            =   675
            TabIndex        =   1
            Top             =   270
            Width           =   1260
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3030
         TabIndex        =   0
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdAcept 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año a Exportar"
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
         Index           =   34
         Left            =   1620
         TabIndex        =   9
         Top             =   750
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Exportar fichero anual de dosis"
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
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5715
      End
   End
End
Attribute VB_Name = "frmExportarAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Aceptar.
Private Sub CmdAcept_Click()
Dim ano As Integer
Dim fechaini As String
Dim fechafin As String
Dim msgerr As Integer
Dim sql As String
Dim fich1 As String
Dim fich2 As String
On Error GoTo ECmdAcept_Click

  ' Comprobamos si es un año válido.
  If IsNumeric(Text1.Text) Then
    ano = CInt(Text1.Text)
  Else
    ano = -1
  End If
    
  If Not (ano < Year(Now) And ano >= 0) Then
    MsgBox "Año incorrecto. Reintroduzca.", vbExclamation, "¡Error!"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  ' Preparativos previos al proceso.
  Screen.MousePointer = vbHourglass
  CmdAcept.Enabled = False
  CmdCancel.Enabled = False
  fechaini = ano & "-01-01"
  fechafin = ano & "-12-31"
  If Not Dir(App.Path & "\temp", vbDirectory) <> "" Then MkDir App.Path & "\temp"
  msgerr = 0
  
  ' En caso de dosimetría de cuerpo o ambas.
  If Option1(0).Value Or Option1(2).Value Then
    sql = "SELECT dni_usuario,dosis_superf,dosis_profunda,f_dosis,instalaciones.c_instalacion,"
    sql = sql & "descripcion,direccion FROM dosiscuerpo LEFT JOIN instalaciones "
    sql = sql & "USING(c_empresa,c_instalacion) WHERE f_dosis BETWEEN '" & fechaini
    sql = sql & "' AND '" & fechafin & "' AND (c_tipo=0 or c_tipo=2)"
    fich1 = App.Path & "\temp\DosisCuerpo_" & ano & ".txt"
    If Not ExportarFichero(sql, fich1) Then msgerr = 1
  End If
  
  ' En caso de dosimetría de área o ambas.
  If Option1(1).Value Or Option1(2).Value Then
    sql = "SELECT instalaciones.c_instalacion,n_dosimetro,dosis_superf,dosis_profunda,f_dosis,"
    sql = sql & "descripcion,direccion FROM dosisarea LEFT JOIN instalaciones "
    sql = sql & "USING(c_empresa,c_instalacion) WHERE f_dosis BETWEEN '" & fechaini
    sql = sql & "' AND '" & fechafin & "' AND (c_tipo=1 or c_tipo=2)"
    fich2 = App.Path & "\temp\DosisArea_" & ano & ".txt"
    If Not ExportarFichero(sql, fich2) Then msgerr = msgerr + 2
  End If
    
  ' Mensaje final.
  Screen.MousePointer = vbDefault
  CmdAcept.Enabled = True
  CmdCancel.Enabled = True
  Select Case msgerr
    Case 0
      Select Case True
        Case Option1(0).Value
          sql = "El fichero ha sido guardado en:" & vbCrLf & vbCrLf & fich1
        Case Option1(1).Value
          sql = "El fichero ha sido guardado en:" & vbCrLf & vbCrLf & fich2
        Case Option1(2).Value
          sql = "Los ficheros han sido guardados en:" & vbCrLf & vbCrLf & fich1 & vbCrLf & fich2
      End Select
      MsgBox "Exportación realizada con éxito. " & sql, vbOKOnly + vbInformation, "Exportación completada"
    Case 1
      If Option1(1).Value Or Option1(2).Value Then
        sql = "El fichero de área ha sido guardado en:" & vbCrLf & vbCrLf & fich2
      Else
        sql = ""
      End If
      MsgBox "No hay datos en dosis de cuerpo para ese año. " & sql, vbOKOnly + vbExclamation, "Exportar Fichero Anual"
    Case 2
      If Option1(0).Value Or Option1(2).Value Then
        sql = "El fichero de cuerpo ha sido guardado en:" & vbCrLf & vbCrLf & fich1
      Else
        sql = ""
      End If
      MsgBox "No hay datos en dosis de área para ese año." & sql, vbOKOnly + vbExclamation, "Exportar Fichero Anual"
    Case 3
      MsgBox "No hay datos en dosis de ninguna dosimetría para ese año.", vbOKOnly + vbExclamation, "Exportar Fichero Anual"
  End Select
  Exit Sub

ECmdAcept_Click:
  Screen.MousePointer = vbDefault
  CmdAcept.Enabled = True
  CmdCancel.Enabled = True
  MsgBox "Error número " & Err.Number & " exportando el fichero de datos:" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly, "¡Error!"
  Err.Clear

End Sub

' Cancel.
Private Sub cmdCancel_Click()
  Unload Me
End Sub

' Exporta los datos a un fichero "empipado".
Private Function ExportarFichero(sql As String, archivo As String) As Boolean
Dim Rs As ADODB.Recordset
Dim NF As Integer
On Error GoTo EExportarFichero

  Set Rs = New ADODB.Recordset
  Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
  
  If Not Rs.EOF Then
    
    ' Recorremos el recordset y vamos escribiendo las lineas en el archivo.
    NF = FreeFile
    Open archivo For Output As #NF
    While Not Rs.EOF
      sql = Rs.GetString(, 1, "|")
      Print #NF, Trim(Left(sql, Len(sql) - 1)) & "|"
    Wend
    
    ' Cerramos, todo ok.
    Close #NF
    ExportarFichero = True
  
  Else
    ExportarFichero = False
  End If
  
  ' Cerramos.
  Rs.Close
  Set Rs = Nothing
  Exit Function
  
EExportarFichero:
  If Not Rs Is Nothing Then
    Rs.Close
    Set Rs = Nothing
  End If
  ExportarFichero = False
  Err.Raise Err.Number, "Exportar Fichero Anual", Err.Description
  
End Function

' El año pasado por defecto.
Private Sub Form_Load()
  Text1.Text = Year(Now) - 1
End Sub

' Impedir salir antes de terminar.
Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not CmdCancel.Enabled Then Cancel = True
End Sub

