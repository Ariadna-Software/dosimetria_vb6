VERSION 5.00
Begin VB.Form frmExportarInstalaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar fichero de códigos de Instalación"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmExportarInstalaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListDosisOpeAcum12 
      Height          =   2670
      Left            =   60
      TabIndex        =   8
      Top             =   -15
      Width           =   6120
      Begin VB.CheckBox Check1 
         Caption         =   "Sin fecha de baja"
         Height          =   195
         Left            =   4305
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   1035
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3060
         TabIndex        =   1
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   2145
         Width           =   975
      End
      Begin VB.CommandButton CmdAcept 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   2145
         Width           =   975
      End
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
         TabIndex        =   9
         Top             =   1395
         Width           =   5670
         Begin VB.OptionButton Option1 
            Caption         =   "Cuerpo"
            Height          =   225
            Index           =   0
            Left            =   675
            TabIndex        =   3
            Top             =   270
            Width           =   1260
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Área"
            Height          =   225
            Index           =   1
            Left            =   2370
            TabIndex        =   4
            Top             =   270
            Width           =   1260
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ambas"
            Height          =   225
            Index           =   2
            Left            =   4080
            TabIndex        =   5
            Top             =   270
            Value           =   -1  'True
            Width           =   1260
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   13
         Top             =   1065
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de alta"
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
         Index           =   2
         Left            =   390
         TabIndex        =   12
         Top             =   720
         Width           =   1125
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   825
         Picture         =   "frmExportarInstalaciones.frx":030A
         Top             =   1050
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   2805
         Picture         =   "frmExportarInstalaciones.frx":040C
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   11
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Exportar fichero de códigos de Instalación"
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
         TabIndex        =   10
         Top             =   240
         Width           =   5715
      End
   End
End
Attribute VB_Name = "frmExportarInstalaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim RC As String

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
  Err.Raise Err.Number, "Exportar Fichero Instalaciones", Err.Description
  
End Function

Private Sub CmdAcept_Click()
Dim sql As String
Dim orden As String
Dim fich1 As String
Dim fich2 As String
Dim msgerr As Integer
On Error GoTo ECmdAcept_Click

  ' Preparamos la fórmula.
  sql = "SELECT DISTINCT c_instalacion,descripcion FROM instalaciones WHERE 1=1"
  If Text1(0).Text <> "" Then sql = sql & " AND f_alta >='" & Format(Text1(0).Text, FormatoFecha) & "'"
  If Text1(1).Text <> "" Then sql = sql & " AND f_alta <='" & Format(Text1(1).Text, FormatoFecha) & "'"
  If Check1.Value = vbChecked Then sql = sql & " AND f_baja IS NULL"
  orden = " ORDER BY c_instalacion,descripcion"
  
  ' Deshabilitar controles.
  Screen.MousePointer = vbHourglass
  CmdAcept.Enabled = False
  CmdCancel.Enabled = False
  msgerr = 0
  
  ' Según el tipo de dosimetría seleccionado, hacemos las llamadas necesarias
  ' a ExportarFichero.
  If Option1(0).Value Or Option1(2).Value Then
    fich1 = App.Path & "\temp\Instalaciones_Cuerpo.txt"
    If Not ExportarFichero(sql & " AND (c_tipo=0 or c_tipo=2)" & orden, fich1) Then msgerr = 1
  End If
  If Option1(1).Value Or Option1(2).Value Then
    fich2 = App.Path & "\temp\Instalaciones_Area.txt"
    If Not ExportarFichero(sql & " AND (c_tipo=1 or c_tipo=2)" & orden, fich2) Then msgerr = msgerr + 2
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
      MsgBox "No hay altas de instalaciones de cuerpo para ese criterio. " & sql, vbOKOnly + vbExclamation, "Exportar Fichero Instalaciones"
    Case 2
      If Option1(0).Value Or Option1(2).Value Then
        sql = "El fichero de cuerpo ha sido guardado en:" & vbCrLf & vbCrLf & fich1
      Else
        sql = ""
      End If
      MsgBox "No hay altas de instalaciones de área para ese criterio." & sql, vbOKOnly + vbExclamation, "Exportar Fichero Instalaciones"
    Case 3
      MsgBox "No hay altas de instalaciones para ese criterio.", vbOKOnly + vbExclamation, "Exportar Fichero Instalaciones"
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


' Formatear la fecha.
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index))
    If Text1(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta: " & Text1(Index), vbExclamation, "¡Error!"
        Text1(Index).Text = ""
        Text1(Index).SetFocus
    End If
End Sub

' Comprueba que la fecha desde es menor o igual a la fecha hasta.
Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text1(Indice1).Text <> "" And Text1(Indice2).Text <> "" Then
        If CDate(Text1(Indice1).Text) > CDate(Text1(Indice2).Text) Then
            MsgBox "Fecha 'desde' mayor que fecha 'hasta'.", vbExclamation, "¡Error!"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function


' Impedir salir antes de terminar.
Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not CmdCancel.Enabled Then Cancel = True
End Sub

' Click en calendarios.
Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.fecha = Now
    If Text1(Index).Text <> "" Then frmC.fecha = CDate(Text1(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


