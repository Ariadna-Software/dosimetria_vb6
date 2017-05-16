VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmGenRecepDosim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Automática de Recepción de Dosímetros"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmGenRecepDosim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListTipoMedicion 
      Height          =   3570
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   7275
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   990
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1500
         Width           =   1185
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3750
         TabIndex        =   0
         Top             =   2460
         Width           =   1425
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   2070
         TabIndex        =   7
         Top             =   2460
         Width           =   1425
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5550
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   1470
         Width           =   1185
      End
      Begin VB.TextBox text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   3390
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   1500
         Width           =   1125
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   300
         Left            =   420
         TabIndex        =   8
         Top             =   2010
         Visible         =   0   'False
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   450
         TabIndex        =   10
         Top             =   1530
         Width           =   390
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   0
         Left            =   3120
         MouseIcon       =   "frmGenRecepDosim.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmGenRecepDosim.frx":0E1C
         ToolTipText     =   "Seleccionar fecha"
         Top             =   1530
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   $"frmGenRecepDosim.frx":0EA7
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1005
         Left            =   540
         TabIndex        =   5
         Top             =   390
         Width           =   6525
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   20
         Left            =   2490
         TabIndex        =   4
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Paridad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   16
         Left            =   4710
         TabIndex        =   3
         Top             =   1530
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmGenRecepDosim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    
    BloqueoManual False, "RECEPCION", ""
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim Tipo As String
Dim tipodos As Byte
Dim Cont As Integer
Dim fecaux As Date
Dim Cad As String

    On Error GoTo eErrorCarga

    Screen.MousePointer = vbHourglass

'   cambiamos la fecha de creacion para que unicamente se produzca el 01 de cada mes
    fecaux = CDate(Text1(0).Text)
    Cad = "01/" & Format(Month(fecaux), "00") & "/" & Format(Year(fecaux), "0000")
    Text1(0).Text = Cad

    Tipo = "I"
    If Combo1.ListIndex = 0 Then Tipo = "P"

    If Combo2.Text = "Personal" Then
        tipodos = 0
    Else 'area
        tipodos = 2
    End If
        
    Conn.BeginTrans

    If Combo2.ListIndex <> -1 Then
        ' no se recepcionan ni organo ni control ni fondo
        sql = "select * from dosimetros where mes_p_i = '" & Trim(Tipo) & "' and "
        sql = sql & "f_retirada is null and tipo_dosimetro = " & Format(tipodos, "0")
        
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    If Cont > 32000 Then Cont = 32000
    pb1.max = Cont + 1
    pb1.Visible = True
    pb1.Value = 0
    Me.Refresh

    
    sql1 = "insert into recepdosim (n_reg_dosimetro, n_dosimetro, c_empresa, "
    sql1 = sql1 & "c_instalacion, dni_usuario, fecha_recepcion,f_creacion_recep,mes_p_i,"
    sql1 = sql1 & "tipo_dosimetro) VALUES ("
    
    Rs.MoveFirst
    
    While Not Rs.EOF
        Cad = ""
        Cad = DevuelveDesdeBD(1, "n_reg_dosimetro", "recepdosim", "n_reg_dosimetro|n_dosimetro|dni_usuario|f_creacion_recep|mes_p_i|tipo_dosimetro|", Rs!n_reg_dosimetro & "|" & Trim(Rs!n_dosimetro) & "|" & Trim(Rs!dni_usuario) & "|" & Format(Text1(0).Text, FormatoFecha) & "|" & Trim(Rs!mes_p_i) & "|" & tipodos & "|", "N|T|T|F|T|N|", 6)
        If Cad = "" Then
            
            
            sql2 = Rs!n_reg_dosimetro & ",'" & Trim(Rs!n_dosimetro) & "','" & Trim(Rs!c_empresa) & "','"
            sql2 = sql2 & Trim(Rs!c_instalacion) & "','" & Trim(Rs!dni_usuario) & "',null,'"
            sql2 = sql2 & Format(Text1(0).Text, FormatoFecha) & "','" & Tipo & "',"
            
    '        If Combo2.ListIndex = 0 Then
                sql2 = sql2 & Format(tipodos, "0") & ")"
    '        Else
    '            Select Case rs!dni_usuario
    '                Case "0" 'fondo
    '                    sql2 = sql2 & "3)"
    '                Case "999999996" 'transito
    '                    sql2 = sql2 & "4)"
    '                Case "666666666" 'libre
    '                    sql2 = sql2 & "6)"
    '                Case "888888888" 'control
    '                    sql2 = sql2 & "5)"
    '                Case Else 'area
    '                    sql2 = sql2 & "2)"
    '            End Select
    '        End If
            
            sql = sql1 & sql2
        
            Conn.Execute sql
            
        End If
            
    
        pb1.Value = pb1.Value + 1
        pb1.Refresh
    
        Rs.MoveNext
    Wend

    Screen.MousePointer = vbDefault


eErrorCarga:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la carga de Recepción de Dosímetros. Revise."
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        MsgBox "Proceso Finalizado Correctamente", vbExclamation, "Recepción de Dosímetros."
        cmdCancelar_Click
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
    If Not BloqueoManual(True, "RECEPCION", "") Then
        MsgBox "Existe algún usuario Recepcionando. Inténtelo más tarde.", vbExclamation, "¡Error!"
        Unload Me
        Exit Sub
    End If
    
    pb1.Visible = False
    
    CargarCombo1
    
    Text1(0).Text = Format(Now, "dd/mm/yyyy")
    Text1(0).Text = "01" & Mid(Text1(0).Text, 3)
    
    If ((Month(Now) / 2) = Round2(Month(Now) / 2)) Then
        Combo1.ListIndex = 1
    Else
        Combo1.ListIndex = 0
    End If
    
    Combo2.ListIndex = 0
    
End Sub

Private Sub CargarCombo1()
Dim Rs As Recordset
Dim sql As String
'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Si, 1-No

    Combo1.Clear
    Combo1.AddItem "Par"
    Combo1.ItemData(Combo1.NewIndex) = 0
    
    Combo1.AddItem "Impar"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo2.Clear
    Combo2.AddItem "Personal"
    Combo2.ItemData(Combo2.NewIndex) = 0
    
    Combo2.AddItem "Area"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    

End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
'    If Modo = 0 Or Modo = 2 Then Exit Sub
    Select Case Index
       Case 0 'fecha
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            ' ### 29/03/2006 DavidV (error en el índice del control Text1)
            If Text1(0).Text <> "" Then
                If IsDate(Text1(0).Text) Then f = Text1(0).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(0).Text = frmC.fecha
            Text1(0).Text = Format(Text1(0).Text, "dd/mm/yyyy")
            Set frmC = Nothing
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Dim sql As String
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
   Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim fec As Date

    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" Then
        MsgBox "Debe de introducir un valor en este campo", vbExclamation
        PonerFoco Text1(0)
    End If
    If Text1(Index).Text <> "" Then
      If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
            Exit Sub
      End If
      Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
      fec = CDate(Text1(Index).Text)
      If ((Month(fec) / 2) = Round2(Month(fec) / 2)) Then
           Combo1.ListIndex = 1
      Else
           Combo1.ListIndex = 0
      End If
    End If
End Sub

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

