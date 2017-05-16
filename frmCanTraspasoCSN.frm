VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmCanTraspasoCSN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación Traspaso Automático de datos a C.S.N"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmCanTraspasoCSN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListTipoMedicion 
      Height          =   5910
      Left            =   30
      TabIndex        =   5
      Top             =   60
      Width           =   7275
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3510
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   0
         Tag             =   "JMCE"
         Top             =   2730
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         Caption         =   "Período de migración"
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
         Height          =   1005
         Left            =   450
         TabIndex        =   8
         Top             =   3180
         Width           =   6195
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   2
            Text            =   "Text5"
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1380
            TabIndex        =   1
            Text            =   "Text5"
            Top             =   420
            Width           =   1215
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   1
            Left            =   4005
            MouseIcon       =   "frmCanTraspasoCSN.frx":0CCA
            MousePointer    =   99  'Custom
            Picture         =   "frmCanTraspasoCSN.frx":0E1C
            ToolTipText     =   "Seleccionar fecha"
            Top             =   420
            Width           =   240
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   0
            Left            =   1080
            MouseIcon       =   "frmCanTraspasoCSN.frx":0EA7
            MousePointer    =   99  'Custom
            Picture         =   "frmCanTraspasoCSN.frx":0FF9
            ToolTipText     =   "Seleccionar fecha"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   3420
            TabIndex        =   10
            Top             =   450
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   420
            TabIndex        =   9
            Top             =   450
            Width           =   735
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3690
         TabIndex        =   4
         Top             =   5040
         Width           =   1425
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   2010
         TabIndex        =   3
         Top             =   5040
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   300
         Left            =   420
         TabIndex        =   7
         Top             =   4380
         Visible         =   0   'False
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label5 
         Caption         =   "Es MUY IMPORTANTE que se tenga claro el grupo de cancelacion, ya que  una vez cancelado es IMPOSIBLE RECUPERAR DICHOS REGISTROS."
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
         Height          =   555
         Left            =   180
         TabIndex        =   13
         Top             =   2040
         Width           =   6945
      End
      Begin VB.Label Label4 
         Caption         =   "Dosis NO Homogeneas - Dosis Homogeneas - Operarios en Instalaciones -    Instalaciones - Empresas"
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
         Height          =   765
         Left            =   1290
         TabIndex        =   12
         Top             =   1080
         Width           =   4125
      End
      Begin VB.Label Label2 
         Caption         =   "CLAVE DE ACCESO:"
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
         Height          =   345
         Left            =   1620
         TabIndex        =   11
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   $"frmCanTraspasoCSN.frx":1084
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
         Height          =   765
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   6915
      End
   End
End
Attribute VB_Name = "FrmCanTraspasoCSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


Dim sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim ContSubgrup As Integer

Dim Cont5 As Integer
Dim Cont6 As Integer
Dim Cont7 As Integer



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
    
    Unload Me
    
End Sub


Private Sub cmdAceptar_Click()
Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim Tipo As String
Dim Cont As Integer

    On Error GoTo eErrorCarga

    Screen.MousePointer = vbHourglass

    If Not ComprobarFechas(0, 1) Then Exit Sub
    
    Cont = RecuentoRegistros

    If Cont = 0 Then
        MsgBox "No hay datos entre estos límites. Reintroduzca.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    Conn.BeginTrans

    If Cont > 32000 Then Cont = 32000
    Pb1.max = Cont + 1
    Pb1.Visible = True
    Pb1.Value = 0
    Me.Refresh

    
    ActualizarTablas
    
    Screen.MousePointer = vbDefault


eErrorCarga:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la cancelación del traspaso de dosis a CSN. Revise."
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        MsgBox "Proceso Finalizado Correctamente", vbExclamation, "Traspaso a CSN."
        cmdCancelar_Click
    End If
End Sub

Private Sub Form_Load()
Dim ano As Currency
Dim Mes As Currency

    ActivarCLAVE
    
    Mes = Month(Now) - 1
    ano = Year(Now)
    If Mes = 0 Then
        Mes = 12
        ano = Year(Now) - 1
    End If
    
    Text1(0).Text = "01/" & Format(Mes, "00") & "/" & Format(ano, "0000")
    Text1(1).Text = CDate("01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BloqueoManual False, "TRASPASO", "TRASPASO"
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    Select Case Index
       Case 0
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(0).Text <> "" Then
                If IsDate(Text1(0).Text) Then f = Text1(0).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(0).Text = frmC.fecha
            mTag.DarFormato Text1(0)
            Set frmC = Nothing
       Case 1
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(1).Text <> "" Then
                If IsDate(Text1(1).Text) Then f = Text1(1).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(1).Text = frmC.fecha
            mTag.DarFormato Text1(1)
            Set frmC = Nothing
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
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
    Dim I As Integer
    Dim sql As String
    Dim mTag As CTag
    Dim valor As Currency
    Dim nomFich As String
    
    ''Quitamos blancos por los lados
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 3
'            ' No dejamos introducir comillas en ningun campo tipo texto
'            If InStr(1, Text1(Index).Text, "'") > 0 Then
'                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation
'                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
'                PonerFoco Text1(Index)
'                Exit Sub
'            End If
'            Text1(Index).Text = Format(Text1(Index).Text, ">")
            
            If Index = 3 Then
                If Trim(Text1(3).Text) <> Trim(Text1(3).Tag) Then
                    MsgBox "    Acceso denegado    ", vbExclamation
                    Text1(3).Text = ""
                    PonerFoco Text1(3)
                Else
                        DesactivarCLAVE
                        PonerFoco Text1(0)
                End If
            End If
            
        Case 0, 1
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
              
              
              'cargamos el campo del fichero a generar
              
              If Text1(0).Text <> "" And Text1(1).Text <> "" Then
              
                If Month(CDate(Text1(0).Text)) = Month(CDate(Text1(1).Text)) And _
                   Year(CDate(Text1(0).Text)) = Year(CDate(Text1(1).Text)) Then
                   
                Else
                
                
                End If
              End If
            End If
            
            
    End Select
    
    '---
End Sub

Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text1(Indice1).Text <> "" And Text1(Indice2).Text <> "" Then
        If CDate(Text1(Indice1).Text) > CDate(Text1(Indice2).Text) Then
            MsgBox "Fecha desde mayor que fecha hasta", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function

Private Sub ActivarCLAVE()
Dim I As Integer
    
    Text1(0).Enabled = False
    Text1(1).Enabled = False

    Text1(3).Enabled = True
    
    Imgppal(0).Enabled = False
    Imgppal(1).Enabled = False
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = True

End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer

    Text1(0).Enabled = True
    Text1(1).Enabled = True

    Text1(3).Text = False
    
    Imgppal(0).Enabled = True
    Imgppal(1).Enabled = True
    
    cmdAceptar.Enabled = True
End Sub


Private Function RecuentoRegistros() As Integer
Dim num As Integer
Dim numTotal As Integer
Dim sql As String
Dim Rs As ADODB.Recordset

    numTotal = 0
    
    Set Rs = New ADODB.Recordset
    
        
    'instalaciones de alta
    sql = "select count(*) from instalaciones "
    sql = sql & "where f_alta >= '" & Format(Text1(0).Text, FormatoFecha)
    sql = sql & "' and f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    
    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num
    
    'instalaciones de baja
    sql = "select count(*) from instalaciones "
    sql = sql & " where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    
    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num
    
    'empresas de alta
    sql = "select count(*) from empresas "
    sql = sql & " where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    
    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num

    'empresas de baja
    sql = "select count(*) from empresas "
    sql = sql & " where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    
    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num
    
    'alta de operarios
    sql = "select count(*) from operarios "
    sql = sql & " where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and "
    sql = sql & " semigracsn = 1 "
    
    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num

    'baja  de operarios
    sql = "select count(*) from operarios "
    sql = sql & " where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and "
    sql = sql & " semigracsn = 1 "
    
    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num
    

    'total de dosis cuerpo
    sql = "select count(*) from dosiscuerpo "
    sql = sql & " where f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "'"

    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num
    
    'total de dosis organo
    sql = "select count(*) from dosisnohomog "
    sql = sql & " where f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    
    Rs.Open sql, Conn, , , adCmdText
    num = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            num = Rs.Fields(0).Value
        End If
    End If
    Rs.Close
    
    numTotal = numTotal + num
    
    RecuentoRegistros = numTotal

End Function

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub ActualizarTablas()
Dim sql As String

    'dosis organo
    sql = "update dosisnohomog set migrado = NULL where f_dosis >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_dosis <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "'"
    
    Conn.Execute sql
    
    'dosis homogeneas
    sql = "update dosiscuerpo set migrado = null where f_dosis >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_dosis <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "'"
    
    Conn.Execute sql
    
    'Baja de Operarios
    sql = "update operarios set migrado = '*' where f_baja >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_baja <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "' and semigracsn = 1 "
'    Sql = Sql & " and operarios.dni<>'0' and operarios.dni<>'999999999' and operarios.dni<>'888888888' "
'    Sql = Sql & " and operarios.dni<>'999999998' and operarios.dni<>'999999997' "
'    Sql = Sql & " and operarios.dni<>'666666666' and operarios.dni<>'777777777' "
'    Sql = Sql & " and operarios.dni<>'999999996' "
    
    
    Conn.Execute sql
    
    'Alta de Operarios
    sql = "update operarios set migrado = NULL where f_alta >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_alta <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "' and semigracsn = 1 "
'    Sql = Sql & " and operarios.dni<>'0' and operarios.dni<>'999999999' and operarios.dni<>'888888888' "
'    Sql = Sql & " and operarios.dni<>'999999998' and operarios.dni<>'999999997' "
'    Sql = Sql & " and operarios.dni<>'666666666' and operarios.dni<>'777777777' "
'    Sql = Sql & " and operarios.dni<>'999999996' "
    
    Conn.Execute sql
    
    'Baja de Instalaciones
    sql = "update instalaciones set migrado = '*' where f_baja >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_baja <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "'"
    sql = sql & " and (c_tipo = 0 or c_tipo = 2) "
    
    Conn.Execute sql
    
    'Alta de Instalaciones
    sql = "update instalaciones set migrado = null where f_alta >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_alta <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "'"
    sql = sql & " and (c_tipo = 0 or c_tipo = 2) "
    
    Conn.Execute sql
    
    'Baja de Empresas
    sql = "update empresas set migrado = '*' where f_baja >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_baja <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "'"
    sql = sql & " and (c_tipo = 0 or c_tipo = 2)  "
    Conn.Execute sql
    
    'Alta de Empresas
    sql = "update empresas set migrado = null where f_alta >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_alta <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "'"
    sql = sql & " and (c_tipo = 0 or c_tipo = 2) "
    
    Conn.Execute sql
    
End Sub
