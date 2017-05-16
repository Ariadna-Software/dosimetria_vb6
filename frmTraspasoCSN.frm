VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmTraspasoCSN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso Automático de datos a C.S.N"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmTraspasoCSN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListTipoMedicion 
      Height          =   5910
      Left            =   30
      TabIndex        =   6
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
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Frame Frame2 
         Caption         =   "Archivo Resultante"
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
         Height          =   675
         Left            =   450
         TabIndex        =   12
         Top             =   3540
         Width           =   6225
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1380
            TabIndex        =   3
            Text            =   "Text5"
            Top             =   270
            Width           =   4215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Período a migrar"
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
         TabIndex        =   9
         Top             =   2430
         Width           =   6195
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   2
            Text            =   "Text5"
            Top             =   390
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
            Left            =   3990
            MouseIcon       =   "frmTraspasoCSN.frx":0CCA
            MousePointer    =   99  'Custom
            Picture         =   "frmTraspasoCSN.frx":0E1C
            ToolTipText     =   "Seleccionar fecha"
            Top             =   390
            Width           =   240
         End
         Begin VB.Image ImgPpal 
            Height          =   240
            Index           =   0
            Left            =   1080
            MouseIcon       =   "frmTraspasoCSN.frx":0EA7
            MousePointer    =   99  'Custom
            Picture         =   "frmTraspasoCSN.frx":0FF9
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
            TabIndex        =   11
            Top             =   420
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
            TabIndex        =   10
            Top             =   420
            Width           =   735
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3690
         TabIndex        =   5
         Top             =   4890
         Width           =   1425
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   2040
         TabIndex        =   4
         Top             =   4890
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   300
         Left            =   420
         TabIndex        =   8
         Top             =   4410
         Visible         =   0   'False
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         Caption         =   "Introduzca la clave de acceso y pulse intro"
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
         Height          =   315
         Left            =   1590
         TabIndex        =   14
         Top             =   1590
         Width           =   3825
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
         TabIndex        =   13
         Top             =   1950
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   $"frmTraspasoCSN.frx":1084
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
         Height          =   1215
         Left            =   570
         TabIndex        =   7
         Top             =   330
         Width           =   5805
      End
   End
End
Attribute VB_Name = "FrmTraspasoCSN"
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
Dim cont As Integer

    On Error GoTo eErrorCarga

    Screen.MousePointer = vbHourglass

    If Not ComprobarFechas(0, 1) Then Exit Sub
    
    cont = RecuentoRegistros

    If cont = 0 Then
        MsgBox "No hay datos entre estos límites. Reintroduzca.", vbExclamation, "¡Error!"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    CargarOperarios

    InsertarRegistros05
    
    If cont > 32000 Then cont = 32000
    PB1.max = cont + 1
    PB1.Visible = True
    PB1.Value = 0
    Me.Refresh

    
    GeneraFichero
    
    
    Screen.MousePointer = vbDefault


eErrorCarga:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la generación del fichero de dosis a CSN. Revise."
    End If
End Sub

Private Sub Form_Load()
Dim ano As Currency
Dim Mes As Currency

    ActivarCLAVE
    Text1(2).Text = Trim(App.Path & "TRASPASOCSN\file")
    
    Mes = Month(Now) - 1
    ano = Year(Now)
    If Mes = 0 Then
        Mes = 12
        ano = Year(Now) - 1
    End If
    
    Text1(0).Text = "01/" & Format(Mes, "00") & "/" & Format(ano, "0000")
    Text1(1).Text = CDate("01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")) - 1
    Text1(2).Text = App.Path & "\TRASPASOCSN\INF" & Format(ano, "0000") & Format(Mes, "00")
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
        Case 2
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, ">")
            
            If Dir(Trim(Text1(2).Text), vbArchive) <> "" Then
               If MsgBox("Este fichero ya existe. Desea reemplazarlo.", vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") = vbNo Then
                 PonerFoco Text1(2)
               Else
                 PonerFoco cmdAceptar
               End If
            End If
        
        Case 3
            If Trim(Text1(3).Text) <> Trim(Text1(3).Tag) Then
              MsgBox "    Acceso denegado    ", vbExclamation, "¡Atención!"
              Text1(3).Text = ""
              PonerFoco Text1(3)
            Else
              DesactivarCLAVE
              PonerFoco Text1(0)
            End If
            
            
        Case 0, 1
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation, "¡Error!"
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
              
              
              'cargamos el campo del fichero a generar
              
              If Text1(0).Text <> "" And Text1(1).Text <> "" Then
              
                If Month(CDate(Text1(0).Text)) = Month(CDate(Text1(1).Text)) And _
                   Year(CDate(Text1(0).Text)) = Year(CDate(Text1(1).Text)) Then
                   
                    If Dir(App.Path & "\TRASPASOCSN", vbDirectory) = "" Then
                        MkDir App.Path & "\TRASPASOCSN"
                    End If
                   
                    nomFich = "INF" & Format(Year(CDate(Text1(0).Text)), "0000") & Format(Month(CDate(Text1(0).Text)), "00")
                    Text1(2).Text = App.Path & "\TRASPASOCSN\" & Trim(nomFich)

'                    If Dir(App.Path & "\TRASPASOCSN\" & Trim(nomFich), vbArchive) <> "" Then
'                        If MsgBox("Este fichero ya existe. Desea reemplazarlo.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                            cmdCancelar_Click
'                        Else
'                            Text1(2).Text = App.Path & "\TRASPASOCSN\" & Trim(nomFich)
'                        End If
'                    End If
                Else
                    If Dir(App.Path & "\TRASPASOCSN", vbDirectory) = "" Then
                        MkDir App.Path & "\TRASPASOCSN"
                    End If
                
                    Text1(2).Text = App.Path & "\TRASPASOCSN\file"
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
            MsgBox "Fecha 'desde' mayor que fecha 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function

Private Sub ActivarCLAVE()
Dim I As Integer
    
    For I = 0 To Text1.Count - 1
        Text1(I).Enabled = False
    Next I

    Imgppal(0).Enabled = False
    Imgppal(1).Enabled = False

    Text1(3).Enabled = True
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = True

End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer

    For I = 0 To Text1.Count - 1
        Text1(I).Enabled = True
    Next I

    Imgppal(0).Enabled = True
    Imgppal(1).Enabled = True

    Text1(3).Text = False
    
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
    sql = "select count(distinct c_instalacion) from instalaciones "
    sql = sql & "where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado is null and "
    sql = sql & " (c_tipo = 0 or c_tipo = 2) "
    
    
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
    sql = "select count(distinct c_instalacion) from instalaciones "
    sql = sql & " where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado = '*' and "
    sql = sql & " (c_tipo = 0 or c_tipo = 2) "
    
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
    sql = "select count(distinct c_empresa)  from empresas "
    sql = sql & " where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado is null and "
    sql = sql & " (c_tipo = 0 or c_tipo = 2) "
    
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
    sql = "select count(distinct c_empresa) from empresas "
    sql = sql & " where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado = '*' and "
    sql = sql & " (c_tipo  = 0 or c_tipo = 2) "
    
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
    sql = "select count(distinct dni) from operarios "
    sql = sql & " where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado is null "
    sql = sql & " and semigracsn = 1 "

    ' 28/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and dni<>'0' and dni<>'999999999' and dni<>'888888888' "
    sql = sql & " and dni<>'999999998' and dni<>'999999997' "
    sql = sql & " and dni<>'666666666' and dni<>'777777777' "
    sql = sql & " and dni<>'999999996' "
    ' 28/02/2006 [DV] Hasta aquí
    
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
    sql = "select count(distinct dni) from operarios "
    sql = sql & " where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado = '*' "
    sql = sql & " and semigracsn = 1 "
    
    ' 28/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and dni<>'0' and dni<>'999999999' and dni<>'888888888' "
    sql = sql & " and dni<>'999999998' and dni<>'999999997' "
    sql = sql & " and dni<>'666666666' and dni<>'777777777' "
    sql = sql & " and dni<>'999999996' "
    ' 28/02/2006 [DV] Hasta aquí
    
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
    sql = "select count(distinct n_registro) from dosiscuerpo, operarios "
    sql = sql & " where f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "' and dosiscuerpo.migrado is null "
    sql = sql & " and  operarios.semigracsn = 1 and operarios.dni = dosiscuerpo.dni_usuario "
    
    
    ' 28/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and dni_usuario<>'0' and dni_usuario <>'999999999' and dni_usuario<>'888888888' "
    sql = sql & " and dni_usuario<>'999999998' and dni_usuario <>'999999997' "
    sql = sql & " and dni_usuario<>'666666666' and dni_usuario <>'777777777' "
    sql = sql & " and dni_usuario<>'999999996' "
    ' 28/02/2006 [DV] Hasta aquí
    
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
    sql = "select count(distinct n_registro) from dosisnohomog, operarios "
    sql = sql & " where f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "' and dosisnohomog.migrado is null and "
    sql = sql & " operarios.semigracsn = 1 and operarios.dni = dosisnohomog.dni_usuario "
    
    ' 28/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and dni_usuario<>'0' and dni_usuario <>'999999999' and dni_usuario<>'888888888' "
    sql = sql & " and dni_usuario<>'999999998' and dni_usuario <>'999999997' "
    sql = sql & " and dni_usuario<>'666666666' and dni_usuario <>'777777777' "
    sql = sql & " and dni_usuario<>'999999996' "
    ' 28/02/2006 [DV] Hasta aquí

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


Private Sub GeneraFichero()
Dim nFich As Integer
Dim Regs As Integer
Dim Importe As Currency
Dim Rs As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim Aux As String
Dim Cad As String
Dim sql As String
Dim sql1 As String

'contadores de totales
Dim ContAltaEmp As Integer
Dim ContBajaEmp As Integer
Dim ContAltaIns As Integer
Dim ContBajaIns As Integer
Dim ContAltaOpe As Integer
Dim ContBajaOpe As Integer
    
    On Error GoTo eGeneraFichero

    Conn.BeginTrans
    
    ContSubgrup = 0
    Cont5 = 0
    Cont6 = 0
    Cont7 = 0
    
    nFich = FreeFile
    Open Trim(Text1(2).Text) For Output As #nFich
    
    Cabecera1 nFich, ""
    
    'alta instalaciones
    Set Rs = New ADODB.Recordset
    
    sql = "select c_instalacion, f_alta, descripcion, direccion, poblacion,"
    sql = sql & "c_postal, distrito, telefono, fax, migrado, rama_gen, rama_especifica from instalaciones "
    sql = sql & "where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and "
    sql = sql & " (c_tipo = 0 or c_tipo = 2) and migrado is null "
    
    ContAltaIns = 0
    
    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            Linea1 nFich, Rs, Cad
        
            ContAltaIns = ContAltaIns + 1
            
            PB1.Value = PB1.Value + 1
            PB1.Refresh
        
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    
    If ContAltaIns <> 0 Then LineaTotales1 nFich, ContAltaIns
    
    
    'baja instalaciones
    sql = "select c_instalacion, descripcion from instalaciones "
    sql = sql & "where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and "
    sql = sql & " (c_tipo = 0 or c_tipo = 2) and migrado = '*' "
    
    ContBajaIns = 0
    
    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            Linea2 nFich, Rs, Cad
            
            ContBajaIns = ContBajaIns + 1
        
            PB1.Value = PB1.Value + 1
            PB1.Refresh
            
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    
    If ContBajaIns <> 0 Then LineaTotales2 nFich, ContBajaIns
    
    
    ' alta de empresas
    sql = "select c_empresa,f_alta,cif_nif, nom_comercial, direccion, "
    sql = sql & "c_postal, distrito, poblacion, tel_contacto, fax, migrado from empresas "
    sql = sql & "where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and "
    sql = sql & " (c_tipo = 0 or c_tipo = 2) and migrado is null "
    
    ContAltaEmp = 0
    
    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            Linea3 nFich, Rs, Cad
            
            ContAltaEmp = ContAltaEmp + 1
            
            PB1.Value = PB1.Value + 1
            PB1.Refresh
        
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    
    If ContAltaEmp <> 0 Then LineaTotales3 nFich, ContAltaEmp
    
        
    'baja empresas
    sql = "select cif_nif, nom_comercial from empresas "
    sql = sql & "where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and "
    sql = sql & " (c_tipo = 0 or c_tipo = 2) and migrado = '*' "
    
    ContBajaEmp = 0
    
    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            Linea4 nFich, Rs, Cad
            
            ContBajaEmp = ContBajaEmp + 1
        
            PB1.Value = PB1.Value + 1
            PB1.Refresh
        
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    
    If ContBajaEmp <> 0 Then LineaTotales4 nFich, ContBajaEmp
    
    
    'alta operarios
    sql = "select operarios.dni, operarios.apellido_1,operarios.apellido_2,operarios.nombre,"
    sql = sql & "operarios.f_nacimiento, operarios.f_alta, "
    sql = sql & "operarios.sexo, operarios.n_carnet_radiolog,operarios.f_emi_carnet_rad "
    sql = sql & " from operarios " ', operainstala "
    sql = sql & " where operarios.f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " operarios.f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and operarios.migrado is null and "
    sql = sql & " operarios.semigracsn = 1 "
    
    ' 28/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and operarios.dni<>'0' and operarios.dni<>'999999999' and operarios.dni<>'888888888' "
    sql = sql & " and operarios.dni<>'999999998' and operarios.dni<>'999999997' "
    sql = sql & " and operarios.dni<>'666666666' and operarios.dni<>'777777777' "
    sql = sql & " and operarios.dni<>'999999996' "
    ' 28/02/2006 [DV] Hasta aquí
    
    ContAltaOpe = 0
    
    
    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            Linea5 nFich, Rs, Cad
            
            ContAltaOpe = ContAltaOpe + 1
        
            PB1.Value = PB1.Value + 1
            PB1.Refresh
        
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    
    If ContAltaOpe <> 0 Then LineaTotales5 nFich, ContAltaOpe
    
    
    'baja operarios
    sql = "select operarios.dni, operarios.apellido_1,operarios.apellido_2,operarios.nombre from operarios " '", operainstala"
    sql = sql & " where operarios.f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " operarios.f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and operarios.migrado = '*' "
    sql = sql & " and operarios.semigracsn = 1 "
    
    ' 28/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and operarios.dni<>'0' and operarios.dni<>'999999999' and operarios.dni<>'888888888' "
    sql = sql & " and operarios.dni<>'999999998' and operarios.dni<>'999999997' "
    sql = sql & " and operarios.dni<>'666666666' and operarios.dni<>'777777777' "
    sql = sql & " and operarios.dni<>'999999996' "
    ' 28/02/2006 [DV] Hasta aquí
    
    ContBajaOpe = 0
    
    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            Linea6 nFich, Rs, Cad
            
            ContBajaOpe = ContBajaOpe + 1
            
            PB1.Value = PB1.Value + 1
            PB1.Refresh
            
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    
    If ContBajaOpe <> 0 Then LineaTotales6 nFich, ContBajaOpe
    
    
    'dosis cuerpo
    sql = "select dosiscuerpo.c_tipo_trabajo,dosiscuerpo.c_instalacion,dosiscuerpo.dni_usuario,dosiscuerpo.c_empresa,dosiscuerpo.dosis_superf,"
    sql = sql & "dosiscuerpo.dosis_profunda,dosiscuerpo.plantilla_contrata,dosiscuerpo.observaciones from dosiscuerpo " ', voperarios " ', dosimetros "
    sql = sql & " where f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "' and dosiscuerpo.migrado is null "

' 07/03/2007 [DV] Evitemos los cruces con otras tablas (una vez más).
'    sql = sql & " and voperarios.codusu = " & vUsu.codigo
'    sql = sql & " and voperarios.semigracsn = 1 and voperarios.dni = dosiscuerpo.dni_usuario "
' 07/03/2007 [DV] Hasta aquí

'    sql = sql & " and dosiscuerpo.n_reg_dosimetro = dosimetros.n_reg_dosimetro  and dosimetros.tipo_dosimetro = 0 "
    
' añadido el enlace entre dosimetros y dosis en lugar de estas lineas
    
    ' 28/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and dni_usuario<>'0' and dni_usuario<>'999999999' and dni_usuario<>'888888888' "
    sql = sql & " and dni_usuario<>'999999998' and dni_usuario<>'999999997' "
    sql = sql & " and dni_usuario<>'666666666' and dni_usuario<>'777777777' "
    sql = sql & " and dni_usuario<>'999999996' "
    ' 28/02/2006 [DV] Hasta aquí
    
    sql = sql & " order by dosiscuerpo.dni_usuario, dosiscuerpo.c_instalacion "

    Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
        
            Linea7 nFich, Rs, Cad
        
            PB1.Value = PB1.Value + 1
            PB1.Refresh
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    
    LineaTotales7 nFich
    
    Regs = ContAltaEmp + ContBajaEmp + ContAltaIns + ContBajaIns + _
           ContAltaOpe + ContBajaOpe + Cont5 + Cont6 + Cont7 + ContSubgrup
    
    LineaTotales8 nFich, Regs
    
    
    Set Rs = Nothing

eGeneraFichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error generando fichero."
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        MsgBox "Proceso Finalizado Correctamente", vbExclamation, "Generación del Fichero."
        cmdCancelar_Click
    End If
    Close (nFich)


End Sub


Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaABlancos = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaABlancos = Right(Cad, Longitud)
    End If
    
End Function

Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaAceros = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaAceros = Right(Cad, Longitud)
    End If
    
End Function

Private Sub Cabecera1(NF As Integer, Cad As String)
Dim ano As Currency
Dim Mes As Currency

    If Year(CDate(Text1(0).Text)) >= 2000 Then
        ano = Year(CDate(Text1(0).Text)) - 2000
    Else
        ano = Year(CDate(Text1(0).Text)) - 1900
    End If
    
    Mes = Month(CDate(Text1(0).Text))

    '-- Registro cabecera
    Cad = ""
    Cad = "0120  "
    Cad = Cad & Mid(CStr(Format(Year(Now), "0000")), 3, 2) & Format(Month(Now), "00") & Format(Day(Now), "00") ' Fecha de presentación
    Cad = Cad & Format(ano, "00")
    Cad = Cad & Format(Mes, "00")
    
    Print #NF, Cad
    
    ContSubgrup = ContSubgrup + 1
    
End Sub

Private Sub Linea1(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim cad1 As String
Dim cad2 As String
Dim I As Integer

    cad2 = Format(CInt(Val(Trim(Rs1!c_postal & ""))), "00") & Format(CInt(Val(Trim(Rs1!distrito & ""))), "000")
    
    Cad = "02"
    
    cad1 = ""
    If Not IsNull(Rs1!descripcion) Then cad1 = Rs1!descripcion
    Cad = Cad & RellenaABlancos(cad1, True, 60)
    
    cad1 = ""
    If Not IsNull(Rs1!direccion) Then cad1 = Rs1!direccion
    Cad = Cad & RellenaABlancos(cad1, True, 50)
    
    cad1 = ""
    If Not IsNull(cad2) Then cad1 = cad2
    Cad = Cad & RellenaABlancos(cad1, True, 5)
    
    cad1 = ""
    If Not IsNull(Rs1!poblacion) Then cad1 = Rs1!poblacion
    Cad = Cad & RellenaABlancos(cad1, True, 40)
    
    cad1 = ""
    If Not IsNull(Rs1!telefono) Then cad1 = Rs1!telefono
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!fax) Then cad1 = Rs1!fax
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    I = 0
    If Not IsNull(Rs1!rama_gen) Then I = CStr(Rs1!rama_gen)
    Cad = Cad & Format(I, "00")
    
    I = 0
    If Not IsNull(Rs1!rama_especifica) Then I = CStr(Rs1!rama_especifica)
    Cad = Cad & Format(I, "00")
    
    
'    Cad = Cad & RellenaABlancos(RS1!direccion, True, 50)
'    Cad = Cad & RellenaABlancos(cad1, True, 5)
'    Cad = Cad & RellenaABlancos(RS1!poblacion, True, 40)
'    Cad = Cad & RellenaABlancos(RS1!telefono, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!fax, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!rama_gen, True, 2)
'    Cad = Cad & RellenaABlancos(RS1!rama_especifica, True, 2)
    
    Print #NF, Cad

End Sub

Private Sub LineaTotales1(NF As Integer, total As Integer)
Dim sql As String
Dim Cad As String

    Cad = "82" & Format(total, "000000")
    
    ContSubgrup = ContSubgrup + 1
    
    Print #NF, Cad
    
    ' actualizamos las instalaciones de alta migrados
    sql = "update instalaciones set migrado = '*' where f_alta >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_alta <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "' and migrado is null and (c_tipo = 0 or c_tipo = 2)"
    
    Conn.Execute sql
    
End Sub

Private Sub Linea2(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim cad1 As String
    
    Cad = "12"
    
    cad1 = ""
    If Not IsNull(Rs1!c_instalacion) Then cad1 = Rs1!c_instalacion
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!descripcion) Then cad1 = Rs1!descripcion
    Cad = Cad & RellenaABlancos(cad1, True, 60)
    
'    Cad = Cad & RellenaABlancos(RS1!c_instalacion, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!descripcion, True, 60)
    
    Print #NF, Cad

End Sub

Private Sub LineaTotales2(NF As Integer, total As Integer)
Dim Cad As String
Dim sql As String

    Cad = "92" & Format(total, "000000")
    
    ContSubgrup = ContSubgrup + 1
    
    Print #NF, Cad

    ' actualizamos las instalaciones de baja migrados
    sql = "update instalaciones set migrado = '*' where f_baja >= '"
    sql = sql & Format(Text1(0).Text, FormatoFecha) & "' and f_baja <= '"
    sql = sql & Format(Text1(1).Text, FormatoFecha) & "' and migrado = '*' and (c_tipo = 0 or c_tipo = 2)"
    
    Conn.Execute sql

End Sub


Private Sub Linea3(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim cad1 As String
Dim cad2 As String

    cad2 = Format(CInt(Val(Trim(Rs1!c_postal & ""))), "00") & Format(CInt(Val(Trim(Rs1!distrito & ""))), "000")
    
    Cad = "03"
    
    cad1 = ""
    If Not IsNull(Rs1!cif_nif) Then cad1 = Rs1!cif_nif
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!nom_comercial) Then cad1 = Rs1!nom_comercial
    Cad = Cad & RellenaABlancos(cad1, True, 50)
    
    cad1 = ""
    If Not IsNull(Rs1!direccion) Then cad1 = Rs1!direccion
    Cad = Cad & RellenaABlancos(cad1, True, 50)
    
    cad1 = ""
    If Not IsNull(cad2) Then cad1 = cad2
    Cad = Cad & RellenaABlancos(cad1, True, 5)
    
    cad1 = ""
    If Not IsNull(Rs1!poblacion) Then cad1 = Rs1!poblacion
    Cad = Cad & RellenaABlancos(cad1, True, 40)
    
    cad1 = ""
    If Not IsNull(Rs1!tel_contacto) Then cad1 = Rs1!tel_contacto
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!fax) Then cad1 = Rs1!fax
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
'    Cad = Cad & RellenaABlancos(RS1!cif_nif, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!nom_comercial, True, 40)
'    Cad = Cad & RellenaABlancos(RS1!direccion, True, 50)
'    Cad = Cad & RellenaABlancos(cad1, True, 5)
'    Cad = Cad & RellenaABlancos(RS1!poblacion, True, 40)
'    Cad = Cad & RellenaABlancos(RS1!telefono, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!fax, True, 11)
    
    Print #NF, Cad


End Sub

Private Sub LineaTotales3(NF As Integer, total As Integer)
Dim sql As String
Dim Cad As String

    Cad = "83" & Format(total, "000000")
    
    ContSubgrup = ContSubgrup + 1
    
    Print #NF, Cad

    ' actualizamos las empresas de alta migradas
    sql = "update empresas set migrado = '*' where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado is null and (c_tipo = 0 or c_tipo = 2)"
    
    Conn.Execute sql

End Sub


Private Sub Linea4(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim cad1 As String
    
    Cad = "13"
    
    cad1 = ""
    If Not IsNull(Rs1!cif_nif) Then cad1 = Rs1!cif_nif
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!nom_comercial) Then cad1 = Rs1!nom_comercial
    Cad = Cad & RellenaABlancos(cad1, True, 50)
    
    
'    Cad = Cad & RellenaABlancos(RS1!cif_nif, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!nom_comercial, True, 40)
    
    Print #NF, Cad

End Sub

Private Sub LineaTotales4(NF As Integer, total As Integer)
Dim Cad As String
Dim sql As String

    Cad = "93" & Format(total, "000000")
    
    ContSubgrup = ContSubgrup + 1
    
    Print #NF, Cad

    ' actualizamos las empresas de baja migradas
    sql = "update empresas set migrado = '*' where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado = '*' and (c_tipo = 0 or c_tipo = 2)"
    
    Conn.Execute sql

End Sub


Private Sub Linea5(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim ano As Currency
Dim Mes As Currency
Dim cad1 As String
    
    ano = Year(Rs1!f_alta)
    If ano >= 2000 Then
        ano = ano - 2000
    Else
        ano = ano - 1900
    End If
    
    Mes = Month(Rs1!f_alta)
    
    Cad = "04"
    
    cad1 = ""
    If Not IsNull(Rs1!dni) Then cad1 = Rs1!dni
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!apellido_1) Then cad1 = Rs1!apellido_1
    Cad = Cad & RellenaABlancos(cad1, True, 25)
    
    cad1 = ""
    If Not IsNull(Rs1!apellido_2) Then cad1 = Rs1!apellido_2
    Cad = Cad & RellenaABlancos(cad1, True, 25)
    
    cad1 = ""
    If Not IsNull(Rs1!nombre) Then cad1 = Rs1!nombre
    Cad = Cad & RellenaABlancos(cad1, True, 20)
    
    cad1 = ""
    If Not IsNull(Rs1!f_nacimiento) Then cad1 = Format(Rs1!f_nacimiento, "YYMMDD")
    Cad = Cad & RellenaABlancos(cad1, True, 6)
    
    cad1 = ""
    If Not IsNull(Rs1!sexo) Then cad1 = Rs1!sexo
    Cad = Cad & RellenaABlancos(cad1, True, 1)
    
    cad1 = ""
    If Not IsNull(Rs1!n_carnet_radiolog) Then cad1 = Rs1!n_carnet_radiolog
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!f_emi_carnet_rad) Then cad1 = Format(Rs1!f_emi_carnet_rad, "YYMMDD")
    Cad = Cad & RellenaABlancos(cad1, True, 6)
    
    
'    Cad = Cad & RellenaABlancos(RS1!dni, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!apellido_1, True, 25)
'    Cad = Cad & RellenaABlancos(RS1!apellido_2, True, 25)
'    Cad = Cad & RellenaABlancos(RS1!nombre, True, 20)
'    Cad = Cad & Format(RS1!f_nacimiento, "YYMMDD")
'    Cad = Cad & RellenaABlancos(RS1!sexo, True, 1)
'    Cad = Cad & RellenaABlancos(RS1!n_carnet_radiologico, True, 11)

    Cad = Cad & "00000000" & "000000000" & "000000000"
    Cad = Cad & Format(ano, "00")
    Cad = Cad & Format(Mes, "00")
    
    Print #NF, Cad


End Sub

Private Sub LineaTotales5(NF As Integer, total As Integer)
Dim Cad As String
Dim sql As String

    Cad = "84" & Format(total, "000000")
    
    ContSubgrup = ContSubgrup + 1
    
    Print #NF, Cad

    ' actualizamos los operarios de alta migrados
    sql = "update operarios set migrado = '*' where f_alta >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_alta <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado is null "
    sql = sql & " and semigracsn = 1"
'    Sql = Sql & " and dni<>'0' and dni<>'999999999' and dni<>'888888888' "
'    Sql = Sql & " and dni<>'999999998' and dni<>'999999997' "
'    Sql = Sql & " and dni<>'666666666' and dni<>'777777777' "
'    Sql = Sql & " and dni<>'999999996' "
    
    Conn.Execute sql

End Sub

Private Sub Linea6(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim cad1 As String
    
    Cad = "14"
    
    cad1 = ""
    If Not IsNull(Rs1!dni) Then cad1 = Rs1!dni
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    cad1 = ""
    If Not IsNull(Rs1!apellido_1) Then cad1 = Rs1!apellido_1
    Cad = Cad & RellenaABlancos(cad1, True, 25)
    
    cad1 = ""
    If Not IsNull(Rs1!apellido_2) Then cad1 = Rs1!apellido_2
    Cad = Cad & RellenaABlancos(cad1, True, 25)
    
    cad1 = ""
    If Not IsNull(Rs1!nombre) Then cad1 = Rs1!nombre
    Cad = Cad & RellenaABlancos(cad1, True, 20)
    
'
'    Cad = Cad & RellenaABlancos(RS1!cif_nif, True, 11)
'    Cad = Cad & RellenaABlancos(RS1!apellido_1, True, 25)
'    Cad = Cad & RellenaABlancos(RS1!apellido_2, True, 25)
'    Cad = Cad & RellenaABlancos(RS1!nombre, True, 20)

    Print #NF, Cad

End Sub

Private Sub LineaTotales6(NF As Integer, total As Integer)
Dim Cad As String
Dim sql As String

    Cad = "94" & Format(total, "000000")
    ContSubgrup = ContSubgrup + 1
    Print #NF, Cad

    ' actualizamos los operarios de baja migrados
    sql = "update operarios set migrado = '**' where f_baja >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " f_baja <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado = '*' "
    sql = sql & " and semigracsn = 1 "
    
'    Sql = Sql & " and dni<>'0' and dni<>'999999999' and dni<>'888888888' "
'    Sql = Sql & " and dni<>'999999998' and dni<>'999999997' "
'    Sql = Sql & " and dni<>'666666666' and dni<>'777777777' "
'    Sql = Sql & " and dni<>'999999996' "
    
    Conn.Execute sql


End Sub

Private Sub Linea7(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim sql As String
Dim Instala As String
Dim rL As ADODB.Recordset
Dim I As Currency
Dim cad1 As String

    sql = ""
    sql = DevuelveDesdeBD(1, "cif_nif", "empresas", "c_empresa|", Trim(Rs1!c_empresa) & "|", "T|", 1)
 
    Cad = "05"
    
    cad1 = ""
    If Not IsNull(Rs1!dni_usuario) Then cad1 = Rs1!dni_usuario
    Cad = Cad & RellenaABlancos(cad1, True, 11)
    
    'If Sql = "AA" Then Stop
    
    Cad = Cad & RellenaABlancos(Trim(sql), True, 11)
    
    I = 0
    If Not IsNull(Rs1!dosis_profunda) Then I = Rs1!dosis_profunda
    Cad = Cad & Format((I * 100), "00000000")
    
    I = 0
    If Not IsNull(Rs1!dosis_superf) Then I = Rs1!dosis_superf
    Cad = Cad & Format((I * 100), "000000000")
    
    I = 0
    If Not IsNull(Rs1!plantilla_contrata) Then I = Rs1!plantilla_contrata
    Cad = Cad & Format(I, "00")
    
    cad1 = ""
    If Not IsNull(Rs1!Observaciones) Then cad1 = Rs1!Observaciones
    Cad = Cad & RellenaABlancos(cad1, True, 120)
    
'    Cad = Cad & RellenaABlancos(RS1!dni_usuario, True, 11)
'    Cad = Cad & RellenaABlancos(Trim(sql), True, 11)
'    Cad = Cad & Format((RS1!dosis_profunda * 100), "00000000")
'    Cad = Cad & Format((RS1!dosis_superf * 100), "000000000")
'    Cad = Cad & Format(CInt(RS1!plantilla_contrata), "00")
'    If Not IsNull(RS1!Observaciones) Then
'        Cad = Cad & RellenaABlancos(RS1!Observaciones, True, 120)
'    End If
    
    Cont5 = Cont5 + 1
    
    Print #NF, Cad

    Instala = Trim(Rs1!c_instalacion)
    
    sql = ""
    sql = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_instalacion|", Instala & "|", "T|", 1)
    
    If Mid(Instala, 1, 1) = "Z" Then Instala = "           " '11 blancos
    
    Cad = "06"
    Cad = Cad & RellenaABlancos(Instala, True, 11)
    Cad = Cad & RellenaABlancos(Trim(sql), True, 60)
    
    Instala = Trim(Rs1!c_instalacion)
    
    
    I = 0
    If Not IsNull(Rs1!c_tipo_trabajo) Then I = Rs1!c_tipo_trabajo
    Cad = Cad & Format(I, "00")
    
'    Cad = Cad & Format(CInt(RS1!c_tipo_trabajo), "00")
    
    Print #NF, Cad

    Cont6 = Cont6 + 1

    ' cursor para las dosis a organo que encontremos de esta instalacion

    sql = "select dosisnohomog.dosis_org, dosisnohomog.tipo_medicion, dosisnohomog.observaciones from dosisnohomog, voperarios " ', dosimetros "
    sql = sql & " where dosisnohomog.dni_usuario = '" & Trim(Rs1!dni_usuario) & "' and "
    sql = sql & " dosisnohomog.c_instalacion = '" & Instala & "' and "
    sql = sql & " voperarios.codusu = " & vUsu.codigo & " and "
    sql = sql & " dosisnohomog.f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " dosisnohomog.f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "' and dosisnohomog.migrado is null and "
    sql = sql & " voperarios.semigracsn = 1 and dosisnohomog.dni_usuario = voperarios.dni "
    
'    sql = sql & " dosimetros.n_reg_dosimetro = dosisnohomog.n_reg_dosimetro and "
    ' 27/02/2006 [DV] Modificación referente a fallos en el envío CSN
    sql = sql & " and dni_usuario<>'0' and dni_usuario<>'999999999' and dni_usuario<>'888888888' "
    sql = sql & " and dni_usuario<>'999999998' and dni_usuario<>'999999997' "
    sql = sql & " and dni_usuario<>'666666666' and dni_usuario<>'777777777' "
    sql = sql & " and dni_usuario<>'999999996' "
    ' 27/02/2006 [DV] Hasta aquí
    
'añadido el enlace entre dosimetros y dosisnohomog en lugar de las lineas siguientes
'    sql = sql & " dosimetros.tipo_dosimetro = 1 "
    
    
    Set rL = New ADODB.Recordset
    rL.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Not rL.EOF Then rL.MoveFirst
    While Not rL.EOF
        Linea8 NF, rL, Cad
        
        Cont7 = Cont7 + 1
    
        rL.MoveNext
    Wend

    Set rL = Nothing
    
End Sub

Private Sub Linea8(NF As Integer, ByRef Rs1 As ADODB.Recordset, ByRef Cad As String)
Dim mAux As String
Dim I As Currency
Dim cad1 As String

    Cad = "07"
    
    I = 0
    If Not IsNull(Rs1!dosis_org) Then I = Rs1!dosis_org
    Cad = Cad & Format(I * 100, "000000000")
    
    I = 0
    If Not IsNull(Rs1!tipo_medicion) Then I = Rs1!tipo_medicion
    Cad = Cad & Format(I, "00")
    
    cad1 = ""
    If Not IsNull(Rs1!Observaciones) Then cad1 = Rs1!Observaciones
    Cad = Cad & RellenaABlancos(cad1, True, 60)
    
'    Cad = Cad & Format(RS1!dosis_org * 100, "000000000")
'    Cad = Cad & Format(CInt(RS1!tipo_medicion), "00")
'    If Not IsNull(RS1!Observaciones) Then
'        Cad = Cad & RellenaABlancos(RS1!Observaciones, True, 60)
'    End If
    
    Print #NF, Cad

End Sub

Private Sub LineaTotales7(NF As Integer)
Dim Cad As String

    Cad = "95" & Format((Cont5 + Cont6 + Cont7), "000000")
    
    ContSubgrup = ContSubgrup + 1
    
    Print #NF, Cad
End Sub

Private Sub LineaTotales8(NF As Integer, total As Integer)
Dim sql As String
Dim sql1 As String
Dim Cad As String

    Cad = "99" & Format(Cont5, "0000000") & Format(total, "0000000")
    
    Print #NF, Cad

    
    ' NOTA : Marco absolutamente todas las dosis sean o no migradas: semigracsn =1 ó 0
    
    ' actualizamos las tablas de dosis cuerpo y organo
    sql = "update dosiscuerpo"
    
    sql1 = " set migrado = '**' where f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql1 = sql1 & " f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "' and migrado is null "
    
    sql = sql & sql1
    
    Conn.Execute sql
    
    sql = "update dosisnohomog "
    sql = sql & sql1
    
    Conn.Execute sql
    
End Sub

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

' este procedimiento crea registros de tipo 05 si no los hay para el 07
Private Sub InsertarRegistros05()
Dim sql As String
Dim sql1 As String
Dim rL As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim NF As Long
Dim plant As String

    ' cursor para buscar aquellas dosis organo sin dosis cuerpo
    sql = "select distinct dosisnohomog.c_empresa, dosisnohomog.c_instalacion, dosisnohomog.dni_usuario from dosisnohomog, voperarios "
    sql = sql & " where dosisnohomog.f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
    sql = sql & " voperarios.codusu = " & vUsu.codigo & " and "
    sql = sql & " dosisnohomog.f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "' and dosisnohomog.migrado is null  and "
    sql = sql & " voperarios.semigracsn = 1 and voperarios.dni= dosisnohomog.dni_usuario"
    
'    Sql = Sql & " dosimetros.n_reg_dosimetro = dosisnohomog.n_reg_dosimetro and "
'    Sql = Sql & " dosimetros.tipo_dosimetro = 1 "
    
    Set rL = New ADODB.Recordset
    rL.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Not rL.EOF Then rL.MoveFirst
    While Not rL.EOF
        sql1 = "select * from dosiscuerpo where dosiscuerpo.c_empresa = '" & Trim(rL.Fields(0).Value) & "' and "
        sql1 = sql1 & " dosiscuerpo.c_instalacion = '" & Trim(rL.Fields(1).Value) & "' and "
        sql1 = sql1 & " dosiscuerpo.dni_usuario = '" & Trim(rL.Fields(2).Value) & "' and "
        sql1 = sql1 & " dosiscuerpo.f_dosis >= '" & Format(Text1(0).Text, FormatoFecha) & "' and "
        sql1 = sql1 & " dosiscuerpo.f_dosis <= '" & Format(Text1(1).Text, FormatoFecha) & "' and dosiscuerpo.migrado is null " 'and "
'        sql1 = sql1 & " dosimetros.n_reg_dosimetro = dosiscuerpo.n_reg_dosimetro and "
'        sql1 = sql1 & " dosimetros.tipo_dosimetro = 0 "
    
        Set Rs = New ADODB.Recordset
        Rs.Open sql1, Conn, adOpenDynamic, adLockOptimistic, adCmdText
        
        If Rs.EOF Then
              sql = "insert into dosiscuerpo (n_registro, n_dosimetro, c_empresa, c_instalacion, "
              sql = sql & "dni_usuario, f_dosis, f_migracion, dosis_superf, dosis_profunda, "
              sql = sql & "plantilla_contrata, rama_generica, rama_especifica, c_tipo_trabajo,"
              sql = sql & "observaciones, migrado, n_reg_dosimetro) VALUES ("
              
              NF = SugerirCodigoSiguiente
                  
              sql = sql & ImporteSinFormato(CStr(NF)) & ",'0','"
              sql = sql & Trim(rL.Fields(0).Value)
              sql = sql & "','" & Trim(rL.Fields(1).Value) & "','" & Trim(rL.Fields(2).Value) & "',"
              sql = sql & "'" & Format(Text1(0).Text, FormatoFecha) & "',"
              sql = sql & "'" & Format(Text1(0).Text, FormatoFecha) & "',0,0,"
              
              ' sacamos los datos de la dosis
              plant = "" ' plantilla/contrata
              plant = DevuelveDesdeBD(1, "plantilla_contrata", "dosisnohomog", "c_empresa|c_instalacion|dni_usuario|", Trim(rL.Fields(0).Value) & "|" & Trim(rL.Fields(1).Value) & "|" & Trim(rL.Fields(2).Value) & "|", "T|T|T|", 3)
              If plant <> "" Then
                    sql = sql & "'" & Format(CInt(plant), "00") & "',"
              Else
                    sql = sql & "null,"
              End If
              
              plant = "" 'rama generica
              plant = DevuelveDesdeBD(1, "rama_generica", "dosisnohomog", "c_empresa|c_instalacion|dni_usuario|", Trim(rL.Fields(0).Value) & "|" & Trim(rL.Fields(1).Value) & "|" & Trim(rL.Fields(2).Value) & "|", "T|T|T|", 3)
              If plant <> "" Then
                    sql = sql & "'" & Format(CStr(plant), "00") & "',"
              Else
                    sql = sql & "null,"
              End If
              
              plant = "" 'rama especifica
              plant = DevuelveDesdeBD(1, "rama_especifica", "dosisnohomog", "c_empresa|c_instalacion|dni_usuario|", Trim(rL.Fields(0).Value) & "|" & Trim(rL.Fields(1).Value) & "|" & Trim(rL.Fields(2).Value) & "|", "T|T|T|", 3)
              If plant <> "" Then
                    sql = sql & "'" & Format(CStr(plant), "00") & "',"
              Else
                    sql = sql & "null,"
              End If
              
              plant = "" 'tipo de trabajo
              plant = DevuelveDesdeBD(1, "c_tipo_trabajo", "dosisnohomog", "c_empresa|c_instalacion|dni_usuario|", Trim(rL.Fields(0).Value) & "|" & Trim(rL.Fields(1).Value) & "|" & Trim(rL.Fields(2).Value) & "|", "T|T|T|", 3)
              If plant <> "" Then
                    sql = sql & "'" & Format(CStr(plant), "00") & "',"
              Else
                    sql = sql & "null,"
              End If
              
              'observaciones/ MIGRADO / n_reg_dosimetro
              sql = sql & "'AUTOMATICA',null,0 )"
              
              Conn.Execute sql
        
        End If
        
        
        Set Rs = Nothing
    
        rL.MoveNext
    Wend
    
    Set rL = Nothing
    
End Sub

Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "Select Max(n_registro) from dosiscuerpo"
    
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

