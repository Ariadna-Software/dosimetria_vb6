VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensaje"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   4365
      Left            =   90
      TabIndex        =   20
      Top             =   75
      Width           =   6525
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Continuar"
         Height          =   645
         Left            =   2580
         TabIndex        =   1
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Si tiene que dar de alta alguna instalación o usuario de una empresa existente, vaya a los mantenimientos correspondientes."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   480
         TabIndex        =   22
         Top             =   1890
         Width           =   5775
      End
      Begin VB.Label Label25 
         Caption         =   $"frmMensajes.frx":030A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   540
         TabIndex        =   21
         Top             =   870
         Width           =   5295
      End
   End
   Begin VB.Frame FrameRemesas 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      TabIndex        =   16
      Top             =   30
      Width           =   6975
      Begin VB.CommandButton CmdRemesa 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   17
         Top             =   4770
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":03A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":5B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":65A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3735
         Left            =   135
         TabIndex        =   18
         Top             =   810
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Remesa"
            Object.Width           =   1994
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "F.Remesa"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "No.Efectos"
            Object.Width           =   1835
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Importe Efectos"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Situacion"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Remesas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   180
         Width           =   5295
      End
   End
   Begin VB.Frame frameCalculoSaldos 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "&Iniciar"
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   12
         Top             =   4800
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":69F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":C1E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":CBFA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   4770
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   135
         TabIndex        =   14
         Top             =   810
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2170
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Cálculo de saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   180
         Width           =   5295
      End
   End
   Begin VB.Timer tCuadre 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6420
      Top             =   5700
   End
   Begin VB.Frame FrameeMPRESAS 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   9
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Regresar"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   8
         Top             =   4800
         Width           =   975
      End
      Begin MSComctlLib.ListView lwE 
         Height          =   3615
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "dsdsd"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Label Label11 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   120
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   2040
      X2              =   3960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label10 
      Caption         =   "años de vida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label10 
      Caption         =   "Valor adquisición"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TABLAS"
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
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '1.- Aviso de utilizacion de ficha de personal
    '2.- Fecha de Apertura de Artículo
    '4.- Seleccionar empresas
    
Public Parametros As String
    '1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private PrimeraVez As Boolean

Dim I As Integer
Dim sql As String
Dim Rs As Recordset
Dim ItmX As ListItem


Private Sub cmdEmpresa_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        sql = ""
        Parametros = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                sql = sql & Me.lwE.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & sql
        'Vemos las conta
        sql = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                sql = sql & Me.lwE.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & sql
    End If
    Unload Me
End Sub

Private Sub CmdRemesa_Click(Index As Integer)
Dim Cad As String

    If ListView2.SelectedItem Is Nothing Then Exit Sub
    If ListView2.SelectedItem = -1 Then
        MsgBox "Ningún registro a devolver.", vbExclamation, "¡Atención!"
        Exit Sub
    End If
    
    Cad = ListView2.SelectedItem & "|"
    Cad = Cad & ListView2.SelectedItem.SubItems(1) & "|"
    
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me

End Sub

Private Sub Command1_Click(Index As Integer)
Unload Me
End Sub

'Private Sub Command2_Click()
'Dim Digitos As Integer
'    ListView1.ListItems.Clear
'    Me.ProgressBar1.Value = 0
'    Me.ProgressBar1.Max = vEmpresa.numnivel + 1
'    Me.ProgressBar1.Visible = True
'    Screen.MousePointer = vbHourglass
'    'Calculamos en historico
''    CalculaSaldosNivel True, Digitos
'    'Iniciamos el calculo de saldos para cada nivel
'    For i = 1 To vEmpresa.numnivel
'        Digitos = DigitosNivel(i)
''        CalculaSaldosNivel False, Digitos
'    Next i
'    Me.ProgressBar1.Visible = False
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 5 Then
            Screen.MousePointer = vbHourglass
            Me.tCuadre.Enabled = True
        End If
    Else
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim W, H
    Me.tCuadre.Enabled = False
    PrimeraVez = True
    Me.frameCalculoSaldos.Visible = False
    Me.FrameeMPRESAS.Visible = False
    Me.FrameRemesas.Visible = False
    Me.Frame3.Visible = False
    
    Select Case Opcion
    Case 1
        Me.Caption = "Remesas"
        W = FrameRemesas.Width
        H = Me.FrameRemesas.Height
        Me.FrameRemesas.Visible = True
'        CargaRemesas
    Case 0
        Me.Caption = "AVISO IMPORTANTE"
        W = Frame3.Width
        H = Me.Frame3.Height
        Me.Frame3.Visible = True
    Case 4
        Me.Caption = "Seleccion"
        W = Me.FrameeMPRESAS.Width
        H = Me.FrameeMPRESAS.Height + 200
        Me.FrameeMPRESAS.Visible = True
        cargaempresas
    End Select
    Me.Width = W + 120
    Me.Height = H + 120
End Sub


Private Sub cargaempresas()
On Error GoTo Ecargaempresas

    sql = "Select * from empresadosis order by codempre"
    Set lwE.SmallIcons = Me.ImageList1
    lwE.ListItems.Clear
    Set Rs = New ADODB.Recordset
    I = -1
    Rs.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set ItmX = lwE.ListItems.Add(, , Rs!nomempre, , 3)
        ItmX.Tag = Rs!codempre
'        If vParam Is Nothing Then
            If ItmX.Tag = DevuelveDesdeBD(1, "codempre", "parametros", "codempre|", "codempre|", "N|", 1) Then
                ItmX.Checked = True
                I = ItmX.Index
            End If
'        Else
'            If ItmX.Tag = vParam.NumeroEmpresa Then
'                ItmX.Checked = True
'                i = ItmX.Index
'            End If
'
'        End If
        ItmX.ToolTipText = Rs!sumi
        Rs.MoveNext
    Wend
    Rs.Close
    If I > 0 Then Set lwE.SelectedItem = lwE.ListItems(I)
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
    Set Rs = Nothing
End Sub

Private Sub KEYpress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'Private Sub CargaRemesas()
'On Error GoTo ECargaRemesas
'
'    Sql = "Select sremes.numremes, sremes.fecremes, count(*), sum(sremes.impefect), if (situacio= 1," & """Abonada""" & "," & """No Abonada" & """)"
'    Sql = Sql & "  from sremes group by 1,2"
'    Set lwE.SmallIcons = Me.ImageList1
'    lwE.ListItems.Clear
'    Set Rs = New ADODB.Recordset
'    i = -1
'    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not Rs.EOF
'        Set ItmX = ListView2.ListItems.Add()
'
'        ItmX.Text = Rs!NumRemes ' VRS:1.0.2(10) Format(rs!numalbpr, "000000")
'        ItmX.SubItems(1) = Format(Rs!fecremes, "dd/mm/yyyy")
'        ItmX.SubItems(2) = Format(Rs.Fields(2).Value, "#,##0")
'        ItmX.SubItems(3) = Format(Rs.Fields(3).Value, "###,###,###,##0.00")
'        ItmX.SubItems(4) = Rs.Fields(4).Value
'        i = ItmX.Index
'        'Sig
'        Rs.MoveNext
'    Wend
'    Rs.Close
'    'If i > 0 Then Set lwE.SelectedItem = lwE.ListItems(i - 1)
'ECargaRemesas:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos remesas"
'    Set Rs = Nothing
'End Sub

'Private Function FechaOk() As Boolean
'    FechaOk = False
'    If Text3.Text <> "" Then
'      If Not EsFechaOK(Text3) Then
'            MsgBox "Fecha incorrecta: " & Text3.Text, vbExclamation
'            Text3.Text = ""
'            PonerFoco Text3
'            Exit Function
'      End If
'      Text3.Text = Format(Text3.Text, "dd/mm/yyyy")
'      FechaOk = True
'    Else
'        MsgBox "Debe introducir un valor. Reintroduzca. ", vbExclamation
'        PonerFoco Text3
'    End If
'End Function

Private Sub CmdOk_Click()
   Unload Me
End Sub


Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Unload(cancel As Integer)
'    If Opcion = 2 Then
'        If Not FechaOk Then
'            Cancel = True
'        Else
'        End If
'
'    End If
End Sub

'Private Sub Text3_GotFocus()
'    Text3.SelStart = 0
'    Text3.SelLength = Len(Text3.Text)
'End Sub
'
'Private Sub Text3_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
'End Sub
'
'Private Sub Text3_LostFocus()
'    If FechaOk Then Exit Sub
'End Sub
