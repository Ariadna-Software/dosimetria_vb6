VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCancelacionMigra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de  Migración "
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   6735
   Icon            =   "frmCancelacionMigra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAceptarCancelacion 
      Caption         =   "&Aceptar"
      Height          =   585
      Left            =   4110
      TabIndex        =   4
      Top             =   4635
      Width           =   1125
   End
   Begin VB.CommandButton CmdCanCancelacion 
      Caption         =   "&Cancelar"
      Height          =   585
      Left            =   5370
      TabIndex        =   5
      Top             =   4620
      Width           =   1095
   End
   Begin VB.Frame FrameCancelacion 
      Height          =   2160
      Left            =   120
      TabIndex        =   6
      Top             =   2220
      Width           =   6390
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1170
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   630
         Width           =   1185
      End
      Begin VB.TextBox txtReg 
         Height          =   285
         Index           =   1
         Left            =   4350
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox txtReg 
         Height          =   285
         Index           =   0
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   5085
         MaxLength       =   15
         TabIndex        =   1
         Top             =   630
         Width           =   1020
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   510
         Left            =   300
         TabIndex        =   7
         Top             =   2310
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Index           =   0
         Left            =   345
         TabIndex        =   15
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   11
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   9
         Left            =   3570
         TabIndex        =   10
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Registro"
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
         Index           =   10
         Left            =   345
         TabIndex        =   9
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   4860
         Picture         =   "frmCancelacionMigra.frx":030A
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Migración"
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
         Index           =   11
         Left            =   3060
         TabIndex        =   8
         Top             =   660
         Width           =   1620
      End
   End
   Begin VB.Label Label5 
      Caption         =   $"frmCancelacionMigra.frx":040C
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
      Height          =   675
      Left            =   300
      TabIndex        =   14
      Top             =   1440
      Width           =   6195
   End
   Begin VB.Label Label2 
      Caption         =   "El programa controla que el N. DE REGISTRO (DESDE-HASTA) para cancelar un  grupo determinado. "
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
      Height          =   555
      Left            =   300
      TabIndex        =   13
      Top             =   750
      Width           =   6195
   End
   Begin VB.Label Label1 
      Caption         =   "Este programa nos permite CANCELAR UNA MIGRACION DEFECTUOSA del archivo de dosis Homogeneas o dosis Area."
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
      Height          =   465
      Left            =   300
      TabIndex        =   12
      Top             =   210
      Width           =   6195
   End
End
Attribute VB_Name = "frmCancelacionMigra"
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
Dim RC As String
 
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


Private Sub CmdAceptarCancelacion_Click()
    If txtReg(0).Text <> "" And txtReg(1).Text <> "" Then
    
        BorradoRegistros
        
    Else
        MsgBox "Debe introducir los valores de número de registro desde y hasta", vbExclamation
        
        PonerFoco txtReg(0)
    
    End If
End Sub

Private Sub CmdCanCancelacion_Click()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    BloqueoManual False, "CANMIGRA", "CANMIGRA"
End Sub

Private Sub ImgFec_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.fecha = Now
    If Text3(Index).Text <> "" Then frmC.fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing

End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    '## A mano
    'Vemos como esta guardado el valor del check
    Me.Top = 0
    Me.Left = 0
      
    txtReg(0).Text = ""
    txtReg(1).Text = ""
    Text3(0).Text = ""
  
    CargarCombo1
    
    Combo2.ListIndex = 0
    
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BorradoRegistros()
Dim sql As String

    On Error GoTo eBorradoRegistros
    
    Conn.BeginTrans
    
    If Combo2.ListIndex = 0 Then
        sql = "delete from dosiscuerpo where n_registro >=" & txtReg(0).Text & " and "
        sql = sql & "n_registro <= " & txtReg(1).Text
    Else
        sql = "delete from dosisarea where n_registro >=" & txtReg(0).Text & " and "
        sql = sql & "n_registro <= " & txtReg(1).Text
    End If
    
    Conn.Execute sql
    
eBorradoRegistros:
    If Err.Number <> 0 Then
        If Combo2.ListIndex = 0 Then
            MuestraError Err.Number, "Error en el borrado de Registros de Dosis Cuerpo"
        Else
            MuestraError Err.Number, "Error en el borrado de Registros de Dosis Area"
        End If
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        LimpiarCampos
        If Combo2.ListIndex = 0 Then
            MsgBox "Borrado de registros de Dosis Cuerpo realizado correctamente.", vbExclamation, "Proceso de borrado."
        Else
            MsgBox "Borrado de registros de Dosis Area realizado correctamente.", vbExclamation, "Proceso de borrado."
        End If
    End If

End Sub


Private Sub txtreg_GotFocus(Index As Integer)
    With txtReg(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtreg_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub txtReg_LostFocus(Index As Integer)
Dim sql As String
Dim valor As Currency

    ''Quitamos blancos por los lados
    txtReg(Index).Text = Trim(txtReg(Index).Text)
    If txtReg(Index).Text = "" Then Exit Sub
    If txtReg(Index).BackColor = vbYellow Then
        txtReg(Index).BackColor = vbWhite
    End If
    
    If txtReg(Index) = "" Then Exit Sub
    
    If ValorAnterior = txtReg(Index).Text Then Exit Sub
    
    If Modo = 3 And ConCaracteresBusqueda(txtReg(Index).Text) Then Exit Sub 'Busquedas
    
    If EsNumerico(txtReg(Index).Text) Then
        If InStr(1, txtReg(Index).Text, ",") > 0 Then
            valor = ImporteFormateado(txtReg(Index).Text)
        Else
            valor = CCur(TransformaPuntosComas(txtReg(Index).Text))
        End If
        
        'miramos que el registro existe en dosiscuerpo y que no está migrado a CSN
        sql = ""
        Dim cad1 As String
        cad1 = "migrado"
        If Combo2.ListIndex = 0 Then
            sql = DevuelveDesdeBD(1, "n_registro", "dosiscuerpo", "n_registro|", txtReg(Index).Text & "|", "N|", 1, cad1)
        Else
            sql = DevuelveDesdeBD(1, "n_registro", "dosisarea", "n_registro|", txtReg(Index).Text & "|", "N|", 1, cad1)
        End If
        If sql = "" Then
            MsgBox "No existe este número registro. Reintroduzca", vbExclamation
            txtReg(Index).Text = ""
            PonerFoco txtReg(Index)
        Else
           If cad1 = "**" Then
                MsgBox "Este número de registro está migrado al CSN. Reintroduzca."
                txtReg(Index).Text = ""
                PonerFoco txtReg(Index)
           End If
        End If
    End If
End Sub

Private Sub LimpiarCampos()
    txtReg(0).Text = ""
    txtReg(1).Text = ""
    Text3(0).Text = ""
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

    Combo2.Clear
    Combo2.AddItem "Personal"
    Combo2.ItemData(Combo2.NewIndex) = 0
    
    Combo2.AddItem "Area"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
End Sub

