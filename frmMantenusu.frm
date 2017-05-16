VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMantenusu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de usuarios"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmMantenusu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameUsuario 
      Height          =   3045
      Left            =   3480
      TabIndex        =   24
      Top             =   2235
      Width           =   5670
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1185
         PasswordChar    =   "*"
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2385
         Width           =   1350
      End
      Begin VB.CommandButton cmdFrameUsu 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4305
         TabIndex        =   16
         Top             =   2535
         Width           =   1215
      End
      Begin VB.CommandButton cmdFrameUsu 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2985
         TabIndex        =   15
         Top             =   2535
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1185
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1860
         Width           =   1350
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1155
         Width           =   4335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMantenusu.frx":030A
         Left            =   1725
         List            =   "frmMantenusu.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   525
         Width           =   2760
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2970
         TabIndex        =   30
         Top             =   1830
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         Height          =   1245
         Left            =   165
         Top             =   1650
         Width           =   2565
      End
      Begin VB.Label Label4 
         Caption         =   "Confirma Password"
         Height          =   420
         Index           =   3
         Left            =   375
         TabIndex        =   29
         Top             =   2310
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   375
         TabIndex        =   28
         Top             =   1905
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel"
         Height          =   255
         Left            =   1725
         TabIndex        =   27
         Top             =   285
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre completo"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   26
         Top             =   915
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Login"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   285
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame FrameNormal 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   3480
         TabIndex        =   18
         Top             =   360
         Width           =   5655
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   480
            Width           =   4335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMantenusu.frx":033C
            Left            =   120
            List            =   "frmMantenusu.frx":0349
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre completo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Nivel"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmMantenusu.frx":036E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo usuario"
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton cmdEmp 
         Height          =   375
         Index           =   0
         Left            =   3480
         Picture         =   "frmMantenusu.frx":0470
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Nueva bloqueo empresa"
         Top             =   5400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "frmMantenusu.frx":0572
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modificar usuario"
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   2
         Left            =   1080
         Picture         =   "frmMantenusu.frx":0674
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar usuario"
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton cmdEmp 
         Height          =   375
         Index           =   1
         Left            =   3960
         Picture         =   "frmMantenusu.frx":0776
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar bloqueo empresa"
         Top             =   5400
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Login"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   3480
         TabIndex        =   3
         Tag             =   $"frmMantenusu.frx":0878
         Top             =   2520
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resum."
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   22
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Empresas NO permitidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   21
         Top             =   2280
         Visible         =   0   'False
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmMantenusu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim sql As String
Dim I As Integer


Private Sub cmdEmp_Click(Index As Integer)
Dim cont As Integer

    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un usuario", vbExclamation
        Exit Sub
    End If
    
    If Index = 0 Then


        'nueva Empresa bloqueada para el usuario
        CadenaDesdeOtroForm = ""
        frmMensajes.Opcion = 4
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            cont = RecuperaValor(CadenaDesdeOtroForm, 1)
            If cont = 0 Then Exit Sub
            For I = 1 To cont
                'No hacemos nada
            Next I
            For I = 0 To cont - 1
                sql = RecuperaValor(CadenaDesdeOtroForm, I + cont + 2)
                InsertarEmpresa CInt(sql)
            Next I
        
        Else
            Exit Sub
        End If
        
    Else
        If ListView2.SelectedItem Is Nothing Then Exit Sub
        sql = "Va a  desbloquear el acceso" & vbCrLf
        sql = sql & vbCrLf & "a la empresa:   " & ListView2.SelectedItem.SubItems(1) & vbCrLf
        sql = sql & "para el usuario:   " & ListView1.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "     ¿Desea continuar?"
        If MsgBox(sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then 'VRS:1.0.1(11)
            sql = "Delete FROM usuarioempresadosis WHERE codusu =" & ListView1.SelectedItem.Text
            sql = sql & " AND codempre = " & ListView2.SelectedItem.Text
            Conn.Execute sql
        Else
            Exit Sub
        End If
    End If
    'Llegados aqui recargamos los datos del usuario
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub


Private Sub InsertarEmpresa(Empresa As Integer)
    sql = "INSERT INTO usuarioempresadosis(codusu,codempre) VALUES ("
    sql = sql & ListView1.SelectedItem.Text & "," & Empresa & ")"
    On Error Resume Next
    Conn.Execute sql
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
    
    End If
    
End Sub


Private Sub cmdFrameUsu_Click(Index As Integer)



    If Index = 0 Then
        For I = 0 To Text2.Count - 1
            Text2(I).Text = Trim(Text2(I).Text)
            If Text2(I).Text = "" Then
                MsgBox Label4(I).Caption & " requerido.", vbExclamation
                Exit Sub
            End If
        Next I
        
        If Combo2.ListIndex < 0 Then
            MsgBox "Seleccione un nivel de acceso", vbExclamation
            Exit Sub
        End If
    
        'Password
        If Text2(2).Text <> Text2(3).Text Then
            MsgBox "Password y confirmacion de password no coinciden", vbExclamation
            Exit Sub
        End If
        
        'Compruebo que el login es unico
        If UCase(Label6.Caption) = "NUEVO" Then
            Set miRsAux = New ADODB.Recordset
            sql = "Select login from Usuarios where login='" & DevNombreSQL(Text2(0).Text) & "'"
            miRsAux.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            sql = ""
            If Not miRsAux.EOF Then sql = "Ya existe en la tabla usuarios uno con el login: " & miRsAux.Fields(0)
            miRsAux.Close
            Set miRsAux = Nothing
            If sql <> "" Then
                MsgBox sql, vbExclamation
                Exit Sub
            End If
        End If
        InsertarModificar
        
        
    End If
    'Cargar usuarios
    If UCase(Label6.Caption) = "NUEVO" Then
        CargaUsuarios
    Else
        'Pero cargamos el tag como coresponde
        ListView1.SelectedItem.Tag = Combo2.ItemData(Combo2.ListIndex) & "|" & Text2(1).Text & "|"
    
        DatosUsusario
    End If
    'Para ambos casos
    Me.FrameUsuario.Visible = False
    Me.FrameNormal.Enabled = True
End Sub


Private Sub InsertarModificar()
On Error GoTo EInsertarModificar

    Set miRsAux = New ADODB.Recordset
    If UCase(Label6.Caption) = "NUEVO" Then
        
        'Nuevo
        sql = "Select max(codusu) from Usuarios"
        miRsAux.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 0
        If Not miRsAux.EOF Then I = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        I = I + 1
        
        sql = "INSERT INTO usuarios (codusu, nomusu,  nivelusu, login, passwordpropio) VALUES ("
        sql = sql & I
        sql = sql & ",'" & DevNombreSQL(Text2(1).Text) & "',"
        'Combo
        sql = sql & Combo2.ItemData(Combo2.ListIndex) & ",'"
        sql = sql & DevNombreSQL(Text2(0).Text) & "','"
        sql = sql & DevNombreSQL(Text2(3).Text) & "')"
        
    Else
        sql = "UPDATE Usuarios Set nomusu='" & DevNombreSQL(Text2(1).Text)
        
        'Si el combo es administrador compruebo que no fuera en un principio SUPERUSUARIO
'        If Combo2.ListIndex = 2 Then
'            'Si el combo1 es 3 entonces es super
'            If Combo1.ListIndex = 3 Then
 ''               I = 0
 '           Else
                I = 1
 '           End If
'        Else
            I = Combo2.ItemData(Combo2.ListIndex)
'        End If
        sql = sql & "' , nivelusu =" & I
        'SQL = SQL & "  , login = '" & Text2(2).Text
        sql = sql & "  , passwordpropio = '" & DevNombreSQL(Text2(3).Text)
        sql = sql & "' WHERE codusu = " & ListView1.SelectedItem.Text
    End If
    Conn.Execute sql

    Exit Sub
EInsertarModificar:
    MuestraError Err.Number, "EInsertarModificar"
End Sub


Private Sub cmdUsu_Click(Index As Integer)

Dim MiRsUsu As ADODB.Recordset
Dim v_aux As Integer
Dim Borrar As Boolean
    
    Select Case Index
    Case 0, 1
        If Index = 0 Then
            'Nuevo usuario
            Limpiar Me
            Label6.Caption = "NUEVO"
            I = 0 'Para el foco
        Else
            'Modificar
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            Label6.Caption = "MODIFICAR"
            Set miRsAux = New ADODB.Recordset
            sql = "Select * from usuarios where codusu = " & ListView1.SelectedItem.Text
            miRsAux.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "Error inesperado: Leer datos usuarios", vbExclamation
            Else
                Text2(0).Text = miRsAux!Login
                Text2(1).Text = miRsAux!nomusu
                Text2(2).Text = miRsAux!passwordpropio
                Text2(3).Text = miRsAux!passwordpropio
                Combo2.ListIndex = miRsAux!NivelUsu
            End If
            I = 1 'Para el foco
        End If
        Text2(0).Enabled = (Index = 0)
        Me.FrameNormal.Enabled = False
        Me.FrameUsuario.Visible = True
        Text2(I).SetFocus
    Case 2
        ' borramos el usuario unicamente si es usuario de nuestra
        ' aplicacion y no de otras
        Set MiRsUsu = New ADODB.Recordset
        sql = "Select * from usuarios where codusu = " & ListView1.SelectedItem.Text
        MiRsUsu.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If MiRsUsu.EOF Then
            MsgBox "Error inesperado: Leer datos usuarios", vbExclamation
        Else
        If (vUsu.NivelUsu > 1) And (vUsu.Codusu <> CLng(ListView1.SelectedItem.Text)) Then
            ' comprobamos que podemos borrar el usuario si
            ' tenemos bastante nivel
            Borrar = True
            For v_aux = 0 To MiRsUsu.Fields.Count - 1
              If MiRsUsu.Fields(v_aux).Name <> "nivelUsu" Then
                If Mid(MiRsUsu.Fields(v_aux).Name, 1, 5) = "nivel" Then
'                    If MiRsUsu.Fields(v_aux).Value > -1 Then
'                        Borrar = False
'                        Exit For
'                    End If
                End If
              End If
            Next v_aux
            If Borrar Then
                sql = "delete from usuarios where codusu = " & ListView1.SelectedItem.Text
                Conn.Execute sql
                sql = "delete from usuarioempresadosis where codusu = " & ListView1.SelectedItem.Text
                Conn.Execute sql
                CargaUsuarios
            End If
        End If
        End If
    
    End Select

End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.ListView1.SmallIcons = frmPpal.ImageList1
        Me.ListView2.SmallIcons = frmPpal.ImageList1
        CargaUsuarios
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.FrameUsuario.Visible = False
    Me.FrameNormal.Enabled = True
End Sub



Private Sub CargaUsuarios()
Dim Itm As ListItem

    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    '                               Aquellos usuarios k tengan nivel usu -1 NO son de conta
    '  QUitamos codusu=0 pq es el usuario ROOT
    sql = "Select * from usuarios where nivelUsu >=0 and codusu > 0 order by codusu"
    miRsAux.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set Itm = ListView1.ListItems.Add
        Itm.Text = miRsAux!Codusu
        Itm.SubItems(1) = miRsAux!Login
        Itm.SmallIcon = 8
        'Nombre y nivel de usuario
        sql = miRsAux!NivelUsu & "|" & miRsAux!nomusu & "|"
        Itm.Tag = sql
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ListView1.ListItems.Count > 0 Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
        DatosUsusario
    End If

End Sub



Private Sub DatosUsusario()
Dim ItmX As ListItem
On Error GoTo EDatosUsu

    If ListView1.SelectedItem Is Nothing Then
        Text4.Text = ""
        Combo1.ListIndex = -1
        Exit Sub
    End If


    Text4.Text = RecuperaValor(ListView1.SelectedItem.Tag, 2)
    'NIVEL
    sql = RecuperaValor(ListView1.SelectedItem.Tag, 1)
    '                           COMBO                      en Bd
    '                       0.- Consulta                     0
    '                       1.- Normal                       1
    '                       2.- Administrador                2
    '                       3.- SuperUsuario (root)          3
    If Not IsNumeric(sql) Then sql = 0
    Combo1.ListIndex = sql
    ListView2.ListItems.Clear
    sql = ListView2.Tag & ListView1.SelectedItem.Text
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set ItmX = ListView2.ListItems.Add
        ItmX.Text = miRsAux.Fields(0)
        ItmX.SubItems(1) = miRsAux!nomempre
        ItmX.SubItems(2) = miRsAux!nomresum
        ItmX.SmallIcon = 9 '20
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Exit Sub
EDatosUsu:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

