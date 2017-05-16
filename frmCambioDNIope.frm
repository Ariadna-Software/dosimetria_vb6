VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmCambioDNIope 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de D.N.I. de Operario"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmCambioDNIope.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListTipoMedicion 
      Height          =   5160
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
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         Caption         =   "DNI Operario"
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
         Height          =   1095
         Left            =   450
         TabIndex        =   8
         Top             =   2280
         Width           =   6525
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   13
            Top             =   300
            Width           =   3990
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   2
            Text            =   "Text5"
            Top             =   660
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   1
            Text            =   "Text5"
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Nuevo"
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
            Left            =   180
            TabIndex        =   10
            Top             =   690
            Width           =   915
         End
         Begin VB.Label Label3 
            Caption         =   "Anterior"
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
            Left            =   180
            TabIndex        =   9
            Top             =   330
            Width           =   735
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3690
         TabIndex        =   4
         Top             =   4200
         Width           =   1425
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   2010
         TabIndex        =   3
         Top             =   4200
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   300
         Left            =   420
         TabIndex        =   7
         Top             =   3570
         Visible         =   0   'False
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         Caption         =   "Se controla que el nuevo DNI no exista ya en nuestra Base de Datos"
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
         Height          =   375
         Left            =   540
         TabIndex        =   12
         Top             =   1080
         Width           =   6165
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
         Top             =   1710
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Este programa nos permite modificar el DNI en todos los modulos que  comprenden esta aplicacion."
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
         Left            =   510
         TabIndex        =   6
         Top             =   300
         Width           =   6045
      End
   End
End
Attribute VB_Name = "FrmCambioDNIope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim ape1 As String
Dim ape2 As String
Dim nombre As String


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

    If Not DatosOk Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Conn.BeginTrans

    pb1.max = 12
    pb1.Visible = True
    pb1.Value = 0
    Me.Refresh

    
    ActualizarTablas
    
    


eErrorCarga:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en el cambio de DNI del Operario."
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        MsgBox "Proceso Finalizado Correctamente", vbExclamation, "Cambio de DNI."
        cmdCancelar_Click
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim ano As Currency
Dim Mes As Currency

    ActivarCLAVE

    Text1(0).Text = ""
    Text1(1).Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    BloqueoManual False, "CAMBIDNI", "CAMBIDNI"
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
        Case 0, 1
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, ">")
            
            If Index = 0 Then
                ape1 = ""
                ape2 = ""
                nombre = ""
                CargarDatosOperarios Text1(0).Text, ape1, ape2, nombre
                Text2(1).Text = Trim(ape1) & " " & Trim(ape2) & ", " & Trim(nombre)
                If Trim(Text2(1).Text) = "," Then
                    MsgBox "Este Dni de operario no existe. Reintroduzca."
                    Text2(1).Text = "NO EXISTE"
                    Text1(0).Text = ""
                    PonerFoco Text1(0)
                End If
            End If
        Case 3
           If Trim(Text1(3).Text) <> Trim(Text1(3).Tag) Then
             MsgBox "    Acceso denegado    ", vbExclamation
             Text1(3).Text = ""
             PonerFoco Text1(3)
           Else
             DesactivarCLAVE
             PonerFoco Text1(0)
           End If
       
    End Select
    
    '---
End Sub

Private Sub ActivarCLAVE()
Dim I As Integer
    
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    Text1(3).Enabled = True
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = True

End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer

    Text1(0).Enabled = True
    Text1(1).Enabled = True
    Text1(3).Text = False
    
    cmdAceptar.Enabled = True
End Sub

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ActualizarTablas()
Dim sql As String
Dim sql1 As String

    pb1.Value = pb1.Value + 1
    Me.Refresh

    sql1 = "CREATE TEMPORARY TABLE temp_operarios ("
    sql1 = sql1 & "dni varchar(11) NOT NULL default '',"
    sql1 = sql1 & "n_seg_social varchar(15) default NULL,"
    sql1 = sql1 & "n_carnet_radiolog varchar(10) default NULL,"
    sql1 = sql1 & "f_emi_carnet_rad date default NULL,"
    sql1 = sql1 & "apellido_1 varchar(15) NOT NULL default '',"
    sql1 = sql1 & "apellido_2 varchar(15) NOT NULL default '',"
    sql1 = sql1 & "nombre varchar(15) NOT NULL default '',"
    sql1 = sql1 & "direccion varchar(40) default NULL,"
    sql1 = sql1 & "poblacion varchar(30) default NULL,"
    sql1 = sql1 & "c_postal char(2) NOT NULL default '',"
    sql1 = sql1 & "distrito char(3) default NULL,"
    sql1 = sql1 & "c_tipo_trabajo char(2) NOT NULL default '',"
    sql1 = sql1 & "f_nacimiento date default NULL,"
    sql1 = sql1 & "profesion_catego varchar(30) default NULL,"
    sql1 = sql1 & "sexo char(1) NOT NULL default '',"
    sql1 = sql1 & "plantilla_contrata char(2) NOT NULL default '',"
    sql1 = sql1 & "f_alta date NOT NULL default '0000-00-00',"
    sql1 = sql1 & "f_baja date default NULL,"
    sql1 = sql1 & "migrado char(2) default NULL,"
    sql1 = sql1 & "cod_rama_gen char(2) NOT NULL default '',"
    sql1 = sql1 & "semigracsn tinyint(1) NOT NULL default '1');"


    Conn.Execute sql1

    'operarios
    sql = "insert into temp_operarios select '" & Trim(Text1(1).Text) & "',"
    sql = sql & " n_seg_social,n_carnet_radiolog,f_emi_carnet_rad,apellido_1,apellido_2,"
    sql = sql & " nombre,direccion,poblacion,c_postal,distrito,c_tipo_trabajo,f_nacimiento,"
    sql = sql & " profesion_catego,sexo,plantilla_contrata,f_alta,f_baja,migrado,cod_rama_gen, "
    sql = sql & " semigracsn "
    sql = sql & " from operarios where dni = '" & Trim(Text1(0).Text) & "'"

    Conn.Execute sql

    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    'operarios
    
    sql = "insert into operarios select * from temp_operarios"
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    'operainstala
    sql = "CREATE TEMPORARY TABLE temp_operainstala ("
    sql = sql & "c_empresa varchar(11),"
    sql = sql & "c_instalacion varchar(11),"
    sql = sql & "dni varchar(11), "
    sql = sql & "f_alta date,"
    sql = sql & "f_baja date,"
    sql = sql & "migrado char(2));"
    
    Conn.Execute sql
    
    sql = "INSERT INTO temp_operainstala select * from operainstala where dni = '" & Trim(Text1(0).Text) & "'"
    Conn.Execute sql
    
    sql = "UPDATE temp_operainstala set dni = '" & Trim(Text1(1).Text) & "'"
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    sql = "INSERT INTO operainstala SELECT * FROM temp_operainstala"
    Conn.Execute sql
    
   
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    'dosimetros
    sql = "update dosimetros set dni_usuario = '" & Trim(Text1(1).Text)
    sql = sql & "' where dni_usuario = '" & Trim(Text1(0).Text) & "'"
    
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    'dosiscuerpo
    sql = "update dosiscuerpo set dni_usuario = '" & Trim(Text1(1).Text)
    sql = sql & "' where dni_usuario = '" & Trim(Text1(0).Text) & "'"
    
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    'recepdosim
    sql = "update recepdosim set dni_usuario = '" & Trim(Text1(1).Text)
    sql = sql & "' where dni_usuario = '" & Trim(Text1(0).Text) & "'"
    
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
     
    
    'dosisarea
    sql = "update dosisarea set dni_usuario = '" & Trim(Text1(1).Text)
    sql = sql & "' where dni_usuario = '" & Trim(Text1(0).Text) & "'"
    
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    'dosisnohomog
    sql = "update dosisnohomog set dni_usuario = '" & Trim(Text1(1).Text)
    sql = sql & "' where dni_usuario = '" & Trim(Text1(0).Text) & "'"
    
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    sql = "delete from operainstala where dni = '" & Trim(Text1(0).Text) & "'"
    Conn.Execute sql
    
    sql = " DROP TABLE IF EXISTS temp_operainstala;"
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    
    sql = "delete from operarios where dni = '" & Trim(Text1(0).Text) & "'"
    Conn.Execute sql
    
    sql = " DROP TABLE IF EXISTS temp_operarios;"
    Conn.Execute sql
    
    pb1.Value = pb1.Value + 1
    Me.Refresh
    
    
End Sub
Private Function DatosOk() As Boolean
    Dim Rs As ADODB.Recordset
    Dim b As Boolean
    Dim Mens As String
    
    DatosOk = True
    
    ' ambos campos deben ser distinto de vacio
    If Text1(0).Text = "" Or Text1(1).Text = "" Then
        MsgBox "Debe de introducir valor en ambos NIF. Revise."
        DatosOk = False
        Exit Function
    End If
    
    '-- Comprobar NIF hasta correcto
    If Not Comprobar_NIF(Text1(1)) Then
        Mens = "El NIF introducido es incorrecto. ¿Desea continuar de todos modos?"
        If MsgBox(Mens, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then 'VRS:1.0.1(11)
            DatosOk = False
            Exit Function
        End If
    End If

    ' el nif hasta no debe de existir en la base de datos
    sql = ""
    sql = DevuelveDesdeBD(1, "dni", "operarios", "dni|", Trim(Text1(1).Text) & "|", "T|", 1)
    If sql <> "" Then
        MsgBox "El nuevo NIF no debe de existir en la base de datos. Reintroduzca.", vbExclamation
        DatosOk = False
        Exit Function
    End If

End Function


