VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmCambioInstala 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de número de Instalación"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmCambioInstala.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Nueva"
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
      Height          =   1110
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   3225
      Width           =   6525
      Begin VB.TextBox Text1 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1050
         TabIndex        =   3
         Top             =   270
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2415
         MaxLength       =   30
         TabIndex        =   18
         Top             =   270
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1050
         TabIndex        =   4
         Top             =   675
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2415
         MaxLength       =   30
         TabIndex        =   16
         Top             =   675
         Width           =   3990
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa"
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
         TabIndex        =   19
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Instalación"
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
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   705
         Width           =   840
      End
   End
   Begin VB.Frame FrameListTipoMedicion 
      Height          =   5460
      Left            =   30
      TabIndex        =   7
      Top             =   60
      Width           =   7275
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3510
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   0
         Tag             =   "JMCE"
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         Caption         =   "Anterior"
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
         Height          =   1110
         Index           =   0
         Left            =   450
         TabIndex        =   10
         Top             =   2040
         Width           =   6525
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2415
            MaxLength       =   30
            TabIndex        =   20
            Top             =   270
            Width           =   3990
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   1
            Left            =   1050
            TabIndex        =   1
            Top             =   270
            Width           =   1275
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2415
            MaxLength       =   30
            TabIndex        =   14
            Top             =   675
            Width           =   3990
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1050
            TabIndex        =   2
            Top             =   675
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Empresa"
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
            Index           =   2
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   840
         End
         Begin VB.Label Label3 
            Caption         =   "Instalación"
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
            TabIndex        =   11
            Top             =   705
            Width           =   765
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3735
         TabIndex        =   6
         Top             =   4635
         Width           =   1425
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   2085
         TabIndex        =   5
         Top             =   4635
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   300
         Left            =   465
         TabIndex        =   9
         Top             =   4305
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Max             =   1000
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Se controla que la nueva Instalación no exista en nuestra Base de Datos"
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
         TabIndex        =   13
         Top             =   1080
         Width           =   6435
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
         TabIndex        =   12
         Top             =   1710
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Este programa nos permite modificar el número de instalacioóin en todos los modulos que  comprenden esta aplicacion."
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
         TabIndex        =   8
         Top             =   300
         Width           =   6045
      End
   End
End
Attribute VB_Name = "FrmCambioInstala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RC As String
Dim rs As Recordset
Dim PrimeraVez As Boolean
Dim SoloCopia As Boolean
Dim ape1 As String
Dim ape2 As String
Dim nombre As String

' Captura de pulsacines de tecla...
Private Sub KEYpress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{tab}"
  Else
    If KeyAscii = 27 Then
      cmdCancelar_Click
    End If
  End If
End Sub

' Adios.
Private Sub cmdCancelar_Click()
  Unload Me
End Sub

' Comenzar el proceso.
Private Sub cmdAceptar_Click()
Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim Tipo As String
Dim cont As Integer
On Error GoTo eErrorCarga
    
  Screen.MousePointer = vbHourglass
  SoloCopia = False
  
  ' Comprobar si los datos son correctos.
  If Not DatosOk Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
    
  ' Comienza la transacción.
  Conn.BeginTrans

  PB1.max = 16
  PB1.Visible = True
  PB1.Value = 0
  Me.Refresh
    
  ' Llamamos al proceso de renombramiento.
  ActualizarTablas

eErrorCarga:
    
    ' En caso de error...
  PB1.Visible = False
  If Err.Number <> 0 Then
    MuestraError Err.Number, "Error en el cambio de Instalación."
    Conn.RollbackTrans
  Else
    Conn.CommitTrans
    PB1.Value = PB1.max
    MsgBox "Proceso Finalizado Correctamente", vbInformation, "Cambio de Instalación."
       
    ' Limpiamos los textbox.
    For cont = 1 To 4
      Text1(cont).Text = ""
      Text2(cont).Text = ""
      PonerFoco Text1(1)
    Next cont
        
  End If
  Screen.MousePointer = vbDefault

End Sub

' Configura los textboxes.
Private Sub Form_Load()
  ActivarCLAVE
End Sub

' Desbloqueamos la aplicación de cambio de instalación al salir.
Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  BloqueoManual False, "CAMBIINS", "CAMBIINS"
End Sub

' Seleccionar todo el contenido del textbox
Private Sub Text1_GotFocus(Index As Integer)
  Text1(Index).SelStart = 0
  Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

' Capturamos las pulsaciones de teclas sobre los textbox.
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

' Mostrar la información correspondiente a los Text1 en los Text2
Private Sub Text1_LostFocus(Index As Integer)
Dim I As Integer
Dim sql As String
Dim mTag As CTag
Dim valor As Currency
Dim nomFich As String
    
  ' Quitamos blancos por los lados
  Text1(Index).Text = Trim(Text1(Index).Text)
  If Text1(Index).BackColor = vbYellow Then
    Text1(Index).BackColor = vbWhite
  End If

  Select Case Index
    
    ' Los textbox de código de empresa e instalación.
    Case 1, 2, 3, 4
      
      ' No dejamos introducir comillas en ningun campo tipo texto
      If InStr(1, Text1(Index).Text, "'") > 0 Then
        MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "Cambio de Instalación."
        Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
        PonerFoco Text1(Index)
        Exit Sub
      End If
      Text1(Index).Text = Format(Text1(Index).Text, ">")
            
      Select Case Index
                  
        ' Código de instalación.
        Case 2, 4
          Text2(Index).Text = ""
          If Text1(Index).Text <> "" Then
            If Text2(Index - 1).Text = "" Then
              MsgBox "Debe introducir el código de empresa correspondiente.", vbExclamation, "Cambio de Instalación."
              Text1(Index).Text = ""
              PonerFoco Text1(Index - 1)
              Exit Sub
            End If
            Text2(Index).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_instalacion|c_empresa|", Text1(Index).Text & "|" & Text1(Index - 1).Text & "|", "T|T|", 2)
            If Text2(Index).Text = "" And Index <> 4 Then
              MsgBox "Ese código de instalación no existe. Reintroduzca.", vbExclamation, "Cambio de Instalación."
              PonerFoco Text1(Index)
            End If
          End If
              
        ' Código de empresa.
        Case 1, 3
          Text2(Index).Text = ""
          If Text1(Index).Text <> "" Then
            Text2(Index).Text = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", Text1(Index).Text & "|", "T|", 1)
            If Text2(Index).Text = "" Then
              MsgBox "Ese código de empresa no existe. Reintroduzca.", vbExclamation, "Cambio de Instalación."
              PonerFoco Text1(Index)
            Else
              If DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_instalacion|c_empresa|", Text1(Index + 1).Text & "|" & Text1(Index).Text & "|", "T|T|", 2) = "" Then
                Text1(Index + 1).Text = ""
                Text2(Index + 1).Text = ""
              End If
              If Index = 1 And Text1(3).Text = "" Then
                Text1(3).Text = Text1(Index).Text
                Text2(3).Text = Text2(Index).Text
              End If
            End If
          End If
      End Select
              
    ' Textbox de la Clave de acceso.
    Case 0
      If Trim(Text1(Index).Text) <> Trim(Text1(Index).Tag) Then
        MsgBox "    Acceso denegado    ", vbExclamation, "Cambio de Instalación."
        Text1(Index).Text = ""
        PonerFoco Text1(Index)
      Else
        DesactivarCLAVE
        PonerFoco Text1(1)
      End If

  End Select
    
    '---
End Sub

' Configura los textbox para introducir la clave de acceso.
Private Sub ActivarCLAVE()
    
  Text1(0).Enabled = True
  Text1(1).Enabled = False
  Text1(2).Enabled = False
  Text1(3).Enabled = False
  Text1(4).Enabled = False
    
  cmdAceptar.Enabled = False
  cmdCancelar.Enabled = True

End Sub

' Configura los textbox para poder operar con la aplicación de cambio de codigo de instalación.
Private Sub DesactivarCLAVE()
    
  Text1(0).Enabled = False
  Text1(1).Enabled = True
  Text1(2).Enabled = True
  Text1(3).Enabled = True
  Text1(4).Enabled = True
    
  cmdAceptar.Enabled = True
    
End Sub

' Un SetFocus libre de errores.
Private Sub PonerFoco(ByRef Text As Object)
  On Error Resume Next
  Text.SetFocus
  If Err.Number <> 0 Then Err.Clear
End Sub

' Actualiza las tablas para el cambio de código de instalación.
Private Sub ActualizarTablas()
Dim rs As ADODB.Recordset
Dim sql As String

  ' Si existen, eliminamos las tablas temporales.
  sql = " DROP TABLE IF EXISTS temp_operainstala;"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  sql = " DROP TABLE IF EXISTS temp_instalaciones;"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
  
  ' Si ya existe una instalación con la misma fecha de alta, no se puede introducir esta nueva,
  ' así que no la creamos.
  If Not SoloCopia Then
    
    ' Creamos la tabla temporal de instalaciones.
    sql = "CREATE TEMPORARY TABLE temp_instalaciones ("
    sql = sql & "c_empresa varchar(11) ,"
    sql = sql & "c_instalacion varchar(11),"
    sql = sql & "f_alta date,"
    sql = sql & "f_baja date,"
    sql = sql & "descripcion varchar(50),"
    sql = sql & "direccion varchar(50),"
    sql = sql & "poblacion varchar(25),"
    sql = sql & "c_postal char(2),"
    sql = sql & "distrito char(3),"
    sql = sql & "telefono varchar(14),"
    sql = sql & "fax varchar(14),"
    sql = sql & "persona_contacto varchar(40),"
    sql = sql & "migrado char(2),"
    sql = sql & "rama_gen char(2),"
    sql = sql & "rama_especifica char(2),"
    sql = sql & "mail_internet varchar(30),"
    sql = sql & "observaciones varchar(80),"
    sql = sql & "c_tipo tinyint(4))"
    Conn.Execute sql
    PB1.Value = PB1.Value + 1
    Me.Refresh

    ' Pasamos a la temporal la instalación que ha sido indicada y cambiamos su
    ' código de instalación y/o de empresa en la temporal.
    sql = "insert into temp_instalaciones select * from instalaciones where c_instalacion = '"
    sql = sql & Trim(Text1(2).Text) & "' and c_empresa='" & Trim(Text1(1).Text) & "'" ' Nuevo
    Conn.Execute sql
    sql = "update temp_instalaciones set c_instalacion= '" & Trim(Text1(4).Text)
    sql = sql & "',c_empresa= '" & Trim(Text1(3).Text) & "'" ' Nuevo
    Conn.Execute sql
    PB1.Value = PB1.Value + 1
    Me.Refresh
    
    ' Traspasamos las instalaciones "modificadas" a la tabla de instalaciones.
    sql = "insert into instalaciones select * from temp_instalaciones "
    Conn.Execute sql
    PB1.Value = PB1.Value + 1
    Me.Refresh
    
  Else
    
    ' Aumentamos la barra de progreso lo correspondiente
    PB1.Value = PB1.Value + 3
    Me.Refresh
  
  End If
  
   ' Creamos la tabla temporal de operainstala.
  sql = "CREATE TEMPORARY TABLE temp_operainstala ("
  sql = sql & "c_empresa varchar(11),"
  sql = sql & "c_instalacion varchar(11),"
  sql = sql & "dni varchar(11), "
  sql = sql & "f_alta date,"
  sql = sql & "f_baja date,"
  sql = sql & "migrado char(2));"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Traspasamos a una tabla temporal todas las relaciones de operarios/instalación
  ' correspondientes a la instalación que hemos cambiado, y actualizamos también
  ' c_empresa/c_instalacion.
  sql = "INSERT INTO temp_operainstala select * from operainstala where c_instalacion= '" & Trim(Text1(2).Text) & "'"
  sql = sql & " and c_empresa='" & Text1(1).Text & "'"
  Conn.Execute sql
  sql = "UPDATE temp_operainstala set c_instalacion = '" & Trim(Text1(4).Text)
  sql = sql & "',c_empresa= '" & Trim(Text1(3).Text) & "'" ' Nuevo
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Si ya existía la "instalación destino", puede que también existan las relaciones
  ' de los operarios con las mismas.
  If Text2(4).Text <> "" Then
  
    ' Recorremos la tabla temporal de relaciones para evitar duplicar claves.
    sql = "SELECT * FROM temp_operainstala"
    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenDynamic, adLockOptimistic
  
    While Not rs.EOF
      sql = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|f_alta|", rs!c_empresa & "|" & rs!c_instalacion & "|" & rs!dni & "|" & Format(rs!f_alta, "yyyy-mm-dd") & "|", "T|T|T|F|", 4)
      
      ' No existe la clave en operainstala, insertamos la relación
      If sql = "" Then
        sql = "INSERT INTO operainstala VALUES('" & rs!c_empresa & "','" & rs!c_instalacion & "','" & rs!dni & "','" & Format(rs!f_alta, "yyyy-mm-dd") & "',"
      
        If Not IsNull(rs!f_baja) Then
          sql = sql & "'" & Format(rs!f_baja, "yyyy-mm-dd") & "',"
        Else
          sql = sql & "NULL,"
        End If
      
        If Not IsNull(rs!migrado) Then
          sql = sql & "'" & rs!migrado & "')"
        Else
          sql = sql & "NULL)"
        End If
        Conn.Execute sql
      End If
    
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
  
  Else
    
    ' Simplemente volcamos la temporal en operainstala.
    sql = "INSERT INTO operainstala SELECT * FROM temp_operainstala"
    Conn.Execute sql
  
  End If
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Cambiamos c_empresa/c_instalacion en la tabla dosimetros.
  sql = "update dosimetros set c_instalacion = '" & Trim(Text1(4).Text)
  sql = sql & "',c_empresa= '" & Trim(Text1(3).Text) ' Nuevo
  sql = sql & "' where c_instalacion = '" & Trim(Text1(2).Text) & "' and "
  sql = sql & "c_empresa='" & Trim(Text1(1).Text) & "'"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Cambiamos c_empresa/c_instalacion en la tabla dosiscuerpo.
  sql = "update dosiscuerpo set c_instalacion = '" & Trim(Text1(4).Text)
  sql = sql & "',c_empresa= '" & Trim(Text1(3).Text) ' Nuevo
  sql = sql & "' where c_instalacion = '" & Trim(Text1(2).Text) & "' and "
  sql = sql & "c_empresa='" & Trim(Text1(1).Text) & "'"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Cambiamos c_empresa/c_instalacion en la tabla dosisarea.
  sql = "update dosisarea set c_instalacion = '" & Trim(Text1(4).Text)
  sql = sql & "',c_empresa= '" & Trim(Text1(3).Text) ' Nuevo
  sql = sql & "' where c_instalacion = '" & Trim(Text1(2).Text) & "' and "
  sql = sql & "c_empresa='" & Trim(Text1(1).Text) & "'"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Cambiamos c_empresa/c_instalacion en la tabla dosisnohomog.
  sql = "update dosisnohomog set c_instalacion = '" & Trim(Text1(4).Text)
  sql = sql & "',c_empresa= '" & Trim(Text1(3).Text) ' Nuevo
  sql = sql & "' where c_instalacion = '" & Trim(Text1(2).Text) & "' and "
  sql = sql & "c_empresa='" & Trim(Text1(1).Text) & "'"
  Conn.Execute sql
    
  ' Cambiamos c_empresa/c_instalacion en la tabla recedosim.
  sql = "update recepdosim set c_instalacion = '" & Trim(Text1(4).Text)
  sql = sql & "',c_empresa= '" & Trim(Text1(3).Text) ' Nuevo
  sql = sql & "' where c_instalacion = '" & Trim(Text1(2).Text) & "' and "
  sql = sql & "c_empresa='" & Trim(Text1(1).Text) & "'"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Eliminamos las entradas con códigos anteriores en operainstala.
  sql = "delete from operainstala where c_instalacion = '" & Trim(Text1(2).Text) & "'"
  sql = sql & "and c_empresa='" & Trim(Text1(1).Text) & "'"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
  
  ' Eliminamos las entradas con códigos anteriores en instalaciones.
  sql = "delete from instalaciones where c_instalacion = '" & Trim(Text1(2).Text) & "'"
  sql = sql & "and c_empresa='" & Trim(Text1(1).Text) & "'"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  ' Nos cargamos las tablas temporales.
  sql = " DROP TABLE IF EXISTS temp_operainstala;"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
    
  sql = " DROP TABLE IF EXISTS temp_instalaciones;"
  Conn.Execute sql
  PB1.Value = PB1.Value + 1
  Me.Refresh
      
End Sub

' Comprueba que todos los datos están bien.
Private Function DatosOk() As Boolean
Dim rs As ADODB.Recordset
Dim b As Boolean
Dim sql As String, fecha As String
    
  DatosOk = True
    
  ' Todos los campos deben ser distinto de vacio.
  If Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Then
    MsgBox "Debe de introducir valor en todos los campos. Revise.", vbExclamation, "¡Error!"
    DatosOk = False
    Exit Function
  End If
    
  ' Hay que impedir el desastre que tendría lugar si al usuario se le ocurriera
  ' poner los mismos valores en c_empresa/c_instalación actual  que en
  ' c_empresa/c_instalación nuevo.
  If Text1(1).Text = Text1(3).Text And Text1(2).Text = Text1(4).Text Then
    MsgBox "No se puede realizar el cambio. Los códigos actuales y nuevos no pueden ser iguales.", vbExclamation, "¡Error!"
    DatosOk = False
    Exit Function
  End If
    
  ' Si Text2(4) <> "", la instalacion "nueva" existe.
  If Text2(4).Text <> "" Then
      
    ' Obtenemos la fecha de alta de la instalación "nueva" (es obligatorio que tenga).
    fecha = DevuelveDesdeBD(1, "f_alta", "instalaciones", "c_empresa|c_instalacion|", Text1(3).Text & "|" & Text1(4).Text & "|", "T|T|", 2)
    If fecha = "" Then
      MsgBox "La instalación " & Text1(4).Text & " no tiene fecha de alta.", vbExclamation, "¡Error!"
      DatosOk = False
      Exit Function
    End If
      
    ' Buscamos en instalaciones la clave c_empresa/c_instalacion/f_alta. Si existe
    ' informamos al usuario y ponemos "SoloCopia" a Verdadero. Esto es para que
    ' no creemos la instalación después (violaría la clave primaria).
    fecha = Format(fecha, "yyyy-mm-dd")
    sql = "SELECT COUNT(*) FROM instalaciones WHERE c_empresa='" & Text1(3).Text & "' AND "
    sql = sql & "c_instalacion='" & Text1(4).Text & "' AND f_alta='" & fecha & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenDynamic, adLockOptimistic
    sql = "0"
    If Not rs.EOF Then
      If Not IsNull(rs.Fields(0).Value) Then sql = rs.Fields(0).Value
    End If
    rs.Close
    Set rs = Nothing
    If Val(sql) > 1 Then
       
      ' Si llegamos a aquí, es porque el usuario quiere cambiar las claves
      ' c_empresa/c_instalacion/f_alta a otras que ya existen en la tabla.
      ' En principio, si la información de ambas instalaciones es la misma (es
      ' decir, que está duplicada), esto sería trasparente para el usuario...
      ' Pero ante la posibilidad de que haya alguna diferencia en la información
      ' de ambas instalaciones, ha de ser el usuario quien decida qué hacer.
      If Not (MsgBox("Ya existe una instalación con código " & Text1(4).Text & " con la misma fecha de alta. Sólo una de las dos puede existir en la base de datos, por lo que la información de la instalación " & Text1(2).Text & " se perderá, quedando únicamente la de la instalación " & Text1(4).Text & ". ¿Desea continuar?", vbExclamation + vbYesNo, "¡Error!") = vbYes) Then
        DatosOk = False
        Exit Function
      End If
        SoloCopia = True
    Else
        
      ' Sabemos que no existe la clave c_empresa/c_instalacion/f_alta, pero aún
      ' hay que contemplar que c_empresa/c_instalacion no tenga f_baja nula.
      sql = "SELECT * FROM instalaciones WHERE c_empresa='" & Text1(3).Text & "' AND c_instalacion='" & Text1(4).Text
      sql = sql & "' AND f_baja IS NULL"
      Set rs = New ADODB.Recordset
      rs.Open sql, Conn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
            
        ' En caso de que exista una c_empresa/c_instalacion sin fecha de baja,
        ' simplemente avisamos al usuario... pero dejamos que decida si quiere
        ' continuar o no.
        If Not (MsgBox("YA EXISTE una instalación sin fecha de baja con ese código (" & Text1(4).Text & "). Es aconsejable darla de baja. ¿Desea continuar a pesar de eso?", vbYesNo + vbExclamation, "¡Atención!") = vbYes) Then
          DatosOk = False
          Exit Function
        End If
        
      ' El c_empresa/c_instalacion tiene fecha de baja. Seguimos pidiendo confirmación.
      ElseIf Not (MsgBox("El código de instalación " & Text1(4).Text & " YA EXISTE en la base de datos, aunque la instalación está de baja. ¿Desea confirmar el cambio de código?", vbYesNo + vbExclamation, "¡Atención!") = vbYes) Then
        DatosOk = False
        Exit Function
      End If
      rs.Close
      Set rs = Nothing
      
    End If
  Else
      
    ' Avisamos de que no existe la instalación destino y pedimos confirmación.
    If Not (MsgBox("No existe la instalación con código " & Text1(4).Text & " para la empresa " & Text2(3).Text & ", por lo que se dará de alta. ¿Desea continuar?", vbYesNo + vbExclamation, "¡Atención!") = vbYes) Then
      DatosOk = False
      Exit Function
    End If
   
  End If
      
End Function


