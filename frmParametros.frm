VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6675
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   6675
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2520
      Top             =   3405
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   450
      Top             =   3405
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3525
      Left            =   45
      TabIndex        =   22
      Top             =   465
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   6218
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmParametros.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Internet"
      TabPicture(1)   =   "frmParametros.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame7 
         Height          =   1455
         Left            =   -74790
         TabIndex        =   38
         Top             =   435
         Width           =   6045
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   9
            Tag             =   "Direccion e-mail|T|S|||parametros|diremail|||"
            Text            =   "3"
            Top             =   300
            Width           =   4395
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   10
            Tag             =   "Servidor SMTP|T|S|||parametros|smtpHost|||"
            Text            =   "3"
            Top             =   660
            Width           =   4395
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   11
            Tag             =   "Usuario SMTP|T|S|||parametros|smtpUser|||"
            Text            =   "3"
            Top             =   1020
            Width           =   2100
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   12
            Left            =   4560
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   12
            Tag             =   "Password SMTP|T|S|||parametros|smtpPass|||"
            Text            =   "3"
            Top             =   1020
            Width           =   1230
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   20
            Left            =   165
            TabIndex        =   43
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            Height          =   195
            Index           =   21
            Left            =   165
            TabIndex        =   42
            Top             =   720
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   22
            Left            =   165
            TabIndex        =   41
            Top             =   1065
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   23
            Left            =   3705
            TabIndex        =   40
            Top             =   1035
            Width           =   840
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
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
            Index           =   0
            Left            =   360
            TabIndex        =   39
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1410
         Left            =   -74790
         TabIndex        =   33
         Top             =   1920
         Width           =   6045
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   1710
            MaxLength       =   100
            TabIndex        =   15
            Tag             =   "Web|T|S|||parametros|webversion|||"
            Text            =   "3"
            Top             =   1005
            Width           =   4110
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   1425
            MaxLength       =   100
            TabIndex        =   14
            Tag             =   "M|T|S|||parametros|mailsoporte|||"
            Text            =   "3"
            Top             =   630
            Width           =   4395
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   1425
            MaxLength       =   100
            TabIndex        =   13
            Tag             =   "W|T|S|||parametros|websoporte|||"
            Text            =   "3"
            Top             =   270
            Width           =   4395
         End
         Begin VB.Label Label8 
            Caption         =   "Soporte"
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
            Left            =   360
            TabIndex        =   37
            Top             =   0
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "Web check version"
            Height          =   195
            Index           =   16
            Left            =   150
            TabIndex        =   36
            Top             =   1020
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Mail soporte"
            Height          =   195
            Index           =   12
            Left            =   135
            TabIndex        =   35
            Top             =   675
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Web de soporte"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   34
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   0
         Tag             =   "Codigo Empresa|N|N|0|999|parametros|codempre|000|S|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1260
      End
      Begin VB.Frame Frame2 
         Height          =   2250
         Left            =   180
         TabIndex        =   23
         Top             =   1050
         Width           =   6120
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   1410
            MaxLength       =   40
            TabIndex        =   1
            Tag             =   "Nombre Empresa|T|S|||parametros|nomempre||N|"
            Text            =   "Text1"
            Top             =   255
            Width           =   3540
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   3660
            MaxLength       =   30
            TabIndex        =   5
            Tag             =   "Provincia|T|S|||parametros|proempre|||"
            Text            =   "Text1"
            Top             =   1365
            Width           =   1305
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   1410
            MaxLength       =   40
            TabIndex        =   2
            Tag             =   "Domicilio Empresa|T|S|||parametros|domempre|||"
            Text            =   "Text1"
            Top             =   630
            Width           =   3540
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   3645
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "Poblacion|T|S|||parametros|pobempre|||"
            Text            =   "Text1"
            Top             =   1005
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1410
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "C.P.Empresa|N|S|0|99999|parametros|codposta|00000||"
            Text            =   "Text1"
            Top             =   1005
            Width           =   1080
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Fax Empresa|T|S|||parametros|faxempre|||"
            Text            =   "Text1"
            Top             =   1725
            Width           =   1020
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   1425
            MaxLength       =   9
            TabIndex        =   6
            Tag             =   "CIF Empresa|T|S|||parametros|cifempre|||"
            Text            =   "Text1"
            Top             =   1365
            Width           =   1080
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   1425
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Telefono Empresa|T|S|||parametros|telempre|||"
            Text            =   "Text1"
            Top             =   1740
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre:"
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   225
            TabIndex        =   31
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "Domicilio:"
            Height          =   225
            Left            =   225
            TabIndex        =   30
            Top             =   660
            Width           =   1485
         End
         Begin VB.Label Label18 
            Caption         =   "Provincia:"
            Height          =   225
            Left            =   2745
            TabIndex        =   29
            Top             =   1410
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Fax:"
            Height          =   225
            Left            =   2760
            TabIndex        =   28
            Top             =   1770
            Width           =   435
         End
         Begin VB.Label Label20 
            Caption         =   "Poblacion:"
            Height          =   225
            Left            =   2745
            TabIndex        =   27
            Top             =   1035
            Width           =   960
         End
         Begin VB.Label Label21 
            Caption         =   "Codigo Postal:"
            Height          =   225
            Left            =   210
            TabIndex        =   26
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "CIF Empresa:"
            Height          =   225
            Left            =   210
            TabIndex        =   25
            Top             =   1395
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Teléfono:"
            Height          =   225
            Left            =   210
            TabIndex        =   24
            Top             =   1740
            Width           =   960
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Código Empresa:"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   195
         TabIndex        =   32
         Top             =   690
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   4245
      Width           =   1110
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5310
      TabIndex        =   20
      Top             =   4245
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   75
      TabIndex        =   17
      Top             =   4035
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   315
         TabIndex        =   19
         Top             =   315
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4050
      TabIndex        =   16
      Top             =   4245
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private WithEvents frmBc As frmBuscaGrid 'Conta
Attribute frmBc.VB_VarHelpID = -1

Dim Rs As ADODB.Recordset
Dim modo As Byte
Dim I As Integer
Private HaDevueltoDatos As Boolean

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim ModificaClaves As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case modo
    Case 0
        'Preparao para modificar
        PonerModo 2
        
    Case 1
        If DatosOk Then
            'Cambiamos el path
            'CambiaPath True
            If InsertarDesdeForm(Me, 1) Then
                PonerModo 0
                ActualizaNombreEmpresa
                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation, "¡Atención!"
            End If

        End If
    Case 2
        'Modificar
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            'CambiaPath True
            
            If Not (vUsu.NivelUsu < 2) Then
                ModificaClaves = True
            End If
            If ModificaClaves Then
'                If ModificaDesdeFormularioClaves(Me, cad) Then
'                    ReestableceVPARAM
'                    PonerModo 0
'                End If
'            Else
                If ModificaDesdeFormulario(Me, 1) Then PonerModo 0
                ActualizaNombreEmpresa
                AbrirConexion
                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation, "¡Atención!"
            End If
        End If
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation, "¡Error!"
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Modificar"
    PonerModo 2
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
'    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
End Sub

Private Sub cmdCancelar_Click()
    If modo = 2 Then PonerCampos
    PonerModo 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim Cad As String

        ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 4
        .Buttons(2).Image = 15
    End With
    
    Me.Left = 0
    Me.Top = 0
    SSTab1.Tab = 0
    
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd

    Cad = "select * from parametros "
    
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = Cad '"Select * from parametros"
    adodc1.Refresh
    
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 2 Then
      Toolbar1.Buttons(1).Visible = False
    End If
    
    If adodc1.Recordset.EOF Then
        'No hay datos
        'quitar###
        Limpiar Me
        
        If (vUsu.NivelUsu < 2) Then
            PonerModo 1
            
            Toolbar1.Buttons(1).Enabled = False
        Else
            MsgBox "No tiene permiso para configurar la aplicacion", vbExclamation, "¡Error!"
            Unload Me
        End If
    Else
        PonerCampos
        PonerModo 0
        'Campos que nos se tocaran los ponemos con colorcitos bonitos
'        If vUsu.Nivel <> 0 Then
'            Text1(0).BackColor = &H80000018
'            Text1(1).BackColor = &H80000018
'        End If
        Toolbar1.Buttons(1).Enabled = Not (vUsu.NivelUsu < 2)
        cmdAceptar.Enabled = Not (vUsu.NivelUsu < 2)
    End If
    
End Sub





'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
   
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
    
End Sub
'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim Cad As String
    Dim sql As String
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)

    FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
    
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 5, 6, 7, 9, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 24, 25, 26, 30
        
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
'            If Modo <> 1 And Text1(Index).Text = "" Then
'                Text1(Index).SetFocus
'                MsgBox "Este campo requiere un valor", vbExclamation
'            End If
'            If Index >= 9 Then Exit Sub
'            Text1(Index).Text = Format(Text1(Index).Text, ">")
            
        Case 2
            If (modo <> 1) Then
                If Text1(Index).Text = "" Then
                    PonerFoco Text1(Index)
                    MsgBox "Debes introducir un valor ", vbExclamation, "¡Error!"
                    
                ElseIf EsNumerico(Text1(Index).Text) Then
                           Text1(Index).Text = Format(Text1(Index).Text, "00000")
                End If
              End If
        
        '....

    End Select
    Text1(Index).Text = Format(Text1(Index).Text, ">")
    '---
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim valor As Boolean
    modo = Kmodo
    Select Case Kmodo
    Case 0
        'Preparamos para ver los datos
        valor = True
        lblIndicador.Caption = ""

    Case 1
        'Preparamos para que pueda insertar
        valor = False
        lblIndicador = "INSERTAR"
        lblIndicador.ForeColor = vbBlue
        
    Case 2
        valor = False
        lblIndicador.Caption = "MODIFICAR"
        lblIndicador.ForeColor = vbRed

    End Select
    cmdAceptar.Visible = modo > 0
    cmdCancelar.Visible = modo > 0
    
    'Ponemos los valores
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = valor
    Next I
    ' el codigo de empresa esta siempre disable
    Text1(8).Locked = Not (modo = 1)
    
    
    Dim T As Object
    
    For Each T In Me.Controls
        If TypeOf T Is ComboBox Then
            T.Locked = valor
        End If
    Next T
    
End Sub

Private Sub PonerCampos()
    Dim Cam As String
    Dim tabla As String
    Dim Cod As String
    Dim Cad As String
        If adodc1.Recordset.EOF Then Exit Sub
        If PonerCamposForma(Me, adodc1) Then
        End If
End Sub
'
Private Function DatosOk() As Boolean
    Dim b As Boolean
    Dim J As Integer
    Dim Mens As String
    
    DatosOk = False
    b = CompForm(Me)
    If Not b Then Exit Function
    
    
    DatosOk = b
    
NoExisteDirectorio:
        
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        'Modificar
         PonerModo 2
        '
    Case 2
        'Salir
        Unload Me
    End Select
End Sub

Private Sub ReestableceVPARAM()
    Set vParam = Nothing
    Set vParam = New Cparametros
    vParam.Leer
End Sub

Private Sub PonerFoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ActualizaNombreEmpresa()
Dim sql As String

    Set Conn = Nothing
    
    If AbrirConexionUsuarios Then
        sql = "update empresadosis set nomempre = '"
        sql = sql & DevNombreSQL(Text1(0).Text) & "' where codempre = "
        sql = sql & Val(Text1(8).Text)
    
        Conn.Execute sql
    Else
        MsgBox "No se pudo acceder a usuarios.", vbCritical, "¡Error!"
        End
    End If
    
End Sub

