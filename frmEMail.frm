VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar E-MAIL"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   2940
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5715
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   2055
         Index           =   3
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Para"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   600
         Picture         =   "frmEMail.frx":030A
         ToolTipText     =   "Buscar contacto"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   5715
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Otro"
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   15
         Top             =   1140
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Error"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sugerencia"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1140
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Mensaje"
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
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
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
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enviar e-Mail Ariadna Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmEMail.frx":040C
      Top             =   3840
      Width           =   480
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0 - Envio del PDF
    '1 - Envio Mail desde menu soporte
    '2 -
    '3 - Reclamaciones via E-MAIL
        'Valores en CadenaDesdeOtroForm
Public MisDatos As String
    'Nombre para|email para|Asunto|Mensaje|

Public Asunto As String
Public NombreMail As String
Public Mail As String
Public Listado As Integer
    
    
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Dim Cad As String
Dim PrimeraVez As Boolean
Dim DatosDelMailEnUsuario As String

Private Sub Enviar()
    Dim imageContentID, success
    Dim mailman As ChilkatMailMan
    Dim Valores As String
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante, es la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
    mailman.LogMailSentFilename = "" ' App.Path & "\mailSent.log"

    
    'Servidor smtp
    Valores = DatosDelMailEnUsuario  'Empipado: smtphost,smtpuser, pass, diremail
    If Valores = "" Then
        MsgBox "Falta configurar en paremtros la opcion de envio mail(servidor, usuario, clave)"
        Exit Sub
    End If
    mailman.SmtpAuthMethod = "NONE"
    mailman.SmtpHost = RecuperaValor(Valores, 2) ' vParam.SmtpHOST
    mailman.SmtpUsername = RecuperaValor(Valores, 3) 'vParam.SmtpUser
    mailman.SmtpPassword = RecuperaValor(Valores, 4) 'vParam.SmtpPass
    
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    'Si es de SOPORTE
'    If Opcion = 1 Then
'         'Obtenemos la pagina web de los parametros
'        cad = DevuelveDesdeBD("mailsoporte", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
'        If cad = "" Then
'            MsgBox "Falta configurar en parametros el mail de soporte", vbExclamation
'            Exit Sub
'        End If
'
'        If cad = "" Then GoTo GotException
'        email.AddTo "Soporte Tesoreria", cad
'        cad = "Soporte Ariconta. "
'        If Option1(0).Value Then cad = cad & Option1(0).Caption
'        If Option1(1).Value Then cad = cad & Option1(1).Caption
'        If Option1(2).Value Then cad = cad & "Otro: " & Text2.Text
'        email.Subject = cad
'
'        'Ahora en text1(3).text generaremos nuestro mensaje
'        cad = "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf
'        cad = cad & "Hora: " & Format(Now, "hh:mm") & vbCrLf
'        cad = cad & "TESORERIA:  " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
'        cad = cad & "Usuario: " & vUsu.nombre & vbCrLf
'        cad = cad & "Nivel USU: " & vUsu.Nivel & vbCrLf
'        cad = cad & "Empresa: " & vEmpresa.nomempre & vbCrLf
'        cad = cad & "&nbsp;<hr>"
'        cad = cad & Text3.Text & vbCrLf & vbCrLf
'        Text1(3).Text = cad
'    Else
        'Envio de mensajes normal
        email.AddTo Text1(0).Text, Text1(1).Text
        email.Subject = Text1(2).Text
'    End If
    
    'El resto lo hacemos comun
    'La imagen
    imageContentID = email.AddRelatedContent(App.Path & "\minilogo.bmp")
    
    
    Cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    Cad = Cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    Cad = Cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    Cad = Cad & "<TR><TD VALIGN=""TOP""><P>"
    FijarTextoMensaje
    Cad = Cad & "</P></TD></TR>"
    Cad = Cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
    'La imagen
    Cad = Cad & "<P ALIGN=""CENTER""><IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & "></P>"
    Cad = Cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa " & App.EXEName & " de "
    Cad = Cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
    Cad = Cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
    Cad = Cad & "<P>Este correo electr�nico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
    Cad = Cad & " los destinatarios especificados. La informaci�n contenida puesde ser CONFIDENCIAL"
    Cad = Cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    Cad = Cad & "<P>Si usted recibe este mensaje por ERROR, por favor comun�queselo inmediatamente al"
    Cad = Cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelaci�n, distribuci�n"
    Cad = Cad & " impresi�n o copia de toda o alguna parte de la informaci�n contenida, Gracias "
    Cad = Cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    Cad = Cad & "</TR></TABLE></BODY></HTML>"
    
    email.SetHtmlBody (Cad)
    
    
    
    email.AddPlainTextAlternativeBody "Programa e-mail NO soporta HTML. " & vbCrLf & Text1(3).Text
    email.From = RecuperaValor(Valores, 1) 'vParam.diremail
    
    If Opcion = 0 Or Opcion = 3 Then
        'ADjunatmos el PDF
        email.AddFileAttachment App.Path & "\docum.pdf"
    End If
        
    
    'email.SendEncrypted = 1
    success = mailman.SendEmail(email)
    If (success = 1) Then
        If Opcion <> 3 Then
            Cad = "Mensaje enviado correctamente."
            MsgBox Cad, vbInformation
            Command2(0).SetFocus
        End If
    Else
        Cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.Path & "\log.xml"
        MsgBox Cad, vbExclamation
    End If
    
    
GotException:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        NumRegElim = 1
    Else
        NumRegElim = 0
    End If
    Set email = Nothing
    Set mailman = Nothing

End Sub

Private Sub Command1_Click()
    If Not DatosOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    Image2.Visible = True
    Me.Refresh
    Enviar
    Image2.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 3 Then
            HacerMultiEnvio
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Image2.Visible = False
    Limpiar Me
    Frame1(0).Visible = (Opcion = 0) Or (Opcion = 3)
    Frame1(1).Visible = (Opcion = 1)
    If Opcion = 1 Then HabilitarText
    Me.Icon = frmPpal.Icon
    PonDisponibilidadEmail
    Me.Command1.Enabled = (DatosDelMailEnUsuario <> "")
    If Opcion = 3 Then
        'Si es masivo na de na
        Command2(0).Enabled = False
        Me.Command1.Enabled = False
        
    End If
    Text1(0).Text = NombreMail
    Text1(1).Text = Mail
    Text1(2).Text = Asunto


End Sub




Private Sub PonDisponibilidadEmail()
    
    Cad = "" 'DevuelveDesdeBD("dirfich", "Usuarios.usuarios", "codusu", (vUsu.codigo Mod 100), "N")
    If Cad = "" Then
        'Primero compruebo si los datos los tengo en el usuario
        Cad = "select diremail,smtphost,smtpuser,smtppass from parametros"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux!SmtpHost) Then
                For NumRegElim = 0 To miRsAux.Fields.Count - 1
                    Cad = Cad & DBLet(miRsAux.Fields(NumRegElim), "T") & "|"
                Next NumRegElim
            End If
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        
    End If
    DatosDelMailEnUsuario = Cad
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Opcion = 0
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)

    Screen.MousePointer = vbHourglass
    Text1(0).Tag = RecuperaValor(CadenaSeleccion, 1)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 2)
    'Si regresa con datos tengo k devolveer desde la bd el campo e-mail
    Cad = DevuelveDesdeBD(1, "mail_internet", "instalaciones", "c_instalacion|", Text1(0).Tag)
    Text1(1).Text = Cad
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Text1(0).Text = RecuperaValor(CadenaDevuelta, 1)
    Text1(1).Text = RecuperaValor(CadenaDevuelta, 3)
End Sub

Private Sub Image1_Click()
    MandaBusquedaPrevia ""
'    Set frmB = New frmBuscaGrid
'    frmB.DatosADevolverBusqueda = "0|1"
'    'frmA.ConfigurarBalances = 5  'NUEVO opcion
'    frmB.Show
'    Set frmB = Nothing
'    If Text1(0).Text <> "" Then Text1(2).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
    HabilitarText
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Function DatosOk() As Boolean
Dim I As Integer

    DatosOk = False
    If Opcion = 0 Then
                'Pocas cosas a comprobar
                For I = 0 To 2
                    Text1(I).Text = Trim(Text1(I).Text)
                    If Text1(I).Text = "" Then
                        MsgBox "El campo: " & Label1(I).Caption & " no puede estar vacio.", vbExclamation
                        Exit Function
                    End If
                Next I
                
                'EL del mail tiene k tener la arroba @
                I = InStr(1, Text1(1).Text, "@")
                If I = 0 Then
                    MsgBox "Direccion e-mail erronea", vbExclamation
                    Exit Function
                End If
    Else
        Text2.Text = Trim(Text2.Text)
        'SOPORTE
        If Trim(Text3.Text) = "" Then
            MsgBox "El mensaje no puede ir en blanco", vbExclamation
            Exit Function
        End If
        If Option1(2).Value Then
            If Text2.Text = "" Then
                MsgBox "El campo 'OTRO asunto' no puede ir en blanco", vbExclamation
                Exit Function
            End If
        End If
    End If
      
    'Llegados aqui OK
    DatosOk = True
        
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then Exit Sub 'Si estamos en el de datos nos salimos
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'El procedimiento servira para ir buscando los vbcrlf y cambiarlos por </p><p>
Private Sub FijarTextoMensaje()
Dim I As Integer
Dim J As Integer

    J = 1
    Do
        I = InStr(J, Text1(3).Text, vbCrLf)
        If I > 0 Then
              Cad = Cad & Mid(Text1(3).Text, J, I - J) & "</P><P>"
        Else
            Cad = Cad & Mid(Text1(3).Text, J)
        End If
        J = I + 2
    Loop Until I = 0
End Sub

Private Sub HabilitarText()
    If Option1(2).Value Then
        Text2.Enabled = True
        Text2.BackColor = vbWhite
    Else
        Text2.Enabled = False
        Text2.BackColor = &H80000018
    End If
End Sub



'Private Function RecuperarDatosEMAILAriadna() As Boolean
'Dim NF As Integer
'
'    RecuperarDatosEMAILAriadna = False
'    NF = FreeFile
'    Open App.Path & "\soporte.dat" For Input As #NF
'    Line Input #NF, cad
'    Close #NF
'    If cad <> "" Then RecuperarDatosEMAILAriadna = True
'
'End Function


'Private Function ObtenerValoresEnvioMail() As String
'    ObtenerValoresEnvioMail = ""
'    Set miRsAux = New ADODB.Recordset
'    cad = "Select diremail,SmtpHost, SmtpUser, SmtpPass  from parametros where"
'    cad = cad & " fechaini='" & Format(vParam.fechaini, FormatoFecha) & "';"
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not miRsAux.EOF Then
'        cad = DBLet(miRsAux!SmtpHost)
'        cad = cad & "|" & DBLet(miRsAux!SmtpUser)
'        cad = cad & "|" & DBLet(miRsAux!SmtpPass)
'        cad = cad & "|" & DBLet(miRsAux!diremail) & "|"
'        ObtenerValoresEnvioMail = cad
'    End If
'    miRsAux.Close
'    Set miRsAux = Nothing
'End Function

Private Sub HacerMultiEnvio()
Dim Rs As ADODB.Recordset
Dim Cont As Integer
Dim I As Integer
    Cad = "Select ztempemail.* from ztempemail where codusu =" & vUsu.codigo & " and fichero <> ''"
    
    'Campos comunes
    'ENVIO MASIVO DE EMAILS
    Text1(2).Text = RecuperaValor(MisDatos, 1)
    Text1(3).Text = RecuperaValor(MisDatos, 2)
    
    Me.Refresh
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText

    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.MoveFirst
    I = 1
    Me.Refresh
    While Not Rs.EOF
        Screen.MousePointer = vbHourglass
        Text1(0).Text = ""
        Text1(0).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_instalacion|", Trim(Rs!c_instalacion) & "|", "T|", 1)
        Text1(1).Text = Rs!email
        Caption = "Enviar E-MAIL (" & I & " de " & Cont & ")"
        Me.Refresh
        
        'De momento volvemos a copiar el archivo como docum.pdf
        FileCopy App.Path & "\temp\A" & Rs!c_instalacion & ".pdf", App.Path & "\docum.pdf"
        Me.Refresh
        NumRegElim = 0
        Enviar
        
        
        If NumRegElim = 1 Then
            'NO SE HA ENVIADO.
            Cad = "UPDATE ztempemail SET  fichero ='' WHERE codusu =" & vUsu.codigo & " AND c_instalacion = '" & Trim(Rs!c_instalacion) & "'"
            Conn.Execute Cad
        End If
        'Siguiente
        Rs.MoveNext
        I = I + 1
    Wend
    Rs.Close
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim HaDevueltoDatos As Boolean
        'Llamamos a al form
        '##A mano
        Cad = ""
'        Cad = Cad & ParaGrid(Text1(0), 10, "C�digo")
'        Cad = Cad & ParaGrid(Text1(1), 60, "Nombre")
'        Cad = Cad & ParaGrid(Text1(3), 60, "Correo Electr�nico")
                
        
        Cad = "Nombre|descripcion|T|50�C�digo|c_instalacion|N|10�Correo El�ctronico|mail_internet|T|60�"
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = "instalaciones"
            frmB.vSql = ""
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Instalaciones"
            frmB.vSelElem = 0
            frmB.vConexionGrid = 1
            'frmB.vBuscaPrevia = chkVistaPrevia
            frmB.vCargaFrame = False
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
'            If HaDevueltoDatos Then
'                MsgBox "hola"
'
''                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                    cmdRegresar_Click
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
''                Text1(kCampo).SetFocus
'            End If
        End If
        Screen.MousePointer = vbDefault

End Sub



