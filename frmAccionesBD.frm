VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAccionesBD2 
   Caption         =   "Conjunto de acciones sobre BD"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "frmAccionesBD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   2835
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   6195
      Begin VB.CommandButton Command3 
         Caption         =   "Recuperar Backup"
         Height          =   1335
         Index           =   0
         Left            =   1380
         TabIndex        =   7
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Una tabla BACKUP"
         Height          =   1335
         Index           =   1
         Left            =   3195
         TabIndex        =   6
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Index           =   2
         X1              =   960
         X2              =   5940
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Insertar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "BackUP"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "INSERTAR"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Utilidades"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar Datos"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   4440
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4140
      Width           =   6075
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Empresa Primera del sector."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6360
   End
End
Attribute VB_Name = "frmAccionesBD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public BD_Intercambio As String

Public Opcion As Byte
    '0 ---->  CONTABILIDAD
    '1 ---->  TESORERIA

Private nombre As String
Dim Rs As ADODB.Recordset
Dim NumTablas As Integer
Dim Contador As Long
Dim tamanyo As Long
Dim sql As String
Dim NivelAnterior As Byte  'Para el aumento de nivel

Dim Campos() As String

Public FormularioOK As String

'-------------------------------------
'Abrir conexion CNN
Private Function AbrirConexion(Usuario As String, Pass As String, Conta As String) As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    AbrirConexion = False
    Set Cnn = Nothing
    Set Cnn = New Connection
    'Conn.CursorLocation = adUseClient
    Cnn.CursorLocation = adUseServer
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & Conta & ";SERVER=" & vConfig.SERVER & ";"
    Cad = Cad & ";UID=" & Usuario
    Cad = Cad & ";PWD=" & Pass
    
    Cnn.ConnectionString = Cad
    Cnn.Open
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión " & Conta, Err.Description
End Function


Private Sub Command3_Click(Index As Integer)
Dim salir As Boolean
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim pos As Integer
Dim I As Byte
Select Case Index
Case 0
    'Preguntas previas
    If MsgBox("Se dispone a restaurar todas las tablas de la base de datos. Se perderá cualquier cambio en los datos desde la fecha de esta copia. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") <> vbYes Then Exit Sub
    
    'Vamos a obtener la carpeta donde estaran los archivos
    If PedirCarpeta Then
        'Importar datos desde un BACKUP
        If MsgBox("La carpeta seleccionada es " & nombre & "." & vbCrLf & "Esta acción es irreversible. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") <> vbYes Then Exit Sub
        RecuperarBACKUP
    End If


Case 1
        
  If MsgBox("Restaurar una única tabla puede causar inconsistencia en la base de datos a causa de las referencias entre las tablas. Debe estar muy seguro/a de lo que está haciendo, pues esta acción es irreversible. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") <> vbYes Then Exit Sub
  If PedirNombreArchivo Then
    sql = "Ha seleccionado: " & nombre & vbCrLf & " ¿Desea continuar?"
    If MsgBox(sql, vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") <> vbYes Then Exit Sub
    If nombre <> "" Then UnSoloBACKUP
  End If
End Select
    
End Sub


Private Sub Form_Load()
Dim OK As Boolean

    OK = Opcion = 1  'SOLO TESORERIA
    With Toolbar1
        .ImageList = frmPpal.ImgListComun
        .Buttons(3).Image = 3
        .Buttons(7).Image = 15
    End With
    Caption = "Conjunto de acciones sobre BD. ("
    If Opcion = 0 Then
        Caption = Caption & "DOSIMETRIA"
        Label2.ForeColor = &H8000&
    End If
    Caption = Caption & ")"
    
    
    
    
    FrameFalse
    PB1.Visible = False
    OK = AbrirConexion(vConfig.User, vConfig.password, "mbgstld4") 'mbgstld4
    Command3(0).Enabled = OK
    Command3(1).Enabled = OK
    Label3.Caption = ""
End Sub


Private Sub PonerProgress(Incremento As Integer)
    Contador = Contador + Incremento
    If Contador < tamanyo Then
        PB1.Value = CInt((Contador / tamanyo) * 100)
    Else
        PB1.Value = 100
    End If
End Sub




Private Function PedirNombreArchivo() As Boolean
    On Error GoTo EP
    PedirNombreArchivo = False
    cd1.CancelError = True
    'Nombre del archivo carpeta
    cd1.DialogTitle = "Nombre del archivo que contiene la tabla a restaurar"
    cd1.InitDir = App.Path & "\backup"
    cd1.ShowOpen
    If cd1.FileTitle = "" Then Exit Function
    nombre = cd1.FileName
    PedirNombreArchivo = True
EP:
End Function

Private Function PedirCarpeta() As Boolean
Dim salir As Boolean

    PedirCarpeta = False

    nombre = GetFolder("Seleccione la carpeta que contiene la copia:")
    
    If nombre <> "" Then PedirCarpeta = True

End Function





''-------------------------------------------
'Private Sub BKTablas(Tabla As String, Optional NombreArchivo As String)
'Dim Cad As String
'Dim NF As Integer
'Dim Izquierda As String
'Dim Derecha As String
'
'    On Error GoTo EBKTablas
'    pb1.Value = 0
'    Label3.Caption = Tabla
'    Me.Refresh
'
'    'Tamanyo
'    Set Rs = New ADODB.Recordset
'    Rs.Open "Select count(*) from " & Tabla, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Tamanyo = 0
'    If Not Rs.EOF Then
'        If Not IsNull(Rs.Fields(0)) Then Tamanyo = Rs.Fields(0)
'    End If
'    Rs.Close
'
'
'    Rs.Open Tabla, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdTable
'    If Rs.EOF Then
'        'No hace falta hacer back up
'
'    Else
'        pb1.Visible = True
'
'        NF = FreeFile
'        If NombreArchivo = "" Then
'            Open Nombre & "\" & Tabla & ".sql" For Output As #NF
'        Else
'            Open NombreArchivo For Output As #NF
'        End If
'
'        BACKUP_TablaIzquierda Rs, Izquierda
'        Contador = 0
'        While Not Rs.EOF
'            Contador = Contador + 1
'            pb1.Value = CInt((Contador / Tamanyo) * 100)
'
'            BACKUP_Tabla Rs, Derecha
'            Cad = "INSERT INTO " & Tabla & " " & Izquierda & " VALUES " & Derecha & ";"
'            Print #NF, Cad
'            Rs.MoveNext
'        Wend
'    End If
'    Rs.Close
'
'EBKTablas:
'    If Err.Number Then
'        If NombreArchivo = "" Then
'            'Es un backup masivo
'            'Llevamos a errores
'            InsertaError Tabla & vbCrLf & Cad, Err.Description
'        Else
'            MsgBox Err.Description & vbCrLf & vbCrLf & Cad, vbExclamation
'        End If
'    End If
'    If Contador > 0 Then Close #NF
'    Set Rs = Nothing
'End Sub
'


Private Sub EliminarDatosTMP()
    sql = "SHOW TABLES"
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    tamanyo = 0
    While Not Rs.EOF
        If Mid(Rs.Fields(0), 1, 3) = "tmp" Then tamanyo = tamanyo + 1
        Rs.MoveNext
    Wend
    Rs.Close
    If tamanyo = 0 Then Exit Sub
    PB1.Visible = True
    Label3.Caption = "Eliminando datos temporales"
    espera 0.1
    Rs.Open sql, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Contador = 0
    While Not Rs.EOF
        If Mid(Rs.Fields(0), 1, 3) = "tmp" Then
            Label3.Caption = Rs.Fields(0)
            Label3.Refresh
            Contador = Contador + 1
            PB1.Value = CInt((Contador / tamanyo) * 100)
            Cnn.Execute "Delete from " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
End Sub



Private Sub ProcesarFichero()
Dim PrimeraVez As Boolean
Dim INC As Integer
Dim f As Integer

On Error GoTo EProcesarFichero
    'El nombre del fichero estara en NOMBRE
    f = FreeFile
    sql = nombre 'Para el error
    tamanyo = FileLen(nombre)
    Contador = 0
    If tamanyo = 0 Then Exit Sub
    PrimeraVez = True
    PB1.Value = 0
    Open nombre For Input As #f
    While Not EOF(f)
        Line Input #f, sql
        If PrimeraVez Then
            INC = Len(sql)
            PrimeraVez = False
        End If
        sql = Trim(sql)
        If sql <> "" Then EjecutaSQL sql
        PonerProgress INC
    Wend
    Close #f
    Exit Sub
EProcesarFichero:
    InsertaError sql, Err.Description
    Err.Clear
End Sub


Private Sub EjecutaSQL(ByRef CADENA As String)
    On Error Resume Next
    Cnn.Execute CADENA
    DoEvents
    If Err.Number <> 0 Then
        InsertaError CADENA, Err.Description
        Err.Clear
    End If
End Sub


Private Sub UnSoloBACKUP()
    PB1.Value = 0
    PB1.Visible = True
    
    ConfigBotones False
    
    'Abrir archvio errores
    IncializaErrores Label3.Caption & vbCrLf & "Restaurando BACKUP"
    
    ' añadido borra la tabla
    BorraTabla nombre
    
    ProcesarFichero
    PB1.Visible = False
    Label3.Caption = ""
    
    'Vemos los errores
    'Vemos si tiene errores
    nombre = TieneErrores
    If nombre <> "" Then
        sql = "Se han producido errores. " & vbCrLf
        sql = sql & "Se ha generado el archivo de error: " & vbCrLf & nombre
        sql = sql & vbCrLf & vbCrLf & vbCrLf & "          ¿Desea verlo ahora?"
        If MsgBox(sql, vbQuestion + vbYesNo, "¡Error!") = vbYes Then
            Screen.MousePointer = vbHourglass
            VolcarFicheroError sql
            frmErrores.Text1 = sql
            frmErrores.Show vbModal
        End If
    Else
      MsgBox "Copia restaurada con éxito.", vbInformation
    End If
    ConfigBotones True
    Screen.MousePointer = vbDefault

End Sub


Private Sub RecuperarBACKUP()
Dim Archivos As String
Dim Numero As Integer
Dim Carpeta As String
Dim I As Integer
    
    ConfigBotones False
    
    'Compruebo k todos los archivos son .sql
    tamanyo = 0  'Los SQL
    Contador = 0 'NO .SQL
    sql = Dir(nombre & "\*.*")
    Do While sql <> ""
        If LCase(Right(sql, 4)) = ".sql" Then
            tamanyo = tamanyo + 1
        Else
            Contador = Contador + 1
        End If
        sql = Dir
    Loop
    
    If tamanyo = 0 Then
        sql = "Ningun archivo .sql en la carpeta"
    Else
        If Contador = 0 Then
            sql = "" 'Para k siga
        Else
            sql = "Hay archivos que no son .sql"
        End If
    End If
    
    If sql <> "" Then
        MsgBox sql, vbExclamation, , "¡Error!"
        Exit Sub
    End If
    
    
    Carpeta = nombre & "\" 'Ya k lugo nombre la utilizo
    
    'AHORA EMPEZAREMOS A METER LOS ARCHIVOS
    '---------------------------------------
    '
    ' TEngo una cadena k tendra los nombres de los archivos(sin la extension .sql)
    ' Separados por PIPES y una variable Contadora para recuperarvalor
    
    'Tengo una linea con los archivos y el orden k deben seguir
    Archivos = "ramagene|provincias|tipmedext|parametros|configuracion|"
    Archivos = Archivos & "ramaespe|tipostrab|"
    Archivos = Archivos & "fondos|fondospana|factcali4400|factcali6600|lotes|lotespana|"
    'EMpresa parametros
    Archivos = Archivos & "usuarios|usuariosempresadosis|empresadosis|pcs|rangoscsn|"
    Archivos = Archivos & "empresas|instalaciones|operarios|operainstala|"
    'Amortizacion
    Archivos = Archivos & "dosimetros|dosisarea|dosiscuerpo|dosisnohomog|"
    'Apuntes
    Archivos = Archivos & "erroresmigra|recepdosim|tempnc|"
    
    'Numero = 30 '52 - 3  'PQ esta comentada una linea hay arriba.. AHI AHI AHI
    
    PB1.Value = 0
    PB1.Visible = True
    
    'borro la bd y la vuelvo a crear
    nombre = App.Path & "\informes\Batch Tablas.sql"
    'Numero = 35
    ProcesarFichero

    Numero = 29
    
    'Abrir archivo errores
    IncializaErrores Label3.Caption & vbCrLf & "Restaurando BACKUP"
    For I = 1 To Numero
        Screen.MousePointer = vbHourglass
        sql = RecuperaValor(Archivos, I)
        
        'Indicadores
        Label3.Caption = sql & " (" & I & " de " & Numero & ")"
        Me.Refresh
        
        
        
        nombre = Carpeta & sql & ".sql"
        If Dir(nombre, vbArchive) <> "" Then
            BorraTabla nombre  ' añadido
            ProcesarFichero
        End If
    Next I
    PB1.Visible = False
    Label3.Caption = ""
    
    'Vemos los errores
    'Vemos si tiene errores
    nombre = TieneErrores
    If nombre <> "" Then
        sql = "Se han producido errores. " & vbCrLf
        sql = sql & "Se ha generado el archivo de error: " & vbCrLf & nombre
        sql = sql & vbCrLf & vbCrLf & vbCrLf & "          ¿Desea verlo ahora?"
        If MsgBox(sql, vbQuestion + vbYesNo, "¡Error!") = vbYes Then
            Screen.MousePointer = vbHourglass
            VolcarFicheroError sql
            frmErrores.Text1 = sql
            frmErrores.Show vbModal
        End If
    Else
      MsgBox "Copia restaurada con éxito.", vbInformation
    End If
    ConfigBotones True
    Screen.MousePointer = vbDefault
End Sub

Private Sub FrameFalse()
    Frame3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not Frame3.Enabled Then
    Cancel = -1
  Else
    Set Cnn = Nothing
    frmPpal.Visible = True
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    FrameFalse
    If Opcion = 1 And Button.Index > 1 Then Exit Sub
    Select Case Button.Index
    Case 3
        Frame3.Visible = True
    Case 7
        Unload Me
    End Select
End Sub


Private Function ComprobarOk(ByRef vNivelAnterior As Byte) As Boolean
Dim vE As String
Dim UltimoNivel As Byte
    On Error GoTo EComprobarOk
    ComprobarOk = False
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '
    'Comprobamos k las tablas siguientes NO tiene registros
    '
    '
    sql = "cabfacte|cabfactprove|linapu|linapue|"  '4
    vE = ""
    NumTablas = 1
    Set Rs = New ADODB.Recordset
    Do
        nombre = RecuperaValor(sql, NumTablas)
        nombre = "Select count(*) from " & nombre
        Rs.Open nombre, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then
                If Rs.Fields(0) > 0 Then vE = vE & RecuperaValor(sql, NumTablas) & vbCrLf
            End If
        End If
        Rs.Close
        NumTablas = NumTablas + 1
    Loop Until NumTablas > 4
    
    If vE <> "" Then
        sql = "Las siguientes tablas tienen datos y deberian estar vacias" & vbCrLf
        sql = sql & vE
        MsgBox sql, vbExclamation, "¡Error!"
        Exit Function
    End If
    'Comprobamos k el ultimo nivel no es 10
    Rs.Open "empresa", Cnn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    vE = ""
    If Rs.EOF Then
        vE = "No esta definida la empresa."
    Else
        UltimoNivel = DBLet(Rs.Fields(3), "N")
        If UltimoNivel = 0 Then
            vE = "Si definir ultimo nivel contable"
        Else
            NumTablas = DBLet(Rs.Fields(3 + UltimoNivel), "N")
            If NumTablas = 0 Then
                vE = "Ultimo nivel es 0. Datos incorrectos"
            Else
                If NumTablas = 10 Then
                    vE = "No se puede ampliar el ultimo nivel. Ya es 10"
                Else
                    'Fale vamos a devolver el nivel anterior al ultimo
                    vNivelAnterior = CByte(DBLet(Rs.Fields(3 + UltimoNivel - 1)))
                    If vNivelAnterior < 3 Or vNivelAnterior > 10 Then vE = "Error obteniendo nivel anterior"
                End If
            End If
        End If
    End If
    Rs.Close
    If vE <> "" Then
        MsgBox vE, vbExclamation, "¡Error!"
        Exit Function
    End If
    ComprobarOk = True
    Exit Function
EComprobarOk:
    MuestraError Err.Number, "ComprobarOk." & Err.Description
End Function


Private Sub DatosTabla(ByRef Rs As ADODB.Recordset, ByRef Derecha As String)
Dim I As Integer
Dim nexo As String
Dim valor As String
Dim Tipo As Integer
    Derecha = ""
    nexo = ""
    For I = 0 To Rs.Fields.Count - 1
        Tipo = Rs.Fields(I).Type
        
        If IsNull(Rs.Fields(I)) Then
            valor = "NULL"
        Else
        
            'pruebas
            Select Case Tipo
            'TEXTO
            Case 129, 200, 201
                valor = Rs.Fields(I)
                NombreSQL valor
                'Si el campo es el codmacta o apudirec lo cambiamos
                valor = "'" & valor & "'"
            'Fecha
            Case 133
                valor = CStr(Rs.Fields(I))
                valor = "'" & Format(valor, "yyyy-mm-dd") & "'"
                
            'Numero normal, sin decimales
            Case 2, 3, 16 To 19
                valor = Rs.Fields(I)
            
            'Numero con decimales
            Case 131
                valor = CStr(Rs.Fields(I))
                valor = TransformaComasPuntos(valor)
            Case Else
                valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                valor = valor & vbCrLf & "SQL: " & Rs.Source
                valor = valor & vbCrLf & "Pos: " & I
                valor = valor & vbCrLf & "Campo: " & Rs.Fields(I).Name
                valor = valor & vbCrLf & "Valor: " & Rs.Fields(I)
                MsgBox valor, vbExclamation, "¡Error!"
                MsgBox "El programa finalizara. Avise al soporte técnico.", vbCritical
                End
            End Select
                        
        End If
        Derecha = Derecha & nexo & valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
End Sub

Private Function CambiaTabla(tabla As String, vCampos As String, NCampos As Integer)
Dim I As Integer

    ReDim Campos(NCampos)
    
    For I = 1 To NCampos
        Campos(I) = RecuperaValor(vCampos, I)
    Next I
    
    Label3.Caption = tabla
    PB1.Value = 0
    Me.Refresh
    CambiaValores tabla, NCampos

End Function




Private Function CambiaValores(tabla As String, numCta As Integer)
Dim sql As String
Dim Cad As String
Dim I As Integer
    Cad = ""
    sql = ""
    On Error GoTo ECambia
    
    For I = 1 To numCta
        'Para bonito
        Label3.Caption = tabla & " (" & I & " de " & numCta & ")"
        PB1.Value = 0
        Me.Refresh
        tamanyo = 0
        'Contador  COUNT(distinct(codmacta))
        sql = "SELECT COUNT(DISTINCT(" & Campos(I) & ")) from " & tabla
        Rs.Open sql, Cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then tamanyo = DBLet(Rs.Fields(0), "N")
        Rs.Close
        

        If tamanyo > 0 Then
            'Updateamos la primera cta
            tamanyo = tamanyo + 1
            sql = "SELECT " & Campos(I) & " FROM " & tabla & " GROUP BY " & Campos(I)
            Rs.Open sql, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Contador = 0
            While Not Rs.EOF
                Contador = Contador + 1
                PonerProgressBar Val((Contador / tamanyo) * 1000)
                If Not IsNull(Rs.Fields(0)) Then
                    Cad = Rs.Fields(0)
                    sql = "UPDATE " & tabla & " SET " & Campos(I) & " = '" & Cad & "'"
                    sql = sql & " WHERE " & Campos(I) & " = '" & Rs.Fields(0) & "'"
                    Cnn.Execute sql
                End If
                'Sig
                Rs.MoveNext
            Wend
            Rs.Close
        End If
    Next I
    Exit Function
ECambia:
    MuestraError Err.Number, Err.Description
End Function

Private Sub PonerProgressBar(valor As Long)
    If valor <= 1000 Then PB1.Value = valor
End Sub


Private Sub BorraTabla(nombre As String)
Dim tabla As String

    tabla = Dir(nombre, vbArchive)
    tabla = Mid(tabla, 1, Len(tabla) - 4)
    Cnn.Execute "set foreign_key_checks=0"
    Cnn.Execute "delete from " & tabla
    Cnn.Execute "set foreign_key_checks=1"

End Sub

Private Sub ConfigBotones(Modo As Boolean)
  Frame3.Enabled = Modo
  Toolbar1.Enabled = Modo
End Sub

