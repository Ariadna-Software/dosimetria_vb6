VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Soporta entidades financieras usadas en los ingresos"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales que contienen valores de propiedad
Private mvarCodigo As Long 'copia local
Private mvarCodusu As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarPasswd As String 'copia local
Private mvarPasswdPROPIO As String 'copia local
Private mvarlogin As String 'copia local
Private mvarDirFich As String 'copia local
Private mvarNivelCon As Long 'nivel de contabilidad
Private mvarNivelUsu As Long 'nivel de usuario
Private mvarNivelGes As Long 'nivel gestion social
Private mvarCadenaConexion As String
Private mvarPC As String

Private Rs As ADODB.Recordset
Private sql As String
Private Mens As String


Public Property Let NivelUsu(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarNivelUsu = vData
End Property


Public Property Get NivelUsu() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    NivelUsu = mvarNivelUsu
End Property


Public Property Let Passwd(ByVal vData As String)
Attribute Passwd.VB_Description = "Password de usuario"
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Passwd = 5
    mvarPasswd = vData
End Property


Public Property Get Passwd() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Passwd
    Passwd = mvarPasswd
End Property



Public Property Let CadenaConexion(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Passwd = 5
    mvarCadenaConexion = vData
End Property


Public Property Get CadenaConexion() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Passwd
    CadenaConexion = mvarCadenaConexion
End Property



Public Property Let PasswdPROPIO(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Passwd = 5
    mvarPasswdPROPIO = vData
End Property


Public Property Get PasswdPROPIO() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Passwd
    PasswdPROPIO = mvarPasswdPROPIO
End Property


Public Property Let nombre(ByVal vData As String)
Attribute nombre.VB_Description = "Nombre del grupo de formas de cobro."
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarNombre = vData
End Property


Public Property Get nombre() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    nombre = mvarNombre
End Property

'---------------------------
'Dirfich
Public Property Let Dirfich(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarDirFich = vData
End Property


Public Property Get Dirfich() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Dirfich = mvarDirFich
End Property

'Login
Public Property Let Login(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarlogin = vData
End Property


Public Property Get Login() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Login = mvarlogin
End Property

Public Property Let codigo(ByVal vData As Long)
Attribute codigo.VB_Description = "Código del grupo de formas de cobro"
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Codigo= 5
    mvarCodigo = vData
End Property


Public Property Get codigo() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    codigo = mvarCodigo
End Property

Public Property Let Codusu(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Codigo= 5
    mvarCodusu = vData
End Property


Public Property Get Codusu() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    Codusu = mvarCodusu
End Property


Public Property Get PC() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Codigo
    PC = mvarPC
End Property

Public Function Leer(ByVal vlogin As String) As Byte
Attribute Leer.VB_Description = "Lee un grupo de formas de cobro en función del código que se le ha pasado. Si devuelve 0 se ha conseguido leer el dato."
Dim I As Integer
    On Error GoTo Err1
    sql = "SELECT * FROM usuarios WHERE login = '" & vlogin & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Rs.EOF Then
        Leer = 1
    Else
        I = DevuelveAumentoPC
        mvarCodigo = Val(DBLet(Rs!Codusu)) + I
        mvarCodusu = Val(DBLet(Rs!Codusu))
        mvarNombre = DBLet(Rs!nomusu)
        mvarNivelUsu = Val(DBLet(Rs!NivelUsu))
        
        mvarPasswdPROPIO = Rs!passwordpropio
        mvarlogin = vlogin
        mvarCadenaConexion = ""  'Se le asigna en tiempo de diseño
        Leer = 0
    End If
    Rs.Close
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation, "¡Error!"
    Rs.Close
    Leer = 1
End Function


Private Function DevuelveAumentoPC() As Integer

DevuelveAumentoPC = 32000
mvarPC = ComputerName
sql = DevuelveDesdeBD(1, "codpc", "pcs", "nompc|", mvarPC & "|", "T|", 1)
If sql <> "" Then DevuelveAumentoPC = Val(sql) * 1000

End Function
