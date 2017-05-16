VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodempre As String 'copia local
Private mvarnomempre As String 'copia local
Private mvarnomresum As String 'copia local

Private sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset
Private RC As Byte


'------------------------------------------------
'Propiedades del modulo Cempresa
'------------------------------------------------


Public Property Let codempre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codempre = 5
     mvarcodempre = vData
End Property

Public Property Get codempre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codempre
     codempre = mvarcodempre
End Property


Public Property Let nomempre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nomempre = 5
     mvarnomempre = vData
End Property

Public Property Get nomempre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     nomempre = mvarnomempre
End Property


Public Property Let nomresum(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nomresum = 5
     mvarnomresum = vData
End Property

Public Property Get nomresum() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomresum
     nomresum = mvarnomresum
End Property



'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------



'****   AGREGAR  ***
Public Function Agregar() As Byte
    On Error GoTo Err1

    sql = "INSERT INTO empresa (codempre, nomempre, nomresum) VALUES ("

    'Datos
    sql = sql & mvarcodempre
    sql = sql & ",'" & mvarnomempre
    sql = sql & "','" & mvarnomresum
    sql = sql & "')"
    
    Conn.Execute sql
    
    Exit Function
Err1:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation, "¡Error!"
    Rs.Close
    Agregar = 1
End Function




'****   LEER  ***
Public Function Leer() As Byte
    On Error GoTo Err3
    sql = "SELECT * FROM empresa"
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If Rs.EOF Then
        Leer = 1
        'ELSE
        Else
            Mens = "Datos basicos"
            mvarcodempre = Rs!codempre
            mvarnomempre = Rs!nomempre
            mvarnomresum = DBLet(Rs!nomresum)
            Leer = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation, "¡Error!"
    Set Rs = Nothing
    Leer = 1
End Function


'****   MODIFICAR  ***
Public Function Modificar() As Byte
    On Error GoTo Err4


    sql = "UPDATE empresa SET "
    sql = sql & " nomempre = '" & mvarnomempre
    sql = sql & "', nomresum = '" & mvarnomresum & "'"
    sql = sql & " WHERE codempre =" & mvarcodempre & ";"
    
    Conn.Execute sql
    
    
    'Modificamos en la tabla Usuarios.empresa
    sql = "UPDATE empresas SET nomempre= '" & mvarnomempre
    sql = sql & "', nomresum= '" & mvarnomresum
    sql = sql & "' WHERE codempre=" & mvarcodempre
    Conn.Execute sql
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation, "¡Error!"
    Rs.Close
    Modificar = 1
End Function
