VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Propiedades privadas.
Private pClave As String
Private pValor As String

' Acceso a las propiedades.
Public Property Get clave() As String
  clave = pClave
End Property
Public Property Let clave(newValue As String)
  pClave = newValue
End Property

Public Property Get valor() As String
  valor = pValor
End Property
Public Property Let valor(newValue As String)
  pValor = newValue
End Property

'------------------------------------------------
'FUNCIONES del modulo CConfiguracion
'------------------------------------------------

'****   AGREGAR  ***
Public Function Agregar() As Boolean
Dim sql As String
On Error GoTo EAgregar

  sql = "INSERT INTO configuracion (clave, valor) VALUES ( '" & pClave
  sql = sql & "' ,'" & pValor & "')"
   
  Conn.Execute sql
  Agregar = True
    
  Exit Function

EAgregar:
  Agregar = False
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Function

'****   LEER  ***
Public Function Leer() As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String
On Error GoTo ELeer
    
  sql = "SELECT * FROM configuracion where clave = '" & pClave & "'"
  Set Rs = New ADODB.Recordset
  Rs.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

  If Not Rs.EOF Then
    pValor = Rs!valor
    Leer = True
  Else
    Leer = False
  End If

  Rs.Close
  Set Rs = Nothing
  Exit Function

ELeer:
  Set Rs = Nothing
  Leer = False
  Err.Raise Err.Number, Err.Source, Err.Description

End Function


'****   MODIFICAR  ***
Public Function Modificar() As Boolean
Dim sql As String
On Error GoTo EModificar

    sql = "UPDATE configuracion SET valor = '" & pValor
    sql = sql & "' where clave = '" & pClave & "'"
    Conn.Execute sql
    Modificar = True
    Exit Function

EModificar:

  Modificar = False
  Err.Raise Err.Number, Err.Source, Err.Description
    
End Function
