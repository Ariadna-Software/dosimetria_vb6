VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarnomempre As String
Private mvarcifempre As String
Private mvardomempre As String
Private mvarpobempre As String
Private mvarcpostalempre As String
Private mvarnumempre As Byte   'copia local
Private mvarmesesdosis As Byte
Private mvardosissup As Integer
Private mvardosisprof As Integer

Private sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As Recordset
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo Cparametros
'------------------------------------------------

Public Property Let NombreEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarnomempre = vData
End Property

Public Property Get NombreEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     NombreEmpresa = mvarnomempre
End Property

Public Property Let NifEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarcifempre = vData
End Property

Public Property Get NifEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     NifEmpresa = mvarcifempre
End Property


Public Property Let NumeroEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarnumempre = vData
End Property

Public Property Get NumeroEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     NumeroEmpresa = mvarnumempre
End Property
Public Property Let DomicilioEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvardomempre = vData
End Property

Public Property Get DomicilioEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     DomicilioEmpresa = mvardomempre
End Property

Public Property Let PoblacionEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarpobempre = vData
End Property

Public Property Get PoblacionEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     PoblacionEmpresa = mvarpobempre
End Property

Public Property Let CPostalEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarcpostalempre = vData
End Property

Public Property Get CPostalEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     CPostalEmpresa = mvarcpostalempre
End Property

Public Property Get MesesDosis() As Byte
  MesesDosis = mvarmesesdosis
End Property
Public Property Let MesesDosis(newValue As Byte)
  mvarmesesdosis = newValue
End Property

Public Property Get DosisSuperficial() As Integer
  DosisSuperficial = mvardosissup
End Property
Public Property Let DosisSuperficial(newValue As Integer)
  mvardosissup = newValue
End Property

Public Property Get DosisProfunda() As Integer
  DosisProfunda = mvardosisprof
End Property
Public Property Let DosisProfunda(newValue As Integer)
  mvardosisprof = newValue
End Property

'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------



''****   AGREGAR  ***
'Public Function Agregar() As Byte
'    On Error GoTo Err1
'    SQL = "SELECT * FROM parametros"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " fechaini = " & mvarfechaini
'   ' Set Rs = DB.OpenRecordSet(Sql, dbOpenDynaset)
'    If Rs.RecordCount > 0 Then
'        Mens = "Ya existe un Registro en la Tabla parametros con esos valores."
'        MsgBox Mens, vbExclamation
'        Agregar = 1
'        'ELSE
'        Else
'            Rs.AddNew
'            Rs!fechaini = mvarfechaini
'            Rs!fechafin = mvarfechafin
'            Rs!autocoste = mvarautocoste
'            Rs!emitedia = mvaremitedia
'            Rs!contabil = mvarcontabil
'            Rs!listahco = mvarlistahco
'            Rs!numdiapr = mvarnumdiapr
'            Rs!concefpr = mvarconcefpr
'            Rs!conceapr = mvarconceapr
'            Rs!numdiacl = mvarnumdiacl
'            Rs!concefcl = mvarconcefcl
'            Rs!conceacl = mvarconceacl
'            Rs!limimpcl = mvarlimimpcl
'            'RS!conpresu = mvarconpresu
'            Rs!periodos = mvarperiodos
'            Rs!grupogto = mvargrupogto
'            Rs!grupovta = mvargrupovta
'            Rs!ctaperga = mvarctaperga
'            Rs!abononeg = mvarabononeg
'            Rs!amortiza = mvaramortiza
'            Rs!intfaccl = mvarintfaccl
'            Rs!intfacpr = mvarintfacpr
'            Rs!intliapu = mvarintliapu
'            Rs!intlihap = mvarintlihap
'            Rs!grupoord = mvargrupoord
'            Rs!tinumfac = mvartinumfac
'            Rs!modhcofa = mvarmodhcofa
'            Rs!anofactu = mvaranofactu
'            Rs!perfactu = mvarperfactu
'            Rs!nctafact = mvarnctafact
'            Rs!AsienActAuto = mvarAsienActAuto
'            Rs.Update
'            Agregar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err1:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Rs.Close
'    Agregar = 1
'End Function


''****   ELIMINAR  ***
'Public Function Eliminar() As Byte
'    On Error GoTo Err2
'    SQL = "SELECT * FROM parametros"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " fechaini = " & mvarfechaini
' '   Set Rs = DB.OpenRecordSet(Sql, dbOpenDynaset)
'    If Rs.RecordCount = 0 Then
'        Mens = "No existe un Registro en la Tabla parametros con esos valores."
'        MsgBox Mens, vbExclamation
'        Eliminar = 1
'        'ELSE
'        Else
'            Rs.Delete
'            Rs.MoveNext
'            If Rs.EOF Then Rs.MoveFirst
'            Eliminar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err2:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Rs.Close
'    Eliminar = 1
'End Function


'****   LEER  ***
Public Function Leer() As Byte
    On Error GoTo Err3
    sql = "SELECT * FROM parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockOptimistic
    If Rs.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarnomempre = DBLet(Rs!nomempre)
            mvardomempre = DBLet(Rs!domempre)
            mvarpobempre = DBLet(Rs!pobempre)
            mvarcpostalempre = Rs!codposta
            mvarcifempre = DBLet(Rs!cifempre)
            mvarnumempre = Rs!codempre
            Leer = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation, "¡Error!"
    Rs.Close
    Leer = 1
End Function
