Attribute VB_Name = "bus"
Option Explicit
' una cosita
'Public DB As Database
Public conn As ADODB.Connection
Public Cnn As ADODB.Connection

Public FormularioHijoModificado As Boolean  'Para cuando llama a un form hijo que puede add or update
Public CadenaDevueltaFormHijo As String
Public vUsu As Usuario  'Datos usuario

Public FormatoFecha As String
Public FormatoImporte As String
Public FormatoHora As String
Public CartaSobredosis As Boolean

Public vParam As Cparametros
Public vConfig As Configuracion
Public CadenaDesdeOtroForm As String
Public vEmpresa As Cempresa 'Los datos de la empresa
Public codEmpresaActual As Integer
Public LineaBlanca As String

Public miRsAux As ADODB.Recordset
Public AnchoLogin As String  'Para fijar los anchos de columna

       
Private Sub Main()
       Load frmIdentifica
       CadenaDesdeOtroForm = ""
       
       'Necesitaremos el archivo arifon.dat
       frmIdentifica.Show vbModal
        
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set conn = Nothing
            End
       End If
       
       CadenaDesdeOtroForm = ""
       frmLogin.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado ninguna empresa
            Set conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass

        'Cerramos la conexion
        conn.Close

        
        If AbrirConexion() = False Then
            MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical, "¡Error!"
            End
        End If
        
        LeerParametros
        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
        GestionaPC
        
        'Otras acciones
        OtrasAcciones
         
        Screen.MousePointer = vbDefault
        frmPpal.Show
        
End Sub

Public Function LeerParametros()
        'Abrimos la empresa
        Set vParam = New Cparametros
        If vParam.Leer() = 1 Then
            MsgBox "No se han podido cargar los parámetros. Debe configurar la aplicación.", vbExclamation, "¡Error!"
            Set vParam = Nothing
        End If
End Function

Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
                       
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    Cad = Cad & ";Persist Security Info=true"
    
    
    conn.ConnectionString = Cad
    conn.Open
    conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión.", Err.Description
End Function




'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "," & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = CADENA
End Function

'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function

'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & ":" & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosHoras = CADENA
End Function

Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"
                    DBLet = ""
                Case "N"
                    DBLet = 0
                Case "F"
                    DBLet = "0:00:00"
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

Public Sub MuestraError(Numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        Cad = Cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If conn.Errors.Count > 0 Then
        ControlamosError Aux
        conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then Cad = Cad & "Número: " & Numero & vbCrLf & "Descripción: " & error(Numero)
    MsgBox Cad, vbExclamation, "¡Error!"
End Sub

Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function

' 1 - suministros
' 2 - contabilidad
' 3 - gestion social
'--
Public Function DevuelveDesdeBD(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional num As Byte, Optional ByRef OtroCampo As String, Optional ByVal orden As String) As String
Dim rs As Recordset
Dim Cad As String
Dim Aux As String
Dim v_aux As Integer
Dim Campo As String
Dim valor As String
Dim tip As String

On Error GoTo EDevuelveDesdeBD
DevuelveDesdeBD = ""

Cad = "Select " & kCampo
If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
Cad = Cad & " FROM " & Ktabla

If Kcodigo <> "" Then Cad = Cad & " where "

For v_aux = 1 To num
    Campo = RecuperaValor(Kcodigo, v_aux)
    valor = RecuperaValor(ValorCodigo, v_aux)
    tip = RecuperaValor(Tipo, v_aux)
        
    If Left(valor, 2) <> ">=" And Left(valor, 2) <> "<=" Then
      Cad = Cad & Campo & "="
    Else
      Cad = Cad & Campo
    End If
    If tip = "" Then Tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                Cad = Cad & valor
            Case "T", "F"
                Cad = Cad & "'" & valor & "'"
            Case Else
                MsgBox "Tipo : " & tip & " no definido", vbExclamation, "¡Error!"
            Exit Function
    End Select
    
    If v_aux < num Then Cad = Cad & " AND "
  Next v_aux

'Creamos el sql
If orden <> "" Then Cad = Cad & " " & orden
Set rs = New ADODB.Recordset
Select Case kBD
    Case 1
        rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Select

If Not rs.EOF Then
    DevuelveDesdeBD = DBLet(rs.Fields(0))
    If OtroCampo <> "" Then OtroCampo = DBLet(rs.Fields(1))
Else
     If OtroCampo <> "" Then OtroCampo = ""
End If
rs.Close
Set rs = Nothing
Exit Function
EDevuelveDesdeBD:
    MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function

Public Function EsNumerico(Texto As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim Cad As String
    
    EsNumerico = False
    Cad = ""
    If Not IsNumeric(Texto) Then
        Cad = "El campo debe ser numérico"
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, Texto, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then Cad = "Numero de puntos incorrecto"
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, Texto, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then Cad = "Numero incorrecto"
        End If
        
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation, "¡Error!"
    Else
        EsNumerico = True
    End If
End Function
'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256.98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

If Importe = "" Then
    ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
End If
End Function

Public Function ImporteSinFormato(CADENA As String) As String
Dim I As Integer
'Quitamos puntos
Do
    I = InStr(1, CADENA, ".")
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Mid(CADENA, I + 1)
Loop Until I = 0
ImporteSinFormato = TransformaPuntosComas(CADENA)
End Function

Public Function EsFechaOK(ByRef T As TextBox) As Boolean
Dim Cad As String
    
    Cad = T.Text
    If InStr(1, Cad, "/") = 0 Then
        If Len(T.Text) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T.Text) = 6 Then
                Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5, 2)
            End If
        End If
    End If
    If IsDate(Cad) Then
        EsFechaOK = True
        T.Text = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOK = False
    End If
End Function



Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
        If Len(T) = 6 Then
             If Mid(Cad, 5, 2) > "20" Then
                Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & "19" & Mid(Cad, 5, 2)
             Else
                Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & "20" & Mid(Cad, 5, 2)
             End If
        End If
        If Len(T) = 8 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
    End If
    If IsDate(Cad) Then
        EsFechaOKString = True
        T = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function


Public Function ComprobarEmpresaBloqueada(Codusu As Integer, ByRef Empresa As String) As Boolean
Dim Cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
conn.Execute "Delete from vBloqBD where codusu=" & Codusu

'Ahora comprobamos k nadie bloquea la BD
Cad = DevuelveDesdeBD(1, "codusu", "vBloqBD", "conta|", Empresa & "|", "T|", 1)
If Cad <> "" Then
    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
    
    Set miRsAux = New ADODB.Recordset
    Cad = "show processlist"
    miRsAux.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(3) = Empresa Then
            Cad = miRsAux.Fields(2)
            miRsAux.MoveLast
        End If
    
        'Siguiente
        miRsAux.MoveNext
    Wend
    
    If Cad = "" Then
        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
        conn.Execute "Delete from Usuario.vBloqBD where conta ='" & Empresa & "'"
        
    Else
        MsgBox "BD bloqueada.", vbCritical, "¡Error!"
        ComprobarEmpresaBloqueada = True
    End If
End If

conn.Execute "commit"
End Function

Public Function AbrirConexionUsuarios() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionUsuarios = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer
    Cad = "DSN=vMbgstld4;DESC=MySQL ODBC 3.51 Driver DSN;"
    'Cad = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & vConfig.SERVER & ";DATABASE=mbgstld4;"
    'Cad = Cad & "UID=" & vConfig.User & ";PWD=" & vConfig.password & ";OPTION=3" ';PORT=3306;STMT=;"
    conn.ConnectionString = Cad
    conn.Open
    AbrirConexionUsuarios = True
    Exit Function

EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión usuarios.", Err.Description
End Function
'/////////////////////////////////////////////////////////////////
'// Se trata de identificar el PC en la BD. Asi conseguiremos tener
'// los nombres de los PC para poder asignarles un codigo
'// UNa vez asignado el codigo  se lo sumaremos (x 1000) al codusu
'// con lo cual el usuario sera distinto( aunque sea con el mismo codigo de entrada)
'// dependiendo desde k PC trabaje

Private Sub GestionaPC()
CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    FormatoFecha = DevuelveDesdeBD(1, "codpc", "pcs", "nompc|", CadenaDesdeOtroForm & "|", "T|", 1)
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        If NumRegElim > 30 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte técnico.", vbCritical, "¡Error!"
            End
        End If
        FormatoFecha = "INSERT INTO pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        conn.Execute FormatoFecha
    End If
End If
End Sub

Private Sub OtrasAcciones()
On Error Resume Next

    'Esto estara en el MAIN
    FormatoFecha = "yyyy-mm-dd"
    FormatoImporte = "#,###,###,##0.00"
    FormatoHora = "yyyy-mm-dd hh:mm:ss"
    LineaBlanca = "                                                                                                                     "
    'Borramos uno de los archivos temporales
    'If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.codigo

    
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub

Public Function EsUnaIP(ByVal equipo As String) As Boolean
Dim I As Integer
Dim cont As Integer
Dim Cad As String
Dim esIP As Boolean

  cont = 0
  esIP = True
  I = InStr(1, equipo, ".")
  
  While I <> 0 And esIP
    cont = cont + 1
    
    Cad = Left(equipo, I - 1)
    If Not IsNumeric(Cad) Then
      esIP = False
    ElseIf Val(Cad) < 0 Or Val(Cad) > 255 Then
      esIP = False
    End If
    
    equipo = Mid(equipo, I + 1)
    I = InStr(1, equipo, ".")
  Wend
    
  If cont <> 3 Or Not esIP Then
    esIP = False
  Else
    If Not IsNumeric(equipo) Then
      esIP = False
    ElseIf Val(equipo) < 0 Or Val(equipo) > 255 Then
      esIP = False
    End If
  End If
    
  EsUnaIP = esIP
  
End Function

Public Function OtrosPCsContraContabiliad() As String
Dim MiRS As Recordset
Dim Cad As String
Dim equipo As String
Dim IP As String
Dim num As Integer

    Set MiRS = New ADODB.Recordset
    Cad = "show processlist"
    MiRS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vUsu.CadenaConexion Then
            equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, equipo, ":")
            If NumRegElim > 0 Then equipo = Mid(equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            If Not EsUnaIP(equipo) Then
              NumRegElim = InStr(1, equipo, ".")
              If NumRegElim > 0 Then
                equipo = Mid(equipo, 1, NumRegElim - 1)
                equipo = UCase(equipo)
              End If
            Else
              If RecuperarIP = equipo Then equipo = vUsu.PC
            End If
            
            If equipo <> vUsu.PC Then
                If MiRS.Fields(2) <> "localhost" Then
                    If MiRS.Fields(2) <> "LOCALHOST" Then
                        If InStr(1, Cad, equipo & "|") = 0 Then Cad = Cad & equipo & "|"
                    End If
                End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = Cad
End Function


Public Function DevNombreSQL(CADENA As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = CADENA
End Function

Public Function ConCaracteresBusqueda(T As String) As Boolean
Dim Cad As String
    
    Cad = T
    ConCaracteresBusqueda = False
    If InStr(1, Cad, ">") <> 0 Or InStr(1, Cad, "<") Or _
       InStr(1, Cad, ">=") <> 0 Or InStr(1, Cad, "<=") Or _
       InStr(1, Cad, ":") <> 0 Or InStr(1, Cad, "<<") Or _
       InStr(1, Cad, ">>") <> 0 Or InStr(1, Cad, "?") Or _
       InStr(1, Cad, "*") Or InStr(1, Cad, "null") Or InStr(1, Cad, "NULL") Then
       ConCaracteresBusqueda = True
    End If
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
End Sub
