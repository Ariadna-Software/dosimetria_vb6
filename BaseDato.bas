Attribute VB_Name = "BaseDato"
Option Explicit


Public Const CadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Conta\BDatos2.mdb;Persist Security Info=False"
'Public Conn As Connection
Public BaseForm As Integer 'Tendremos la base donde mostrar
                            'los formularios
Public NumRegElim As Long
Public MostrarErrores As Boolean
Public mConfig As Configuracion
Public Directorio As String ' es el directorio que me guardo de la migracion

' declaracion de las constantes utilizadas en los mantenimientos como opciones de menu
Public Const vbPrimero = 36
Public Const vbUltimo = 35
Public Const vbAnterior = 33
Public Const vbSiguiente = 34
Public Const vbAñadir = 65
Public Const vbEliminar = 69
Public Const vbModificar = 77
Public Const vbImprimir = 73
Public Const vbLineas = 76
Public Const vbRecepcion = 82
Public Const vbSalir = 83
Public Const vbBuscar = 66
Public Const vbVerTodos = 86
Public Const vbESC = 27

Private sql As String

Dim ImpD As Currency
Dim ImpH As Currency
Dim RT As ADODB.Recordset


Dim d As String
Dim H As String
'Para los balances
Dim M1 As Integer   ' años y meses para el balance
Dim M2 As Integer
Dim M3 As Integer
Dim A1 As Integer
Dim A2 As Integer
Dim A3 As Integer
Dim vCta As String
Dim ImAcD As Currency  'importes
Dim ImAcH As Currency
Dim ImPerD As Currency  'importes
Dim ImPerH As Currency
Dim ImCierrD As Currency  'importes
Dim ImCierrH As Currency

Dim Aux As String
Dim vFecha1 As Date
Dim vFecha2 As Date
Dim VFecha3 As Date
Dim codigo As String
Dim EjerciciosCerrados As Boolean
Dim NumAsiento As Integer
Public NoExistenDatos As Boolean

'--------------------------------------------------------------------
'--------------------------------------------------------------------
Private Function ImporteASQL(ByRef Importe As Currency) As String
ImporteASQL = ","
If Importe = 0 Then
    ImporteASQL = ImporteASQL & "NULL"
Else
    ImporteASQL = ImporteASQL & TransformaComasPuntos(CStr(Importe))
End If
End Function

'--------------------------------------------------------------------
'--------------------------------------------------------------------



Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
Dim Cad As String
Dim Aux As String
Dim Ch As String
Dim FIN As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
Cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    I = CararacteresCorrectos(CADENA, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = "(" & Campo & " >= " & Cad & ") AND (" & Campo & " <= " & Aux & ")"
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    FIN = False
                    I = 1
                    Cad = ""
                    Aux = "NO ES NUMERO"
                    While Not FIN
                        Ch = Mid(CADENA, I, 1)
                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                            Cad = Cad & Ch
                            Else
                                Aux = Mid(CADENA, I)
                                FIN = True
                        End If
                        I = I + 1
                        If I > Len(CADENA) Then FIN = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Cad = "" Then Cad = " = "
                    DevSQL = "(" & Campo & " " & Cad & " " & Aux & ")" ' aqui
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        'antes estaba esto, lo he cambiado por lo de abajo
        'If Not IsDate(Cad) Or Not IsDate(Aux) Then Exit Function  'No son numeros
        If Not EsFechaOKString(Cad) Or Not EsFechaOKString(Aux) Then Exit Function  'No son numeros
        
        'Intervalo correcto
        'Construimos la cadena
        Cad = Format(Cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = "(" & Campo & " >='" & Cad & "')  AND (" & Campo & " <= '" & Aux & "')" ' aqui
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                FIN = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not FIN
                    Ch = Mid(CADENA, I, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        Cad = Cad & Ch
                        Else
                            Aux = Mid(CADENA, I)
                            FIN = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then FIN = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If Cad = "" Then Cad = " = "
                DevSQL = "(" & Campo & " " & Cad & " " & Aux & ")" ' aqui
            End If
        End If
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(CADENA, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    'Cambiamos el * por % puesto que en ADO es el caraacter para like
    I = 1
    Aux = CADENA
    While I <> 0
        I = InStr(1, Aux, "*")
        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
    Wend
    'Cambiamos el ? por la _ pue es su omonimo
    I = 1
    While I <> 0
        I = InStr(1, Aux, "?")
        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
    Wend
    Cad = Mid(CADENA, 1, 2)
    If Cad = "<>" Then
        Aux = Mid(Aux, 3)
        DevSQL = "(" & Campo & " LIKE '!" & Aux & "')" 'aqui
        Else
        DevSQL = "(" & Campo & " LIKE '" & Aux & "')" 'aqui
    End If
    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, CADENA, "<>")
    If I = 0 Then
        'IGUAL A valor
        Cad = " = "
        Else
            'Distinto a valor
        Cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, CADENA, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = "(" & Campo & " " & Cad & " " & Aux & ")" ' aqui
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
Dim I As Integer
Dim Ch As String
Dim error As Boolean

CararacteresCorrectos = 1
error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "=", "."
            Case Else
                error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "a" To "z"
            Case "A" To "Z"
            Case "0" To "9"
            Case ".", "," ' Añadidos porque sus códigos contienen estos caracteres.
            Case "*", "%", " ", "-", "?", "_", "\", "/", ":", "'" ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case Else
                error = True
                Exit For
        End Select
    Next I
Case "F"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                error = True
                Exit For
        End Select
    Next I
Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not error Then CararacteresCorrectos = 0
End Function

Public Function BloqueoManual(Bloquear As Boolean, tabla As String, clave As String) As Boolean
    If Bloquear Then
        sql = "INSERT INTO zbloqueos (codusu, tabla, clave) VALUES (" & vUsu.codigo
        sql = sql & ",'" & UCase(tabla) & "','" & UCase(clave) & "')"
    Else
        sql = "DELETE FROM zbloqueos where codusu = " & vUsu.codigo & " AND tabla ='"
        sql = sql & DevNombreSQL(tabla) & "'"
    End If
    On Error Resume Next
    Conn.Execute sql
    If Err.Number <> 0 Then
        Err.Clear
        BloqueoManual = False
    Else
        BloqueoManual = True
    End If
End Function

'Public Function NdigUltnivel() As Integer
'Dim Cad As String
'Dim UltNivel As Integer
'
'    Cad = ""
'    Cad = DevuelveDesdeBD(2, "numnivel", "empresa", "codempre|", vParam.NumeroContabilidad & "|", "N|", 1)
'
'    UltNivel = CInt(Cad)
'
'    Cad = ""
'    Cad = DevuelveDesdeBD(2, "numdigi" & CStr(UltNivel), "empresa", "codempre|", vParam.NumeroContabilidad & "|", "N|", 1)
'
'    NdigUltnivel = CInt(Cad)
'
'End Function
'

'Public Function EsDeApuntesDirectos(Cuenta As String) As Boolean
'
'    EsDeApuntesDirectos = (Len(Cuenta) = NdigUltnivel)
'
'End Function

'Public Function GenerarAlbaran(Pedido As Long, FechaAlb As Date, Albaran As String, Tipodtos As Byte) As String
'    Dim SQL As String
'    Dim Sql1 As String
'    Dim Sql2 As String
'    Dim Importe As Currency
'    Dim Diferencia As Currency
'    Dim mC As Contadores
'    Dim Numlinea As Integer
'    Dim proveedor As Long
'    Dim RT As ADODB.Recordset
'    Dim RT1 As ADODB.Recordset
'    Dim Categoria As String
'    Dim fechainv As String
'    Dim horainve As String
'    Dim b As Boolean
'    Dim CMov As CMovimientos
'
'    On Error GoTo Etmpconext
'
'    GenerarAlbaran = "-1"
'
'    'Conn.Execute "Delete from tmpconext where codusu =" & vUsu.Codigo
'
'    Conn.BeginTrans
'
'
''    Set mC = New Contadores
''    mC.ConseguirContador "1", 4, True
'
'    Set RT = New ADODB.Recordset
'    Set RT1 = New ADODB.Recordset
'
'    SQL = "Select * from scappr where  numpedpr=" & Pedido
'    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    proveedor = RT.Fields(2).Value
'    SQL = "INSERT INTO scaalp (numalbpr,fecalbar,codprove, codforpa,dtognral,observac,numpedpr,fecpedpr)"
'    SQL = SQL & " VALUES ('" & DevNombreSQL(Trim(Albaran)) & "','" & Format(FechaAlb, FormatoFecha) & "'," 'VRS:1.0.1(12)
'    SQL = SQL & RT.Fields(2).Value & "," & RT.Fields(3).Value & "," & TransformaComasPuntos(RT.Fields(4))
'    If IsNull(RT.Fields(5).Value) Then
'        SQL = SQL & "," & ValorNulo & ","
'    Else
'        SQL = SQL & ",'" & DevNombreSQL(RT.Fields(5).Value) & "',"
'    End If
'
'    SQL = SQL & Pedido & ",'" & Format(RT!fecpedpr, FormatoFecha) & "')"
'
'
'    Conn.Execute SQL
'    RT.Close
'
'    SQL = "select * from slippr where numpedpr= " & Pedido & " and "
'    SQL = SQL & "recibida is not null order by numlinea"
'    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    'If Not RT.EOF Then RT.MoveFirst
'    Numlinea = 0
'    While Not RT.EOF
'        Numlinea = Numlinea + 1
'
'        Importe = Round2(RT!recibida * RT!Precioar, 2)
'        If Tipodtos = 0 Then
'            Importe = Round2(Importe - (Importe * RT!dtoline1 / 100) - (Importe * RT!dtoline2 / 100), 2)
'        Else
'            Importe = Importe - (Importe * RT!dtoline1 / 100)
'            Importe = Importe - (Importe * RT!dtoline2 / 100)
'            Importe = Round2(Importe, 2)
'        End If
'
'        SQL = "insert into slialp (numalbpr, codprove, numlinea, codartic, regfitosanitario,"
'        SQL = SQL & " cantidad, precioar, dtoline1, dtoline2, implinea, ampliaci, numpedpr) "
'        SQL = SQL & " VALUES ('" & DevNombreSQL(Trim(Albaran)) & "', " & proveedor & "," & Numlinea & "," 'VRS:1.0.1(12)
'        SQL = SQL & RT!codArtic & ",'" & DevNombreSQL(RT!Regfitosanitario) & "'," & TransformaComasPuntos(RT!recibida) & ","
'        SQL = SQL & TransformaComasPuntos(RT!Precioar) & "," & TransformaComasPuntos(RT!dtoline1) & "," & TransformaComasPuntos(RT!dtoline2) & ","
'        SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ",'" & DevNombreSQL(RT!ampliaci) & "'," & Pedido & ")"
'
'        Conn.Execute SQL
'
'        ' actualizamos el stock de sartic
'        b = ModDatosCompraArt(CStr(RT!codArtic), CStr(RT.Fields(4).Value), CStr(Importe), CStr(RT!recibida), "0", CStr(FechaAlb), CStr(RT!Precioar))  'VRS:1.0.5(6) añadido el precio
'
'        'VRS:1.0.4(14)
'        If b Then
'            Set CMov = New CMovimientos
'
'            CMov.almacen = 0
'            CMov.articulo = CStr(RT!codArtic)
'            'cantidad
'            CMov.cantidad = RT!recibida
'            'importe movimiento
'            CMov.ImporteM = Importe
'            CMov.DetaMov = "ALB"
'            CMov.ConnBD = Conn
'            CMov.Documento = Albaran
'            CMov.Linea = Numlinea
'            CMov.ClieProv = proveedor
'            CMov.Fechamov = CDate(FechaAlb)
'            CMov.HoraMov = CDate(Format(FechaAlb, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"))
'            CMov.LetraSerie = ""
'            CMov.InsertarMovimiento
'
'            Set CMov = Nothing
'        End If
'
'
'        ' si es fitosanitario y cambia de registro hay que insertarlo
'        ' en el hco de regfito
'        If EsArticuloFitosanitario(CLng(RT!codArtic)) Then
'            CambioRegistroFitosanitario CStr(RT!codArtic), CStr(RT!Regfitosanitario), CStr(Albaran), CStr(proveedor), CStr(FechaAlb), CStr(Numlinea)
'
'        End If
'
'
'        ' insertamos lo que hay en la temporal tmplotes del usuario, en la tabla de
'        ' movimientos slotes.
'        SQL = "select * from tmplotes "
'
'        Sql2 = "where codusu = " & vUsu.codigo
'        Sql2 = Sql2 & " and numpedid = " & Pedido
'        Sql2 = Sql2 & " and numlinea = " & RT!Numlinea
'
'        SQL = SQL & Sql2
'
'        RT1.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RT1.EOF Then RT1.MoveFirst
'        While Not RT1.EOF
'            Categoria = DevuelveDesdeBD(1, "categoria", "sartic", "codartic|", RT!codArtic & "|", "N|", 1)
'
'            Sql1 = "insert into slotes (codartic, numlinea, sublinea, fechamov, nrolotes, "
'            Sql1 = Sql1 & "cantidad, categoria, codprove, codsocio, tipomovi, documento, regfitosanitario) VALUES ("
'            Sql1 = Sql1 & RT1!codArtic & "," & Numlinea & "," & RT1!Sublinea & ",'"
'            Sql1 = Sql1 & Format(FechaAlb, FormatoFecha) & "','" & DevNombreSQL(RT1!nrolotes) & "',"
'            Sql1 = Sql1 & TransformaComasPuntos(RT1!cantidad) & ",'" & DevNombreSQL(Categoria) & "',"
'            Sql1 = Sql1 & proveedor & "," & ValorNulo & ",0,'"
'            Sql1 = Sql1 & DevNombreSQL(Trim(Albaran)) & "','" & DevNombreSQL(RT!Regfitosanitario) & "')" 'VRS:1.0.1(12)
'
'            Conn.Execute Sql1
'
'            ' insertamos las sublineas en caso de ser fitosanitario
'            Sql1 = "insert into slialp1 (numalbpr, codprove, numlinea, "
'            Sql1 = Sql1 & "sublinea, codartic, nrolotes, regfitosanitario, "
'            Sql1 = Sql1 & "cantidad) VALUES ('" & DevNombreSQL(Trim(Albaran)) & "'," 'VRS:1.0.1(12)
'            Sql1 = Sql1 & proveedor & "," & Numlinea & ","
'            Sql1 = Sql1 & RT1!Sublinea & "," & RT1!codArtic & ",'"
'            Sql1 = Sql1 & DevNombreSQL(RT1!nrolotes) & "','" & DevNombreSQL(RT!Regfitosanitario) & "',"
'            Sql1 = Sql1 & TransformaComasPuntos(RT1!cantidad) & ")"
'
'            Conn.Execute Sql1
'
'            ' actualizamos el stock de sarticlotes
'            fechainv = ""
'            fechainv = DevuelveDesdeBD(1, "fechainv", "sartic", "codartic|", RT1!codArtic & "|", "N|", 1)
'            If fechainv = "" Then fechainv = "01/01/1900" 'CStr(Format(Now - 1, "dd/mm/yyyy"))
'            horainve = ""
'            horainve = DevuelveDesdeBD(1, "horainve", "sartic", "codartic|", RT1!codArtic & "|", "N|", 1)
'            If horainve = "" Then horainve = "01/01/1900 00:00:00" 'CStr(Format(Now - 1, "dd/mm/yyyy"))
'
'
'            If CDate(fechainv) < CDate(FechaAlb) Or _
'               (CDate(fechainv) = CDate(FechaAlb) And horainve < Now) Then
'                ActualizarStockLotes RT1!codArtic, RT1!nrolotes, RT!Regfitosanitario, CDbl(RT1!cantidad), 0
'            End If
'
'            Conn.Execute "delete from tmplotes " & Sql2
'
'            RT1.MoveNext
'        Wend
'        RT1.Close
'
'        ' fin del lotes
'
'
'        If RT!cantidad - RT!recibida <= 0 Then
'            ' tenemos que borrar la linea del pedido
'            Sql1 = "delete from slippr where numpedpr =" & Pedido & " and numlinea =" & RT!Numlinea
'            Conn.Execute Sql1
'
'            ' decrementamos el numero de linea de las lineas posteriores
''            sql1 = "update slippr set numlinea = numlinea - 1 where numpedpr = " & Pedido & " and "
''            sql1 = sql1 & " numlinea > " & RT!Numlinea
''            Conn.Execute sql1
'        Else
'            ' ponemos en cantidad la diferencia entre cantidad y recibida
'            ' y en recibida ponemos 0
'            Diferencia = RT!cantidad - RT!recibida
'            Importe = Round2(Diferencia * RT!Precioar, 2)
'            If Tipodtos = 0 Then
'                Importe = Round2(Importe - (Importe * RT!dtoline1 / 100) - (Importe * RT!dtoline2 / 100), 2)
'            Else
'                Importe = Importe - (Importe * RT!dtoline1 / 100)
'                Importe = Importe - (Importe * RT!dtoline2 / 100)
'                Importe = Round2(Importe, 2)
'            End If
'
'
'            Sql1 = "update slippr set cantidad = cantidad - recibida, implinea = " & TransformaComasPuntos(CStr(Importe))
'            Sql1 = Sql1 & ", recibida = " & ValorNulo & " where numpedpr = " & Pedido & " and "
'            Sql1 = Sql1 & " numlinea = " & RT!Numlinea
'            Conn.Execute Sql1
'
'        End If
'
'        RT.MoveNext
'
'    Wend
'    RT.Close
'
'    ' si no hay lineas tendremos que borrar la cabecera
'    SQL = "Select * from slippr where  numpedpr=" & Pedido
'    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If RT.EOF Then
'        SQL = "delete from scappr where numpedpr = " & Pedido
'        Conn.Execute SQL
'    End If
'
'    GenerarAlbaran = Albaran
'
'Etmpconext:
'    If GenerarAlbaran <> "-1" Then
'        Conn.CommitTrans
'    Else
'        Conn.RollbackTrans
'    End If
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Generando el albaran de compra"
'    End If
'    Set RT = Nothing
'End Function


Public Function GenerarListadoCSN(Usuario As Long, desdefec As Date, hastafec As Date, Tipo As Byte, TipoDosimetria As Byte) As Boolean
' tipo = 0 dosis superficial
' tipo = 1 dosis profunda

'tipodosimetria = 0 personal
'tipodosimetria = 1 area
Dim I As Integer
Dim Encontrado As Boolean
Dim rL As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim valor As Currency
Dim valormax As Currency
Dim continua As Boolean
Dim semigracsn As Byte

    On Error GoTo EGenerarlistadocsn
    GenerarListadoCSN = False
                                                     
    ' borramos primero lo que hay en la temporal del usuario
    sql = "delete from zlistadocsn where codusu = " & Usuario
    Conn.Execute sql
    
' antes estaba esto: al cargar la tabla de rangos, solo hay que copiarla
'    For i = 1 To 11
'        sql = "insert into zlistadocsn (codusu, tipo, descripcion, numeroreg, dosisacum)"
'        sql = sql & "VALUES (" & Usuario & "," & i & ","
'        Select Case i
'            Case 1
'                sql = sql & "'Dosis < NR'"
'            Case 2
'                sql = sql & "'NR    <= Dosis <= 01.00'"
'            Case 3
'                sql = sql & "'01.00 < Dosis  <= 02.00'"
'            Case 4
'                sql = sql & "'02.00 < Dosis  <= 03.00'"
'            Case 5
'                sql = sql & "'03.00 < Dosis  <= 04.00'"
'            Case 6
'                sql = sql & "'04.00 < Dosis  <= 05.00'"
'            Case 7
'                sql = sql & "'05.00 < Dosis  <= 06.00'"
'            Case 8
'                sql = sql & "'06.00 < Dosis  <= 10.00'"
'            Case 9
'                sql = sql & "'10.00 < Dosis  <= 20.00'"
'            Case 10
'                sql = sql & "'20.00 < Dosis  <= 50.00'"
'            Case 11
'                sql = sql & "'Dosis > 50.00         '"
'        End Select
'        sql = sql & ",0,0)"
'        Conn.Execute sql
'    Next i
    
    sql = "insert into zlistadocsn (codusu, tipo, descripcion, numeroreg, dosisacum)"
    sql = sql & "select " & Usuario & ", orden , ' ', 0, 0 from rangoscsn where tipo = " & Tipo
    
    Conn.Execute sql
    
    ' actualizamos la tabla temporal que vamos a imprimir
    If TipoDosimetria = 0 Then
        sql = "select dosiscuerpo.dni_usuario, sum("
    Else
        sql = "select dosisarea.n_dosimetro, sum("
    End If
    
    ' 06/03/2007 [DV] Cruzar la tabla de dosimetros con la de dosis no es necesario y repercute
    ' caóticamente en los resultados.
    If TipoDosimetria = 0 Then
        If Tipo = 0 Then
'            sql = sql & "dosis_superf)  from dosiscuerpo, dosimetros "
            sql = sql & "dosis_superf)  from dosiscuerpo "
        Else
'            sql = sql & "dosis_profunda)  from dosiscuerpo, dosimetros "
            sql = sql & "dosis_profunda)  from dosiscuerpo "
        End If
    Else
        If Tipo = 0 Then
'            sql = sql & "dosis_superf)  from dosisarea, dosimetros "
            sql = sql & "dosis_superf)  from dosisarea "
        Else
'            sql = sql & "dosis_profunda)  from dosisarea, dosimetros "
            sql = sql & "dosis_profunda)  from dosisarea "
        End If
    End If
    
    If TipoDosimetria = 0 Then
        sql = sql & " where f_dosis >= '" & Format(desdefec, FormatoFecha) & "' and "
        sql = sql & " f_dosis <= '" & Format(hastafec, FormatoFecha) & "' and "
        ' 06/03/2007 [DV]
        'sql = sql & " dosimetros.tipo_dosimetro = 0 and dosiscuerpo.n_reg_dosimetro = dosimetros.n_reg_dosimetro"
        ' 27/02/2006 [DV] Modificación referente a fallos en el envío CSN
        sql = sql & " dosiscuerpo.dni_usuario <> '0' and "
        sql = sql & " dosiscuerpo.dni_usuario<>'999999999' and "
        sql = sql & " dosiscuerpo.dni_usuario <> '999999998' and "
        sql = sql & " dosiscuerpo.dni_usuario <> '999999997' and "
        sql = sql & " dosiscuerpo.dni_usuario <> '999999996' and "
        sql = sql & " dosiscuerpo.dni_usuario <> '777777777' and "
        sql = sql & " dosiscuerpo.dni_usuario <> '666666666' and "
        sql = sql & " dosiscuerpo.dni_usuario <> '888888888'"
        ' 27/02/2006 [DV] Hasta aquí
        sql = sql & " group by dosiscuerpo.dni_usuario"
    Else
        sql = sql & " where f_dosis >= '" & Format(desdefec, FormatoFecha) & "' and "
        sql = sql & " f_dosis <= '" & Format(hastafec, FormatoFecha) & "' and "
        'sql = sql & " dosimetros.tipo_dosimetro = 2 and " ' 06/03/2007 [DV]
        sql = sql & " dosisarea.dni_usuario = '777777777' and "
        
        ' 27/02/2006 [DV] Modificación referente a fallos en el envío CSN
        sql = sql & " dosisarea.dni_usuario <> '0' and "
        sql = sql & " dosisarea.dni_usuario <> '999999996' and "
        sql = sql & " dosisarea.dni_usuario <> '666666666' and "
        sql = sql & " dosisarea.dni_usuario <> '888888888'" ' and "
        ' 27/02/2006 [DV] Hasta aquí
        
        'sql = sql & " dosisarea.n_reg_dosimetro = dosimetros.n_reg_dosimetro"
        sql = sql & " group by dosisarea.n_dosimetro "
    End If
    
    Set RT = New ADODB.Recordset
    RT.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then RT.MoveFirst
    While Not RT.EOF
        continua = True
        If TipoDosimetria = 0 Then
            semigracsn = 0
            semigracsn = DevuelveDesdeBD(1, "semigracsn", "operarios", "dni|", Trim(RT.Fields(0).Value) & "|", "T|", 1)
            If semigracsn = 0 Then continua = False
'        Else
'            If Trim(RT.Fields(0).Value) <> "777777777" Then
'                semigracsn = 0
'                semigracsn = DevuelveDesdeBD(1, "semigracsn", "operarios", "dni|", Trim(RT.Fields(0).Value) & "|", "T|", 1)
'                If semigracsn = 0 Then continua = False
'            End If
        End If
        sql = "update zlistadocsn set numeroreg = numeroreg + 1,"
        sql = sql & "dosisacum = dosisacum + "
        
        valor = 0
        If IsNull(RT.Fields(1).Value) Then
            valor = 0
        Else
            valor = RT.Fields(1).Value
        End If
'            If Valor = 0.11 Then Stop
        
        sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(valor)))
        sql = sql & " WHERE codusu = " & Usuario & " and tipo = "
    
        sql1 = "select orden, desde, hasta from rangoscsn where tipo = " & Tipo
        sql1 = sql1 & " order by tipo, orden"
        
        Set rL = New ADODB.Recordset
        rL.Open sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rL.EOF Then rL.MoveFirst
        Encontrado = False


        While Not rL.EOF And Not Encontrado
            
            If IsNull(rL.Fields(2).Value) Then
                valormax = 99999999
            Else
                valormax = rL.Fields(2).Value
            End If
'            If rL.Fields(1).Value = 0.11 Then Stop
            
''            ' Cambiado TAL Y COMO tiene que ser realmente. No se contemplaban los
''            ' casos 1 y 2 (siempre se usaba la condición del "Case Else")
''            Select Case rL!orden
''              Case 1
''                If (CCur(valor) < CCur(valormax)) Then Encontrado = True
''              Case 2
''                If (CCur(rL.Fields(1).Value) <= CCur(valor)) And (CCur(valor) <= CCur(valormax)) Then Encontrado = True
''              Case Else
                If (CCur(rL.Fields(1).Value) < CCur(valor)) And (CCur(valor) <= CCur(valormax)) Then Encontrado = True
''            End Select
            
            If Encontrado Then
               sql = sql & rL.Fields(0).Value
               If continua Then Conn.Execute sql
            End If
        
            rL.MoveNext
        Wend
        Set rL = Nothing
' antes estaba esto
'        If RT.Fields(1).Value < 0.1 Then
'            sql = sql & "1"
'        ElseIf (0.1 <= RT.Fields(1).Value) And (RT.Fields(1).Value <= 1) Then
'            sql = sql & "2"
'        ElseIf (1 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 2) Then
'            sql = sql & "3"
'        ElseIf (2 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 3) Then
'            sql = sql & "4"
'        ElseIf (3 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 4) Then
'            sql = sql & "5"
'        ElseIf (4 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 5) Then
'            sql = sql & "6"
'        ElseIf (5 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 6) Then
'            sql = sql & "7"
'        ElseIf (6 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 10) Then
'            sql = sql & "8"
'        ElseIf (10 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 20) Then
'            sql = sql & "9"
'        ElseIf (20 < RT.Fields(1).Value) And (RT.Fields(1).Value <= 50) Then
'            sql = sql & "10"
'        ElseIf (50 < RT.Fields(1).Value) Then
'            sql = sql & "11"
'        End If
'
'        Conn.Execute Sql
        
        RT.MoveNext
        
    Wend
    RT.Close
        
    Set RT = Nothing
    
    GenerarListadoCSN = True
    
    Exit Function

EGenerarlistadocsn:
    MuestraError Err.Number, "Generando Listado CSN."
End Function
 

'-- Esta librería contiene un conjunto de funciones de utilidad general
Public Function Comprueba_CC(CC As String) As Boolean
    Dim ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim I, i2, i3, i4 As Integer
    Dim NumCC As String ' Número de cuenta propiamente dicho
    '-- Esta función comprueba la corrección de un número de cuenta pasado en CC
    If Len(CC) <> 20 Then Exit Function '-- Las cuentas deben contener 20 dígitos en total
    '-- Calculamos el primer dígito de control
    I = Val(Mid(CC, 1, 1)) * 4
    I = I + Val(Mid(CC, 2, 1)) * 8
    I = I + Val(Mid(CC, 3, 1)) * 5
    I = I + Val(Mid(CC, 4, 1)) * 10
    I = I + Val(Mid(CC, 5, 1)) * 9
    I = I + Val(Mid(CC, 6, 1)) * 7
    I = I + Val(Mid(CC, 7, 1)) * 3
    I = I + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 9, 1)) Then Exit Function '-- El primer dígito de control no coincide
    '-- Calculamos el segundo dígito de control
    I = Val(Mid(CC, 11, 1)) * 1
    I = I + Val(Mid(CC, 12, 1)) * 2
    I = I + Val(Mid(CC, 13, 1)) * 4
    I = I + Val(Mid(CC, 14, 1)) * 8
    I = I + Val(Mid(CC, 15, 1)) * 5
    I = I + Val(Mid(CC, 16, 1)) * 10
    I = I + Val(Mid(CC, 17, 1)) * 9
    I = I + Val(Mid(CC, 18, 1)) * 7
    I = I + Val(Mid(CC, 19, 1)) * 3
    I = I + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 10, 1)) Then Exit Function '-- El segundo dígito de control no coincide
    '-- Si llega aquí ambos figitos de control son correctos
    Comprueba_CC = True
End Function

Public Function Comprobar_NIF(nif As String) As Boolean
    '-- Comprobación general de NIF
    If Len(nif) <> 9 Then
        Comprobar_NIF = False
        Exit Function
    Else
        If IsNumeric(Mid(nif, 1, 1)) Then
            '-- Comienza por número
            If IsNumeric(Mid(nif, 9, 1)) Then
                Comprobar_NIF = False
                Exit Function
            Else
                Comprobar_NIF = Comprobar_NIF_PersonaFisica(nif)
            End If
        Else
            '-- comienza por letra
            If IsNumeric(Mid(nif, 9, 1)) Then
                '-- Acaba en número
                If InStr(1, "ABCDEFGHPQSN", Mid(nif, 1, 1)) <> 0 Then
                    '-- Es una sociedad
                    Comprobar_NIF = Comprobar_NIF_Sociedad(nif)
                ElseIf InStr(1, "T", Mid(nif, 1, 1)) <> 0 Then
                    '-- Es un NIF antiguo que no lleva comprobación
                    Comprobar_NIF = True
                End If
            Else
                '-- Acaba en letra
                If InStr(1, "ABCDEFGHPQSN", Mid(nif, 1, 1)) <> 0 Then
                    '-- Es una sociedad
                    Comprobar_NIF = Comprobar_NIF_Sociedad(nif)
                ElseIf InStr(1, "MX", Mid(nif, 1, 1)) <> 0 Then
                    '-- Es un extranjero
                    Comprobar_NIF = Comprobar_NIF_PersonaExtranjera(nif)
                ElseIf InStr(1, "KL", Mid(nif, 1, 1)) <> 0 Then
                    '-- Es un NIF antiguo que no lleva comprobación
                    Comprobar_NIF = True
                End If
            End If
        End If
    End If
End Function

Public Function Comprobar_NIF_PersonaFisica(nif As String) As Boolean
    Dim mCadena As String
    Dim mLetra As String
    Dim m23 As Integer
    mCadena = "TRWAGMYFPDXBNJZSQVHLCKE"
    '-- Tomamos el NIF propiamente dicho y calculamos el módulo 23
    m23 = Val(Mid(nif, 1, 8)) Mod 23
    mLetra = Mid(mCadena, m23 + 1, 1)
    '-- Validamos que la letra es correcta
    If Mid(nif, 9, 1) = mLetra Then
        Comprobar_NIF_PersonaFisica = True
    Else
        Comprobar_NIF_PersonaFisica = False
    End If
End Function

Public Function Comprobar_NIF_PersonaExtranjera(nif As String) As Boolean
    Dim mCadena As String
    Dim mLetra As String
    Dim m23 As Integer
    mCadena = "DTRWAGMYFPXBNJZSQVHLCKE"
    '-- Tomamos el NIF propiamente dicho y calculamos el módulo 23
    m23 = Val(Mid(nif, 2, 7)) Mod 23
    mLetra = Mid(mCadena, m23 + 1, 1)
    '-- Validamos que la letra es correcta
    If Mid(nif, 9, 1) = mLetra Then
        Comprobar_NIF_PersonaExtranjera = True
    Else
        Comprobar_NIF_PersonaExtranjera = False
    End If
End Function

Public Function Comprobar_NIF_Sociedad(nif As String) As Boolean
    Dim mCadena As String
    Dim mLetra As String
    Dim vNIF As String
    Dim mN2 As String
    Dim I, i2 As Integer
    Dim suma, Control As Long
    mCadena = "ABCDEFGHIJ"
    vNIF = Mid(nif, 2, 7)
    '-- Sumamos las cifras pares
    For I = 2 To Len(vNIF) Step 2
        suma = suma + Val(Mid(vNIF, I, 1))
    Next I
    '-- Ahora las impares * 2, y sumando las cifras del resultado
    For I = 1 To Len(vNIF) Step 2
        mN2 = CStr(Val(Mid(vNIF, I, 1)) * 2)
        For i2 = 1 To Len(mN2)
            suma = suma + Val(Mid(mN2, i2, 1))
        Next i2
    Next I
    '-- Ya tenemos la suma y calculamos el control
    Control = 10 - (suma Mod 10)
    If Control = 10 Then Control = 0
    mLetra = Mid(nif, 9, 1)
    If IsNumeric(mLetra) Then
        If Val(mLetra) = Control Then
            Comprobar_NIF_Sociedad = True
        Else
            Comprobar_NIF_Sociedad = False
        End If
    Else
        If Control = 0 Then Control = 10
        If mLetra = Chr(64 + Control) Then
            Comprobar_NIF_Sociedad = True
        Else
            Comprobar_NIF_Sociedad = False
        End If
    End If
End Function

Public Function FrmtStr(Campo As String, Longitud As Integer) As String
'    Dim LineaBlanca As String
    Dim CADENA As String
    
    FrmtStr = ""
    CADENA = Trim(Campo) & LineaBlanca
    FrmtStr = Mid(CADENA, 1, Longitud)
    
End Function



'funcion que comprueba que los clientes/proveedores tienen la CCC asignada
Public Function ComprobarCCCLineas(Rs As ADODB.Recordset, Tipo As Byte) As Boolean
' tipo : 0 = cartera de cobros
'        1 = cartera de pagos
Dim codigo As Long
Dim Cad As String


    ComprobarCCCLineas = True
    If Not Rs.EOF Then Rs.MoveFirst
    
    If Tipo = 0 Then
        codigo = Rs!codsocio
    Else
        codigo = Rs!codprove
    End If
    
    While Not Rs.EOF
        If IsNull(Rs!codbanco) Or IsNull(Rs!codsucur) _
           Or IsNull(Rs!digcontr) Or IsNull(Rs!cuentaba) Then
           If Tipo = 0 Then
               MsgBox "El Socio " & Format(codigo, "000000") & " no tiene asignada CCC. Revise Socio y cartera.", vbExclamation, "¡Error!"
           Else
               MsgBox "El Proveedor " & Format(codigo, "000000") & " no tiene asignada CCC. Revise Proveedor y cartera.", vbExclamation, "¡Error!"
           End If
           ComprobarCCCLineas = False
           Exit Function
' de momento no he incluido que la CCC que me incluyen sea correcta
'        Else
'            Cad = Format(RS!codbanco, "0000") & Format(RS!codsucur, "0000") & FrmtStr(RS!digcontr, 2)
'            Cad = Cad & FrmtStr(RS!cuentaba, 10)
'            If Not Comprueba_CC(Cad) Then
'                If tipo = 0 Then
'                    MsgBox "El Socio " & Format(codigo, "000000") & " tiene la CCC incorrecta. Revise Socio y cartera.", vbExclamation
'                Else
'                    MsgBox "El Proveedor " & Format(codigo, "000000") & " tiene la CCC incorrecta. Revise Proveedor y cartera.", vbExclamation
'                End If
'
'                ComprobarCCCLineas = False
'                Exit Function
'            End If
        End If
        Rs.MoveNext
    Wend
    
End Function

Public Sub PonerIndicador(lblIndicador As Label, Modo As Byte)
' modificandolineas = 1 insertar lineas
'                   = 0 modificando lineas

'Pone el titulo del label lblIndicador
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub


Public Sub CargarDatosOperarios(nif As String, ByRef ape1 As String, ByRef ape2 As String, ByRef nombre As String)
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "Select distinct apellido_1, apellido_2, nombre from operarios"
    sql = sql & " where dni = '" & Trim(nif) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    
    ape1 = ""
    ape2 = ""
    nombre = ""
    If Not Rs.EOF Then
        ape1 = Rs.Fields(0).Value
        ape2 = Rs.Fields(1).Value
        nombre = Rs.Fields(2).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
End Sub



Public Sub CargarEmpresas()
Dim sql As String

    sql = "delete from vempresas where codusu = " & vUsu.codigo
    Conn.Execute sql

    sql = "insert into vempresas "
    sql = sql & "select c_empresa, max(f_alta), f_baja, cif_nif, nom_comercial, direccion, poblacion,"
    sql = sql & "c_postal, distrito, tel_contacto, fax, pers_contacto, migrado, mail_internet,"
    sql = sql & "c_tipo, " & vUsu.codigo & " from empresas "
    sql = sql & "group by c_empresa"
    
'    , f_baja, cif_nif, nom_comercial, direccion, poblacion, "
'    Sql = Sql & "c_postal, distrito, tel_contacto, fax, pers_contacto, migrado, mail_internet,"
'    Sql = Sql & "c_tipo , 16"
    
    Conn.Execute sql

End Sub

Public Sub CargarEmpresas1()
Dim Rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim existe As String

    Conn.Execute "delete from vempresas where codusu = " & vUsu.codigo
    
    sql = "select * from empresas order by c_empresa, f_alta desc "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    While Not Rs.EOF
        existe = ""
        existe = DevuelveDesdeBD(1, "c_empresa", "vempresas", "c_empresa|codusu|", Trim(Rs!c_empresa) & "|" & vUsu.codigo & "|", "T|N|", 2)
        If existe = "" Then
            sql1 = "insert into vempresas values ('" & Trim(Rs!c_empresa) & "','" & Format(Rs!f_alta, FormatoFecha) & "',"
            
            If Not IsNull(Rs!f_baja) Then
                sql1 = sql1 & "'" & Format(Rs!f_baja, FormatoFecha) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!cif_nif) Then
                sql1 = sql1 & "'" & Trim(Rs!cif_nif) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!nom_comercial)) & "',"
            
            If Not IsNull(Rs!direccion) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!direccion)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!poblacion) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!poblacion)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & Format(Trim(Rs!c_postal), "00") & "',"
            
            If Not IsNull(Rs!distrito) Then
                sql1 = sql1 & "'" & Trim(Rs!distrito) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!tel_contacto) Then
                sql1 = sql1 & "'" & Trim(Rs!tel_contacto) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!fax) Then
                sql1 = sql1 & "'" & Trim(Rs!fax) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!pers_contacto) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!pers_contacto)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!migrado) Then
                sql1 = sql1 & "'" & Trim(Rs!migrado) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!mail_internet) Then
                sql1 = sql1 & "'" & Trim(Rs!mail_internet) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & Format(Rs!c_tipo, "0") & "," & vUsu.codigo & ")"
            
            Conn.Execute sql1
        End If
        
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing
    
End Sub

Public Sub CargarOperarios()
Dim sql As String

    sql = "delete from voperarios where codusu =  " & vUsu.codigo
    Conn.Execute sql
    
    sql = "insert into  voperarios "
    sql = sql & "select dni, n_seg_social, n_carnet_radiolog, f_emi_carnet_rad, apellido_1,"
    sql = sql & "apellido_2, nombre, direccion, poblacion, c_postal, distrito, c_tipo_trabajo,"
    sql = sql & "f_nacimiento, profesion_catego, sexo, plantilla_contrata, max(f_alta),f_baja,"
    sql = sql & "migrado, cod_rama_gen, semigracsn, " & vUsu.codigo & " from operarios "
    sql = sql & "group by dni"
    
'    , n_seg_social, n_carnet_radiolog, f_emi_carnet_rad, apellido_1,"
'    Sql = Sql & "apellido_2, nombre, direccion, poblacion, c_postal, distrito, c_tipo_trabajo,"
'    Sql = Sql & "f_nacimiento, profesion_catego, sexo, plantilla_contrata, f_baja,"
'    Sql = Sql & "migrado , cod_rama_gen, semigracsn, 22"
    
    Conn.Execute sql

End Sub

Public Sub CargarOperarios1()
Dim Rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim existe As String

    Conn.Execute "delete from voperarios where codusu = " & vUsu.codigo
    
    sql = "select * from operarios order by dni, f_alta desc "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    While Not Rs.EOF
        existe = ""
        existe = DevuelveDesdeBD(1, "dni", "voperarios", "dni|codusu|", Trim(Rs!dni) & "|" & vUsu.codigo & "|", "T|N|", 2)
        
        If existe = "" Then
            sql1 = "insert into voperarios values ('" & Trim(Rs!dni) & "',"
            
            
            If Not IsNull(Rs!n_seg_social) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!n_seg_social)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!n_carnet_radiolog) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!n_carnet_radiolog)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!f_emi_carnet_rad) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!f_emi_carnet_rad)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!apellido_1)) & "',"
            sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!apellido_2)) & "',"
            sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!nombre)) & "',"
            
            If Not IsNull(Rs!direccion) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!direccion)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!poblacion) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!poblacion)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & Format(Trim(Rs!c_postal), "00") & "',"
            
            If Not IsNull(Rs!distrito) Then
                sql1 = sql1 & "'" & Trim(Rs!distrito) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & Trim(Rs!c_tipo_trabajo) & "',"
            
            If Not IsNull(Rs!f_nacimiento) Then
                sql1 = sql1 & "'" & Format(Rs!f_nacimiento, FormatoFecha) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!profesion_catego) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!profesion_catego)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & Trim(Rs!sexo) & "',"
            
            sql1 = sql1 & "'" & Trim(Rs!plantilla_contrata) & "',"
            
            sql1 = sql1 & "'" & Format(Rs!f_alta, FormatoFecha) & "',"
            
            If Not IsNull(Rs!f_baja) Then
                sql1 = sql1 & "'" & Format(Rs!f_baja, FormatoFecha) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!migrado) Then
                sql1 = sql1 & "'" & Trim(Rs!migrado) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & Format(Rs!cod_rama_gen, "00") & "',"
            
            sql1 = sql1 & Format(Rs!semigracsn, "0") & "," & vUsu.codigo & ")"
            
            Conn.Execute sql1
        End If
        
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing

End Sub
Public Sub CargarInstalaciones()
Dim sql As String

    sql = "delete from vinstalaciones where codusu = " & vUsu.codigo
    Conn.Execute sql

    sql = "insert into vinstalaciones "
    sql = sql & "select c_empresa, c_instalacion, max(f_alta), f_baja, descripcion, direccion,"
    sql = sql & "poblacion, c_postal, distrito, telefono, fax, persona_contacto, migrado,"
    sql = sql & "rama_gen , rama_especifica, mail_internet, Observaciones, c_tipo, " & vUsu.codigo
    sql = sql & " from instalaciones "
    sql = sql & "group by c_empresa, c_instalacion"
'    , f_baja, descripcion, direccion,"
'    Sql = Sql & "poblacion, c_postal, distrito, telefono, fax, persona_contacto, migrado,"
'    Sql = Sql & "rama_gen , rama_especifica, mail_internet, Observaciones, c_tipo, 19"

    Conn.Execute sql

End Sub

Public Sub CargarInstalaciones1()
Dim Rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim existe As String


    Conn.Execute "delete from vinstalaciones where codusu = " & vUsu.codigo
    
    sql = "select * from instalaciones order by c_empresa, c_instalacion, f_alta desc "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    While Not Rs.EOF
        existe = ""
        existe = DevuelveDesdeBD(1, "c_instalacion", "vinstalaciones", "c_empresa|c_instalacion|codusu|", Rs!c_empresa & "|" & Rs!c_instalacion & vUsu.codigo & "|", "T|T|N|", 3)
        
        If existe = "" Then
            sql1 = "insert into vinstalaciones values ('" & Trim(Rs!c_empresa) & "','"
            sql1 = sql1 & Trim(Rs!c_instalacion) & "','"
            sql1 = sql1 & Format(Rs!f_alta, FormatoFecha) & "',"
            
            If Not IsNull(Rs!f_baja) Then
                sql1 = sql1 & "'" & Format(Rs!f_baja, FormatoFecha) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!descripcion)) & "',"
            
            If Not IsNull(Rs!direccion) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!direccion)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!poblacion) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!poblacion)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & Format(Trim(Rs!c_postal), "00") & "',"
            
            If Not IsNull(Rs!distrito) Then
                sql1 = sql1 & "'" & Trim(Rs!distrito) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!telefono) Then
                sql1 = sql1 & "'" & Trim(Rs!telefono) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!fax) Then
                sql1 = sql1 & "'" & Trim(Rs!fax) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!persona_contacto) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!persona_contacto)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!migrado) Then
                sql1 = sql1 & "'" & Trim(Rs!migrado) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & "'" & Format(Trim(Rs!rama_gen), "00") & "',"
            sql1 = sql1 & "'" & Format(Trim(Rs!rama_especifica), "00") & "',"
            
            If Not IsNull(Rs!mail_internet) Then
                sql1 = sql1 & "'" & Trim(Rs!mail_internet) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            If Not IsNull(Rs!Observaciones) Then
                sql1 = sql1 & "'" & DevNombreSQL(Trim(Rs!Observaciones)) & "',"
            Else
                sql1 = sql1 & "null,"
            End If
            
            sql1 = sql1 & Format(Rs!c_tipo, "0") & "," & vUsu.codigo & ")"
            
            Conn.Execute sql1
        End If
        
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing




End Sub

Public Sub CommitConexion()
    On Error Resume Next
    Conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function TieneFechaRetirada(dosim As String, Tipo As Byte) As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String

    TieneFechaRetirada = True
        
    sql = "select n_dosimetro from dosimetros where n_dosimetro = '" & Trim(dosim) & "' and "
    If Tipo = 0 Or Tipo = 2 Then
        sql = sql & "(tipo_dosimetro = 0 or  tipo_dosimetro = 2) and "
    Else
        sql = sql & "tipo_dosimetro = 1 and "
    End If
    
    sql = sql & "(f_retirada is null or f_retirada ='')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    sql = ""
    If Not Rs.EOF Then
        TieneFechaRetirada = False
'        rs.Close
'        ' miramos en la tabla de dosimetros area
'        sql = "select n_dosimetro from dosisarea where n_dosimetro = '" & Trim(dosim) & "' and "
'        sql = sql & "(f_retirada is null or f_retirada = '')"
'
'        rs.Open sql, Conn, , , adCmdText
'        If Not rs.EOF Then TieneFechaRetirada = False
'    Else
'        TieneFechaRetirada = False
    End If
    Set Rs = Nothing

End Function


Public Function DosimetroEnUso(dosim As String, nreg As Long, Tipo As Byte, Sist As String) As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String

    DosimetroEnUso = False
        
    sql = "select n_dosimetro from dosimetros where n_dosimetro = '" & Trim(dosim) & "' and "
    sql = sql & " n_reg_dosimetro <> " & nreg & " and sistema = '" & Sist & "' and "
    
    If Tipo = 1 Then
        sql = sql & "tipo_dosimetro = 1 and "
    Else
        sql = sql & "(tipo_dosimetro = 0 or tipo_dosimetro = 2) and "
    End If
    
    sql = sql & "(f_retirada is null or f_retirada ='')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    sql = ""
    If Not Rs.EOF Then
        DosimetroEnUso = True
    End If
    Set Rs = Nothing

End Function


