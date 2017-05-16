Attribute VB_Name = "Errores"
Option Explicit

Private NErrores As Long
Private NomArchivo As String
Private NF As Integer

'-----------------------------------------------------------------
'
' Si tiene error el objeto conn entonces lo mostramos
'
'
'---------------------------------------------------------------

Public Sub ControlamosError(ByRef CADENA As String)

Select Case Conn.Errors(0).NativeError
Case 0
    CADENA = "El controlador ODBC no admite las propiedades solicitadas."
Case 1044
    CADENA = "Acceso denegado para usuario: " & CadenaDesde(15, Conn.Errors(0).Description, ":")
Case 1045
    CADENA = "Acceso denegado para usuario: " & CadenaDesde(15, Conn.Errors(0).Description, ":")
Case 1048
    CADENA = "Columna no puede ser nula: " & CadenaDesde(1, Conn.Errors(0).Description, ":")
Case 1049
    CADENA = "Base de datos desconocida: " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 1052
    CADENA = "La columna :" & CadenaDesde(1, Conn.Errors(0).Description, "'") & " tiene un nombre ambiguo "
Case 1054
    CADENA = "Columna desconocida en cadena SQL."
Case 1062
    CADENA = "Entrada duplicada en BD." & vbCrLf & CadenaDesde(60, Conn.Errors(0).Description, "'")
Case 1064
    CADENA = "Error en el SQL."
Case 1109
    CADENA = "Tabla desconocida:  " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 1110
    CADENA = "Columna : " & CadenaDesde(1, Conn.Errors(0).Description, "'") & " especificada dos veces"
Case 1146
    CADENA = "Tabla no existe:  " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 1136
    CADENA = "N� de columnas en el SQL incorrectos."
Case 1205
    CADENA = "Tabla bloqueada. Tiempo espera excedido"
Case 1216
    CADENA = "Imposible a�adir una columna hija. Fallo en la clave referencial"
Case 1217
    CADENA = "El registro es clave referencial en otras tablas"
Case 2003
    CADENA = "Imposible conectar con el servidor " & CadenaDesde(15, Conn.Errors(0).Description, "'")
Case 2005
    CADENA = "Servidor host MYSQL desconocido:  " & CadenaDesde(1, Conn.Errors(0).Description, "'")
Case 2013
    CADENA = "Se ha perdido la conexi�n con el servidor MySQL durante la ejecuci�n."
Case Else
    CADENA = ""
End Select
End Sub


Private Function CadenaDesde(Inicio As Integer, CADENA As String, Caracter As String) As String
Dim I, J
CadenaDesde = ""
I = InStr(Inicio, CADENA, Caracter)
If I >= Inicio Then
    J = InStr(I + 1, CADENA, Caracter)
    I = I + 1
    If J > 0 Then CadenaDesde = Mid(CADENA, I, J - I)
End If
End Function


Public Sub InsertaError(Cadena1 As String, vError As String)
    Print #NF, " ---- ---- ----"
    Print #NF, Cadena1
    Print #NF, " -> " & vError
    NErrores = NErrores + 1
End Sub


Public Sub IncializaErrores(Lugar As String)
Dim I As Integer

Dim Aux2 As String

    Aux2 = App.Path & "\ErrorCC"
    If Dir(Aux2, vbDirectory) = "" Then MkDir Aux2
    
    I = 0
    Aux2 = Aux2 & "\" & Format(Now, "yyyymmdd")
    Do
        NomArchivo = Aux2 & "." & Format(I, "000")
        If Dir(NomArchivo) = "" Then
            'Este nombre es el correcto
            I = -1
        Else
            I = I + 1
            If I > 900 Then
                MsgBox "IMPOSIBLE ERROR. DDCCAAMM. Finalizara", vbCritical, "�Error!"
                End
            End If
        End If
    Loop Until I < 0
    
    
    'Ya tenemos el nombre
    NErrores = 0
    NF = FreeFile
    Open NomArchivo For Output As #NF
    Print #NF, "Fecha/hora: " & Now
    Print #NF, Lugar
    Print #NF,
End Sub

Public Function TieneErrores() As String
    Print #NF,
    Print #NF,
    Print #NF,
    Print #NF,
    Print #NF, " FIN ARCHIVO: " & Now
    Close #NF
    espera 0.2
    If NErrores > 0 Then
        TieneErrores = NomArchivo
    Else
        TieneErrores = ""
        Kill NomArchivo
    End If
End Function

Public Sub VolcarFicheroError(Texto As String)
Dim Cad As String
    Texto = ""
    On Error GoTo EVolcarFicheroError
    NF = FreeFile
    Open NomArchivo For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        Texto = Texto & Cad & vbCrLf
    Wend
    Close (NF)
    Exit Sub
EVolcarFicheroError:
    MsgBox "Fayo leyendo el fichero de error: " & NomArchivo & vbCrLf & Err.Description, vbCritical, "�Error!"
End Sub

'
'/* Copyright Abandoned 1997 TCX DataKonsult AB & Monty Program KB & Detron HB
'   This file is public domain and comes with NO WARRANTY of any kind
'   Traduccion por Miguel Angel Fernandez Roiz -- LoboCom Sistemas, s.l.
'   From June 28, 2001 translated by Miguel Solorzano miguel@mysql.com */
'"hashchk",
'"isamchk",
'"NO",
'"SI",
'"No puedo crear archivo '%-.64s' (Error: %d)",
'"No puedo crear tabla '%-.64s' (Error: %d)",
'"No puedo crear base de datos '%-.64s'. Error %d",
'"No puedo crear base de datos '%-.64s'. La base de datos ya existe",
'"No puedo eliminar base de datos '%-.64s'. La base de datos no existe",
'"Error eliminando la base de datos(no puedo borrar '%-.64s', error %d)",
'"Error eliminando la base de datos (No puedo borrar directorio '%-.64s', error %d)",
'"Error en el borrado de '%-.64s' (Error: %d)",
'"No puedo leer el registro en la tabla del sistema",
'"No puedo obtener el estado de '%-.64s' (Error: %d)",
'"No puedo acceder al directorio (Error: %d)",
'"No puedo bloquear archivo: (Error: %d)",
'"No puedo abrir archivo: '%-.64s'. (Error: %d)",
'"No puedo encontrar archivo: '%-.64s' (Error: %d)",
'"No puedo leer el directorio de '%-.64s' (Error: %d)",
'"No puedo cambiar al directorio de '%-.64s' (Error: %d)",
'"El registro ha cambiado desde la ultima lectura de la tabla '%-.64s'",
'"Disco lleno (%s). Esperando para que se libere algo de espacio....",
'"No puedo escribir, clave duplicada en la tabla '%-.64s'",
'"Error en el cierre de '%-.64s' (Error: %d)",
'"Error leyendo el fichero '%-.64s' (Error: %d)",
'"Error en el renombrado de '%-.64s' a '%-.64s' (Error: %d)",
'"Error escribiendo el archivo '%-.64s' (Error: %d)",
'"'%-.64s' esta bloqueado contra cambios",
'"Ordeancion cancelada",
'"La vista '%-.64s' no existe para '%-.64s'",
'"Error %d desde el manejador de la tabla",
'"El manejador de la tabla de '%-.64s' no tiene esta opcion",
'"No puedo encontrar el registro en '%-.64s'",
'"Informacion erronea en el archivo: '%-.64s'",
'"Clave de archivo erronea para la tabla: '%-.64s'. Intente repararlo",
'"Clave de archivo antigua para la tabla '%-.64s'; Reparelo!",
'"'%-.64s' es de solo lectura",
'"Memoria insuficiente. Reinicie el demonio e intentelo otra vez (necesita %d bytes)",
'"Memoria de ordenacion insuficiente. Incremente el tamano del buffer de ordenacion",
'"Inesperado fin de ficheroU mientras leiamos el archivo '%-.64s' (Error: %d)",
'"Demasiadas conexiones",
'"Memoria/espacio de tranpaso insuficiente",
'"No puedo obtener el nombre de maquina de tu direccion",
'"Protocolo erroneo",
'"Acceso negado para usuario: '%-.32s@%-.64s' para la base de datos '%-.64s'",
'"Acceso negado para usuario: '%-.32s@%-.64s' (Usando clave: %s)",
'"Base de datos no seleccionada",
'"Comando desconocido",
'"La columna '%-.64s' no puede ser nula",
'"Base de datos desconocida '%-.64s'",
'"La tabla  '%-.64s' ya existe",
'"Tabla '%-.64s' desconocida",
'"La columna: '%-.64s' en %s es ambigua",
'"Desconexion de servidor en proceso",
'"La columna '%-.64s' en %s es desconocida",
'"Usado '%-.64s' el cual no esta group by",
'"No puedo agrupar por '%-.64s'",
'"El estamento tiene funciones de suma y columnas en el mismo estamento",
'"La columna con count no tiene valores para contar",
'"El nombre del identificador '%-.64s' es demasiado grande",
'"Nombre de columna duplicado '%-.64s'",
'"Nombre de clave duplicado '%-.64s'",
'"Entrada duplicada '%-.64s' para la clave %d",
'"Especificador de columna erroneo para la columna '%-.64s'",
'"%s cerca '%-.64s' en la linea %d",
'"La query estaba vacia",
'"Tabla/alias: '%-.64s' es no unica",
'"Valor por defecto invalido para '%-.64s'",
'"Multiples claves primarias definidas",
'"Demasiadas claves primarias declaradas. Un maximo de %d claves son permitidas",
'"Demasiadas partes de clave declaradas. Un maximo de %d partes son permitidas",
'"Declaracion de clave demasiado larga. La maxima longitud de clave es %d",
'"La columna clave '%-.64s' no existe en la tabla",
'"La columna Blob '%-.64s' no puede ser usada en una declaracion de clave",
'"Longitud de columna demasiado grande para la columna '%-.64s' (maximo = %d).Usar BLOB en su lugar",
'"Puede ser solamente un campo automatico y este debe ser definido como una clave",
'"%s: preparado para conexiones\n",
'"%s: Apagado normal\n",
'"%s: Recibiendo signal %d. Abortando!\n",
'"%s: Apagado completado\n",
'"%s: Forzando a cerrar el thread %ld  usuario: '%-.64s'\n",
'"No puedo crear IP socket",
'"La tabla '%-.64s' no tiene indice como el usado en CREATE INDEX. Crea de nuevo la tabla",
'"Los separadores de argumentos del campo no son los especificados. Comprueba el manual",
'"No puedes usar longitudes de filas fijos con BLOBs. Por favor usa 'campos terminados por '.",
'"El archivo '%-.64s' debe estar en el directorio de la base de datos o ser de lectura por todos",
'"El archivo '%-.64s' ya existe",
'"Registros: %ld  Borrados: %ld  Saltados: %ld  Peligros: %ld",
'"Registros: %ld  Duplicados: %ld",
'"Parte de la clave es erronea. Una parte de la clave no es una cadena o la longitud usada es tan grande como la parte de la clave",
'"No puede borrar todos los campos con ALTER TABLE. Usa DROP TABLE para hacerlo",
'"No puedo ELIMINAR '%-.64s'. compuebe que el campo/clave existe",
'"Registros: %ld  Duplicados: %ld  Peligros: %ld",
'"INSERT TABLE '%-.64s' no esta permitido en FROM tabla lista",
'"Identificador del thread: %lu  desconocido",
'"Tu no eres el propietario del thread%lu",
'"No ha tablas usadas",
'"Muchas strings para columna %s y SET",
'"No puede crear un unico archivo log %s.(1-999)\n",
'"Tabla '%-.64s' fue trabada con un READ lock y no puede ser actualizada",
'"Tabla '%-.64s' no fue trabada con LOCK TABLES",
'"Campo Blob '%-.64s' no puede tener valores patron",
'"Nombre de base de datos ilegal '%-.64s'",
'"Nombre de tabla ilegal '%-.64s'",
'"El SELECT puede examinar muchos registros y probablemente con mucho tiempo. Verifique tu WHERE y usa SET OPTION SQL_BIG_SELECTS=1 si el SELECT esta correcto",
'"Error desconocido",
'"Procedimiento desconocido %s",
'"Equivocado parametro count para procedimiento %s",
'"Equivocados parametros para procedimiento %s",
'"Tabla desconocida '%-.64s' in %s",
'"Campo '%-.64s' especificado dos veces",
'"Invalido uso de funci�n en grupo",
'"Tabla '%-.64s' usa una extensi�n que no existe en esta MySQL versi�n",
'"Una tabla debe tener al menos 1 columna",
'"La tabla '%-.64s' est� llena",
'"Juego de caracteres desconocido: '%-.64s'",
'"Muchas tablas. MySQL solamente puede usar %d tablas en un join",
'"Muchos campos",
'"Tama�o de l�nea muy grande. M�ximo tama�o de l�nea, no contando blob, es %d. Tu tienes que cambiar algunos campos para blob",
'"Sobrecarga de la pila de thread:  Usada: %ld de una %ld pila.  Use 'mysqld -O thread_stack=#' para especificar una mayor pila si necesario",
'"Dependencia cruzada encontrada en OUTER JOIN.  Examine su condici�n ON",
'"Columna '%-.32s' es usada con UNIQUE o INDEX pero no est� definida como NOT NULL",
'"No puedo cargar funci�n '%-.64s'",
'"No puedo inicializar funci�n '%-.64s'; %-.80s",
'"No pasos permitidos para librarias conjugadas",
'"Funci�n '%-.64s' ya existe",
'"No puedo abrir libraria conjugada '%-.64s' (errno: %d %s)",
'"No puedo encontrar funci�n '%-.64s' en libraria'",
'"Funci�n '%-.64s' no est� definida",
'"Servidor '%-.64s' est� bloqueado por muchos errores de conexi�n.  Desbloquear con 'mysqladmin flush-hosts'",
'"Servidor '%-.64s' no est� permitido para conectar con este servidor MySQL",
'"Tu est�s usando MySQL como un usuario anonimo y usuarios anonimos no tienen permiso para cambiar las claves",
'"Tu debes de tener permiso para actualizar tablas en la base de datos mysql para cambiar las claves para otros",
'"No puedo encontrar una l�nea correponsdiente en la tabla user",
'"L�neas correspondientes: %ld  Cambiadas: %ld  Avisos: %ld",
'"No puedo crear un nuevo thread (errno %d). Si tu est� con falta de memoria disponible, tu puedes consultar el Manual para posibles problemas con SO",
'"El n�mero de columnas no corresponde al n�mero en la l�nea %ld",
'"No puedo reabrir tabla: '%-.64s',
'"Invalido uso de valor NULL",
'"Obtenido error '%-.64s' de regexp",
'"Mezcla de columnas GROUP (MIN(),MAX(),COUNT()...) con no GROUP columnas es ilegal si no hat la clausula GROUP BY",
'"No existe permiso definido para usuario '%-.32s' en el servidor '%-.64s'",
'"%-.16s comando negado para usuario: '%-.32s@%-.64s' para tabla '%-.64s'",
'"%-.16s comando negado para usuario: '%-.32s@%-.64s' para columna '%-.64s' en la tabla '%-.64s'",
'"Ilegal comando GRANT/REVOKE. Por favor consulte el manual para cuales permisos pueden ser usados.",
'"El argumento para servidor o usuario para GRANT es demasiado grande",
'"Tabla '%-64s.%s' no existe",
'"No existe tal permiso definido para usuario '%-.32s' en el servidor '%-.64s' en la tabla '%-.64s'",
'"El comando usado no es permitido con esta versi�n de MySQL",
'"Algo est� equivocado en su sintax",
'"Thread de inserci�n retarda no pudiendo bloquear para la tabla %-.64s",
'"Muchos threads retardados en uso",
'"Conexi�n abortada %ld para db: '%-.64s' usuario: '%-.64s' (%s)",
'"Obtenido un paquete mayor que 'max_allowed_packet'",
'"Obtenido un error de lectura de la conexi�n pipe",
'"Obtenido un error de fcntl()",
'"Obtenido paquetes desordenados",
'"No puedo descomprimir paquetes de comunicaci�n",
'"Obtenido un error leyendo paquetes de comunicaci�n"
'"Obtenido timeout leyendo paquetes de comunicaci�n",
'"Obtenido un error de escribiendo paquetes de comunicaci�n",
'"Obtenido timeout escribiendo paquetes de comunicaci�n",
'"La string resultante es mayor que max_allowed_packet",
'"El tipo de tabla usada no permite soporte para columnas BLOB/TEXT",
'"El tipo de tabla usada no permite soporte para columnas AUTO_INCREMENT",
'"INSERT DELAYED no puede ser usado con tablas '%-.64s', porque esta bloqueada con LOCK TABLES",
'"Incorrecto nombre de columna '%-.100s'",
'"El manipulador de tabla usado no puede indexar columna '%-.64s'",
'"Todas las tablas en la MERGE tabla no estan definidas identicamente",
'"No puedo escribir, debido al �nico constraint, para tabla '%-.64s'",
'"Columna BLOB column '%-.64s' usada en especificaci�n de clave sin tama�o de la clave",
'"Todas las partes de un PRIMARY KEY deben ser NOT NULL;  Si necesitas NULL en una clave, use UNIQUE",
'"Resultado compuesto de mas que una l�nea",
'"Este tipo de tabla necesita de una primary key",
'"Esta versi�n de MySQL no es compilada con soporte RAID",
'"Tu est�s usando modo de actualizaci�n segura y tentado actualizar una tabla sin un WHERE que usa una KEY columna",
'"Clave '%-.64s' no existe en la tabla '%-.64s'",
'"No puedo abrir tabla",
'"El manipulador de la tabla no permite soporte para check/repair",
'"No tienes el permiso para ejecutar este comando en una transici�n",
'"Obtenido error %d durante COMMIT",
'"Obtenido error %d durante ROLLBACK",
'"Obtenido error %d durante FLUSH_LOGS",
'"Obtenido error %d durante CHECKPOINT",
'"Abortada conexi�n %ld para db: '%-.64s' usuario: '%-.32s' servidor: `%-.64s' (%-.64s)",
'"El manipulador de tabla no soporta dump para tabla binaria",
'"Binlog cerrado mientras tentaba el FLUSH MASTER",
'"Falla reconstruyendo el indice de la tabla dumped '%-.64s'",
'"Error del master: '%-.64s'",
'"Error de red leyendo del master",8
'"Error de red escribiendo para el master",
'"No puedo encontrar �ndice FULLTEXT correspondiendo a la lista de columnas",
'"No puedo ejecutar el comando dado porque tienes tablas bloqueadas o una transici�n activa",
'"Desconocida variable de sistema '%-.64s'",
'"Tabla '%-.64s' est� marcada como crashed y debe ser reparada",
'"Tabla '%-.64s' est� marcada como crashed y la �ltima reparaci�n (automactica?) fall�",
'"Aviso:  Algunas tablas no transancionales no pueden tener rolled back",
'"Multipla transici�n necesita mas que 'max_binlog_cache_size' bytes de almacenamiento. Aumente esta variable mysqld y tente de nuevo',
'"Esta operaci�n no puede ser hecha con el esclavo funcionando, primero use SLAVE ",
'"Esta operaci�n necesita el esclavo funcionando, configure esclavo y haga el SLAVE START",
'"El servidor no est� configurado como esclavo, edite el archivo config file o con CHANGE MASTER TO",
'"No puedo inicializar la estructura info del master, verifique permisiones en el master.info",
'"No puedo crear el thread esclavo, verifique recursos del sistema",
'"Usario %-.64s ya tiene mas que 'max_user_connections' conexiones activas",
'"Tu solo debes usar expresiones constantes con SET",
'"Tiempo de bloqueo de espera excedido",
'"El n�mero total de bloqueos excede el tama�o de bloqueo de la tabla",
'"Bloqueos de actualizaci�n no pueden ser adqueridos durante una transici�n READ UNCOMMITTED",
'"DROP DATABASE no permitido mientras un thread est� ejerciendo un bloqueo de lectura global",
'"CREATE DATABASE no permitido mientras un thread est� ejerciendo un bloqueo de lectura global",
'"Wrong arguments to %s",
'"%-.32s@%-.64s is not allowed to create new users",
'"Incorrect table definition; All MERGE tables must be in the same database",
'"Deadlock found when trying to get lock; Try restarting transaction",
'"The used table type doesn't support FULLTEXT indexes",
'"Cannot add foreign key constraint",
'"Cannot add a child row: a foreign key constraint fails",
'"Cannot delete a parent row: a foreign key constraint fails",
'
