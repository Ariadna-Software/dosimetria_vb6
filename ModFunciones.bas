Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const ValorNulo = "Null"

Public Function CompForm(ByRef Formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Carga As Boolean
    Dim Correcto As Boolean
       
    CompForm = False
    Set mTag = New CTag
    For Each Control In Formulario.Controls

        'TEXT BOX
        If TypeOf Control Is TextBox Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag, "¡Error!"
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag, "¡Error!"
                    Exit Function
                    
                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                       If Control.Text = "" Then
                            MsgBox "Seleccione una dato para: " & mTag.nombre, vbExclamation, "¡Error!"
                            Exit Function
                       End If
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True

End Function


Public Sub Limpiar(ByRef Formulario As Form)
    Dim Control As Object
    
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub


Public Function CampoSiguiente(ByRef Formulario As Form, valor As Integer) As Control
Dim FIN As Boolean
Dim Control As Object

On Error GoTo ECampoSiguiente

    'Debug.Print "Llamada:  " & Valor
    'Vemos cual es el siguiente
    Do
        valor = valor + 1
        For Each Control In Formulario.Controls
            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
            'Si es texto monta esta parte de sql
            If Control.TabIndex = valor Then
                    Set CampoSiguiente = Control
                    FIN = True
                    Exit For
            End If
        Next Control
        If Not FIN Then
            valor = -1
        End If
    Loop Until FIN
    Exit Function
ECampoSiguiente:
    Set CampoSiguiente = Nothing
    Err.Clear
End Function




Private Function ValorParaSQL(valor, ByRef vTag As CTag) As String
Dim Dev As String
Dim d As Single
Dim V
    Dev = ""
    If valor <> "" Then
        Select Case vTag.TipoDato
        Case "N"
            V = valor
            If InStr(1, valor, ",") Then
                V = CSng(valor)
                valor = V
            End If
            Dev = TransformaComasPuntos(CStr(valor))
            
        Case "F"
            Dev = "'" & Format(valor, FormatoFecha) & "'"
        Case "T"
            Dev = CStr(valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & valor & "'"
        End Select
        
    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vTag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function

Public Function InsertarDesdeForm(ByRef Formulario As Form, Bd As Byte) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Izda As String
    Dim Der As String
    Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
'        Debug.Print Control.Tag
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.Columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        If mTag.TipoDato = "T" Then
                            Cad = "'" & DevNombreSQL(Control.List(Control.ListIndex)) & "'"
                        Else
                            Cad = Control.ItemData(Control.ListIndex)
                        End If
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    
   
    Select Case Bd
        Case 1
            ' suministros
            Conn.Execute Cad, , adCmdText
'        Case 2
'            ' contabilidad
'            ConnConta.Execute Cad, , adCmdText
'        Case 3
'            ' gestion social
'            ConnGestion.Execute Cad, , adCmdText
    End Select
    
    
    
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function PonerCamposForma(ByRef Formulario As Form, ByRef vData As Adodc) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Cad As String
    Dim valor As Variant
    Dim Campo As String  'Campo en la base de datos
    Dim I As Integer
    
    'On Error GoTo EPonerCamtrol.TagposForma '#QUITAR
    'Exit Function
    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In Formulario.Controls
        'TEXTO
        'Debug.Print Control.Tag
'        Debug.Print Control.Name
'
        If TypeOf Control Is TextBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.Columna <> "" Then
                        Campo = mTag.Columna
                        If mTag.Vacio = "S" Then
                            valor = DBLet(vData.Recordset.Fields(Campo))
                        Else
                            valor = vData.Recordset.Fields(Campo) & ""
                        End If
                        If mTag.Formato <> "" And CStr(valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    Campo = mTag.Columna
                    valor = vData.Recordset.Fields(Campo)
                    Else
                        valor = 0
                End If
                Control.Value = valor
            End If
            
         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Campo = mTag.Columna
                    valor = vData.Recordset.Fields(Campo)
                    
                    If mTag.TipoDato = "T" Then
'                        ' MODIFICADO
'                       i = 0
'                        For i = 0 To Control.ListCount - 1
'                            If Control.ItemData(i) = val(valor) Then
'                                Control.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                        If i = Control.ListCount Then Control.ListIndex = -1
                     
                    
                    
                        Control.Text = valor
                    Else
                        I = 0
                        For I = 0 To Control.ListCount - 1
                            If Control.ItemData(I) = Val(valor) Then
                                Control.ListIndex = I
                                Exit For
                            End If
                        Next I
                        If I = Control.ListCount Then Control.ListIndex = -1
                    End If
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control
    
    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

Private Function ObtenerMaximoMinimo(ByRef vSql As String) As String
Dim Rs As Recordset
ObtenerMaximoMinimo = ""
Set Rs = New ADODB.Recordset
Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.EOF) Then
        ObtenerMaximoMinimo = CStr(Rs.Fields(0))
    End If
End If
Rs.Close
Set Rs = Nothing
End Function


Public Function ObtenerBusqueda(ByRef Formulario As Form) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim sql As String
    Dim tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = DevNombreSQL(Trim(Control.Text))
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        Cad = " MAX(" & mTag.Columna & ")"
                    Else
                        Cad = " MIN(" & mTag.Columna & ")"
                    End If
                    sql = "Select " & Cad & " from " & mTag.tabla
                    sql = ObtenerMaximoMinimo(sql)
                    Select Case mTag.TipoDato
                    Case "N"
                        sql = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(sql)
                    Case "F"
                        sql = mTag.tabla & "." & mTag.Columna & " = '" & Format(sql, "yyyy-mm-dd") & "'"
                    Case Else
                        sql = mTag.tabla & "." & mTag.Columna & " = '" & sql & "'"
                    End Select
                    sql = "(" & sql & ")"
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = DevNombreSQL(Trim(Control.Text))
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    sql = mTag.tabla & "." & mTag.Columna & " is NULL"
                    sql = "(" & sql & ")"
                    Control.Text = ""
                End If
            Else
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If mTag.Columna = "migrado" Then
                        Select Case UCase(Aux)
                            Case "**"
                                sql = mTag.tabla & "." & mTag.Columna & " = '**'"
                                sql = "(" & sql & ")"
                                Control.Text = ""
                            
                            Case "*"
                                sql = mTag.tabla & "." & mTag.Columna & " = '*'"
                                sql = "(" & sql & ")"
                                Control.Text = ""
                        End Select
                    End If
                End If
            End If
        End If
    Next
    

    'Recorremos los textbox
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
           If Control.Enabled = True Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            
            
            If Carga Then
                Aux = DevNombreSQL(Trim(Control.Text))
                If Aux <> "" Then
                    If mTag.tabla <> "" Then
                        tabla = mTag.tabla & "."
                        Else
                        tabla = ""
                    End If
                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, Cad)
                    If RC = 0 Then
                        If sql <> "" Then sql = sql & " AND "
                        sql = sql & Cad
                        'sql = sql & "(" & Cad & ")"
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag, "¡Error!"
                Exit Function
            End If
           End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Enabled Then
        If Control.Text = "Panasonic" Then
          Debug.Print
        End If
            
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato = "T" Then
                            Cad = Control.List(Control.ListIndex)
                            Cad = mTag.tabla & "." & mTag.Columna & " = '" & Cad & "'"
                        Else
                            Cad = Control.ItemData(Control.ListIndex)
                            Cad = mTag.tabla & "." & mTag.Columna & " = " & Cad
                        End If
                        If sql <> "" Then sql = sql & " AND "
                        sql = sql & "(" & Cad & ")"
                    End If
                End If
            End If
        End If
    Next Control
    ObtenerBusqueda = sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function




Public Function ModificaDesdeFormulario(ByRef Formulario As Form, Bd As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim Campos As Variant
Dim Valores As Variant

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            'Debug.Print Control.Tag
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.Columna & " = " & Aux & ")"
                             
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                  If Control.ListIndex = -1 Then
                      Aux = ValorNulo
                  Else
'                        Aux = Control.ItemData(Control.ListIndex)
                      If mTag.TipoDato = "T" Then
                          Aux = "'" & Control.List(Control.ListIndex) & "'"
                      Else
                          Aux = Control.ItemData(Control.ListIndex)
                      End If
                  End If
                  
                  If mTag.EsClave Then
                    'Lo pondremos para el WHERE
                    If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                    cadWHERE = cadWHERE & "(" & mTag.Columna & " = " & Aux & ")"
                  Else
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                  End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        
        If Formulario.Name = "frmDosisExtremidades" Then
          With Formulario
            .adodc1.Recordset.Fields("n_dosimetro") = .txtAux(0).Text
            .adodc1.Recordset.Fields("cristal_2") = TransformaComasPuntos(.txtAux(1).Text)
            .adodc1.Recordset.Fields("fecha_lectura") = Format(.txtAux(2).Text, "yyyy-MM-dd")
            .adodc1.Recordset.Update
          End With
        Else
          MsgBox "No se ha definido ninguna clave principal.", vbExclamation, "¡Error!"
          Exit Function
        End If
    Else
      Aux = "UPDATE " & mTag.tabla
      Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    
      Select Case Bd
         Case 1
            Conn.Execute Aux, , adCmdText
      End Select
    End If

ModificaDesdeFormulario = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim Cad As String

'Montamos al final: "Cod Diag.|idDiag|N|10·"

ParaGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Desc <> "" Then
                Cad = Desc
            Else
                Cad = mTag.nombre
            End If
            Cad = Cad & "|"
            Cad = Cad & mTag.Columna & "|"
            Cad = Cad & mTag.TipoDato & "|"
            Cad = Cad & AnchoPorcentaje & "·"
            
                
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            
        ElseIf TypeOf Control Is ComboBox Then
                If Desc <> "" Then
                    Cad = Desc
                Else
                    Cad = mTag.nombre
                End If
                Cad = Cad & "|"
                Cad = Cad & mTag.Columna & "|"
                Cad = Cad & mTag.TipoDato & "|"
                Cad = Cad & AnchoPorcentaje & "·"
        
        End If 'De los elseif
    End If
Set mTag = Nothing
ParaGrid = Cad
End If



End Function

'////////////////////////////////////////////////////
' Monta a partir de una cadena devuelta por el formulario
'de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, orden As Integer) As String
Dim mTag As CTag
Dim Cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

ValorDevueltoFormGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            Aux = RecuperaValor(CadenaDevuelta, orden)
            If Aux <> "" Then Cad = mTag.Columna & " = " & ValorParaSQL(Aux, mTag)
                
            
            
                
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
       
        ElseIf TypeOf Control Is ComboBox Then
            Aux = RecuperaValor(CadenaDevuelta, orden)
            If Aux <> "" Then
                Cad = mTag.Columna & " = " & ValorParaSQL(Aux, mTag) & ""
            End If
        End If 'De los elseif
    End If
End If
Set mTag = Nothing
ValorDevueltoFormGrid = Cad
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

I = 0
Cont = 1
Cad = ""
Do
    J = I + 1
    I = InStr(J, CADENA, "|")
    If I > 0 Then
        If Cont = orden Then
            Cad = Mid(CADENA, J, I - J)
            I = Len(CADENA) 'Para salir del bucle
            Else
                Cont = Cont + 1
        End If
    End If
Loop Until I = 0
RecuperaValor = Cad

End Function




'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef Formulario As Form)
Dim I As Integer
Dim J As Integer
Dim ctrl As Variant

On Error GoTo EPonerOpcionesMenuGeneral

'Añadir, modificar y borrar deshabilitados si no nivel
With Formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J > vUsu.NivelUsu Then
                .Toolbar1.Buttons(I).Visible = False
            End If
        End If
    Next I
    
    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next
    
    On Error Resume Next
    
    'Los MENUS
      
    ' ### DavidV: Se configuran los menús según el nivel de usuario.
    For Each ctrl In .Controls
      If TypeName(ctrl) = "Menu" And Left(ctrl.Name, 2) = "mn" Then
        If Val(ctrl.HelpContextID) > vUsu.NivelUsu Then ctrl.Visible = False
      End If
    Next ctrl
    
    On Error GoTo 0
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef Formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadWHERE = Claves
    'Construimos el SQL
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation, "¡Error!"
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    Conn.Execute Aux, , adCmdText






ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function







Public Function BLOQUEADesdeFormulario(ByRef Formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                         cadWHERE = cadWHERE & "(" & mTag.Columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control
    
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation, "¡Error!"
        
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"
        
        'Intenteamos bloquear
        PreparaBloquear
        Conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function

Public Function BloqueaRegistroForm(ByRef Formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    'Debug.Print Control.Name
 '   Debug.Print Control.Index
    'Debug.Print Control.Tag
    
    Next Control
    
    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation, "¡Error!"
        
    Else
        Aux = "Insert into zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        Conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation, "¡Error!"
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim sql As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
        sql = "DELETE from zbloqueos where codusu=" & vUsu.codigo & " and tabla='" & mTag.tabla & "'"
        Conn.Execute sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function


Public Function DesbloqueaRegistroForm1(ByRef Formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EDesBLOQ
    DesbloqueaRegistroForm1 = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    'Debug.Print Control.Name
 '   Debug.Print Control.Index
    'Debug.Print Control.Tag
    
    Next Control
    
    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation, "¡Error!"
        
    Else
        Aux = "delete from zbloqueos where codusu = " & vUsu.codigo & " and tabla = '" & mTag.tabla
        Aux = Aux & "' and clave = '" & DevNombreSQL(AuxDef) & "'"
        Conn.Execute Aux
        DesbloqueaRegistroForm1 = True
    End If
EDesBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 2013 Then
                '¡se ha perdido de la conexion
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Desbloqueo tabla"
        Else
            MsgBox "Se ha perdido la conexion", vbExclamation, "¡Error!"
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim ent As Integer
Dim Cad As String
  
  ' Comprobaciones
  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If
  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If
  
  ' Redondeo.
  Cad = "0"
  If NumDigitsAfterDecimals <> 0 Then Cad = Cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Val(TransformaComasPuntos(Format(Number, Cad)))
  
End Function

