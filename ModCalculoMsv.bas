Attribute VB_Name = "ModCalculoMsv"
Option Explicit

Public Sistema As String

Public Sub CalculoMsv(ByVal Tipo As Byte, ByVal Sist As String)
Dim rs As ADODB.Recordset
Dim rL As ADODB.Recordset
Dim rf As ADODB.Recordset

' Factores de fondo 2 y 3 para extremidad y solapa.
Dim Fond2_ext As String
Dim Fond3_ext As String
Dim Fond2_sol As String
Dim Fond3_sol As String

' Factores de calibración 1 y 2 para anillo y pulsera, y 2 y 3 para solapa.
Dim FCal2_sol As String
Dim FCal3_sol As String
Dim FCal1_ani As String
Dim FCal2_ani As String
Dim FCal1_pul As String
Dim FCal2_pul As String

' Varios para la fórmula.
Dim Tipo_dos As String
Dim Tipo_med As String
Dim Fondo1 As String
Dim Fondo2 As String
Dim Calib1 As String
Dim Calib2 As String
Dim Fact_dos As Single
Dim Fact_lot As Single

Dim sql As String
Dim sql2 As String
Dim sql1 As String
Dim mSv2 As Single
Dim mSv3 As Single
Dim ErrorLectura As Boolean
Dim DosisElevada As Boolean
Dim tabla As String

Dim Observaciones As String
Dim NF As Currency

Dim f_dosis As Date
Dim f_migracion As Date
Dim ndosi As String
Dim punt_error As String
Dim dni_usuario As String
Dim c_empresa As String
Dim c_instalacion As String
Dim c_tipo_trabajo As String
Dim plantilla_contrata As String
Dim n_reg_dosimetro As String
Dim rama_generica As String
Dim rama_especifica As String
Dim dato As String
Dim PanaObj As CPanasonic

On Error GoTo eCalculoMsv

    Sistema = Sist
    'Sist = IIf(Sist = "H", "6600", "Panasonic")
    tabla = IIf(Sistema = "H", "lotes", "lotespana")
    conn.BeginTrans

    ' borramos la tabla auxiliar del listado
    conn.Execute "delete from zlistadomigracion where codusu = " & vUsu.codigo


    sql = "select fecha_lectura, hora_lectura, n_dosimetro, cristal_2, cristal_3, cristal_1, cristal_4 from tempnc "
    sql = sql & " where codusu = " & vUsu.codigo & " and sistema = '" & Sistema & "'"
    sql = sql & " order by n_dosimetro"

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText

    If Not rs.EOF Then

      ' Los de solapa SIEMPRE se cargan, porque los de extremidad de abdomen usan esos factores...
      ' Cargamos factores de fondo de Solapa.
      Fond2_sol = "1"
      Fond3_sol = "1"
      If Not CargarFondos(Fond2_sol, Fond3_sol, "S") Then
        MsgBox "No existe un registro de fondo de Solapa con fecha fin vacia. Revise", vbExclamation, "¡Atención!"
        conn.RollbackTrans
        Exit Sub
      End If
      
      If Tipo = 1 Then
        ' Cargamos factores de fondo de Extremidad.
        Fond2_ext = "1"
        Fond3_ext = "1"
        If Not CargarFondos(Fond2_ext, Fond3_ext, "E") Then
          MsgBox "No existe un registro de fondo de Extremidad con fecha fin vacia. Revise", vbExclamation, "¡Atención!"
          conn.RollbackTrans
          Exit Sub
        End If
      End If
      
      ' Para que estos factores no afecten en la fórmula siendo cálculo de Panasonic.
      FCal2_sol = "1"
      FCal3_sol = "1"
      FCal1_ani = "1"
      FCal2_ani = "1"
      FCal1_pul = "1"
      FCal2_pul = "1"
      
      ' Los factores de calibración sólo valen para el sistema Harshaw.
      If Sistema = "H" Then
        
        ' Los de solapa SIEMPRE se cargan, porque los de extremidad de abdomen usan esos factores...
        ' Cargamos factores de calibración de Solapa.
        If Not CargarFactores(FCal2_sol, FCal3_sol, "S") Then
          MsgBox "No existe un registro de factores 6600 de Solapa con fecha fin vacia. Revise."
          conn.RollbackTrans
          Exit Sub
        End If
        
        If Tipo = 1 Then
          ' Cargamos factores de calibración de Anillo.
          If Not CargarFactores(FCal1_ani, FCal2_ani, "A") Then
            MsgBox "No existe un registro de factores 6600 de Anillo con fecha fin vacia. Revise."
            conn.RollbackTrans
            Exit Sub
          End If
 
          ' Cargamos factores de calibración de Pulsera.
          If Not CargarFactores(FCal1_pul, FCal2_pul, "P") Then
            MsgBox "No existe un registro de factores 6600 de Pulsera con fecha fin vacia. Revise.", vbExclamation, "¡Atención!"
            conn.RollbackTrans
            Exit Sub
          End If
        
        End If
      
      End If
      rs.MoveFirst
    
    End If

    While Not rs.EOF
      ErrorLectura = False
      DosisElevada = False
      Observaciones = ""
      punt_error = ""

      ' ### [DavidV] 03/04/2006: Depende del tipo pasado como argumento,
      ' es personal(0) o área(2).
      ' (VRS 1.2.2) Ojo ahora se leen también los valores de dos cristales más
      ' en el docsímetro.
      sql1 = "select c_empresa, c_instalacion, dni_usuario, c_tipo_trabajo, "
      sql1 = sql1 & "plantilla_contrata, n_reg_dosimetro, cristal_a, cristal_b, tipo_medicion, cristal_c, cristal_d from dosimetros "
      sql1 = sql1 & " where n_dosimetro = '" & Trim(rs.Fields(2).Value) & "' and "
      sql1 = sql1 & " (f_retirada is null or f_retirada = '0000-00-00') and tipo_dosimetro = "
      sql1 = sql1 & Tipo & " and sistema = '" & Sistema & "'"

      Set rL = New ADODB.Recordset
      rL.Open sql1, conn, adOpenDynamic, adLockOptimistic

      If Not rL.EOF Then
        rL.MoveFirst
        ndosi = rL.Fields(5).Value
      Else
        ErrorLectura = True
        Observaciones = "DOSIMETRO NO ENCONTRADO"
        ndosi = "-1"
      End If

      ' Depende del tipo de dosímetro, se usan unos fondos y factores distintos.
      If Tipo <> 1 Then
        ' Cuerpo o Área, son de Solapa.
        Fondo1 = Fond2_sol
        Fondo2 = Fond3_sol
        Calib1 = FCal2_sol
        Calib2 = FCal3_sol
        Tipo_dos = "S"
      Else
        ' Órgano.
        Tipo_dos = "E"
        Fondo1 = Fond2_ext
        Fondo2 = Fond3_ext
        If Not rL.EOF Then Tipo_med = rL!tipo_medicion & "" Else Tipo_med = ""
        If Tipo_med = "" Then
          Tipo_med = DevuelveDesdeBD(1, "tipo_medicion", "dosisnohomog", "n_dosimetro|n_reg_dosimetro|", rs.Fields(2) & "|" & ndosi & "|", "T|N|", 2, , "order by f_migracion desc")
        End If
        
        Select Case Tipo_med
          Case "01", "05"
            ' Pulsera.
            Calib1 = FCal1_pul
            Calib2 = FCal2_pul
          Case "06", "07"
            ' Anillo.
            Calib1 = FCal1_ani
            Calib2 = FCal2_ani
          Case "08"
            ' Abdomen (este es un caso raro de Solapa).
            Fondo1 = Fond2_sol
            Fondo2 = Fond3_sol
            Calib1 = FCal2_sol
            Calib2 = FCal3_sol
            Tipo_dos = "S"
          Case ""
            Calib1 = "1"
            Calib2 = "1"
            Tipo_med = "XX"
          Case Else
            Calib1 = "1"
            Calib2 = "1"
            
        End Select
      End If
      
      ' Si es Panasonic, se ha de aplicar el algoritmo correspondiente.
      If Sistema = "P" Then
        Set PanaObj = New CPanasonic
        With PanaObj
          .E1 = Val(TransformaComasPuntos(rs.Fields(5).Value & ""))
          .E2 = Val(TransformaComasPuntos(rs.Fields(3).Value & ""))
          .E3 = Val(TransformaComasPuntos(rs.Fields(4).Value & ""))
          .E4 = Val(TransformaComasPuntos(rs.Fields(6).Value & ""))
          If Not rL.EOF Then
            .corrE1 = Val(TransformaComasPuntos(rL!cristal_a))
            .corrE2 = Val(TransformaComasPuntos(rL!cristal_b))
            .corrE3 = Val(TransformaComasPuntos(rL!cristal_c))
            .corrE4 = Val(TransformaComasPuntos(rL!cristal_d))
          Else
            .corrE1 = 1
            .corrE2 = 1
            .corrE3 = 1
            .corrE4 = 1
          End If
            dato = DevuelveDesdeBD(1, "cristal_a", "lotespana", "dosimetro_inicial|dosimetro_final|tipo|", "<=" & rs.Fields(2) & "|>=" & rs.Fields(2) & "|" & Tipo_dos & "|", "N|N|T|", 3)
            If dato <> "" Then
              .corrLote = CSng(dato)
            Else
              .corrLote = 1
            End If
          '.procesar (antiguo)
          'If Rs!n_dosimetro = 14368 Then Stop ' Ojo quitar
          .procesar2 rs!n_dosimetro
          mSv2 = .Hs
          mSv3 = .Hd
        End With
        Set PanaObj = Nothing
      Else
        mSv2 = Val(TransformaComasPuntos(rs.Fields(3).Value & ""))
        mSv3 = Val(TransformaComasPuntos(rs.Fields(4).Value & ""))
      End If
      
      ' Calculando el cristal A.
      If rs.Fields(3).Value <> "" Then
        
        ' Obtenemos el valor de corrección del dosímetro para el primer cristal.
        ' (1 por defecto).
        If Not rL.EOF Then
          dato = rL!cristal_a & ""
          If dato <> "" Then
            Fact_dos = CSng(dato)
          Else
            Fact_dos = 1
          End If
        Else
          Fact_dos = 1
        End If
        
        ' Pillamos el factor de corrección de lote para el primer cristal.
        ' (1 por defecto).
        dato = DevuelveDesdeBD(1, "cristal_a", tabla, "dosimetro_inicial|dosimetro_final|tipo|", "<=" & rs.Fields(2) & "|>=" & rs.Fields(2) & "|" & Tipo_dos & "|", "N|N|T|", 3)
        If dato <> "" Then
          Fact_lot = CSng(dato)
        Else
          Fact_lot = 1
        End If
        '-- (VRS 1.2.2) Ahora se diferencia entre Panasonic y no Panasonic, en el primer caso se aplican los
        '   factores de correción antes.
        If Sistema = "P" Then
            mSv2 = Round2((mSv2 - CSng(Fondo1)) * CSng(Calib1), 3)
        Else
            mSv2 = Round2((mSv2 - CSng(Fondo1)) * CSng(Calib1) * Fact_dos * Fact_lot, 3)
        End If
      End If
      
      ' Calculando el cristal B.
      If rs.Fields(4).Value <> "" Then
        
        ' Obtenemos el valor de corrección del dosímetro para el segundo cristal.
        ' (1 por defecto).
        If Not rL.EOF Then
          dato = rL!cristal_b & ""
          If dato <> "" Then
            Fact_dos = CSng(dato)
          Else
            Fact_dos = 1
          End If
        Else
          Fact_dos = 1
        End If

        ' Pillamos el factor de corrección de lote para el segundo cristal.
        ' (1 por defecto).
        dato = DevuelveDesdeBD(1, "cristal_b", tabla, "dosimetro_inicial|dosimetro_final|tipo|", "<=" & rs.Fields(2) & "|>=" & rs.Fields(2) & "|" & Tipo_dos & "|", "N|N|T|", 3)
        If dato <> "" Then
          Fact_lot = CSng(dato)
        Else
          Fact_lot = 1
        End If
        '-- (VRS 1.2.2) Ahora se diferencia entre Panasonic y no Panasonic, en el primer caso se aplican los
        '   factores de correción antes.
        If Sistema = "P" Then
            mSv3 = Round2((mSv3 - CSng(Fondo2)) * CSng(Calib2), 3)
        Else
            mSv3 = Round2((mSv3 - CSng(Fondo2)) * CSng(Calib2) * Fact_dos * Fact_lot, 3)
        End If
      End If

      If mSv2 < 0.1 Then mSv2 = 0
      If mSv3 < 0.1 Then mSv3 = 0

      ' datos de la rama especifica y generica
      
      
      If Not ErrorLectura Then

        sql1 = "select rama_gen, rama_especifica from instalaciones where c_instalacion = '"
        sql1 = sql1 & Trim(rL.Fields(1).Value) & "'"

        Set rf = New ADODB.Recordset

        rf.Open sql1, conn, adOpenDynamic, adLockOptimistic

        If Not rf.EOF Then rf.MoveFirst

      End If
      NF = SugerirCodigoSiguiente(Tipo)

      f_dosis = DateAdd("m", -1, CDate(rs.Fields(0).Value))
      f_migracion = Now

      If mSv2 > 4 Or mSv3 > 4 Then
        Observaciones = "DOSIS ELEVADA"
        DosisElevada = True
      End If

      If ErrorLectura Or DosisElevada Then
        sql2 = "insert into erroresmigra (n_registro, descripcion, c_tipo) VALUES ("
        sql2 = sql2 & ImporteSinFormato(CStr(NF)) & ",'" & Trim(Observaciones) & "'," & Format(Tipo, "0") & ")"

        conn.Execute sql2
      End If
      If ErrorLectura Then
        punt_error = "**"
        dni_usuario = "999999999"
        c_empresa = "DESCON"
        c_instalacion = "DESCON"
        c_tipo_trabajo = "99"
        plantilla_contrata = "00"
        n_reg_dosimetro = ""
        rama_generica = "99"
        rama_especifica = "99"
      Else
        If DosisElevada Then punt_error = "**"
        dni_usuario = rL.Fields(2).Value
        c_empresa = rL.Fields(0).Value
        c_instalacion = rL.Fields(1).Value
        c_tipo_trabajo = rL.Fields(3).Value
        plantilla_contrata = rL.Fields(4).Value
        n_reg_dosimetro = ndosi
        rama_generica = rf.Fields(0).Value
        rama_especifica = rf.Fields(1).Value
      End If

      If Tipo = 0 Then
        'personal (homogéneas)
        sql2 = "dosiscuerpo"
      ElseIf Tipo = 2 Then
        'área
        sql2 = "dosisarea"
      End If

      If Tipo = 1 Then
        sql2 = "insert into dosisnohomog (n_registro, n_dosimetro, c_empresa, c_instalacion, "
        sql2 = sql2 & " dni_usuario, f_dosis, f_migracion, tipo_medicion, dosis_org, "
        sql2 = sql2 & " plantilla_contrata, rama_generica, rama_especifica, c_tipo_trabajo, "
        sql2 = sql2 & " observaciones, migrado, n_reg_dosimetro) values ("
        sql2 = sql2 & ImporteSinFormato(CStr(NF)) & ",'" & Trim(rs.Fields(2).Value) & "','"  'n_dosimetro
        sql2 = sql2 & Trim(c_empresa) & "','"          ' empresa
        sql2 = sql2 & Trim(c_instalacion) & "','"          ' instalacion
        sql2 = sql2 & Trim(dni_usuario) & "','"          ' dni de usuario
        sql2 = sql2 & Format(f_dosis, FormatoFecha) & "','"     'fecha de dosis
        sql2 = sql2 & Format(f_migracion, FormatoFecha) & "',"  ' fecha de migracion
        sql2 = sql2 & "'" & Tipo_med & "',"   'dosis superficial
      Else
        sql2 = "insert into " & sql2 & " (n_registro, n_dosimetro, c_empresa, c_instalacion, "
        sql2 = sql2 & " dni_usuario, f_dosis, f_migracion, dosis_superf, dosis_profunda, "
        sql2 = sql2 & " plantilla_contrata, rama_generica, rama_especifica, c_tipo_trabajo, "
        sql2 = sql2 & " observaciones, migrado, n_reg_dosimetro) values ("
        sql2 = sql2 & ImporteSinFormato(CStr(NF)) & ",'" & Trim(rs.Fields(2).Value) & "','"  'n_dosimetro
        sql2 = sql2 & Trim(c_empresa) & "','"          ' empresa
        sql2 = sql2 & Trim(c_instalacion) & "','"          ' instalacion
        sql2 = sql2 & Trim(dni_usuario) & "','"          ' dni de usuario
        sql2 = sql2 & Format(f_dosis, FormatoFecha) & "','"     'fecha de dosis
        sql2 = sql2 & Format(f_migracion, FormatoFecha) & "',"  ' fecha de migracion
        If Sistema = "P" And (Tipo = 0 Or Tipo = 2) Then ' solo cuando el sistema es panasonic (v 2.0.2) personal o área
            sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv2))) & ","   'tipo medicion (V 2.0.1)
        Else
            sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv3))) & ","   'tipo medicion
        End If
      End If
      If Sistema = "P" And (Tipo = 0 Or Tipo = 2) Then ' solo cuando es panasonic (v 2.0.2) personal o área
          sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv3))) & ",'"  'dosis profunda (v 2.0.1)
      Else
          sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv2))) & ",'"  'dosis profunda (RF01)
      End If
      sql2 = sql2 & Trim(plantilla_contrata) & "','"  'plantilla contrata
      sql2 = sql2 & Trim(rama_generica) & "','"  'rama generica
      sql2 = sql2 & Trim(rama_especifica) & "','"  'rama especifica
      sql2 = sql2 & Trim(c_tipo_trabajo) & "','"  ' tipo de trabajo
      sql2 = sql2 & Trim(Observaciones) & "',null,'"  'observaciones, migrado
      sql2 = sql2 & Trim(n_reg_dosimetro) & "')"  'n_reg_dosimetro
      conn.Execute sql2
      
      ' tenemos que insertar en la tabla temporal para poder imprimir
      ' (V 2.0.1) Aunque los valores mSv2 corresponde a superficial y mSv3 a profunda lo dejamos "mal" por como está el listado
      sql2 = "insert into zlistadomigracion (codusu, n_registro, n_dosimetro, dni_usuario, cristal2,"
      sql2 = sql2 & "cristal3, f_migracion, punt_error) values (" & vUsu.codigo & ","
      sql2 = sql2 & ImporteSinFormato(CStr(NF)) & ",'"  'numero de registro
      sql2 = sql2 & Trim(rs.Fields(2).Value) & "','" & Trim(dni_usuario) & "',"
      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv2))) & ","  'dosis profunda
      sql2 = sql2 & TransformaComasPuntos(ImporteSinFormato(CStr(mSv3))) & ",'"   'dosis superficial
      sql2 = sql2 & Format(f_migracion, FormatoFecha) & "','"  ' fecha de migracion
      sql2 = sql2 & Trim(punt_error) & "')"
      conn.Execute sql2

      Set rL = Nothing
      Set rf = Nothing
      rs.MoveNext

    Wend

    Set rs = Nothing

eCalculoMsv:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en el cálculo de mSvs"
        
        conn.RollbackTrans
        Screen.MousePointer = vbDefault
    Else
        conn.CommitTrans
    End If

End Sub

Public Function CargarFondos(ByRef Fondo1 As String, ByRef Fondo2 As String, ByVal Tipo As String) As Boolean
Dim rf As ADODB.Recordset
Dim sql As String
Dim tabla As String

    CargarFondos = False
    tabla = IIf(Sistema = "H", "fondos", "fondospana")
    sql = "select fondo_2, fondo_3 from " & tabla & " where f_fin is null and tipo = '" & Tipo & "'"
    Set rf = New ADODB.Recordset
    
    rf.Open sql, conn, adOpenDynamic, adLockOptimistic
    If Not rf.EOF Then
        rf.MoveFirst
        Fondo1 = rf.Fields(0).Value
        Fondo2 = rf.Fields(1).Value
        CargarFondos = True
    End If
    rf.Close
    Set rf = Nothing
End Function

Private Function CargarFactores(ByRef Factor1 As String, ByRef Factor2 As String, ByVal Tipo As String) As Boolean
Dim rf As ADODB.Recordset
Dim sql As String
Dim tabla As String

    CargarFactores = False
    tabla = IIf(Sistema = "H", "factcali6600", "factcalipana")
    sql = "select cristal_a, cristal_b, f_inicio from " & tabla & " where f_fin is null and tipo = '" & Tipo & "'"
    Set rf = New ADODB.Recordset
    
    rf.Open sql, conn, adOpenDynamic, adLockOptimistic
    If Not rf.EOF Then
        rf.MoveFirst
        Factor1 = rf.Fields(0).Value
        Factor2 = rf.Fields(1).Value
        CargarFactores = True
    End If
End Function

Private Function SugerirCodigoSiguiente(Tipo As Byte) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    If Tipo = 0 Then
        sql = "Select Max(n_registro) from dosiscuerpo "
    ElseIf Tipo = 1 Then
        sql = "Select Max(n_registro) from dosisnohomog "
    Else
        sql = "Select Max(n_registro) from dosisarea "
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, , , adCmdText
    sql = "1"
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then
            sql = CStr(rs.Fields(0) + 1)
        End If
    End If
    rs.Close
    SugerirCodigoSiguiente = sql
End Function

