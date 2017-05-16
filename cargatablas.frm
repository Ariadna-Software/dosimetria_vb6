VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Cargatablas 
   Caption         =   "Carga de tablas"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "cargatablas.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Pb2 
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   1590
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdCan 
      Caption         =   "Cancelar"
      Height          =   555
      Left            =   2520
      TabIndex        =   5
      Top             =   2310
      Width           =   1095
   End
   Begin VB.CommandButton CmdAcep 
      Caption         =   "Aceptar"
      Height          =   555
      Left            =   1110
      TabIndex        =   4
      Top             =   2310
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   180
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1110
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   870
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "Directorio:"
      Height          =   255
      Left            =   270
      TabIndex        =   2
      Top             =   210
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Tabla:"
      Height          =   255
      Left            =   270
      TabIndex        =   1
      Top             =   900
      Width           =   855
   End
End
Attribute VB_Name = "Cargatablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim CADENA As String
Dim Cad As String
Dim cad1 As String
Dim NF As Integer
Dim Tam2 As Long
Dim tamanyo As Long
Dim Primeralinea As Boolean
Dim Rs As ADODB.Recordset

Private Sub CmdAcep_Click()

On Error GoTo eCmdAceptar

    Select Case Combo1.ListIndex
        Case 0
            CADENA = Trim(Text1.Text) & "provinci.unl"
        Case 1
            CADENA = Trim(Text1.Text) & "ramagene.unl"
        Case 2
            CADENA = Trim(Text1.Text) & "ramaespe.unl"
        Case 3
            CADENA = Trim(Text1.Text) & "tipostra.unl"
        Case 4
            CADENA = Trim(Text1.Text) & "tipmedex.unl"
        Case 5
            CADENA = Trim(Text1.Text) & "fondos.unl"
        Case 6
            CADENA = Trim(Text1.Text) & "factcal4.unl"
        Case 7
            CADENA = Trim(Text1.Text) & "factcali.unl"
        Case 8
            CADENA = Trim(Text1.Text) & "empresas.unl"
        Case 9
            CADENA = Trim(Text1.Text) & "instalac.unl"
        Case 10
            CADENA = Trim(Text1.Text) & "operario.unl"
        Case 11
            CADENA = Trim(Text1.Text) & "operario.unl"
        Case 12
            CADENA = Trim(Text1.Text) & "dosimetr.unl"
        Case 13
            CADENA = Trim(Text1.Text) & "dosimorg.unl"
        Case 14
            CADENA = Trim(Text1.Text) & "dosiscue.unl"
        Case 15
            CADENA = Trim(Text1.Text) & "dosisnoh.unl"
        Case 16 ' empresa de area
            CADENA = Trim(Text1.Text) & "empresas.unl"
        Case 17 ' instalaciones de area
            CADENA = Trim(Text1.Text) & "instalac.unl"
        Case 18
            InsertaRegistro
            Exit Sub
        Case 19 ' dosimetros de area
            CADENA = Trim(Text1.Text) & "dosimetr.unl"
        Case 20 'dosis de area
            CADENA = Trim(Text1.Text) & "dosiscue.unl"
        Case 21 ' tabla recepdosim de cuerpo
            CADENA = Trim(Text1.Text) & "recepdos.unl"
        Case 22 ' tabla recepdosim de area
            CADENA = Trim(Text1.Text) & "recepdos.unl"
        
    End Select

    NF = FreeFile
    tamanyo = FileLen(CADENA)
    pb2.Value = 0
    Primeralinea = True

    Open Trim(CADENA) For Input As #NF
    
    While Not EOF(NF)
        Line Input #NF, Cad
        
        If Primeralinea Then
            Primeralinea = False
            If Cad <> "" Then
                Tam2 = Len(Cad) + 2
            Else
                Tam2 = 20
            End If
            Tam2 = tamanyo \ Tam2
            If Tam2 < 32000 Then
                pb2.max = Tam2 + 2
            Else
                pb2.max = 32100
            End If
        End If
        
        If Cad <> "" Then InsertaRegistro
        
    Wend

    Close #NF
    MsgBox "Terminado", vbExclamation, "Carga de Tablas."
    
eCmdAceptar:
    If Err.Number <> 0 Then
        MsgBox "Error insertando en " & Combo1.Text & " registro: " & Cad & Err.Description, vbCritical, "¡Error!"
        Close #NF
    End If

End Sub

Private Sub CmdCan_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim Cad As String
    
    Text1.Text = "c:\mbgstld4\datos\"
    
    CargarCombo


End Sub


Private Sub InsertaRegistro()
Dim sql As String
Dim sql1 As String
Dim fec As String
Dim dni As String
Dim existe As String
Dim Empresa As String
Dim Instala As String


'On Error GoTo eInsertarRegistro

    Select Case Combo1.ListIndex
        Case 0
            fec = RecuperaValor(Cad, 1)
            fec = Trim(Format(CInt(fec), "00"))
        
            sql = "insert into provincias values ('"
            sql = sql & Trim(fec) & "','"
            sql = sql & RevisaCaracterMultibase(Trim(RecuperaValor(Cad, 2))) & "',"
            If RecuperaValor(Cad, 3) = "" Then
                sql = sql & "null)"
            Else
                sql = sql & "'" & Trim(RecuperaValor(Cad, 3)) & "')"
            End If
            
            Conn.Execute sql
            
        Case 1
            sql = "insert into ramagene values ('"
            sql = sql & Trim(Format(CInt(RecuperaValor(Cad, 1)), "00")) & "','"
            sql = sql & RevisaCaracterMultibase(Trim(RecuperaValor(Cad, 2))) & "')"
            
            Conn.Execute sql
            
        Case 2
            sql = "insert into ramaespe values ('"
            sql = sql & Trim(Format(RecuperaValor(Cad, 1), "00")) & "','"
            sql = sql & Trim(Format(RecuperaValor(Cad, 2), "00")) & "','"
            sql = sql & RevisaCaracterMultibase(Trim(RecuperaValor(Cad, 3))) & "')"
            
            Conn.Execute sql
        
        Case 3
            sql = "insert into tipostrab values ('"
            sql = sql & Trim(Format(RecuperaValor(Cad, 1), "00")) & "','"
            sql = sql & Trim(Format(RecuperaValor(Cad, 2), "00")) & "','"
            sql = sql & RevisaCaracterMultibase(Trim(RecuperaValor(Cad, 3))) & "')"
            
            Conn.Execute sql
           
        Case 4
            sql = "insert into tipmedext values ('"
            sql = sql & Trim(Format(RecuperaValor(Cad, 1), "00")) & "','"
            sql = sql & RevisaCaracterMultibase(Trim(RecuperaValor(Cad, 2))) & "')"
            Conn.Execute sql
            
        Case 5
            sql = "insert into fondos values ("
            sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(RecuperaValor(Cad, 1)))) & ","
            sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(RecuperaValor(Cad, 2)))) & ","
            sql = sql & "'" & Format(RecuperaValor(Cad, 3), FormatoFecha) & "',"
            fec = RecuperaValor(Cad, 4)
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "','S')"
            Else
                sql = sql & "null,'S')"
            End If
            
            Conn.Execute sql
            
        Case 6
            sql = "insert into factcali4400 values ("
            sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(RecuperaValor(Cad, 1)))) & ","
            sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(RecuperaValor(Cad, 2)))) & "'"
            sql = sql & "'" & Format(RecuperaValor(Cad, 3), FormatoFecha) & "',"
            fec = RecuperaValor(Cad, 4)
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "')"
            Else
                sql = sql & "null)"
            End If
            
            Conn.Execute sql
            
        Case 7
            sql = "insert into factcali6600 values ("
            sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(RecuperaValor(Cad, 1)))) & ","
            sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(RecuperaValor(Cad, 2)))) & ","
            sql = sql & "'" & Format(RecuperaValor(Cad, 3), FormatoFecha) & "',"
            fec = RecuperaValor(Cad, 4)
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "','S')"
            Else
                sql = sql & "null,'S')"
            End If
            
            Conn.Execute sql
        
        Case 8, 16
            sql = "insert into empresas values ('"
        
        
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 1))
            
            If Empresa = " " Then Empresa = "DESCON"
        
            sql = sql & Empresa & "',"
            
            fec = RecuperaValor(Cad, 2) 'fecha alta
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format("01/01/1900", FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 3) 'fecha baja
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            sql = sql & "'" & Trim(RecuperaValor(Cad, 4)) & "'," 'cif/nif
            
            ' nombre comercial
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 5))
            NombreSQL fec
            sql = sql & "'" & Trim(fec) & "',"
            
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 6)) ' direccion
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 7)) ' poblacion
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 8) ' c.postal
            If fec <> "" Then
                sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            Else
                sql = sql & "'46',"
            End If
            fec = RecuperaValor(Cad, 9) ' distrito
            If fec <> "" Then
                sql = sql & "'" & Trim(Format(CInt(fec), "000")) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 10) 'telefono
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 11) ' fax
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 12)) ' persona contacto
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 13) 'migrado
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 14) 'mail internet
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            If Combo1.ListIndex = 8 Then
                sql = sql & "0)" ' tipo de empresa 0 personal
                                                  '1 Area
            Else
                ' miramos que la empresa no exista en empresas
                ' si existe tipo = 2 else tipo = 1
'                Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 1))
                sql1 = ""
                sql1 = DevuelveDesdeBD(1, "c_empresa", "empresas", "c_empresa|", Trim(Empresa) & "|", "T|", 1)
                If sql1 = "" Then
                    sql = sql & "1)"
                Else
                    sql = "update empresas set c_tipo = 2 where c_empresa = '" & Trim(Empresa)
                    sql = sql & "' "
                End If
            End If
            
            Conn.Execute sql
            
        Case 9, 17
            sql = "insert into instalaciones values ('"
        
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 1))
            If Empresa = " " Then Empresa = "DESCON"
            
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 2))
            If Instala = " " Then Instala = "DESCON"
        
        
            sql = sql & Empresa & "','" ' empresa
            sql = sql & Instala & "',"  ' instalacion
            
            fec = RecuperaValor(Cad, 3) 'fecha alta
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format("01/01/1900", FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 4) 'fecha baja
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'descripcion
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 5))
            NombreSQL fec
            sql = sql & "'" & Trim(fec) & "',"
            
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 6)) ' direccion
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 7)) ' poblacion
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 8) ' c.postal
            If fec <> "" Then
                sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            Else
                sql = sql & "'46',"
            End If
            fec = RecuperaValor(Cad, 9) ' distrito
            If fec <> "" Then
                sql = sql & "'" & Trim(Format(CInt(fec), "000")) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 10) 'telefono
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 11) ' fax
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 12)) ' persona contacto
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 13) 'migrado
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 14) 'rama generica
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            fec = RecuperaValor(Cad, 15) 'rama especifica
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            fec = RecuperaValor(Cad, 16) 'mail internet
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            ' observaciones campo que no estaba anteriormente
            sql = sql & "null,"
            If Combo1.ListIndex = 9 Then
                sql = sql & "0)" ' tipo de empresa 0 personal
                                 '                 1 area
            Else
                Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 1))     ' empresa
                If Empresa = " " Then Empresa = "DESCON"
                Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 2)) ' instalacion
                If Instala = " " Then Instala = "DESCON"
                fec = RecuperaValor(Cad, 3) 'fecha alta
                
                sql1 = ""
                sql1 = DevuelveDesdeBD(1, "c_instalacion", "instalaciones", "c_empresa|c_instalacion|", Empresa & "|" & Instala & "|" & Format(fec, FormatoFecha), "T|T|", 2)
                If sql1 = "" Then
                    sql = sql & "1)"
                Else
                    sql = "update instalaciones set c_tipo = 2 where c_empresa = '" & Trim(Empresa)
                    sql = sql & "' and c_instalacion = '" & Trim(Instala) & "'"
                End If
            End If
            Conn.Execute sql
        
        Case 10 ' operarios
            ' comprobamos que un dni esté únicamente una vez
            
            dni = RecuperaValor(Cad, 2)
            fec = RecuperaValor(Cad, 18) 'fecha alta
            
            sql = ""
            sql = DevuelveDesdeBD(1, "dni", "operarios", "dni|f_alta|", dni & "|" & Format(fec, FormatoFecha) & "|", "T|F|", 2)
            If sql = "" Then
                ' si no existe lo insertamos
                
                sql = "insert into operarios values('"
            
                sql = sql & RecuperaValor(Cad, 2) & "'," ' dni operario
                
                 ' nsegsocial
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 3))
                If fec <> "" Then
                    NombreSQL fec
                    sql = sql & "'" & Trim(fec) & "',"
                Else
                    sql = sql & "null,"
                End If
                
                ' n carnet radiologico
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
                If fec <> "" Then
                    NombreSQL fec
                    sql = sql & "'" & Trim(fec) & "',"
                Else
                    sql = sql & "null,"
                End If
                
                fec = RecuperaValor(Cad, 5) 'fecha emision carnet radilogico
                If fec <> "" Then
                    sql = sql & "'" & Format(fec, FormatoFecha) & "',"
                Else
                    sql = sql & "null,"
                End If
                
                'apellido 1
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 6))
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
                
                'apellido 2
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 7))
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
                
                'nombre
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 8))
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
                
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 9)) ' direccion
                If fec <> "" Then
                    NombreSQL fec
                    sql = sql & "'" & Trim(fec) & "',"
                Else
                    sql = sql & "null,"
                End If
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 10)) ' poblacion
                If fec <> "" Then
                    NombreSQL fec
                    sql = sql & "'" & Trim(fec) & "',"
                Else
                    sql = sql & "null,"
                End If
                fec = CInt(RecuperaValor(Cad, 11)) ' c.postal
                If fec <> "" Then
                    sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
                Else
                    sql = sql & "'46',"
                End If
                fec = RecuperaValor(Cad, 12) ' distrito
                If fec <> "" Then
                    sql = sql & "'" & Trim(Format(CInt(fec), "000")) & "',"
                Else
                    sql = sql & "null,"
                End If
                
                'tipo de trabajo
                fec = Mid(RecuperaValor(Cad, 13), 1, 2)
'                If fec = "TE" Then Stop
                sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
                
                fec = RecuperaValor(Cad, 14) 'fecha nacimiento
                If fec <> "" Then
                    sql = sql & "'" & Format(fec, FormatoFecha) & "',"
                Else
                    sql = sql & "null,"
                End If
                
                'profesion categoria
                fec = RevisaCaracterMultibase(RecuperaValor(Cad, 15))
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
                
                fec = RecuperaValor(Cad, 16) 'sexo
                sql = sql & "'" & Trim(fec) & "',"
                
                
                fec = RecuperaValor(Cad, 17) ' plantilla/contrata
                sql = sql & "'" & Trim(fec) & "',"
                
                fec = RecuperaValor(Cad, 18) 'fecha alta
                If fec <> "" Then
                    sql = sql & "'" & Format(fec, FormatoFecha) & "',"
                Else
                    sql = sql & "'" & Format("01/01/1900", FormatoFecha)
                End If
                fec = RecuperaValor(Cad, 19) 'fecha baja
                If fec <> "" Then
                    sql = sql & "'" & Format(fec, FormatoFecha) & "',"
                Else
                    sql = sql & "null,"
                End If
                
                fec = RecuperaValor(Cad, 20) 'migrado
                If fec <> "" Then
                    sql = sql & "'" & Trim(fec) & "',"
                Else
                    sql = sql & "null,"
                End If
                
                fec = ""
                fec = DevuelveDesdeBD(1, "cod_rama_gen", "tipostrab", "c_tipo_trabajo" & "|", Format(CInt(Mid(RecuperaValor(Cad, 13), 1, 2)), "00") & "|", "T|", 1)
                sql = sql & "'" & Trim(fec) & "',"
                If dni = "0" Or dni = "999999999" Or dni = "999999998" Or dni = "666666666" _
                   Or dni = "666666666" Or dni = "888888888" Or dni = "555555555" Then
                    sql = sql & "0)"
                Else
                    sql = sql & "1)"
                End If
                
                Conn.Execute sql
        
            End If
        
        
        
        Case 11 ' operarios por instalacion
            ' comprobamos que un dni esté en la tabla de operarios
            
            dni = RecuperaValor(Cad, 2)
            sql = ""
            sql = DevuelveDesdeBD(1, "dni", "operarios", "dni|", dni & "|", "T|", 1)
            If sql = "" Then   ' si no está no insertamos
                Exit Sub
            End If
            
            
            sql = "insert into operainstala values('"
        
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 1))
            If Empresa = " " Then Empresa = "DESCON"
            sql = sql & Empresa & "','"  ' empresa
            
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 21))
            If Instala = " " Then Instala = "DESCON"
            sql = sql & Instala & "','" ' instalacion
            sql = sql & Trim(dni) & "',"               ' dni del operario
            
            
            
            fec = RecuperaValor(Cad, 18) 'fecha alta
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format("01/01/1900", FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 19) 'fecha baja
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            
            fec = RecuperaValor(Cad, 20) 'migrado
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "')"
            Else
                sql = sql & "null)"
            End If
            
            
            Conn.Execute sql
       
        
        Case 12 'dosimetros de personal
            ' vemos previamente las referenciales
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 3))
            If Empresa = " " Then Empresa = "DESCON"
            
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
            If Instala = " " Then Instala = "DESCON"
            dni = RecuperaValor(Cad, 5)
'            If dni = "54040798" Then Stop
            existe = ""
            existe = DevuelveDesdeBD(1, "c_instalacion", "instalaciones", "c_empresa|c_instalacion|", Empresa & "|" & Instala & "|", "T|T|", 2)
            If existe = "" Then
                Empresa = "CREADA"
                Instala = "CREADA"
                dni = "CREADA"
            Else
                existe = ""
                existe = DevuelveDesdeBD(1, "dni", "operarios", "dni|", Trim(dni) & "|", "T|", 1)
                If existe = "" Then
                    Empresa = "CREADA"
                    Instala = "CREADA"
                    dni = "CREADA"
                Else
                    existe = ""
                    existe = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", Trim(Empresa) & "|" & Trim(Instala) & "|" & Trim(dni) & "|", "T|T|T|", 3)
                    If existe = "" Then
                        fec = RecuperaValor(Cad, 8) 'fecha asignacion dosimetro
                        If fec <> "" Then
                            fec = Format(fec, FormatoFecha)
                        Else
                            fec = Format("01/01/1900", FormatoFecha)
                        End If
                    
                        sql = "insert into operainstala values ('" & Trim(Empresa) & "','" & Trim(Instala) & "','"
                        sql = sql & Trim(dni) & "','" & Format(fec, FormatoFecha) & "',null,null)"
                        Conn.Execute sql
                    End If
                End If
            End If
                
            
'                Sql = "insert into operainstala values ('" & Trim(empresa) & "','" & Trim(Instala) & "','"
'                Sql = Sql & Trim(dni) & "','" & Format(Now, FormatoFecha) & "',null,null)"
'                Conn.Execute Sql
'            End If
            
            sql = "insert into dosimetros values ("
        
            sql = sql & RecuperaValor(Cad, 1) & ",'" ' n_reg_dosimetro
            sql = sql & RecuperaValor(Cad, 2) & "','" ' n_dosimetro
            sql = sql & Empresa & "','" ' c_empresa
            sql = sql & Instala & "','" ' c_instalacion
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 3)) & "','" ' c_empresa
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 4)) & "','" 'c_instalacion
'            Sql = Sql & RecuperaValor(Cad, 5) & "','"  'dni
            sql = sql & Trim(dni) & "','"
            sql = sql & Format(CInt(RecuperaValor(Cad, 6)), "00") & "'," 'tipo de trabajo
            sql = sql & RecuperaValor(Cad, 7) & "," ' plantilla /contrata
            
            fec = RecuperaValor(Cad, 8) 'fecha asignacion dosimetro
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format("01/01/1900", FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 9) 'fecha baja
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            
            
            'mes p i
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 10))
            NombreSQL fec
            sql = sql & "'" & Trim(fec) & "',0," ' el tipo de dosimetro por defecto es 0

            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 11)) ' observaciones
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',1,1,'H',0,null)"
            Else
                sql = sql & "null,1,1,'H',0,null)"
            End If
            
            Conn.Execute sql
        
        
        Case 13 ' dosimetros de organo
            ' vemos previamente las referenciales
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 3))
            If Empresa = " " Then Empresa = "DESCON"
            
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
            If Instala = " " Then Instala = "DESCON"
            dni = RecuperaValor(Cad, 5)
            
'            existe = ""
'            existe = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", empresa & "|" & Instala & "|" & dni & "|", "T|T|T|", 3)
'            If existe = "" Then
'                Sql = "insert into operainstala values ('" & Trim(empresa) & "','" & Trim(Instala) & "','"
'                Sql = Sql & Trim(dni) & "','" & Format(Now, FormatoFecha) & "',null,null)"
'                Conn.Execute Sql
'            End If
            
            existe = ""
            existe = DevuelveDesdeBD(1, "c_instalacion", "instalaciones", "c_empresa|c_instalacion|", Empresa & "|" & Instala & "|", "T|T|", 2)
            If existe = "" Then
                Empresa = "CREADA"
                Instala = "CREADA"
                dni = "CREADA"
            Else
                existe = ""
                existe = DevuelveDesdeBD(1, "dni", "operarios", "dni|", Trim(dni) & "|", "T|", 1)
                If existe = "" Then
                    Empresa = "CREADA"
                    Instala = "CREADA"
                    dni = "CREADA"
                Else
                    existe = ""
                    existe = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", Trim(Empresa) & "|" & Trim(Instala) & "|" & Trim(dni) & "|", "T|T|T|", 3)
                    If existe = "" Then
                        sql = "insert into operainstala values ('" & Trim(Empresa) & "','" & Trim(Instala) & "','"
                        sql = sql & Trim(dni) & "','" & Format(Now, FormatoFecha) & "',null,null)"
                        Conn.Execute sql
                    End If
                End If
            End If
            
            
            sql = "insert into dosimetros values ("
        
            sql = sql & RecuperaValor(Cad, 1) & ",'" ' n_reg_dosimetro
            sql = sql & RecuperaValor(Cad, 2) & "','" ' n_dosimetro
            sql = sql & Empresa & "','" ' c_empresa
            sql = sql & Instala & "','" 'c_instalacion
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 3)) & "','" ' c_empresa
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 4)) & "','" 'c_instalacion
            sql = sql & RecuperaValor(Cad, 5) & "','"  'dni
            sql = sql & Format(CInt(RecuperaValor(Cad, 6)), "00") & "'," 'tipo de trabajo
            sql = sql & RecuperaValor(Cad, 7) & "," ' plantilla /contrata
            
            fec = RecuperaValor(Cad, 8) 'fecha asignacion dosimetro
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format("01/01/1900", FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 9) 'fecha baja
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            
            
            'mes p i
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 10))
            NombreSQL fec
            sql = sql & "'" & Trim(fec) & "',1," ' el tipo de dosimetro por defecto es 1 dosimetro de cuerpo

            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 11)) ' observaciones
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',1,1,'H',0,null)"
            Else
                sql = sql & "null,1,1,'H',0,null)"
            End If
            
            Conn.Execute sql
        
        Case 14 ' dosis cuerpo
            sql = "insert into dosiscuerpo values ("
        
            sql = sql & RecuperaValor(Cad, 1) & ",'" ' n_reg_dosimetro
            sql = sql & RecuperaValor(Cad, 2) & "','" ' n_dosimetro
            
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 3))
            If Empresa = " " Then Empresa = "DESCON"
            
            
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
            If Instala = " " Then Instala = "DESCON"
            
            existe = ""
            existe = DevuelveDesdeBD(1, "c_empresa", "instalaciones", "c_empresa|c_instalacion|", Trim(Empresa) & "|" & Trim(Instala) & "|", "T|T|", 2)
            If existe = "" Then
                Empresa = "CREADA"
                Instala = "CREADA"
            End If
            
            sql = sql & Empresa & "','" 'c_empresa
            sql = sql & Instala & "','" 'c_instalacion
            
            sql = sql & RecuperaValor(Cad, 5) & "',"  'dni
            
            fec = RecuperaValor(Cad, 6) 'fecha dosis
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format(Now, FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 7) 'fecha migracion
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            'dosis superficial
            If RecuperaValor(Cad, 8) = "" Then
                sql = sql & "0,"
            Else
                sql = sql & TransformaComasPuntos(ImporteSinFormato(RecuperaValor(Cad, 8))) & ","
            End If
            
            'dosis profunda
            If RecuperaValor(Cad, 9) = "" Then
                sql = sql & "0,"
            Else
                sql = sql & TransformaComasPuntos(ImporteSinFormato(RecuperaValor(Cad, 9))) & ","
            End If
            
            'plantilla/contrata
            fec = RecuperaValor(Cad, 10) ' plantilla/contrata
            sql = sql & "'" & Trim(fec) & "',"
            
            fec = RecuperaValor(Cad, 11) 'rama generica
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            fec = RecuperaValor(Cad, 12) 'rama especifica
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            'c_tipotrabajo
            fec = RecuperaValor(Cad, 13)
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            'observaciones
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 14)) ' observaciones
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'migrado
            fec = RecuperaValor(Cad, 15) 'migrado
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'n-regdosimetro
            fec = ""
            fec = RecuperaValor(Cad, 16)
            If fec = "" Then fec = 0
            sql = sql & fec & ")" ' n_reg_dosimetro
            
            
            Conn.Execute sql
        
        
        Case 15 ' dosis no homogeneas
            sql = "insert into dosisnohomog values ("
        
            sql = sql & RecuperaValor(Cad, 1) & ",'" ' n_reg_dosimetro
            sql = sql & RecuperaValor(Cad, 2) & "','" ' n_dosimetro
            sql = sql & RecuperaValor(Cad, 3) & "','" ' dni
            
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 4)) & "','" 'c_empresa
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 5)) & "',"  'c_instalacion
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
            If Empresa = " " Then Empresa = "DESCON"
            
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 5))
            If Instala = " " Then Instala = "DESCON"
            
            existe = ""
            existe = DevuelveDesdeBD(1, "c_empresa", "instalaciones", "c_empresa|c_instalacion|", Trim(Empresa) & "|" & Trim(Instala) & "|", "T|T|", 2)
            If existe = "" Then
                Empresa = "CREADA"
                Instala = "CREADA"
            End If
            
            sql = sql & Empresa & "','" 'c_empresa
            sql = sql & Instala & "'," 'c_instalacion
            
            fec = RecuperaValor(Cad, 6) 'fecha dosis
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format(Now, FormatoFecha) & "',"
            End If
            fec = RecuperaValor(Cad, 7) 'fecha migracion
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            
            ' tipo de medicion
            sql = sql & "'" & Format(CInt(Trim(RecuperaValor(Cad, 8))), "00") & "',"
            
            'dosis organo
            If RecuperaValor(Cad, 9) = "" Then
                sql = sql & "0,"
            Else
                sql = sql & TransformaComasPuntos(ImporteSinFormato(RecuperaValor(Cad, 9))) & ","
            End If
            
            'plantilla/contrata
            fec = RecuperaValor(Cad, 10) ' plantilla/contrata
            sql = sql & "'" & Trim(fec) & "',"
            
            fec = RecuperaValor(Cad, 11) 'rama generica
            sql = sql & "'" & Format(CInt(Trim(fec)), "00") & "',"
            
            fec = RecuperaValor(Cad, 12) 'rama especifica
            sql = sql & "'" & Format(CInt(Trim(fec)), "00") & "',"
            
            'c_tipotrabajo
            fec = RecuperaValor(Cad, 13)
            sql = sql & "'" & Format(CInt(Trim(fec)), "00") & "',"
            
            'observaciones
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 14)) ' observaciones
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'migrado
            fec = RecuperaValor(Cad, 15) 'migrado
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'n-regdosimetro
            fec = ""
            fec = RecuperaValor(Cad, 16)
            If fec = "" Then fec = 0
            sql = sql & fec & ")" ' n_reg_dosimetro
            
            
            Conn.Execute sql
        

        Case 18 ' operario 777777777 y operarios por instalacion en operainstala de dosimetria de area
            sql = ""
            sql = DevuelveDesdeBD(1, "dni", "operarios", "dni|", "777777777|", "T|", 1)
            If sql = "" Then
                sql = "insert into operarios values ('777777777','FICTICIO','FICTICIO','1900-01-01',"
                sql = sql & "'FICTICIO','FICTICIO','FICTICIO','FICTICIO','FICTICIO','46','000','99','1900-01-01',"
                sql = sql & "'FICTICIO','V','01','1900-01-01',null,null,'99',0)"
            
                Conn.Execute sql
            End If
            
            sql = "select * from instalaciones where c_tipo = 1 or c_tipo = 2"
            
            Set Rs = New ADODB.Recordset
            
            Rs.Open sql, Conn, , , adCmdText
            While Not Rs.EOF
                existe = ""
                existe = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", Rs!c_empresa & "|" & Rs!c_instalacion & "|" & "777777777|", "T|T|T|", 3)
                If existe = "" Then
                    sql1 = "insert into operainstala values ('" & Trim(Rs!c_empresa) & "','" & Trim(Rs!c_instalacion) & "',"
                    sql1 = sql1 & "'777777777','" & Format(Now, FormatoFecha) & "',null,null)"
                    
                    Conn.Execute sql1
                End If
                    
                Rs.MoveNext
            Wend
            Rs.Close
        
        Case 19 ' dosimetros de area
            ' vemos previamente las referenciales
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 3))
            If Empresa = " " Then Empresa = "DESCON"
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
            If Instala = " " Then Instala = "DESCON"
            dni = Trim(RecuperaValor(Cad, 5))

'            existe = ""
'            existe = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", empresa & "|" & Instala & "|" & dni & "|", "T|T|T|", 3)
'            If existe = "" Then
'                Sql = "insert into operainstala values ('" & Trim(empresa) & "','" & Trim(Instala) & "','"
'                Sql = Sql & Trim(dni) & "','" & Format(Now, FormatoFecha) & "',null,null)"
'                Conn.Execute Sql
'            End If
            
            existe = ""
            existe = DevuelveDesdeBD(1, "c_instalacion", "instalaciones", "c_empresa|c_instalacion|", Empresa & "|" & Instala & "|", "T|T|", 2)
            If existe = "" Then
                Empresa = "CREADA"
                Instala = "CREADA"
                dni = "CREADA"
            Else
                existe = ""
                existe = DevuelveDesdeBD(1, "dni", "operarios", "dni|", Trim(dni) & "|", "T|", 1)
                If existe = "" Then
                    Empresa = "CREADA"
                    Instala = "CREADA"
                    dni = "CREADA"
                Else
                    existe = ""
                    existe = DevuelveDesdeBD(1, "dni", "operainstala", "c_empresa|c_instalacion|dni|", Trim(Empresa) & "|" & Trim(Instala) & "|" & Trim(dni) & "|", "T|T|T|", 3)
                    If existe = "" Then
                        sql = "insert into operainstala values ('" & Trim(Empresa) & "','" & Trim(Instala) & "','"
                        sql = sql & Trim(dni) & "','" & Format(Now, FormatoFecha) & "',null,null)"
                        Conn.Execute sql
                    End If
                End If
            End If
            
            
            
            sql = "insert into dosimetros values ("
        
            sql = sql & RecuperaValor(Cad, 1) & ",'" ' n_reg_dosimetro
            sql = sql & RecuperaValor(Cad, 2) & "','" ' n_dosimetro
            sql = sql & Empresa & "','" ' c_empresa
            sql = sql & Instala & "','" ' c_instalacion
            sql = sql & dni & "','" ' DNI
            
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 3)) & "','" ' c_empresa
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 4)) & "','" 'c_instalacion
'            Sql = Sql & RecuperaValor(Cad, 5) & "','"  'dni
            sql = sql & Format(CInt(RecuperaValor(Cad, 6)), "00") & "'," 'tipo de trabajo
            sql = sql & RecuperaValor(Cad, 7) & "," ' plantilla /contrata
            
            fec = RecuperaValor(Cad, 8) 'fecha asignacion dosimetro
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format("01/01/1900", FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 9) 'fecha baja
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            
            
            'mes p i
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 10))
            NombreSQL fec
            sql = sql & "'" & Trim(fec) & "',2," ' el tipo de dosimetro es 2 porque es de area

            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 11)) ' observaciones
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',1,1,'H',0,null)"
            Else
                sql = sql & "null,1,1,'H',0,null)"
            End If
            
            Conn.Execute sql

        Case 20 ' dosis area
            sql = "insert into dosisarea values ("
        
            sql = sql & RecuperaValor(Cad, 1) & ",'" ' n_reg_dosimetro
            sql = sql & RecuperaValor(Cad, 2) & "','" ' n_dosimetro
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 3))
            If Empresa = " " Then Empresa = "DESCON"
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
            If Instala = " " Then Instala = "DESCON"
            
            existe = ""
            existe = DevuelveDesdeBD(1, "c_empresa", "instalaciones", "c_empresa|c_instalacion|", Trim(Empresa) & "|" & Trim(Instala) & "|", "T|T|", 2)
            If existe = "" Then
                Empresa = "CREADA"
                Instala = "CREADA"
            End If
            
            sql = sql & Empresa & "','" ' c_empresa
            sql = sql & Instala & "','" 'c_instalacion

'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 3)) & "','" ' c_empresa
'            sql = sql & RevisaCaracterMultibase(RecuperaValor(Cad, 4)) & "','" 'c_instalacion
            sql = sql & RecuperaValor(Cad, 5) & "',"  'dni
            
            fec = RecuperaValor(Cad, 6) 'fecha dosis
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format(Now, FormatoFecha)
            End If
            fec = RecuperaValor(Cad, 7) 'fecha migracion
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'dosis superficial
            If RecuperaValor(Cad, 8) = "" Then
                sql = sql & "0,"
            Else
                sql = sql & TransformaComasPuntos(ImporteSinFormato(RecuperaValor(Cad, 8))) & ","
            End If
            
            'dosis profunda
            If RecuperaValor(Cad, 9) = "" Then
                sql = sql & "0,"
            Else
                sql = sql & TransformaComasPuntos(ImporteSinFormato(RecuperaValor(Cad, 9))) & ","
            End If
            
            'plantilla/contrata
            fec = RecuperaValor(Cad, 10) ' plantilla/contrata
            sql = sql & "'" & Trim(fec) & "',"
            
            fec = RecuperaValor(Cad, 11) 'rama generica
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            fec = RecuperaValor(Cad, 12) 'rama especifica
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            'c_tipotrabajo
            fec = RecuperaValor(Cad, 13)
            sql = sql & "'" & Trim(Format(CInt(fec), "00")) & "',"
            
            'observaciones
            fec = RevisaCaracterMultibase(RecuperaValor(Cad, 14)) ' observaciones
            If fec <> "" Then
                NombreSQL fec
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'migrado
            fec = RecuperaValor(Cad, 15) 'migrado
            If fec <> "" Then
                sql = sql & "'" & Trim(fec) & "',"
            Else
                sql = sql & "null,"
            End If
            
            'n-regdosimetro
            fec = ""
            fec = RecuperaValor(Cad, 16)
            If fec = "" Then fec = 0
            sql = sql & fec & ")" ' n_reg_dosimetro
            
            
            Conn.Execute sql
        
        Case 21, 22 'recpcion de dosimetros
            '21 cuerpo
            '22 area
            sql = "insert into recepdosim values ("
        
            sql = sql & RecuperaValor(Cad, 1) & ",'" ' n_reg_dosimetro
            sql = sql & RecuperaValor(Cad, 2) & "','" ' n_dosimetro
            sql = sql & RecuperaValor(Cad, 3) & "','" ' dni_usuario
            Empresa = RevisaCaracterMultibase(RecuperaValor(Cad, 4))
            If Empresa = " " Then Empresa = "DESCON"
            Instala = RevisaCaracterMultibase(RecuperaValor(Cad, 5))
            If Instala = " " Then Instala = "DESCON"
            sql = sql & Empresa & "','" ' c_empresa
            sql = sql & Instala & "'," 'c_instalacion

            fec = RecuperaValor(Cad, 6) 'fecha recepcion
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "null,"
            End If
            fec = RecuperaValor(Cad, 7) 'fecha de creacion
            If fec <> "" Then
                sql = sql & "'" & Format(fec, FormatoFecha) & "',"
            Else
                sql = sql & "'" & Format(Now, FormatoFecha) & "',"
            End If
            
            'mes p/i
            sql = sql & "'" & RecuperaValor(Cad, 8) & "',"
            
            If Combo1.ListIndex = 21 Then
                sql = sql & "0,'H')"
            Else
                sql = sql & "2,'H')"
            End If
            
            Conn.Execute sql
        
            
    
    End Select
    IncPb2
'eInsertarRegistro:
'    If Err.Number <> 0 Then
'        MsgBox "Error insertando en " & Combo1.Text & " regsitro: " & Cad & Err.Description
'    End If

End Sub

Private Sub IncPb2()
On Error Resume Next
pb2.Value = pb2.Value + 1
If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub CargarCombo()
    Combo1.Clear
    Combo1.AddItem "Provincias"
    Combo1.ItemData(Combo1.NewIndex) = 0

    Combo1.AddItem "Ramas Genéricas"
    Combo1.ItemData(Combo1.NewIndex) = 1

    Combo1.AddItem "Ramas Específicas"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
    Combo1.AddItem "Tipos Trabajo"
    Combo1.ItemData(Combo1.NewIndex) = 3
    
    Combo1.AddItem "Tipos Med.Extremidades"
    Combo1.ItemData(Combo1.NewIndex) = 4
    
    Combo1.AddItem "Fondos"
    Combo1.ItemData(Combo1.NewIndex) = 5
    
    Combo1.AddItem "Fact.Cali 4400"
    Combo1.ItemData(Combo1.NewIndex) = 6
    
    Combo1.AddItem "Fact.Cali 6600"
    Combo1.ItemData(Combo1.NewIndex) = 7
    
    Combo1.AddItem "Empresas"
    Combo1.ItemData(Combo1.NewIndex) = 8
    
    Combo1.AddItem "Instalaciones"
    Combo1.ItemData(Combo1.NewIndex) = 9
    
    Combo1.AddItem "Operarios"
    Combo1.ItemData(Combo1.NewIndex) = 10
    
    Combo1.AddItem "Operarios/instalacion"
    Combo1.ItemData(Combo1.NewIndex) = 11
    
    Combo1.AddItem "Dosimetros"
    Combo1.ItemData(Combo1.NewIndex) = 12
    
    Combo1.AddItem "Dosimetros Organos"
    Combo1.ItemData(Combo1.NewIndex) = 13
    
    Combo1.AddItem "Dosis Homogeneas"
    Combo1.ItemData(Combo1.NewIndex) = 14
    
    Combo1.AddItem "Dosis Organo"
    Combo1.ItemData(Combo1.NewIndex) = 15
    
    
    Combo1.AddItem "Empresas area"
    Combo1.ItemData(Combo1.NewIndex) = 16
    
    
    Combo1.AddItem "Instalaciones Area"
    Combo1.ItemData(Combo1.NewIndex) = 17
    
    
    Combo1.AddItem "Operarios Area"
    Combo1.ItemData(Combo1.NewIndex) = 18
    
    
    Combo1.AddItem "Dosimetros Area"
    Combo1.ItemData(Combo1.NewIndex) = 19
    
    
    Combo1.AddItem "Dosis Area"
    Combo1.ItemData(Combo1.NewIndex) = 20

    Combo1.AddItem "Recepcion dosim cuerpo"
    Combo1.ItemData(Combo1.NewIndex) = 21

    Combo1.AddItem "Recepcion dosim area"
    Combo1.ItemData(Combo1.NewIndex) = 22

End Sub

