VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmPenalizacionDosis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalización Automática de Dosímetros sin dosis"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmPenalizacionDosis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListTipoMedicion 
      Height          =   5940
      Left            =   30
      TabIndex        =   7
      Top             =   60
      Width           =   7275
      Begin VB.TextBox Text1 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   5820
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   2685
         Width           =   555
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1815
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3510
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   0
         Tag             =   "JMCE"
         Top             =   1935
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dosis de Penalización "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   420
         TabIndex        =   10
         Top             =   3240
         Width           =   6195
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   4410
            TabIndex        =   4
            Text            =   "Text5"
            Top             =   420
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1380
            TabIndex        =   3
            Text            =   "Text5"
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Profunda:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   16
            Top             =   450
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Superficial:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   15
            Top             =   450
            Width           =   1155
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3720
         TabIndex        =   6
         Top             =   4920
         Width           =   1425
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   2145
         TabIndex        =   5
         Top             =   4920
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   300
         Left            =   360
         TabIndex        =   9
         Top             =   4380
         Visible         =   0   'False
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Máximo de Meses sin Dosis:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3510
         TabIndex        =   17
         Top             =   2715
         Width           =   2265
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Desde:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   570
         TabIndex        =   14
         Top             =   2730
         Width           =   1005
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmPenalizacionDosis.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmPenalizacionDosis.frx":0E1C
         ToolTipText     =   "Seleccionar fecha"
         Top             =   2700
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Se asigna una dosis de penalización del valor que se indique abajo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   540
         TabIndex        =   13
         Top             =   930
         Width           =   5805
      End
      Begin VB.Label Label4 
         Caption         =   "Introduzca la clave de acceso y pulse intro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1590
         TabIndex        =   12
         Top             =   1590
         Width           =   3825
      End
      Begin VB.Label Label2 
         Caption         =   "CLAVE DE ACCESO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   1620
         TabIndex        =   11
         Top             =   1950
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Este programa genera una dosis homogénea de penalización por cada dosímetro no recibido en el plazo de 6 meses."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   570
         TabIndex        =   8
         Top             =   330
         Width           =   5805
      End
   End
End
Attribute VB_Name = "frmPenalizacionDosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim dospactual As Currency
Dim dossactual As Currency
Dim mesespenal As Integer

Dim sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean
Dim fec As Date
Dim sql1 As String
Dim ContSubgrup As Integer
Dim fecdosis As String
Dim Cont As Integer
Dim ModoLocal As Integer

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    
    Unload Me
    
End Sub


Private Sub cmdAceptar_Click()
Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim Tipo As String
Dim nRegs As Long
On Error GoTo eErrorCarga

  Screen.MousePointer = vbHourglass
  If DatosOk Then
    nRegs = 0
    nRegs = CargarPbarMax

    If nRegs = 0 Then
      MsgBox "No hay datos entre estos límites. Reintroduzca.", vbExclamation, "¡Error!"
      Screen.MousePointer = vbDefault
      Exit Sub
    Else
      
      FrameListTipoMedicion.Enabled = False
      
      Pb1.Visible = True
      Pb1.max = nRegs + 1
      Pb1.Value = 0
         
      CargaTablaIntermedia
         
      Pb1.Visible = False
      Pb1.Value = 0
            
      ' ### [DavidV] 14/04/2006: Mostramos el formulario de selección de penalizaciones.
      frmSelecPenalizaciones.Show vbModal
      If CadenaDevueltaFormHijo <> "Ningún dosímetro a penalizar." Then
        ' ### [DavidV] 14/04/2006: ¿Desea Imprimir?
        If MsgBox("¿Desea imprimir los dosimetros que van a ser penalizados?", vbQuestion + vbYesNo + vbDefaultButton1, "¡Atención!") = vbYes Then
          frmImprimir.Opcion = 32
          frmImprimir.Show vbModal
        End If
        
        ' ### [DavidV] 05/04/2006: Confirmar la penalización.
        If MsgBox("Por favor, confirme el proceso de penalización.", vbQuestion + vbYesNo + vbDefaultButton1, "¡Atención!") = vbYes Then
          ActualizarRegistros
        End If
      End If
    End If
  End If
      
  FrameListTipoMedicion.Enabled = True
  ActivarCLAVE
  Screen.MousePointer = vbDefault
  Exit Sub

eErrorCarga:

    FrameListTipoMedicion.Enabled = True
    ActivarCLAVE
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Error en la generación del fichero de dosis a CSN. Revise."

End Sub

Private Sub Form_Load()
Dim ano As Currency
Dim Mes As Currency
Dim config As CConfiguracion

    ActivarCLAVE
    
    ano = Year(Now)
    Mes = Month(Now)
    Mes = Mes - 5
    If Mes <= 0 Then
        Mes = Mes + 12
        ano = ano - 1
    End If
    fecdosis = Format(Now, "dd/mm/yyyy")
    fec = CDate("01/" & Format(Mes, "00") & "/" & Format(ano, "0000"))
    Text1(1).Text = Format(DateAdd("d", -1, fec), "dd/mm/yyyy")
    
    Set config = New CConfiguracion
    config.clave = "meses_penalizacion"
    config.Leer
    mesespenal = Val(config.valor)
    config.clave = "dosis_superficial"
    config.Leer
    dossactual = Val(config.valor)
    config.clave = "dosis_profunda"
    config.Leer
    dospactual = Val(config.valor)
    Set config = Nothing
    
    Text1(4).Text = mesespenal 'meses
    Text1(0).Text = Format(dossactual, "0.00") 'superficial
    Text1(2).Text = Format(dospactual, "0.00") 'profunda
        
    ModoLocal = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim config As CConfiguracion
    
  If Not FrameListTipoMedicion.Enabled Then
    Cancel = vbCancel
    Exit Sub
  End If
  
  If IsNumeric(Text1(4).Text) And IsNumeric(Text1(0).Text) And IsNumeric(Text1(2).Text) Then
    BloqueoManual False, "TRASPASO", "TRASPASO"
      If CInt(Text1(4).Text) <> CInt(mesespenal) Or Val(Text1(0).Text) <> Val(dossactual) Or Val(Text1(2).Text) <> Val(dospactual) Then
        If MsgBox("Ha cambiado la configuración de penalizaciones. ¿Está seguro/a?", vbYesNo + vbQuestion, "¡Atención!") = vbYes Then
          Set config = New CConfiguracion
          With config
            .clave = "meses_penalizacion"
            .valor = CInt(Text1(4).Text)
            .Modificar
            .clave = "dosis_superficial"
            .valor = Val(TransformaComasPuntos(Text1(0).Text))
            .Modificar
            .clave = "dosis_profunda"
            .valor = Val(TransformaComasPuntos(Text1(2).Text))
            .Modificar
          End With
          Set config = Nothing
        End If
      End If
  End If
      
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    Select Case Index
       Case 1
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(1).Text <> "" Then
                If IsDate(Text1(1).Text) Then f = Text1(1).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            Text1(1).Text = frmC.fecha
            mTag.DarFormato Text1(1)
            Set frmC = Nothing
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
   Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim I As Integer
    Dim sql As String
    Dim mTag As CTag
    Dim valor As Currency
    Dim nomFich As String
    
    ''Quitamos blancos por los lados
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    'If Text1(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 3
            ' No dejamos introducir comillas en ningun campo tipo texto
            'If InStr(1, Text1(Index).Text, "'") > 0 Then
            '    MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
            '    Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
            '    PonerFoco Text1(Index)
            '    Exit Sub
            'End If
            'Text1(Index).Text = Format(Text1(Index).Text, ">")
            
            If Text1(3).Text = "" Then Exit Sub
            
            If Trim(Text1(3).Text) <> Trim(Text1(3).Tag) Then
                MsgBox "    Acceso denegado    ", vbExclamation, "¡Atención!"
                Text1(3).Text = ""
                ActivarCLAVE
                PonerFoco Text1(3)
            Else
                DesactivarCLAVE
                PonerFoco Text1(1)
            End If
            
        Case 4
            ' comprobamos cosas
            If Text1(Index).Text <> "" Then
                If EsNumerico(Text1(Index).Text) Then
                    If InStr(1, Text1(Index).Text, ",") > 0 Or InStr(1, Text1(Index).Text, ".") > 0 Then
                       MsgBox "El valor ha de ser un entero.", vbOKOnly + vbExclamation, "¡Error!"
                       PonerFoco Text1(Index)
                    End If
                    If InStr(1, Text1(Index).Text, "-") > 0 Then
                       MsgBox "El valor no puede ser negativo.", vbOKOnly + vbExclamation, "¡Error!"
                       PonerFoco Text1(Index)
                    End If
                Else
                  PonerFoco Text1(Index)
                End If
                
            End If
     
        Case 0, 2
            ' comprobamos cosas
            If Text1(Index).Text <> "" Then
                If EsNumerico(Text1(Index).Text) Then
                    If InStr(1, Text1(Index).Text, ",") > 0 Then
                        valor = ImporteFormateado(Text1(Index).Text)
                    Else
                        valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                    End If
                    
                    Text1(Index).Text = Format(valor, "##0.00")
                Else
                  PonerFoco Text1(Index)
                End If
            End If
          
        Case 1
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation, "¡Error!"
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
              

            End If
              
    End Select
    
    '---
End Sub

Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text1(Indice1).Text <> "" And Text1(Indice2).Text <> "" Then
        If CDate(Text1(Indice1).Text) > CDate(Text1(Indice2).Text) Then
            MsgBox "Fecha 'desde' mayor que fecha 'hasta'.", vbExclamation, "¡Error!"
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function

Private Sub ActivarCLAVE()
Dim I As Integer
    
    For I = 0 To Text1.Count - 1
        Text1(I).Enabled = False
    Next I

    Imgppal(1).Enabled = False

    Text1(3).Enabled = True
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = True

End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer

    For I = 0 To Text1.Count - 1
        Text1(I).Enabled = True
    Next I

    Imgppal(1).Enabled = True

    Text1(3).Text = ""
    
    cmdAceptar.Enabled = True
End Sub

' Carga la tabla intermedia de dosímetros a penalizar.
Private Function CargaTablaIntermedia() As Integer
Dim sql As String
Dim parimpar As String
Dim fecha As String
Dim mesespenal As Integer
Dim fechapenal As String
Dim dosiprof As Currency
Dim dosissup As Currency
Dim Rs As ADODB.Recordset
Dim porcentaje As Currency
Dim nreg As Long

  ' Determinamos la fecha a partir de la cual se penaliza, entre otras cosas.
  fecha = Format(Text1(1).Text, "yyyy-MM-01")
  mesespenal = (CInt(Text1(4).Text) * 2) - 2
  fechapenal = Format(DateAdd("M", -mesespenal, fecha), "yyyy-MM-01")
  parimpar = IIf((Month(fecha) And 1) = 1, "I", "P")
  dosissup = Round2(CCur(Text1(0).Text), 2)
  dosiprof = Round2(CCur(Text1(2).Text), 2)
  porcentaje = CCur(100 / Pb1.max)
  Pb1.max = 100
  
  ' Carga de las tablas para el informe.
  CargarInstalaciones
  CargarOperarios
  
  ' Inicializamos la tabla para este usuario.
  sql = "delete from zdosisacum where codusu = " & vUsu.codigo
  Conn.Execute sql

  ' Seleccionamos aquellos dosímetros susceptibles de ser penalizados: cruza la tabla de
  ' dosímetros con la de dosis de cuerpo cuya fecha de dosis esté dentro del periodo correcto.
  ' Aquellos que no estén en ese periodo, serán NULL debido al LEFT JOIN. Después sólo hay
  ' que seleccionar los que son NULL.
  sql = "SELECT T1.n_reg_dosimetro Reg1, T1.dni_usuario, T1.c_empresa, T1.c_instalacion,"
  sql = sql & "T1.n_dosimetro, T2.n_reg_dosimetro Reg2 FROM (" ' Primer elemento del LEFT JOIN (T1).
  sql = sql & "SELECT d.* "
  sql = sql & "FROM dosimetros d, operarios o WHERE d.n_dosimetro NOT LIKE 'VIRTUAL%' AND "
  sql = sql & "d.mes_p_i = '" & parimpar & "' AND d.f_retirada IS NULL AND "
  sql = sql & "d.f_asig_dosimetro < '" & fechapenal & "' AND d.tipo_dosimetro = 0 AND "
  sql = sql & "o.semigracsn = 1 AND o.f_baja IS NULL AND d.dni_usuario = o.dni and "
  sql = sql & "o.dni <> '888888888' and o.dni <> '999999996' and o.dni <> '0') T1 "
  sql = sql & "LEFT JOIN (" ' Segundo elemento del LEFT JOIN (T2).
  sql = sql & "SELECT DISTINCT(n_reg_dosimetro) FROM dosiscuerpo WHERE f_dosis >= '"
  sql = sql & fechapenal & "') T2 "
  sql = sql & "ON T1.n_reg_dosimetro = T2.n_reg_dosimetro "
  sql = sql & "WHERE T2.n_reg_dosimetro IS NULL"
  

  Set Rs = New ADODB.Recordset
  Rs.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
  nreg = 0
  While Not Rs.EOF
    
    ' Insertamos en la tabla temporal las dosis de penalización correspondientes.
    sql = "insert into zdosisacum (codusu, c_empresa, c_instalacion, dni_usuario, "
    sql = sql & "n_dosimetro, dosissuper, dosisprofu, n_reg_dosimetro) values (" & vUsu.codigo & ","
    sql = sql & "'" & Trim(Rs!c_empresa) & "','" & Trim(Rs!c_instalacion) & "',"
    sql = sql & "'" & Trim(Rs!dni_usuario) & "','" & Trim(Rs!n_dosimetro) & "',"
    sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(dosissup))) & ","
    sql = sql & TransformaComasPuntos(ImporteSinFormato(CStr(dosiprof))) & ","
    sql = sql & Rs!Reg1 & ")"

    Conn.Execute sql
      
    ' Barra de progreso.
    nreg = nreg + 1
    If CInt(nreg * porcentaje) < Pb1.max Then
      Pb1.Value = CInt(nreg * porcentaje)
    Else
      Debug.Print nreg
    End If
    DoEvents
    
    Rs.MoveNext

  Wend

  Rs.Close
  Set Rs = Nothing
  Pb1.Value = 100
  
End Function

Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaABlancos = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaABlancos = Right(Cad, Longitud)
    End If
    
End Function

Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaAceros = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaAceros = Right(Cad, Longitud)
    End If
    
End Function

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "Select Max(n_registro) from dosiscuerpo"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    sql = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            sql = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = sql
End Function

Private Function DatosOk() As Boolean

    DatosOk = True
    If Text1(1).Text = "" Then
        MsgBox "Debe introducir un valor en el campo fecha para poder hacer los cálculos. Reintroduzca.", vbExclamation, , "¡Error!"
        DatosOk = False
        Exit Function
    End If

End Function

Private Sub ActualizarRegistros()

Dim Rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim nreg As Currency
Dim rama_gen As String
Dim rama_esp As String
Dim tipo_tra As String
Dim plan_con As String
Dim f_migracion As String
Dim fecha As String

On Error GoTo eActualizarRegistros
    
    Conn.BeginTrans
    
    
    fecha = Format(DateAdd("d", -1, Format(DateAdd("M", 1, Text1(1).Text), "yyyy-MM-01")), "yyyy-MM-dd")
    f_migracion = Format(DateAdd("M", 1, fecha), "yyyy-MM-dd")

    sql = "select * from zdosisacum where codusu = " & vUsu.codigo
    
    Set Rs = New ADODB.Recordset
    
    Rs.Open sql, Conn, , , adCmdText
    If Not Rs.EOF Then Rs.MoveFirst
    ' Cambios, más cambios...
    
    
    While Not Rs.EOF
        
        Pb1.Value = Pb1.Value + 1
        Pb1.Refresh
        
        nreg = SugerirCodigoSiguiente
    
        rama_gen = ""
        rama_esp = ""
        tipo_tra = ""
        plan_con = ""
        rama_gen = DevuelveDesdeBD(1, "rama_gen", "instalaciones", "c_empresa|c_instalacion|", Trim(Rs!c_empresa) & "|" & Trim(Rs!c_instalacion) & "|", "T|T|", 2)
        rama_esp = DevuelveDesdeBD(1, "rama_especifica", "instalaciones", "c_empresa|c_instalacion|", Trim(Rs!c_empresa) & "|" & Trim(Rs!c_instalacion) & "|", "T|T|", 2)
        tipo_tra = DevuelveDesdeBD(1, "c_tipo_trabajo", "dosimetros", "n_reg_dosimetro|", Rs!n_reg_dosimetro & "|", "N|", 1)
        plan_con = DevuelveDesdeBD(1, "plantilla_contrata", "dosimetros", "n_reg_dosimetro|", Rs!n_reg_dosimetro & "|", "N|", 1)
    
        sql1 = "insert into dosiscuerpo (n_registro, n_dosimetro, c_empresa, c_instalacion, dni_usuario, "
        sql1 = sql1 & "f_dosis, f_migracion, dosis_superf, dosis_profunda, plantilla_contrata, "
        sql1 = sql1 & "rama_generica, rama_especifica, c_tipo_trabajo, observaciones, migrado, n_reg_dosimetro) "
        sql1 = sql1 & " values (" & nreg & ",'" & Trim(Rs!n_dosimetro) & "','" & Trim(Rs!c_empresa) & "','"
        sql1 = sql1 & Trim(Rs!c_instalacion) & "','" & Trim(Rs!dni_usuario) & "','" & fecha & "',"
        sql1 = sql1 & "'" & Trim(f_migracion) & "', " 'fecha de migracion
        sql1 = sql1 & TransformaComasPuntos(ImporteSinFormato(CStr(Rs!dosissuper))) & ","
        sql1 = sql1 & TransformaComasPuntos(ImporteSinFormato(CStr(Rs!dosisprofu))) & ","
        sql1 = sql1 & "'" & Format(CInt(plan_con), "00") & "'," ' plantilla contrata
        sql1 = sql1 & "'" & Trim(rama_gen) & "'," ' rama generica
        sql1 = sql1 & "'" & Trim(rama_esp) & "'," ' rama especifica
        sql1 = sql1 & "'" & Trim(tipo_tra) & "'," ' tipo de trabajo
        sql1 = sql1 & "'ASIGNACION DOSIS POR NO RECAMBIO DEL DOSIMETRO + DE " & Text1(4).Text & " MESES'," ' observaciones
        sql1 = sql1 & "null," ' migrado
        sql1 = sql1 & Rs!n_reg_dosimetro & ")" ' me falta el numero de registro de dosimetro
        
        Conn.Execute sql1
        
        ' Ahora no debe de cambiarse la fecha de recepción como control de
        ' penalización. Estoy impaciente por saber el siguiente "cambio".
'        sql1 = "update recepdosim set fecha_recepcion = '" & Format(fecha, FormatoFecha) & "' "
'        sql1 = sql1 & " where n_reg_dosimetro = " & Rs!n_reg_dosimetro & " and tipo_dosimetro = 0 "
'        sql1 = sql1 & " and fecha_recepcion is null "
'        sql1 = sql1 & " and f_creacion_recep >= '" & Format(f_migracion, FormatoFecha) & "'"
'
'        Conn.Execute sql1
        
        Rs.MoveNext
    Wend
    
eActualizarRegistros:
    If Err.Number <> 0 Then
        Conn.RollbackTrans
        MuestraError Err.Number, "Error en la actualizacion de registros"
    Else
        Conn.CommitTrans
        MsgBox "Proceso realizado con éxito.", vbInformation, "Penalización de dosímetros."
    End If

End Sub

Private Function CargarPbarMax() As Long
Dim Rs As ADODB.Recordset
Dim sql As String
Dim parimpar As String
Dim fecha As String
Dim fechapenal As String

  Set Rs = New ADODB.Recordset

  ' Determinamos la fecha a partir de la cual se penaliza, entre otras cosas.
  fecha = Format(Text1(1).Text, "yyyy-MM-01")
  fechapenal = Format(DateAdd("M", -CInt(Text1(4).Text), fecha), "yyyy-MM-01")
  parimpar = IIf((Month(fecha) And 1) = 1, "I", "P")
  
  sql = "SELECT COUNT(*) FROM (" ' Primer elemento del LEFT JOIN (T1).
  sql = sql & "SELECT d.* "
  sql = sql & "FROM dosimetros d, operarios o WHERE d.n_dosimetro NOT LIKE 'VIRTUAL%' AND "
  sql = sql & "d.mes_p_i = '" & parimpar & "' AND d.f_retirada IS NULL AND "
  sql = sql & "d.f_asig_dosimetro < '" & fecha & "' AND d.tipo_dosimetro = 0 AND "
  sql = sql & "o.semigracsn = 1 AND o.f_baja IS NULL AND d.dni_usuario = o.dni ) T1 "
  sql = sql & "LEFT JOIN (" ' Segundo elemento del LEFT JOIN (T2).
  sql = sql & "SELECT DISTINCT(n_reg_dosimetro) FROM dosiscuerpo WHERE f_dosis >= '"
  sql = sql & fechapenal & "') T2 "
  sql = sql & "ON T1.n_reg_dosimetro = T2.n_reg_dosimetro "
  sql = sql & "WHERE T2.n_reg_dosimetro IS NULL"
  
  Rs.Open sql, Conn, , , adCmdText
  If Not Rs.EOF Then CargarPbarMax = Rs.Fields(0).Value Else CargarPbarMax = 0
  Set Rs = Nothing

End Function
