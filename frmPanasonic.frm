VERSION 5.00
Begin VB.Form frmPanasonic 
   Caption         =   "Carga Automática de Dosis Panasonic"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   Icon            =   "frmPanasonic.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4845
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4050
      TabIndex        =   4
      Top             =   4260
      Width           =   855
   End
   Begin VB.CommandButton BtnCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   3210
      TabIndex        =   3
      Top             =   4260
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   2730
      Pattern         =   "*.asc"
      TabIndex        =   2
      Top             =   780
      Width           =   2040
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   450
      TabIndex        =   1
      Top             =   780
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   300
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "N. Registro"
      Height          =   255
      Index           =   8
      Left            =   465
      TabIndex        =   21
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   1665
      TabIndex        =   20
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   3810
      TabIndex        =   19
      Top             =   3660
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   3810
      TabIndex        =   18
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   3810
      TabIndex        =   17
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   3810
      TabIndex        =   16
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1650
      TabIndex        =   15
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1650
      TabIndex        =   14
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   1650
      TabIndex        =   13
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  4"
      Height          =   255
      Index           =   7
      Left            =   3090
      TabIndex        =   12
      Top             =   3660
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  3"
      Height          =   255
      Index           =   6
      Left            =   3090
      TabIndex        =   11
      Top             =   3420
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  2"
      Height          =   255
      Index           =   5
      Left            =   3090
      TabIndex        =   10
      Top             =   3180
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  1"
      Height          =   255
      Index           =   4
      Left            =   3090
      TabIndex        =   9
      Top             =   2940
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Dosimetro"
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   8
      Top             =   3420
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Hora..."
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   7
      Top             =   3180
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha..."
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   6
      Top             =   2940
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Archivo del Panasonic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2730
      TabIndex        =   5
      Top             =   420
      Width           =   2040
   End
End
Attribute VB_Name = "frmPanasonic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Valor_Minimo_Luz_C4 As Double
Public Valor_Maximo_Luz_C4 As Double
Public Valor_Minimo_Luz_C2 As Double
Public Valor_Maximo_Luz_C2 As Double
Public Valor_Minimo_Ruido_C4 As Double
Public Valor_Minimo_Ruido_C2 As Double
Public Valor_Maximo_Ruido_C4 As Double
Public Valor_Maximo_Ruido_C2 As Double
Public Descrip_Dosimetro As String
Public N_Linea As Integer


Dim Caracter As String

Dim sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Private Sub Form_Load()
    
  
   'antes esto
   ' Dir1.Path = "c:\MBgstld4\Migra\"
   Dir1.Path = App.Path & "\migra\"
   'MsgBox Dir1.Path, vbExclamation
   
    File1.Path = Dir1.Path
    File1.Pattern = "*.DAT"
    '
    'Ponemos Valores Predeterminados
    '
    Valor_Minimo_Luz_C4 = 999999
    Valor_Minimo_Luz_C2 = 999999
    Valor_Minimo_Ruido_C4 = 9999
    Valor_Minimo_Ruido_C2 = 9999
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    'Shell "notepad.exe 'c:\mbgstld4\migra\informe.txt'", vbMaximizedFocus
    Directorio = App.Path & "\migra\"
    
    ' ### [DavidV] ¿Esto pa qué?
    'If Dir(Directorio & "\informe.txt", vbArchive) = "Informe.txt" Then
    '    Shell "notepad.exe '" & Directorio & "\informe.txt'", vbMaximizedFocus
    'End If
    
    'formulario de la tabla intermedia
    
    
    Unload Me
End Sub

' ### [DavidV] Al parecer esto está obsoleto.
' hemos de ejecutar esta sintaxis
'
'    load data infile 'c:/programas/datos/sprovi.dat'
'    into table sprovi
'    fields terminated by '|'
'    lines terminated by '\n'
Private Sub CargaTemporal1()
Dim sql As String
Dim fic As String
    fic = Trim(Dir1.Path) & "\Harsw6600.txt"
    fic = Replace(fic, "\", "/")
    
    'sql = "load data infile 'c:/Mbgstld4/Migra/Harsw6600.txt' "
    sql = "load data infile '" & Trim(fic) & "' "
    sql = sql & "into table tempnc1 "
    sql = sql & "fields terminated by '|' "
    sql = sql & "lines terminated by '\n'"
    
    Conn.Execute sql
    
End Sub


Private Sub InsertarTEMPNC(fecha As String, hora As String, dosimetro As String, c1 As String, c2 As String, c3 As String, c4 As String)
Dim sql As String
    
    sql = "insert into tempnc (codusu, fecha_lectura, hora_lectura, n_dosimetro, "
    sql = sql & "cristal_1, cristal_2, cristal_3, cristal_4, sistema) VALUES (" & vUsu.codigo & ",'"
    sql = sql & Format(fecha, FormatoFecha) & "','" & Format(hora, FormatoHora) & "','"
    sql = sql & Val(Trim(dosimetro)) & "'," & TransformaComasPuntos(c1) & ","
    sql = sql & TransformaComasPuntos(c2) & "," & TransformaComasPuntos(c3) & ","
    sql = sql & TransformaComasPuntos(c4) & ", 'P')"
    
    Conn.Execute sql
      
End Sub

' Cargar fichero de panasonic.
Private Sub BtnCargar_Click()
Dim Numlinea As Integer
Dim linea As String
Dim panaobj As CPanasonic
Dim I As Integer
Dim destino As String
On Error GoTo eBtnCargar

  If File1.FileName = "" Then
    MsgBox "Debe de seleccionar un fichero a migrar", vbExclamation, "¡Error!"
    Exit Sub
  End If
  
  ' Abrimos el fichero.
  Screen.MousePointer = vbHourglass
  Conn.Execute "delete from tempnc where codusu = " & vUsu.codigo
  If Right(Dir1.Path, 1) <> "\" Then Caracter = "\" Else Caracter = ""
  Open Dir1.Path + Caracter + File1.FileName For Input As #1
  destino = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "migra\"
  
  ' Abrimos los ficheros para escribir los resúmenes.
  Open destino & "Panasonic.txt" For Output As #2 Len = 128
  Open destino & "InformePana.txt" For Output As #3
  
  ' Empezamos a leer el fichero y guardar la info en la tabla temporal.
  Numlinea = 0
  I = 0
  While Not EOF(1)
 
    Input #1, linea

    ' Si la línea empieza por 2, es importante para nosotros.
    If Left(linea, 1) = "2" Then
      Set panaobj = New CPanasonic
      If Not panaobj.Cargar(linea) Then
        Close #1
        Close #2
        Close #3
        Exit Sub
      End If
      
      ' Escribimos la cabecera, si procede.
      If Numlinea = 0 Or Numlinea > 35 Then
        Numlinea = 0
        Print #3, "                 INFORME DE MIGRACION PANASONIC       archivo.." + File1.FileName
        Print #3, " "
        Print #3, " "
        Print #3, " "
        Print #3, "Tipo    Fecha     Hora Lec.  Dosimet Ref   Cristal 1       Cristal 2      Cristal 3      Cristal 4"
        Print #3, "==== ===========  =========  ======= ===  ============   ============   ============   ============"
      End If
      Numlinea = Numlinea + 1
      
      ' Escribimos la linea correspondiente en cada fichero.
      Print #3, "---"; Spc(3); panaobj.fecha; Spc(3); panaobj.hora; Spc(3); panaobj.dosimetro; Spc(3); " "; Spc(3); FormateaCeros(CStr(panaobj.E1)); Spc(3); FormateaCeros(CStr(panaobj.E2)); Spc(3); FormateaCeros(CStr(panaobj.E3)); Spc(3); FormateaCeros(CStr(panaobj.E4))
      Print #2, panaobj.fecha; "|"; panaobj.hora; "|"; panaobj.dosimetro; "|"; panaobj.E1; "|"; panaobj.E2; "|"; panaobj.E3; "|"; panaobj.E4; "|"
      
      ' Insertamos el registro en la tabla temporal.
      InsertarTEMPNC panaobj.fecha, panaobj.hora, panaobj.dosimetro, panaobj.E1, panaobj.E2, panaobj.E3, panaobj.E4
      
      ' Mostramos la información por pantalla.
      I = I + 1
      Label3(0).Caption = panaobj.fecha
      Label3(1).Caption = panaobj.hora
      Label3(2).Caption = panaobj.dosimetro
      Label3(4).Caption = CStr(panaobj.E1)
      Label3(5).Caption = CStr(panaobj.E2)
      Label3(6).Caption = CStr(panaobj.E3)
      Label3(7).Caption = CStr(panaobj.E4)
      Label3(8).Caption = I
      Set panaobj = Nothing
    
    End If
  
  Wend
  
  ' Mostramos el mensaje de Fin de carga y cerramos los ficheros y esas cosas.
  MsgBox "Archivo Transferido... ", vbExclamation, "Carga del archivo."
  Screen.MousePointer = vbDefault
  Close #1
  Close #2
  Close #3
  MsgBox "RECUERDE DE PONER EL PAPEL EN APAISADO SI DESEA IMPRIMIR ESTE INFORME..", , "¡Atención!"
  Unload Me
  
  Exit Sub

eBtnCargar:
  
  Screen.MousePointer = vbDefault
  Close #1
  Close #2
  Close #3
  If Err.Number = 70 Then
    MsgBox "El dispositivo se encuentra protegido contra escritura.", vbExclamation, "¡Error!"
  End If

End Sub

'Dim fondoS As Single
'Dim fondoD As Single
'Dim lista As Dictionary
'Dim auxS As String
'Dim auxD As String

'  ' Creamos una lista con todos los dosímetros y sus valores.
'  fondoS = 0
'  fondoD = 0
'  Set lista = New Dictionary


'      If Not panaobj.procesar Then
'        Close #1
'        Exit Sub
'      End If
'
'      lista.Add lista.Count, panaobj
'      fondoS = fondoS + panaobj.Hs
'      fondoD = fondoD + panaobj.Hd

'  ' Sacamos la media de los fondos.
'  fondoS = fondoS / lista.Count
'  fondoD = fondoD / lista.Count

'    auxS = CStr(lista(I).Hs - fondoS)
'    auxD = CStr(lista(I).Hd - fondoD)



Private Function FormateaCeros(Numero As String) As String
    Dim Entero As String
    Dim resto As String
    Dim solucion As String
    If Numero = 0 Then
        FormateaCeros = "0000.000    "
        Exit Function
    End If
    
    If Int(CDbl(Numero)) = 0 Then
        Entero = "0000"
    Else
        If Len(CStr(Int(CDbl(Numero)))) > 4 Then
          Numero = CStr(Round2(Val(Numero), 2))
        End If
        Entero = String(4 - Len(CStr(Int(CDbl(Numero)))), "0") + CStr(Int(Numero))
    End If
    resto = Right(Numero, Len(Numero) - InStr(1, Numero, ","))
    If Len(resto) > 5 Then resto = Left(resto, 5)
    solucion = Entero + "." + resto
    FormateaCeros = solucion + String(12 - Len(solucion), " ")
    
End Function
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo eDrive1change

Dir1.Path = Drive1.Drive

eDrive1change:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
End Sub

'Private Sub Form_Load()
'Dir1.Path = "c:\MBgstld4\Migra\"
'File1.Path = Dir1.Path
''
''Ponemos Valores Predeterminados
''
'Valor_Minimo_Luz_C4 = 999999
'Valor_Minimo_Luz_C2 = 999999
'Valor_Minimo_Ruido_C4 = 9999
'Valor_Minimo_Ruido_C2 = 9999
'End Sub







