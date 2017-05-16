VERSION 5.00
Begin VB.Form FrmHarshaw6600 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Automática de Dosis HARSHAW 6.600"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frmHarshaw6600.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   450
      TabIndex        =   4
      Top             =   300
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   450
      TabIndex        =   3
      Top             =   780
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   2730
      Pattern         =   "*.asc"
      TabIndex        =   2
      Top             =   780
      Width           =   1935
   End
   Begin VB.CommandButton BtnCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   3210
      TabIndex        =   1
      Top             =   4260
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4050
      TabIndex        =   0
      Top             =   4260
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Archivo del Harshaw"
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
      TabIndex        =   23
      Top             =   420
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha..."
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   22
      Top             =   2940
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Hora..."
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   21
      Top             =   3180
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Dosimetro"
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   20
      Top             =   3420
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Dosimetro"
      Height          =   255
      Index           =   3
      Left            =   450
      TabIndex        =   19
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  1"
      Height          =   255
      Index           =   4
      Left            =   3090
      TabIndex        =   18
      Top             =   2940
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  2"
      Height          =   255
      Index           =   5
      Left            =   3090
      TabIndex        =   17
      Top             =   3180
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  3"
      Height          =   255
      Index           =   6
      Left            =   3090
      TabIndex        =   16
      Top             =   3420
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cristal  4"
      Height          =   255
      Index           =   7
      Left            =   3090
      TabIndex        =   15
      Top             =   3660
      Width           =   735
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   1650
      TabIndex        =   14
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1650
      TabIndex        =   13
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1650
      TabIndex        =   12
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1650
      TabIndex        =   11
      Top             =   3660
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   3810
      TabIndex        =   10
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   3810
      TabIndex        =   9
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   3810
      TabIndex        =   8
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   3810
      TabIndex        =   7
      Top             =   3660
      Width           =   975
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   1650
      TabIndex        =   6
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "N. Registro"
      Height          =   255
      Index           =   8
      Left            =   450
      TabIndex        =   5
      Top             =   4380
      Width           =   1095
   End
End
Attribute VB_Name = "FrmHarshaw6600"
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
    File1.Pattern = "*.ASC"
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
    sql = sql & TransformaComasPuntos(c4) & ", 'H')"
    
    Conn.Execute sql
    
    
End Sub


'??????????????????????
'Public Valor_Minimo_Luz_C4 As Double
'Public Valor_Maximo_Luz_C4 As Double
'Public Valor_Minimo_Luz_C2 As Double
'Public Valor_Maximo_Luz_C2 As Double
'Public Valor_Minimo_Ruido_C4 As Double
'Public Valor_Minimo_Ruido_C2 As Double
'Public Valor_Maximo_Ruido_C4 As Double
'Public Valor_Maximo_Ruido_C2 As Double
'Public Descrip_Dosimetro As String
'Public N_Linea As Integer


Private Sub BtnCargar_Click()
Dim N_registro As Integer

Dim fecha As String
Dim hora As String
Dim dosimetro As String
Dim tipo_dosimetro As String
Dim C_1 As String
Dim C_2 As String
Dim C_3 As String
Dim C_4 As String
Dim Campos(29) As String
Dim Archivo_origen As String
Dim f As Integer
Dim salto As Integer
Dim destino As String

On Error GoTo eBtnCargar

    
If File1.FileName = "" Then
    MsgBox "Debe de seleccionar un fichero a migrar", vbExclamation, "¡Error!"

Else
'
'ABRIMOS EL FICHERO SECUENCIAL PARA GENERARL
'CON EL FORMATO DE TUBERIAS
'
Conn.Execute "delete from tempnc where codusu = " & vUsu.codigo
Screen.MousePointer = vbHourglass

If Right(Dir1.Path, 1) <> "\" Then Caracter = "\" Else Caracter = ""

Archivo_origen = Dir1.Path + Caracter + File1.FileName
destino = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "migra\"
Open Archivo_origen For Input As #1 Len = 128
'
'Abrimos el Archivo Destino
'
'¡antes estaba esto
'Open "c:\Mbgstld4\Migra\Harsw6600.txt" For Output As #2 Len = 128
'Open "c:\Mbgstld4\Migra\Informe.txt" For Output As #3
'ahora
Open destino & "Harsw6600.txt" For Output As #2 Len = 128
Open destino & "Informe6600.txt" For Output As #3

'
'Imprime Cabecera
'
GoSub Cabecera

'
'Desempaquetamos el registro
'
N_registro = 0
While Not EOF(1)
For f = 1 To 29
    Input #1, Campos(f)
Next f
'
'Organizamos la Informacion Desempaquetada
'
fecha = Mid$(Campos(12), 1, 8)
fecha = Right(fecha, 2) + "/" + Mid(fecha, 5, 2) + "/" + Left(fecha, 4)
hora = Campos(13)
hora = Left(hora, 2) + ":" + Mid(hora, 3, 2) + ":" + Right(hora, 2)
dosimetro = Campos(14)

dosimetro = dosimetro + String(5 - Len(dosimetro), " ")

tipo_dosimetro = Campos(16)
If tipo_dosimetro = "" Then tipo_dosimetro = " "
C_1 = Val(Campos(19)) / 16000
C_2 = Val(Campos(22)) / 16000
C_3 = Val(Campos(25)) / 16000
C_4 = Val(Campos(28)) / 16000
'
'Presentamos los Datos en Pantalla
'
Label3(0).Caption = fecha
Label3(1).Caption = hora
Label3(2).Caption = dosimetro
Label3(3).Caption = tipo_dosimetro
Label3(4).Caption = C_1
Label3(5).Caption = C_2
Label3(6).Caption = C_3
Label3(7).Caption = C_4
N_registro = N_registro + 1
Label3(8).Caption = N_registro
'
'Imprimimos Informacion de las Luces de Referencias
'
Select Case tipo_dosimetro
       Case "R"
       Descrip_Dosimetro = "LUZ"
       '
       'Comprobamos la luz de referencia del canal B
       '
       If Valor_Maximo_Luz_C4 < Val(C_4) Then Valor_Maximo_Luz_C4 = Val(C_4)
       '
       'Comprobamos el Valor Minimo
       '
       If Val(C_4) < Valor_Minimo_Luz_C4 And Val(C_4) <> 0 Then Valor_Minimo_Luz_C4 = Val(C_4)
       '
       'Comprobamos la Luz de referencia del Canal A
       '
       If Valor_Maximo_Luz_C2 < Val(C_2) Then Valor_Maximo_Luz_C2 = Val(C_2)
       '
       'Comprobamos el Valor Minimo
       '
       If Val(C_2) < Valor_Minimo_Luz_C2 And Val(C_2) <> 0 Then Valor_Minimo_Luz_C2 = Val(C_2)
       '
       Case "P"
       '
       Descrip_Dosimetro = "PMT"
       '
       'Comprobamos Ruido PMT del canal B
       '
       If Valor_Maximo_Ruido_C4 < (C_4) Then Valor_Maximo_Ruido_C4 = (C_4)
       '
       'Comprobamos el Valor Minimo
       '
       If (C_4) < Valor_Minimo_Ruido_C4 And (C_4) <> 0 Then Valor_Minimo_Ruido_C4 = (C_4)
       '
       'Comprobamos Ruido PMT del canal A
       '
       If Valor_Maximo_Ruido_C2 < (C_2) Then Valor_Maximo_Ruido_C2 = (C_2)
       '
       'Comprobamos el Valor Minimo
       '
       If (C_2) < Valor_Minimo_Ruido_C2 And (C_2) <> 0 Then Valor_Minimo_Ruido_C2 = (C_2)
       '
       Case Else
       Descrip_Dosimetro = "---"
End Select
'
'Imprimimos el Registro para el Wordpad
'
Print #3, Descrip_Dosimetro; Spc(3); fecha; Spc(3); hora; Spc(3); dosimetro; Spc(3); tipo_dosimetro; Spc(3); FormateaCeros(C_1); Spc(3); FormateaCeros(C_2); Spc(3); FormateaCeros(C_3); Spc(3); FormateaCeros(C_4)
N_Linea = N_Linea + 1
If N_Linea > 35 Then
                 N_Linea = 0
                 GoSub Cabecera
End If
'
'Montamos el Registro Destino
'
If tipo_dosimetro <> "R" And tipo_dosimetro <> "P" Then
       Print #2, fecha; "|"; hora; "|"; dosimetro; "|"; C_1; "|"; C_2; "|"; C_3; "|"; C_4; "|"
       
       InsertarTEMPNC fecha, hora, dosimetro, C_1, C_2, C_3, C_4
End If
If N_registro = 38 Then
Debug.Print
End If
Wend
'
'Cerramos los Archivos
'
Close #1
Close #2
'
'Forzamos un salto a otra pagina si no cabe el resumen del informe
'
If (35 - N_Linea) < 21 Then
        For salto = 0 To 35 - N_Linea
        Print #3, ""
        Next salto
        N_Linea = 0
        GoSub Cabecera
End If

MsgBox "Archivo Transferido... ", vbExclamation, "Carga del archivo."
'
'Presenta en el Informe las Desviaciones de lectura
'
Print #3, ""
Print #3, "LUZ DE REFERENCIA.... "
Print #3, "===================== "
Print #3, "       CANAL A....."; "                            CANAL B....."
Print #3, ""
Print #3, "VALOR MAXIMO........"; FormateaCeros(CStr(Valor_Maximo_Luz_C2)); Spc(5); "VALOR MAXIMO........"; FormateaCeros(CStr(Valor_Maximo_Luz_C4))
Print #3, "DESVIACION.........."; FormateaCeros(CStr((Valor_Maximo_Luz_C2 - Valor_Minimo_Luz_C2) / Valor_Minimo_Luz_C2 * 100)); " %"; Spc(3); "DESVIACION.........."; FormateaCeros(CStr((Valor_Maximo_Luz_C4 - Valor_Minimo_Luz_C4) / Valor_Minimo_Luz_C4 * 100)); " %"
Print #3, "VALOR MINIMO........"; FormateaCeros(CStr(Valor_Minimo_Luz_C2)); Spc(5); "VALOR MINIMO........"; FormateaCeros(CStr(Valor_Minimo_Luz_C4))
Print #3, ""
Print #3, ""
Print #3, "RUIDO DEL PMT........ "
Print #3, "===================== "
Print #3, "       CANAL A....."; "                            CANAL B....."
Print #3, ""
Print #3, "VALOR MAXIMO........"; FormateaCeros(CStr(Valor_Maximo_Ruido_C2)); Spc(5); "VALOR MAXIMO........"; FormateaCeros(CStr(Valor_Maximo_Ruido_C4))
Print #3, "DESVIACION.........."; FormateaCeros(CStr((Valor_Maximo_Ruido_C2 - Valor_Minimo_Ruido_C2) / Valor_Minimo_Ruido_C2 * 100)); " %"; Spc(3); "DESVIACION.........."; FormateaCeros(CStr((Valor_Maximo_Ruido_C4 - Valor_Minimo_Ruido_C4) / Valor_Minimo_Ruido_C4 * 100)); " %"
Print #3, "VALOR MINIMO........"; FormateaCeros(CStr(Valor_Minimo_Ruido_C2)); Spc(5); "VALOR MINIMO........"; FormateaCeros(CStr(Valor_Minimo_Ruido_C4))
Print #3, ""
Print #3, ""
Close #3
'
'Abrimos el Documento para Imprimir
'
MsgBox "RECUERDE DE PONER EL PAPEL EN APAISADO SI DESEA IMPRIMIR ESTE INFORME..", , "¡Atención!"

'Shell ("WORDPAD.EXE c:\mbgstld4\migra\informe.txt"), vbMaximizedFocus

Unload Me
End If
Screen.MousePointer = vbDefault
Exit Sub

Cabecera:

Print #3, "                 INFORME DE MIGRACION HARSHAW 6.600   archivo.." + File1.FileName
Print #3, " "
Print #3, " "
Print #3, " "
Print #3, "Tipo    Fecha     Hora Lec.  Dosimet Ref   Cristal 1       Cristal 2      Cristal 3      Cristal 4"
Print #3, "==== ===========  =========  ======= ===  ============   ============   ============   ============"

Return

eBtnCargar:
    If Err.Number = 70 Then
        MsgBox "El dispositivo se encuentra protegido contra escritura.", vbExclamation, "¡Error!"
    End If

End Sub
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






