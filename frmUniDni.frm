VERSION 5.00
Begin VB.Form frmUniDni 
   Caption         =   "Unificación de operaciones por DNI (DOSIMETRIA)"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Frame frmDni2 
      Caption         =   "DNI 2"
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   6015
      Begin VB.TextBox txtDni2 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtDni2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "DNI:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre operario:"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame frmDNI1 
      Caption         =   "DNI 1"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   6015
      Begin VB.TextBox txtDni1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtDni1 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre operario"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "DNI:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label lblVersion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Caption         =   "Información de proceso"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   6015
   End
   Begin VB.Label lblIntro 
      Caption         =   $"frmUniDni.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "frmUniDni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------------------------
' UnificarOperacionesPorDNI
' (c) Ariadna Software S.L.
' Autor: Rafael García (rafa@myariadna.com)
'--------------------------------------------------------------------------------
' Version: 1.0.0
' Fecha: 05/04/2007
'--------------------------------------------------------------------------------
' Esta utilidad pretente pasar todos los registros asignados a un determinado DNI (Dni1)
' a otro DNI (Dni2). De este modo se salva un problema provocado por los usuarios
' que dan de alta dos usuarios (que en realidad son el mismo) y asignan operacines a
' cada uno de ellos. Cuando se dan cuenta deben corregir las operacions asignadas al
' usuario incorrecto y pasarlas al correcto. [JIRA: DSMT-19]
'--------------------------------------------------------------------------------
' Version: 1.0.1
' Fecha: 05/04/2007
'--------------------------------------------------------------------------------
' [1] Se actualiza operariosinstala para evitar error de claves referenciales
' al actualizar dosímetros

Public vConfig As Configuracion ' lleva los datos de configuración actual
Dim resultado As Integer ' guarda resultados de llamadas a métodos y funciones
Dim Mens As String ' variable auxiliar para mensajes

Dim conn As ADODB.Connection ' conexión con la base de datos
Dim sql As String ' ' variable auxiliar para instrucciones SQL
Dim rs As ADODB.Recordset ' recordset auxiliar

Private Sub cmdSalir_Click()
    '-- Nos vamos
    conn.Close
    End
End Sub

Private Sub Form_Load()
    '-- Mostrar versión
    lblVersion = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    Set vConfig = New Configuracion
    resultado = vConfig.Leer
    If resultado Then
        ' hay algun problema leyendo la configuración
        Mens = "No puede acceder a la configuración de la aplicación. " & _
                "Verifique que el fichero ConfigDosis.ini se encuentra en el directorio desde el que se ejecuta la utilidad."
        MsgBox Mens, vbCritical
        Unload Me
    End If
    If Not AbrirConexiones() Then
        ' no se puede acceder a la base de datos
        Mens = "No se puede acceder a la base de datos. La utilidad no puede continuar"
        MsgBox Mens, vbCritical
        Unload Me
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim Dni1 As String ' dni antiguo o incorreto
    Dim Dni2 As String ' dni nuevo o correcto
    Dim I As Long
    '-- Realiza coprobaciones básicas antes de lanzar la operación.
    If txtDni1(1) = "" Or txtDni2(1) = "" Then
        MsgBox "Debe introducir DNI1 y DNI2", vbExclamation
        Exit Sub
    End If
    If txtDni1(1) = "NO EXISTE" Or txtDni2(1) = "NO EXISTE" Then
        MsgBox "Los DNI introducidos deben estar dados de alta en operarios", vbExclamation
        Exit Sub
    End If
    '-- Pregunta final
    Mens = "Va a asignar todas la operaciones del DNI: " & txtDni1(0) & _
            " (" & txtDni1(1) & ") " & " al DNI:" & _
            txtDni2(0) & " (" & txtDni2(1) & ") " & vbCrLf & _
            "¿Desea continuar?"
    If MsgBox(Mens, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
    '-- Tomamos los valores de los DNI a actualizar
    Dni1 = Trim(txtDni1(0))
    Dni2 = Trim(txtDni2(0))
    '-- Y aqui es donde se hacen realmente los cambios protegidos por transaccion
    On Error GoTo Err_Aceptar:
        conn.Execute "START TRANSACTION"
            ' dosiscuerpo
            MuestraInformacion "Actualizando dosiscuerpo..."
            sql = "update dosiscuerpo set dni_usuario = '" & Dni2 & "'" & _
                    " where dni_usuario = '" & Dni1 & "'"
            conn.Execute sql
            ' recepdosim
            MuestraInformacion "Actualizando recepdosim..."
            sql = "update recepdosim set dni_usuario = '" & Dni2 & "'" & _
                    " where dni_usuario = '" & Dni1 & "'"
            conn.Execute sql
            ' dosisarea
            MuestraInformacion "Actualizando dosisarea..."
            sql = "update dosisarea set dni_usuario = '" & Dni2 & "'" & _
                    " where dni_usuario = '" & Dni1 & "'"
            conn.Execute sql
            ' dosisnohomog
            MuestraInformacion "Actualizando dosisnohomog..."
            sql = "update dosisnohomog set dni_usuario = '" & Dni2 & "'" & _
                    " where dni_usuario = '" & Dni1 & "'"
            conn.Execute sql
            ' dosimetros
            MuestraInformacion "Actualizando dosimetros..."
            '-- VRS1.0.1[1] Se actualiza operariosinstala para evitar errores en claves referenciales
            conn.Execute "SET FOREIGN_KEY_CHECKS = 0"
                sql = "update operainstala set dni = '" & Dni2 & "'" & _
                        " where dni = '" & Dni1 & "'"
                conn.Execute sql
            conn.Execute "SET FOREIGN_KEY_CHECKS = 1"
            sql = "update dosimetros set dni_usuario = '" & Dni2 & "'" & _
                    " where dni_usuario = '" & Dni1 & "'"
            conn.Execute sql
        conn.Execute "COMMIT" ' si todo acaba bien se confirma transacción
    Mens = "El proceso de unificación ha finalizado. ¿Desea unificar otro DNI?"
    If MsgBox(Mens, vbYesNo + vbQuestion) = vbYes Then
        '-- Se queda
        For I = 0 To 1
            txtDni1(I) = ""
            txtDni2(I) = ""
        Next I
        MuestraInformacion "Información de proceso"
        txtDni1(0).SetFocus
    Else
        '-- Se va
        cmdSalir_Click
    End If
    Exit Sub
Err_Aceptar:
    conn.Execute "ROLLBACK" ' si acaba mal se echa para atrás la tranasacción
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Err_Aceptar"
    Mens = "Se ha producido un error en la unificación de operaciones. Habitualmente este error es debido a que no ha sido posible actualizar" & _
            " la tabla de dosimetros porque los DNI introducidos no coinciden en empresa e instalación." & vbCrLf & _
            "Revise esta circunstancia antes de volver a ejecutar la utilidad"
    MsgBox Mens, vbInformation
    MuestraInformacion "Proceso no finalizado."
End Sub

Private Function AbrirConexiones() As Boolean
    '-- AbrirConexiones:
    '   Abre las conexiones necesarias para el programa si se produce algún error
    '   devuleve (false) si todo va bien devuelve (true)
On Error GoTo Err_AbrirConexiones
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=" & _
                    ";DATABASE=mbgstld4" & _
                    ";SERVER=" & vConfig.SERVER & _
                    ";UID=" & vConfig.User & _
                    ";PWD=" & vConfig.password & _
                    ";PORT=3306;OPTION=3;STMT="
    conn.Open
    AbrirConexiones = True
    Exit Function
Err_AbrirConexiones:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Err_AbrirConexiones"
End Function

Private Sub txtDni1_LostFocus(Index As Integer)
    '-- Muestra el nombre si se ha introducido un dni
    If Index = 0 Then
        txtDni1(0) = UCase(txtDni1(0))
        txtDni1(1) = NombreOperario(txtDni1(0))
    End If
End Sub

Private Sub txtDni2_LostFocus(Index As Integer)
    '-- Muestra el nombre si se ha introducido un dni
    If Index = 0 Then
        txtDni2(0) = UCase(txtDni2(0))
        txtDni2(1) = NombreOperario(txtDni2(0))
    End If
End Sub

Private Function NombreOperario(dni As String) As String
    '-- NombreOperario:
    '   Devuelve el nombre del operario al que corresponde el DNI pasado en
    '   parámetros. Si el operario no existe devuelve "No existe"
On Error GoTo Err_NombreOperario
    sql = "select concat(nombre, ' ', apellido_1, ' ' , apellido_2) from operarios" & _
            " where dni = '" & Trim(dni) & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly
    If Not rs.EOF Then
         NombreOperario = rs.Fields(0)
    Else
        NombreOperario = "NO EXISTE"
    End If
    Exit Function
Err_NombreOperario:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Err_NombreOperario"
End Function

Private Sub MuestraInformacion(Mensaje As String)
    '-- MuestraInformación
    '   Muestra el mensaje que se le pasa en un control label
    lblInf.Caption = Mensaje
    lblInf.Refresh
    DoEvents
End Sub
