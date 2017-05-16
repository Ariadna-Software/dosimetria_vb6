VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFichero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ----------------------------------------------------------------------------
' Nombre de Clase   : CFichero
' Autor             : David Vallés Galve
' Dependencias      : N/A
' Fecha de Creación : 08/02/2006
' Ultima Revisión   : 13/02/2006
' ----------------------------------------------------------------------------

Option Explicit

' Nombre de la clase.
' -------------------
Const CLASS_NAME = "Fichero"

' Variables privadas.
' -------------------
Dim vFichero        As Integer
Dim vAbierto        As Boolean
Dim vTamanyo        As Long

' Propiedades de la clase.
' ------------------------
Private pNombre         As String
Private pLinea          As String
Private pPorcentaje     As Single

' Interface de acceso a la clase.
' -------------------------------

' Nombre del Fichero
Public Property Get nombre() As String
  nombre = pNombre
End Property
Public Property Let nombre(ByVal newValue As String)
  pNombre = newValue
End Property

' Línea (sólo lectura, de momento)
Public Property Get linea() As String
  linea = pLinea
End Property
'Public Property Let linea(ByVal newValue As String)
'  pLinea = newValue
'End Property

' Porcentaje (sólo lectura)
Public Property Get porcentaje() As String
  porcentaje = pPorcentaje
End Property
'Public Property Let porcentaje(ByVal newValue As String)
'  pPorcentaje = newValue
'End Property

' Funciones públicas.
' -------------------

' Abrir el fichero.
Public Sub abrir(Optional nFich As String, Optional escrib As Boolean)
On Error GoTo ECargar

  ' No se ha dado el nombre del fichero.
  If pNombre = "" And nFich = "" Then
    Err.Raise 1, CLASS_NAME & ".cargar", "Falta indicar el nombre del fichero"
  End If
  
  ' Si le pasamos parámetro a abrir, actualizamos la propiedad.
  If nFich <> "" Then pNombre = nFich
  
  ' El fichero no existe
  If Dir(pNombre) = "" And Not escrib Then
    Err.Raise 1, CLASS_NAME & ".cargar", "Fichero no encontrado: " & pNombre
  End If
  
  ' Abrimos el fichero.
  vFichero = FreeFile
  If escrib Then
    Open pNombre For Append As #vFichero
    vAbierto = True
  Else
    vTamanyo = FileLen(pNombre)
    pPorcentaje = 0
    
    Open pNombre For Input As #vFichero
    vAbierto = True
  End If
  Exit Sub
  
ECargar:
  
  ' Provocamos el error.
  Err.Raise Err.Number, CLASS_NAME & ".cargar", Err.Description

End Sub
        
' Leer Línea.
Public Function leerLinea() As Boolean
On Error GoTo ELeerLinea

  leerLinea = False
  
  ' Si el fichero está abierto...
  If vAbierto Then
    If Not EOF(vFichero) Then
      ' Leemos una linea y devolvemos verdadero
      Line Input #vFichero, pLinea
      leerLinea = True
      pPorcentaje = pPorcentaje + ((Len(pLinea) + 2) * 100) / vTamanyo
      If pPorcentaje >= 100 Then pPorcentaje = 100
    Else
      ' Hemos llegado al final del fichero, lo cerramos.
      cerrar
    End If
  End If
  Exit Function

ELeerLinea:

  ' Provocamos el error.
  Err.Raise Err.Number, CLASS_NAME & ".leerLinea", Err.Description
  
End Function

' Escribir Línea.
Public Function escribirLinea(str As String) As Boolean
On Error GoTo EescribirLinea

  escribirLinea = False
  
  ' Si el fichero está abierto...
  If vAbierto Then
    
    ' Escribimos una linea y devolvemos verdadero
      Print #vFichero, str
      escribirLinea = True
  End If
  Exit Function

EescribirLinea:

  ' Provocamos el error.
  Err.Raise Err.Number, CLASS_NAME & ".escribirLinea", Err.Description
  
End Function


' Cerrar el fichero.
Public Sub cerrar()
On Error GoTo ECerrar
    
  ' Cerramos el fichero, a no ser que ya esté cerrado.
  If vAbierto Then
    Close #vFichero
    vAbierto = False
  End If
  
  Exit Sub

ECerrar:

  ' Provocamos el error.
  Err.Raise Err.Number, CLASS_NAME & ".cerrar", Err.Description
  
End Sub

' Métodos y funciones privadas.
' -----------------------------

' Al morir la clase, cerramos el fichero.
Private Sub Class_Terminate()
  cerrar
End Sub
