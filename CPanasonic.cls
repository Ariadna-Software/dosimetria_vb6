VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPanasonic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Propiedades privadas.
Private pFecha As String
Private pHora As String
Private pDosimetro As String
Private pHs As Single
Private pHe As Single
Private pHd As Single
Private pE1 As Single
Private pE2 As Single
Private pE3 As Single
Private pE4 As Single
'variables locales para almacenar los valores de las propiedades
Private pcorrE1 As Single 'copia local
Private pcorrE2 As Single 'copia local
Private pcorrE3 As Single 'copia local
Private pcorrE4 As Single 'copia local
'variables locales para almacenar los valores de las propiedades
Private pcorrLote As Single 'copia local
'----
Dim A As Single
Dim b As Single
Dim C As Single
Dim d As Single
Dim E As Single
Dim f As Single
Dim XG As Single
Dim FG As Single
Dim FX As Single
Dim NEUTRON As Single
Dim GAMMA As Single
Dim BETAs As Single
Dim BETAe As Single
Dim BFs As Single
Dim BFe As Single
Dim NF As Single
Dim dosx As String
'---


Public Property Let corrLote(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.corrLote = 5
    pcorrLote = vData
End Property


Public Property Get corrLote() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.corrLote
    corrLote = pcorrLote
End Property



Public Property Let corrE4(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.corrE4 = 5
    pcorrE4 = vData
End Property


Public Property Get corrE4() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.corrE4
    corrE4 = pcorrE4
End Property



Public Property Let corrE3(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.corrE3 = 5
    pcorrE3 = vData
End Property


Public Property Get corrE3() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.corrE3
    corrE3 = pcorrE3
End Property



Public Property Let corrE2(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.corrE2 = 5
    pcorrE2 = vData
End Property


Public Property Get corrE2() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.corrE2
    corrE2 = pcorrE2
End Property



Public Property Let corrE1(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.corrE1 = 5
    pcorrE1 = vData
End Property


Public Property Get corrE1() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.corrE1
    corrE1 = pcorrE1
End Property




' Acceso de SÓLO lectura a las propiedades privadas.
Public Property Get fecha() As String
  fecha = pFecha
End Property
Public Property Get hora() As String
  hora = pHora
End Property
Public Property Get dosimetro() As String
  dosimetro = pDosimetro
End Property
Public Property Get Hs() As Single
  Hs = pHs
End Property
Public Property Get He() As Single
  He = pHe
End Property
Public Property Get Hd() As Single
  Hd = pHd
End Property

' Acceso TOTAL a otras propiedades privadas.
Public Property Get E1() As Single
  E1 = pE1
End Property
Public Property Let E1(ByVal newValue As Single)
  pE1 = newValue
End Property
Public Property Get E2() As Single
  E2 = pE2
End Property
Public Property Let E2(ByVal newValue As Single)
  pE2 = newValue
End Property
Public Property Get E3() As Single
  E3 = pE3
End Property
Public Property Let E3(ByVal newValue As Single)
  pE3 = newValue
End Property
Public Property Get E4() As Single
  E4 = pE4
End Property
Public Property Let E4(ByVal newValue As Single)
  pE4 = newValue
End Property

' Carga las propiedades a partir de una línea formateada de Panasonic.
Public Function Cargar(ByVal linea As String) As Boolean
On Error GoTo ECargar
  
  ' Fecha, hora y nº de dosímetro.
  pFecha = Mid(linea, 13, 2) & "/" & Mid(linea, 11, 2) & "/" & Mid(linea, 9, 2)
  pHora = Mid(linea, 15, 2) + ":" + Mid(linea, 17, 2)
  pDosimetro = Mid(linea, 24, 7)

  ' Esto es una movida para calcular en mSv el valor correspondiente a cada cristal,
  ' basándonos en el dígito previo a cada valor, el cual utilizo para la potencia de
  ' 10 que multiplicará dicho valor. El valor siempre viene en formato 000.
  '
  '   Dígito  Unidades  Formato
  '     4        Sv      0.00
  '     3       mSv       000
  '     2       mSv      00.0
  '     1       mSv      0.00
  '     0       µSv       000
  pE1 = Val(Mid(linea, 32, 3)) * (10 ^ (Val(Mid(linea, 31, 1)) - 3))
  pE2 = Val(Mid(linea, 36, 3)) * (10 ^ (Val(Mid(linea, 35, 1)) - 3))
  pE3 = Val(Mid(linea, 40, 3)) * (10 ^ (Val(Mid(linea, 39, 1)) - 3))
  pE4 = Val(Mid(linea, 44, 3)) * (10 ^ (Val(Mid(linea, 43, 1)) - 3))
  Cargar = True
  Exit Function
  
ECargar:
  Cargar = False
  MsgBox "Error " & Err.Number & " cargando línea del fichero Panasonic: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "¡Error!"
  Err.Clear
End Function

' He hecho las menos variaciones posibles respecto a la lógica del algoritmo
' original con el fin de que sea más fácil su mantenimiento, pero he tenido que
' variar algún aspecto lógico para ahorrar complejidad al implementarlo.
Public Function procesar() As Boolean
Dim A As Single
Dim b As Single
Dim C As Single
Dim d As Single
Dim E As Single
Dim f As Single
Dim XG As Single
Dim FG As Single
Dim FX As Single
Dim NEUTRON As Single
Dim GAMMA As Single
Dim BETAs As Single
Dim BETAe As Single
Dim BFs As Single
Dim BFe As Single
On Error GoTo EProcesar
    '-- (VRS 1.2.2) Aplicamos las correcciones
    pE1 = pE1 * pcorrE1 * pcorrLote
    pE2 = pE2 * pcorrE2 * pcorrLote
    pE3 = pE3 * pcorrE3 * pcorrLote
    pE4 = pE4 * pcorrE4 * pcorrLote
  
  ' Primeras asignaciones.
  If pE2 <> 0 Then
    A = Round2(pE3 / pE2, 3) ' VRS:1.3.5
    b = Round2(pE1 / pE2, 3) ' VRS:1.3.5
    d = Round2(pE4 / pE2, 3) ' VRS:1.3.5
  End If
  If pE4 <> 0 Then
    C = Round2(pE3 / pE4, 3) ' VRS:1.3.5
    E = Round2(pE1 / pE4, 3) ' VRS:1.3.5
  End If
  If pE1 - pE4 <> 0 Then f = Round2((pE2 - pE4) / (pE1 - pE4), 3) ' VRS:1.3.5
  
  If b > 8 Or (A < 1.5 And E >= 1.1) Then
    ' GAMMA+BETA OR GAMMA+NEUTRON
    GAMMA = pE4
    '-- (VRS 1.2.2) Ahora el camino es siempre BETA PLUS GAMMA
      BFs = 1.95 - (3.81 * f)
      '-- (VRS 1.2.2) Control de valores límite, nuevo también
      If BFs < 0.8 Then BFs = 0.8
      If BFs > 1.95 Then BFs = 1.95
      If f > 0.25 Then
        BFe = 0.12
      Else
        BFe = 0
      End If
      BETAs = Round2(BFs * (pE1 - pE4), 3) ' VRS:1.3.5
      BETAe = Round2(BFe * (pE1 - pE4), 2) ' VRS:1.3.5
      pHs = Round2(BETAs + GAMMA, 3) ' VRS:1.3.5
      pHe = Round2(BETAe + GAMMA, 3) ' VRS:1.3.5
      pHd = GAMMA
'    If f < 0.5 Then
'       BETA PLUS GAMMA
'      BFs = 1.95 - (3.81 * f)
'      If f > 0.25 Then
'        BFe = 0.12
'      Else
'        BFe = 0
'      End If
'      BETAs = BFs * (pE1 - pE4)
'      BETAe = BFe * (pE1 - pE4)
'      pHs = BETAs + GAMMA
'      pHe = BETAe + GAMMA
'      pHd = GAMMA
'    Else
'       NEUTRON PLUS GAMMA
'      NEUTRON = 0.7 * (pE2 - pE4)
'      pHs = NEUTRON + GAMMA
'      pHe = NEUTRON + GAMMA
'      pHd = NEUTRON + GAMMA
'    End If
  Else
    ' PHOTONS
    If C <= 1.5 Then
      ' GAMMA
      pHs = pE2
      pHe = pE2
      pHd = pE2
    ElseIf C >= 24 Then
      ' LG
      pHs = 1.85 * pE2
      pHe = 1.43 * pE2
      pHd = 0.8 * pE2
    ElseIf d <= 0.9 Then
      ' LG PLUS GAMMA
      XG = -0.3873 + (0.08195 * C) + (0.002234 * (C ^ 2))
      FG = 1 / (1 + XG)
      FX = 1 - FG
      pHs = (1.85 * FX * pE2) + (FG + pE2) 'ojo rafa
      pHe = (1.43 * FX * pE2) + (FG + pE2)
      pHd = (0.8 * FX * pE2) + (FG + pE2)
    ElseIf E >= 0.95 And C > 3.5 Then
      If A > 13.5 Then
        ' LI
        pHs = 1.52 * pE2
        pHe = 1.36 * pE2
        pHd = 1.1 * pE2
      Else
        ' LI PLUS GAMMA
        XG = -2.487 + (0.7897 * C) - (0.02831 * (C ^ 2))
        FG = 1 / (1 + XG)
        FX = 1 - FG
        pHs = (1.52 * FX * pE2) + (FG * pE2)
        pHe = (1.36 * FX * pE2) + (FG * pE2)
        pHd = (1.1 * FX * pE2) + (FG * pE2)
      End If
    Else
      If C > 10 And A > 13 Then
        ' LK
        pHs = 1.37 * pE2
        pHe = 1.26 * pE2
        pHd = 1.14 * pE2
      ElseIf C > 7 And A > 10 Then
        ' MFG
        pHs = 1.3 * pE2
        pHe = 1.25 * pE2
        pHd = 1.2 * pE2
      ElseIf d > 1.3 Then
        ' MFI
        pHs = 1.36 * pE2
        pHe = 1.34 * pE2
        pHd = 1.31 * pE2
      Else
        ' MID
        pHs = 1.18 * pE2
        pHe = 1.14 * pE2
        pHd = 1.11 * pE2
      End If
    End If
  End If
  procesar = True
  Exit Function
  
EProcesar:
  procesar = False
  MsgBox "Error " & Err.Number & " procesando fichero Panasonic: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "¡Error!"
  Err.Clear
End Function

'-------
' procesar2:
' Nueva función de cálculo del algoritmo Panasonic.
'-------
Public Function procesar2(Optional dosim As String) As Boolean

On Error GoTo EProcesar2
    dosx = dosim
    '-- (VRS 1.2.2) Aplicamos las correcciones
    pE1 = pE1 * pcorrE1 * pcorrLote
    pE2 = pE2 * pcorrE2 * pcorrLote
    pE3 = pE3 * pcorrE3 * pcorrLote
    pE4 = pE4 * pcorrE4 * pcorrLote
    '----
    A = pE3 / pE2
    b = pE1 / pE2
    C = pE3 / pE4
    d = pE4 / pE2
    E = pE1 / pE4
    f = (pE2 - pE4) / (pE1 - pE4)
    f = 0.3 '' 08/07/2010 vicent i javier
    ''If f < 0 Then f = 0 ' 08/07/2010 lo dice javier
    '-----
    If b > 8 Then
        GAMMABETA
    Else
        PR1
    End If
    
    procesar2 = True
    Exit Function
EProcesar2:
    MsgBox Err.Description & vbCrLf & "E:" & Err.Number, vbCritical, "PROCESAR 2"
    procesar2 = False
End Function

'--- PR1
Private Sub PR1()
    If A >= 1.5 Then
        PHOTONS
    ElseIf E < 1.1 Then
        PHOTONS
    Else
        GAMMABETA
    End If
End Sub

'--- PHOTONS
Private Sub PHOTONS()
    If C > 1.5 Then
        PR2
    Else
        'GAMMA
        pHs = pE2
        pHe = pE2
        pHd = pE2
    End If
End Sub

'--- PR2
Private Sub PR2()
    If C < 24 Then
        PR3
    Else
        'LG
        pHs = 1.85 * pE2
        pHe = 1.42 * pE2
        pHd = 0.8 * pE2
    End If
End Sub

'--- PR3
Private Sub PR3()
    If d > 0.9 Then
        PR4
    Else
        'LG PLUS GAMMA
        XG = -0.3873 + (0.08195 * C) + (0.002234 * (C ^ 2))
        FG = 1 / (1 + XG)
        FX = 1 - FG
        pHs = (1.85 * FX * pE2) + (FG * pE2)
        pHe = (1.43 * FX * pE2) + (FG * pE2)
        pHd = (0.8 * FX * pE2) + (FG * pE2)
    End If
End Sub

'--- GAMMABETA
Private Sub GAMMABETA()
    GAMMA = pE4
    If f < 0.5 Then
        BETAPLUSGAMMA
    Else
'        MsgBox "El dosímetro: " & dosx & " se va por la via NEUTRON PLUS GAMMA" & _
'                vbCrLf & "Valor F = " & f, vbInformation, "PROCESAR 2"
        'NEUTRONPLUSGAMMA
        f = 0.3
        BETAPLUSGAMMA '' 08/07/2010 lo dice javier
    End If

End Sub

'--- BETAPLUSGAMMA
Private Sub BETAPLUSGAMMA()
    BFs = 1.95 - (3.81 * f)
    '-- Control de error
    If (BFs < 0.8 Or BFs > 1.95) Then
        '-- Esto en teoría es erróneo.
        MsgBox "El dosímetro: " & dosx & " tiene un valor BFs incorrecto" & _
                vbCrLf & "Valor BFs = " & BFs, vbCritical, "PROCESAR 2"
        Exit Sub
    End If
    If f > 0.25 Then
        BFe = 0.12
    Else
        BFe = 0#
    End If
    BETAs = BFs * (pE1 - pE4)
    BETAe = BFe * (pE1 - pE4)
    pHs = BETAs + GAMMA
    pHe = BETAe + GAMMA
    pHd = GAMMA
End Sub

'--- NEUTRONPLUSGAMMA
Private Sub NEUTRONPLUSGAMMA()
    NF = 0.7
    NEUTRON = NF * (pE2 - pE4)
    pHs = NEUTRON + GAMMA
    pHe = NEUTRON + GAMMA
    pHd = NEUTRON + GAMMA
End Sub

'----PR4
Private Sub PR4()
    If E >= 0.95 Then
        PR5
    Else
        PR6
    End If
End Sub

'--- PR5
Private Sub PR5()
    If C > 3.5 Then
        If A > 13.5 Then
            '-- LI
            pHs = 1.52 * pE2
            pHe = 1.36 * pE2
            pHd = 1.1 * pE2
        Else
            '-- LI PLUS GAMMA
            XG = -2.487 + (0.787 * C) - (0.02831 * (C ^ 2))
            FG = 1 / (1 + XG)
            FX = 1 - FG
            pHs = (1.52 * FX * pE2) + (FG * pE2)
            pHe = (1.36 * FX * pE2) + (FG * pE2)
            pHd = (1.1 * FX * pE2) + (FG * pE2)
        End If
    Else
        PR6
    End If
End Sub

'--- PR6
Private Sub PR6()
    If C > 10 Then
        If A <= 13 Then
            PR7
        Else
            '-- LK
            pHs = 1.37 * pE2
            pHe = 1.26 * pE2
            pHd = 1.14 * pE2
        End If
    Else
        PR7
    End If
End Sub

'--- PR7
Private Sub PR7()
    If C > 7 Then
        If A <= 10 Then
            PR8
        Else
            'MFG
            pHs = 1.3 * pE2
            pHe = 1.25 * pE2
            pHd = 1.2 * pE2
        End If
    Else
        PR8
    End If
    

End Sub

'-- PR7
Private Sub PR8()
    If d > 1.3 Then
        pHs = 1.36 * E2
        pHe = 1.24 * pE2
        pHd = 1.31 * pE2
    Else
        pHs = 1.18 * pE2
        pHe = 1.14 * pE2
        pHd = 1.11 * pE2
    End If
End Sub

