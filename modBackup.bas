Attribute VB_Name = "modBackup"
Option Explicit


Public Sub BACKUP_TablaIzquierda(ByRef Rs As ADODB.Recordset, ByRef CADENA As String)
Dim I As Integer
Dim nexo As String

    CADENA = ""
    nexo = ""
    For I = 0 To Rs.Fields.Count - 1
        CADENA = CADENA & nexo & Rs.Fields(I).Name
        nexo = ","
    Next I
    CADENA = "(" & CADENA & ")"
End Sub


'---------------------------------------------------
'El fichero siempre sera NF
Public Sub BACKUP_Tabla(ByRef Rs As ADODB.Recordset, ByRef Derecha As String)
Dim I As Integer
Dim nexo As String
Dim valor As String
Dim Tipo As Integer
    Derecha = ""
    nexo = ""
    For I = 0 To Rs.Fields.Count - 1
        Tipo = Rs.Fields(I).Type
        
        If IsNull(Rs.Fields(I)) Then
            valor = "NULL"
        Else
            
            'pruebas
            Select Case Tipo
            'TEXTO
            Case 129, 200, 201
                valor = Rs.Fields(I)
                NombreSQL valor    '.-----------> 23 Octubre 2003.
                valor = "'" & valor & "'"
            'Fecha
            Case 133
                valor = CStr(Rs.Fields(I))
                valor = "'" & Format(valor, FormatoFecha) & "'"
                
            ' campo hora añadido
            Case 134, 135
                valor = CStr((Rs.Fields(I)))
                
                valor = "'" & Format(valor, FormatoHora) & "'"
                
            'Numero normal, sin decimales
            Case 2, 3, 16 To 19
                valor = Rs.Fields(I)
            
            'Numero con decimales
            Case 131
                valor = CStr(Rs.Fields(I))
                valor = TransformaComasPuntos(valor)
            Case Else
                valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                valor = valor & vbCrLf & "SQL: " & Rs.Source
                valor = valor & vbCrLf & "Pos: " & I
                valor = valor & vbCrLf & "Campo: " & Rs.Fields(I).Name
                valor = valor & vbCrLf & "Valor: " & Rs.Fields(I)
                MsgBox valor, vbExclamation, "¡Error!"
                MsgBox "El programa finalizara. Avise al soporte técnico.", vbCritical
                End
            End Select
        End If
        Derecha = Derecha & nexo & valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
End Sub
