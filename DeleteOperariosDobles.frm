VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Dim Sql As String
    Dim sql2 As String
    Dim dni As String
    Dim Rs As ADODB.Recordset
    
    Sql = "Select dni,c_empresa, c_instalacion,f_alta  from operarios ORDER BY dni,c_empresa, c_instalacion,f_alta"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        dni = Rs.Fields(0).Value
    End If
    dni = ""
    While Not Rs.EOF
       
        If Trim(Rs.Fields(0).Value) = Trim(dni) Then
            sql2 = "delete from operarios where dni = '" & Trim(Rs.Fields(0).Value) & "' and "
            sql2 = sql2 & " c_empresa = '" & Trim(Rs.Fields(1).Value) & "' and "
            sql2 = sql2 & " c_instalacion = '" & Trim(Rs.Fields(2).Value) & "' and "
            sql2 = sql2 & " f_alta = '" & Format(Rs.Fields(3).Value, FormatoFecha) & "'"
            Conn.Execute sql2
        Else
            dni = Rs.Fields(0).Value
        End If
        Rs.MoveNext
    Wend
    
'        sql2 = "delete from operarios where dni= '" & Trim(Rs.Fields(0).Value) & "' and "
'        sql2 = sql2 & "c_empresa = '" & Trim(Rs.Fields(1).Value) & "' and c_instalacion = '"
'        sql2 = sql2 & Trim(Rs.Fields(2).Value) & "' and f_alta <> '" & Format(Rs.Fields(3).Value, FormatoFecha) & "'"
'
'        Conn.Execute sql2
'        Rs.MoveNext
'    Wend
    
    Set Rs = Nothing
    
End Sub
