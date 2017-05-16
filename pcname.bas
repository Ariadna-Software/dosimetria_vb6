Attribute VB_Name = "pcname"
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

'Función Api CopyMemory
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Función Api GetIpAddrTable para obtener la tabla de direcciones IP
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

Public Const MAX_COMPUTERNAME_LENGTH = 255
'Estructuras

Private Type IPINFO
dwAddr As Long
dwIndex As Long
dwMask As Long
dwBCastAddr As Long
dwReasmSize As Long
unused1 As Integer
unused2 As Integer
End Type

Private Type MIB_IPADDRTABLE
dEntrys As Long 'Numero de entradas de la tabla
mIPInfo(5) As IPINFO 'Array de entradas de direcciones Ip
End Type

Private Type IP_Array
mBuffer As MIB_IPADDRTABLE
BufferLen As Long
End Type

Public Function ComputerName() As String
    'Devuelve el nombre del equipo actual
    Dim sComputerName As String
    Dim ComputerNameLength As Long
    
    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
     ComputerName = Mid(sComputerName, 1, ComputerNameLength)
    
End Function


'Función para Convertir el valor de tipo Long a un string
Private Function ConvertirDirecciónAstring(longAddr As Long) As String
Dim myByte(3) As Byte 'array de tipo Byte
Dim Cnt As Long
CopyMemory myByte(0), longAddr, 4
For Cnt = 0 To 3
ConvertirDirecciónAstring = ConvertirDirecciónAstring + CStr(myByte(Cnt)) + "."
Next Cnt
ConvertirDirecciónAstring = Left$(ConvertirDirecciónAstring, Len(ConvertirDirecciónAstring) - 1)
End Function

'Función que retorna un string con la dirección Ip final
Public Function RecuperarIP() As String

Dim Ret As Long, Tel As Long
Dim bBytes() As Byte
Dim TempList() As String
Dim TempIP As String
Dim Tempi As Long
Dim Listing As MIB_IPADDRTABLE
Dim L3 As String

On Error GoTo errSub
GetIpAddrTable ByVal 0&, Ret, True

If Ret <= 0 Then Exit Function
ReDim bBytes(0 To Ret - 1) As Byte
ReDim TempList(0 To Ret - 1) As String

'recuperamos la tabla con las ip
GetIpAddrTable bBytes(0), Ret, False


CopyMemory Listing.dEntrys, bBytes(0), 4

For Tel = 0 To Listing.dEntrys - 1
'Copiamos la estructura entera a la lista
CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))

TempList(Tel) = ConvertirDirecciónAstring(Listing.mIPInfo(Tel).dwAddr)

Next Tel

TempIP = TempList(0)
For Tempi = 0 To Listing.dEntrys - 1
L3 = Left(TempList(Tempi), 3)
If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
TempIP = TempList(Tempi)
End If
Next Tempi
RecuperarIP = TempIP

Exit Function
errSub:

RecuperarIP = ""

End Function



