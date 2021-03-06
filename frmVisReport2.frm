VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor de informes"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmVisReport2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3615
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      lastProp        =   600
      _cx             =   7646
      _cy             =   6376
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisReport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+-     Autor: DAVID     +-+-
' +-+- Alguns canvis: C�SAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit
Public Informe As String
Public InfConta As Boolean 'Enlazar a la Contabilidas

'SubInforme con conexion a la contabilidad. Conectar a las
'tablas de la BDatos correspondiente a la empresa: conta1, conta2, etc.

Public ConSubInforme As Boolean 'Si tiene subinforme ejecta la funcion AbrirSubInforme para enlazar esta a la BD correspondiente


'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean
Public Opcion As Integer
Public ExportarPDF As Boolean
Public EstaImpreso As Boolean


Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim smrpt As CRAXDRT.Report

Dim Argumentos() As String
Dim PrimeraVez As Boolean

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    If mrpt.PrinterSetupEx(0) = 0 Then
        mrpt.PrintOut False, 1
        EstaImpreso = True
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
            Screen.MousePointer = vbHourglass
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim I As Integer
On Error GoTo Err_Carga

    Screen.MousePointer = vbHourglass
    
    ' ### 29/03/2006 DavidV (peque�o cambio para que los informes funcionen tambien en mi
    ' ordenador y, por lo tanto, en los que sean tan "raros" como el m�o)
    'Set mapp = CreateObject("CrystalRuntime.Application")
    Set mapp = New CRAXDRT.Application
'    Informe = "C:\Programas\Ariges 4\Informes\rptMarcas.rpt"
    Set mrpt = mapp.OpenReport(Informe)



    'Conectar a la BD de la Empresa
    '####Descomentar
'    For i = 1 To mrpt.Database.Tables.Count
'       mrpt.Database.Tables(i).SetLogOnInfo "vAriges", vEmpresa.BDAriges, vConfig.User, vConfig.Password
'       If InStr(1, mrpt.Database.Tables(i).Name, "_") = 0 Then
'               mrpt.Database.Tables(i).Location = vEmpresa.BDAriges & "." & mrpt.Database.Tables(i).Name
'       End If
'    Next i

' lo he quitado yo
'    If InfConta Then
'        For i = 1 To mrpt.Database.Tables.Count
'           mrpt.Database.Tables(i).SetLogOnInfo "vconta", vEmpresa.BDConta, "root", "aritel"
'           If InStr(1, mrpt.Database.Tables(i).Name, "_") = 0 Then
'                   mrpt.Database.Tables(i).Location = vEmpresa.BDConta & "." & mrpt.Database.Tables(i).Name
'           End If
'        Next i
'    End If

'    If SubInformeConta <> "" Then
'        Set smrpt = mrpt.OpenSubreport(SubInformeConta)
'        For i = 1 To smrpt.Database.Tables.Count
'            smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamConta.NumeroConta, vParamConta.UsuarioConta, vParamConta.PasswordConta
'            smrpt.Database.Tables(i).Location = "conta" & vParamConta.NumeroConta & "." & smrpt.Database.Tables(i).Name
'        Next i
'    End If



    PrimeraVez = True

    CargaArgumentos
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    mrpt.RecordSelectionFormula = FormulaSeleccion

    'poner en la select del subinforme los mismos criterios que los del informe
    If ConSubInforme Then SelectSubreport

'    If Opcion = 6 Then
'        Dim crD As CRAXDRT.DatabaseFieldDefinition
'        Dim crF As CRAXDRT.FieldObject
'        Dim crS As CRAXDRT.Section
'        Set crD = mrpt.Database.Tables(1).Fields(6)
'        mrpt.AddGroup 0, crD, crGCAnyValue, crAscendingOrder
'        Set crS = mrpt.Sections.Item("GH")
'        Set crF = crS.AddFieldObject("{sfamia.nomfamia}", 100, 0)
''        mrpt.RecordSortFields.Item(3).Parent = mrpt.RecordSortFields.Item(1)
''        mrpt.RecordSortFields.Item(3).Parent = mrpt.RecordSortFields.Item(2)
'    End If


    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If

    'lOS MARGENES
    PonerMargen

    EstaImpreso = False
    CRViewer1.ReportSource = mrpt


    If SoloImprimir Then
        'mrpt.PrintOut False
        mrpt.PrintOut True
        EstaImpreso = True
    Else
        CRViewer1.ViewReport
    End If
    Exit Sub
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical, "�Error!"
    Set mapp = Nothing
    Set mrpt = Nothing
    Set smrpt = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub CargaArgumentos()
Dim Parametro As String
Dim I As Integer
    'El primer parametro es el nombre de la empresa para todas las empresas
    ' Por lo tanto concaatenaremos con otros parametros
    ' Y sumaremos uno
    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
    ' La cadena de parametros
    'Si esta asignaremos su valor

'    OtrosParametros = "|Emp= """ & vEmpresa.nomempre & """|" & OtrosParametros
    OtrosParametros = "|" & OtrosParametros
Select Case NumeroParametros
Case 0
    '====Comenta: LAura
    'Solo se vacian los campos de formula que empiezan con "p" ya que estas
    'formulas se corresponden con paso de parametros al Report
    For I = 1 To mrpt.FormulaFields.Count
        If Left(Mid(mrpt.FormulaFields(I).Name, 3), 1) = "p" Then
            mrpt.FormulaFields(I).Text = """"""
        End If
    Next I
    '====
Case 1

    For I = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(I).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        'Debug.Print Parametro
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(I).Text = Parametro
        Else
'            mrpt.FormulaFields(I).Text = """"""
        End If
    Next I

Case Else
    'NumeroParametros = NumeroParametros + 1

    For I = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(I).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(I).Text = Parametro
        End If
    Next I
'    mrpt.RecordSelectionFormula
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
    Set smrpt = Nothing
End Sub


Private Function DevuelveValor(ByRef valor As String) As Boolean
Dim I As Integer
Dim J As Integer

    valor = "|" & valor & "="
    DevuelveValor = False
    I = InStr(1, OtrosParametros, valor, vbTextCompare)
    If I > 0 Then
        I = I + Len(valor)
        J = InStr(I, OtrosParametros, "|")
        If J > 0 Then
            valor = Mid(OtrosParametros, I, J - I)
            If valor = "" Then
                valor = " "
            Else
                CompruebaComillas valor
            End If
            DevuelveValor = True
        End If
    End If
End Function


Private Sub CompruebaComillas(ByRef Valor1 As String)
Dim Aux As String
Dim J As Integer
Dim I As Integer

    If Mid(Valor1, 1, 1) = Chr(34) Then
        'Tiene comillas. Con lo cual tengo k poner las dobles
        Aux = Mid(Valor1, 2, Len(Valor1) - 2)
        I = -1
        Do
            J = I + 2
            I = InStr(J, Aux, """")
            If I > 0 Then
              Aux = Mid(Aux, 1, I - 1) & """" & Mid(Aux, I)
            End If
        Loop Until I = 0
        Aux = """" & Aux & """"
        Valor1 = Aux
    End If
End Sub

Private Sub Exportar()
    mrpt.ExportOptions.DiskFileName = App.Path & "\docum.pdf"
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    mrpt.ExportOptions.PDFExportAllPages = True
    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
    mrpt.Export False
    'Si ha generado bien entonces
    CadenaDesdeOtroForm = "OK"
End Sub

Private Sub PonerMargen()
Dim Cad As String
Dim I As Integer
    On Error GoTo EPon
    Cad = Dir(App.Path & "\*.mrg")
    If Cad <> "" Then
        I = InStr(1, Cad, ".")
        If I > 0 Then
            Cad = Mid(Cad, 1, I - 1)
            If IsNumeric(Cad) Then
                If Val(Cad) > 4000 Then Cad = "4000"
                If Val(Cad) > 0 Then
                    mrpt.BottomMargin = mrpt.BottomMargin + Val(Cad)
                End If
            End If
        End If
    End If
    Exit Sub
EPon:
    Err.Clear
End Sub



Private Sub SelectSubreport()
'Para cada subReport que encuentre en el Informe pone a la del subReport
'la select del report
Dim crxSection As CRAXDRT.Section
Dim crxObject As Object
Dim crxSubreportObject As CRAXDRT.SubreportObject
'Dim i As Byte
'
    For Each crxSection In mrpt.Sections
        For Each crxObject In crxSection.ReportObjects
             If TypeOf crxObject Is SubreportObject Then
                Set crxSubreportObject = crxObject
'                If crxSubreportObject.SubreportName <> SubInformeConta Then
                    Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
                    smrpt.RecordSelectionFormula = mrpt.RecordSelectionFormula

'                    For i = 1 To smrpt.Database.Tables.Count
'                         smrpt.Database.Tables(i).SetLogOnInfo "vAriges", vEmpresa.BDAriges, vConfig.User, vConfig.Password
'                         If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
'                            smrpt.Database.Tables(i).Location = vEmpresa.BDAriges & "." & smrpt.Database.Tables(i).Name
'                         End If
'                    Next i
'                End If
             End If
        Next crxObject
    Next crxSection
'
    Set crxSubreportObject = Nothing

End Sub




Private Sub AbrirSubreport()
'Para cada subReport que encuentre en el Informe pone las tablas del subReport
'apuntando a la BD correspondiente
'Dim crxSection As CRAXDRT.Section
'Dim crxObject As Object
'Dim crxSubreportObject As CRAXDRT.SubreportObject
'Dim i As Byte
'
'    For Each crxSection In mrpt.Sections
'        For Each crxObject In crxSection.ReportObjects
'             If TypeOf crxObject Is SubreportObject Then
'                Set crxSubreportObject = crxObject
'                If crxSubreportObject.SubreportName <> SubInformeConta Then
'                    Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
'                    For i = 1 To smrpt.Database.Tables.Count
'                         smrpt.Database.Tables(i).SetLogOnInfo "vAriges", vEmpresa.BDAriges, vConfig.User, vConfig.Password
'                         If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
'                            smrpt.Database.Tables(i).Location = vEmpresa.BDAriges & "." & smrpt.Database.Tables(i).Name
'                         End If
'                    Next i
'                End If
'             End If
'        Next crxObject
'    Next crxSection
'
'    Set crxSubreportObject = Nothing

End Sub

