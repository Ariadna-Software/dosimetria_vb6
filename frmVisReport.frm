VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport 
   Caption         =   "Visor de informes"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   Icon            =   "frmVisReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5925
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      lastProp        =   600
      _cx             =   9340
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
Attribute VB_Name = "frmVisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Informe As String

'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean
Public BaseDatos As Byte
Public CampoOrden As Integer
Public Impresora As String
Public EstaImpreso As Boolean
Public ExportarPDF As Boolean
Public ConSubInforme As Boolean

Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim smrpt As CRAXDRT.Report
Dim Argumentos() As String
Dim PrimeraVez As Boolean

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    EstaImpreso = True
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
Dim campo1 As CRFields
Dim Campo2 As CRFields

On Error GoTo Err_Carga

    Dim I As Integer
    Screen.MousePointer = vbHourglass
    ' ### 29/03/2006 DavidV (pequeño cambio para que los informes funcionen tambien en mi
    ' ordenador y, por lo tanto, en los que sean tan "raros" como el mío)
    'Set mapp = CreateObject("CrystalRuntime.Application")
    Set mapp = New CRAXDRT.Application
    'Informe = "C:\Programas\Conta\Contabilidad\InformesD\sumas12.rpt"
    Set mrpt = mapp.OpenReport(Informe)


    If BaseDatos = 3 Then
        For I = 1 To mrpt.Database.Tables.Count
            mrpt.Database.Tables(I).SetLogOnInfo "Suministros", "gessocial", vConfig.User, vConfig.password
            mrpt.Database.Tables(I).Location = "gessocial." & mrpt.Database.Tables(I).Name
        Next I
    Else
        For I = 1 To mrpt.Database.Tables.Count
            'mrpt.Database.Tables(I).SetLogOnInfo  "vMbgstld4", vUsu.CadenaConexion, vConfig.User, vConfig.password
            If mrpt.Database.Tables(I).Name <> "sql" Then
                mrpt.Database.Tables(I).Location = vUsu.CadenaConexion & "." & mrpt.Database.Tables(I).Name
            End If
        Next I
    End If

    If ConSubInforme Then AbrirSubreport


    PrimeraVez = True
    EstaImpreso = False

    '@@@@
    If Impresora <> "" Then
        mrpt.SelectPrinter "", Impresora, ""
    End If

    CargaArgumentos
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree

    If CampoOrden <> 0 Then
        If CampoOrden < 0 Then
          mrpt.RecordSortFields.Add mrpt.Database.Tables(2).Fields(-CampoOrden), crAscendingOrder
        Else
          mrpt.RecordSortFields.Add mrpt.Database.Tables(1).Fields(CampoOrden), crAscendingOrder
        End If
    End If

    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If
    CRViewer1.ReportSource = mrpt
    If SoloImprimir Then
        mrpt.PrintOut False
        EstaImpreso = True
    Else
        CRViewer1.ViewReport
    End If
    Exit Sub
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "¡Error!"
    Set mapp = Nothing
    Set mrpt = Nothing
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

    OtrosParametros = "|Emp= """ & vParam.NombreEmpresa & """|" & OtrosParametros
    NumeroParametros = NumeroParametros + 1
    For I = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(I).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(I).Text = Parametro
            'Stop
        End If
    Next I

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
End Sub


Private Function DevuelveValor(ByRef valor As String) As Boolean
Dim I As Integer
Dim J As Integer
    valor = "|" & valor & "="
    DevuelveValor = False
    I = InStr(1, OtrosParametros, valor, vbTextCompare)
    If I > 0 Then
        I = I + Len(valor) + 1
        J = InStr(I, OtrosParametros, "|")
        If J > 0 Then
            valor = Mid(OtrosParametros, I, J - I)
            If valor = "" Then valor = " "
            DevuelveValor = True
        End If
    End If
End Function


Private Sub Exportar()
    mrpt.ExportOptions.DiskFileName = App.Path & "\docum.pdf"
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    mrpt.ExportOptions.PDFExportAllPages = True
    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
    mrpt.Export False

    'Si ha generado bien entonces
    CadenaDesdeOtroForm = "OK"

End Sub



Private Sub AbrirSubreport()
'Para cada subReport que encuentre en el Informe pone las tablas del subReport
'apuntando a la BD correspondiente
Dim crxSection As CRAXDRT.Section
Dim crxObject As Object
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim I As Byte

    For Each crxSection In mrpt.Sections
        For Each crxObject In crxSection.ReportObjects
             If TypeOf crxObject Is SubreportObject Then
                Set crxSubreportObject = crxObject
                Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
                For I = 1 To smrpt.Database.Tables.Count
                     smrpt.Database.Tables(I).SetLogOnInfo "Suministros", vUsu.CadenaConexion, vConfig.User, vConfig.password
                     smrpt.Database.Tables(I).Location = vUsu.CadenaConexion & "." & smrpt.Database.Tables(I).Name
                Next I
             End If
        Next crxObject
    Next crxSection

    Set crxSubreportObject = Nothing

   ' Get the ReportObject by name and cast it as a
   ' SubreportObject.
'   mrpt.Sections.Item(6).ReportObjects.Item(1).Name
'   If TypeOf mrpt.Sections.Item(6).ReportObjects.Item(1) Is subreportObject Then
'   End If

'   If TypeOf mrpt.ReportDefinition.ReportObjects.Item(reportObjectName) Is subreportObject Then
'      subreportObject = mrpt.ReportDefinition.ReportObjects.Item(reportObjectName)
'
'     ' Get the subreport name.
'      subreportName = subreportObject.subreportName
'      ' Open the subreport as a ReportDocument.
'      subreport = mrpt.OpenSubreport(subreportName)
'      ' Preview the subreport.
'
'   End If
End Sub





