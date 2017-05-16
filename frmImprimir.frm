VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión listados"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2430
      TabIndex        =   2
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkEMAIL 
         Caption         =   "Enviar e-mail"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   5535
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer
    '1 .- Listado Empresas
    '2 .- Listado de instalaciones
    '3 .- Listado de Operarios en instalaciones
    '4 .- Listado de Dosimetros (todos menos area)
    '5 .- Listado de Dosimetros area
    '7 .- Listado de Factores de calibracion 4400
    '8 .- Listado de Factores de calibracion 6600
    '9 .- Listado de Dosis homogeneas por instalacion
    '10.- Listado de Dosis no homogeneas por instalacion
    '11.- Listado de Dosis Area por instalación
    '12.- Listado de Dosis CSN
    '13.- Listado de provincias
    '14.- Listado de Tipos de medicion
    '15.- Listado de Ramas Genericas
    '16.- Listado de Ramas Específicas
    '17.- Listado de Tipos de trabajo
    '18.- Listado de errores de migracion
    '19.- Listado de fondos
    '20.- Listado de dosis no homogeneas por operario
    '21.- Carta de reclamacion de dosimetros no recibidos
    '22.- Listado de dosis acumuladas 12 meses por operarios
    '23.- Carta de Sobredosis al CSN
    '24.- Listado de Recepcion de dosimetros de cuerpo
    '25.- Etiquetas de Empresas
    '26.- Etiquetas de Instalaciones
    '27.- Etiquetas de Operarios
    '28.- Errores de migracion area
    '29.- Informe de migración Msv Personal (6600)
    '30.- Informe de migracion mSv Area (6600)
    '31.- Listado de Operarios con Sobredosis
    '32.- Listado de dosimetros penalizados
    '33.- Listado de lotes 6600
    '34.- Listado de lotes panasonic
    '35.- Listado de Factores de calibración panasonic
    '36.- Listado de Fondos Panasonic
    '37.- Informe de migracion mSv Personal (Panasonic)
    '38.- Informe de migracion mSv Área (Panasonic)
    '39.- Informe de migracion mSv No homogénea (6600)
    '39.- Informe de migracion mSv No homogénea (Panasonic)
    '41.- Listado de Dosis homogeneas por instalacion (fichero migrado)
    '42.- Listado de Dosis no homogeneas por instalacion (fichero migrado)
    '43.- Listado de Dosis Area por instalación (fichero migrado)
    'xx.- Listado de Operarios
    '-------------------------------------------------
    
    
    
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public CampoOrden As Integer
Public NomDocu As String

Public Titulo As String

Private MostrarTree As Boolean
Private nombre As String
Private MIPATH As String
Private Lanzado As Boolean
Private PrimeraVez As Boolean
Private AntPredeterminada As Printer
Private ImpresoraSalida As String
Private ConSubInforme As Boolean


Public email As Boolean
Public NombreMail As String


'Private ReestableceSoloImprimir As Boolean

Private Sub chkEMAIL_Click()
    If chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 0
End Sub

Private Sub chkSoloImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 Then Me.chkEMAIL.Value = 0
End Sub

Private Sub cmdConfigImpre_Click()
    Screen.MousePointer = vbHourglass
    'Me.CommonDialog1.Flags = cdlPDPageNums
    CommonDialog1.PrinterDefault = True
    CommonDialog1.ShowPrinter
    
    PonerNombreImpresora
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 And Me.chkEMAIL.Value = 1 Then
        MsgBox "Si desea enviar por mail no debe marcar vista preliminar", vbExclamation, "¡Atención!"
        Exit Sub
    End If
    'Form2.Show vbModal
    Imprime
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
If PrimeraVez Then
    espera 0.1
    CommitConexion
    If SoloImprimir Then
        Imprime
        Unload Me
    Else
        If email Then
            Me.Hide
            chkEMAIL.Value = 1
            Imprime
            Unload Me
        End If
    End If
End If
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Dim Cad As String

    Set AntPredeterminada = Printer
    
    PrimeraVez = True
    Lanzado = False
    CargaICO
    Cad = Dir(App.Path & "\impre.dat", vbArchive)
    ConSubInforme = False
    'If Opcion = 3 Then ConSubInforme = True
    'ReestableceSoloImprimir = False
    If Cad = "" Then
        chkSoloImprimir.Value = 0
        Else
        chkSoloImprimir.Value = 1
        'ReestableceSoloImprimir = True
    End If
    cmdImprimir.Enabled = True
    If SoloImprimir Then
        chkSoloImprimir.Value = 0
        Me.Frame2.Enabled = False
        chkSoloImprimir.Visible = False
    Else
        Frame2.Enabled = True
        chkSoloImprimir.Visible = True
    End If
    PonerNombreImpresora
    MostrarTree = False
    ImpresoraSalida = ""
    'A partir del infome 26, se trabajaba sobre la b de datos de informes(USUARIOS)
    
    frmVisReport.BaseDatos = 1
    
    MIPATH = App.Path & "\Informes\"
    
    Select Case Opcion
    Case 1
        Text1.Text = "Empresas"
        nombre = "empresas.rpt"
    
    Case 2
        Text1.Text = "Instalaciones"
        nombre = "instalaciones.rpt"
        
    Case 3
        Text1.Text = "Operarios en Instalaciones"
        If CampoOrden = -1 Then
          nombre = "UsuariosInstalacionesDNI.rpt"
        Else
          nombre = "UsuariosInstalacionesEmpre.rpt"
        End If
        CampoOrden = 0
        
    Case 4
'        Text1.Text = "Dosímetros a Cuerpo"
        Text1.Text = "Dosímetros "
        nombre = "DosimetrosApa.rpt"
        
    Case 5
        Text1.Text = "Dosímetros Area"
        nombre = "DosimetrosApaArea.rpt"
    
    Case 7
        Text1.Text = "Factores de Calibración 4400"
        nombre = "FactoresCalib4400.rpt"
    
    Case 8
        Text1.Text = "Factores de Calibración 6600"
        nombre = "FactoresCalib6600.rpt"
        
    Case 9
        Text1.Text = "Dosis Homogénea por Instalación (migrado)"
        nombre = "DosisHomoInstalacionNew.rpt"
 
    Case 10
        Text1.Text = "Dosis No Homogénea por Instalación"
        nombre = "DosisNoHomoInstalacion.rpt"
    
    Case 11
        Text1.Text = "Dosis Area por Instalación"
        nombre = "DosisAreaInstalacionNew.rpt"
    
    Case 12
        Text1.Text = Titulo
        nombre = "DosisColectivasCSN.rpt"
        
    Case 13
        Text1.Text = "Provincias"
        nombre = "Provincias.rpt"
        
    Case 14
        Text1.Text = "Tipos de Medición"
        nombre = "TiposMedicion.rpt"
    
    Case 15
        Text1.Text = "Ramas Genéricas"
        nombre = "RamasGenericas.rpt"
    
    Case 16
        Text1.Text = "Ramas Específicas"
        nombre = "RamasEspecificas.rpt"
        
    Case 17
        Text1.Text = "Tipos de Trabajo"
        nombre = "TiposTrabajo.rpt"
    
    Case 18
        Text1.Text = "Errores de Migración"
        nombre = "ErroresMigracion.rpt"
    
    Case 19
        Text1.Text = "Fondos Harshaw 6600"
        nombre = "Fondos.rpt"
    
    Case 20
        Text1.Text = "Dosis no Homogéneas por Operario"
        nombre = "DosisNHomogOpe.rpt"
    
    Case 21
        Text1.Text = "Cartas de Reclamación de Dosímetros no recibidos"
        nombre = "CartaDosimNRec.rpt"
    
    Case 22
        Text1.Text = "Informe de Dosis por Operario Año Oficial"
        nombre = "DosisOpeAcum12.rpt"
    
    Case 23
        Text1.Text = "Carta de Sobredosis al CSN"
        nombre = "CartaSobredosis.rpt"
    
    Case 24
        Text1.Text = Titulo
        nombre = NomDocu
'        If CampoOrden = 0 Then
'            nombre = "RecepDosimCuerpo.rpt"
'        Else
'            nombre = "RecepDosimCuerpoInstala.rpt"
'        End If
    Case 25
        Text1.Text = "Etiquetas de Empresas"
        nombre = "EtiqEmpresas.rpt"
    
    Case 26
        Text1.Text = "Etiquetas de Instalaciones"
        nombre = "EtiqInstalaciones.rpt"
    
    Case 27
        Text1.Text = "Etiquetas de Operarios"
        nombre = "EtiqUsuarios1.rpt"
    
    Case 28
        Text1.Text = "Errores de Migración Area"
        nombre = "ErroresMigracionArea.rpt"
    
    Case 29
        Text1.Text = "Informe de Migración mSv Personal (Harshaw 6600)"
        nombre = "InformeMigracionmSv.rpt"
        
    Case 30
        Text1.Text = "Informe de Migración mSv Area (Harshaw 6600)"
        nombre = "InformeMigracionmSv.rpt"
    
    Case 31
        Text1.Text = "Informe de Operarios con Sobredosis"
        nombre = "OperariosSobredosis.rpt"
    
    Case 32
        Text1.Text = "Dosimetros Penalizados"
        OtrosParametros = "usu= " & vUsu.codigo & "|"
        nombre = "DosimetrosPenalizados.rpt"
    
    Case 33
        Text1.Text = "Lotes Harshaw 6600"
        nombre = "Lotes.rpt"
   
    Case 34
        Text1.Text = "Lotes Panasonic"
        nombre = "LotesPana.rpt"
    
    Case 35
        Text1.Text = "Factores de Calibración Panasonic"
        nombre = "FactoresCalibPana.rpt"
   
    Case 36
        Text1.Text = "Fondos Panasonic"
        nombre = "FondosPana.rpt"
        
    Case 37
        Text1.Text = "Informe de Migración mSv Personal (Panasonic)"
        nombre = "InformeMigracionmSvPana.rpt"
        
    Case 38
        Text1.Text = "Informe de Migración mSv Área (Panasonic)"
        nombre = "InformeMigracionmSvPana.rpt"
    
    Case 39
        Text1.Text = "Informe de Migración mSv No Homogénea (6600)"
        nombre = "InformeMigracionmSvNoHomog.rpt"
        
    Case 40
        Text1.Text = "Informe de Migración mSv No Homogénea (Panasonic)"
        nombre = "InformeMigracionmSvNoHomogPana.rpt"
    
    Case 41
        Text1.Text = "Dosis Homogénea por Instalación (migrado)"
        nombre = "DosisHomoInstalacionNew.rpt"

    Case 42
        Text1.Text = "Dosis No Homogénea por Instalación (migrado)"
        nombre = "DosisNoHomoInstalacionMigra.rpt"
        
    Case 43
        Text1.Text = "Dosis Area por Instalación (migrado)"
        nombre = "DosisAreaInstalacionNew.rpt"
        
    Case Else
        Text1.Text = "Opcion incorrecta"
        Me.cmdImprimir.Enabled = False
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Function Imprime() As Boolean
Dim Seguir As Boolean
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParametros
        .NumeroParametros = NumeroParametros
        .MostrarTree = MostrarTree
        .Informe = MIPATH & nombre
        .ExportarPDF = (chkEMAIL.Value = 1)
        .CampoOrden = CampoOrden
        .Impresora = ImpresoraSalida
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
    
    If Me.chkEMAIL.Value = 1 Then
        If Not email Then
            frmEMail.Asunto = Text1.Text
            frmEMail.NombreMail = NombreMail
            frmEMail.Mail = ""
            frmEMail.Show vbModal
        End If
        CadenaDesdeOtroForm = ""
    End If
    Unload Me
 
End Function


Private Sub Form_Unload(Cancel As Integer)
    If Me.chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 1
    'If ReestableceSoloImprimir Then SoloImprimir = False
    OperacionesArchivoDefecto
    
    Dim I As Printer
    For Each I In Printers
        If I.DeviceName = AntPredeterminada.DeviceName Then
            Set Printer = I
            Exit For
        End If
    Next
        
End Sub

Private Sub OperacionesArchivoDefecto()
Dim crear  As Boolean
On Error GoTo ErrOperacionesArchivoDefecto

crear = (Me.chkSoloImprimir.Value = 1)
'crear = crear Or ReestableceSoloImprimir
If Not crear Then
    Kill App.Path & "\impre.dat"
    Else
        FileCopy App.Path & "\Vacio.dat", App.Path & "\impre.dat"
End If
ErrOperacionesArchivoDefecto:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Text1_DblClick()
Frame2.Tag = Val(Frame2.Tag) + 1
If Val(Frame2.Tag) > 2 Then
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next
    Label1.Caption = Printer.DeviceName
    If Err.Number <> 0 Then
        Label1.Caption = "No hay impresora instalada"
        Err.Clear
    End If
End Sub

Private Sub CargaICO()
    On Error Resume Next
    Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
    If Err.Number <> 0 Then Err.Clear
End Sub

