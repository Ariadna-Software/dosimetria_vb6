VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOperariosInstala 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operarios en Instalaciones"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13215
   Icon            =   "frmOperariosInstala.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   13215
   Begin VB.TextBox Txtaux 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   12690
      MaxLength       =   4
      TabIndex        =   5
      Tag             =   "Migrado|T|S|||empresas|migrado|||"
      Top             =   5490
      Width           =   495
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   2
      Left            =   5730
      TabIndex        =   2
      Tag             =   "Dni de operario|T|N|||operainstala|dni||S|"
      Text            =   "Dato2"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.CommandButton ImgPpal 
      Height          =   255
      Index           =   2
      Left            =   7140
      TabIndex        =   20
      Top             =   5520
      Width           =   255
   End
   Begin VB.CommandButton CmdFec 
      Height          =   255
      Index           =   0
      Left            =   11190
      TabIndex        =   19
      Top             =   5520
      Width           =   255
   End
   Begin VB.CommandButton CmdFec 
      Height          =   255
      Index           =   1
      Left            =   12540
      TabIndex        =   18
      Top             =   5490
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3180
      MaxLength       =   30
      TabIndex        =   17
      Top             =   5490
      Width           =   2520
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7410
      MaxLength       =   30
      TabIndex        =   16
      Top             =   5490
      Width           =   2670
   End
   Begin VB.CommandButton ImgPpal 
      Height          =   255
      Index           =   1
      Left            =   2910
      TabIndex        =   15
      Top             =   5490
      Width           =   255
   End
   Begin VB.CommandButton ImgPpal 
      Height          =   255
      Index           =   0
      Left            =   1290
      TabIndex        =   14
      Top             =   5490
      Width           =   255
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   4
      Left            =   11460
      TabIndex        =   4
      Tag             =   "Fecha de baja|F|S|||operainstala|f_baja|dd/mm/yyyy||"
      Text            =   "Dato4"
      Top             =   5490
      Width           =   1335
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   3
      Left            =   10110
      TabIndex        =   3
      Tag             =   "Fecha de alta|F|N|||operainstala|f_alta|dd/mm/yyyy|S|"
      Text            =   "Dato3"
      Top             =   5520
      Width           =   1305
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10830
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11970
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Tag             =   "Código de Instalacion|T|N|||operainstala|c_instalacion||S|"
      Text            =   "Dato2"
      Top             =   5490
      Width           =   1395
   End
   Begin VB.TextBox Txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   260
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Tag             =   "Código de Empresa|T|N|||operainstala|c_empresa||S|"
      Text            =   "Dat"
      Top             =   5490
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11970
      TabIndex        =   10
      Top             =   5730
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   975
      Left            =   1440
      Top             =   3480
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1720
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":1A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":22F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":2BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":38F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":3A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":3B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperariosInstala.frx":42A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmOperariosInstala.frx":43BA
      Height          =   4965
      Left            =   60
      TabIndex        =   11
      Top             =   480
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   8758
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOperariosInstala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Empresa As String
Public instalacion As String
Public dni As String
'Public fechaalta As String

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmEmp As frmEmpresas
Attribute frmEmp.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmBc As frmBuscaGrid
Attribute frmBc.VB_VarHelpID = -1
Private WithEvents frmBd As frmBuscaGrid
Attribute frmBd.VB_VarHelpID = -1
Private WithEvents frmIns As frmInstalaciones
Attribute frmIns.VB_VarHelpID = -1
Private WithEvents frmOpe As frmOperarios
Attribute frmOpe.VB_VarHelpID = -1

Private CadenaConsulta As String
Private HaDevueltoDatos As Boolean
Dim fich As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim I As Integer
Dim ape1 As String
Dim ape2 As String
Dim nombre As String

Private ValorAnterior As String
'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim b As Boolean
Dim I As Integer

    Modo = vModo
    
    b = (Modo = 0)
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).Visible = Not b
    Next I
    For I = 0 To CmdFec.Count - 1
        CmdFec(I).Visible = Not b
    Next I
    For I = 0 To Imgppal.Count - 1
        Imgppal(I).Visible = Not b
    Next I
    For I = 0 To Text2.Count - 1
        Text2(I).Visible = Not b
    Next I
    
'    If Modo = 2 Then
'        imgppal(0).Visible = False
'        imgppal(1).Visible = False
'        imgppal(2).Visible = False
'    End If
    
    Toolbar1.Buttons(1).Enabled = b
    Toolbar1.Buttons(2).Enabled = b
    Toolbar1.Buttons(6).Enabled = b
    Toolbar1.Buttons(7).Enabled = b
    Toolbar1.Buttons(8).Enabled = b
    
    cmdAceptar.Visible = Not b
    cmdCancelar.Visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    End If
    'Si estamo mod or insert
    If Modo = 2 Then
       txtAux(0).BackColor = &H80000018
       txtAux(1).BackColor = &H80000018
       txtAux(2).BackColor = &H80000018
    Else
       txtAux(0).BackColor = &H80000005
       txtAux(1).BackColor = &H80000005
       txtAux(2).BackColor = &H80000005
    End If
    txtAux(2).Enabled = (Modo <> 2)
    
    txtAux(0).Enabled = (Empresa = "" And instalacion = "")
    txtAux(1).Enabled = (Empresa = "" And instalacion = "")
    txtAux(2).Enabled = (dni = "")
'    Txtaux(3).Enabled = (fechaalta = "")
    Imgppal(0).Enabled = txtAux(0).Enabled And (Modo <> 2)
    Imgppal(1).Enabled = txtAux(0).Enabled And (Modo <> 2)
    Imgppal(2).Enabled = txtAux(2).Enabled And (Modo <> 2)
    
    If Empresa <> "" Or instalacion <> "" Then
        txtAux(0).Text = Trim(Empresa)
        txtAux(1).Text = Trim(instalacion)
        Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", Trim(txtAux(0).Text) & "|" & Trim(txtAux(1).Text) & "|", "T|T|", 2)
    Else
        If dni <> "" Then txtAux(2).Text = Trim(dni)
        CargarDatosOperarios Trim(txtAux(2).Text), ape1, ape2, nombre
        Text2(1).Text = Trim(ape1) & " " & Trim(ape2) & ", " & Trim(nombre)
'        Txtaux(3) = Format(fechaalta, "dd/mm/yyyy")
    End If


End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    Dim I As Integer
    
    'Obtenemos la siguiente numero de factura
    'NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
          
    If DataGrid1.Row < 0 Then
        anc = 690
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 495 '545
    End If
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    LLamaLineas anc, 0
    
    'Ponemos el foco
    PonerFoco txtAux(0)
    
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
Dim I As Integer

    CadenaConsulta = "Select distinct operainstala.c_empresa, "
    CadenaConsulta = CadenaConsulta & " operainstala.c_instalacion, instalaciones.descripcion, "
    CadenaConsulta = CadenaConsulta & " operainstala.dni, "
    CadenaConsulta = CadenaConsulta & " concat(voperarios.apellido_1, ' ', voperarios.apellido_2, ',', nombre), "
    CadenaConsulta = CadenaConsulta & " operainstala.f_alta, operainstala.f_baja, operainstala.migrado  "
    CadenaConsulta = CadenaConsulta & "from operainstala, voperarios, instalaciones "
    CadenaConsulta = CadenaConsulta & "where operainstala.dni  = voperarios.dni and "
    CadenaConsulta = CadenaConsulta & "voperarios.codusu = " & vUsu.codigo & " and "
'    CadenaConsulta = CadenaConsulta & "operarios.f_baja is null and "
    CadenaConsulta = CadenaConsulta & "operainstala.c_empresa = instalaciones.c_empresa and "
    CadenaConsulta = CadenaConsulta & "operainstala.c_instalacion = instalaciones.c_instalacion "
    
    If Empresa <> "" And instalacion <> "" Then
        CadenaConsulta = CadenaConsulta & " and operainstala.c_empresa = '" & Trim(Empresa) & "' "
        CadenaConsulta = CadenaConsulta & " and operainstala.c_instalacion = '" & Trim(instalacion) & "'"
    ElseIf dni <> "" Then 'And fechaalta <> "" Then
        CadenaConsulta = CadenaConsulta & " and operainstala.dni = '" & Trim(dni) & "'"
'        CadenaConsulta = CadenaConsulta & " and operainstala.f_alta = '" & Format(fechaalta, FormatoFecha) & "'"
    
    End If

    CargaGrid ("operainstala.f_alta = '9999-12-31'")
    Me.lblIndicador.Caption = "BUSQUEDA"
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    LLamaLineas DataGrid1.Top + 206, 2
    
    PonerFoco txtAux(0)

End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 495 '545
    End If
    Cad = ""
    For I = 0 To 3
        Cad = Cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    Text2(0).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = DataGrid1.Columns(3).Text
    Text2(1).Text = DataGrid1.Columns(4).Text
    txtAux(3).Text = DataGrid1.Columns(5).Text
    txtAux(4).Text = DataGrid1.Columns(6).Text
    
'    i = adodc1.Recordset!tipoconce
'    Combo1.ListIndex = i - 1
    LLamaLineas anc, 1
   
   'Como es modificar
   PonerFoco txtAux(3)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    PonerModo xModo + 1
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        txtAux(I).Top = alto
    Next I
    Imgppal(0).Top = alto
    Imgppal(1).Top = alto
    Imgppal(2).Top = alto
    CmdFec(0).Top = alto
    CmdFec(1).Top = alto
    For I = 0 To Text2.Count - 1
        Text2(I).Top = alto
    Next I
End Sub

Private Sub BotonEliminar()
Dim sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    sql = "Seguro que desea eliminar el operario de la instalación:"
    sql = sql & vbCrLf & "Código de Empresa: " & adodc1.Recordset.Fields(0)
    sql = sql & vbCrLf & "Código de Instalación: " & adodc1.Recordset.Fields(1)
    sql = sql & vbCrLf & "Dni Operario: " & adodc1.Recordset.Fields(3)
    sql = sql & vbCrLf & "Fecha de Alta: " & adodc1.Recordset.Fields(5)
    
    If MsgBox(sql, vbQuestion + vbYesNoCancel, "¡Atención!") = vbYes Then
        'Hay que eliminar
        sql = "Delete from operainstala where c_empresa='" & Trim(adodc1.Recordset!c_empresa) & "' and "
        sql = sql & " c_instalacion = '" & Trim(adodc1.Recordset!c_instalacion) & "' and "
        sql = sql & " dni = '" & Trim(adodc1.Recordset!dni) & "' and "
        sql = sql & " f_alta = '" & Format(adodc1.Recordset!f_alta, FormatoFecha) & "'"
        
        Conn.Execute sql
        CargaGrid ""
    End If

Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Operario en la Instalación"
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me, 1) Then
                   ' i = Adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                  '  Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " ='" & i & "'")
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        Else
            MsgBox vbCrLf & "  Debe introducir alguna condición de búsqueda. " & vbCrLf, vbExclamation, "¡Error!"
            PonerModo 0
        End If
    End Select
    PonerFoco DataGrid1
    
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1
        DataGrid1.AllowAddNew = False
        'CargaGrid
        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        
    Case 3
        CargaGrid
    End Select
    PonerModo 0
    If Not adodc1.Recordset.EOF Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    Else
        lblIndicador.Caption = ""
    End If
    PonerFoco DataGrid1
End Sub


Private Sub cmdFec_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    Select Case Index
       Case 0 'fecha de inicio
            f = Now
            If txtAux(3).Text <> "" Then
                If IsDate(txtAux(3).Text) Then f = txtAux(3).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            txtAux(3).Text = frmC.fecha
                mTag.DarFormato txtAux(3)
            Set frmC = Nothing
       Case 1 ' fecha finalizacion
            f = Now
            If txtAux(4).Text <> "" Then
                If IsDate(txtAux(4).Text) Then f = txtAux(4).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            txtAux(4).Text = frmC.fecha
            mTag.DarFormato txtAux(4)
            Set frmC = Nothing
    End Select
End Sub

Private Sub cmdRegresar_Click()
    Dim Cad As String
    
    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation, "¡Atención!"
        Exit Sub
    End If
    
    Cad = adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & adodc1.Recordset.Fields(2) & "|"
    Cad = Cad & adodc1.Recordset.Fields(3) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
           Case vbESC, vbSalir
                If Modo = 0 Then
                    Unload Me
                Else
                    cmdCancelar_Click
                End If
           Case vbBuscar
                 Toolbar1_ButtonClick Toolbar1.Buttons(1)
           Case vbVerTodos
                 Toolbar1_ButtonClick Toolbar1.Buttons(2)
           Case vbAñadir
                 Toolbar1_ButtonClick Toolbar1.Buttons(6)
           Case vbModificar
                 Toolbar1_ButtonClick Toolbar1.Buttons(7)
           Case vbEliminar
                 Toolbar1_ButtonClick Toolbar1.Buttons(8)
           Case vbImprimir
                 Toolbar1_ButtonClick Toolbar1.Buttons(11)
        End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 0 Then lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    '## A mano
    'Vemos como esta guardado el valor del check
    
'    Me.Top = 0
'    Me.Height = 0
    
    chkVistaPrevia.Value = CheckValueLeer(Name)
      
    CargarOperarios
      
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        '.Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    'Cadena consulta
'    PonerOpcionesMenuGeneral Me

    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
    End If
    
    CadenaConsulta = "Select distinct operainstala.c_empresa, "
    CadenaConsulta = CadenaConsulta & " operainstala.c_instalacion, instalaciones.descripcion, "
    CadenaConsulta = CadenaConsulta & " operainstala.dni, "
    CadenaConsulta = CadenaConsulta & " concat(voperarios.apellido_1, ' ', voperarios.apellido_2, ',', nombre), "
    CadenaConsulta = CadenaConsulta & " operainstala.f_alta, operainstala.f_baja, operainstala.migrado  "
    CadenaConsulta = CadenaConsulta & "from operainstala, voperarios, instalaciones "
    CadenaConsulta = CadenaConsulta & "where operainstala.dni  = voperarios.dni and "
    CadenaConsulta = CadenaConsulta & "voperarios.codusu = " & vUsu.codigo & " and "
'    CadenaConsulta = CadenaConsulta & "operarios.f_baja is null and "
    CadenaConsulta = CadenaConsulta & "operainstala.c_empresa = instalaciones.c_empresa and "
    CadenaConsulta = CadenaConsulta & "operainstala.c_instalacion = instalaciones.c_instalacion "
    
    If Empresa <> "" And instalacion <> "" Then
        CadenaConsulta = CadenaConsulta & " and operainstala.c_empresa = '" & Trim(Empresa) & "' "
        CadenaConsulta = CadenaConsulta & " and operainstala.c_instalacion = '" & Trim(instalacion) & "'"
    ElseIf dni <> "" Then 'And fechaalta <> "" Then
        CadenaConsulta = CadenaConsulta & " and operainstala.dni = '" & Trim(dni) & "' "
'        CadenaConsulta = CadenaConsulta & " and operainstala.f_alta = '" & Format(fechaalta, FormatoFecha) & "' "
    End If
    CargaGrid

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmBc_Selecionado(CadenaDevuelta As String)
Dim ape1 As String
Dim ape2 As String
Dim nombre As String

    If CadenaDevuelta <> "" Then
        txtAux(2).Text = RecuperaValor(CadenaDevuelta, 1)
        ape1 = RecuperaValor(CadenaDevuelta, 2)
        ape2 = RecuperaValor(CadenaDevuelta, 3)
        nombre = RecuperaValor(CadenaDevuelta, 4)
        
        Text2(1).Text = Trim(ape1) & " " & Trim(ape2) & ", " & Trim(nombre)
    End If

End Sub

Private Sub frmEmp_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmIns_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub
'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Function SugerirCodigoSiguiente() As String
    Dim sql As String
    Dim Rs As ADODB.Recordset
    
    sql = "Select Max(codcomar) from scomar where codprovi = '" & txtAux(0).Text
    sql = sql & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, , , adCmdText
    sql = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            sql = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = sql
End Function

Private Sub frmOpe_DatoSeleccionado(CadenaSeleccion As String)
Dim ape1 As String
Dim ape2 As String
Dim nombre As String

    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1)
    CargarDatosOperarios Trim(txtAux(2).Text), ape1, ape2, nombre
    Text2(1).Text = Trim(ape1) & " " & Trim(ape2) & ", " & Trim(nombre)
    
End Sub

Private Sub imgppal_Click(Index As Integer)
    Select Case Index
        Case 0
            'codigo de empresa
'            Set frmEmp = New frmEmpresas
'            frmEmp.DatosADevolverBusqueda = "0|1|"
'            frmEmp.Show
            MandaBusquedaPreviaEmpresa ""
        Case 1
            'codigo de instalacion
'            Set frmIns = New frmInstalaciones
'            frmIns.DatosADevolverBusqueda = "0|13|1|"
'            frmIns.Show
            MandaBusquedaPreviaInstala ""
        Case 2
            'codigo de operario
            MandaBusquedaPreviaDni ""
'            Set frmOpe = New frmOperarios
'            frmOpe.DatosADevolverBusqueda = "9|13|10|5|"
'            frmOpe.Show
    
    
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
            BotonBuscar
    Case 2
            BotonVerTodos
    Case 6
            BotonAnyadir
    Case 7
            BotonModificar
    Case 8
            BotonEliminar
    Case 11
            'Imprimimos el listado
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 3 'Listado de operarios en instalaciones
            FrmListado.Show
    Case 12
            Unload Me
    Case Else
    
    End Select
End Sub

Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub

Private Sub CargaGrid(Optional sql As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    DataGrid1.Enabled = False
    adodc1.ConnectionString = Conn
    If sql <> "" Then
        sql = CadenaConsulta & " AND " & sql
        Else
        sql = CadenaConsulta
    End If
    sql = sql & " ORDER BY c_empresa, c_instalacion, dni, f_alta"
    adodc1.RecordSource = sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    DataGrid1.Enabled = True
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
    
    'codigo de empresa
    I = 0
        DataGrid1.Columns(I).Caption = "Empresa"
        DataGrid1.Columns(I).Width = 1400

    'instalacion
    I = 1
        DataGrid1.Columns(I).Caption = "Instalación"
        DataGrid1.Columns(I).Width = 1400
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'Descripcion de la instalacion
    I = 2
        DataGrid1.Columns(I).Caption = "Nombre"
        DataGrid1.Columns(I).Width = 3500
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'dni del operario
    I = 3
        DataGrid1.Columns(I).Caption = "DNI Operario"
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'nombre
    I = 4
        DataGrid1.Columns(I).Caption = "Nombre"
        DataGrid1.Columns(I).Width = 2300
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    I = 5
        DataGrid1.Columns(I).Caption = "Fecha Alta"
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    I = 6
        DataGrid1.Columns(I).Caption = "Fecha Baja"
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
        'añadido
        
    I = 7
        DataGrid1.Columns(I).Caption = "M"
        DataGrid1.Columns(I).Width = 300
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        Text2(0).Width = DataGrid1.Columns(2).Width - 60
        txtAux(2).Width = DataGrid1.Columns(3).Width - 60
        Text2(1).Width = DataGrid1.Columns(4).Width - 60
        txtAux(3).Width = DataGrid1.Columns(5).Width - 60
        txtAux(4).Width = DataGrid1.Columns(6).Width - 60
        txtAux(5).Width = DataGrid1.Columns(7).Width - 60
        
        txtAux(0).Left = DataGrid1.Left + 340
        Imgppal(0).Left = txtAux(0).Left + txtAux(0).Width - Imgppal(0).Width '- 55
        txtAux(1).Left = Imgppal(0).Left + Imgppal(0).Width + 55
        Imgppal(1).Left = txtAux(1).Left + txtAux(1).Width - Imgppal(1).Width '- 55
        Text2(0).Left = Imgppal(1).Left + Imgppal(1).Width + 55
        txtAux(2).Left = Text2(0).Left + Text2(0).Width + 55
        Imgppal(2).Left = txtAux(2).Left + txtAux(2).Width - Imgppal(1).Width '- 55
        Text2(1).Left = Imgppal(2).Left + Imgppal(2).Width + 55
        txtAux(3).Left = Text2(1).Left + Text2(1).Width + 55
        txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 55
        CmdFec(0).Left = txtAux(4).Left - 55 - CmdFec(0).Width
        CmdFec(1).Left = txtAux(4).Left + txtAux(4).Width - CmdFec(1).Width
        txtAux(5).Left = CmdFec(1).Left + CmdFec(1).Width + 60
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF

   If Not adodc1.Recordset.EOF Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        lblIndicador.Caption = ""
   End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    With txtAux(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim sql As String
Dim valor As Currency

    ''Quitamos blancos por los lados
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then Exit Sub
    If txtAux(Index).BackColor = vbYellow Then
        txtAux(Index).BackColor = vbWhite
    End If
    
    If txtAux(Index) = "" Then Exit Sub
    
    If ValorAnterior = txtAux(Index).Text Then Exit Sub
    
    If Modo = 3 And ConCaracteresBusqueda(txtAux(Index).Text) Then Exit Sub 'Busquedas
    
    Select Case Index
      Case 0, 1, 2, 5 ' TEXTOS
            If InStr(1, txtAux(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                txtAux(Index).Text = Replace(Format(txtAux(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco txtAux(Index)
                Exit Sub
            End If
                        
            If txtAux(0).Text <> "" Then
                sql = ""
                sql = DevuelveDesdeBD(1, "nom_comercial", "empresas", "c_empresa|", txtAux(0).Text & "|", "T|", 1)
                If sql = "" Then
                    MsgBox "Código de Empresa no existe. Reintroduzca.", vbExclamation, "¡Error!"
                    txtAux(0).Text = ""
                    PonerFoco txtAux(0)
                    Exit Sub
                End If
            End If
            
            If txtAux(0).Text <> "" And txtAux(1).Text <> "" Then
                Text2(0).Text = ""
                Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "instalaciones", "c_empresa|c_instalacion|", txtAux(0).Text & "|" & txtAux(1).Text & "|", "T|T|", 2)
                If Text2(0).Text = "" Then
                    MsgBox "Código de instalación no existe. Reintroduzca.", vbExclamation, "¡Error!"
                    txtAux(1).Text = ""
                    PonerFoco txtAux(1)
                    Exit Sub
                End If
            End If
            
            If txtAux(2).Text <> "" Then
                sql = ""
                sql = DevuelveDesdeBD(1, "apellido_1", "operarios", "dni|", txtAux(2).Text & "|", "T|", 1)
                If sql = "" Then
                    MsgBox "Dni de operario no existe. Reintroduzca.", vbExclamation, "¡Error!"
                    txtAux(2).Text = ""
                    PonerFoco txtAux(2)
                    Exit Sub
                Else
                    CargarDatosOperarios txtAux(2).Text, ape1, ape2, nombre
                    Text2(1).Text = Trim(ape1) & " " & Trim(ape2) & ", " & Trim(nombre)
                End If
            End If
        
      Case 3, 4 ' fechas
            If txtAux(Index).Text <> "" Then
              If Not EsFechaOK(txtAux(Index)) Then
                    MsgBox "Fecha incorrecta: " & txtAux(Index).Text, vbExclamation, "¡Error!"
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                    Exit Sub
              End If
              txtAux(Index).Text = Format(txtAux(Index).Text, "dd/mm/yyyy")
            End If
      
    End Select
    txtAux(Index).Text = Format(txtAux(Index).Text, ">")
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset

b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 1 Then
        'Estamos insertando
         Datos = DevuelveDesdeBD(1, "f_alta", "operainstala", "c_empresa|c_instalacion|dni|f_alta|", Trim(txtAux(0).Text) & "|" & Trim(txtAux(1).Text) & "|" & txtAux(2).Text & "|" & Format(txtAux(3).Text, "yyyy-MM-dd") & "|", "T|T|T|F|", 4)
         If Datos <> "" Then
            MsgBox "Ya existe el operario en la instalacion con esa fecha de alta. Reintroduzca.", vbExclamation, "¡Error!"
            b = False
         Else
           Set Rs = New ADODB.Recordset
           sql = "select * from operainstala where c_empresa='" & Trim(txtAux(0).Text) & "' and c_instalacion='" & Trim(txtAux(1).Text) & "' and dni='" & txtAux(2).Text & "' and f_baja is null"
           Rs.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
           If Not Rs.EOF Then
             MsgBox "Ya existe una relación activa entre ese operario y esa instalación. Reintroduzca.", vbExclamation, "¡Error!"
             b = False
           End If
           Rs.Close
           Set Rs = Nothing
           
         End If
    End If
    
    If txtAux(4).Text <> "" Then
        If txtAux(3).Text <> "" Then
            If CDate(txtAux(3).Text) > CDate(txtAux(4).Text) Then
                MsgBox "La fecha de baja no puede ser inferior a la fecha de alta. Reintroduzca.", vbExclamation, "¡Error!"
                b = False
            End If
        End If
    End If
    DatosOk = b
End Function

Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub MandaBusquedaPreviaDni(CadB As String)
Dim Cad As String
Dim tabla As String
Dim Titulo As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(txtAux(2), 15, "DNI")
        Cad = Cad & "Apellido 1|apellido_1|T|20·"
        Cad = Cad & "Apellido 2|apellido_2|T|20·"
        Cad = Cad & "Nombre|nombre|T|20·"
        Cad = Cad & "Profesion|profesion_catego|T|25·"
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmBc = New frmBuscaGrid
            frmBc.vCampos = Cad
            frmBc.vTabla = "operarios"
            frmBc.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmBc.vDevuelve = "0|1|2|3|"
            frmBc.vTitulo = "Operarios"
            frmBc.vSelElem = 1
            frmBc.vConexionGrid = 1
            frmBc.vCargaFrame = False
            '#
            frmBc.Show vbModal
            Set frmBc = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
        End If
        Screen.MousePointer = vbDefault

End Sub

Private Sub MandaBusquedaPreviaInstala(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(txtAux(0), 20, "Empresa")
        Cad = Cad & ParaGrid(txtAux(1), 20, "Código")
        Cad = Cad & "Nombre|descripcion|T|60·"
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = "instalaciones"
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Instalaciones"
            frmB.vSelElem = 0
            frmB.vConexionGrid = 1
            'frmB.vBuscaPrevia = chkVistaPrevia
            '#
            frmB.Show vbModal
            Set frmB = Nothing
        End If
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    If CadenaDevuelta <> "" Then
        txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
        txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
        Text2(0).Text = RecuperaValor(CadenaDevuelta, 3)
    End If
End Sub

Private Sub MandaBusquedaPreviaEmpresa(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(txtAux(0), 20, "Código")
        Cad = Cad & "Nombre Comercial|nom_comercial|T|60·"
        Cad = Cad & "CIF|cif_nif|T|15·"
        
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmBd = New frmBuscaGrid
            frmBd.vCampos = Cad
            frmBd.vTabla = "empresas"
            frmBd.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmBd.vDevuelve = "0|"
            frmBd.vTitulo = "Empresas"
            frmBd.vSelElem = 0
            frmBd.vConexionGrid = 1
            'frmB.vBuscaPrevia = chkVistaPrevia
            '#
            frmBd.Show vbModal
            Set frmBd = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
        End If
End Sub

Private Sub frmBd_Selecionado(CadenaDevuelta As String)
    If CadenaDevuelta <> "" Then
        txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
    End If
End Sub

