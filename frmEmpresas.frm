VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frmEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7950
   Icon            =   "frmEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   7950
   Begin VB.Frame Frame4 
      Height          =   1665
      Left            =   210
      TabIndex        =   34
      Top             =   3870
      Width           =   7605
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1710
         MaxLength       =   40
         TabIndex        =   15
         Tag             =   "Persona Contacto|T|S|||empresas|pers_contacto|||"
         Text            =   "Text1"
         Top             =   750
         Width           =   5520
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "Mail|T|S|||empresas|mail_internet|||"
         Text            =   "Text1"
         Top             =   1170
         Width           =   5520
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   3810
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Fax|T|S|||empresas|fax|||"
         Text            =   "Text1"
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Telefono|T|S|||empresas|tel_contacto|||"
         Text            =   "Text1"
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label Label15 
         Caption         =   "Pers.Contacto:"
         Height          =   255
         Left            =   390
         TabIndex        =   38
         Top             =   750
         Width           =   1140
      End
      Begin VB.Label Label16 
         Caption         =   "Mail:"
         Height          =   255
         Left            =   390
         TabIndex        =   37
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   3090
         TabIndex        =   36
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   390
         TabIndex        =   35
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1545
      Left            =   210
      TabIndex        =   29
      Top             =   2340
      Width           =   7605
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2190
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "Distrito|T|S|||empresas|distrito|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Direccion|T|S|||empresas|direccion|||"
         Text            =   "Text1"
         Top             =   300
         Width           =   5520
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1740
         MaxLength       =   5
         TabIndex        =   9
         Tag             =   "C.Postal|T|N|||empresas|c_postal|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "Poblacion|T|S|||empresas|poblacion|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   3420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1140
         Width           =   3420
      End
      Begin VB.Image ImgPro 
         Height          =   240
         Left            =   1470
         MouseIcon       =   "frmEmpresas.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "frmEmpresas.frx":0E1C
         ToolTipText     =   "Buscar código postal"
         Top             =   735
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Código Postal:"
         Height          =   255
         Left            =   300
         TabIndex        =   33
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Domicilio:"
         Height          =   195
         Left            =   315
         TabIndex        =   32
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Poblacion:"
         Height          =   255
         Left            =   2940
         TabIndex        =   31
         Top             =   750
         Width           =   930
      End
      Begin VB.Label Label13 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   2940
         TabIndex        =   30
         Top             =   1155
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   210
      TabIndex        =   26
      Top             =   1230
      Width           =   7605
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5820
         TabIndex        =   7
         Tag             =   "Tipo Dosimetria|N|N|||empresas|c_tipo||N|"
         Text            =   "Combo2"
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "Migrado|T|S|||empresas|migrado|||"
         Text            =   "Text1"
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   1950
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Fecha Baja|F|S|||empresas|f_baja|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   570
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   420
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Fecha Alta|F|N|||empresas|f_alta|dd/mm/yyyy|S|"
         Text            =   "Text1"
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Dosimetria"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   5
         Left            =   5820
         TabIndex        =   40
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Migrado CSN"
         Height          =   255
         Left            =   3600
         TabIndex        =   39
         Top             =   300
         Width           =   1170
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   0
         Left            =   420
         MouseIcon       =   "frmEmpresas.frx":0F1E
         MousePointer    =   99  'Custom
         Picture         =   "frmEmpresas.frx":1070
         ToolTipText     =   "Seleccionar fecha"
         Top             =   300
         Width           =   240
      End
      Begin VB.Image ImgPpal 
         Height          =   240
         Index           =   1
         Left            =   1950
         MouseIcon       =   "frmEmpresas.frx":10FB
         MousePointer    =   99  'Custom
         Picture         =   "frmEmpresas.frx":124D
         ToolTipText     =   "Seleccionar fecha"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Baja"
         Height          =   255
         Left            =   2250
         TabIndex        =   28
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label21 
         Caption         =   "Fecha Alta"
         Height          =   255
         Left            =   750
         TabIndex        =   27
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   6660
      MaxLength       =   15
      TabIndex        =   3
      Tag             =   "CIF|T|S|||empresas|cif_nif|||"
      Text            =   "Text1"
      Top             =   870
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Nombre|T|N|||empresas|nom_comercial|||"
      Text            =   "Text1"
      Top             =   870
      Width           =   4710
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6690
      TabIndex        =   18
      Top             =   5730
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   0
      Left            =   420
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Codigo|T|N|||empresas|c_empresa||S|"
      Text            =   "Text1"
      Top             =   870
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6705
      TabIndex        =   19
      Top             =   5700
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   210
      TabIndex        =   20
      Top             =   5580
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5445
      TabIndex        =   17
      Top             =   5730
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   465
      Left            =   360
      Top             =   5580
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   820
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
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
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "L"
                  Object.Tag             =   "L"
                  Text            =   "Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "E"
                  Object.Tag             =   "E"
                  Text            =   "Etiquetas"
               EndProperty
            EndProperty
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
         Left            =   5250
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label10 
      Caption         =   "CIF/NIF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6690
      TabIndex        =   25
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre Comercial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1950
      TabIndex        =   24
      Top             =   570
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   420
      TabIndex        =   22
      Top             =   570
      Width           =   735
   End
End
Attribute VB_Name = "frmEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmPro As frmProvincias
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private ValorAnterior As String

Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
   If Modo = 2 Or Modo = 0 Then
        Select Case KeyCode
            Case vbESC '27
                If Modo = 0 Then
                    Toolbar1_ButtonClick Toolbar1.Buttons(12)
                Else
                    PonerModo 0
                End If
            Case vbAnterior '33
                If Modo = 2 Then Desplazamiento (1)
            Case vbSiguiente '34
                If Modo = 2 Then Desplazamiento (2)
            Case vbPrimero  ' 36 ' inicio
                If Modo = 2 Then Desplazamiento (0)
            Case vbUltimo '35 ' fin
                If Modo = 2 Then Desplazamiento (3)
           Case vbBuscar
                Toolbar1_ButtonClick Toolbar1.Buttons(1)
           Case vbVerTodos
                Toolbar1_ButtonClick Toolbar1.Buttons(2)
           Case vbAñadir
                Toolbar1_ButtonClick Toolbar1.Buttons(6)
           Case vbModificar
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(7)
            Case vbEliminar
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(8)
            Case vbImprimir
                 If Modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(11)
            Case vbSalir
                  Toolbar1_ButtonClick Toolbar1.Buttons(12)
        End Select
   End If


End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                If SituarData1 Then
                    PonerModo 2
                    PonerCampos
                Else
                    LimpiarCampos
                    PonerModo 0
                End If
            End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me, 1) Then
'                    lblIndicador.Caption = ""
                    If SituarData1 Then
                        PonerModo 2
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    Case 1
        HacerBusqueda
    End Select
    

Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation, "¡Error!"
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        Case 4
            PonerModo 2
            PonerCampos
        End Select
End Sub

' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim sql As String
    On Error GoTo ESituarData1
            'Actualizamos el recordset
            Data1.Refresh
            '********* canviar la clau primaria codsocio per la que siga *********
            'El sql para que se situe en el registro en especial es el siguiente
            sql = " c_empresa = '" & Trim(Text1(0).Text) & "'"
            '*****************************************************************
            Data1.Recordset.Find sql
            If Data1.Recordset.EOF Then GoTo ESituarData1
            SituarData1 = True
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
'        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    DespalzamientoVisible False
    '###A mano
    PonerFoco Text1(0)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
'        lblIndicador.Caption = "BUSCAR"
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Text1(kCampo).SetFocus
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
    Case 1
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 3
        Data1.Recordset.MoveLast
End Select
PonerCampos
lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
'    lblIndicador.Caption = "Modificar"
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer
    Dim sql As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    '******* canviar el mensage i la cadena *********************
    Cad = "Seguro que desea eliminar la empresa:"
    Cad = Cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(4)
    Cad = Cad & vbCrLf & "Fecha: " & Text1(11).Text
    '**********************************************************
    I = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") 'VRS:1.0.1(11)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        sql = "delete from empresas where c_empresa = '" & Data1.Recordset.Fields(0) & "' and "
        sql = sql & "f_alta = '" & Format(Text1(11).Text, FormatoFecha) & "'"
        
        
        conn.Execute sql
        
        NumRegElim = Data1.Recordset.AbsolutePosition
'        DataGrid1.Enabled = False
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            PonerModo 0
            Else
                Data1.Recordset.MoveFirst
                NumRegElim = NumRegElim - 1
                If NumRegElim > 1 Then
                    For I = 1 To NumRegElim - 1
                        Data1.Recordset.MoveNext
                    Next I
                End If
                PonerCampos
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Empresas"
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

If Data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation, "¡Atención!"
    Exit Sub
End If

Cad = ""
I = 0
Do
    J = I + 1
    I = InStr(J, DatosADevolverBusqueda, "|")
    If I > 0 Then
        Aux = Mid(DatosADevolverBusqueda, J, I - J)
        J = Val(Aux)
        Cad = Cad & Text1(J).Text & "|"
    End If
Loop Until I = 0
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
      
End Sub

Private Sub Form_Load()
Dim I As Integer
    
    Me.Top = 0
    Me.Left = 0

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

    LimpiarCampos
    
    ' Usuario restringido a consultas.
    If vUsu.NivelUsu < 1 Then
      Toolbar1.Buttons(6).Visible = False
      Toolbar1.Buttons(7).Visible = False
      Toolbar1.Buttons(8).Visible = False
    End If
    
    '***** canviar el nom de la taula i el ORDER BY ********
    NombreTabla = "empresas"
    Ordenacion = " ORDER BY c_empresa"
    '******************************************************+
        
'    PonerOpcionesMenu
    
    'Para todos
'    Data1.UserName = vUsu.Login
'    Me.Data1.password = vUsu.Passwd
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
'    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        '### A mano
        Text1(0).BackColor = vbYellow
    End If
    
    CargarCombo
End Sub

Private Sub LimpiarCampos()
    Limpiar Me
    
    Combo3.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
'        Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    If Modo = 0 Or Modo = 2 Then Exit Sub
    Select Case Index
       Case 0
            f = Now
            If Text1(11).Text <> "" Then
                If IsDate(Text1(11).Text) Then f = Text1(11).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(11).Text = frmC.fecha
                mTag.DarFormato Text1(11)
            End If
            Set frmC = Nothing
       Case 1
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(12).Text <> "" Then
                If IsDate(Text1(12).Text) Then f = Text1(12).Text
            End If
            Set frmC = New frmCal
            frmC.fecha = f
            frmC.Show vbModal
            If Modo = 3 Or Modo = 4 Or Modo = 1 Then
                Text1(12).Text = frmC.fecha
                mTag.DarFormato Text1(12)
            End If
            Set frmC = Nothing
   
   End Select
End Sub

Private Sub Imgpro_Click()
    Set frmPro = New frmProvincias
    frmPro.DatosADevolverBusqueda = "0|1|"
    frmPro.Show
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
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
    Else
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim I As Integer
    Dim sql As String
    Dim mTag As CTag
    Dim valor As Currency
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" Then Exit Sub
    
    If Modo = 1 And ConCaracteresBusqueda(Text1(Index).Text) Then Exit Sub
    
    Select Case Index
        Case 0, 3, 5, 6, 7, 8, 10, 13
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            If Modo = 1 Then Exit Sub
            If Modo <> 1 And Text1(Index).Text = "" Then
                PonerFoco Text1(Index)
                MsgBox "Este campo requiere un valor", vbExclamation, "¡Error!"
            End If
            If Index <> 10 Then
                Text1(Index).Text = Format(Text1(Index).Text, ">")
            End If
            If Index = 3 Then
                If Text1(Index).Text <> "" And Modo <> 1 Then
                    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
                    If Text2(0).Text = "" Then
                        MsgBox "Código de provincia no existe. Reintroduzca.", vbExclamation, "¡Error!"
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
              
        Case 11, 12 ' campos de fechas
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation, "¡Error!"
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
            End If
              
    End Select
    Text1(Index).Text = Format(Text1(Index).Text, ">")
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String
    CadB = ObtenerBusqueda(Me)
    If CadB = "" Then
        MsgBox vbCrLf & "  Debe introducir alguna condición de búsqueda. " & vbCrLf, vbExclamation, "¡Error!"
        PonerModo 0
        Exit Sub
    End If
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
        Else
            'Se muestran en el mismo form
            If CadB <> "" Then
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 20, "Código")
        Cad = Cad & ParaGrid(Text1(1), 60, "Nombre")
        Cad = Cad & ParaGrid(Text1(6), 20, "Cif")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Empresas"
            frmB.vSelElem = 0
            frmB.vConexionGrid = 1
            'frmB.vBuscaPrevia = chkVistaPrevia
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                Text1(kCampo).SetFocus
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation, "¡Atención!"
    Screen.MousePointer = vbDefault
    PonerModo 0
    Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
End If

Screen.MousePointer = vbDefault
Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    Dim I As Integer
    Dim mTag As CTag
    Dim sql As String
    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
    
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub
'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Dim b As Boolean
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For I = 0 To Text1.Count - 1
            Text1(I).BackColor = vbWhite
        Next I
    End If
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    If Modo = 0 Then LimpiarCampos
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    'Modificar
    Toolbar1.Buttons(7).Enabled = b 'And vUsu.NivelUsu <= 2
 '   mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b 'And vUsu.NivelUsu <= 2
    Toolbar1.Buttons(11).Enabled = (Modo = 2)

'    mnModificar.Enabled = b
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    Else
        cmdRegresar.Visible = False
    End If
    
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = b Or Modo = 1
    cmdCancelar.Visible = b Or Modo = 1
'    mnOpciones.Enabled = Not b
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(6).Enabled = Not b And Modo <> 1 'And vUsu.NivelUsu <= 2
    Toolbar1.Buttons(1).Enabled = Not b And Modo <> 1
    Toolbar1.Buttons(2).Enabled = Not b And Modo <> 1
    
    b = (Modo = 2) Or Modo = 0
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = b
            Text1(I).BackColor = vbWhite
    Next I
    Combo3.Enabled = Not b
    
    If Modo = 3 Then ValoresPorDefecto
    
    ImgPro.Enabled = Not b
    
    PonerFoco chkVistaPrevia
    
    
End Sub

Private Function DatosOk() As Boolean
Dim rs As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim Datos As String
Dim Mens As String

    b = CompForm(Me)
    
    If (b = True) And ((Modo = 3) Or (Modo = 4)) Then
'        For I = 0 To Text1.Count - 1
'             If InStr(1, Text1(I).Text, "'") > 0 Then
'                MsgBox "No puede introducir el carácter ' en ese campo.", vbExclamation, "¡Error!"
'                DatosOk = False
'                Exit Function
'             End If
'        Next I
        Datos = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
        If Datos = "" Then
            MsgBox "No existe la provincia.", vbExclamation, "¡Error!"
            DatosOk = False
            Exit Function
        End If
        
        If Modo = 3 Then
             Datos = DevuelveDesdeBD(1, "c_empresa", "empresas", "c_empresa|f_alta|", Text1(0).Text & "|" & Text1(11).Text, "T|F|", 2)
             If Datos <> "" Then
                MsgBox "Ya existe el código de empresa : " & Text1(0).Text, vbExclamation, "¡Error!"
                DatosOk = False
                Exit Function
             End If
        End If
        
        ' el tipo de dosimetria ha de ser igual en instalacion y empresa
        If Modo = 4 Then
            If HayInstalacionesDeOtroTipo(Text1(0).Text, Combo3.ListIndex) Then
                If MsgBox("La empresa tiene alguna instalacion que no pertenece al mismo tipo de dosimetria. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "¡Atención!") = vbNo Then
                    DatosOk = False
                    Exit Function
                End If
            End If
        End If
        
    End If
    DatosOk = b
End Function


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
        Case 12
            mnSalir_Click
        Case 14 To 17
            Desplazamiento (Button.Index - 14)
        Case 11
'            Screen.MousePointer = vbHourglass
'            FrmListado.Opcion = 1 'Listado de empresas
'            FrmListado.Show
        Case Else
    
    End Select
End Sub

Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub ValoresPorDefecto()
    Text1(3).Text = "46"
    Text2(0).Text = DevuelveDesdeBD(1, "descripcion", "provincias", "c_postal|", Text1(3).Text & "|", "T|", 1)
    Text1(10).Text = "mail"
    Text1(11).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu
        Case "Listado"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 1 'Listado de empresas
            FrmListado.Show
        Case "Etiquetas"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 24 'Listado de etiquetas
            FrmListado.Show
    End Select
End Sub

Private Sub CargarCombo()
'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Si, 1-No
    Combo3.Clear
    Combo3.AddItem "Personal"
    Combo3.ItemData(Combo3.NewIndex) = 0
    
    Combo3.AddItem "Area"
    Combo3.ItemData(Combo3.NewIndex) = 1
    
    Combo3.AddItem "Personal/Area"
    Combo3.ItemData(Combo3.NewIndex) = 2
    
End Sub

Private Function HayInstalacionesDeOtroTipo(Emp As String, Tipo As Byte)
Dim rs As ADODB.Recordset
Dim sql As String
Dim b As Boolean

    b = True
        
    sql = "select c_tipo from instalaciones where c_empresa = '" & Trim(Emp) & "' and c_tipo <> " & Tipo
    
    Set rs = New ADODB.Recordset
    
    rs.Open sql, conn, , , adCmdText
    
    If rs.EOF Then
        b = False
    Else
        If IsNull(rs.Fields(0)) Then
            b = False
        End If
    End If
    rs.Close
    Set rs = Nothing
    HayInstalacionesDeOtroTipo = b

End Function
