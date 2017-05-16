VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      _Version        =   524288
      _ExtentX        =   8281
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2003
      Month           =   7
      Day             =   24
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   1605
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCal 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3165
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Public fecha As Date
Private Dia As Integer
Private Mes As Integer
Private Año As Integer
Public Event Selec(vFecha As Date)

Private Sub Calendar1_Click()
    fecha = Calendar1.Value
End Sub

Private Sub Calendar1_DblClick()
    Calendar1_Click
    cmdCal_Click (0)
End Sub

Private Sub cmdCal_Click(Index As Integer)
    Select Case Index
        Case 0
            '-- Aceptar
            RaiseEvent Selec(fecha)
        Case 1
    End Select
    Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dia = Val(Format(fecha, "dd"))
    Mes = Val(Format(fecha, "mm"))
    Año = Val(Format(fecha, "yyyy"))
    Calendar1.Day = Dia
    Calendar1.Month = Mes
    Calendar1.Year = Año
End Sub
