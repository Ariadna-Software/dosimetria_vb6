VERSION 5.00
Begin VB.Form frmErrores 
   Caption         =   "Se han producido errores"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "frmErrores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmErrores.frx":030A
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 1000 Then Me.Width = 1000
    If Me.Height < 1000 Then Me.Height = 1000
    Me.Command1.Top = Me.Height - Command1.Height - 500
    Me.Command1.Left = Me.Width - Command1.Width - 500
    Text1.Width = Me.Width - 300
    Text1.Height = Command1.Top - 300
End Sub
