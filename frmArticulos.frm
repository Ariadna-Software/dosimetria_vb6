VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmArticulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Articulos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9420
   Icon            =   "frmArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   9420
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   5160
      MaxLength       =   13
      TabIndex        =   3
      Tag             =   "Codigo EAN|T|N|||sartic|codigean|||"
      Text            =   "1234567890123"
      Top             =   765
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1530
      MaxLength       =   35
      TabIndex        =   2
      Tag             =   "Descripcion|T|N|||sartic|nomartic|||"
      Text            =   "12345678901234567890123456789012345"
      Top             =   765
      Width           =   3315
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8010
      TabIndex        =   34
      Top             =   5550
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Index           =   0
      Left            =   495
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Codigo|N|N|0|999999|sartic|codartic|000000|S|"
      Text            =   "Text1"
      Top             =   765
      Width           =   750
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4200
      Left            =   225
      TabIndex        =   39
      Top             =   1215
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   7408
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Datos &Generales"
      TabPicture(0)   =   "frmArticulos.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImgTun"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ImgFam"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Datos Almacen"
      TabPicture(1)   =   "frmArticulos.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Text1(10)"
      Tab(1).Control(3)=   "Text1(11)"
      Tab(1).Control(4)=   "Text1(12)"
      Tab(1).Control(5)=   "Text1(13)"
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(8)=   "Label8"
      Tab(1).Control(9)=   "Label9"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Datos &Fitosanitarios"
      TabPicture(2)   =   "frmArticulos.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Txtaux(2)"
      Tab(2).Control(1)=   "Txtaux(1)"
      Tab(2).Control(2)=   "Txtaux(0)"
      Tab(2).Control(3)=   "Text1(19)"
      Tab(2).Control(4)=   "Text1(18)"
      Tab(2).Control(5)=   "Text1(23)"
      Tab(2).Control(6)=   "DataGrid1"
      Tab(2).Control(7)=   "Label18"
      Tab(2).Control(8)=   "Label17"
      Tab(2).Control(9)=   "Label23"
      Tab(2).ControlCount=   10
      Begin VB.TextBox Txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   -69105
         MaxLength       =   15
         TabIndex        =   30
         Top             =   3825
         Width           =   2355
      End
      Begin VB.TextBox Txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   -71805
         MaxLength       =   15
         TabIndex        =   29
         Top             =   3825
         Width           =   2625
      End
      Begin VB.TextBox Txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   -74415
         MaxLength       =   15
         TabIndex        =   28
         Top             =   3825
         Width           =   2580
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   -2430
         Top             =   1665
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
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
         UserName        =   "root"
         Password        =   "aritel"
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
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   19
         Left            =   -70365
         MaxLength       =   3
         TabIndex        =   26
         Tag             =   "Categoria|T|S|||sartic|categoria|||"
         Text            =   "Text1"
         Top             =   765
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   -73290
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Nro.Registro Fitosanitario|T|S|||sartic|regfitosanitario|||"
         Text            =   "123456789012345"
         Top             =   795
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   23
         Left            =   -67410
         MaxLength       =   5
         TabIndex        =   27
         Tag             =   "Nro.ADRa|T|S|||sartic|numeroadr|||"
         Text            =   "Text1"
         Top             =   765
         Width           =   570
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1845
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Familia|N|N|0|999|sartic|codfamia|000||"
         Text            =   "Text1"
         Top             =   900
         Width           =   525
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text2"
         Top             =   1350
         Width           =   3915
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Tipo Unidad|N|N|0|99|sartic|codtipun|00||"
         Text            =   "Text1"
         Top             =   1350
         Width           =   525
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "123456789012345678901234567890"
         Top             =   915
         Width           =   3915
      End
      Begin VB.Frame Frame2 
         Caption         =   "Precios"
         Height          =   1860
         Left            =   180
         TabIndex        =   52
         Top             =   2025
         Width           =   8460
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   22
            Left            =   5895
            MaxLength       =   40
            TabIndex        =   7
            Tag             =   "Fec.Cambio Precio|F|N|||sartic|feccambioprec|dd/mm/yyyy||"
            Text            =   "Text1"
            Top             =   450
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   7560
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "%Aumento Socio|N|N|0.00|999.99|sartic|aumsocio|##0.00||"
            Text            =   "123456"
            Top             =   900
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   7560
            MaxLength       =   6
            TabIndex        =   13
            Tag             =   "%Aumento cliente|N|N|0.00|999.99|sartic|aumcliente|##0.00||"
            Text            =   "Text1"
            Top             =   1380
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2250
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   1665
            MaxLength       =   40
            TabIndex        =   6
            Tag             =   "Codigo Iva|N|N|1|9|sartic|codi_iva|0||"
            Text            =   "Text1"
            Top             =   420
            Width           =   480
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   5085
            MaxLength       =   15
            TabIndex        =   12
            Tag             =   "Pr.IVA Cliente|N|N|0|99999999.9999|sartic|preciov4|##,###,##0.0000||"
            Text            =   "Text1"
            Top             =   1380
            Width           =   1920
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   5085
            MaxLength       =   15
            TabIndex        =   9
            Tag             =   "Pr.IVA Socio|N|N|0|99999999.9999|sartic|preciov3|##,###,##0.0000||"
            Text            =   "Text1"
            Top             =   885
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1665
            MaxLength       =   15
            TabIndex        =   11
            Tag             =   "Pr.Base Cliente|N|N|0|99999999.9999|sartic|preciov2|##,###,##0.0000||"
            Text            =   "Text1"
            Top             =   1365
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1665
            MaxLength       =   15
            TabIndex        =   8
            Tag             =   "Pr.Base Socio|N|N|0|99999999.9999|sartic|preciov1|##,###,##0.0000||"
            Text            =   "Text1"
            Top             =   885
            Width           =   1965
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   2
            Left            =   5535
            Picture         =   "frmArticulos.frx":0D1E
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label21 
            Caption         =   "Fecha Cambio Precios:"
            Height          =   255
            Left            =   3735
            TabIndex        =   65
            Top             =   465
            Width           =   1710
         End
         Begin VB.Label Label20 
            Caption         =   "%Aum:"
            Height          =   255
            Left            =   7065
            TabIndex        =   64
            Top             =   900
            Width           =   570
         End
         Begin VB.Label Label19 
            Caption         =   "%Aum:"
            Height          =   255
            Left            =   7065
            TabIndex        =   63
            Top             =   1395
            Width           =   570
         End
         Begin VB.Image ImgIva 
            Height          =   240
            Left            =   1215
            MouseIcon       =   "frmArticulos.frx":0E20
            MousePointer    =   99  'Custom
            Picture         =   "frmArticulos.frx":0F72
            ToolTipText     =   "Buscar socio"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label16 
            Caption         =   "Código IVA:"
            Height          =   255
            Left            =   210
            TabIndex        =   59
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Precio IVA Cliente:"
            Height          =   255
            Left            =   3735
            TabIndex        =   56
            Top             =   1395
            Width           =   1470
         End
         Begin VB.Label Label14 
            Caption         =   "Precio IVA Socio:"
            Height          =   255
            Left            =   3735
            TabIndex        =   55
            Top             =   900
            Width           =   1350
         End
         Begin VB.Label Label13 
            Caption         =   "Precio Base Cliente:"
            Height          =   255
            Left            =   225
            TabIndex        =   54
            Top             =   1395
            Width           =   1560
         End
         Begin VB.Label Label12 
            Caption         =   "Precio Base Socio:"
            Height          =   255
            Left            =   225
            TabIndex        =   53
            Top             =   915
            Width           =   1470
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Inventario"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   47
         Top             =   2370
         Width           =   3885
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Index           =   25
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   20
            Tag             =   "Hora Inventario|H|S|||sartic|horainve|hh:mm||"
            Text            =   "commor"
            Top             =   1020
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   15
            Left            =   120
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "fechainv|F|S|||sartic|fechainv|dd/mm/yyyy||"
            Text            =   "commor"
            Top             =   1020
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   18
            Tag             =   "Stock Inventario|N|N|-99999999.99|99999999.99|sartic|stockinv|##,###,##0.00||"
            Text            =   "Text1"
            Top             =   345
            Width           =   2055
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmArticulos.frx":1074
            Left            =   2340
            List            =   "frmArticulos.frx":1076
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Tag             =   "Sit.Inventario|N|N|0|9|sartic|statusin|||"
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Hora"
            Height          =   195
            Index           =   2
            Left            =   1350
            TabIndex        =   71
            Top             =   780
            Width           =   450
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   0
            Left            =   720
            Picture         =   "frmArticulos.frx":1078
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label22 
            Caption         =   "Situación:"
            Height          =   255
            Left            =   2370
            TabIndex        =   66
            Top             =   780
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   49
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label10 
            Caption         =   "Stock :"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   345
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Precios Valoración"
         Height          =   1575
         Left            =   -70380
         TabIndex        =   46
         Top             =   2370
         Width           =   4065
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "Precio Ult.Compra|N|N|0|99999999.9999|sartic|preciouc|##,###,##0.0000||"
            Text            =   "Text1"
            Top             =   690
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   17
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   24
            Tag             =   "Fec.Ult.Compra|F|S|||sartic|ultfecco|dd/mm/yyyy||"
            Text            =   "commor"
            Top             =   1095
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   22
            Tag             =   "Precio Med.Ponderado|N|N|||sartic|preciomp|##,###,##0.0000||"
            Text            =   "Text1"
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label Label24 
            Caption         =   "Última Compra:"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   750
            Width           =   1335
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   1
            Left            =   1350
            Picture         =   "frmArticulos.frx":117A
            Top             =   1140
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fec.Ult.Compra"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   1170
            Width           =   1140
         End
         Begin VB.Label Label11 
            Caption         =   "Medio Ponderado:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   330
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   -73425
         MaxLength       =   40
         TabIndex        =   14
         Tag             =   "Stock Minimo|N|N|0|99999999.99|sartic|stockmin|##,###,##0.00||"
         Text            =   "Text1"
         Top             =   1020
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   -73425
         MaxLength       =   40
         TabIndex        =   15
         Tag             =   "Punto Pedido|N|N|0|99999999.99|sartic|puntoped|##,###,##0.00||"
         Text            =   "Text1"
         Top             =   1545
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   -68700
         MaxLength       =   40
         TabIndex        =   16
         Tag             =   "Stock Maximo|N|N|0|99999999.99|sartic|stockmax|##,###,##0.00||"
         Text            =   "Text1"
         Top             =   1050
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   -68670
         MaxLength       =   40
         TabIndex        =   17
         Tag             =   "Stock Actual|N|N|-99999999.99|99999999.99|sartic|stockact|##,###,##0.00||"
         Text            =   "Text1"
         Top             =   1530
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmArticulos.frx":127C
         Height          =   2880
         Left            =   -74775
         TabIndex        =   31
         Top             =   1170
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   5080
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
      Begin VB.Label Label18 
         Caption         =   "Categoria:"
         Height          =   255
         Left            =   -71220
         TabIndex        =   69
         Top             =   825
         Width           =   1020
      End
      Begin VB.Label Label17 
         Caption         =   "Nro. Registro:"
         Height          =   255
         Left            =   -74595
         TabIndex        =   68
         Top             =   825
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Número ADR:"
         Height          =   255
         Left            =   -68550
         TabIndex        =   67
         Top             =   825
         Width           =   1020
      End
      Begin VB.Label Label4 
         Caption         =   "Stock Minimo:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   45
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Punto Pedido:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   44
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Stock Máximo:"
         Height          =   255
         Left            =   -70170
         TabIndex        =   43
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Stock Actual:"
         Height          =   255
         Left            =   -70170
         TabIndex        =   42
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Familia:"
         Height          =   255
         Left            =   540
         TabIndex        =   41
         Top             =   930
         Width           =   615
      End
      Begin VB.Image ImgFam 
         Height          =   240
         Left            =   1395
         MouseIcon       =   "frmArticulos.frx":1291
         MousePointer    =   99  'Custom
         Picture         =   "frmArticulos.frx":13E3
         ToolTipText     =   "Buscar socio"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image ImgTun 
         Height          =   240
         Left            =   1395
         MouseIcon       =   "frmArticulos.frx":14E5
         MousePointer    =   99  'Custom
         Picture         =   "frmArticulos.frx":1637
         ToolTipText     =   "Buscar socio"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Unidad:"
         Height          =   255
         Left            =   540
         TabIndex        =   40
         Top             =   1350
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8070
      TabIndex        =   33
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   315
      TabIndex        =   35
      Top             =   5490
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6750
      TabIndex        =   32
      Top             =   5550
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   405
      Top             =   5850
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      TabIndex        =   38
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
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
            Object.ToolTipText     =   "Modificar Lineas"
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
         Left            =   5220
         TabIndex        =   0
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Codigo EAN"
      Height          =   255
      Left            =   5220
      TabIndex        =   62
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   1575
      TabIndex        =   61
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   495
      TabIndex        =   37
      Top             =   540
      Width           =   735
   End
End
Attribute VB_Name = "frmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmBc As frmBuscaGrid 'Conta
Attribute frmBc.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1

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
Private modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private ValorAnterior As String
Private IndiceErroneo As Byte
Private conexion As Integer
Private familia As Integer
Dim SumaLinea As Currency
Dim i As Integer
Dim Numlinea As Integer
Dim Aux As Currency
Dim PulsadoSalir As Boolean
Private ModificandoLineas As Byte
Dim AntiguoText1 As String

' campo que indica si la familia es fitosanitaria
' si lo es: obligamos a introducir los campos de fitos.
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim FechaApertura As Date

Private Sub chkVistaPrevia_KeyDown(KeyCode As Integer, Shift As Integer)
    AsignarTeclasFuncion KeyCode
End Sub


Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim i As Integer
    Dim v_aux As Integer
    Dim SQL As String
    
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me, 1) Then
                'MsgBox "Registro insertado.", vbInformation
                If EsFitosanitario(Text1(3).Text) Then
                    If SituarData1 Then
                        PonerModo 5
                        'Haremos como si pulsamo el boton de insertar nuevas lineas
                        'Ponemos el importe en AUX
                        Aux = ImporteFormateado(Text1(13).Text)
                        cmdCancelar.Caption = "&Cabecera"
                        ModificandoLineas = 0
                        'Bloqueamos pa' k nadie entre
                        BloqueaRegistroForm Me
                        AnyadirLinea True, False
                    Else
                        SQL = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                        SQL = SQL & vbCrLf & " CLAVE: FrmFacturas. cmdAceptar. SituarData1"
                        MsgBox SQL, vbCritical
                        Exit Sub
                    End If
                 Else
                    CargaGrid False
                    PonerModo 0
'                    lblIndicador.Caption = ""
                 End If
                 'VRS:1.0.5(14)
                 'hemos de insertar el movimiento de cierre que es desde donde se inicia
                 InsertarMovimientoAlmacen
                 
            End If
        End If
    
    Case 4
        'Modificar
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If ModificaDesdeFormulario(Me, 1) Then
                If ComprobarStocksLineas Then
                    DesBloqueaRegistroForm Text1(0)
'                    lblIndicador.Caption = ""
                    If SituarData1 Then
                        PonerModo 2
                        PonerCampos
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                Else
                    ' no coinciden stocks
                    DesBloqueaRegistroForm Text1(0)
                    mnLineas_Click
                End If
            End If
'            SSTab1.Tab = 0
'        Else
'                Text1(IndiceErroneo).SetFocus
'                If IndiceErroneo < 10 Then SSTab1.Tab = 0
        End If


    Case 5
        Cad = AuxOK
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
        Else
            'Insertaremos, o modificaremos
            If InsertarModificar Then
                'Reestablecemos los campos
                'y ponemos el grid
                cmdAceptar.Visible = False
                DataGrid1.AllowAddNew = False
                CargaGrid True
                lblIndicador.Caption = "Lineas detalle"

                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    ModificandoLineas = 0
'                    cmdAceptar.Visible = True
                    cmdCancelar.Caption = "&Cabecera"

                    For v_aux = 0 To 2
                        Txtaux(v_aux).Text = ""
                    Next v_aux

                    AnyadirLinea True, False

                    PonerFoco Txtaux(0)
                Else
                    Txtaux(0).Enabled = True
                    Txtaux(1).Enabled = True

                    ModificandoLineas = 0
                    CamposAux False, 0, False
                    cmdCancelar.Caption = "&Cabecera"
                End If
            End If
        End If

    Case 1
        HacerBusqueda
    End Select
    

Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    Select Case modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        
        Case 4
            'Modificar
'            lblIndicador.Caption = ""
            DesBloqueaRegistroForm Text1(0)
            PonerModo 2
            PonerCampos
        
        Case 5
            CamposAux False, 0, False
            lblIndicador.Caption = "Lineas detalle"
            'Si esta insertando/modificando lineas haremos unas cosas u otras
            DataGrid1.Enabled = True
            If ModificandoLineas = 0 Then
                'NUEVO
                AntiguoText1 = ""
                If adodc1.Recordset.EOF Then
                    AntiguoText1 = "El articulo no tiene cantidad en stock. ¿SEGURO que desea salir?"
                    If MsgBox(AntiguoText1, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then  'VRS:1.0.1(11)
                        AntiguoText1 = ""
                    Else
                        'Para k no muestre el siguiente punto de error
                        AntiguoText1 = "###"
                    End If
                End If
                'Else
                    'Comprobamos que el total de factura es el de suma
                   ObtenerSigueinteNumeroLinea
                   If Aux <> 0 Then AntiguoText1 = "La suma de lineas no coincide con la cantidad de stock: " & Format(Aux, FormatoImporte)
                'End If
                If AntiguoText1 <> "" Then
                    If AntiguoText1 <> "###" Then MsgBox AntiguoText1, vbExclamation
                    Exit Sub
                End If
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                CargaGrid True
                If Not SituarData1 Then
                    PonerModo 0
                Else
                    PonerCampos
                    PonerModo 2
                End If

                Txtaux(0).Enabled = True
                Txtaux(1).Enabled = True

                DesBloqueaRegistroForm Me.Text1(0)
            Else
                If ModificandoLineas = 1 Then
                     DataGrid1.AllowAddNew = False
                     If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
                     DataGrid1.Refresh
                End If

                cmdAceptar.Visible = False
                cmdCancelar.Caption = "&Cabecera"
                ModificandoLineas = 0
            End If
        
        End Select
        
'    SSTab1.Tab = 0
End Sub

' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim SQL As String
    On Error GoTo ESituarData1
            'Actualizamos el recordset
            Data1.Refresh
            '********* canviar la clau primaria codsocio per la que siga *********
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = "codartic = " & Text1(0).Text & ""
            '*****************************************************************
            Data1.Recordset.Find SQL
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
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
'    lblIndicador.Caption = "INSERTANDO"
    SugerirCodigoSiguiente
    '###A mano
    'precios
    For i = 6 To 9
        Text1(i).Text = Format(0, "###,###,##0.0000")
    Next i
    
    Text1(20).Text = Format(0, "##0.00")
    Text1(21).Text = Format(0, "##0.00")
    'cantidades almacen
    For i = 10 To 13
        Text1(i).Text = Format(0, "###,###,##0.00")
    Next i
    ' inventario
    Text1(14).Text = Format(0, "###,###,##0.000")
    Text1(15).Text = ""
    Combo2.Text = "No"
    'compra
    Text1(16).Text = Format(0, "###,###,##0.0000")
    Text1(24).Text = Format(0, "###,###,##0.0000")
    Text1(17).Text = ""
    
    Text1(22).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(0)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If modo <> 1 Then
        LimpiarCampos
'        lblIndicador.Caption = "BUSCAR"
        PonerModo 1
        CargaGrid False
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarillo
         Select Case SSTab1.Tab
             Case 0
                 PonerFoco Text1(0)
                 Text1(0).BackColor = vbYellow
             Case 1
                 PonerFoco Text1(10)
                 Text1(10).BackColor = vbYellow
             Case 2
                 PonerFoco Text1(18)
                 Text1(18).BackColor = vbYellow
        End Select
        
'        Text1(0).SetFocus
'        Text1(0).BackColor = vbYellow
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
    CargaGrid False
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
    If Not BloqueaRegistroForm(Me) Then Exit Sub
   
    If Combo2.ListIndex = 0 Then
        MsgBox "Artículo inventariandose. No se puede modificar.", vbExclamation
        DesBloqueaRegistroForm Text1(0)
        Exit Sub
    End If
   
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
'    lblIndicador.Caption = "Modificar"
    cmdCancelar.Caption = "&Cancelar"
    
    'VRS:1.0.1(5)
    familia = Text1(3).Text
    
    Select Case SSTab1.Tab
        Case 0
            PonerFoco Text1(1)
        Case 1
            PonerFoco Text1(10)
        Case 2
            PonerFoco Text1(18)
   End Select

End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim i As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    ' si el articulo está inventariandose no se puede eliminar
    If Combo2.ListIndex = 0 Then
        MsgBox "Artículo inventariandose. No se puede eliminar.", vbExclamation
        Exit Sub
    End If

    
    '******* canviar el mensage i la cadena *********************
    Cad = "Seguro que desea eliminar el articulo:"
    Cad = Cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    '**********************************************************
    i = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) 'VRS:1.0.1(11)
    
   'Borramos
    If i <> vbYes Then
        DesbloqueaRegistroForm1 Me 'Me.Text1(4)
        Exit Sub
    End If
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    If Not Eliminar Then Exit Sub
   
    NumRegElim = Data1.Recordset.AbsolutePosition
    DataGrid1.Enabled = False
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid False
        PonerModo 0
        Else
            If NumRegElim > Data1.Recordset.RecordCount Then
                Data1.Recordset.MoveLast
            Else
                Data1.Recordset.MoveFirst
                Data1.Recordset.Move NumRegElim - 1
            End If
            PonerCampos
            DataGrid1.Enabled = True
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Artículo"
End Sub

Private Sub cmdCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
    If modo = 5 Then
        Select Case KeyCode
           Case vbESC
                cmdCancelar_Click
           Case vbAñadir
                 Toolbar1_ButtonClick Toolbar1.Buttons(6)
           Case vbModificar
                 Toolbar1_ButtonClick Toolbar1.Buttons(7)
           Case vbEliminar
                 Toolbar1_ButtonClick Toolbar1.Buttons(8)
           Case vbSalir
                 cmdCancelar_Click
        End Select
   End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

If Data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

Cad = ""
i = 0
Do
    J = i + 1
    i = InStr(J, DatosADevolverBusqueda, "|")
    If i > 0 Then
        Aux = Mid(DatosADevolverBusqueda, J, i - J)
        J = Val(Aux)
        Cad = Cad & Text1(J).Text & "|"
    End If
Loop Until i = 0
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If modo = 5 Then
        Select Case KeyCode
           Case vbESC
                cmdCancelar_Click
           Case vbAñadir
                 Toolbar1_ButtonClick Toolbar1.Buttons(6)
           Case vbModificar
                 Toolbar1_ButtonClick Toolbar1.Buttons(7)
           Case vbEliminar
                 Toolbar1_ButtonClick Toolbar1.Buttons(8)
           Case vbSalir
                 cmdCancelar_Click
        End Select
   End If

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    CargaGrid (modo = 2)
End Sub

Private Sub Form_Load()
Dim i As Integer

'    If vParam.HayContabilidad Then
'        If AbrirConexionConta(vParam.NombreHost, vParam.NombreUsuario) = False Then
'            MsgBox "La aplicación no puede continuar sin acceso a la Contabilidad. ", vbCritical
'            End
'        End If
'    End If


    Me.Top = 0
    Me.Left = 0
    PulsadoSalir = False

      ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    
    'Los campos auxiliares
    CamposAux False, 0, True

    LimpiarCampos
    
    'Como son cuentas, como mucho seran
'    For i = 4 To 8
'        Text1(i).MaxLength = vEmpresa.DigitosUltimoNivel
'    Next i
    
    '***** canviar el nom de la taula i el ORDER BY ********
    NombreTabla = "sartic"
    Ordenacion = " ORDER BY sartic.codartic"
    '******************************************************+
        
    PonerOpcionesMenu
    
    'Para todos
'    Data1.UserName = vUsu.Login
'    Me.Data1.password = vUsu.Passwd
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    End If
    
    CargarCombo
    
    SSTab1.Tab = 0
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
'    lblIndicador.Caption = ""
    
    'Aqui va el especifico de cada form es
    '### a mano
    Combo2.ListIndex = -1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If

    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
'    If vParam.HayContabilidad Then ConnConta.Close
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        'If CadB <> "" Then CadB = CadB & " AND "
        'CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmBc_Selecionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text1(5).Text = Format(Text1(5).Text, "00")
        Text2(2).Text = RecuperaValor(CadenaSeleccion, 3)
        Text2(2).Text = Format(Text2(2).Text, "#0.00")
        RecalculoIva (6)
        RecalculoIva (7)
    End If
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 2)
    Text1(3).Text = Format(Text1(3).Text, "000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    FechaApertura = CDate(RecuperaValor(CadenaSeleccion, 1))
End Sub

Private Sub frmTun_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(4).Text = Format(Text1(4).Text, "00")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0")
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
    Text2(2).Text = Format(Text2(2).Text, "##0.00")
End Sub

Private Sub ImgFam_Click()
    Set frmFam = New frmFamilias
    frmFam.DatosADevolverBusqueda = "0|1|"
    frmFam.Show
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

Private Sub mnLineas_Click()
Dim b As Button
    Set b = Toolbar1.Buttons(10)
    Toolbar1_ButtonClick b
    Set b = Nothing
End Sub

Private Sub mnSalir_Click()
    If modo = 5 Then
        Exit Sub
    Else
        PulsadoSalir = True
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub ImgIva_Click()
    If vParam.HayContabilidad = False Then
        Set frmIva = New frmTiposIva
        frmIva.DatosADevolverBusqueda = "0|1|"
        frmIva.Show
    Else
        MandaBusquedaIvas ""
        'llamamos a contabilidad los tipos de iva
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    Dim f As Date
    Dim vFecRec As Date
    Dim mTag As New CTag
    If modo = 0 Or modo = 2 Then Exit Sub
    Select Case Index
       Case 0
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(15).Text <> "" Then
                If IsDate(Text1(15).Text) Then f = Text1(15).Text
            End If
            Set frmC = New frmCal
            frmC.Fecha = f
            frmC.Show vbModal
            If modo = 3 Or modo = 4 Or modo = 1 Then
                Text1(15).Text = frmC.Fecha
                mTag.DarFormato Text1(15)
            End If
            Set frmC = Nothing
       Case 1
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(17).Text <> "" Then
                If IsDate(Text1(17).Text) Then f = Text1(17).Text
            End If
            Set frmC = New frmCal
            frmC.Fecha = f
            frmC.Show vbModal
            If modo = 3 Or modo = 4 Or modo = 1 Then
                Text1(17).Text = frmC.Fecha
                mTag.DarFormato Text1(17)
            End If
            Set frmC = Nothing
       Case 2
            'En los tag
            'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
            f = Now
            If Text1(22).Text <> "" Then
                If IsDate(Text1(22).Text) Then f = Text1(22).Text
            End If
            Set frmC = New frmCal
            frmC.Fecha = f
            frmC.Show vbModal
            If modo = 3 Or modo = 4 Or modo = 1 Then
                Text1(22).Text = frmC.Fecha
                mTag.DarFormato Text1(22)
            End If
            Set frmC = Nothing
   
   End Select
End Sub

Private Sub ImgTun_Click()
    Set frmTun = New frmTiposUnidad
    frmTun.DatosADevolverBusqueda = "0|1|"
    frmTun.Show
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 18 Then Exit Sub  ' caso de pulsar ALT
    AsignarTeclasFuncion KeyCode
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
   Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If

End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim SQL As String

    kCampo = Index
    
    If modo = 1 Then
        Text1(Index).BackColor = vbYellow
    Else
        'en los datos fitosanitarios solo entraremos si la
        ' familia lo permite
        If Index = 18 Or Index = 19 Or Index = 23 Then
            If Text1(3).Text = "" Then
                PonerFoco Text1(5)
            Else
                If Not EsFitosanitario(CInt(Text1(3).Text)) Then
                    Text1(Index).Text = ""
                    PonerFoco Text1(5)
                Else
                    Text1(Index).SelStart = 0
                    Text1(Index).SelLength = Len(Text1(Index).Text)
                End If
            End If
        Else
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
        End If
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 20 Then
            SSTab1.Tab = 1
            PonerFoco Text1(10)
        Else
        If Index = 17 Then
            If modo <> 1 Then
                If EsFitosanitario(CInt(Text1(3).Text)) Then
                    SSTab1.Tab = 2
                    PonerFoco Text1(18)
                Else ' VRS:1.0.4(3)
                    PonerFoco cmdAceptar
                End If
            Else
                SSTab1.Tab = 2
                PonerFoco Text1(18)
            End If
        Else
            SendKeys "{tab}"
        End If
        End If
   Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim i As Integer
    Dim SQL As String
    Dim mTag As CTag
    Dim valor As Currency
    ''Quitamos blancos por los lados
   
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite
    End If

    If Text1(Index).Text = "" Then Exit Sub
    
    If modo = 1 And ConCaracteresBusqueda(Text1(Index).Text) Then Exit Sub
    
    Select Case Index
        Case 1, 2, 18, 19, 23
            ' No dejamos introducir comillas en ningun campo tipo texto
            If InStr(1, Text1(Index).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ningún campo de texto", vbExclamation
                Text1(Index).Text = Replace(Format(Text1(Index).Text, ">"), "'", "", , , vbTextCompare)
                PonerFoco Text1(Index)
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, ">")
            
            'VRS:1.0.1(1)
            If Index = 2 Then
                If Len(Text1(Index).Text) <> 8 And Len(Text1(Index).Text) <> 13 Then
                        MsgBox "La longitud de un Código EAN debe de ser 8 ó 13. Reintroduzca", vbExclamation
                        Text1(Index).Text = "" 'VRS:1.0.2(1)
                        PonerFoco Text1(Index)
                        Exit Sub
                Else
                    If Not CodigoEanCorrecto(Text1(Index).Text, Len(Text1(Index).Text)) Then
                        MsgBox "El Código EAN introducido no es correcto. Reintroduzca", vbExclamation
                        Text1(Index).Text = "" 'VRS:1.0.2(1)
                        PonerFoco Text1(Index)
                        Exit Sub
                    End If
                End If
            End If
            
        Case 0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 20, 21, 24
'            If (Modo <> 1) Then
'                If Text1(Index).Text = "" Then
'                    Text1(Index).SetFocus
'                    MsgBox "Debes introducir un valor ", vbExclamation
'
'                Else
             If Text1(Index).Text <> "" Then
                If EsNumerico(Text1(Index).Text) Then
                            Select Case Index
                                Case 0
                                    Text1(Index).Text = Format(Text1(Index).Text, "000000")
                                Case 3
                                    Text1(Index).Text = Format(Text1(Index).Text, "000")
                                    Text2(1).Text = DevuelveDesdeBD(1, "nomfamia", "sfamia", "codfamia|", Text1(Index).Text & "|", "N|", 1)
                                    SSTab1.TabEnabled(2) = EsFitosanitario(CInt(Text1(3).Text))
                                    If modo <> 1 And Text2(1).Text = "" Then
                                        MsgBox "Código no existe. Reintroduzca.", vbExclamation
                                        Text1(Index).Text = ""
                                        PonerFoco Text1(Index)
                                    End If
                                Case 4
                                    Text1(Index).Text = Format(Text1(Index).Text, "00")
                                    Text2(0).Text = DevuelveDesdeBD(1, "nomtipun", "stipun", "codtipun|", Text1(Index).Text & "|", "N|", 1)
                                    If modo <> 1 And Text2(0).Text = "" Then
                                        MsgBox "Código no existe. Reintroduzca.", vbExclamation
                                        Text1(Index).Text = ""
                                        PonerFoco Text1(Index)
                                    End If
                                    
                                Case 5
                                    Text1(Index).Text = Format(Text1(Index).Text, "00")
                                    If Text1(Index).Text <> "" Then
                                        If Not vParam.HayContabilidad Then
                                            Text2(2).Text = DevuelveDesdeBD(1, "porceiva", "tiposiva", "codigiva|", Text1(Index).Text & "|", "N|", 1)
                                        Else
                                            Text2(2).Text = DevuelveDesdeBD(2, "porceiva", "tiposiva", "codigiva|", Text1(Index).Text & "|", "N|", 1)
                                        End If
                                        If modo <> 1 And Text2(2).Text = "" Then
                                            MsgBox "Código no existe. Reintroduzca.", vbExclamation
                                            Text1(Index).Text = ""
                                            PonerFoco Text1(Index)
                                        End If
                                    Else
                                        Text2(2).Text = ""
                                    End If
                                    
                                    If Text2(2).Text <> "" Then
                                        Text2(2).Text = Format(Text2(2).Text, "##0.00")
                                        RecalculoIva (6)
                                        RecalculoIva (7)
                                    End If
                                     
                                Case 6, 7, 8, 9, 16, 24  'VRS:1.0.5(6)
                                    If modo = 1 Then Exit Sub
                                    
                                    If InStr(1, Text1(Index).Text, ",") > 0 Then
                                        valor = ImporteFormateado(Text1(Index).Text)
                                    Else
                                        valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                                    End If
                                    
                                    Text1(Index).Text = Format(valor, "##,###,##0.0000")   'VRS:1.0.5(6)
                            
                                    Select Case Index
                                        Case 6, 7
                                            RecalculoIva (Index)
                                        Case 8, 9
                                            RecalculoBase (Index)
                                    End Select
                                
                                Case 10, 11, 12, 13, 14
                                    If modo = 1 Then Exit Sub
                                    If InStr(1, Text1(Index).Text, ",") > 0 Then
                                        valor = ImporteFormateado(Text1(Index).Text)
                                    Else
                                        valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                                    End If
                                    'VRS:1.0.1(2)
                                    If Index = 13 Then
                                        If valor < 0 Then MsgBox "La cantidad de stock actual introducida es negativa", vbInformation
                                    End If
                                    
                                    
                                    Text1(Index).Text = Format(valor, "##,###,##0.00")
                                Case 20, 21
                                    If modo = 1 Then Exit Sub
                                    If InStr(1, Text1(Index).Text, ",") > 0 Then
                                        valor = ImporteFormateado(Text1(Index).Text)
                                    Else
                                        valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                                    End If
                                    Text1(Index).Text = Format(valor, "##0.00")
                            
                            End Select
                Else
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
              End If
        Case 15, 17, 22
            If Text1(Index).Text <> "" Then
              If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                    Exit Sub
              End If
              Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
            End If
    End Select
    
'    If Index = 9 Then SSTab1.TabIndex = 0
    'Si queremos hacer algo ..
'    Select Case Index
'        Case 2, 3
'
    '---
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
Dim CadB As String
CadB = ObtenerBusqueda(Me)

If CadB = "" Then
    MsgBox vbCrLf & "  Debe introducir alguna condición de búsqueda. " & vbCrLf, vbExclamation
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
Dim tabla As String
Dim titulo As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 10, "Código")
        Cad = Cad & ParaGrid(Text1(1), 60, "Nombre")
        Cad = Cad & ParaGrid(Text1(3), 16, "Familia")
        Cad = Cad & ParaGrid(Text1(4), 15, "Unidad")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|3|"
            frmB.vTitulo = "Articulos"
            frmB.vSelElem = 0
            frmB.vConexionGrid = 1
            'frmB.vBuscaPrevia = chkVistaPrevia
            frmB.vCargaFrame = False
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
        Screen.MousePointer = vbDefault

End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
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
    Dim i As Integer
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    
    If PonerCamposForma(Me, Data1) Then
        Combo2.Text = DevuelveDesdeBD(1, "descsino", "tiposino", "tiposino|", Data1.Recordset!statusin & "|", "N|", 1)
    End If
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If modo = 2 Then DataGrid1.Enabled = True
    
    Text2(1).Text = DevuelveDesdeBD(1, "nomfamia", "sfamia", "codfamia|", Text1(3).Text & "|", "N|", 1)
    Text2(0).Text = DevuelveDesdeBD(1, "nomtipun", "stipun", "codtipun|", Text1(4).Text & "|", "N|", 1)
    If vParam.HayContabilidad Then
        Text2(2).Text = DevuelveDesdeBD(2, "porceiva", "tiposiva", "codigiva|", Text1(5).Text & "|", "N|", 1)
    Else
        Text2(2).Text = DevuelveDesdeBD(1, "porceiva", "tiposiva", "codigiva|", Text1(5).Text & "|", "N|", 1)
    End If
    Text2(2).Text = Format(Text2(2).Text, "#0.00")
    
    '    PonerCtasIVA
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub
'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim i As Integer
    Dim b As Boolean
    
    
    If modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nuevo"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar"
    End If
        
    'ASIGNAR MODO
    modo = Kmodo
    
    PonerIndicador lblIndicador, modo
    If modo = 0 Then LimpiarCampos
    
    If modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva Linea de Lote"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar Linea de Lote"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar Linea de Lote"
        
        
        Text1(0).Enabled = False
        Text1(1).Enabled = False
        Text1(2).Enabled = False
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.Tab = 2
    Else
        Text1(0).Enabled = True
        Text1(1).Enabled = True
        Text1(2).Enabled = True
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
    End If
    
    b = (modo < 5)
    chkVistaPrevia.Visible = b
    
    
    If modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For i = 0 To Text1.Count - 1
            'Text1(I).BackColor = vbWhite
            Text1(i).BackColor = &H80000018
        Next i
        'Combo1.BackColor = &H80000018
        'chkVistaPrevia.Visible = False
    End If
    'Modo = Kmodo
    'chkVistaPrevia.Visible = (Modo = 1)
    
    b = (modo = 0) Or (modo = 5) Or (modo = 2)
    Toolbar1.Buttons(6).Enabled = (b And vUsu.NivelSumi <= 2)
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    Toolbar1.Buttons(10).Enabled = (b And vUsu.NivelSumi <= 2) 'Lineas factur
    Toolbar1.Buttons(11).Enabled = b
    
    b = b Or (modo = 5)
    DataGrid1.Enabled = b
    'Modificar
    Toolbar1.Buttons(7).Enabled = (b And vUsu.NivelSumi <= 2)
    'eliminar
    Toolbar1.Buttons(8).Enabled = (b And vUsu.NivelSumi <= 2)
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    Else
        cmdRegresar.Visible = False
    End If
    
    'Modo insertar o modificar
    b = (Kmodo >= 3)  '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = b Or modo = 1
    cmdCancelar.Visible = b Or modo = 1
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(1).Enabled = Not b And modo <> 1
    Toolbar1.Buttons(2).Enabled = Not b And modo <> 1
    
'    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    b = (modo = 2) Or modo = 0
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = b
        Text1(i).BackColor = vbWhite
    Next i
    Frame4.Enabled = (modo <> 4)
    
    ImgFam.Enabled = Not b
    ImgTun.Enabled = Not b
    ImgIva.Enabled = Not b
    Combo2.Enabled = Not b
    
    If modo = 4 Then
        SSTab1.TabEnabled(2) = EsFitosanitario(Text1(3).Text)
    Else
        SSTab1.TabEnabled(2) = True
    End If
   
    Frame4.Enabled = Not (modo = 3 Or modo = 4)
    Frame5.Enabled = Not (modo = 4)
    
    If modo = 5 Then
        PonerFoco DataGrid1
    Else
        PonerFoco chkVistaPrevia
    End If
End Sub

Private Function DatosOk() As Boolean
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim Datos As String
Dim Cad As String

b = CompForm(Me)
IndiceErroneo = 0
If (b = True) And ((modo = 3) Or (modo = 4)) Then

        For i = 0 To 23
             If InStr(1, Text1(i).Text, "'") > 0 Then
                MsgBox "No puede introducir el carácter ' en ningún campo de texto", vbExclamation
                IndiceErroneo = i
                DatosOk = False
                Exit Function
             End If
        Next i
        
        'VRS:1.0.1(1)
        If Len(Text1(2).Text) <> 8 And Len(Text1(2).Text) <> 13 Then
                MsgBox "La longitud de un Código EAN debe de ser 8 ó 13. Reintroduzca", vbExclamation
                IndiceErroneo = 2
                DatosOk = False
                Exit Function
        Else  'VRS:1.0.2(1)
            If Not CodigoEanCorrecto(Text1(2).Text, Len(Text1(2).Text)) Then
                MsgBox "El Código EAN introducido no es correcto. Reintroduzca", vbExclamation
                IndiceErroneo = 2
                DatosOk = False
                Exit Function
            End If
        End If
        
        ' el codigo de iva no puede ser 0
        If Text1(5).Text = "" Then
            MsgBox "El código de Iva no puede ser vacío.", vbExclamation
            IndiceErroneo = 5
            DatosOk = False
            Exit Function
        End If
        ' el codigo de iva no puede ser 0
        If Text1(5).Text = "" Then
            MsgBox "El código de Iva no puede ser nulo", vbExclamation
            IndiceErroneo = 5
            DatosOk = False
            Exit Function
        End If
        
        ' comprobamos que: stock minimo <= puntopedido <= stock maximo
        If CCur(Text1(10).Text) > CCur(Text1(11).Text) Then
            MsgBox "El Stock Mínimo no puede ser superior al Punto de Pedido", vbExclamation
            DatosOk = False
            IndiceErroneo = 10
            Exit Function
        Else
        If CCur(Text1(11).Text) > CCur(Text1(12).Text) Then
            MsgBox "El Punto de Pedido no puede ser superior al Stock Máximo.", vbExclamation
            DatosOk = False
            IndiceErroneo = 11
            Exit Function
        End If
        End If
        
        ' si la familia es fito debemos tener los datos correspondientes
        If EsFitosanitario(CInt(Text1(3).Text)) Then
            If Text1(18).Text = "" Or Text1(19).Text = "" Or _
               Text1(23).Text = "" Then
               MsgBox "La familia del artículo es fitosanitaria. Debe introducir los datos correspondientes.", vbExclamation
               SSTab1.Tab = 2
               PonerFoco Text1(18)
               DatosOk = False
               Exit Function
            End If
        Else
            'VRS:1.0.1(5)
            If EsFitosanitario(familia) And modo = 4 Then
                 Cad = "El artículo va a pasar a ser de una familia no fitosanitaria y " & vbCrLf
                 Cad = Cad & "se borrarán todos los datos fitosanitarios." & vbCrLf & vbCrLf
                 Cad = Cad & "           ¿Seguro que desea continuar?" & vbCrLf & vbCrLf
                 i = MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) 'VRS:1.0.1(11)
                 If i = vbYes Then
                    ' si no es familia fitosanitaria limpiamos los datos fitosanitarios.
                        Text1(18).Text = ""
                        Text1(19).Text = ""
                        Text1(23).Text = ""
             
                        BorrarStocksFitosanitarios CLng(Text1(0).Text)
                 Else
                        DatosOk = False
                        Exit Function
                 End If
                
            End If
            
        End If
        
End If

If (b = True) And (modo = 3) Then

    'Estamos insertando
    'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
    'la N es pa dir que es numeric
     Datos = DevuelveDesdeBD(1, "codartic", "sartic", "codartic|", Text1(0).Text & "|", "N|", 1)
     If Datos <> "" Then
        MsgBox "Ya existe el codigo de articulo : " & Text1(0).Text, vbExclamation
        DatosOk = False
        IndiceErroneo = 0
        Exit Function
    End If
End If

DatosOk = b
End Function

'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()
 
    Dim SQL As String
    Dim RS As ADODB.Recordset

    '***** canviar el SQL *********************
    SQL = "Select Max(codartic) from " & NombreTabla
    '******************************************
    Text1(0).Text = 1
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            Text1(0).Text = RS.Fields(0) + 1
        End If
    End If
    RS.Close
    
    ValorAnterior = Text1(0).Text

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonBuscar
        Case 2
            BotonVerTodos
        Case 6
            If modo <> 5 Then
                BotonAnyadir
            Else
                'AÑADIR linea factura
                AnyadirLinea True, True
            End If
        Case 7
            If modo <> 5 Then
                'Intentamos bloquear la cuenta
                BotonModificar
            Else
                'MODIFICAR linea factura
                ModificarLinea
            End If
        Case 8
            If modo <> 5 Then
                BotonEliminar
            Else
                'ELIMINAR linea de lote
                EliminarLineaLotes
            End If
        Case 10
            If Not BloqueaRegistroForm(Me) Then Exit Sub
            'Nuevo Modo
            If EsFitosanitario(CInt(Text1(3).Text)) Then
                If Combo2.ListIndex = 0 Then
                    MsgBox "Artículo inventariandose. No se puede modificar.", vbExclamation
                    DesBloqueaRegistroForm Text1(0)
                    Exit Sub
                End If
                PonerModo 5
                'Fuerzo que se vean las lineas
                cmdAceptar.Visible = False
                cmdCancelar.Caption = "&Cabecera"
                lblIndicador.Caption = "Lineas detalle"
            Else
                DesBloqueaRegistroForm Text1(0)
                
            End If
            
        Case 12
            mnSalir_Click
        Case 14 To 17
            Desplazamiento (Button.Index - 14)
        Case 11
'            MsgBox "Hola"
        
        
'            Screen.MousePointer = vbHourglass
'            frmListado.Opcion = 2 'Listado de articulos
'            frmListado.Show
        Case Else
    
    End Select
End Sub

Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = Bol
    Next i
End Sub

Private Sub PonerFoco(ByRef Text As Object)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False
    ' tabla de lineas de hco de facturas
    SQL = DevuelveDesdeBD(1, "codartic", "slifpc", "codartic|", Data1.Recordset!codArtic & "|", "N|", 1)
    If SQL <> "" Then
        MsgBox "Este articulo está vinculado con registros de la tabla de lineas de hco de facturas", vbExclamation
        Exit Function
    End If
    
    ' tabla de lineas de pedidos de compra
    SQL = DevuelveDesdeBD(1, "codartic", "slippr", "codartic|", Data1.Recordset!codArtic & "|", "N|", 1)
    If SQL <> "" Then
       MsgBox "Este articulo está vinculado con registros de la tabla de lineas de pedidos de compra", vbExclamation
        Exit Function
    End If
    
    ' tabla de lineas de albaranes de compra
    SQL = DevuelveDesdeBD(1, "codartic", "slialp", "codartic|", Data1.Recordset!codArtic & "|", "N|", 1)
    If SQL <> "" Then
        MsgBox "Este articulo está vinculado con registros de la tabla de lineas de albaranes de compras", vbExclamation
        Exit Function
    End If
    
    ' tabla de lineas de albaranes de venta
    SQL = DevuelveDesdeBD(1, "codartic", "slialb", "codartic|", Data1.Recordset!codArtic & "|", "N|", 1)
    If SQL <> "" Then
       MsgBox "Este articulo está vinculado con registros de la tabla de lineas de albaranes de venta", vbExclamation
        Exit Function
    End If
    
    ' tabla de lotes
    SQL = DevuelveDesdeBD(1, "codartic", "slotes", "codartic|", Data1.Recordset!codArtic & "|", "N|", 1)
    If SQL <> "" Then
       MsgBox "Este articulo está vinculado con registros de la tabla de registros fitosanitarios", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True
End Function

Private Sub RecalculoIva(Index As Integer)
Dim Base As Currency
Dim Iva As Currency
Dim ImpIva As Currency

    Iva = ImporteFormateado(Text2(2).Text)
    Select Case Index
        Case 6
            Base = ImporteFormateado(Text1(Index).Text)
        Case 7
            Base = ImporteFormateado(Text1(Index).Text)
    End Select
    
    If Iva <> 0 Then
        ImpIva = Round(Base * Iva / 100, 4)   'VRS:1.0.5(6)
    Else
        ImpIva = 0
    End If
    
    Select Case Index
        Case 6
            Text1(8).Text = Format((Base + ImpIva), "##,###,##0.0000")  'VRS:1.0.5(6)
        Case 7
            Text1(9).Text = Format((Base + ImpIva), "##,###,##0.0000")  'VRS:1.0.5(6)
    End Select
    
End Sub

Private Sub RecalculoBase(Index As Integer)
Dim Base As Currency
Dim Iva As Currency
Dim ImpIva As Currency

    Iva = ImporteFormateado(Text2(2).Text)
    Select Case Index
        Case 8
            ImpIva = ImporteFormateado(Text1(Index).Text)
        Case 9
            ImpIva = ImporteFormateado(Text1(Index).Text)
    End Select
    Base = Round(ImpIva / (1 + (Iva / 100)), 4)     'VRS:1.0.5(6)
    Select Case Index
        Case 8
            Text1(6).Text = Format(Base, "##,###,##0.0000") 'VRS:1.0.5(6)
        Case 9
            Text1(7).Text = Format(Base, "##,###,##0.0000") 'VRS:1.0.5(6)
    End Select
    
End Sub

Private Sub MandaBusquedaIvas(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = "Codigo|codigiva|N|20" & "·" & "Descripcion|nombriva|T|40" & "·" & "Porcentaje|porceiva|N|40" & "·"
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmBc = New frmBuscaGrid 'Conta
            frmBc.vCampos = Cad
            frmBc.vTabla = "tiposiva"
            frmBc.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmBc.vDevuelve = "0|1|2|"
            frmBc.vTitulo = "Tipos de Iva"
            frmBc.vSelElem = 0
            frmBc.vConexionGrid = 2
            'frmBc.vBuscaPrevia = True
            frmBc.vCargaFrame = False
            '#
            frmBc.Show
            
' ahora ya no va a ser modal tendremos que recuperar datos en el activate
'            vbModal
'            Set frmBc = Nothing
' fin ya no es modal

            
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
'            If HaDevueltoDatos Then
'                If (Not adodc1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                txtAux(2).SetFocus
'            End If
        End If
End Sub

Private Sub CargarCombo()
'###
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Si, 1-No
    
    Combo2.Clear
    Combo2.AddItem "Si"
    Combo2.ItemData(Combo2.NewIndex) = 0

    Combo2.AddItem "No"
    Combo2.ItemData(Combo2.NewIndex) = 1

End Sub
Private Sub CargaGrid(Enlaza As Boolean)
Dim b As Boolean
    b = DataGrid1.Enabled
    DataGrid1.Enabled = False
    CargaGrid2 Enlaza
    DataGrid1.Enabled = b
End Sub

Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = MontaSQLCarga(Enlaza)
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockPessimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    DataGrid1.Tag = "Asignando"
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-numaspre  1.- Lin aspre
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Visible = False
    
    'Cuenta
    DataGrid1.Columns(2).Caption = "Nro.Lote"
    DataGrid1.Columns(2).Width = 2900
   
    DataGrid1.Columns(3).Caption = "Reg.Fitosanitario"
    DataGrid1.Columns(3).Width = 2900 '4395

    DataGrid1.Columns(4).Caption = "Cantidad"
    DataGrid1.Columns(4).Width = 1870
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).NumberFormat = "###,###,##0.00"
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        anc = 323
        Txtaux(0).Left = DataGrid1.Left + 330
        Txtaux(0).Width = DataGrid1.Columns(2).Width - 60

        Txtaux(1).Left = DataGrid1.Columns(3).Left + 240
        Txtaux(1).Width = DataGrid1.Columns(3).Width - 60

        Txtaux(2).Left = DataGrid1.Columns(4).Left + 260
        Txtaux(2).Width = DataGrid1.Columns(4).Width - 40

        CadAncho = True
    End If
        
    For i = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(i).AllowSizing = False
    Next i
    
    DataGrid1.Tag = "Calculando"
    'Obtenemos las sumas
'    ObtenerSumas
    
    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Function MontaSQLCarga(Enlaza As Boolean) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    ' Si ENLAZA -> Enlaza con el data1
    '           -> Si no lo cargamos sin enlazar a nngun campo
    '--------------------------------------------------------------------
    Dim SQL As String
    SQL = "SELECT sarticlotes.codartic, sarticlotes.numlinea, sarticlotes.nrolotes, "
    SQL = SQL & " sarticlotes.regfitosanitario,sarticlotes.cantidad "
    SQL = SQL & " FROM (sarticlotes LEFT JOIN sartic ON sarticlotes.codartic ="
    SQL = SQL & " sartic.codartic)"
    If Enlaza Then
        SQL = SQL & " WHERE sarticlotes.codartic = " & Data1.Recordset!codArtic
    Else
        SQL = SQL & " WHERE sarticlotes.codartic = -1"
    End If
    SQL = SQL & " ORDER BY sarticlotes.numlinea"
    MontaSQLCarga = SQL
End Function

Private Sub AnyadirLinea(Limpiar As Boolean, DesdeBoton As Boolean)
    Dim anc As Single
    Dim i As Single
    
'    If ModificandoLineas <> 0 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    
    Numlinea = ObtenerSigueinteNumeroLinea
    If Aux = 0 Then
        If DesdeBoton Then MsgBox "No se pueden insertar mas lineas. La cantidad es igual al stock actual.", vbExclamation
        LLamaLineas anc, 0, True
        cmdCancelar.Caption = "&Cabecera"
        Exit Sub
    End If
    
    lblIndicador.Caption = "AÑADIR"
    
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        'DataGrid1.Row = DataGrid1.Row + 1
    End If
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        If DataGrid1.Row >= 7 Then
            DataGrid1.Row = DataGrid1.Row + 1
            anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
        Else
        
        anc = anc + DataGrid1.RowTop(DataGrid1.Row + 1) + 15
        End If
        
    End If
    LLamaLineas anc, 1, Limpiar
    
    Txtaux(0).Enabled = True
    Txtaux(1).Enabled = True
    Txtaux(2).Text = Format(Aux, "0.00")
    Txtaux(1).Text = Text1(18).Text
    'Ponemos el foco
    PonerFoco Txtaux(0)
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    
    Numlinea = adodc1.Recordset!Numlinea
    Me.lblIndicador.Caption = "MODIFICAR"
     
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    Txtaux(0).Text = adodc1.Recordset.Fields!nrolotes
    Txtaux(1).Text = adodc1.Recordset.Fields!Regfitosanitario
    Txtaux(2).Text = adodc1.Recordset.Fields!cantidad
    Txtaux(2).Text = Format(Txtaux(2).Text, "###,###,##0.00")
    
    LLamaLineas anc, 2, False
    PonerFoco Txtaux(2)

End Sub

Private Sub EliminarLineaLotes()
Dim fechainv As Date
Dim SQL As String

On Error GoTo FinEliminarLineaLotes

    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    
    
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Lineas de Lotes." & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar el Nro.Lote: "
    SQL = SQL & adodc1.Recordset.Fields!nrolotes
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? " 'VRS:1.0.1(11)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
        
        'borramos la linea
        SQL = "Delete from sarticlotes"
        SQL = SQL & " WHERE sarticlotes.numlinea = " & adodc1.Recordset!Numlinea
        SQL = SQL & " AND sarticlotes.codartic= " & Data1.Recordset!codArtic & ";"
        DataGrid1.Enabled = False
        Conn.Execute SQL
        
        CargaGrid (Not Data1.Recordset.EOF)
        DataGrid1.Enabled = True
        
    End If

FinEliminarLineaLotes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar linea lotes"
    End If
End Sub

Private Function ObtenerSigueinteNumeroLinea() As Long
    Dim RS As ADODB.Recordset
    Dim i As Long
    Dim SQL As String
    
    Set RS = New ADODB.Recordset
    SQL = " WHERE sarticlotes.codartic=" & Data1.Recordset!codArtic & ";"
    RS.Open "SELECT Max(numlinea) FROM sarticlotes" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    i = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then i = RS.Fields(0)
    End If
    RS.Close
    
    'La suma
    SumaLinea = 0
    If i > 0 Then
        RS.Open "SELECT sum(cantidad) FROM sarticlotes" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then SumaLinea = RS.Fields(0)
        End If
        RS.Close
    End If
    Set RS = Nothing
    
    'Lo que falta lo fijamos en aux. El importe es de la bASE IMPONIBLE si fuera del total seria Text2(4).Text
    Aux = ImporteFormateado(Text1(13).Text)
    Aux = Aux - SumaLinea
    
    ObtenerSigueinteNumeroLinea = i + 1
End Function

Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim b As Boolean
    DeseleccionaGrid
    cmdCancelar.Caption = "&Cancelar"
    ModificandoLineas = xModo
    b = (xModo = 0)
    cmdAceptar.Visible = Not b
    'cmdCancelar.Visible = Not b
    
    CamposAux Not b, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim i As Integer
    Dim J As Integer
    
    DataGrid1.Enabled = Not Visible

    J = Txtaux.Count - 1
    For i = 0 To J
        If i <> 7 Then
            Txtaux(i).Visible = Visible
            Txtaux(i).Top = Altura
        End If
    Next i
    
    If Limpiar Then
        For i = 0 To J
            Txtaux(i).Text = ""
        Next i
    End If
    
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu
        Case "Listado"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 2 'Listado de articulos
            FrmListado.Show
        Case "Etiquetas"
            Screen.MousePointer = vbHourglass
            FrmListado.Opcion = 17 'Listado de articulos
            FrmListado.Show
    End Select
    
    
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    With Txtaux(Index)
        AntiguoText1 = .Text
        
        If Index <> 5 Then
             .SelStart = 0
            .SelLength = Len(.Text)
        Else
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtaux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        'Esto sera k hemos pulsado el ENTER
        txtAux_LostFocus Index
        cmdAceptar_Click
    Else
        If KeyCode = 113 Then
            'Esto sera k pedimos la calculadora
            PideCalculadora
        End If
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            cmdCancelar_Click
        End If
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Sng As Double
    Dim Preciouc As Double
    Dim valor As Double
    Dim vRegfito As String
        
    If ModificandoLineas = 0 Then Exit Sub
    
    Txtaux(Index).Text = Trim(Txtaux(Index).Text)
    If Txtaux(Index).BackColor = vbYellow Then
        Txtaux(Index).BackColor = vbWhite
    End If

    If Txtaux(Index).Text = "" Then Exit Sub
    
    Select Case Index
    
        Case 2
            If (modo <> 1) Then
                If Txtaux(Index).Text = "" Then
                    PonerFoco Txtaux(Index)
                    MsgBox "Debes introducir un valor ", vbExclamation
                    
                Else
                   If EsNumerico(Txtaux(Index).Text) Then
                            If InStr(1, Txtaux(Index).Text, ",") > 0 Then
                                valor = ImporteFormateado(Txtaux(Index).Text)
                            Else
                                valor = CCur(TransformaPuntosComas(Txtaux(Index).Text))
                            End If
                            
                            Txtaux(Index).Text = Format(valor, "##,###,##0.00")
                            If valor = 0 And Txtaux(0).Text <> "" Then
                                    MsgBox "Debe introducir una cantidad distinta de cero", vbExclamation
                                    PonerFoco Txtaux(Index)
                             'VRS:1.0.1(3)
                            Else
                                If valor < 0 Then
                                    MsgBox "La cantidad introducida es negativa.", vbInformation
                                End If
                            End If
                    Else
                        Txtaux(Index).Text = ""
                        PonerFoco Txtaux(Index)
                    End If
              End If
            End If
        Case 0, 1 'ampliacion es un campo de txto
'            If InStr(1, txtAux(Index).Text, "'") > 0 Then
'                MsgBox "No puede introducir el carácter ' en ningún campo de texto", vbExclamation
'                txtAux(Index).Text = Replace(Format(txtAux(Index).Text, ">"), "'", "", , , vbTextCompare)
'                PonerFoco txtAux(Index)
'                Exit Sub
'            End If
            Txtaux(Index).Text = Format(Txtaux(Index).Text, ">")
        
    End Select
End Sub

Private Function InsertarModificar() As Boolean
Dim EraFito As String
Dim SQL As String

    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    Conn.BeginTrans
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS
        SQL = "INSERT INTO sarticlotes (codartic, numlinea, nrolotes, regfitosanitario, cantidad) "
        SQL = SQL & "VALUES (" & Text1(0).Text & "," & Numlinea & ",'"
        SQL = SQL & DevNombreSQL(Txtaux(0).Text) & "','" & DevNombreSQL(Txtaux(1).Text) & "',"
        
        If Txtaux(2).Text = "" Then
          SQL = SQL & ValorNulo & ")"
        Else
          SQL = SQL & TransformaComasPuntos(ImporteFormateado(Txtaux(2).Text)) & ")"
        End If
        
    Else
        'MODIFICAR
        SQL = "UPDATE sarticlotes SET "
        
        SQL = SQL & " nrolotes = '" & DevNombreSQL(Txtaux(0).Text) & "',"
        '--regfitosanitario
        SQL = SQL & " regfitosanitario = '" & DevNombreSQL(Txtaux(1).Text) & "',"
        'cantidad
        If Txtaux(2).Text = "" Then
          SQL = SQL & " cantidad = " & ValorNulo
        Else
          SQL = SQL & " cantidad = " & TransformaComasPuntos(ImporteFormateado(Txtaux(2).Text))
        End If
        SQL = SQL & " WHERE sarticlotes.codartic = " & Text1(0).Text
        SQL = SQL & " AND sarticlotes.numlinea = " & Numlinea & ";"
    End If
    
    Conn.Execute SQL
    
    Conn.CommitTrans
   
    InsertarModificar = True
    Exit Function

EInsertarModificar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "InsertarModificar linea de albarán.", Err.Description
        'If ModificandoLineas = 0 Then Conn.RollbackTrans
        Conn.RollbackTrans
    End If
End Function

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub

Private Function Eliminar() As Boolean
Dim i As Integer
Dim SQL As String

        SQL = " WHERE codartic=" & Data1.Recordset!codArtic
        
        'Lineas
        Conn.Execute "Delete  from sarticlotes " & SQL
        
        'Cabeceras
        Conn.Execute "Delete  from sartic " & SQL
       
        Eliminar = True
        
End Function

Private Sub PideCalculadora()
On Error GoTo EPideCalculadora
    Shell App.Path & "\arical.exe", vbNormalFocus
    Exit Sub
EPideCalculadora:
    Err.Clear
End Sub

Private Function AuxOK() As String
Dim SQL As String

    'Nrolotes  no puede estar vacio
    If Txtaux(0).Text = "" Then
        AuxOK = "Número de lote de Artículo no puede estar vacio."
        Exit Function
    End If
    'Registro Fitosanitario no puede estar vacio
    If Txtaux(1).Text = "" Then
        AuxOK = "Número de registro fitosanitario no puede estar vacio."
        Exit Function
    End If
    'Cantidad no puede estar vacio
    If Txtaux(2).Text = "" Then
        AuxOK = "Cantidad no puede estar vacia."
        Exit Function
    Else
        If Not IsNumeric(Txtaux(2).Text) Then
            AuxOK = "La cantidad debe de ser numérico."
            Exit Function
        End If
    End If
    
    If RegFitosanitarioExistente Then
        AuxOK = "Nro.Lote de registro fitosanitario existente"
        Exit Function
    End If
    AuxOK = ""
End Function

Private Function RegFitosanitarioExistente() As Boolean
Dim SQL As String
Dim cantidad As Long
Dim RS As ADODB.Recordset
    
    RegFitosanitarioExistente = False
    
    If modo = 5 And ModificandoLineas <> 1 Then Exit Function
        
    SQL = ""
    SQL = DevuelveDesdeBD(1, "nrolotes", "sarticlotes", "codartic|nrolotes|regfitosanitario|", Text1(0).Text & "|" & Txtaux(0).Text & "|" & Txtaux(1).Text & "|", "N|T|T|", 3)
        
    RegFitosanitarioExistente = (SQL <> "")
    
End Function

Private Function ComprobarStocksLineas() As Boolean
Dim SQL As String
Dim cantidad As Double
Dim RS As ADODB.Recordset
    
    If EsFitosanitario(Text1(3).Text) Then
    
        Set RS = New ADODB.Recordset
        SQL = "select sum(cantidad) from sarticlotes where codartic = " & Text1(0).Text
        RS.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
        cantidad = 0
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then cantidad = RS.Fields(0)
        End If
        RS.Close
    
        ComprobarStocksLineas = (cantidad = CLng(Text1(13).Text))
    Else
        ComprobarStocksLineas = True
    End If

End Function


Private Sub BorrarStocksFitosanitarios(articulo As Long)
Dim Cad As String
    
    Cad = "delete from sarticlotes where codartic = " & Format(articulo, "000000")
    
    Conn.Execute Cad

End Sub


Private Sub AsignarTeclasFuncion(key As Integer)

    If modo = 2 Or modo = 0 Then
        Select Case key
            Case vbESC '27
                If modo = 0 Then
                    Toolbar1_ButtonClick Toolbar1.Buttons(12)
                Else
                    PonerModo 0
                End If
            Case vbAnterior '33
                If modo = 2 Then Desplazamiento (1)
            Case vbSiguiente '34
                If modo = 2 Then Desplazamiento (2)
            Case vbPrimero  ' 36 ' inicio
                If modo = 2 Then Desplazamiento (0)
            Case vbUltimo '35 ' fin
                If modo = 2 Then Desplazamiento (3)
           Case vbBuscar
                Toolbar1_ButtonClick Toolbar1.Buttons(1)
           Case vbVerTodos
                Toolbar1_ButtonClick Toolbar1.Buttons(2)
           Case vbAñadir
                Toolbar1_ButtonClick Toolbar1.Buttons(6)
           Case vbModificar
                 If modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(7)
            Case vbEliminar
                 If modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(8)
            Case vbLineas
                 If modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(10)
            Case vbImprimir
                 If modo = 2 Then Toolbar1_ButtonClick Toolbar1.Buttons(13)
            Case vbSalir
                  Toolbar1_ButtonClick Toolbar1.Buttons(12)
        End Select
   End If

End Sub


Private Sub InsertarMovimientoAlmacen()
Dim CMov As CMovimientos

    Set frmMens = New frmMensajes
    
    frmMens.Opcion = 2
    frmMens.DatosADevolverBusqueda = "1|"
    frmMens.Text3 = Format(Now, "dd/mm/yyyy")
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    Set CMov = New CMovimientos
    CMov.almacen = 0
    CMov.articulo = CStr(CLng(Text1(0).Text))
    CMov.cantidad = 0
    CMov.ConnBD = Conn
    CMov.Fechamov = FechaApertura
    CMov.HoraMov = CDate(Format(CStr(FechaApertura), "dd/mm/yyyy") & " " & Format(Now, "hh:mm:ss"))
    CMov.LetraSerie = ""
    CMov.preciomp = ImporteFormateado(Text1(24).Text)
    CMov.InsertarMovimientoCierre

End Sub

