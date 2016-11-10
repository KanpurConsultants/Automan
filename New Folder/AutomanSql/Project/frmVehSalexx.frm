VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehSale 
   Appearance      =   0  'Flat
   BackColor       =   &H00BEE4D3&
   Caption         =   "Vehicle Sale Bill"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11775
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CancelBill 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancel Bill"
      Height          =   315
      Left            =   7335
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   45
      Width           =   1995
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   59
      Left            =   5475
      TabIndex        =   47
      Top             =   6345
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   58
      Left            =   4905
      TabIndex        =   46
      Top             =   6345
      Width           =   510
   End
   Begin VB.Frame FrmPrn 
      BackColor       =   &H00CAECF0&
      Caption         =   "Printing Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3000
      Left            =   1665
      TabIndex        =   113
      Top             =   1500
      Visible         =   0   'False
      Width           =   8400
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5580
         TabIndex        =   116
         Top             =   375
         Width           =   2325
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   117
         Text            =   "12-MAR-2003"
         Top             =   945
         Width           =   1305
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4170
         MaxLength       =   15
         TabIndex        =   123
         Top             =   2010
         Width           =   2325
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5955
         TabIndex        =   118
         Top             =   945
         Width           =   540
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2490
         MaxLength       =   15
         TabIndex        =   124
         Text            =   "12-MAR-2003"
         Top             =   2280
         Width           =   1305
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4170
         MaxLength       =   20
         TabIndex        =   120
         Top             =   1215
         Width           =   2325
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1950
         TabIndex        =   119
         Top             =   1215
         Width           =   480
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2250
         TabIndex        =   122
         Top             =   2010
         Width           =   480
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   1950
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   121
         Top             =   1485
         Width           =   4545
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmVehSale.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   6675
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Printer "
         Top             =   1950
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmVehSale.frx":030A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   6675
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Screen"
         Top             =   2280
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmVehSale.frx":0614
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   6675
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Printer "
         Top             =   2610
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   15
         Picture         =   "frmVehSale.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Screen"
         Top             =   2640
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00CAECF0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   8085
         MousePointer    =   99  'Custom
         Picture         =   "frmVehSale.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Delete Current Record"
         Top             =   15
         Width           =   315
      End
      Begin VB.OptionButton OptPlain 
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Plain"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2145
         TabIndex        =   114
         Top             =   315
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Optpre 
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "PrePrinted "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2145
         TabIndex        =   115
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   50
         Left            =   4440
         TabIndex        =   154
         Top             =   390
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   48
         Left            =   330
         TabIndex        =   153
         Top             =   960
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         Height          =   1695
         Left            =   225
         Top             =   900
         Width           =   6330
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Printer Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   45
         TabIndex        =   139
         Top             =   15
         Width           =   8085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTO Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   49
         Left            =   3075
         TabIndex        =   138
         Top             =   2025
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temp. Sale Certificate Y/N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   51
         Left            =   3735
         TabIndex        =   137
         Top             =   960
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Certificate Print Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   52
         Left            =   315
         TabIndex        =   136
         Top             =   2295
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Body"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   53
         Left            =   3105
         TabIndex        =   135
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seating Capacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   54
         Left            =   330
         TabIndex        =   134
         Top             =   1230
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight In Printing Y/N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   55
         Left            =   330
         TabIndex        =   133
         Top             =   2025
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CertificateNarration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   56
         Left            =   345
         TabIndex        =   132
         Top             =   1500
         Width           =   1590
      End
      Begin VB.Label LblPrinter 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Current Active Printer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   345
         TabIndex        =   131
         Top             =   2640
         Width           =   6315
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Stationary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   41
         Left            =   420
         TabIndex        =   130
         Top             =   465
         Width           =   825
      End
      Begin VB.Line Line6 
         X1              =   1695
         X2              =   1395
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line8 
         X1              =   1710
         X2              =   1710
         Y1              =   720
         Y2              =   420
      End
      Begin VB.Line Line2 
         X1              =   1710
         X2              =   2040
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line3 
         X1              =   1710
         X2              =   2055
         Y1              =   735
         Y2              =   735
      End
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   56
      Left            =   9135
      MaxLength       =   8
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   2190
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   57
      Left            =   10410
      MaxLength       =   12
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   2190
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   54
      Left            =   1470
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   855
      Width           =   360
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   55
      Left            =   1845
      MaxLength       =   40
      TabIndex        =   10
      Top             =   855
      Width           =   4635
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   53
      Left            =   1470
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   585
      Width           =   360
   End
   Begin MSDataGridLib.DataGrid DGFin 
      Height          =   3885
      Left            =   6030
      Negotiate       =   -1  'True
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   -3465
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   6853
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Financier Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Financier"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2865
      Left            =   -825
      Negotiate       =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   7455
      Visible         =   0   'False
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   5054
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Model Help"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Model Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Model Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Chas_Type"
         Caption         =   "Chassis Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6075.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1349.858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   2445
      Left            =   10875
      Negotiate       =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   5790
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   4313
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Chassis Help"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Chassis No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "EngineNo"
         Caption         =   "Engine No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Model"
         Caption         =   "Model"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Col_desc"
         Caption         =   "Colour"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PBill_No"
         Caption         =   "TelcoBillNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PBill_Date"
         Caption         =   "TelcoDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "PurVNo"
         Caption         =   "Purch No."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Pur_VDate"
         Caption         =   "PurchDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Rate"
         Caption         =   "Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Al_Name"
         Caption         =   "Alloted"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2039.811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   50
      Left            =   9465
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "01234"
      Top             =   1590
      Width           =   585
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   49
      Left            =   8535
      MaxLength       =   20
      TabIndex        =   57
      Text            =   "01234567890123456789"
      Top             =   6615
      Width           =   2265
   End
   Begin MSDataGridLib.DataGrid DGBook 
      Height          =   2175
      Left            =   -8340
      Negotiate       =   -1  'True
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Booking Help"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "code"
         Caption         =   "Booking No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ord_date"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Name"
         Caption         =   "Party"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   5790.047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   47
      Left            =   5475
      TabIndex        =   49
      Top             =   6885
      Width           =   1335
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   10335
      TabIndex        =   79
      Top             =   7065
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   510
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   0
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   16379351
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   7620
      Negotiate       =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   7275
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Voucher No"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Voucher No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2865.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   4935
      Left            =   9255
      Negotiate       =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7110
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tax Form Help "
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Form Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   10275
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "0123456789"
      Top             =   2475
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   22
      Left            =   10275
      MaxLength       =   12
      TabIndex        =   25
      Top             =   2745
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   20
      Left            =   5985
      TabIndex        =   21
      Top             =   2745
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   23
      Left            =   10275
      TabIndex        =   26
      Top             =   3015
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   5475
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6615
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   42
      Left            =   8535
      TabIndex        =   50
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   38
      Left            =   5475
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5535
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   6000
      TabIndex        =   15
      Top             =   1935
      Width           =   480
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   8535
      MaxLength       =   10
      TabIndex        =   56
      Top             =   6345
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   8535
      MaxLength       =   15
      TabIndex        =   58
      Text            =   "012345678901234"
      Top             =   6885
      Width           =   2265
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   46
      Left            =   8535
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   5805
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   44
      Left            =   8535
      TabIndex        =   52
      Top             =   5265
      Width           =   3255
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   43
      Left            =   8535
      TabIndex        =   53
      Top             =   5535
      Width           =   2265
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   45
      Left            =   8535
      TabIndex        =   51
      Top             =   4995
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5115
      MaxLength       =   25
      TabIndex        =   23
      Top             =   3015
      Width           =   2670
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   40
      Left            =   5475
      TabIndex        =   45
      Top             =   6075
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4275
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2475
      Width           =   2205
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   1245
      MaxLength       =   15
      TabIndex        =   18
      Top             =   2475
      Width           =   1890
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4590
      MaxLength       =   15
      TabIndex        =   17
      Top             =   2205
      Width           =   1890
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1245
      MaxLength       =   15
      TabIndex        =   16
      Top             =   2205
      Width           =   2205
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9465
      TabIndex        =   6
      Top             =   1860
      Width           =   1770
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1245
      MaxLength       =   40
      TabIndex        =   14
      Text            =   " "
      Top             =   1935
      Width           =   2205
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   13
      Text            =   " 0123456789012345678901234567890123456789"
      Top             =   1665
      Width           =   5010
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   12
      Text            =   " "
      Top             =   1395
      Width           =   5010
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1125
      Width           =   5010
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   8385
      Negotiate       =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   7995
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Site Help"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Site Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "code"
         Caption         =   "Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   1815
      TabIndex        =   33
      Top             =   5805
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   4905
      TabIndex        =   38
      Top             =   4725
      Width           =   510
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   4905
      TabIndex        =   40
      Top             =   4995
      Width           =   510
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   1815
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6885
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   5475
      TabIndex        =   41
      Top             =   4995
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   5475
      TabIndex        =   42
      Top             =   5265
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9465
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1050
      Width           =   2085
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   1815
      TabIndex        =   32
      Top             =   5535
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   1815
      TabIndex        =   31
      Top             =   5265
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   5475
      TabIndex        =   39
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8D8FE&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   5070
      TabIndex        =   27
      Top             =   4230
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSDataGridLib.DataGrid DGADItem 
      Height          =   4935
      Left            =   8685
      Negotiate       =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   7815
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Addition Fitments Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Item Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   1755
      MaxLength       =   25
      TabIndex        =   22
      Top             =   3015
      Width           =   2220
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   24
      Left            =   1815
      TabIndex        =   29
      Top             =   4725
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "99,99,999.99"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   25
      Left            =   1815
      TabIndex        =   30
      Text            =   "99999999.99"
      Top             =   4995
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   10080
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1590
      Width           =   1155
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   1845
      MaxLength       =   40
      TabIndex        =   8
      Top             =   585
      Width           =   4635
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9465
      MaxLength       =   21
      TabIndex        =   1
      Top             =   510
      Width           =   2235
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   48
      Left            =   8535
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   6075
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   1815
      TabIndex        =   36
      Top             =   6615
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   1815
      TabIndex        =   35
      Top             =   6345
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   39
      Left            =   5475
      TabIndex        =   44
      Top             =   5805
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   29
      Left            =   1815
      TabIndex        =   34
      Top             =   6075
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   1245
      MaxLength       =   25
      TabIndex        =   20
      Text            =   "0123456789012345678901234"
      Top             =   2745
      Width           =   2730
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9465
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1320
      Width           =   1770
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1335
      Left            =   135
      TabIndex        =   28
      Top             =   3345
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   11
      BackColorFixed  =   12632319
      ForeColorFixed  =   128
      BackColorSel    =   16703741
      ForeColorSel    =   12582912
      BackColorBkg    =   13298928
      GridColor       =   12632319
      GridColorFixed  =   8421631
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "SrNo.|Add/Del Item |Type      |Qty|Rate  |Amount|Itemcode"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   52
      Left            =   7830
      Locked          =   -1  'True
      TabIndex        =   144
      TabStop         =   0   'False
      Text            =   "VFa"
      Top             =   2745
      Width           =   1275
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   51
      Left            =   7830
      Locked          =   -1  'True
      TabIndex        =   145
      TabStop         =   0   'False
      Text            =   "0123456789"
      Top             =   2475
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Over Tax  @"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   57
      Left            =   3435
      TabIndex        =   155
      Top             =   6360
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Doc No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Index           =   47
      Left            =   7725
      TabIndex        =   152
      Top             =   2205
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   28
      Left            =   135
      TabIndex        =   149
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label LblAcPostBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Posting By"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   6615
      TabIndex        =   147
      Top             =   2490
      Width           =   1155
   End
   Begin VB.Label LblAcPostDt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   7380
      TabIndex        =   146
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Axle No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   4
      Left            =   7005
      TabIndex        =   142
      Top             =   6630
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Fuel Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   36
      Left            =   3435
      TabIndex        =   140
      Top             =   6900
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   25
      Left            =   3510
      TabIndex        =   112
      Top             =   2220
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   26
      Left            =   135
      TabIndex        =   111
      Top             =   2490
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Round Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   46
      Left            =   3435
      TabIndex        =   108
      Top             =   6630
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Invoice Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   20
      Left            =   7035
      TabIndex        =   107
      Top             =   4740
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sob Total [B]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   19
      Left            =   3435
      TabIndex        =   106
      Top             =   5550
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt Y/N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   8
      Left            =   5235
      TabIndex        =   105
      Top             =   1950
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Book No.*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   45
      Left            =   7005
      TabIndex        =   104
      Top             =   6360
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   44
      Left            =   7005
      TabIndex        =   103
      Top             =   6900
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Total Adv."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Index           =   35
      Left            =   7035
      TabIndex        =   102
      Top             =   5820
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   34
      Left            =   7035
      TabIndex        =   101
      Top             =   5280
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source of Fund"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   33
      Left            =   7035
      TabIndex        =   100
      Top             =   5550
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financed Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   31
      Left            =   7035
      TabIndex        =   99
      Top             =   5010
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RTO Office"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   30
      Left            =   4155
      TabIndex        =   98
      Top             =   3030
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax on Other Fitment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   29
      Left            =   3435
      TabIndex        =   97
      Top             =   6090
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   27
      Left            =   3345
      TabIndex        =   96
      Top             =   2490
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   24
      Left            =   135
      TabIndex        =   95
      Top             =   2220
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No.*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   23
      Left            =   8280
      TabIndex        =   94
      Top             =   1875
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   21
      Left            =   135
      TabIndex        =   93
      Top             =   1950
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   14
      Left            =   135
      TabIndex        =   92
      Top             =   1140
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mfg. Bill No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   5
      Left            =   9255
      TabIndex        =   91
      Top             =   2490
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   7
      Left            =   9780
      TabIndex        =   89
      Top             =   2745
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Octroi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   6
      Left            =   165
      TabIndex        =   88
      Top             =   5550
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NDP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   11
      Left            =   9825
      TabIndex        =   87
      Top             =   3015
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SobTotal [A]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   18
      Left            =   165
      TabIndex        =   86
      Top             =   6900
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surch. on Tax  @"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   17
      Left            =   3435
      TabIndex        =   85
      Top             =   5010
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. Charges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   16
      Left            =   3435
      TabIndex        =   84
      Top             =   5280
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   15
      Left            =   8280
      TabIndex        =   83
      Top             =   1065
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incidental charges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   13
      Left            =   165
      TabIndex        =   82
      Top             =   5280
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax                    @"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   12
      Left            =   3435
      TabIndex        =   81
      Top             =   4740
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   32
      Left            =   135
      TabIndex        =   77
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add./ Ded./ Short.*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   0
      Left            =   135
      TabIndex        =   76
      Top             =   3030
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   10
      Left            =   165
      TabIndex        =   75
      Top             =   4740
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rebate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   9
      Left            =   165
      TabIndex        =   74
      Top             =   5010
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008080FF&
      Height          =   1695
      Left            =   8190
      Top             =   465
      Width           =   3540
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   7380
      TabIndex        =   73
      Top             =   435
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill  No.*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   8280
      TabIndex        =   72
      Top             =   1605
      Width           =   825
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division           "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   8280
      TabIndex        =   71
      Top             =   795
      Width           =   1155
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   10140
      TabIndex        =   70
      Top             =   795
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Y/N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   43
      Left            =   4950
      TabIndex        =   69
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   42
      Left            =   8280
      TabIndex        =   68
      Top             =   525
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net OutStanding"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   41
      Left            =   7035
      TabIndex        =   66
      Top             =   6090
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   40
      Left            =   165
      TabIndex        =   65
      Top             =   6630
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MVT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   39
      Left            =   165
      TabIndex        =   64
      Top             =   6360
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Fitment Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   38
      Left            =   3435
      TabIndex        =   63
      Top             =   5820
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transit Insurance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   37
      Left            =   165
      TabIndex        =   62
      Top             =   6090
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Temp. Registration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   22
      Left            =   165
      TabIndex        =   61
      Top             =   5820
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   3
      Left            =   135
      TabIndex        =   60
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   2
      Left            =   8280
      TabIndex        =   59
      Top             =   1335
      Width           =   765
   End
End
Attribute VB_Name = "frmVehSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mSP5 As String = "     "
Dim mInvPrefixHt As Integer
Dim RsChassis As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim RsMod As ADODB.Recordset
Dim RsSite As ADODB.Recordset
Dim rsFin As ADODB.Recordset
Dim RsADItem As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim RSBook As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim GridKey As Integer
Dim DocId As String * 21
Public mVType As String
Dim VoucherEditFlag As Boolean
Dim FinAcCode As String
Dim vPrefix As String
Dim CancelBillY_N As Boolean

'Grid color scheme
'Private Const CellBackColLeave As String = &HBAD3C9
'Private Const CellForeColLeave As String = &H0&
'Private Const CellBackColEnter As String = &HC0E0FF
'Private Const GridBackColorBkg As String = &HCAECF0
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const TxtDocId As Byte = 0
Private Const SiteCode As Byte = 1
Private Const Vdate As Byte = 2
Private Const SerialNo As Byte = 3
Private Const BookNo As Byte = 4
Private Const Party As Byte = 5
Private Const Add1 As Byte = 6
Private Const Add2 As Byte = 7
Private Const Add3 As Byte = 8
Private Const City As Byte = 9
Private Const Govt_YN As Byte = 10
Private Const Model As Byte = 11
Private Const ChassisNo As Byte = 12
Private Const EngineNo As Byte = 13
Private Const Colours As Byte = 14
Private Const FormType As Byte = 15
Private Const ADType As Byte = 16
Private Const RTO  As Byte = 17
Private Const SpclInfo  As Byte = 18
Private Const SrvBookNo As Byte = 19
Private Const Taxable As Byte = 20
Private Const TelcoInvNo As Byte = 21
Private Const TelcoInvDate As Byte = 22
Private Const NDP As Byte = 23
Private Const SaleRate As Byte = 24
Private Const Rebate As Byte = 25
Private Const IncCharge As Byte = 26
Private Const Octroi As Byte = 27
Private Const TempReg As Byte = 28
Private Const TransIns As Byte = 29
Private Const MVT As Byte = 30
Private Const Transportation As Byte = 31
Private Const SubTotA As Byte = 32
Private Const TaxPer As Byte = 33
Private Const TaxAmt As Byte = 34
Private Const TaxSurPer As Byte = 35
Private Const TaxSurch As Byte = 36
Private Const MisCharge As Byte = 37
Private Const SubTotB As Byte = 38
Private Const OthFitAmt As Byte = 39
Private Const OthFitTax As Byte = 40
Private Const ROff  As Byte = 41
Private Const GTotAmt As Byte = 42
Private Const FundSource As Byte = 43
Private Const FB_Code As Byte = 44
Private Const FinAmt As Byte = 45
Private Const AdvAmt As Byte = 46
Private Const FuelAmt As Byte = 47
Private Const NetOStng As Byte = 48
Private Const TransAxlNo As Byte = 49
Private Const InvPrefix As Byte = 50    'Invoice Prefix used in DocID 12-04-03
Private Const AcPostByName As Byte = 51
Private Const AcPostDate As Byte = 52
Private Const NamePrefix As Byte = 53
Private Const FNamePrefix As Byte = 54
Private Const fname As Byte = 55
Private Const DelChNo As Byte = 56
Private Const DelChDate As Byte = 57
Private Const TOTPer As Byte = 58
Private Const TOTAmt As Byte = 59

Private Const ADItem As Byte = 1
Private Const Qty As Byte = 2
Private Const Rate As Byte = 3
Private Const Amt As Byte = 4
Private Const TaxPer1 As Byte = 5
Private Const TaxAmt1 As Byte = 6
Private Const TaxSurPer1 As Byte = 7
Private Const TaxSurAmt1 As Byte = 8
Private Const FinalAmt As Byte = 9
Private Const ADItemCode  As Byte = 10

Private Const TempInvDate As Byte = 0
Private Const CertiTempYN As Byte = 1
Private Const Seet As Byte = 2
Private Const Body As Byte = 3
Private Const Narr As Byte = 4
Private Const WtPrn As Byte = 5
Private Const RTOName As Byte = 6
Private Const CertiPrnDate As Byte = 7
Private Const DocType As Byte = 8

Private Const DocInv As Byte = 0
Private Const DocSaleCert As Byte = 1
Private Const DocForm22 As Byte = 2
Private Const DocForm22A As Byte = 3

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName$, mRepNameCert$, mRepName22$, mRepName22A$, mOldChasis$

Private Sub CancelBill_Click()
    Dim I As Integer, Rst As ADODB.Recordset
    CancelBillY_N = True
    If TopCtrl1.TopText2 <> "Browse" Then
        MsgBox "Cancellation Denied in this mode !", vbInformation
        CancelBillY_N = False
        Exit Sub
    End If
          
    'OfftakeIncentiveSrlNo
    'SubventionSrlNo
    GSQL = "Select OfftakeIncentiveSrlNo,SubventionSrlNo from veh_stock where Sal_DocId='" & Txt(TxtDocId) & "' and ChassisNo  = '" & Txt(ChassisNo).TEXT & "'"
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        If Rst!OfftakeIncentiveSrlNo <> "" Or Rst!SubventionSrlNo <> "" Then
            MsgBox "Offtake Incentive Claim / Subvention Letter made." & vbCrLf & "Deletion denied!", vbCritical, "Deletion Denied"
            Set Rst = Nothing
            Exit Sub
        End If
    End If
    Set Rst = Nothing
    If GCn.Execute("Select DelCh_DocId from  veh_order where Inv_DocId = '" & Master!SearchCode & "'").Fields(0).Value <> "" Then
        MsgBox "Delivery has been made against this Invoice", vbInformation, "Deletion Denied": Set Rst = Nothing: Exit Sub
    End If
If AcPostAuthorisation(Txt(AcPostByName)) = False Then Exit Sub
    If MsgBox(" Are You Sure to Cancel the Invoice ? ", vbInformation + vbYesNo, "Cancel Bill Message") = vbYes Then
        DocId = Txt(TxtDocId)
        GCn.Execute ("Insert into Veh_Order1 select * from Veh_Order where Inv_DocId='" & DocId & "'")
        GCn.Execute ("Update Veh_Order1 set Inv_UEntDt=#" & date & "# where Inv_DocId='" & DocId & "'")
        GCn.BeginTrans
        GCnFaV.BeginTrans
'*******START POSTING
        Dim MsgStr$, rsCtrlAc As ADODB.Recordset, rsTemp As ADODB.Recordset, mPostFinAmt As Byte
        Dim mGTotAmt As Double, mTOT_Ac_Code$, mCommNarr$
        
        Set rsCtrlAc = New ADODB.Recordset
        rsCtrlAc.CursorLocation = adUseClient
        rsCtrlAc.Open "Select Fitment_Ac,Fuel_Ac,VehROff_Ac From AcControls", GCnFaV, adOpenStatic, adLockReadOnly
        If rsCtrlAc.RecordCount <= 0 Then
            MsgStr = "Please Add Records in A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        If IsNull(rsCtrlAc!Fitment_Ac) Or rsCtrlAc!Fitment_Ac = "" Or _
            IsNull(rsCtrlAc!Fuel_Ac) Or rsCtrlAc!Fuel_Ac = "" Or _
            IsNull(rsCtrlAc!VehROff_Ac) Or rsCtrlAc!VehROff_Ac = "" Then
            MsgStr = "Please define Fitment,Fuel and Round Off A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        rsForm.MoveFirst        'Vehicle Sale A/c Code, Tax A/c Code, Surcharge A/c Code
        rsForm.FIND "Name ='" & Txt(FormType) & "'"
        If IsNull(rsForm!PurSal_Ac_Code) Or rsForm!PurSal_Ac_Code = "" Then
            MsgStr = "Please Define Sale A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        'Tax A/c Code Checking
        If Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(OthFitTax)) <> 0 Then
            If IsNull(rsForm!Tax_Ac_Code) Or rsForm!Sur_Ac_Code = "" Then
                MsgStr = "Please Define Tax A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
                'CancelBillLedgerPost = False
                GoTo lblExit
            End If
        End If
        'Financier A/c Checking
        mTOT_Ac_Code = G_FaCn.Execute("select iif(isnull(totax_ac),'',TOTax_Ac) as TOT_Ac from AcControls where Div_Code='" & PubDivCode & "'").Fields(0).Value
        If Val(Txt(TOTAmt)) <> 0 And mTOT_Ac_Code = "" Then
            MsgStr = "Please define TOT A/c Code in Vehicle Controls" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        mPostFinAmt = GCn.Execute("select iif(isnull(PostFinAmt),0,postfinamt) as PostFinAmt from Syctrl").Fields(0).Value
        If mPostFinAmt = 1 And Val(Txt(FinAmt)) <> 0 Then
            If Txt(FundSource) = "Hypothication" Or Txt(FundSource) = "Hire Purchase" Then
                Set rsTemp = New ADODB.Recordset
                rsTemp.CursorLocation = adUseClient
                rsTemp.Open "Select switch(Ac_YN='1','Y',Ac_YN<>'1','N') as ACYN,AcCode From ContractFinance where FinCode='" & Txt(FB_Code).Tag & "' ", GCn, adOpenStatic, adLockReadOnly
                If rsTemp!AcYN = "Y" Then
                    If rsTemp!AcCode = "" Or IsNull(rsTemp!AcCode) Then
                        MsgStr = "Please define A/c Code in Financier Master" & vbCrLf & "A/c Posting Aborted !"
                        GoTo lblExit
                    End If
                End If
            End If
        End If
       ' If CheckCtrls Then 'Control setting found Ok
            'CancelBillLedgerPost = True: Exit Function
        'End If
        
        'A/c Posting related declarations
        Dim mBookDocID$
        Dim LedgAry(7) As LedgRec, mResult As Byte, mNarr$
        
        'Sale Party A/c
        mBookDocID = GCn.Execute("select OrdDocId from Veh_Order where Inv_DocId='" & Txt(TxtDocId) & "'").Fields(0).Value
        mNarr = "By Cancelled Sales Invoice No." & Txt(InvPrefix) & Txt(SerialNo) & " Dt. " & date & " Chassis " & Txt(ChassisNo)
        mCommNarr = mNarr & "[Common]"
        I = 0
        LedgAry(I).SubCode = Txt(Party).Tag
        mGTotAmt = Val(Txt(GTotAmt))
        If mPostFinAmt = 0 Then
            mGTotAmt = Val(Txt(GTotAmt)) + Val(Txt(FinAmt))
        End If
        LedgAry(I).AmtCr = Round(Val(Txt(GTotAmt)), 2)
        LedgAry(I).Narration = mNarr
        'Vehicle Sale A/c
        'Modi LPS 05.12.2003
        If Val(Txt(SubTotA)) + Val(Txt(MisCharge)) - Val(Txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsForm!PurSal_Ac_Code
            LedgAry(I).AmtDr = Round(Val(Txt(SubTotA)) + Val(Txt(MisCharge)), 2)
            LedgAry(I).Narration = mNarr
        End If
        'eof Modi
        'Fitment Amount
        If Val(Txt(OthFitAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fitment_Ac
            LedgAry(I).AmtDr = Round(Val(Txt(OthFitAmt)), 2)
            LedgAry(I).Narration = mNarr & " Additional Fitments on Vehicle Sale Bill"
        End If
        'Tax Amt
        If Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(OthFitTax)) <> 0 Then
            If rsForm!Tax_Ac_Code <> "" And rsForm!Sur_Ac_Code <> "" _
                 And rsForm!Tax_Ac_Code <> rsForm!Sur_Ac_Code Then
                If Val(Txt(TaxAmt)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                    LedgAry(I).AmtDr = Round(Val(Txt(TaxAmt)) + Val(Txt(OthFitTax)), 2)
                    LedgAry(I).Narration = mNarr & " Sale Tax"
                End If
                If Val(Txt(TaxSurch)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Sur_Ac_Code
                    LedgAry(I).AmtDr = Round(Val(Txt(TaxSurch)), 2)
                    LedgAry(I).Narration = mNarr & " Surcharge on Sales Tax"
                End If
            Else
                I = I + 1
                LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                LedgAry(I).AmtDr = Round(Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(OthFitTax)), 2)
                LedgAry(I).Narration = mNarr & " Sales Tax & Surcharge"
            End If
        End If
        If Val(Txt(TOTAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = mTOT_Ac_Code
            LedgAry(I).AmtDr = Val(Txt(TOTAmt))
            LedgAry(I).Narration = mNarr & " TOT Amt"
        End If
        If Val(Txt(ROff)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!VehROff_Ac
            If Val(Txt(ROff)) > 0 Then
                LedgAry(I).AmtDr = Round(Val(Txt(ROff)), 2)
            Else
                LedgAry(I).AmtCr = Round(Abs(Val(Txt(ROff))), 2)
            End If
            LedgAry(I).Narration = mNarr & " Round Off"
        End If
        'Fuel Amount
        If Val(Txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fuel_Ac
            LedgAry(I).AmtCr = Round(Val(Txt(FuelAmt)), 2)
            LedgAry(I).Narration = mNarr & " Fuel Amount"
        End If
        
        If mPostFinAmt = 1 And Val(Txt(FinAmt)) <> 0 Then
            If Txt(FundSource) = "Hypothication" Or Txt(FundSource) = "Hire Purchase" Then
                If rsTemp!AcCode = "" Or IsNull(rsTemp!AcCode) Then
                Else
                    I = I + 1
                    LedgAry(I).SubCode = rsTemp!AcCode
                    LedgAry(I).AmtCr = Round(Val(Txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amt."
                    I = I + 1
                    LedgAry(I).SubCode = Txt(Party).Tag
                    LedgAry(I).AmtDr = Round(Val(Txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amount."
                End If
            End If
        End If
        DocId = left(Txt(TxtDocId), 8) & "Cancl" & Right(Txt(TxtDocId), 8)
        mResult = LedgerPost("C", LedgAry, GCnFaV, DocId, CDate(Txt(Vdate)), mCommNarr)
        If mResult <> 1 Then
            MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
           ' ProcAcPost = False
        Else
            'ProcAcPost = True
        End If
lblExit:
If MsgStr <> "" Then
    MsgBox MsgStr, vbCritical, "A/c Posting"
ElseIf Err.NUMBER > 0 Then
    MsgBox Err.Description, vbCritical, "A/c Posting"
End If
Set rsCtrlAc = Nothing
Set rsTemp = Nothing
If MsgBox("Print Cancelled bill Copy ?", vbInformation + vbYesNo, "Print Information") = vbYes Then
    Call Cmdprint_Click(5)
End If
'****************END POSTING
        'Unposting of Ledger completed
        'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
        'GCn.Execute ("update hiscard set Dealer_Code='', CouponNo='', TransAxelNo='', Supplier_BillNo='', Supplier_BillDate=Null, Name='" & PubComp_Name & "',Add1='',Add2='',Add3='',CityCode='',Govt_YN =  0 " & _
            " where Chassis ='" & txt(ChassisNo) & "'")
        GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_VDate=Null, Srv_BookNo = '', TransAxlNo='' " & _
            "where ChassisNo  = '" & Txt(ChassisNo) & "' and Sal_DocId = '" & Txt(TxtDocId) & "'"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ADItem) <> "" Then
                GCn.Execute ("delete from veh_purch2 where DocId='" & Txt(TxtDocId) & "'")
            End If
        Next
        GCn.Execute ("update veh_order  " & _
            "set Inv_DocId='',Inv_DocIDHelp='' ,Inv_SiteCode='',Inv_VType='',Inv_No=Null ,Inv_Date=null,Form_Code='', " & _
            "TAX_Per=0,TAX_Amt=0,Surcharge_Per=0,Surcharge_Amt=0,MARGINE=0,VRATE=0,REBATE=0, " & _
            "InciChrg=0,Octroi=0,RegTemp=0,TransitInsu=0,Transport=0,MVT=0,OtherChrg=0,FIT_AMT=0,FIT_TAX=0, " & _
            "DieselAmt=0,MISC_INFO='',RTO='',Round_off=0, " & _
            "FB_Code='' , FIN_AcCode='', FIN_AMT=0, " & _
            "TrnType_Prn=0,Fund_Source=0,Chassis='' , Srv_BookNo='', " & _
            "Inv_UName='', Inv_UEntDt=null, Inv_UAE= '',Inv_AcPostByUName='',Inv_AcPostByUEntDt=Null " & _
            "where Inv_DocId='" & Txt(TxtDocId) & "'")
        GCnFaV.CommitTrans
        GCn.CommitTrans
        Master.Requery
        RSBook.Requery
        Call MoveRec
        BUTTONS True, Me, Master, 0
        CancelBillY_N = False
End If
Exit Sub

End Sub

Private Sub DGBook_Click()
    If RSBook.RecordCount > 0 Then
        Txt(BookNo).TEXT = RSBook!Code
        Txt(BookNo).Tag = RSBook!OrdDocId
        FillRecords
    End If
    Txt(BookNo).SetFocus
    DGBook.Visible = False
End Sub

Private Sub DGFin_Click()
    If rsFin.RecordCount > 0 Then
        Txt(FB_Code).TEXT = rsFin!Name
        Txt(FB_Code).Tag = rsFin!Code
        FinAcCode = rsFin!Code
    End If
    Txt(FB_Code).SetFocus
    DGFin.Visible = False
End Sub

Private Sub DGSite_Click()
    If RsSite.RecordCount > 0 Then
        Txt(SiteCode).TEXT = RsSite!Name
        Txt(SiteCode).Tag = RsSite!Code
    End If
    Txt(SiteCode).SetFocus
    DGSite.Visible = False
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
MsgBox Err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
'On Error GoTo ELoop
Dim I As Byte
TopCtrl1.Tag = PubUParam: WinSetting Me:     Ini_Grid

    mVType = "V_SB"
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Inv_DocId as SearchCode,Veh_Order.* from Veh_Order where left(Inv_DocID,1)='" & PubDivCode & "' and trim(mid(Inv_DocID,4,5))= '" & mVType & "' Order by Inv_Date desc,Inv_DocId desc", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
    Set RSBook = New ADODB.Recordset
    RSBook.CursorLocation = adUseClient
    RSBook.Open "Select distinct OrdDocID, trim(str(veh_order.Ord_No)) as code,veh_order.ord_date,subgroup.Name  from Veh_Order left join subgroup on subgroup.subcode = veh_order.partycode where left(veh_order.OrdDocId,1)= '" & PubDivCode & "' and veh_order.Inv_DocId='' ", GCn, adOpenDynamic, adLockOptimistic
    Set DGBook.DataSource = RSBook
    
    Set rsFin = New ADODB.Recordset
    rsFin.CursorLocation = adUseClient
    rsFin.Open "select fincode as code,finname & ',' & City.CityName as name,AcCode,FinBankCode from ContractFinance " & _
    "left join city on left(ContractFinance.City,4)=City.CityCode where fincatg = 0  order by finname", GCn, adOpenDynamic, adLockOptimistic
    Set DGFin.DataSource = rsFin
  
    Set RsMod = New ADODB.Recordset
    RsMod.CursorLocation = adUseClient
    RsMod.Open "select Model as code,Model_Desc as NAME, Chas_Type from Model where Div_Code='" & PubDivCode & "' order by Model", GCn, adOpenDynamic, adLockOptimistic
    Set DGMod.DataSource = RsMod
    
    Set rsForm = New ADODB.Recordset
    With rsForm
        .CursorLocation = adUseClient
        .Open "SELECT T.Form_Code as Code,T.Form_Desc as Name,T.Tax_Sur_Per,T.Tax_Per,T1.Tax_Ac_Code,T1.Sur_Ac_Code,T1.PurSal_Ac_Code " & _
            "FROM TaxForms as T left join TaxFormsAc as T1 on  T.Form_Code&'" & PubDivCode & "'=T1.Form_Code&T1.Div_Code " & _
            "where T.Vehicle_YN = 1 and T.Trn_Type = 'Sale' Order by Form_Desc ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGForm.DataSource = rsForm
    
    Set RsADItem = New ADODB.Recordset
    With RsADItem
        .CursorLocation = adUseClient
        .Open "SELECT  Prod_Code as code,Prod_name as name,Rate  FROM veh_amdModel order by  veh_amdModel.Prod_name ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGADItem.DataSource = RsADItem
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Exit Sub
ELoop:    MsgBox Err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsMod = Nothing
Set RsADItem = Nothing
Set RsSite = Nothing
Set rsForm = Nothing
Set RsVno = Nothing
Set RsChassis = Nothing
Set rsFin = Nothing
Set Master = Nothing
Set mListItem = Nothing
End Sub
Private Sub ListView_Click()
If FrmPrn.Visible = False Then
    Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    Txt(Val(ListView.Tag)).SetFocus
Else
    txtPrint(DocType).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txtPrint(DocType).SetFocus
End If
End Sub

Private Sub OptPlain_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Optpre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    Txt(SiteCode).SetFocus
     Txt(TOTPer) = MainLib.TOTCal()
Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer, Rst As ADODB.Recordset
Dim LedgAry(1) As LedgRec, mResult As Byte
    'OfftakeIncentiveSrlNo
    'SubventionSrlNo
    GSQL = "Select OfftakeIncentiveSrlNo,SubventionSrlNo from veh_stock where Sal_DocId='" & Txt(TxtDocId) & "' and ChassisNo  = '" & Txt(ChassisNo).TEXT & "'"
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        If Rst!OfftakeIncentiveSrlNo <> "" Or Rst!SubventionSrlNo <> "" Then
            MsgBox "Offtake Incentive Claim / Subvention Letter made." & vbCrLf & "Deletion denied!", vbCritical, "Deletion Denied"
            Set Rst = Nothing
            Exit Sub
        End If
    End If
    Set Rst = Nothing
    If GCn.Execute("Select DelCh_DocId from  veh_order where Inv_DocId = '" & Master!SearchCode & "'").Fields(0).Value <> "" Then
        MsgBox "Delivery has been made against this Invoice", vbInformation, "Deletion Denied": Set Rst = Nothing: Exit Sub
    End If
    
If AcPostAuthorisation(Txt(AcPostByName)) = False Then Exit Sub

If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    GCn.BeginTrans
    GCnFaV.BeginTrans
    'Unpost Ledger a/c
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, Txt(TxtDocId))
    If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    'Unposting of Ledger completed
'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
'    GCn.Execute ("update hiscard set Dealer_Code='', CouponNo='', TransAxelNo='', Supplier_BillNo='', Supplier_BillDate=Null, Name='" & PubComp_Name & "',Add1='',Add2='',Add3='',CityCode='',Govt_YN =  0 " & _
        " where Chassis ='" & txt(ChassisNo) & "'")
    GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_VDate=Null, Srv_BookNo = '', TransAxlNo='' " & _
        "where ChassisNo  = '" & Txt(ChassisNo) & "' and Sal_DocId = '" & Txt(TxtDocId) & "'"
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ADItem) <> "" Then
            GCn.Execute ("delete from veh_purch2 where DocId='" & Txt(TxtDocId) & "'")
        End If
    Next
    GCn.Execute ("update veh_order  " & _
        "set Inv_DocId='',Inv_DocIDHelp='' ,Inv_SiteCode='',Inv_VType='',Inv_No=Null ,Inv_Date=null,Form_Code='', " & _
        "TAX_Per=0,TAX_Amt=0,Surcharge_Per=0,Surcharge_Amt=0,MARGINE=0,VRATE=0,REBATE=0, " & _
        "InciChrg=0,Octroi=0,RegTemp=0,TransitInsu=0,Transport=0,MVT=0,OtherChrg=0,FIT_AMT=0,FIT_TAX=0, " & _
        "DieselAmt=0,MISC_INFO='',RTO='',Round_off=0, " & _
        "FB_Code='' , FIN_AcCode='', FIN_AMT=0, " & _
        "TrnType_Prn=0,Fund_Source=0,Chassis='' , Srv_BookNo='', " & _
        "Inv_UName='', Inv_UEntDt=null, Inv_UAE= '',Inv_AcPostByUName='',Inv_AcPostByUEntDt=Null " & _
        "where Inv_DocId='" & Txt(TxtDocId) & "'")
    GCnFaV.CommitTrans
    GCn.CommitTrans
    Master.Requery
    RSBook.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If Err.NUMBER <> 0 Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
'lp 10-03-03
    If Txt(DelChNo) <> "" Then MsgBox "Delivery Made, Edit denied !", vbInformation, "Validation": Exit Sub
'eof lp
    If AcPostAuthorisation(Txt(AcPostByName)) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    Txt(Vdate).SetFocus
    FGrid.AddItem FGrid.Rows
    Exit Sub
eloop1:
    If Err.NUMBER <> 0 Then
        MsgBox Err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, Master, 1
  Call MoveRec
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, Master, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, Master, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, Master, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox Err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
txtPrint(DocType) = "Sale Bill"
txtPrint(TempInvDate) = Txt(Vdate)
txtPrint(RTOName) = Txt(RTO)
'txtPrint(Seet) = GCn.Execute("Select SEAT from Model where Model='" & Txt(Model) & "'").Fields(0).Value
'txtPrint(Body) = GCn.Execute("Select INTD_USE from Veh_Order where inv_Docid='" & Txt(TxtDocId) & "'").Fields(0).Value
'txtPrint(Narr) = "Signature of Manufacturer/Dealer or Officer of Defence Department"

mRepName = IIf(OptPlain.Value = True, "VehSale", "VehSale")
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
CmdPrint(PWindows).SetFocus
If PubSpeedPrint Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
End Sub

Private Sub TopCtrl1_eRef()
    RsMod.Requery
    RSBook.Requery
    RsSite.Requery
    rsForm.Requery
    RsADItem.Requery
'    RsChassis.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    Dim DocIdHlp$, sqlstr$, mDlrID$, mPBIllNo$, mPBIllDate$
    Dim mFundSource As Byte
    Dim mTrntypeprn As Byte, mQuotDocID$, mQuotDocIDSrlNo As Integer
    On Error GoTo errlbl

If TxtGrid(0).Visible = True Then
    If TxtGridLeave = False Then
        TxtGrid(0).SetFocus
        Exit Sub
    End If
End If
Grid_Hide

If IsValid(Txt(SiteCode), "SiteCode") = False Then Exit Sub
If IsValid(Txt(Vdate), "Bill Date") = False Then Exit Sub
If IsValid(Txt(SerialNo), "Bill Number") = False Then Exit Sub
If IsValid(Txt(Party), "Party Name") = False Then Exit Sub
If IsValid(Txt(BookNo), "Booking No.") = False Then Exit Sub
If IsValid(Txt(FormType), "Form Type") = False Then Exit Sub
If IsValid(Txt(Model), "Model") = False Then Exit Sub
If IsValid(Txt(ChassisNo), "Chassis") = False Then Exit Sub

If Val(Txt(FinAmt)) > 0 Then
    If IsValid(Txt(FB_Code), "Financier") = False Then Exit Sub
    If IsValid(Txt(FundSource), "Source of Fund") = False Then Exit Sub
    If Txt(FundSource) <> "Hypothication" And _
        Txt(FundSource) <> "Hire Purchase" And _
        Txt(FundSource) <> "Lease" And _
        Txt(FundSource) <> "Agreement" And _
        Txt(FundSource) <> "Lease & Agreement" And _
        Txt(FundSource) <> "Loan Cum Hypt." Then
        
        MsgBox "Invalid Source of Fund !", vbCritical, "Fund Source Validation"
        Txt(FundSource).SetFocus
        Exit Sub
    End If
Else
'    If txt(FundSource) <> "Own Fund" Then
'        MsgBox "Financed Amount is zero, Correct Source of Fund", vbCritical, "Fund Source Validation"
'        txt(FundSource).SetFocus
'        Exit Sub
'    End If
End If
If IsValid(Txt(SrvBookNo), "Service Book No.") = False Then Exit Sub
GSQL = "Select Model,ChassisNo,Sal_DocID from Veh_Stock where ChassisNo<>'" & Txt(ChassisNo) & "' and Srv_BookNo = '" & Txt(SrvBookNo) & "'" ' and Model='" & Txt(Model) & "'"
sqlstr = "Select Model,ChassisNo,Sal_DocID from Veh_Stock where ChassisNo<>'" & Txt(ChassisNo) & "' and TransAxlNo = '" & Txt(TransAxlNo) & "'" ' and Model='" & Txt(Model) & "'"
'Service Book No. checking
Set Rst = New ADODB.Recordset
Rst.CursorLocation = adUseClient
Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
If Rst.RecordCount > 0 Then
    MsgBox "Service Book No. " & Txt(SrvBookNo) & " is already allocated/issued for " & vbCrLf & "Model " & Rst!Model & " and Chassis No." & Rst!ChassisNo, vbCritical, "Duplicate Service Book No."
    Txt(SrvBookNo).SetFocus
    Set Rst = Nothing
    Exit Sub
End If
'TransAxlNo checking
If Txt(TransAxlNo) <> "" Then
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open sqlstr, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        MsgBox "Trans Axle No. " & Txt(TransAxlNo) & " is already allocated/issued for " & vbCrLf & "Model " & Rst!Model & " and Chassis No." & Rst!ChassisNo, vbCritical, "Duplicate TransAxle No."
        Txt(TransAxlNo).SetFocus
        Set Rst = Nothing
        Exit Sub
    End If
End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ADItem) <> "" Then
            If Val(FGrid.TextMatrix(I, Qty)) = 0 Then MsgBox "Fill Quantity in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Qty: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
        End If
    Next
    Amt_Cal
    Select Case Txt(FundSource).TEXT
        Case "Hypothication"
            mFundSource = 0
        Case "Hire Purchase"
            mFundSource = 1
        Case "Lease"
            mFundSource = 3
        Case "Agreement"
            mFundSource = 4
        Case "Lease & Agreement"
            mFundSource = 5
        Case "Loan Cum Hypt."
            mFundSource = 6
        Case Else
            mFundSource = 2 'Own Fund
    End Select
    Select Case Txt(ADType).TEXT
        Case "No Detail"
            mTrntypeprn = 0
        Case "Name/Qty"
            mTrntypeprn = 1
        Case "Name/Qty/Amount"
            mTrntypeprn = 2
    End Select
    '********* cHECKING pOSTING cOTROLS
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        If ProcAcPost(True) = False Then Me.ActiveControl.SetFocus: Exit Sub
        Txt(AcPostByName) = pubUName
        Txt(AcPostDate) = PubServerDate
    End If
    '**********
    mDlrID = GCn.Execute("Select Dealer_ID from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    Set Rst = New ADODB.Recordset
    Rst.Open "Select PBILL_NO,PBILL_DATE from Veh_Stock where ChassisNo = '" & Txt(ChassisNo) & "'", GCn, adOpenStatic, adLockReadOnly
    mPBIllNo = IIf(IsNull(Rst!PBILL_NO), "", Rst!PBILL_NO)
    mPBIllDate = IIf(IsNull(Rst!PBILL_DATE), "", Rst!PBILL_DATE)
    Set Rst = Nothing
    
'    If TopCtrl1.TopText2.CAPTION = "Add" Then
'    '   lp 11-03-03
'        DocId = Txt(TxtDocId)
'        If GCn.Execute("select count(*) from veh_order where inv_DocID='" & Txt(TxtDocId) & "'").Fields(0) > 0 Then
'            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
'                MsgBox "Bill No. already exists, Retry", vbCritical, "Validation Error"
'                Txt(SerialNo).SetFocus
'                Exit Sub
'            Else
'                Txt(TxtDocId) = GetDocIDVBill(GCnFaV, mVType, Txt(Vdate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
'                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocId, Document_No)) Then
'                    MsgBox "Bill No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
'                    Exit Sub
'                End If
'            End If
'        End If
'   End If
   DocIdHlp = Replace(Txt(TxtDocId), " ", "")
    GCn.BeginTrans
    GCnFaV.BeginTrans
    mTrans = True
    
    If TopCtrl1.TopText2 = "Add" Then
    '   lp 21-05-03
        DocId = Txt(TxtDocId)
        If GCn.Execute("select count(*) from veh_order where inv_DocID='" & Txt(TxtDocId) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
                MsgBox "Bill No. already exists, Retry", vbCritical, "Validation Error"
                Txt(SerialNo).SetFocus
                GoTo errlbl
            Else
                Txt(TxtDocId) = GetDocIDVBill(GCnFaV, mVType, Txt(Vdate), VoucherEditFlag, Txt(SerialNo), LblVPrefix, Txt(SiteCode).Tag)
                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocId, Document_No)) Then
                    MsgBox "Bill No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If
        DocIdHlp = Replace(Txt(TxtDocId), " ", "")
        '** eof
        GCn.Execute ("update veh_order" & _
            " set Inv_DocId='" & Txt(TxtDocId) & "',Inv_DocIDHelp='" & DocIdHlp & "' ,Inv_SiteCode='" & PubSiteCode & Txt(SiteCode).Tag & "',Inv_VType='" & mVType & "',Inv_No=" & Val(Txt(SerialNo).TEXT) & " ,Inv_Date=" & ConvertDate(Txt(Vdate)) & " ,Form_Code='" & Txt(FormType).Tag & _
            "',TAX_Per=" & Val(Txt(TaxPer)) & ",model = '" & Txt(Model).TEXT & "', TAX_Amt=" & Val(Txt(TaxAmt)) & ",Surcharge_Per=" & Val(Txt(TaxSurPer)) & ",Surcharge_Amt=" & Val(Txt(TaxSurch)) & ",MARGINE=" & (Val(Txt(SaleRate)) - Val(Txt(NDP))) & ",VRATE=" & Val(Txt(NDP)) & ",REBATE=" & Val(Txt(Rebate)) & _
            " ,InciChrg=" & Val(Txt(IncCharge)) & ",Octroi=" & Val(Txt(Octroi)) & " ,RegTemp=" & Val(Txt(TempReg)) & ",TransitInsu=" & Val(Txt(TransIns)) & ",Transport=" & Val(Txt(Transportation)) & ",MVT=" & Val(Txt(MVT)) & ",OtherChrg=" & Val(Txt(MisCharge)) & ",FIT_AMT=" & Val(Txt(OthFitAmt)) & ",FIT_TAX=" & Val(Txt(OthFitTax)) & _
            " ,TOT_Per=" & Val(Txt(TOTPer)) & ",TOT_Amt=" & Val(Txt(TOTAmt)) & ",DieselAmt=" & Val(Txt(FuelAmt)) & ",MISC_INFO='" & Txt(SpclInfo) & "',RTO='" & Txt(RTO) & "',Round_off=" & Val(Txt(ROff)) & _
            " ,FB_Code='" & Txt(FB_Code).Tag & "' , FIN_AcCode='" & FinAcCode & "', FIN_AMT=" & Val(Txt(FinAmt)) & _
            " ,Net_Amount = " & Val(Txt(GTotAmt)) & ", TrnType_Prn=" & mTrntypeprn & ",Fund_Source=" & mFundSource & ",Chassis='" & Txt(ChassisNo) & _
            "',Inv_UName='" & pubUName & "', Inv_UEntDt=#" & PubServerDate & "#, Inv_UAE= 'A' " & _
            " ,Inv_AcPostByUName='" & Txt(AcPostByName) & "',Inv_AcPostByUEntDt=" & ConvertDate(Txt(AcPostDate)) & _
            " where OrdDocId = '" & Txt(BookNo).Tag & "'")
        GCn.Execute "Update Veh_Stock set Sal_DocId = '" & Txt(TxtDocId) & "',Sal_VDate=" & ConvertDate(Txt(Vdate)) & ", Srv_BookNo = '" & Txt(SrvBookNo) & "', TransAxlNo='" & Txt(TransAxlNo) & "' where ChassisNo  = '" & Txt(ChassisNo) & "' and Model='" & Txt(Model) & "'"
    'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
    '    GCn.Execute ("update hiscard set Dealer_Code='" & mDlrID & "', CouponNo='" & txt(SrvBookNo) & "', TransAxelNo='" & txt(TransAxlNo) & "', Supplier_BillNo='" & mPBIllNo & "', Supplier_BillDate=" & ConvertDate(mPBIllDate) & _
    '        ", Name='" & txt(Party) & "',Add1='" & txt(Add1) & "',Add2='" & txt(Add2) & "',Add3='" & txt(Add3) & "',CityCode='" & txt(City).Tag & "',Govt_YN = " & IIf(txt(Govt_YN) = "Yes", 1, 0) & _
    '        " where Chassis ='" & txt(ChassisNo) & "' and Model='" & txt(Model) & "'")
        mQuotDocID = GCn.Execute("select iif(isnull(Quot_DocID),'',Quot_DocID) as QuotID from Veh_Order where OrdDocID = '" & Txt(BookNo).Tag & "'").Fields(0).Value
        mQuotDocIDSrlNo = GCn.Execute("select iif(isnull(QuotSrl_No),'',QuotSrl_No) as QuotSrlNo from Veh_Order where OrdDocID = '" & Txt(BookNo).Tag & "'").Fields(0).Value
        If mQuotDocID <> "" Then
            GCn.Execute "Update Veh_SubGroupQuot set Got_Lost='Got',GotLost_Date=" & ConvertDate(Txt(Vdate)) & " where QuotDocId='" & mQuotDocID & "' and QuotSrl_No=" & mQuotDocIDSrlNo & ""
        End If
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo>Table.SerialNo
        UpdVouSrlNo GCnFaV, Txt(TxtDocId), Txt(Vdate)
    Else
        If Txt(ChassisNo) <> mOldChasis Then
    'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
    '        GCn.Execute ("update hiscard set Dealer_Code='', CouponNo='', TransAxelNo='', Supplier_BillNo='', Supplier_BillDate=Null, Name='" & PubComp_Name & "',Add1='',Add2='',Add3='',CityCode='',Govt_YN =  0 " & _
    '            " where Chassis ='" & txt(ChassisNo).Tag & "' and Model='" & txt(Model).Tag & "'")
        GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_VDate=Null, Srv_BookNo = '', TransAxlNo='' " & _
                "where ChassisNo  = '" & mOldChasis & "' and Sal_DocId = '" & Txt(TxtDocId) & "'"
        
         End If
        GCn.Execute ("update veh_order  " & _
            "set Form_Code='" & Txt(FormType).Tag & "', Inv_Date=" & ConvertDate(Txt(Vdate)) & ", " & _
            "TAX_Per=" & Val(Txt(TaxPer)) & ",model = '" & Txt(Model) & "',TAX_Amt=" & Val(Txt(TaxAmt)) & ",Surcharge_Per=" & Val(Txt(TaxSurPer)) & ",Surcharge_Amt=" & Val(Txt(TaxSurch)) & ",MARGINE=" & (Val(Txt(SaleRate)) - Val(Txt(NDP))) & ",VRATE=" & Val(Txt(NDP)) & ",REBATE=" & Val(Txt(Rebate)) & ", " & _
            "InciChrg=" & Val(Txt(IncCharge)) & ",Octroi=" & Val(Txt(Octroi)) & ",RegTemp=" & Val(Txt(TempReg)) & ",TransitInsu=" & Val(Txt(TransIns)) & ",Transport=" & Val(Txt(Transportation)) & ",MVT=" & Val(Txt(MVT)) & ",OtherChrg=" & Val(Txt(MisCharge)) & ",FIT_AMT=" & Val(Txt(OthFitAmt)) & ",FIT_TAX=" & Val(Txt(OthFitTax)) & ", " & _
            "TOT_Per=" & Val(Txt(TOTPer)) & ",TOT_Amt=" & Val(Txt(TOTAmt)) & ",DieselAmt=" & Val(Txt(FuelAmt)) & ",MISC_INFO='" & Txt(SpclInfo) & "',RTO='" & Txt(RTO) & "',Round_off=" & Val(Txt(ROff)) & ", " & _
            "FB_Code='" & Txt(FB_Code).Tag & "' , FIN_AcCode='" & FinAcCode & "', FIN_AMT=" & Val(Txt(FinAmt)) & ", " & _
            "Net_Amount = " & Val(Txt(GTotAmt)) & ",TrnType_Prn=" & mTrntypeprn & ",Fund_Source=" & mFundSource & ",Chassis='" & Txt(ChassisNo) & "', Srv_BookNo='" & Txt(SrvBookNo) & "', " & _
            "Inv_UName='" & pubUName & "', Inv_UEntDt=#" & PubServerDate & "#, Inv_UAE= 'E', " & _
            "Inv_AcPostByUName='" & Txt(AcPostByName) & "',Inv_AcPostByUEntDt=" & ConvertDate(Txt(AcPostDate)) & _
            " where Inv_DocId='" & Txt(TxtDocId) & "'")
        GCn.Execute "Update Veh_Stock set Sal_DocId = '" & Txt(TxtDocId) & "',Sal_VDate=" & ConvertDate(Txt(Vdate)) & ", Srv_BookNo = '" & Txt(SrvBookNo) & "', TransAxlNo='" & Txt(TransAxlNo) & _
            "' where ChassisNo  = '" & Txt(ChassisNo) & "' and Model='" & Txt(Model) & "'"
    'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
    '    GCn.Execute ("update hiscard set Dealer_Code='" & mDlrID & "', CouponNo='" & txt(SrvBookNo) & "', TransAxelNo='" & txt(TransAxlNo) & "', Supplier_BillNo='" & mPBIllNo & "', Supplier_BillDate=" & ConvertDate(mPBIllDate) & _
            ", Name='" & txt(Party) & "',Add1='" & txt(Add1) & "',Add2='" & txt(Add2) & "',Add3='" & txt(Add3) & "',CityCode='" & txt(City).Tag & "',Govt_YN = " & IIf(txt(Govt_YN) = "Yes", 1, 0) & _
            " where Chassis ='" & txt(ChassisNo) & "' and Model='" & txt(Model) & "'")
    End If
    GCn.Execute ("delete from veh_purch2 where DocId='" & Txt(TxtDocId) & "'")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ADItem) <> "" And Val(FGrid.TextMatrix(I, Qty)) <> 0 Then
            GCn.Execute ("insert into veh_purch2(DocId,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE,TAX_PER,TAX_AMT,TaxSur_Per,TaxSur_AMT, U_Name, U_EntDt, U_AE) " & _
                "values('" & Txt(TxtDocId).TEXT & "'," & I & ",'" & PubSiteCode & Txt(SiteCode).Tag & "','" & mVType & "','" & Txt(SerialNo).TEXT & "', " & _
                "'" & FGrid.TextMatrix(I, ADItemCode) & "','A'," & Val(FGrid.TextMatrix(I, Qty)) & "," & Val(FGrid.TextMatrix(I, Rate)) & "," & Val(FGrid.TextMatrix(I, TaxPer1)) & ", " & _
                "" & Val(FGrid.TextMatrix(I, TaxAmt1)) & "," & Val(FGrid.TextMatrix(I, TaxSurPer1)) & "," & Val(FGrid.TextMatrix(I, TaxSurAmt1)) & ",'" & pubUName & "',#" & PubServerDate & "#,'" & left(TopCtrl1.TopText2.CAPTION, 1) & "')")
        End If
    Next
    'A/c Posting
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        ProcAcPost
    End If
    'EOF of A/c Posting Section
GCnFaV.CommitTrans
GCn.CommitTrans
Set Rst = Nothing
mTrans = False
    Master.Requery
    RSBook.Requery
    Master.FIND "Inv_DocId = '" & Txt(TxtDocId) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(Txt(SerialNo)) > Val(DeCodeDocID(DocId, Document_No)) Then
            MsgBox "Bill No." & Trim(DeCodeDocID(DocId, Document_No)) & " already exists ! " & vbCrLf & "New No. " & Txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select Inv_DocId as SearchCode,Inv_No,Inv_Date,Ord_No,SG.Name,Model,Chassis,Inv_DocId " & _
        " from Veh_Order left join SubGroup SG on Veh_Order.PartyCode=SG.Subcode where left(Inv_DocID,1)='" & PubDivCode & "' and trim(mid(Inv_DocID,4,5)) = '" & mVType & "' Order By Inv_Date Desc"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox Err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox Err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus Txt(Index)
Grid_Hide
Dim I As Integer, XXA() As String  'Rst As ADODB.Recordset,
Select Case Index
    Case InvPrefix
        Set GRs = New ADODB.Recordset
        With GRs
             .CursorLocation = adUseClient
             .Open "SELECT Prefix from VehBill_Counter where Div_Code='" & PubDivCode & "'", GCnFaV, adOpenDynamic, adLockOptimistic
        End With
        Do While Not GRs.EOF
            I = I
            ReDim Preserve XXA(I)
            XXA(I) = GRs!Prefix
            I = I + 1
            GRs.MoveNext
        Loop
        Set mListItem = ListView_Items(ListView, Txt, InvPrefix, XXA, GRs.RecordCount)
        mInvPrefixHt = GRs.RecordCount * 260
        Set GRs = Nothing
        Txt(Index) = ListView.SelectedItem.TEXT
    Case ADType
        ListArray = Array("No Detail", "Name/Qty", "Name/Qty/Amount")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 3)
    Case FundSource
        ListArray = Array("Hypothication", "Hire Purchase", "Own Fund", "Lease", "Agreement", "Lease & Agreement", "Loan Cum Hypt.")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 7)
    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(FB_Code).TEXT = "" Then Exit Sub
        If Txt(FB_Code).TEXT <> rsFin!Name Then
            rsFin.MoveFirst
            rsFin.FIND "name ='" & Txt(FB_Code).TEXT & "'"
        End If
    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or Txt(BookNo).TEXT = "" Then Exit Sub
        If Txt(BookNo).TEXT <> RSBook!Code Then
            RSBook.MoveFirst
            RSBook.FIND "code ='" & Txt(BookNo).TEXT & "'"
        End If
    Case ChassisNo
        If Txt(Model) = "" Then MsgBox "Select Model First", vbInformation, "Validation": Txt(Model).SetFocus: Exit Sub
        '14-05-03 lps
        Set RsChassis = GCn.Execute("SELECT VStk.ChassisNo as code, VStk.EngineNo, VStk.MODEL, VStk.Srv_BookNo, VStk.VRATE, VStk.Colour_Code, " & _
            " ColMast.Col_Desc,VStk.PBILL_NO,VStk.PBILL_DATE, mid(VStk.Pur_DocId,14,8) as PurVNo,VStk.Pur_VDate, VStk.AL_Name,VStk.tax_yn,VStk.RSO_WORK,VStk.INDATE " & _
            " FROM (Veh_Stock VStk LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
            " left join Model on VStk.Model=Model.Model " & _
            " where Model.Div_Code='" & PubDivCode & "' and (VStk.Sal_DocId='" & Txt(TxtDocId) & "' or VStk.Sal_DocId= '' or isnull(VStk.Sal_DocId)) " & _
            " and (Vstk.Pur_VDate<=" & ConvertDate(Txt(Vdate)) & " or VStk.Pur_VDate is null)")
        'eof
        Set DgChassis.DataSource = RsChassis
    Case SiteCode
        Set DGSite.DataSource = RsSite
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If Txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            Txt(Index).Tag = RsSite!Code
            Txt(Index).TEXT = RsSite!Name
        Else
            If Txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & Txt(Index).TEXT & "'"
        End If
'    Case SerialNo, TaxAmt, TaxSurch, TaxPer, TaxSurPer, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation
'        SendKeys "{HOME}+{END}"
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case InvPrefix
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, mInvPrefixHt
    Case BookNo
        DGridTxtKeyDown DGBook, Txt, Index, RSBook, KeyCode, False, 0
    Case ADType
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 900
    Case FundSource
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width + 50, 1600
    Case SiteCode
        DGridTxtKeyDown DGSite, Txt, Index, RsSite, KeyCode, False, 1
'    Case Model
'        DGridTxtKeyDown DGMod, txt, Index, RsMod, KeyCode, False, 0, frmModel
'        If DGMod.Visible = True Then txt(ChassisNo).Text = ""
    Case ChassisNo
        DGridTxtKeyDown DgChassis, Txt, Index, RsChassis, KeyCode, False, 0
    Case FormType
        DGridTxtKeyDown DGForm, Txt, FormType, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
    Case FB_Code
        DGridTxtKeyDown DGFin, Txt, Index, rsFin, KeyCode, False, 1, frmFinMast, "frmFinMast"
    Case Model
        DGridTxtKeyDown DGMod, Txt, Index, RsMod, KeyCode, False, 0, frmModel, "frmModel"
    Case ChassisNo
        DGridTxtKeyDown DgChassis, Txt, Index, RsChassis, KeyCode, False, 0
End Select
If FrmList.Visible = False And DGBook.Visible = False And DGSite.Visible = False And DGFin.Visible = False _
    And DgChassis.Visible = False And DGMod.Visible = False And DGForm.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Vdate Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> SpclInfo Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = SpclInfo Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SiteCode Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> BookNo Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case BookNo
        If DGBook.Visible = True Then DGridTxtKeyPress Txt, Index, RSBook, KeyAscii, "Code"
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress Txt, Index, RsSite, KeyAscii, "Name"
'    Case Model
'        If DGMod.Visible = True Then DGridTxtKeyPress txt, Index, RsMod, KeyAscii, "code"
    Case ChassisNo
        If DgChassis.Visible = True Then DGridTxtKeyPress Txt, Index, RsChassis, KeyAscii, "code", False
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress Txt, FormType, rsForm, KeyAscii, "Name"
    Case SerialNo
        Call NumPress(Txt(Index), KeyAscii, 7, 0)
    Case SaleRate, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation
        Call NumPress(Txt(Index), KeyAscii, 8, 2)
    Case TaxAmt, TaxSurch, MisCharge, FinAmt, TOTAmt
        Call NumPress(Txt(Index), KeyAscii, 7, 2)
    Case FuelAmt
        Call NumPress(Txt(Index), KeyAscii, 4, 2)
    Case TaxPer, TaxSurPer, TOTPer
        Call NumPress(Txt(Index), KeyAscii, 2, 2)
End Select
'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
Select Case Index
    Case InvPrefix, FundSource, ADType
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
    Case FormType
        If DGForm.Visible = True Then
            If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
            Txt(TaxPer).TEXT = IIf(IsNull(rsForm!Tax_Per), 0, rsForm!Tax_Per)
            Txt(TaxAmt).TEXT = Val(Txt(SubTotA).TEXT) * Val(Txt(TaxPer).TEXT) / 100
            Txt(TaxSurPer).TEXT = IIf(IsNull(rsForm!Tax_Sur_Per), 0, rsForm!Tax_Sur_Per)
            Txt(TaxSurch).TEXT = Val(Txt(TaxSurPer).TEXT) * Val(Txt(TaxAmt).TEXT) / 100
            Amt_Cal
        End If
    Case TaxPer, TaxSurPer, TOTPer
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
        Amt_Cal
    Case SaleRate, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation
         Amt_Cal
    Case MisCharge, FinAmt, AdvAmt, FuelAmt
         Amt_Cal
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Integer
Select Case Index
    Case FundSource, ADType
        If Txt(Index) <> "" Then Txt(Index) = ListView.SelectedItem.TEXT
    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RSBook!Code
            Txt(Index).Tag = RSBook!OrdDocId
        End If
        Cancel = Not FillRecords '= False Then  = True
'    Case Model
'        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or txt(Index).Text = "" Then
'            txt(Index).Text = ""
'            txt(Index).Tag = ""
'        Else
'            txt(Index).Text = RsMod!Code
'            txt(Index).Tag = RsMod!Code
'        End If
'        If IsValid(txt(Index), "Model") = False Then Cancel = True: GoTo lblExitSub
'        txt(ChassisNo).SetFocus
    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
            FinAcCode = ""
        Else
            Txt(Index).TEXT = rsFin!Name
            Txt(Index).Tag = rsFin!Code
            FinAcCode = rsFin!AcCode
        End If
    Case ChassisNo
        IsValid Txt(Index), "Chassis No.", True
        If RsChassis.RecordCount = 0 Or (RsChassis.EOF = True Or RsChassis.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(ChassisNo) = ""
            Cancel = Not Fill_Data(False)
        Else
            Txt(ChassisNo) = RsChassis!Code
            Cancel = Not Fill_Data(True)
        End If
'        If Txt(ChassisNo).Text = "" And RsChassis.RecordCount > 0 Then
'            MsgBox "chassis no is required", vbInformation, "Validation  Check"
'            Cancel = True: GoTo lblExitSub
'        End If
    Case SiteCode
        If IsValid(Txt(Index), "Site Code") = False Then Cancel = True: GoTo lblExitSub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsSite!Name
            Txt(Index).Tag = RsSite!Code
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = rsForm!Name
            Txt(Index).Tag = rsForm!Code
        End If
    Case Vdate
        If Len(Trim(Txt(Vdate).TEXT)) = 0 Then
            Txt(Vdate).TEXT = PubLoginDate
        Else
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
        If CheckFinYear(Txt(Index)) Then
'            txt(TxtDocId) = GetDocIDVBill(GCnFaV, mVType, txt(Vdate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
'            DocId = txt(TxtDocId)
        Else
            Cancel = True
        End If
    Case InvPrefix
        If Txt(Index) <> "" Then Txt(Index) = ListView.SelectedItem.TEXT
        LblVPrefix = Txt(Index)
        Txt(TxtDocId) = GetDocIDVBill(GCnFaV, mVType, Txt(Vdate), VoucherEditFlag, Txt(SerialNo), LblVPrefix, Txt(SiteCode).Tag)
        DocId = Txt(TxtDocId)
    Case SerialNo
        If IsValid(Txt(SerialNo), "Serial No.") = False Then Cancel = True:  GoTo lblExitSub
        'If VoucherEditFlag Then      ' Manual
            Txt(TxtDocId) = GetDocIDVBill(GCnFaV, mVType, Txt(Vdate), VoucherEditFlag, Txt(SerialNo), LblVPrefix, Txt(SiteCode).Tag)
            DocId = Txt(TxtDocId)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select Inv_DocId From veh_Order Where Inv_DocID='" & Txt(TxtDocId) & "'", GCn, adOpenStatic, adLockReadOnly
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Set Rst = Nothing
                Cancel = True
                Txt(SerialNo).SetFocus
         'End If
        End If
    Case Rebate
         Amt_Cal
         Txt(Index).TEXT = IIf(Val(Txt(Index)) <> 0, Format(Txt(Index), "0.00"), "")
         If Val(Txt(Index)) > Val(Txt(SaleRate)) Then
            MsgBox "Rebate Rs." & Txt(Rebate) & " is greater than Sale Rate Rs." & Txt(SaleRate), vbOKOnly, "Validation"
            Cancel = True
        End If
    Case TaxPer, TaxAmt, TaxSurPer, TaxSurch, SaleRate, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation, MisCharge, OthFitAmt, OthFitTax, FinAmt, AdvAmt, FuelAmt
         Txt(Index).TEXT = IIf(Val(Txt(Index)) <> 0, Format(Txt(Index), "0.00"), "")
         Amt_Cal
End Select
lblExitSub:
Set Rst = Nothing
End Sub

Private Sub DGADItem_Click()
    If RsADItem.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsADItem!Name
         FGrid.TextMatrix(FGrid.Row, ADItem) = RsADItem!Name
         FGrid.TextMatrix(FGrid.Row, ADItemCode) = RsADItem!Code
    End If
    TxtGrid(0).SetFocus
    DGADItem.Visible = False
End Sub

Private Sub DgChassis_Click()
    If RsChassis.RecordCount > 0 Then
        Txt(ChassisNo).TEXT = RsChassis!Code
        Fill_Data True
    End If
    Txt(ChassisNo).SetFocus
    DgChassis.Visible = False
End Sub

Private Sub DGMod_Click()
If RsMod.RecordCount > 0 Then
    Txt(Model) = RsMod!Code
End If
Txt(Model).SetFocus
DGMod.Visible = False
End Sub

Private Sub DGForm_Click()
    If rsForm.RecordCount > 0 Then
        Txt(FormType).TEXT = rsForm!Name
        Txt(FormType).Tag = rsForm!Code
    End If
    Txt(FormType).SetFocus
    DGForm.Visible = False
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
Next I
Txt(Model).Tag = ""
Txt(ChassisNo).Tag = ""
End Sub

Private Sub MoveRec()
Dim Rst As Recordset
Dim I As Integer
On Error GoTo error1
Grid_Hide
If Master.RecordCount > 0 Then
    DocId = Master!Inv_DocId
    Txt(TxtDocId) = Master!Inv_DocId
    LblDiv.CAPTION = "Division : " & left(Master!Inv_DocId, 1)
    LblSite.CAPTION = "Site Code : " & mID(Master!Inv_SiteCode, 1, 1)
    Txt(SiteCode).Tag = mID(Master!Inv_SiteCode, 2, 1)
    Txt(SiteCode).TEXT = GCn.Execute("select site_desc from site where site_code = '" & Txt(SiteCode).Tag & "'").Fields(0).Value
    LblVPrefix.CAPTION = mID(Master!Inv_DocId, 8, 5)
    Txt(InvPrefix) = DeCodeDocID(Master!Inv_DocId, Document_Prefix)
    Txt(SerialNo).TEXT = Master!Inv_No
    Txt(Vdate).TEXT = Master!Inv_Date
    Txt(BookNo).TEXT = Master!Ord_No
    Txt(BookNo).Tag = Master!OrdDocId
    Txt(DelChNo) = DeCodeDocID(Master!DelCh_DocId, Document_No)
    Txt(DelChDate) = IIf(IsNull(Master!DelCh_DT), "", Master!DelCh_DT)
    '*** A/c Posting Status
    Txt(AcPostByName) = IIf(IsNull(Master!Inv_AcPostByUName), "", Master!Inv_AcPostByUName)
    Txt(AcPostDate) = IIf(IsNull(Master!Inv_AcPostByUEntDt), "", Master!Inv_AcPostByUEntDt)
    '***
    If Not IsNull(Master!Fund_Source) Then
        Select Case Master!Fund_Source
            Case 0 '0 Hypothication ,1 Hire purchase ,2 Own Fund,3 Lease
                Txt(FundSource).TEXT = "Hypothication"
            Case 1
                Txt(FundSource).TEXT = "Hire Purchase"
'            Case 2
'                txt(FundSource).Text = "Own Fund"
            Case 3
                Txt(FundSource).TEXT = "Lease"
            Case 4
                Txt(FundSource).TEXT = "Agreement"
            Case 5
                Txt(FundSource).TEXT = "Lease & Agreement"
            Case 6
                Txt(FundSource).TEXT = "Loan Cum Hypt."
            Case Else
                Txt(FundSource).TEXT = "Own Fund"
        End Select
    Else
        Txt(FundSource).TEXT = ""
    End If
    If Not IsNull(Master!TrnType_Prn) Then
        Select Case Master!TrnType_Prn
            Case 0
                Txt(ADType).TEXT = "No Detail"
            Case 1
                Txt(ADType).TEXT = "Name/Qty"
            Case 2
                Txt(ADType).TEXT = "Name/Qty/Amount"
        End Select
    Else
        Txt(ADType).TEXT = ""
    End If
    
    Txt(Party).Tag = IIf(IsNull(Master!PartyCode), "", Master!PartyCode)
    If Txt(Party).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NamePrefix,name,FPrefix,FName,add1,add2,add3,CityCode from SubGroup where Subcode = '" & Txt(Party).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        Txt(NamePrefix).TEXT = IIf(IsNull(Rst!NamePrefix), "", Rst!NamePrefix)
        Txt(Party).TEXT = Rst!Name
        Txt(FNamePrefix).TEXT = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
        Txt(fname).TEXT = IIf(IsNull(Rst!fname), "", Rst!fname)
        Txt(Add1).TEXT = IIf(IsNull(Rst!Add1), "", Rst!Add1)
        Txt(Add2).TEXT = IIf(IsNull(Rst!Add2), "", Rst!Add2)
        Txt(Add3).TEXT = IIf(IsNull(Rst!Add3), "", Rst!Add3)
        Txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
        If Txt(City).Tag <> "" Then
            Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
        End If
    End If
    Txt(Model).TEXT = Master!Model
    Txt(Model).Tag = Master!Model
    Txt(Govt_YN).TEXT = IIf(Master!Govt_YN = 1, "Yes", "No")
    Txt(Colours).Tag = IIf(IsNull(Master!Colour_Code), "", Master!Colour_Code)
    If Txt(Colours).Tag <> "" Then
        Txt(Colours).TEXT = GCn.Execute("select col_desc from colmast where col_code = '" & Txt(Colours).Tag & "'").Fields(0).Value
    End If
    
    Txt(FormType).Tag = IIf(IsNull(Master!Form_Code), "", Master!Form_Code)
    If Txt(FormType).Tag <> "" Then
        Txt(FormType).TEXT = GCn.Execute("select form_desc from taxforms where form_code = '" & Txt(FormType).Tag & "'").Fields(0).Value
    Else
        Txt(FormType).TEXT = ""
    End If
    Txt(NDP).TEXT = IIf(IsNull(Master!vrate) Or Master!vrate = 0, "", Format(Master!vrate, "0.00"))
    
    Txt(SaleRate).TEXT = Format(Val(Txt(NDP)) + IIf(IsNull(Master!Margine), 0, Master!Margine), "0.00")
    Txt(Rebate).TEXT = IIf(IsNull(Master!Rebate) Or Master!Rebate = 0, "", Format(Master!Rebate, "0.00"))
    Txt(IncCharge).TEXT = IIf(IsNull(Master!InciChrg) Or Master!InciChrg = 0, "", Format(Master!InciChrg, "0.00"))
    Txt(Octroi).TEXT = IIf(IsNull(Master!Octroi) Or Master!Octroi = 0, "", Format(Master!Octroi, "0.00"))
    Txt(TempReg).TEXT = IIf(IsNull(Master!RegTemp) Or Master!RegTemp = 0, "", Format(Master!RegTemp, "0.00"))
    Txt(TransIns).TEXT = IIf(IsNull(Master!TransitInsu) Or Master!TransitInsu = 0, "", Format(Master!TransitInsu, "0.00"))
    Txt(MVT).TEXT = IIf(IsNull(Master!MVT) Or Master!MVT = 0, "", Format(Master!MVT, "0.00"))
    Txt(Transportation).TEXT = IIf(IsNull(Master!Transport) Or Master!Transport = 0, "", Format(Master!Transport, "0.00"))
    Txt(SubTotA) = Format((Val(Txt(SaleRate)) - Val(Txt(Rebate)) + Val(Txt(IncCharge)) + Val(Txt(Octroi)) + Val(Txt(TempReg)) + Val(Txt(TransIns)) + Val(Txt(MVT)) + Val(Txt(Transportation))), "0.00")
    
    Txt(TaxPer).TEXT = IIf(IsNull(Master!Tax_Per) Or Master!Tax_Per = 0, "", Format(Master!Tax_Per, "0.00"))
    Txt(TaxAmt).TEXT = IIf(IsNull(Master!Tax_Amt) Or Master!Tax_Amt = 0, "", Format(Master!Tax_Amt, "0.00"))
    Txt(TaxSurPer).TEXT = IIf(IsNull(Master!surcharge_per) Or Master!surcharge_per = 0, "", Format(Master!surcharge_per, "0.00"))
    Txt(TaxSurch).TEXT = IIf(IsNull(Master!Surcharge_Amt) Or Master!Surcharge_Amt = 0, "", Format(Master!Surcharge_Amt, "0.00"))
    Txt(MisCharge).TEXT = IIf(IsNull(Master!OtherChrg) Or Master!OtherChrg = 0, "", Format(Master!OtherChrg, "0.00"))
    Txt(SubTotB) = Format((Val(Txt(SubTotA)) + Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(MisCharge))), "0.00")
        
    Txt(OthFitAmt).TEXT = IIf(IsNull(Master!Fit_Amt) Or Master!Fit_Amt = 0, "", Format(Master!Fit_Amt, "0.00"))
    Txt(OthFitTax).TEXT = IIf(IsNull(Master!Fit_Tax) Or Master!Fit_Tax = 0, "", Format(Master!Fit_Tax, "0.00"))
    Txt(TOTPer) = IIf(IsNull(Master!TOT_Per) Or Master!TOT_Per = 0, "", Format(Master!TOT_Per, "0.00"))
    Txt(TOTAmt) = IIf(IsNull(Master!Tot_Amt) Or Master!Tot_Amt = 0, "", Format(Master!Tot_Amt, "0.00"))
    Txt(FuelAmt).TEXT = IIf(IsNull(Master!DieselAmt) Or Master!DieselAmt = 0, "", Format(Master!DieselAmt, "0.00"))
    Txt(ROff).TEXT = IIf(IsNull(Master!Round_off) Or Master!Round_off = 0, "", Format(Master!Round_off, "0.00"))
    'Modi LPS 05.12.2003
    Txt(GTotAmt) = Format((Val(Txt(SubTotB)) + Val(Txt(OthFitAmt)) + Val(Txt(OthFitTax)) + Val(Txt(TOTAmt)) - Val(Txt(FuelAmt)) + Val(Txt(ROff))), "0.00")
    'eof Modi
    
    'modified for docid / invdate by lps
    Txt(AdvAmt) = IIf(PartyAdvance(Master!OrdDocId, Txt(Vdate)) <> 0, Format(PartyAdvance(Master!OrdDocId, Txt(Vdate)), "0.00"), "")
  ' Txt(AdvAmt) = Format(IIf(IsNull(Master!P_Amount), 0, Master!P_Amount), "0.00")
  ' end modi
    Txt(NetOStng) = Format((Val(Txt(GTotAmt)) - Val(Txt(AdvAmt))), "0.00")
    Txt(FinAmt).TEXT = IIf(IsNull(Master!FIN_AMT) Or Master!FIN_AMT = 0, "", Format(Master!FIN_AMT, "0.00"))
    Txt(FB_Code).Tag = IIf(IsNull(Master!FB_Code), "", Master!FB_Code)
    If Txt(FB_Code).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select fincode as code,finname & ',' & City.CityName as name,AcCode " & _
        " from ContractFinance left join city on left(ContractFinance.City,4)=City.CityCode " & _
        " where fincatg = 0 and  fincode = '" & Txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        Txt(FB_Code).TEXT = Rst!Name
        FinAcCode = IIf(IsNull(Rst!AcCode), "", Rst!AcCode)
    Else
        Txt(FB_Code).TEXT = ""
        FinAcCode = ""
    End If
    
    Txt(SpclInfo).TEXT = IIf(IsNull(Master!MISC_INFO), "", Master!MISC_INFO)
    Txt(RTO).TEXT = IIf(IsNull(Master!RTO), "", Master!RTO)
    Txt(ChassisNo).TEXT = IIf(IsNull(Master!Chassis), "", Master!Chassis)
    Txt(ChassisNo).Tag = IIf(IsNull(Master!Chassis), "", Master!Chassis)
    mOldChasis = Txt(ChassisNo).Tag
    Set Rst = New Recordset
    Rst.Open "SELECT Veh_Stock.TransAxlNo,Veh_Stock.Srv_BookNo,Veh_Stock.EngineNo,Veh_Stock.VehSerialNo,Veh_Stock.tax_yn,Veh_Stock.PBILL_NO,Veh_Stock.PBILL_DATE FROM Veh_Stock where Veh_Stock.MODEL  = '" & Txt(Model) & "' and Veh_Stock.ChassisNo = '" & Txt(ChassisNo) & "' and Veh_Stock.Sal_DocId= '" & Master!Inv_DocId & "'", GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        Txt(TransAxlNo).TEXT = IIf(IsNull(Rst!TransAxlNo), "", Rst!TransAxlNo)
        Txt(SrvBookNo).TEXT = IIf(IsNull(Rst!Srv_BookNo), "", Rst!Srv_BookNo)
        Txt(EngineNo).TEXT = IIf(IsNull(Rst!EngineNo), "", Rst!EngineNo)
        Txt(TelcoInvNo).TEXT = IIf(IsNull(Rst!PBILL_NO), "", Rst!PBILL_NO)
        Txt(TelcoInvDate).TEXT = IIf(IsNull(Rst!PBILL_DATE), "", Rst!PBILL_DATE)
        Txt(Taxable).TEXT = IIf(Rst!Tax_YN = 1, "Yes", "No")
    Else
        Txt(TransAxlNo).TEXT = ""
        Txt(SrvBookNo).TEXT = ""
        Txt(EngineNo).TEXT = ""
        Txt(TelcoInvNo).TEXT = ""
        Txt(TelcoInvDate).TEXT = ""
        Txt(Taxable).TEXT = ""
    End If
    Set Rst = New Recordset
    Set Rst = GCn.Execute("SELECT Veh_AMDModel.Prod_Name, Veh_Purch2.Srl_No, Veh_Purch2.PROD_CODE, Veh_Purch2.QTY, Veh_Purch2.RATE,Veh_Purch2.TAX_PER,Veh_Purch2.TAX_AMT,Veh_Purch2.TaxSur_Per,Veh_Purch2.TaxSur_AMT " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_purch2.DocId = '" & Master!Inv_DocId & "'")
    FGrid.Rows = 1
    I = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            With FGrid
                .AddItem ""
                .TextMatrix(I, 0) = Rst!Srl_No
                .TextMatrix(I, ADItem) = Rst!Prod_Name
                .TextMatrix(I, Qty) = Format(IIf(IsNull(Rst!Qty), "", Rst!Qty), "0")
                .TextMatrix(I, Rate) = Format(IIf(IsNull(Rst!Rate), "", Rst!Rate), "0.00")
                .TextMatrix(I, Amt) = Format(.TextMatrix(I, Qty) * .TextMatrix(I, Rate), "0.00")
                .TextMatrix(I, TaxPer1) = Format(IIf(IsNull(Rst!Tax_Per), "", Rst!Tax_Per), "0.00")
                .TextMatrix(I, TaxAmt1) = Format(IIf(IsNull(Rst!Tax_Amt), "", Rst!Tax_Amt), "0.00")
                .TextMatrix(I, TaxSurPer1) = Format(IIf(IsNull(Rst!TaxSur_Per), "", Rst!TaxSur_Per), "0.00")
                .TextMatrix(I, TaxSurAmt1) = Format(IIf(IsNull(Rst!TaxSur_Amt), "", Rst!TaxSur_Amt), "0.00")
                .TextMatrix(I, FinalAmt) = Format((Val(.TextMatrix(I, Amt)) + Val(.TextMatrix(I, TaxAmt1)) + Val(.TextMatrix(I, TaxSurAmt1))), "0.00")
                .TextMatrix(I, ADItemCode) = Rst!Prod_Code
            End With
            Rst.MoveNext
           I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Set Rst = Nothing
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
'lp 10-03-03
'Amt_Cal
Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte
    
    With FGrid
        .left = Me.left '+45
        .top = 3345
        .Cols = 11
'        .BackColor = CellBackColLeave
'        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight

        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, ADItem) = "Accessories / Additional Fitments"
        .ColAlignment(ADItem) = flexAlignLeftCenter
        .ColWidth(ADItem) = 3720
        
       
        .TextMatrix(0, Qty) = "Qty"
        .ColAlignmentFixed(Qty) = flexAlignRightCenter
        .ColWidth(Qty) = 540

        .TextMatrix(0, Rate) = "Rate"
        .ColAlignmentFixed(Rate) = flexAlignRightCenter
        .ColWidth(Rate) = 855
        
        .TextMatrix(0, Amt) = "Amount"
        .ColAlignmentFixed(Amt) = flexAlignRightCenter
        .ColWidth(Amt) = 1065
        
        .TextMatrix(0, TaxPer1) = "Tax%"
        .ColAlignmentFixed(TaxPer1) = flexAlignRightCenter
        .ColWidth(TaxPer1) = 690
        
        .TextMatrix(0, TaxAmt1) = "TaxAmt"
        .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
        .ColWidth(TaxAmt1) = 990
        
        .TextMatrix(0, TaxSurPer1) = "Surch%"
        .ColAlignmentFixed(TaxSurPer1) = flexAlignRightCenter
        .ColWidth(TaxSurPer1) = 720
        
        .TextMatrix(0, TaxSurAmt1) = "SurchAmt"
        .ColAlignmentFixed(TaxSurAmt1) = flexAlignRightCenter
        .ColWidth(TaxSurAmt1) = 990
  
        .TextMatrix(0, FinalAmt) = "NetAmt"
        .ColAlignmentFixed(FinalAmt) = flexAlignRightCenter
        .ColWidth(FinalAmt) = 1065
        .ColWidth(ADItemCode) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel

DGSite.left = 4260: DGSite.top = mTopScale
DGForm.left = Me.width - (DGForm.width + mRtScale): DGForm.top = mTopScale
DGADItem.left = Me.width - (DGADItem.width + mRtScale): DGADItem.top = mTopScale
DGBook.left = 0: DGBook.top = FGrid.top: DGBook.width = Me.width - 90: DGBook.height = Me.height - (DGBook.top + mBotScale)
DGMod.left = 0: DGMod.width = Me.width - 90: DGMod.top = FGrid.top: DGMod.height = Me.height - (DGMod.top + mBotScale)
DgChassis.left = 0: DgChassis.width = Me.width - 90: DgChassis.top = FGrid.top: DgChassis.height = Me.height - (DgChassis.top + mBotScale)
With DgChassis
    .Columns(0).width = 1769.953
    .Columns(1).width = 2055.118
    .Columns(2).width = 1709.858
    .Columns(3).width = 1019.906
    .Columns(4).width = 1184.882
    .Columns(5).width = 1275.024
    .Columns(6).width = 929.7639
    .Columns(7).width = 1230.236
    .Columns(8).width = 0
    .Columns(9).width = 2039.811
End With
DGFin.left = Me.width - (DGFin.width + mRtScale): DGFin.top = mTopScale

End Sub
Private Sub Disp_Text(Enb As Boolean)

Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    Txt(I).ForeColor = CtrlFColOrg
Next

If TopCtrl1.TopText2 = "Edit" Then
    Txt(SiteCode).Enabled = False
    Txt(Vdate).Enabled = True
    Txt(SerialNo).Enabled = False
    Txt(InvPrefix).Enabled = False
    Txt(BookNo).Enabled = False
    If GCn.Execute("select DelCh_DocId from veh_stock where chassisNo ='" & Txt(ChassisNo) & "'").Fields(0).Value = "" Then
'        Txt(Model).Enabled = True
        Txt(ChassisNo).Enabled = True
    Else
'        Txt(Model).Enabled = False
        Txt(ChassisNo).Enabled = False
    End If
End If
If TopCtrl1.TopText2 = "Add" Then
    Txt(SerialNo).Enabled = True
    Txt(InvPrefix).Enabled = True
End If

Txt(TxtDocId).Enabled = False
Txt(Taxable).Enabled = False
Txt(NamePrefix).Enabled = False
Txt(Party).Enabled = False
Txt(FNamePrefix).Enabled = False
Txt(fname).Enabled = False
Txt(Add1).Enabled = False
Txt(Add2).Enabled = False
Txt(Add3).Enabled = False
Txt(City).Enabled = False
Txt(Govt_YN).Enabled = False
Txt(TelcoInvNo).Enabled = False
Txt(TelcoInvDate).Enabled = False
Txt(Model).Enabled = False
Txt(EngineNo).Enabled = False
Txt(Colours).Enabled = False
Txt(NDP).Enabled = False
Txt(SubTotA).Enabled = False
Txt(OthFitAmt).Enabled = False
Txt(OthFitTax).Enabled = False
Txt(SubTotB).Enabled = False
Txt(TaxPer).Enabled = False: Txt(TaxAmt).Enabled = False
Txt(TaxSurPer).Enabled = False: Txt(TaxSurch).Enabled = False
Txt(ROff).Enabled = False
Txt(GTotAmt).Enabled = False
Txt(NetOStng).Enabled = False
Txt(DelChNo).Enabled = False
Txt(DelChDate).Enabled = False

txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGBook.Visible = True Then DGBook.Visible = False
    If DGFin.Visible = True Then DGFin.Visible = False
    If DGMod.Visible = True Then DGMod.Visible = False
    If DgChassis.Visible = True Then DgChassis.Visible = False
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGADItem.Visible = True Then DGADItem.Visible = False
    If DGSite.Visible = True Then DGSite.Visible = False
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ADItemCode) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, ADItem) <> RsADItem!Code Then
                RsADItem.MoveFirst
                RsADItem.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, ADItemCode) & "'"
            End If
        Case Qty, TaxPer1, TaxAmt1, TaxSurPer1, TaxSurAmt1
'            SendKeys "{HOME}+{END}"
     End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TxtGrid(0) = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case ADItem    '1
            DGridTxtKeyDown DGADItem, TxtGrid, Index, RsADItem, KeyCode, True, 1, frmVehAMDMast, "frmVehAMDMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, TaxSurAmt1
                End If
            End If
        Case Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, TaxSurAmt1
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case ADItem
        If DGADItem.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsADItem, KeyAscii, "name"
    Case Rate, TaxAmt1, TaxSurAmt1
        Call NumPress(TxtGrid(Index), KeyAscii, 8, 2)
    Case TaxPer1, TaxSurPer1
        Call NumPress(TxtGrid(Index), KeyAscii, 2, 2)
    Case Qty
        Call NumPress(TxtGrid(Index), KeyAscii, 6, 0)
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
        Select Case FGrid.Col
            Case ADItem
                If KeyCode <> 13 And DGADItem.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsADItem, KeyCode, "name", True
            Case Qty
                FGrid.TextMatrix(FGrid.Row, Qty) = Format(Val(TxtGrid(Index).TEXT), "0")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxSurPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case Rate
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxSurPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxAmt1
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                Else
                    FGrid.TextMatrix(FGrid.Row, TaxPer1) = Format((100 * Val(TxtGrid(Index).TEXT)) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxPer1
                FGrid.TextMatrix(FGrid.Row, TaxPer1) = TxtGrid(Index).TEXT
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxSurAmt1
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                Else
                   FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = Format((100 * Val(TxtGrid(Index).TEXT)) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxSurPer1
                FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = TxtGrid(Index).TEXT
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
        End Select
        If KeyCode = vbKeyEscape Then
            FGrid.SetFocus
            TxtGrid(0).Visible = False
            Grid_Hide
        End If
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case ADItem
        If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, ADItem) = ""
            FGrid.TextMatrix(FGrid.Row, ADItemCode) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, ADItemCode) = RsADItem!Code
            FGrid.TextMatrix(FGrid.Row, ADItem) = RsADItem!Name
            FGrid.TextMatrix(FGrid.Row, Rate) = Format(IIf(IsNull(RsADItem!Rate), 0, RsADItem!Rate), "0.00")
        End If
        If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
    Case Rate, TaxPer1, TaxAmt1, TaxSurPer1, TaxSurAmt1
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
        Amt_Cal
    Case Qty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0")
        Amt_Cal
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
'    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell--> Enter Cell-->KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    SendKeys vbTab
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case ADItem
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case Qty
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Amt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, FinalAmt) = ""
        Case Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Amt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, FinalAmt) = ""
        Case TaxPer1, TaxAmt1
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxPer1) = ""
        Case TaxSurPer1, TaxSurAmt1
            FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = ""
    End Select
    Amt_Cal
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case ADItem, Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.Rows > 2 Then
                FGrid.RemoveItem (FGrid.Row)
            Else
                FGrid.Rows = 1
                FGrid.AddItem FGrid.Rows
                FGrid.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid.Rows - 1
            FGrid.TextMatrix(I, 0) = I
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case ADItem
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case Amt
        FGrid.Col = FGrid.Col + 1
    Case Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
       Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub Amt_Cal()
Dim I As Byte
Dim Tottax As Double
Dim TotAdd As Double, SubTotC As Double, TurnOverTax As Double
For I = 1 To FGrid.Rows - 1
   If FGrid.TextMatrix(I, ADItem) <> "" Then
        TotAdd = TotAdd + Val(FGrid.TextMatrix(I, Amt))
        Tottax = Tottax + Val(FGrid.TextMatrix(I, TaxAmt1)) + Val(FGrid.TextMatrix(I, TaxSurAmt1))
   End If
Next
Txt(OthFitAmt) = IIf(TotAdd <> 0, Format(TotAdd, "0.00"), "")
Txt(OthFitTax) = IIf(Tottax <> 0, Format(Tottax, "0.00"), "")
Txt(SubTotA) = Format((Val(Txt(SaleRate)) - Val(Txt(Rebate)) + Val(Txt(IncCharge)) + Val(Txt(Octroi)) + Val(Txt(TempReg)) + Val(Txt(TransIns)) + Val(Txt(MVT)) + Val(Txt(Transportation))), "0.00")
Txt(TaxAmt) = Format(Val(Txt(SubTotA)) * Val(Txt(TaxPer)) / 100, "0.00")
Txt(TaxSurch) = Format(Val(Txt(TaxSurPer)) * Val(Txt(TaxAmt)) / 100, "0.00")
Txt(SubTotB) = Format((Val(Txt(SubTotA)) + Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(MisCharge))), "0.00")
SubTotC = Val(Txt(SubTotB)) + Val(Txt(OthFitAmt)) + Val(Txt(OthFitTax))
Txt(TOTAmt) = Format(SubTotC * Val(Txt(TOTPer)) / 100, "0.00")
'txt(ROff) = dmRoundOff(Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)))
Txt(ROff) = dmRoundOff(SubTotC + Val(Txt(TOTAmt)))
'txt(GTotAmt) = Format((Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)) + Val(txt(ROff))), "0.00")
Txt(GTotAmt) = Format((SubTotC + Val(Txt(TOTAmt)) - Val(Txt(FuelAmt)) + Val(Txt(ROff))), "0.00")
Txt(NetOStng) = Format((Val(Txt(GTotAmt)) - Val(Txt(AdvAmt))), "0.00")
End Sub

Private Function Fill_Data(Enb As Boolean) As Boolean
Dim Rst As ADODB.Recordset
Dim Margin As Double
If Enb Then
    If Txt(Model) <> RsChassis!Model Then
        If MsgBox("Model changed Continue Yes/No ? ", vbYesNo + vbCritical + vbDefaultButton2, "Check") = vbNo Then
            GoTo NXT
        Else
            Txt(Model) = RsChassis!Model
        End If
    End If
    If IsNull(RsChassis!InDate) Then
        If MsgBox("Vehicle In Transit Continue Yes/No ? ", vbYesNo + vbCritical + vbDefaultButton2, "Check") = vbNo Then
            GoTo NXT
        End If
    End If
    Txt(EngineNo) = IIf(IsNull(RsChassis!EngineNo), "", RsChassis!EngineNo)
    Txt(SrvBookNo) = IIf(IsNull(RsChassis!Srv_BookNo), "", RsChassis!Srv_BookNo)
    Txt(Colours) = IIf(IsNull(RsChassis!Col_Desc), "", RsChassis!Col_Desc)
    Txt(Colours).Tag = IIf(IsNull(RsChassis!Colour_Code), "", RsChassis!Colour_Code)
    Txt(TelcoInvDate).TEXT = IIf(IsNull(RsChassis!PBILL_DATE), "", RsChassis!PBILL_DATE)
    Txt(TelcoInvNo).TEXT = IIf(IsNull(RsChassis!PBILL_NO), "", RsChassis!PBILL_NO)
    Txt(Taxable).TEXT = IIf(RsChassis!Tax_YN = 1, "Yes", "No")
    Txt(NDP).TEXT = IIf(IsNull(RsChassis!vrate) Or RsChassis!vrate = 0, "", Format(RsChassis!vrate, "0.00"))
    Set Rst = New Recordset
    Rst.Open "Select P_RATE,s_rate,INCI_CHRG,OCTROI,REG_TEMP,INS_TRN,TRANSPORT,MVT,REG_FEE,INS_FEE from veh_rate where model = '" & Txt(Model).TEXT & "' and Effective_Date <= " & ConvertDate(Txt(Vdate)) & " and RSO_WORK = " & RsChassis!RSO_WORK & " and TAXABLE_YN = " & RsChassis!Tax_YN & " order by Effective_Date DESC", GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
         If Val(Txt(NDP).TEXT) = 0 Then
            Txt(NDP).TEXT = IIf(IsNull(Rst!p_rate) Or Rst!p_rate = 0, "", Format(Rst!p_rate, "0.00"))
         End If
         Margin = IIf(IsNull(Rst!S_Rate), 0, Rst!S_Rate) - IIf(IsNull(Rst!p_rate), 0, Rst!p_rate)
         Txt(SaleRate).TEXT = Format((Val(Txt(NDP)) + Margin), "0.00")
    Else
         Margin = 0
         Txt(SaleRate).TEXT = Format((Val(Txt(NDP)) + Margin), "0.00")
    End If
    If Rst.RecordCount > 0 Then
        Txt(IncCharge) = IIf(IsNull(Rst!INCI_CHRG) Or Rst!INCI_CHRG = 0, "", Format(Rst!INCI_CHRG, "0.00"))
        Txt(Octroi) = IIf(IsNull(Rst!Octroi) Or Rst!Octroi = 0, "", Format(Rst!Octroi, "0.00"))
        Txt(TempReg) = IIf(IsNull(Rst!REG_TEMP) Or Rst!REG_TEMP = 0, "", Format(Rst!REG_TEMP, "0.00"))
        Txt(TransIns) = IIf(IsNull(Rst!INS_TRN) Or Rst!INS_TRN = 0, "", Format(Rst!INS_TRN, "0.00"))
        Txt(MVT) = IIf(IsNull(Rst!MVT) Or Rst!MVT = 0, "", Format(Rst!MVT, "0.00"))
        Txt(Transportation) = IIf(IsNull(Rst!Transport) Or Rst!Transport = 0, "", Format(Rst!Transport, "0.00"))
    Else
        Txt(IncCharge) = ""
        Txt(Octroi) = ""
        Txt(TempReg) = ""
        Txt(TransIns) = ""
        Txt(MVT) = ""
        Txt(Transportation) = ""
    End If
    Amt_Cal
    Fill_Data = True
    Exit Function
End If
Fill_Data = True
NXT:
    Txt(EngineNo) = ""
    Txt(SrvBookNo) = ""
    Txt(Colours) = ""
    Txt(Colours).Tag = ""
    Txt(TelcoInvDate).TEXT = ""
    Txt(TelcoInvNo).TEXT = ""
    Txt(Taxable).TEXT = ""
    Txt(NDP).TEXT = ""
    Txt(SaleRate).TEXT = ""
    Amt_Cal
End Function

Private Function FillRecords() As Boolean
On Error GoTo error1
Dim Rst As ADODB.Recordset
Dim RsBooking  As ADODB.Recordset
    Set RsBooking = New Recordset
    RsBooking.CursorLocation = adUseClient
    RsBooking.Open "SELECT Veh_Order.OrdDocID,Veh_Order.Inv_DocId,Veh_Order.PartyCode, " & _
    "Veh_Order.GOVT_YN, Veh_Order.MODEL, Veh_Order.Chassis, Veh_Order.Srv_BookNo, Veh_Order.RATE, Veh_Order.Fund_Source, Veh_Order.FB_CODE, Veh_Order.FIN_AcCode, Veh_Order.FIN_AcCode, Veh_Order.Colour_Code, Veh_Order.FIN_AMT " & _
    "FROM Veh_Order " & _
    "where OrdDocid = '" & Txt(BookNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
    
    If RsBooking.RecordCount = 0 Then
        MsgBox "Booking No. Not Exist", vbInformation, "Booking Not Found"
        Txt(NamePrefix).TEXT = ""
        Txt(Party).TEXT = ""
        Txt(Party).Tag = ""
        Txt(FNamePrefix).TEXT = ""
        Txt(fname).TEXT = ""
        Txt(Add1).TEXT = ""
        Txt(Add2).TEXT = ""
        Txt(Add3).TEXT = ""
        Txt(City).TEXT = ""
        Txt(Model).TEXT = ""
        Txt(Govt_YN).TEXT = ""
        Txt(ChassisNo).TEXT = ""
        Txt(Colours).Tag = ""
        Txt(Colours).TEXT = ""
        Txt(SrvBookNo).TEXT = ""
        Txt(NDP).TEXT = ""
        Txt(FundSource).TEXT = ""
        Txt(FB_Code).TEXT = ""
        Txt(FB_Code).Tag = ""
        FinAcCode = ""
        Txt(FinAmt).TEXT = ""
        Txt(BookNo).Tag = ""
        Txt(BookNo).SetFocus
        Set RsBooking = Nothing
        FillRecords = False
        Exit Function
    Else
        If RsBooking!Inv_DocId <> Null Or RsBooking!Inv_DocId <> "" Then
            MsgBox "Invoice Exist Against Booking No", vbInformation, "Validation Check"
            Txt(NamePrefix).TEXT = ""
            Txt(Party).TEXT = ""
            Txt(Party).Tag = ""
            Txt(FNamePrefix).TEXT = ""
            Txt(fname).TEXT = ""
            Txt(Add1).TEXT = ""
            Txt(Add2).TEXT = ""
            Txt(Add3).TEXT = ""
            Txt(City).TEXT = ""
            Txt(Model).TEXT = ""
            Txt(Govt_YN).TEXT = ""
            Txt(ChassisNo).TEXT = ""
            Txt(Colours).Tag = ""
            Txt(Colours).TEXT = ""
            Txt(SrvBookNo).TEXT = ""
            Txt(NDP).TEXT = ""
            Txt(FundSource).TEXT = ""
            Txt(FB_Code).TEXT = ""
            Txt(FB_Code).Tag = ""
            FinAcCode = ""
            Txt(FinAmt).TEXT = ""
            Txt(BookNo).Tag = ""
            Txt(BookNo).SetFocus
            Set RsBooking = Nothing
            FillRecords = False
            Exit Function
        End If
        Txt(AdvAmt) = Format(PartyAdvance(RsBooking!OrdDocId, Txt(Vdate)), "0.00")
        Txt(Party).Tag = IIf(IsNull(RsBooking!PartyCode), "", RsBooking!PartyCode)
        If Txt(Party).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select NamePrefix,name,FPrefix,FName,add1,add2,add3,CityCode from SubGroup where Subcode = '" & Txt(Party).Tag & "'", GCn, adOpenStatic, adLockReadOnly
            Txt(NamePrefix).TEXT = IIf(IsNull(Rst!NamePrefix), "", Rst!NamePrefix)
            Txt(Party).TEXT = Rst!Name
            Txt(FNamePrefix).TEXT = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
            Txt(fname).TEXT = IIf(IsNull(Rst!fname), "", Rst!fname)
            Txt(Add1).TEXT = IIf(IsNull(Rst!Add1), "", Rst!Add1)
            Txt(Add2).TEXT = IIf(IsNull(Rst!Add2), "", Rst!Add2)
            Txt(Add3).TEXT = IIf(IsNull(Rst!Add3), "", Rst!Add3)
            Txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
            If Txt(City).Tag <> "" Then
                Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
            End If
            Txt(RTO).TEXT = Txt(City).TEXT
        End If
        Txt(Model).TEXT = RsBooking!Model
        Txt(Govt_YN).TEXT = IIf(IsNull(RsBooking!Govt_YN), "", RsBooking!Govt_YN)
        Txt(Colours).Tag = IIf(IsNull(RsBooking!Colour_Code), "", RsBooking!Colour_Code)
        If Txt(Colours).Tag <> "" Then
            Txt(Colours).TEXT = GCn.Execute("select col_desc from colmast where col_code = '" & Txt(Colours).Tag & "'").Fields(0).Value
        End If
        Select Case RsBooking!Fund_Source
            Case 0 '0 Hypothication ,1 Hire purchase ,2 Own Fund,3 Lease
                Txt(FundSource).TEXT = "Hypothication"
            Case 1
                Txt(FundSource).TEXT = "Hire Purchase"
            Case 3
                Txt(FundSource).TEXT = "Lease"
            Case 4
                Txt(FundSource).TEXT = "Agreement"
            Case 5
                Txt(FundSource).TEXT = "Lease & Agreement"
            Case 6
                Txt(FundSource).TEXT = "Loan Cum Hypt."
            Case Else
                Txt(FundSource).TEXT = "Own Fund"
        End Select
        Txt(FB_Code).Tag = IIf(IsNull(RsBooking!FB_Code), "", RsBooking!FB_Code)
        If Txt(FB_Code).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select fincode as code,finname & ',' & City.CityName as name,AcCode,FinBankCode " & _
                    " from ContractFinance left join city on left(ContractFinance.City,4)=City.CityCode " & _
                    " where fincatg = 0 and  fincode = '" & Txt(FB_Code).Tag & "'", GCn, adOpenStatic, adLockReadOnly
            Txt(FB_Code).TEXT = Rst!Name
            FinAcCode = IIf(IsNull(Rst!AcCode), "", Rst!AcCode)
        Else
            Txt(FB_Code).TEXT = ""
            FinAcCode = ""
        End If
        Txt(FinAmt).TEXT = IIf(IsNull(RsBooking!FIN_AMT), "", RsBooking!FIN_AMT)
    End If
Set Rst = Nothing
FillRecords = True
error1:
    CheckError

End Function
'************************ PRINTING CODE ******************

Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case DocType
        ListArray = Array("Sale Bill", "Sale Certificate", "Form22", "Form22A", "Declaration")
        Set mListItem = ListView_Items(ListView, txtPrint, Index, ListArray, 5)
End Select
End Sub

Private Sub TxtPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case DocType
        ListView_KeyDown FrmList, ListView, txtPrint, Index, KeyCode, Shift, FrmPrn.left + txtPrint(Index).left, (FrmPrn.top + txtPrint(Index).top + txtPrint(Index).height), txtPrint(Index).width, 1200
End Select
If FrmList.Visible = False And DGSite.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> TempInvDate Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case CertiTempYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            txtPrint(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txtPrint(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txtPrint(Index) = ""
        End If
        KeyAscii = 0
        FldEnabled1 (IIf(txtPrint(Index) = "Yes", True, False))
    Case WtPrn
        If UCase(Chr(KeyAscii)) = "Y" Then
            txtPrint(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txtPrint(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txtPrint(Index) = ""
        End If
        KeyAscii = 0
    Case Seet
       Call NumPress(txtPrint(Index), KeyAscii, 2, 0)
End Select
'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub txtPrint_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case DocType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txtPrint, Index, KeyCode, mListItem
End Select
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
  Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case DocType
        If txtPrint(Index).TEXT <> "" Then txtPrint(Index).TEXT = ListView.SelectedItem.TEXT
        If txtPrint(Index).TEXT = "Sale Certificate" Then FldEnabled True Else FldEnabled False
    Case TempInvDate, CertiPrnDate
        txtPrint(Index).TEXT = RetDate(txtPrint(Index))
End Select
End Sub

Private Sub FldEnabled(Enb As Boolean)
    txtPrint(RTOName).Enabled = Enb
    txtPrint(CertiPrnDate).Enabled = Enb
    txtPrint(CertiTempYN).Enabled = Enb
    txtPrint(Seet).Enabled = Enb
    txtPrint(Body).Enabled = Enb
    txtPrint(Narr).Enabled = Enb
    txtPrint(WtPrn).Enabled = Enb
    If Enb = False Then
        txtPrint(CertiPrnDate).TEXT = ""
        txtPrint(CertiTempYN).TEXT = ""
        txtPrint(Seet).TEXT = ""
        txtPrint(Body).TEXT = ""
        txtPrint(Narr).TEXT = ""
    End If
End Sub
Private Sub FldEnabled1(Enb As Boolean)
    txtPrint(Seet).Enabled = Enb
    txtPrint(Body).Enabled = Enb
    txtPrint(Narr).Enabled = Enb
    txtPrint(WtPrn).Enabled = Enb
    txtPrint(RTOName).Enabled = Enb
    If Enb = False Then
        txtPrint(Seet).TEXT = ""
        txtPrint(Body).TEXT = ""
        txtPrint(Narr).TEXT = ""
        txtPrint(WtPrn).TEXT = ""
    End If
End Sub

Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    FrmPrn.Visible = False
    If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
        If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
        Disp_Text SETS("INI", Me, Master)
        Call MoveRec
    End If
End If
End Sub
Private Sub Cmdprint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
'*****For Calcelled bill Dos printing
If CancelBillY_N = True Then
    txtPrint(DocType) = "Sale Bill"
    Index = 2
End If
'*******
If IsValid(txtPrint(DocType), "Print Document") = False Then Exit Sub
'"Sale Bill", "Sale Certificate", "Form22", "Form22A", "Declaration"
Select Case txtPrint(DocType)
    Case "Sale Bill"
        'mRepName = IIf(OptPlain.Value = True, "VehSale", "VehSale")
        'modi lps 15-04-2003
        mRepName = GCn.Execute("Select VBilRptName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
        GSQL = "SELECT VO.OrdDocID,Ord_No,Ord_Date,VO.Model,VO.PartyCode,VO.RATE,VO.Fund_Source,VO.FIN_YN,VO.FB_CODE,VO.FIN_AMT,VO.Inv_DocId,VO.Inv_VType,VO.Inv_No,VO.Inv_Date," & _
            " VO.Form_Code,VO.TrnType_Prn,VO.VRATE,VO.MARGINE,VO.REBATE,VO.InciChrg,VO.Octroi,VO.RegTemp,VO.TransitInsu,VO.Transport,VO.MVT,VO.TAX_Per,VO.TAX_Amt,VO.Surcharge_Per," & _
            " VO.Surcharge_Amt,VO.TOT_Per,VO.TOT_Amt,VO.OtherChrg,VO.FIT_AMT,VO.FIT_TAX,VO.Round_off,VO.DieselAmt,VO.BillPrn_YN,VO.DETAILS_YN,VO.INS_FEE,VO.INS_NOTE,VO.S_CHARGE,VO.RoundOff_YN,VO.Net_Amount," & _
            " VO.Inv_UName,VO.Inv_UEntDt,Veh_Purch1.gate,CF.finname,City_1.CityName as fincity, CF.Add1 as finadd1,CF.Add2 as finadd2,FinBank.FinBankName,site.site_desc,VStk.Pur_DocId," & _
            " VStk.Sal_DocId,VStk.ChassisNo, VStk.EngineNo, VStk.PBILL_NO, VStk.PBILL_DATE, " & _
            " M.Model_Desc,M.Model_Desc1, ColMast.Col_Desc,SG.NamePrefix, SG.Name, " & _
            " iif(isnull(SG.Add1) or SG.Add1='',SG.TAdd1,SG.Add1) as PAdd1, " & _
            " iif(isnull(SG.Add2) or SG.Add2='',SG.TAdd2,SG.Add2) as PAdd2, " & _
            " iif(isnull(SG.Add3) or SG.Add3='',SG.TAdd3,SG.Add3) as PAdd3, " & _
            " iif(isnull(City.CityName) or City.CityName='',TCity.CityName,City.CityName) as PCityName, " & _
            " SG.FPrefix,SG.FName,M.WHEELBASE,M.RIMS,M.TYRES,M.TyreDetails,M.GearBoxNo,TaxForms.Printing_Desc,SG.PANNo, #" & txtPrint(TempInvDate) & "# as InvDate,VO.TOT_Per,VO.TOT_Amt " & _
            " FROM (((((((((((Veh_Order as VO LEFT JOIN Veh_Stock as VStk ON VO.Inv_DocId = VStk.Sal_DocId) " & _
            " LEFT JOIN TaxForms ON VO.Form_Code = TaxForms.Form_Code) " & _
            " LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
            " LEFT JOIN Model as M ON VO.MODEL = M.MODEL) " & _
            " LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
            " LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
            " LEFT JOIN City TCity ON SG.TCityCode = TCity.CityCode) " & _
            " LEFT JOIN ContractFinance as CF ON VO.FB_CODE = CF.FinCode) " & _
            " LEFT JOIN Site ON right(VO.Inv_SiteCode,1) = Site.Site_Code) " & _
            " LEFT JOIN FinBank ON CF.FinBankCode = FinBank.FinBankCode) " & _
            " LEFT JOIN City AS City_1 ON CF.City = City_1.CityCode) " & _
            " LEFT JOIN Veh_Purch1 ON VStk.Pur_DocId = Veh_Purch1.DocID  " & _
            " where VO.Inv_DocId = '" & Master!SearchCode & "'"
    Case "Sale Certificate"
        mRepName = IIf(OptPlain.Value = True, "VehSaleCert", "VehSaleCert")
        GSQL = "SELECT VO.INTD_USE, VP1.Tot_Amount,VStk.vrate, VStk.Mfg_Month ,VStk.Mfg_Yr, " & _
            "VO.CertiPrn_YN, VO.TCertiPrn_YN,SG.FPrefix, SG.FName, city_1.cityname as TCity,SG.TAdd1, " & _
            "SG.TAdd2, SG.TAdd3,SG.TPIN, VO.Inv_DocId, #" & txtPrint(TempInvDate) & "# as Inv_Date, VO.Fund_Source, VO.P_AMOUNT, " & _
            "VO.DelCh_DT, '" & txtPrint(RTOName) & "' as RTO, Model_Grp.ModelGrp_Name, city.CityName, Fincity.cityname as FinCity," & _
            "SG.Name, ColMast.Col_Desc, M.MODEL, M.Vehicle_Type, " & _
            "M.Model_Desc, M.Model_Desc1, M.Model_Desc2,M.Wheel_Catg, " & _
            "M.TYRES, M.TYRE_F, M.TYRE_M, M.TYRE_R,M.TYRE_FS, M.TYRE_MS, M.TYRE_RS, " & _
            "M.TyreDetails,M.SEAT,M.RLW,M.HORSEPOWER , M.FRONT_A_WT, M.REAR_A_WT," & _
            "M.UNLADEN_WT,M.GROSS_WT, M.WHEELBASE, M.CYLINDER, M.FUEL,M.TRADE_NO, M.Manufacturer, " & _
            "VStk.ChassisNo, VStk.EngineNo,Finbank.FinBankName,CF.FinName, CF.Add1 as FAdd1," & _
            "CF.add2 as Fadd2, CF.PinCode as FPin , SG.Add1, SG.Add2, SG.Add3, SG.PIN ,vo.Inv_UName,vo.Inv_UEntDt " & _
            "FROM ((((((((((Veh_Order VO " & _
            "LEFT JOIN veh_Stock VStk ON VO.Inv_DocId = VStk.Sal_DocId) " & _
            "LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
            "LEFT JOIN Veh_Purch1 VP1 ON VStk.Pur_DocId = VP1.DocID) " & _
            "LEFT JOIN ContractFinance CF ON VO.FB_CODE = CF.FinCode) " & _
            "LEFT JOIN Subgroup SG ON VO.PartyCode = SG.SubCode) " & _
            "LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
            "LEFT JOIN Model_Grp ON M.Grp_Code = Model_Grp.ModelGrp_Code) " & _
            "LEFT JOIN city AS fincity ON CF.City = fincity.CityCode) " & _
            "LEFT JOIN Finbank ON CF.FinBankCode = Finbank.FinBankCode) " & _
            "LEFT JOIN city ON SG.CityCode = city.CityCode) " & _
            "LEFT JOIN City AS city_1 ON SG.TCityCode = city_1.CityCode " & _
            "where VO.Inv_DocId = '" & Master!SearchCode & "' and VO.DelCh_docid <> Null"
    Case "Form22A", "Form22"
        If txtPrint(DocType) = "Form22" Then
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert22", "VehSaleCert22")
        Else
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert22A", "VehSaleCert22A")
        End If
        GSQL = "SELECT M.Manufacturer, D.MfgAdd1,D.MfgAdd2,D.MfgAdd3," & _
            "M.MODEL,M.Chas_Type,M.Vehicle_Type, M.Sales_Desc," & _
            "M.Model_Desc, M.Model_Desc1, M.Model_Desc2, " & _
            "M.WHEELBASE,M.Fuel, VStk.ChassisNo , VStk.EngineNo,vo.Inv_UName,vo.Inv_UEntDt " & _
            "FROM ((veh_order as VO LEFT JOIN veh_Stock as VStk ON VO.Inv_DocId = VStk.Sal_DocId) " & _
            "LEFT JOIN Model as M ON VO.MODEL = M.MODEL) " & _
            "Left Join Division as D on D.Div_Code=left(VO.Inv_DocId,1) " & _
            " where VO.Inv_DocId = '" & Master!SearchCode & "' and VO.DelCh_docid <> Null"
End Select
Select Case Index
    Case PScreen, PWindows
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        If txtPrint(DocType) = "Sale Certificate" Then
            SpeedPrintCerti GSQL
        ElseIf txtPrint(DocType) = "Form22A" Then
            SpeedPrint22A GSQL
        ElseIf txtPrint(DocType) = "Form22" Then
            SpeedPrint22 GSQL
        ElseIf txtPrint(DocType) = "Sale Bill" Then
            SpeedPrintInv GSQL
        Else
            SpeedPrintDeclar
        End If
        FrmPrn.Visible = False
    Case PSetUp
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox Err.Description, vbCritical, Me.CAPTION

End Sub

Private Sub WindowsPrint(Index As Integer, mQRY As String)
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset, RstSub2 As ADODB.Recordset
Dim I As Integer, cnt As Integer, Foot1 As String, Foot2 As String, Foot3 As String, Foot4 As String
Dim Foot5 As String, Foot6 As String, Foot7 As String, Foot8 As String, Foot9 As String
Dim RstCompDet As ADODB.Recordset, J As Integer, Footer As String
Dim Rst2 As ADODB.Recordset
'On Error GoTo ERRORHANDLER

If txtPrint(DocType) = "Sale Certificate" Then
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    Set RstCompDet = GCn.Execute("select V_SecPAN_No,V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("CompPanNo")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecPAN_No & "'"
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecSpeciality & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecGram & "'"
            Case UCase("RTOName")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(RTOName) & "'"
            Case UCase("PrnDate")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(CertiPrnDate) & "'"
            Case UCase("TempYN")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(CertiTempYN) & "'"
            Case UCase("Seet")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(Seet) & "'"
            Case UCase("Body")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(Body) & "'"
            Case UCase("Narr")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(Narr) & "'"
            Case UCase("WtPrn")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(WtPrn) & "'"
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Select Case Index
        Case PWindows
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case UCase("comp_name")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
                    Case UCase("comp_add1")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
                    Case UCase("comp_add2")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_Add2 & "'"
                    Case UCase("comp_city")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
                    Case UCase("Title")
                        rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
                End Select
            Next
            rpt.PrintOut False
            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                If txtPrint(CertiTempYN) = "Yes" Then
                    GCn.Execute "update veh_order set CertiPrn_YN = 1  where where veh_order.Inv_DocId = '" & Master!SearchCode & "' And Veh_Order.DelCh_docid <> Null"
                Else
                    GCn.Execute "update veh_order set TCertiPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "' And Veh_Order.DelCh_docid <> Null"
                End If
            End If
            Set Rst = Nothing
            Set RstCompDet = Nothing
            Set rpt = Nothing
        Case PScreen 'screen
            Call Report_View(rpt, Me.CAPTION, , True)
            Set Rst = Nothing
            Set RstCompDet = Nothing
    End Select
ElseIf txtPrint(DocType) = "Form22A" Or txtPrint(DocType) = "Form22" Then
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
'            Case UCase("PrnDate")
'                rpt.FormulaFields(i).Text = "'" & txtPrint(PrnDate) & "'"
            Case UCase("Narr")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(Narr) & "'"
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Select Case Index
        Case PWindows
            rpt.PrintOut False
            Set Rst = Nothing
'            Set Rst1 = Nothing
            Set rpt = Nothing
        Case PScreen 'screen
            Call Report_View(rpt, Me.CAPTION, , True)
            Set Rst = Nothing
'            Set Rst1 = Nothing
    End Select
Else    'Sale Bill
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    'Recordset is made for subreport1
    mQRY = "SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
    "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    "where Veh_Purch2.DocId = '" & Master!SearchCode & "'"
    
    Set RstSub1 = New Recordset
    RstSub1.CursorLocation = adUseClient
    RstSub1.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic

   'Recordset is made for subreport2
   
    mQRY = "SELECT Veh_Purch2.Trn_Type, Veh_Purch2.DocID, Veh_Purch2.QTY, Veh_Purch2.RATE, Veh_AMDModel.Prod_Name " & _
    "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code where veh_stock.Chassisno = '" & Txt(ChassisNo) & "'"
        
    Set RstSub2 = New Recordset
    RstSub2.CursorLocation = adUseClient
    RstSub2.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
'modi LPS 15-04-2003
'    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
'    CreateFieldDefFile RstSub1, PubRepoPath + "\" & mRepName & "1.ttx", True
'    CreateFieldDefFile RstSub2, PubRepoPath + "\" & mRepName & "2.ttx", True
    CreateFieldDefFile Rst, PubRepoPath + "\VehSale.ttx", True
    CreateFieldDefFile RstSub1, PubRepoPath + "\VehSale1.ttx", True
    CreateFieldDefFile RstSub2, PubRepoPath + "\VehSale2.ttx", True
    
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    Set RstCompDet = New ADODB.Recordset
    RstCompDet.CursorLocation = adUseClient
    RstCompDet.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    J = 1
    cnt = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Select Case cnt
            Case 1
                Foot1 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 2
                Foot2 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 3
                Foot3 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 4
                Foot4 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 5
                Foot5 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 6
                Foot6 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 7
                Foot7 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 8
                Foot8 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            Case 9
                Foot9 = left(RTrim(mID(Footer, J, I - J - 1)), 130)
            End Select
            cnt = cnt + 1
            J = I + 1
        End If
    Next
    
    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select SupInvOnVehSaleInv , TaxDetOnVehInv from Syctrl", GCn, adOpenDynamic, adLockOptimistic
        For I = 1 To rpt.ParameterFields.Count
            Select Case UCase(rpt.ParameterFields(I).ParameterFieldName)
                Case UCase("PrePrinted")
                    rpt.ParameterFields(I).AddCurrentValue (IIf(OptPlain.Value, False, True))
            End Select
        Next
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecSpeciality & "'"
            Case UCase("AmtPrefix")
                rpt.FormulaFields(I).TEXT = "'" & PubAmountPrefix & "'"
            Case UCase("TelcoInvYN")
                rpt.FormulaFields(I).TEXT = "" & Rst2!SupInvOnVehSaleInv & ""
            Case UCase("TaxDetYN")
                rpt.FormulaFields(I).TEXT = "" & Rst2!TaxDetOnVehInv & ""
'            Case UCase("InvPrefix")
'                rpt.FormulaFields(i).Text = "'" & Rst2!VehSaleInv_Prefix & "'"
            Case UCase("LST")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecGram & "'"
            Case UCase("SubRep1")
                rpt.FormulaFields(I).TEXT = "" & IIf(RstSub1.RecordCount = 0, 0, 1) & ""
            Case UCase("SubRep2")
                rpt.FormulaFields(I).TEXT = "" & IIf(RstSub2.RecordCount = 0, 0, 1) & ""
           Case UCase("AddDet")
                rpt.FormulaFields(I).TEXT = "" & IIf(Txt(ADType) = "No Detail", 0, IIf(Txt(ADType) = "Name/Qty", 1, 2)) & ""
           Case UCase("Foot1")
                rpt.FormulaFields(I).TEXT = "'" & Foot1 & "'"
            Case UCase("Foot2")
                rpt.FormulaFields(I).TEXT = "'" & Foot2 & "'"
            Case UCase("Foot3")
                rpt.FormulaFields(I).TEXT = "'" & Foot3 & "'"
            Case UCase("Foot4")
                rpt.FormulaFields(I).TEXT = "'" & Foot4 & "'"
            Case UCase("Foot5")
                rpt.FormulaFields(I).TEXT = "'" & Foot5 & "'"
            Case UCase("Foot6")
                rpt.FormulaFields(I).TEXT = "'" & Foot6 & "'"
            Case UCase("Foot7")
                rpt.FormulaFields(I).TEXT = "'" & Foot7 & "'"
            Case UCase("Foot8")
                rpt.FormulaFields(I).TEXT = "'" & Foot8 & "'"
            Case UCase("Foot9")
                rpt.FormulaFields(I).TEXT = "'" & Foot9 & "'"
        End Select
    Next
    For I = 1 To rpt.OpenSubreport("SubRep2").FormulaFields.Count
        Select Case UCase(rpt.OpenSubreport("SubRep2").FormulaFields(I).FormulaFieldName)
            Case UCase("AddDet")
            rpt.OpenSubreport("SubRep2").FormulaFields(I).TEXT = "" & IIf(Txt(ADType) = "No Detail", 0, IIf(Txt(ADType) = "Name/Qty", 1, 2)) & ""
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstSub1
    rpt.OpenSubreport("SubRep2").Database.SetDataSource RstSub2
    rpt.ReadRecords
    Select Case Index
        Case PWindows  'Printer
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case UCase("comp_name")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
                    Case UCase("comp_add1")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
                    Case UCase("comp_add2")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_Add2 & "'"
                    Case UCase("comp_city")
                        rpt.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
                    Case UCase("Title")
                        rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
                End Select
            Next
            rpt.PrintOut False
            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
            End If
            Set Rst = Nothing
            Set RstCompDet = Nothing
            Set rpt = Nothing
        Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
            Set Rst = Nothing
            Set RstCompDet = Nothing
    End Select
End If
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox Err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub
Private Sub SpeedPrint22A(mQRY$)
'On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, J As Integer
    Dim RstCert As ADODB.Recordset
    Dim PageWidth As Byte, PageLength As Integer
    Dim mHeader As Byte, mFooter As Byte
    Dim fob As New FileSystemObject

    Set RstCert = GCn.Execute(mQRY)

    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
 
    PageLength = PubPageLength
    PageWidth = 80 '34
    mHeader = 0   'Ideal 17
    mFooter = 2
        
    Print #1, Chr(27) + Chr(67) + Chr(36) & PRN_TIT(Trim(RstCert!Manufacturer), "B", PageWidth) 'small paper size
'        Print #1, PRN_TIT(Trim(RstCert!Manufacturer), "C", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCert!MfgAdd1) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd1)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd2) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd2)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd3) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd3)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PRN_TIT("F O R M - 22-A", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[See Rule 47 (g),124,126a and 127]", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("Part 1", "B", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("(Issued By The Manufacturer)", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "Certified of compliance and pollution standards/safety of components"
        mHeader = mHeader + 1
        Print #1, mSP5 & "Road Worthiness (for vehicle whose Body is fabricated separately)"
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "Certified that Tata " & mChr17 & RstCert!FUEL & mChr18 & " Vehicle Model " & mEmph & RstCert!Model_Desc & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP5 & RstCert!WHEELBASE & "MM Wheel Base (Brand name of the vehicle) Truck/Bus/Car bearing "
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Chassis Number :" & RstCert!ChassisNo
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Engine Number  :" & RstCert!EngineNo
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "complies with the provisions of  the  Motor Vehicles Act, 1988 and the rule"
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "made there under."
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("For " & RstCert!Manufacturer, PageWidth, , AlignRight) & mEmph1
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, PSTR("Signature of the manufacturer", PageWidth, , AlignRight)
        Do Until mHeader >= PageLength - mFooter - 6
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        Print #1, mSP5 & Replace(Space(PageWidth), " ", "-")
'        Print #1, mSP5 & mChr17 & RstCert!Inv_UName & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(RstCert!Inv_UName)) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrint22(mQRY As String)
'On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, J As Integer
    Dim RstCert As ADODB.Recordset
    Dim PageWidth As Byte, PageLength As Integer
    Dim mHeader As Byte, mFooter As Byte
    Dim fob As New FileSystemObject

    Set RstCert = GCn.Execute(mQRY)

    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1

    PageLength = PubPageLength
    PageWidth = 80 '34
    mHeader = 0   'Ideal 17
    mFooter = 2
    Print #1, Chr(27) + Chr(67) + Chr(36) & PRN_TIT(Trim(RstCert!Manufacturer), "B", PageWidth) 'small paper size
        
'        Print #1, PRN_TIT(Trim(RstCert!Manufacturer), "B", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCert!MfgAdd1) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd1)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd2) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd2)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd3) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd3)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PRN_TIT("F O R M - 22", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[See Rule 47 (g), and 127]", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PRN_TIT("Initial certificate of Road Worthiness", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("(To be Issued By The Manufacturer)", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "Certified that Tata " & mChr17 & RstCert!FUEL & mChr18 & " Vehicle Model " & mEmph & RstCert!Model_Desc & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP5 & RstCert!WHEELBASE & "MM Wheel Base (Brand name of the vehicle) Truck/Bus/Car bearing "
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Chassis Number :" & RstCert!ChassisNo
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Engine Number  :" & RstCert!EngineNo
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "complies with the provisions of  the  Motor Vehicles Act, 1988 and the rule"
        mHeader = mHeader + 1
        Print #1, mSP5 & "made there under."
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("For " & RstCert!Manufacturer, PageWidth, , AlignRight) & mEmph1
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, PSTR("Signature of the manufacturer", PageWidth, , AlignRight)
        
        Do Until mHeader >= PageLength - mFooter - 6
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        Print #1, mSP5 & Replace(Space(PageWidth), " ", "-")
'        Print #1, mChr17 & RstCert!Inv_UName & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(RstCert!Inv_UName)) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintCerti(mQRY$)
'On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, J As Integer
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstCert As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer1$, Footer2$, Footer3$, Footer4$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim NetAmt As Double
    
    Dim mPAdd1$, mPAdd2$, mPAdd3$, mPCity$, mPPin$
    Dim mTAdd1$, mTAdd2$, mTAdd3$, mTCITY$, mTPin$
    Dim mComp_Add$, mComp_Add2$, mComp_City$, mPhone$, mFax$

    Set RstCert = GCn.Execute(mQRY)
    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Set RstCert = Nothing: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
 
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 4
    
    mComp_Add = Trim(PubComp_Add)
    mComp_Add2 = Trim(PubComp_Add2)
    mComp_City = Trim(PubComp_City)
'    If XNull(RstCompDet!V_SecPhone) <> "" Then
'        mPhone = "PHONE : " & RstCompDet!V_SecPhone
'    End If
'    If XNull(RstCompDet!V_SecFax) <> "" Then
'        mFax = "  Fax :" & RstCompDet!V_SecFax
'    End If
    
    'Header
    'Form22A,RTOName,PrnDate,CertiTempYN,Seet,Body,Narr,WtPrn,TempRto
    
    If txtPrint(CertiTempYN) = "Yes" Then
'        mDocStr = "Temporary Sale Certificate"
        If RstCert!TCertiPrn_YN = 1 Then
            mDupStr = " (Duplicate)"
        End If
        mPAdd1 = XNull(RstCert!TAdd1)
        mPAdd2 = XNull(RstCert!TAdd2)
        mPAdd3 = XNull(RstCert!TAdd3)
        mPCity = XNull(RstCert!TCity)
        mPPin = XNull(RstCert!TPin)
        
        mTAdd1 = XNull(RstCert!Add1)
        mTAdd2 = XNull(RstCert!Add2)
        mTAdd3 = XNull(RstCert!Add3)
        mTCITY = XNull(RstCert!CityName)
        mTPin = XNull(RstCert!Pin)
    Else
 '       mDocStr = "Sale Certificate"
        If RstCert!CertiPrn_YN = 1 Then
            mDupStr = " (Duplicate)"
        End If
        mPAdd1 = XNull(RstCert!Add1)
        mPAdd2 = XNull(RstCert!Add2)
        mPAdd3 = XNull(RstCert!Add3)
        mPCity = XNull(RstCert!CityName)
        mPPin = XNull(RstCert!Pin)
        
        mTAdd1 = XNull(RstCert!TAdd1)
        mTAdd2 = XNull(RstCert!TAdd2)
        mTAdd3 = XNull(RstCert!TAdd3)
        mTCITY = XNull(RstCert!TCity)
        mTPin = XNull(RstCert!TPin)
    End If
    mDocStr = "Sale Certificate"
    
    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!V_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!V_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
    If PubComp_Add2 <> "" Or PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_Add2 & IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", "  Fax : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PRN_TIT("Form-21", "B", PageWidth)
    mHeader = mHeader + 1
    Print #1, PRN_TIT("[See Rule 47(a) and (d)]", "C", PageWidth)
    mHeader = mHeader + 1
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth) & mChr18 & mEmph
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("The Registration Authority,", 40) & mEmph & "     Invoice No.  : " & PSTR(Trim(mID(RstCert!Inv_DocId, 9, 5)) & "-" & Trim(mID(RstCert!Inv_DocId, 14, 8)), 14, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR(txtPrint(RTOName), 40) & mEmph & "     Invoice Date : " & txtPrint(TempInvDate) & mEmph1
    mHeader = mHeader + 1
'    Print #1, mSP5 & "Ex. Factory Price Rs.: " & Format(RstCert!VRATE, "0.00")
'    mHeader = mHeader + 1
    
    Print #1, mSP5 & "(To be issued by the manufacturer, Dealer or Officer or defence (in Case of "
    mHeader = mHeader + 1
    Print #1, mSP5 & "Military auctioned vehicles)for presentation along with the application For"
    mHeader = mHeader + 1
    Print #1, mSP5 & "registration of a motor vehicle.)"
    mHeader = mHeader + 1
    
    Print #1, mSP5 & "Certified that - " & mEmph & "One " & RstCert!Model_Desc & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP5 & "Has been delivered by us on " & mEmph & RstCert!DelCh_DT & mEmph1 & " to :- "
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Name of Buyer", 28) & " : " & RstCert!Name
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Son/Wife/Daughter of ", 28) & " : " & XNull(RstCert!FPrefix) & " " & XNull(RstCert!fname)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Address(Permanent)", 28) & " : " & mPAdd1
    mHeader = mHeader + 1
    If mPAdd2 <> "" Then
        Print #1, mSP5 & Space(31) & mPAdd2
        mHeader = mHeader + 1
    End If
    If mPAdd3 <> "" Then
        Print #1, mSP5 & Space(31) & mPAdd3
        mHeader = mHeader + 1
    End If
    Print #1, mSP5 & Space(31) & mPCity & " " & mPPin
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Address(Temporary)", 28) & " : " & mTAdd1
    mHeader = mHeader + 1
    If mTAdd2 <> "" Then
        Print #1, mSP5 & Space(31) & mTAdd2
        mHeader = mHeader + 1
    End If
    If mTAdd3 <> "" Then
        Print #1, mSP5 & Space(31) & mTAdd3
        mHeader = mHeader + 1
    End If
    Print #1, mSP5 & Space(31) & mTCITY & " " & mTPin
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    If RstCert!Fund_Source = 0 Then
        Print #1, mSP5 & "The vehicle is held under agreement of Hypothication with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 1 Then
        Print #1, mSP5 & "The vehicle is held under agreement of Hire purchase with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 3 Then
        Print #1, mSP5 & "The vehicle is held under agreement of Lease with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 4 Then
        Print #1, mSP5 & "The vehicle is held under Hire purchase finance agreement with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 5 Then
        Print #1, mSP5 & "The vehicle is held under Hire purchase finance Lease&agreement with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 6 Then
        Print #1, mSP5 & "The vehicle is held under Loan Cum Hypothication Agreement with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    
        
    Else
        Print #1, ""
        mHeader = mHeader + 1
    End If
    
    Print #1, mSP5 & XNull(RstCert!FAdd1) & " " & XNull(RstCert!FAdd2)
    mHeader = mHeader + 1
    Print #1, mSP5 & XNull(RstCert!FinCity) & " " & XNull(RstCert!FPin)
    mHeader = mHeader + 1
    
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mSP5 & "The details of the vehicle are given below :  "
    mHeader = mHeader + 1
    
    Print #1, mSP5 & PSTR("1. Class of Vehicle", 32) & " : " & XNull(RstCert!Vehicle_Type)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("2. Maker's Name", 32) & " : " & XNull(RstCert!Manufacturer)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("3. Chassis No.", 32) & " : " & XNull(RstCert!ChassisNo)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("4. Engine No.", 32) & " : " & XNull(RstCert!EngineNo)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("5. Horse Power/Cubic Capacity", 32) & " : " & XNull(RstCert!HorsePower)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("6. Fuel Used", 32) & " : " & XNull(RstCert!FUEL)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("7. No. of Cylinders ", 32) & " : " & XNull(RstCert!Cylinder)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("8. Month & Year of Mfg.", 32) & " : " & XNull(RstCert!Mfg_Month) & " " & XNull(RstCert!Mfg_Yr)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("9. Seating Capacity(Incld. Driver)", 32) & " : " & txtPrint(Seet)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("10.Unleaden Weight", 32) & " : " & RstCert!Unladen_Wt '& "Kg."
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("11.Maximum Axle Weight and number and Description of tyres", 60) & " : "
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("   (in Case of Transport Vehicle)", 32)
    mHeader = mHeader + 1
    Print #1, mSP5 & "    " & PSTR("Front Axle", 28) & " : " & RstCert!Front_A_Wt ' & "Kg."
    mHeader = mHeader + 1
    Print #1, mSP5 & "    " & PSTR("Rear Axle", 28) & " : " & RstCert!Rear_A_Wt '& "Kg."
    mHeader = mHeader + 1
    Print #1, mSP5 & "    " & PSTR("Any Other Axle", 28) & " : "
    mHeader = mHeader + 1
    Print #1, mSP5 & "   " & Space(32) & PSTR("Front", 16, , AlignLeft) & PSTR("Middle", 16, , AlignLeft) & PSTR("Rear", 16, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mSP5 & "    " & PSTR("No. of Tyres", 28) & " : " & PSTR(CStr(RstCert!Tyre_F), 16, , AlignLeft) & PSTR(CStr(RstCert!Tyre_M), 16, , AlignLeft) & PSTR(CStr(RstCert!Tyre_R), 16, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mSP5 & "    " & PSTR("Size of Tyres", 28) & " : " & PSTR(RstCert!Tyre_FS, 16, , AlignLeft) & PSTR(RstCert!Tyre_MS, 16, , AlignLeft) & PSTR(RstCert!Tyre_RS, 16, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mSP5 & "    " & PSTR("Other Details", 28) & " : " & RstCert!TyreDetails
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("12.Colour of Body", 32) & " : " & RstCert!Col_Desc
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("13.Gross Vehicle Weight", 32) & " : " & RstCert!Gross_Wt '& "Kg."
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("14.Type of Body", 32) & " : " & txtPrint(Body)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("15.Trade No", 32) & " : " & RstCert!Trade_NO
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("16.WheelBase", 32) & " : " & RstCert!WHEELBASE & " MM"
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mSP5 & mChr17 & "*Strike out whichever is inapplicable" & mChr18
    mHeader = mHeader + 1
    Do Until mHeader >= PageLength - mFooter
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    Print #1, mSP5 & mEmph & PSTR("For " & PubComp_Name, PageWidth, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, mSP5 & mChr17 & txtPrint(Narr) & mChr18
    Print #1, mSP5 & PSTR("Signature of Manufacturer/Dealer or Officer of Defence", 58) & "Authorised Signatory"
    'Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
    'End Of Page 1 For SAle Certificate
    Print #1, mEject
    
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        If txtPrint(CertiTempYN) = "Yes" Then
            GCn.Execute "update veh_order set CertiPrn_YN = 1  where where veh_order.Inv_DocId = '" & Master!SearchCode & "' And Veh_Order.DelCh_docid <> Null"
        Else
            GCn.Execute "update veh_order set TCertiPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "' And Veh_Order.DelCh_docid <> Null"
        End If
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintInv(mQRY$)
On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, J As Integer
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim mSaleRate As Single, mNetAmt As Single, mInv_No$

     Set Rstsale = GCn.Execute(mQRY)
    
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 19
    mFooter = mFooter + FooterCnt
    
    ' Header
    If CancelBillY_N = True Then
        mDocStr = "Sale Invoice (Credit Note)"
    Else
        mDocStr = "Sale Invoice"
    End If
    mDupStr = IIf(Rstsale!BillPrn_YN = 0, "", " (Duplicate)")
 '0 -Hypothication ,1- Hire purchase ,2 -Own Fund,3- Lease, 4-Agreement, 5-Lease & Agreement

    If Rstsale!Fund_Source = 0 Then   'Hypothication
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Hypothication to  "
        Right2 = XNull(Rstsale!FinBankName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!FIN_AMT, "0.00")
        
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Leaser  "
        Right2 = XNull(Rstsale!FinBankName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Lease Amount :" & Format(Rstsale!FIN_AMT, "0.00")
        
    ElseIf Rstsale!Fund_Source = 6 Then
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Loan Cum Hypt. Agreement with  "
        Right2 = XNull(Rstsale!FinBankName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!FIN_AMT, "0.00")
    
        
    ElseIf Rstsale!Fund_Source = 1 Or _
           Rstsale!Fund_Source = 4 Or _
           Rstsale!Fund_Source = 5 Then
        
        Left1 = "Sold to under HPA with, "      '1-Hire Purchase
        If Rstsale!Fund_Source = 4 Then         '4-Agreement
            Left1 = "Hire Purchase Finance Agreement with, "
        ElseIf Rstsale!Fund_Source = 5 Then     '5-Lease & Agreement
            Left1 = "Hire Purchase Finance Lease&Agreement with, "
        
        End If
        Left2 = " U/F " & XNull(Rstsale!FinBankName)
        Left3 = XNull(Rstsale!FinAdd1)
        Left4 = XNull(Rstsale!FinAdd2)
        Left5 = XNull(Rstsale!FinCity)
        Left6 = ""
        
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!PAdd1)
        Right5 = XNull(Rstsale!PAdd2)
        Right6 = XNull(Rstsale!PAdd3) & XNull(Rstsale!PCityName)
        
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
    End If
    

    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
    mInv_No = Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_Prefix)) & " - " & Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_No))
    
    '************
    mSaleRate = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate + Rstsale!InciChrg _
         + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
         + Rstsale!MVT + Rstsale!Transport
         
    mNetAmt = mSaleRate + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!Tot_Amt + Rstsale!OtherChrg + Rstsale!Fit_Amt _
        + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off
    '***********
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!V_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!V_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
         
    If PubComp_Add2 <> "" Or PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_Add2 & IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignRight)
    mHeader = mHeader + 1

    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, mChr18 & mEmph & PSTR(Left1, 40) & PSTR(Right1, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(Left2, 40) & PSTR(Right2, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left3, 40) & PSTR(Right3, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left4, 40) & PSTR(Right4, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left5, 40) & PSTR(Right5, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left6, 40) & PSTR(Right6, 40)
    mHeader = mHeader + 1
        
    Print #1, IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Telco Invoice No.: " & XNull(Rstsale!PBILL_NO) & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 40), Space(40)) & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("Telco Gate Pass No. : " & XNull(Rstsale!GATE), 40) & mEmph & "Invoice Date : " & str(Rstsale!Inv_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, "Booking No. & Date  : " & str(Rstsale!Ord_No) & "    " & IIf(IsNull(Rstsale!Ord_Date), "", (str(Rstsale!Ord_Date)))
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, PSTR("Model : " & Rstsale!Model_Desc, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight)
    mHeader = mHeader + 1
    Print #1, PSTR(Rstsale!Model_Desc1, 45) & PSTR(IIf(Rstsale!Tax_Per = 0, "", "Tax @ " & Format(Rstsale!Tax_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tax_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Colour      : " & Rstsale!Col_Desc, 40) & PSTR(IIf(Rstsale!surcharge_per = 0, "", "Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %"), 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Chassis No. : " & Rstsale!ChassisNo, 45) & PSTR(IIf(Rstsale!TOT_Per = 0, "", "TOT @ " & Format(Rstsale!TOT_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Engine No.  : " & Rstsale!EngineNo, 40) & PSTR("Other Charges", 27, , AlignRight) & ": " & PSTR(Rstsale!OtherChrg, 11, 2)
    mHeader = mHeader + 1
                
    Print #1, "Other Fitments Details : " & mEmph1
    mHeader = mHeader + 1
    
    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("Sr", 3) & PSTR("Item Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, 2, AlignRight) & " " & PSTR("<----Tax---->", 13) & " " & PSTR("<-Sur.On Tax->", 14) & " " & PSTR("Amount", 9, , AlignRight)
    mHeader = mHeader + 1
    Print #1, "No." & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
        
    Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")
    cnt = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            Print #1, mChr17 & str(cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) + Rst!Tax_Amt + Rst!TaxSur_Amt), 10, 2)
            mHeader = mHeader + 1
            cnt = cnt + 1
            Rst.MoveNext
        Loop
    End If
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
        "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where veh_stock.Chassisno = '" & Txt(ChassisNo) & "' " & _
        "Group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
    If Rst.RecordCount > 0 Then
        Print #1, mDoub & PSTR("Addition/Deletion/Shortage Detail", 52) & PSTR("Qty", 13, , AlignRight) & PSTR("Amount", 15, , AlignRight) & mDoub1
        mHeader = mHeader + 1
        Do Until Rst.EOF
            Print #1, PSTR(IIf(Rst!Trn_Type = "A", "Addition", IIf(Rst!Trn_Type = "D", "Deletion", "Shortage")), 52) & PSTR(Rst!TotQty, 13, 2) & PSTR(Rst!Amt, 15, 2)
            mHeader = mHeader + 1
            Rst.MoveNext
        Loop
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
    End If
        
    Do Until mHeader >= PageLength - mFooter
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    Print #1, PSTR(IIf(Rstsale!Round_off = 0, "", "Round Off"), 65, , AlignRight) & " : " & PSTR(Rstsale!Round_off, 12, 2)
    Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
    Print #1, mEmph & PSTR("Bill Amount", 65, , AlignRight) & " : " & PSTR(Amount_Fill((mNetAmt), PubAmountPrefix), 12, 2, AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, ntow(mNetAmt, "Rupees", "Paise") & mEmph1
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, "Complete With Tools and equipment as supplied by the manufacturer including "
    Print #1, "excise duty,Sales tax & delivery & handing charges."
    Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, ""
    Print #1, "Accountant" & PSTR("Authorised Signatory", PageWidth - 10, , AlignRight)
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")

    Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
        
    Footer = Footer & vbLf
    J = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, J, I - J))
            J = I + 1
        End If
    Next
    
    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
           
    Print #1, mChr17 & Rstsale!Inv_UName & " " & str(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & str(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If

    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Sub SpeedPrintDeclar()
'On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, J As Integer, mQRY As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim cnt As Byte, NetAmt As Double, PrnStr As String, PrnStr1 As String, mRegCert As String
    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set Rstsale = GCn.Execute("SELECT veh_order.*,City.CityName,  " & _
        " Veh_Stock.ChassisNo, Veh_Stock.EngineNo,Model.Model_Desc,Model.Model_Desc1, " & _
        " SubGroup.Name, SubGroup.Add1,SubGroup.Add2,SubGroup.Add3,SubGroup.Tadd1,SubGroup.Tadd2,SubGroup.Tadd3 FROM  " & _
        "(((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode where veh_order.Inv_DocId = '" & Master!SearchCode & "'")

    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1


    FooterCnt = 1
    Footer = ""

'    For i = 1 To Len(Footer)
'        If Mid(Footer, i, 1) = vbLf Then
'            FooterCnt = FooterCnt + 1
'        End If
'    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 6
    mFooter = mFooter + FooterCnt

    ' Header

    mDocStr = "DECLARATION"

        Print #1, PRN_TIT(mDocStr, "A", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[Under rule 2148 (1)]", "C", PageWidth) & mChr18 & mEmph
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "Declartion No. : " & Space(10) & "Date : " & mEmph1
        mHeader = mHeader + 1
        Print #1, "I/We declare that the following consignment of notified comodity is "
        mHeader = mHeader + 1
        Print #1, "Despatched from a place within West Bengal : "
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "1. " & "Name & address of the consignor : " & mEmph & PubComp_Name & mEmph1
        Print #1, Space(25) & " : " & mEmph & PubComp_Add & mEmph1
        mHeader = mHeader + 1
        Print #1, Space(25) & " : " & mEmph & PubComp_Add2 & mEmph1
        mHeader = mHeader + 1
        Print #1, Space(25) & " : " & mEmph & PubComp_City & mEmph1
        mHeader = mHeader + 1
        Print #1, Space(25) & " : " & IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax)
        mHeader = mHeader + 1
        If txtPrint(CertiTempYN) = "Yes" Then
            Print #1, PSTR("2. a) " & "Name and address ", 25) & " : " & mEmph & XNull(Rstsale!Name) & mEmph1
            mHeader = mHeader + 1
            Print #1, PSTR("of the Consignee", 25) & " : " & mEmph & XNull(Rstsale!TAdd1) & mEmph1
            mHeader = mHeader + 1
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!TAdd2) & XNull(Rstsale!TAdd3) & mEmph1
            mHeader = mHeader + 1
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, PSTR("   b) " & "Temporary Address", 25) & " : " & XNull(Rstsale!Add1)
            mHeader = mHeader + 1
            If XNull(Rstsale!Add2) <> "" Then
                Print #1, Space(25) & " : " & Rstsale!Add2
                mHeader = mHeader + 1
            End If
            If XNull(Rstsale!Add3) <> "" Then
                Print #1, Space(23) & " : " & Rstsale!Add3
                mHeader = mHeader + 1
            End If
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!CityName) & mEmph1
            mHeader = mHeader + 1
        Else
            Print #1, PSTR("2. a) " & "Name and address ", 25) & " : " & mEmph & XNull(Rstsale!Name) & mEmph1
            mHeader = mHeader + 1
            Print #1, PSTR("of the Consignee", 25) & " : " & mEmph & XNull(Rstsale!Add1) & mEmph1
            mHeader = mHeader + 1
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!Add2) & XNull(Rstsale!Add3) & mEmph1
            mHeader = mHeader + 1
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!CityName) & mEmph1
            mHeader = mHeader + 1
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, PSTR("   b) " & "Temporary Address", 25) & " : " & XNull(Rstsale!TAdd1)
            mHeader = mHeader + 1
            If XNull(Rstsale!TAdd2) <> "" Then
                Print #1, Space(25) & " : " & Rstsale!TAdd2
                mHeader = mHeader + 1
            End If
            If XNull(Rstsale!TAdd3) <> "" Then
                Print #1, Space(23) & " : " & Rstsale!TAdd3
                mHeader = mHeader + 1
            End If
            
        End If
        
        Print #1, "   c) " & "Registration certificate No. of the consignee [if registered under "
        mHeader = mHeader + 1
        Print #1, "the West Bengal Sales Tax Act. 1994(West Ben. Act. XLIX of 1994)/ "
        mHeader = mHeader + 1
        mRegCert = XNull(GCn.Execute("select RegCertNo from Syctrl").Fields(0).Value)
        Print #1, "the central Sals Tax Act. 1956(74 of 1956) ] : " & mEmph & "Nil" & mEmph1
        mHeader = mHeader + 1

        Print #1, "3. " & "Place Of Dispatch : " & mEmph & PubComp_City & mEmph1
        mHeader = mHeader + 1
        Print #1, "4. " & "Destination : "
        mHeader = mHeader + 1
        Print #1, "5. " & "Description of consignment  : " & mEmph & Rstsale!Model_Desc & mEmph1
        mHeader = mHeader + 1
        Print #1, "6. " & PSTR("Quantity  : ", 15) & mEmph & "1 No. (One)" & mEmph1
        mHeader = mHeader + 1
        Print #1, "7. " & PSTR(" Weight : ", 15)
        mHeader = mHeader + 1
        NetAmt = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate + Rstsale!InciChrg _
        + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
        + Rstsale!MVT + Rstsale!Transport + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!OtherChrg + Rstsale!Fit_Amt + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off

        Print #1, "8. " & PSTR("Value  : ", 15) & mEmph & NetAmt & mEmph1
        mHeader = mHeader + 1
        Print #1, "9. " & "Consignor Bill/Cash Memo/Other"
        mHeader = mHeader + 1
        Print #1, "   " & "Document(Specify) No. and date :" & mEmph & Rstsale!Inv_No & " Dt. " & Rstsale!Inv_Date & mEmph1
        mHeader = mHeader + 1
        Print #1, "   " & "Consignment or deleivery note No. and Date : " & mEmph & "delivery Receipt Dt.____________ "
        mHeader = mHeader + 1
        Print #1, "Chassis Deliverd at " & PubComp_City & " on ___________ and now "
        mHeader = mHeader + 1
        Print #1, "transported by the customer by his/her/their own mode" & mEmph1
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "I/We declare that I/we hold the registration certificate No " & mRegCert
        mHeader = mHeader + 1
        Print #1, "Under the West Bengal Sales Tax Act. 1994.(West Ben. Act XLIX of 1994)."
        mHeader = mHeader + 1
        Print #1, "We have not manufactured the comodity in West Bengal/not transported the "
        mHeader = mHeader + 1
        Print #1, "commodities from outside of West Bengal"
        mHeader = mHeader + 1
        Print #1, "The Above Statement are true to the best of my/our knowledge and belief"
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mEmph & "Chassis No :" & Rstsale!ChassisNo
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "Engine No :" & Rstsale!EngineNo & mEmph1
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "Temp. Regn No. ______________________ :"
        mHeader = mHeader + 1
        Print #1, Space(40) & "Signature ______________________ :"
        mHeader = mHeader + 1
        Print #1, Space(40) & "Status of the declarent  ______________________ :" & mEmph1
        mHeader = mHeader + 1
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Function ProcAcPost(Optional CheckCtrls As Boolean) As Boolean
On Error GoTo lblExit
        Dim MsgStr$, rsCtrlAc As ADODB.Recordset, rsTemp As ADODB.Recordset, mPostFinAmt As Byte
        Dim mGTotAmt As Double, mTOT_Ac_Code$, mCommNarr$
        
        Set rsCtrlAc = New ADODB.Recordset
        rsCtrlAc.CursorLocation = adUseClient
        rsCtrlAc.Open "Select Fitment_Ac,Fuel_Ac,VehROff_Ac From AcControls", GCnFaV, adOpenStatic, adLockReadOnly
        If rsCtrlAc.RecordCount <= 0 Then
            MsgStr = "Please Add Records in A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        If IsNull(rsCtrlAc!Fitment_Ac) Or rsCtrlAc!Fitment_Ac = "" Or _
            IsNull(rsCtrlAc!Fuel_Ac) Or rsCtrlAc!Fuel_Ac = "" Or _
            IsNull(rsCtrlAc!VehROff_Ac) Or rsCtrlAc!VehROff_Ac = "" Then
            MsgStr = "Please define Fitment,Fuel and Round Off A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        rsForm.MoveFirst        'Vehicle Sale A/c Code, Tax A/c Code, Surcharge A/c Code
        rsForm.FIND "Name ='" & Txt(FormType) & "'"
        If IsNull(rsForm!PurSal_Ac_Code) Or rsForm!PurSal_Ac_Code = "" Then
            MsgStr = "Please Define Sale A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        'Tax A/c Code Checking
        If Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(OthFitTax)) <> 0 Then
            If IsNull(rsForm!Tax_Ac_Code) Or rsForm!Sur_Ac_Code = "" Then
                MsgStr = "Please Define Tax A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
                ProcAcPost = False
                GoTo lblExit
            End If
        End If
        'Financier A/c Checking
        mTOT_Ac_Code = G_FaCn.Execute("select iif(isnull(totax_ac),'',TOTax_Ac) as TOT_Ac from AcControls where Div_Code='" & PubDivCode & "'").Fields(0).Value
        If Val(Txt(TOTAmt)) <> 0 And mTOT_Ac_Code = "" Then
            MsgStr = "Please define TOT A/c Code in Vehicle Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        mPostFinAmt = GCn.Execute("select iif(isnull(PostFinAmt),0,postfinamt) as PostFinAmt from Syctrl").Fields(0).Value
        If mPostFinAmt = 1 And Val(Txt(FinAmt)) <> 0 Then
            If Txt(FundSource) = "Hypothication" Or Txt(FundSource) = "Hire Purchase" Then
                Set rsTemp = New ADODB.Recordset
                rsTemp.CursorLocation = adUseClient
                rsTemp.Open "Select switch(Ac_YN='1','Y',Ac_YN<>'1','N') as ACYN,AcCode From ContractFinance where FinCode='" & Txt(FB_Code).Tag & "' ", GCn, adOpenStatic, adLockReadOnly
                If rsTemp!AcYN = "Y" Then
                    If rsTemp!AcCode = "" Or IsNull(rsTemp!AcCode) Then
                        MsgStr = "Please define A/c Code in Financier Master" & vbCrLf & "A/c Posting Aborted !"
                        GoTo lblExit
                    End If
                End If
            End If
        End If
        If CheckCtrls Then 'Control setting found Ok
            ProcAcPost = True: Exit Function
        End If
        
        'A/c Posting related declarations
        Dim I As Integer, mBookDocID$
        Dim LedgAry(7) As LedgRec, mResult As Byte, mNarr$
        
        'Sale Party A/c
        mBookDocID = GCn.Execute("select OrdDocId from Veh_Order where Inv_DocId='" & Txt(TxtDocId) & "'").Fields(0).Value
        mNarr = "By Sales Invoice No." & Txt(InvPrefix) & Txt(SerialNo) & " Dt. " & Txt(Vdate) & " Chassis " & Txt(ChassisNo)
        mCommNarr = mNarr & "[Common]"
        I = 0
        LedgAry(I).SubCode = Txt(Party).Tag
        mGTotAmt = Val(Txt(GTotAmt))
        If mPostFinAmt = 0 Then
            mGTotAmt = Val(Txt(GTotAmt)) + Val(Txt(FinAmt))
        End If
        LedgAry(I).AmtDr = Round(Val(Txt(GTotAmt)), 2)
        LedgAry(I).Narration = mNarr
        'Vehicle Sale A/c
        'Modi LPS 05.12.2003
        If Val(Txt(SubTotA)) + Val(Txt(MisCharge)) - Val(Txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsForm!PurSal_Ac_Code
            LedgAry(I).AmtCr = Round(Val(Txt(SubTotA)) + Val(Txt(MisCharge)), 2)
            LedgAry(I).Narration = mNarr
        End If
        'eof Modi
        'Fitment Amount
        If Val(Txt(OthFitAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fitment_Ac
            LedgAry(I).AmtCr = Round(Val(Txt(OthFitAmt)), 2)
            LedgAry(I).Narration = mNarr & " Additional Fitments on Vehicle Sale Bill"
        End If
        'Tax Amt
        If Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(OthFitTax)) <> 0 Then
            If rsForm!Tax_Ac_Code <> "" And rsForm!Sur_Ac_Code <> "" _
                 And rsForm!Tax_Ac_Code <> rsForm!Sur_Ac_Code Then
                If Val(Txt(TaxAmt)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                    LedgAry(I).AmtCr = Round(Val(Txt(TaxAmt)) + Val(Txt(OthFitTax)), 2)
                    LedgAry(I).Narration = mNarr & " Sale Tax"
                End If
                If Val(Txt(TaxSurch)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Sur_Ac_Code
                    LedgAry(I).AmtCr = Round(Val(Txt(TaxSurch)), 2)
                    LedgAry(I).Narration = mNarr & " Surcharge on Sales Tax"
                End If
            Else
                I = I + 1
                LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                LedgAry(I).AmtCr = Round(Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(OthFitTax)), 2)
                LedgAry(I).Narration = mNarr & " Sales Tax & Surcharge"
            End If
        End If
        If Val(Txt(TOTAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = mTOT_Ac_Code
            LedgAry(I).AmtCr = Val(Txt(TOTAmt))
            LedgAry(I).Narration = mNarr & " TOT Amt"
        End If
        If Val(Txt(ROff)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!VehROff_Ac
            If Val(Txt(ROff)) > 0 Then
                LedgAry(I).AmtCr = Round(Val(Txt(ROff)), 2)
            Else
                LedgAry(I).AmtDr = Round(Abs(Val(Txt(ROff))), 2)
            End If
            LedgAry(I).Narration = mNarr & " Round Off"
        End If
        'Fuel Amount
        If Val(Txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fuel_Ac
            LedgAry(I).AmtDr = Round(Val(Txt(FuelAmt)), 2)
            LedgAry(I).Narration = mNarr & " Fuel Amount"
        End If
        
        If mPostFinAmt = 1 And Val(Txt(FinAmt)) <> 0 Then
            If Txt(FundSource) = "Hypothication" Or Txt(FundSource) = "Hire Purchase" Then
                If rsTemp!AcCode = "" Or IsNull(rsTemp!AcCode) Then
                Else
                    I = I + 1
                    LedgAry(I).SubCode = rsTemp!AcCode
                    LedgAry(I).AmtDr = Round(Val(Txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amt."
                    I = I + 1
                    LedgAry(I).SubCode = Txt(Party).Tag
                    LedgAry(I).AmtCr = Round(Val(Txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amount."
                End If
            End If
        End If
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, Txt(TxtDocId), CDate(Txt(Vdate)), mCommNarr)
        If mResult <> 1 Then
            MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
            ProcAcPost = False
        Else
            ProcAcPost = True
        End If
lblExit:
If MsgStr <> "" Then
    MsgBox MsgStr, vbCritical, "A/c Posting"
ElseIf Err.NUMBER > 0 Then
    MsgBox Err.Description, vbCritical, "A/c Posting"
End If
Set rsCtrlAc = Nothing
Set rsTemp = Nothing
End Function

Public Function GetDocIDVBill(FACn As ADODB.Connection, ByVal VType As String, ByVal Vdate As String, _
    ByRef VoucherEditFlag As Boolean, ByRef TxtSrlNo As Object, _
    ByRef lblPrefix As Object, Optional ForSiteCode As String) As String
'FACn As ADODB.Connection,
Dim Rst As ADODB.Recordset, VNo As Long, NotExists As Boolean
Dim TEMPSQL$, DivBaseNumber As Boolean, FaVoucher As Boolean
'12-04-03
'Voucher_Prefix replaced with VehBill_Counter table
'Change in connection CGN to FACn
    If FACn.Execute("Select distinct Category,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'").RecordCount <= 0 Then
        MsgBox "Please Add Record in Voucher Type Table in FA Data" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error": GetDocIDVBill = "": Exit Function
        GetDocIDVBill = ""
        GoTo errlbl
    Else
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Set Rst = FACn.Execute("Select distinct switch(Category='FA',True,Category<>'FA',False) as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
    End If
    FaVoucher = Rst!FaVoucher
    DivBaseNumber = IIf(Rst!DivBaseNumber = 0, False, True)
    If Rst.RecordCount <= 0 Then
        MsgBox "Please Define Document Numbering System  " & vbCrLf & " in Voucher Controls under Utility Menu", vbCritical, "System Configuration"
        GetDocIDVBill = ""
        GoTo errlbl
    End If
    
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    'No Division Base No. in FA /(divison base no introduced by lps at udaipur
    If FaVoucher Then MsgBox "Please Category in Voucher Type Table changed" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error": GetDocIDVBill = "": GoTo errlbl
        'Voucher No's other than FA (Division Base No possible, Voucher No. table from FAData)
        'Voucher No. From FA Data as per connection passed
        TEMPSQL = "Select Top 1 VT.Number_Method,VP.Prefix,VP.Start_Srl_No+1 as Start_Srl_No from Voucher_Type VT Left Join VehBill_Counter VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "' And VP.Prefix='" & lblPrefix & "'"
        If DivBaseNumber Then
            TEMPSQL = TEMPSQL & " and VP.Div_Code='" & PubDivCode & "'"
        End If
        TEMPSQL = TEMPSQL & " and VP.Date_From<=#" & Format(Vdate, "dd/MMM/yyyy") & "# Order By VP.Date_From DESC"
        If FACn.Execute(TEMPSQL).RecordCount > 0 Then
            Rst.Open TEMPSQL, FACn, adOpenStatic, adLockReadOnly
        Else
            'Applicable for No Records in Prefix Table & Manual Only
            'Rst.Open "Select VT.Number_Method,VT.SerialNo_From_Table,VT.V_Type From Voucher_Type VT ", FACn, adOpenDynamic, adLockOptimistic
            MsgBox "Please Add Record in Vehicle Bill Counter table " & vbCrLf & "Vehicle Sale Invoice No. Creation failed!", vbCritical, "Fatal Error": GetDocIDVBill = "": Exit Function
            GetDocIDVBill = ""
            GoTo errlbl
        End If
        '*---------
'        lblPrefix = Rst!Prefix
        If IsMissing(ForSiteCode) Then
            ForSiteCode = PubSiteCode
        ElseIf ForSiteCode = "" Then
            ForSiteCode = PubSiteCode
        End If
        If Rst!Number_Method = "Manual" Then
            VoucherEditFlag = True
            TxtSrlNo.Enabled = True
            If Val(TxtSrlNo) > 0 Then
                VNo = Val(TxtSrlNo)
            Else
                VNo = Rst!start_srl_no
            End If
        Else    'Automatic No.
            VoucherEditFlag = False
            If TopCtrl1.TopText2 = "Add" Then
                 TxtSrlNo.Enabled = True
                 VNo = Rst!start_srl_no
            Else
                TxtSrlNo.Enabled = False
                VNo = Val(Txt(SerialNo))
            End If
           
        End If
    If Val(Txt(SerialNo)) > 0 And Val(Txt(SerialNo)) < VNo Then
       VNo = Val(Txt(SerialNo))
       TxtSrlNo = VNo
    Else
        TxtSrlNo = VNo
    End If
    GetDocIDVBill = PubDivCode + PubSiteCode + ForSiteCode + Space(5 - Len(CStr(VType))) + VType + Space(5 - Len(CStr(Rst!Prefix))) + Rst!Prefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
errlbl:
    Set Rst = Nothing
End Function

