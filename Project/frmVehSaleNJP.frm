VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehSaleNJP 
   Appearance      =   0  'Flat
   BackColor       =   &H00BEE4D3&
   Caption         =   "Vehicle Sale Bill"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11820
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
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
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
      Height          =   3450
      Left            =   375
      TabIndex        =   108
      Top             =   855
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
         Index           =   13
         Left            =   3765
         TabIndex        =   117
         Top             =   960
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
         Index           =   1
         Left            =   8070
         TabIndex        =   129
         Top             =   705
         Visible         =   0   'False
         Width           =   360
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
         Left            =   7965
         TabIndex        =   128
         Top             =   750
         Visible         =   0   'False
         Width           =   465
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
         Left            =   5235
         TabIndex        =   111
         Top             =   420
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
         Index           =   5
         Left            =   5760
         MaxLength       =   15
         TabIndex        =   118
         Top             =   975
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
         Index           =   4
         Left            =   1950
         TabIndex        =   112
         Top             =   975
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
         Left            =   5775
         TabIndex        =   119
         Top             =   1245
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
         Index           =   6
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   113
         Top             =   1245
         Width           =   1770
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
         Index           =   9
         Left            =   5760
         MaxLength       =   20
         TabIndex        =   120
         Top             =   1515
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
         Index           =   8
         Left            =   1950
         TabIndex        =   114
         Top             =   1515
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
         Index           =   11
         Left            =   1950
         TabIndex        =   116
         Top             =   2055
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
         Index           =   10
         Left            =   1950
         MaxLength       =   100
         TabIndex        =   115
         Top             =   1785
         Width           =   6135
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
         Index           =   12
         Left            =   5760
         MaxLength       =   15
         TabIndex        =   121
         Top             =   2055
         Width           =   2325
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmVehSaleNJP.frx":0000
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
         TabIndex        =   122
         ToolTipText     =   "Printer "
         Top             =   2415
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmVehSaleNJP.frx":030A
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
         TabIndex        =   123
         ToolTipText     =   "Screen"
         Top             =   2745
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmVehSaleNJP.frx":0614
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
         TabIndex        =   124
         ToolTipText     =   "Printer "
         Top             =   3075
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
         Picture         =   "frmVehSaleNJP.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Screen"
         Top             =   3105
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
         Picture         =   "frmVehSaleNJP.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   126
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
         TabIndex        =   109
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
         TabIndex        =   110
         Top             =   615
         Width           =   1230
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
         Left            =   7830
         TabIndex        =   127
         Top             =   735
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temp Address"
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
         Index           =   28
         Left            =   2505
         TabIndex        =   145
         Top             =   975
         Width           =   1215
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
         Index           =   4
         Left            =   4095
         TabIndex        =   144
         Top             =   435
         Width           =   960
      End
      Begin VB.Shape Shape2 
         Height          =   1470
         Left            =   240
         Top             =   915
         Width           =   8010
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
         Left            =   0
         TabIndex        =   141
         Top             =   15
         Width           =   8085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New RTO Name"
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
         Left            =   4365
         TabIndex        =   140
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form -22 A Print"
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
         Left            =   390
         TabIndex        =   139
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temporary Y/N"
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
         Left            =   4110
         TabIndex        =   138
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Date"
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
         Left            =   390
         TabIndex        =   137
         Top             =   1260
         Width           =   810
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
         Left            =   4110
         TabIndex        =   136
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seeting Capacity"
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
         Left            =   390
         TabIndex        =   135
         Top             =   1530
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight In Printing"
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
         Left            =   390
         TabIndex        =   134
         Top             =   2070
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration"
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
         Left            =   390
         TabIndex        =   133
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temporary RTO"
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
         Index           =   57
         Left            =   4110
         TabIndex        =   132
         Top             =   2070
         Width           =   1305
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
         Top             =   3105
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
   Begin MSDataGridLib.DataGrid DGBook 
      Height          =   2175
      Left            =   1320
      Negotiate       =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   2130
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
      TabIndex        =   44
      Top             =   6690
      Width           =   1335
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   10320
      TabIndex        =   72
      Top             =   6105
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   510
         TabIndex        =   73
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
      Left            =   6420
      Negotiate       =   -1  'True
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   6750
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
      Left            =   9240
      Negotiate       =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   6975
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
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   2445
      Left            =   2010
      Negotiate       =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   7455
      Visible         =   0   'False
      Width           =   11760
      _ExtentX        =   20743
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
      Caption         =   "Chassis Help"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Chassis No"
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
      BeginProperty Column03 
         DataField       =   "TelcoNo"
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
      BeginProperty Column04 
         DataField       =   "ChlNo"
         Caption         =   "ChallanNo"
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
      BeginProperty Column06 
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
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1769.953
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
      Left            =   9450
      MaxLength       =   15
      TabIndex        =   22
      Top             =   2475
      Width           =   1665
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
      Left            =   9450
      MaxLength       =   12
      TabIndex        =   23
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
      Left            =   9450
      TabIndex        =   21
      Top             =   2205
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
      Left            =   9450
      TabIndex        =   24
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
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   6420
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
      Left            =   5475
      TabIndex        =   46
      Top             =   6960
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
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5610
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2175
      Left            =   510
      Negotiate       =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   7290
      Visible         =   0   'False
      Width           =   8670
      _ExtentX        =   15293
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
      Caption         =   "Model Help"
      ColumnCount     =   2
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
      Index           =   10
      Left            =   6510
      TabIndex        =   11
      Top             =   1665
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
      Left            =   5115
      MaxLength       =   10
      TabIndex        =   20
      Top             =   3015
      Width           =   1875
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
      Left            =   1755
      MaxLength       =   15
      TabIndex        =   19
      Top             =   3015
      Width           =   1470
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
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5610
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
      TabIndex        =   49
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
      TabIndex        =   48
      Top             =   4995
      Width           =   2325
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
      TabIndex        =   47
      Top             =   4725
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
      TabIndex        =   18
      Top             =   2745
      Width           =   1875
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
      TabIndex        =   43
      Top             =   6150
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
      Left            =   1755
      MaxLength       =   20
      TabIndex        =   14
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
      Index           =   14
      Left            =   5100
      MaxLength       =   15
      TabIndex        =   15
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
      Index           =   12
      Left            =   5100
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1935
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
      Left            =   1755
      MaxLength       =   15
      TabIndex        =   12
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
      Index           =   4
      Left            =   9465
      TabIndex        =   5
      Top             =   1860
      Width           =   1470
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
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   10
      Text            =   " "
      Top             =   1665
      Width           =   3495
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
      Left            =   1755
      MaxLength       =   50
      TabIndex        =   9
      Text            =   " "
      Top             =   1395
      Width           =   5235
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
      Left            =   1755
      MaxLength       =   50
      TabIndex        =   8
      Text            =   " "
      Top             =   1125
      Width           =   5235
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
      Left            =   1755
      MaxLength       =   50
      TabIndex        =   7
      Top             =   855
      Width           =   5235
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   8385
      Negotiate       =   -1  'True
      TabIndex        =   85
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
      TabIndex        =   31
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
      TabIndex        =   36
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
      TabIndex        =   38
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
      TabIndex        =   35
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
      TabIndex        =   39
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
      TabIndex        =   40
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   37
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
      TabIndex        =   25
      Top             =   4230
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSDataGridLib.DataGrid DGADItem 
      Height          =   4935
      Left            =   8685
      Negotiate       =   -1  'True
      TabIndex        =   71
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
      TabIndex        =   17
      Top             =   2745
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
      TabIndex        =   27
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
      TabIndex        =   28
      Text            =   "99999999.99"
      Top             =   4995
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   10245
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1590
      Width           =   975
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
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
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   6
      Top             =   585
      Width           =   5235
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
      Width           =   2100
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
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   42
      Top             =   5880
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
      TabIndex        =   32
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
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   16
      Top             =   2475
      Width           =   5235
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
      Width           =   1470
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1335
      Left            =   135
      TabIndex        =   26
      Top             =   3345
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   12243913
      Cols            =   11
      BackColorFixed  =   12632319
      ForeColorFixed  =   128
      BackColorSel    =   16703741
      BackColorBkg    =   13298928
      GridColor       =   12632319
      GridColorFixed  =   8421631
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
   Begin MSDataGridLib.DataGrid DGFin 
      Height          =   4935
      Left            =   7230
      Negotiate       =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   6840
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
      TabIndex        =   142
      Top             =   6705
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No"
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
      Left            =   4050
      TabIndex        =   107
      Top             =   1950
      Width           =   990
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
      Left            =   4590
      TabIndex        =   106
      Top             =   2220
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
      TabIndex        =   103
      Top             =   6435
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
      Left            =   3435
      TabIndex        =   102
      Top             =   6975
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
      TabIndex        =   101
      Top             =   5625
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
      Left            =   5745
      TabIndex        =   100
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Book No."
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
      Left            =   3690
      TabIndex        =   99
      Top             =   3030
      Width           =   1395
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
      Left            =   135
      TabIndex        =   98
      Top             =   3030
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
      Left            =   7050
      TabIndex        =   97
      Top             =   5625
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
      TabIndex        =   96
      Top             =   5280
      Width           =   720
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
      TabIndex        =   95
      Top             =   5010
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   31
      Left            =   7050
      TabIndex        =   94
      Top             =   4740
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
      TabIndex        =   93
      Top             =   2760
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
      TabIndex        =   92
      Top             =   6165
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
      Left            =   135
      TabIndex        =   91
      Top             =   2220
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      TabIndex        =   90
      Top             =   1950
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookng No"
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
      TabIndex        =   89
      Top             =   1875
      Width           =   915
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
      TabIndex        =   88
      Top             =   1680
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
      TabIndex        =   87
      Top             =   870
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
      Left            =   8265
      TabIndex        =   86
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
      Height          =   225
      Index           =   7
      Left            =   8280
      TabIndex        =   84
      Top             =   2760
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
      TabIndex        =   83
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
      Height          =   225
      Index           =   11
      Left            =   8280
      TabIndex        =   82
      Top             =   3030
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   3435
      X2              =   6780
      Y1              =   5550
      Y2              =   5550
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
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   77
      Top             =   5280
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code"
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
      TabIndex        =   76
      Top             =   1065
      Width           =   810
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
      TabIndex        =   75
      Top             =   5280
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax                     @"
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
      TabIndex        =   74
      Top             =   4740
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type"
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
      TabIndex        =   70
      Top             =   2490
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add./ Ded./ Short."
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
      TabIndex        =   69
      Top             =   2760
      Width           =   1410
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      Left            =   9465
      TabIndex        =   66
      Top             =   1575
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill  No."
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
      TabIndex        =   65
      Top             =   1605
      Width           =   750
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
      TabIndex        =   64
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
      TabIndex        =   63
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
      Left            =   8280
      TabIndex        =   62
      Top             =   2220
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
      TabIndex        =   61
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
      Left            =   7050
      TabIndex        =   59
      Top             =   5895
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
      TabIndex        =   58
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
      TabIndex        =   57
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
      TabIndex        =   56
      Top             =   5895
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
      TabIndex        =   55
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
      TabIndex        =   54
      Top             =   5820
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
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
      TabIndex        =   53
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      TabIndex        =   52
      Top             =   1335
      Width           =   690
   End
End
Attribute VB_Name = "frmVehSaleNJP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsChassis As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim RsMod As ADODB.Recordset
Dim rsSite As ADODB.Recordset
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

'Grid color scheme
Private Const CellBackColLeave As String = &HBAD3C9
Private Const CellForeColLeave As String = &H0&
Private Const CellBackColEnter As String = &HC0E0FF
Private Const GridBackColorBkg As String = &HCAECF0

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
Private Const SrBookNo As Byte = 19
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

Private Const SiteCode1 As Byte = 0
Private Const FromVno As Byte = 1
Private Const ToVno As Byte = 2
Private Const DocType As Byte = 3
Private Const Form22A As Byte = 4
Private Const NewRTOName As Byte = 5
Private Const PrnDate As Byte = 6
Private Const TempYN As Byte = 7
Private Const Seet As Byte = 8
Private Const Body As Byte = 9
Private Const Narr As Byte = 10
Private Const WtPrn As Byte = 11
Private Const TempRTO As Byte = 12
Private Const Tempadd As Byte = 13

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub DGBook_Click()
    DGBook.Visible = False
    If RSBook.RecordCount > 0 Then
        txt(BookNo).Text = RSBook!Code
        FillRecords
    End If
    txt(BookNo).SetFocus
End Sub

Private Sub DGFin_Click()
    DGFin.Visible = False
    If rsFin.RecordCount > 0 Then
        txt(FB_Code).Text = rsFin!Name
        txt(FB_Code).Tag = rsFin!Code
        FinAcCode = rsFin!Code
    End If
    txt(FB_Code).SetFocus
End Sub

'Inv_DocId,Inv_DocIDHelp,Inv_SiteCode,Inv_VType,Inv_No,Inv_DT,Form_Code,
'TAX_Per,TAX_Amt,Surcharge_Per,Surcharge_Amt,MARGINE,AMOUNT,REBATE
'InciChrg,Octroi,RegTemp,TransitInsu,Transport,MVT,OtherChrg,FIT_AMT,FIT_TAX
'"DieselAmt,MISC_INFO,RTO,Round_off,FB_Code,FIN_AcCode, FIN_AMT
'TrnType_Prn,Fund_Source,Chassis,Srv_BookNo

'Txt(FormType),Txt(TaxPer),Txt(TaxAmt),Txt(TaxSurPer),Txt(TaxSurch),Txt(SaleRate),Txt(NDP),Txt(Rebate)
'Txt(IncCharge),Txt(Octroi),Txt(TempReg),Txt(TransIns),Txt(Transportation),Txt(MVT),Txt(MisCharge),Txt(OthFitAmt),Txt(OthFitTax))
'Txt(FuelAmt),Txt(SpclInfo) ,Txt(RTO),Txt(ROff),Txt(FinAmt),Txt(ChassisNo),Txt(SrBookNo)
Private Sub DGVno_Click()
Dim Index As Integer
If DGVno.Tag = "1" Then
    Index = ToVno
Else
    Index = FromVno
End If
    DGVno.Visible = False
    If RsVno.RecordCount > 0 Then
        txtPrint(Index).Text = RsVno!Code
    End If
    txtPrint(Index).SetFocus
End Sub

Private Sub DGsite_Click()
If FrmPrn.Visible = False Then
    DGSite.Visible = False
    If rsSite.RecordCount > 0 Then
        txt(SiteCode).Text = rsSite!Name
        txt(SiteCode).Tag = rsSite!Code
    End If
    txt(SiteCode).SetFocus
Else
    DGSite.Visible = False
    If rsSite.RecordCount > 0 Then
        txtPrint(SiteCode1).Text = rsSite!Name
        txtPrint(SiteCode1).Tag = rsSite!Code
    End If
    txtPrint(SiteCode1).SetFocus
End If
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
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim i As Byte
TopCtrl1.Tag = UserPermission(Me.Name): WinSetting Me:     Ini_Grid

    mVType = "V_SB"
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Inv_DocId as searchcode,Veh_Order.* from Veh_Order  where Inv_vtype = '" & mVType & "'", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsSite = New ADODB.Recordset
    rsSite.CursorLocation = adUseClient
    rsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = rsSite
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct Inv_No as code from Veh_Order", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno
    
    Set RSBook = New ADODB.Recordset
    RSBook.CursorLocation = adUseClient
    RSBook.Open "Select distinct trim(str(veh_order.Ord_No)) as code,veh_order.ord_date,subgroup.Name  from Veh_Order left join subgroup on subgroup.subcode = veh_order.partycode where left(veh_order.OrdDocId,1)= '" & PubDivCode & "' and veh_order.Inv_DocId=''", GCn, adOpenDynamic, adLockOptimistic
    Set DGBook.DataSource = RSBook
    
    Set rsFin = New ADODB.Recordset
    rsFin.CursorLocation = adUseClient
    rsFin.Open "select fincode as code,finname as name,AcCode from ContractFinance where fincatg = 0  order by finname", GCn, adOpenDynamic, adLockOptimistic
    Set DGFin.DataSource = rsFin
  
    Set RsMod = New ADODB.Recordset
    RsMod.CursorLocation = adUseClient
    RsMod.Open "select Model as code,Model_Desc as NAME from Model order by Model", GCn, adOpenDynamic, adLockOptimistic
    Set DGMod.DataSource = RsMod
    
'    Set RsChassis = New ADODB.Recordset
'    RsChassis.CursorLocation = adUseClient
'    RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code, Veh_Stock.EngineNo, Veh_Stock.Srv_BookNo, Veh_Stock.VRATE, Veh_Stock.Colour_Code, (Veh_Stock.PBILL_NO+' Dt '+CStr(Veh_Stock.PBILL_DATE)) AS TelcoNo, (Right(Veh_Stock.Pur_DocId,13)+' Dt '+CStr(Veh_Stock.Pur_VDate)) AS chlNo, ColMast.Col_Desc, Veh_Stock.AL_Name FROM Veh_Stock LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
'        "where Veh_Stock.MODEL  = '" & Txt(Model) & "' and (Veh_Stock.Sal_DocId= '' or isnull(Veh_Stock.Sal_DocId))"), GCn, adOpenDynamic, adLockOptimistic

'    Set DgChassis.DataSource = RsChassis
    
    Set rsForm = New ADODB.Recordset
    With rsForm
        .CursorLocation = adUseClient
        .Open "SELECT Form_Code as Code,Form_Desc as Name,Tax_Sur_Per,Tax_Per,Tax_Ac_Code,Sur_Ac_Code,PurSal_Ac_Code FROM TaxForms where Vehicle_YN = 1 and trn_Type = 'Sale' order by Form_Desc ", GCn, adOpenDynamic, adLockOptimistic
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
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsMod = Nothing
Set RsADItem = Nothing
Set rsSite = Nothing
Set rsForm = Nothing
Set RsVno = Nothing
Set RsChassis = Nothing
Set rsFin = Nothing
Set Master = Nothing
Set mListItem = Nothing
End Sub


Private Sub ListView_Click()
If FrmPrn.Visible = False Then
    txt(Val(ListView.Tag)).Text = ListView.SelectedItem.Text
    FrmList.Visible = False
    txt(Val(ListView.Tag)).SetFocus
Else
    txtPrint(DocType).Text = ListView.SelectedItem.Text
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

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim i As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.Caption = "Division : " & PubDivCode
    LblSite.Caption = "Site Code : " & PubSiteCode
    LblVPrefix.Caption = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    txt(Vdate) = PubLoginDate
    txt(SiteCode).SetFocus
    
Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim i As Integer
Dim LedgAry(1) As LedgRec, mResult As Byte
    If GCn.Execute("Select  DelCh_DocId from  veh_order where Inv_DocId = '" & Master!SearchCode & "'").Fields(0).Value <> "" Then
        MsgBox "Delivery has been made against this Invoice", vbInformation, "deletion Denied": Exit Sub
    End If

If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then

    GCn.BeginTrans
    GCnFaV.BeginTrans
    'Unpost Ledger a/c
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, txt(TxtDocId))
    If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    'Unposting of Ledger completed
    
    GCn.Execute ("update veh_order  " & _
        "set Inv_DocId='',Inv_DocIDHelp='' ,Inv_SiteCode='',Inv_VType='',Inv_No=Null ,Inv_Date=null,Form_Code='', " & _
        "TAX_Per=0,TAX_Amt=0,Surcharge_Per=0,Surcharge_Amt=0,MARGINE=0,VRATE=0,REBATE=0, " & _
        "InciChrg=0,Octroi=0,RegTemp=0,TransitInsu=0,Transport=0,MVT=0,OtherChrg=0,FIT_AMT=0,FIT_TAX=0, " & _
        "DieselAmt=0,MISC_INFO='',RTO='',Round_off=0, " & _
        "FB_Code='' , FIN_AcCode='', FIN_AMT=0, " & _
        "TrnType_Prn=0,Fund_Source=0,Chassis='' , Srv_BookNo='', " & _
        "Inv_UName='', Inv_UEntDt=null, Inv_UAE= '' where Inv_DocId='" & txt(TxtDocId) & "'")
    
    GCn.Execute "Update Veh_Stock set Sal_DocId = '', VehSerialNo = '' where ChassisNo  = '" & txt(ChassisNo).Text & "'"
    
    For i = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(i, ADItem) <> "" Then
            GCn.Execute ("delete from veh_purch2 where DocId='" & txt(TxtDocId) & "'")
        End If
    Next
    GCnFaV.CommitTrans
    GCn.CommitTrans
    Master.Requery
    RSBook.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If err.NUMBER <> 0 Then GCn.RollbackTrans: GCnFa.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(FormType).SetFocus
    FGrid.AddItem FGrid.Rows
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
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
Dim i As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.Caption = Printer.DeviceName
txtPrint(DocType) = "Sale Bill"
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
End Sub

Private Sub TopCtrl1_eRef()
    RsMod.Requery
    rsFin.Requery
    RSBook.Requery
    rsSite.Requery
    rsForm.Requery
    RsADItem.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim i As Integer
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    Dim DocIdHlp As String
    Dim mFundSource As Byte
    Dim mTrntypeprn As Byte
'    On Error GoTo errlbl

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(txt(SiteCode), "SiteCode") = False Then Exit Sub
    If IsValid(txt(Vdate), "Bill Date") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Bill Number") = False Then Exit Sub
    If IsValid(txt(Party), "Supplier Name") = False Then Exit Sub
    If IsValid(txt(BookNo), "Booking No") = False Then Exit Sub
    If IsValid(txt(FormType), "Form Type") = False Then Exit Sub
    If IsValid(txt(Model), "Model") = False Then Exit Sub
    If IsValid(txt(ChassisNo), "Chassis") = False Then Exit Sub
    If IsValid(txt(SrBookNo), "Service Book No") = False Then Exit Sub
    If Val(txt(FinAmt)) > 0 Then
        If IsValid(txt(FundSource), "Source of Fund") = False Then Exit Sub
        If IsValid(txt(FB_Code), "Financier") = False Then Exit Sub
    End If
  If TopCtrl1.TopText2 = "Edit" Then
    If txt(SrBookNo).Text = Master!Srv_BookNo Then GoTo NXT
    If GCn.Execute("select Inv_No from veh_order where Srv_BookNo = '" & txt(SrBookNo) & "' and (inv_docid <> '' or inv_docid <> null)").RecordCount > 0 Then
       MsgBox "This Service Book no. is already issued For Site " & txt(SiteCode).Text & " through SaleBill No. -->  " _
           & GCn.Execute("select Inv_No from veh_order where Srv_BookNo = '" & txt(SrBookNo) & "' and (inv_docid <> '' or not isnull(inv_docid))").Fields(0).Value, vbInformation, "Duplicate Entry"
       Exit Sub
    End If
  Else
    If GCn.Execute("select Inv_No from veh_order where Srv_BookNo = '" & txt(SrBookNo) & "' and (inv_docid <> '' or inv_docid <> null)").RecordCount > 0 Then
        MsgBox "This Service Book no. is already issued For Site " & txt(SiteCode).Text & " through SaleBill No. -->  " _
            & GCn.Execute("select Inv_No from veh_order where Srv_BookNo = '" & txt(SrBookNo) & "' and (inv_docid <> '' or not isnull(inv_docid))").Fields(0).Value, vbInformation, "Duplicate Entry"
        Exit Sub
    End If
 End If
NXT:

    For i = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(i, ADItem) <> "" Then
            If Val(FGrid.TextMatrix(i, Qty)) = 0 Then MsgBox "Fill Quantity in Row No " & i, vbInformation, "Required data": FGrid.Row = i: FGrid.Col = Qty: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
        End If
    Next
    Amt_Cal
    If TopCtrl1.TopText2.Caption = "Add" Then
        If GCn.Execute("select count(*) from veh_order where inv_DocID='" & txt(TxtDocId) & "'").Fields(0) > 0 Then
            MsgBox "Duplicate Bill No.", vbCritical, "Validation Ertror"
            txt(TxtDocId) = VoucherNo
            Exit Sub
        End If
    Else
        txt(TxtDocId) = Master!Inv_docid
    End If
    DocIdHlp = Replace(txt(TxtDocId), " ", "")
    GCn.BeginTrans
    GCnFaV.BeginTrans
    mTrans = True
    Select Case txt(FundSource).Text
        Case "Hypothication"
            mFundSource = 0
        Case "Hire Purchase"
            mFundSource = 1
        Case "Lease"
            mFundSource = 3
        Case Else
            mFundSource = 2 'Own Fund
    End Select
    Select Case txt(ADType).Text
        Case "No Detail"
            mTrntypeprn = 0
        Case "Name/Qty"
            mTrntypeprn = 1
        Case "Name/Qty/Amount"
            mTrntypeprn = 2
    End Select
If TopCtrl1.TopText2 = "Add" Then
    GCn.Execute ("update veh_order  " & _
        "set Inv_DocId='" & DocId & "',Inv_DocIDHelp='" & DocIdHlp & "' ,Inv_SiteCode='" & PubSiteCode & txt(SiteCode).Tag & "',Inv_VType='" & mVType & "',Inv_No=" & Val(txt(SerialNo).Text) & " ,Inv_Date=" & ConvertDate(txt(Vdate)) & " ,Form_Code='" & txt(FormType).Tag & "', " & _
        "TAX_Per=" & Val(txt(TaxPer)) & ",model = '" & txt(Model).Text & "', TAX_Amt=" & Val(txt(TaxAmt)) & ",Surcharge_Per=" & Val(txt(TaxSurPer)) & ",Surcharge_Amt=" & Val(txt(TaxSurch)) & ",MARGINE=" & (Val(txt(SaleRate)) - Val(txt(NDP))) & ",VRATE=" & Val(txt(NDP)) & ",REBATE=" & Val(txt(Rebate)) & ", " & _
        "InciChrg=" & Val(txt(IncCharge)) & ",Octroi=" & Val(txt(Octroi)) & " ,RegTemp=" & Val(txt(TempReg)) & ",TransitInsu=" & Val(txt(TransIns)) & ",Transport=" & Val(txt(Transportation)) & ",MVT=" & Val(txt(MVT)) & ",OtherChrg=" & Val(txt(MisCharge)) & ",FIT_AMT=" & Val(txt(OthFitAmt)) & ",FIT_TAX=" & Val(txt(OthFitTax)) & ", " & _
        "DieselAmt=" & Val(txt(FuelAmt)) & ",MISC_INFO='" & txt(SpclInfo) & "',RTO='" & txt(RTO) & "',Round_off=" & Val(txt(ROff)) & ", " & _
        "FB_Code='" & txt(FB_Code).Tag & "' , FIN_AcCode='" & FinAcCode & "', FIN_AMT=" & Val(txt(FinAmt)) & ", " & _
        "net_Amount = " & Val(txt(GTotAmt)) & ", TrnType_Prn=" & mTrntypeprn & ",Fund_Source=" & mFundSource & ",Chassis='" & txt(ChassisNo) & "' ,Colour_Code = '" & txt(Colours).Tag & "', Srv_BookNo='" & txt(SrBookNo) & "', " & _
        "Inv_UName='" & pubUName & "', Inv_UEntDt=#" & PubServerDate & "#, Inv_UAE= 'A' where ord_no = " & Val(txt(BookNo).Text) & " and left(Ord_SiteCode,1) = '" & txt(SiteCode).Tag & "'")
        
        GCn.Execute "Update Veh_Stock set Sal_DocId = '" & DocId & "', VehSerialNo = '" & txt(SrBookNo).Text & "' where ChassisNo  = '" & txt(ChassisNo).Text & "'"
        
        GCn.Execute ("update hiscard set Name='" & txt(Party) & "',Add1='" & txt(Add1) & "', " & _
        "Add2='" & txt(Add2) & "',Add3='" & txt(Add3) & "',CityCode='" & txt(City).Tag & "', " & _
        "Govt_YN = " & IIf(txt(Govt_YN) = "Yes", 1, 0) & "  where Chassis ='" & txt(ChassisNo) & "'")
Else
    GCn.Execute ("update veh_order  " & _
        "set Form_Code='" & txt(FormType).Tag & "', " & _
        "TAX_Per=" & Val(txt(TaxPer)) & ",model = '" & txt(Model).Text & "',TAX_Amt=" & Val(txt(TaxAmt)) & ",Surcharge_Per=" & Val(txt(TaxSurPer)) & ",Surcharge_Amt=" & Val(txt(TaxSurch)) & ",MARGINE=" & (Val(txt(SaleRate)) - Val(txt(NDP))) & ",VRATE=" & Val(txt(NDP)) & ",REBATE=" & Val(txt(Rebate)) & ", " & _
        "InciChrg=" & Val(txt(IncCharge)) & ",Octroi=" & Val(txt(Octroi)) & " ,RegTemp=" & Val(txt(TempReg)) & ",TransitInsu=" & Val(txt(TransIns)) & ",Transport=" & Val(txt(Transportation)) & ",MVT=" & Val(txt(MVT)) & ",OtherChrg=" & Val(txt(MisCharge)) & ",FIT_AMT=" & Val(txt(OthFitAmt)) & ",FIT_TAX=" & Val(txt(OthFitTax)) & ", " & _
        "DieselAmt=" & Val(txt(FuelAmt)) & ",MISC_INFO='" & txt(SpclInfo) & "',RTO='" & txt(RTO) & "',Round_off=" & Val(txt(ROff)) & ", " & _
        "FB_Code='" & txt(FB_Code).Tag & "' , FIN_AcCode='" & FinAcCode & "', FIN_AMT=" & Val(txt(FinAmt)) & ", " & _
        "net_Amount = " & Val(txt(GTotAmt)) & ",TrnType_Prn=" & mTrntypeprn & ",Fund_Source=" & mFundSource & ",Chassis='" & txt(ChassisNo) & "',Colour_Code = '" & txt(Colours).Tag & "' , Srv_BookNo='" & txt(SrBookNo) & "', " & _
        "Inv_UName='" & pubUName & "', Inv_UEntDt=#" & PubServerDate & "#, Inv_UAE= 'A' where Inv_DocId='" & DocId & "'")
        
        GCn.Execute "Update Veh_Stock set VehSerialNo = '" & txt(SrBookNo).Text & "' where ChassisNo  = '" & txt(ChassisNo).Text & "'"
End If
    GCn.Execute ("delete from veh_purch2 where DocId='" & txt(TxtDocId) & "'")
    For i = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(i, ADItem) <> "" And Val(FGrid.TextMatrix(i, Qty)) <> 0 Then
            GCn.Execute ("insert into veh_purch2(DocId,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE,TAX_PER,TAX_AMT,TaxSur_Per,TaxSur_AMT, U_Name, U_EntDt, U_AE) " & _
                "values('" & txt(TxtDocId).Text & "'," & i & ",'" & PubSiteCode & txt(SiteCode).Tag & "','" & mVType & "','" & txt(SerialNo).Text & "', " & _
                "'" & FGrid.TextMatrix(i, ADItemCode) & "','A'," & Val(FGrid.TextMatrix(i, Qty)) & "," & Val(FGrid.TextMatrix(i, Rate)) & "," & Val(FGrid.TextMatrix(i, TaxPer1)) & ", " & _
                "" & Val(FGrid.TextMatrix(i, TaxAmt1)) & "," & Val(FGrid.TextMatrix(i, TaxSurPer1)) & "," & Val(FGrid.TextMatrix(i, TaxSurAmt1)) & ",'" & pubUName & "',#" & PubServerDate & "#,'E')")
        End If
    Next
    If TopCtrl1.TopText2.Caption = "Add" Then
'        If VoucherEditFlag = False Then
            UpdVouSrlNo mVType, txt(Vdate), Val(txt(SerialNo))
'            Set Rst = New ADODB.Recordset
'            Rst.CursorLocation = adUseClient
'            Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & mVType & "' And VP.Date_From<=#" & Format(txt(Vdate).Text, "dd/MMM/yyyy") & "# Order By VP.Date_From DESC", GCn, adOpenDynamic, adLockOptimistic
'            If Rst.RecordCount > 0 Then
'                GCn.Execute "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_Type & "' and Date_From=#" & Format(Rst!Date_From, "dd/MMM/yyyy") & "#"
'            End If
'        End If
    End If
    'A/c Posting
'    mNarr = "Through Vehicle Sale Bill"
'        mTOT = (AMOUNT - Rebate) + MARGINE + INCI_CHRG + Octroi + REG_TEMP + INS_TRN + Transport + MVT + TAX + SUR_AMT + OTHER
'        mTOT=IIF(ORDER->ROFYN='Y',ROUN(mTOT,0),mTOT)
'        DO &VSALE_PRO WITH ORDER->P_CODE,ORDER->PARTYCODE,mTOT,ORDER->TAX,ORDER->SUR_AMT,ORDER->DIESEL,ORDER->FIT_AMT,ORDER->FIT_TAX,ORDER->DATE1,ORDER->B_CODE1,ORDER->V_NO1,"S",mBILL_NO
'        DO &VSALE_PRO WITH ORDER->P_CODE,ORDER->PARTYCODE,(mTOT+mDIESEL),mTAX,mSUR_AMT,mDIESEL,mFIT_TOT,mFIT_TAX,mINV_DATE,mBRCO1,mINV_NO,"S",mBILL_NO
'    MsgBox "Apply New A/c Posting System"
        ProcAcPost
    'EOF of A/c Posting Section
    
GCnFaV.CommitTrans
GCn.CommitTrans
Set Rst = Nothing
mTrans = False
    Master.Requery
    RSBook.Requery
    Master.FIND "Inv_DocId = '" & txt(TxtDocId) & "'"
'    If TopCtrl1.TopText2.Caption = "Add" Then TopCtrl1_eAdd: Exit Sub
'    Disp_Text SETS("INI", Me, Master)
'    Call MoveRec
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
Exit Sub
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    'modishekhar24jan
    GSQL = "select Inv_DocId as searchcode,Inv_DocId as docid,Veh_Order.Inv_No, Veh_Order.Inv_Date, SubGroup.Name as CustomerName, City.CityName, Veh_Order.MODEL, Veh_Order.Ord_No, Veh_Order.Ord_Date " & _
        "FROM Veh_Order LEFT JOIN (SubGroup LEFT JOIN City ON SubGroup.CityCode = City.CityCode) ON Veh_Order.PartyCode = SubGroup.SubCode where Inv_vtype = '" & mVType & "'  order by Inv_DocId "
    'end modi
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
'If Index <> SiteCode Or Index <> Vdate Then If Txt(Vdate) = "" Then MsgBox "Fill Bill Date ": Txt(Vdate).SetFocus
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case ADType
        ListArray = Array("No Detail", "Name/Qty", "Name/Qty/Amount")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
    Case FundSource
        ListArray = Array("Hypothication", "Hire Purchase", "Own Fund", "Lease")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 4)
    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or txt(FB_Code).Text = "" Then Exit Sub
        If txt(FB_Code).Text <> rsFin!Name Then
            rsFin.MoveFirst
            rsFin.FIND "name ='" & txt(FB_Code).Text & "'"
        End If
        

    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or txt(BookNo).Text = "" Then Exit Sub
        If txt(BookNo).Text <> RSBook!Code Then
            RSBook.MoveFirst
            RSBook.FIND "code ='" & txt(BookNo).Text & "'"
        End If
    Case Model
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or txt(Model) = "" Then Exit Sub
        If txt(Model) <> RsMod!Code Then
            RsMod.MoveFirst
            RsMod.FIND "code ='" & txt(Model) & "'"
        End If
    Case ChassisNo
        If txt(Model) = "" Then MsgBox "Select Model First", vbInformation, "Validation": txt(Model).SetFocus: Exit Sub
        Set RsChassis = GCn.Execute("SELECT Veh_Stock.ChassisNo as code, Veh_Stock.EngineNo, Veh_Stock.Srv_BookNo, Veh_Stock.VRATE, Veh_Stock.Colour_Code, (Veh_Stock.PBILL_NO+' Dt '+ iif(isnull(Veh_Stock.PBILL_DATE),'',CStr(Veh_Stock.PBILL_DATE))) AS TelcoNo, (Right(Veh_Stock.Pur_DocId,13)+' Dt '+ iif(isnull(Veh_Stock.Pur_VDate),'',CStr(Veh_Stock.Pur_VDate))) AS chlNo, ColMast.Col_Desc, Veh_Stock.AL_Name,Veh_Stock.tax_yn,Veh_Stock.RSO_WORK,Veh_Stock.PBILL_NO,Veh_Stock.PBILL_DATE,Veh_Stock.INDATE " & _
            " FROM Veh_Stock LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code where Veh_Stock.MODEL  = '" & txt(Model) & "' and (Veh_Stock.Sal_DocId= '' or isnull(Veh_Stock.Sal_DocId))")
        Set DgChassis.DataSource = RsChassis
    Case SiteCode
        Set DGSite.DataSource = rsSite
        If rsSite.RecordCount = 0 Or (rsSite.EOF = True Or rsSite.BOF = True) Then Exit Sub
        If txt(Index).Text = "" Then
            rsSite.MoveFirst
            rsSite.FIND "code ='" & PubSiteCode & "'"
            txt(Index).Tag = rsSite!Code
            txt(Index).Text = rsSite!Name
        Else
            If txt(Index).Text <> rsSite!Name Then
                rsSite.MoveFirst
                rsSite.FIND "name ='" & txt(Index).Text & "'"
            End If
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).Text = "" Then Exit Sub
        If txt(Index).Text <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & txt(Index).Text & "'"
        End If
    Case SerialNo, TaxAmt, TaxSurch, TaxPer, TaxSurPer, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation
        SendKeys "{HOME}+{END}"
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
    Case BookNo
         DGridTxtKeyDown DGBook, txt, Index, RSBook, KeyCode, False, 0, frmVehBook
    Case ADType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).Height), txt(Index).width, 900
    Case FundSource
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).Height), txt(Index).width, 1200
    Case SiteCode
        DGridTxtKeyDown DGSite, txt, Index, rsSite, KeyCode, False, 1
    Case Model
        DGridTxtKeyDown DGMod, txt, Index, RsMod, KeyCode, False, 0, frmModel
        If DGMod.Visible = True Then txt(ChassisNo).Text = ""
    Case ChassisNo
        DGridTxtKeyDown DgChassis, txt, Index, RsChassis, KeyCode, False, 0
    Case FormType
        DGridTxtKeyDown DGForm, txt, FormType, rsForm, KeyCode, False, 1, frmTaxForms
    Case FB_Code
        DGridTxtKeyDown DGFin, txt, Index, rsFin, KeyCode, False, 1, frmFinMast
    Case Model
        DGridTxtKeyDown DGMod, txt, Index, RsMod, KeyCode, False, 0, frmModel
    Case ChassisNo
        DGridTxtKeyDown DgChassis, txt, Index, RsChassis, KeyCode, False, 0
End Select
If FrmList.Visible = False And DGBook.Visible = False And DGSite.Visible = False And DGFin.Visible = False And DgChassis.Visible = False And DGMod.Visible = False And DGForm.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Vdate Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> FB_Code Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = FB_Code Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.Caption = "Add" And Index <> SiteCode Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.Caption = "Edit" And Index <> BookNo Then
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
        If DGBook.Visible = True Then DGridTxtKeyPress txt, Index, RSBook, KeyAscii, "Code"
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress txt, Index, rsSite, KeyAscii, "Name"
    Case Model
        If DGMod.Visible = True Then DGridTxtKeyPress txt, Index, RsMod, KeyAscii, "code"
    Case ChassisNo
        If DgChassis.Visible = True Then DGridTxtKeyPress txt, Index, RsChassis, KeyAscii, "code", False
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, FormType, rsForm, KeyAscii, "Name"
    Case SerialNo
        Call NumPress(txt(Index), KeyAscii, 6, 0)
    Case TaxAmt, TaxSurch, MisCharge, FinAmt, FuelAmt
        Call NumPress(txt(Index), KeyAscii, 8, 2)
    Case TaxPer, TaxSurPer
        Call NumPress(txt(Index), KeyAscii, 3, 2)
End Select
'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs

Select Case Index
    Case FundSource, ADType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case FormType
        If DGForm.Visible = True Then
            If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).Text = "" Then Exit Sub
            txt(TaxPer).Text = IIf(IsNull(rsForm!Tax_Per), 0, rsForm!Tax_Per)
            txt(TaxAmt).Text = Val(txt(SubTotA).Text) * Val(txt(TaxPer).Text) / 100
            txt(TaxSurPer).Text = IIf(IsNull(rsForm!Tax_Sur_Per), 0, rsForm!Tax_Sur_Per)
            txt(TaxSurch).Text = Val(txt(TaxSurPer).Text) * Val(txt(TaxAmt).Text) / 100
            Amt_Cal
        End If
    Case TaxPer
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
        txt(TaxAmt).Text = Format(Val(txt(SubTotA).Text) * Val(txt(TaxPer).Text) / 100, "0.00")
        txt(TaxSurch).Text = Format(Val(txt(TaxSurPer).Text) * Val(txt(TaxAmt).Text) / 100, "0.00")
        Amt_Cal
    Case TaxSurPer
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
        txt(TaxSurch).Text = Format(Val(txt(TaxSurPer).Text) * Val(txt(TaxAmt).Text) / 100, "0.00")
        Amt_Cal
    Case SaleRate, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation
         Amt_Cal
    Case MisCharge, OthFitAmt, OthFitTax, FinAmt, AdvAmt, FuelAmt
         Amt_Cal
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim i As Integer
Select Case Index
    Case FundSource, ADType
        If txt(Index).Text <> "" Then txt(Index).Text = ListView.SelectedItem.Text
    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or txt(Index).Text = "" Then
            txt(Index).Text = ""
        Else
            txt(Index).Text = RSBook!Code
        End If
        If FillRecords = False Then Cancel = True
    Case Model
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or txt(Index).Text = "" Then
            txt(Index).Text = ""
            txt(Index).Tag = ""
        Else
            txt(Index).Text = RsMod!Code
            txt(Index).Tag = RsMod!Code
        End If
        If IsValid(txt(Index), "Model") = False Then Cancel = True: GoTo lblExitSub
        txt(ChassisNo).SetFocus
    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or txt(Index).Text = "" Then
            txt(Index).Text = ""
            txt(Index).Tag = ""
            FinAcCode = ""
        Else
            txt(Index).Text = rsFin!Name
            txt(Index).Tag = rsFin!Code
            FinAcCode = rsFin!AcCode
        End If
    Case ChassisNo
        IsValid txt(Index), "ChassisNo.", True
        If RsChassis.RecordCount = 0 Or (RsChassis.EOF = True Or RsChassis.BOF = True) Or txt(Index).Text = "" Then
            txt(ChassisNo) = ""
            Fill_Data False
        Else
            txt(ChassisNo) = RsChassis!Code
            Fill_Data True
'            If UCase(Trim(Txt(Index).Text)) <> UCase(RsChassis!Code) Then
'                If Txt(ChassisNo) <> Txt(Index).Text Then Fill_Data False
'                Txt(ChassisNo) = Txt(Index).Text
'            Else
'               Txt(ChassisNo) = Txt(Index).Text
'               Fill_Data True
'            End If
        End If
'        If Txt(ChassisNo).Text = "" And RsChassis.RecordCount > 0 Then
'            MsgBox "chassis no is required", vbInformation, "Validation  Check"
'            Cancel = True: GoTo lblExitSub
'        End If
    Case SiteCode
        If IsValid(txt(Index), "Site Code") = False Then Cancel = True: GoTo lblExitSub
        If rsSite.RecordCount = 0 Or (rsSite.EOF = True Or rsSite.BOF = True) Or txt(Index).Text = "" Then
            txt(Index).Text = ""
            txt(Index).Tag = ""
        Else
            txt(Index).Text = rsSite!Name
            txt(Index).Tag = rsSite!Code
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).Text = "" Then
            txt(Index).Text = ""
            txt(Index).Tag = ""
        Else
            txt(Index).Text = rsForm!Name
            txt(Index).Tag = rsForm!Code
        End If
    Case Vdate
        If Len(Trim(txt(Vdate).Text)) = 0 Then
             txt(Vdate).Text = PubLoginDate
        Else
            txt(Index).Text = RetDate(txt(Index))
        End If
        txt(TxtDocId).Text = VoucherNo
    Case SerialNo
        If IsValid(txt(SerialNo), "Serial No.") = False Then Cancel = True:  GoTo lblExitSub
        If VoucherEditFlag = True Then      ' Manual
            DocId = VoucherNo
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select * From veh_purch1 Where docid='" & DocId & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(SerialNo).SetFocus
            End If
        End If
    Case TaxPer, TaxAmt, TaxSurPer, TaxSurch, SaleRate, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation, MisCharge, OthFitAmt, OthFitTax, FinAmt, AdvAmt, FuelAmt
         txt(Index).Text = Format(txt(Index).Text, "0.00")
         Amt_Cal
End Select
lblExitSub:
Set Rst = Nothing
End Sub

Private Sub DGADItem_Click()
    DGADItem.Visible = False
    If RsADItem.RecordCount > 0 Then
        TxtGrid(0).Text = RsADItem!Name
         FGrid.TextMatrix(FGrid.Row, ADItem) = RsADItem!Name
         FGrid.TextMatrix(FGrid.Row, ADItemCode) = RsADItem!Code
    End If
    TxtGrid(0).SetFocus
End Sub
Private Sub DgChassis_Click()
    DgChassis.Visible = False
    If RsChassis.RecordCount > 0 Then
        txt(ChassisNo).Text = RsChassis!Code
        Fill_Data True
    End If
    txt(ChassisNo).SetFocus
End Sub
Private Sub DGMod_Click()
DGMod.Visible = False
If RsMod.RecordCount > 0 Then
    txt(Model) = RsMod!Code
End If
    txt(Model).SetFocus
End Sub

Private Sub DGForm_Click()
    DGForm.Visible = False
    If rsForm.RecordCount > 0 Then
        txt(FormType).Text = rsForm!Name
        txt(FormType).Tag = rsForm!Code
    End If
    txt(FormType).SetFocus
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
For i = 0 To txt.Count - 1
    txt(i).Text = ""
Next i
End Sub

Private Sub MoveRec()
Dim Rst As Recordset
Dim i As Integer
On Error GoTo error1
If Master.RecordCount > 0 Then
    DocId = Master!Inv_docid
    txt(TxtDocId).Text = Master!Inv_docid
    LblDiv.Caption = "Division : " & left(Master!Inv_docid, 1)
    LblSite.Caption = "Site Code : " & Mid(Master!Inv_SiteCode, 1, 1)
    txt(SiteCode).Tag = Mid(Master!Inv_SiteCode, 2, 1)
    txt(SiteCode).Text = GCn.Execute("select site_desc from site where site_code = '" & txt(SiteCode).Tag & "'").Fields(0).Value
    LblVPrefix.Caption = Mid(Master!Inv_docid, 8, 5)
    txt(SerialNo).Text = Master!inv_No
    txt(Vdate).Text = Master!inv_date
    txt(BookNo).Text = Master!ord_no
    If Not IsNull(Master!Fund_Source) Then
        Select Case Master!Fund_Source
            Case 0 '0 Hypothication ,1 Hire purchase ,2 Own Fund,3 Lease
                txt(FundSource).Text = "Hypothication"
            Case 1
                txt(FundSource).Text = "Hire Purchase"
'            Case 2
'                txt(FundSource).Text = "Own Fund"
            Case 3
                txt(FundSource).Text = "Lease"
            Case Else
                txt(FundSource).Text = "Own Fund"
        End Select
    Else
        txt(FundSource).Text = ""
    End If
    If Not IsNull(Master!TrnType_Prn) Then
        Select Case Master!TrnType_Prn
            Case 0
                txt(ADType).Text = "No Detail"
            Case 1
                txt(ADType).Text = "Name/Qty"
            Case 2
                txt(ADType).Text = "Name/Qty/Amount"
        End Select
    Else
        txt(ADType).Text = ""
    End If
    
    txt(Party).Tag = IIf(IsNull(Master!PartyCode), "", Master!PartyCode)
    If txt(Party).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select name,add1,add2,add3,CityCode from SubGroup where Subcode = '" & txt(Party).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        txt(Party).Text = Rst!Name
        txt(Add1).Text = IIf(IsNull(Rst!Add1), "", Rst!Add1)
        txt(Add2).Text = IIf(IsNull(Rst!Add2), "", Rst!Add2)
        txt(Add3).Text = IIf(IsNull(Rst!Add3), "", Rst!Add3)
        txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
        If txt(City).Tag <> "" Then
            txt(City).Text = GCn.Execute("select cityname from city where citycode = '" & txt(City).Tag & "'").Fields(0).Value
        End If
    End If
    txt(Model).Text = Master!Model
    txt(Govt_YN).Text = IIf(Master!Govt_YN = 1, "Yes", "No")
    txt(Colours).Tag = IIf(IsNull(Master!Colour_Code), "", Master!Colour_Code)
    If txt(Colours).Tag <> "" Then
        txt(Colours).Text = GCn.Execute("select col_desc from colmast where col_code = '" & txt(Colours).Tag & "'").Fields(0).Value
    End If
    
    txt(FormType).Tag = IIf(IsNull(Master!form_Code), "", Master!form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType).Text = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType).Text = ""
    End If
    txt(NDP).Text = Format(IIf(IsNull(Master!VRATE), 0, Master!VRATE), "0.00")
    
    txt(SaleRate).Text = Format(Val(txt(NDP)) + Val(IIf(IsNull(Master!MARGINE), 0, Master!MARGINE)), "0.00")
    txt(Rebate).Text = Format(IIf(IsNull(Master!Rebate), 0, Master!Rebate), "0.00")
    txt(IncCharge).Text = Format(IIf(IsNull(Master!InciChrg), 0, Master!InciChrg), "0.00")
    txt(Octroi).Text = Format(IIf(IsNull(Master!Octroi), 0, Master!Octroi), "0.00")
    txt(TempReg).Text = Format(IIf(IsNull(Master!RegTemp), 0, Master!RegTemp), "0.00")
    txt(TransIns).Text = Format(IIf(IsNull(Master!TransitInsu), 0, Master!TransitInsu), "0.00")
    txt(MVT).Text = Format(IIf(IsNull(Master!MVT), 0, Master!MVT), "0.00")
    txt(Transportation).Text = Format(IIf(IsNull(Master!Transport), 0, Master!Transport), "0.00")
    txt(SubTotA) = Format((Val(txt(SaleRate)) - Val(txt(Rebate)) + Val(txt(IncCharge)) + Val(txt(Octroi)) + Val(txt(TempReg)) + Val(txt(TransIns)) + Val(txt(MVT)) + Val(txt(Transportation))), "0.00")
    
    txt(TaxPer).Text = Format(IIf(IsNull(Master!Tax_Per), 0, Master!Tax_Per), "0.00")
    txt(TaxAmt).Text = Format(IIf(IsNull(Master!Tax_Amt), 0, Master!Tax_Amt), "0.00")
    txt(TaxSurPer).Text = Format(IIf(IsNull(Master!surcharge_per), 0, Master!surcharge_per), "0.00")
    txt(TaxSurch).Text = Format(IIf(IsNull(Master!Surcharge_Amt), 0, Master!Surcharge_Amt), "0.00")
    txt(MisCharge).Text = Format(IIf(IsNull(Master!OtherChrg), 0, Master!OtherChrg), "0.00")
    txt(SubTotB) = Format((Val(txt(SubTotA)) + Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(MisCharge))), "0.00")
    
    txt(OthFitAmt).Text = Format(IIf(IsNull(Master!FIT_AMT), 0, Master!FIT_AMT), "0.00")
    txt(OthFitTax).Text = Format(IIf(IsNull(Master!FIT_TAX), 0, Master!FIT_TAX), "0.00")
    txt(FuelAmt).Text = Format(IIf(IsNull(Master!DieselAmt), 0, Master!DieselAmt), "0.00")
    txt(ROff).Text = Format(IIf(IsNull(Master!Round_off), 0, Master!Round_off), "0.00")
    txt(GTotAmt) = Format((Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)) + Val(txt(ROff))), "0.00")
    
    Dim RstTemp As ADODB.Recordset
    Set RstTemp = GCn.Execute("Select sum(iif(DrCr ='C',Amount,Amount*-1)) as AdvAmt from Rect where mid(Ord_DocId,14,8) = " & Val(txt(BookNo)) & " and left(Ord_SiteCode,1)= '" & txt(SiteCode).Tag & "'")
    If RstTemp.RecordCount > 0 Then
        txt(AdvAmt) = Format(IIf(IsNull(RstTemp!AdvAmt), 0, RstTemp!AdvAmt), "0.00")
    End If
 
  ' Txt(AdvAmt) = Format(IIf(IsNull(Master!P_Amount), 0, Master!P_Amount), "0.00")
    txt(NetOStng) = Format((Val(txt(GTotAmt)) - Val(txt(AdvAmt))), "0.00")
    
    txt(FinAmt).Text = Format(IIf(IsNull(Master!FIN_AMT), 0, Master!FIN_AMT), "0.00")
    txt(FB_Code).Tag = IIf(IsNull(Master!FB_Code), "", Master!FB_Code)
    If txt(FB_Code).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select fincode as code,finname as name,AcCode from ContractFinance where fincatg = 0 and  fincode = '" & txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        txt(FB_Code).Text = Rst!Name
        FinAcCode = IIf(IsNull(Rst!AcCode), "", Rst!AcCode)
    Else
        txt(FB_Code).Text = ""
        FinAcCode = ""
    End If
    
    txt(SpclInfo).Text = IIf(IsNull(Master!MISC_INFO), "", Master!MISC_INFO)
    txt(RTO).Text = IIf(IsNull(Master!RTO), "", Master!RTO)
    txt(SrBookNo).Text = IIf(IsNull(Master!Srv_BookNo), "", Master!Srv_BookNo)
    txt(ChassisNo).Text = IIf(IsNull(Master!Chassis), "", Master!Chassis)
    
    Set Rst = New Recordset
    Rst.Open "SELECT Veh_Stock.EngineNo,Veh_Stock.VehSerialNo,Veh_Stock.tax_yn,Veh_Stock.PBILL_NO,Veh_Stock.PBILL_DATE FROM Veh_Stock where Veh_Stock.MODEL  = '" & txt(Model) & "' and Veh_Stock.ChassisNo = '" & txt(ChassisNo) & "' and Veh_Stock.Sal_DocId= '" & Master!Inv_docid & "'", GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
        txt(EngineNo).Text = IIf(IsNull(Rst!EngineNo), "", Rst!EngineNo)
        txt(TelcoInvNo).Text = IIf(IsNull(Rst!PBILL_NO), "", Rst!PBILL_NO)
        txt(TelcoInvDate).Text = IIf(IsNull(Rst!PBILL_DATE), "", Rst!PBILL_DATE)
        txt(Taxable).Text = IIf(Rst!TAX_YN = 1, "Yes", "No")
        txt(SrBookNo).Text = IIf(IsNull(Rst!VehSerialNo), "", Rst!VehSerialNo)
    End If
    Set Rst = New Recordset
    Set Rst = GCn.Execute("SELECT Veh_AMDModel.Prod_Name, Veh_Purch2.Srl_No, Veh_Purch2.PROD_CODE, Veh_Purch2.QTY, Veh_Purch2.RATE,Veh_Purch2.TAX_PER,Veh_Purch2.TAX_AMT,Veh_Purch2.TaxSur_Per,Veh_Purch2.TaxSur_AMT " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_purch2.DocId = '" & Master!Inv_docid & "'")
    FGrid.Rows = 1
    i = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            With FGrid
                .AddItem ""
                .TextMatrix(i, 0) = Rst!Srl_No
                .TextMatrix(i, ADItem) = Rst!Prod_Name
                .TextMatrix(i, Qty) = Format(IIf(IsNull(Rst!Qty), "", Rst!Qty), "0")
                .TextMatrix(i, Rate) = Format(IIf(IsNull(Rst!Rate), "", Rst!Rate), "0.00")
                .TextMatrix(i, Amt) = Format(.TextMatrix(i, Qty) * .TextMatrix(i, Rate), "0.00")
                .TextMatrix(i, TaxPer1) = Format(IIf(IsNull(Rst!Tax_Per), "", Rst!Tax_Per), "0.00")
                .TextMatrix(i, TaxAmt1) = Format(IIf(IsNull(Rst!Tax_Amt), "", Rst!Tax_Amt), "0.00")
                .TextMatrix(i, TaxSurPer1) = Format(IIf(IsNull(Rst!TaxSur_Per), "", Rst!TaxSur_Per), "0.00")
                .TextMatrix(i, TaxSurAmt1) = Format(IIf(IsNull(Rst!TaxSur_Amt), "", Rst!TaxSur_Amt), "0.00")
                .TextMatrix(i, FinalAmt) = Format((Val(.TextMatrix(i, Amt)) + Val(.TextMatrix(i, TaxAmt1)) + Val(.TextMatrix(i, TaxSurAmt1))), "0.00")
                .TextMatrix(i, ADItemCode) = Rst!Prod_Code
            End With
            Rst.MoveNext
           i = i + 1
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
Grid_Hide
Amt_Cal
Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
Dim i As Byte
DGFin.left = 6630: DGFin.top = mTopScale
DgChassis.left = Me.left + 45: DgChassis.top = 3450
DGMod.left = Me.left + 45: DGMod.top = 3450
DGSite.left = 4260: DGSite.top = mTopScale
DGForm.left = 6630: DGForm.top = mTopScale
DGADItem.left = 6630: DGADItem.top = mTopScale
DGVno.left = 4260: DGVno.top = mTopScale
DGBook.left = 1320: DGBook.top = txt(BookNo).top + txt(BookNo).Height + 20
    
    With FGrid
        .left = Me.left '+45
        .top = 3345
        .Cols = 11
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight

        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, ADItem) = "Additional Fitments"
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
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim i As Integer
For i = 0 To txt.Count - 1
    txt(i).Enabled = Enb
    txt(i).ForeColor = CtrlFColOrg
Next

If TopCtrl1.TopText2 = "Edit" Then
    txt(SiteCode).Enabled = False
    txt(Vdate).Enabled = False
    txt(SerialNo).Enabled = False
    txt(BookNo).Enabled = False
    If GCn.Execute("select DelCh_DocId from veh_stock where chassisNo ='" & txt(ChassisNo) & "'").Fields(0).Value = "" Then
        txt(Model).Enabled = True
        txt(ChassisNo).Enabled = True
    Else
        txt(Model).Enabled = False
        txt(ChassisNo).Enabled = False
    End If
End If

txt(TxtDocId).Enabled = False
txt(Taxable).Enabled = False
txt(Party).Enabled = False
txt(Add1).Enabled = False
txt(Add2).Enabled = False
txt(Add3).Enabled = False
txt(City).Enabled = False
txt(Govt_YN).Enabled = False
txt(TelcoInvNo).Enabled = False
txt(TelcoInvDate).Enabled = False
txt(EngineNo).Enabled = False
txt(Colours).Enabled = False
txt(NDP).Enabled = False
txt(SubTotA).Enabled = False
txt(SubTotB).Enabled = False
txt(TaxPer).Enabled = False: txt(TaxAmt).Enabled = False
txt(TaxSurPer).Enabled = False: txt(TaxSurch).Enabled = False
txt(ROff).Enabled = False
txt(GTotAmt).Enabled = False
txt(NetOStng).Enabled = False

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
    If DGVno.Visible = True Then DGVno.Visible = False
End Sub

Public Function VoucherNo() As String
Dim Rst As ADODB.Recordset, VouType As String, VNo As Long
    VouType = mVType
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VouType & "' And VP.Date_From<=#" & Format(txt(Vdate).Text, "dd/MMM/yyyy") & "# Order By VP.Date_From DESC", GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
        If Rst!Number_Method = "Manual" Then
             VoucherEditFlag = True
             vPrefix = Rst!Prefix
            If txt(SerialNo).Text = "" Then
                VNo = Rst!start_srl_no + 1
            Else
                VNo = Val(txt(SerialNo).Text)
            End If
             txt(SerialNo).Enabled = True
             txt(SerialNo).BackColor = CtrlBColOrg
        Else
            txt(SerialNo).Enabled = False
            txt(SerialNo).BackColor = CtrlBColDisabled
            vPrefix = Rst!Prefix
            VNo = Rst!start_srl_no + 1
            VoucherEditFlag = False
        End If
    End If
    DocId = PubDivCode + PubSiteCode & txt(SiteCode).Tag + Space(5 - Len(CStr(VouType))) + VouType + Space(5 - Len(CStr(vPrefix))) + vPrefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
    LblVPrefix.Caption = vPrefix
    txt(TxtDocId).Text = DocId
    txt(SerialNo).Text = VNo
    VoucherNo = DocId
Set Rst = Nothing
End Function

Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
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
            If KeyCode = vbKeyEscape Then
                TxtGrid(0).Text = TxtGrid(0).Tag
                TxtGrid_KeyUp Index, KeyCode, Shift
                TxtGrid(0).Visible = False
                Grid_Hide
                FGrid.SetFocus
                Exit Sub
            End If
            Select Case FGrid.Col
 
                Case ADItem    '1
                    DGridTxtKeyDown DGADItem, TxtGrid, Index, RsADItem, KeyCode, True, 1, frmVehAMDMast
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
                FGrid.TextMatrix(FGrid.Row, Qty) = Format(Val(TxtGrid(Index).Text), "0")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxSurPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case Rate
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(Index).Text), "0.00")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxSurPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxAmt1
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(TxtGrid(Index).Text), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                Else
                    FGrid.TextMatrix(FGrid.Row, TaxPer1) = Format((100 * Val(TxtGrid(Index).Text)) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxPer1
                FGrid.TextMatrix(FGrid.Row, TaxPer1) = TxtGrid(Index).Text
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(TxtGrid(Index).Text) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxSurAmt1
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(TxtGrid(Index).Text), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                Else
                   FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = Format((100 * Val(TxtGrid(Index).Text)) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
            Case TaxSurPer1
                FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = TxtGrid(Index).Text
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(TxtGrid(Index).Text) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal
        End Select
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Dim j As Integer
Select Case FGrid.Col
        Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or TxtGrid(0).Text = "" Then
                FGrid.TextMatrix(FGrid.Row, ADItem) = ""
                FGrid.TextMatrix(FGrid.Row, ADItemCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, ADItemCode) = RsADItem!Code
                FGrid.TextMatrix(FGrid.Row, ADItem) = RsADItem!Name
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(IIf(IsNull(RsADItem!Rate), 0, RsADItem!Rate), "0.00")
            End If
            
            If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
        Case Rate, TaxPer1, TaxAmt1, TaxSurPer1, TaxSurAmt1
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).Text), "0.00")
                Amt_Cal
        Case Qty
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).Text), "0")
                Amt_Cal
End Select
End Sub

Private Function TxtGridLeave() As Boolean
Dim j As Integer
Dim GridCol As Byte
GridCol = FGrid.Col
Select Case GridCol
        Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or TxtGrid(0).Text = "" Then
                FGrid.TextMatrix(FGrid.Row, ADItem) = ""
                FGrid.TextMatrix(FGrid.Row, ADItemCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, ADItemCode) = RsADItem!Code
                FGrid.TextMatrix(FGrid.Row, ADItem) = RsADItem!Name
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(IIf(IsNull(RsADItem!Rate), 0, RsADItem!Rate), "0.00")
            End If
            If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
        Case Rate, TaxPer1, TaxAmt1, TaxSurPer1, TaxSurAmt1
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).Text), "0.00")
                Amt_Cal
        Case Qty
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).Text), "0")
                Amt_Cal
End Select
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function

Private Sub FGrid_Click()
If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell--> Enter Cell-->KeyDown
If TopCtrl1.TopText2.Caption = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    FGrid.CellBackColor = CellBackColLeave
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
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "0.000"
            FGrid.TextMatrix(FGrid.Row, Amt) = "0.00"
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = "0.00"
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
            FGrid.TextMatrix(FGrid.Row, FinalAmt) = "0.00"
        Case Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "0.00"
            FGrid.TextMatrix(FGrid.Row, Amt) = "0.00"
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = "0.00"
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
            FGrid.TextMatrix(FGrid.Row, FinalAmt) = "0.00"
        Case TaxPer1, TaxAmt1
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = "0.00"
            FGrid.TextMatrix(FGrid.Row, TaxPer1) = "0.00"
        Case TaxSurPer1, TaxSurAmt1
            FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
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

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case ADItem, Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub

Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
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
         For i = 1 To FGrid.Rows - 1
            FGrid.TextMatrix(i, 0) = i
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If

FGrid.SetFocus
End If
Exit Sub
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case ADItem
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
        Case Amt
            FGrid_LeaveCell
            FGrid.Col = FGrid.Col + 1
            FGrid_EnterCell
            FGrid.SetFocus
        Case Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
           Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
'    FGrid.CellForeColor = CellForeColLeave
End Sub
 
Private Sub Amt_Cal()
Dim i As Byte
Dim Tottax As Double
Dim TotAdd As Double
For i = 1 To FGrid.Rows - 1
   If FGrid.TextMatrix(i, ADItem) <> "" Then
           TotAdd = TotAdd + Val(FGrid.TextMatrix(i, Amt))
           Tottax = Tottax + Val(FGrid.TextMatrix(i, TaxAmt1)) + Val(FGrid.TextMatrix(i, TaxSurAmt1))
   End If
Next
txt(OthFitAmt) = Format(TotAdd, "0.00")
txt(OthFitTax) = Format(Tottax, "0.00")
'SaleRate, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportatio
'MisCharge, OthFitAmt, OthFitTax, FinAmt, AdvAmt, FuelAmt
'SubTotA , TaxAmt, TaxSurch, SubTotB, ROff, GTotAmt, NetOStng
txt(SubTotA) = Format((Val(txt(SaleRate)) - Val(txt(Rebate)) + Val(txt(IncCharge)) + Val(txt(Octroi)) + Val(txt(TempReg)) + Val(txt(TransIns)) + Val(txt(MVT)) + Val(txt(Transportation))), "0.00")
txt(TaxAmt).Text = Format(Val(txt(SubTotA).Text) * Val(txt(TaxPer).Text) / 100, "0.00")
txt(TaxSurch).Text = Format(Val(txt(TaxSurPer).Text) * Val(txt(TaxAmt).Text) / 100, "0.00")
txt(SubTotB) = Format((Val(txt(SubTotA)) + Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(MisCharge))), "0.00")
txt(ROff) = dmRoundOff(Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)))
txt(GTotAmt) = Format((Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)) + Val(txt(ROff))), "0.00")
txt(NetOStng) = Format((Val(txt(GTotAmt)) - Val(txt(AdvAmt))), "0.00")
End Sub

Private Sub Fill_Data(Enb As Boolean)
Dim Rst As ADODB.Recordset
Dim Margin As Double
If Enb = True Then
    If IsNull(RsChassis!InDate) Then
        If MsgBox("Vehicle In Transit Continue Yes/No ? ", vbYesNo + vbCritical + vbDefaultButton2, "Check") = vbNo Then GoTo NXT
    End If
    txt(EngineNo) = IIf(IsNull(RsChassis!EngineNo), "", RsChassis!EngineNo)
    txt(SrBookNo) = IIf(IsNull(RsChassis!Srv_BookNo), "", RsChassis!Srv_BookNo)
    txt(Colours) = IIf(IsNull(RsChassis!Col_Desc), "", RsChassis!Col_Desc)
    txt(Colours).Tag = IIf(IsNull(RsChassis!Colour_Code), "", RsChassis!Colour_Code)
    txt(TelcoInvDate).Text = IIf(IsNull(RsChassis!PBILL_DATE), "", RsChassis!PBILL_DATE)
    txt(TelcoInvNo).Text = IIf(IsNull(RsChassis!PBILL_NO), "", RsChassis!PBILL_NO)
    txt(Taxable).Text = IIf(RsChassis!TAX_YN = 1, "Yes", "No")
    txt(NDP).Text = Format(IIf(IsNull(RsChassis!VRATE), "", RsChassis!VRATE), "0.00")
    
    
    Set Rst = New Recordset
    Rst.Open "Select P_RATE,s_rate,INCI_CHRG,OCTROI,REG_TEMP,INS_TRN,TRANSPORT,MVT,REG_FEE,INS_FEE from veh_rate where model = '" & txt(Model).Text & "' and Effective_Date <= " & ConvertDate(txt(Vdate)) & " and RSO_WORK = " & RsChassis!RSO_WORK & " and TAXABLE_YN = " & RsChassis!TAX_YN & " order by Effective_Date DESC", GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
         If Val(txt(NDP).Text) = 0 Then
            txt(NDP).Text = IIf(IsNull(Rst!p_rate), 0, Rst!p_rate)
         End If
         Margin = IIf(IsNull(Rst!S_Rate), 0, Rst!S_Rate) - IIf(IsNull(Rst!p_rate), 0, Rst!p_rate)
         txt(SaleRate).Text = Format((Val(txt(NDP)) + Margin), "0.00")
    Else
         Margin = 0
         txt(SaleRate).Text = Format((Val(txt(NDP)) + Margin), "0.00")
    End If
    If Rst.RecordCount > 0 Then
        txt(IncCharge) = Format(Rst!INCI_CHRG, "0.00")
        txt(Octroi) = Format(Rst!Octroi, "0.00")
        txt(TempReg) = Format(Rst!REG_TEMP, "0.00")
        txt(TransIns) = Format(Rst!INS_TRN, "0.00")
        txt(MVT) = Format(Rst!MVT, "0.00")
        txt(Transportation) = Format(Rst!Transport, "0.00")
    Else
        txt(IncCharge) = Format(0, "0.00")
        txt(Octroi) = Format(0, "0.00")
        txt(TempReg) = Format(0, "0.00")
        txt(TransIns) = Format(0, "0.00")
        txt(MVT) = Format(0, "0.00")
        txt(Transportation) = Format(0, "0.00")
    End If
    Amt_Cal
    Exit Sub
NXT:
    txt(EngineNo) = ""
    txt(SrBookNo) = ""
    txt(Colours) = ""
    txt(Colours).Tag = ""
    txt(TelcoInvDate).Text = ""
    txt(TelcoInvNo).Text = ""
    txt(Taxable).Text = ""
    txt(NDP).Text = "0.00"
    txt(SaleRate).Text = "0.00"
    Amt_Cal
Else
    txt(EngineNo) = ""
    txt(SrBookNo) = ""
    txt(Colours) = ""
    txt(Colours).Tag = ""
    txt(TelcoInvDate).Text = ""
    txt(TelcoInvNo).Text = ""
    txt(Taxable).Text = ""
    txt(NDP).Text = "0.00"
    txt(SaleRate).Text = "0.00"
    Amt_Cal
End If
End Sub

Private Function FillRecords() As Boolean
Dim Rst As ADODB.Recordset
Dim RsBooking  As ADODB.Recordset
    Set RsBooking = New Recordset
    RsBooking.CursorLocation = adUseClient
    RsBooking.Open "SELECT Veh_Order.Inv_DocId,Veh_Order.PartyCode, " & _
    "Veh_Order.GOVT_YN, Veh_Order.MODEL, Veh_Order.Chassis, Veh_Order.Srv_BookNo, Veh_Order.RATE, Veh_Order.Fund_Source, Veh_Order.FB_CODE, Veh_Order.FIN_AcCode, Veh_Order.FIN_AcCode, Veh_Order.Colour_Code, Veh_Order.FIN_AMT FROM Veh_Order where Ord_No = " & Val(txt(BookNo)) & " and left(OrdDocId,1)= '" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    If RsBooking.RecordCount = 0 Then
        MsgBox "Booking No Not Exist", vbInformation, "Booking Not Found"
        txt(Party).Text = ""
        txt(Party).Tag = ""
        txt(Add1).Text = ""
        txt(Add2).Text = ""
        txt(Add3).Text = ""
        txt(City).Text = ""
        txt(Model).Text = ""
        txt(Govt_YN).Text = ""
        txt(ChassisNo).Text = ""
        txt(Colours).Tag = ""
        txt(Colours).Text = ""
        txt(SrBookNo).Text = ""
        txt(NDP).Text = ""
        txt(FundSource).Text = ""
        txt(FB_Code).Text = ""
        txt(FB_Code).Tag = ""
        FinAcCode = ""
        txt(FinAmt).Text = ""
        txt(BookNo).SetFocus
        FillRecords = False
        Exit Function
    Else
        If RsBooking!Inv_docid <> Null Or RsBooking!Inv_docid <> "" Then
        MsgBox "Invoice Exist Against Booking No", vbInformation, "Validation Check"
        txt(Party).Text = ""
        txt(Party).Tag = ""
        txt(Add1).Text = ""
        txt(Add2).Text = ""
        txt(Add3).Text = ""
        txt(City).Text = ""
        txt(Model).Text = ""
        txt(Govt_YN).Text = ""
        txt(ChassisNo).Text = ""
        txt(Colours).Tag = ""
        txt(Colours).Text = ""
        txt(SrBookNo).Text = ""
        txt(NDP).Text = ""
        txt(FundSource).Text = ""
        txt(FB_Code).Text = ""
        txt(FB_Code).Tag = ""
        FinAcCode = ""
        txt(FinAmt).Text = ""
        txt(BookNo).SetFocus
        FillRecords = False
        Exit Function
        End If
        Dim RstTemp As ADODB.Recordset
        Set RstTemp = GCn.Execute("Select sum(iif(DrCr ='C',Amount,Amount*-1)) as AdvAmt from Rect where mid(Ord_DocId,14,8) = " & Val(txt(BookNo)) & " and left(Ord_SiteCode,1)= '" & txt(SiteCode).Tag & "'")
        If RstTemp.RecordCount > 0 Then
            txt(AdvAmt) = Format(IIf(IsNull(RstTemp!AdvAmt), 0, RstTemp!AdvAmt), "0.00")
        End If
        txt(Party).Tag = IIf(IsNull(RsBooking!PartyCode), "", RsBooking!PartyCode)
        If txt(Party).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select name,add1,add2,add3,CityCode from SubGroup where Subcode = '" & txt(Party).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
            txt(Party).Text = Rst!Name
            txt(Add1).Text = IIf(IsNull(Rst!Add1), "", Rst!Add1)
            txt(Add2).Text = IIf(IsNull(Rst!Add2), "", Rst!Add2)
            txt(Add3).Text = IIf(IsNull(Rst!Add3), "", Rst!Add3)
            txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
            If txt(City).Tag <> "" Then
                txt(City).Text = GCn.Execute("select cityname from city where citycode = '" & txt(City).Tag & "'").Fields(0).Value
            End If
            txt(RTO).Text = txt(City).Text
        End If
        txt(Model).Text = RsBooking!Model
        txt(Govt_YN).Text = IIf(IsNull(RsBooking!Govt_YN), "", RsBooking!Govt_YN)
        txt(Colours).Tag = IIf(IsNull(RsBooking!Colour_Code), "", RsBooking!Colour_Code)
        If txt(Colours).Tag <> "" Then
            txt(Colours).Text = GCn.Execute("select col_desc from colmast where col_code = '" & txt(Colours).Tag & "'").Fields(0).Value
        End If
        Select Case RsBooking!Fund_Source
            Case 0 '0 Hypothication ,1 Hire purchase ,2 Own Fund,3 Lease
                txt(FundSource).Text = "Hypothication"
            Case 1
                txt(FundSource).Text = "Hire Purchase"
            Case 3
                txt(FundSource).Text = "Lease"
            Case Else
                txt(FundSource).Text = "Own Fund"
        End Select
            
        txt(FB_Code).Tag = IIf(IsNull(RsBooking!FB_Code), "", RsBooking!FB_Code)
        If txt(FB_Code).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select fincode as code,finname as name,AcCode from ContractFinance where fincatg = 0 and  fincode = '" & txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
            txt(FB_Code).Text = Rst!Name
            FinAcCode = IIf(IsNull(Rst!AcCode), "", Rst!AcCode)
        Else
            txt(FB_Code).Text = ""
            FinAcCode = ""
        End If
        txt(FinAmt).Text = IIf(IsNull(RsBooking!FIN_AMT), "", RsBooking!FIN_AMT)
    End If
    
FillRecords = True
Set Rst = Nothing
End Function
'************************ PRINTING CODE ******************


Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case DocType
        ListArray = Array("Sale Bill", "Sale Certificate", "Form22", "Form22A", "Declaration")
        Set mListItem = ListView_Items(ListView, txtPrint, Index, ListArray, 5)
    Case FromVno, ToVno
            RsVno.Close
            RsVno.Open "Select Inv_No as code from Veh_Order where right(veh_order.Inv_SiteCode,1)='" & txtPrint(SiteCode1).Tag & "' and  veh_order.inv_VType='V_SB'", GCn, adOpenDynamic, adLockOptimistic
            Set DGVno.DataSource = RsVno
            If txtPrint(Index).Text <> RsVno!Code Then
                RsVno.MoveFirst
                RsVno.FIND "code ='" & txtPrint(Index).Text & "'"
            End If
            If Index = ToVno Then DGVno.Tag = "1" Else DGVno.Tag = "2"
    Case SiteCode1
        If rsSite.RecordCount = 0 Or (rsSite.EOF = True Or rsSite.BOF = True) Then Exit Sub
        If txtPrint(Index).Text = "" Then
            rsSite.MoveFirst
            rsSite.FIND "code ='" & PubSiteCode & "'"
            txtPrint(Index).Tag = rsSite!Code
            txtPrint(Index).Text = rsSite!Name
        Else
            If txtPrint(Index).Text <> rsSite!Name Then
                rsSite.MoveFirst
                rsSite.FIND "name ='" & txtPrint(Index).Text & "'"
            End If
        End If
    
End Select
End Sub

Private Sub TxtPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case DocType
        ListView_KeyDown FrmList, ListView, txtPrint, Index, KeyCode, Shift, FrmPrn.left + txtPrint(Index).left, (FrmPrn.top + txtPrint(Index).top + txtPrint(Index).Height), txtPrint(Index).width, 1200
    Case FromVno, ToVno
        DGridTxtKeyDown DGVno, txtPrint, Index, RsVno, KeyCode, False, 0
    Case SiteCode1
        DGridTxtKeyDown DGSite, txtPrint, Index, rsSite, KeyCode, False, 1
End Select
If FrmList.Visible = False And DGSite.Visible = False And DGVno.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> SiteCode1 Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    
    Case TempYN, Tempadd
        If UCase(Chr(KeyAscii)) = "Y" Then
            txtPrint(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txtPrint(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txtPrint(Index) = ""
        End If
        KeyAscii = 0
        If txtPrint(Index).Text = "Yes" Then FldEnabled1 True Else FldEnabled1 False
    Case Form22A, WtPrn
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
    Case FromVno, ToVno
        If DGVno.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVno, KeyAscii, "Code"
    Case SiteCode1
        If DGSite.Visible = True Then DGridTxtKeyPress txtPrint, Index, rsSite, KeyAscii, "Name"
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
            If txtPrint(Index).Text <> "" Then txtPrint(Index).Text = ListView.SelectedItem.Text
            If txtPrint(Index).Text = "Sale Certificate" Then FldEnabled True Else FldEnabled False
    Case PrnDate
        txtPrint(Index).Text = RetDate(txtPrint(Index))
    Case ToVno, FromVno
        If RsVno.RecordCount = 0 Or (RsVno.EOF = True Or RsVno.BOF = True) Or txtPrint(Index).Text = "" Then
            txtPrint(Index).Text = ""
        Else
            txtPrint(Index).Text = RsVno!Code
        End If
    Case SiteCode1
        If IsValid(txtPrint(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If rsSite.RecordCount = 0 Or (rsSite.EOF = True Or rsSite.BOF = True) Or txtPrint(Index).Text = "" Then
            txtPrint(Index).Text = ""
            txtPrint(Index).Tag = ""
        Else
            txtPrint(Index).Text = rsSite!Name
            txtPrint(Index).Tag = rsSite!Code
        End If

End Select
End Sub


Private Sub FldEnabled(Enb As Boolean)
    txtPrint(Form22A).Enabled = Enb
    txtPrint(NewRTOName).Enabled = Enb
    txtPrint(PrnDate).Enabled = Enb
    txtPrint(TempYN).Enabled = Enb
    txtPrint(Seet).Enabled = Enb
    txtPrint(Body).Enabled = Enb
    txtPrint(Narr).Enabled = Enb
    txtPrint(WtPrn).Enabled = Enb
    txtPrint(TempRTO).Enabled = Enb
    If Enb = False Then
        txtPrint(Form22A).Text = ""
        txtPrint(NewRTOName).Text = ""
        txtPrint(PrnDate).Text = ""
        txtPrint(TempYN).Text = ""
        txtPrint(Seet).Text = ""
        txtPrint(Body).Text = ""
        txtPrint(Narr).Text = ""
        txtPrint(WtPrn).Text = ""
        txtPrint(TempRTO).Text = ""
    End If
End Sub
Private Sub FldEnabled1(Enb As Boolean)
    txtPrint(Seet).Enabled = Enb
    txtPrint(Body).Enabled = Enb
    txtPrint(Narr).Enabled = Enb
    txtPrint(WtPrn).Enabled = Enb
    txtPrint(TempRTO).Enabled = Enb
    If Enb = False Then
        txtPrint(Seet).Text = ""
        txtPrint(Body).Text = ""
        txtPrint(Narr).Text = ""
        txtPrint(WtPrn).Text = ""
        txtPrint(TempRTO).Text = ""
    End If
End Sub

Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    FrmPrn.Visible = False
    If Index <> PSetUp And TopCtrl1.TopText2.Caption <> "Browse" Then
        If TopCtrl1.TopText2.Caption = "Add" Then TopCtrl1_eAdd: Exit Sub
        Disp_Text SETS("INI", Me, Master)
        Call MoveRec
    End If
End If
End Sub
Private Sub Cmdprint_Click(Index As Integer)
'On Error GoTo ERRORHANDLER
If IsValid(txtPrint(DocType), "Print Document") = False Then Exit Sub
Select Case Index
    Case PScreen, PWindows
        If txtPrint(DocType).Text = "Sale Certificate" Then
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert", "VehSaleCert")
            Call WindowsPrint(Index)
            If txtPrint(Form22A).Text = "Yes" Then
                mRepName = IIf(OptPlain.Value = True, "VehSaleCert22A", "VehSaleCert22A")
                Call WindowsPrint(Index)
            End If
        ElseIf txtPrint(DocType).Text = "Form22A" Then
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert22A", "VehSaleCert22A")
            Call WindowsPrint(Index)
        ElseIf txtPrint(DocType).Text = "Form22" Then
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert22A", "VehSaleCert22A")
            Call WindowsPrint(Index)
        ElseIf txtPrint(DocType).Text = "Sale Bill" Then
            mRepName = IIf(OptPlain.Value = True, "Vehsale", "Vehsale")
            Call WindowsPrint(Index)
        End If
        
        FrmPrn.Visible = False
    Case PDos
        If txtPrint(DocType).Text = "Sale Certificate" Then
            Call SpeedPrint1
        ElseIf txtPrint(DocType).Text = "Form22A" Then
            Call SpeedPrint22A
        ElseIf txtPrint(DocType).Text = "Form22" Then
            Call SpeedPrint22
        ElseIf txtPrint(DocType).Text = "Sale Bill" Then
            Call SpeedPrint
        Else
            Call SpeedPrintDeclar
        End If
        FrmPrn.Visible = False
    Case PSetUp
        If txtPrint(DocType).Text = "Sale Certificate" Then
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert", "VehSaleCert")
        Else
            mRepName = IIf(OptPlain.Value = True, "Vehsale", "Vehsale")
        End If
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.Caption <> "Browse" Then
    If TopCtrl1.TopText2.Caption = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.Caption

End Sub

Private Sub WindowsPrint(Index As Integer)
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset, RstSub2 As ADODB.Recordset, mQry As String
Dim i As Integer, Cnt As Integer, Foot1 As String, Foot2 As String, Foot3 As String, Foot4 As String
Dim Foot5 As String, Foot6 As String, Foot7 As String, Foot8 As String, Foot9 As String
Dim Rst1 As ADODB.Recordset, j As Integer, Footer As String
Dim Rst2 As ADODB.Recordset
'On Error GoTo ERRORHANDLER
    
If txtPrint(DocType).Text = "Sale Certificate" Then
    mQry = "SELECT veh_order.INTD_USE, Veh_Purch1.Tot_Amount,veh_Stock.vrate, veh_Stock.Mfg_Month , " & _
        "veh_Stock.Mfg_Yr, veh_order.CertiPrn_YN, veh_order.TCertiPrn_YN, " & _
        "veh_order.Inv_UName, veh_order.Inv_UEntDt, Subgroup.FPrefix, Subgroup.FName, " & _
        "city_1.cityname as TCity,   Subgroup.TAdd1, Subgroup.TAdd2, Subgroup.TAdd3," & _
        "Subgroup.TPIN, veh_order.Inv_DocId, veh_order.Fund_Source, veh_order.P_AMOUNT, " & _
        "veh_order.DelCh_DT, veh_order.Inv_Date, veh_order.Inv_SiteCode, veh_order.RTO," & _
        "Model_Grp.ModelGrp_Name, city.CityName, Fincity.cityname as FinCity," & _
        "Subgroup.Name, ColMast.Col_Desc, Model.MODEL, Model.Site_Code, Model.Chas_Type," & _
        "Model.Vehicle_Type, Model.Model_Type, Model.Model_Ind, Model.Sales_Desc," & _
        "Model.Model_Desc, Model.Model_Desc1, Model.Model_Desc2, Model.Grp_Code," & _
        "Model.Cat_Code, Model.Div_Code, Model.Wheel_Catg, Model.Active_YN, " & _
        "Model.STAT_IND, Model.TYRES, Model.TYRE_F, Model.TYRE_M, Model.TYRE_R," & _
        "Model.TYRE_FS, Model.TYRE_MS, Model.TYRE_RS, Model.RIMS, Model.SEAT," & _
        "Model.RLW,Model.HORSEPOWER , Model.FRONT_A_WT, Model.REAR_A_WT," & _
        "Model.UNLADEN_WT, Model.GROSS_WT, Model.WHEELBASE, Model.CYLINDER, Model.FUEL," & _
        "Model.TRADE_NO, Model.Manufacturer, Model.Warr_KMS, Model.Warr_Mth, Model.U_Name," & _
        "Model.U_EntDt, Model.U_AE, Model.Trf_Date, veh_Stock.ChassisNo , veh_Stock.EngineNo , Finbank.FinBankName,ContractFinance.FinName, ContractFinance.Add1 as FAdd1," & _
        "ContractFinance.add2 as Fadd2 , ContractFinance.PinCode as FPin , Subgroup.Add1, Subgroup.Add2, Subgroup.Add3, Subgroup.PIN " & _
        "FROM ((((((((((veh_order LEFT JOIN veh_Stock ON veh_order.Inv_DocId = veh_Stock.Sal_DocId) LEFT JOIN ColMast ON veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN Veh_Purch1 ON veh_Stock.Pur_DocId = Veh_Purch1.DocID) " & _
        "LEFT JOIN ContractFinance ON veh_order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Subgroup ON veh_order.PartyCode = Subgroup.SubCode) LEFT JOIN Model ON veh_order.MODEL = Model.MODEL) " & _
        "LEFT JOIN Model_Grp ON Model.Grp_Code = Model_Grp.ModelGrp_Code) LEFT JOIN city AS fincity ON ContractFinance.City = fincity.CityCode) LEFT JOIN Finbank ON ContractFinance.FinBankCode = Finbank.FinBankCode) " & _
        "LEFT JOIN city ON Subgroup.CityCode = city.CityCode) LEFT JOIN City AS city_1 ON Subgroup.TCityCode = city_1.CityCode " & _
        "where veh_order.Inv_DocId = '" & Master!SearchCode & "' and  Veh_Order.DelCh_docid <> Null"
        
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    Set Rst1 = New ADODB.Recordset
    Rst1.CursorLocation = adUseClient
    Rst1.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    For i = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
            Case UCase("TempAdd")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Tempadd) & "'"
            Case UCase("SubTitle")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecSpeciality & "'"
            Case UCase("LST")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecGram & "'"
            Case UCase("Form22A")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Form22A) & "'"
            Case UCase("NewRTOName")
                rpt.FormulaFields(i).Text = "'" & txtPrint(NewRTOName) & "'"
            Case UCase("PrnDate")
                rpt.FormulaFields(i).Text = "'" & txtPrint(PrnDate) & "'"
            Case UCase("TempYN")
                rpt.FormulaFields(i).Text = "'" & txtPrint(TempYN) & "'"
            Case UCase("Seet")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Seet) & "'"
            Case UCase("Body")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Body) & "'"
            Case UCase("Narr")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Narr) & "'"
            Case UCase("WtPrn")
                rpt.FormulaFields(i).Text = "'" & txtPrint(WtPrn) & "'"
            Case UCase("TempRto")
                rpt.FormulaFields(i).Text = "'" & txtPrint(TempRTO) & "'"
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Select Case Index
        Case PWindows
            For i = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
                    Case UCase("comp_name")
                        rpt.FormulaFields(i).Text = "'" & PubComp_Name & "'"
                    Case UCase("comp_add1")
                        rpt.FormulaFields(i).Text = "'" & PubComp_Add & "'"
                    Case UCase("comp_add2")
                        rpt.FormulaFields(i).Text = "'" & PubComp_Add2 & "'"
                    Case UCase("comp_city")
                        rpt.FormulaFields(i).Text = "'" & PubComp_City & "'"
                    Case UCase("Title")
                        rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                End Select
            Next
            rpt.PrintOut False
            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                If txtPrint(TempYN) = "Yes" Then
                    GCn.Execute "update veh_order set CertiPrn_YN = 1  where where veh_order.Inv_DocId = '" & Master!SearchCode & "' And Veh_Order.DelCh_docid <> Null"
                Else
                    GCn.Execute "update veh_order set TCertiPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "' And Veh_Order.DelCh_docid <> Null"
                End If
            End If
            Set Rst = Nothing
            Set Rst1 = Nothing
            Set rpt = Nothing
        Case PScreen 'screen
            Call Report_View(rpt, Me.Caption, , True)
            Set Rst = Nothing
            Set Rst1 = Nothing
    End Select
ElseIf txtPrint(DocType).Text = "Form22A" Or txtPrint(DocType).Text = "Form22" Then
    mQry = "SELECT veh_order.INTD_USE, Veh_Purch1.Tot_Amount,veh_Stock.vrate, veh_Stock.Mfg_Month , " & _
        "veh_Stock.Mfg_Yr, veh_order.CertiPrn_YN, veh_order.TCertiPrn_YN, " & _
        "veh_order.Inv_UName, veh_order.Inv_UEntDt, Subgroup.FPrefix, Subgroup.FName, " & _
        "city_1.cityname as TCity,   Subgroup.TAdd1, Subgroup.TAdd2, Subgroup.TAdd3," & _
        "Subgroup.TPIN, veh_order.Inv_DocId, veh_order.Fund_Source, veh_order.P_AMOUNT, " & _
        "veh_order.DelCh_DT, veh_order.Inv_Date, veh_order.Inv_SiteCode, veh_order.RTO," & _
        "Model_Grp.ModelGrp_Name, city.CityName, Fincity.cityname as FinCity," & _
        "Subgroup.Name, ColMast.Col_Desc, Model.MODEL, Model.Site_Code, Model.Chas_Type," & _
        "Model.Vehicle_Type, Model.Model_Type, Model.Model_Ind, Model.Sales_Desc," & _
        "Model.Model_Desc, Model.Model_Desc1, Model.Model_Desc2, Model.Grp_Code," & _
        "Model.Cat_Code, Model.Div_Code, Model.Wheel_Catg, Model.Active_YN, " & _
        "Model.STAT_IND, Model.TYRES, Model.TYRE_F, Model.TYRE_M, Model.TYRE_R," & _
        "Model.TYRE_FS, Model.TYRE_MS, Model.TYRE_RS, Model.RIMS, Model.SEAT," & _
        "Model.RLW,Model.HORSEPOWER , Model.FRONT_A_WT, Model.REAR_A_WT," & _
        "Model.UNLADEN_WT, Model.GROSS_WT, Model.WHEELBASE, Model.CYLINDER, Model.FUEL," & _
        "Model.TRADE_NO, Model.Manufacturer, Model.Warr_KMS, Model.Warr_Mth, Model.U_Name," & _
        "Model.U_EntDt, Model.U_AE, Model.Trf_Date, veh_Stock.ChassisNo , veh_Stock.EngineNo , Finbank.FinBankName,ContractFinance.FinName, ContractFinance.Add1 as FAdd1," & _
        "ContractFinance.add2 as Fadd2 , ContractFinance.PinCode as FPin , Subgroup.Add1, Subgroup.Add2, Subgroup.Add3, Subgroup.PIN " & _
        "FROM ((((((((((veh_order LEFT JOIN veh_Stock ON veh_order.Inv_DocId = veh_Stock.Sal_DocId) LEFT JOIN ColMast ON veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN Veh_Purch1 ON veh_Stock.Pur_DocId = Veh_Purch1.DocID) " & _
        "LEFT JOIN ContractFinance ON veh_order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Subgroup ON veh_order.PartyCode = Subgroup.SubCode) LEFT JOIN Model ON veh_order.MODEL = Model.MODEL) " & _
        "LEFT JOIN Model_Grp ON Model.Grp_Code = Model_Grp.ModelGrp_Code) LEFT JOIN city AS fincity ON ContractFinance.City = fincity.CityCode) LEFT JOIN Finbank ON ContractFinance.FinBankCode = Finbank.FinBankCode) " & _
        "LEFT JOIN city ON Subgroup.CityCode = city.CityCode) LEFT JOIN City AS city_1 ON Subgroup.TCityCode = city_1.CityCode " & _
        "where veh_order.Inv_DocId = '" & Master!SearchCode & "' and  Veh_Order.DelCh_docid <> Null"
        
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    Set Rst1 = New ADODB.Recordset
    Rst1.CursorLocation = adUseClient
    Rst1.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    For i = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
            Case UCase("FormType")
                rpt.FormulaFields(i).Text = "'" & IIf(txtPrint(DocType).Text = "Form22", "22", "22A") & "'"
            Case UCase("Form22A")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Form22A) & "'"
            Case UCase("NewRTOName")
                rpt.FormulaFields(i).Text = "'" & txtPrint(NewRTOName) & "'"
            Case UCase("PrnDate")
                rpt.FormulaFields(i).Text = "'" & txtPrint(PrnDate) & "'"
            Case UCase("TempYN")
                rpt.FormulaFields(i).Text = "'" & txtPrint(TempYN) & "'"
            Case UCase("Seet")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Seet) & "'"
            Case UCase("Body")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Body) & "'"
            Case UCase("Narr")
                rpt.FormulaFields(i).Text = "'" & txtPrint(Narr) & "'"
            Case UCase("WtPrn")
                rpt.FormulaFields(i).Text = "'" & txtPrint(WtPrn) & "'"
            Case UCase("TempRto")
                rpt.FormulaFields(i).Text = "'" & txtPrint(TempRTO) & "'"
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Select Case Index
        Case PWindows
            
            rpt.PrintOut False
            Set Rst = Nothing
            Set Rst1 = Nothing
            Set rpt = Nothing
        Case PScreen 'screen
            Call Report_View(rpt, Me.Caption, , True)
            Set Rst = Nothing
            Set Rst1 = Nothing
    End Select
    
Else
    mQry = "SELECT veh_order.*,Veh_Purch1.gate,City_1.CityName as fincity, ContractFinance.Add1 as finadd1, ContractFinance.Add2 as finadd2,finbank.finbankname,site.site_desc,ContractFinance.finname,   " & _
        " Veh_Stock.Pur_DocId, Veh_Stock.Sal_DocId, Veh_Stock.ChassisNo, Veh_Stock.EngineNo, Veh_Stock.PBILL_NO, Veh_Stock.PBILL_DATE, " & _
        " Model.Model_Desc,Model.Model_Desc1, " & _
        "ColMast.Col_Desc, SubGroup.Name, SubGroup.Add1, " & _
        "SubGroup.Add2,SubGroup.Add3,SubGroup.FPrefix,SubGroup.FName,City.CityName FROM ((((((((((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN ColMast ON " & _
        "Veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode) " & _
        "LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Site ON right(Veh_Order.Inv_SiteCode,1) = Site.Site_Code) LEFT JOIN FinBank ON ContractFinance.FinBankCode = FinBank.FinBankCode) LEFT JOIN City AS City_1 ON ContractFinance.City = City_1.CityCode) " & _
        "LEFT JOIN Veh_Purch1 ON Veh_Stock.Pur_DocId = Veh_Purch1.DocID  " & _
        "where veh_order.Inv_DocId = '" & Master!SearchCode & "' "
        
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
    
    'Recordset is made for subreport1
'    mQRY = "SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
'        "FROM (Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN (Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code) ON Veh_Stock.Sal_DocId = Veh_Purch2.DocID " & _
'        "where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
        
    mQry = "SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
    "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    "where Veh_Purch2.DocId = '" & Master!SearchCode & "'"
    
    Set RstSub1 = New Recordset
    RstSub1.CursorLocation = adUseClient
    RstSub1.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

   'Recordset is made for subreport2
   
    mQry = "SELECT Veh_Purch2.Trn_Type, Veh_Purch2.DocID, Veh_Purch2.QTY, Veh_Purch2.RATE, Veh_AMDModel.Prod_Name " & _
    "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code where veh_stock.Chassisno = '" & txt(ChassisNo) & "'"
        
    Set RstSub2 = New Recordset
    RstSub2.CursorLocation = adUseClient
    RstSub2.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    CreateFieldDefFile RstSub1, PubRepoPath + "\" & mRepName & "1.ttx", True
    CreateFieldDefFile RstSub2, PubRepoPath + "\" & mRepName & "2.ttx", True
    
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    Set Rst1 = New ADODB.Recordset
    Rst1.CursorLocation = adUseClient
    Rst1.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    j = 1
    Cnt = 1
    For i = 1 To Len(Footer)
        If Mid(Footer, i, 1) = vbLf Then
            Select Case Cnt
            Case 1
                Foot1 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 2
                Foot2 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 3
                Foot3 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 4
                Foot4 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 5
                Foot5 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 6
                Foot6 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 7
                Foot7 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 8
                Foot8 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            Case 9
                Foot9 = left(RTrim(Mid(Footer, j, i - j - 1)), 130)
            End Select
            Cnt = Cnt + 1
            j = i + 1
            
        End If
    Next
    
    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from Syctrl", GCn, adOpenDynamic, adLockOptimistic
    For i = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
            Case UCase("SubTitle")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecSpeciality & "'"
            Case UCase("AmtPrefix")
                rpt.FormulaFields(i).Text = "'" & PubAmountPrefix & "'"
            Case UCase("TelcoInvYN")
                rpt.FormulaFields(i).Text = "" & Rst2!SupInvOnVehSaleInv & ""
            Case UCase("TaxDetYN")
                rpt.FormulaFields(i).Text = "" & Rst2!TaxDetOnVehInv & ""
            Case UCase("InvPrefix")
                rpt.FormulaFields(i).Text = "'" & Rst2!VehSaleInv_Prefix & "'"
            Case UCase("LST")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(i).Text = "'" & Rst1!V_SecGram & "'"
            Case UCase("SubRep1")
                rpt.FormulaFields(i).Text = "" & IIf(RstSub1.RecordCount = 0, 0, 1) & ""
            Case UCase("SubRep2")
                rpt.FormulaFields(i).Text = "" & IIf(RstSub2.RecordCount = 0, 0, 1) & ""
           Case UCase("AddDet")
                rpt.FormulaFields(i).Text = "" & IIf(txt(ADType) = "No Detail", 0, IIf(txt(ADType) = "Name/Qty", 1, 2)) & ""
           Case UCase("Foot1")
                rpt.FormulaFields(i).Text = "'" & Foot1 & "'"
            Case UCase("Foot2")
                rpt.FormulaFields(i).Text = "'" & Foot2 & "'"
            Case UCase("Foot3")
                rpt.FormulaFields(i).Text = "'" & Foot3 & "'"
            Case UCase("Foot4")
                rpt.FormulaFields(i).Text = "'" & Foot4 & "'"
            Case UCase("Foot5")
                rpt.FormulaFields(i).Text = "'" & Foot5 & "'"
            Case UCase("Foot6")
                rpt.FormulaFields(i).Text = "'" & Foot6 & "'"
            Case UCase("Foot7")
                rpt.FormulaFields(i).Text = "'" & Foot7 & "'"
            Case UCase("Foot8")
                rpt.FormulaFields(i).Text = "'" & Foot8 & "'"
            Case UCase("Foot9")
                rpt.FormulaFields(i).Text = "'" & Foot9 & "'"
        End Select
    Next
    For i = 1 To rpt.OpenSubreport("SubRep2").FormulaFields.Count
        Select Case UCase(rpt.OpenSubreport("SubRep2").FormulaFields(i).FormulaFieldName)
            Case UCase("AddDet")
            rpt.OpenSubreport("SubRep2").FormulaFields(i).Text = "" & IIf(txt(ADType) = "No Detail", 0, IIf(txt(ADType) = "Name/Qty", 1, 2)) & ""
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstSub1
    rpt.OpenSubreport("SubRep2").Database.SetDataSource RstSub2
    rpt.ReadRecords
    Select Case Index
        Case PWindows  'Printer
            For i = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
                    Case UCase("comp_name")
                        rpt.FormulaFields(i).Text = "'" & PubComp_Name & "'"
                    Case UCase("comp_add1")
                        rpt.FormulaFields(i).Text = "'" & PubComp_Add & "'"
                    Case UCase("comp_add2")
                        rpt.FormulaFields(i).Text = "'" & PubComp_Add2 & "'"
                    Case UCase("comp_city")
                        rpt.FormulaFields(i).Text = "'" & PubComp_City & "'"
                    Case UCase("Title")
                        rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                End Select
            Next
            rpt.PrintOut False
            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
            End If
            Set Rst = Nothing
            Set Rst1 = Nothing
            Set rpt = Nothing
        Case PScreen  'screen
            Call Report_View(rpt, Me.Caption, , True)
            Set Rst = Nothing
            Set Rst1 = Nothing
    End Select
End If
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.Caption = rpt.PrinterName
End Sub
Private Sub SpeedPrint22A()
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
    Dim i As Integer, j As Integer, mQry As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstCert As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer1 As String, Footer2 As String, Footer3 As String, Footer4 As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim Fob As New FileSystemObject
    Dim mJuriCity As String
    Dim Cnt As Byte, mAmt As Double, PrnStr As String, PrnStr1 As String
    Dim Left1 As String, Left2 As String, Left3 As String
    Dim Left4 As String, Left5 As String, Left6 As String, Left7 As String
    Dim Right1 As String, Right2 As String, Right3 As String
    Dim Right4 As String, Right5 As String, Right6 As String, Right7 As String
    Dim NetAmt As Double

    Set RstCert = GCn.Execute("SELECT veh_order.INTD_USE, Veh_Purch1.Tot_Amount,veh_Stock.vrate, veh_Stock.Mfg_Month , " & _
        "veh_Stock.Mfg_Yr, veh_order.CertiPrn_YN, veh_order.TCertiPrn_YN, " & _
        "veh_order.Inv_UName, veh_order.Inv_UEntDt, Subgroup.FPrefix, Subgroup.FName, " & _
        "city_1.cityname as TCity,   Subgroup.TAdd1, Subgroup.TAdd2, Subgroup.TAdd3," & _
        "Subgroup.TPIN, veh_order.Inv_DocId, veh_order.Fund_Source, veh_order.P_AMOUNT, " & _
        "veh_order.DelCh_DT, veh_order.Inv_Date, veh_order.Inv_SiteCode, veh_order.RTO," & _
        "Model_Grp.ModelGrp_Name, city.CityName, Fincity.cityname as FinCity," & _
        "Subgroup.Name, ColMast.Col_Desc, Model.MODEL, Model.Site_Code, Model.Chas_Type," & _
        "Model.Vehicle_Type, Model.Model_Type, Model.Model_Ind, Model.Sales_Desc," & _
        "Model.Model_Desc, Model.Model_Desc1, Model.Model_Desc2, Model.Grp_Code," & _
        "Model.Cat_Code, Model.Div_Code, Model.Wheel_Catg, Model.Active_YN, " & _
        "Model.STAT_IND, Model.TYRES, Model.TYRE_F, Model.TYRE_M, Model.TYRE_R," & _
        "Model.TYRE_FS, Model.TYRE_MS, Model.TYRE_RS, Model.RIMS, Model.SEAT," & _
        "Model.RLW,Model.HORSEPOWER , Model.FRONT_A_WT, Model.REAR_A_WT," & _
        "Model.UNLADEN_WT, Model.GROSS_WT, Model.WHEELBASE, Model.CYLINDER, Model.FUEL," & _
        "Model.TRADE_NO, Model.Manufacturer, Model.Warr_KMS, Model.Warr_Mth, Model.U_Name," & _
        "Model.U_EntDt, Model.U_AE, Model.Trf_Date, veh_Stock.ChassisNo , veh_Stock.EngineNo , Finbank.FinBankName,ContractFinance.FinName, ContractFinance.Add1 as FAdd1," & _
        "ContractFinance.add2 as Fadd2 , ContractFinance.PinCode as FPin , Subgroup.Add1, Subgroup.Add2, Subgroup.Add3, Subgroup.PIN " & _
        "FROM ((((((((((veh_order LEFT JOIN veh_Stock ON veh_order.Inv_DocId = veh_Stock.Sal_DocId) LEFT JOIN ColMast ON veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN Veh_Purch1 ON veh_Stock.Pur_DocId = Veh_Purch1.DocID) " & _
        "LEFT JOIN ContractFinance ON veh_order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Subgroup ON veh_order.PartyCode = Subgroup.SubCode) LEFT JOIN Model ON veh_order.MODEL = Model.MODEL) " & _
        "LEFT JOIN Model_Grp ON Model.Grp_Code = Model_Grp.ModelGrp_Code) LEFT JOIN city AS fincity ON ContractFinance.City = fincity.CityCode) LEFT JOIN Finbank ON ContractFinance.FinBankCode = Finbank.FinBankCode) " & _
        "LEFT JOIN city ON Subgroup.CityCode = city.CityCode) LEFT JOIN City AS city_1 ON Subgroup.TCityCode = city_1.CityCode " & _
        "where veh_order.Inv_DocId = '" & Master!SearchCode & "' and Veh_Order.DelCh_docid <> Null")

      
    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
    If Fob.FileExists("C:\RepPrint.Txt") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If Fob.FileExists("C:\RepPrint.Bat") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
    
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
 
    PageLength = PubPageLength
    PageWidth = 34
    mHeader = 0   'Ideal 17
    
        mHeader = 0
        mFooter = 2
        
        Print #1, mChr18 & PRN_TIT(RstCert!Manufacturer, "A", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT(IIf(RstCert!CertiPrn_YN = 1, "DUPLICATE", ""), "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("F O R M - 22-A", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[See Rule 47 (g),124,126A AND 127]", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("Part 1", "B", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("(Issued By The Manufacturer)", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, "Certified that Tata " & RstCert!Model_Desc & "(Brand name of the vehicle)"
        mHeader = mHeader + 1
        Print #1, "bearing Chassis Number " & RstCert!ChassisNo & "  and Engine Number " & RstCert!EngineNo
        mHeader = mHeader + 1
        Print #1, "complies with the provisions of  the  Motor Vehicles Act, 1988 and the rule made thereunder."
        mHeader = mHeader + 1
        Print #1, PSTR("Signature of the manufacturer", PageWidth, , AlignRight)
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("For " & RstCert!Manufacturer, PageWidth, , AlignRight) & mEmph1
        mHeader = mHeader + 1
        
        Do Until mHeader >= PageLength - mFooter
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        Print #1, Replace(Space(PageWidth), " ", "-")
        
        Print #1, mChr17 & RstCert!Inv_UName & " " & str(RstCert!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(RstCert!Inv_UName & " " & str(RstCert!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If Fob.FolderExists("c:\WinNt") Then
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
    Else
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Close #1
    
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub




Private Sub SpeedPrint22()
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
    Dim i As Integer, j As Integer, mQry As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstCert As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer1 As String, Footer2 As String, Footer3 As String, Footer4 As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim Fob As New FileSystemObject
    Dim mJuriCity As String
    Dim Cnt As Byte, mAmt As Double, PrnStr As String, PrnStr1 As String
    Dim Left1 As String, Left2 As String, Left3 As String
    Dim Left4 As String, Left5 As String, Left6 As String, Left7 As String
    Dim Right1 As String, Right2 As String, Right3 As String
    Dim Right4 As String, Right5 As String, Right6 As String, Right7 As String
    Dim NetAmt As Double

    Set RstCert = GCn.Execute("SELECT veh_order.INTD_USE, Veh_Purch1.Tot_Amount,veh_Stock.vrate, veh_Stock.Mfg_Month , " & _
        "veh_Stock.Mfg_Yr, veh_order.CertiPrn_YN, veh_order.TCertiPrn_YN, " & _
        "veh_order.Inv_UName, veh_order.Inv_UEntDt, Subgroup.FPrefix, Subgroup.FName, " & _
        "city_1.cityname as TCity,   Subgroup.TAdd1, Subgroup.TAdd2, Subgroup.TAdd3," & _
        "Subgroup.TPIN, veh_order.Inv_DocId, veh_order.Fund_Source, veh_order.P_AMOUNT, " & _
        "veh_order.DelCh_DT, veh_order.Inv_Date, veh_order.Inv_SiteCode, veh_order.RTO," & _
        "Model_Grp.ModelGrp_Name, city.CityName, Fincity.cityname as FinCity," & _
        "Subgroup.Name, ColMast.Col_Desc, Model.MODEL, Model.Site_Code, Model.Chas_Type," & _
        "Model.Vehicle_Type, Model.Model_Type, Model.Model_Ind, Model.Sales_Desc," & _
        "Model.Model_Desc, Model.Model_Desc1, Model.Model_Desc2, Model.Grp_Code," & _
        "Model.Cat_Code, Model.Div_Code, Model.Wheel_Catg, Model.Active_YN, " & _
        "Model.STAT_IND, Model.TYRES, Model.TYRE_F, Model.TYRE_M, Model.TYRE_R," & _
        "Model.TYRE_FS, Model.TYRE_MS, Model.TYRE_RS, Model.RIMS, Model.SEAT," & _
        "Model.RLW,Model.HORSEPOWER , Model.FRONT_A_WT, Model.REAR_A_WT," & _
        "Model.UNLADEN_WT, Model.GROSS_WT, Model.WHEELBASE, Model.CYLINDER, Model.FUEL," & _
        "Model.TRADE_NO, Model.Manufacturer, Model.Warr_KMS, Model.Warr_Mth, Model.U_Name," & _
        "Model.U_EntDt, Model.U_AE, Model.Trf_Date, veh_Stock.ChassisNo , veh_Stock.EngineNo , Finbank.FinBankName,ContractFinance.FinName, ContractFinance.Add1 as FAdd1," & _
        "ContractFinance.add2 as Fadd2 , ContractFinance.PinCode as FPin , Subgroup.Add1, Subgroup.Add2, Subgroup.Add3, Subgroup.PIN " & _
        "FROM ((((((((((veh_order LEFT JOIN veh_Stock ON veh_order.Inv_DocId = veh_Stock.Sal_DocId) LEFT JOIN ColMast ON veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN Veh_Purch1 ON veh_Stock.Pur_DocId = Veh_Purch1.DocID) " & _
        "LEFT JOIN ContractFinance ON veh_order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Subgroup ON veh_order.PartyCode = Subgroup.SubCode) LEFT JOIN Model ON veh_order.MODEL = Model.MODEL) " & _
        "LEFT JOIN Model_Grp ON Model.Grp_Code = Model_Grp.ModelGrp_Code) LEFT JOIN city AS fincity ON ContractFinance.City = fincity.CityCode) LEFT JOIN Finbank ON ContractFinance.FinBankCode = Finbank.FinBankCode) " & _
        "LEFT JOIN city ON Subgroup.CityCode = city.CityCode) LEFT JOIN City AS city_1 ON Subgroup.TCityCode = city_1.CityCode " & _
        "where veh_order.Inv_DocId = '" & Master!SearchCode & "' and Veh_Order.DelCh_docid <> Null")

      
    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
    If Fob.FileExists("C:\RepPrint.Txt") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If Fob.FileExists("C:\RepPrint.Bat") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
    
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
 
    PageLength = PubPageLength
    PageWidth = 34
    mHeader = 0   'Ideal 17
    
        mHeader = 0
        mFooter = 2
        
        Print #1, mChr18 & PRN_TIT(RstCert!Manufacturer, "A", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT(IIf(RstCert!CertiPrn_YN = 1, "DUPLICATE", ""), "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("F O R M - 22", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[See Rule 47 (g),115(6),124 & 127]", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("Part 1", "B", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("(Issued By The Manufacturer)", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, "Certified that Tata " & RstCert!Model_Desc & "(Brand name of the vehicle)"
        mHeader = mHeader + 1
        Print #1, "bearing Chassis Number " & RstCert!ChassisNo & "  and Engine Number " & RstCert!EngineNo
        mHeader = mHeader + 1
        Print #1, "complies with the provisions of  the  Motor Vehicles Act, 1988 and the rule made thereunder."
        mHeader = mHeader + 1
        Print #1, PSTR("Signature of the manufacturer", PageWidth, , AlignRight)
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("For " & RstCert!Manufacturer, PageWidth, , AlignRight) & mEmph1
        mHeader = mHeader + 1
        
        Do Until mHeader >= PageLength - mFooter
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        Print #1, Replace(Space(PageWidth), " ", "-")
        
        Print #1, mChr17 & RstCert!Inv_UName & " " & str(RstCert!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(RstCert!Inv_UName & " " & str(RstCert!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If Fob.FolderExists("c:\WinNt") Then
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
    Else
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Close #1
    
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub



Private Sub SpeedPrint1()
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
    Dim i As Integer, j As Integer, mQry As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstCert As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer1 As String, Footer2 As String, Footer3 As String, Footer4 As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim Fob As New FileSystemObject
    Dim mJuriCity As String
    Dim Cnt As Byte, mAmt As Double, PrnStr As String, PrnStr1 As String
    Dim Left1 As String, Left2 As String, Left3 As String
    Dim Left4 As String, Left5 As String, Left6 As String, Left7 As String
    Dim Right1 As String, Right2 As String, Right3 As String
    Dim Right4 As String, Right5 As String, Right6 As String, Right7 As String
    Dim NetAmt As Double

    Set RstCert = GCn.Execute("SELECT veh_order.INTD_USE, Veh_Purch1.Tot_Amount,veh_Stock.vrate, veh_Stock.Mfg_Month , " & _
        "veh_Stock.Mfg_Yr, veh_order.CertiPrn_YN, veh_order.TCertiPrn_YN, " & _
        "veh_order.Inv_UName, veh_order.Inv_UEntDt, Subgroup.FPrefix, Subgroup.FName, " & _
        "city_1.cityname as TCity,   Subgroup.TAdd1, Subgroup.TAdd2, Subgroup.TAdd3," & _
        "Subgroup.TPIN, veh_order.Inv_DocId, veh_order.Fund_Source, veh_order.P_AMOUNT, " & _
        "veh_order.DelCh_DT, veh_order.Inv_Date, veh_order.Inv_SiteCode, veh_order.RTO," & _
        "Model_Grp.ModelGrp_Name, city.CityName, Fincity.cityname as FinCity," & _
        "Subgroup.Name, ColMast.Col_Desc, Model.MODEL, Model.Site_Code, Model.Chas_Type," & _
        "Model.Vehicle_Type, Model.Model_Type, Model.Model_Ind, Model.Sales_Desc," & _
        "Model.Model_Desc, Model.Model_Desc1, Model.Model_Desc2, Model.Grp_Code," & _
        "Model.Cat_Code, Model.Div_Code, Model.Wheel_Catg, Model.Active_YN, " & _
        "Model.STAT_IND, Model.TYRES, Model.TYRE_F, Model.TYRE_M, Model.TYRE_R," & _
        "Model.TYRE_FS, Model.TYRE_MS, Model.TYRE_RS, Model.RIMS, Model.SEAT," & _
        "Model.RLW,Model.HORSEPOWER , Model.FRONT_A_WT, Model.REAR_A_WT," & _
        "Model.UNLADEN_WT, Model.GROSS_WT, Model.WHEELBASE, Model.CYLINDER, Model.FUEL," & _
        "Model.TRADE_NO, Model.Manufacturer, Model.Warr_KMS, Model.Warr_Mth, Model.U_Name," & _
        "Model.U_EntDt, Model.U_AE, Model.Trf_Date, veh_Stock.ChassisNo , veh_Stock.EngineNo , Finbank.FinBankName,ContractFinance.FinName, ContractFinance.Add1 as FAdd1," & _
        "ContractFinance.add2 as Fadd2 , ContractFinance.PinCode as FPin , Subgroup.Add1, Subgroup.Add2, Subgroup.Add3, Subgroup.PIN " & _
        "FROM ((((((((((veh_order LEFT JOIN veh_Stock ON veh_order.Inv_DocId = veh_Stock.Sal_DocId) LEFT JOIN ColMast ON veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN Veh_Purch1 ON veh_Stock.Pur_DocId = Veh_Purch1.DocID) " & _
        "LEFT JOIN ContractFinance ON veh_order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Subgroup ON veh_order.PartyCode = Subgroup.SubCode) LEFT JOIN Model ON veh_order.MODEL = Model.MODEL) " & _
        "LEFT JOIN Model_Grp ON Model.Grp_Code = Model_Grp.ModelGrp_Code) LEFT JOIN city AS fincity ON ContractFinance.City = fincity.CityCode) LEFT JOIN Finbank ON ContractFinance.FinBankCode = Finbank.FinBankCode) " & _
        "LEFT JOIN city ON Subgroup.CityCode = city.CityCode) LEFT JOIN City AS city_1 ON Subgroup.TCityCode = city_1.CityCode " & _
        "where veh_order.Inv_DocId = '" & Master!SearchCode & "' and Veh_Order.DelCh_docid <> Null")

      
    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
    If Fob.FileExists("C:\RepPrint.Txt") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If Fob.FileExists("C:\RepPrint.Bat") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
    
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
 
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 4
    
    'Header
    'Form22A,NewRTOName,PrnDate,TempYN,Seet,Body,Narr,WtPrn,TempRto
    
    If txtPrint(TempYN) = "Yes" Then
        If RstCert!TCertiPrn_YN = 0 Then
            mDocStr = "Temporary Sale Certificate"
        Else
            mDocStr = "Temporary Sale Certificate (Duplicate)"
        End If
    Else
        If RstCert!CertiPrn_YN = 0 Then
            mDocStr = "Sale Certificate"
        Else
            mDocStr = "Sale Certificate (Duplicate)"
        End If
    End If
    
    mDupStr = ""

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
         Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & str(RstCompDet!V_SecCST_Date)), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & str(RstCompDet!V_SecLST_Date)), 40, , AlignRight)
        mHeader = mHeader + 1

        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth) & mChr18 & mEmph
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[Form 21 See Rule 47 (a) and (d)]", "C", PageWidth)
        mHeader = mHeader + 1
        
        Print #1, PSTR("The Registration Authority,", 40) & "Invoice No. : " & PSTR(str(Mid(RstCert!Inv_docid, 14, 8)), 8, , AlignLeft) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR((IIf(txtPrint(TempYN) = "No", IIf(txtPrint(NewRTOName) = "", RstCert!RTO, txtPrint(NewRTOName)), txtPrint(TempRTO))), 40) & mEmph & "Invoice Date : " & str(RstCert!inv_date) & mEmph1
        mHeader = mHeader + 1
        Print #1, "Ex. factory Price : " & Format(RstCert!VRATE, "0.00")
        mHeader = mHeader + 1
        
        Print #1, "(To be issued by the manufacturer,Dealer or Officer or defence(in Case of "
        mHeader = mHeader + 1
        Print #1, "Military auctioned vehicles)for presentation along with the application For"
        mHeader = mHeader + 1
        Print #1, "registration of a motor vehicle.)"
        mHeader = mHeader + 1
        
        Print #1, "Certified that - One " & RstCert!Model_Desc
        mHeader = mHeader + 1
        Print #1, "Has been delivered by us on " & RstCert!DelCh_Dt & " to :- "
        mHeader = mHeader + 1
        Print #1, PSTR("Name Of Buyer", 22) & " : " & RstCert!Name
        mHeader = mHeader + 1
        Print #1, PSTR("Son/Wife/Daughter of ", 22) & " : " & XNull(RstCert!FPrefix) & " " & XNull(RstCert!fname)
        mHeader = mHeader + 1
        If txtPrint(Tempadd) = "Yes" Then
            Print #1, PSTR("Address(Permanent)", 22) & " : " & XNull(RstCert!TAdd1) & XNull(RstCert!TAdd2)
            mHeader = mHeader + 1
            Print #1, XNull(RstCert!TAdd3) & XNull(RstCert!TCity) & XNull(RstCert!TPin)
            mHeader = mHeader + 1
            
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, PSTR("Address(Temporary)", 22) & " : " & XNull(RstCert!Add1) & XNull(RstCert!Add2)
            mHeader = mHeader + 1
            Print #1, XNull(RstCert!Add3) & XNull(RstCert!CityName) & XNull(RstCert!Pin)
            mHeader = mHeader + 1
        Else
            Print #1, PSTR("Address(Permanent)", 22) & " : " & " & XNull(RstCert!Add1) & XNull(RstCert!Add2) "
            mHeader = mHeader + 1
            Print #1, XNull(RstCert!Add3) & XNull(RstCert!CityName) & XNull(RstCert!Pin)
            mHeader = mHeader + 1
            
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, PSTR("Address(Temporary)", 22) & " : " & XNull(RstCert!TAdd1) & XNull(RstCert!TAdd2)
            mHeader = mHeader + 1
            Print #1, XNull(RstCert!TAdd3) & XNull(RstCert!TCity) & XNull(RstCert!TPin)
            mHeader = mHeader + 1
        End If

        If RstCert!Fund_Source = 0 Then
            Print #1, "The vehicle is held under agreement of Hypothication with " & RstCert!FinName
            mHeader = mHeader + 1
        ElseIf RstCert!Fund_Source = 1 Then
            Print #1, "The vehicle is held under agreement of Hire purchase with " & RstCert!FinName
            mHeader = mHeader + 1
        ElseIf RstCert!Fund_Source = 3 Then
            Print #1, "The vehicle is held under agreement of Lease with " & RstCert!FinName
            mHeader = mHeader + 1
        Else
            Print #1, ""
            mHeader = mHeader + 1
        End If
            
        Print #1, XNull(RstCert!FAdd1) & XNull(RstCert!FAdd2)
        mHeader = mHeader + 1
        Print #1, XNull(RstCert!FinCity) & XNull(RstCert!FPin)
        mHeader = mHeader + 1
        
        Print #1, "The detail of the vehicle are given below :  "
        mHeader = mHeader + 1
        
        Print #1, PSTR("Class of Vehicle", 22) & " : " & XNull(RstCert!ModelGrp_Name)
        mHeader = mHeader + 1
        Print #1, PSTR("Maker's Name", 22) & " : " & XNull(RstCert!Manufacturer)
        mHeader = mHeader + 1
        Print #1, PSTR("Chassis No.", 22) & " : " & XNull(RstCert!ChassisNo)
        mHeader = mHeader + 1
        Print #1, PSTR("Egine No.", 22) & " : " & XNull(RstCert!EngineNo)
        mHeader = mHeader + 1
        Print #1, PSTR("Horse Power Or Cubic Capacity", 22) & " : " & XNull(RstCert!HorsePower)
        mHeader = mHeader + 1
        Print #1, PSTR("Fuel Used", 22) & " : " & XNull(RstCert!FUEL)
        mHeader = mHeader + 1
        Print #1, PSTR("No. of cylenders ", 22) & " : " & XNull(RstCert!Cylinder)
        mHeader = mHeader + 1
        Print #1, PSTR("Month & Year Of Mfg.", 22) & " : " & XNull(RstCert!Mfg_Month) & " " & XNull(RstCert!Mfg_Yr)
        mHeader = mHeader + 1
        Print #1, PSTR("Seating Capacity(Incld. Driver)", 32) & " : " & IIf(txtPrint(TempYN) = "Yes", txtPrint(Seet), str(RstCert!Seat))
        mHeader = mHeader + 1
        Print #1, PSTR("Unleaden Weight", 22) & " : " & RstCert!Unladen_Wt & "Kg."
        mHeader = mHeader + 1
        Print #1, PSTR("Maximum Axle Weight and no.", 32) & " : "
        mHeader = mHeader + 1
        Print #1, PSTR("Maximum Axle Weight and no.", 32) & " : "
        mHeader = mHeader + 1
        Print #1, PSTR("and Description of tyres", 32) & " : "
        mHeader = mHeader + 1
        Print #1, PSTR("(In Case of Transport Vehicle)", 32) & " : "
        mHeader = mHeader + 1
        
        Print #1, "  " & PSTR("Front Axle", 22) & " : " & RstCert!Front_A_Wt
        mHeader = mHeader + 1
        Print #1, "  " & PSTR("Rear Axle", 22) & " : " & RstCert!Rear_A_Wt
        mHeader = mHeader + 1
        Print #1, "  " & PSTR("Any Other Axle", 22) & " : "
        mHeader = mHeader + 1
        Print #1, "  " & Space(25) & PSTR("Front", 10, , AlignRight) & PSTR("Moddle", 10, , AlignRight) & PSTR("Rear", 10, , AlignRight)
        mHeader = mHeader + 1
        Print #1, PSTR("No. Of Tyres", 24) & " : " & PSTR(RstCert!Tyre_F, 10) & PSTR(RstCert!Tyre_M, 10) & PSTR(RstCert!Tyre_R, 10)
        mHeader = mHeader + 1
        Print #1, PSTR("Size Of Tyres", 24) & " : " & PSTR(RstCert!Tyre_FS, 10) & PSTR(RstCert!Tyre_MS, 10) & PSTR(RstCert!Tyre_RS, 10)
        mHeader = mHeader + 1
        Print #1, PSTR("Colour Of Body", 24) & " : " & RstCert!Col_Desc
        mHeader = mHeader + 1
        Print #1, PSTR("Type of Body", 24) & " : " & IIf(txtPrint(TempYN) = "Yes", txtPrint(Body), RstCert!Intd_use)
        mHeader = mHeader + 1
        Print #1, PSTR("Trade No", 24) & " : " & RstCert!Trade_NO
        mHeader = mHeader + 1
        Print #1, PSTR("WheelBase", 24) & " : " & RstCert!WHEELBASE
        mHeader = mHeader + 1
        
        Do Until mHeader >= PageLength - mFooter
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        
        Print #1, mEmph & PSTR("For " & PubComp_Name, PageWidth, , AlignRight) & mEmph1
        Print #1, ""
        
        Print #1, IIf(txtPrint(TempYN) = "Yes", txtPrint(Narr), "Signature of Manufacturer/Dealer or Officer of Defence Department")
        Print #1, PSTR((IIf(txtPrint(TempYN) = "Yes", "", "*Strike out whichever is inapplicable")), 60) & "Authorised Signatory"
        
              
        
        'End Of Page 1 For SAle Certificate
        Print #1, mEject
        
'        mHeader = 0
'        mFooter = 2
'
'        Print #1, PRN_TIT(RstCert!Manufacturer, "A", PageWidth)
'        mHeader = mHeader + 1
'        Print #1, PRN_TIT(IIf(RstCert!CertiPrn_YN = 1, "DUPLICATE", ""), "C", PageWidth)
'        mHeader = mHeader + 1
'        Print #1, PRN_TIT("F O R M - 22-A", "C", PageWidth)
'        mHeader = mHeader + 1
'        Print #1, PRN_TIT("[See Rule 47 (g),124,126A AND 127]", "C", PageWidth)
'        mHeader = mHeader + 1
'        Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
'        mHeader = mHeader + 1
'        Print #1, PRN_TIT("Part 1", "B", PageWidth)
'        mHeader = mHeader + 1
'        Print #1, PRN_TIT("(Issued By The Manufacturer)", "C", PageWidth)
'        mHeader = mHeader + 1
'        Print #1, "Certified that Tata " & RstCert!Model_Desc & "(Brand name of the vehicle)"
'        mHeader = mHeader + 1
'        Print #1, "bearing Chassis Number " & RstCert!ChassisNo & "  and Engine Number " & RstCert!EngineNo
'        mHeader = mHeader + 1
'        Print #1, "complies with the provisions of  the  Motor Vehicles Act, 1988 and the rule made thereunder."
'        mHeader = mHeader + 1
'        Print #1, PSTR("Signature of the manufacturer", PageWidth, , AlignRight)
'        mHeader = mHeader + 1
'        Print #1, " "
'        mHeader = mHeader + 1
'        Print #1, " "
'        mHeader = mHeader + 1
'        Print #1, mEmph & PSTR("For " & RstCert!Manufacturer, PageWidth, , AlignRight) & mEmph1
'        mHeader = mHeader + 1
'
'        Do Until mHeader >= PageLength - mFooter
'            Print #1, ""
'            mHeader = mHeader + 1
'        Loop
'        Print #1, Replace(Space(PageWidth), " ", "-")
'
'        Print #1, mChr17 & RstCert!Inv_UName & " " & str(RstCert!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(RstCert!Inv_UName & " " & str(RstCert!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
'    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If Fob.FolderExists("c:\WinNt") Then
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
    Else
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        If txtPrint(TempYN) = "Yes" Then
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



Private Sub SpeedPrint()
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
    Dim i As Integer, j As Integer, mQry As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim Fob As New FileSystemObject
    Dim mJuriCity As String
    Dim Cnt As Byte, mAmt As Double, PrnStr As String, PrnStr1 As String
    Dim Left1 As String, Left2 As String, Left3 As String
    Dim Left4 As String, Left5 As String, Left6 As String, Left7 As String
    Dim Right1 As String, Right2 As String, Right3 As String
    Dim Right4 As String, Right5 As String, Right6 As String, Right7 As String
    Dim NetAmt As Double

     Set Rstsale = GCn.Execute("SELECT veh_order.*,Veh_Purch1.gate,City_1.CityName as fincity, ContractFinance.Add1 as finadd1, ContractFinance.Add2 as finadd2,finbank.finbankname,site.site_desc,ContractFinance.finname, SubGroup.Add3, City.CityName,  " & _
        " Veh_Stock.Pur_DocId, Veh_Stock.Sal_DocId, Veh_Stock.ChassisNo, Veh_Stock.EngineNo, Veh_Stock.PBILL_NO, Veh_Stock.PBILL_DATE, " & _
        "Model.Model_Desc,Model.Model_Desc1, " & _
        "ColMast.Col_Desc, SubGroup.Name, SubGroup.Add1, " & _
        "SubGroup.Add2,SubGroup.Add3,SubGroup.FPrefix,SubGroup.FName FROM ((((((((((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN ColMast ON " & _
        "Veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Site ON right(Veh_Order.Inv_SiteCode,1) = Site.Site_Code) " & _
        "LEFT JOIN FinBank ON ContractFinance.FinBankCode = FinBank.FinBankCode) " & _
        "LEFT JOIN City AS City_1 ON ContractFinance.City = City_1.CityCode) " & _
        "LEFT JOIN Veh_Purch1 ON Veh_Stock.Pur_DocId = Veh_Purch1.DocID  " & _
        "where veh_order.Inv_DocId = '" & Master!SearchCode & "'")
      
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
    If Fob.FileExists("C:\RepPrint.Txt") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If Fob.FileExists("C:\RepPrint.Bat") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For i = 1 To Len(Footer)
        If Mid(Footer, i, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 17
    mFooter = mFooter + FooterCnt
    
    ' Header
          
    mDocStr = "Sale Invoice"
    mDupStr = IIf(Rstsale!BillPrn_YN = 0, "", " (Duplicate)")

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
         Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
         mHeader = mHeader + 1
         Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignRight)
         mHeader = mHeader + 1

        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth) & mChr18 & mEmph
        mHeader = mHeader + 1
        
 '0 -Hypothication ,1- Hire purchase ,2 -Own Fund,3- Lease



    If Rstsale!Fund_Source = 0 Then   'Hypothication
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!Add1)
        Left5 = XNull(Rstsale!Add2)
        Left6 = XNull(Rstsale!Add3) & IIf(XNull(Rstsale!CityName) = "" Or XNull(Rstsale!Add3) = "", "", ",") & XNull(Rstsale!CityName)
        
        Right1 = "Under Hypothication to  "
        Right2 = XNull(Rstsale!FinBankName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!FIN_AMT, "0.00")
        
    ElseIf Rstsale!Fund_Source = 1 Then  'Hire Purchase
        Left1 = "Sold to under HPA with, "
        Left2 = " U/F " & XNull(Rstsale!FinBankName)
        Left3 = XNull(Rstsale!FinAdd1)
        Left4 = XNull(Rstsale!FinAdd2)
        Left5 = XNull(Rstsale!FinCity)
        Left6 = ""
           
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!Add1)
        Right5 = XNull(Rstsale!Add2)
        Right6 = XNull(Rstsale!Add3) & XNull(Rstsale!CityName)
    
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!Add1)
        Left5 = XNull(Rstsale!Add2)
        Left6 = XNull(Rstsale!Add3) & IIf(XNull(Rstsale!CityName) = "" Or XNull(Rstsale!Add3) = "", "", ",") & XNull(Rstsale!CityName)
        
        Right1 = "Leaser  "
        Right2 = XNull(Rstsale!FinBankName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Lease Amount :" & Rstsale!FIN_AMT
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!Add1)
        Left5 = XNull(Rstsale!Add2)
        Left6 = XNull(Rstsale!Add3) & IIf(XNull(Rstsale!CityName) = "" Or XNull(Rstsale!Add3) = "", "", ",") & IIf(XNull(Rstsale!CityName) = "" Or XNull(Rstsale!Add3) = "", "", ",") & XNull(Rstsale!CityName)
    End If

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
        
        Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
        
        Print #1, IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Telco Invoice No.: " & XNull(Rstsale!PBILL_NO) & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 40), Space(40)) & "Invoice No. : " & RstInvDet!VehSaleInv_Prefix & " " & PSTR(str(Rstsale!inv_No), 8, , AlignLeft) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR("Telco Gate Pass No. : " & XNull(Rstsale!GATE), 40) & mEmph & "Invoice Date : " & str(Rstsale!inv_date) & mEmph1
        mHeader = mHeader + 1
        Print #1, "Booking No. & Date : " & str(Rstsale!ord_no) & IIf(IsNull(Rstsale!ord_date), "", (str(Rstsale!ord_date)))
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        NetAmt = Rstsale!VRATE + Rstsale!MARGINE - Rstsale!Rebate _
        + Rstsale!InciChrg + Rstsale!Octroi + Rstsale!RegTemp _
        + Rstsale!TransitInsu + Rstsale!MVT + Rstsale!Transport
        
        Print #1, PSTR("Model : " & Rstsale!Model_Desc, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(NetAmt, "0.00"), 11, 2, AlignRight)
        mHeader = mHeader + 1
        Print #1, PSTR(Rstsale!Model_Desc1, 45) & PSTR(IIf(Rstsale!Tax_Per = 0, "", "Tax @ " & Format(Rstsale!Tax_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tax_Per, 11, 2)
        mHeader = mHeader + 1
        Print #1, PSTR("Colour      : " & Rstsale!Col_Desc, 40) & PSTR(IIf(Rstsale!surcharge_per = 0, "", "Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %"), 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2)
        mHeader = mHeader + 1
        
        Print #1, PSTR("Chassis No. : " & Rstsale!ChassisNo, 40) & PSTR("Other Charges", 27, , AlignRight) & ": " & PSTR(Rstsale!OtherChrg, 11, 2)
        mHeader = mHeader + 1
        Print #1, "Engine No.  : " & Rstsale!EngineNo & mEmph
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
        
        Cnt = 1
        
    Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
    "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")

     
        If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            Print #1, mChr17 & str(Cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) & Rst!Tax_Amt & Rst!TaxSur_Amt), 10, 2)
            mHeader = mHeader + 1
            Cnt = Cnt + 1
            Rst.MoveNext
        Loop
        End If
        
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        
        Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
        "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where veh_stock.Chassisno = '" & txt(ChassisNo) & "' " & _
        "group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
        
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
        
        NetAmt = Rstsale!VRATE + Rstsale!MARGINE - Rstsale!Rebate + Rstsale!InciChrg _
        + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
        + Rstsale!MVT + Rstsale!Transport + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!OtherChrg + Rstsale!FIT_AMT + Rstsale!FIT_TAX - Rstsale!DieselAmt + Rstsale!Round_off
        
        Print #1, PSTR(IIf(Rstsale!Round_off = 0, "", "Round Off"), 65, , AlignRight) & " : " & PSTR(Rstsale!Round_off, 12, 2)
        Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
        Print #1, mEmph & PSTR("Bill Amount", 65, , AlignRight) & " : " & PSTR(Amount_Fill(NetAmt, PubAmountPrefix), 12, 2, AlignRight)
        
        Print #1, ntow(NetAmt, "Rupees", "Paise") & mEmph1
        Print #1, ""
        
        Print #1, "Complete With Tools and equipment as supplied by the manufacturer including excise duty,Sales tax & delivery & handing charges."
        
        Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
        Print #1, ""
        Print #1, ""
        Print #1, "Accountant" & PSTR("Authorised Signatory", PageWidth - 10, , AlignRight)
        Print #1, ""
        Print #1, Replace(Space(PageWidth), " ", "-")
    
        Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
        
        Footer = Footer & vbLf
        j = 1
        For i = 1 To Len(Footer)
            If Mid(Footer, i, 1) = vbLf Then
                Print #1, RTrim(Mid(Footer, j, i - j))
                j = i + 1
            End If
        Next
        
        Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
               
        Print #1, mChr17 & Rstsale!Inv_UName & " " & str(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & str(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If Fob.FolderExists("c:\WinNt") Then
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
    Else
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If
   
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
    Dim i As Integer, j As Integer, mQry As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim Fob As New FileSystemObject
    Dim Cnt As Byte, NetAmt As Double, PrnStr As String, PrnStr1 As String, mRegCert As String
    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set Rstsale = GCn.Execute("SELECT veh_order.*,City.CityName,  " & _
        " Veh_Stock.ChassisNo, Veh_Stock.EngineNo,Model.Model_Desc,Model.Model_Desc1, " & _
        " SubGroup.Name, SubGroup.Add1,SubGroup.Add2,SubGroup.Add3,SubGroup.Tadd1,SubGroup.Tadd2,SubGroup.Tadd3 FROM  " & _
        "(((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode where veh_order.Inv_DocId = '" & Master!SearchCode & "'")

    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
    If Fob.FileExists("C:\RepPrint.Txt") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If Fob.FileExists("C:\RepPrint.Bat") = False Then
        Fob.CreateTextFile ("C:\RepPrint.Bat")
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
        If txtPrint(Tempadd) = "Yes" Then
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

        Print #1, "3. " & "Place Of Despatch : " & mEmph & PubComp_City & mEmph1
        mHeader = mHeader + 1
        Print #1, "4. " & "Destination : "
        mHeader = mHeader + 1
        Print #1, "5. " & "Description of consignment  : " & mEmph & Rstsale!Model_Desc & mEmph1
        mHeader = mHeader + 1
        Print #1, "6. " & PSTR("Quantity  : ", 15) & mEmph & "1 No. (One)" & mEmph1
        mHeader = mHeader + 1
        Print #1, "7. " & PSTR(" Weight : ", 15)
        mHeader = mHeader + 1
        NetAmt = Rstsale!VRATE + Rstsale!MARGINE - Rstsale!Rebate + Rstsale!InciChrg _
        + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
        + Rstsale!MVT + Rstsale!Transport + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!OtherChrg + Rstsale!FIT_AMT + Rstsale!FIT_TAX - Rstsale!DieselAmt + Rstsale!Round_off

        Print #1, "8. " & PSTR("Value  : ", 15) & mEmph & NetAmt & mEmph1
        mHeader = mHeader + 1
        Print #1, "9. " & "Consignor Bill/Cash Memo/Other"
        mHeader = mHeader + 1
        Print #1, "   " & "Document(Specify) No. and date :" & mEmph & Rstsale!inv_No & " Dt. " & Rstsale!inv_date & mEmph1
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
    If Fob.FolderExists("c:\WinNt") Then
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
    Else
        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub



Private Sub ProcAcPost()
On Error GoTo lblExit
        'A/c Posting related declarations
        Dim i As Integer, msgStr$, mBookDocID$
        Dim LedgAry(7) As LedgRec, mResult As Byte, mNarr$, rsCtrlAc As ADODB.Recordset
        Dim rsTemp As ADODB.Recordset
        
        Set rsCtrlAc = New ADODB.Recordset
        rsCtrlAc.CursorLocation = adUseClient
        rsCtrlAc.Open "Select Fitment_Ac,Fuel_Ac,VehROff_Ac From AcControls", GCnFa, adOpenStatic, adLockReadOnly
        If rsCtrlAc.RecordCount <= 0 Then
            msgStr = "Please Add Records in A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            GoTo lblExit
        End If
        If IsNull(rsCtrlAc!Fitment_Ac) Or rsCtrlAc!Fitment_Ac = "" Or _
            IsNull(rsCtrlAc!Fuel_Ac) Or rsCtrlAc!Fuel_Ac = "" Or _
            IsNull(rsCtrlAc!VehROff_Ac) Or rsCtrlAc!VehROff_Ac = "" Then
            msgStr = "Please define Fitment,Fuel and Round Off A/c's in A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            GoTo lblExit
        End If
        rsForm.MoveFirst        'Vehicle Sale A/c Code, Tax A/c Code, Surcharge A/c Code
        rsForm.FIND "Name ='" & txt(FormType) & "'"
        If IsNull(rsForm!PurSal_Ac_Code) Or rsForm!PurSal_Ac_Code = "" Then
            msgStr = "Please Define Sale A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
            GoTo lblExit
        End If
        'Tax A/c Code Checking
        If Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)) <> 0 Then
            If IsNull(rsForm!Tax_Ac_Code) Or rsForm!Sur_Ac_Code = "" Then
                msgStr = "Please Define Tax A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
                GoTo lblExit
            End If
        End If
        'Financier A/c Checking
        If Val(txt(FinAmt)) <> 0 Then
            If txt(FundSource) = "Hypothication" Or txt(FundSource) = "Hire Purchase" Then
                Set rsTemp = New ADODB.Recordset
                rsTemp.CursorLocation = adUseClient
                rsTemp.Open "Select Ac_Yn,AcCode From ContractFinance where FinCode='" & txt(FB_Code).Tag & "' ", GCn, adOpenStatic, adLockReadOnly
                If rsTemp!Ac_YN = "Y" Then
                    If rsTemp!AcCode = "" Or IsNull(rsTemp!AcCode) Then
                        msgStr = "Please define A/c Code in Financier Master" & vbCrLf & "A/c Posting Aborted !"
                        GoTo lblExit
                    End If
                End If
            End If
        End If
        
        'Sale Party A/c
        mBookDocID = GCn.Execute("select OrdDocId from Veh_Order where Inv_DocId='" & txt(TxtDocId) & "'").Fields(0).Value
        mNarr = "By Sales (Booking No." & DeCodeDocID(mBookDocID, For_Site_Code) & DeCodeDocID(mBookDocID, Current_Site) & _
            DeCodeDocID(mBookDocID, Document_Prefix) & Trim(DeCodeDocID(mBookDocID, Document_No)) & ")," & txt(Model)
        i = 0
        LedgAry(i).SubCode = txt(Party).Tag
        LedgAry(i).AmtDr = Round(Val(txt(GTotAmt)), 2)
        LedgAry(i).Narration = mNarr & " Telco Inv. No." & txt(TelcoInvNo)
        'Vehicle Sale A/c
        If Val(txt(SubTotB)) - (Val(txt(TaxAmt)) + Val(txt(TaxSurch))) <> 0 Then
            i = i + 1
            LedgAry(i).SubCode = rsForm!PurSal_Ac_Code
            LedgAry(i).AmtCr = Round(Val(txt(SubTotB)) - (Val(txt(TaxAmt)) + Val(txt(TaxSurch))), 2)
            LedgAry(i).Narration = mNarr & " Telco Inv. No." & txt(TelcoInvNo)
        End If
        'Fitment Amount
        If Val(txt(OthFitAmt)) <> 0 Then
            i = i + 1
            LedgAry(i).SubCode = rsCtrlAc!Fitment_Ac
            LedgAry(i).AmtCr = Round(Val(txt(OthFitAmt)), 2)
            LedgAry(i).Narration = mNarr & " Telco Inv. No." & txt(TelcoInvNo) & " Additional Fitments on Vehicle Sale Bill"
        End If
        'Tax Amt
        If Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)) <> 0 Then
            If rsForm!Tax_Ac_Code <> "" And rsForm!Sur_Ac_Code <> "" _
                 And rsForm!Tax_Ac_Code <> rsForm!Sur_Ac_Code Then
                If Val(txt(TaxAmt)) <> 0 Then
                    i = i + 1
                    LedgAry(i).SubCode = rsForm!Tax_Ac_Code
                    LedgAry(i).AmtCr = Round(Val(txt(TaxAmt)) + Val(txt(OthFitTax)), 2)
                    LedgAry(i).Narration = mNarr & " Sale Tax"
                End If
                If Val(txt(TaxSurch)) <> 0 Then
                    i = i + 1
                    LedgAry(i).SubCode = rsForm!Sur_Ac_Code
                    LedgAry(i).AmtCr = Round(Val(txt(TaxSurch)), 2)
                    LedgAry(i).Narration = mNarr & " Surcharge on Sales Tax"
                End If
            Else
                i = i + 1
                LedgAry(i).SubCode = rsForm!Tax_Ac_Code
                LedgAry(i).AmtCr = Round(Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)), 2)
                LedgAry(i).Narration = mNarr & " Sales Tax & Surcharge"
            End If
        End If
        If Val(txt(ROff)) <> 0 Then
            i = i + 1
            LedgAry(i).SubCode = rsCtrlAc!VehROff_Ac
            If Val(txt(ROff)) > 0 Then
                LedgAry(i).AmtCr = Round(Val(txt(ROff)), 2)
            Else
                LedgAry(i).AmtDr = Round(Abs(Val(txt(ROff))), 2)
            End If
            LedgAry(i).Narration = mNarr & " Round Off"
        End If
        'Fuel Amount
        If Val(txt(FuelAmt)) <> 0 Then
            i = i + 1
            LedgAry(i).SubCode = rsCtrlAc!Fuel_Ac
            LedgAry(i).AmtDr = Round(Val(txt(FuelAmt)), 2)
            LedgAry(i).Narration = mNarr & " Fuel Amount"
'            i = i + 1
'            LedgAry(i).SubCode = Txt(Party).Tag
'            LedgAry(i).AmtCr = Round(Val(Txt(FuelAmt)), 2)
'            LedgAry(i).Narration = mNarr & " Fuel Amount"
        End If
        
        If Val(txt(FinAmt)) <> 0 Then
            If txt(FundSource) = "Hypothication" Or txt(FundSource) = "Hire Purchase" Then
                If rsTemp!AcCode = "" Or IsNull(rsTemp!AcCode) Then
                Else
                    i = i + 1
                    LedgAry(i).SubCode = rsTemp!AcCode
                    LedgAry(i).AmtDr = Round(Val(txt(FinAmt)), 2)
                    LedgAry(i).Narration = mNarr & " Finance Amt."
                    i = i + 1
                    LedgAry(i).SubCode = txt(Party).Tag
                    LedgAry(i).AmtCr = Round(Val(txt(FinAmt)), 2)
                    LedgAry(i).Narration = mNarr & " Finance Amount."
                End If
            End If
        End If
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFa, txt(TxtDocId), CDate(txt(Vdate)))
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
lblExit:
If msgStr <> "" Then
    MsgBox msgStr, vbCritical, "A/c Posting"
ElseIf err.NUMBER > 0 Then
    MsgBox err.Description, vbCritical, "A/c Posting"
End If
Set rsCtrlAc = Nothing
Set rsTemp = Nothing
End Sub


