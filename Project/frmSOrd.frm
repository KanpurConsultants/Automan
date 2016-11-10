VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSOrd 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Sales Order Entry"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13710
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   13710
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Txt 
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
      Left            =   2475
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2445
      Width           =   6105
   End
   Begin VB.TextBox lblGroup 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   555
      TabIndex        =   119
      Top             =   3555
      Visible         =   0   'False
      Width           =   1785
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
      Height          =   1605
      Left            =   645
      TabIndex        =   105
      Top             =   4080
      Visible         =   0   'False
      Width           =   5025
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
         Left            =   4695
         MousePointer    =   99  'Custom
         Picture         =   "frmSOrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Delete Current Record"
         Top             =   0
         Width           =   315
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
         Picture         =   "frmSOrd.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmSOrd.frx":0678
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
         Left            =   3420
         MaskColor       =   &H00FFC0FF&
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmSOrd.frx":0982
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
         Left            =   3420
         MaskColor       =   &H00EFD5B8&
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmSOrd.frx":0C8C
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
         Left            =   3420
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Printer "
         Top             =   285
         Width           =   1590
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
         Left            =   7425
         TabIndex        =   110
         Top             =   555
         Visible         =   0   'False
         Width           =   300
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
         Left            =   7080
         TabIndex        =   109
         Top             =   285
         Visible         =   0   'False
         Width           =   330
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
         Left            =   7470
         TabIndex        =   108
         Top             =   300
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton Optpre 
         Alignment       =   1  'Right Justify
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
         Left            =   1725
         TabIndex        =   107
         Top             =   720
         Width           =   1200
      End
      Begin VB.OptionButton OptPlain 
         Alignment       =   1  'Right Justify
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
         Left            =   300
         TabIndex        =   106
         Top             =   720
         Width           =   750
      End
      Begin VB.Line Line8 
         X1              =   1470
         X2              =   1470
         Y1              =   510
         Y2              =   600
      End
      Begin VB.Line Line7 
         X1              =   2820
         X2              =   2820
         Y1              =   630
         Y2              =   735
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   360
         Y1              =   615
         Y2              =   720
      End
      Begin VB.Line Line6 
         X1              =   2820
         X2              =   345
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Index           =   41
         Left            =   -180
         TabIndex        =   118
         Top             =   255
         Width           =   3315
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
         Left            =   330
         TabIndex        =   117
         Top             =   1275
         Width           =   4650
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
         TabIndex        =   116
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   195
      Negotiate       =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   6810
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   4710
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Part No."
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
         Caption         =   "Part Name"
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
         DataField       =   "Unit"
         Caption         =   "Unit"
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
         DataField       =   "MRP"
         Caption         =   "      MRP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TB_SRate"
         Caption         =   "   TB Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TP_SRate"
         Caption         =   "   TP Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "LName"
         Caption         =   "Part Local Name"
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
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2564.788
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   6315
      TabIndex        =   72
      Top             =   7755
      Visible         =   0   'False
      Width           =   6285
      Begin VB.Line Line3 
         X1              =   3750
         X2              =   3750
         Y1              =   1035
         Y2              =   2070
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   2475
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line1 
         X1              =   1755
         X2              =   75
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bin Location"
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
         Index           =   12
         Left            =   3765
         TabIndex        =   103
         Top             =   255
         Width           =   1020
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Bin Loca>"
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
         Height          =   225
         Index           =   3
         Left            =   4920
         TabIndex        =   102
         Top             =   255
         Width           =   930
      End
      Begin VB.Label LblFrm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Part No.>"
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
         Index           =   0
         Left            =   1140
         TabIndex        =   101
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
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
         Index           =   20
         Left            =   75
         TabIndex        =   100
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   1800
         TabIndex        =   99
         Top             =   930
         Width           =   660
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000.00"
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
         Index           =   9
         Left            =   2745
         TabIndex        =   98
         Top             =   1185
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   225
         Index           =   23
         Left            =   4920
         TabIndex        =   97
         Top             =   1185
         Width           =   360
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   255
         Index           =   11
         Left            =   3285
         TabIndex        =   96
         Top             =   1875
         Width           =   360
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   210
         Index           =   14
         Left            =   5460
         TabIndex        =   95
         Top             =   1185
         Width           =   765
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   255
         Index           =   6
         Left            =   2115
         TabIndex        =   94
         Top             =   1635
         Width           =   360
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   255
         Index           =   7
         Left            =   2115
         TabIndex        =   93
         Top             =   1875
         Width           =   360
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   255
         Index           =   10
         Left            =   3285
         TabIndex        =   92
         Top             =   1635
         Width           =   360
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   13
         Left            =   5460
         TabIndex        =   91
         Top             =   1657
         Width           =   765
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000.000"
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
         Index           =   8
         Left            =   5130
         TabIndex        =   90
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   255
         Index           =   4
         Left            =   2100
         TabIndex        =   89
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Part Local Name>"
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
         Height          =   225
         Index           =   2
         Left            =   1140
         TabIndex        =   88
         Top             =   675
         Width           =   1590
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   5460
         TabIndex        =   87
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
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
         Height          =   225
         Index           =   19
         Left            =   4920
         TabIndex        =   86
         Top             =   1410
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00BBDBB3&
         BackStyle       =   0  'Transparent
         Caption         =   "Pur. Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   225
         Index           =   18
         Left            =   3930
         TabIndex        =   85
         Top             =   1185
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Name"
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
         Height          =   225
         Index           =   17
         Left            =   75
         TabIndex        =   84
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
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
         Index           =   16
         Left            =   4920
         TabIndex        =   83
         Top             =   1650
         Width           =   345
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
         Index           =   14
         Left            =   2805
         TabIndex        =   82
         Top             =   930
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Paid"
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
         Height          =   255
         Index           =   13
         Left            =   75
         TabIndex        =   81
         Top             =   1875
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Taxable"
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
         Index           =   11
         Left            =   75
         TabIndex        =   80
         Top             =   1185
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock"
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
         Height          =   255
         Index           =   10
         Left            =   3930
         TabIndex        =   79
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item Detail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   78
         Top             =   0
         Width           =   6285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable"
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
         Height          =   255
         Index           =   8
         Left            =   75
         TabIndex        =   77
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Name"
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
         Index           =   0
         Left            =   75
         TabIndex        =   76
         Top             =   465
         Width           =   885
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Part Name>"
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
         Index           =   1
         Left            =   1140
         TabIndex        =   75
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Taxpaid"
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
         Height          =   225
         Index           =   1
         Left            =   75
         TabIndex        =   74
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   2115
         TabIndex        =   73
         Top             =   1395
         Width           =   360
      End
      Begin VB.Line Line4 
         X1              =   3660
         X2              =   3885
         Y1              =   1035
         Y2              =   1035
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3330
      Left            =   4185
      Negotiate       =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4020
      Visible         =   0   'False
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   5874
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Party Name"
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
         DataField       =   "AcCode"
         Caption         =   "Ac.Code"
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
         DataField       =   "Add1"
         Caption         =   "Add1"
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
         DataField       =   "Add2"
         Caption         =   "Add2"
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
         DataField       =   "CreditDays"
         Caption         =   "Credit Days"
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
         DataField       =   "GovtParty"
         Caption         =   "Govt Party"
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
         DataField       =   "CityName"
         Caption         =   "City"
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
            DividerStyle    =   3
            ColumnWidth     =   4275.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDF4B5&
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
      Left            =   2550
      TabIndex        =   18
      Top             =   3975
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Txt 
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
      Index           =   4
      Left            =   1080
      MaxLength       =   40
      TabIndex        =   5
      Top             =   825
      Width           =   3900
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   9525
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   615
      Width           =   2295
   End
   Begin VB.TextBox Txt 
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
      Left            =   1080
      MaxLength       =   40
      TabIndex        =   4
      Top             =   555
      Width           =   3900
   End
   Begin VB.TextBox Txt 
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
      Left            =   10215
      MaxLength       =   11
      TabIndex        =   1
      Top             =   1185
      Width           =   1605
   End
   Begin VB.TextBox Txt 
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
      Index           =   5
      Left            =   1080
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1095
      Width           =   3900
   End
   Begin VB.TextBox Txt 
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
      Left            =   3690
      MaxLength       =   11
      TabIndex        =   10
      Top             =   1635
      Width           =   1290
   End
   Begin VB.TextBox Txt 
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
      Index           =   11
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1365
      Width           =   945
   End
   Begin VB.TextBox Txt 
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
      Height          =   255
      Index           =   20
      Left            =   3510
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6210
      Width           =   1260
   End
   Begin VB.TextBox Txt 
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
      Index           =   19
      Left            =   3510
      TabIndex        =   23
      Top             =   5940
      Width           =   1260
   End
   Begin VB.TextBox Txt 
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
      Index           =   18
      Left            =   2730
      TabIndex        =   22
      Top             =   5940
      Width           =   660
   End
   Begin VB.TextBox Txt 
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
      Height          =   585
      Index           =   13
      Left            =   5040
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   1050
      Width           =   3555
   End
   Begin VB.TextBox Txt 
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
      Height          =   555
      Index           =   14
      Left            =   5040
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   1860
      Width           =   3555
   End
   Begin VB.TextBox Txt 
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
      Left            =   10680
      TabIndex        =   17
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   2175
      Width           =   615
   End
   Begin VB.TextBox Txt 
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
      Index           =   17
      Left            =   3510
      TabIndex        =   21
      Top             =   5670
      Width           =   1260
   End
   Begin VB.TextBox Txt 
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
      Index           =   16
      Left            =   2730
      TabIndex        =   20
      Top             =   5670
      Width           =   660
   End
   Begin VB.TextBox Txt 
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   12
      Top             =   2175
      Width           =   3420
   End
   Begin VB.TextBox Txt 
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
      Left            =   1560
      MaxLength       =   35
      TabIndex        =   11
      Top             =   1905
      Width           =   3420
   End
   Begin VB.TextBox Txt 
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
      Left            =   5880
      MaxLength       =   20
      TabIndex        =   14
      Top             =   585
      Width           =   2700
   End
   Begin VB.TextBox Txt 
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
      Left            =   10815
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1455
      Width           =   1005
   End
   Begin VB.TextBox Txt 
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
      Index           =   12
      Left            =   3690
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1365
      Width           =   1290
   End
   Begin VB.TextBox Txt 
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
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1635
      Width           =   1485
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2535
      Left            =   60
      TabIndex        =   19
      Top             =   2820
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   26
      BackColorFixed  =   13623520
      ForeColorFixed  =   8388608
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      FocusRect       =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "HH"
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   26
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Order No. Detail :"
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
      Index           =   11
      Left            =   165
      TabIndex        =   120
      Top             =   2445
      Width           =   2280
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   27
      Left            =   4755
      TabIndex        =   71
      Top             =   5430
      Width           =   180
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   28
      Left            =   2445
      TabIndex        =   70
      Top             =   5430
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   25
      Left            =   3480
      TabIndex        =   69
      Top             =   5430
      Width           =   1170
   End
   Begin VB.Label LblIVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   68
      Top             =   5430
      Width           =   105
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5175
      TabIndex        =   67
      Top             =   5430
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   66
      Top             =   5430
      Width           =   1470
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   29
      Left            =   1410
      TabIndex        =   65
      Top             =   5430
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1260
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   3240
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      Height          =   270
      Left            =   8700
      TabIndex        =   64
      Top             =   900
      Width           =   660
   End
   Begin VB.Label LblSite 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code"
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
      Height          =   270
      Left            =   10200
      TabIndex        =   63
      Top             =   900
      Width           =   810
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
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
      Left            =   10200
      TabIndex        =   62
      Top             =   1455
      Width           =   600
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   61
      Top             =   825
      Width           =   690
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   4
      Left            =   915
      TabIndex        =   60
      Top             =   825
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc. ID"
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
      Index           =   31
      Left            =   8700
      TabIndex        =   59
      Top             =   615
      Width           =   585
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   25
      Left            =   9345
      TabIndex        =   58
      Top             =   615
      Width           =   45
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   57
      Top             =   1635
      Width           =   390
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   91
      Left            =   3570
      TabIndex        =   56
      Top             =   1635
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despatch Mode"
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
      Index           =   6
      Left            =   165
      TabIndex        =   55
      Top             =   1920
      Width           =   1290
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   54
      Top             =   1905
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party"
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
      Height          =   255
      Index           =   9
      Left            =   180
      TabIndex        =   53
      Top             =   555
      Width           =   405
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   3
      Left            =   915
      TabIndex        =   52
      Top             =   555
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Days"
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
      Height          =   255
      Index           =   5
      Left            =   2550
      TabIndex        =   51
      Top             =   1365
      Width           =   960
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   24
      Left            =   3570
      TabIndex        =   50
      Top             =   1365
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt.Party"
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
      Height          =   255
      Index           =   3
      Left            =   165
      TabIndex        =   49
      Top             =   1365
      Width           =   810
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   23
      Left            =   1440
      TabIndex        =   48
      Top             =   1365
      Width           =   195
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   22
      Left            =   2445
      TabIndex        =   47
      Top             =   6210
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Payable"
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
      Height          =   255
      Index           =   30
      Left            =   180
      TabIndex        =   46
      Top             =   6210
      Width           =   990
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   21
      Left            =   2445
      TabIndex        =   45
      Top             =   5940
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Other Charges"
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
      Height          =   255
      Index           =   29
      Left            =   180
      TabIndex        =   44
      Top             =   5940
      Width           =   1575
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   13
      Left            =   7515
      TabIndex        =   43
      Top             =   795
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Terms && Condition"
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
      Height          =   255
      Index           =   21
      Left            =   5040
      TabIndex        =   42
      Top             =   810
      Width           =   2415
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   270
      Index           =   12
      Left            =   7020
      TabIndex        =   41
      Top             =   1635
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Our Terms && Condition"
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
      Height          =   270
      Index           =   20
      Left            =   5040
      TabIndex        =   40
      Top             =   1635
      Width           =   1890
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   11
      Left            =   10545
      TabIndex        =   39
      Top             =   2175
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel Rest Orders"
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
      Height          =   255
      Index           =   19
      Left            =   8700
      TabIndex        =   38
      Top             =   2175
      Width           =   1635
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   7
      Left            =   2445
      TabIndex        =   37
      Top             =   5670
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   36
      Top             =   5670
      Width           =   735
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   35
      Top             =   570
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
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
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   34
      Top             =   570
      Width           =   690
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery To"
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
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   33
      Top             =   2175
      Width           =   900
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   32
      Top             =   2175
      Width           =   195
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   90
      Left            =   10050
      TabIndex        =   31
      Top             =   1185
      Width           =   195
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   92
      Left            =   1440
      TabIndex        =   30
      Top             =   1635
      Width           =   195
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   93
      Left            =   10035
      TabIndex        =   29
      Top             =   1455
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Order No"
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
      Height          =   255
      Index           =   0
      Left            =   8700
      TabIndex        =   28
      Top             =   1455
      Width           =   1170
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Ref.No"
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
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   27
      Top             =   1635
      Width           =   1020
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Order Date"
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
      Height          =   255
      Index           =   1
      Left            =   8715
      TabIndex        =   26
      Top             =   1185
      Width           =   1320
   End
End
Attribute VB_Name = "frmSOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsParty As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mVType As String
Dim mVPrefix As String
Dim ExitCtrl As Boolean
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function
Dim mSearchCode As String
'grid color scheme
Private Const CellBackColLeave As String = &HEDF7FE
'Private Const CellForeColLeave As String = &HFF00FF
'Private Const CellBackColEnter As String = &HF0D5BF
Private Const GridBackColorBkg As String = &HCFE0E0
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation


Dim mCheckNegetiveStockSiteWise As Boolean
Private Const DocID As Byte = 0                 ' Doc.ID
Private Const VDate As Byte = 1                 ' Date
Private Const SerialNo As Byte = 2              ' Sale Order No.
Private Const Party As Byte = 3                 ' Party Name
Private Const Address1 As Byte = 4              ' Address1
Private Const Address2 As Byte = 5              ' Address2
Private Const PartyRefNo As Byte = 6            ' Party Ref. No.
Private Const PartyRefDate As Byte = 7          ' Party Ref. Date
Private Const DispMode As Byte = 8              ' Dispatch Mode
Private Const Through As Byte = 9               ' Through
Private Const DeliveryTo As Byte = 10           ' Delivery To
Private Const GovtParty As Byte = 11            ' Govt.Party
Private Const CrDays As Byte = 12               ' Credit Days
Private Const CustTerms As Byte = 13            ' Customer Terms & Condition
Private Const OurTerms As Byte = 14             ' Our Terms && Condition
Private Const CancelRestOrders As Byte = 15     ' Cancel Rest Orders
Private Const DiscPer As Byte = 16              ' Discount%
Private Const DiscAmt As Byte = 17              ' Discount Amt.
Private Const AddOtherChrPer As Byte = 18       ' Add Other Charges%
Private Const AddOtherChrAmt As Byte = 19       ' Add Other Charges Amt.
Private Const NetAmt As Byte = 20               ' Net Amt
Private Const CustOrdDet As Byte = 21

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_PNo As Byte = 1               ' Part No
Private Const Col_Unit As Byte = 2              ' Unit
Private Const Col_MRP As Byte = 3               ' MRP Yes/No
Private Const Col_Taxable As Byte = 4           ' Taxable Yes/No
Private Const Col_Qty As Byte = 5               ' Qty
Private Const Col_Rate As Byte = 6              ' Rate
Private Const Col_MRPRate As Byte = 7           ' MRP Rate
Private Const Col_Amt As Byte = 8               ' Amt
Private Const Col_DiscPer As Byte = 9           ' Disc. %
Private Const Col_DiscAmt As Byte = 10          ' Disc. Amt.
Private Const Col_ItemVal As Byte = 11          ' Item Value
Private Const Col_PName As Byte = 12            ' Part Name
Private Const Col_LName As Byte = 13            ' Local Name
Private Const Col_MRPStkTB As Byte = 14         ' MRP Qty TB 'Current Stock Qty
Private Const Col_MRPStkTP As Byte = 15         ' MRP Qty TP
Private Const Col_TBStk As Byte = 16            ' Taxbale Qty
Private Const Col_TPStk As Byte = 17            ' Tax Paid Qty
Private Const Col_TBRate As Byte = 18           ' Taxbale Rate
Private Const Col_TPRate As Byte = 19           ' Tax Paid Rate
Private Const Col_Bin As Byte = 20              ' Bin
Private Const Col_LastRate As Byte = 21         ' Last Purchase Rate
Private Const Col_HPRate As Byte = 22           ' High Purchase Rate
Private Const Col_LPRate As Byte = 23           ' Low Purchase Rate
Private Const Col_PartGrade As Byte = 24        ' Part Grade (Used for Oil Item)
Private Const Col_EffectDate As Byte = 25       ' MRP Effective Date/TB Effective Date
''* Item Detail Column Declaration
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        If I = DocID Or I = Address1 Or I = Address2 Or I = GovtParty Or I = CrDays _
            Or I = NetAmt Then
        Else
            Txt(I).Enabled = Enb
        End If
    Next
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select SPO.OrderId As SearchCode " _
            & "From SP_Order SPO " _
            & "Where left(SPO.OrderId,1)='" & PubDivCode & "' and SPO.Order_Type='" & mVType & "' And SPO.OrderId = '" & MyValue & "'  " _
            & "Order by SPO.V_Date desc,SPO.Order_Type")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub
'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
    Next I
    Txt(DocID).Tag = ""
    LblDiv.CAPTION = "Division : "
    LblSite.CAPTION = "Site Code : "
    LblVPrefix.CAPTION = ""
    LblIVal.CAPTION = ""
    LblQty.CAPTION = ""

    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub
'* Used for intialize grid columns
Private Sub Grid_Ini()
'Serial No  | Part No |Unit | MRP Yes/No | Taxable Yes/No  | Qty | Rate | Amt | Disc. % | Disc. Amt. | Item Value | Part Name | Local Name
    With FGrid
        .left = Me.left '+ 60
        .width = Me.width - 90
        .top = 2825
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220

        .Cols = 26

        .TextMatrix(0, Col_SrNo) = "S.No."
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450

        .TextMatrix(0, Col_PNo) = "Part No."
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 1500

        .TextMatrix(0, Col_PName) = "Part Name"
        .ColAlignment(Col_PName) = flexAlignLeftCenter
        .ColWidth(Col_PName) = 2500

        .TextMatrix(0, Col_Unit) = "Unit"
        .ColAlignment(Col_Unit) = flexAlignLeftCenter
        .ColWidth(Col_Unit) = 550

        .TextMatrix(0, Col_MRP) = "MRP"
        .ColAlignment(Col_MRP) = flexAlignLeftCenter
        .ColWidth(Col_MRP) = 450

        .TextMatrix(0, Col_Taxable) = "Tax"
        .ColAlignment(Col_Taxable) = flexAlignLeftCenter
        .ColWidth(Col_Taxable) = 420

        .TextMatrix(0, Col_Qty) = "Qty"
        .ColAlignmentFixed(Col_Qty) = flexAlignRightCenter
        .ColWidth(Col_Qty) = 960

        .TextMatrix(0, Col_Rate) = "Rate"
        .ColAlignmentFixed(Col_Rate) = flexAlignRightCenter
        .ColWidth(Col_Rate) = 870

        .TextMatrix(0, Col_MRPRate) = "MRP Rate"
        .ColAlignmentFixed(Col_MRPRate) = flexAlignRightCenter
        .ColWidth(Col_MRPRate) = 0

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
        .ColWidth(Col_DiscPer) = 555

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
        .ColWidth(Col_DiscAmt) = 840

        .TextMatrix(0, Col_ItemVal) = "Item Value"
        .ColAlignmentFixed(Col_ItemVal) = flexAlignRightCenter
        .ColWidth(Col_ItemVal) = 1095

        .TextMatrix(0, Col_LName) = "Local Name"
        .ColAlignment(Col_LName) = flexAlignLeftCenter
        .ColWidth(Col_LName) = 2000

        .TextMatrix(0, Col_MRPStkTB) = "MRP Qty TB"
        .ColAlignmentFixed(Col_MRPStkTB) = flexAlignRightCenter
        .ColWidth(Col_MRPStkTB) = 0

        .TextMatrix(0, Col_MRPStkTP) = "MRP Qty TP"
        .ColAlignmentFixed(Col_MRPStkTP) = flexAlignRightCenter
        .ColWidth(Col_MRPStkTP) = 0

        .TextMatrix(0, Col_TBStk) = "Taxable Qty"
        .ColAlignmentFixed(Col_TBStk) = flexAlignRightCenter
        .ColWidth(Col_TBStk) = 0

        .TextMatrix(0, Col_TPStk) = "Tax Paid Qty"
        .ColAlignmentFixed(Col_TPStk) = flexAlignRightCenter
        .ColWidth(Col_TPStk) = 0

        .TextMatrix(0, Col_TBRate) = "Taxbale Rate"
        .ColAlignmentFixed(Col_TBRate) = flexAlignRightCenter
        .ColWidth(Col_TBRate) = 0

        .TextMatrix(0, Col_TPRate) = "Tax Paid Rate"
        .ColAlignmentFixed(Col_TPRate) = flexAlignRightCenter
        .ColWidth(Col_TPRate) = 0

        .TextMatrix(0, Col_Bin) = "Bin"
        .ColAlignmentFixed(Col_Bin) = flexAlignRightCenter
        .ColWidth(Col_Bin) = 0

        .TextMatrix(0, Col_LastRate) = "Last Purchase Rate"
        .ColAlignmentFixed(Col_LastRate) = flexAlignRightCenter
        .ColWidth(Col_LastRate) = 0

        .TextMatrix(0, Col_HPRate) = "High Purchase Rate"
        .ColAlignmentFixed(Col_HPRate) = flexAlignRightCenter
        .ColWidth(Col_HPRate) = 0

        .TextMatrix(0, Col_LPRate) = "Low Purchase Rate"
        .ColAlignmentFixed(Col_LPRate) = flexAlignRightCenter
        .ColWidth(Col_LPRate) = 0

        .TextMatrix(0, Col_PartGrade) = "Part Grade"
        .ColWidth(Col_PartGrade) = 0

        .TextMatrix(0, Col_EffectDate) = "Rate Effective Date"
        .ColWidth(Col_EffectDate) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (DGPart.top + mBotScale)
    FrmDetail.width = 6285: FrmDetail.left = 5595: FrmDetail.top = 405: FrmDetail.height = 2130
    DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale + 300
    
    FrmPrn.left = (Me.width - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
End Sub

Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub
Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
Dim Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
    FrmDetail.Visible = False
    If Master.RecordCount > 0 Then
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select Order_DocId from SP_Stock Where Order_DocId='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            TopCtrl1.tEdit = False
            TopCtrl1.tDel = False
        Else
            If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
            If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then Me.TopCtrl1.tDel = True
        End If
        Set Rst = Nothing
        
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select SP_Order.*,S.SubCode,S.Name As Party,S.Add1,S.Add2,S.Party_Type,S.Govt_YN,S.CreditDays " _
            & "From SP_Order " _
            & "Left Join SubGroup S on SP_Order.Party_Code=S.SubCode " _
            & "Where SP_Order.OrderId='" & Master!SearchCode & "' ", GCn, adOpenStatic, adLockReadOnly
        
        Txt(DocID).TEXT = Master1!OrderID
        Txt(DocID).Tag = Txt(DocID)
        mSearchCode = Txt(DocID)
        mVType = Master1!Order_Type
        mVPrefix = Master1!Order_Prefix
        LblDiv.CAPTION = "Division : " & left(Master1!OrderID, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        LblVPrefix.CAPTION = mID(Master1!OrderID, 9, 5)
        Txt(VDate).TEXT = Master1!V_DATE
        Txt(SerialNo).TEXT = Master1!Order_NO
        mPartyType = Master1!Party_Type
        Txt(Party).Tag = Master1!Party_code
        Txt(Party).TEXT = Master1!Party
        Txt(Address1).TEXT = Master1!Add1
        Txt(Address2).TEXT = Master1!Add2
        Txt(GovtParty).TEXT = IIf(Master1!Govt_YN = 0, "No", "Yes")
        Txt(CrDays).TEXT = VNull(Master1!CreditDays)
        Txt(PartyRefNo).TEXT = Master1!Order_Reg_No
        Txt(PartyRefDate).TEXT = IIf(IsNull(Master1!Order_Reg_Dt), "", Master1!Order_Reg_Dt)
        Txt(DispMode).TEXT = Master1!Dispatch_Mode
        Txt(Through).TEXT = Master1!Through
        Txt(CustOrdDet).TEXT = Master1!CustOrd_Det
        Txt(DeliveryTo).TEXT = Master1!Delivery_To
        Txt(CustTerms).TEXT = Master1!Terms
        Txt(OurTerms).TEXT = Master1!Terms2
        Txt(CancelRestOrders).TEXT = IIf(Master1!Cancel_RestOrders = 0, "No", "Yes")

        LblIVal.CAPTION = Format(Master1!Tot_Items, "0")
        LblQty.CAPTION = Format(Master1!Tot_Qty, "0.000")
        Txt(DiscPer).TEXT = Format(Master1!Disc_Per, "0.00")
        Txt(DiscAmt).TEXT = Format(Master1!Disc_Amt, "0.00")
        Txt(AddOtherChrPer).TEXT = Format(Master1!Add_Charge_Per, "0.00")
        Txt(AddOtherChrAmt).TEXT = Format(Master1!Add_Charge, "0.00")
        Txt(NetAmt).TEXT = Format(Master1!Tot_Amount, "0.00")
        
        FGrid.Redraw = False
        FGrid.Rows = 1
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select P.Part_Name,P.Local_Name,P.Unit,P.MRP,P.MRP_Effect_Dt, P.TB_SRate ,P.TP_SRate ,P.TB_Effect_Dt ,P.Part_Grade ,P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk ,P.Cur_TP_Stk ,P.Bin_Loca ,P.High_Pur_Rate ,P.Low_Pur_Rate,SP_Order1.*,SP_Order1.Rate as MRP_Rate From SP_Order1 Left Join Part P On SP_Order1.Part_No=P.Part_No and P.Div_Code = left(SP_Order1.OrderID,1) Where SP_Order1.OrderId='" & Master1!OrderID & "' order by Srl_No", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            I = 1
            Do Until Rst.EOF
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Qty) = Format(Rst!Qty, "0.000")

                    .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                    .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")
                    If Rst!MRP_YN = 1 Then
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty * Rst!MRP_Rate), "0.00")
                    Else
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty * Rst!Rate), "0.00")
                    End If
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                    .TextMatrix(I, Col_ItemVal) = Format(Rst!Amount, "0.00")
                    .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                    .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                    .TextMatrix(I, Col_MRPStkTP) = IIf(IsNull(Rst!Cur_MRP_TPStk), "", Rst!Cur_MRP_TPStk)
                    .TextMatrix(I, Col_MRPStkTB) = IIf(IsNull(Rst!Cur_MRP_TbStk), "", Rst!Cur_MRP_TbStk)
                    .TextMatrix(I, Col_TBStk) = IIf(IsNull(Rst!Cur_TB_STk), "", Rst!Cur_TB_STk)
                    .TextMatrix(I, Col_TPStk) = IIf(IsNull(Rst!Cur_TP_Stk), "", Rst!Cur_TP_Stk)
                    .TextMatrix(I, Col_TBRate) = IIf(IsNull(Rst!TB_SRate), "", Rst!TB_SRate)
                    .TextMatrix(I, Col_TPRate) = IIf(IsNull(Rst!TP_SRate), "", Rst!TP_SRate)
                    .TextMatrix(I, Col_Bin) = IIf(IsNull(Rst!Bin_Loca), "", Rst!Bin_Loca)
                    .TextMatrix(I, Col_LastRate) = ""
                    .TextMatrix(I, Col_HPRate) = IIf(IsNull(Rst!high_pur_rate), "", Rst!high_pur_rate)
                    .TextMatrix(I, Col_LPRate) = IIf(IsNull(Rst!low_pur_rate), "", Rst!low_pur_rate)
                    .TextMatrix(I, Col_PartGrade) = IIf(IsNull(Rst!Part_Grade), "", Rst!Part_Grade)
                    .TextMatrix(I, Col_EffectDate) = Format(IIf(Rst!MRP_YN = 1, IIf(IsNull(Rst!MRP_Effect_Dt), "", Rst!MRP_Effect_Dt), IIf(IsNull(Rst!TB_Effect_Dt), "", Rst!TB_Effect_Dt)), "dd/MMM/yyyy")
                End With

'                If Rst!Tax_YN = 1 Then
'                    mItemDiscTotTB = mItemDiscTotTB + Rst!Disc_Amt
'                Else
'                    mItemDiscTotTP = mItemDiscTotTP + Rst!Disc_Amt
'                End If
                Rst.MoveNext
                I = I + 1
            Loop
            FGrid.FixedRows = 1
            CountItem
        Else
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
        End If
        FGrid.Redraw = True
    Else
        BlankText
    End If
    Grid_Hide
Set Rst = Nothing
Set Master1 = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
    Select Case FGrid.Col
        Case Col_PNo, Col_PName, Col_LName
            If RsPart.RecordCount = 0 Then TxtGridLeave = False: ExitCtrl = False: Exit Function
            If ChkDuplicate = False Then TxtGridLeave = False: ExitCtrl = False: Exit Function
            TxtGridValid_PNo
            
        Case Col_Taxable, Col_MRP
            If ChkDuplicate = False Then TxtGridLeave = False: ExitCtrl = False: Exit Function
            TxtGridValid_TaxMRP
        Case Col_DiscPer, Col_DiscAmt, Col_Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(0).TEXT), "0.000")
            Amt_Cal
        Case Col_DiscAmt
            If Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) < Val(TxtGrid(0)) Then
                MsgBox "Item-wsie Disc. Amount is greater than Item Value", vbOKOnly, "Item-wise Disc. Checking"
                TxtGridLeave = False: Exit Function
            End If
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
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

Private Function ChkDuplicate() As Boolean
Dim I As Integer, X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte
    Select Case FGrid.Col
    Case Col_PNo, Col_PName, Col_LName
        Col1 = Col_MRP
        Col2 = Col_Taxable
        Col3 = FGrid.Col
    Case Col_MRP
        Col1 = Col_PNo
        Col2 = Col_Taxable
        Col3 = Col_MRP
    Case Col_Taxable
        Col1 = Col_PNo
        Col2 = Col_MRP
        Col3 = Col_Taxable
    End Select
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col2))) + CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))) + CStr(Trim(FGrid.TextMatrix(I, Col3))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
            TxtGrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function
'* Used for Calculate the Amount
Private Sub Amt_Cal()
Dim I As Integer
Dim NetAmount As Double
    FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
    FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) * Val(FGrid.TextMatrix(FGrid.Row, Col_DiscPer))) / 100), "0.00")
    FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))), "0.00")
    For I = 1 To FGrid.Rows - 1
        NetAmount = NetAmount + Val(FGrid.TextMatrix(I, Col_ItemVal))
    Next
    Txt(DiscAmt) = Format((NetAmount * Val(Txt(DiscPer))) / 100, "0.00")
    Txt(AddOtherChrAmt) = Format((NetAmount + Val(Txt(DiscAmt))) * Val(Txt(AddOtherChrPer)) / 100, "0.00")
    Txt(NetAmt) = Format(((NetAmount - Val(Txt(DiscAmt))) + Val(Txt(AddOtherChrAmt))), "0.00")
End Sub

Private Sub CountItem()
Dim I As Integer, TotItems As Integer, TotQty As Double
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            TotQty = TotQty + Val(FGrid.TextMatrix(I, Col_Qty))
            TotItems = TotItems + 1
        End If
    Next I
    LblIVal.CAPTION = Format(TotItems, "0")
    LblQty.CAPTION = Format(TotQty, "0.000")
End Sub


Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub TopCtrl1_ePrn()
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus

End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        Txt(Party).TEXT = RsParty!Name
        Txt(Party).Tag = RsParty!Code
        Txt(Address1).TEXT = IIf(IsNull(RsParty!Add1), "", RsParty!Add1)
        Txt(Address2).TEXT = IIf(IsNull(RsParty!Add2), "", RsParty!Add2)
        Txt(GovtParty).TEXT = IIf(RsParty!Govt_YN = 0, "No", "Yes")
        Txt(CrDays).TEXT = IIf(IsNull(RsParty!CreditDays), "", RsParty!CreditDays)
    End If
    Txt(Party).SetFocus
    DGParty.Visible = False
    lblGroup.Visible = False
End Sub

Private Sub DGPart_Click()
On Error GoTo ELoop
    If RsPart.RecordCount > 0 Then
        Select Case FGrid.Col
            Case Col_PNo
                TxtGrid(0).TEXT = RsPart!Code
            Case Col_PName
                TxtGrid(0).TEXT = RsPart!Name
            Case Col_LName
                TxtGrid(0).TEXT = RsPart!LName
        End Select
        TxtGridValid_PNo
    End If
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGPart.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte

    TopCtrl1.Tag = PubUParam:    WinSetting Me:     Grid_Ini
    Call Ini_Pub
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
'    Hook TxtGrid(0).hWnd
    Txt(VDate).Tag = PubLoginDate
    mVType = "S_SO"
    
    Set DGPart.DataSource = RsPart

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Add1,Add2,Govt_YN,CreditDays,Transporter,Party_Type,City.CityName from ((SubGroup " & _
        "left Join City on City.CityCode=SubGroup.CityCode) " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode) " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        "order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Dim sitecond As String
    sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("SPO.OrderId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubMoveRecYn Then
        Set Master = GCn.Execute("Select SPO.OrderId As SearchCode " _
            & "From SP_Order SPO " _
            & "Where left(SPO.OrderId,1)='" & PubDivCode & "' and SPO.Order_Type='" & mVType & "' " & sitecond & " " _
            & "Order by SPO.V_Date desc,SPO.Order_Type")
    Else
        Set Master = GCn.Execute("Select Top 1 SPO.OrderId As SearchCode " _
            & "From SP_Order SPO " _
            & "Where left(SPO.OrderId,1)='" & PubDivCode & "' and SPO.Order_Type='" & mVType & "' " & sitecond & " " _
            & "Order by SPO.V_Date desc,SPO.Order_Type")
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsParty = Nothing
    Set Master = Nothing
End Sub

Private Sub FrmDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmDetail.MousePointer = 15
End Sub

Private Sub FrmDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmDetail.MousePointer = 0
FrmDetail.Move X, Y
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    Txt(VDate).TEXT = Txt(VDate).Tag
    Txt(CancelRestOrders).TEXT = "No"
    mPartyType = 0
    Txt(DocID) = GetDocID(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
    Txt(VDate).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset
    'Check for existance of transactions
'    Set Rst = New ADODB.Recordset
'    Rst.CursorLocation = adUseClient
'    Rst.Open "Select Order_DocId from SP_Stock Where Order_DocId='" & Txt(DocId) & "'", GCn, adOpenDynamic, adLockOptimistic
'    If Rst.RecordCount  > 0 Then
'        MsgBox "Dispatch Challan Exists of this Sale Order, " & vbCrLf & "Can't Edit the Reocord", vbInformation, "Validation"
'        Exit Sub
'    End If

    Disp_Text SETS("EDIT", Me, Master)
    Txt(VDate).Enabled = False
    Txt(SerialNo).Enabled = False
    FGrid.AddItem FGrid.Rows
    Txt(Party).SetFocus
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant, Rst As ADODB.Recordset, mTrans As Boolean
'Check for existance of transactions
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select Order_DocId from SP_Stock Where Order_DocId='" & Txt(DocID) & "'", GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
        MsgBox "Dispatch Challan Exists of this Sale Order, " & vbCrLf & "Can't Delete the Reocord", vbInformation, "Validation"
        Exit Sub
    End If
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
            mTrans = True
                GCn.Execute ("Delete From SP_Order1 Where OrderId='" & Txt(DocID) & "'")
                GCn.Execute ("Delete From SP_Order Where OrderId='" & Txt(DocID) & "'")
            GCn.CommitTrans
            mTrans = False
            
            Master.Requery
            If Master.RecordCount > 0 Then
                If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            End If
            BUTTONS True, Me, Master, 0
            MoveRec
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
Set Rst = Nothing
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub

Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub

Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
     Dim sitecond As String
     sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("SPO.OrderId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    
    
    GSQL = "Select SPO.OrderId As SearchCode,SPO.Site_Code,SPO.Order_Prefix As VPrefix, " & cCStr("SPO.Order_No") & " As Order_No, " & cDt("SPO.V_Date") & " AS VDate, SubGroup.Name as PartyName FROM SP_Order SPO Left Join SubGroup On SPO.Party_Code = SubGroup.SubCode Where left(SPO.OrderId,1)='" & PubDivCode & "' and SPO.Order_Type='S_SO' " & sitecond & " Order by SPO.V_Date Desc,SPO.Order_Type"
    
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
    RsPart.Requery
    RsParty.Requery
    Master.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, DocIdHlp$, mGridFilled As Boolean
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(Txt(VDate), "Sale Order Date") = False Then Exit Sub
    If IsValid(Txt(SerialNo), "Sale Order Number") = False Then Exit Sub
    If IsValid(Txt(Party), "Party Name") = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Please Specify MRP Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Please Specify Taxable Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Please Specify Quantity in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Rate)) = 0 Then
'                If PubULabel <> "Y" Then
                    MsgBox "Please Specify Rate in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
'                End If
            End If
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_PNo: FGrid.SetFocus: Exit Sub
    GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2 = "Add" Then
        'lp 12-03-03
        Txt(DocID).Tag = Txt(DocID)
        If GCn.Execute("Select Count(*) From SP_Order Where OrderId='" & Txt(DocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Sale Order No. already exists, Retry", vbCritical, "Validation Error"
                Txt(SerialNo).SetFocus
                GoTo ELoop
            Else
                Txt(DocID) = GetDocID(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(Txt(DocID).Tag, Document_No)) Then
                    MsgBox "Sale Order No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo ELoop
                End If
            End If
        End If
        DocIdHlp = UCase(Replace(Txt(DocID), " ", ""))
        '********
        GCn.Execute "Insert Into SP_Order(" _
            & "OrderId,OrderIDHelp,Order_Type,Order_Prefix,Order_No," _
            & "Site_Code,V_Date,Party_Code,Order_Reg_No,Order_Reg_Dt," _
            & "Dispatch_Mode,Through,Delivery_To,Terms,Terms2," _
            & "Cancel_RestOrders,Tot_Items,Tot_Qty,Disc_Per,Disc_Amt," _
            & "Add_Charge_Per,Add_Charge,Tot_Amount,U_Name,U_EntDt," _
            & "U_AE,CustOrd_Det) " _
            & "Values(" _
            & "'" & Txt(DocID) & "','" & DocIdHlp & "','" & mVType & "','" & DeCodeDocID(Txt(DocID), Document_Prefix) & "'," & Txt(SerialNo) & "," _
            & "'" & PubSiteCode & PubSiteCode & "'," & ConvertDate(Format(Txt(VDate).TEXT, "dd/MMM/yyyy")) & ",'" & Txt(Party).Tag & "','" & Txt(PartyRefNo).TEXT & "'," & ConvertDate(Txt(PartyRefDate).TEXT) & "," _
            & "'" & Txt(DispMode).TEXT & "','" & Txt(Through).TEXT & "','" & Txt(DeliveryTo).TEXT & "','" & Txt(CustTerms).TEXT & "','" & Txt(OurTerms).TEXT & "'," _
            & "" & IIf(Txt(CancelRestOrders) = "Yes", 1, 0) & "," & Val(LblIVal.CAPTION) & "," & Val(LblQty.CAPTION) & "," & Val(Txt(DiscPer).TEXT) & "," & Val(Txt(DiscAmt).TEXT) & "," _
            & "" & Val(Txt(AddOtherChrPer).TEXT) & "," & Val(Txt(AddOtherChrAmt).TEXT) & "," & Val(Txt(NetAmt).TEXT) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & "," _
            & "'A','" & Txt(CustOrdDet) & "')"
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, Txt(DocID), Txt(VDate)
    Else
        GCn.Execute ("Delete From SP_Order1 Where OrderId='" & Txt(DocID) & "'")
        GCn.Execute "Update SP_Order Set " _
            & "Party_Code='" & Txt(Party).Tag & "',Order_Reg_No='" & Txt(PartyRefNo).TEXT & "'," _
            & "Order_Reg_Dt=" & ConvertDate(Txt(PartyRefDate).TEXT) & ",Dispatch_Mode='" & Txt(DispMode).TEXT & "'," _
            & "Through='" & Txt(Through).TEXT & "',Delivery_To='" & Txt(DeliveryTo).TEXT & "'," _
            & "Terms='" & Txt(CustTerms).TEXT & "',Terms2='" & Txt(OurTerms).TEXT & "'," _
            & "Cancel_RestOrders=" & IIf(Txt(CancelRestOrders) = "Yes", 1, 0) & ",Tot_Items=" & Val(LblIVal.CAPTION) & "," _
            & "Tot_Qty=" & Val(LblQty.CAPTION) & ",Disc_Per=" & Val(Txt(DiscPer).TEXT) & "," _
            & "Disc_Amt=" & Val(Txt(DiscAmt).TEXT) & ",Add_Charge_Per=" & Val(Txt(AddOtherChrPer).TEXT) & "," _
            & "Add_Charge=" & Val(Txt(AddOtherChrAmt).TEXT) & ",Tot_Amount=" & Val(Txt(NetAmt).TEXT) & "," _
            & "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & "," _
            & "U_AE='E', " _
            & "CustOrd_Det='" & Txt(CustOrdDet) & "'" _
            & "Where OrderId='" & Txt(DocID).TEXT & "'"
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            GCn.Execute "Insert Into SP_Order1(" _
                & "OrderId,Srl_No,Order_Type,Site_Code,V_Date,Party_Code," _
                & "Part_No,Qty,Tax_YN,MRP_YN,Rate," _
                & "Disc_Per,Disc_Amt,Amount,U_Name,U_EntDt," _
                & "U_AE) " _
                & "Values(" _
                & "'" & Txt(DocID) & "'," & I & ",'" & mVType & "','" & PubSiteCode & PubSiteCode & "'," & ConvertDate(Format(Txt(VDate).TEXT, "dd/MMM/yyyy")) & ",'" & Txt(Party).Tag & "'," _
                & "'" & FGrid.TextMatrix(I, Col_PNo) & "'," & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & "," _
                & "" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & "," _
                & "'" & left(TopCtrl1.TopText2, 1) & "')"
        End If
    Next
    GCn.CommitTrans
    mTrans = False
    mSearchCode = Txt(DocID)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select SPO.OrderId As SearchCode " _
            & "From SP_Order SPO " _
            & "Where left(SPO.OrderId,1)='" & PubDivCode & "' and SPO.Order_Type='" & mVType & "' And SPO.OrderId = '" & mSearchCode & "'  " _
            & "Order by SPO.V_Date desc,SPO.Order_Type")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"

    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(Txt(SerialNo)) > Val(DeCodeDocID(Txt(DocID).Tag, Document_No)) Then
            MsgBox "Sale Order No." & Trim(DeCodeDocID(Txt(DocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & Txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        Txt(VDate).Tag = Txt(VDate).TEXT
        TopCtrl1_eAdd
        Exit Sub
    End If
    TopCtrl1_ePrn
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
'        Master.FIND "OrderID='" & mSearchCode & "'"
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
    TxtGrid(0).Visible = False
    Grid_Hide
    Select Case Index
        Case Party
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).TEXT <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.FIND "Name ='" & Txt(Index).TEXT & "'"
            End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case Party
            If RsParty.RecordCount > 0 Then
                DGridTxtKeyDown DGParty, Txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            
            Else
                Txt_Validate Index, True
            End If
        Case DiscPer, AddOtherChrPer
            NumDown Txt(Index), KeyCode, 3, 2
        Case DiscAmt, AddOtherChrAmt
            NumDown Txt(Index), KeyCode, 8, 2
    End Select
    If DGParty.Visible = False Then
        If Index <> AddOtherChrAmt Then
            If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = AddOtherChrAmt Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If Index <> VDate And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If Index <> Party And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
    Select Case Index
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress Txt, Party, RsParty, KeyAscii, "Name"
        lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
    Case CancelRestOrders
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                Txt(Index).TEXT = "Yes"
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                Txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    Case DiscPer, AddOtherChrPer
        NumPress Txt(Index), KeyAscii, 3, 2
    Case DiscAmt, AddOtherChrAmt
        NumPress Txt(Index), KeyAscii, 8, 2
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case DiscAmt, DiscPer, AddOtherChrAmt, AddOtherChrPer
            Amt_Cal
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, I As Byte
On Error GoTo ELoop
    Select Case Index
        Case VDate
            Txt(Index).TEXT = RetDate(Txt(Index))
            Cancel = Not CheckFinYear(Txt(Index))
            If Cancel = False Then
                Txt(DocID) = GetDocID(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                Txt(DocID).Tag = Txt(DocID)
            End If
        Case SerialNo
            If VoucherEditFlag Then      ' Manual
                Txt(DocID) = GetDocID(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                Txt(DocID).Tag = Txt(DocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select Order_No From SP_Order Where OrderId='" & Txt(DocID).TEXT & "'", GCn, adOpenStatic, adLockReadOnly
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    Txt(SerialNo).SetFocus
                End If
            End If
        Case Party
            If Trim(Txt(Index).TEXT = "") Then
                MsgBox "Please Select Party", vbInformation, "Information"
                Txt(Index).SetFocus
                Cancel = True
                Exit Sub
            End If
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Party).Tag = ""
                Txt(Party).TEXT = ""
                Txt(Address1).TEXT = ""
                Txt(Address2).TEXT = ""
                Txt(GovtParty).TEXT = ""
                Txt(CrDays).TEXT = ""
                mPartyType = 0
            Else
                Txt(Party).Tag = RsParty!Code
                Txt(Party).TEXT = RsParty!Name
                Txt(Address1).TEXT = IIf(IsNull(RsParty!Add1), "", RsParty!Add1)
                Txt(Address2).TEXT = IIf(IsNull(RsParty!Add2), "", RsParty!Add2)
                Txt(GovtParty).TEXT = IIf(RsParty!Govt_YN = 0, "No", "Yes") 'IsNull(RsParty!GovtParty), "", RsParty!GovtParty)
                Txt(CrDays).TEXT = IIf(IsNull(RsParty!CreditDays), "", RsParty!CreditDays)
                mPartyType = RsParty!Party_Type
            End If
        Case PartyRefDate
            Txt(Index).TEXT = RetDate(Txt(Index))
        Case DiscPer, DiscAmt, AddOtherChrPer, AddOtherChrAmt
            If Val(Txt(Index).TEXT) = 0 Then
                Txt(Index).TEXT = ""
            Else
                Txt(Index).TEXT = Format(Txt(Index), "0.00")
            End If
    End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
'    Ctrl_GetFocus TxtGrid(Index)
    Grid_Hide
    If FrmDetail.Visible = False Then FrmDetail.Visible = True
'    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case Col_PNo
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "Code"
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "Code='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Col_PName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "Name"
            If FGrid.TextMatrix(FGrid.Row, Col_PName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_PName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Col_LName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "LName"
            If FGrid.TextMatrix(FGrid.Row, Col_LName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "LName ='" & FGrid.TextMatrix(FGrid.Row, Col_LName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        FGrid.SetFocus
        TxtGrid(0).Visible = False
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Col_PNo
            If DGPart.Visible = False Then DGridColSwap DGPart, 0
            DGridTxtKeyDown DGPart, TxtGrid, 0, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, 1
                End If
            End If
        Case Col_PName
            If DGPart.Visible = False Then DGridColSwap DGPart, 1
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 1, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                End If
            End If
        Case Col_LName
            If DGPart.Visible = False Then DGridColSwap DGPart, 2
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                End If
            End If
        Case Col_Taxable, Col_MRP
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_DiscAmt
                End If
            End If
        Case Col_Qty
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_DiscAmt
                End If
            End If
            CountItem
        Case Col_Rate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_DiscAmt, 2
                End If
            End If
        Case Col_DiscPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_DiscAmt
                End If
            End If
        Case Col_DiscAmt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_DiscAmt, 1
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    CheckQuote KeyAscii
    Select Case FGrid.Col
    Case Col_PNo
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Code"
    Case Col_PName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Name"
    Case Col_LName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "LName"
    Case Col_Qty
        NumPress TxtGrid(Index), KeyAscii, 8, 3
    Case Col_DiscPer
        NumPress TxtGrid(Index), KeyAscii, 2, 2
    Case Col_Rate, Col_DiscAmt
        NumPress TxtGrid(Index), KeyAscii, 8, 2
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case 0
        Select Case FGrid.Col
            Case Col_PNo
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Code", True
            Case Col_PName
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Name", True
            Case Col_LName
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "LName", True
            Case Col_Taxable, Col_MRP
                If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                    TxtGrid(Index) = ""
                ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
                    TxtGrid(Index) = "Yes"
                Else
                    TxtGrid(Index) = "No"
                End If
            Case Col_DiscPer, Col_DiscAmt, Col_Rate
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
            Case Col_Qty
                FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0.000")
                CountItem
        End Select
        Amt_Cal
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
    TxtGrid(Index).Visible = True
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.Col = Col_Unit Then Exit Sub
    Select Case FGrid.Col
        Case Col_PNo, Col_PName, Col_LName
            GridDblClick Me, FGrid, TxtGrid, 0
        Case Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_DiscPer, Col_DiscAmt
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                GridDblClick Me, FGrid, TxtGrid, 0
            End If
    End Select
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_EnterCell()
'    FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
'    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
    If TopCtrl1.TopText2 <> "Browse" Then
'        If FrmDetail.Visible = False Then
            MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
                FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
                Col_MRPStkTB, Col_MRPStkTP, _
                Col_TBStk, Col_TPStk, _
                Col_MRPRate, Col_TBRate, _
                Col_TPRate, Col_Bin, _
                Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
'        End If
        FrmDetail.Visible = True
    End If
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
        SendKeysA vbKeyTab, True
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid.Col
            Case Col_MRP, Col_Taxable
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            Case Col_Qty, Col_Amt, Col_DiscPer, Col_DiscAmt
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                Amt_Cal
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case FGrid.Col
            Case Col_PNo, Col_PName, Col_LName
                GridDblClick Me, FGrid, TxtGrid, 0
                TAddMode = False
            Case Col_Taxable, Col_MRP, Col_Qty, Col_Rate, Col_DiscPer, Col_DiscAmt
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    GridDblClick Me, FGrid, TxtGrid, 0
                    TAddMode = False
                End If
        End Select
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    Select Case FGrid.Col
        Case Col_PNo, Col_PName, Col_LName
           Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case Col_Unit
            FGrid_LeaveCell
            FGrid.Col = FGrid.Col + 1
            FGrid_EnterCell
            FGrid.SetFocus
        Case Col_MRP, Col_Taxable
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or _
                    UCase(Chr(KeyAscii)) = "Y" Or UCase(Chr(KeyAscii)) = "N" Then
                    'Allow keyascii
                Else
                    KeyAscii = 0
                End If
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            End If
        Case Col_Qty, Col_Rate, Col_DiscPer, Col_DiscAmt
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            End If
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
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
                FGrid.TextMatrix(I, Col_SrNo) = I
            Next
            CountItem
            Amt_Cal
            MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
                FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
                Col_MRPStkTB, Col_MRPStkTP, _
                Col_TBStk, Col_TPStk, _
                Col_MRPRate, Col_TBRate, _
                Col_TPRate, Col_Bin, _
                Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LeaveCell()
'    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
    If FrmDetail.Visible = True Then FrmDetail.Visible = False
End Sub

Private Sub FGrid_RowColChange()
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
            FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
            Col_MRPStkTB, Col_MRPStkTP, _
            Col_TBStk, Col_TPStk, _
            Col_MRPRate, Col_TBRate, _
            Col_TPRate, Col_Bin, _
            Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
    End If
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
'    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub TxtGridValid_PNo()
'Called from TxtGrid_Validate & TxtGridLeave procedures
Dim OldPNo$
If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, Col_PNo) = ""
    FGrid.TextMatrix(FGrid.Row, Col_PName) = ""
    FGrid.TextMatrix(FGrid.Row, Col_LName) = ""
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        "", "", "", _
        Col_Unit, Col_MRP, Col_Taxable, Col_MRPStkTB, Col_MRPStkTP, _
        Col_TBStk, Col_TPStk, _
        Col_MRPRate, Col_TBRate, _
        Col_TPRate, Col_Bin, _
        Col_HPRate, Col_LPRate, _
        Col_LastRate, Col_PartGrade, _
        Col_EffectDate, Col_DiscPer, mCheckNegetiveStockSiteWise
Else
    OldPNo = FGrid.TextMatrix(FGrid.Row, Col_PNo)
    FGrid.TextMatrix(FGrid.Row, Col_PNo) = RsPart!Code
    FGrid.TextMatrix(FGrid.Row, Col_PName) = RsPart!Name
    FGrid.TextMatrix(FGrid.Row, Col_LName) = RsPart!LName
    
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        RsPart!Code, RsPart!Name, RsPart!LName, _
        Col_Unit, Col_MRP, Col_Taxable, Col_MRPStkTB, Col_MRPStkTP, _
        Col_TBStk, Col_TPStk, _
        Col_MRPRate, Col_TBRate, _
        Col_TPRate, Col_Bin, _
        Col_HPRate, Col_LPRate, _
        Col_LastRate, Col_PartGrade, _
        Col_EffectDate, Col_DiscPer, mCheckNegetiveStockSiteWise
    
'by LPS 27-04-2K2
    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> OldPNo Then
            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(Txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
'            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
        End If
    End If
End If
If FGrid.TextMatrix(FGrid.Rows - 1, Col_PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

Private Sub TxtGridValid_TaxMRP()
FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'   If TopCtrl1.TopText2 = "Add" Or _
        TopCtrl1.TopText2 = "Edit" And Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) = 0 Then
        FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(Txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
'   End If
End If
Amt_Cal
End Sub

Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmPrn.Visible = False
        If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
            If TopCtrl1.TopText2.CAPTION = "Add" Then
                Txt(VDate).Tag = Txt(VDate).TEXT
                TopCtrl1_eAdd
                Exit Sub
            End If
            Disp_Text SETS("INI", Me, Master)
            MoveRec
        End If
    End If
End Sub
Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "SOrdPkLst", "SOrdPkLst")
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "SOrdPkLst", "SOrdPkLst")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        Txt(VDate).Tag = Txt(VDate).TEXT
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub WindowsPrint(Index As Integer)
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset
Dim mQry$
Dim I As Integer
Dim Rst2 As ADODB.Recordset
On Error GoTo ERRORHANDLER
'''        mQry = "SELECT " & cMID("SPO1.OrderId", "9", "13") & " as OrderId, SPO1.Srl_No, SPO1.QTY, " & cIIF("SPO1.TAX_YN = 0", "'No'", "'Yes'") & ", " & cIIF("SPO1.MRP_YN = 0", "'No'", "'Yes'") & ",SPO1.U_Name, SPO1.V_Date,SPO1.Order_Type,Part.PART_NO, Part.Part_Name, Part.Bin_Loca, SG.Name, SG.Add1,SG.Add2, SG.Add3, City.CityName ,SPO1.Rate,SPO1.Amount  " & _
'''            " FROM ((SP_Order1 SPO1 LEFT JOIN Part ON SPO1.PART_NO = Part.PART_NO and Part.Div_Code = left(SPO1.OrderID,1)) " & _
'''            " LEFT JOIN SubGroup SG ON SPO1.Party_Code = SG.SubCode) " & _
'''            " LEFT JOIN City ON SG.CityCode = City.CityCode" & _
'''            " where SPO1.OrderID='" & Txt(DocID) & "'"
          mQry = "SELECT " & cMID("SPO1.OrderId", "9", "13") & " as OrderId, SPO1.Srl_No, SPO1.QTY,Part.PART_NO, Part.Part_Name, Part.Bin_Loca,SPO1.Rate,SPO1.Amount, SG.Name, SG.Add1,SG.Add2, SG.Add3, City.CityName ,SPO1.Order_Type, SPO1.V_Date, " & cIIF("SPO1.TAX_YN = 0", "'No'", "'Yes'") & ", " & cIIF("SPO1.MRP_YN = 0", "'No'", "'Yes'") & ",SPO1.U_Name  " & _
                 " FROM ((SP_Order1 SPO1 LEFT JOIN Part ON SPO1.PART_NO = Part.PART_NO and Part.Div_Code = left(SPO1.OrderID,1)) " & _
                 " LEFT JOIN SubGroup SG ON SPO1.Party_Code = SG.SubCode) " & _
                 " LEFT JOIN City ON SG.CityCode = City.CityCode" & _
                 " where SPO1.OrderID='" & Txt(DocID) & "'"
        
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

        CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
        If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath & "\" & mRepName & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
Select Case Index
    Case 0  'Printer
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
            GCn.Execute "update SP_Order set Printed = 1  where SP_Order.OrderId='" & Master!SearchCode & "' "
        End If
        Set Rst = Nothing
        Set rpt = Nothing
    Case 1  'screen
        Call Report_View(rpt, Me.CAPTION, , True)
End Select
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrint()
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
    Dim I As Integer, j As Integer
    Dim PrintStr As String
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstOrder As ADODB.Recordset, RstOrder1 As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    
    Set RstOrder = GCn.Execute("SELECT SP_Order.Printed,SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, City.CityName, SP_Order.Order_No, SP_Order.V_Date, SP_Order.U_Name, SP_Order.U_EntDt " & _
    "FROM (SP_Order LEFT JOIN SubGroup ON SP_Order.Party_Code = SubGroup.SubCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode where SP_Order.OrderId = '" & Master!SearchCode & "'")

    Set RstOrder1 = GCn.Execute("SELECT SP_Order1.Srl_No,Part.Part_Name,Part.Bin_Loca ,SP_Order1.Part_No,SP_Order1.QTY ,SP_Order1.Rate ,SP_Order1.Amount  " & _
    "FROM SP_Order1 LEFT JOIN Part ON SP_Order1.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Order1.OrderID,1) where OrderId = '" & Master!SearchCode & "'")

    If RstOrder.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
   
    
   
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0
    mFooter = 3


    'Sale Bill Header

      mDocStr = "SALE ORDER PICK LIST"
      mDupStr = IIf(RstOrder!Printed = 1, "(DUPLICATE)", "")
      Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")

        Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
        mHeader = mHeader + 1
         If XNull(RstCompDet!S_SecSpeciality) <> "" Then
             Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
             mHeader = mHeader + 1
         End If
         Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
         If PubComp_Add2 <> "" Then
             Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
             mHeader = mHeader + 1
         End If
         If PubComp_City <> "" Then
             Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
             mHeader = mHeader + 1
         End If
         Print #1, PSTR(XNull(RstCompDet!S_SecLST) & IIf(XNull(RstCompDet!S_SecLST_Date) = "", "", " Dt. " & RstCompDet!S_SecLST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!S_SecPhone)), 27, , AlignRight, " ")
         mHeader = mHeader + 1
         Print #1, PSTR(XNull(RstCompDet!S_SecCST) & IIf(XNull(RstCompDet!S_SecCST_Date) = "", "", " Dt. " & RstCompDet!S_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!S_SecFax)), 27, , AlignRight, " ")
         mHeader = mHeader + 1

        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
        mHeader = mHeader + 1

        Print #1, mChr18 & "To," & mEmph
        mHeader = mHeader + 1
        Print #1, PSTR("M/s " & RstOrder!Name, 40) & Space(1) & PSTR("Sale Order No.", 16) & " : " & PSTR(STR(RstOrder!Order_NO), 14) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstOrder!Add1), 40) & Space(1) & mEmph & PSTR("Sale Order Date ", 16) & " : " & PSTR(STR(RstOrder!V_DATE), 14) & mEmph1
        mHeader = mHeader + 1
        Print #1, XNull(RstOrder!Add2)
        mHeader = mHeader + 1
        Print #1, XNull(RstOrder!Add3) & IIf(XNull(RstOrder!CityName) <> "", ",", "") & XNull(RstOrder!CityName)
        mHeader = mHeader + 1

        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        mHeader = mHeader + 1
        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 16) & PSTR("DESCRIPTION", 28) & PSTR("BIN LOCATION", 15) & PSTR("QUANTITY", 12, , AlignRight) & mDoub1 & PSTR("RATE", 12, , AlignRight) & mDoub1 & PSTR("AMOUNT", 12, , AlignRight) & mDoub1
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        mFix = PageLength - (mHeader + mFooter)
        Page = 1
        mLine = 1
        mSlNo = 1
        
        If RstOrder1.RecordCount > 0 Then
            Do Until RstOrder1.EOF
                If mLine > mFix Then
                    Page = Page + 1
                    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-")
                    Print #1, Space(PageWidth - Len("Contd. on next page.." + STR(Page))) & "Contd. on next page.." & STR(Page)
                    Do Until mLine >= mFix + mFooter - 2
                        Print #1, ""
                        mLine = mLine + 1
                    Loop
                    Print #1, mEject

                    'Header On Second Page
                    mHeader = 0
                    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                    mHeader = mHeader + 1

                    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
                    mHeader = mHeader + 1

                    Print #1, mChr18 & "To," & mEmph
                    mHeader = mHeader + 1
                    Print #1, PSTR("M/s " & RstOrder!Name, 40) & Space(1) & PSTR("Sale Order No.", 16) & " : " & PSTR(STR(RstOrder!Order_NO), 14) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, PSTR(XNull(RstOrder!Add1), 40) & Space(1) & mEmph & PSTR("Sale Order Date ", 16) & " : " & PSTR(STR(RstOrder!V_DATE), 14) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, XNull(RstOrder!Add2)
                    mHeader = mHeader + 1
                    Print #1, XNull(RstOrder!Add3) & IIf(XNull(RstOrder!CityName) <> "", ",", "") & XNull(RstOrder!CityName)
                    mHeader = mHeader + 1
            
                    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
                    mHeader = mHeader + 1
                    Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 16) & PSTR("DESCRIPTION", 28) & PSTR("BIN LOCATION", 15) & PSTR("QUANTITY", 12, , AlignRight) & mDoub1
                    mHeader = mHeader + 1
                    Print #1, Replace(Space(PageWidth), " ", "-")
                    mHeader = mHeader + 1
                    mFix = PageLength - (mHeader + mFooter)
                    mLine = 1
                End If
                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & mChr17 & PSTR(RstOrder1!Part_No, 27, , AlignLeft) & PSTR(RstOrder1!Part_Name, 50) & mChr18 & PSTR(RstOrder1!Bin_Loca, 15) & PSTR(RstOrder1!Qty, 12, 3) & PSTR(RstOrder1!Rate, 12, 3) & PSTR(RstOrder1!Amount, 12, 3)
                Print #1, PrintStr
            RstOrder1.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, "User : " & RstOrder!U_Name & " " & STR(RstOrder!U_EntDt) & mChr17
    'Print #1, Space(((PageWidth * 1.7) - Len("")) / 2) & "" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update SP_Order set Printed = 1  where SP_Order.OrderId='" & Master!SearchCode & "' "
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub


Sub Ini_Pub()
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl")
    If RsTemp.RecordCount > 0 Then
        mCheckNegetiveStockSiteWise = VNull(RsTemp!CheckNegetiveStockSiteWise)
    End If
End Sub

