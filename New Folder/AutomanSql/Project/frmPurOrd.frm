VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPurOrd 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Purchase Order"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12765
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
   ScaleHeight     =   8850
   ScaleWidth      =   12765
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Create CSV File"
      Height          =   315
      Left            =   6945
      TabIndex        =   148
      Top             =   30
      Width           =   2445
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
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   147
      Text            =   "444444"
      Top             =   3600
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
      Left            =   3000
      TabIndex        =   133
      Top             =   4320
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
         Picture         =   "frmPurOrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   143
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
         Picture         =   "frmPurOrd.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmPurOrd.frx":0678
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
         TabIndex        =   141
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmPurOrd.frx":0982
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
         TabIndex        =   140
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmPurOrd.frx":0C8C
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
         TabIndex        =   139
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         Left            =   -105
         TabIndex        =   146
         Top             =   300
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
         TabIndex        =   145
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
         TabIndex        =   144
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   -1815
      Negotiate       =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   7485
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
      ColumnCount     =   8
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
         DataField       =   "CurrStk"
         Caption         =   "      Stk."
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
            ColumnWidth     =   1094.74
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
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2564.788
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   3375
      Negotiate       =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   2985
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   10845
      TabIndex        =   100
      Top             =   -960
      Visible         =   0   'False
      Width           =   6285
      Begin VB.Line Line4 
         X1              =   3660
         X2              =   3885
         Y1              =   1035
         Y2              =   1035
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
         TabIndex        =   131
         Top             =   1395
         Width           =   360
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
         Index           =   48
         Left            =   75
         TabIndex        =   130
         Top             =   1410
         Width           =   1080
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
         TabIndex        =   129
         Top             =   465
         Width           =   1095
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
         Index           =   15
         Left            =   75
         TabIndex        =   128
         Top             =   465
         Width           =   885
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
         TabIndex        =   127
         Top             =   1635
         Width           =   645
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
         TabIndex        =   126
         Top             =   0
         Width           =   6285
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
         TabIndex        =   125
         Top             =   915
         Width           =   1110
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
         TabIndex        =   124
         Top             =   1185
         Width           =   1080
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
         TabIndex        =   123
         Top             =   1875
         Width           =   705
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
         TabIndex        =   122
         Top             =   930
         Width           =   810
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
         TabIndex        =   121
         Top             =   1650
         Width           =   345
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
         TabIndex        =   120
         Top             =   675
         Width           =   1005
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
         TabIndex        =   119
         Top             =   1185
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
         TabIndex        =   118
         Top             =   1410
         Width           =   390
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
         TabIndex        =   117
         Top             =   1410
         Width           =   765
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
         TabIndex        =   116
         Top             =   675
         Width           =   1590
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
         TabIndex        =   115
         Top             =   1170
         Width           =   360
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
         TabIndex        =   114
         Top             =   930
         Width           =   1095
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
         TabIndex        =   113
         Top             =   1657
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
         Index           =   10
         Left            =   3285
         TabIndex        =   112
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
         TabIndex        =   111
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
         Index           =   6
         Left            =   2115
         TabIndex        =   110
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
         ForeColor       =   &H00C000C0&
         Height          =   210
         Index           =   14
         Left            =   5460
         TabIndex        =   109
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   3285
         TabIndex        =   108
         Top             =   1875
         Width           =   360
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
         TabIndex        =   107
         Top             =   1185
         Width           =   360
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
         TabIndex        =   106
         Top             =   1185
         Width           =   885
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
         TabIndex        =   105
         Top             =   930
         Width           =   660
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
         TabIndex        =   104
         Top             =   270
         Width           =   660
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
         TabIndex        =   103
         Top             =   255
         Width           =   870
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
         TabIndex        =   101
         Top             =   255
         Width           =   1020
      End
      Begin VB.Line Line1 
         X1              =   1755
         X2              =   75
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   2475
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line3 
         X1              =   3750
         X2              =   3750
         Y1              =   1035
         Y2              =   2070
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   661
      Enabled         =   -1  'True
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   8805
      TabIndex        =   94
      Top             =   6945
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   95
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
         BackColor       =   16777152
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Height          =   930
      Index           =   23
      Left            =   1980
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   6090
      Width           =   4995
   End
   Begin MSDataGridLib.DataGrid DGJCNo 
      Height          =   2145
      Left            =   2055
      Negotiate       =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   2895
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   3784
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Job Card  No"
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
            ColumnWidth     =   2745.071
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
      Index           =   24
      Left            =   9480
      MaxLength       =   15
      TabIndex        =   28
      Top             =   6360
      Width           =   570
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
      Index           =   22
      Left            =   8535
      MaxLength       =   12
      TabIndex        =   27
      Top             =   6090
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
      Index           =   21
      Left            =   8535
      MaxLength       =   35
      TabIndex        =   26
      Top             =   5820
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
      Index           =   19
      Left            =   8535
      MaxLength       =   35
      TabIndex        =   25
      Top             =   5550
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
      Index           =   20
      Left            =   1980
      MaxLength       =   20
      TabIndex        =   23
      Top             =   5820
      Width           =   2355
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
      Left            =   1980
      MaxLength       =   12
      TabIndex        =   22
      Top             =   5550
      Width           =   2355
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Left            =   540
      MaxLength       =   25
      TabIndex        =   20
      Top             =   4380
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   4935
      Left            =   3270
      Negotiate       =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5985
      Visible         =   0   'False
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
      ColumnCount     =   3
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
         DataField       =   "Add1"
         Caption         =   "Address"
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
            ColumnWidth     =   4545.071
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2580
      Left            =   15
      TabIndex        =   21
      Top             =   2655
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   4551
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   0
      Cols            =   24
      BackColorFixed  =   15259902
      ForeColorFixed  =   8388608
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   15259902
      GridColorFixed  =   8421504
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"frmPurOrd.frx":0F96
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
      _Band(0).Cols   =   24
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
      Left            =   10065
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
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
      Left            =   9330
      MaxLength       =   16
      TabIndex        =   2
      Top             =   1140
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
      Index           =   4
      Left            =   9600
      MaxLength       =   12
      TabIndex        =   18
      Top             =   2070
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
      Index           =   5
      Left            =   9600
      MaxLength       =   16
      TabIndex        =   19
      Top             =   2340
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   3
      Left            =   1590
      MaxLength       =   40
      TabIndex        =   5
      Top             =   615
      Width           =   4845
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   9330
      MaxLength       =   21
      TabIndex        =   1
      Top             =   615
      Width           =   2340
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
      Left            =   1545
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1110
      Width           =   1320
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
      Left            =   4245
      MaxLength       =   20
      TabIndex        =   11
      Top             =   1110
      Width           =   1080
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
      Index           =   12
      Left            =   1545
      MaxLength       =   16
      TabIndex        =   9
      Top             =   1920
      Width           =   1470
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
      Index           =   9
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   10
      Top             =   2190
      Width           =   795
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
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1650
      Width           =   1905
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
      Left            =   5400
      MaxLength       =   15
      TabIndex        =   15
      Top             =   1920
      Width           =   1905
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
      Left            =   5835
      MaxLength       =   15
      TabIndex        =   17
      Top             =   2190
      Width           =   1470
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEE0FD&
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
      Left            =   3375
      MaxLength       =   6
      TabIndex        =   16
      Top             =   2190
      Width           =   990
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
      Index           =   11
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1650
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
      Index           =   25
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1380
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
      Index           =   13
      Left            =   6090
      MaxLength       =   7
      TabIndex        =   12
      Top             =   1110
      Width           =   780
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
      Left            =   9330
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1410
      Width           =   2325
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
      Left            =   1545
      MaxLength       =   14
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel Remaining Order"
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
      Left            =   7065
      TabIndex        =   80
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spare I/C Name"
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
      Index           =   43
      Left            =   7035
      TabIndex        =   79
      Top             =   5820
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I/C Designation"
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
      Index           =   46
      Left            =   7050
      TabIndex        =   78
      Top             =   6090
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terms && Condition"
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
      Index           =   44
      Left            =   210
      TabIndex        =   64
      Top             =   6090
      Width           =   1545
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   25
      Left            =   1860
      TabIndex        =   63
      Top             =   6090
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   24
      Left            =   9330
      TabIndex        =   62
      Top             =   6360
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   23
      Left            =   8400
      TabIndex        =   61
      Top             =   6090
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   22
      Left            =   8400
      TabIndex        =   60
      Top             =   5820
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price List Reference"
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
      Index           =   42
      Left            =   90
      TabIndex        =   59
      Top             =   5550
      Width           =   1665
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   41
      Left            =   1065
      TabIndex        =   58
      Top             =   5820
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode of Dispatch"
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
      Index           =   40
      Left            =   6930
      TabIndex        =   57
      Top             =   5550
      Width           =   1425
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   21
      Left            =   8400
      TabIndex        =   56
      Top             =   5550
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   20
      Left            =   1860
      TabIndex        =   55
      Top             =   5820
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   19
      Left            =   1860
      TabIndex        =   54
      Top             =   5550
      Width           =   45
   End
   Begin VB.Label LblAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   9435
      TabIndex        =   51
      Top             =   5280
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amoumt"
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
      Index           =   22
      Left            =   7725
      TabIndex        =   50
      Top             =   5280
      Width           =   1185
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   8955
      TabIndex        =   49
      Top             =   5280
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   3
      Left            =   5595
      TabIndex        =   46
      Top             =   5280
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   1875
      TabIndex        =   45
      Top             =   5280
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Index           =   25
      Left            =   4770
      TabIndex        =   44
      Top             =   5280
      Width           =   705
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
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2190
      TabIndex        =   43
      Top             =   5280
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
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6135
      TabIndex        =   42
      Top             =   5280
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   285
      TabIndex        =   41
      Top             =   5280
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   1575
      TabIndex        =   40
      Top             =   5280
      Width           =   45
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   1500
      Left            =   7845
      Top             =   525
      Width           =   3930
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division            :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   8010
      TabIndex        =   98
      Top             =   885
      Width           =   1245
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   9720
      TabIndex        =   97
      Top             =   885
      Width           =   1125
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9390
      TabIndex        =   96
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00D7C6C8&
      Caption         =   "VOR Detail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   24
      Left            =   150
      TabIndex        =   93
      Top             =   900
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   1500
      Left            =   30
      Top             =   1020
      Width           =   7380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Steering Make"
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
      Index           =   68
      Left            =   3960
      TabIndex        =   92
      Top             =   1935
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Radiator Make"
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
      Index           =   71
      Left            =   4425
      TabIndex        =   91
      Top             =   2205
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brake Type"
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
      Index           =   70
      Left            =   2235
      TabIndex        =   90
      Top             =   2205
      Width           =   915
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   70
      Left            =   3240
      TabIndex        =   89
      Top             =   2205
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   71
      Left            =   5700
      TabIndex        =   88
      Top             =   2205
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   68
      Left            =   5220
      TabIndex        =   87
      Top             =   1905
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   62
      Left            =   5910
      TabIndex        =   86
      Top             =   1110
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KMs "
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
      Index           =   62
      Left            =   5430
      TabIndex        =   85
      Top             =   1110
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   64
      Left            =   5220
      TabIndex        =   84
      Top             =   1395
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
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
      Index           =   64
      Left            =   4095
      TabIndex        =   83
      Top             =   1395
      Width           =   1035
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   66
      Left            =   5220
      TabIndex        =   82
      Top             =   1665
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Steering Type"
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
      Index           =   66
      Left            =   4005
      TabIndex        =   81
      Top             =   1665
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No."
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
      Index           =   63
      Left            =   150
      TabIndex        =   76
      Top             =   1395
      Width           =   930
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   63
      Left            =   1365
      TabIndex        =   75
      Top             =   1395
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   69
      Left            =   1230
      TabIndex        =   74
      Top             =   2205
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel Base"
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
      Index           =   69
      Left            =   150
      TabIndex        =   73
      Top             =   2205
      Width           =   1005
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
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   65
      Left            =   1365
      TabIndex        =   72
      Top             =   1635
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   67
      Left            =   1365
      TabIndex        =   71
      Top             =   1935
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Sales"
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
      Index           =   67
      Left            =   150
      TabIndex        =   70
      Top             =   1935
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
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
      Index           =   65
      Left            =   150
      TabIndex        =   69
      Top             =   1665
      Width           =   1035
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   61
      Left            =   4110
      TabIndex        =   68
      Top             =   1110
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Card No."
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
      Index           =   61
      Left            =   2925
      TabIndex        =   67
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of  VOR"
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
      Index           =   60
      Left            =   150
      TabIndex        =   66
      Top             =   1110
      Width           =   1050
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   60
      Left            =   1365
      TabIndex        =   65
      Top             =   1110
      Width           =   45
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   4
      Left            =   7995
      TabIndex        =   53
      Top             =   615
      Width           =   585
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   9165
      TabIndex        =   52
      Top             =   615
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Reg. No."
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
      Height          =   285
      Index           =   5
      Left            =   7980
      TabIndex        =   48
      Top             =   2055
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
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
      Index           =   3
      Left            =   150
      TabIndex        =   47
      Top             =   615
      Width           =   1245
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
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   9435
      TabIndex        =   39
      Top             =   2325
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Reg. Date"
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
      Height          =   285
      Index           =   6
      Left            =   7980
      TabIndex        =   38
      Top             =   2325
      Width           =   1335
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
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   88
      Left            =   9435
      TabIndex        =   37
      Top             =   2055
      Width           =   45
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
      Left            =   1470
      TabIndex        =   36
      Top             =   615
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   91
      Left            =   9165
      TabIndex        =   35
      Top             =   1140
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   92
      Left            =   9165
      TabIndex        =   34
      Top             =   1680
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   93
      Left            =   9165
      TabIndex        =   33
      Top             =   1410
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Type"
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
      Index           =   0
      Left            =   7995
      TabIndex        =   32
      Top             =   1410
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Sr. No."
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
      Index           =   1
      Left            =   7995
      TabIndex        =   31
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
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
      Index           =   2
      Left            =   7995
      TabIndex        =   30
      Top             =   1140
      Width           =   930
   End
End
Attribute VB_Name = "frmPurOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Const CellBackColLeave As String = &HEDF7FE
'Private Const CellForeColLeave As String = &H0&
'Private Const CellBackColEnter As String = &HFED7CF
'Private Const GridBackColorBkg As String = &HD7C6C8
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Dim RsParty As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim RsJCNo As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mVType As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function

Dim mCheckNegetiveStockSiteWise As Boolean

Dim FirmAddFlag As Byte
Dim GridKey As Integer
Dim DocID As String * 21
Dim OrderType As String
Dim LastNo As Integer
Dim StartNo As Integer
Dim txtKey As Boolean
Private Const TxtDocID As Byte = 0
Private Const OrderNo As Byte = 1
Private Const VDate As Byte = 2
Private Const PartyCode As Byte = 3
Private Const OrderRegNo As Byte = 4
Private Const OrderRegDt As Byte = 5
Private Const TypeVOR As Byte = 6
Private Const VType As Byte = 7
Private Const JobNo As Byte = 8
Private Const WHEELBASE As Byte = 9
Private Const EngineNo As Byte = 10
Private Const ChassisNo As Byte = 11
Private Const DateSale As Byte = 12
Private Const KMReading As Byte = 13
Private Const VehType As Byte = 25
Private Const SteerType As Byte = 14
Private Const SteerMake As Byte = 15
Private Const BreakType  As Byte = 16
Private Const Radiator As Byte = 17
Private Const ListRef As Byte = 18
Private Const DispMode As Byte = 19
Private Const Through As Byte = 20
Private Const ICName As Byte = 21
Private Const ICDesc As Byte = 22
Private Const Terms As Byte = 23
Private Const CancelOrd As Byte = 24

'FGrid  Col Declaration
Private Const Col_SrNo As Byte = 0
Private Const Col_PNo As Byte = 1
Private Const Col_Unit As Byte = 2
Private Const Col_MRP As Byte = 3
Private Const Col_Taxable As Byte = 4
Private Const Col_Qty As Byte = 5
Private Const Col_NDP As Byte = 6
Private Const Col_Amt As Byte = 7
Private Const Col_DiscPer As Byte = 8
Private Const Col_DiscAmt As Byte = 9
Private Const Col_FRate As Byte = 10
Private Const Col_Rate As Byte = 10 'Col_FRate=Col_Rate
Private Const Col_ItemVal As Byte = 11
Private Const Col_PName As Byte = 12
Private Const Col_LName As Byte = 13
Private Const Col_MRPStkTP As Byte = 14
Private Const Col_MRPStkTB As Byte = 15
Private Const Col_TBStk As Byte = 16
Private Const Col_TPStk As Byte = 17
Private Const Col_TBRate As Byte = 18
Private Const Col_TPRate As Byte = 19
Private Const Col_Bin As Byte = 20
Private Const Col_LastRate As Byte = 21
Private Const Col_HPRate As Byte = 22
Private Const Col_LPRate As Byte = 23
Private Const Col_PartGrade As Byte = 24
Private Const Col_EffectDate As Byte = 25
Private Const Col_MRPRate As Byte = 6

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Private Const FromVno As Byte = 0
Private Const ToVno As Byte = 1
Private Const VType1 As Byte = 2

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Private Sub Command1_Click()
Dim Rst As ADODB.Recordset
Dim I As Double, mQry$
Dim PrintStr$, CSVFName$
Dim fob As New FileSystemObject
CSVFName = Replace(txt(TxtDocID).TEXT, " ", "") & ".csv"
If fob.FileExists("C:\" & CSVFName) = False Then
    fob.CreateTextFile ("C:\" & CSVFName)
End If
Close #1
Open "C:\" & CSVFName For Output As #1

mQry = "SELECT Syctrl.SprPurOrdFooter,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
    "SPO.OrderID,SPO.Order_Type,SPO.Order_Prefix,SPO.Order_No,SPO.V_Date,SPO.Party_Code,SPO.VOR_Type,SPO.Job_Type,SPO.Job_DocID," & _
    "SPO.Vehicle_Type,SPO.Steer_Type,SPO.Steer_Make,SPO.Brake_Type,SPO.Radiator_Make,SPO.Order_Reg_No,SPO.Order_Reg_Dt," & _
    "SPO.Tot_Items,SPO.Tot_Qty,SPO.Disc_Per,SPO.Disc_Amt,SPO.Add_Charge_Per,SPO.Add_Charge,SPO.Tot_Amount,SPO.Spr_PriceList," & _
    "SPO.Dispatch_Mode,SPO.Through,SPO.Delivery_To,SPO.IC_Name,SPO.IC_Desig,SPO.Terms,SPO.Terms2,SPO.Printed,SPO.U_Name,SPO.U_EntDt," & _
    "SPO1.OrderId,SPO1.Srl_No,SPO1.PART_NO,Part.Part_Name,SPO1.QTY,SPO1.Rate,SPO1.Disc_Amt,SPO1.Amount " & _
"FROM (((SP_Order SPO LEFT JOIN SP_Order1 SPO1 ON SPO.OrderId = SPO1.OrderId) " & _
    "LEFT JOIN Part ON SPO1.PART_NO = Part.PART_NO and Part.Div_Code = left(SPO1.OrderID,1)) " & _
    "LEFT JOIN (SubGroup SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON SPO.Party_Code = SG.SubCode) " & _
    "LEFT JOIN Syctrl ON  Syctrl.LinkTable  >= SPO.U_AE " & _
"where SPO.OrderId = '" & Master!SearchCode & "'"
I = 1
Set Rst = GCn.Execute(mQry)
Do Until Rst.EOF
    Print #1, I & "," & Rst!Part_No & "," & Rst!Qty & ""
    Rst.MoveNext
    I = I + 1
Loop
Close #1
MsgBox "CSV File " & vbCrLf & CSVFName & vbCrLf & "Made!"
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

Dim I As Byte
'Dim RstMain As Recordset
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    Call Ini_Pub
    
'    PubVCompCode , PubSCompCode, PubWCompCode
    '** Hide Job Card Details if Only Spare Section is activated
    If PubWCompCode = "" Then
        Label3(24).Visible = False
        Shape1.Visible = False
        For I = 60 To 71
            Label3(I).Visible = False
            LblColon(I).Visible = False
        Next
        txt(6).Visible = False
        txt(25).Visible = False
        For I = 8 To 17
            txt(I).Visible = False
        Next
        Label3(3).top = 1110
        LblColon(90).top = 1110
        txt(3).top = 1110
    End If
    txt(VDate).Tag = PubLoginDate

    Set DGPart.DataSource = RsPart
    
    Dim SiteCond As String
    SiteCond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("sp_order.OrderId", "3", "1") & "='" & PubSiteCode & "'"
    End If


    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select sp_order.OrderId as searchcode from sp_order " & _
                    "where left(OrderId,1)='" & PubDivCode & "' " & SiteCond & " and left(order_type,4) = 'S_PO' order by Order_no desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("select Top 1 sp_order.OrderId as searchcode from sp_order " & _
                    "where left(OrderId,1)='" & PubDivCode & "' " & SiteCond & "  and left(order_type,4) = 'S_PO' order by Order_no desc")
    End If
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct Order_No as code from SP_Order where left(OrderId,1)='" & PubDivCode & "' order by order_No", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno

    Set RsJCNo = New ADODB.Recordset
    With RsJCNo
        .CursorLocation = adUseClient
        .Open "SELECT DocId as Code,str(Job_No) as Name,AtKMsHrs from Job_Card where left(docid,1)='" & PubDivCode & "' and (JobcloseDate is null or JobCloseDate<= " & ConvertDate(PubLoginDate) & ") order by Job_No", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGJCNo.DataSource = RsJCNo
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select Subcode as code,NAME,Party_Type from SubGroup Where firmCode = '" & PubFirmCode & "' and Nature='Supplier' Order by name", GCn, adOpenDynamic, adLockOptimistic
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        "order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsParty = Nothing
Set RsJCNo = Nothing
Set RsVno = Nothing
Set Master = Nothing
End Sub


Private Sub ListView_Click()
If FrmPrn.Visible = False Then
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    txt(Val(ListView.Tag)).SetFocus
    FrmList.Visible = False
Else
    txtPrint(VType1).TEXT = ListView.SelectedItem.TEXT
    txtPrint(VType1).SetFocus
    FrmList.Visible = False
End If
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    txt(TxtDocID).Enabled = False
    mPartyType = 0
    txt(VDate) = txt(VDate).Tag
    txt(VDate).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim mTrans As Boolean
If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    If GCn.Execute("select count(*) from sp_stock where Order_DocId = '" & Master!SearchCode & "'").Fields(0).Value > 0 Then
         MsgBox "Dispatch Challan Exists of this Purchase Order, " & vbCrLf & "Can't Delete the Reocord", vbInformation, "Validation"
         Exit Sub
    End If
    GCn.BeginTrans
    mTrans = True
    GCn.Execute ("delete from sp_order where OrderId = '" & Master!SearchCode & "'")
    GCn.Execute ("delete from sp_order1 where OrderId = '" & Master!SearchCode & "'")
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    BUTTONS True, Me, Master, 0
    Call MoveRec
End If
eloop1:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
     If GCn.Execute("select count(*) from sp_stock where Order_DocId = '" & Master!SearchCode & "'").Fields(0).Value > 0 Then
          MsgBox "Dispatch Challan Exists of this Purchase Order, " & vbCrLf & "Can't Edit the Reocord", vbInformation, "Validation"
          Exit Sub
    End If
    Disp_Text SETS("EDIT", Me, Master)
    txt(PartyCode).SetFocus

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

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    Dim SiteCond As String
    SiteCond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("sp_order.OrderId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "SELECT sp_order.OrderId as searchcode,sp_order.OrderId, sp_order.Order_Type, sp_order.Order_Prefix, " & cCStr("Sp_Order.Order_No") & " As Order_No, sp_order.Site_Code, " & cDt("sp_order.V_Date") & " AS VoucherDate, SubGroup.Name as PartyName FROM sp_order LEFT JOIN SubGroup ON sp_order.Party_Code = SubGroup.Subcode where left(OrderId,1)='" & PubDivCode & "' " & SiteCond & " and left(sp_order.order_type,4) = 'S_PO' order by sp_order.OrderId"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("select sp_order.OrderId as searchcode from sp_order " & _
                    "where left(OrderId,1)='" & PubDivCode & "' and left(order_type,4) = 'S_PO' And sp_order.OrderId = '" & MyValue & "' order by Order_no desc")
    
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
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
    MsgBox err.Description, vbInformation, "Information": Exit Sub
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

Private Sub TopCtrl1_eRef()
    RsParty.Requery
    RsPart.Requery
    RsJCNo.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, mGridFilled As Boolean
Dim DocIdHlp As String
On Error GoTo errlbl
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    
    If IsValid(txt(VDate), "Order Date") = False Then Exit Sub
    If IsValid(txt(VType), "Order Type") = False Then Exit Sub
    If IsValid(txt(OrderNo), "Order Serial Number") = False Then Exit Sub
    If IsValid(txt(PartyCode), "Supplier Name") = False Then Exit Sub
   
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Fill Taxable Yes/No in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
            If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Fill MRP Yes/No in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
            If Val(FGrid.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Fill Quantity in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_PNo: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter

    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If GCn.Execute("select count(*) from sp_order where orderid='" & DocID & "'").Fields(0) > 0 Then MsgBox "Duplicate Order No.", vbCritical, "Validation Error": Exit Sub
    Else
        DocID = Master!SearchCode
    End If
    
    DocIdHlp = Replace(DocID, " ", "")
    RemoveTxtNull
    GCn.BeginTrans
    mTrans = True
 
    If TopCtrl1.TopText2.CAPTION = "Add" Then
'        GCn.Execute ("delete from sp_order where orderid='" & DocId & "'")
        GCn.Execute "insert into sp_order(OrderId , OrderIDHelp, Order_Type, Order_Prefix,Order_No, Site_Code, V_Date, Party_Code,Order_Reg_No , Order_Reg_Dt, Tot_Items, Tot_Qty, Tot_Amount, U_Name, U_EntDt, U_AE, VOR_TYPE , JOB_TYPE , JOB_DocID ,Vehicle_Type ,Steer_Type , Steer_Make, Brake_Type , Radiator_Make, Spr_PriceList , Dispatch_Mode, Through , IC_Name , IC_Desig , Terms ,Cancel_RestOrders) " & _
            " values('" & DocID & "','" & DocIdHlp & "','" & OrderType & "','" & OrderType & "'," & Val(txt(OrderNo).TEXT) & ",'" & PubSiteCode & PubSiteCode & "'," & ConvertDate(txt(VDate).TEXT) & ",'" & txt(PartyCode).Tag & "', " & _
            " '" & txt(OrderRegNo).TEXT & "'," & ConvertDate(txt(OrderRegDt).TEXT) & "," & Val(LblIVal.CAPTION) & "," & Val(LblQty.CAPTION) & "," & Val(LblAmt.CAPTION) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & left(txt(TypeVOR).TEXT, 1) & "', ' ', '" & txt(JobNo).TEXT & "' ,'" & txt(VehType).TEXT & "' , '" & txt(SteerType).TEXT & "', '" & txt(SteerMake).TEXT & "',  '" & txt(BreakType).TEXT & "', '" & txt(Radiator).TEXT & "', '" & txt(ListRef).TEXT & "', '" & txt(DispMode).TEXT & "','" & txt(Through).TEXT & "', '" & txt(ICName).TEXT & "', '" & txt(ICDesc).TEXT & "', '" & txt(Terms).TEXT & "', '" & IIf(txt(CancelOrd).TEXT = "Yes", 1, 0) & "')"
    Else
        GCn.Execute ("update sp_order set Party_Code='" & txt(PartyCode).Tag & "',v_date = " & ConvertDate(txt(VDate).TEXT) & ",Order_Reg_No='" & txt(OrderRegNo).TEXT & "' , Order_Reg_Dt=" & ConvertDate(txt(OrderRegDt).TEXT) & ", Tot_Items=" & Val(LblIVal.CAPTION) & ",  " & _
            " Tot_Qty=" & Val(LblQty.CAPTION) & ", Tot_Amount=" & Val(LblAmt.CAPTION) & ", U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E', VOR_TYPE ='" & left(txt(TypeVOR).TEXT, 1) & "', JOB_TYPE = ' ', JOB_DocID = '" & txt(JobNo).Tag & "' ,Vehicle_Type = '" & txt(VehType).TEXT & "' ,Steer_Type = '" & txt(SteerType).TEXT & "', Steer_Make='" & txt(SteerMake).TEXT & "', Brake_Type = '" & txt(BreakType).TEXT & "', Radiator_Make = '" & txt(Radiator).TEXT & "',  Spr_PriceList = '" & txt(ListRef).TEXT & "', Dispatch_Mode = '" & txt(DispMode).TEXT & "', Through = '" & txt(Through).TEXT & "', IC_Name ='" & txt(ICName).TEXT & "', IC_Desig ='" & txt(ICDesc).TEXT & "', Terms = '" & txt(Terms).TEXT & "',Cancel_RestOrders = '" & IIf(txt(CancelOrd).TEXT = "Yes", 1, 0) & "' where orderid = '" & DocID & "'")
    End If
    GCn.Execute ("delete from sp_order1 where orderid='" & DocID & "'")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" And Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
            GCn.Execute ("insert into sp_order1(OrderId , Srl_No, Order_Type, " & _
                " Site_Code, V_Date, Party_Code, Order_Reg_No, Order_Reg_Dt, Part_No,TAX_YN, " & _
                " MRP_YN,QTY, Rate, NDP, Disc_Per, Disc_Amt, Amount, U_Name, U_EntDt, U_AE) " & _
                " values('" & DocID & "'," & I & ",'" & OrderType & "','" & PubSiteCode & PubSiteCode & "'," & ConvertDate(txt(VDate).TEXT) & ",'" & txt(PartyCode).Tag & "','" & txt(OrderRegNo).TEXT & "'," & ConvertDate(txt(OrderRegDt).TEXT) & ",'" & FGrid.TextMatrix(I, Col_PNo) & "', " & _
                "" & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & " , " & _
                "" & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & Val(FGrid.TextMatrix(I, Col_FRate)) & "," & Val(FGrid.TextMatrix(I, Col_NDP)) & ",  " & _
                "" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & ", " & _
                " " & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
        End If
    Next
GCn.CommitTrans
mTrans = False
    If TopCtrl1.TopText2.CAPTION = "Add" Then txt(VDate).Tag = txt(VDate)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select sp_order.OrderId as searchcode from sp_order " & _
                    "where left(OrderId,1)='" & PubDivCode & "' and left(order_type,4) = 'S_PO' And sp_order.OrderId = '" & DocID & "' order by Order_no desc")
    End If
    Master.FIND "SearchCode = '" & DocID & "'"
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
If txt(VType).TEXT = "" And Index <> VDate Then txt(VType).SetFocus
TxtGrid(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case VType
        ListArray = Array("Annual", "Quarterly", "Monthly", "General(Casual)", "VOR")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 5)
    Case BreakType
        ListArray = Array("DAOH", "S-CAM", "OTHER")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
    Case PartyCode
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case TypeVOR
        ListArray = Array("Workshop", "VOR")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    Case OrderNo
        If IsValid(txt(VType), "Order Type") = False Then Exit Sub
    Case JobNo
        Set DGJCNo.DataSource = RsJCNo
        If RsJCNo.RecordCount = 0 Or (RsJCNo.EOF = True Or RsJCNo.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsJCNo!Name Then
            RsJCNo.MoveFirst
            RsJCNo.FIND "code ='" & txt(Index).Tag & "'"
        End If
    
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case VType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1500
    Case TypeVOR
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case BreakType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 900
    Case PartyCode
        DGridTxtKeyDown DGParty, txt, PartyCode, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    Case JobNo
        DGridTxtKeyDown DGJCNo, txt, JobNo, RsJCNo, KeyCode, False, 1
End Select
If FrmList.Visible = False And DGParty.Visible = False And DGJCNo.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Terms And Index <> CancelOrd Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = CancelOrd Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> Terms And Index <> VDate Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Terms And Index <> PartyCode Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case Index
Case PartyCode
    If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "Name"
    lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.BackColor = vbBlack: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
Case JobNo
    If DGJCNo.Visible = True Then DGridTxtKeyPress txt, Index, RsJCNo, KeyAscii, "Name"
Case OrderNo
    Call NumPress(txt(Index), KeyAscii, 6, 0)
Case KMReading
    Call NumPress(txt(Index), KeyAscii, 7, 0)
Case CancelOrd
     If UCase(Chr(KeyAscii)) = "N" Then
         txt(Index) = "No"
     ElseIf UCase(Chr(KeyAscii)) = "Y" Then
         txt(Index) = "Yes"
     Else
         txt(Index) = ""
     End If
     KeyAscii = 0
     
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case VType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case TypeVOR
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case BreakType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Dim rsPOCounter As ADODB.Recordset
Select Case Index
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
           txt(VDate).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        Cancel = Not CheckFinYear(txt(Index))
        If Cancel = False Then
            If Len(txt(OrderRegDt).TEXT) > 0 Then
                If Len(txt(VDate).TEXT) > 0 Then
                    txt(VDate).TEXT = RetDate(txt(Index))
                    If CDate(Format(txt(VDate).TEXT, "dd/mm/yyyy")) > CDate(Format(txt(OrderRegDt).TEXT, "dd/mm/yyyy")) Then
                        MsgBox "Order Registration Date is less than Order date", vbOKOnly, "Validation Check"
                        Cancel = True
                    End If
                End If
            End If
        End If
    Case VType
        If IsValid(txt(VType), "Order Type") = False Then Cancel = True:   Exit Sub
        txt(VType).TEXT = ListView.SelectedItem.TEXT
        'PO No. generation
        Set rsPOCounter = New Recordset
        rsPOCounter.CursorLocation = adUseClient
        rsPOCounter.Open "select * FROM SP_OrdCoun WHERE Div_Code='" & PubDivCode & "' and DETAILS= '" & txt(VType) & "'", GCn, adOpenStatic, adLockReadOnly
        LastNo = IIf(IsNull(rsPOCounter!end_no), 0, rsPOCounter!end_no)
        StartNo = IIf(IsNull(rsPOCounter!start_no), 0, rsPOCounter!start_no)
        DocID = PubDivCode & PubSiteCode & PubSiteCode & Space(5 - Len(Trim(rsPOCounter!ord_type))) + Trim(rsPOCounter!ord_type) + Space(5 - Len(Trim(rsPOCounter!Prefix))) + Trim(rsPOCounter!Prefix)
        If GCn.Execute("select count(*) from sp_order where left(orderId,13) = '" & DocID & "'").Fields(0).Value > 0 Then
            txt(OrderNo) = GCn.Execute("select max(Order_No) from sp_order where left(orderid,1) = '" & PubDivCode & "' and left(orderid,13) = '" & DocID & "'").Fields(0).Value + 1
            txt(OrderRegNo) = txt(OrderNo)
        Else
            txt(OrderNo) = rsPOCounter!start_no
            txt(OrderRegNo) = txt(OrderNo)
        End If
        DocID = PubDivCode & PubSiteCode & PubSiteCode & Space(5 - Len(Trim(rsPOCounter!ord_type))) + Trim(rsPOCounter!ord_type) + Space(5 - Len(Trim(rsPOCounter!Prefix))) + Trim(rsPOCounter!Prefix) + Space(8 - Len(Trim(txt(OrderNo)))) + Trim(txt(OrderNo))
        txt(TxtDocID).TEXT = DocID
        OrderType = rsPOCounter!ord_type
        LblVPrefix.CAPTION = OrderType
        Set rsPOCounter = Nothing
        
        If txt(Index).TEXT = "VOR" Then
            FldEnabled True
        Else
            FldEnabled False
        End If
        txt(VType).Tag = txt(VType).TEXT
    Case OrderNo
        If IsValid(txt(Index), "Order No") = False Then Cancel = True:   Exit Sub
        If LastNo = 0 Then
            If Val(txt(Index).TEXT) < StartNo Then
                MsgBox "Invalid Serial No", vbInformation, "Validation Check": Cancel = True: Exit Sub
            End If
        Else
            If Val(txt(Index).TEXT) < StartNo Or Val(txt(Index).TEXT) > LastNo Then
                MsgBox "Invalid Serial No", vbInformation, "Validation Check": Cancel = True: Exit Sub
            End If
        End If
        DocID = left(DocID, 13) + Space(8 - Len(Trim(txt(OrderNo).TEXT))) + Trim(txt(OrderNo).TEXT)    'DivCode(1)+SiteCode(2)+vtype(5)+prefix(5)+no(8)
        txt(TxtDocID).TEXT = DocID
        If GCn.Execute("select COUNT(*) from sp_order where orderid = '" & DocID & "'").Fields(0).Value > 0 Then
            MsgBox "Duplicate Order No ", vbInformation, "Validation Check": Cancel = True: Exit Sub
        End If
    Case PartyCode
        If IsValid(txt(Index), "Party") = False Then Cancel = True: Exit Sub
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            mPartyType = 0
        Else
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
            mPartyType = RsParty!Party_Type
        End If
    Case BreakType, TypeVOR
        If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case JobNo
        If RsJCNo.RecordCount = 0 Or (RsJCNo.EOF = True Or RsJCNo.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsJCNo!Name
            txt(Index).Tag = RsJCNo!Code
            Dim RsCard As ADODB.Recordset
            Set RsCard = New ADODB.Recordset
            With RsCard
                .CursorLocation = adUseClient
                .Open "SELECT HisCard.Chassis, HisCard.engine, HisCard.Supplier_BillDate, Model.WHEELBASE, Model.Wheel_Catg " & _
                    " FROM (HisCard LEFT JOIN Model ON HisCard.Model = Model.MODEL) RIGHT JOIN Job_Card ON HisCard.CardNo = Job_Card.CardNo  where job_card.DocId = '" & RsJCNo!Code & "' ", GCn, adOpenDynamic, adLockOptimistic
            End With
            If RsCard.RecordCount > 0 Then
                txt(EngineNo).TEXT = IIf(IsNull(RsCard!Engine), "", RsCard!Engine)
                txt(ChassisNo).TEXT = IIf(IsNull(RsCard!Chassis), "", RsCard!Chassis)
                txt(DateSale).TEXT = IIf(IsNull(RsCard!Supplier_BillDate), "", RsCard!Supplier_BillDate)
                txt(WHEELBASE).TEXT = IIf(IsNull(RsCard!WHEELBASE), "", RsCard!WHEELBASE)
                txt(KMReading).TEXT = IIf(IsNull(RsCard!Wheel_Catg), "", RsCard!Wheel_Catg)
            End If
            Set RsCard = Nothing
        End If
    Case OrderRegDt
        txt(OrderRegDt).TEXT = Trim(txt(OrderRegDt).TEXT)
        If Len(txt(OrderRegDt).TEXT) > 0 Then
            If Len(txt(VDate).TEXT) > 0 Then
                txt(OrderRegDt).TEXT = RetDate(txt(Index))
                If CDate(Format(txt(VDate).TEXT, "dd/mm/yyyy")) > CDate(Format(txt(OrderRegDt).TEXT, "dd/mm/yyyy")) Then
                    MsgBox "Order Registration Date is less than Order date", vbOKOnly, "Validation Check"
                    Cancel = True
                End If
            End If
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGPart_Click()
On Error GoTo ELoop
If RsPart.RecordCount > 0 Then
    Select Case FGrid.Col
        Case Col_PNo
            TxtGrid(0).TEXT = RsPart!Code
        Case Col_PName
            TxtGrid(0) = RsPart!Name
        Case Col_LName
            TxtGrid(0) = RsPart!LName
    End Select
End If
    TxtGridValid_PNo
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGPart.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGJCNo_Click()
    If RsJCNo.RecordCount > 0 Then
        txt(JobNo).TEXT = RsJCNo!Name
        txt(JobNo).Tag = RsJCNo!Code
    End If
    txt(JobNo).SetFocus
    DGJCNo.Visible = False
End Sub

Private Sub DGVno_Click()
Dim Index As Integer
If DGVno.Tag = "1" Then
    Index = ToVno
Else
    Index = FromVno
End If
    If RsVno.RecordCount > 0 Then
        txtPrint(Index).TEXT = RsVno!Code
    End If
    txtPrint(Index).SetFocus
    DGVno.Visible = False
End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(PartyCode).TEXT = RsParty!Name
        txt(PartyCode).Tag = RsParty!Code
        mPartyType = RsParty!Party_Type
    End If
    txt(PartyCode).SetFocus
    DGParty.Visible = False
    lblGroup.Visible = False
End Sub

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If txt(VType).TEXT = "" Then txt(VType).SetFocus: Exit Sub
Select Case FGrid.Col
    Case Col_PNo, Col_PName, Col_LName
        Call GridDblClick(Me, FGrid, TxtGrid, 0)
    Case Col_Taxable, Col_MRP, Col_Qty, Col_NDP, Col_DiscPer, Col_DiscAmt
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
        End If
End Select
TAddMode = False
End Sub

Private Sub FGrid_EnterCell()
'FGrid.CellBackColor = CellBackColEnter
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
                Col_MRPStkTB, Col_MRPStkTP, Col_TBStk, Col_TPStk, _
                Col_MRPRate, Col_TBRate, Col_TPRate, Col_Bin, _
                Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
'        End If
        FrmDetail.Visible = True
    End If
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
'If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If txt(VType).TEXT = "" Then txt(VType).SetFocus: Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
'    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
'    FGrid.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case Col_MRP, Col_Taxable
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case Col_Qty, Col_NDP, Col_DiscPer, Col_DiscAmt
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "" '"0.00"
    End Select
    Amt_Cal1
    Amt_Cal
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Col_PNo, Col_PName, Col_LName
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
        Case Col_Taxable, Col_MRP, Col_Qty, Col_NDP, Col_DiscPer, Col_DiscAmt
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Call GridDblClick(Me, FGrid, TxtGrid, 0)
                TAddMode = False
            End If
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
Select Case FGrid.Col
    Case Col_PNo, Col_PName, Col_LName
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case Col_Unit, Col_Amt, Col_FRate, Col_ItemVal
        FGrid_LeaveCell
        FGrid.Col = FGrid.Col + 1
        FGrid_EnterCell
        FGrid.SetFocus
    Case Col_Taxable, Col_MRP
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
        End If
    Case Col_Qty, Col_NDP, Col_DiscPer, Col_DiscAmt
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
           Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
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
            FGrid.TextMatrix(I, 0) = I
         Next
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
        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB) = GetMrpTBStk(FGrid.TextMatrix(FGrid.Row, Col_PNo))
        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP) = GetMrpTPStk(FGrid.TextMatrix(FGrid.Row, Col_PNo))
        FGrid.TextMatrix(FGrid.Row, Col_TBStk) = GetTBStk(FGrid.TextMatrix(FGrid.Row, Col_PNo))
        FGrid.TextMatrix(FGrid.Row, Col_TPStk) = GetTPStk(FGrid.TextMatrix(FGrid.Row, Col_PNo))
    
    
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
'FGrid.CellBackColor = CellBackColLeave
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub

Private Sub MoveRec()
Dim Master1 As Recordset
Dim Rst As Recordset, I As Integer
Dim mVor As String
On Error GoTo error1
If Master.RecordCount > 0 Then
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select sp_order.* from sp_order where orderID='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
    
    LblDiv.CAPTION = "Division : " & left(Master1!OrderID, 1)
    LblSite.CAPTION = "Site Code : " & Master1!Site_Code
    LblVPrefix.CAPTION = mID(Master1!OrderID, 8, 5)
    txt(TxtDocID).TEXT = Master1!OrderID
    txt(OrderNo).TEXT = Master1!Order_NO
    txt(VDate).TEXT = Master1!V_Date
    Select Case Right(Master1!Order_Type, 1)
        Case "A"
            txt(VType).TEXT = "Annual"
        Case "Q"
            txt(VType).TEXT = "Quarterly"
        Case "M"
            txt(VType).TEXT = "Monthly"
        Case "G"
            txt(VType).TEXT = "General(Casual)"
        Case "V"
            txt(VType).TEXT = "VOR"
    End Select
    OrderType = Right(Master1!Order_Type, 1)
    txt(PartyCode).Tag = Master1!Party_code
    txt(PartyCode).TEXT = GCn.Execute("select name from SubGroup where Subcode = '" & Master1!Party_code & "'").Fields(0).Value
    txt(OrderRegNo).TEXT = IIf(IsNull(Master1!Order_Reg_No), "", Master1!Order_Reg_No)
    txt(OrderRegDt).TEXT = IIf(IsNull(Master1!Order_Reg_Dt), "", Master1!Order_Reg_Dt)
    
    mVor = IIf(IsNull(Master1!VOR_TYPE), "", Master1!VOR_TYPE)
    Select Case mVor
        Case "W"
            txt(TypeVOR).TEXT = "WorkShop"
        Case "V"
            txt(TypeVOR).TEXT = "VOR"
    End Select

'    txt(JobType).Text = IIf(IsNull(Master1!JOB_TYPE), "", Master1!JOB_TYPE)
    txt(JobNo).TEXT = IIf(IsNull(Master1!job_docid), "", Right(Master1!job_docid, 8))
    txt(JobNo).Tag = IIf(IsNull(Master1!job_docid), "", Master1!job_docid)
    txt(VehType).TEXT = IIf(IsNull(Master1!Vehicle_Type), "", Master1!Vehicle_Type)
    txt(SteerType).TEXT = IIf(IsNull(Master1!Steer_Type), "", Master1!Steer_Type)
    txt(SteerMake).TEXT = IIf(IsNull(Master1!Steer_Make), "", Master1!Steer_Make)
    txt(BreakType).TEXT = IIf(IsNull(Master1!Brake_Type), "", Master1!Brake_Type)
    txt(Radiator).TEXT = IIf(IsNull(Master1!Radiator_Make), "", Master1!Radiator_Make)
    txt(ListRef).TEXT = IIf(IsNull(Master1!Spr_PriceList), "", Master1!Spr_PriceList)
    txt(DispMode).TEXT = IIf(IsNull(Master1!Dispatch_Mode), "", Master1!Dispatch_Mode)
    txt(Through).TEXT = IIf(IsNull(Master1!Through), "", Master1!Through)
    txt(ICName).TEXT = IIf(IsNull(Master1!IC_Name), "", Master1!IC_Name)
    txt(ICDesc).TEXT = IIf(IsNull(Master1!IC_Desig), "", Master1!IC_Desig)
    txt(Terms).TEXT = IIf(IsNull(Master1!Terms), "", Master1!Terms)
    txt(CancelOrd).TEXT = IIf(Master1!Cancel_RestOrders = 1, "Yes", "No")
    
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "SELECT P.Part_Name, P.Local_Name,P.Part_Grade,p.MRP_Effect_Dt,P.TB_Effect_Dt,P.unit,P.Cur_MRP_TBStk, P.Cur_MRP_TPStk,P.Cur_TB_Stk,P.Cur_TP_Stk, P.MRP,P.TP_SRate, P.TB_SRate, P.Bin_Loca,P.high_pur_rate, P.low_pur_rate, SP_Order1.* FROM SP_Order1 LEFT JOIN Part P ON SP_Order1.PART_NO = P.PART_NO and P.Div_Code = left(SP_Order1.OrderID,1) where SP_Order1.OrderId = '" & Master1!OrderID & "' order by SP_Order1.OrderId,SP_Order1.Srl_No", GCn, adOpenStatic, adLockReadOnly
    FGrid.Redraw = False
    FGrid.Rows = 1
    If Rst.RecordCount > 0 Then
        I = 1
        Do Until Rst.EOF
'            FGrid.AddItem rs!Srl_No & Chr(9) & rs!Part_No & Chr(9) & rs!Unit & Chr(9) & IIf(rs!Tax_YN = 0, "No", "Yes") & Chr(9) & IIf(rs!MRP_YN = 0, "No", "Yes") & Chr(9) & Format(rs!Qty, "0.000") & Chr(9) & Format(rs!Rate, "0.00") & Chr(9) & Format((rs!Qty * rs!NDP), "0.00") & Chr(9) & Format(rs!Disc_Per, "0.00") & Chr(9) & Format(rs!Disc_Amt, "0.00") & Chr(9) & Format(rs!AMOUNT, "0.00") _
            & Chr(9) & Format(rs!NDP, "0.00") & Chr(9) & rs!Part_Name & Chr(9) & rs!Local_Name & Chr(9) & IIf(IsNull(rs!Curstk), 0, rs!Curstk) & Chr(9) & IIf(IsNull(rs!MRPQty), 0, rs!MRPQty) & Chr(9) & IIf(IsNull(rs!Cur_TB_Stk), 0, rs!Cur_TB_Stk) & Chr(9) & IIf(IsNull(rs!Cur_TP_Stk), 0, rs!Cur_TP_Stk) & Chr(9) & IIf(IsNull(rs!TB_SRate), 0, rs!TB_SRate) & Chr(9) & IIf(IsNull(rs!TP_SRate), 0, rs!TP_SRate) & Chr(9) & IIf(IsNull(rs!Bin_Loca), 0, rs!Bin_Loca) & Chr(9) & " " & Chr(9) & IIf(IsNull(rs!high_pur_rate), 0, rs!high_pur_rate) & Chr(9) & IIf(IsNull(rs!low_pur_rate), 0, rs!low_pur_rate)
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, Col_SrNo) = Rst!Srl_No
                .TextMatrix(I, Col_PNo) = Rst!Part_No
                .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, Col_Qty) = Format(Rst!Qty, "0.000")
                .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                .TextMatrix(I, Col_Amt) = Format((Rst!Qty * Rst!NDP), "0.00")
'                    .TextMatrix(i, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")
'                    If Rst!MRP_YN = 1 Then
'                        .TextMatrix(i, Col_Amt) = Format((Rst!Qty * Rst!MRP_Rate), "0.00")
'                    Else
'                        .TextMatrix(i, Col_Amt) = Format((Rst!Qty * Rst!Rate), "0.00")
'                    End If
                .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                .TextMatrix(I, Col_ItemVal) = Format(Rst!Amount, "0.00")
                .TextMatrix(I, Col_NDP) = Format(Rst!NDP, "0.00")
                .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)

'Private Const FRate As Byte = 10



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
            Rst.MoveNext
            I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    FGrid.Redraw = True
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
Set Rst = Nothing
Set Master1 = Nothing
Grid_Hide
Amt_Cal1
Amt_Cal

Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte
'0|Part No.1|Part Name 2|Unit 3|Taxable 4|MRP Y/N 5|Quantity 6|NDP 7|Amount 8|Dis % 9 |Dis Rs 10|Item Value 11|Loal Name    12|Curr Stk Qty 13|MRP Qty 14|Taxable Qty 15|TaxPaid Qty 16|Taxable Rate 17|TaxPaid Rate 18|Bin Location 19|Last Purch Rate 20|High Purch Rate 21|Low Purch Rate 22
'FGrid.left = 60: FGrid.width = 11775

    With FGrid
        .left = Me.left '+ 45
        .width = Me.width - 90
        .top = 2655
        .Cols = 27
        .RowHeightMin = PubGridRowHeight
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, Col_PNo) = "Part No"
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

        .TextMatrix(0, Col_Qty) = "Quantity"
        .ColAlignmentFixed(Col_Qty) = flexAlignRightCenter
        .ColWidth(Col_Qty) = 960

       
        .TextMatrix(0, Col_NDP) = "Rate"
        .ColAlignmentFixed(Col_NDP) = flexAlignRightCenter
        .ColWidth(Col_NDP) = 870

        .TextMatrix(0, Col_FRate) = "NDP"
        .ColAlignmentFixed(Col_FRate) = flexAlignRightCenter
        .ColWidth(Col_FRate) = 870

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
        .ColAlignmentFixed(Col_LName) = flexAlignLeftCenter
        .ColWidth(Col_LName) = 2000
    End With
    For I = 14 To 25
        FGrid.ColWidth(I) = 0
    Next
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    
    FGrid.ColAlignment(13) = flexAlignLeftCenter
    
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height ': DGPart.height = (Me.height - DGPart.top) - 90  '  2350
    FrmDetail.width = 6285: FrmDetail.left = Me.width - (FrmDetail.width + mRtScale): FrmDetail.top = 405: FrmDetail.height = 2130
    FrmPrn.left = (Me.width - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
    DGVno.left = 5145: DGVno.top = mTopScale
    DGParty.width = 5130:   DGParty.left = 6700
    DGParty.top = mTopScale '390
    DGParty.height = 4935
    DGJCNo.left = txt(JobNo).left + txt(JobNo).width
    DGJCNo.top = txt(JobNo).top
    With DGPart
        .Columns(6).width = 2564.788
        .Columns(5).width = 1005.165
        .Columns(4).width = 1005.165
        .Columns(3).width = 1005.165
        .Columns(2).width = 494.9292
        .Columns(1).width = 3225.26
        .Columns(0).width = 1950.236
    End With
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next
txt(TxtDocID).Enabled = False
If TopCtrl1.TopText2 = "Edit" Then
    txt(VDate).Enabled = False
    txt(VType).Enabled = False
    txt(OrderNo).Enabled = False
    txt(PartyCode).SetFocus
    If txt(VType).TEXT = "VOR" Then
        FldEnabled True
    Else
        FldEnabled False
    End If
End If

txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
    If DGJCNo.Visible = True Then DGJCNo.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub
 Private Sub Amt_Cal()
 Dim I As Double
 Dim IQty As Double
 Dim ICnt As Integer
 Dim IAmt As Double
 For I = 1 To FGrid.Rows - 1
    If FGrid.TextMatrix(I, Col_PNo) <> "" Then
        IQty = IQty + Val(FGrid.TextMatrix(I, Col_Qty))
        IAmt = IAmt + Val(FGrid.TextMatrix(I, Col_ItemVal))
        ICnt = ICnt + 1
    End If
Next I
    LblIVal.CAPTION = Format(ICnt, "0")
    LblQty.CAPTION = Format(IQty, "0.000")
    LblAmt.CAPTION = Format(IAmt, "0.00")
 End Sub

Private Sub Amt_Cal1()
      FGrid.TextMatrix(FGrid.Row, Col_FRate) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_NDP)) - Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))), "0.00")
      FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_NDP)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty)), "0.00")
      FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_FRate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    If FrmDetail.Visible = False Then FrmDetail.Visible = True
'    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
         Case Col_PNo
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Col_PName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "name"
            If FGrid.TextMatrix(FGrid.Row, Col_PName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Col_PName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Col_LName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "LName"
            If FGrid.TextMatrix(FGrid.Row, Col_LName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "lname ='" & FGrid.TextMatrix(FGrid.Row, Col_LName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        FGrid.SetFocus
        TxtGrid(0).Visible = False
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Col_PNo    '1
            'If DGPart.Visible = False Then DGridColSwap DGPart, 0
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
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
        Case Col_LName   '3
            If DGPart.Visible = False Then DGridColSwap DGPart, 2
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                End If
            End If
        Case Col_Taxable, Col_MRP, Col_Qty, Col_NDP, Col_DiscPer, Col_DiscAmt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_DiscAmt ', 3
                End If
                If FGrid.Col - 1 = Col_MRP Then
                    TxtGridValid_PNo
                End If
               
            End If
        
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case Col_PNo
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "CODE"
    Case Col_PName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "name"
    Case Col_LName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "Lname"
    Case Col_NDP, Col_Amt, Col_DiscAmt, Col_ItemVal
        Call NumPress(TxtGrid(Index), KeyAscii, 8, 2)
    Case Col_DiscPer
        NumPress TxtGrid(Index), KeyAscii, 2, 2
    Case Col_Qty
        Call NumPress(TxtGrid(Index), KeyAscii, 8, 3)
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case Col_PNo
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "CODE", True
    Case Col_PName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "name", True
    Case Col_LName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Lname", True
    Case Col_MRP, Col_Taxable
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            TxtGrid(Index) = ""
        ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
            TxtGrid(Index) = "Yes"
        Else
            TxtGrid(Index) = "No"
        End If
    Case Col_Qty
        FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0.000")
    Case Col_NDP
        FGrid.TextMatrix(FGrid.Row, Col_NDP) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(Val(TxtGrid(Index).TEXT) * Val(FGrid.TextMatrix(FGrid.Row, Col_DiscPer)) / 100, "0.00")
    Case Col_DiscPer
        FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = TxtGrid(Index).TEXT
        FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_NDP)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
    Case Col_DiscAmt
        FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        If Val(FGrid.TextMatrix(FGrid.Row, Col_NDP)) = 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = ""
        Else
           FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format((100 * Val(TxtGrid(Index).TEXT)) / Val(FGrid.TextMatrix(FGrid.Row, Col_NDP)), "0.00")
        End If
End Select
Amt_Cal1
Amt_Cal
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim j As Integer
    Select Case FGrid.Col
        Case Col_PNo, Col_PName, Col_LName
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_PNo
        Case Col_Taxable, Col_MRP
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        Case Col_NDP
            FGrid.TextMatrix(FGrid.Row, Col_NDP) = Format(Val(TxtGrid(0).TEXT), "0.00")
        Case Col_DiscAmt
            FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(Val(TxtGrid(0).TEXT), "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt)) > Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) Then
                TxtGridLeave = False: Exit Function
            End If
        Case Col_DiscPer
            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(Val(TxtGrid(0).TEXT), "0.00")
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(0).TEXT), "0.000")
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
Dim I As Integer
Dim X As String, Y As String
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
            If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Sub FldEnabled(Enb As Boolean)
If Enb = True Then
    txt(TypeVOR).Enabled = True
    txt(JobNo).Enabled = True
    txt(WHEELBASE).Enabled = False
    txt(EngineNo).Enabled = False
    txt(ChassisNo).Enabled = False
    txt(DateSale).Enabled = False
    txt(KMReading).Enabled = False
    txt(SteerType).Enabled = True
    txt(VehType).Enabled = True
    txt(SteerMake).Enabled = True
    txt(BreakType).Enabled = True
    txt(Radiator).Enabled = True
Else
    txt(TypeVOR).Enabled = False: txt(TypeVOR).TEXT = ""
    txt(JobNo).Enabled = False: txt(JobNo).TEXT = ""
    txt(WHEELBASE).Enabled = False: txt(WHEELBASE).TEXT = ""
    txt(EngineNo).Enabled = False: txt(EngineNo).TEXT = ""
    txt(ChassisNo).Enabled = False: txt(ChassisNo).TEXT = ""
    txt(DateSale).Enabled = False: txt(DateSale).TEXT = ""
    txt(KMReading).Enabled = False: txt(KMReading).TEXT = ""
    txt(SteerType).Enabled = False: txt(SteerType).TEXT = ""
    txt(VehType).Enabled = False: txt(VehType).TEXT = ""
    txt(SteerMake).Enabled = False: txt(SteerMake).TEXT = ""
    txt(BreakType).Enabled = False: txt(BreakType).TEXT = ""
    txt(Radiator).Enabled = False: txt(Radiator).TEXT = ""
End If
End Sub
Private Sub RemoveTxtNull()
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).TEXT = IIf(IsNull(txt(I).TEXT), "Null", txt(I).TEXT)
Next I
End Sub

'************ PRINTING OPTION*************
Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case VType1
        ListArray = Array("Annual", "Quarterly", "Monthly", "General(Casual)", "VOR")
        Set mListItem = ListView_Items(ListView, txtPrint, Index, ListArray, 5)
    Case FromVno, ToVno
            If IsValid(txt(VType1), "Voucher Type") = False Then Exit Sub
            RsVno.Close
            RsVno.Open "Select Order_no as code from Sp_Order where left(OrderID,1)='" & PubDivCode & "' and Sp_Order.Order_Type ='" & txtPrint(VType1).Tag & "'  ", GCn, adOpenDynamic, adLockOptimistic
            Set DGVno.DataSource = RsVno
            If RsVno.RecordCount = 0 Or (RsVno.EOF = True Or RsVno.BOF = True) Or txtPrint(Index).TEXT = "" Then Exit Sub
            If txtPrint(Index).TEXT <> RsVno!Code Then
                RsVno.MoveFirst
                RsVno.FIND "code =" & txtPrint(Index).TEXT & ""
            End If
            If Index = ToVno Then DGVno.Tag = "1" Else DGVno.Tag = "2"
End Select
End Sub

Private Sub TxtPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case VType1
        ListView_KeyDown FrmList, ListView, txtPrint, Index, KeyCode, Shift, txtPrint(Index).left + FrmPrn.left, (FrmPrn.top + txtPrint(Index).top + txtPrint(Index).height), txtPrint(Index).width, 1500
    Case FromVno, ToVno
        DGridTxtKeyDown DGVno, txtPrint, Index, RsVno, KeyCode, False, 0
End Select
If DGVno.Visible = False And FrmList.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> VType1 Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case FromVno, ToVno
        If DGVno.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVno, KeyAscii, "Code"
End Select

'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub txtPrint_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case VType1
        If FrmList.Visible = True Then ListView_KeyUp ListView, txtPrint, Index, KeyCode, mListItem
End Select
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
  Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case VType1
        If IsValid(txt(VType1), "Voucher Type") = False Then Cancel = True:   Exit Sub
        If txtPrint(VType1).TEXT <> "" Then txtPrint(VType1).TEXT = ListView.SelectedItem.TEXT
        Select Case left(txtPrint(VType1).TEXT, 1)
        Case "A"
            txtPrint(VType1).Tag = "S_POA"
        Case "Q"
            txtPrint(VType1).Tag = "S_POQ"
        Case "M"
            txtPrint(VType1).Tag = "S_POM"
        Case "G"
            txtPrint(VType1).Tag = "S_POG"
        Case "V"
            txtPrint(VType1).Tag = "S_POV"
    End Select

    Case ToVno, FromVno
        If RsVno.RecordCount = 0 Or (RsVno.EOF = True Or RsVno.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
        Else
            txtPrint(Index).TEXT = RsVno!Code
        End If
End Select
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
''by LPS 27-04-2K2
'    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> OldPNo Then
'            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(FGrid, CDate(Txt(Vdate).Text), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
''            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
'        End If
'    End If
End If
If FGrid.TextMatrix(FGrid.Rows - 1, Col_PNo) <> "" Then FGrid.AddItem FGrid.Rows
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
Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
Dim mQry$
mQry = "SELECT Syctrl.SprPurOrdFooter,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
    "SPO.OrderID,SPO.Order_Type,SPO.Order_Prefix,SPO.Order_No,SPO.V_Date,SPO.Party_Code,SPO.VOR_Type,SPO.Job_Type,SPO.Job_DocID," & _
    "SPO.Vehicle_Type,SPO.Steer_Type,SPO.Steer_Make,SPO.Brake_Type,SPO.Radiator_Make,SPO.Order_Reg_No,SPO.Order_Reg_Dt," & _
    "SPO.Tot_Items,SPO.Tot_Qty,SPO.Disc_Per,SPO.Disc_Amt,SPO.Add_Charge_Per,SPO.Add_Charge,SPO.Tot_Amount,SPO.Spr_PriceList," & _
    "SPO.Dispatch_Mode,SPO.Through,SPO.Delivery_To,SPO.IC_Name,SPO.IC_Desig,SPO.Terms,SPO.Terms2,SPO.Printed,SPO.U_Name,SPO.U_EntDt," & _
    "SPO1.OrderId,SPO1.Srl_No,SPO1.PART_NO,Part.Part_Name,SPO1.QTY,SPO1.Rate,SPO1.Disc_Amt,SPO1.Amount " & _
"FROM (((SP_Order SPO LEFT JOIN SP_Order1 SPO1 ON SPO.OrderId = SPO1.OrderId) " & _
    "LEFT JOIN Part ON SPO1.PART_NO = Part.PART_NO and Part.Div_Code = left(SPO1.OrderID,1)) " & _
    "LEFT JOIN (SubGroup SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON SPO.Party_Code = SG.SubCode) " & _
    "LEFT JOIN Syctrl ON  Syctrl.LinkTable  >= SPO.U_AE " & _
"where SPO.OrderId = '" & Master!SearchCode & "'"

Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "SpPurOrd", "SpPurOrd")
        Call WindowsPrint(Index, mQry)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(mQry)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "SpPurOrd", "SpPurOrd")
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
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrint(Index As Integer, mQry As String)
Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset, Rst2 As ADODB.Recordset
Dim DealerID As String
Dim RstSub1 As ADODB.Recordset
Dim I As Integer

On Error GoTo ERRORHANDLER
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select dealer_id from Division where Div_Code='" & PubDivCode & "'", GCn, adOpenStatic, adLockReadOnly
    If Rst2.RecordCount > 0 Then DealerID = IIf(IsNull(Rst2!Dealer_ID), "", Rst2!Dealer_ID) Else DealerID = ""
            
    Set RST1 = New ADODB.Recordset
    RST1.CursorLocation = adUseClient
    RST1.Open "select S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'", GCn, adOpenStatic, adLockReadOnly
  
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DealerId")
                rpt.FormulaFields(I).TEXT = "'" & DealerID & "'"
            Case UCase("LST")
                rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecFax & "'"
        End Select
    Next
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
                GCn.Execute "update Sp_order set Printed = 1  where Sp_order.orderid='" & Master!OrderID & "'"
            End If
        Case 1  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
    End Select
Set Rst = Nothing
Set RstSub1 = Nothing
Set RST1 = Nothing
Set Rst2 = Nothing
Exit Sub
CmdPrint(PSetUp).Tag = ""
ERRORHANDLER:
        CheckError
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrint(mQry As String)
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
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstPurOrd As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mQty As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, Footer2Cnt As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject, UserPrintStr$, mDealerID$
    
    mDealerID = GCn.Execute("Select Dealer_ID from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    Set RstPurOrd = GCn.Execute(mQry)
    If RstPurOrd.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select SprPurOrdFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
    
    Footer2Cnt = 1
    For I = 1 To Len(RstPurOrd!Terms)
        If mID(RstPurOrd!Terms, I, 1) = vbLf Then
            Footer2Cnt = Footer2Cnt + 1
        End If
    Next
    
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 12
    mFooter = mFooter + FooterCnt + Footer2Cnt
    
    'Sale Bill Header
    mDocStr = "PURCHASE ORDER"
    mDupStr = IIf(RstPurOrd!Printed = 1, "(DUPLICATE)", "")
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")

    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
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
    Print #1, PRN_TIT("* " & mDocStr & mDupStr & " *", "A", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, mChr18 & "To," & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR(RstPurOrd!NamePrefix & " " & RstPurOrd!Name, 40) & Space(1) & PSTR("Order No.", 11) & ": " & PSTR(STR(RstPurOrd!Order_NO), 14) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstPurOrd!Add1), 40) & Space(1) & PSTR("Order Date", 11) & ": " & PSTR(STR(RstPurOrd!V_Date), 14)
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstPurOrd!Add2), 40) & Space(1) & PSTR("Dealer Code", 11) & ": " & mDealerID
    mHeader = mHeader + 1
    Print #1, XNull(RstPurOrd!Add3) & IIf(XNull(RstPurOrd!Add3) <> "" And XNull(RstPurOrd!CityName) <> "", ",", "") & XNull(RstPurOrd!CityName)
    mHeader = mHeader + 1
    
    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 16) & PSTR("DESCRIPTION", 28) & PSTR("QUANTITY", 10, , AlignRight) & PSTR("RATE", 9, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & mDoub1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    If RstPurOrd.RecordCount > 0 Then
            Do Until RstPurOrd.EOF
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
                      Print #1, PSTR("M/s " & RstPurOrd!Name, 40) & Space(1) & PSTR("Order No.", 11) & ": " & PSTR(STR(RstPurOrd!Order_NO), 14) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, PSTR(XNull(RstPurOrd!Add1), 40) & Space(1) & mEmph & PSTR("Order Date", 11) & ": " & PSTR(STR(RstPurOrd!V_Date), 14) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, XNull(RstPurOrd!Add2)
                    mHeader = mHeader + 1
                    Print #1, XNull(RstPurOrd!Add3) & IIf(XNull(RstPurOrd!Add3) <> "" And XNull(RstPurOrd!CityName) <> "", ",", "") & XNull(RstPurOrd!CityName)
                    mHeader = mHeader + 1
           
                    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
                    mHeader = mHeader + 1
                    Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 16) & PSTR("DESCRIPTION", 28) & PSTR("QUANTITY", 10, , AlignRight) & PSTR("RATE", 9, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & mDoub1
                    mHeader = mHeader + 1
                    Print #1, Replace(Space(PageWidth), " ", "-")
                    mHeader = mHeader + 1
                    mFix = PageLength - (mHeader + mFooter)
                    mLine = 1
                End If
                
                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & mChr17 & PSTR(RstPurOrd!Part_No, 27, , AlignLeft) & PSTR(RstPurOrd!Part_Name, 48) & mChr18 & PSTR(RstPurOrd!Qty, 10, 3) & PSTR(RstPurOrd!Rate, 9, 2) & PSTR(RstPurOrd!Amount, 10, 3)
                mQty = mQty + RstPurOrd!Qty: mAmount = mAmount + RstPurOrd!Amount
            Print #1, PrintStr
            RstPurOrd.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    RstPurOrd.MoveFirst
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, PSTR("Total  > > ", 51, , AlignRight) & PSTR(mQty, 10, 3) & Space(9) & PSTR(mAmount, 10, 2)
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, "Please send the Parts at your earliest and inform us."
    Print #1, ""
    'Print #1, Space(Len("Please send the Parts at your earliest and inform us.")) & "Thanking You"
    'Print #1, PSTR("Sincerely Your's", 70, , AlignRight)
    Print #1, PSTR("For " & mEmph & PubComp_Name & mEmph1, PageWidth, , AlignRight)
    Print #1, ""
    Print #1, "Our Terms : " & PSTR("Autorised Signatory", PageWidth - Len("Terms And Condition :"), , AlignRight) & mDoub1 & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(RstPurOrd!Terms & vbLf)
        If mID(RstPurOrd!Terms & vbLf, I, 1) = vbLf Then
            Print #1, RTrim(mID(RstPurOrd!Terms & vbLf, j, I - j))
            j = I + 1
        End If
    Next
    Print #1, mChr17 & "By: " & RstPurOrd!U_Name & "  Dt." & RstPurOrd!U_EntDt & mChr18
    Print #1, "Terms & Conditions : " '& PSTR("Autorised Signatory", PageWidth - Len("Terms And Condition :"), , AlignRight) & mDoub1 & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    Print #1, Space((PageWidth * 1.7) - Len(UserPrintStr) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
'        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt > Prn"
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
        GCn.Execute "update Sp_order set Printed = 1  where Sp_order.orderid='" & Master!SearchCode & "'"
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

