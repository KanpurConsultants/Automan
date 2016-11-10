VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPurRet 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Purchase Return"
   ClientHeight    =   8610
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   11820
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   28
      Left            =   5985
      MaxLength       =   40
      TabIndex        =   128
      Top             =   6360
      Width           =   1500
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
      Left            =   165
      TabIndex        =   127
      Top             =   3120
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   27
      Left            =   10050
      MaxLength       =   40
      TabIndex        =   27
      Top             =   6360
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   10050
      MaxLength       =   40
      TabIndex        =   29
      Top             =   6615
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   7
      Left            =   9435
      TabIndex        =   28
      Top             =   6615
      Width           =   570
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   26
      Left            =   10050
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6870
      Width           =   1500
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
      Left            =   -930
      TabIndex        =   109
      Top             =   7695
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
         Picture         =   "frmPurRet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   119
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
         Picture         =   "frmPurRet.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmPurRet.frx":0678
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
         TabIndex        =   117
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmPurRet.frx":0982
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
         TabIndex        =   116
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmPurRet.frx":0C8C
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   -10905
      Negotiate       =   -1  'True
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   7920
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
      Left            =   7260
      TabIndex        =   76
      Top             =   -1920
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
         Index           =   45
         Left            =   3765
         TabIndex        =   107
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
         TabIndex        =   106
         Top             =   255
         Width           =   930
      End
      Begin VB.Label LblFrm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Part No>"
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
         TabIndex        =   105
         Top             =   255
         Width           =   825
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         Index           =   15
         Left            =   2805
         TabIndex        =   86
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
         Index           =   14
         Left            =   75
         TabIndex        =   85
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
         Index           =   13
         Left            =   75
         TabIndex        =   84
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
         Index           =   12
         Left            =   3930
         TabIndex        =   83
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
         Index           =   11
         Left            =   0
         TabIndex        =   82
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
         Index           =   10
         Left            =   75
         TabIndex        =   81
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
         Index           =   9
         Left            =   75
         TabIndex        =   80
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
         TabIndex        =   79
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
         Index           =   8
         Left            =   75
         TabIndex        =   78
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
         TabIndex        =   77
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
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   4140
      Negotiate       =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   7890
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
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   9
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   16
      Top             =   2220
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   9930
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1080
      Width           =   1680
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   10635
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1590
      Width           =   975
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   25
      Left            =   1605
      MaxLength       =   8
      TabIndex        =   8
      Top             =   945
      Width           =   1530
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   24
      Left            =   1605
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1200
      Width           =   4800
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -1530
      TabIndex        =   63
      Top             =   8340
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   64
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   1605
      MaxLength       =   40
      TabIndex        =   5
      Top             =   435
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   9930
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1335
      Width           =   1680
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
      TabIndex        =   17
      Top             =   3870
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   8910
      MaxLength       =   21
      TabIndex        =   1
      Top             =   540
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   23
      Left            =   10050
      MaxLength       =   40
      TabIndex        =   26
      Top             =   6105
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   22
      Left            =   10050
      MaxLength       =   40
      TabIndex        =   25
      Top             =   5850
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   21
      Left            =   5985
      MaxLength       =   40
      TabIndex        =   24
      Top             =   6615
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   20
      Left            =   5985
      MaxLength       =   40
      TabIndex        =   23
      Top             =   6105
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   19
      Left            =   5985
      MaxLength       =   40
      TabIndex        =   22
      Top             =   5850
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   18
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   21
      Top             =   6360
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   17
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   20
      Top             =   6105
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   16
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   19
      Top             =   5850
      Width           =   1500
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   13
      Left            =   4485
      MaxLength       =   15
      TabIndex        =   9
      Top             =   945
      Width           =   1920
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   12
      Left            =   1605
      MaxLength       =   30
      TabIndex        =   13
      Top             =   1710
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   11
      Left            =   5910
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1455
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   10
      Left            =   1605
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1455
      Width           =   1530
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   15
      Left            =   4935
      MaxLength       =   12
      TabIndex        =   15
      Top             =   1965
      Width           =   1470
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   14
      Left            =   1605
      MaxLength       =   15
      TabIndex        =   14
      Top             =   1965
      Width           =   1530
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   6
      Left            =   4485
      MaxLength       =   12
      TabIndex        =   7
      Top             =   690
      Width           =   1920
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   1605
      MaxLength       =   15
      TabIndex        =   6
      Top             =   690
      Width           =   1530
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   4935
      Left            =   7995
      Negotiate       =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   8055
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
      Caption         =   "Form Help"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2835
      Left            =   0
      TabIndex        =   18
      Top             =   2625
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   5001
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   40
      ForeColorFixed  =   0
      BackColorSel    =   15595518
      ForeColorSel    =   12582912
      BackColorBkg    =   14737632
      GridColor       =   0
      FocusRect       =   0
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"frmPurRet.frx":0F96
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
      _Band(0).Cols   =   40
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   4935
      Left            =   -555
      Negotiate       =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   8130
      Visible         =   0   'False
      Width           =   8955
      _ExtentX        =   15796
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
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGTrans 
      Height          =   4935
      Left            =   -3630
      Negotiate       =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   8250
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Transporter"
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
   Begin MSDataGridLib.DataGrid DGGod 
      Height          =   2145
      Left            =   5985
      Negotiate       =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3784
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
      Caption         =   "Godown Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Godown Name"
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
      Caption         =   "Additional Tax"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   24
      Left            =   3870
      TabIndex        =   129
      Top             =   6375
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   7935
      TabIndex        =   126
      Top             =   6375
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Tax @"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   48
      Left            =   7935
      TabIndex        =   125
      Top             =   6630
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   49
      Left            =   7935
      TabIndex        =   124
      Top             =   6870
      Width           =   1140
   End
   Begin VB.Label LblCancel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Cancelled*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6645
      TabIndex        =   123
      Top             =   2130
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   73
      Top             =   2235
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   31
      Left            =   7770
      TabIndex        =   72
      Top             =   1095
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   1500
      Left            =   7635
      Top             =   450
      Width           =   4050
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9900
      TabIndex        =   70
      Top             =   1605
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   7770
      TabIndex        =   69
      Top             =   1605
      Width           =   1050
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division          :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7770
      TabIndex        =   68
      Top             =   825
      Width           =   1350
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10050
      TabIndex        =   67
      Top             =   825
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase  Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   44
      Left            =   105
      TabIndex        =   66
      Top             =   945
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FormType"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   43
      Left            =   105
      TabIndex        =   65
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   42
      Left            =   7770
      TabIndex        =   61
      Top             =   555
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Return Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   41
      Left            =   7935
      TabIndex        =   59
      Top             =   6120
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   40
      Left            =   7935
      TabIndex        =   58
      Top             =   5865
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Addition"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   39
      Left            =   3870
      TabIndex        =   57
      Top             =   6630
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   38
      Left            =   3870
      TabIndex        =   56
      Top             =   6120
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Goods Value"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   37
      Left            =   3870
      TabIndex        =   55
      Top             =   5865
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Order Discount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   36
      Left            =   210
      TabIndex        =   54
      Top             =   6375
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   35
      Left            =   210
      TabIndex        =   53
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   34
      Left            =   210
      TabIndex        =   52
      Top             =   5865
      Width           =   1140
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
      Height          =   225
      Index           =   14
      Left            =   10425
      TabIndex        =   51
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Goods Amt."
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
      Index           =   33
      Left            =   8880
      TabIndex        =   50
      Top             =   5520
      Width           =   1410
   End
   Begin VB.Label LblAmt 
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
      Height          =   225
      Left            =   10755
      TabIndex        =   49
      Top             =   5520
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return  Mode "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   32
      Left            =   3225
      TabIndex        =   48
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transporter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   30
      Left            =   105
      TabIndex        =   47
      Top             =   1725
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Case"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   29
      Left            =   3225
      TabIndex        =   46
      Top             =   1455
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Case Marking"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   28
      Left            =   105
      TabIndex        =   45
      Top             =   1470
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GR/ Bilty Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   27
      Left            =   3225
      TabIndex        =   44
      Top             =   1965
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GR/Bilty No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   26
      Left            =   105
      TabIndex        =   43
      Top             =   1965
      Width           =   1050
   End
   Begin VB.Label LblPQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.000"
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
      Left            =   7710
      TabIndex        =   42
      Top             =   5520
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity(Phy)"
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
      Index           =   22
      Left            =   5790
      TabIndex        =   41
      Top             =   5520
      Width           =   1530
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
      Height          =   225
      Index           =   4
      Left            =   7500
      TabIndex        =   40
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Inv No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   105
      TabIndex        =   39
      Top             =   690
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   38
      Top             =   420
      Width           =   1260
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
      Height          =   225
      Index           =   3
      Left            =   4395
      TabIndex        =   37
      Top             =   5520
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity(Doc)"
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
      Index           =   25
      Left            =   2730
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label LblIVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   35
      Top             =   5520
      Width           =   105
   End
   Begin VB.Label LblDQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.000"
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
      Left            =   4575
      TabIndex        =   34
      Top             =   5520
      Visible         =   0   'False
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   7
      Left            =   210
      TabIndex        =   33
      Top             =   5520
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   3225
      TabIndex        =   32
      Top             =   690
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Type"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   7770
      TabIndex        =   31
      Top             =   1350
      Width           =   1050
   End
End
Attribute VB_Name = "frmPurRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsParty As ADODB.Recordset
Dim rsGod As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim rsTrans As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim rsCtrlAc As ADODB.Recordset
Dim rsTaxPer As ADODB.Recordset

Dim mCheckNegetiveStockSiteWise As Boolean
Dim FirmAddFlag As Byte
Dim GridKey As Integer
'Dim DocId As String * 21
Dim mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String
Dim ChCr As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function
Dim mVatYn As Byte
Dim mSatYn As Boolean

Dim mGatePassNo As Long

'Private Const CellBackColLeave As String = &HEDF7FE
'Private Const CellForeColLeave As String = &HFF00FF
'Private Const CellBackColEnter As String = &HF0D5BF
'Private Const GridBackColorBkg As String = &HCFE0E0
Private Const BackColorSelEnter As String = &HF8D7FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const PRetCashVType As String = "SYPRC"
Private Const PRetCrVType As String = "SYPRR"
Private Const TrfRetRecVType As String = "SYPRT"

Private Const TxtDocID As Byte = 0
Private Const SerialNo As Byte = 1
Private Const VDate As Byte = 2
Private Const VType As Byte = 3
Private Const Party As Byte = 4
Private Const SuppChlNo As Byte = 5
Private Const SuppChlDate As Byte = 6
Private Const FormType As Byte = 24
Private Const Remark As Byte = 9
Private Const CaseMark As Byte = 10
Private Const CaseNo As Byte = 11
Private Const Transport As Byte = 12
Private Const LC As Byte = 25
Private Const SupplyMode As Byte = 13
Private Const GrNo As Byte = 14
Private Const GrDate As Byte = 15
Private Const TOTAmt As Byte = 16
Private Const TotDis As Byte = 17
Private Const TotOrdDis As Byte = 18
Private Const TotGoods As Byte = 19
Private Const TaxAmt  As Byte = 20
Private Const Addition As Byte = 21
Private Const Deduction As Byte = 22
Private Const NetAmt As Byte = 23
Private Const EntryTaxPer As Byte = 7
Private Const EntryTaxAmt As Byte = 8
Private Const TotRetAmt As Byte = 26
Private Const Transportation As Byte = 27
Private Const SatAmt As Byte = 28

' Col Declaration

Private Const PNo       As Byte = 1
Private Const Unit      As Byte = 2
Private Const MRP       As Byte = 4
Private Const Taxable   As Byte = 5
Private Const DQty      As Byte = 6
Private Const PQty      As Byte = 7
Private Const FRate     As Byte = 8
Private Const Amt       As Byte = 9
Private Const DisPer    As Byte = 10
Private Const DisRs     As Byte = 11
Private Const DisOrd    As Byte = 12
Private Const DisOrdRs  As Byte = 13
Private Const TaxPer    As Byte = 14
Private Const TaxAmt1   As Byte = 15
Private Const SatPer    As Byte = 16
Private Const SatAmt1   As Byte = 17

Private Const NDP       As Byte = 18
Private Const ItemVal   As Byte = 19
Private Const Godown    As Byte = 20
Private Const PartSrlNo As Byte = 21         ' Part Serial No
Private Const PName     As Byte = 22
Private Const LName     As Byte = 23
Private Const MRPStkTB  As Byte = 24
Private Const MRPStkTP  As Byte = 25
Private Const TBStk     As Byte = 26
Private Const TPStk     As Byte = 27
Private Const TBRate    As Byte = 28
Private Const TPRate    As Byte = 29
Private Const Bin       As Byte = 30
Private Const LastRate  As Byte = 31
Private Const HPRate    As Byte = 32
Private Const LPRate    As Byte = 33
Private Const God       As Byte = 34
Private Const PONOCode  As Byte = 35
Private Const POSrlNo   As Byte = 36
Private Const PartGrade As Byte = 37
Private Const EffectDate As Byte = 38
Private Const MRPRate   As Byte = 39

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

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
If rsCtrlAc.RecordCount <= 0 Then
    MsgStr = "No Records in Spare A/c Controls"
    UnLoadFrm = True
End If
If rsCtrlAc!SprCash_Ac = "" Then
    MsgStr = "Please Fill Spare Purchase "
    UnLoadFrm = True
End If
'EOF Spare A/c control checking
If UnLoadFrm Then
    MsgBox "Spare Purchase Return Entry Loading Aborted !" & vbCrLf & MsgStr & " A/c Controls through Utility Menu", vbInformation, "Validation"
    Unload Me
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
TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    mVatYn = PubVATYN
    Call Ini_Pub
    mVType = PRetCashVType
    'A/c Pstong Control Checking
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
    'CSSprAc=Temp Sale A/c
    rsCtrlAc.Open "Select SprPurTrans_Ac,EntryTax_Ac,SprCash_Ac From AcControls", GCnFaS, adOpenDynamic, adLockOptimistic
    'eof checking


  Dim SiteCond As String
    SiteCond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
'    Master.Open "select DocID as searchcode,Sp_Purch.* from Sp_Purch where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PRetCashVType & "','" & PRetCrVType & "','" & TrfRetRecVType & "') Order By V_Date Desc, docid desc", GCn, adOpenDynamic, adLockOptimistic
    If PubMoveRecYn Then
        Master.Open "select DocID as searchcode from Sp_Purch where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PRetCashVType & "','" & PRetCrVType & "','" & TrfRetRecVType & "') " & SiteCond & " Order By V_Date Desc, docid desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("select Top 1 DocID as searchcode from Sp_Purch where left(DocID,1)='" & PubDivCode & "' " & SiteCond & " and v_type in ('" & PRetCashVType & "','" & PRetCrVType & "','" & TrfRetRecVType & "') Order By V_Date Desc, docid desc")
    End If
    
    Set DGPart.DataSource = RsPart
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct V_No as code from SP_Purch where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PRetCashVType & "','" & PRetCrVType & "','" & TrfRetRecVType & "') order by V_No", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno
    
    Set rsForm = New ADODB.Recordset
    With rsForm
        .CursorLocation = adUseClient
        .Open "SELECT  TaxForms.Form_Code as code,TaxForms.form_Desc as name FROM TaxForms where Trn_Type='Purchase' and Spare_YN = 1 Order by TaxForms.form_Desc ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGForm.DataSource = rsForm
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME,Party_Type from SubGroup Where firmCode = '" & PubFirmCode & "' and Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type,SubGroup.Add1,City.CityName from ((SubGroup " & _
        "left join City on City.CityName =Subgroup.CityCode) " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode)" & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set rsGod = New ADODB.Recordset
    rsGod.CursorLocation = adUseClient
    rsGod.Open "select god_code as code,god_name as name from godown where appli_for=0 order by god_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGod.DataSource = rsGod
    
    Set rsTrans = New ADODB.Recordset
    rsTrans.CursorLocation = adUseClient
    rsTrans.Open "select distinct transport as name from  sp_Purch  where  transport <> '' order by transport", GCn, adOpenDynamic, adLockOptimistic
    Set DGTrans.DataSource = rsTrans
    
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsParty = Nothing
Set rsGod = Nothing
Set rsForm = Nothing
Set Master = Nothing
Set rsTrans = Nothing
Set RsVno = Nothing
Set mListItem = Nothing
Set rsCtrlAc = Nothing
End Sub

Private Sub LV_Click()
'txt(VType).Text = LV.SelectedItem.Text
'LV.Visible = False
End Sub

Private Sub ListView_Click()
If FrmPrn.Visible = False Then
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txt(Val(ListView.Tag)).SetFocus
Else
    txtPrint(VType1).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txtPrint(VType1).SetFocus
End If
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    mVatYn = PubVATYN
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    If PubSatYn = 1 Then mSatYn = True
    LblVPrefix.CAPTION = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    txt(TxtDocID).Enabled = False
    mPartyType = 0
    txt(VDate).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$
If GCn.Execute("Select CancelYN from SP_Purch where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
    MsgStr = "Are You Sure To Delete This ? "
    mTitle = "Delete Entry!"
Else
    MsgStr = "Are You Sure To Cancel This ? "
    mTitle = "Cancel Entry!"
End If

If MsgBox(MsgStr, vbYesNo + vbCritical + vbDefaultButton2, mTitle) = vbYes Then
    vBook = Master.AbsolutePosition
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    'GCn.Execute ("delete from Sp_Purch where docId = '" & Master!DocId & "'")
    GCn.Execute ("delete from Sp_Stock where docId = '" & Master!SearchCode & "'")
    If mTitle = "Delete Entry!" Then
        GCn.Execute ("delete from Sp_Purch where docId = '" & Master!DocID & "'")
    Else
        GCn.Execute ("update sp_purch set CancelYN=1," & _
            " Tot_Amt=0,Tot_Disc_Amt=0,Tot_Ord_DiscAmt=0,Tot_Goods_Value=0,Tax_Amt=0,Addition=0," & _
            " Deduction=0,NET_AMT=0,U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E'," & _
            " EntryTaxamt=0, Transportation=0 where docid = '" & txt(TxtDocID) & "'")
    End If
        'Unpost Ledger a/c
        If txt(VType).TEXT = "Cash" Then
            'A/c Posting
            ProcAcPost rsCtrlAc
            'EOF Posting
        Else
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, txt(TxtDocID))
            If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
        End If
        'Unposting of Ledger completed
        GCnFaS.CommitTrans
        GCn.CommitTrans
        mTrans = False
        Master.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        Call MoveRec
    End If
eloop1:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1

    Disp_Text SETS("EDIT", Me, Master)
    txt(Party).SetFocus
    FGrid.AddItem FGrid.Rows
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub TopCtrl1_eExit()
'    Master.Cancel
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
    rsGod.Requery
    rsForm.Requery
    rsTrans.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean
Dim Rst As ADODB.Recordset, DocIdHlp As String, mGridFilled As Boolean
On Error GoTo errlbl

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(txt(TxtDocID), "Document ID") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Serial Number") = False Then Exit Sub
    If IsValid(txt(VDate), "Return Date") = False Then Exit Sub
    If IsValid(txt(VType), "Return Type") = False Then Exit Sub
    If IsValid(txt(VType), "Cash Credit") = False Then Exit Sub
    If IsValid(txt(Party), "Supplier Name") = False Then Exit Sub
    If IsValid(txt(LC), "Purchase Type") = False Then Exit Sub
    If IsValid(txt(FormType), "Form Type") = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" Then
            If FGrid.TextMatrix(I, Taxable) = "" Then MsgBox "Fill Taxable Yes/No in Row No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Taxable: FGrid.SetFocus: Exit Sub  ': FGrid.CellBackColor = CellBackColEnter
            If FGrid.TextMatrix(I, MRP) = "" Then MsgBox "Fill MRP Yes/No in Row No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = MRP: FGrid.SetFocus: Exit Sub  ': FGrid.CellBackColor = CellBackColEnter
            If Val(FGrid.TextMatrix(I, PQty)) = 0 Then MsgBox "Fill Quantity in Row No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = PQty: FGrid.SetFocus: Exit Sub  ': FGrid.CellBackColor = CellBackColEnter
            If FGrid.TextMatrix(I, God) = "" Then MsgBox "Fill Godown in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Godown: FGrid.SetFocus: Exit Sub  ': FGrid.CellBackColor = CellBackColEnter
            If Val(FGrid.TextMatrix(I, FRate)) = 0 Then
'                If PubULabel <> "Y" Then
                    MsgBox "Please Specify Rate in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = FRate: FGrid.SetFocus: Exit Sub   ': FGrid.CellBackColor = CellBackColEnter
'                End If
            End If
            If Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrd)) > Val(FGrid.TextMatrix(I, Amt)) Then
                MsgBox "Discount is greater than Item Value in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = FRate: FGrid.SetFocus: Exit Sub   ': FGrid.CellBackColor = CellBackColEnter
            End If
            If FGrid.TextMatrix(I, God) = "" Then MsgBox "Fill Godown in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Godown: FGrid.SetFocus: Exit Sub    ': FGrid.CellBackColor = CellBackColEnter
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = PNo: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
    If Val(txt(EntryTaxAmt)) <> 0 Then
        If IsNull(rsCtrlAc!EntryTax_Ac) Or rsCtrlAc!EntryTax_Ac = "" Then
            MsgBox "Please Fill Entry Tax A/c Code in " & vbCrLf & "Spare System Controls", vbInformation, "Control A/c Not Defined": txt(EntryTaxAmt).SetFocus: Exit Sub
        End If
    End If
    If Val(txt(Transportation)) <> 0 Then
        If IsNull(rsCtrlAc!SprPurTrans_Ac) Or rsCtrlAc!SprPurTrans_Ac = "" Then
            MsgBox "Please Fill Spare Purchase Transportation A/c in " & vbCrLf & "Spare System Controls", vbOKOnly, "Control A/c Not Defined"
            txt(Transportation).SetFocus: Exit Sub
        End If
    End If

    RemoveTxtNull
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        'lps 12-03-03
        txt(TxtDocID).Tag = txt(TxtDocID)
        GSQL = "select count(*) from sp_purch where Left(DocID,1)='" & PubDivCode & "' And V_Type = '" & mVType & "' And V_No = " & Val(txt(SerialNo)) & ""
        If VoucherEditFlag Then 'And txt(BookNo).Visible Then
            If GCn.Execute(GSQL).Fields(0) > 0 Then
                MsgBox "Serial No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                GoTo errlbl
            End If
        Else
            If GCn.Execute(GSQL).Fields(0) > 0 Then
                txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) Then
                    MsgBox "Serial No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If
        
        mGatePassNo = GCn.Execute("Select " & vIsNull("GatePassNo", "0") & " + 1 From Sp_Purch  ").Fields(0).Value
        DocIdHlp = Replace(txt(TxtDocID), " ", "")
        '*********
        GCn.Execute "insert into sp_purch(DocID,DocIDHelp,V_Type,V_No,Site_Code," _
            & "V_Date,Cash_Credit,Party_Code,Party_Name,Party_Doc_No," _
            & "Party_Doc_Date,GR_RR_No,GR_RR_Date," _
            & "L_C,form_code,Tot_No_of_Items,Tot_Doc_Qty,Tot_Phy_Qty," _
            & "Tot_Amt,Tot_Disc_Amt,Tot_Ord_DiscAmt,Tot_Goods_Value,Tax_Amt, SatAmt," _
            & "Addition,Deduction,NET_AMT,EntryTaxPer,EntryTaxAmt,Case_no,Case_Mark," _
            & "Transport,Supply_Mode,U_Name,U_EntDt,U_AE,Remarks,Transportation, GatePassNo, Sat_Yn) values(" _
            & "'" & txt(TxtDocID) & "','" & DocIdHlp & "','" & mVType & "'," & Val(txt(SerialNo)) & ",'" & PubSiteCode & PubSiteCode & "'," _
            & "" & ConvertDate(txt(VDate)) & ",'" & ChCr & "','" & txt(Party).Tag & "','" & txt(Party) & "','" & txt(SuppChlNo) & "'," _
            & "" & ConvertDate(txt(SuppChlDate)) & ",'" & txt(GrNo) & "'," & ConvertDate(txt(GrDate)) & "," _
            & "'" & left(txt(LC), 1) & "','" & txt(FormType).Tag & "'," & Val(LblIVal.CAPTION) & "," & Val(LblDQty.CAPTION) & "," & Val(LblPQty.CAPTION) & "," _
            & "" & Val(txt(TOTAmt)) & "," & Val(txt(TotDis)) & "," & Val(txt(TotOrdDis)) & "," & Val(txt(TotGoods)) & "," & Val(txt(TaxAmt)) & ", " & Val(txt(SatAmt)) & "," _
            & "" & Val(txt(Addition)) & "," & Val(txt(Deduction)) & "," & Val(txt(NetAmt)) & "," & Val(txt(EntryTaxPer)) & "," & Val(txt(EntryTaxAmt)) & "," & Val(txt(CaseNo)) & ",'" & txt(CaseMark) & "'," _
            & "'" & txt(Transport) & "','" & txt(SupplyMode) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & txt(Remark) & "'," & Val(txt(Transportation)) & ", " & mGatePassNo & ", " & IIf(mSatYn, 1, 0) & ")"
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, txt(TxtDocID), txt(VDate)
    Else
        GCn.Execute ("update sp_purch set Party_Code = '" & txt(Party).Tag & "', Party_Name= '" & txt(Party) & "', Party_Doc_No ='" & txt(SuppChlNo) & "', " & _
            " Party_Doc_Date =" & ConvertDate(txt(SuppChlDate)) & ",GR_RR_No='" & txt(GrNo) & "',GR_RR_Date=" & ConvertDate(txt(GrDate)) & ",L_C = '" & left(txt(LC), 1) & "',form_code = '" & txt(FormType).Tag & "',Tot_No_of_Items = " & Val(LblIVal.CAPTION) & " ,Tot_Doc_Qty = " & Val(LblDQty.CAPTION) & ",Tot_Phy_Qty = " & Val(LblPQty.CAPTION) & ",Tot_Amt = " & Val(txt(TOTAmt)) & ",Tot_Disc_Amt= " & Val(txt(TotDis)) & ",Tot_Ord_DiscAmt=" & Val(txt(TotOrdDis)) & " ," & _
            " Tot_Goods_Value=" & Val(txt(TotGoods)) & ", Tax_Amt=" & Val(txt(TaxAmt)) & ", SatAmt = " & Val(txt(SatAmt)) & ", Addition =" & Val(txt(Addition)) & "  , Deduction=" & Val(txt(Deduction)) & ",NET_AMT = " & Val(txt(NetAmt)) & ",Case_no=" & Val(txt(CaseNo)) & ",Case_Mark='" & txt(CaseMark) & "',Transport = '" & txt(Transport) & "',Supply_Mode = '" & txt(SupplyMode) & "',U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E',remarks = '" & txt(Remark) & "', " & _
            " EntryTaxPer=" & Val(txt(EntryTaxPer)) & ",EntryTaxAmt=" & Val(txt(EntryTaxAmt)) & ", Transportation=" & Val(txt(Transportation)) & _
            " Where DocId = '" & txt(TxtDocID) & "'")
    End If
    UpdStkTableToTable txt(TxtDocID), "+", "I"
    
    GCn.Execute ("delete from sp_stock where docid='" & txt(TxtDocID) & "'")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" And Val(FGrid.TextMatrix(I, PQty)) <> 0 Then
            GCn.Execute ("insert into sp_stock(DocID,Srl_No,V_Type,V_No,V_Date,Party_Code,L_C, " & _
                " Part_No, Godown, Qty_Doc, Qty_iss, Tax_YN, MRP_YN, Rate, V_Rate, " & _
                " Disc_Per,Disc_Amt , Amount, Ord_DiscPer, Ord_DiscAmt, TaxPer, TaxAmt, SatPer, SatAmt, Net_Amt, " & _
                " Part_SrlNo, Site_Code, U_Name, U_EntDt, U_AE) " & _
                " values('" & txt(TxtDocID) & "'," & I & ",'" & mVType & "'," & Val(txt(SerialNo)) & "," & ConvertDate(txt(VDate)) & ",'" & txt(Party).Tag & "','" & left(txt(LC), 1) & _
                "','" & FGrid.TextMatrix(I, PNo) & "','" & FGrid.TextMatrix(I, God) & "'," & Val(FGrid.TextMatrix(I, DQty)) & ", " & Val(FGrid.TextMatrix(I, PQty)) & "," & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ", " & IIf(FGrid.TextMatrix(I, MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, FRate)) & " ," & Val(FGrid.TextMatrix(I, NDP)) & _
                " , " & Val(FGrid.TextMatrix(I, DisPer)) & "," & Val(FGrid.TextMatrix(I, DisRs)) & "," & Val(FGrid.TextMatrix(I, Amt)) & _
                " , " & Val(FGrid.TextMatrix(I, DisOrd)) & "," & Val(FGrid.TextMatrix(I, DisOrdRs)) & ", " & Val(FGrid.TextMatrix(I, TaxPer)) & ", " & Val(FGrid.TextMatrix(I, TaxAmt1)) & ", " & Val(FGrid.TextMatrix(I, SatPer)) & ", " & Val(FGrid.TextMatrix(I, SatAmt1)) & "," & Val(FGrid.TextMatrix(I, ItemVal)) & _
                " ,'" & FGrid.TextMatrix(I, PartSrlNo) & "','" & PubSiteCode & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
            Call UpdStkGridToTable(FGrid.TextMatrix(I, PNo), "-", FGrid.TextMatrix(I, MRP), FGrid.TextMatrix(I, Taxable), FGrid.TextMatrix(I, PQty))
        End If
    Next
    'A/c Posting
    '************
    ProcAcPost rsCtrlAc
    'EOF Posting
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select DocID as searchcode from Sp_Purch where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PRetCashVType & "','" & PRetCrVType & "','" & TrfRetRecVType & "') And DocId = '" & txt(TxtDocID) & "' Order By V_Date Desc, docid desc")
    End If
    rsTrans.Requery
    Master.FIND "SearchCode = '" & txt(TxtDocID) & "'"
    'lp 12-03-03
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) Then
            MsgBox "Serial No." & Trim(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
    Dim SiteCond As String
    SiteCond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("sp_purch.Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
       
    GSQL = "SELECT sp_purch.DocId as searchcode, sp_purch.DocId, sp_purch.v_Type, sp_purch.v_No, sp_purch.Site_Code, sp_purch.V_Date AS VoucherDate, SubGroup.Name as PartyName FROM sp_purch LEFT JOIN SubGroup ON sp_purch.Party_Code = SubGroup.Subcode where  left(DocID,1)='" & PubDivCode & "' " & SiteCond & " and v_type in ('" & PRetCashVType & "','" & PRetCrVType & "','" & TrfRetRecVType & "') order by sp_purch.docId"
    Set SearchForm = Me
    FIND.Show vbModal
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
        Set Master = GCn.Execute("select DocID as searchcode from Sp_Purch where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PRetCashVType & "','" & PRetCrVType & "','" & TrfRetRecVType & "') And DocId = '" & MyValue & "' Order By V_Date Desc, docid desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
If txt(VType).TEXT = "" And Index <> VDate Then txt(VType).SetFocus
TxtGrid(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case VType
        ListArray = Array("Cash", "Credit", "Stock Transfer")
        Set mListItem = ListView_Items(ListView, txt, VType, ListArray, 3)
    Case SerialNo
        If IsValid(txt(VType), "Return Type") = False Then Exit Sub
    Case LC
        ListArray = Array("Local", "Central")
        Set mListItem = ListView_Items(ListView, txt, LC, ListArray, 2)
    Case Party
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case FormType
        Set DGForm.DataSource = rsForm
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case SerialNo, CaseNo, Addition, Deduction, TaxAmt, NetAmt, SatAmt
        SendKeys "{HOME}+{END}"
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
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 900
    Case LC
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case Party
        If ChCr = "Credit" Then
            DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        End If
    Case Transport
        DGridTxtKeyDown_Mast DGTrans, txt, Transport, rsTrans, KeyCode, False, 0
    Case FormType
        DGridTxtKeyDown DGForm, txt, FormType, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
End Select
If FrmList.Visible = False And DGTrans.Visible = False And DGGod.Visible = False And DGParty.Visible = False And DGForm.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> EntryTaxAmt Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = EntryTaxAmt Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> VDate Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Party Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case SerialNo, CaseNo
        Call NumPress(txt(Index), KeyAscii, 6, 0)
    Case Party
        If txt(VType).TEXT = "Credit" Then
            If DGParty.Visible = True Then DGridTxtKeyPress txt, Party, RsParty, KeyAscii, "Name"
            lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.BackColor = vbBlack: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
        End If
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, FormType, rsForm, KeyAscii, "Name"
    Case EntryTaxPer
        Call NumPress(txt(Index), KeyAscii, 2, 2)
    Case Addition, Deduction, TaxAmt, EntryTaxAmt, Transportation, SatAmt
        Call NumPress(txt(Index), KeyAscii, 8, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case VType
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case LC
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case Transport
        If DGTrans.Visible = True Then DGridTxtKeyUp_Mast txt, Transport, rsTrans, KeyCode, "Name"
    Case Addition, Deduction, TaxAmt, EntryTaxPer, EntryTaxAmt, Transportation, SatAmt
        Amt_Cal Index
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Select Case Index
    Case EntryTaxPer, Addition, Deduction, TaxAmt, EntryTaxAmt, Transportation, SatAmt
        txt(Index) = Format(txt(Index), "0.00")
        
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
            txt(VDate).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        Cancel = Not CheckFinYear(txt(Index))
    Case VType
        If IsValid(txt(VType), "Cash Credit") = False Then Cancel = True:   Exit Sub
        If txt(VType).TEXT <> "" Then txt(VType).TEXT = ListView.SelectedItem.TEXT
        If txt(Index).TEXT = "Cash" Then
            mVType = PRetCashVType
            ChCr = "Cash"
            txt(Party).TEXT = "Cash"
        ElseIf txt(Index).TEXT = "Credit" Then
            mVType = PRetCrVType
            ChCr = "Credit"
        Else
            mVType = TrfRetRecVType
            ChCr = "Credit"
        End If
        If txt(VType).TEXT = "Cash" Then
            txt(Party).TEXT = "Cash"
            txt(Index).Tag = PubSprCashAc
        End If
        txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        txt(TxtDocID).Tag = txt(TxtDocID)
        txt(VType).Tag = txt(VType).TEXT
    Case SerialNo
        If IsValid(txt(SerialNo), "SerialNo") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag Then      ' Manual
            txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            txt(TxtDocID).Tag = txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select docid From sp_purch Where docid='" & txt(TxtDocID) & "'", GCn, adOpenStatic, adLockReadOnly
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(SerialNo).SetFocus
            End If
        End If
    Case LC
        If IsValid(txt(LC), "Purchase Type") = False Then Cancel = True:   Exit Sub
        If txt(LC).TEXT <> "" Then txt(LC).TEXT = ListView.SelectedItem.TEXT
    Case Party
        If IsValid(txt(Index), "Party") = False Then Cancel = True: Exit Sub
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
            If txt(VType).TEXT = "Cash" Then
                txt(Index).Tag = PubSprCashAc
            ElseIf txt(VType).TEXT = "Credit" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Else
            If txt(VType).TEXT = "Cash" Then
                txt(Index).Tag = PubSprCashAc
            ElseIf txt(VType).TEXT = "Credit" Then
                txt(Index).TEXT = RsParty!Name
                txt(Index).Tag = RsParty!Code
            End If
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsForm!Name
            txt(Index).Tag = rsForm!Code
        End If
    Case SuppChlDate, GrDate
        txt(Index).TEXT = RetDate(txt(Index))
End Select
Set Rst = Nothing
End Sub

Private Sub DGPart_Click()
If RsPart.RecordCount > 0 Then
    Select Case FGrid.Col
        Case PNo
            TxtGrid(0).TEXT = RsPart!Code
        Case PName
            TxtGrid(0) = RsPart!Name
        Case LName
            TxtGrid(0) = RsPart!LName
    End Select
End If
    TxtGridValid_PNo
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGPart.Visible = False
End Sub
Private Sub DGTrans_Click()
    If rsTrans.RecordCount > 0 Then
        txt(Transport).TEXT = rsTrans!Name
    End If
    txt(Transport).SetFocus
    DGTrans.Visible = False
End Sub

Private Sub DGForm_Click()
    If rsForm.RecordCount > 0 Then
        txt(FormType).TEXT = rsForm!Name
        txt(FormType).Tag = rsForm!Code
    End If
    txt(FormType).SetFocus
    DGForm.Visible = False
End Sub
Private Sub DGGod_Click()
    If rsGod.RecordCount > 0 Then
        TxtGrid(0).TEXT = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
    End If
   TxtGrid(0).SetFocus
    DGGod.Visible = False
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
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
    End If
    txt(Party).SetFocus
    DGParty.Visible = False
    lblGroup.Visible = False
End Sub

Private Sub FGrid_Click()
TxtGrid(0).Visible = False
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub
Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
    If TopCtrl1.TopText2 <> "Browse" Then
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, PNo), _
            FGrid.TextMatrix(FGrid.Row, PName), FGrid.TextMatrix(FGrid.Row, LName), _
            MRPStkTB, MRPStkTP, TBStk, TPStk, _
            MRPRate, TBRate, TPRate, Bin, _
            LastRate, HPRate, LPRate, mCheckNegetiveStockSiteWise
        FrmDetail.Visible = True
    End If
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If txt(VType).TEXT = "" Then txt(VType).SetFocus: Exit Sub
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
        Case PQty, PartSrlNo
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case DisRs, DisPer
            FGrid.TextMatrix(FGrid.Row, DisRs) = ""
            FGrid.TextMatrix(FGrid.Row, DisPer) = ""
        Case DisOrd, DisOrdRs
            FGrid.TextMatrix(FGrid.Row, DisOrd) = ""
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = ""
        Case NDP
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
    Amt_Cal1
    Amt_Cal
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case PNo, PName, LName
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
        Case Taxable, Godown, FRate, MRP, PQty, DisPer, DisOrd, DisRs, DisOrdRs, PartSrlNo, TaxPer, SatPer
            If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
                Call GridDblClick(Me, FGrid, TxtGrid, 0)
                TAddMode = False
            End If
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
Select Case FGrid.Col
    Case PNo, PName, LName
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case Unit, Amt, ItemVal
        FGrid.Col = FGrid.Col + 1
        FGrid.SetFocus
    Case PartSrlNo, Godown, MRP, Taxable
        If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
           Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        End If
    Case PQty, FRate, DisPer, DisOrd, DisRs, DisOrdRs, TaxPer, SatPer
        If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
           Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
Dim I As Integer
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
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, PNo), _
            FGrid.TextMatrix(FGrid.Row, PName), FGrid.TextMatrix(FGrid.Row, LName), _
            MRPStkTB, MRPStkTP, TBStk, TPStk, _
            MRPRate, TBRate, TPRate, Bin, _
            LastRate, HPRate, LPRate, mCheckNegetiveStockSiteWise
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
    If FrmDetail.Visible = True Then FrmDetail.Visible = False
End Sub

Private Sub FGrid_RowColChange()
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, PNo), _
           FGrid.TextMatrix(FGrid.Row, PName), FGrid.TextMatrix(FGrid.Row, LName), _
           MRPStkTB, MRPStkTP, TBStk, TPStk, _
           MRPRate, TBRate, TPRate, Bin, _
           LastRate, HPRate, LPRate, mCheckNegetiveStockSiteWise
    End If
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
txt(TxtDocID).Tag = ""
End Sub

Private Sub MoveRec()
Dim Rs As Recordset, Master1 As ADODB.Recordset, I As Integer
On Error GoTo error1
If Master.RecordCount > 0 Then
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select SubGroup.Name,SubGroup.Party_Type,SP_Purch.* from SP_Purch " _
        & " left join SubGroup on SP_Purch.Party_Code=SubGroup.SubCode " _
        & " where DocID='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly

    txt(TxtDocID) = Master!SearchCode
    txt(TxtDocID).Tag = Master!SearchCode
    
    mVType = Master1!V_Type
    LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
    LblSite.CAPTION = "Site Code : " & Master1!Site_Code
    LblVPrefix.CAPTION = mID(Master1!DocID, 8, 5)
    txt(SerialNo) = Master1!V_NO
    
    txt(VDate) = Master1!V_Date
    mVType = Master1!V_Type
    If mVType = PRetCashVType Then
        txt(VType) = "Cash"
        ChCr = "Cash"
    ElseIf mVType = PRetCrVType Then
        txt(VType) = "Credit"
        ChCr = "Credit"
    Else
        txt(VType) = "Stock Transfer"
        ChCr = "Credit"
    End If
    
    
    mVatYn = PubVATYN
    If StrCmp(left(PubComp_Name, 3), "jmk") And CDate(Master1!V_Date) < CDate("01/Jan/2008") Then
        mVatYn = 0
    End If
    
    
    txt(VType) = Master1!Cash_Credit
    txt(Party).Tag = Master1!Party_code
    If Master1!Cash_Credit = "Cash" Then
        txt(Party) = Master1!Party_Name
        mPartyType = 0
    Else
        txt(Party) = IIf(IsNull(Master1!Name), "", Master1!Name)
        mPartyType = Master1!Party_Type
    End If
    
    If PubBackEnd = "A" Then
        mSatYn = IIf(VNull(Master1!SAT_YN) = 1, True, False)
    Else
        mSatYn = IIf(VNull(Master1!SAT_YN) = True, True, False)
    End If
    
    txt(SuppChlNo) = IIf(IsNull(Master1!Party_Doc_No), "", Master1!Party_Doc_No)
    txt(SuppChlDate) = IIf(IsNull(Master1!Party_Doc_Date), "", Master1!Party_Doc_Date)
    txt(GrNo) = IIf(IsNull(Master1!GR_RR_No), "", Master1!GR_RR_No)
    txt(GrDate) = IIf(IsNull(Master1!GR_RR_Date), "", Master1!GR_RR_Date)
    txt(LC) = IIf(Master1!L_C = "L", "Local", "Central")
    txt(FormType).Tag = IIf(IsNull(Master1!Form_Code), "", Master1!Form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType) = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType) = ""
    End If
    LblIVal.CAPTION = Format(IIf(IsNull(Master1!Tot_No_of_Items), 0, Master1!Tot_No_of_Items), "0")
    LblDQty.CAPTION = Format(IIf(IsNull(Master1!Tot_Doc_Qty), 0, Master1!Tot_Doc_Qty), "0.000")
    LblPQty.CAPTION = Format(IIf(IsNull(Master1!Tot_Phy_Qty), 0, Master1!Tot_Phy_Qty), "0.000")
    LblAmt.CAPTION = Format(IIf(IsNull(Master1!Tot_Amt), 0, Master1!Tot_Amt), "0.00")
    txt(TOTAmt) = Format(IIf(IsNull(Master1!Tot_Amt), 0, Master1!Tot_Amt), "0.00")
    txt(TotDis) = Format(IIf(IsNull(Master1!Tot_Disc_Amt), 0, Master1!Tot_Disc_Amt), "0.00")
    txt(TotOrdDis) = Format(IIf(IsNull(Master1!Tot_Ord_DiscAmt), 0, Master1!Tot_Ord_DiscAmt), "0.00")
    txt(TotGoods) = Format(IIf(IsNull(Master1!Tot_Goods_Value), 0, Master1!Tot_Goods_Value), "0.00")
    txt(TaxAmt) = Format(IIf(IsNull(Master1!Tax_Amt), 0, Master1!Tax_Amt), "0.00")
    txt(SatAmt) = Format(VNull(Master1!SatAmt), "0.00")
    txt(Addition) = Format(IIf(IsNull(Master1!Addition), 0, Master1!Addition), "0.00")
    txt(Deduction) = Format(IIf(IsNull(Master1!Deduction), 0, Master1!Deduction), "0.00")
    txt(NetAmt) = Format(IIf(IsNull(Master1!Net_Amt), 0, Master1!Net_Amt), "0.00")
    txt(EntryTaxPer) = Format(IIf(IsNull(Master1!EntryTaxPer), 0, Master1!EntryTaxPer), "0.00")
    txt(EntryTaxAmt) = Format(IIf(IsNull(Master1!EntryTaxAmt), 0, Master1!EntryTaxAmt), "0.00")
    txt(Transportation) = Format(IIf(IsNull(Master1!Transportation), 0, Master1!Transportation), "0.00")
    txt(TotRetAmt) = Format(Val(txt(NetAmt)) + Val(txt(EntryTaxAmt)) + Val(txt(Transportation)), "0.00")
    txt(CaseNo) = IIf(IsNull(Master1!Case_No), "", Master1!Case_No)
    txt(CaseMark) = IIf(IsNull(Master1!Case_Mark), "", Master1!Case_Mark)
    txt(Transport) = IIf(IsNull(Master1!Transport), "", Master1!Transport)
    txt(SupplyMode) = IIf(IsNull(Master1!Supply_Mode), "", Master1!Supply_Mode)
    txt(Remark) = IIf(IsNull(Master1!Remarks), "", Master1!Remarks)

    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT P.Part_Name, P.Local_Name, P.UNIT, P.MRP,P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk, P.Cur_TP_Stk, " & _
            " P.TP_SRate, P.TB_SRate, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate, Sp_Stock.*, Godown.God_Name" & _
            " FROM (Sp_Stock LEFT JOIN Part P ON Sp_Stock.Part_No = P.PART_NO and P.Div_Code = left(SP_Stock.DocID,1)) LEFT JOIN Godown ON Sp_Stock.Godown = Godown.God_Code " & _
            " where Sp_Stock.docId = '" & Master1!DocID & "'")
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        I = 1
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = Rs!Srl_No
                .TextMatrix(I, PNo) = Rs!Part_No
                .TextMatrix(I, Unit) = IIf(IsNull(Rs!Unit), "", Rs!Unit)
                .TextMatrix(I, MRP) = IIf(Rs!MRP_YN = 1, "Yes", "No")
                .TextMatrix(I, Taxable) = IIf(Rs!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, DQty) = ""   'Format(rs!Qty_doc, "0.000")
                .TextMatrix(I, PQty) = Format(Rs!Qty_Iss, "0.000")
                .TextMatrix(I, FRate) = Format(Rs!Rate, "0.00")   'Rahul U.N. Automobiles 10-04-2003
                .TextMatrix(I, Amt) = Format(Rs!Amount, "0.00")
                .TextMatrix(I, DisPer) = Format(Rs!Disc_Per, "0.00")
                .TextMatrix(I, DisRs) = Format(Rs!Disc_Amt, "0.00")
                .TextMatrix(I, DisOrd) = Format(Rs!ord_Discper, "0.00")
                .TextMatrix(I, DisOrdRs) = Format(Rs!ord_Discamt, "0.00")
                .TextMatrix(I, TaxPer) = Format(VNull(Rs!TaxPer), "0.00")
                .TextMatrix(I, TaxAmt1) = Format(VNull(Rs!TaxAmt), "0.00")
                If PubVATYN = 1 Then
                    .TextMatrix(I, TaxPer) = VNull(Rs!TaxPer)
                    .TextMatrix(I, TaxAmt1) = Format(VNull(Rs!TaxAmt), "0.00")
                    If mSatYn Then
                        .TextMatrix(I, SatPer) = VNull(Rs!SatPer)
                        .TextMatrix(I, SatAmt1) = Format(VNull(Rs!SatAmt), "0.00")
                    End If
                End If
                
                .TextMatrix(I, NDP) = Format(Rs!V_Rate, "0.00")
                .TextMatrix(I, ItemVal) = Format(Rs!Net_Amt, "0.00")
                .TextMatrix(I, God) = Rs!Godown
                .TextMatrix(I, Godown) = IIf(IsNull(Rs!God_Name), "", Rs!God_Name)
                .TextMatrix(I, PName) = IIf(IsNull(Rs!Part_Name), "", Rs!Part_Name)
                .TextMatrix(I, LName) = IIf(IsNull(Rs!Local_Name), "", Rs!Local_Name)
                .TextMatrix(I, MRPStkTB) = IIf(IsNull(Rs!Cur_MRP_TbStk), "", Rs!Cur_MRP_TbStk)
                .TextMatrix(I, MRPStkTP) = IIf(IsNull(Rs!Cur_MRP_TPStk), "", Rs!Cur_MRP_TPStk)
                .TextMatrix(I, MRPRate) = Format(Rs!MRP, "0.00")
                .TextMatrix(I, TBStk) = IIf(IsNull(Rs!Cur_TB_STk), "", Rs!Cur_TB_STk)
                .TextMatrix(I, TPStk) = IIf(IsNull(Rs!Cur_TP_Stk), "", Rs!Cur_TP_Stk)
                .TextMatrix(I, TBRate) = IIf(IsNull(Rs!TB_SRate), "", Rs!TB_SRate)
                .TextMatrix(I, TPRate) = IIf(IsNull(Rs!TP_SRate), "", Rs!TP_SRate)
                .TextMatrix(I, Bin) = IIf(IsNull(Rs!Bin_Loca), "", Rs!Bin_Loca)
'                    .TextMatrix(i, LastRate) = ""
                .TextMatrix(I, HPRate) = IIf(IsNull(Rs!high_pur_rate), "", Rs!high_pur_rate)
                .TextMatrix(I, LPRate) = IIf(IsNull(Rs!low_pur_rate), "", Rs!low_pur_rate)
            End With
            Rs.MoveNext
            I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Set Rs = Nothing
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
Set Rs = Nothing
Set Master1 = Nothing
Grid_Hide
'Call Amt_Cal
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte
' |Part No.1|Part Name2|Unit 3|PO No 4|Taxable 5|MRP6|Qty(Doc)7|Qty(Phy)8|NDP 9 |Amount 10
' |Dis %11|Ord Dis %12|Amount 13|Loal Name 14|Curr Stk Qty 15|MRP Qty 16 |Taxable Qty 17|TaxPaid Qty 18|Taxable Rate 19|TaxPaid Rate 20|Bin Location 21|Last Purch Rate 22|High Purch Rate 23|Low Purch Rate 24
FrmPrn.left = (Me.width - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
DGVno.left = 5145: DGVno.top = mTopScale
FGrid.left = Me.left: FGrid.width = Me.width - 90: FGrid.top = 2610  ': FGrid.height = 2895
DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (DGPart.top + mBotScale)
FrmDetail.width = 6285: FrmDetail.left = Me.width - (FrmDetail.width + mRtScale): FrmDetail.top = mTopScale: FrmDetail.height = 2130
DGGod.left = Me.width - (DGGod.width + mRtScale): DGGod.top = mTopScale

'DGParty.width = 5130:   DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale '390
'DGParty.height = 4935

DGParty.width = 9500:   DGParty.left = 1500: DGParty.top = mTopScale '390
DGParty.height = 4935
DGTrans.width = DGParty.width: DGTrans.left = DGParty.left: DGTrans.top = DGParty.top: DGTrans.height = DGParty.height
DGForm.width = DGParty.width: DGForm.left = DGParty.left: DGForm.top = DGParty.top: DGForm.height = DGParty.height
'    FGrid.FormatString = "SrNo.|Part No.            |Part Name             |Unit |Godown          |PO No.         |Tax Y/N|MRP Y/N| Qty(Doc)|Qty(Phy)|Rate     |Amount    |Dis %    |Dis Rs   |Ord Dis %  |Ord Dis Rs  |NDP     |ItemValue   |Local Name|Curr Stk Qty|MRP Qty|Taxable Qty|TaxPaid Qty|Taxable Rate|TaxPaid Rate|Bin Location|Last Purch Rate|High Purch Rate|Low Purch Rate"
    'SrNo.1|Part No.2|Part Name3|Unit 4|Godown5|PO No.6|Tax Y/N 7|MRP Y/N8| Qty(Doc)9|Qty(Phy)10|Rate 11|Amount12|Dis %13|Dis Rs14|Ord Dis %15|Ord Dis Rs16|NDP 17|ItemValue 18|Local Name19|Curr Stk Qty20|MRP Qty21|Taxable Qty22|TaxPaid Qty23|Taxable Rate24|TaxPaid Rate25|Bin Location26|Last Purch Rate27|High Purch Rate28|Low Purch Rate29"
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    With FGrid
        .RowHeightMin = PubGridRowHeight
        
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, PNo) = "Part No"
        .ColAlignment(PNo) = flexAlignLeftCenter
        .ColWidth(PNo) = 1500

        .ColWidth(3) = 0 'PONO

        .TextMatrix(0, Unit) = "Unit"
        .ColAlignment(Unit) = flexAlignLeftCenter
        .ColWidth(Unit) = 550

        .TextMatrix(0, MRP) = "MRP"
        .ColAlignment(MRP) = flexAlignLeftCenter
        .ColWidth(MRP) = 450

        .TextMatrix(0, Taxable) = "Tax"
        .ColAlignment(Taxable) = flexAlignLeftCenter
        .ColWidth(Taxable) = 420

        .TextMatrix(0, PQty) = "Quantity"
        .ColAlignmentFixed(PQty) = flexAlignRightCenter
        .ColWidth(PQty) = 960

'        .TextMatrix(0, DQty) = "Qty(Doc)"
'        .ColAlignment(DQty) = flexAlignRightCenter
        .ColWidth(DQty) = 0
        
        .TextMatrix(0, NDP) = "Rate"
        .ColAlignmentFixed(NDP) = flexAlignRightCenter
        .ColWidth(NDP) = 870

        .TextMatrix(0, FRate) = "Rate" 'NDP"
        .ColAlignmentFixed(FRate) = flexAlignRightCenter
        .ColWidth(FRate) = 870

        .TextMatrix(0, Amt) = "Amount"
        .ColAlignmentFixed(Amt) = flexAlignRightCenter
        .ColWidth(Amt) = 1065

        .TextMatrix(0, DisOrd) = "ODis%"
        .ColAlignmentFixed(DisOrd) = flexAlignRightCenter
        .ColWidth(DisOrd) = 555

        .TextMatrix(0, DisOrdRs) = "OrdDisc"
        .ColAlignmentFixed(DisOrdRs) = flexAlignRightCenter
        .ColWidth(DisOrdRs) = 840
        
        .TextMatrix(0, TaxPer) = "Tax%"
        .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
        .ColWidth(TaxPer) = 555

        .TextMatrix(0, TaxAmt1) = "Tax Amt"
        .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
        .ColWidth(TaxAmt1) = 840
        
        If PubSatYn = 1 Then
            .TextMatrix(0, SatPer) = "SAT %"
            .ColAlignmentFixed(SatPer) = flexAlignRightCenter
            .ColWidth(SatPer) = 555
    
            .TextMatrix(0, SatAmt1) = "SAT Amt"
            .ColAlignmentFixed(SatAmt1) = flexAlignRightCenter
            .ColWidth(SatAmt1) = 840
        End If
        
        
        
        .TextMatrix(0, DisPer) = "Disc%"
        .ColAlignmentFixed(DisPer) = flexAlignRightCenter
        .ColWidth(DisPer) = 555

        .TextMatrix(0, DisRs) = "Disc.Amt"
        .ColAlignmentFixed(DisRs) = flexAlignRightCenter
        .ColWidth(DisRs) = 840

        .TextMatrix(0, ItemVal) = "Item Value"
        .ColAlignmentFixed(ItemVal) = flexAlignRightCenter
        .ColWidth(ItemVal) = 1095
        
        .ColWidth(God) = 0
        .TextMatrix(0, Godown) = "Godown"
        .ColAlignmentFixed(Godown) = flexAlignRightCenter
        .ColWidth(Godown) = 1095

        .TextMatrix(0, PartSrlNo) = "Part SrlNo"
        .ColAlignmentFixed(PartSrlNo) = flexAlignLeftCenter
        .ColAlignment(PartSrlNo) = flexAlignLeftCenter
        .ColWidth(PartSrlNo) = 1095
        
        .TextMatrix(0, PName) = "Part Name"
        .ColAlignment(PName) = flexAlignLeftCenter
        .ColWidth(PName) = 2500
        
        .TextMatrix(0, LName) = "Local Name"
        .ColAlignment(LName) = flexAlignLeftCenter
        .ColWidth(LName) = 2000
    End With
    For I = 19 To 35
        FGrid.ColWidth(I) = 0
    Next
    
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
    txt(SerialNo).Enabled = False
    txt(VType).Enabled = False
End If
txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol

End Sub

Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
    If DGTrans.Visible = True Then DGTrans.Visible = False
    If DGGod.Visible = True Then DGGod.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub
Private Sub Amt_Cal1()
Dim mAmount As Double
Dim DisAmt As Double
Dim OrdDisAmt1 As Double
Dim mAddAmt As Double
Dim mTaxableAmt As Double
    If FGrid.TextMatrix(FGrid.Row, DisPer) <> "" Then
        FGrid.TextMatrix(FGrid.Row, DisRs) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, DisPer)) / 100, "0.00")
    End If
    If FGrid.TextMatrix(FGrid.Row, DisOrd) <> "" Then
        FGrid.TextMatrix(FGrid.Row, DisOrdRs) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs))) * Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) / 100, "0.00")
    End If
    FGrid.TextMatrix(FGrid.Row, ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs)) - Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs)), "0.00")
    If Val(FGrid.TextMatrix(FGrid.Row, PQty)) <> 0 Then
        FGrid.TextMatrix(FGrid.Row, NDP) = Format(Val(FGrid.TextMatrix(FGrid.Row, ItemVal)) / Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
    Else
        FGrid.TextMatrix(FGrid.Row, NDP) = ""
    End If
    
    
    
    If mVatYn = 1 Then
        If FGrid.TextMatrix(FGrid.Row, TaxPer) <> "" Then
            mAmount = Val(FGrid.TextMatrix(FGrid.Row, Amt))
            DisAmt = Val(FGrid.TextMatrix(FGrid.Row, DisRs))
            OrdDisAmt1 = Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs))
            
            
            If FGrid.TextMatrix(FGrid.Row, MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes" Then
                If mSatYn Then
                    mTaxableAmt = Format((mAmount - (DisAmt + OrdDisAmt1)) * 100 / (100 + Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) + Val(FGrid.TextMatrix(FGrid.Row, SatPer))), "0.00")
                    FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                    FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                Else
                    FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1) + mAddAmt) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / (100 + Val(FGrid.TextMatrix(FGrid.Row, TaxPer))), "0.00")
                End If
                
                FGrid.TextMatrix(FGrid.Row, ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, ItemVal)) - Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)), "0.00")
            ElseIf FGrid.TextMatrix(FGrid.Row, MRP) = "No" And FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes" Then
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1) + mAddAmt) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                If mSatYn Then
                    FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                End If
                
            Else
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
            End If
        End If
    End If
    
End Sub

Private Sub Amt_Cal(Optional Index As Integer)
 Dim I As Byte
 Dim IQty As Double, DQty1 As Double, ICnt As Integer, IGAmt As Double, IDic As Double
 Dim TaxPer As Double, SurPer As Double, TaxAmountMRP As Double, TaxAmount As Double
 Dim mTaxPer As Double
 Dim mTaxableAmt As Double
 Dim SatAmount As Double
 Dim IOrdDic As Double, IAmt As Double, mNetAmt As Double
 IQty = 0: DQty1 = 0: IGAmt = 0: IAmt = 0: IDic = 0: IOrdDic = 0

For I = 1 To FGrid.Rows - 1
    If FGrid.TextMatrix(I, PNo) <> "" Then
        IQty = IQty + Val(FGrid.TextMatrix(I, PQty))
        DQty1 = DQty1 + Val(FGrid.TextMatrix(I, DQty))
        IGAmt = IGAmt + Val(FGrid.TextMatrix(I, ItemVal))
        IAmt = IAmt + Val(FGrid.TextMatrix(I, Amt))
        IDic = IDic + Val(FGrid.TextMatrix(I, DisRs))
        IOrdDic = IOrdDic + Val(FGrid.TextMatrix(I, DisOrdRs))
        
        
        
        If txt(FormType) <> "" Then
            TaxPer = GCn.Execute("Select Tax_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'").Fields(0).Value
            SurPer = GCn.Execute("Select Tax_Sur_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'").Fields(0).Value
        End If
        
        If mVatYn = 1 Then
            If FGrid.TextMatrix(I, MRP) = "Yes" Then
                TaxAmountMRP = TaxAmountMRP + Val(FGrid.TextMatrix(I, TaxAmt1))
                TaxAmount = TaxAmount + Val(FGrid.TextMatrix(I, TaxAmt1))
            Else
                TaxAmount = TaxAmount + Val(FGrid.TextMatrix(I, TaxAmt1))
            End If
            SatAmount = SatAmount + Val(FGrid.TextMatrix(I, SatAmt1))
        Else
            If txt(FormType) <> "" Then
                mTaxPer = TaxPer + (TaxPer * SurPer / 100)
                If FGrid.TextMatrix(I, MRP) = "Yes" Then
                    TaxAmount = TaxAmount + Round(((Val(FGrid.TextMatrix(I, Amt)) - Val(FGrid.TextMatrix(I, DisRs)) - Val(FGrid.TextMatrix(I, DisOrdRs))) * mTaxPer) / (100 + mTaxPer), 2)
                Else
                    TaxAmount = TaxAmount + Round(((Val(FGrid.TextMatrix(I, Amt)) - Val(FGrid.TextMatrix(I, DisRs)) - Val(FGrid.TextMatrix(I, DisOrdRs))) * mTaxPer) / 100, 2)
                End If
            End If
        End If
        
        
        
        ICnt = ICnt + 1
    End If
Next I


        






    LblIVal.CAPTION = Format(ICnt, "0")
    LblPQty.CAPTION = Format(IQty, "0.000")
    LblDQty.CAPTION = Format(DQty1, "0.000")
    LblAmt.CAPTION = Format(IGAmt, "0.00")
    txt(TOTAmt) = Format(IAmt, "0.00")
    txt(TotDis) = Format(IDic, "0.00")
    txt(TotOrdDis) = Format(IOrdDic, "0.00")
    txt(TotGoods) = Format(IGAmt, "0.00")
    
    If Not StrCmp(txt(LC), "Local") Then
        mTaxableAmt = Val(txt(TotGoods))
        txt(TaxAmt) = Format(mTaxableAmt * TaxPer / 100, "0.00")
    Else
        txt(TaxAmt) = Format(TaxAmount, "0.00")
    End If
    
    txt(SatAmt) = Format(SatAmount, "0.00")
    
    txt(NetAmt) = Format((IGAmt + Val(txt(TaxAmt)) + Val(txt(SatAmt)) + Val(txt(Addition)) - Val(txt(Deduction))), "0.00")
    'For Entry Tax
    mNetAmt = Val(txt(NetAmt)) + Val(txt(Transportation))
    If Index <> EntryTaxAmt Then
        txt(EntryTaxAmt) = Format(mNetAmt * Val(txt(EntryTaxPer)) / 100, "0.00")
    End If
    txt(TotRetAmt) = Format(mNetAmt + Val(txt(EntryTaxAmt)), "0.00")
    'Eof Entry Tax
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
Grid_Hide
If FrmDetail.Visible = False Then FrmDetail.Visible = True
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
TxtGrid(Index).MaxLength = 0
    Select Case FGrid.Col
         Case PNo
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, PNo) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case PName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "name"
            If FGrid.TextMatrix(FGrid.Row, PName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, PName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case LName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "lname"
            If FGrid.TextMatrix(FGrid.Row, LName) = "" Then
                RsPart.MoveFirst
                RsPart.FIND "lname ='" & FGrid.TextMatrix(FGrid.Row, LName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Godown) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Godown) <> rsGod!Name Then
                rsGod.MoveFirst
                rsGod.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Godown) & "'"
            End If
        Case PartSrlNo
            TxtGrid(Index).MaxLength = 20
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then TxtGrid(0).TEXT = TxtGrid(0).Tag: Exit Sub
Select Case FGrid.Col
    Case PNo    '1
        If DGPart.Visible = False Then DGridColSwap DGPart, 0
        DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, 2
            End If
        End If
    Case Godown
        DGridTxtKeyDown DGGod, TxtGrid, 0, rsGod, KeyCode, True, 1, frmGodown, "frmGodown"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, PartSrlNo
            End If
        End If
    Case MRP
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
            End If
        End If
    Case Taxable, PQty, FRate, DisPer, DisOrd, DisRs, DisOrdRs, PartSrlNo, TaxAmt1, FRate, SatAmt1
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 18
            End If
        End If
    Case TaxPer, SatPer
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 1
            End If
        End If
    Case PName
        If DGPart.Visible = False Then DGridColSwap DGPart, 1
        DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 1, frmPartMast, "frmPartMast"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 18
            End If
        End If
    Case LName   '3
        If DGPart.Visible = False Then DGridColSwap DGPart, 2
        DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 18
            End If
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case PNo
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "CODE"
    Case PName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "name"
    Case LName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "Lname"
    Case Godown
        If DGGod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsGod, KeyAscii, "Name"
    Case FRate, DisRs, DisOrdRs
        NumPress TxtGrid(Index), KeyAscii, 8, 2
    Case DisPer, DisOrd, TaxPer
        NumPress TxtGrid(Index), KeyAscii, 2, 2
    Case PQty
        NumPress TxtGrid(Index), KeyAscii, 8, 3
    Case Godown
        If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
            If DGGod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsGod, KeyAscii, "Name"
        End If
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case PNo
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "CODE", True
    Case PName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "name", True
    Case LName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Lname", True
    Case Godown
        If KeyCode <> 13 And DGGod.Visible = False Then
            TxtGrid_KeyDown Index, GridKey, 0
            If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                DGridTxtKeyPress TxtGrid, 0, rsGod, KeyCode, "Name", True
            End If
        End If
    Case MRP, Taxable
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            TxtGrid(Index) = ""
        ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
            TxtGrid(Index) = "Yes"
        Else
            TxtGrid(Index) = "No"
        End If
    Case FRate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
    Case PQty
        FGrid.TextMatrix(FGrid.Row, PQty) = Format(Val(TxtGrid(Index).TEXT), "0.000")
    Case DisPer
        If TxtGrid(Index) <> "" Then
            FGrid.TextMatrix(FGrid.Row, DisPer) = Format(TxtGrid(Index).TEXT, "0.00")
            FGrid.TextMatrix(FGrid.Row, DisRs) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, DisPer) = ""
            FGrid.TextMatrix(FGrid.Row, DisRs) = ""
        End If
    Case DisRs
        If TxtGrid(Index) <> "" Then
            FGrid.TextMatrix(FGrid.Row, DisRs) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, DisPer) = ""
            FGrid.TextMatrix(FGrid.Row, DisRs) = ""
        End If
    Case DisOrd
        If Val(TxtGrid(Index)) <> 0 Then
            FGrid.TextMatrix(FGrid.Row, DisOrd) = Format(TxtGrid(Index).TEXT, "0.00")
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, DisOrd) = ""
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = ""
        End If
    Case DisOrdRs
        If Val(TxtGrid(Index)) <> 0 Then
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Else
           FGrid.TextMatrix(FGrid.Row, DisOrd) = ""
           FGrid.TextMatrix(FGrid.Row, DisOrdRs) = ""
        End If
    Case PartSrlNo
        FGrid.TextMatrix(FGrid.Row, PartSrlNo) = TxtGrid(Index)
End Select
Amt_Cal1
Amt_Cal Index
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
Dim j As Integer
Dim rsTaxPer As ADODB.Recordset
Select Case FGrid.Col
    Case PNo, PName, LName
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        TxtGridValid_PNo
    Case MRP, Taxable
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
            FGrid.TextMatrix(FGrid.Row, FRate) = Format(GetRate(mPartyType, FGrid, CDate(txt(VDate)), FGrid.TextMatrix(FGrid.Row, PNo), MRP, Val(FGrid.TextMatrix(FGrid.Row, MRPRate)), Taxable, Val(FGrid.TextMatrix(FGrid.Row, TBRate)), Val(FGrid.TextMatrix(FGrid.Row, TPRate)), EffectDate, MRPRate), "0.00")
        End If
    Case Godown
        If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, Godown) = ""
            FGrid.TextMatrix(FGrid.Row, God) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
            FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
        End If
    Case FRate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(Index) = "", "", Format(Val(TxtGrid(Index).TEXT), "0.00"))
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
        If PubVATYN = 1 Then
           If txt(FormType).Tag <> "" Then
                Set rsTaxPer = GCn.Execute("Select Tax_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'")
                 If rsTaxPer.RecordCount > 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = rsTaxPer!Tax_Per
                 End If
           End If
        End If
        
        Amt_Cal1
        Amt_Cal Index
    Case PQty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(Index) = "", "", Format(Val(TxtGrid(Index).TEXT), "0.000"))
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
        
        If mVatYn = 1 Then
           If txt(FormType).Tag <> "" Then
                Set rsTaxPer = GCn.Execute("Select Tax_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'")
                 If rsTaxPer.RecordCount > 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = rsTaxPer!Tax_Per
                 End If
           End If
        End If
        
        
        Amt_Cal1
        Amt_Cal Index
        
        If rsGod.RecordCount > 0 And Trim(FGrid.TextMatrix(FGrid.Row, Godown)) = "" Then
            rsGod.MoveFirst
            rsGod.FIND "Code ='" & PubSprCounterGodown & "'"
            FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
            FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
        End If
    Case DisPer
        FGrid.TextMatrix(FGrid.Row, DisPer) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If FGrid.TextMatrix(FGrid.Row, DisPer) = "" Then FGrid.TextMatrix(FGrid.Row, DisRs) = ""
        Amt_Cal1
        Amt_Cal Index
    Case DisRs
        FGrid.TextMatrix(FGrid.Row, DisRs) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisRs)) + Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs)) > Val(FGrid.TextMatrix(FGrid.Row, Amt)) Then
            TxtGridLeave = False: Exit Function
        End If
        Amt_Cal1
        Amt_Cal Index
    Case DisOrd
        FGrid.TextMatrix(FGrid.Row, DisOrd) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) = 0 Then FGrid.TextMatrix(FGrid.Row, DisOrdRs) = 0
        Amt_Cal1
        Amt_Cal Index
    Case DisOrdRs
        FGrid.TextMatrix(FGrid.Row, DisOrdRs) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisRs)) + Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs)) > Val(FGrid.TextMatrix(FGrid.Row, Amt)) Then
            TxtGridLeave = False: Exit Function
        End If
        Amt_Cal1
        Amt_Cal Index
     Case TaxPer
        FGrid.TextMatrix(FGrid.Row, TaxPer) = TxtGrid(0)
        If FGrid.TextMatrix(FGrid.Row, TaxPer) = "" Then FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
        Amt_Cal1
        Amt_Cal
     Case SatPer
        FGrid.TextMatrix(FGrid.Row, SatPer) = TxtGrid(0)
        If FGrid.TextMatrix(FGrid.Row, SatPer) = "" Then FGrid.TextMatrix(FGrid.Row, SatAmt1) = ""
        Amt_Cal1
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
Dim I As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte, Col4 As Byte
    Select Case FGrid.Col
    Case PNo, PName, LName
        Col1 = MRP
        Col2 = Taxable
        Col3 = FGrid.Col
    Case MRP
        Col1 = PNo
        Col2 = Taxable
        Col3 = MRP
    Case Taxable
        Col1 = PNo
        Col2 = MRP
        Col3 = Taxable
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

' Used For Updation of Order in case of Edit and Delete
Private Sub UpdateOrderQty()
Dim Rst As ADODB.Recordset, I As Byte
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select * From SP_Stock Where DocId='" & txt(TxtDocID).TEXT & "'", GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
    While Not Rst.EOF
        If Rst!Order_DocId <> "" Then
            GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty-" & Rst!Qty_Iss & " Where OrderId='" & Rst!Order_DocId & "' and Part_No='" & Rst!Part_No & "'"
        End If
        Rst.MoveNext
    Wend
    End If
Set Rst = Nothing
End Sub

Private Sub RemoveTxtNull()
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).TEXT = IIf(IsNull(txt(I).TEXT), "", txt(I).TEXT)
Next I
End Sub

'************ PRINTING OPTION*************
Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case VType1
        ListArray = Array("Cash", "Credit", "Stock Transfer")
        Set mListItem = ListView_Items(ListView, txtPrint, VType1, ListArray, 3)
    Case FromVno, ToVno
        If IsValid(txt(VType1), "Voucher Type") = False Then Exit Sub
        RsVno.Close
        RsVno.Open "Select V_no as code from Sp_Purch where Sp_Purch.V_Type ='" & txtPrint(VType1).Tag & "'  ", GCn, adOpenDynamic, adLockOptimistic
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
        If txtPrint(VType1).TEXT = "Cash" Then
            txtPrint(VType1).Tag = PRetCashVType
        ElseIf txtPrint(VType1).TEXT = "Credit" Then
            txtPrint(VType1).Tag = PRetCrVType
        Else
            txtPrint(VType1).Tag = TrfRetRecVType
        End If

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
    FGrid.TextMatrix(FGrid.Row, PNo) = ""
    FGrid.TextMatrix(FGrid.Row, PName) = ""
    FGrid.TextMatrix(FGrid.Row, LName) = ""
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        "", "", "", Unit, MRP, Taxable, _
        MRPStkTB, MRPStkTP, TBStk, TPStk, _
        MRPRate, TBRate, TPRate, Bin, _
        HPRate, LPRate, LastRate, PartGrade, _
        EffectDate, DisPer, mCheckNegetiveStockSiteWise, True
Else
    OldPNo = FGrid.TextMatrix(FGrid.Row, PNo)
    FGrid.TextMatrix(FGrid.Row, PNo) = RsPart!Code
    FGrid.TextMatrix(FGrid.Row, PName) = RsPart!Name
    FGrid.TextMatrix(FGrid.Row, LName) = RsPart!LName
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        RsPart!Code, RsPart!Name, RsPart!LName, _
        Unit, MRP, Taxable, _
        MRPStkTB, MRPStkTP, TBStk, TPStk, _
        MRPRate, TBRate, TPRate, Bin, _
        HPRate, LPRate, LastRate, PartGrade, _
        EffectDate, DisPer, mCheckNegetiveStockSiteWise, True
''by LPS 27-04-2K2
'    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> OldPNo Then
'            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(FGrid, CDate(Txt(Vdate).Text), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
''            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
'        End If
'    End If

    If PubVATYN = 1 Then
       If txt(FormType).Tag <> "" Then
            Set rsTaxPer = GCn.Execute("Select Tax_Per, AddTaxPer,L_C from TaxForms where Form_Code='" & txt(FormType).Tag & "'")
             If rsTaxPer.RecordCount > 0 Then
                FGrid.TextMatrix(FGrid.Row, TaxPer) = rsTaxPer!Tax_Per
                FGrid.TextMatrix(FGrid.Row, SatPer) = XNull(rsTaxPer!AddTaxPer)
                
                If UTrim(XNull(rsTaxPer!L_C)) = "LOCAL" Then
                   Set rsTaxPer = GCn.Execute("Select VatPer, AddTaxPer From Part_Grade Where PartGrade_Code='" & FGrid.TextMatrix(FGrid.Row, PartGrade) & "'")
                   If rsTaxPer.RecordCount > 0 Then
                       If VNull(rsTaxPer!VatPer) > 0 Then FGrid.TextMatrix(FGrid.Row, TaxPer) = Format(rsTaxPer!VatPer, "0.00")
                       If VNull(rsTaxPer!AddTaxPer) > 0 Then FGrid.TextMatrix(FGrid.Row, SatPer) = Format(rsTaxPer!AddTaxPer, "0.00")
                   End If
                End If
             End If
       End If
    End If
End If

If FGrid.TextMatrix(FGrid.Rows - 1, PNo) <> "" Then FGrid.AddItem FGrid.Rows
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
    GSQL = "SELECT SP.DocID,SP.V_Type,SP.V_No,SP.V_Date,SP.Cash_Credit,SP.L_C,SP.Party_Code,SG.NamePrefix,SP.Party_Name," & _
        "SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone,SP.Form_Code,SP.Party_Doc_No,SP.Party_Doc_Date," & _
        "SP.GR_RR_No,SP.GR_RR_Date,SP.RoadPermit_No,SP.Tot_No_of_Items,SP.Tot_Doc_Qty,SP.Tot_Phy_Qty," & _
        "SP.OilAmt,SP.SprAmt,SP.Tot_Amt,SP.Tot_Disc_Amt,SP.Tot_Ord_DiscAmt,SP.Tot_Goods_Value,SP.Tax_Amt," & _
        "SP.Addition,SP.Deduction,SP.Net_Amt,SP.Case_No,SP.Case_Mark,SP.Transport, Sp.Transportation, Sp.EntryTaxAmt," & _
        "SP.Supply_Mode,SP.Invoice_DocID,SP.Printed_YN,SP.CancelYN,SP.U_Name,SP.U_EntDt," & _
        "Stk.Srl_No,Part.Part_Name,Stk.Part_No,Stk.Qty_Rec,Stk.Qty_Iss,Stk.Qty_Ret,Stk.Rate," & _
        "Stk.Disc_Per,Stk.Disc_Amt,Stk.Ord_DiscAmt,Stk.Net_Amt as INetAmt,Syctrl.SprPurRetFooter,SP.SatAmt " & _
    "FROM ((((SP_Purch SP LEFT JOIN SP_Stock Stk ON SP.DocID = Stk.DocID) " & _
        "LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1)) " & _
        "LEFT JOIN (SubGroup SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON SP.Party_Code = SG.SubCode) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>SP.U_AE) " & _
    "where SP.docid = '" & Master!SearchCode & "'"

Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "SpPurRet", "SpPurRet")
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(GSQL)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "SpPurRet", "SpPurRet")
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
Private Sub WindowsPrint(Index As Integer, mQry$)
Dim Rst As ADODB.Recordset
Dim I As Integer, RST1 As ADODB.Recordset, Rst2 As ADODB.Recordset
On Error GoTo ERRORHANDLER

        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
        CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
        If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
        Set RST1 = GCn.Execute("select S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
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
                Case UCase("VouType")
                    rpt.FormulaFields(I).TEXT = "'" & TrfRetRecVType & "'"
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
                GCn.Execute "update Sp_purch set Printed_YN = 1  where Sp_purch.docid='" & Master!SearchCode & "'"
            End If
    Case 1  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
    End Select
    
Set Rst = Nothing
Set RST1 = Nothing
Set Rst2 = Nothing
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
        CheckError
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrint(mQry$)
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
    Dim I As Integer, j As Integer
    Dim PrintStr As String
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstPurRet As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mQty As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    
'    Set RstPurRet = GCn.Execute("SELECT City.CityName, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, SubGroup.Phone, Part.Part_Name, Syctrl.SprPurRetFooter,SP_Stock.srl_no, SP_Stock.Part_No, SP_Stock.Qty_Rec, SP_Stock.Qty_Iss, SP_Stock.Qty_Ret, SP_Stock.rate,SP_Stock.Disc_Amt, SP_Stock.Ord_DiscAmt, SP_Stock.Net_Amt as NAmt, SP_Purch.* " & _
        "FROM (((SP_Purch LEFT JOIN SP_Stock ON SP_Purch.DocID = SP_Stock.DocID) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.DocID,1)) LEFT JOIN (SubGroup LEFT JOIN City ON SubGroup.CityCode = City.CityCode) ON SP_Purch.Party_Code = SubGroup.SubCode) LEFT JOIN Syctrl ON  Syctrl.LinkTable  >= SP_Purch.U_AE " & _
        "where SP_Purch.docid = '" & Master!searchcode & "'")
    
    Set RstPurRet = GCn.Execute(mQry)
    If RstPurRet.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select SprPurRetFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
    
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 16
    mFooter = mFooter + FooterCnt
    
      
    'Sale Bill Header
      
      mDocStr = IIf(mVType = TrfRetRecVType, "TRANSFER RETURN", "PURCHASE RETURN")
      mDupStr = IIf(RstPurRet!Printed_YN = 1, "(DUPLICATE)", "")
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
        Print #1, PSTR(RstPurRet!NamePrefix & " " & RstPurRet!Party_Name, 40) & Space(1) & PSTR("Pur. Return No.", 16) & " : " & PSTR(STR(RstPurRet!V_NO), 14) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstPurRet!Add1), 40) & Space(1) & mEmph & PSTR("Pur. Return Date", 16) & " : " & PSTR(STR(RstPurRet!V_Date), 14) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstPurRet!Add2), 40) & Space(1) & PSTR("Party Doc. No", 16) & " : " & PSTR(XNull(RstPurRet!Party_Doc_No), 14)
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstPurRet!Add3) & IIf(XNull(RstPurRet!CityName) <> "" And XNull(RstPurRet!Add3) <> "", ",", "") & XNull(RstPurRet!CityName), 40) & Space(1) & PSTR("Party Doc. Date", 16) & " : " & STR(RstPurRet!Party_Doc_Date)
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
        If RstPurRet.RecordCount > 0 Then
            Do Until RstPurRet.EOF
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
                    Print #1, PSTR(RstPurRet!NamePrefix & " " & RstPurRet!Party_Name, 40) & Space(1) & PSTR("Pur. Return No.", 16) & " : " & PSTR(STR(RstPurRet!V_NO), 14) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, PSTR(XNull(RstPurRet!Add1), 40) & Space(1) & mEmph & PSTR("Pur. Return Date", 16) & " : " & PSTR(STR(RstPurRet!V_Date), 14) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, PSTR(XNull(RstPurRet!Add2), 40)
                    mHeader = mHeader + 1
                    Print #1, XNull(RstPurRet!Add3) & IIf(XNull(RstPurRet!CityName) <> "" And XNull(RstPurRet!Add3) <> "", ",", "") & XNull(RstPurRet!CityName)
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
                
                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & mChr17 & PSTR(RstPurRet!Part_No, 27, , AlignLeft) & PSTR(RstPurRet!Part_Name, 48) & mChr18 & PSTR(RstPurRet!Qty_Iss, 10, 3) & PSTR(RstPurRet!Rate, 9, 2) & PSTR(RstPurRet!INetAmt, 10, 2)
                mQty = mQty + RstPurRet!Qty_Iss: mAmount = mAmount + RstPurRet!INetAmt
            Print #1, PrintStr
            RstPurRet.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    RstPurRet.MoveFirst
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, PSTR("Total  > > ", 51, , AlignRight) & PSTR(mQty, 10, 3) & Space(9) & PSTR(mAmount, 10, 2)
    Print #1, Replace(Space(PageWidth), " ", "-")
    
    Print #1, Space(38) & " | " & PSTR("Total Goods Value", 22) & " : " & PSTR(RstPurRet!Tot_Goods_Value, 13, 2)
    Print #1, Space(38) & " | " & PSTR("Tax Amount", 22) & " : " & PSTR(RstPurRet!Tax_Amt, 13, 2)
     Print #1, Space(38) & " | " & PSTR("Additional Tax", 22) & " : " & PSTR(RstPurRet!SatAmt, 13, 2)
    
    Print #1, Space(38) & " | " & PSTR("Addition", 22) & " : " & PSTR(RstPurRet!Addition, 13, 2)
    Print #1, Space(38) & " | " & PSTR("Deduction", 22) & " : " & PSTR(RstPurRet!Deduction, 13, 2)
    
    Print #1, Space(38) & " | " & mEmph & PSTR("Net Payble Rs.", 22) & " : " & PSTR(RstPurRet!Net_Amt, 13, 2) & mEmph1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(RstPurRet!Net_Amt, "Rupees", "Paise") & mDoub1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, "Prepared By " & RstPurRet!U_Name & " Dt. " & STR(RstPurRet!U_EntDt) & mEmph & PSTR("For " & PubComp_Name, PageWidth - Len("Prepared By " & RstPurRet!U_Name & " Dt. " & STR(RstPurRet!U_EntDt)), , AlignRight) & mEmph1 & mDoub
    Print #1, ""
    Print #1, "Terms And Condition : " & PSTR("Autorised Signatury", PageWidth - Len("Terms And Condition :"), , AlignRight) & mDoub1 & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
     Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt  > Prn"
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
        GCn.Execute "update Sp_purch set Printed_YN = 1  where Sp_purch.docid='" & Master!SearchCode & "'"
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub ProcAcPost(rsCtrlAc As ADODB.Recordset)
On Error GoTo lblExit
'Dim xNetAmt As Double, xEntryTaxAmt As Double, xTransportation As Double
''A/c Posting related declarations
'Dim LedgAry() As LedgRec, mCommNarr$
'Dim mResult As Byte, mNarr$, I As Integer
'Dim PartyCode$, mFADocID$
'
'    If txt(VType) = "Cash" Then
'        xEntryTaxAmt = VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & "' and CancelYN=0 ").Fields(0).Value)
'        xTransportation = VNull(GCn.Execute("select sum(Transportation) from SP_Purch where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & "' and CancelYN=0 ").Fields(0).Value)
'        GSQL = "select TF.PurSal_Ac_Code,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt " & _
'            "from SP_Purch " & _
'            "Left join TaxFormsAc TF on SP_Purch.Form_Code+'" & PubDivCode & "' =TF.Form_Code+TF.Div_Code " & _
'            "where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & _
'            "' group by TF.PurSal_Ac_Code"
'        mNarr = "Through Spare Cash Purchase Return (Daily Posting)"
'        mCommNarr = mNarr & " [Common]"
'        'Undelete old Posting (individual if any)
''        LedgerUnPost GCnFaS, Txt(TxtDocId)
'        'Create FA DocID for Daily Posting
'        mFADocID = left(txt(TxtDocID), 8) & "QQQQQ" & "  " & Format(PubStartDate, "yy") & Format(txt(VDate), "mmdd")
'        PartyCode = PubSprCashAc
'    Else
'        PartyCode = txt(Party).Tag
'        mFADocID = txt(TxtDocID)
'        mNarr = "Through Spare Cr Purchase Return"
'        mCommNarr = mNarr & " [Common]"
'        xEntryTaxAmt = Val(txt(EntryTaxAmt)) ' VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where docid='" & Txt(TxtDocId) & "' and CancelYN=0 ").Fields(0).Value)
'        xTransportation = Val(txt(Transportation)) ' VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where docid='" & Txt(TxtDocId) & "' and CancelYN=0 ").Fields(0).Value)
'        GSQL = "select TF.PurSal_Ac_Code,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt " & _
'            "from SP_Purch " & _
'            "Left Join TaxFormsAc TF on SP_Purch.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
'            "where docid='" & txt(TxtDocID) & _
'            "' Group by TF.PurSal_Ac_Code"
'    End If
'    Set GRs = New ADODB.Recordset
'    GRs.CursorLocation = adUseClient
'    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
'
'    '*** pURCHASE Amount Row
''   0.Purchase A/c
''   1.Party A/c or Cash A/c
''*********
'    ReDim Preserve LedgAry(1)
'    Do While GRs.EOF = False
'        If IsNull(GRs!PurSal_Ac_Code) Or GRs!PurSal_Ac_Code = "" Then
'            MsgBox "Please Define Purchase A/c in Tax Forms " & GRs!PurSal_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
'            GoTo lblExit
'        End If
'        I = UBound(LedgAry) + 1
'        ReDim Preserve LedgAry(I)
'        LedgAry(I).SubCode = GRs!PurSal_Ac_Code
'        LedgAry(I).AmtCr = IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
'        LedgAry(I).Narration = mNarr
'        LedgAry(I).ContraSub = PartyCode
'
'        xNetAmt = xNetAmt + IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
'        GRs.MoveNext
'    Loop
'    If xTransportation <> 0 Then
'        I = UBound(LedgAry) + 1
'        ReDim Preserve LedgAry(I)
'        LedgAry(I).SubCode = rsCtrlAc!SprPurTrans_Ac
'        LedgAry(I).AmtDr = xTransportation
'        LedgAry(I).Narration = mNarr
'    End If
'
'    If xEntryTaxAmt <> 0 Then
'        I = UBound(LedgAry) + 1
'        ReDim Preserve LedgAry(I)
'        LedgAry(I).SubCode = rsCtrlAc!EntryTax_Ac
'        LedgAry(I).AmtDr = xEntryTaxAmt
'        LedgAry(I).Narration = mNarr
'    End If
'    I = UBound(LedgAry) + 1
'    ReDim Preserve LedgAry(I)
'    LedgAry(I).SubCode = PartyCode
'    LedgAry(I).AmtDr = (xNetAmt - (xEntryTaxAmt + xTransportation))
'    LedgAry(I).Narration = mNarr
'
'    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, mFADocID, CDate(txt(VDate)), mCommNarr)
'    If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
'lblExit:
'    Set GRs = Nothing
'    If err.NUMBER <> 0 Then MsgBox err.Description


On Error GoTo lblExit
Dim xNetAmt As Double, xEntryTaxAmt As Double, xTransportation As Double, TransAc$
'A/c Posting related declarations
Dim LedgAry() As LedgRec, mCommNarr$, ContraCodeCr$
Dim mResult As Byte, mNarr$, TaxSQL$, I As Double, j As Double
Dim mSprPurPfx$, mFADocID$, PartyCode$

    mSprPurPfx = "PPPPP"
    If txt(VType) = "Cash" Then
        xTransportation = VNull(GCn.Execute("select sum(Transportation) from SP_Purch where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & "' and CancelYN=0 ").Fields(0).Value)
        xEntryTaxAmt = VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & "' and CancelYN=0 ").Fields(0).Value)
        GSQL = "select TF.PurSal_Ac_Code,TF.Tax_Ac_Code, AddTaxAc,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt,sum(Tax_Amt) as TaxAmt, Sum(SatAmt) As SatAmt,TaxForms.L_C " & _
            "from (SP_Purch " & _
            "left join TaxFormsAc as TF on SP_Purch.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code) Left Join TaxForms on SP_Purch.Form_Code=TaxForms.Form_Code  " & _
            "where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & _
            "' and " & vIsNull("CancelYN", "0") & "=0 Group by TF.PurSal_Ac_Code,TF.Tax_Ac_Code, tf.AddTaxAc,TaxForms.L_C"
        mNarr = "Through Spare Cash Purchase (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        'Undelete old Posting (individual if any)
        'LedgerUnPost GCnFaS, Txt(TxtDocId)
        'Create FA DocID for Daily Posting
        mFADocID = left(txt(TxtDocID), 8) & mSprPurPfx & "  " & Format(PubStartDate, "yy") & Format(txt(VDate), "mmdd")
        PartyCode = PubSprCashAc
    Else
        PartyCode = txt(Party).Tag
        mFADocID = txt(TxtDocID)
        mNarr = "Cr Purchase "
        If txt(SuppChlNo) <> "" Then
            mNarr = mNarr & " Party Document No." & txt(SuppChlNo)
        End If
        If txt(SuppChlDate) <> "" Then
            mNarr = mNarr & " Date " & txt(SuppChlDate)
        End If
        mCommNarr = mNarr & " [Common]"
        xEntryTaxAmt = Val(txt(EntryTaxAmt)) ' VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where docid='" & Txt(TxtDocId) & "' and CancelYN=0 ").Fields(0).Value)
        xTransportation = Val(txt(Transportation)) ' VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where docid='" & Txt(TxtDocId) & "' and CancelYN=0 ").Fields(0).Value)
        GSQL = "select TF.PurSal_Ac_Code,TF.Tax_Ac_Code, Tf.AddTaxAc,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt,sum(Tax_Amt) as TaxAmt, Sum(SatAmt) As SatAmt,TaxForms.L_C " & _
            "from (SP_Purch " & _
            "left join TaxFormsAc as TF on SP_Purch.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code ) Left Join TaxForms on SP_Purch.Form_Code=TaxForms.Form_Code  " & _
            "where docid='" & txt(TxtDocID) & _
            "' group by TF.PurSal_Ac_Code,TF.Tax_Ac_Code, Tf.AddTaxAc,TaxForms.L_C"
    End If
    
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    '*** pURCHASE Amount Row
'   0.Purchase A/c
'   1.Party A/c or Cash A/c
'    Dim LedgAry() As LedgRec
    
'*********
    I = -1
    If I = -1 Then
         ReDim Preserve LedgAry(1)
         I = 0
    Else
         I = UBound(LedgAry) + 1
         ReDim Preserve LedgAry(I)
    End If
    Do While GRs.EOF = False
        If PubVATYN = 1 And XNull(GRs!L_C) = "Local" Then
            If IsNull(GRs!PurSal_Ac_Code) Or GRs!PurSal_Ac_Code = "" Then
                MsgBox "Please Define Purchase A/c in Tax Forms " & GRs!PurSal_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                GoTo lblExit
            End If
            If IsNull(GRs!Tax_Ac_Code) Or GRs!Tax_Ac_Code = "" Then
                MsgBox "Please Define Tax A/c in Tax Forms " & GRs!Tax_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                GoTo lblExit
            End If
            If mSatYn And XNull(GRs!L_C) = "Local" Then
                If IsNull(GRs!AddTaxAc) Or GRs!AddTaxAc = "" Then
                    MsgBox "Please Define Additional Tax A/c in Tax Forms " & XNull(GRs!AddTaxAc) & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                    GoTo lblExit
                End If
            End If
            If I = -1 Then
                ReDim Preserve LedgAry(1)
                I = 0
            Else
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
            End If
            LedgAry(I).SubCode = GRs!PurSal_Ac_Code
            LedgAry(I).AmtCr = VNull(GRs!NetAmt) - VNull(GRs!TaxAmt) - VNull(GRs!SatAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode
 
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            
            LedgAry(I).SubCode = GRs!Tax_Ac_Code
            LedgAry(I).AmtCr = VNull(GRs!TaxAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode

            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            
            LedgAry(I).SubCode = XNull(GRs!AddTaxAc)
            LedgAry(I).AmtCr = VNull(GRs!SatAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode

            xNetAmt = xNetAmt + IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
            
        Else
            If IsNull(GRs!PurSal_Ac_Code) Or GRs!PurSal_Ac_Code = "" Then
                MsgBox "Please Define Purchase A/c in Tax Forms " & GRs!PurSal_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                GoTo lblExit
            End If
            If I = -1 Then
                ReDim Preserve LedgAry(1)
                I = 0
            Else
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
            End If
            LedgAry(I).SubCode = GRs!PurSal_Ac_Code
            LedgAry(I).AmtCr = IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode
            
            xNetAmt = xNetAmt + IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
        End If
        GRs.MoveNext
    Loop
    If xTransportation <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprPurTrans_Ac
        LedgAry(I).AmtDr = xTransportation
        LedgAry(I).Narration = mNarr
    End If
    If xEntryTaxAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!EntryTax_Ac
        LedgAry(I).AmtDr = xEntryTaxAmt
        LedgAry(I).Narration = mNarr
        
    End If
    I = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(I)
    LedgAry(I).SubCode = PartyCode
    LedgAry(I).AmtDr = (xNetAmt - (xEntryTaxAmt + xTransportation))
    LedgAry(I).Narration = mNarr

    
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, mFADocID, CDate(txt(VDate)), mCommNarr)
    If mResult <> 1 Then
        MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    End If
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
    End If

End Sub




Sub Ini_Pub()
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl")
    If RsTemp.RecordCount > 0 Then
        mCheckNegetiveStockSiteWise = VNull(RsTemp!CheckNegetiveStockSiteWise)
    End If
End Sub



Sub DispText_Vat()
    With FGrid
        If PubVATYN = 1 Then
            .TextMatrix(0, TaxPer) = "TaxPer"
            .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
            .ColWidth(TaxPer) = 840
            
            .TextMatrix(0, TaxAmt1) = "TaxAmt"
            .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
            .ColWidth(TaxAmt1) = 840
            
            If mSatYn Then
                .TextMatrix(0, SatPer) = "Sat %"
                .ColAlignmentFixed(SatPer) = flexAlignRightCenter
                .ColWidth(SatPer) = 840
                
                .TextMatrix(0, SatAmt1) = "Sat Amt"
                .ColAlignmentFixed(SatAmt1) = flexAlignRightCenter
                .ColWidth(SatAmt1) = 840
            Else
                .ColWidth(SatPer) = 0
                .ColWidth(SatAmt1) = 0
            End If
        Else
            .TextMatrix(0, TaxPer) = ""
            .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
            .ColWidth(TaxPer) = 0
        
            .TextMatrix(0, TaxAmt1) = ""
            .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
            .ColWidth(TaxAmt1) = 0
            
            .ColWidth(SatPer) = 0
            .ColWidth(SatAmt1) = 0
        End If
    
    End With
End Sub

