VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmCustRect 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "Customer Transaction Entry"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10890
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
   ScaleHeight     =   5250
   ScaleWidth      =   10890
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox TxtGrid2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDF4B5&
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
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   7725
      TabIndex        =   103
      Top             =   3330
      Visible         =   0   'False
      Width           =   975
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
      Index           =   37
      Left            =   1635
      TabIndex        =   19
      Text            =   "RectCatg"
      Top             =   2280
      Width           =   4305
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   36
      Left            =   1635
      TabIndex        =   32
      Text            =   "99999999.99"
      Top             =   3360
      Width           =   1245
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   35
      Left            =   1635
      TabIndex        =   31
      Top             =   3090
      Width           =   4305
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   330
      Left            =   7320
      TabIndex        =   96
      Top             =   15
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6615
      Visible         =   0   'False
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
      Index           =   34
      Left            =   7035
      TabIndex        =   27
      Top             =   6615
      Visible         =   0   'False
      Width           =   1080
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   17
      Left            =   1635
      TabIndex        =   20
      Top             =   2550
      Width           =   4305
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   255
      Index           =   28
      Left            =   7035
      TabIndex        =   21
      Top             =   5805
      Visible         =   0   'False
      Width           =   1080
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
      Left            =   7035
      TabIndex        =   23
      Text            =   "999999.99"
      Top             =   6075
      Visible         =   0   'False
      Width           =   1080
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
      Index           =   32
      Left            =   7035
      TabIndex        =   25
      Top             =   6345
      Visible         =   0   'False
      Width           =   1080
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
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6345
      Visible         =   0   'False
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
      Index           =   29
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6075
      Visible         =   0   'False
      Width           =   510
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
      Left            =   9495
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Text            =   "VFa"
      Top             =   4770
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
      Index           =   13
      Left            =   9495
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Text            =   "0123456789"
      Top             =   4500
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   3825
      Left            =   4260
      Negotiate       =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   7500
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   6747
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
         Caption         =   "City Name"
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
            ColumnWidth     =   4410.142
         EndProperty
      EndProperty
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
      Left            =   225
      TabIndex        =   70
      Top             =   6930
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
         Picture         =   "frmCustRect.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   80
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
         Picture         =   "frmCustRect.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmCustRect.frx":0678
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
         TabIndex        =   78
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmCustRect.frx":0982
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
         TabIndex        =   77
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmCustRect.frx":0C8C
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   4440
      Negotiate       =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   7230
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
   Begin MSDataGridLib.DataGrid DGVType 
      Height          =   2175
      Left            =   2415
      Negotiate       =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   7185
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
         DataField       =   "Name"
         Caption         =   "Voucher Type"
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
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -30
      TabIndex        =   66
      Top             =   7005
      Visible         =   0   'False
      Width           =   1920
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   30
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   120
         Width           =   1860
         _ExtentX        =   3281
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
         BackColor       =   12640511
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
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3825
      Left            =   330
      Negotiate       =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   6747
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
      ColumnCount     =   2
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
         DataField       =   "Father"
         Caption         =   "Father / Husband Name"
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
            ColumnWidth     =   5520.189
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   6765
      Negotiate       =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   7440
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
   Begin MSDataGridLib.DataGrid DGBook 
      Height          =   2085
      Left            =   -30
      Negotiate       =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   7125
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3678
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Site_Desc"
         Caption         =   "Site"
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
         DataField       =   "Code"
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
      BeginProperty Column02 
         DataField       =   "Ord_Date"
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
      BeginProperty Column03 
         DataField       =   "MODEL"
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
      BeginProperty Column04 
         DataField       =   "PaidAmt"
         Caption         =   "AmtPaid"
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
         DataField       =   "Inv_No"
         Caption         =   "Invoice No"
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
         DataField       =   "Inv_dt"
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
      BeginProperty Column07 
         DataField       =   "Net_AMOUNT"
         Caption         =   "Inv. Amt"
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
      BeginProperty Column08 
         DataField       =   "FinName"
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
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   4424.882
         EndProperty
      EndProperty
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
      TabIndex        =   30
      Text            =   "23-MAR-2002"
      Top             =   2820
      Width           =   1245
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
      Left            =   4575
      MaxLength       =   10
      TabIndex        =   29
      Top             =   2820
      Width           =   1365
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
      Index           =   5
      Left            =   6795
      TabIndex        =   6
      Text            =   "RectCatg"
      Top             =   5535
      Visible         =   0   'False
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
      Index           =   2
      Left            =   8580
      TabIndex        =   3
      Top             =   1305
      Width           =   2085
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
      Left            =   5715
      TabIndex        =   10
      Top             =   930
      Width           =   1365
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
      Left            =   1635
      TabIndex        =   7
      Top             =   660
      Width           =   1935
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
      Index           =   7
      Left            =   5715
      TabIndex        =   8
      Top             =   660
      Width           =   1365
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
      Index           =   23
      Left            =   1635
      MaxLength       =   40
      TabIndex        =   35
      Top             =   3900
      Width           =   5445
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
      Left            =   1815
      TabIndex        =   15
      Top             =   5835
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   18
      Left            =   1635
      TabIndex        =   28
      Text            =   "99999999.99"
      Top             =   2820
      Width           =   1245
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
      Left            =   1635
      MaxLength       =   100
      TabIndex        =   33
      Top             =   3630
      Width           =   5445
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
      Left            =   1635
      TabIndex        =   18
      Top             =   2010
      Width           =   5460
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   3435
      TabIndex        =   13
      Top             =   1470
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
      Index           =   15
      Left            =   1635
      TabIndex        =   14
      Top             =   1740
      Width           =   2475
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
      Left            =   1635
      TabIndex        =   9
      Top             =   930
      Width           =   1935
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   1635
      TabIndex        =   12
      Text            =   " "
      Top             =   1470
      Width           =   1185
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
      Left            =   8580
      TabIndex        =   2
      Top             =   1035
      Width           =   2085
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
      Index           =   27
      Left            =   5730
      TabIndex        =   17
      Top             =   1740
      Width           =   1365
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   24
      Left            =   1635
      MaxLength       =   40
      TabIndex        =   36
      Top             =   4170
      Width           =   5445
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
      Index           =   26
      Left            =   1815
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   1470
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
      Index           =   4
      Left            =   9345
      TabIndex        =   5
      Top             =   1845
      Width           =   1320
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   1635
      TabIndex        =   11
      Top             =   1200
      Width           =   5445
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
      Left            =   8580
      TabIndex        =   1
      Top             =   540
      Width           =   2100
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
      Left            =   3870
      TabIndex        =   34
      Top             =   5505
      Visible         =   0   'False
      Width           =   495
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
      Index           =   3
      Left            =   8580
      TabIndex        =   4
      Top             =   1575
      Width           =   2085
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid2 
      Height          =   1980
      Left            =   7350
      TabIndex        =   104
      Top             =   2370
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   3493
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   13623520
      ForeColorFixed  =   128
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   32896
      FocusRect       =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "ddd"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks  From"
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
      Index           =   24
      Left            =   195
      TabIndex        =   102
      Top             =   4170
      Width           =   1140
   End
   Begin VB.Label Label3 
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
      Index           =   16
      Left            =   2895
      TabIndex        =   101
      Top             =   1485
      Width           =   390
   End
   Begin VB.Label LblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
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
      Left            =   210
      TabIndex        =   100
      Top             =   4455
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card No."
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
      Index           =   14
      Left            =   210
      TabIndex        =   99
      Top             =   2295
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. A/c"
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
      Index           =   11
      Left            =   210
      TabIndex        =   98
      Top             =   3105
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Amt (Rs.)"
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
      Index           =   10
      Left            =   210
      TabIndex        =   97
      Top             =   3375
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Turn OverTax  @"
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
      Left            =   5055
      TabIndex        =   95
      Top             =   6630
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Amount"
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
      Left            =   5700
      TabIndex        =   94
      Top             =   5820
      Visible         =   0   'False
      Width           =   1275
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
      Left            =   5055
      TabIndex        =   93
      Top             =   6090
      Visible         =   0   'False
      Width           =   1365
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
      Left            =   5055
      TabIndex        =   92
      Top             =   6360
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label LblCancel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Cancelled*"
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
      Left            =   3930
      TabIndex        =   91
      Top             =   480
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label LblPartyBal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bal."
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
      Left            =   5130
      TabIndex        =   90
      Top             =   1485
      Width           =   315
   End
   Begin VB.Label lblAcBal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bal."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5985
      TabIndex        =   89
      Top             =   2550
      Width           =   270
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
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   9060
      TabIndex        =   88
      Top             =   4785
      Width           =   390
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
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   8295
      TabIndex        =   87
      Top             =   4515
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Trn Catg."
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
      Index           =   0
      Left            =   5820
      TabIndex        =   84
      Top             =   5550
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DD/Chq No. && Date"
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
      Index           =   22
      Left            =   2970
      TabIndex        =   63
      Top             =   2835
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party A/c"
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
      Index           =   3
      Left            =   210
      TabIndex        =   62
      Top             =   1215
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Date"
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
      Index           =   27
      Left            =   7425
      TabIndex        =   61
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Trn Y/N"
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
      Index           =   18
      Left            =   4425
      TabIndex        =   60
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received With"
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
      Index           =   8
      Left            =   210
      TabIndex        =   57
      Top             =   3915
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAN/GIR No."
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
      Index           =   45
      Left            =   120
      TabIndex        =   56
      Top             =   5850
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount (Rs.)"
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
      Index           =   34
      Left            =   210
      TabIndex        =   55
      Top             =   2835
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drawn On/Narr."
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
      Left            =   210
      TabIndex        =   54
      Top             =   3645
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger A/c"
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
      Index           =   30
      Left            =   210
      TabIndex        =   53
      Top             =   2565
      Width           =   915
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   21
      Left            =   210
      TabIndex        =   52
      Top             =   2025
      Width           =   765
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   20
      Left            =   210
      TabIndex        =   51
      Top             =   1755
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Location"
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
      Index           =   23
      Left            =   210
      TabIndex        =   50
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No."
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
      Index           =   19
      Left            =   210
      TabIndex        =   49
      Top             =   1485
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provisional No."
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
      Index           =   5
      Left            =   210
      TabIndex        =   48
      Top             =   675
      Width           =   1245
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   5235
      TabIndex        =   47
      Top             =   675
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name"
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
      Index           =   15
      Left            =   7425
      TabIndex        =   46
      Top             =   1065
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Declaration Under"
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
      Index           =   13
      Left            =   4200
      TabIndex        =   45
      Top             =   1755
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printing of Party A/c Name on Receipt Y/N"
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
      Index           =   32
      Left            =   345
      TabIndex        =   44
      Top             =   5520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Circle Ward No."
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
      Index           =   9
      Left            =   510
      TabIndex        =   43
      Top             =   6135
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1755
      Left            =   7335
      Top             =   465
      Width           =   3465
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
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
      Left            =   8580
      TabIndex        =   42
      Top             =   1845
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr No."
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
      Index           =   1
      Left            =   7425
      TabIndex        =   41
      Top             =   1845
      Width           =   510
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
      Left            =   7425
      TabIndex        =   40
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
      Left            =   9240
      TabIndex        =   39
      Top             =   795
      Width           =   990
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
      Left            =   7425
      TabIndex        =   38
      Top             =   555
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Type"
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
      Index           =   2
      Left            =   7425
      TabIndex        =   37
      Top             =   1590
      Width           =   600
   End
End
Attribute VB_Name = "frmCustRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PubReceiptType As String
Dim RsSite As ADODB.Recordset
Dim RSBook As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsVType As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim DocID As String * 21
Dim BookDocId As String * 21
Public mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String



Private Const TxtDocID As Byte = 0
Private Const SiteCode As Byte = 1
Private Const VDate As Byte = 2
Private Const VType As Byte = 3
Private Const SerialNo As Byte = 4
Private Const RectCatg As Byte = 5
Private Const ProNo As Byte = 6
Private Const ProDate As Byte = 7
Private Const RecLocation As Byte = 8
Private Const VehTrnYN As Byte = 9
Private Const Party As Byte = 10
Private Const BookNo As Byte = 11
Private Const BookDate As Byte = 12
Private Const Model As Byte = 15
Private Const FB_Code  As Byte = 16
Private Const AcHead As Byte = 17
Private Const Amt  As Byte = 18
Private Const DDNo As Byte = 19
Private Const DDDate  As Byte = 20
Private Const Narr As Byte = 21
Private Const PrnNameYN As Byte = 22
Private Const Rem1 As Byte = 23
Private Const Rem2 As Byte = 24
Private Const PanNo As Byte = 25

Private Const CircleNo As Byte = 26
Private Const FormType As Byte = 27
Private Const VehAmt As Byte = 28
Private Const TaxPer As Byte = 29
Private Const TaxAmt As Byte = 30
Private Const TaxSurPer As Byte = 31
Private Const TaxSurch As Byte = 32
Private Const TOTPer As Byte = 33
Private Const TOTAmt As Byte = 34
Private Const DiscAcName As Byte = 35
Private Const DiscAmt As Byte = 36
Private Const CreditCardNo As Byte = 37

Private Const AcPostByName As Byte = 13
Private Const AcPostDate As Byte = 14

Private Const SiteCode1 As Byte = 0
Private Const VType1 As Byte = 1
Private Const FromVno As Byte = 2
Private Const ToVno As Byte = 3

Dim ListArray As Variant
Dim mListItem As ListItem
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim mTaxAcHead$, mTaxSurAcHead$, mTOTAcHead$

Private Const CnsPrinter As Byte = 0
Private Const CnsScreen As Byte = 1
Private Const CnsPrinterSetup As Byte = 2
Private Const CnsClose As Byte = 3
Private Const mCustBR$ = "G_ABR"  'Customer Bank Receipt
Private Const mCustCR$ = "G_ACR"  'Customer Cash Receipt
Private Const mCustBP$ = "G_BBP"  'Customer Bank Payment
Private Const mCustCP$ = "G_BCP"  'Customer Cash Payment
Private Const mCustCRN$ = "G_CRN" 'Credit Note
Private Const mCustDRN$ = "G_DRN" 'Debit Note
Private Const mTelcoRct$ = "G_TLR" 'Telco Receipt


Private Const Col_Code   As Byte = 0
Private Const Col_ChqNo   As Byte = 1
Private Const Col_ChqDate    As Byte = 2
Private Const Col_ChqAmt    As Byte = 3

Dim TAddMode        As Boolean
Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Dim GridKey As Integer
Dim mMultipleChqNo As Boolean


Private Sub cmdPost_Click()
Dim I As Integer, j As Integer
Dim LedgAry(4) As LedgRec, mNarr$, mResult As Byte

    Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        If CDate(Format(txt(VDate).TEXT, "dd/MMM/yyyy")) < CDate("01/Apr/2005") Then
            'MsgBox ""
        End If
        If CDate(Format(txt(VDate).TEXT, "dd/MMM/yyyy")) < PubStartDate Or CDate(Format(txt(VDate).TEXT, "dd/MMM/yyyy")) > PubEndDate Then GoTo MyNextRecord
        Call TopCtrl1_eEdit
        'A/c Posting
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            Select Case txt(VType).Tag
                Case "G_ABR", "G_ACR", "G_TLR", "G_JV"   'Receipt
                    I = -1
                    
                    If Not mMultipleChqNo Then
                        I = I + 1
                        LedgAry(I).SubCode = txt(AcHead).Tag
                        LedgAry(I).AmtDr = Val(txt(Amt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = txt(Party).Tag
                    Else
                        For j = 0 To FGrid2.Rows - 1
                            I = I + 1
                            
                            LedgAry(I).SubCode = txt(AcHead).Tag
                            LedgAry(I).AmtDr = Val(FGrid2.TextMatrix(j, Col_ChqAmt))
                            LedgAry(I).Chq_No = FGrid2.TextMatrix(j, Col_ChqNo)
                            LedgAry(I).Chq_Date = FGrid2.TextMatrix(j, Col_ChqDate)
                            LedgAry(I).Narration = mNarr
                            LedgAry(I).ContraSub = txt(Party).Tag
                            
                        Next j
                    End If
                    I = I + 1
                    LedgAry(I).SubCode = txt(Party).Tag
                    LedgAry(I).AmtCr = Val(txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = txt(AcHead).Tag
                Case "G_BBP", "G_BCP", "G_DRN"   'payment
                    I = 0
                    LedgAry(I).SubCode = txt(Party).Tag
                    LedgAry(I).AmtDr = Val(txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = txt(AcHead).Tag
                    
                    
                    If Not mMultipleChqNo Then
                    
                        I = I + 1
                        LedgAry(I).SubCode = txt(AcHead).Tag
                        LedgAry(I).AmtCr = Val(txt(Amt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = txt(Party).Tag
                    Else
                        For j = 0 To FGrid2.Rows - 1
                            I = I + 1
                            
                            LedgAry(I).SubCode = txt(AcHead).Tag
                            LedgAry(I).AmtCr = Val(FGrid2.TextMatrix(j, Col_ChqAmt))
                            LedgAry(I).Chq_No = FGrid2.TextMatrix(j, Col_ChqNo)
                            LedgAry(I).Chq_Date = FGrid2.TextMatrix(j, Col_ChqDate)
                            LedgAry(I).Narration = mNarr
                            LedgAry(I).ContraSub = txt(Party).Tag
                            
                        Next j
                    
                    End If
                    
                Case mCustCRN
                    I = 0
                    LedgAry(I).SubCode = txt(Party).Tag
                    LedgAry(I).AmtCr = Val(txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = txt(AcHead).Tag
                    I = I + 1
                    LedgAry(I).SubCode = txt(AcHead).Tag
                    If Val(txt(VehAmt)) + Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(TOTAmt)) = 0 Then
                        LedgAry(I).AmtDr = Val(txt(Amt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = txt(Party).Tag
                    Else
                        LedgAry(I).AmtDr = Val(txt(VehAmt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
        
                        I = I + 1
                        LedgAry(I).SubCode = mTaxAcHead
                        LedgAry(I).AmtDr = Val(txt(TaxAmt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
                        I = I + 1
                        LedgAry(I).SubCode = mTaxSurAcHead
                        LedgAry(I).AmtDr = Val(txt(TaxSurch))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
                        I = I + 1
                        LedgAry(I).SubCode = mTOTAcHead
                        LedgAry(I).AmtDr = Val(txt(TOTAmt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
                    End If
            End Select
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(TxtDocID), CDate(txt(VDate)), mNarr)
            If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        End If
        'eof posting
MyNextRecord:
        Disp_Text SETS("INI", Me, Master)
        Master.MoveNext
    Loop

End Sub

Private Sub DGCity_Click()
    DgCity.Visible = False
    If RsCity.RecordCount > 0 Then
        txt(RecLocation).Tag = RsCity!Code
        txt(RecLocation).TEXT = RsCity!Name
    End If
    txt(RecLocation).SetFocus
End Sub

Private Sub DGSite_Click()
If FrmPrn.Visible = False Then
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        txt(SiteCode).TEXT = RsSite!Name
        txt(SiteCode).Tag = RsSite!Code
    End If
    txt(SiteCode).SetFocus
Else
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        txtPrint(SiteCode1).TEXT = RsSite!Name
        txtPrint(SiteCode1).Tag = RsSite!Code
    End If
    txtPrint(SiteCode1).SetFocus
End If
End Sub

Private Sub DGBook_Click()
    DGBook.Visible = False
    If RSBook.RecordCount > 0 Then
        txt(BookNo).TEXT = RSBook!Code
        FillRecords RSBook
    End If
    txt(BookNo).SetFocus
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
    End If
    DGParty.Visible = False
    txt(Party).SetFocus
End Sub

Private Sub DGVno_Click()
Dim Index As Integer
If DGVno.Tag = "1" Then
    Index = ToVno
Else
    Index = FromVno
End If
    DGVno.Visible = False
    If RsVno.RecordCount > 0 Then
        txtPrint(Index).TEXT = RsVno!Code
    End If
    txtPrint(Index).SetFocus
End Sub

Private Sub DGVType_Click()
If FrmPrn.Visible = False Then
    If RsVType.RecordCount > 0 Then
        txt(VType).TEXT = RsVType!Name
        txt(VType).Tag = RsVType!Code
    End If
    DGVType.Visible = False
    txt(VType).SetFocus
Else
    If RsVType.RecordCount > 0 Then
        txtPrint(VType1).TEXT = RsVType!Name
        txtPrint(VType1).Tag = RsVType!Code
    End If
    DGVType.Visible = False
    txtPrint(VType1).SetFocus
End If
End Sub

Private Sub FGrid2_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
txtgrid2(0).Visible = False
End Sub
Private Sub FGrid2_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid2.Col
    Case Col_ChqNo, Col_ChqDate, Col_ChqAmt
        Call GridDblClick(Me, FGrid2, txtgrid2, 0)
End Select
TAddMode = False
End Sub
Private Sub FGrid2_GotFocus()
    FGrid2.BackColorSel = BackColorSelEnter
    FGrid2.ForeColorSel = ForeColorSelEnter
    
    txtgrid2(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid2.Tag) = (FGrid2.Rows - (FGrid2.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid2.Tag) = FGrid2.Rows - 1 Then
    SendKeys vbTab
    KeyCode = 0
End If

GridKey = KeyCode
FGrid2.Tag = FGrid2.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid2.Col
        Case Col_ChqNo, Col_ChqAmt, Col_ChqDate
            FGrid2 = ""
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid2.Col
        Case Col_ChqNo, Col_ChqDate, Col_ChqAmt
            DoEvents
            Call GridDblClick(Me, FGrid2, txtgrid2, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub
Private Sub FGrid2_KeyPress(KeyAscii As Integer)

Select Case FGrid2.Col
    Case Col_ChqNo, Col_ChqDate
       Call Get_Text(Me, FGrid2, txtgrid2, 0, False, KeyAscii)
    Case Col_ChqAmt
       Call Get_Text(Me, FGrid2, txtgrid2, 0, True, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid2.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid2.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid2.Rows > 2 Then
                FGrid2.RemoveItem (FGrid2.Row)
            Else
                FGrid2.Rows = 1
                FGrid2.AddItem FGrid2.Rows
                FGrid2.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid2.Rows - 1
            FGrid2.TextMatrix(I, 0) = I
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
   
FGrid2.SetFocus
End If
Exit Sub
End Sub
Private Sub FGrid2_LostFocus()
FGrid2.BackColorSel = BackColorSelLeave
FGrid2.ForeColorSel = FGrid2.ForeColor
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
    TopCtrl1.Tag = PubUParam: WinSetting Me, 5760, , 850, 465
    DGBook.left = 0: DGBook.width = Me.width - 90: DGBook.top = Me.height - (DGBook.height + mBotScale)
    DGParty.left = 0: DGParty.width = Me.width - 90: DGParty.top = txt(Amt).top: DGParty.height = Me.height - (DGParty.top + mBotScale)
    DGSite.left = 4500: DGSite.top = mTopScale
    DGVno.left = 4500: DGVno.top = mTopScale
    DGVType.left = 4500: DGVType.top = mTopScale
    DgCity.left = 4500: DgCity.top = mTopScale
    FrmPrn.left = 525: FrmPrn.top = 2220
    Ini_Grid
    If StrCmp(left(PubComp_Name, 6), "J.M.A.") Then
        mMultipleChqNo = True
    End If
    '** Hide Vehicle Details if Only Vehicle Section is not activated
    If PubVCompCode = "" Then
        For I = 18 To 21
            Label3(I).Visible = False
        Next
        txt(9).Visible = False
        txt(11).Visible = False
        txt(12).Visible = False
        txt(15).Visible = False
        txt(16).Visible = False
    End If

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
      Dim sitecond As String
    sitecond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    If PubMoveRecYn Then
        Master.Open "select DocId as searchcode,rect.* " & _
                    "from rect Where V_Date>=" & ConvertDate(PubStartDate) & " " & sitecond & " order by V_Date desc,docid", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("select Top 1 DocId as searchcode,rect.* " & _
                    "from rect Where V_Date>=" & ConvertDate(PubStartDate) & " " & sitecond & "order by V_Date desc,docid")
    End If
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
    Set RSBook = New ADODB.Recordset
    RSBook.CursorLocation = adUseClient
    RSBook.Open "SELECT " & cCStr("Veh_Order.Ord_No") & " as code,Veh_Order.OrdDocId, Veh_Order.Net_AMOUNT,Veh_Order.Ord_SiteCode,  Veh_Order.Ord_Date, Veh_Order.PartyCode, Veh_Order.MODEL, Veh_Order.FB_CODE, Veh_Order.Inv_No, Veh_Order.Inv_Date, Site.Site_Desc, ContractFinance.FinName,sum(" & cIIF("Rect.DrCr = 'D'", "Rect.AMOUNT", "Rect.AMOUNT*-1") & ") as AmtPaid " & _
        "FROM ((Veh_Order LEFT JOIN Site ON right(Veh_Order.Ord_SiteCode,1) = Site.Site_Code) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Rect ON (Veh_Order.OrdDocId = Rect.Ord_DocId) AND (Veh_Order.Ord_SiteCode = Rect.Ord_SiteCode) " & _
        "where veh_order.PartyCode ='' " & _
        "group by Veh_Order.OrdDocId, Veh_Order.Ord_SiteCode, Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.PartyCode, Veh_Order.MODEL, Veh_Order.FB_CODE, Veh_Order.Inv_No, Veh_Order.Inv_Date, Site.Site_Desc, ContractFinance.FinName, Veh_Order.Net_AMOUNT", GCn, adOpenDynamic, adLockOptimistic
    Set DGBook.DataSource = RSBook
    
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "select citycode as code,cityname as name from city order by cityname,citycode", GCn, adOpenDynamic, adLockOptimistic
    Set DgCity.DataSource = RsCity
    
    Set RsVType = New ADODB.Recordset
    RsVType.CursorLocation = adUseClient
    RsVType.Open "select Voucher_Type.v_type as code,Description as name from Voucher_Type where category='GENFA' order by v_type ", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGVType.DataSource = RsVType
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select SubGroup.SUBCODE as code,SubGroup.NAME, FPrefix + ' ' + FName as Father,Curr_Bal,ITWARD_NO,PANNO from SubGroup  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    'RsParty.Open "select SubGroup.SUBCODE as code,SubGroup.NAME, FPrefix + ' ' + FName as Father,Curr_Bal,ITWARD_NO,PANNO from SubGroup Where Nature In ('Customer', 'Supplier', 'Cash', 'Bank', 'Employee')  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct v_no as code from rect ", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
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
Set RSBook = Nothing
Set Master = Nothing
Set RsParty = Nothing
Set RsCity = Nothing
Set RsVno = Nothing
Set RsSite = Nothing
Set RsVType = Nothing
End Sub

Private Sub ListView_Click()
txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
txt(Val(ListView.Tag)).SetFocus
FrmList.Visible = False
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""
    txt(VehTrnYN) = IIf(PubVCompCode <> "", "No", "")
    txt(PrnNameYN) = "Yes"
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        txt(SiteCode).Tag = PubSiteCode
        txt(SiteCode) = PubSiteName
        txt(VDate).SetFocus
    Else
        txt(SiteCode).SetFocus
    End If
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean ',i As Integer
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$

If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub

If GCn.Execute("Select CancelYN from Rect where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
    MsgStr = "Are You Sure To Delete This ? "
    mTitle = "Delete Entry!"
Else
    MsgStr = "Are You Sure To Cancel This ? "
    mTitle = "Cancel Entry!"
End If
If MsgBox(MsgStr, vbYesNo + vbCritical + vbDefaultButton2, mTitle) = vbYes Then
    vBook = Master.AbsolutePosition
    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    'Unpost Ledger a/c
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(TxtDocID))
    If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    'Unposting of Ledger completed
    If GCn.Execute("Select CancelYN from Rect where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
        GCn.Execute ("delete from Rect where DocId='" & txt(TxtDocID) & "'")
    Else
        GCn.Execute "update rect set " & _
            "CancelYN=1,AMOUNT =0, AcCode= '" & txt(AcHead).Tag & "'," & _
            "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " & _
            " where docid = '" & txt(TxtDocID) & "'"
    End If
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    If Master.RecordCount > 0 Then
        If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
    End If
    BUTTONS True, Me, Master, 0
    Call MoveRec
End If
Exit Sub

eloop1:
    If mTrans = True Then GCn.RollbackTrans: G_FaCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    txt(Party).SetFocus
    If txt(PrnNameYN) = "Yes" Then
        txt(Rem1).Enabled = False
        txt(Rem2).Enabled = False
        txt(Rem1).BackColor = CtrlBColDisabled
        txt(Rem2).BackColor = CtrlBColDisabled
        txt(2).Enabled = True
    Else
        txt(Rem1).Enabled = True
        txt(Rem2).Enabled = True
        txt(Rem1).BackColor = CtrlBColOrg
        txt(Rem2).BackColor = CtrlBColOrg
        
    End If
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then CheckError
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
    RSBook.Requery
    RsSite.Requery
    RsParty.Requery
    RsCity.Requery
    RsVType.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, j As Integer
    Dim mQry As String
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    Dim DocIdHlp As String
    Dim mDrCr As String
    Dim LedgAry() As LedgRec, mNarr$, mResult As Byte
    Dim mMultipleChqFlag As Boolean
    
'    On Error GoTo errlbl
    
    Grid_Hide
    If IsValid(txt(SiteCode), "Site Name") = False Then Exit Sub
    If IsValid(txt(VDate), "Date") = False Then Exit Sub
    If IsValid(txt(VType), "Voucher Type") = False Then Exit Sub
    If txt(SerialNo).Enabled = True Then
        If txt(SerialNo).TEXT = "" Then MsgBox "SerialNo is required field", vbInformation, "Validation Check": txt(SerialNo).SetFocus: Exit Sub
    Else
        If txt(SerialNo).TEXT = "" Then MsgBox "SerialNo is required field", vbInformation, "Validation Check": txt(VType).SetFocus: Exit Sub
    End If
    If txt(VehTrnYN) = "Yes" Then
        If IsValid(txt(BookNo), "Booking No.") = False Then Exit Sub
    End If
    If txt(Party).Tag = txt(AcHead).Tag Then
        MsgBox "Party A/c and Ledger A/c both same !" & vbCrLf & "Correct A/c Selection ", vbCritical, "A/c Checking"
        txt(AcHead).SetFocus: Exit Sub
    End If
    If Val(txt(Amt)) <= 0 Then
        MsgBox "Please Enter Amount", vbCritical, "Validation"
        txt(Amt).SetFocus: Exit Sub
    End If
    If Val(txt(DiscAmt)) > 0 And txt(DiscAcName).Tag = "" Then
        MsgBox "Please Enter Disc. A/c", vbCritical, "Validation"
        txt(DiscAcName).SetFocus: Exit Sub
    End If
    '********* cHECKING pOSTING cOTROLS
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        txt(AcPostByName) = pubUName
        txt(AcPostDate) = PubServerDate
    End If
    '*********
    ReDim Preserve LedgAry(4)
    mMultipleChqFlag = False
    If mMultipleChqNo Then
        Dim mSum As Double
        mSum = 0
        For I = 1 To FGrid2.Rows - 1
            If FGrid2.TextMatrix(I, Col_ChqNo) <> "" Then
                mSum = mSum + Val(FGrid2.TextMatrix(I, Col_ChqAmt))
            End If
        Next I
        
        If mSum <> Val(txt(Amt)) And mSum > 0 Then
            MsgBox "Detail of Cheque/DD can't be differ from total Amount"
            FGrid2.SetFocus
            Exit Sub
        ElseIf mSum > 0 Then
            mMultipleChqFlag = True
        End If
    End If
    
    If TopCtrl1.TopText2.CAPTION = "Add" Then
    'lp 11-03-03
        DocID = txt(TxtDocID)
        If GCn.Execute("select count(*) from rect where DocId='" & txt(TxtDocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
                MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                Exit Sub
            Else
                txt(TxtDocID) = GetDocID(G_FaCn, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    SetMax_VoucherPrefix "DocID", txt(VType).Tag, "Rect", "V_date"
                    txt(TxtDocID) = GetDocID(G_FaCn, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                    If Val(txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                        MsgBox "Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    DocIdHlp = Replace(txt(TxtDocID), " ", "")
    Select Case txt(VType).Tag
        Case "G_ABR", "G_ACR", "G_DRN", "G_TLR", "G_JV"
            mDrCr = "C"
        Case "G_BBP", "G_BCP", "G_CRN"
            mDrCr = "D"
    End Select

    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    mNarr = txt(Narr)
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute "insert into rect(DocId,DocIDHelp,V_Date,V_Type,V_No,site_code, " & _
            "Prov_No,Prov_Date,Prov_Location, " & _
            "Vehicle_YN,Ord_SiteCode,Ord_DocId,PartyCode," & _
            "Veh_Amt,Tax_Amt,Surcharge_Amt,TOT_Amt,AMOUNT, " & _
            "DrCr,Narration,AcCode, " & _
            "DDNo,DDDate,PrintParty_YN," & _
            "RWTF1,RWTF2,IFORM,RectCatg,CancelYN, CreditCardNo," & _
            "U_Name,U_EntDt,U_AE,AcPostByU_Name,AcPostByU_EntDt,AddBy, AddDate,DiscAc,DiscAmt) values( " & _
            "  '" & txt(TxtDocID) & "','" & DocIdHlp & "'," & ConvertDate(txt(VDate)) & ",'" & mVType & "'," & Val(txt(SerialNo)) & ",'" & PubSiteCode & txt(SiteCode).Tag & _
            "', " & Val(txt(ProNo)) & "," & ConvertDate(txt(ProDate)) & ",'" & txt(RecLocation).Tag & _
            "', " & IIf(txt(VehTrnYN) = "Yes", 1, 0) & ",'" & mID(txt(BookNo).Tag, 2, 2) & "','" & txt(BookNo).Tag & "','" & txt(Party).Tag & _
            "', " & Val(txt(VehAmt)) & ", " & Val(txt(TaxAmt)) & ", " & Val(txt(TaxSurch)) & "," & Val(txt(TOTAmt)) & "," & Val(txt(Amt)) & _
            " ,'" & mDrCr & "','" & mNarr & "', '" & txt(AcHead).Tag & _
            "','" & txt(DDNo).TEXT & "'," & ConvertDate(txt(DDDate)) & ", " & IIf(txt(PrnNameYN) = "Yes", 1, 0) & _
            " ,'" & txt(Rem1) & "','" & txt(Rem2) & "','" & txt(FormType).TEXT & "','" & txt(RectCatg).TEXT & "',0, '" & txt(CreditCardNo) & "'," & _
            "  '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & txt(AcPostByName) & "'," & ConvertDate(txt(AcPostDate)) & ", '" & pubUName & "', " & ConvertDateTime(PubServerDate) & ",'" & txt(DiscAcName).Tag & "'," & Val(txt(DiscAmt)) & " )"
            'Voucher Serial No. Updation LPS 21-05-03
            'update Table only when DocSrlNo >Table.SerialNo
            UpdVouSrlNo G_FaCn, txt(TxtDocID), txt(VDate)
    Else
        GCn.Execute "update rect set V_date = " & ConvertDate(txt(VDate)) & ",RectCatg='" & txt(RectCatg).TEXT & "', Prov_No=" & Val(txt(ProNo)) & ", " & _
            "Prov_Date=" & ConvertDate(txt(ProDate)) & ",Prov_Location='" & txt(RecLocation).Tag & "', Vehicle_YN=" & IIf(txt(VehTrnYN) = "Yes", 1, 0) & ", " & _
            "Ord_SiteCode='" & mID(txt(BookNo).Tag, 2, 2) & "',Ord_DocId='" & txt(BookNo).Tag & "',PartyCode='" & txt(Party).Tag & "', " & _
            "Veh_Amt=" & Val(txt(VehAmt)) & ",Tax_Amt= " & Val(txt(TaxAmt)) & ",Surcharge_Amt= " & Val(txt(TaxSurch)) & ",TOT_Amt=" & Val(txt(TOTAmt)) & ",AMOUNT =" & Val(txt(Amt)) & _
            ",DrCr='" & mDrCr & "' ,Narration='" & mNarr & "',AcCode= '" & txt(AcHead).Tag & "',DDNo='" & txt(DDNo).TEXT & "',DDDate=" & ConvertDate(txt(DDDate)) & " , " & _
            "PrintParty_YN=" & IIf(txt(PrnNameYN) = "Yes", 1, 0) & ",RWTF1='" & txt(Rem1) & "',RWTF2='" & txt(Rem2) & "', " & _
            "IFORM='" & txt(FormType).TEXT & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E', " & _
            "AcPostByU_Name='" & txt(AcPostByName) & "',AcPostByU_EntDt=" & ConvertDate(txt(AcPostDate)) & ", ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ",DiscAc='" & txt(DiscAcName).Tag & "',DiscAmt=" & Val(txt(DiscAmt)) & ", CreditCardNo = '" & txt(CreditCardNo) & "' " & _
            " where docid = '" & txt(TxtDocID) & "'"
    End If
    
    mQry = "Delete From Rect1 Where DocId = '" & txt(TxtDocID) & "'"
    GCn.Execute mQry
    For I = 1 To FGrid2.Rows - 1
        If FGrid2.TextMatrix(I, Col_ChqNo) <> "" Then
            mQry = "INSERT INTO dbo.Rect1(DocID,Sr,ChqNo,ChqDate,ChqAmt) " & _
                   "VALUES ('" & txt(TxtDocID) & "'," & I & ",'" & FGrid2.TextMatrix(I, Col_ChqNo) & "'," & ConvertDate(FGrid2.TextMatrix(I, Col_ChqDate)) & "," & Val(FGrid2.TextMatrix(I, Col_ChqAmt)) & ")"
            GCn.Execute mQry
        End If
    Next I
    
    
    'A/c Posting
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        Select Case txt(VType).Tag
            Case "G_ABR", "G_ACR", "G_TLR", "G_JV", "SBLCQ", "SBLCS", "SBLRO"  'Receipt
                I = 0
                
                If Not mMultipleChqFlag Then
                    LedgAry(I).SubCode = txt(AcHead).Tag
                    LedgAry(I).AmtDr = Val(txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = txt(Party).Tag
                Else
                    For j = 0 To FGrid2.Rows - 1
                        I = I + 1
                        ReDim Preserve LedgAry(UBound(LedgAry) + 1)
                        LedgAry(I).SubCode = txt(AcHead).Tag
                        LedgAry(I).AmtDr = Val(FGrid2.TextMatrix(j, Col_ChqAmt))
                        LedgAry(I).Chq_No = FGrid2.TextMatrix(j, Col_ChqNo)
                        LedgAry(I).Chq_Date = FGrid2.TextMatrix(j, Col_ChqDate)
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = txt(Party).Tag
                    Next j
                End If
                I = I + 1
                LedgAry(I).SubCode = txt(Party).Tag
                LedgAry(I).AmtCr = Val(txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = txt(AcHead).Tag
            Case "G_DRN"
                I = 0
                LedgAry(I).SubCode = txt(AcHead).Tag
                LedgAry(I).AmtCr = Val(txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = txt(Party).Tag
                I = I + 1
                LedgAry(I).SubCode = txt(Party).Tag
                LedgAry(I).AmtDr = Val(txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = txt(AcHead).Tag
            Case "G_BBP", "G_BCP"
                I = 0
                LedgAry(I).SubCode = txt(Party).Tag
                LedgAry(I).AmtDr = Val(txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = txt(AcHead).Tag
                I = I + 1
                LedgAry(I).SubCode = txt(AcHead).Tag
                LedgAry(I).AmtCr = Val(txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = txt(Party).Tag
                
            Case mCustCRN
                I = 0
                LedgAry(I).SubCode = txt(Party).Tag
                LedgAry(I).AmtCr = Val(txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = txt(AcHead).Tag
                I = I + 1
                LedgAry(I).SubCode = txt(AcHead).Tag
                If Val(txt(VehAmt)) + Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(TOTAmt)) = 0 Then
                    LedgAry(I).AmtDr = Val(txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = txt(Party).Tag
                Else
                    LedgAry(I).AmtDr = Val(txt(VehAmt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
    
                    I = I + 1
                    LedgAry(I).SubCode = mTaxAcHead
                    LedgAry(I).AmtDr = Val(txt(TaxAmt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
                    I = I + 1
                    LedgAry(I).SubCode = mTaxSurAcHead
                    LedgAry(I).AmtDr = Val(txt(TaxSurch))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
                    I = I + 1
                    LedgAry(I).SubCode = mTOTAcHead
                    LedgAry(I).AmtDr = Val(txt(TOTAmt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
                End If
        End Select
        If txt(VType).Tag = "G_ACR" Then
            If Val(txt(DiscAmt)) > 0 Then
                I = I + 1
                LedgAry(I).SubCode = txt(DiscAcName).Tag
                LedgAry(I).AmtDr = Val(txt(DiscAmt))
                LedgAry(I).Narration = "Being Cash Discount Given to Party"
                LedgAry(I).ContraSub = txt(Party).Tag
                I = I + 1
                LedgAry(I).SubCode = txt(Party).Tag
                LedgAry(I).AmtCr = Val(txt(DiscAmt))
                LedgAry(I).Narration = "Being Cash Discount Given to Party"
                LedgAry(I).ContraSub = txt(DiscAcName).Tag
            End If
        End If
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(TxtDocID), CDate(txt(VDate)), mNarr)
        G_FaCn.Execute ("Update Ledger set chq_no='" & txt(DDNo) & "',Chq_Date=" & ConvertDate(txt(DDDate)) & " where DocId='" & txt(TxtDocID) & "'")
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    End If
    'eof posting
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Set Rst = Nothing
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select DocId as searchcode,rect.* " & _
                    "from rect Where V_Date>=" & ConvertDate(PubStartDate) & " And DocId = '" & txt(TxtDocID) & "' order by V_Date desc,docid")
    End If
    RSBook.Requery
    Master.FIND "DocId = '" & txt(TxtDocID) & "'"

    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > DeCodeDocID(DocID, Document_No) Then
            MsgBox "Document No." & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans: G_FaCn.RollbackTrans
    CheckError
Exit Sub
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    Dim sitecond As String
    sitecond = " Where V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = sitecond & " and " & cMID("r.DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "select R.DocId as searchcode, " & cDt("R.V_Date") & " As V_Date, R.V_Type, " & cCStr("R.V_No", 10) & " As V_No, " & cCStr("R.Prov_No", 15) & " As Prov_No, R.Prov_Date As Prov_Date, R.Narration from Rect R " & sitecond & " "
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
        Set Master = GCn.Execute("select DocId as searchcode,rect.* " & _
                    "from rect Where V_Date>=" & ConvertDate(PubStartDate) & " And DocId = '" & MyValue & "' order by V_Date desc,docid")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
If txt(SiteCode).TEXT <> "" Then
    If txt(VDate).TEXT = "" Then txt(VDate).SetFocus: Ctrl_GetFocus txt(Index): Exit Sub
    If txt(VType).TEXT = "" Then txt(VType).SetFocus: Ctrl_GetFocus txt(Index): Exit Sub
End If
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case VType
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsVType!Name Then
            RsVType.MoveFirst
            RsVType.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case RectCatg
        ListArray = Array("     ", "M.M.", "BAL", "FULL", "Staff")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 5)

    Case FormType
        ListArray = Array("Form-60", "Form-61", "N/A")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
    Case SiteCode
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txt(Index).Tag = RsSite!Code
            txt(Index).TEXT = RsSite!Name
        Else
            If txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txt(Index).TEXT & "'"
            End If
        End If
    Case RecLocation
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsCity!Name Then
            RsCity.MoveFirst
            RsCity.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case Party
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case BookNo
         Set RSBook = GCn.Execute("SELECT " & cCStr("Veh_Order.Ord_No") & " as code,Veh_Order.OrdDocId, Veh_Order.Net_AMOUNT, Veh_Order.Ord_SiteCode,  Veh_Order.Ord_Date, Veh_Order.PartyCode, Veh_Order.MODEL, Veh_Order.FB_CODE, Veh_Order.Inv_No, Veh_Order.Inv_Date, Site.Site_Desc, ContractFinance.FinName, sum(" & cIIF("Rect.DrCr = 'D'", "Rect.AMOUNT", "Rect.AMOUNT*-1") & ") as AmtPaid " & _
            "FROM ((Veh_Order LEFT JOIN Site ON right(Veh_Order.Ord_SiteCode,1) = Site.Site_Code) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Rect ON (Veh_Order.OrdDocId = Rect.Ord_DocId) AND (Veh_Order.Ord_SiteCode = Rect.Ord_SiteCode) " & _
            "WHERE Veh_Order.PartyCode = '" & txt(Party).Tag & "'  " & _
            "group by Veh_Order.OrdDocId, Veh_Order.Ord_SiteCode, Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.PartyCode, Veh_Order.MODEL, Veh_Order.FB_CODE, Veh_Order.Inv_No, Veh_Order.Inv_Date, Site.Site_Desc, ContractFinance.FinName, Veh_Order.Net_AMOUNT")
        Set DGBook.DataSource = RSBook
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RSBook!Code Then
            RSBook.MoveFirst
            RSBook.FIND "code ='" & txt(Index).TEXT & "'"
        End If
    Case AcHead, DiscAcName
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & txt(Index).TEXT & "'"
        End If
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case VType
        DGridTxtKeyDown DGVType, txt, Index, RsVType, KeyCode, False, 1
    Case RectCatg
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1500
    Case FormType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 900
    Case SiteCode
        DGridTxtKeyDown DGSite, txt, Index, RsSite, KeyCode, False, 1
    Case BookNo
        DGridTxtKeyDown DGBook, txt, Index, RSBook, KeyCode, False, 0
    Case RecLocation
        DGridTxtKeyDown DgCity, txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
    Case Party, AcHead, DiscAcName
        DGridTxtKeyDown DGParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
End Select
If FrmList.Visible = False And DGVType.Visible = False And DgCity.Visible = False And DGParty.Visible = False And DGBook.Visible = False And DGSite.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
        If (txt(Rem2).Enabled = True And Index <> Rem2) Or (txt(Rem2).Enabled = False And Index <> PrnNameYN) Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (txt(Rem2).Enabled = True And Index = Rem2) Or (txt(Rem2).Enabled = False And Index = PrnNameYN) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SiteCode Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> ProNo Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Dim mVehYN As Boolean
Call CheckQuote(KeyAscii)

Select Case Index
    Case VType
        If DGVType.Visible = True Then DGridTxtKeyPress txt, Index, RsVType, KeyAscii, "name"
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress txt, Index, RsSite, KeyAscii, "Name"
    Case BookNo
        If DGBook.Visible = True Then DGridTxtKeyPress txt, Index, RSBook, KeyAscii, "Code"
    Case RecLocation
        If DgCity.Visible = True Then DGridTxtKeyPress txt, Index, RsCity, KeyAscii, "Name"
    Case SerialNo
        Call NumPress(txt(Index), KeyAscii, 6, 0)
    Case Party, AcHead, DiscAcName
        If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "Name"
    Case VehTrnYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            txt(Index) = "Yes"
            mVehYN = True
'            txt(BookNo).Enabled = True
        ElseIf UCase(Chr(KeyAscii)) = "N" Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txt(Index) = "No"
'            txt(BookNo).Enabled = False
'        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
'            txt(Index) = "No"
'            txt(BookNo).Enabled = False
        End If
        If txt(VehTrnYN) = "Yes" Then
            mVehYN = True
        End If
        txt(BookNo).Enabled = mVehYN
        If mVType <> mCustCRN Then
            mVehYN = False
        End If
        txt(VehAmt).Enabled = mVehYN
        txt(TaxPer).Enabled = mVehYN
        txt(TaxAmt).Enabled = mVehYN
        txt(TaxSurPer).Enabled = mVehYN
        txt(TaxSurch).Enabled = mVehYN
        txt(TOTPer).Enabled = mVehYN
        txt(TOTAmt).Enabled = mVehYN
        
        KeyAscii = 0
        If txt(BookNo).Enabled = False Then
            txt(BookNo).BackColor = CtrlBColDisabled
            txt(VehAmt).BackColor = CtrlBColDisabled
            txt(TaxPer).BackColor = CtrlBColDisabled
            txt(TaxAmt).BackColor = CtrlBColDisabled
            txt(TaxSurPer).BackColor = CtrlBColDisabled
            txt(TaxSurch).BackColor = CtrlBColDisabled
            txt(TOTPer).BackColor = CtrlBColDisabled
            txt(TOTAmt).BackColor = CtrlBColDisabled
        Else
            txt(BookNo).BackColor = CtrlBColOrg
            txt(VehAmt).Enabled = CtrlBColOrg
            txt(TaxPer).Enabled = CtrlBColOrg
            txt(TaxAmt).Enabled = CtrlBColOrg
            txt(TaxSurPer).Enabled = CtrlBColOrg
            txt(TaxSurch).Enabled = CtrlBColOrg
            txt(TOTPer).Enabled = CtrlBColOrg
            txt(TOTAmt).Enabled = CtrlBColOrg
        End If
    Case VehAmt, TaxAmt, TaxSurch, TOTAmt, Amt
        Call NumPress(txt(Index), KeyAscii, 8, 2)
    Case PrnNameYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            txt(Index) = "Yes"
            txt(Rem1).Enabled = False
            txt(Rem2).Enabled = False
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txt(Index) = "No"
            txt(Rem1).Enabled = True
            txt(Rem2).Enabled = True
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txt(Index) = ""
            txt(Rem1).Enabled = False
            txt(Rem2).Enabled = False
        End If
        KeyAscii = 0
        If txt(Rem1).Enabled = False Then
            txt(Rem1).BackColor = CtrlBColDisabled
        Else
            txt(Rem1).BackColor = CtrlBColOrg
        End If
        If txt(Rem2).Enabled = False Then
            txt(Rem2).BackColor = CtrlBColDisabled
        Else
            txt(Rem2).BackColor = CtrlBColOrg
        End If
End Select
'KeyAscii = RetDGKeyAscii()
End Sub



Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs

Select Case Index
    Case FormType, RectCatg
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case VehAmt, TaxAmt, TaxSurch, TOTAmt
        txt(Amt) = Val(txt(VehAmt)) + Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(TOTAmt))
    Case CreditCardNo
        If txt(Index) <> "" Then
            txt(AcHead).Enabled = False
        Else
            txt(AcHead).Enabled = True
        End If
End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim I As Integer, mEnb As Boolean, mVehYN As Boolean
Select Case Index
    Case VehTrnYN
        If txt(Index) <> "Yes" Then
            txt(BookNo).TEXT = ""
            txt(BookNo).Tag = ""
            txt(BookDate).TEXT = ""
            BookDocId = ""
            txt(Model).TEXT = ""
            txt(FB_Code).Tag = ""
            txt(FB_Code).TEXT = ""
            txt(VehAmt) = ""
            txt(TaxPer) = ""
            txt(TaxAmt) = ""
            txt(TaxSurPer) = ""
            txt(TaxSurch) = ""
            txt(TOTPer) = ""
            txt(TOTAmt) = ""
         Else
            If txt(VType) <> "Customer Credit Note" Then
                txt(VehAmt).Enabled = False
                txt(TaxPer).Enabled = False
                txt(TaxAmt).Enabled = False
                txt(TaxSurPer).Enabled = False
                txt(TaxSurch).Enabled = False
                txt(TOTPer).Enabled = False
                txt(TOTAmt).Enabled = False
            End If
            
        End If
        
    Case VType
        If IsValid(txt(Index), "Voucher Type") = False Then Cancel = True: Exit Sub
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsVType!Name
            txt(Index).Tag = RsVType!Code
            mVType = txt(Index).Tag
            If RSOJPR = True Then
                If mVType = mCustCR Then
                    txt(DiscAcName).Visible = True
                    txt(DiscAmt).Visible = True
                    Label3(10).Visible = True
                    Label3(11).Visible = True
                Else
                    txt(DiscAcName).Visible = False
                    txt(DiscAmt).Visible = False
                    Label3(10).Visible = False
                    Label3(11).Visible = False
                End If
            End If
            If mVType = mCustBP Or mVType = mCustBR Then
                mEnb = True
                txt(DDNo).BackColor = CtrlBColOrg
                txt(DDDate).BackColor = CtrlBColOrg
            Else
                txt(DDNo).BackColor = CtrlBColDisabled
                txt(DDDate).BackColor = CtrlBColDisabled
            End If
            txt(DDNo).Enabled = mEnb
            txt(DDDate).Enabled = mEnb
            If mVType = mCustCRN Then
                mVehYN = True
                txt(VehAmt).BackColor = CtrlBColOrg
                txt(TaxAmt).BackColor = CtrlBColOrg
                txt(TaxSurch).BackColor = CtrlBColOrg
                txt(TOTAmt).BackColor = CtrlBColOrg
                txt(VehAmt).Enabled = mVehYN
                txt(TaxAmt).Enabled = mVehYN
                txt(TaxSurch).Enabled = mVehYN
                txt(TOTAmt).Enabled = mVehYN
            Else
                txt(VehAmt).BackColor = CtrlBColDisabled
                txt(TaxAmt).BackColor = CtrlBColDisabled
                txt(TaxSurch).BackColor = CtrlBColDisabled
                txt(TOTAmt).BackColor = CtrlBColDisabled
            End If
            
            
            
            'DocID
            txt(TxtDocID) = GetDocID(G_FaCn, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            DocID = txt(TxtDocID)
            
            
            If txt(VType) <> "Customer Bank Receipt" Then
                txt(CreditCardNo).Enabled = False
            Else
                txt(CreditCardNo).Enabled = True
            End If
            
        End If
    Case SerialNo
        If IsValid(txt(SerialNo), "Serial No.") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag = True Then      ' Manual
            txt(TxtDocID) = GetDocID(G_FaCn, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            DocID = txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select * From Rect Where docid='" & txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(SerialNo).SetFocus
            End If
        End If
    Case FormType, RectCatg
        If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case SiteCode
        If IsValid(txt(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsSite!Name
            txt(Index).Tag = RsSite!Code
        End If
    Case RecLocation
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsCity!Name
            txt(Index).Tag = RsCity!Code
        End If
    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            txt(BookDate).TEXT = ""
            BookDocId = ""
            txt(Model).TEXT = ""
            txt(FB_Code).Tag = ""
            txt(FB_Code).TEXT = ""
            txt(VehAmt) = ""
            txt(TaxPer) = ""
            txt(TaxAmt) = ""
            txt(TaxSurPer) = ""
            txt(TaxSurch) = ""
            txt(TOTPer) = ""
            txt(TOTAmt) = ""
        Else
            txt(Index).TEXT = RSBook!Code
            txt(Index).Tag = RSBook!OrdDocId
            FillRecords RSBook
        End If
    Case Party
        If IsValid(txt(Index), "Party Name") = False Then Cancel = True: Exit Sub
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
            txt(PanNo) = IIf(IsNull(RsParty!PanNo), "", RsParty!PanNo)
            txt(CircleNo) = IIf(IsNull(RsParty!ITWARD_NO), "", RsParty!ITWARD_NO)
            LblPartyBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
            LblPartyBal = LblPartyBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
            txt(FormType).Enabled = IIf(IsNull(RsParty!PanNo) Or RsParty!PanNo = "", True, False)
            txt(FormType).TEXT = IIf(txt(FormType).Enabled = True, "N/A", "")
        End If
    Case AcHead
        If IsValid(txt(Index), "A/C Head") = False Then Cancel = True: Exit Sub
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
            lblAcBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
            lblAcBal = lblAcBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
        End If
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
            txt(VDate).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        Cancel = Not CheckFinYear(txt(VDate))
       If Cancel = False And TopCtrl1.TopText2 = "Add" Then txt(VType).SetFocus
    Case ProDate
        txt(Index).TEXT = RetDate(txt(Index))
        If txt(ProDate) <> "" Then
            If CDate(txt(ProDate)) > CDate(txt(VDate)) Then
                MsgBox "Provisional Date  > Vr Date", vbInformation, "Validation"
                Cancel = True
                txt(ProDate).SetFocus
            End If
        End If
    Case BookDate, DDDate
        txt(Index).TEXT = RetDate(txt(Index))
    Case Amt
        txt(Index).TEXT = Format(txt(Index).TEXT, "0.00")
    Case DiscAmt
        If Val(txt(DiscAmt)) > Val(txt(Amt)) Then
            MsgBox "Please give proper discount."
            txt(DiscAmt) = "": txt(DiscAmt).SetFocus
            Cancel = True
            Exit Sub
        End If
        txt(Index).TEXT = Format(txt(Index).TEXT, "0.00")
    Case CreditCardNo
        If txt(CreditCardNo) <> "" Then
            Set RsTemp = GCn.Execute("Select A.CreditCardAc, S.Name From AcControls A Left Join SubGroup S On A.CreditCardAc=S.SubCode")
            If RsTemp.RecordCount > 0 Then
                txt(AcHead).Tag = XNull(RsTemp!CreditCardAc)
                txt(AcHead) = XNull(RsTemp!Name)
                txt(AcHead).Enabled = False
                
            End If
        Else
            txt(AcHead).Enabled = True
        End If
End Select
Ctrl_validate txt(Index)
Set Rst = Nothing
End Sub

'*** Fuctions ********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
    txt(I).Tag = ""
Next I

'FGrid2.Clear
FGrid2.Rows = 1
FGrid2.AddItem ""
FGrid2.FixedRows = 1

End Sub
Private Sub MoveRec()
Dim Rst As Recordset
Dim I As Integer
Dim TmpRst As Recordset
On Error GoTo error1
If Master.RecordCount > 0 Then
    If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
'    If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then Me.TopCtrl1.tDel = True
    If Master!CancelYN = 1 Then
        LblCancel.Visible = True
        TopCtrl1.tEdit = False
    Else
        LblCancel.Visible = False
    End If
    DocID = Master!DocID
    txt(TxtDocID).TEXT = Master!DocID
    LblDiv.CAPTION = "Division : " & left(Master!DocID, 1)
    LblSite.CAPTION = "Site Code : " & mID(Master!Site_Code, 1, 1)
    txt(SiteCode).Tag = mID(Master!Site_Code, 2, 1)
    txt(SiteCode).TEXT = GCn.Execute("select site_desc from site where site_code = '" & txt(SiteCode).Tag & "'").Fields(0).Value
    LblUser = IIf(Not IsNull(Master!AddDate), "Add By : " & XNull(Master!AddBy) & "  Dated : " & XNull(Master!AddDate), "") & IIf(Not IsNull(Master!ModifyDate), "     Modify By : " & XNull(Master!ModifyBy) & "  Dated : " & XNull(Master!ModifyDate), "")
    LblVPrefix.CAPTION = mID(Master!DocID, 8, 5)
    txt(SerialNo).TEXT = Master!V_NO
    txt(VDate).TEXT = Master!V_DATE
    txt(RectCatg).TEXT = IIf(IsNull(Master!RectCatg), "", Master!RectCatg)
    mVType = Master!V_Type
    txt(VType).Tag = mVType
    txt(VType).TEXT = G_FaCn.Execute("select Description from Voucher_Type where category='GENFA' and v_type = '" & txt(VType).Tag & "'").Fields(0).Value
    '*** A/c Posting Status
    txt(AcPostByName) = IIf(IsNull(Master!AcPostByU_Name), "", Master!AcPostByU_Name)
    txt(AcPostDate) = IIf(IsNull(Master!AcPostByU_EntDt), "", Master!AcPostByU_EntDt)
    '***
    txt(BookNo).Tag = Master!Ord_DocId
    If txt(BookNo).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "SELECT VO.Inv_Date,VO.OrdDocId,VO.Ord_no,VO.Ord_Date, VO.MODEL, VO.FB_CODE FROM veh_order as VO where OrdDocId = '" & txt(BookNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount > 0 Then
            txt(BookNo).TEXT = Rst!Ord_No
        End If
        FillRecords Rst
    Else
        txt(BookNo).TEXT = ""
        txt(BookDate).TEXT = ""
        BookDocId = ""
        txt(Model).TEXT = ""
        txt(FB_Code).Tag = ""
        txt(FB_Code).TEXT = ""
    End If
    txt(Party).Tag = Master!PartyCode
    If txt(Party).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NAME,Curr_Bal,PanNo,ITWARD_NO from SubGroup where Subcode = '" & txt(Party).Tag & "'", GCn, adOpenDynamic, adLockBatchOptimistic
        txt(Party) = Rst!Name
        txt(PanNo) = IIf(IsNull(Rst!PanNo), "", Rst!PanNo)
        txt(CircleNo) = IIf(IsNull(Rst!ITWARD_NO), "", Rst!ITWARD_NO)
        LblPartyBal = "Bal. " & Format(Abs(Rst!Curr_Bal), "0.00")
        LblPartyBal = LblPartyBal & IIf(Rst!Curr_Bal > 0, " Cr", IIf(Rst!Curr_Bal < 0, " Dr", ""))
        txt(FormType).Enabled = IIf(IsNull(Rst!PanNo) Or Rst!PanNo = "", True, False)
        If TopCtrl1.TopText2 = "Browse" Then txt(FormType).Enabled = False
        txt(FormType).TEXT = IIf(txt(FormType).Enabled = True, "N/A", "")
    Else
        txt(Party).TEXT = ""
    End If
    txt(ProNo).TEXT = IIf(IsNull(Master!Prov_No) Or Master!Prov_No = 0, "", Master!Prov_No)
    txt(ProDate).TEXT = IIf(IsNull(Master!Prov_Date), "", Master!Prov_Date)
    txt(RecLocation).Tag = IIf(IsNull(Master!Prov_Location), "", Master!Prov_Location)
    If txt(RecLocation).Tag <> "" Then
        txt(RecLocation).TEXT = GCn.Execute("select cityname from city where citycode = '" & txt(RecLocation).Tag & "'").Fields(0).Value
    End If
    txt(DDNo).TEXT = IIf(IsNull(Master!DDNo), "", Master!DDNo)
    txt(DDDate).TEXT = IIf(IsNull(Master!DDDate), "", Master!DDDate)
    txt(VehTrnYN).TEXT = IIf(Master!Vehicle_YN = 1, "Yes", "No")
    txt(PrnNameYN).TEXT = IIf(Master!PrintParty_YN = 1, "Yes", "No")
    txt(Narr).TEXT = IIf(IsNull(Master!Narration), "", Master!Narration)
    txt(Rem1).TEXT = IIf(IsNull(Master!RWTF1), "", Master!RWTF1)
    txt(Rem2).TEXT = IIf(IsNull(Master!RWTF2), "", Master!RWTF2)
    txt(FormType) = IIf(IsNull(Master!IForm), "", Master!IForm)
    txt(CreditCardNo) = XNull(Master!CreditCardNo)
    If mVType = mCustCRN Then
        txt(VehAmt) = IIf(IsNull(Master!Veh_Amt), "", Format(Master!Veh_Amt, "0.00"))
        txt(TaxAmt) = IIf(IsNull(Master!Tax_Amt), "", Format(Master!Tax_Amt, "0.00"))
        txt(TaxSurch) = IIf(IsNull(Master!Surcharge_Amt), "", Format(Master!Surcharge_Amt, "0.00"))
        txt(TOTAmt) = IIf(IsNull(Master!Tot_Amt), "", Format(Master!Tot_Amt, "0.00"))
    Else
        txt(VehAmt) = ""
        txt(TaxAmt) = ""
        txt(TaxSurch) = ""
        txt(TOTAmt) = ""
    End If
    txt(Amt).TEXT = Format(Master!Amount, "0.00")
    
    txt(AcHead).Tag = Master!AcCode
    If txt(AcHead).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NAME,Curr_Bal from SubGroup where Subcode = '" & txt(AcHead).Tag & "'", GCn, adOpenDynamic, adLockBatchOptimistic
        txt(AcHead) = XNull(Rst!Name)
        lblAcBal = "Bal. " & Format(Abs(Rst!Curr_Bal), "0.00")
        lblAcBal = lblAcBal & IIf(Rst!Curr_Bal > 0, " Cr", IIf(Rst!Curr_Bal < 0, " Dr", ""))
    Else
        txt(AcHead).TEXT = ""
    End If
    txt(DiscAcName).Tag = XNull(Master!DiscAc)
    If txt(DiscAcName).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NAME from SubGroup where Subcode = '" & txt(DiscAcName).Tag & "'", GCn, adOpenDynamic, adLockBatchOptimistic
        txt(DiscAcName) = Rst!Name
    Else
        txt(DiscAcName).TEXT = ""
    End If
    txt(DiscAmt) = Format(VNull(Master!DiscAmt), "0.00")
    
    
    
    Set TmpRst = GCn.Execute("Select M.* from Rect1 M where M.DocID='" & Master!SearchCode & "'")
    With FGrid2
        .Rows = 1
        If TmpRst.RecordCount > 0 Then
            TmpRst.MoveFirst
            For I = 1 To TmpRst.RecordCount
                .AddItem ""
                .TextMatrix(I, 0) = I
                .TextMatrix(I, Col_ChqNo) = XNull(TmpRst!ChqNo)
                .TextMatrix(I, Col_ChqDate) = XNull(TmpRst!ChqDate)
                .TextMatrix(I, Col_ChqAmt) = Format(VNull(TmpRst!ChqAmt), "0.00")
                TmpRst.MoveNext
            Next
            .FixedRows = 1
        Else
            .AddItem ""
            .FixedRows = 1
        End If
    End With
    
    
Else
    Call BlankText
End If
Grid_Hide
Set Rst = Nothing
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next


If UCase(left(PubComp_Name, 4)) = "ENAR" Then txt(SiteCode).Enabled = False

If TopCtrl1.TopText2 = "Edit" Then
    txt(SiteCode).Enabled = False
    txt(VDate).Enabled = False
    txt(SerialNo).Enabled = False
    txt(VType).Enabled = False
    If mVType <> mCustCRN Then
        txt(VehAmt).Enabled = False
        txt(TaxAmt).Enabled = False
        txt(TaxSurch).Enabled = False
        txt(TOTAmt).Enabled = False
    End If
End If
If RSOJPR = True Then
    Label3(10).Visible = True
    Label3(11).Visible = True
    txt(DiscAcName).Visible = True
    txt(DiscAmt).Visible = True
Else
    Label3(10).Visible = False
    Label3(11).Visible = False
    txt(DiscAcName).Visible = False
    txt(DiscAmt).Visible = False
End If
txt(Rem1).Enabled = False
txt(Rem2).Enabled = False
txt(TxtDocID).Enabled = False
txt(BookNo).Enabled = False
txt(BookDate).Enabled = False
txt(Model).Enabled = False
txt(FB_Code).Enabled = False
txt(Model).Enabled = False
txt(FB_Code).Enabled = False
txt(PanNo).Enabled = False
txt(CircleNo).Enabled = False
txtDisabled_Color Me

If txt(CreditCardNo) <> "" Then
    txt(AcHead).Enabled = False
End If

End Sub

Private Sub Grid_Hide()
    If DGBook.Visible = True Then DGBook.Visible = False
    If DGSite.Visible = True Then DGSite.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If DgCity.Visible = True Then DgCity.Visible = False
    If DGVType.Visible = True Then DGVType.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
End Sub
 
Private Sub FillRecords(RSBook As ADODB.Recordset)
Dim Rst As ADODB.Recordset
    If RSBook.RecordCount > 0 Then
        txt(BookDate).TEXT = RSBook!Ord_Date
        BookDocId = RSBook!OrdDocId
        txt(Model).TEXT = RSBook!Model
        txt(FB_Code).Tag = IIf(IsNull(RSBook!FB_Code), "", RSBook!FB_Code)
        If txt(FB_Code).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select fincode as code,finname as name,AcCode from ContractFinance where fincatg = 0 and  fincode = '" & txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
            txt(FB_Code).TEXT = Rst!Name
        Else
            txt(FB_Code).TEXT = ""
        End If
        If Not IsNull(RSBook!Inv_Date) Then
            If mVType = mCustCRN Then
                 GSQL = "SELECT V.Tax_Per,V.Tax_Amt,V.Surcharge_Per,V.Surcharge_Amt,V.TOT_Per,V.TOT_Amt," & _
                            "T.Tax_Ac_Code,T.Sur_Ac_Code,T.PurSal_Ac_Code " & _
                            "from Veh_Order as V left Join TaxFormsAc as T on V.Form_Code&'" & PubDivCode & "'=T.Form_Code&T.Div_Code " & _
                            "where V.OrdDocId='" & txt(BookNo).Tag & "'"
                
                Set Rst = New Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
                txt(TaxPer) = IIf(IsNull(Rst!Tax_Per), "", Rst!Tax_Per)
                txt(TaxSurPer) = IIf(IsNull(Rst!surcharge_per), "", Rst!surcharge_per)
                txt(TOTPer) = IIf(IsNull(Rst!TOT_Per), "", Rst!TOT_Per)
                
                'mHead = Rst!PurSal_Ac_Code 'Veh Amount
                 mTaxAcHead = Rst!Tax_Ac_Code
                 mTaxSurAcHead = Rst!Sur_Ac_Code
                 mTOTAcHead = G_FaCn.Execute("Select TOTax_Ac From AcControls").Fields(0)
             End If
         End If
        Set Rst = Nothing
    Else
        txt(BookDate).TEXT = ""
        BookDocId = ""
        txt(Model).TEXT = ""
        txt(FB_Code).Tag = ""
        txt(FB_Code).TEXT = ""
        txt(VehAmt) = ""
        txt(TaxPer) = ""
        txt(TaxAmt) = ""
        txt(TaxSurPer) = ""
        txt(TaxSurch) = ""
        txt(TOTPer) = ""
        txt(TOTAmt) = ""
    End If
Set Rst = Nothing
End Sub

Private Sub TxtGrid2_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid2(Index)
    Grid_Hide
    txtgrid2(0).MaxLength = 0
    txtgrid2(0).Tag = FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col)
    Select Case FGrid2.Col
        Case Col_ChqNo
            txtgrid2(0).MaxLength = 10
    End Select
End Sub
Private Sub TxtGrid2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim MaxCol As Byte
    If KeyCode = vbKeyEscape Then
        txtgrid2(0).TEXT = txtgrid2(0).Tag
        TxtGrid2_KeyUp Index, KeyCode, Shift
        FGrid2.SetFocus
        txtgrid2(0).Visible = False
        Exit Sub
    End If
    
    Select Case FGrid2.Col
        Case Col_ChqNo
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid2Leave = True Then
                    DoEvents
                    GridTxtDown FGrid2, txtgrid2, Index, KeyCode, TAddMode, Col_ChqNo
                End If
            End If
            
        Case Col_ChqDate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid2Leave = True Then
                    DoEvents
                    GridTxtDown FGrid2, txtgrid2, Index, KeyCode, TAddMode, Col_ChqAmt
                End If
            End If
        Case Col_ChqAmt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid2Leave = True Then
                    DoEvents
                    GridTxtDown FGrid2, txtgrid2, Index, KeyCode, TAddMode, Col_ChqNo
                End If
            End If
            
    End Select
        
End Sub
Private Sub txtGrid2_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case FGrid2.Col
    End Select
End Sub
Private Sub TxtGrid2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid2.Col
                
        End Select
End Sub

Private Sub TxtGrid2_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid2Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub
Private Function TxtGrid2Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim Repeat$
Dim TempRs As ADODB.Recordset
Dim I As Double, X As Double
With FGrid2
    If .Col = Col_ChqNo Then
        For I = 1 To .Rows - 1
            For X = I + 1 To .Rows - 1
                If UCase(.TextMatrix(I, Col_ChqNo)) = UCase(txtgrid2(0).TEXT) And I <> .Row Then
                    MsgBox "Cheque/DD No Should Not Be Duplicate.", vbOKOnly
                    Exit Function
                End If
            Next X
        Next I
    End If
End With
        
Select Case FGrid2.Col
    Case Col_ChqNo
        FGrid2 = txtgrid2(0)
    Case Col_ChqDate
        FGrid2 = RetDate(txtgrid2(0))
    Case Col_ChqAmt
    FGrid2 = Format(txtgrid2(0), "0.00")
End Select
TxtGrid2Leave = True
If ValidateCall = False Then
    FGrid2.SetFocus
    txtgrid2(0).Visible = False
End If
End Function



'
'************ PRINTING CODE ******************
Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case FromVno, ToVno
        RsVno.Close
        RsVno.Open "Select v_no as code from rect where right(Rect.Site_Code,1)='" & txtPrint(SiteCode1).Tag & "' and  Rect.V_Type='" & txtPrint(VType1).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        Set DGVno.DataSource = RsVno
        If txtPrint(Index).TEXT <> RsVno!Code Then
            RsVno.MoveFirst
            RsVno.FIND "code ='" & txtPrint(Index).TEXT & "'"
        End If
        If Index = ToVno Then DGVno.Tag = "1" Else DGVno.Tag = "2"
    Case VType1
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txtPrint(Index).TEXT = "" Then Exit Sub
        If txtPrint(Index).TEXT <> RsVType!Name Then
            RsVType.MoveFirst
            RsVType.FIND "Name ='" & txtPrint(Index).TEXT & "'"
        End If
    Case SiteCode1
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txtPrint(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txtPrint(Index).Tag = RsSite!Code
            txtPrint(Index).TEXT = RsSite!Name
        Else
            If txtPrint(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txtPrint(Index).TEXT & "'"
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
    Case FromVno, ToVno
        DGridTxtKeyDown DGVno, txtPrint, Index, RsVno, KeyCode, False, 0
    Case VType1
        DGridTxtKeyDown DGVType, txtPrint, Index, RsVType, KeyCode, False, 1
    Case SiteCode1
        DGridTxtKeyDown DGSite, txtPrint, Index, RsSite, KeyCode, False, 1
End Select
If DGVType.Visible = False And DGSite.Visible = False And DGVno.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> SiteCode1 Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case FromVno, ToVno
        If DGVno.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVno, KeyAscii, "Code"
    Case VType1
        If DGVType.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVType, KeyAscii, "name"
    Case SiteCode1
        If DGSite.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsSite, KeyAscii, "Name"
End Select

'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
  Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case ToVno, FromVno
        If RsVno.RecordCount = 0 Or (RsVno.EOF = True Or RsVno.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
        Else
            txtPrint(Index).TEXT = RsVno!Code
        End If
    Case VType1
        If IsValid(txtPrint(Index), "Voucher Type") = False Then Cancel = True: Exit Sub
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
            txtPrint(Index).Tag = ""
        Else
            txtPrint(Index).TEXT = RsVType!Name
            txtPrint(Index).Tag = RsVType!Code
        End If
    Case SiteCode1
        If IsValid(txtPrint(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
            txtPrint(Index).Tag = ""
        Else
            txtPrint(Index).TEXT = RsSite!Name
            txtPrint(Index).Tag = RsSite!Code
        End If
End Select
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
'On Error GoTo ERRORHANDLER

Dim mChqStr As String
Dim I As Integer

For I = 0 To FGrid2.Rows - 1
    If FGrid2.TextMatrix(I, Col_ChqNo) <> "" Then
        If mChqStr <> "" Then mChqStr = mChqStr + ", "
        mChqStr = mChqStr + FGrid2.TextMatrix(I, Col_ChqNo)
    End If
Next I


GSQL = "SELECT SG.NamePrefix, SG.Name as PartyName,SG.FPrefix,SG.FName,SG.Add1,SG.Add2,SG.Add3,SG.PANNo,SG.ITWARD_NO, City.CityName, SG1.Name as AcName,Voucher_Type.Description,  Rect.*, Syctrl.SprMoneyRectFooter,model_Grp.ModelGrp_Name as model, VO.Ord_No, VO.Ord_Date,VO.Chassis,Veh_Stock.EngineNo, CF.FinName, '" & mChqStr & "' as ChqStr " & _
    " FROM (((((((((Rect LEFT JOIN " & FaTable("Voucher_Type") & " ON Rect.V_Type = Voucher_Type.V_Type) " & _
    " LEFT JOIN SubGroup SG on Rect.PartyCode = SG.SubCode ) " & _
    " LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
    " LEFT JOIN SubGroup SG1 ON Rect.AcCode = SG1.SubCode) " & _
    " LEFT JOIN Veh_Order VO ON Rect.Ord_DocId = VO.OrdDocId) " & _
    " LEFT JOIN Veh_Stock ON Veh_Stock.Sal_DocId = VO.Inv_DocId) " & _
    " LEFT JOIN Model on Model.Model = VO.Model ) " & _
    " LEFT JOIN Model_Grp ON Model.Grp_Code = Model_Grp.ModelGrp_Code) " & _
    " LEFT JOIN ContractFinance CF ON VO.FB_CODE = CF.FinCode) " & _
    " LEFT JOIN Syctrl ON Syctrl.LinkTable  > Rect.U_AE   " & _
    " where Rect.Docid ='" & Master!SearchCode & "'"

Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "SprCustRect", "SprCustRect")
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(GSQL)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "SprCustRect", "SprCustRect")
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

Dim I As Integer
Dim mDivSName$
Dim Rst As ADODB.Recordset
Dim RST1 As ADODB.Recordset
Dim RstSub As ADODB.Recordset

On Error GoTo ERRORHANDLER
'    DataPathFA = PubVFADataPath
    mDivSName = IIf(PubDivSName = "", "", "-" & PubDivSName & " ")
 
   
'    mQry = "SELECT SG.NamePrefix, SG.Name as PartyName,SG.FPrefix,SG.FName,SG.Add1,SG.Add2,SG.Add3,SG.PANNo,SG.ITWARD_NO, City.CityName, SG1.Name as AcName,Voucher_Type.Description,  Rect.*, Syctrl.SprMoneyRectFooter,VO.model, VO.Ord_No, VO.Ord_Date, CF.FinName " & _
        " FROM ((((((Rect LEFT JOIN [" & PubVFADataPath & "].Voucher_Type ON Rect.V_Type = Voucher_Type.V_Type) " & _
        " LEFT JOIN SubGroup SG on Rect.PartyCode = SG.SubCode ) " & _
        " LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
        " LEFT JOIN SubGroup SG1 ON Rect.AcCode = SG1.SubCode) " & _
        " LEFT JOIN Veh_Order VO ON Rect.Ord_DocId = VO.OrdDocId) " & _
        " LEFT JOIN ContractFinance CF ON VO.FB_CODE = CF.FinCode) " & _
        " LEFT JOIN Syctrl ON Syctrl.LinkTable  > Rect.U_AE   " & _
        " where Rect.Docid ='" & Master!searchcode & "'"
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    
    mQry = "SELECT DocID, Sr, ChqNo, ChqDate, ChqAmt FROM dbo.Rect1 Where DocID = '" & Master!SearchCode & "'"
    Set RstSub = GCn.Execute(mQry)
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    CreateFieldDefFile RstSub, PubRepoPath + "\" & mRepName & "1.ttx", True
    
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstSub
    
        If PubReceiptType = "Vehicle" Then
            Set RST1 = GCn.Execute("select V_SecSpeciality AS S_SecSpeciality,V_SecLST AS S_SecLST,V_SecLST_Date AS S_SecLST_Date,V_SecCST AS S_SecCST,V_SecCST_Date AS S_SecCST_Date,V_SecPhone AS S_SecPhone,V_SecFax AS S_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'")
        ElseIf PubReceiptType = "Spare" Then
            Set RST1 = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
        Else
            Set RST1 = GCn.Execute("select W_SecSpeciality AS S_SecSpeciality,W_SecLST AS S_SecLST,W_SecLST_Date AS S_SecLST_Date,W_SecCST AS S_SecCST,W_SecCST_Date AS S_SecCST_Date,W_SecPhone AS S_SecPhone,W_SecFax AS S_SecFax from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
        End If
                
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("SubTitle")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecSpeciality & "'"
                Case UCase("Phone")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecPhone & "'"
                Case UCase("Fax")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecFax & "'"
                Case UCase("DivSName")
                    rpt.FormulaFields(I).TEXT = "'" & mDivSName & "'"
                Case UCase("AmtPrefix")
                    rpt.FormulaFields(I).TEXT = "'" & PubAmountPrefix & "'"
'                Case UCase("VouType")
'                    If Txt(Vtype).Tag = "G_ABR" Or Txt(Vtype).Tag = "G_BBP" Or Txt(Vtype).Tag = "G_TLR" Then
'                        rpt.FormulaFields(i).Text = "'Bank'"
'                    Else
'                        rpt.FormulaFields(i).Text = "''"
'                    End If
            End Select
        Next
     rpt.Database.SetDataSource Rst
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
                GCn.Execute "update Rect set Printed = 1 where DocID='" & txt(TxtDocID) & "'"
            End If
            Set RST1 = Nothing
            Set rpt = Nothing
        Case PScreen 'screen
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

Private Sub SpeedPrint(mQry As String)
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
    Dim AcAdd, AcCity As String
    Dim I As Integer, j As Integer
    Dim PrintStr As String
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstCust As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mQty As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject, UserPrintStr$, mPName$, mFName$, mDivSName$
    
'    Set RstCust = GCn.Execute("SELECT SG.NamePrefix, SG.Name as PartyName,SG.FPrefix,SG.FName, SG.Add1, SG.Add2, SG.Add3,SG.PANNo,SG.ITWARD_NO, City.CityName, SG1.Name as AcName, Rect.*, Syctrl.SprMoneyRectFooter,VO.model, VO.Ord_No, VO.Ord_Date, CF.FinName " & _
        " FROM (((((Rect " & _
        " LEFT JOIN SubGroup SG on Rect.PartyCode = SG.SubCode) " & _
        " LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
        " LEFT JOIN SubGroup SG1 ON Rect.AcCode = SG1.SubCode) " & _
        " LEFT JOIN Veh_Order VO ON Rect.Ord_DocId = VO.OrdDocId) " & _
        " LEFT JOIN ContractFinance CF ON VO.FB_CODE = CF.FinCode) " & _
        " LEFT JOIN Syctrl ON Syctrl.LinkTable  > Rect.U_AE   " & _
        " where Rect.Docid ='" & Master!searchcode & "'")
 Set RstCust = GCn.Execute(mQry)
''    Set RstCust = GCn.Execute("SELECT SG.NamePrefix, SG.Name as PartyName,SG.FPrefix,SG.FName, SG.Add1, SG.Add2, SG.Add3,SG.PANNo,SG.ITWARD_NO, City.CityName, Voucher_Type.Description, SG1.Name as AcName, Rect.*, Syctrl.SprMoneyRectFooter,VO.model, VO.Ord_No, VO.Ord_Date, CF.FinName " & _
''        " FROM ((((((Rect LEFT JOIN Voucher_Type ON Rect.V_Type = Voucher_Type.V_Type) " & _
''        " LEFT JOIN SubGroup SG on Rect.PartyCode = SG.SubCode ) " & _
''        " LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
''        " LEFT JOIN SubGroup SG1 ON Rect.AcCode = SG1.SubCode) " & _
''        " LEFT JOIN Veh_Order VO ON Rect.Ord_DocId = VO.OrdDocId) " & _
''        " LEFT JOIN ContractFinance CF ON VO.FB_CODE = CF.FinCode) " & _
''        " LEFT JOIN Syctrl ON Syctrl.LinkTable  > Rect.U_AE   " & _
''        " where Rect.Docid ='" & Master!SearchCode & "'")

    If RstCust.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select SprMoneyRectFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
    PageWidth = 80
    PageLength = PubPageLengthHalf
    mHeader = 0   'Ideal 17
    mFooter = 7 + FooterCnt
    'Document Header
    
    Dim PartyHeader$
    Select Case RstCust!V_Type
        Case "G_ABR"
            PartyHeader = "Received with thanks from : "
        Case "G_ACR"
            PartyHeader = "Received with thanks from : "
        Case "G_BBP"
            PartyHeader = "Paid with thanks to : "
        Case "G_BCP"
            PartyHeader = "Paid with thanks to : "
        Case "G_CRN"
            PartyHeader = ""
        Case "G_DRN"
            PartyHeader = ""
        Case "G_TLR"
            PartyHeader = ""
    End Select
    mDivSName = IIf(PubDivSName = "", "", "-" & PubDivSName & " ")
    '********* < Rahul U.N.Automobiles 10-04-2003  >
    Set GRs = G_FaCn.Execute("Select Description From Voucher_Type Where V_Type='" & RstCust!V_Type & "'")
    If GRs.RecordCount > 0 Then
        mDocStr = GRs!Description
    End If
    Set GRs = Nothing
    mDupStr = IIf(RstCust!Printed = 1, "(Duplicate)", "")
    If PubReceiptType = "Vehicle" Then
        Set RstCompDet = GCn.Execute("select V_SecSpeciality AS S_SecSpeciality,V_SecLST AS S_SecLST,V_SecLST_Date AS S_SecLST_Date,V_SecCST AS S_SecCST,V_SecCST_Date AS S_SecCST_Date,V_SecPhone AS S_SecPhone,V_SecFax AS S_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'")
    ElseIf PubReceiptType = "Spare" Then
        Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    Else
        Set RstCompDet = GCn.Execute("select W_SecSpeciality AS S_SecSpeciality,W_SecLST AS S_SecLST,W_SecLST_Date AS S_SecLST_Date,W_SecCST AS S_SecCST,W_SecCST_Date AS S_SecCST_Date,W_SecPhone AS S_SecPhone,W_SecFax AS S_SecFax,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    End If
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
  
    Print #1, Chr(27) + Chr(67) + Chr(36) & PRN_TIT(PubComp_Name, "A", PageWidth) 'small paper size

'    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!S_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
    If PubComp_Add2 <> "" Or PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_Add2 & IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", "   Fax : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PRN_TIT("* " & UCase(mDocStr) & mDivSName & mDupStr & " *", "A", PageWidth) & mEmph
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PSTR(mDocStr & " No. : " & PubDivCode & XNull(RstCust!Site_Code) & "/" & RstCust!V_NO, 40) & PSTR(mDocStr & " Date : " & RstCust!V_DATE, 40, , AlignRight) + mEmph1
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    If RstCust!PrintParty_YN = 1 Then
        mPName = Trim(RstCust!NamePrefix) & " " & Trim(RstCust!PartyName)
        mFName = Trim(RstCust!FPrefix) & " " & Trim(RstCust!fname)
    Else
        mPName = XNull(RstCust!RWTF1)
        mFName = ""
    End If
    Print #1, PartyHeader & mEmph & mPName & mEmph1
    mHeader = mHeader + 1
    If mFName <> "" Then
        Print #1, Space(Len(PartyHeader)) & mEmph & mFName & mEmph1
        mHeader = mHeader + 1
    End If
    If RstCust!PrintParty_YN = 1 Then
       Print #1, Space(20) & XNull(RstCust!Add1) & IIf(XNull(RstCust!Add1) <> "" And XNull(RstCust!Add2) <> "", ",", "") & XNull(RstCust!Add2)
       mHeader = mHeader + 1
       Print #1, Space(20) & XNull(RstCust!Add3) & IIf(XNull(RstCust!CityName) <> "" And XNull(RstCust!Add3) <> "", ",", "") & XNull(RstCust!CityName)
       mHeader = mHeader + 1
    Else
       Print #1, Space(20) & XNull(RstCust!RWTF2)
       mHeader = mHeader + 1
       Print #1, "On Account of : " & RstCust!PartyName
       mHeader = mHeader + 1
       Print #1, "Debit A/c " & XNull(RstCust!AcName) & mEmph
       mHeader = mHeader + 1
    End If
    Print #1, "The sum of Rs." & mEmph & Format(RstCust!Amount, "0.00") & mChr17 & "( " & ntow(RstCust!Amount, "Rupees", "Paise") & " )" & mEmph1 & mChr18
    mHeader = mHeader + 1
    
    Select Case RstCust!V_Type
        Case "G_ABR", "G_BBP"
            If XNull(RstCust!ChqStr) = "" Then
                PrintStr = "By Pay Order/Chq/Draft No." & RstCust!DDNo & " Dated " & RstCust!DDDate & " Drawn On " & RstCust!Narration
            Else
               ' PrintStr = "By Pay Order/Chq/Draft No." & XNull(RstCust!ChqStr) & " Drawn On " & RstCust!Narration
                Print #1, "By Pay Order/Chq/Draft No." & RstCust!DDNo & " Dated " & RstCust!DDDate & " Drawn On " & RstCust!Narration
                
            End If
            If Len(PrintStr) > 75 Then
                Print #1, left(PrintStr, 75) & " -"
                Print #1, Right(PrintStr, Len(PrintStr) - 75)
            Else
                Print #1, PrintStr
            End If
            mHeader = mHeader + 1
        Case "G_ACR", "G_BCP"
            Print #1, "By Cash " & RstCust!Narration
            mHeader = mHeader + 1
        Case "G_CRN", "G_DRN"
            Print #1, "Due To " & RstCust!Narration
            mHeader = mHeader + 1
        Case "G_TLR"
            Print #1, IIf(RstCust!DDNo = "", "By Cash " & RstCust!Narration, "By PO/Ch/Draft No " & RstCust!DDNo & " Dated " & RstCust!DDDate & " Drawn On " & RstCust!Narration)
            mHeader = mHeader + 1
    End Select
    If RstCust!Vehicle_YN = 1 Then
        Print #1, PSTR("Towards :", 10) & PSTR("Booking of", 25) & ": " & RstCust!Model
        mHeader = mHeader + 1
        Print #1, Space(10) & PSTR("Booking No. & Date", 25) & ": " & PSTR(RstCust!Ord_No, 8, , AlignRight) & " " & RstCust!Ord_Date
        mHeader = mHeader + 1
        Print #1, Space(10) & PSTR("Financed/Hypothecated By", 25) & ": " & RstCust!FinName & IIf(mID(RstCust!DocID, 4, 5) = "G_CRN", Space(8) & PSTR("Veh Amt.", 8) & ": " & PSTR(Format(RstCust!Veh_Amt, "0.00"), 9, , AlignRight), "")
        mHeader = mHeader + 1
        If mID(RstCust!DocID, 4, 5) = "G_CRN" Then
            Print #1, Space(10) & PSTR("Chassis No.", 25) & ": " & RstCust!Chassis & Space(25) & PSTR("Tax Amt.", 8) & ": " & PSTR(Format(RstCust!Tax_Amt, "0.00"), 9, , AlignRight)
            mHeader = mHeader + 1
            Print #1, Space(10) & PSTR("Engine No.", 25) & ": " & RstCust!EngineNo & Space(24) & PSTR("Sur Amt.", 8) & ": " & PSTR(Format(RstCust!Surcharge_Amt, "0.00"), 9, , AlignRight)
            mHeader = mHeader + 1
        End If
    End If
    If IsNull(RstCust!PanNo) Or RstCust!PanNo = "" Then
        Print #1, Space(10) & PSTR("Declaration under", 25) & ": " & RstCust!IForm & IIf(mID(RstCust!DocID, 4, 5) = "G_CRN", Space(18) & PSTR("Tot Amt.", 8) & ": " & PSTR(Format(RstCust!Tot_Amt, "0.00"), 9, , AlignRight), "")
    Else
        Print #1, Space(10) & PSTR("PAN/GIR NO.      ", 25) & ": " & RstCust!PanNo & IIf(mID(RstCust!DocID, 4, 5) = "G_CRN", Space(18) & PSTR("Tot Amt.", 8) & ": " & PSTR(Format(RstCust!Tot_Amt, "0.00"), 9, , AlignRight), "")
    End If
    mHeader = mHeader + 1
    Do Until mHeader <= PageLength - mFooter
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
        Print #1, Space(20) & " " & Replace(Space(8), " ", "-") & Space(8) & "Catg.: " & RstCust!RectCatg
        Print #1, Space(20) & "|" & Space(8) & "|"
        Print #1, Space(20) & "|" & Space(8) & "|"
    '    Print #1, mUnd & PSTR("Rs. " & Amount_Fill(RstCust!AMOUNT, PubAmountPrefix), 20) & mUnd1 & "|" & Space(8) & "|" & PSTR("For " & PubComp_Name, 50, , AlignRight) & mEmph1
        Print #1, Space(20) & "|" & Space(8) & "|" & PSTR("For " & PubComp_Name, 50, , AlignRight) & mEmph1
        Print #1, Space(20) & " " & Replace(Space(8), " ", "-")
    Else
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
          Print #1, ""
    End If
   
    Print #1, ""
    Print #1, PSTR("Cashier", 20) & PSTR("Party Signature", 25) & PSTR("Authorised Signatory", PageWidth - 45, , AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
    
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    
    UserPrintStr = "* a dataman software *" & "By: " & RstCust!U_Name & "   " & RstCust!U_EntDt
    Print #1, "* a dataman software .*" & Space((PageWidth * 1.7) - Len(UserPrintStr)) & "By: " & RstCust!U_Name & "   " & RstCust!U_EntDt & mChr18
    Print #1, Chr(27) + Chr(67) + Chr(72)
    Print #1, ""
    Print #1, ""
    Print #1, ""
        
    If UCase(left(PubComp_Name, 5)) = "SOCIE" And txt(VType).Tag = "G_ABR" Then
        For j = 1 To 4
            Print #1, ""
        Next
        Print #1, mChr14 & mEmph & txt(AcHead).TEXT & mEmph1 & mChr18
        AcAdd = GCn.Execute("Select Add1 & Add2 from SubGroup where SubCode='" & txt(AcHead).Tag & "'").Fields(0)
       'AcCity = GCn.Execute("Select isNull(City.cITYName,'') from SubGroup Left Join City on City.CityCode=SubGroup.CityCode where SubGroup.SubCode='" & Txt(AcHead).Tag & "'").Fields(0)
        Print #1, SETW(AcAdd & AcCity, 32) & Space(20) & Chr(27) & Chr(69) & "CHEQUE DEPOSIT SLIP" & Chr(27) & Chr(70)
        Print #1, "Account Title:" & SETW(PubComp_Name, 35) & Space(5) & "Date:" & SETW(txt(VDate), 12)
        Print #1, Space(54) & "Bank Rec.No." & SETW(txt(SerialNo), 12)
        Print #1, "Rs:" & ntow(Val(txt(Amt)), "Rupees", "Paise")
        Print #1, mChr14
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, " Customer            Chq/DD No.   Bank/Branch         Amount Realisation Remark "
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, SETW(txt(Party), 20) & Space(1) & SETW(txt(DDNo), 12) & Space(1) & SETW(txt(Narr), 15) & Space(1) & SETN(Format(txt(Amt), "0.00"), 10)
        For j = 1 To 5
            Print #1, ""
        Next
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, Space(43) & "Total: " & SETN(Format(txt(Amt), "0.00"), 10)
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, ""
        Print #1, Space(40) & "DEPOSITED               RECEIVED/APPROVED"
        Print #1, mChr18
         
        
    End If
    
    
        
    
    
    Close #1
    
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
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
        GCn.Execute "update Rect set Printed = 1 where DocID='" & Master!SearchCode & "'"
    End If
    Exit Sub
ELoop:
   ' Close #1: CheckError
    'EOF Speed Printing Section
    MsgBox err.Description & vbCr & "In PreedPrint Procedure Of " & Me.Name
End Sub

Private Sub Ini_Grid()
   
    
    With FGrid2
        .Cols = 4
        .width = 4000
        
        .TextMatrix(0, 0) = "Srl"
        .TextMatrix(1, 0) = "1"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 500
        
        .TextMatrix(0, Col_Code) = ""
        .ColAlignment(Col_Code) = flexAlignLeftCenter
        .ColWidth(Col_Code) = 0
        
        .TextMatrix(0, Col_ChqNo) = "Chq No."
        .ColAlignment(Col_ChqNo) = flexAlignLeftCenter
        .ColWidth(Col_ChqNo) = 1200
        
        .TextMatrix(0, Col_ChqDate) = "Chq Dt."
        .ColAlignment(Col_ChqDate) = flexAlignLeftCenter
        .ColWidth(Col_ChqDate) = 1200
        
        .TextMatrix(0, Col_ChqAmt) = "Amount"
        .ColAlignment(Col_ChqAmt) = flexAlignRightCenter
        .ColWidth(Col_ChqAmt) = 1200
        
    End With
    
    BackColorSelLeave = FGrid2.BackColorSel
    ForeColorSelEnter = FGrid2.ForeColorSel
        
End Sub

