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
        For I = 0 To Txt.Count - 1
            Txt(I).Refresh
        Next
        If CDate(Format(Txt(VDate).TEXT, "dd/MMM/yyyy")) < CDate("01/Apr/2005") Then
            'MsgBox ""
        End If
        If CDate(Format(Txt(VDate).TEXT, "dd/MMM/yyyy")) < PubStartDate Or CDate(Format(Txt(VDate).TEXT, "dd/MMM/yyyy")) > PubEndDate Then GoTo MyNextRecord
        Call TopCtrl1_eEdit
        'A/c Posting
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            Select Case Txt(VType).Tag
                Case "G_ABR", "G_ACR", "G_TLR", "G_JV"   'Receipt
                    I = -1
                    
                    If Not mMultipleChqNo Then
                        I = I + 1
                        LedgAry(I).SubCode = Txt(AcHead).Tag
                        LedgAry(I).AmtDr = Val(Txt(Amt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = Txt(Party).Tag
                    Else
                        For j = 0 To FGrid2.Rows - 1
                            I = I + 1
                            
                            LedgAry(I).SubCode = Txt(AcHead).Tag
                            LedgAry(I).AmtDr = Val(FGrid2.TextMatrix(j, Col_ChqAmt))
                            LedgAry(I).Chq_No = FGrid2.TextMatrix(j, Col_ChqNo)
                            LedgAry(I).Chq_Date = FGrid2.TextMatrix(j, Col_ChqDate)
                            LedgAry(I).Narration = mNarr
                            LedgAry(I).ContraSub = Txt(Party).Tag
                            
                        Next j
                    End If
                    I = I + 1
                    LedgAry(I).SubCode = Txt(Party).Tag
                    LedgAry(I).AmtCr = Val(Txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = Txt(AcHead).Tag
                Case "G_BBP", "G_BCP", "G_DRN"   'payment
                    I = 0
                    LedgAry(I).SubCode = Txt(Party).Tag
                    LedgAry(I).AmtDr = Val(Txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = Txt(AcHead).Tag
                    
                    
                    If Not mMultipleChqNo Then
                    
                        I = I + 1
                        LedgAry(I).SubCode = Txt(AcHead).Tag
                        LedgAry(I).AmtCr = Val(Txt(Amt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = Txt(Party).Tag
                    Else
                        For j = 0 To FGrid2.Rows - 1
                            I = I + 1
                            
                            LedgAry(I).SubCode = Txt(AcHead).Tag
                            LedgAry(I).AmtCr = Val(FGrid2.TextMatrix(j, Col_ChqAmt))
                            LedgAry(I).Chq_No = FGrid2.TextMatrix(j, Col_ChqNo)
                            LedgAry(I).Chq_Date = FGrid2.TextMatrix(j, Col_ChqDate)
                            LedgAry(I).Narration = mNarr
                            LedgAry(I).ContraSub = Txt(Party).Tag
                            
                        Next j
                    
                    End If
                    
                Case mCustCRN
                    I = 0
                    LedgAry(I).SubCode = Txt(Party).Tag
                    LedgAry(I).AmtCr = Val(Txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = Txt(AcHead).Tag
                    I = I + 1
                    LedgAry(I).SubCode = Txt(AcHead).Tag
                    If Val(Txt(VehAmt)) + Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(TOTAmt)) = 0 Then
                        LedgAry(I).AmtDr = Val(Txt(Amt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = Txt(Party).Tag
                    Else
                        LedgAry(I).AmtDr = Val(Txt(VehAmt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
        
                        I = I + 1
                        LedgAry(I).SubCode = mTaxAcHead
                        LedgAry(I).AmtDr = Val(Txt(TaxAmt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
                        I = I + 1
                        LedgAry(I).SubCode = mTaxSurAcHead
                        LedgAry(I).AmtDr = Val(Txt(TaxSurch))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
                        I = I + 1
                        LedgAry(I).SubCode = mTOTAcHead
                        LedgAry(I).AmtDr = Val(Txt(TOTAmt))
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = "" 'txt(Party).Tag
                    End If
            End Select
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, Txt(TxtDocID), CDate(Txt(VDate)), mNarr)
            If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        End If
        'eof posting
MyNextRecord:
        Disp_Text SETS("INI", Me, Master)
        Master.MoveNext
    Loop

End Sub

Private Sub DGCity_Click()
    DGCity.Visible = False
    If RsCity.RecordCount > 0 Then
        Txt(RecLocation).Tag = RsCity!Code
        Txt(RecLocation).TEXT = RsCity!Name
    End If
    Txt(RecLocation).SetFocus
End Sub

Private Sub DGSite_Click()
If FrmPrn.Visible = False Then
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        Txt(SiteCode).TEXT = RsSite!Name
        Txt(SiteCode).Tag = RsSite!Code
    End If
    Txt(SiteCode).SetFocus
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
        Txt(BookNo).TEXT = RSBook!Code
        FillRecords RSBook
    End If
    Txt(BookNo).SetFocus
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        Txt(Party).TEXT = RsParty!Name
        Txt(Party).Tag = RsParty!Code
    End If
    DGParty.Visible = False
    Txt(Party).SetFocus
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
        Txt(VType).TEXT = RsVType!Name
        Txt(VType).Tag = RsVType!Code
    End If
    DGVType.Visible = False
    Txt(VType).SetFocus
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
    PubUParam = MDIForm1.Permission("Customer Receipt")
    TopCtrl1.Tag = PubUParam: WinSetting Me, 5760, , 850, 465
    DGBook.left = 0: DGBook.width = Me.width - 90: DGBook.top = Me.height - (DGBook.height + mBotScale)
    DGParty.left = 0: DGParty.width = Me.width - 90: DGParty.top = Txt(Amt).top: DGParty.height = Me.height - (DGParty.top + mBotScale)
    DGSite.left = 4500: DGSite.top = mTopScale
    DGVno.left = 4500: DGVno.top = mTopScale
    DGVType.left = 4500: DGVType.top = mTopScale
    DGCity.left = 4500: DGCity.top = mTopScale
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
        Txt(9).Visible = False
        Txt(11).Visible = False
        Txt(12).Visible = False
        Txt(15).Visible = False
        Txt(16).Visible = False
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
    Set DGCity.DataSource = RsCity
    
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
Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
Txt(Val(ListView.Tag)).SetFocus
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
    Txt(VehTrnYN) = IIf(PubVCompCode <> "", "No", "")
    Txt(PrnNameYN) = "Yes"
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        Txt(SiteCode).Tag = PubSiteCode
        Txt(SiteCode) = PubSiteName
        Txt(VDate).SetFocus
    Else
        Txt(SiteCode).SetFocus
    End If
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean ',i As Integer
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$

If AcPostAuthorisation(Txt(AcPostByName)) = False Then Exit Sub

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
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, Txt(TxtDocID))
    If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    'Unposting of Ledger completed
    If GCn.Execute("Select CancelYN from Rect where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
        GCn.Execute ("delete from Rect where DocId='" & Txt(TxtDocID) & "'")
    Else
        GCn.Execute "update rect set " & _
            "CancelYN=1,AMOUNT =0, AcCode= '" & Txt(AcHead).Tag & "'," & _
            "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " & _
            " where docid = '" & Txt(TxtDocID) & "'"
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
    If AcPostAuthorisation(Txt(AcPostByName)) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    Txt(Party).SetFocus
    If Txt(PrnNameYN) = "Yes" Then
        Txt(Rem1).Enabled = False
        Txt(Rem2).Enabled = False
        Txt(Rem1).BackColor = CtrlBColDisabled
        Txt(Rem2).BackColor = CtrlBColDisabled
        Txt(2).Enabled = True
    Else
        Txt(Rem1).Enabled = True
        Txt(Rem2).Enabled = True
        Txt(Rem1).BackColor = CtrlBColOrg
        Txt(Rem2).BackColor = CtrlBColOrg
        
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
    If IsValid(Txt(SiteCode), "Site Name") = False Then Exit Sub
    If IsValid(Txt(VDate), "Date") = False Then Exit Sub
    If IsValid(Txt(VType), "Voucher Type") = False Then Exit Sub
    If Txt(SerialNo).Enabled = True Then
        If Txt(SerialNo).TEXT = "" Then MsgBox "SerialNo is required field", vbInformation, "Validation Check": Txt(SerialNo).SetFocus: Exit Sub
    Else
        If Txt(SerialNo).TEXT = "" Then MsgBox "SerialNo is required field", vbInformation, "Validation Check": Txt(VType).SetFocus: Exit Sub
    End If
    If Txt(VehTrnYN) = "Yes" Then
        If IsValid(Txt(BookNo), "Booking No.") = False Then Exit Sub
    End If
    If Txt(Party).Tag = Txt(AcHead).Tag Then
        MsgBox "Party A/c and Ledger A/c both same !" & vbCrLf & "Correct A/c Selection ", vbCritical, "A/c Checking"
        Txt(AcHead).SetFocus: Exit Sub
    End If
    If Val(Txt(Amt)) <= 0 Then
        MsgBox "Please Enter Amount", vbCritical, "Validation"
        Txt(Amt).SetFocus: Exit Sub
    End If
    If Val(Txt(DiscAmt)) > 0 And Txt(DiscAcName).Tag = "" Then
        MsgBox "Please Enter Disc. A/c", vbCritical, "Validation"
        Txt(DiscAcName).SetFocus: Exit Sub
    End If
    '********* cHECKING pOSTING cOTROLS
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        Txt(AcPostByName) = pubUName
        Txt(AcPostDate) = PubServerDate
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
        
        If mSum <> Val(Txt(Amt)) And mSum > 0 Then
            MsgBox "Detail of Cheque/DD can't be differ from total Amount"
            FGrid2.SetFocus
            Exit Sub
        ElseIf mSum > 0 Then
            mMultipleChqFlag = True
        End If
    End If
    
    If TopCtrl1.TopText2.CAPTION = "Add" Then
    'lp 11-03-03
        DocID = Txt(TxtDocID)
        If GCn.Execute("select count(*) from rect where DocId='" & Txt(TxtDocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
                MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
                Txt(SerialNo).SetFocus
                Exit Sub
            Else
                Txt(TxtDocID) = GetDocID(G_FaCn, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    SetMax_VoucherPrefix "DocID", Txt(VType).Tag, "Rect", "V_date"
                    Txt(TxtDocID) = GetDocID(G_FaCn, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                    If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                        MsgBox "Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    DocIdHlp = Replace(Txt(TxtDocID), " ", "")
    Select Case Txt(VType).Tag
        Case "G_ABR", "G_ACR", "G_DRN", "G_TLR", "G_JV"
            mDrCr = "C"
        Case "G_BBP", "G_BCP", "G_CRN"
            mDrCr = "D"
    End Select

    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    mNarr = Txt(Narr)
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute "insert into rect(DocId,DocIDHelp,V_Date,V_Type,V_No,site_code, " & _
            "Prov_No,Prov_Date,Prov_Location, " & _
            "Vehicle_YN,Ord_SiteCode,Ord_DocId,PartyCode," & _
            "Veh_Amt,Tax_Amt,Surcharge_Amt,TOT_Amt,AMOUNT, " & _
            "DrCr,Narration,AcCode, " & _
            "DDNo,DDDate,PrintParty_YN," & _
            "RWTF1,RWTF2,IFORM,RectCatg,CancelYN, CreditCardNo," & _
            "U_Name,U_EntDt,U_AE,AcPostByU_Name,AcPostByU_EntDt,AddBy, AddDate,DiscAc,DiscAmt) values( " & _
            "  '" & Txt(TxtDocID) & "','" & DocIdHlp & "'," & ConvertDate(Txt(VDate)) & ",'" & mVType & "'," & Val(Txt(SerialNo)) & ",'" & PubSiteCode & Txt(SiteCode).Tag & _
            "', " & Val(Txt(ProNo)) & "," & ConvertDate(Txt(ProDate)) & ",'" & Txt(RecLocation).Tag & _
            "', " & IIf(Txt(VehTrnYN) = "Yes", 1, 0) & ",'" & mID(Txt(BookNo).Tag, 2, 2) & "','" & Txt(BookNo).Tag & "','" & Txt(Party).Tag & _
            "', " & Val(Txt(VehAmt)) & ", " & Val(Txt(TaxAmt)) & ", " & Val(Txt(TaxSurch)) & "," & Val(Txt(TOTAmt)) & "," & Val(Txt(Amt)) & _
            " ,'" & mDrCr & "','" & mNarr & "', '" & Txt(AcHead).Tag & _
            "','" & Txt(DDNo).TEXT & "'," & ConvertDate(Txt(DDDate)) & ", " & IIf(Txt(PrnNameYN) = "Yes", 1, 0) & _
            " ,'" & Txt(Rem1) & "','" & Txt(Rem2) & "','" & Txt(FormType).TEXT & "','" & Txt(RectCatg).TEXT & "',0, '" & Txt(CreditCardNo) & "'," & _
            "  '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & Txt(AcPostByName) & "'," & ConvertDate(Txt(AcPostDate)) & ", '" & pubUName & "', " & ConvertDateTime(PubServerDate) & ",'" & Txt(DiscAcName).Tag & "'," & Val(Txt(DiscAmt)) & " )"
            'Voucher Serial No. Updation LPS 21-05-03
            'update Table only when DocSrlNo >Table.SerialNo
            UpdVouSrlNo G_FaCn, Txt(TxtDocID), Txt(VDate)
    Else
        GCn.Execute "update rect set V_date = " & ConvertDate(Txt(VDate)) & ",RectCatg='" & Txt(RectCatg).TEXT & "', Prov_No=" & Val(Txt(ProNo)) & ", " & _
            "Prov_Date=" & ConvertDate(Txt(ProDate)) & ",Prov_Location='" & Txt(RecLocation).Tag & "', Vehicle_YN=" & IIf(Txt(VehTrnYN) = "Yes", 1, 0) & ", " & _
            "Ord_SiteCode='" & mID(Txt(BookNo).Tag, 2, 2) & "',Ord_DocId='" & Txt(BookNo).Tag & "',PartyCode='" & Txt(Party).Tag & "', " & _
            "Veh_Amt=" & Val(Txt(VehAmt)) & ",Tax_Amt= " & Val(Txt(TaxAmt)) & ",Surcharge_Amt= " & Val(Txt(TaxSurch)) & ",TOT_Amt=" & Val(Txt(TOTAmt)) & ",AMOUNT =" & Val(Txt(Amt)) & _
            ",DrCr='" & mDrCr & "' ,Narration='" & mNarr & "',AcCode= '" & Txt(AcHead).Tag & "',DDNo='" & Txt(DDNo).TEXT & "',DDDate=" & ConvertDate(Txt(DDDate)) & " , " & _
            "PrintParty_YN=" & IIf(Txt(PrnNameYN) = "Yes", 1, 0) & ",RWTF1='" & Txt(Rem1) & "',RWTF2='" & Txt(Rem2) & "', " & _
            "IFORM='" & Txt(FormType).TEXT & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E', " & _
            "AcPostByU_Name='" & Txt(AcPostByName) & "',AcPostByU_EntDt=" & ConvertDate(Txt(AcPostDate)) & ", ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ",DiscAc='" & Txt(DiscAcName).Tag & "',DiscAmt=" & Val(Txt(DiscAmt)) & ", CreditCardNo = '" & Txt(CreditCardNo) & "' " & _
            " where docid = '" & Txt(TxtDocID) & "'"
    End If
    
    mQry = "Delete From Rect1 Where DocId = '" & Txt(TxtDocID) & "'"
    GCn.Execute mQry
    For I = 1 To FGrid2.Rows - 1
        If FGrid2.TextMatrix(I, Col_ChqNo) <> "" Then
            mQry = "INSERT INTO dbo.Rect1(DocID,Sr,ChqNo,ChqDate,ChqAmt) " & _
                   "VALUES ('" & Txt(TxtDocID) & "'," & I & ",'" & FGrid2.TextMatrix(I, Col_ChqNo) & "'," & ConvertDate(FGrid2.TextMatrix(I, Col_ChqDate)) & "," & Val(FGrid2.TextMatrix(I, Col_ChqAmt)) & ")"
            GCn.Execute mQry
        End If
    Next I
    
    
    'A/c Posting
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        Select Case Txt(VType).Tag
            Case "G_ABR", "G_ACR", "G_TLR", "G_JV", "SBLCQ", "SBLCS", "SBLRO"  'Receipt
                I = 0
                
                If Not mMultipleChqFlag Then
                    LedgAry(I).SubCode = Txt(AcHead).Tag
                    LedgAry(I).AmtDr = Val(Txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = Txt(Party).Tag
                Else
                    For j = 0 To FGrid2.Rows - 1
                        I = I + 1
                        ReDim Preserve LedgAry(UBound(LedgAry) + 1)
                        LedgAry(I).SubCode = Txt(AcHead).Tag
                        LedgAry(I).AmtDr = Val(FGrid2.TextMatrix(j, Col_ChqAmt))
                        LedgAry(I).Chq_No = FGrid2.TextMatrix(j, Col_ChqNo)
                        LedgAry(I).Chq_Date = FGrid2.TextMatrix(j, Col_ChqDate)
                        LedgAry(I).Narration = mNarr
                        LedgAry(I).ContraSub = Txt(Party).Tag
                    Next j
                End If
                I = I + 1
                LedgAry(I).SubCode = Txt(Party).Tag
                LedgAry(I).AmtCr = Val(Txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = Txt(AcHead).Tag
            Case "G_DRN"
                I = 0
                LedgAry(I).SubCode = Txt(AcHead).Tag
                LedgAry(I).AmtCr = Val(Txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = Txt(Party).Tag
                I = I + 1
                LedgAry(I).SubCode = Txt(Party).Tag
                LedgAry(I).AmtDr = Val(Txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = Txt(AcHead).Tag
            Case "G_BBP", "G_BCP"
                I = 0
                LedgAry(I).SubCode = Txt(Party).Tag
                LedgAry(I).AmtDr = Val(Txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = Txt(AcHead).Tag
                I = I + 1
                LedgAry(I).SubCode = Txt(AcHead).Tag
                LedgAry(I).AmtCr = Val(Txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = Txt(Party).Tag
                
            Case mCustCRN
                I = 0
                LedgAry(I).SubCode = Txt(Party).Tag
                LedgAry(I).AmtCr = Val(Txt(Amt))
                LedgAry(I).Narration = mNarr
                LedgAry(I).ContraSub = Txt(AcHead).Tag
                I = I + 1
                LedgAry(I).SubCode = Txt(AcHead).Tag
                If Val(Txt(VehAmt)) + Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(TOTAmt)) = 0 Then
                    LedgAry(I).AmtDr = Val(Txt(Amt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = Txt(Party).Tag
                Else
                    LedgAry(I).AmtDr = Val(Txt(VehAmt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
    
                    I = I + 1
                    LedgAry(I).SubCode = mTaxAcHead
                    LedgAry(I).AmtDr = Val(Txt(TaxAmt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
                    I = I + 1
                    LedgAry(I).SubCode = mTaxSurAcHead
                    LedgAry(I).AmtDr = Val(Txt(TaxSurch))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
                    I = I + 1
                    LedgAry(I).SubCode = mTOTAcHead
                    LedgAry(I).AmtDr = Val(Txt(TOTAmt))
                    LedgAry(I).Narration = mNarr
                    LedgAry(I).ContraSub = "" 'txt(Party).Tag
                End If
        End Select
        If Txt(VType).Tag = "G_ACR" Then
            If Val(Txt(DiscAmt)) > 0 Then
                I = I + 1
                LedgAry(I).SubCode = Txt(DiscAcName).Tag
                LedgAry(I).AmtDr = Val(Txt(DiscAmt))
                LedgAry(I).Narration = "Being Cash Discount Given to Party"
                LedgAry(I).ContraSub = Txt(Party).Tag
                I = I + 1
                LedgAry(I).SubCode = Txt(Party).Tag
                LedgAry(I).AmtCr = Val(Txt(DiscAmt))
                LedgAry(I).Narration = "Being Cash Discount Given to Party"
                LedgAry(I).ContraSub = Txt(DiscAcName).Tag
            End If
        End If
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, Txt(TxtDocID), CDate(Txt(VDate)), mNarr)
        G_FaCn.Execute ("Update Ledger set chq_no='" & Txt(DDNo) & "',Chq_Date=" & ConvertDate(Txt(DDDate)) & " where DocId='" & Txt(TxtDocID) & "'")
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
                    "from rect Where V_Date>=" & ConvertDate(PubStartDate) & " And DocId = '" & Txt(TxtDocID) & "' order by V_Date desc,docid")
    End If
    RSBook.Requery
    Master.FIND "DocId = '" & Txt(TxtDocID) & "'"

    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(Txt(SerialNo)) > DeCodeDocID(DocID, Document_No) Then
            MsgBox "Document No." & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & Txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
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
If Txt(SiteCode).TEXT <> "" Then
    If Txt(VDate).TEXT = "" Then Txt(VDate).SetFocus: Ctrl_GetFocus Txt(Index): Exit Sub
    If Txt(VType).TEXT = "" Then Txt(VType).SetFocus: Ctrl_GetFocus Txt(Index): Exit Sub
End If
Ctrl_GetFocus Txt(Index)
Grid_Hide
Select Case Index
    Case VType
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsVType!Name Then
            RsVType.MoveFirst
            RsVType.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case RectCatg
        ListArray = Array("     ", "M.M.", "BAL", "FULL", "Staff")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 5)

    Case FormType
        ListArray = Array("Form-60", "Form-61", "N/A")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 3)
    Case SiteCode
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
    Case RecLocation
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsCity!Name Then
            RsCity.MoveFirst
            RsCity.FIND "name ='" & Txt(Index).TEXT & "'"
        End If
    Case Party
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & Txt(Index).TEXT & "'"
        End If
    Case BookNo
         Set RSBook = GCn.Execute("SELECT " & cCStr("Veh_Order.Ord_No") & " as code,Veh_Order.OrdDocId, Veh_Order.Net_AMOUNT, Veh_Order.Ord_SiteCode,  Veh_Order.Ord_Date, Veh_Order.PartyCode, Veh_Order.MODEL, Veh_Order.FB_CODE, Veh_Order.Inv_No, Veh_Order.Inv_Date, Site.Site_Desc, ContractFinance.FinName, sum(" & cIIF("Rect.DrCr = 'D'", "Rect.AMOUNT", "Rect.AMOUNT*-1") & ") as AmtPaid " & _
            "FROM ((Veh_Order LEFT JOIN Site ON right(Veh_Order.Ord_SiteCode,1) = Site.Site_Code) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) LEFT JOIN Rect ON (Veh_Order.OrdDocId = Rect.Ord_DocId) AND (Veh_Order.Ord_SiteCode = Rect.Ord_SiteCode) " & _
            "WHERE Veh_Order.PartyCode = '" & Txt(Party).Tag & "'  " & _
            "group by Veh_Order.OrdDocId, Veh_Order.Ord_SiteCode, Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.PartyCode, Veh_Order.MODEL, Veh_Order.FB_CODE, Veh_Order.Inv_No, Veh_Order.Inv_Date, Site.Site_Desc, ContractFinance.FinName, Veh_Order.Net_AMOUNT")
        Set DGBook.DataSource = RSBook
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RSBook!Code Then
            RSBook.MoveFirst
            RSBook.FIND "code ='" & Txt(Index).TEXT & "'"
        End If
    Case AcHead, DiscAcName
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & Txt(Index).TEXT & "'"
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
        DGridTxtKeyDown DGVType, Txt, Index, RsVType, KeyCode, False, 1
    Case RectCatg
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1500
    Case FormType
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 900
    Case SiteCode
        DGridTxtKeyDown DGSite, Txt, Index, RsSite, KeyCode, False, 1
    Case BookNo
        DGridTxtKeyDown DGBook, Txt, Index, RSBook, KeyCode, False, 0
    Case RecLocation
        DGridTxtKeyDown DGCity, Txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
    Case Party, AcHead, DiscAcName
        DGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
End Select
If FrmList.Visible = False And DGVType.Visible = False And DGCity.Visible = False And DGParty.Visible = False And DGBook.Visible = False And DGSite.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
        If (Txt(Rem2).Enabled = True And Index <> Rem2) Or (Txt(Rem2).Enabled = False And Index <> PrnNameYN) Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (Txt(Rem2).Enabled = True And Index = Rem2) Or (Txt(Rem2).Enabled = False And Index = PrnNameYN) Then
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
        If DGVType.Visible = True Then DGridTxtKeyPress Txt, Index, RsVType, KeyAscii, "name"
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress Txt, Index, RsSite, KeyAscii, "Name"
    Case BookNo
        If DGBook.Visible = True Then DGridTxtKeyPress Txt, Index, RSBook, KeyAscii, "Code"
    Case RecLocation
        If DGCity.Visible = True Then DGridTxtKeyPress Txt, Index, RsCity, KeyAscii, "Name"
    Case SerialNo
        Call NumPress(Txt(Index), KeyAscii, 6, 0)
    Case Party, AcHead, DiscAcName
        If DGParty.Visible = True Then DGridTxtKeyPress Txt, Index, RsParty, KeyAscii, "Name"
    Case VehTrnYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
            mVehYN = True
'            txt(BookNo).Enabled = True
        ElseIf UCase(Chr(KeyAscii)) = "N" Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = "No"
'            txt(BookNo).Enabled = False
'        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
'            txt(Index) = "No"
'            txt(BookNo).Enabled = False
        End If
        If Txt(VehTrnYN) = "Yes" Then
            mVehYN = True
        End If
        Txt(BookNo).Enabled = mVehYN
        If mVType <> mCustCRN Then
            mVehYN = False
        End If
        Txt(VehAmt).Enabled = mVehYN
        Txt(TaxPer).Enabled = mVehYN
        Txt(TaxAmt).Enabled = mVehYN
        Txt(TaxSurPer).Enabled = mVehYN
        Txt(TaxSurch).Enabled = mVehYN
        Txt(TOTPer).Enabled = mVehYN
        Txt(TOTAmt).Enabled = mVehYN
        
        KeyAscii = 0
        If Txt(BookNo).Enabled = False Then
            Txt(BookNo).BackColor = CtrlBColDisabled
            Txt(VehAmt).BackColor = CtrlBColDisabled
            Txt(TaxPer).BackColor = CtrlBColDisabled
            Txt(TaxAmt).BackColor = CtrlBColDisabled
            Txt(TaxSurPer).BackColor = CtrlBColDisabled
            Txt(TaxSurch).BackColor = CtrlBColDisabled
            Txt(TOTPer).BackColor = CtrlBColDisabled
            Txt(TOTAmt).BackColor = CtrlBColDisabled
        Else
            Txt(BookNo).BackColor = CtrlBColOrg
            Txt(VehAmt).Enabled = CtrlBColOrg
            Txt(TaxPer).Enabled = CtrlBColOrg
            Txt(TaxAmt).Enabled = CtrlBColOrg
            Txt(TaxSurPer).Enabled = CtrlBColOrg
            Txt(TaxSurch).Enabled = CtrlBColOrg
            Txt(TOTPer).Enabled = CtrlBColOrg
            Txt(TOTAmt).Enabled = CtrlBColOrg
        End If
    Case VehAmt, TaxAmt, TaxSurch, TOTAmt, Amt
        Call NumPress(Txt(Index), KeyAscii, 8, 2)
    Case PrnNameYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
            Txt(Rem1).Enabled = False
            Txt(Rem2).Enabled = False
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
            Txt(Rem1).Enabled = True
            Txt(Rem2).Enabled = True
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
            Txt(Rem1).Enabled = False
            Txt(Rem2).Enabled = False
        End If
        KeyAscii = 0
        If Txt(Rem1).Enabled = False Then
            Txt(Rem1).BackColor = CtrlBColDisabled
        Else
            Txt(Rem1).BackColor = CtrlBColOrg
        End If
        If Txt(Rem2).Enabled = False Then
            Txt(Rem2).BackColor = CtrlBColDisabled
        Else
            Txt(Rem2).BackColor = CtrlBColOrg
        End If
End Select
'KeyAscii = RetDGKeyAscii()
End Sub



Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs

Select Case Index
    Case FormType, RectCatg
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
    Case VehAmt, TaxAmt, TaxSurch, TOTAmt
        Txt(Amt) = Val(Txt(VehAmt)) + Val(Txt(TaxAmt)) + Val(Txt(TaxSurch)) + Val(Txt(TOTAmt))
    Case CreditCardNo
        If Txt(Index) <> "" Then
            Txt(AcHead).Enabled = False
        Else
            Txt(AcHead).Enabled = True
        End If
End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim I As Integer, mEnb As Boolean, mVehYN As Boolean
Select Case Index
    Case VehTrnYN
        If Txt(Index) <> "Yes" Then
            Txt(BookNo).TEXT = ""
            Txt(BookNo).Tag = ""
            Txt(BookDate).TEXT = ""
            BookDocId = ""
            Txt(Model).TEXT = ""
            Txt(FB_Code).Tag = ""
            Txt(FB_Code).TEXT = ""
            Txt(VehAmt) = ""
            Txt(TaxPer) = ""
            Txt(TaxAmt) = ""
            Txt(TaxSurPer) = ""
            Txt(TaxSurch) = ""
            Txt(TOTPer) = ""
            Txt(TOTAmt) = ""
         Else
            If Txt(VType) <> "Customer Credit Note" Then
                Txt(VehAmt).Enabled = False
                Txt(TaxPer).Enabled = False
                Txt(TaxAmt).Enabled = False
                Txt(TaxSurPer).Enabled = False
                Txt(TaxSurch).Enabled = False
                Txt(TOTPer).Enabled = False
                Txt(TOTAmt).Enabled = False
            End If
            
        End If
        
    Case VType
        If IsValid(Txt(Index), "Voucher Type") = False Then Cancel = True: Exit Sub
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsVType!Name
            Txt(Index).Tag = RsVType!Code
            mVType = Txt(Index).Tag
            If RSOJPR = True Then
                If mVType = mCustCR Then
                    Txt(DiscAcName).Visible = True
                    Txt(DiscAmt).Visible = True
                    Label3(10).Visible = True
                    Label3(11).Visible = True
                Else
                    Txt(DiscAcName).Visible = False
                    Txt(DiscAmt).Visible = False
                    Label3(10).Visible = False
                    Label3(11).Visible = False
                End If
            End If
            If mVType = mCustBP Or mVType = mCustBR Then
                mEnb = True
                Txt(DDNo).BackColor = CtrlBColOrg
                Txt(DDDate).BackColor = CtrlBColOrg
            Else
                Txt(DDNo).BackColor = CtrlBColDisabled
                Txt(DDDate).BackColor = CtrlBColDisabled
            End If
            Txt(DDNo).Enabled = mEnb
            Txt(DDDate).Enabled = mEnb
            If mVType = mCustCRN Then
                mVehYN = True
                Txt(VehAmt).BackColor = CtrlBColOrg
                Txt(TaxAmt).BackColor = CtrlBColOrg
                Txt(TaxSurch).BackColor = CtrlBColOrg
                Txt(TOTAmt).BackColor = CtrlBColOrg
                Txt(VehAmt).Enabled = mVehYN
                Txt(TaxAmt).Enabled = mVehYN
                Txt(TaxSurch).Enabled = mVehYN
                Txt(TOTAmt).Enabled = mVehYN
            Else
                Txt(VehAmt).BackColor = CtrlBColDisabled
                Txt(TaxAmt).BackColor = CtrlBColDisabled
                Txt(TaxSurch).BackColor = CtrlBColDisabled
                Txt(TOTAmt).BackColor = CtrlBColDisabled
            End If
            
            
            
            'DocID
            Txt(TxtDocID) = GetDocID(G_FaCn, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
            DocID = Txt(TxtDocID)
            
            
            If Txt(VType) <> "Customer Bank Receipt" Then
                Txt(CreditCardNo).Enabled = False
            Else
                Txt(CreditCardNo).Enabled = True
            End If
            
        End If
    Case SerialNo
        If IsValid(Txt(SerialNo), "Serial No.") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag = True Then      ' Manual
            Txt(TxtDocID) = GetDocID(G_FaCn, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
            DocID = Txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select * From Rect Where docid='" & Txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                Txt(SerialNo).SetFocus
            End If
        End If
    Case FormType, RectCatg
        If Txt(Index).TEXT <> "" Then Txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case SiteCode
        If IsValid(Txt(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsSite!Name
            Txt(Index).Tag = RsSite!Code
        End If
    Case RecLocation
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsCity!Name
            Txt(Index).Tag = RsCity!Code
        End If
    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
            Txt(BookDate).TEXT = ""
            BookDocId = ""
            Txt(Model).TEXT = ""
            Txt(FB_Code).Tag = ""
            Txt(FB_Code).TEXT = ""
            Txt(VehAmt) = ""
            Txt(TaxPer) = ""
            Txt(TaxAmt) = ""
            Txt(TaxSurPer) = ""
            Txt(TaxSurch) = ""
            Txt(TOTPer) = ""
            Txt(TOTAmt) = ""
        Else
            Txt(Index).TEXT = RSBook!Code
            Txt(Index).Tag = RSBook!OrdDocId
            FillRecords RSBook
        End If
    Case Party
        If IsValid(Txt(Index), "Party Name") = False Then Cancel = True: Exit Sub
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsParty!Name
            Txt(Index).Tag = RsParty!Code
            Txt(PanNo) = IIf(IsNull(RsParty!PanNo), "", RsParty!PanNo)
            Txt(CircleNo) = IIf(IsNull(RsParty!ITWARD_NO), "", RsParty!ITWARD_NO)
            LblPartyBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
            LblPartyBal = LblPartyBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
            Txt(FormType).Enabled = IIf(IsNull(RsParty!PanNo) Or RsParty!PanNo = "", True, False)
            Txt(FormType).TEXT = IIf(Txt(FormType).Enabled = True, "N/A", "")
        End If
    Case AcHead
        If IsValid(Txt(Index), "A/C Head") = False Then Cancel = True: Exit Sub
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsParty!Name
            Txt(Index).Tag = RsParty!Code
            lblAcBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
            lblAcBal = lblAcBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
        End If
    Case VDate
        If Len(Trim(Txt(VDate).TEXT)) = 0 Then
            Txt(VDate).TEXT = PubLoginDate
        Else
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
        Cancel = Not CheckFinYear(Txt(VDate))
       If Cancel = False And TopCtrl1.TopText2 = "Add" Then Txt(VType).SetFocus
    Case ProDate
        Txt(Index).TEXT = RetDate(Txt(Index))
        If Txt(ProDate) <> "" Then
            If CDate(Txt(ProDate)) > CDate(Txt(VDate)) Then
                MsgBox "Provisional Date  > Vr Date", vbInformation, "Validation"
                Cancel = True
                Txt(ProDate).SetFocus
            End If
        End If
    Case BookDate, DDDate
        Txt(Index).TEXT = RetDate(Txt(Index))
    Case Amt
        Txt(Index).TEXT = Format(Txt(Index).TEXT, "0.00")
    Case DiscAmt
        If Val(Txt(DiscAmt)) > Val(Txt(Amt)) Then
            MsgBox "Please give proper discount."
            Txt(DiscAmt) = "": Txt(DiscAmt).SetFocus
            Cancel = True
            Exit Sub
        End If
        Txt(Index).TEXT = Format(Txt(Index).TEXT, "0.00")
    Case CreditCardNo
        If Txt(CreditCardNo) <> "" Then
            Set RsTemp = GCn.Execute("Select A.CreditCardAc, S.Name From AcControls A Left Join SubGroup S On A.CreditCardAc=S.SubCode")
            If RsTemp.RecordCount > 0 Then
                Txt(AcHead).Tag = XNull(RsTemp!CreditCardAc)
                Txt(AcHead) = XNull(RsTemp!Name)
                Txt(AcHead).Enabled = False
                
            End If
        Else
            Txt(AcHead).Enabled = True
        End If
End Select
Ctrl_validate Txt(Index)
Set Rst = Nothing
End Sub

'*** Fuctions ********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
    Txt(I).Tag = ""
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
    Txt(TxtDocID).TEXT = Master!DocID
    LblDiv.CAPTION = "Division : " & left(Master!DocID, 1)
    LblSite.CAPTION = "Site Code : " & mID(Master!Site_Code, 1, 1)
    Txt(SiteCode).Tag = mID(Master!Site_Code, 2, 1)
    Txt(SiteCode).TEXT = GCn.Execute("select site_desc from site where site_code = '" & Txt(SiteCode).Tag & "'").Fields(0).Value
    LblUser = IIf(Not IsNull(Master!AddDate), "Add By : " & XNull(Master!AddBy) & "  Dated : " & XNull(Master!AddDate), "") & IIf(Not IsNull(Master!ModifyDate), "     Modify By : " & XNull(Master!ModifyBy) & "  Dated : " & XNull(Master!ModifyDate), "")
    LblVPrefix.CAPTION = mID(Master!DocID, 8, 5)
    Txt(SerialNo).TEXT = Master!V_NO
    Txt(VDate).TEXT = Master!V_DATE
    Txt(RectCatg).TEXT = IIf(IsNull(Master!RectCatg), "", Master!RectCatg)
    mVType = Master!V_Type
    Txt(VType).Tag = mVType
    Txt(VType).TEXT = G_FaCn.Execute("select Description from Voucher_Type where category='GENFA' and v_type = '" & Txt(VType).Tag & "'").Fields(0).Value
    '*** A/c Posting Status
    Txt(AcPostByName) = IIf(IsNull(Master!AcPostByU_Name), "", Master!AcPostByU_Name)
    Txt(AcPostDate) = IIf(IsNull(Master!AcPostByU_EntDt), "", Master!AcPostByU_EntDt)
    '***
    Txt(BookNo).Tag = Master!Ord_DocId
    If Txt(BookNo).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "SELECT VO.Inv_Date,VO.OrdDocId,VO.Ord_no,VO.Ord_Date, VO.MODEL, VO.FB_CODE FROM veh_order as VO where OrdDocId = '" & Txt(BookNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount > 0 Then
            Txt(BookNo).TEXT = Rst!Ord_No
        End If
        FillRecords Rst
    Else
        Txt(BookNo).TEXT = ""
        Txt(BookDate).TEXT = ""
        BookDocId = ""
        Txt(Model).TEXT = ""
        Txt(FB_Code).Tag = ""
        Txt(FB_Code).TEXT = ""
    End If
    Txt(Party).Tag = Master!PartyCode
    If Txt(Party).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NAME,Curr_Bal,PanNo,ITWARD_NO from SubGroup where Subcode = '" & Txt(Party).Tag & "'", GCn, adOpenDynamic, adLockBatchOptimistic
        Txt(Party) = Rst!Name
        Txt(PanNo) = IIf(IsNull(Rst!PanNo), "", Rst!PanNo)
        Txt(CircleNo) = IIf(IsNull(Rst!ITWARD_NO), "", Rst!ITWARD_NO)
        LblPartyBal = "Bal. " & Format(Abs(Rst!Curr_Bal), "0.00")
        LblPartyBal = LblPartyBal & IIf(Rst!Curr_Bal > 0, " Cr", IIf(Rst!Curr_Bal < 0, " Dr", ""))
        Txt(FormType).Enabled = IIf(IsNull(Rst!PanNo) Or Rst!PanNo = "", True, False)
        If TopCtrl1.TopText2 = "Browse" Then Txt(FormType).Enabled = False
        Txt(FormType).TEXT = IIf(Txt(FormType).Enabled = True, "N/A", "")
    Else
        Txt(Party).TEXT = ""
    End If
    Txt(ProNo).TEXT = IIf(IsNull(Master!Prov_No) Or Master!Prov_No = 0, "", Master!Prov_No)
    Txt(ProDate).TEXT = IIf(IsNull(Master!Prov_Date), "", Master!Prov_Date)
    Txt(RecLocation).Tag = IIf(IsNull(Master!Prov_Location), "", Master!Prov_Location)
    If Txt(RecLocation).Tag <> "" Then
        Txt(RecLocation).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(RecLocation).Tag & "'").Fields(0).Value
    End If
    Txt(DDNo).TEXT = IIf(IsNull(Master!DDNo), "", Master!DDNo)
    Txt(DDDate).TEXT = IIf(IsNull(Master!DDDate), "", Master!DDDate)
    Txt(VehTrnYN).TEXT = IIf(Master!Vehicle_YN = 1, "Yes", "No")
    Txt(PrnNameYN).TEXT = IIf(Master!PrintParty_YN = 1, "Yes", "No")
    Txt(Narr).TEXT = IIf(IsNull(Master!Narration), "", Master!Narration)
    Txt(Rem1).TEXT = IIf(IsNull(Master!RWTF1), "", Master!RWTF1)
    Txt(Rem2).TEXT = IIf(IsNull(Master!RWTF2), "", Master!RWTF2)
    Txt(FormType) = IIf(IsNull(Master!IForm), "", Master!IForm)
    Txt(CreditCardNo) = XNull(Master!CreditCardNo)
    If mVType = mCustCRN Then
        Txt(VehAmt) = IIf(IsNull(Master!Veh_Amt), "", Format(Master!Veh_Amt, "0.00"))
        Txt(TaxAmt) = IIf(IsNull(Master!Tax_Amt), "", Format(Master!Tax_Amt, "0.00"))
        Txt(TaxSurch) = IIf(IsNull(Master!Surcharge_Amt), "", Format(Master!Surcharge_Amt, "0.00"))
        Txt(TOTAmt) = IIf(IsNull(Master!Tot_Amt), "", Format(Master!Tot_Amt, "0.00"))
    Else
        Txt(VehAmt) = ""
        Txt(TaxAmt) = ""
        Txt(TaxSurch) = ""
        Txt(TOTAmt) = ""
    End If
    Txt(Amt).TEXT = Format(Master!Amount, "0.00")
    
    Txt(AcHead).Tag = Master!AcCode
    If Txt(AcHead).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NAME,Curr_Bal from SubGroup where Subcode = '" & Txt(AcHead).Tag & "'", GCn, adOpenDynamic, adLockBatchOptimistic
        Txt(AcHead) = XNull(Rst!Name)
        lblAcBal = "Bal. " & Format(Abs(Rst!Curr_Bal), "0.00")
        lblAcBal = lblAcBal & IIf(Rst!Curr_Bal > 0, " Cr", IIf(Rst!Curr_Bal < 0, " Dr", ""))
    Else
        Txt(AcHead).TEXT = ""
    End If
    Txt(DiscAcName).Tag = XNull(Master!DiscAc)
    If Txt(DiscAcName).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NAME from SubGroup where Subcode = '" & Txt(DiscAcName).Tag & "'", GCn, adOpenDynamic, adLockBatchOptimistic
        Txt(DiscAcName) = Rst!Name
    Else
        Txt(DiscAcName).TEXT = ""
    End If
    Txt(DiscAmt) = Format(VNull(Master!DiscAmt), "0.00")
    
    
    
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
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    Txt(I).ForeColor = CtrlFColOrg
Next


If UCase(left(PubComp_Name, 4)) = "ENAR" Then Txt(SiteCode).Enabled = False

If TopCtrl1.TopText2 = "Edit" Then
    Txt(SiteCode).Enabled = False
    Txt(VDate).Enabled = False
    Txt(SerialNo).Enabled = False
    Txt(VType).Enabled = False
    If mVType <> mCustCRN Then
        Txt(VehAmt).Enabled = False
        Txt(TaxAmt).Enabled = False
        Txt(TaxSurch).Enabled = False
        Txt(TOTAmt).Enabled = False
    End If
End If
If RSOJPR = True Then
    Label3(10).Visible = True
    Label3(11).Visible = True
    Txt(DiscAcName).Visible = True
    Txt(DiscAmt).Visible = True
Else
    Label3(10).Visible = False
    Label3(11).Visible = False
    Txt(DiscAcName).Visible = False
    Txt(DiscAmt).Visible = False
End If
Txt(Rem1).Enabled = False
Txt(Rem2).Enabled = False
Txt(TxtDocID).Enabled = False
Txt(BookNo).Enabled = False
Txt(BookDate).Enabled = False
Txt(Model).Enabled = False
Txt(FB_Code).Enabled = False
Txt(Model).Enabled = False
Txt(FB_Code).Enabled = False
Txt(PanNo).Enabled = False
Txt(CircleNo).Enabled = False
txtDisabled_Color Me

If Txt(CreditCardNo) <> "" Then
    Txt(AcHead).Enabled = False
End If

End Sub

Private Sub Grid_Hide()
    If DGBook.Visible = True Then DGBook.Visible = False
    If DGSite.Visible = True Then DGSite.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGCity.Visible = True Then DGCity.Visible = False
    If DGVType.Visible = True Then DGVType.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
End Sub
 
Private Sub FillRecords(RSBook As ADODB.Recordset)
Dim Rst As ADODB.Recordset
    If RSBook.RecordCount > 0 Then
        Txt(BookDate).TEXT = RSBook!Ord_Date
        BookDocId = RSBook!OrdDocId
        Txt(Model).TEXT = RSBook!Model
        Txt(FB_Code).Tag = IIf(IsNull(RSBook!FB_Code), "", RSBook!FB_Code)
        If Txt(FB_Code).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select fincode as code,finname as name,AcCode from ContractFinance where fincatg = 0 and  fincode = '" & Txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
            Txt(FB_Code).TEXT = Rst!Name
        Else
            Txt(FB_Code).TEXT = ""
        End If
        If Not IsNull(RSBook!Inv_Date) Then
            If mVType = mCustCRN Then
                 GSQL = "SELECT V.Tax_Per,V.Tax_Amt,V.Surcharge_Per,V.Surcharge_Amt,V.TOT_Per,V.TOT_Amt," & _
                            "T.Tax_Ac_Code,T.Sur_Ac_Code,T.PurSal_Ac_Code " & _
                            "from Veh_Order as V left Join TaxFormsAc as T on V.Form_Code&'" & PubDivCode & "'=T.Form_Code&T.Div_Code " & _
                            "where V.OrdDocId='" & Txt(BookNo).Tag & "'"
                
                Set Rst = New Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
                Txt(TaxPer) = IIf(IsNull(Rst!Tax_Per), "", Rst!Tax_Per)
                Txt(TaxSurPer) = IIf(IsNull(Rst!surcharge_per), "", Rst!surcharge_per)
                Txt(TOTPer) = IIf(IsNull(Rst!TOT_Per), "", Rst!TOT_Per)
                
                'mHead = Rst!PurSal_Ac_Code 'Veh Amount
                 mTaxAcHead = Rst!Tax_Ac_Code
                 mTaxSurAcHead = Rst!Sur_Ac_Code
                 mTOTAcHead = G_FaCn.Execute("Select TOTax_Ac From AcControls").Fields(0)
             End If
         End If
        Set Rst = Nothing
    Else
        Txt(BookDate).TEXT = ""
        BookDocId = ""
        Txt(Model).TEXT = ""
        Txt(FB_Code).Tag = ""
        Txt(FB_Code).TEXT = ""
        Txt(VehAmt) = ""
        Txt(TaxPer) = ""
        Txt(TaxAmt) = ""
        Txt(TaxSurPer) = ""
        Txt(TaxSurch) = ""
        Txt(TOTPer) = ""
        Txt(TOTAmt) = ""
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
                GCn.Execute "update Rect set Printed = 1 where DocID='" & Txt(TxtDocID) & "'"
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
    
    UserPrintStr = "" & "By: " & RstCust!U_Name & "   " & RstCust!U_EntDt
    Print #1, "" & Space((PageWidth * 1.7) - Len(UserPrintStr)) & "By: " & RstCust!U_Name & "   " & RstCust!U_EntDt & mChr18
    Print #1, Chr(27) + Chr(67) + Chr(72)
    Print #1, ""
    Print #1, ""
    Print #1, ""
        
    If UCase(left(PubComp_Name, 5)) = "SOCIE" And Txt(VType).Tag = "G_ABR" Then
        For j = 1 To 4
            Print #1, ""
        Next
        Print #1, mChr14 & mEmph & Txt(AcHead).TEXT & mEmph1 & mChr18
        AcAdd = GCn.Execute("Select Add1 & Add2 from SubGroup where SubCode='" & Txt(AcHead).Tag & "'").Fields(0)
       'AcCity = GCn.Execute("Select isNull(City.cITYName,'') from SubGroup Left Join City on City.CityCode=SubGroup.CityCode where SubGroup.SubCode='" & Txt(AcHead).Tag & "'").Fields(0)
        Print #1, SETW(AcAdd & AcCity, 32) & Space(20) & Chr(27) & Chr(69) & "CHEQUE DEPOSIT SLIP" & Chr(27) & Chr(70)
        Print #1, "Account Title:" & SETW(PubComp_Name, 35) & Space(5) & "Date:" & SETW(Txt(VDate), 12)
        Print #1, Space(54) & "Bank Rec.No." & SETW(Txt(SerialNo), 12)
        Print #1, "Rs:" & ntow(Val(Txt(Amt)), "Rupees", "Paise")
        Print #1, mChr14
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, " Customer            Chq/DD No.   Bank/Branch         Amount Realisation Remark "
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, SETW(Txt(Party), 20) & Space(1) & SETW(Txt(DDNo), 12) & Space(1) & SETW(Txt(Narr), 15) & Space(1) & SETN(Format(Txt(Amt), "0.00"), 10)
        For j = 1 To 5
            Print #1, ""
        Next
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, Space(43) & "Total: " & SETN(Format(Txt(Amt), "0.00"), 10)
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

