VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmQuot 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Quotation"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13755
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
   ScaleHeight     =   9180
   ScaleWidth      =   13755
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
      Height          =   1605
      Left            =   810
      TabIndex        =   125
      Top             =   2805
      Visible         =   0   'False
      Width           =   5025
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
         TabIndex        =   135
         Top             =   720
         Width           =   750
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
         TabIndex        =   134
         Top             =   720
         Width           =   1200
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
         TabIndex        =   133
         Top             =   300
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   132
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
         Index           =   2
         Left            =   7425
         TabIndex        =   131
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmQuot.frx":0000
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
         TabIndex        =   130
         ToolTipText     =   "Printer "
         Top             =   285
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmQuot.frx":030A
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
         TabIndex        =   129
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmQuot.frx":0614
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
         TabIndex        =   128
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
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
         Picture         =   "frmQuot.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Screen"
         Top             =   1275
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
         Left            =   4695
         MousePointer    =   99  'Custom
         Picture         =   "frmQuot.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Delete Current Record"
         Top             =   0
         Width           =   315
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
         Left            =   -15
         TabIndex        =   138
         Top             =   -15
         Width           =   4695
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
         TabIndex        =   137
         Top             =   1275
         Width           =   4650
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
         TabIndex        =   136
         Top             =   300
         Width           =   3315
      End
      Begin VB.Line Line6 
         X1              =   2820
         X2              =   345
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   360
         Y1              =   615
         Y2              =   720
      End
      Begin VB.Line Line7 
         X1              =   2820
         X2              =   2820
         Y1              =   630
         Y2              =   735
      End
      Begin VB.Line Line8 
         X1              =   1470
         X2              =   1470
         Y1              =   510
         Y2              =   600
      End
   End
   Begin MSDataGridLib.DataGrid DgModelGroup 
      Height          =   2835
      Left            =   1380
      Negotiate       =   -1  'True
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   8325
      Visible         =   0   'False
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   5001
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Model Group"
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
            ColumnWidth     =   3209.953
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   3705
      Left            =   60
      Negotiate       =   -1  'True
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   5850
      Visible         =   0   'False
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   6535
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
         DataField       =   "Sale_Rate"
         Caption         =   "Rate"
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
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6075.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1379.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPurpose 
      Height          =   4935
      Left            =   3750
      Negotiate       =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   5130
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
      Caption         =   "Purpose Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Purpose"
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
      Index           =   40
      Left            =   1830
      TabIndex        =   25
      Text            =   "0123456789012345678901234"
      Top             =   2655
      Width           =   4260
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
      Index           =   41
      Left            =   1830
      MaxLength       =   20
      TabIndex        =   26
      Top             =   2925
      Width           =   4260
   End
   Begin MSDataGridLib.DataGrid DGADItem 
      Height          =   4935
      Left            =   8880
      Negotiate       =   -1  'True
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   6810
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
      Caption         =   "Add/Del Item Help"
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
   Begin VB.TextBox TxtGrid1 
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
      Left            =   945
      TabIndex        =   35
      Top             =   5655
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSDataGridLib.DataGrid DGFin 
      Height          =   4935
      Left            =   8310
      Negotiate       =   -1  'True
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   6960
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
      Caption         =   "Financier Help "
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
   Begin MSDataGridLib.DataGrid DGProf 
      Height          =   4935
      Left            =   11490
      Negotiate       =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   7215
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
      Caption         =   "Profession Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Profession"
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
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3540
      Left            =   2205
      Negotiate       =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   6244
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
      Caption         =   "Prospective Customers Help"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "name"
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
         DataField       =   "NSuffix"
         Caption         =   "Sfix"
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
         Caption         =   "Address-1"
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
         Caption         =   "Address-2"
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
            ColumnWidth     =   3704.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2475.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2865.26
         EndProperty
         BeginProperty Column04 
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
      Left            =   3180
      MaxLength       =   15
      TabIndex        =   28
      Text            =   "Christian"
      Top             =   3195
      Width           =   945
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
      Left            =   7515
      MaxLength       =   40
      TabIndex        =   12
      Top             =   2115
      Width           =   4275
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
      Left            =   7515
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1845
      Width           =   4275
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
      Left            =   7515
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1575
      Width           =   4275
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
      Left            =   7515
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3465
      Width           =   4275
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
      Index           =   34
      Left            =   7515
      MaxLength       =   6
      TabIndex        =   15
      Text            =   "123456"
      Top             =   2385
      Width           =   1020
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
      Index           =   16
      Left            =   7515
      MaxLength       =   35
      TabIndex        =   17
      Top             =   2925
      Width           =   4275
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
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   14
      Text            =   "123456"
      Top             =   1305
      Width           =   810
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
      Left            =   7515
      MaxLength       =   35
      TabIndex        =   16
      Top             =   2655
      Width           =   4275
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
      Left            =   7515
      MaxLength       =   24
      TabIndex        =   20
      Top             =   3735
      Width           =   4275
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
      Left            =   7515
      MaxLength       =   24
      TabIndex        =   18
      Top             =   3195
      Width           =   4275
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
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1035
      Width           =   4260
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
      Left            =   1830
      MaxLength       =   4
      TabIndex        =   7
      Top             =   765
      Width           =   630
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
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   8
      Top             =   765
      Width           =   4275
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
      Left            =   6765
      MaxLength       =   1
      TabIndex        =   6
      Top             =   495
      Width           =   255
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
      Left            =   1830
      MaxLength       =   4
      TabIndex        =   4
      Top             =   495
      Width           =   630
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
      Index           =   13
      Left            =   1830
      MaxLength       =   25
      TabIndex        =   13
      Text            =   "0123456789012345678901234"
      Top             =   1305
      Width           =   3000
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
      Index           =   27
      Left            =   1830
      MaxLength       =   6
      TabIndex        =   31
      Top             =   3735
      Width           =   1095
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
      Index           =   2
      Left            =   9915
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1260
      Width           =   975
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   39
      Left            =   9255
      MaxLength       =   12
      TabIndex        =   42
      Top             =   5610
      Width           =   1455
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
      Left            =   5670
      MaxLength       =   4
      TabIndex        =   29
      Text            =   "yes"
      Top             =   3195
      Width           =   420
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
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   21
      Top             =   1575
      Width           =   4260
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7530
      TabIndex        =   78
      Top             =   7005
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   -90
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   60
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
      Align           =   3  'Align Left
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
      Index           =   4
      Left            =   2475
      MaxLength       =   40
      TabIndex        =   5
      Top             =   495
      Width           =   4275
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
      Index           =   38
      Left            =   9255
      MaxLength       =   12
      TabIndex        =   41
      Top             =   5340
      Width           =   1455
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
      Left            =   9210
      MaxLength       =   21
      TabIndex        =   1
      Top             =   450
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
      Index           =   37
      Left            =   7935
      MaxLength       =   21
      TabIndex        =   46
      Text            =   "visfalse"
      Top             =   7185
      Visible         =   0   'False
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
      Index           =   36
      Left            =   7935
      MaxLength       =   12
      TabIndex        =   45
      Text            =   "visfalse"
      Top             =   6915
      Visible         =   0   'False
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
      Index           =   30
      Left            =   1095
      MaxLength       =   20
      TabIndex        =   38
      Top             =   6390
      Width           =   4440
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
      Index           =   33
      Left            =   9255
      MaxLength       =   15
      TabIndex        =   43
      Top             =   5880
      Width           =   1455
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
      Index           =   32
      Left            =   1095
      MaxLength       =   50
      TabIndex        =   40
      Top             =   6930
      Width           =   4440
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
      Index           =   31
      Left            =   1095
      MaxLength       =   50
      TabIndex        =   39
      Top             =   6660
      Width           =   4440
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
      Index           =   35
      Left            =   9255
      MaxLength       =   4
      TabIndex        =   44
      Top             =   6150
      Visible         =   0   'False
      Width           =   675
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   9255
      MaxLength       =   10
      TabIndex        =   37
      Top             =   5070
      Width           =   1455
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
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   24
      Top             =   2385
      Width           =   4260
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
      Height          =   255
      Index           =   26
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   30
      Top             =   3465
      Width           =   4260
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
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   22
      Top             =   1845
      Width           =   4260
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
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   23
      Top             =   2115
      Width           =   4260
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
      Index           =   28
      Left            =   4725
      MaxLength       =   4
      TabIndex        =   32
      Top             =   3735
      Width           =   420
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
      Left            =   1395
      MaxLength       =   4
      TabIndex        =   27
      Top             =   3195
      Width           =   420
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
      Left            =   9240
      MaxLength       =   12
      TabIndex        =   2
      Top             =   990
      Width           =   1470
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1005
      Left            =   30
      TabIndex        =   34
      Top             =   4035
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1773
      _Version        =   393216
      BackColor       =   12243913
      Cols            =   12
      BackColorFixed  =   13300221
      ForeColorFixed  =   16384
      BackColorSel    =   12243913
      ForeColorSel    =   -2147483640
      BackColorBkg    =   13623520
      GridColor       =   13298928
      GridColorFixed  =   49344
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"frmQuot.frx":0F96
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
      _Band(0).Cols   =   12
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   4935
      Left            =   7575
      Negotiate       =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   6825
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
      Caption         =   "City Help"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGArea 
      Height          =   4935
      Left            =   3450
      Negotiate       =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   7170
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
      Caption         =   "Area Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Area Name"
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
   Begin MSDataGridLib.DataGrid DGRep 
      Height          =   4935
      Left            =   11460
      Negotiate       =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   7095
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
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Representative Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Representative Name"
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
   Begin MSDataGridLib.DataGrid DGRef 
      Height          =   4935
      Left            =   7560
      Negotiate       =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   6885
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
      Caption         =   "Refer Person Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Reference Person"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1200
      Left            =   45
      TabIndex        =   36
      Top             =   5085
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2117
      _Version        =   393216
      BackColor       =   12243913
      Cols            =   7
      BackColorFixed  =   13300221
      ForeColorFixed  =   16384
      BackColorSel    =   12243913
      ForeColorSel    =   -2147483640
      BackColorBkg    =   13623520
      GridColor       =   13298928
      GridColorFixed  =   49344
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "SrNo.|Additional Fitments |Type      |Qty|Rate  |Amount|Itemcode"
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
      _Band(0).Cols   =   7
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
      Left            =   6150
      TabIndex        =   33
      Top             =   1965
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   15
      Left            =   270
      TabIndex        =   142
      Top             =   2670
      Width           =   720
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   18
      Left            =   1710
      TabIndex        =   141
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Intended Use"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   14
      Left            =   270
      TabIndex        =   140
      Top             =   2940
      Width           =   1110
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   14
      Left            =   1710
      TabIndex        =   139
      Top             =   2910
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   33
      Left            =   3060
      TabIndex        =   121
      Top             =   3195
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   45
      Left            =   2295
      TabIndex        =   120
      Top             =   3195
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   30
      Left            =   7380
      TabIndex        =   119
      Top             =   1560
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   29
      Left            =   7380
      TabIndex        =   118
      Top             =   3450
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   28
      Left            =   7380
      TabIndex        =   117
      Top             =   2910
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   27
      Left            =   7380
      TabIndex        =   116
      Top             =   3180
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   26
      Left            =   7380
      TabIndex        =   115
      Top             =   3720
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   4
      Left            =   7380
      TabIndex        =   114
      Top             =   2640
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
      Height          =   285
      Index           =   3
      Left            =   7290
      TabIndex        =   113
      Top             =   4425
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   2
      Left            =   7380
      TabIndex        =   112
      Top             =   2370
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   1
      Left            =   1710
      TabIndex        =   111
      Top             =   1020
      Width           =   45
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   13
      Left            =   6300
      TabIndex        =   110
      Top             =   1590
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   12
      Left            =   6300
      TabIndex        =   109
      Top             =   3480
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STD Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   11
      Left            =   6300
      TabIndex        =   108
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Res)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   10
      Left            =   6765
      TabIndex        =   107
      Top             =   2940
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   9
      Left            =   4860
      TabIndex        =   106
      Top             =   1320
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   8
      Left            =   6300
      TabIndex        =   105
      Top             =   3210
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   7
      Left            =   6300
      TabIndex        =   104
      Top             =   3750
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (Off)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   37
      Left            =   6300
      TabIndex        =   103
      Top             =   2670
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   6
      Left            =   270
      TabIndex        =   102
      Top             =   1050
      Width           =   1320
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   0
      Left            =   1710
      TabIndex        =   101
      Top             =   750
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   5
      Left            =   270
      TabIndex        =   100
      Top             =   780
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   4
      Left            =   270
      TabIndex        =   94
      Top             =   1320
      Width           =   300
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   5
      Left            =   1710
      TabIndex        =   93
      Top             =   1290
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Call Status*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   24
      Left            =   270
      TabIndex        =   92
      Top             =   3750
      Width           =   975
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   6
      Left            =   1710
      TabIndex        =   91
      Top             =   3720
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008080&
      Height          =   1140
      Left            =   7500
      Top             =   420
      Width           =   4005
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
      Left            =   9210
      TabIndex        =   90
      Top             =   1245
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   1
      Left            =   7920
      TabIndex        =   89
      Top             =   1245
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   92
      Left            =   9075
      TabIndex        =   88
      Top             =   1275
      Width           =   45
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division           :"
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
      Left            =   7920
      TabIndex        =   87
      Top             =   735
      Width           =   1200
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    :"
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
      Left            =   9795
      TabIndex        =   86
      Top             =   735
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quotation Expiry Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   31
      Left            =   7365
      TabIndex        =   85
      Top             =   5625
      Width           =   1755
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   12
      Left            =   9135
      TabIndex        =   84
      Top             =   5370
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   25
      Left            =   5550
      TabIndex        =   83
      Top             =   3195
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finance Y/N*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   44
      Left            =   4455
      TabIndex        =   82
      Top             =   3195
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   43
      Left            =   270
      TabIndex        =   81
      Top             =   1590
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   24
      Left            =   1710
      TabIndex        =   80
      Top             =   1560
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
      Height          =   225
      Index           =   23
      Left            =   9075
      TabIndex        =   76
      Top             =   465
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   42
      Left            =   7920
      TabIndex        =   75
      Top             =   465
      Width           =   630
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   20
      Left            =   975
      TabIndex        =   74
      Top             =   6375
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Follow Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   39
      Left            =   45
      TabIndex        =   73
      Top             =   6405
      Width           =   825
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   19
      Left            =   9135
      TabIndex        =   72
      Top             =   5865
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month Action Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   38
      Left            =   7365
      TabIndex        =   71
      Top             =   5895
      Width           =   1455
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   17
      Left            =   975
      TabIndex        =   70
      Top             =   6645
      Width           =   45
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   36
      Left            =   45
      TabIndex        =   69
      Top             =   6675
      Width           =   765
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   16
      Left            =   9135
      TabIndex        =   68
      Top             =   6135
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RoundOff YN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   35
      Left            =   7365
      TabIndex        =   67
      Top             =   6165
      Visible         =   0   'False
      Width           =   1065
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   15
      Left            =   9135
      TabIndex        =   66
      Top             =   5055
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Rs."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   34
      Left            =   7365
      TabIndex        =   65
      Top             =   5085
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profession"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   32
      Left            =   270
      TabIndex        =   64
      Top             =   2400
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   13
      Left            =   1710
      TabIndex        =   63
      Top             =   2370
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   30
      Left            =   270
      TabIndex        =   62
      Top             =   3480
      Width           =   765
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   11
      Left            =   1710
      TabIndex        =   61
      Top             =   3450
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   10
      Left            =   1710
      TabIndex        =   60
      Top             =   1830
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Executive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   29
      Left            =   270
      TabIndex        =   59
      Top             =   1860
      Width           =   1350
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   9
      Left            =   1710
      TabIndex        =   58
      Top             =   2100
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reffered By"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   28
      Left            =   270
      TabIndex        =   57
      Top             =   2130
      Width           =   945
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   8
      Left            =   4605
      TabIndex        =   56
      Top             =   3720
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Vehicle YN*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   27
      Left            =   3165
      TabIndex        =   55
      Top             =   3750
      Width           =   1365
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   7
      Left            =   1275
      TabIndex        =   54
      Top             =   3195
      Width           =   45
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   26
      Left            =   270
      TabIndex        =   53
      Top             =   3195
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   3
      Left            =   270
      TabIndex        =   52
      Top             =   510
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   90
      Left            =   1710
      TabIndex        =   51
      Top             =   480
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
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   91
      Left            =   9075
      TabIndex        =   50
      Top             =   975
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   93
      Left            =   9135
      TabIndex        =   49
      Top             =   5625
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expected Delv.  Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   0
      Left            =   7365
      TabIndex        =   48
      Top             =   5370
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   2
      Left            =   7920
      TabIndex        =   47
      Top             =   1005
      Width           =   390
   End
End
Attribute VB_Name = "frmQuot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BackColorSelLeave As String

Dim RsParty As ADODB.Recordset
Dim rsFin As ADODB.Recordset
Dim RsMod  As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsRef As ADODB.Recordset
Dim RsRep As ADODB.Recordset
Dim RsProf As ADODB.Recordset
Dim RsPurpose  As ADODB.Recordset
Dim RsModelGroup As ADODB.Recordset
Dim RsArea As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsADItem As ADODB.Recordset
Dim CustFlag As Boolean
Dim GridKey As Integer
Dim DocID As String * 21
Dim mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String

Private Const TxtDocID As Byte = 0
Private Const VDate As Byte = 1
Private Const SerialNo As Byte = 2
Private Const NPrefix As Byte = 3
Private Const Party As Byte = 4
Private Const NSuffix As Byte = 5
Private Const FPrefix As Byte = 6
Private Const fname As Byte = 7
Private Const Profession As Byte = 8
Private Const ConPerson As Byte = 9
Private Const Add1 As Byte = 10
Private Const Add2 As Byte = 11
Private Const Add3 As Byte = 12
Private Const City As Byte = 13
Private Const Pin As Byte = 14
Private Const STD As Byte = 34
Private Const PhoneOff As Byte = 15
Private Const PhoneResi As Byte = 16
Private Const Mobile  As Byte = 17
Private Const EMail As Byte = 18
Private Const FAx As Byte = 19
Private Const Area As Byte = 20
Private Const REF_CODE As Byte = 21
Private Const REP_CODE As Byte = 22
Private Const Govt_YN As Byte = 23
Private Const Religion As Byte = 24
Private Const FIN_YN As Byte = 25
Private Const FB_Code As Byte = 26
Private Const Call_Status As Byte = 27
Private Const FirstVeh_YN As Byte = 28
Private Const Amount As Byte = 29
Private Const FOLLOW_UP As Byte = 30
Private Const NARR1 As Byte = 31
Private Const NARR2 As Byte = 32
Private Const MAP As Byte = 33
Private Const RoundOff_YN As Byte = 35
Private Const Book_SiteCode  As Byte = 36
Private Const Book_DocId As Byte = 37
Private Const DEL_DATE As Byte = 38
Private Const EXP_DATE As Byte = 39
Private Const Purpose As Byte = 40
Private Const IndUse As Byte = 41

' Col Declaration

Private Const ModelGroup As Byte = 1
Private Const Model As Byte = 2
Private Const Taxable As Byte = 3
Private Const RW As Byte = 4
Private Const Qty As Byte = 5
Private Const Rate As Byte = 6
Private Const Amt  As Byte = 7
Private Const TaxPer As Byte = 8
Private Const TaxAmt As Byte = 9
Private Const SurchPer  As Byte = 10
Private Const SurchAmt  As Byte = 11
Private Const Namt  As Byte = 12
Private Const BookId As Byte = 13
Private Const ModelGroupCode As Byte = 14

Private Const ADItem  As Byte = 1
Private Const ADType  As Byte = 2
Private Const Qty1  As Byte = 3
Private Const Rate1 As Byte = 4
Private Const Amt1 As Byte = 5
Private Const ADItemCode  As Byte = 6

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub DGADItem_Click()
    DGADItem.Visible = False
    If RsADItem.RecordCount > 0 Then
        txtgrid1(0).TEXT = RsADItem!Name
         FGrid1.TextMatrix(FGrid1.Row, ADItem) = RsADItem!Name
         FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = RsADItem!Code
    End If
   txtgrid1(0).SetFocus
End Sub

Private Sub DGPurpose_Click()
If RsPurpose.RecordCount > 0 Then
    Txt(Purpose).TEXT = RsPurpose!Name
    Txt(Purpose).Tag = RsPurpose!Code
End If
Txt(Purpose).SetFocus
DGPurpose.Visible = False
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
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
    Dim sitecond As String
    sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and " & cMID("DocId", "3", "1") & " ='" & PubSiteCode & "'"
    End If
    
    If PubMoveRecYn Then
        Master.Open "select DocID as searchcode,docid from Veh_Quot where left(docid,1)='" & PubDivCode & "' " & sitecond & " order by V_date desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 DocID as searchcode,docid from Veh_Quot where left(docid,1)='" & PubDivCode & "' " & sitecond & " order by V_date desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "select citycode as code,cityname as name from city order by cityname,citycode", GCn, adOpenDynamic, adLockOptimistic
    Set DGCity.DataSource = RsCity
    
    Set RsRef = New ADODB.Recordset
    RsRef.CursorLocation = adUseClient
    RsRef.Open "select RefCode as code,RefName as name from reffered order by Refname", GCn, adOpenDynamic, adLockOptimistic
    Set DGRef.DataSource = RsRef
    
    Set rsFin = New ADODB.Recordset
    rsFin.CursorLocation = adUseClient
'    rsFin.Open "select fincode as code,finname as name from ContractFinance where fincatg = 0  order by finname", GCn, adOpenDynamic, adLockOptimistic
    rsFin.Open "select fincode as code,finname + ',' + " & xIsNull("City.CityName", "") & " as name,AcCode,FinBankCode from ContractFinance " & _
    "left join city on left(ContractFinance.City,4)=City.CityCode where fincatg = 0  order by finname", GCn, adOpenDynamic, adLockOptimistic
    Set DGFin.DataSource = rsFin
  
    Set RsArea = New ADODB.Recordset
    RsArea.CursorLocation = adUseClient
    RsArea.Open "select AreaCode as code,AreaName as name from Area order by AreaName", GCn, adOpenDynamic, adLockOptimistic
    Set DGArea.DataSource = RsArea
    
    Set RsModelGroup = GCn.Execute("Select ModelGrp_Code As Code, ModelGrp_Name As Name From Model_Grp Order By ModelGrp_Name")
    Set DgModelGroup.DataSource = RsModelGroup
    
    
    Set RsProf = New ADODB.Recordset
    RsProf.CursorLocation = adUseClient
    RsProf.Open "select ProfessionCode as code,Professionname as name from Profession order by Professionname", GCn, adOpenDynamic, adLockOptimistic
    Set DGProf.DataSource = RsProf
  
    Set RsPurpose = New ADODB.Recordset
    RsPurpose.CursorLocation = adUseClient
    RsPurpose.Open "select PurposeCode as code,Purposename as name from Purpose order by PurposeName", GCn, adOpenDynamic, adLockOptimistic
    Set DGPurpose.DataSource = RsPurpose
  
    Set RsRep = New ADODB.Recordset
    RsRep.CursorLocation = adUseClient
    RsRep.Open "select Emp_code as code,emp_name as name from emp_mast where emp_type = 0  order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGRep.DataSource = RsRep

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select ProspectiveCust.Cust_code as code,ProspectiveCust.name  as name,ProspectiveCust.*,City.CityName from (ProspectiveCust Left Join City On City.CityCode=Prospectivecust.CityCode) order by ProspectiveCust.Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set RsMod = New ADODB.Recordset
    RsMod.CursorLocation = adUseClient
    RsMod.Open "select Model as code,Model_Desc as NAME,Chas_Type,Sale_Rate, Grp_Code From model Where (Div_Code='" & PubDivCode & "' or Div_Code='') Order by model", GCn, adOpenDynamic, adLockOptimistic
    Set DGMod.DataSource = RsMod
    
    Set RsADItem = New ADODB.Recordset
    With RsADItem
        .CursorLocation = adUseClient
        .Open "SELECT Prod_Code as code,Prod_name as name,Rate  FROM veh_amdModel order by  veh_amdModel.Prod_name ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGADItem.DataSource = RsADItem
    
    mVType = "V_QOT"
    Txt(VDate).Tag = PubLoginDate
    
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
Set RsParty = Nothing
Set rsFin = Nothing
Set RsMod = Nothing
Set RsCity = Nothing
Set RsProf = Nothing
Set RsPurpose = Nothing
Set RsRep = Nothing
Set RsRef = Nothing
Set RsArea = Nothing
Set Master = Nothing
Set mListItem = Nothing
End Sub

Private Sub ListView_Click()
If txtgrid1(0).Visible = False Then
    Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    Txt(Val(ListView.Tag)).SetFocus
Else
    txtgrid1(0).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txtgrid1(0).SetFocus
End If
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    CustFlag = False
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""
    Txt(TxtDocID).Enabled = False
    Txt(Govt_YN) = "No"
    Txt(Call_Status) = "Cold"
    Txt(VDate) = Txt(VDate).Tag
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    Txt(VDate).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer, mTrans As Boolean
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, BookId) <> "" Then
            MsgBox "Booking has been made against this Quotation", vbInformation, "deletion Denied": Exit Sub
        End If
    Next

If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then

    GCn.BeginTrans
    mTrans = True
    GCn.Execute ("delete from Veh_Quot where docId = '" & Master!DocID & "'")
    GCn.Execute ("delete from Veh_Quot1 where docId = '" & Master!DocID & "'")
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
       If mTrans Then GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    CustFlag = False
    FGrid.AddItem FGrid.Rows
    FGrid1.AddItem FGrid1.Rows
    FGrid.SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
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
    If TxtGrid(0).Visible = True Then TxtGrid(0).Visible = False
    If txtgrid1(0).Visible = True Then txtgrid1(0).Visible = False
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
    rsFin.Requery
    RsCity.Requery
    RsRep.Requery
    RsRef.Requery
    RsProf.Requery
    RsPurpose.Requery
    RsArea.Requery
    RsMod.Requery
    RsADItem.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    Dim DocIdHlp As String
    Dim Relg As Integer, StCall As Integer
    Dim Cnt As Integer
    On Error GoTo errlbl

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    If txtgrid1(0).Visible = True Then
        If TxtGridLeave1 = False Then
            txtgrid1(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(Txt(VDate), "Quotation Date") = False Then Exit Sub
    If IsValid(Txt(SerialNo), "Serial Number") = False Then Exit Sub
    If IsValid(Txt(Party), "Party Name") = False Then Exit Sub
    If Txt(Area) = "" Then
        MsgBox "Area Name is required", vbOKOnly, "Validation"
        Txt(Area).Enabled = True
        Txt(Area).SetFocus
        Exit Sub
    End If
    If Txt(REP_CODE) = "" Then
        MsgBox "Representative Person is required", vbOKOnly, "Validation"
        Txt(REP_CODE).Enabled = True
        Txt(REP_CODE).SetFocus
        Exit Sub
    End If
    If IsValid(Txt(REF_CODE), "Refer Person") = False Then Exit Sub
'    If IsValid(Txt(Profession), "Profession Name") = False Then Exit Sub
    If IsValid(Txt(Purpose), "Purpuse") = False Then Exit Sub
    
    If FGrid.Rows = 2 And FGrid.TextMatrix(1, Model) = "" Then MsgBox "Fill Transaction Data", vbInformation, "Required data": FGrid.Row = 1: FGrid.Col = Model: FGrid.SetFocus: Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Model) <> "" Then
            Cnt = 1
            If FGrid.TextMatrix(I, Taxable) = "" Then MsgBox "Fill Tax Yes/No in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Taxable: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, RW) = "" Then MsgBox "Fill RSO Yes/No in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Taxable: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Qty)) = 0 Then MsgBox "Fill Quantity in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Qty: FGrid.SetFocus:  Exit Sub
            If Val(FGrid.TextMatrix(I, Rate)) = 0 Then MsgBox "Fill Rate in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Rate: FGrid.SetFocus:  Exit Sub
        End If
    Next
    If Cnt = 0 Then MsgBox "Select Model", vbInformation, "Validation Chech": FGrid.SetFocus: Exit Sub
    Relg = FxReligion(Txt(Religion))
    Select Case Txt(Call_Status).TEXT
        Case "Cold"
            StCall = 0
        Case "Warm"
            StCall = 1
        Case "Hot"
            StCall = 2
    End Select
    GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If CustFlag Then
            If IsValid(Txt(ConPerson), "Contact Person") = False Then GoTo errlbl
        End If
    '   lp 11-03-03
        DocID = Txt(TxtDocID)
        If GCn.Execute("select count(*) from Veh_Quot where DocID='" & Txt(TxtDocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Quotation No. already exists, Retry", vbCritical, "Validation Error"
                Txt(SerialNo).SetFocus
                GoTo errlbl
            Else
                Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    MsgBox "Quotation No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If
        DocIdHlp = UCase(Replace(Txt(TxtDocID), " ", ""))
        If CustFlag Then
'            GCn.Execute ("delete from ProspectiveCust where cust_code='" & Txt(Party).Tag & "'")
            GCn.Execute ("insert into ProspectiveCust(cust_Code,Site_Code,NPrefix, Name,NSuffix, " & _
                " FPrefix,FName,Govt_YN,ConPerson,Add1, " & _
                "Add2,Add3,CityCode,PIN,STD, " & _
                "PhoneOff,PhoneResi,Mobile,FAX,EMail, " & _
                "AREA,REF_CODE,REP_CODE,Profession, " & _
                "Religion,FirstVeh_YN,Call_Status, " & _
                "U_Name, U_EntDt, U_AE ) " & _
                " values('" & Txt(Party).Tag & "','" & PubSiteCode & "','" & Txt(NPrefix).TEXT & "','" & Txt(Party).TEXT & "' ,'" & Txt(NSuffix).TEXT & "', " & _
                "'" & Txt(FPrefix).TEXT & "','" & Txt(fname).TEXT & "'," & IIf(Txt(Govt_YN).TEXT = "Yes", 1, 0) & ",'" & Txt(ConPerson).TEXT & "','" & Txt(Add1).TEXT & "', " & _
                "'" & Txt(Add2).TEXT & "','" & Txt(Add3).TEXT & "','" & Txt(City).Tag & "','" & Txt(Pin).TEXT & "','" & Txt(STD).TEXT & "', " & _
                "'" & Txt(PhoneOff).TEXT & "','" & Txt(PhoneResi).TEXT & "','" & Txt(Mobile).TEXT & "','" & Txt(FAx).TEXT & "','" & Txt(EMail).TEXT & "', " & _
                "'" & Txt(Area).Tag & "','" & Txt(REF_CODE).Tag & "','" & Txt(REP_CODE).Tag & "','" & Txt(Profession).Tag & "', " & _
                "" & Relg & "," & IIf(Txt(FirstVeh_YN).TEXT = "Yes", 1, 0) & "," & StCall & ", " & _
                "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
        End If
        GCn.Execute "insert into Veh_Quot(DocId,DocIDHelp,V_Type,V_No,Site_Code, " & _
            "V_Date,Party_Code,CityCode, " & _
            "Call_Status,AREA,REF_CODE,REP_CODE,Profession,Purpose,Intd_Use, " & _
            "FIN_YN,FB_CODE , Govt_YN, FirstVeh_YN, AMOUNT, " & _
            "RoundOff_YN,NARR1 , NARR2, MAP, FOLLOW_UP, " & _
            "Del_DATE,Exp_DATE, U_Name, U_EntDt, U_AE ) " & _
            "values('" & Txt(TxtDocID) & "','" & DocIdHlp & "','" & mVType & "'," & Val(Txt(SerialNo).TEXT) & ",'" & PubSiteCode & PubSiteCode & "'," & _
            "" & ConvertDate(Txt(VDate).TEXT) & ",'" & Txt(Party).Tag & "','" & Txt(City).Tag & "'," & _
            "" & StCall & ",'" & Txt(Area).Tag & "','" & Txt(REF_CODE).Tag & "','" & Txt(REP_CODE).Tag & "','" & Txt(Profession).Tag & "', '" & Txt(Purpose).Tag & "', '" & Txt(IndUse) & "'," & _
            "" & IIf(Txt(FIN_YN) = "Yes", 1, 0) & ",'" & Txt(FB_Code).Tag & "'," & IIf(Txt(Govt_YN) = "Yes", 1, 0) & "," & IIf(Txt(FirstVeh_YN) = "Yes", 1, 0) & "," & Val(Txt(Amount).TEXT) & "," & _
            "" & IIf(Txt(RoundOff_YN) = "Yes", 1, 0) & ",'" & Txt(NARR1).TEXT & "','" & Txt(NARR2).TEXT & "','" & Txt(MAP) & "','" & Txt(FOLLOW_UP).TEXT & "'," & _
            "" & ConvertDate(Txt(DEL_DATE).TEXT) & "," & ConvertDate(Txt(EXP_DATE).TEXT) & "," & _
            "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, Txt(TxtDocID), Txt(VDate)
    Else
        DocIdHlp = UCase(Replace(Txt(TxtDocID), " ", ""))
        GCn.Execute ("update Veh_Quot set Party_Code='" & Txt(Party).Tag & "',CityCode='" & Txt(City).Tag & "', " & _
            "Call_Status=" & StCall & ",AREA='" & Txt(Area).Tag & "',REF_CODE='" & Txt(REF_CODE).Tag & "',  " & _
            "REP_CODE='" & Txt(REP_CODE).Tag & "',Profession='" & Txt(Profession).Tag & "',PURPOSE='" & Txt(Purpose).Tag & "',INTD_USE='" & Txt(IndUse) & "',FIN_YN=" & IIf(Txt(FIN_YN) = "Yes", 1, 0) & ", " & _
            "FB_CODE='" & Txt(FB_Code).Tag & "' , Govt_YN=" & IIf(Txt(Govt_YN) = "Yes", 1, 0) & ", FirstVeh_YN=" & IIf(Txt(FirstVeh_YN) = "Yes", 1, 0) & ", " & _
            "AMOUNT=" & Val(Txt(Amount).TEXT) & ",RoundOff_YN=" & IIf(Txt(RoundOff_YN) = "Yes", 1, 0) & ",NARR1='" & Txt(NARR1).TEXT & "' , NARR2='" & Txt(NARR2).TEXT & "', " & _
            "MAP='" & Txt(MAP) & "', FOLLOW_UP='" & Txt(FOLLOW_UP).TEXT & "', " & _
            " Del_DATE=" & ConvertDate(Txt(DEL_DATE).TEXT) & ",Exp_DATE=" & ConvertDate(Txt(EXP_DATE).TEXT) & ", " & _
            " U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E'   " & _
            " where docid = '" & Txt(TxtDocID) & "'")
    End If
    GCn.Execute ("delete from Veh_Quot1 where docid='" & Txt(TxtDocID) & "'")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Model) <> "" And Val(FGrid.TextMatrix(I, Qty)) <> 0 Then
            GCn.Execute ("insert into Veh_Quot1(DocId,srl_no,DocIDHelp,V_Type,V_No,Site_Code, " & _
            "MODEL,QTY,RSO_WORK,TaxableRate_YN,RATE, " & _
            "SURCHARGE_Per,SUR_AMT,TAX_Per,TAX_AMT,AMOUNT,Book_DocId, " & _
            " U_Name, U_EntDt, U_AE ) " & _
            " values('" & Txt(TxtDocID) & "'," & I & ",'" & DocIdHlp & "','" & mVType & "'," & Val(Txt(SerialNo).TEXT) & ",'" & PubSiteCode & "', " & _
            " '" & FGrid.TextMatrix(I, Model) & "'," & Val(FGrid.TextMatrix(I, Qty)) & "," & IIf(FGrid.TextMatrix(I, RW) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ", " & Val(FGrid.TextMatrix(I, Rate)) & ", " & _
            " " & Val(FGrid.TextMatrix(I, SurchPer)) & "," & Val(FGrid.TextMatrix(I, SurchAmt)) & "," & Val(FGrid.TextMatrix(I, TaxPer)) & "," & Val(FGrid.TextMatrix(I, TaxAmt)) & "," & Val(FGrid.TextMatrix(I, Namt)) & ",'" & FGrid.TextMatrix(I, BookId) & "', " & _
            " '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2.CAPTION, 1) & "')")
        End If
    Next
    GCn.Execute ("delete from Veh_Quot2 where DocId='" & Txt(TxtDocID) & "'")
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, ADItem) <> "" And Val(FGrid1.TextMatrix(I, Qty1)) <> 0 Then
            GCn.Execute ("insert into Veh_Quot2(DocId,DocIdHelp,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE, U_Name, U_EntDt, U_AE) " & _
                "values('" & Txt(TxtDocID) & "','" & DocIdHlp & "'," & I & ",'" & PubSiteCode & PubSiteCode & "','" & mVType & "','" & Txt(SerialNo).TEXT & "', " & _
                "'" & FGrid1.TextMatrix(I, ADItemCode) & "','A'," & Val(FGrid1.TextMatrix(I, Qty1)) & "," & Val(FGrid1.TextMatrix(I, Rate1)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2.CAPTION, 1) & "')")
        End If
    Next
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Model) <> "" And Val(FGrid.TextMatrix(I, Qty)) <> 0 Then
            GCn.Execute ("delete from Veh_SubGroupQuot where PartyCode ='" & Txt(Party).Tag & "' and  StartDate =" & ConvertDate(Txt(VDate)) & " and Model = '" & FGrid.TextMatrix(I, Model) & "'")
            GCn.Execute ("insert into Veh_SubGroupQuot(PartyCode , StartDate, Model, " & _
                "ProspectiveCust_SubGroup, QuotDocId, QuotSrl_No, " & _
                "Site_Code,REP_CODE,Call_Status, " & _
                " U_Name, U_EntDt, U_AE ) " & _
                " values('" & Txt(Party).Tag & "'," & ConvertDate(Txt(VDate)) & ",'" & FGrid.TextMatrix(I, Model) & "', " & _
                " 0,'" & Txt(TxtDocID) & "'," & I & ", " & _
                " '" & PubSiteCode & PubSiteCode & "','" & Txt(REP_CODE).Tag & "',0, " & _
                " '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2.CAPTION, 1) & "')")
        End If
    Next
    'Update required fields in Master , by lps at un 11-04-03
    GCn.Execute ("Update ProspectiveCust set U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ",AREA='" & Txt(Area).Tag & "',REP_CODE='" & Txt(REP_CODE).Tag & "' Where cust_code='" & Txt(Party).Tag & "'")
    'eof
GCn.CommitTrans
mTrans = False
RsParty.Requery

If PubMoveRecYn Then
    Master.Requery
Else
    Set Master = GCn.Execute("select DocID as searchcode,docid from Veh_Quot where left(docid,1)='" & PubDivCode & "' And DocID = '" & Txt(TxtDocID) & "' order by V_date desc")
End If
Master.FIND "DocId = '" & Txt(TxtDocID) & "'"


If TopCtrl1.TopText2.CAPTION = "Add" Then
    If Val(Txt(SerialNo)) > DeCodeDocID(DocID, Document_No) Then
        MsgBox "Quotation No. " & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & Txt(SerialNo) & " alloted ", vbCritical, "Document No. Changed"
    End If
    Txt(VDate).Tag = Txt(VDate)
End If
TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    Dim sitecond As String
    sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
    
       sitecond = sitecond & " and " & cMID("Veh_Quot.DocId", "3", "1") & " ='" & PubSiteCode & "'"
    End If
    
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select DocID as searchcode, " & cDt("Veh_Quot.V_Date") & " As V_Date, " & cCStr("Veh_Quot.V_No", 10) & " As V_No, PC.Name, PC.ConPerson " & _
        " from Veh_Quot left join ProspectiveCust PC on Veh_Quot.Party_Code=PC.Cust_Code " & _
        " Where left(docID,1)='" & PubDivCode & "' " & sitecond & " order by V_date Desc"
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
        Set Master = GCn.Execute("select DocID as searchcode,docid from Veh_Quot where left(docid,1)='" & PubDivCode & "' And DocID = '" & MyValue & "' order by V_date desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
TxtGrid(0).Visible = False
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
Select Case Index
    Case Religion
        ListArray = Array("N/A", "Hindu", "Muslim", "Sikh", "Christian")
        Set mListItem = ListView_Items(ListView, Txt, Religion, ListArray, 5)
    Case NPrefix
        ListArray = Array("Mr.", "Mrs.", "Miss", "M/S")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 4)
    Case FPrefix
        ListArray = Array("S/O", "W/O", "D/O", "C/O", "And ", "U/C")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 6)
    Case Call_Status
        ListArray = Array("Cold", "Warm", "Hot")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 3)
    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(FB_Code).TEXT = "" Then Exit Sub
        If Txt(FB_Code).TEXT <> rsFin!Name Then
            rsFin.MoveFirst
            rsFin.FIND "name ='" & Txt(FB_Code).TEXT & "'"
        End If
    Case Party
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & Txt(Index).TEXT & "'"
        End If
    Case City
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(City).TEXT = "" Then Exit Sub
        If Txt(City).TEXT <> RsCity!Name Then
            RsCity.MoveFirst
            RsCity.FIND "name ='" & Txt(City).TEXT & "'"
        End If
    Case Area
        If RsArea.RecordCount = 0 Or (RsArea.EOF = True Or RsArea.BOF = True) Or Txt(Area).TEXT = "" Then Exit Sub
        If Txt(Area).TEXT <> RsArea!Name Then
            RsArea.MoveFirst
            RsArea.FIND "name ='" & Txt(Area).TEXT & "'"
        End If

    Case REF_CODE
        If RsRef.RecordCount = 0 Or (RsRef.EOF = True Or RsRef.BOF = True) Or Txt(REF_CODE).TEXT = "" Then Exit Sub
        If Txt(REF_CODE).TEXT <> RsRef!Name Then
            RsRef.MoveFirst
            RsRef.FIND "name ='" & Txt(REF_CODE).TEXT & "'"
        End If
    Case REP_CODE
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(REP_CODE).TEXT = "" Then Exit Sub
        If Txt(REP_CODE).TEXT <> RsRep!Name Then
            RsRep.MoveFirst
            RsRep.FIND "name ='" & Txt(REP_CODE).TEXT & "'"
        End If
    Case Profession
        If RsProf.RecordCount = 0 Or (RsProf.EOF = True Or RsProf.BOF = True) Or Txt(Profession).TEXT = "" Then Exit Sub
        If Txt(Profession).TEXT <> RsProf!Name Then
            RsProf.MoveFirst
            RsProf.FIND "name ='" & Txt(Profession).TEXT & "'"
        End If
    Case Purpose
        If RsPurpose.RecordCount = 0 Or (RsPurpose.EOF = True Or RsPurpose.BOF = True) Or Txt(Purpose).TEXT = "" Then Exit Sub
        If Txt(Purpose).TEXT <> RsPurpose!Name Then
            RsPurpose.MoveFirst
            RsPurpose.FIND "name ='" & Txt(Purpose).TEXT & "'"
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
    Case NPrefix
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1200
    Case FPrefix
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1800
    Case Religion
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1500
    Case Call_Status
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 900
    Case Party
        DGridTxtKeyDown_Mast DGParty, Txt, Party, RsParty, KeyCode, False, 1
    Case FB_Code
        DGridTxtKeyDown DGFin, Txt, Index, rsFin, KeyCode, False, 1, frmFinMast, "frmFinMast"
    Case City
        DGridTxtKeyDown DGCity, Txt, City, RsCity, KeyCode, False, 1, frmCity, "frmCity"
    Case Area
        DGridTxtKeyDown DGArea, Txt, Index, RsArea, KeyCode, False, 1, frmArea, "frmArea"
    Case REF_CODE
        DGridTxtKeyDown DGRef, Txt, Index, RsRef, KeyCode, False, 1
    Case REP_CODE
        DGridTxtKeyDown DGRep, Txt, Index, RsRep, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
    Case Profession
        DGridTxtKeyDown DGProf, Txt, Index, RsProf, KeyCode, False, 1
    Case Purpose
        DGridTxtKeyDown DGPurpose, Txt, Index, RsPurpose, KeyCode, False, 1
End Select
If DGPurpose.Visible = False And FrmList.Visible = False And DGFin.Visible = False And DGCity.Visible = False And DGRep.Visible = False And DGRef.Visible = False And DGProf.Visible = False And DGArea.Visible = False And DGParty.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VDate Then Txt_Validate Index, True
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> MAP Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = MAP Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> VDate Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> FIN_YN Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case Index
    'Case Party
    '    If DGParty.Visible = True Then DGridTxtKeyPress Txt, Index, RsParty, KeyAscii, "Name"
    Case FB_Code
        If DGFin.Visible = True Then DGridTxtKeyPress Txt, Index, rsFin, KeyAscii, "Name"
    Case City
        If DGCity.Visible = True Then DGridTxtKeyPress Txt, Index, RsCity, KeyAscii, "Name"
    Case Area
        If DGArea.Visible = True Then DGridTxtKeyPress Txt, Index, RsArea, KeyAscii, "Name"
    Case REP_CODE
        If DGRep.Visible = True Then DGridTxtKeyPress Txt, Index, RsRep, KeyAscii, "Name"
    Case REF_CODE
        If DGRef.Visible = True Then DGridTxtKeyPress Txt, Index, RsRef, KeyAscii, "Name"
    Case Profession
        If DGProf.Visible = True Then DGridTxtKeyPress Txt, Index, RsProf, KeyAscii, "Name"
    Case Purpose
        If DGPurpose.Visible = True Then DGridTxtKeyPress Txt, Index, RsPurpose, KeyAscii, "Name"
    Case Govt_YN, RoundOff_YN, FirstVeh_YN
        If UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
        ElseIf UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
    Case FIN_YN
        If UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
        ElseIf UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
        Txt(FB_Code).TEXT = ""
        Txt(FB_Code).Tag = ""
    Case SerialNo
        Call NumPress(Txt(Index), KeyAscii, 6, 0)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyUp_Mast Txt, Index, RsParty, KeyCode, "Name"
    Case Religion, Call_Status, NPrefix, FPrefix
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
End Select
Amt_Cal
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Select Case Index
    Case Religion, Call_Status, NPrefix, FPrefix
        If Txt(Index).TEXT <> "" Then Txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case Party
        If IsValid(Txt(Index), "Party") = False Then Cancel = True: Exit Sub
        If Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        End If
        'AT UDAIPUR
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
            'Txt(Index).Text = ""
            Txt(Index).Tag = ""
            CreateNewParty (Index)
        Else
            If UCase(Trim(Txt(Index))) <> UCase(Trim(RsParty!Name)) Then
                Txt(Index).Tag = ""
                CreateNewParty (Index)
            Else
                Txt(Index).TEXT = RsParty!Name
                Txt(Index).Tag = RsParty!Code
                Fill_Data
            End If
        End If
    Case NSuffix
        GSQL = "Select Name from ProspectiveCust where " & cUCase(cTrim("Name") & " + " & cTrim("Nsuffix")) & " ='" & UCase(Trim(Txt(Party)) + Trim(Txt(NSuffix))) & "'"
        If GCn.Execute(GSQL).RecordCount <= 0 Then
            CreateNewParty Party, True
        End If
    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = rsFin!Name
            Txt(Index).Tag = rsFin!Code
        End If
    Case City
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsCity!Name
            Txt(Index).Tag = RsCity!Code
        End If
    Case Area
        If RsArea.RecordCount = 0 Or (RsArea.EOF = True Or RsArea.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsArea!Name
            Txt(Index).Tag = RsArea!Code
        End If
    Case REP_CODE
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsRep!Name
            Txt(Index).Tag = RsRep!Code
        End If
    Case Profession
        If RsProf.RecordCount = 0 Or (RsProf.EOF = True Or RsProf.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsProf!Name
            Txt(Index).Tag = RsProf!Code
        End If
    Case Purpose
        If RsPurpose.RecordCount = 0 Or (RsPurpose.EOF = True Or RsPurpose.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsPurpose!Name
            Txt(Index).Tag = RsPurpose!Code
        End If
    Case REF_CODE
        If RsRef.RecordCount = 0 Or (RsRef.EOF = True Or RsRef.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsRef!Name
            Txt(Index).Tag = RsRef!Code
        End If
    Case DEL_DATE
        Txt(Index).TEXT = RetDate(Txt(Index))
    Case EXP_DATE
        Txt(Index).TEXT = RetDate(Txt(Index))
    Case VDate
        If Len(Trim(Txt(VDate).TEXT)) = 0 Then
            Txt(VDate).TEXT = PubLoginDate
        Else
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
        If CheckFinYear(Txt(Index)) Then
            Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
            DocID = Txt(TxtDocID)
            Dim ValidQuot As Integer
            ValidQuot = GCn.Execute("select Valid_Day from syctrl").Fields(0).Value
            Txt(EXP_DATE).TEXT = Format(DateAdd("D", ValidQuot, Txt(VDate).TEXT), "dd/mmm/yyyy")
        Else
            Cancel = True
        End If
    Case SerialNo
        If IsValid(Txt(SerialNo), "SerialNo") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag = True Then      ' Manual
            Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
            DocID = Txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select * From Veh_Quot Where docid='" & Txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                Txt(SerialNo).SetFocus
            End If
        End If
End Select
Set Rst = Nothing
End Sub

Private Sub DGMod_Click()
If RsMod.RecordCount > 0 Then
    TxtGrid(0).TEXT = RsMod!Code
    FGrid.TextMatrix(FGrid.Row, Model) = RsMod!Code
End If
TxtGrid(0).SetFocus
DGMod.Visible = False
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        Txt(Party).TEXT = RsParty!Name
        Txt(Party).Tag = RsParty!Code
    End If
    Txt(Party).SetFocus
    DGParty.Visible = False
End Sub
Private Sub DGCity_Click()
If RsCity.RecordCount > 0 Then
    Txt(City).TEXT = RsCity!Name
    Txt(City).Tag = RsCity!Code
End If
Txt(City).SetFocus
DGCity.Visible = False
End Sub
Private Sub DGFin_Click()
If rsFin.RecordCount > 0 Then
    Txt(FB_Code).TEXT = rsFin!Name
    Txt(FB_Code).Tag = rsFin!Code
End If
Txt(FB_Code).SetFocus
DGFin.Visible = False
End Sub
Private Sub DGArea_Click()
    If RsArea.RecordCount > 0 Then
        Txt(Area).TEXT = RsArea!Name
        Txt(Area).Tag = RsArea!Code
    End If
    Txt(Area).SetFocus
    DGArea.Visible = False
End Sub

Private Sub DGProf_Click()
    If RsProf.RecordCount > 0 Then
        Txt(Profession).TEXT = RsProf!Name
        Txt(Profession).Tag = RsProf!Code
    End If
    Txt(Profession).SetFocus
    DGProf.Visible = False
End Sub

Private Sub DGRef_Click()
    If RsRef.RecordCount > 0 Then
        Txt(REF_CODE).TEXT = RsRef!Name
        Txt(REF_CODE).Tag = RsRef!Code
    End If
    Txt(REF_CODE).SetFocus
    DGRef.Visible = False
End Sub

Private Sub DGRep_Click()
    If RsRep.RecordCount > 0 Then
        Txt(REP_CODE).TEXT = RsRep!Name
        Txt(REP_CODE).Tag = RsRep!Code
    End If
    Txt(REP_CODE).SetFocus
    DGRep.Visible = False
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress (vbKeyReturn)
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
    If FGrid.BackColorSel = BackColorSelLeave Then FGrid.Col = 1
    FGrid.BackColorSel = BackColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case Model
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case ModelGroup
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, ModelGroupCode) = ""
        Case Qty
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Amt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = ""
            FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
        Case Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Amt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = ""
            FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
        Case TaxAmt
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = ""
            FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
        Case TaxPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt) = ""
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = ""
            FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
        Case SurchAmt
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
        Case SurchPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = ""
    End Select
    Amt_Cal1
    Amt_Cal
End If
If KeyCode = 13 Then TAddMode = False
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.Row > 3 Then MsgBox "You can't take more than three models in a Quotation", vbInformation: FGrid.SetFocus:   Exit Sub
Select Case FGrid.Col
    Case Model, Taxable, RW, ModelGroup
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case Amt
        FGrid_LeaveCell
        FGrid.Col = FGrid.Col + 1
'        FGrid_EnterCell
        FGrid.SetFocus
    Case Qty, Rate, TaxPer, SurchPer
       Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
End Select
Amt_Cal1
Amt_Cal
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
         Amt_Cal1
         Amt_Cal
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
DGMod.Visible = False
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
'    FGrid.CellForeColor = CellForeColLeave
End Sub

Private Sub FGrid_LostFocus()
    If TxtGrid(0).Visible = False Then FGrid.BackColorSel = BackColorSelLeave
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
    If I <> VDate Then
        Txt(I).Tag = ""
    End If
Next I
DocID = ""
End Sub

Private Sub MoveRec()
Dim Master1 As Recordset, Rs As Recordset
Dim RsPro As ADODB.Recordset, I As Integer
On Error GoTo error1
If Master.RecordCount > 0 Then
    Set Master1 = New ADODB.Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select veh_quot.* from Veh_Quot where DocID = '" & Master!SearchCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    DocID = Master1!DocID
    Txt(TxtDocID).TEXT = Master1!DocID
    mVType = Master1!V_Type
    LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
    LblSite.CAPTION = "Site Code : " & Master1!Site_Code
    LblVPrefix.CAPTION = DeCodeDocID(Master1!DocID, Document_Prefix)
    Txt(SerialNo).TEXT = Master1!V_NO
    Txt(VDate).TEXT = Master1!V_DATE
    
    Txt(Purpose).Tag = IIf(IsNull(Master1!Purpose), "", Master1!Purpose)
    If Txt(Purpose).Tag <> "" Then
        Set Rs = GCn.Execute("select PurposeName from Purpose where Purposecode = '" & Txt(Purpose).Tag & "'")
        If Rs.RecordCount > 0 Then Txt(Purpose).TEXT = Rs(0)
    Else
        Txt(Purpose).TEXT = ""
    End If
    Txt(IndUse) = IIf(IsNull(Master1!INTD_USE), "", Master1!INTD_USE)
    
    Txt(Party).Tag = IIf(IsNull(Master1!Party_code), "", Master1!Party_code)
    If Txt(Party).Tag <> "" And GCn.Execute("select ProspectiveCust.* from ProspectiveCust where ProspectiveCust.Cust_code = '" & Txt(Party).Tag & "'").RecordCount > 0 Then
        Set RsPro = New Recordset
        Set RsPro = GCn.Execute("select ProspectiveCust.* from ProspectiveCust where ProspectiveCust.Cust_code = '" & Txt(Party).Tag & "'")
        Txt(Party) = IIf(IsNull(RsPro!Name), "", RsPro!Name)
        Txt(NPrefix) = IIf(IsNull(RsPro!NPrefix), "", RsPro!NPrefix)
        Txt(NSuffix) = IIf(IsNull(RsPro!NSuffix), "", RsPro!NSuffix)
        Txt(ConPerson) = IIf(IsNull(RsPro!ConPerson), "", RsPro!ConPerson)
        Txt(FPrefix) = IIf(IsNull(RsPro!FPrefix), "", RsPro!FPrefix)
        Txt(fname) = IIf(IsNull(RsPro!fname), "", RsPro!fname)
        Txt(Religion).TEXT = FxReligion(IIf(IsNull(RsPro!Religion), 0, RsPro!Religion))
        If Not IsNull(RsPro!Call_Status) Then
            Select Case RsPro!Call_Status
                Case 0
                   Txt(Call_Status) = "Cold"
                Case 1
                   Txt(Call_Status) = "Warm"
                Case 2
                    Txt(Call_Status) = "Hot"
            End Select
        End If
        Txt(Add1) = IIf(IsNull(RsPro!Add1), "", RsPro!Add1)
        Txt(Add2) = IIf(IsNull(RsPro!Add2), "", RsPro!Add2)
        Txt(Add3) = IIf(IsNull(RsPro!Add3), "", RsPro!Add3)
        
        Txt(Profession).Tag = IIf(IsNull(RsPro!Profession), "", RsPro!Profession)
        If Txt(Profession).Tag <> "" And GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").RecordCount > 0 Then
            Txt(Profession).TEXT = GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").Fields(0).Value
        Else
            Txt(Profession).TEXT = ""
        End If

        Txt(City).Tag = IIf(IsNull(RsPro!CityCode), "", RsPro!CityCode)
        If Txt(City).Tag <> "" And GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").RecordCount > 0 Then
            Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
        Else
            Txt(City).TEXT = ""
        End If
        Txt(Area).Tag = IIf(IsNull(RsPro!Area), "", RsPro!Area)
        If Txt(Area).Tag <> "" And GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").RecordCount > 0 Then
            Txt(Area).TEXT = GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").Fields(0).Value
        Else
            Txt(Area).TEXT = ""
        End If
        Txt(REP_CODE).Tag = IIf(IsNull(RsPro!REP_CODE), "", RsPro!REP_CODE)
        If Txt(REP_CODE).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").RecordCount > 0 Then
            Txt(REP_CODE).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").Fields(0).Value
        Else
            Txt(REP_CODE).TEXT = ""
        End If
        Txt(REF_CODE).Tag = IIf(IsNull(RsPro!REF_CODE), "", RsPro!REF_CODE)
        If Txt(REF_CODE).Tag <> "" And GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").RecordCount > 0 Then
            Txt(REF_CODE).TEXT = GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").Fields(0).Value
        Else
            Txt(REF_CODE).TEXT = ""
        End If
        
        Txt(Pin) = IIf(IsNull(RsPro!Pin), "", RsPro!Pin)
        Txt(STD) = IIf(IsNull(RsPro!STD), "", RsPro!STD)
        Txt(PhoneOff) = IIf(IsNull(RsPro!PhoneOff), "", RsPro!PhoneOff)
        Txt(PhoneResi) = IIf(IsNull(RsPro!PhoneResi), "", RsPro!PhoneResi)
        Txt(Mobile) = IIf(IsNull(RsPro!Mobile), "", RsPro!Mobile)
        Txt(EMail) = IIf(IsNull(RsPro!EMail), "", RsPro!EMail)
        Txt(FAx) = IIf(IsNull(RsPro!FAx), "", RsPro!FAx)
        Txt(Govt_YN) = IIf(RsPro!Govt_YN = 1, "Yes", "No")
        Txt(FirstVeh_YN) = IIf(RsPro!FirstVeh_YN = 1, "Yes", "No")
    End If
    Txt(FIN_YN) = IIf(Master1!FIN_YN = 1, "Yes", "No")
    Txt(RoundOff_YN) = IIf(Master1!RoundOff_YN = 1, "Yes", "No")
    Txt(Amount) = Format(IIf(IsNull(Master1!Amount), 0, Master1!Amount), "0.00")
    Txt(NARR1) = IIf(IsNull(Master1!NARR1), "", Master1!NARR1)
    Txt(NARR2) = IIf(IsNull(Master1!NARR2), "", Master1!NARR2)
    Txt(FOLLOW_UP) = IIf(IsNull(Master1!FOLLOW_UP), "", Master1!FOLLOW_UP)
    Txt(MAP) = IIf(IsNull(Master1!MAP), "", Master1!MAP)
'    txt(Book_SiteCode) = IIf(IsNull(Master1!Book_SiteCode), "", Master1!Book_SiteCode)
'    txt(Book_DocId) = IIf(IsNull(Master1!Book_DocId), "", Master1!Book_DocId)
    Txt(DEL_DATE) = IIf(IsNull(Master1!DEL_DATE), "", Master1!DEL_DATE)
    Txt(EXP_DATE) = IIf(IsNull(Master1!EXP_DATE), "", Master1!EXP_DATE)
    Txt(FB_Code).Tag = IIf(IsNull(Master1!FB_Code), "", Master1!FB_Code)
    If Txt(FB_Code).Tag <> "" And GCn.Execute("select fincode as code,finname as name from ContractFinance where fincatg = 0 and  fincode = '" & Txt(FB_Code).Tag & "'").RecordCount > 0 Then
'        txt(FB_Code).Text = GCn.Execute("select finname as name from ContractFinance where fincatg = 0 and  fincode = '" & txt(FB_Code).Tag & "'").Fields(0).Value
        Set Rs = New Recordset
        Rs.CursorLocation = adUseClient
        Rs.Open "select fincode as code,finname + ',' + " & xIsNull("City.CityName", "") & " as name,AcCode " & _
        " from ContractFinance left join city on left(ContractFinance.City,4)=City.CityCode " & _
        " where fincatg = 0 and  fincode = '" & Txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        Txt(FB_Code).TEXT = XNull(Rs!Name)
    Else
        Txt(FB_Code).TEXT = ""
        Txt(FB_Code).Tag = ""
    End If
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Q.*, MG.ModelGrp_Code, MG.ModelGrp_Name " & _
                         "FROM veh_Quot1   Q " & _
                         "LEFT JOIN model M ON Q.MODEL =M.MODEL " & _
                         "LEFT JOIN Model_Grp MG ON M.Grp_Code = MG.ModelGrp_Code " & _
                         "where Q.docId = '" & Master1!DocID & "'")
    I = 1
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            With FGrid
                .AddItem ""
                .TextMatrix(I, 0) = Rs!Srl_No
                .TextMatrix(I, Model) = Rs!Model
                .TextMatrix(I, ModelGroup) = Rs!ModelGrp_Name
                .TextMatrix(I, ModelGroupCode) = Rs!ModelGrp_Code
                .TextMatrix(I, Taxable) = IIf(Rs!TaxableRate_YN = 0, "No", "Yes")
                .TextMatrix(I, RW) = IIf(Rs!RSO_WORK = 0, "No", "Yes")
                .TextMatrix(I, Qty) = IIf(Rs!Qty <> 0, Format(Rs!Qty, "0"), "")
                .TextMatrix(I, Rate) = IIf(Rs!Rate <> 0, Format(Rs!Rate, "0.00"), "")
                .TextMatrix(I, Amt) = Format((Rs!Qty * Rs!Rate), "0.00")
                .TextMatrix(I, TaxPer) = IIf(Rs!Tax_Per <> 0, Format(Rs!Tax_Per, "0.00"), "")
                .TextMatrix(I, TaxAmt) = Format(Rs!Tax_Amt, "0.00")
                .TextMatrix(I, SurchPer) = IIf(Rs!surcharge_per <> 0, Format(Rs!surcharge_per, "0.00"), "")
                .TextMatrix(I, SurchAmt) = Format(Rs!sur_amt, "0.00")
                .TextMatrix(I, Namt) = Format(Rs!Amount, "0.00")
                .TextMatrix(I, BookId) = Rs!Book_DocId
            End With
            Rs.MoveNext
            I = I + 1
'            FGrid.AddItem rs!Srl_No & Chr(9) & rs!Model & Chr(9) & IIf(rs!TaxableRate_YN = 0, "No", "Yes") & Chr(9) & IIf(rs!RSO_WORK = 0, "No", "Yes") & Chr(9) & IIf(rs!Qty <> 0, Format(rs!Qty, "0"), "") & Chr(9) & _
            Format(rs!Rate, "0.00") & Chr(9) & Format((rs!Qty * rs!Rate), "0.00") & Chr(9) & IIf(rs!Tax_Per <> 0, Format(rs!Tax_Per, "0.00"), "") & Chr(9) & Format(rs!Tax_Amt, "0.00") & Chr(9) & IIf(rs!surcharge_per <> 0, Format(rs!surcharge_per, "0.00"), "") & Chr(9) & Format(rs!sur_amt, "0.00") & Chr(9) & Format(rs!AMOUNT, "0.00") & Chr(9) & rs!Book_DocId
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Veh_AMDModel.Prod_Name, Veh_Quot2.Srl_No, Veh_Quot2.PROD_CODE, Veh_Quot2.QTY, Veh_Quot2.RATE, Veh_Quot2.Trn_Type " & _
        "FROM Veh_Quot2 LEFT JOIN Veh_AMDModel ON Veh_Quot2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_Quot2.DocId = '" & Master1!DocID & "'")
    FGrid1.Rows = 1: FGrid1.Redraw = False
    I = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            With FGrid1
                .AddItem ""
                .TextMatrix(I, 0) = Rs!Srl_No
                .TextMatrix(I, ADItem) = Rs!Prod_Name
                .TextMatrix(I, ADType) = IIf(Rs!Trn_Type = "A", "Addition", IIf(Rs!Trn_Type = "D", "Deletion", "Shortage"))
                .TextMatrix(I, Qty1) = Format(IIf(IsNull(Rs!Qty), "", Rs!Qty), "0")
                .TextMatrix(I, Rate1) = Format(IIf(IsNull(Rs!Rate), "", Rs!Rate), "0.00")
                .TextMatrix(I, Amt1) = Format(.TextMatrix(I, Qty1) * .TextMatrix(I, Rate1), "0.00")
                .TextMatrix(I, ADItemCode) = Rs!Prod_Code
            End With
            Rs.MoveNext
           I = I + 1
        Loop
        FGrid1.FixedRows = 1
    Else
        FGrid1.AddItem FGrid1.Rows
        FGrid1.FixedRows = 1
    End If
    FGrid1.Redraw = True
    Set Rs = Nothing
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
Grid_Hide
'Call Amt_Cal
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte
'SrNo.0|Model1|RSO/Work2|Tax3|Quantiy4|Rate5|Tax%6|TaxAmt7|Surch%8|SurchAmt9|Amount10
    With FGrid
        .Cols = 15
        .left = Me.left '+ 60
        .width = Me.width - 90
        .top = 4035
        .RowHeightMin = PubGridRowHeight
        
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, ModelGroup) = "Model Grp"
        .ColAlignment(ModelGroup) = flexAlignLeftCenter
        .ColWidth(ModelGroup) = 1800


        .TextMatrix(0, Model) = "Model"
        .ColAlignment(Model) = flexAlignLeftCenter
        .ColWidth(Model) = 2200
      
        .TextMatrix(0, Taxable) = "Tax"
        .ColAlignment(Taxable) = flexAlignLeftCenter
        .ColWidth(Taxable) = 450
        
        .TextMatrix(0, RW) = "RSO"
        .ColAlignment(RW) = flexAlignLeftCenter
        .ColWidth(RW) = 450

        .ColAlignmentFixed(Qty) = flexAlignRightCenter
        .TextMatrix(0, Qty) = "Qty"
        .ColAlignment(Qty) = flexAlignRightCenter
        .ColWidth(Qty) = 495
        
        .TextMatrix(0, Rate) = "Rate"
        .ColAlignmentFixed(Rate) = flexAlignRightCenter
        .ColWidth(Rate) = 1000

        .TextMatrix(0, Namt) = "NetAmt"
        .ColAlignmentFixed(Namt) = flexAlignRightCenter
        .ColWidth(Namt) = 1065

        .TextMatrix(0, TaxPer) = "Tax%"
        .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
        .ColWidth(TaxPer) = 600

        .TextMatrix(0, TaxAmt) = "TaxAmt"
        .ColAlignmentFixed(TaxAmt) = flexAlignRightCenter
        .ColWidth(TaxAmt) = 840
        
        .TextMatrix(0, SurchPer) = "Surch%"
        .ColAlignmentFixed(SurchPer) = flexAlignRightCenter
        .ColWidth(SurchPer) = 750

        .TextMatrix(0, SurchAmt) = "SurchAmt"
        .ColAlignmentFixed(SurchAmt) = flexAlignRightCenter
        .ColWidth(SurchAmt) = 800
        
        .TextMatrix(0, Amt) = "Amount"
        .ColAlignmentFixed(Amt) = flexAlignRightCenter
        .ColWidth(Amt) = 1065
        
        .TextMatrix(0, BookId) = "BookDocId"
        .ColAlignment(BookId) = flexAlignLeftCenter
        .ColWidth(BookId) = 1300
        
        .TextMatrix(0, ModelGroupCode) = "ModelGroupCode"
        .ColAlignment(ModelGroupCode) = flexAlignLeftCenter
        .ColWidth(ModelGroupCode) = 1300
        
    End With
        
    With FGrid1
        .left = Me.left
        .width = 7200
        .top = 5070
        .RowHeightMin = PubGridRowHeight

        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, ADItem) = "Additional Fitments"
        .ColAlignment(ADItem) = flexAlignLeftCenter
        .ColWidth(ADItem) = 2400
'
'        .TextMatrix(0, ADType) = "Type"
'        .ColAlignment(ADType) = flexAlignLeftCenter
'        .ColWidth(ADType) = 1200
         .ColWidth(ADType) = 0
       
        .TextMatrix(0, Qty1) = "Qty"
        .ColAlignmentFixed(Qty1) = flexAlignRightCenter
        .ColWidth(Qty1) = 645

        .TextMatrix(0, Rate1) = "Rate"
        .ColAlignmentFixed(Rate1) = flexAlignRightCenter
        .ColWidth(Rate1) = 855
        
        .TextMatrix(0, Amt1) = "Amount"
        .ColAlignmentFixed(Amt1) = flexAlignRightCenter
        .ColWidth(Amt1) = 1065
        
        .ColWidth(ADItemCode) = 0
        
    End With
    
    BackColorSelLeave = FGrid.BackColor
DGParty.left = 0: DGParty.top = FGrid.top
DGParty.width = Me.width - 90:: DGParty.height = Me.height - (DGParty.top + mBotScale)

DGCity.left = Me.width - (DGCity.width + mRtScale): DGCity.top = mTopScale
DGRep.left = Me.width - (DGRep.width + mRtScale): DGRep.top = mTopScale
DGFin.left = Me.width - (DGFin.width + mRtScale): DGFin.top = mTopScale
DGProf.left = Me.width - (DGProf.width + mRtScale): DGProf.top = mTopScale
DGPurpose.left = Me.width - (DGPurpose.width + mRtScale): DGPurpose.top = mTopScale
DGRef.left = Me.width - (DGRef.width + mRtScale): DGRef.top = mTopScale
DGArea.left = Me.width - (DGArea.width + mRtScale): DGArea.top = mTopScale
DGMod.left = Me.width - (DGMod.width + mRtScale): DGMod.top = mTopScale
DGMod.height = FGrid.top - DGMod.top
DGADItem.left = Me.width - (DGADItem.width + mRtScale): DGADItem.top = mTopScale
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
'    txt(i).ForeColor = CtrlFColOrg
Next
Txt(TxtDocID).Enabled = False
If TopCtrl1.TopText2 = "Edit" Then
    Txt(VDate).Enabled = False
    Txt(SerialNo).Enabled = False
    For I = 1 To 24
        Txt(I).Enabled = False
    Next
    Txt(34).Enabled = False
End If
txtDisabled_Color Me
TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
txtgrid1(0).BackColor = CtrlBCol
txtgrid1(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If DGCity.Visible Then DGCity.Visible = False
    If DGFin.Visible Then DGFin.Visible = False
    If DGRep.Visible Then DGRep.Visible = False
    If DGRef.Visible Then DGRef.Visible = False
    If DGProf.Visible Then DGProf.Visible = False
    If DGPurpose.Visible Then DGPurpose.Visible = False
    If DGArea.Visible Then DGArea.Visible = False
    If FrmList.Visible Then FrmList.Visible = False
    If DGParty.Visible Then DGParty.Visible = False
    If DGADItem.Visible Then DGADItem.Visible = False
    If DGMod.Visible Then DGMod.Visible = False
End Sub
Private Sub Amt_Cal1()
    If FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes" Then
        FGrid.TextMatrix(FGrid.Row, Namt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) + Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))), "0.00")
    Else
        FGrid.TextMatrix(FGrid.Row, Namt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
        FGrid.TextMatrix(FGrid.Row, TaxPer) = 0: FGrid.TextMatrix(FGrid.Row, TaxAmt) = 0
        FGrid.TextMatrix(FGrid.Row, SurchPer) = 0: FGrid.TextMatrix(FGrid.Row, SurchAmt) = 0
    End If
End Sub
 
 Private Sub Amt_Cal()
 Dim I As Byte
 Dim IAmt As Double
 Dim ICnt As Integer
 Dim TotAdd As Double
 Dim TotDel As Double
 Dim TotAdd1 As Double
 Dim TotDel1 As Double

 For I = 1 To FGrid.Rows - 1
    If FGrid.TextMatrix(I, Model) <> "" Then
        IAmt = IAmt + Val(FGrid.TextMatrix(I, Namt))
        ICnt = ICnt + 1
    End If
 Next I
 
     For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, ADItem) <> "" Then
'            If FGrid1.TextMatrix(i, ADType) = "Shortage" Then
'                FGrid1.TextMatrix(i, Rate1) = "0.00"
'                FGrid1.TextMatrix(i, Amt1) = "0.00"
'            End If
'            If FGrid1.TextMatrix(i, ADType) = "Addition" Then
                TotAdd = TotAdd + Val(FGrid1.TextMatrix(I, Amt1))
'            ElseIf FGrid1.TextMatrix(i, ADType) = "Deletion" Then
'                TotDel = TotDel + Val(FGrid1.TextMatrix(i, Amt1))
'            End If
        End If
    Next
    TotAdd1 = TotAdd * ICnt
    'TotDel1 = TotDel * ICnt

 
     Txt(Amount).TEXT = Format(IAmt + TotAdd1 - TotDel1, "0.00")
 End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
         Case Model
            If FGrid.TextMatrix(FGrid.Row, ModelGroupCode) <> "" Then
                RsMod.Filter = adFilterNone
                RsMod.Filter = "Grp_Code ='" & FGrid.TextMatrix(FGrid.Row, ModelGroupCode) & "'"
            End If
            
            If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Model) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Model) <> RsMod!Code Then
                RsMod.MoveFirst
                RsMod.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, Model) & "'"
            End If
            
        Case ModelGroup
            DgModelGroup.Move TxtGrid(Index).left, TxtGrid(Index).top + TxtGrid(Index).height + 30
            If RsModelGroup.RecordCount = 0 Or (RsModelGroup.EOF = True Or RsModelGroup.BOF = True) Or Txt(ModelGroup).TEXT = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, ModelGroupCode) <> RsModelGroup!Code Then
                RsModelGroup.MoveFirst
                RsModelGroup.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, ModelGroupCode) & "'"
            End If
        
        Case Rate, TaxPer, TaxAmt, SurchPer, SurchAmt, Qty
'           SendKeys "{HOME}+{END}"
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        FGrid.SetFocus
        TxtGrid(0).Visible = False
        Grid_Hide
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Model    '1
            DGridTxtKeyDown DGMod, TxtGrid, Index, RsMod, KeyCode, True, 0, frmModel, "frmModel"
            If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SurchPer
                    End If
            End If
            
        Case ModelGroup
            DGridTxtKeyDown DgModelGroup, TxtGrid, Index, RsModelGroup, KeyCode, True, 0, frmModelGrp, "frmModelGrp"
            If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SurchPer
                    End If
            End If
            
        Case TaxPer, SurchPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SurchPer, 1
                End If
            End If
        Case Taxable, RW, Qty, Rate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SurchPer
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case Model
        If DGMod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsMod, KeyAscii, "code"
        
    Case ModelGroup
        If DgModelGroup.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsModelGroup, KeyAscii, "Name"
        
    Case Taxable
        If UCase(Chr(KeyAscii)) = "Y" Then
            TxtGrid(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            TxtGrid(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            TxtGrid(Index) = ""
        End If
        KeyAscii = 0
        
    Case RW
        If UCase(Chr(KeyAscii)) = "Y" Then
            TxtGrid(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            TxtGrid(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            TxtGrid(Index) = ""
        End If
        KeyAscii = 0
    Case TaxAmt, SurchAmt
        KeyAscii = 0
    Case TaxPer, SurchPer
        Call NumPress(TxtGrid(Index), KeyAscii, 3, 2)
    Case Rate, TaxPer, TaxAmt, SurchPer, SurchAmt
        Call NumPress(TxtGrid(Index), KeyAscii, 8, 2)
    Case Qty
        Call NumPress(TxtGrid(Index), KeyAscii, 2, 0)
End Select
End Sub


Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
        Select Case FGrid.Col
            Case Model
                If KeyCode <> 13 And DGMod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsMod, KeyCode, "code", True
            Case ModelGroup
                If KeyCode <> 13 And DgModelGroup.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsModelGroup, KeyCode, "Name", True
                
            Case Qty
                FGrid.TextMatrix(FGrid.Row, Qty) = Format(Val(TxtGrid(Index).TEXT), "0")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
                Else
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, TaxAmt))) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
                Else
                   FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
                End If
            Case Rate
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
                Else
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, TaxAmt))) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
                Else
                   FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
                End If
            Case TaxAmt
                FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
                Else
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, TaxAmt))) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
                Else
                   FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
                End If
            Case TaxPer
                FGrid.TextMatrix(FGrid.Row, TaxPer) = TxtGrid(Index).TEXT
                FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
                Else
                   FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
                End If
            Case SurchAmt
                FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
                Else
                   FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
                End If
            Case SurchPer
                FGrid.TextMatrix(FGrid.Row, SurchPer) = TxtGrid(Index).TEXT
                FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
            Case Taxable, RW
                If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                    TxtGrid(Index) = ""
                ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
                    TxtGrid(Index) = "Yes"
                Else
                TxtGrid(Index) = "No"
                End If
        End Select
        Amt_Cal1
        Amt_Cal
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index)
Exit Sub
ELoop:
    CheckError
End Sub
Private Function TxtGridLeave(Optional Index As Integer) As Boolean
Dim j As Integer
Select Case FGrid.Col
        Case Model
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Model) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Model) = RsMod!Code
                FGrid.TextMatrix(FGrid.Row, Rate) = RsMod!Sale_Rate
            End If
            If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
        Case ModelGroup
            If RsModelGroup.RecordCount = 0 Or (RsModelGroup.EOF = True Or RsModelGroup.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, ModelGroup) = ""
                FGrid.TextMatrix(FGrid.Row, ModelGroupCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, ModelGroupCode) = RsModelGroup!Code
                FGrid.TextMatrix(FGrid.Row, ModelGroup) = RsModelGroup!Name
            End If
            
        Case Taxable
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
            If Txt(VDate) <> "" And FGrid.TextMatrix(FGrid.Row, Model) <> "" _
                And FGrid.TextMatrix(FGrid.Row, Taxable) <> "" And FGrid.TextMatrix(FGrid.Row, RW) <> "" Then
                'FGrid.TextMatrix(FGrid.Row, Rate) = Format(VehSRate(txt(VDate), FGrid.TextMatrix(FGrid.Row, Model), FGrid.TextMatrix(FGrid.Row, Taxable), FGrid.TextMatrix(FGrid.Row, RW)), "0.00")
            End If
        Case RW
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
            If Txt(VDate) <> "" And FGrid.TextMatrix(FGrid.Row, Model) <> "" _
                And FGrid.TextMatrix(FGrid.Row, Taxable) <> "" And FGrid.TextMatrix(FGrid.Row, RW) <> "" Then
                'FGrid.TextMatrix(FGrid.Row, Rate) = Format(VehSRate(txt(VDate), FGrid.TextMatrix(FGrid.Row, Model), FGrid.TextMatrix(FGrid.Row, Taxable), FGrid.TextMatrix(FGrid.Row, RW)), "0.00")
            End If
        Case Qty
            FGrid.TextMatrix(FGrid.Row, Qty) = Format(Val(TxtGrid(0).TEXT), "0")
            FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
            FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, TaxPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, TaxAmt))) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
            End If
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
            Else
               FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
            End If
        Case Rate
            FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(0).TEXT), "0.00")
            FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
            FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, TaxPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, TaxAmt))) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
            End If
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
            Else
               FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
            End If
        Case TaxAmt
            FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(TxtGrid(0).TEXT), "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, TaxPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, TaxAmt))) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
            End If
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
            Else
               FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
            End If
        Case TaxPer
            FGrid.TextMatrix(FGrid.Row, TaxPer) = TxtGrid(0).TEXT
            FGrid.TextMatrix(FGrid.Row, TaxAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
            Else
               FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
            End If
        Case SurchAmt
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(TxtGrid(0).TEXT), "0.00")
            If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, SurchPer) = ""
            Else
               FGrid.TextMatrix(FGrid.Row, SurchPer) = Format((100 * Val(FGrid.TextMatrix(FGrid.Row, SurchAmt))) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)), "0.00")
            End If
        Case SurchPer
            FGrid.TextMatrix(FGrid.Row, SurchPer) = TxtGrid(0).TEXT
            FGrid.TextMatrix(FGrid.Row, SurchAmt) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt)) * Val(FGrid.TextMatrix(FGrid.Row, SurchPer)) / 100, "0.00")
    End Select
    Amt_Cal1
    Amt_Cal
    TxtGridLeave = True
End Function

Private Function ChkDuplicate() As Boolean
Dim I As Integer
Dim X As String, Y As String, addstring As String
Dim Col1 As Byte, Col2 As Byte

    Select Case FGrid.Col
    Case Model
        Col2 = Model
        Col1 = Taxable
    Case Taxable
        Col1 = Model
        Col2 = Taxable
    End Select
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))))
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

Private Sub Fill_Data()
Dim I As Integer

    CustFlag = False
    Txt(NPrefix) = IIf(IsNull(RsParty!NPrefix), "", RsParty!NPrefix)
    Txt(NSuffix) = IIf(IsNull(RsParty!NSuffix), "", RsParty!NSuffix)
    Txt(FPrefix) = IIf(IsNull(RsParty!FPrefix), "", RsParty!FPrefix)
    Txt(fname) = IIf(IsNull(RsParty!fname), "", RsParty!fname)
    Txt(Religion).TEXT = FxReligion(IIf(IsNull(RsParty!Religion), 0, RsParty!Religion))
    If Not IsNull(RsParty!Call_Status) Then
        Select Case RsParty!Call_Status
            Case 0
               Txt(Call_Status) = "Cold"
            Case 1
               Txt(Call_Status) = "Warm"
            Case 2
                Txt(Call_Status) = "Hot"
        End Select
    End If
    
    Txt(Add1) = IIf(IsNull(RsParty!Add1), "", RsParty!Add1)
    Txt(Add2) = IIf(IsNull(RsParty!Add2), "", RsParty!Add2)
    Txt(Add3) = IIf(IsNull(RsParty!Add3), "", RsParty!Add3)
    Txt(ConPerson) = IIf(IsNull(RsParty!ConPerson), "", RsParty!ConPerson)
    Txt(Profession).Tag = IIf(IsNull(RsParty!Profession), "", RsParty!Profession)
    If Txt(Profession).Tag <> "" And GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").RecordCount > 0 Then
        Txt(Profession).TEXT = GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").Fields(0).Value
    Else
        Txt(Profession).TEXT = ""
    End If
    Txt(City).Tag = IIf(IsNull(RsParty!CityCode), "", RsParty!CityCode)
    If Txt(City).Tag <> "" And GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").RecordCount > 0 Then
        Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
    Else
        Txt(City).TEXT = ""
    End If
    Txt(Area).Tag = IIf(IsNull(RsParty!Area), "", RsParty!Area)
    If Txt(Area).Tag <> "" And GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").RecordCount > 0 Then
        Txt(Area).TEXT = GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").Fields(0).Value
    Else
        Txt(Area).TEXT = ""
    End If
    Txt(REP_CODE).Tag = IIf(IsNull(RsParty!REP_CODE), "", RsParty!REP_CODE)
    If Txt(REP_CODE).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").RecordCount > 0 Then
        Txt(REP_CODE).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").Fields(0).Value
    Else
        Txt(REP_CODE).TEXT = ""
    End If

    Txt(REF_CODE).Tag = IIf(IsNull(RsParty!REF_CODE), "", RsParty!REF_CODE)
    If Txt(REF_CODE).Tag <> "" And GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").RecordCount > 0 Then
        Txt(REF_CODE).TEXT = GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").Fields(0).Value
    Else
        Txt(REF_CODE).TEXT = ""
    End If
    
    Txt(Pin) = IIf(IsNull(RsParty!Pin), "", RsParty!Pin)
    Txt(STD) = IIf(IsNull(RsParty!STD), "", RsParty!STD)
    Txt(PhoneOff) = IIf(IsNull(RsParty!PhoneOff), "", RsParty!PhoneOff)
    Txt(PhoneResi) = IIf(IsNull(RsParty!PhoneResi), "", RsParty!PhoneResi)
    Txt(Mobile) = IIf(IsNull(RsParty!Mobile), "", RsParty!Mobile)
    Txt(EMail) = IIf(IsNull(RsParty!EMail), "", RsParty!EMail)
    Txt(FAx) = IIf(IsNull(RsParty!FAx), "", RsParty!FAx)
    Txt(Govt_YN) = IIf(RsParty!Govt_YN = 1, "Yes", "No")
    Txt(FirstVeh_YN) = IIf(RsParty!FirstVeh_YN = 1, "Yes", "No")
    For I = 5 To 24
        Txt(I).Enabled = False
    Next
    Txt(34).Enabled = False
    Txt(NSuffix).Enabled = True
End Sub

Private Sub CreateNewParty(Index As Integer, Optional SuffixCall As Boolean)
 Dim VNo As Long
 Dim I As Integer
    If GCn.Execute("select Cust_Code from ProspectiveCust where site_code = '" & PubSiteCode & "'").RecordCount > 0 Then
        VNo = GCn.Execute("select MAX(right(cust_CODE,7)) from ProspectiveCust where site_code = '" & PubSiteCode & "'").Fields(0).Value + 1
    Else
        VNo = 1
    End If
    Txt(Index).Tag = PubSiteCode + Space(7 - Len(CStr(VNo))) + CStr(VNo)
    CustFlag = True
    For I = 5 To 24
        Txt(I).Enabled = True
        If SuffixCall = False Then
            Txt(I).TEXT = ""
        End If
    Next
    Txt(STD).Enabled = True
    If SuffixCall = False Then
        Txt(STD).TEXT = ""
    End If
    Txt(Govt_YN) = "No"
    Txt(Call_Status) = "Cold"
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)

    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    Select Case FGrid1.Col
        Case ADType
            ListArray = Array("Addition", "Deletion", "Shortage")
            Set mListItem = ListView_Items(ListView, txtgrid1, 0, ListArray, 3)
         Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = "" Then Exit Sub
            If FGrid1.TextMatrix(FGrid1.Row, ADItem) <> RsADItem!Code Then
                RsADItem.MoveFirst
                RsADItem.FIND "code ='" & FGrid1.TextMatrix(FGrid1.Row, ADItemCode) & "'"
            End If
         Case Rate1, Qty1
'                         SendKeys "{HOME}+{END}"
     End Select

End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyEscape Then
                txtgrid1(0).TEXT = txtgrid1(0).Tag
                TxtGrid1_KeyUp Index, KeyCode, Shift
                txtgrid1(0).Visible = False
                Grid_Hide
                FGrid1.SetFocus
                Exit Sub
            End If
            Select Case FGrid1.Col
                Case ADType
                    ListView_KeyDown FrmList, ListView, txtgrid1, 0, KeyCode, Shift, txtgrid1(0).left, (txtgrid1(0).top + txtgrid1(0).height + 25), txtgrid1(0).width, 900
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave1 = True Then
                             GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, Rate1
                        End If
                    End If
                Case ADItem    '1
                    DGridTxtKeyDown DGADItem, txtgrid1, Index, RsADItem, KeyCode, True, 1, frmVehAMDMast, "frmVehAMDMast"
                    If KeyCode = vbKeyReturn Then
                            If TxtGridLeave1 = True Then
                                 GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, Rate1
                            End If
                    End If

                Case Qty1, Rate1
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave1 = True Then
                             GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, Rate1
                        End If
                End If
                End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case FGrid1.Col
    Case ADItem
        If DGADItem.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsADItem, KeyAscii, "name"
    Case Rate1
        Call NumPress(txtgrid1(Index), KeyAscii, 8, 2)
    Case Qty1
        Call NumPress(txtgrid1(Index), KeyAscii, 5, 3)
End Select
End Sub


Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
        Select Case FGrid1.Col
            Case ADType
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                ListView_KeyUp ListView, txtgrid1, 0, KeyCode, mListItem
            Case ADItem
                If KeyCode <> 13 And DGADItem.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsADItem, KeyCode, "name", True
            Case Qty1
                FGrid1.TextMatrix(FGrid1.Row, Qty1) = Format(Val(txtgrid1(Index).TEXT), "0.000")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal
            Case Rate1
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(Val(txtgrid1(Index).TEXT), "0.00")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal

        End Select
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave1(Optional Index As Integer) As Boolean
Dim j As Integer
Dim GridCol As Byte
GridCol = FGrid1.Col
Select Case GridCol
        Case ADType
            If txtgrid1(0).TEXT <> "" Then txtgrid1(0).TEXT = ListView.SelectedItem.TEXT
            FGrid1.TextMatrix(FGrid1.Row, ADType) = txtgrid1(0).TEXT
            If FGrid1.TextMatrix(FGrid1.Row, ADType) = "Shortage" Then
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = "0.00"
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = "0.00"
            End If
        Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or txtgrid1(0).TEXT = "" Then
                FGrid1.TextMatrix(FGrid1.Row, ADItem) = ""
                FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = ""
            Else
                FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = RsADItem!Code
                FGrid1.TextMatrix(FGrid1.Row, ADItem) = RsADItem!Name
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(IIf(IsNull(RsADItem!Rate), 0, RsADItem!Rate), "0.00")
            End If
            If FGrid1.TextMatrix(FGrid1.Rows - 1, 1) <> "" Then FGrid1.AddItem FGrid1.Rows
        Case Qty1
                FGrid1.TextMatrix(FGrid1.Row, Qty1) = Format(Val(txtgrid1(0).TEXT), "0.000")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal
        Case Rate1
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(Val(txtgrid1(0).TEXT), "0.00")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal
End Select
    TxtGridLeave1 = True
End Function

Private Sub FGrid1_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub
Private Sub FGrid1_DblClick()
FGrid1_KeyPress (vbKeyReturn)
TAddMode = False
End Sub

Private Sub FGrid1_EnterCell()
'FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    If FGrid1.BackColorSel = BackColorSelLeave Then FGrid1.Col = 1
    FGrid1.BackColorSel = BackColorSelEnter
    txtgrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
    FGrid1.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
    FGrid1.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid1.Col
        Case ADType
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
        Case Qty1
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "0.000"
            FGrid1.TextMatrix(FGrid1.Row, Amt1) = "0.00"
        Case Rate1
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "0.00"
            FGrid1.TextMatrix(FGrid1.Row, Amt1) = "0.00"
    End Select
    Amt_Cal
End If
KeyCode = 0
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid1.Row > 6 Then MsgBox "You can't take more than three Items in a Quotation", vbInformation: FGrid1.SetFocus:   Exit Sub
Select Case FGrid1.Col
    Case ADItem                                ', ADType
       Call Get_Text(Me, FGrid1, txtgrid1, 0, False, KeyAscii)
    Case Amt1
        FGrid1_LeaveCell
        FGrid1.Col = FGrid1.Col + 1
        FGrid1_EnterCell
        FGrid1.SetFocus
    Case Qty1, Rate1
       Call Get_Text(Me, FGrid1, txtgrid1, 0, True, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid1.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid1.Rows > 2 Then
                FGrid1.RemoveItem (FGrid1.Row)
            Else
                FGrid1.Rows = 1
                FGrid1.AddItem FGrid1.Rows
                FGrid1.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid1.Rows - 1
            FGrid1.TextMatrix(I, 0) = I
         Next
        Amt_Cal
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If

FGrid1.SetFocus
End If
Exit Sub
End Sub


Private Sub FGrid1_Scroll()
txtgrid1(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid1_LeaveCell()
    FGrid1.CellBackColor = CellBackColLeave
'    fgrid1.CellForeColor = CellForeColLeave
End Sub
Private Sub FGrid1_LostFocus()
    If txtgrid1(0).Visible = False Then FGrid1.BackColorSel = BackColorSelLeave
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
    FGrid1.CellBackColor = CellBackColLeave
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
GSQL = "SELECT PC.NPrefix,PC.Name,PC.Nsuffix,PC.Add1,PC.Add2,PC.Add3,City.CityName,PC.PhoneOff, PC.PhoneResi," & _
    "M.Model_Desc,M.Model_Desc1,M.Model_Desc2,M.TYRES,M.RIMS,M.WHEELBASE," & _
    "VQ1.MODEL, VQ1.QTY, VQ1.RATE, VQ1.RSO_WORK, VQ1.TaxableRate_YN, VQ1.TAX_Per, VQ1.TAX_AMT, VQ1.SUR_AMT, VQ1.SURCHARGE_Per, VQ1.AMOUNT AS ModelAmt," & _
    "VQ.DocId,VQ.V_Date,VQ.Party_Code,VQ.CityCode,VQ.Call_Status,VQ.AREA,VQ.REF_CODE,VQ.REP_CODE,VQ.Profession,VQ.PURPOSE,VQ.FIN_YN," & _
    "VQ.FB_CODE,VQ.GOVT_YN,VQ.FirstVeh_YN,VQ.AMOUNT,VQ.RoundOff_YN,VQ.NARR1,VQ.NARR2,VQ.Printed_YN,VQ.INTD_USE,VQ.DEL_DATE,VQ.U_Name,VQ.U_EntDt,VQ.U_AE,Reffered.RefName,M.Sales_Desc, CF.FinName As Financer_Name  " & _
    "FROM ((((((Veh_Quot as VQ left JOIN Veh_Quot1 as VQ1 ON VQ.DocId = VQ1.DocId)" & _
    "LEFT JOIN ProspectiveCust as PC ON VQ.Party_Code = PC.cust_Code) " & _
    "LEFT JOIN City ON PC.CityCode = City.CityCode) " & _
    "LEFT JOIN Model as M ON VQ1.MODEL = M.MODEL) " & _
    "LEFT JOIN Reffered ON VQ.REF_CODE = Reffered.RefCode)" & _
    "LEFT JOIN ContractFinance CF ON VQ.FB_CODE = CF.FinCode)" & _
    "where VQ.DocID = '" & Master!SearchCode & "'"

Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "VehQuot", "VehQuot")
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(GSQL)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "VehQuot", "VehQuot")
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
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset
Dim I As Integer, Cnt As Integer, Foot1$, Foot2$, Foot3$, Foot4$
Dim Foot5$, Foot6$, Foot7$, Foot8$, Foot9$
Dim RST1 As ADODB.Recordset, j As Integer, Footer$
Dim Rst2 As ADODB.Recordset
On Error GoTo ERRORHANDLER
  
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    GSQL = "SELECT " & Rst.RecordCount & " as Cnt, Veh_AMDModel.Prod_Name,Veh_Quot2.QTY  as ADQty,Veh_Quot2.RATE as ADrate, Veh_Quot2.Trn_Type " & _
    "FROM veh_Quot2 LEFT JOIN Veh_AMDModel ON Veh_Quot2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    "where veh_quot2.docid = '" & Master!SearchCode & "'"
    
    Set RstSub1 = New Recordset
    RstSub1.CursorLocation = adUseClient
    RstSub1.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

   'Recordset is made for subreport2
            
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    CreateFieldDefFile RstSub1, PubRepoPath + "\" & mRepName & "1.ttx", True
        
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    Footer = XNull(GCn.Execute("select VehQuotFooter from Syctrl").Fields(0).Value)
    j = 1
    Cnt = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Select Case Cnt
            Case 1
                Foot1 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            Case 2
                Foot2 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            Case 3
                Foot3 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            Case 4
                Foot4 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            Case 5
                Foot5 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            Case 6
                Foot6 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            Case 7
                Foot7 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
             End Select
            Cnt = Cnt + 1
            j = I + 1
        End If
    Next
    
    Set RST1 = New ADODB.Recordset
    RST1.CursorLocation = adUseClient
    RST1.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecSpeciality & "'"
            Case UCase("LST")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecGram & "'"
            Case UCase("AmtPrefix")
                rpt.FormulaFields(I).TEXT = "'" & PubAmountPrefix & "'"
            Case UCase("TitleType")
                rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            Case UCase("JuriCity")
                rpt.FormulaFields(I).TEXT = "'" & XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value) & "'"
            Case UCase("SubRep")
                rpt.FormulaFields(I).TEXT = "" & IIf(RstSub1.RecordCount > 0, 1, 0) & ""
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
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstSub1
        
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
                    Case UCase("TitleType")
                        rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "' + '" & IIf(GCn.Execute("select Printed_YN from veh_quot").Fields(0).Value = 0, "", " (Duplicate)") & "'"
                End Select
            Next
            rpt.PrintOut False
'            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
'                GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!searchcode & "'"
'            End If
            Set Rst = Nothing
            Set RST1 = Nothing
            Set rpt = Nothing
        Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
            Set Rst = Nothing
            Set RST1 = Nothing
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
If rpt.PrinterName <> "" Then LblPrinter.CAPTION = rpt.PrinterName
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
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstQuot As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, PrintStr$
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double
    
    Set RstQuot = GCn.Execute(mQry)
    If RstQuot.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
        
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
        
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehQuotFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
  
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 31
    mFooter = mFooter + FooterCnt
    
    'Header
    mDocStr = "Vehicle Quotation"
    mDupStr = ""
    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
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
        If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
            Print #1, PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignLeft)
            mHeader = mHeader + 1
        Else
            Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignRight)
            mHeader = mHeader + 1
        End If
         mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth) & mChr18 & mEmph
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        
        Print #1, mChr18 & mEmph & PSTR("Party : " & RstQuot!NPrefix & " " & RstQuot!Name & " " & RstQuot!NSuffix, 50) & "Financer : " & XNull(RstQuot!Financer_Name)
        mHeader = mHeader + 1
        Print #1, PSTR((XNull(RstQuot!Add1) & IIf(XNull(RstQuot!Add1) = "" Or XNull(RstQuot!Add2) = "", "", ",") & XNull(RstQuot!Add2)), 50) & "Quotation No.  : " & PSTR(PrinID(RstQuot!DocID), 18, , AlignLeft) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstQuot!Add3), 50) & mEmph & "Quotation Date : " & RstQuot!V_DATE & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(RstQuot!CityName, 40) & Space(10) & "Reffered By : " & RstQuot!refname
        mHeader = mHeader + 1
        Print #1, PSTR((IIf(XNull(RstQuot!PhoneResi) <> "" Or XNull(RstQuot!PhoneOff) <> "", " Phone : ", "") & XNull(RstQuot!PhoneOff) & IIf(XNull(RstQuot!PhoneOff) <> "", "(0)", "") & XNull(RstQuot!PhoneOff) & IIf(XNull(RstQuot!PhoneResi) <> "", "(R)", "")), 50) & "Expected Del. Date : " & IIf(IsNull(RstQuot!DEL_DATE), "", Format((RstQuot!DEL_DATE), "dd/MMM/yyyy"))
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        mHeader = mHeader + 1
        
        Print #1, PSTR("Sr", 3) & PSTR("Model Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, , AlignRight) & " " & PSTR("<----Tax---- >", 13) & " " & PSTR("<-Sur.On Tax- >", 14) & " " & PSTR("Amount", 9, , AlignRight)
        mHeader = mHeader + 1
        Print #1, PSTR("No.", 3) & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        
        Cnt = 1
        Do Until RstQuot.EOF
            Print #1, Cnt & ". " & mDoub & PSTR(RstQuot!Model, 22) & mDoub1 & " " & PSTR(RstQuot!Qty, 3) & " " & PSTR(RstQuot!Rate, 11, 2) & " " & PSTR(RstQuot!Tax_Per, 5, 2) & " " & PSTR(RstQuot!Tax_Amt, 7, 2) & " " & PSTR(RstQuot!surcharge_per, 5, 2) & " " & PSTR(RstQuot!sur_amt, 7, 2) & " " & PSTR(RstQuot!modelAmt, 10, 2)
             mHeader = mHeader + 1
             If RstQuot!Model_Desc <> "" Then
                Print #1, mChr17 & RstQuot!Model_Desc & mChr18
                mHeader = mHeader + 1
             End If
             If RstQuot!Model_Desc1 <> "" Then
                Print #1, mChr17 & PSTR(RstQuot!Model_Desc1, 50) & mChr18
                mHeader = mHeader + 1
             End If
             Print #1, mChr17 & "<WheelBase " & STR(RstQuot!WHEELBASE) & " With " & STR(RstQuot!Tyres) & " Tyres And " & STR(RstQuot!Rims) & " Rims  >" & mChr18
             mHeader = mHeader + 1
             Cnt = Cnt + 1
            RstQuot.MoveNext
        Loop
        
        RstQuot.MoveFirst
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        
        mQry = "SELECT Veh_AMDModel.Prod_Name,Veh_Quot2.QTY  as ADQty , Veh_Quot2.RATE as ADrate, Veh_Quot2.Trn_Type " & _
        "FROM veh_Quot2 LEFT JOIN Veh_AMDModel ON Veh_Quot2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where veh_quot2.docid = '" & Master!SearchCode & "'"
    
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount > 0 Then
            Print #1, mEmph & "Additional Fitments Detail : " & mEmph1 & mDoub
            mHeader = mHeader + 1
            Print #1, PSTR("Item Name", 52) & PSTR("Qty", 6, , AlignRight) & PSTR("Rate", 11, , AlignRight) & PSTR("Amount", 11, , AlignRight) & mDoub1
            mHeader = mHeader + 1
            Do Until Rst.EOF
                Print #1, PSTR(Rst!Prod_Name, 52) & PSTR(Rst!adqty, 6, 2) & PSTR(Rst!adrate, 11, 2) & PSTR(Rst!adqty * Rst!adrate, 11, 2)
                mAmt = mAmt + (Rst!adqty * Rst!adrate)
                mHeader = mHeader + 1
                Rst.MoveNext
            Loop
            mAmt = mAmt * RstQuot.RecordCount
            Print #1, Replace(Space(PageWidth), " ", "-")
            mHeader = mHeader + 1
            Print #1, PSTR("Total Additional Fitments Amount : " & Format(mAmt, "0.00"), PageWidth, , AlignRight)
            mHeader = mHeader + 1
        End If
        
        Do Until mHeader >= PageLength - mFooter
            Print #1, ""
            mHeader = mHeader + 1
        Loop
                
        Print #1, mDoub & PSTR(("Bill Amount : " & Amount_Fill(RstQuot!Amount, PubAmountPrefix)), PageWidth, , AlignRight)
        Print #1, ntow(RstQuot!Amount, "Rupees", "Paise") & mDoub1 & mChr17
        Print #1, ""
        
        Print #1, "Equipment Specification,Price,Excise duty & Sales Tax quoted above are subject to change without notice.The price,excise duty & sales tax"
        Print #1, "prevailing at the time of delivery of vehicle will be charged irrespective of when the order is placed or part or full payment is accepted."
        Print #1, ""
        
        Print #1, "With the acceptance of this offer, the contract shall be construed as having been entered into at " & mJuriCity & "  The court at " & mJuriCity & " alone "
        Print #1, "will have jurisdiction to try any suits in respect of any claim/dispute arising out of the said contract"
        Print #1, ""
        
        Print #1, mChr18 & mEmph & PSTR("For " & PubComp_Name, PageWidth, , AlignRight) & mEmph1
        Print #1, ""
        Print #1, ""
        Print #1, PSTR("Authorised Signatory", PageWidth, , AlignRight)
        Print #1, ""
        Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
    
        Footer = Footer & vbLf
        j = 1
        For I = 1 To Len(Footer)
            If mID(Footer, I, 1) = vbLf Then
                Print #1, RTrim(mID(Footer, j, I - j))
                j = I + 1
            End If
        Next

        Print #1, "" & mChr18 & mEmph
        
    If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
        Print #1, "Note : " & mEmph1 & mChr17
   
        Print #1, "1.This is to inform all our esteemed customers that ANY ADVANCE payments for purchase  of  vehicle made  by them to us are our OWN"
        Print #1, "LIABILITY and our principals essers Telco are in No WAY , Implicitly or explicitly  responsible for any vicarious liability"
        Print #1, "for the refund of advance or delivery of vehicles thereof,as they deal with us on  a  principal to principal basis"
        
        Print #1, "2.Financiers are requested to make the payment of financed vehicle accomponied with a formal letter giving correct"
        Print #1, "full particulars in respect of name and address directly to us. In case payment is handedover to hirer and if it does"
        Print #1, "not reach us, we will not be responsible for any further disputes against the same."
        
        Print #1, "P.S. : In case the delivery of vehicle delayed beyond 15 days from dt. of receipt of payment in our bank, interest @12% p.a."
        Print #1, "will be payable after 15th day till the date of invoicing of TELCO LTD."
    Else
        Print #1, "P.S. : In case the delivery of vehicle delayed beyond 15 days from dt. of receipt "
        Print #1, "of payment in our bank, interest @6% p.a. will be payable after 15th day till the "
        Print #1, "date of invoicing of TATA MOTORS LTD."
    End If
     
        Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
               
        Print #1, mChr17 & RstQuot!U_Name & " " & STR(RstQuot!U_EntDt) & Space(((PageWidth * 1.7) - Len("") - Len(RstQuot!U_Name & " " & STR(RstQuot!U_EntDt))) / 2) & "" & mChr18
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
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub


