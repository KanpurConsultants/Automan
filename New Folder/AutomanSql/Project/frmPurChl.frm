VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPurChl 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Goods Receipts"
   ClientHeight    =   9135
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
   ScaleHeight     =   9135
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
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
      Index           =   28
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   150
      Top             =   6180
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
      Height          =   255
      Index           =   27
      Left            =   6060
      MaxLength       =   40
      TabIndex        =   143
      Top             =   5910
      Width           =   1470
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4830
      Top             =   7710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   0
      TabIndex        =   142
      Top             =   3255
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
      Left            =   -2145
      TabIndex        =   127
      Top             =   7620
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
         Picture         =   "frmPurChl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   137
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
         Picture         =   "frmPurChl.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmPurChl.frx":0678
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
         TabIndex        =   135
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmPurChl.frx":0982
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
         TabIndex        =   134
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmPurChl.frx":0C8C
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGPONo 
      Height          =   2775
      Left            =   -105
      Negotiate       =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   9015
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4895
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
      RowHeight       =   16
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "PO Reg. No."
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
         DataField       =   "Order_Reg_Dt"
         Caption         =   "PO Reg.Date"
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
         DataField       =   "Code"
         Caption         =   "OrderID"
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
         DataField       =   "OurDocNo"
         Caption         =   "PO No. "
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
         DataField       =   "v_date"
         Caption         =   "PO Date"
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
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGod 
      Height          =   2145
      Left            =   6705
      Negotiate       =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   8100
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
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
   Begin MSDataGridLib.DataGrid DGTrans 
      Height          =   4935
      Left            =   10065
      Negotiate       =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   8100
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
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
      RowHeight       =   16
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
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   4935
      Left            =   1725
      Negotiate       =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   8355
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
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
            ColumnWidth     =   3360.189
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGDrAc 
      Height          =   4935
      Left            =   1470
      Negotiate       =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   9075
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
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
         Caption         =   "A/c Name"
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
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   -3465
      Negotiate       =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   9000
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
      Left            =   3270
      TabIndex        =   94
      Top             =   2595
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
         Index           =   47
         Left            =   3765
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
      Height          =   240
      Index           =   26
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   9
      Top             =   915
      Width           =   3540
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   11610
      Negotiate       =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   2460
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
   Begin MSDataGridLib.DataGrid DGOrdPart 
      Height          =   2625
      Left            =   11055
      Negotiate       =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   2385
      Visible         =   0   'False
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   4630
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Srl_No"
         Caption         =   "Srl.No."
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
         DataField       =   "Part_No"
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
      BeginProperty Column02 
         DataField       =   "Part_Name"
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
      BeginProperty Column03 
         DataField       =   "Qty"
         Caption         =   "Ord.Qty"
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
         DataField       =   "Sup_Qty"
         Caption         =   "Sup.Qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PendQty"
         Caption         =   "Pending Qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
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
            Alignment       =   1
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3149.858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1289.764
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   14
      Left            =   6330
      MaxLength       =   15
      TabIndex        =   18
      Top             =   1935
      Width           =   1725
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
      Height          =   240
      Index           =   13
      Left            =   6330
      MaxLength       =   15
      TabIndex        =   16
      Top             =   1680
      Width           =   1725
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
      Height          =   240
      Index           =   7
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1680
      Width           =   1995
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
      Height          =   240
      Index           =   8
      Left            =   3975
      MaxLength       =   10
      TabIndex        =   13
      Text            =   "0123456789"
      Top             =   1680
      Width           =   1110
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
      Height          =   240
      Index           =   1
      Left            =   10290
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1590
      Width           =   1185
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
      Height          =   240
      Index           =   9
      Left            =   9555
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   1920
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
      Height          =   240
      Index           =   25
      Left            =   1545
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1170
      Width           =   1470
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
      Height          =   240
      Index           =   24
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1425
      Width           =   3540
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   10875
      TabIndex        =   76
      Top             =   8010
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   -150
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   810
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
      Height          =   240
      Index           =   4
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   6
      Top             =   405
      Width           =   3540
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
      Height          =   240
      Index           =   3
      Left            =   9555
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1860
      Width           =   1200
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
      Left            =   300
      TabIndex        =   20
      Top             =   3435
      Visible         =   0   'False
      Width           =   1395
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
      Height          =   240
      Index           =   0
      Left            =   9135
      MaxLength       =   21
      TabIndex        =   1
      Top             =   510
      Width           =   2565
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
      Index           =   23
      Left            =   9750
      MaxLength       =   40
      TabIndex        =   29
      Top             =   5610
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
      Height          =   255
      Index           =   22
      Left            =   9750
      MaxLength       =   40
      TabIndex        =   28
      Top             =   5340
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
      Height          =   255
      Index           =   21
      Left            =   6060
      MaxLength       =   40
      TabIndex        =   27
      Top             =   6180
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
      Height          =   255
      Index           =   20
      Left            =   6060
      MaxLength       =   40
      TabIndex        =   26
      Top             =   5640
      Width           =   1470
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
      Index           =   19
      Left            =   6060
      MaxLength       =   40
      TabIndex        =   25
      Top             =   5370
      Width           =   1470
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
      Height          =   255
      Index           =   18
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   24
      Top             =   5910
      Width           =   1470
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
      Height          =   255
      Index           =   17
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   23
      Top             =   5640
      Width           =   1470
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
      Height          =   255
      Index           =   16
      Left            =   2205
      MaxLength       =   40
      TabIndex        =   22
      Top             =   5370
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   1545
      MaxLength       =   30
      TabIndex        =   17
      Top             =   2190
      Width           =   3540
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
      Height          =   240
      Index           =   11
      Left            =   4590
      MaxLength       =   4
      TabIndex        =   15
      Top             =   1935
      Width           =   495
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
      Height          =   240
      Index           =   10
      Left            =   1545
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1935
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
      Height          =   240
      Index           =   15
      Left            =   6330
      MaxLength       =   12
      TabIndex        =   19
      Top             =   2190
      Width           =   1725
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
      Height          =   240
      Index           =   6
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   8
      Top             =   660
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   1545
      MaxLength       =   15
      TabIndex        =   7
      Top             =   660
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
      Height          =   240
      Index           =   2
      Left            =   9555
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1050
      Width           =   1245
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   4935
      Left            =   3765
      Negotiate       =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   8415
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
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
      Caption         =   "Tax Form Help"
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
      Height          =   2550
      Left            =   45
      TabIndex        =   21
      Top             =   2460
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   4498
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   0
      Cols            =   31
      BackColorFixed  =   13300221
      ForeColorFixed  =   128
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16761024
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "SrNo."
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
      _Band(0).Cols   =   31
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total SFC"
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
      Index           =   24
      Left            =   300
      TabIndex        =   151
      Top             =   6195
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   149
      Top             =   0
      Width           =   45
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   330
      TabIndex        =   148
      Top             =   6525
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
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
      Left            =   3600
      TabIndex        =   147
      Top             =   1695
      Width           =   285
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   46
      Left            =   3120
      TabIndex        =   146
      Top             =   675
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax Amount"
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
      Index           =   45
      Left            =   3945
      TabIndex        =   145
      Top             =   5925
      Width           =   1815
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   26
      Left            =   5865
      TabIndex        =   144
      Top             =   5895
      Width           =   45
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6690
      TabIndex        =   141
      Top             =   420
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr A/c Name"
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
      Index           =   6
      Left            =   45
      TabIndex        =   92
      Top             =   915
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permit Form"
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
      Left            =   45
      TabIndex        =   87
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      Height          =   1740
      Left            =   8175
      Top             =   435
      Width           =   3585
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
      Height          =   270
      Left            =   9555
      TabIndex        =   86
      Top             =   1575
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt No."
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
      Index           =   1
      Left            =   8265
      TabIndex        =   85
      Top             =   1605
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   92
      Left            =   9390
      TabIndex        =   84
      Top             =   1605
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
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   8265
      TabIndex        =   83
      Top             =   795
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
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   10155
      TabIndex        =   82
      Top             =   795
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Type"
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
      Index           =   31
      Left            =   8265
      TabIndex        =   81
      Top             =   1335
      Width           =   1095
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   12
      Left            =   9390
      TabIndex        =   80
      Top             =   1335
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Type"
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
      Height          =   225
      Index           =   44
      Left            =   45
      TabIndex        =   79
      Top             =   1170
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type"
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
      Height          =   225
      Index           =   43
      Left            =   45
      TabIndex        =   78
      Top             =   1425
      Width           =   885
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
      Left            =   9015
      TabIndex        =   73
      Top             =   510
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
      Height          =   255
      Index           =   42
      Left            =   8250
      TabIndex        =   72
      Top             =   510
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   22
      Left            =   9555
      TabIndex        =   70
      Top             =   5610
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
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
      Index           =   41
      Left            =   7635
      TabIndex        =   69
      Top             =   5640
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   21
      Left            =   9585
      TabIndex        =   68
      Top             =   5340
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction"
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
      Index           =   40
      Left            =   7665
      TabIndex        =   67
      Top             =   5370
      Width           =   840
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   20
      Left            =   5865
      TabIndex        =   66
      Top             =   6165
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Addition"
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
      Index           =   39
      Left            =   3945
      TabIndex        =   65
      Top             =   6195
      Width           =   660
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   19
      Left            =   5865
      TabIndex        =   64
      Top             =   5625
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Amount"
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
      Index           =   38
      Left            =   3945
      TabIndex        =   63
      Top             =   5655
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   18
      Left            =   5865
      TabIndex        =   62
      Top             =   5355
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Goods Value"
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
      Index           =   37
      Left            =   3930
      TabIndex        =   61
      Top             =   5385
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   17
      Left            =   2070
      TabIndex        =   60
      Top             =   5895
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Order Discount"
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
      Index           =   36
      Left            =   300
      TabIndex        =   59
      Top             =   5925
      Width           =   1695
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   16
      Left            =   2100
      TabIndex        =   58
      Top             =   5625
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
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
      Index           =   35
      Left            =   300
      TabIndex        =   57
      Top             =   5655
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   15
      Left            =   2100
      TabIndex        =   56
      Top             =   5355
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Index           =   34
      Left            =   300
      TabIndex        =   55
      Top             =   5385
      Width           =   1080
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   14
      Left            =   10710
      TabIndex        =   54
      Top             =   5070
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Goods Amount"
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
      Index           =   33
      Left            =   8940
      TabIndex        =   53
      Top             =   5070
      Width           =   1680
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
      Left            =   10965
      TabIndex        =   52
      Top             =   5070
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supply Mode "
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
      Index           =   32
      Left            =   5190
      TabIndex        =   51
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transporter"
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
      Index           =   30
      Left            =   45
      TabIndex        =   50
      Top             =   2190
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Case"
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
      Index           =   29
      Left            =   3105
      TabIndex        =   49
      Top             =   1935
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Case Marking"
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
      Index           =   28
      Left            =   45
      TabIndex        =   48
      Top             =   1935
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GR/ Bilty Date"
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
      Index           =   27
      Left            =   5190
      TabIndex        =   47
      Top             =   2205
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GR/Bilty No."
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
      Left            =   5190
      TabIndex        =   46
      Top             =   1935
      Width           =   975
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
      Left            =   7665
      TabIndex        =   45
      Top             =   5070
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   22
      Left            =   5850
      TabIndex        =   44
      Top             =   5070
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   4
      Left            =   7455
      TabIndex        =   43
      Top             =   5070
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Inv. No."
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
      Index           =   5
      Left            =   45
      TabIndex        =   42
      Top             =   660
      Width           =   1335
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   45
      TabIndex        =   41
      Top             =   405
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   3
      Left            =   4515
      TabIndex        =   40
      Top             =   5070
      Width           =   120
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   2
      Left            =   1920
      TabIndex        =   39
      Top             =   5070
      Width           =   45
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   25
      Left            =   2850
      TabIndex        =   38
      Top             =   5070
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
      Left            =   2220
      TabIndex        =   37
      Top             =   5070
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
      Left            =   4695
      TabIndex        =   36
      Top             =   5070
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   7
      Left            =   300
      TabIndex        =   35
      Top             =   5070
      Width           =   1440
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   1
      Left            =   1560
      TabIndex        =   34
      Top             =   5070
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   91
      Left            =   9390
      TabIndex        =   33
      Top             =   1035
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   93
      Left            =   9390
      TabIndex        =   32
      Top             =   1875
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Credit"
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
      Index           =   0
      Left            =   8280
      TabIndex        =   31
      Top             =   1875
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Date"
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
      Index           =   2
      Left            =   8265
      TabIndex        =   30
      Top             =   1065
      Width           =   1080
   End
End
Attribute VB_Name = "frmPurChl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Const CellBackColLeave As String = &HEDF7FE
'Private Const CellForeColLeave As String = &HFF00FF
'Private Const CellBackColEnter As String = &HCAF1FD   '&HF0D5BF    '&HFFC0C0
Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Private Const ChalVType As String = "SXGR"
Private Const TrfVType As String = "SXGRT"


Dim mCheckNegetiveStockSiteWise As Boolean
Dim RsVno As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim RsDrAc As ADODB.Recordset
Dim rsPONo As ADODB.Recordset
Dim rsGod As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim rsForm31 As ADODB.Recordset
Dim rsTrans As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mDrAcFlag As Boolean
Dim FirmAddFlag As Byte
Dim GridKey As Integer
'Dim Docid As String * 21
Dim mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function

Private Const TxtDocID As Byte = 0
Private Const SerialNo As Byte = 1
Private Const VDate As Byte = 2
Private Const VType As Byte = 3
Private Const Party As Byte = 4
Private Const SuppChlNo As Byte = 5
Private Const SuppChlDate As Byte = 6
Private Const FormType As Byte = 24
Private Const PermitType As Byte = 7
Private Const FormNo As Byte = 8
Private Const ChlType As Byte = 9
Private Const CaseMark As Byte = 10
Private Const CaseNo As Byte = 11
Private Const Transport As Byte = 12
Private Const LC As Byte = 25
Private Const SupplyMode As Byte = 13
Private Const GrNo      As Byte = 14
Private Const GrDate    As Byte = 15
Private Const TOTAmt    As Byte = 16
Private Const TotDis    As Byte = 17
Private Const TotOrdDis As Byte = 18
Private Const TotGoods  As Byte = 19
Private Const TaxAmt    As Byte = 20
Private Const Addition  As Byte = 21
Private Const Deduction As Byte = 22
Private Const NetAmt    As Byte = 23
Private Const DrAcCode  As Byte = 26
Private Const SatAmt    As Byte = 27
Private Const SFCAmt    As Byte = 28

' Col Declaration
Private Const PONo As Byte = 1
Private Const PNo As Byte = 2
Private Const Unit As Byte = 3
Private Const MRP As Byte = 4
Private Const Taxable As Byte = 5
Private Const DQty As Byte = 6
Private Const PQty As Byte = 7
Private Const FRate  As Byte = 8
Private Const Amt  As Byte = 9
Private Const DisPer  As Byte = 10
Private Const DisRs  As Byte = 11
Private Const DisOrd  As Byte = 12
Private Const DisOrdRs  As Byte = 13
Private Const SFCPer As Byte = 14
Private Const SFCAmt1 As Byte = 15

Private Const TaxPer As Byte = 16
Private Const TaxAmt1 As Byte = 17
Private Const SatPer As Byte = 18
Private Const SatAmt1 As Byte = 19
Private Const ItemVal As Byte = 20
Private Const God As Byte = 21 '30
Private Const Godown As Byte = 22
Private Const NDP As Byte = 23
Private Const PartSrlNo As Byte = 24         ' Part Serial No
Private Const PName As Byte = 25
Private Const LName As Byte = 26
Private Const MRPStkTB As Byte = 27
Private Const MRPStkTP As Byte = 28
Private Const TBStk As Byte = 29
Private Const TPStk As Byte = 30
Private Const TBRate As Byte = 31
Private Const TPRate As Byte = 32
Private Const Bin As Byte = 33
Private Const LastRate As Byte = 34
Private Const HPRate As Byte = 35
Private Const LPRate As Byte = 36
Private Const PONOCode As Byte = 37
Private Const POSrlNo As Byte = 38
Private Const PartGrade As Byte = 39
Private Const EffectDate As Byte = 40
Private Const MRPRate As Byte = 41

Private Const FromVno As Byte = 0
Private Const ToVno As Byte = 1

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim rsTaxPer As ADODB.Recordset

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Dim mSatYn As Boolean

'Private Sub CmdBrowse_Click()
'CommonDialog1.ShowOpen
'Dim TmpArr() As String
'Dim I As Double
'Dim varScrptObj As New Scripting.FileSystemObject, varTxtstrm As Scripting.TextStream, varTxtstrm1 As Scripting.TextStream
'
'Set varTxtstrm = varScrptObj.OpenTextFile(CommonDialog1.FileName)
'While Not varTxtstrm.AtEndOfStream = True
'    TmpArr = Split(varTxtstrm.ReadLine, ",")
'    FGrid.TextMatrix(I, PNo) = TmpArr(0)
'    FGrid.TextMatrix(I, Taxable) = "Yes"
'    FGrid.TextMatrix(I, DQty) = TmpArr(1)
'    FGrid.TextMatrix(I, PQty) = TmpArr(1)
'    FGrid.TextMatrix(I, FRate) = TmpArr(2)
'    I = I + 1
'Wend
'End Sub
Private Sub DGDrAc_Click()
    If RsDrAc.RecordCount > 0 Then
        txt(DrAcCode).TEXT = RsDrAc!Name
        txt(DrAcCode).Tag = RsDrAc!Code
    End If
    txt(DrAcCode).SetFocus
    DGDrAc.Visible = False
End Sub

Private Sub DGOrdPart_dblClick()
FGrid.TextMatrix(FGrid.Row, POSrlNo) = GRs!Srl_No
Set GRs = Nothing
FGrid.SetFocus
DGOrdPart.Visible = False
End Sub
' Col Declaration
' |Part No.1|Part Name2|Unit 3|PO No 4|Taxable 5|MRP6|Qty(Doc)7|Qty(Phy)8|NDP 9 |Amount 10
' |Dis %11  |Ord Dis %12|Amount 13|Loal Name 14|Curr Stk Qty 15|MRP Qty 16 |Taxable Qty 17|TaxPaid Qty 18|Taxable Rate 19|TaxPaid Rate 20|Bin Location 21|Last Purch Rate 22|High Purch Rate 23|Low Purch Rate 24

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
    Call Ini_Pub
    
    Label3(4) = PubForm31Caption
    'Label3(24) = PubForm31Caption & " No."
    mVType = ChalVType
    txt(VDate).Tag = PubLoginDate
    
    
     Dim sitecond As String
     sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
'    Master.Open "select DocID as searchcode,Sp_Purch.* from Sp_Purch  where v_type in ('" & ChalVType & "','" & TrfVType & "')", GCn, adOpenDynamic, adLockOptimistic
    If PubMoveRecYn Then
        Master.Open "select DocID as SearchCode, DocID from Sp_Purch where left(DocId,1) = '" & PubDivCode & "' " & sitecond & " and v_type in ('" & ChalVType & "','" & TrfVType & "') and v_date<=" & ConvertDate(Format(PubEndDate, "dd-mm-yyyy")) & " Order by V_Date Desc, DocId desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("select Top 1 DocID as SearchCode, DocID from Sp_Purch where left(DocId,1) = '" & PubDivCode & "' " & sitecond & " and v_type in ('" & ChalVType & "','" & TrfVType & "') and v_date<=" & ConvertDate(Format(PubEndDate, "dd-mm-yyyy")) & " Order by V_Date Desc, DocId desc")
    End If
    
    Set DGPart.DataSource = RsPart
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct V_No as code from SP_Purch where left(DocId,1) = '" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno
    
    Set rsForm = New ADODB.Recordset
    With rsForm
        .CursorLocation = adUseClient
        .Open "SELECT TaxForms.Form_Code as code,TaxForms.form_Desc as name FROM TaxForms where Spare_YN = 1 and trn_Type = 'Purchase' order by  TaxForms.form_Desc", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGForm.DataSource = rsForm
        
    Set rsForm31 = New ADODB.Recordset
    With rsForm31
        .CursorLocation = adUseClient
        .Open "SELECT TaxForms.Form_Code as code,TaxForms.form_Desc as name FROM TaxForms where Spare_YN = 1 and trn_Type = 'Permit' order by  TaxForms.form_Desc", GCn, adOpenDynamic, adLockOptimistic
    End With
    
    Set rsPONo = New ADODB.Recordset
    With rsPONo
        .CursorLocation = adUseClient
        .Open "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, Right(OrderID,13) as OurDocNo,V_Date From SP_Order Where left(OrderId,1) = '" & PubDivCode & "' and left(Order_Type,4)='S_PO' Order By OrderID", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGPONo.DataSource = rsPONo
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME,Party_Type from SubGroup Where firmCode = '" & PubFirmCode & "' and Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type,SubGroup.Add1,City.CityName from ((SubGroup " & _
        "left join City on City.CityCode = SubGroup.CityCode) " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode) " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup " & _
        "left join " & FaTable("AcGroup") & "  on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) not in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        "order by SubGroup.name"
    Set RsDrAc = New ADODB.Recordset
    RsDrAc.CursorLocation = adUseClient
    RsDrAc.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGDrAc.DataSource = RsDrAc
    
    Set rsGod = New ADODB.Recordset
    rsGod.CursorLocation = adUseClient
    rsGod.Open "select god_code as code,god_name as name from godown where Appli_For=0 order by god_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGod.DataSource = rsGod
    
    Set rsTrans = New ADODB.Recordset
    rsTrans.CursorLocation = adUseClient
    rsTrans.Open "select distinct transport as name from  sp_Purch  where  transport <>   '' order by transport", GCn, adOpenDynamic, adLockOptimistic
    Set DGTrans.DataSource = rsTrans
    
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    If PubVATYN = 1 Then
        Label3(38).CAPTION = "V A T"
    End If
    
    Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsParty = Nothing
Set rsPONo = Nothing
Set rsGod = Nothing
Set rsForm = Nothing
Set RsVno = Nothing
Set Master = Nothing
Set rsTrans = Nothing
Set mListItem = Nothing
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
    mSatYn = IIf(PubSatYn = 1, True, False)
    DispText_Vat
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    txt(TxtDocID).Enabled = False
    mPartyType = 0
    txt(VDate) = txt(VDate).Tag
    txt(VDate).SetFocus
    FGrid.Col = PONo
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$
If Master.RecordCount > 0 Then
    If GCn.Execute("select count(*) from sp_Purch where Invoice_DocId = '" & Master!DocID & "'").Fields(0).Value > 0 Then
         MsgBox "Purchase Bill Exists of this Purchase Challan, " & vbCrLf & "Can't Delete the Reocord", vbInformation, "Validation"
         Exit Sub
    End If
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
        If left(txt(ChlType).TEXT, 1) = "S" Then
            'Unpost Ledger a/c
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, txt(TxtDocID))
            If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
            'Unposting of Ledger completed
        End If
        UpdateOrderQty
        UpdStkTableToTable Master!DocID, "-", "R"
        GCn.Execute ("delete from Sp_Stock where docId = '" & Master!DocID & "'")
        If GCn.Execute("Select CancelYN from SP_Purch where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
            GCn.Execute ("delete from Sp_Purch where docId = '" & Master!DocID & "'")
        Else
            GCn.Execute ("update sp_purch set " & _
                " CancelYN=1,RoadPermit_Formcode='',RoadPermit_no='',Tot_Amt = 0,Tot_Disc_Amt= 0,Tot_Ord_DiscAmt=0," & _
                " Tot_Goods_Value=0,Tax_Amt=0,Addition=0,Deduction=0,NET_AMT =0,U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E', DrAc_Code='" & txt(DrAcCode).Tag & _
                "' where docid = '" & txt(TxtDocID) & "'")
        End If
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
End If
eloop1:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    If GCn.Execute("select count(*) from sp_Purch where Invoice_DocId = '" & Master!DocID & "'").Fields(0).Value > 0 Then
         MsgBox "Purchase Bill Exists of this Purchase Challan, " & vbCrLf & "Can't Edit the Reocord", vbInformation, "Validation"
         Exit Sub
    End If
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    FGrid.AddItem FGrid.Rows
    'modi lps 19-02-2004 at Cuttack
    'Txt(Party).SetFocus
    txt(VType).Enabled = True
    txt(VType).SetFocus
    'eof modi lps
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
    rsPONo.Requery
    rsGod.Requery
    rsForm.Requery
    rsTrans.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean
Dim Rst As ADODB.Recordset, DocIdHlp As String, mGridFilled As Boolean
On Error GoTo errlbl
Dim mItemVal As Double, mItemQty As Double, mTotDiffAmt As Double
Dim mDiffPerc As Single, mDiffAmt As Double
Dim mDiffPosted As Double, LastI As Integer
Dim LedgAry(1) As LedgRec, mNarr$, mResult As Byte
Dim OilAmtMrpTP As Double, OilAmtMrpTB As Double, OilAmtTP As Double, OilAmtTB As Double
Dim SprAmtMrpTP As Double, SprAmtMrpTB As Double, SprAmtTP As Double, SprAmtTB As Double
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(txt(VDate), Label3(2)) = False Then Exit Sub
    If IsValid(txt(ChlType), Label3(31)) = False Then Exit Sub
    If IsValid(txt(VType), "Cash Credit") = False Then Exit Sub
    If IsValid(txt(SerialNo), Label3(1)) = False Then Exit Sub
    If IsValid(txt(Party), Label3(3)) = False Then Exit Sub
    If IsValid(txt(LC), Label3(44)) = False Then Exit Sub
    If txt(SuppChlDate) <> "" Then
        If CDate(RetDate(txt(SuppChlDate))) > CDate(RetDate(txt(VDate))) Then
            MsgBox "Supplier Document Date  > Bill Date", vbOKOnly, "Validation": txt(SuppChlDate).SetFocus: Exit Sub
        End If
    End If
    
    If GCn.Execute("Select * from SP_Purch where Party_Code='" & txt(Party).Tag & "' and Party_Doc_No='" & txt(SuppChlNo).TEXT & "'").RecordCount > 0 And TopCtrl1.TopText2 = "Add" Then
        MsgBox " This Supplier Document No for the Party Already Exists.", vbInformation + vbOKOnly, "Validation": txt(SuppChlNo).SetFocus: Exit Sub
    End If
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" Then
            If FGrid.TextMatrix(I, MRP) = "" Then MsgBox "Fill MRP Yes/No in Row No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = MRP: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Taxable) = "" Then MsgBox "Fill Taxable Yes/No in Row No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Taxable: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, PQty)) = 0 Then MsgBox "Fill Quantity in Row No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = PQty: FGrid.SetFocus: Exit Sub
            'Check Removed By Nra For FOC Item Entry
'            If Val(FGrid.TextMatrix(I, FRate)) = 0 Then
''                If PubULabel <> "Y" Then
'                    MsgBox "Please Specify Rate in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = FRate: FGrid.SetFocus: Exit Sub
''                End If
'            End If
            If Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrd)) > Val(FGrid.TextMatrix(I, Amt)) Then
                MsgBox "Discount is greater than Item Value in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = FRate: FGrid.SetFocus: Exit Sub
            End If
            If FGrid.TextMatrix(I, God) = "" Then MsgBox "Fill Godown in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Godown: FGrid.SetFocus: Exit Sub
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = PNo: FGrid.SetFocus: Exit Sub
    'Calculating Landed Rate for Each Part
    mTotDiffAmt = Val(txt(TaxAmt)) + Val(txt(Addition)) + Val(txt(Deduction))
    mDiffPosted = 0
    If mTotDiffAmt <> 0 Then
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, PNo) <> "" Then
                mItemVal = Val(FGrid.TextMatrix(I, ItemVal))
                mItemQty = Val(FGrid.TextMatrix(I, PQty))
                mDiffPerc = Round((mItemVal * 100) / Val(txt(TotGoods)), 2)
                mDiffAmt = Round(mTotDiffAmt * mDiffPerc / 100, 2)
                mDiffPosted = mDiffPosted + mDiffAmt
                FGrid.TextMatrix(I, NDP) = Round((mItemVal + mDiffAmt) / mItemQty, 2)
                LastI = I
            End If
        Next
    End If
    If mTotDiffAmt - mDiffPosted <> 0 Then
        mItemVal = Val(FGrid.TextMatrix(LastI, ItemVal))
        mItemQty = Val(FGrid.TextMatrix(LastI, PQty))
        FGrid.TextMatrix(LastI, NDP) = Round((mItemVal + mTotDiffAmt - mDiffPosted) / mItemQty, 2)
    End If
    'EOF Landed Rate Calculation
    
'modishekhar
'calculation for distinguish spr and oil amt
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" Then
            If FGrid.TextMatrix(I, PartGrade) = PubPartGrade_Lub Then
                If FGrid.TextMatrix(I, MRP) = "Yes" Then
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        OilAmtMrpTB = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtMrpTB
                    Else
                        OilAmtMrpTP = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtMrpTP
                    End If
                Else
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        OilAmtTB = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtTB
                    Else
                        OilAmtTP = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtTP
                    End If
                End If
            Else
                If FGrid.TextMatrix(I, MRP) = "Yes" Then
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        SprAmtMrpTB = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtMrpTB
                    Else
                        SprAmtMrpTP = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtMrpTP
                    End If
                Else
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        SprAmtTB = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtTB
                    Else
                        SprAmtTP = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtTP
                    End If
                End If
            End If
        End If
    Next
    RemoveTxtNull
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        txt(TxtDocID).Tag = txt(TxtDocID)
'        If GCn.Execute("select count(*) from sp_purch where Left(DocID,1)='" & PubDivCode & "' And V_Type = '" & mVType & "' And V_No=" & Val(txt(SerialNo)) & "").Fields(0) > 0 Then
'            If VoucherEditFlag Then 'And txt(BookNo).Visible Then
'                MsgBox "Challan No. already exists, Retry", vbCritical, "Validation Error"
'                txt(SerialNo).SetFocus
'                GoTo errlbl
'            Else
'                txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
'                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) Then
'                    MsgBox "Challan No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
'                    GoTo errlbl
'                End If
'            End If
'        End If
 If GCn.Execute("select count(*) from sp_purch where DocID ='" & txt(TxtDocID) & "' And V_Type = '" & mVType & "' And V_No=" & Val(txt(SerialNo)) & "").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And txt(BookNo).Visible Then
                MsgBox "Challan No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                GoTo errlbl
            Else
                txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) Then
                    MsgBox "Challan No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If
        
        DocIdHlp = Replace(txt(TxtDocID), " ", "")
        '************
        GCn.Execute "insert into sp_purch(DocID,DocIDHelp,V_Type,V_No,Site_Code," _
            & "V_Date,Cash_Credit,Party_Code,Party_Name,Party_Doc_No," _
            & "Party_Doc_Date,RoadPermit_Formcode,RoadPermit_no,GR_RR_No,GR_RR_Date," _
            & "L_C,form_code,Tot_No_of_Items,Tot_Doc_Qty,Tot_Phy_Qty," _
            & "Tot_Amt,Tot_Disc_Amt,Tot_Ord_DiscAmt,Tot_Goods_Value,Tax_Amt," _
            & "Addition,Deduction,NET_AMT,Case_no,Case_Mark," _
            & "Transport,Supply_Mode,U_Name,U_EntDt,U_AE, AddBy, AddDate,DrAc_Code, SatAmt, Sat_Yn,SFCAMT) values(" _
            & "'" & txt(TxtDocID) & "','" & DocIdHlp & "','" & mVType & "'," & Val(txt(SerialNo).TEXT) & ",'" & PubSiteCode & PubSiteCode & "'," _
            & "" & ConvertDate(txt(VDate).TEXT) & ",'" & txt(VType).TEXT & "','" & txt(Party).Tag & "','" & txt(Party).TEXT & "','" & txt(SuppChlNo).TEXT & "'," _
            & "" & ConvertDate(txt(SuppChlDate).TEXT) & ",'" & txt(PermitType).Tag & "','" & txt(FormNo).TEXT & "','" & txt(GrNo).TEXT & "'," & ConvertDate(txt(GrDate).TEXT) & "," _
            & "'" & left(txt(LC).TEXT, 1) & "','" & txt(FormType).Tag & "'," & Val(LblIVal.CAPTION) & "," & Val(LblDQty.CAPTION) & "," & Val(LblPQty.CAPTION) & "," _
            & "" & Val(txt(TOTAmt).TEXT) & "," & Val(txt(TotDis).TEXT) & "," & Val(txt(TotOrdDis).TEXT) & "," & Val(txt(TotGoods).TEXT) & "," & Val(txt(TaxAmt).TEXT) & "," _
            & "" & Val(txt(Addition).TEXT) & "," & Val(txt(Deduction).TEXT) & "," & Val(txt(NetAmt).TEXT) & "," & Val(txt(CaseNo).TEXT) & ",'" & txt(CaseMark).TEXT & "'," _
            & "'" & txt(Transport).TEXT & "','" & txt(SupplyMode).TEXT & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & ",'" & txt(DrAcCode).Tag & "', " & Val(txt(SatAmt)) & ", " & IIf(mSatYn, 1, 0) & ", " & Val(txt(SFCAmt)) & ")"
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, txt(TxtDocID), txt(VDate)
    Else
        UpdateOrderQty
        'Stock unposting
        UpdStkTableToTable txt(TxtDocID), "-", "R"
        'eof stock unposting
        GCn.Execute ("delete from sp_stock where docid='" & txt(TxtDocID) & "'")
        GCn.Execute ("update sp_purch set V_Date=" & ConvertDate(txt(VDate).TEXT) & ", Cash_Credit = '" & txt(VType) & "', Party_Code = '" & txt(Party).Tag & "', Party_Name= '" & txt(Party) & "', Party_Doc_No ='" & txt(SuppChlNo) & "',Party_Doc_Date =" & ConvertDate(txt(SuppChlDate)) & _
            ",RoadPermit_Formcode='" & txt(PermitType).Tag & "',RoadPermit_no='" & txt(FormNo) & "',GR_RR_No='" & txt(GrNo) & "',GR_RR_Date=" & ConvertDate(txt(GrDate)) & ",L_C = '" & left(txt(LC), 1) & "',form_code = '" & txt(FormType).Tag & _
            "',Tot_No_of_Items = " & Val(LblIVal.CAPTION) & ",Tot_Doc_Qty = " & Val(LblDQty.CAPTION) & ",Tot_Phy_Qty = " & Val(LblPQty.CAPTION) & ",Tot_Amt = " & Val(txt(TOTAmt)) & ",Tot_Disc_Amt= " & Val(txt(TotDis)) & _
            ",Tot_Ord_DiscAmt=" & Val(txt(TotOrdDis)) & ",Tot_Goods_Value=" & Val(txt(TotGoods)) & ",Tax_Amt=" & Val(txt(TaxAmt)) & ",   Addition =" & Val(txt(Addition)) & "  , Deduction=" & Val(txt(Deduction)) & _
            ",NET_AMT = " & Val(txt(NetAmt)) & ",SFCAMT=" & Val(txt(SFCAmt)) & ",Case_no=" & Val(txt(CaseNo)) & ",Case_Mark='" & txt(CaseMark) & "',Transport = '" & txt(Transport) & "',Supply_Mode = '" & txt(SupplyMode) & _
            "',U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E', ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ", SatAmt = " & Val(txt(SatAmt)) & ", DrAc_Code='" & txt(DrAcCode).Tag & _
            "' where docid = '" & txt(TxtDocID) & "'")
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" And Val(FGrid.TextMatrix(I, PQty)) <> 0 Then
            GCn.Execute ("insert into sp_stock(DocID,Srl_No,V_Type,V_No,V_Date,Party_Code,L_C,Order_DocId,Order_Srl_No, " & _
                " Part_No, Godown, Qty_Doc, Qty_Rec, Tax_YN, MRP_YN, Rate, V_Rate, " & _
                " Disc_Per,Disc_Amt, Amount, Ord_DiscPer, Ord_DiscAmt, Net_Amt, " & _
                " Part_SrlNo, Site_Code, U_Name, U_EntDt, U_AE,TaxPer,TaxAmt, SatPer, SatAmt, SFCPer, SFCAmt) " & _
                " values('" & txt(TxtDocID) & "'," & I & ",'" & mVType & "'," & Val(txt(SerialNo).TEXT) & "," & ConvertDate(txt(VDate).TEXT) & ",'" & txt(Party).Tag & "','" & left(txt(LC).TEXT, 1) & "','" & FGrid.TextMatrix(I, PONOCode) & "'," & Val(FGrid.TextMatrix(I, POSrlNo)) & _
                ",'" & FGrid.TextMatrix(I, PNo) & "','" & FGrid.TextMatrix(I, God) & "'," & Val(FGrid.TextMatrix(I, DQty)) & ", " & Val(FGrid.TextMatrix(I, PQty)) & "," & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ", " & IIf(FGrid.TextMatrix(I, MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, FRate)) & " ," & Val(FGrid.TextMatrix(I, NDP)) & _
                "," & Val(FGrid.TextMatrix(I, DisPer)) & "," & Val(FGrid.TextMatrix(I, DisRs)) & "," & Val(FGrid.TextMatrix(I, Amt)) & ", " & Val(FGrid.TextMatrix(I, DisOrd)) & "," & Val(FGrid.TextMatrix(I, DisOrdRs)) & "," & Val(FGrid.TextMatrix(I, ItemVal)) & _
                ",'" & FGrid.TextMatrix(I, PartSrlNo) & "','" & PubSiteCode & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "'," & Val(FGrid.TextMatrix(I, TaxPer)) & " ," & Val(FGrid.TextMatrix(I, TaxAmt1)) & "," & Val(FGrid.TextMatrix(I, SatPer)) & " ," & Val(FGrid.TextMatrix(I, SatAmt1)) & "," & Val(FGrid.TextMatrix(I, SFCPer)) & " ," & Val(FGrid.TextMatrix(I, SFCAmt1)) & ")")
                
            If FGrid.TextMatrix(I, PONo) <> "" And FGrid.TextMatrix(I, POSrlNo) <> "" Then
                GCn.Execute "Update SP_Order1 Set Sup_Qty=" & Val(FGrid.TextMatrix(I, PQty)) & " Where OrderId='" & FGrid.TextMatrix(I, PONOCode) & "' and Srl_No=" & FGrid.TextMatrix(I, POSrlNo) & ""
            End If
            '" & IIf( = "Yes", 1, 0) & ", " & IIf(FGrid.TextMatrix(i, MRP) = "Yes", 1, 0) & "
            Call UpdStkGridToTable(FGrid.TextMatrix(I, PNo), "+", FGrid.TextMatrix(I, MRP), FGrid.TextMatrix(I, Taxable), FGrid.TextMatrix(I, PQty))
        End If
    Next
    'Update Last purch rate
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" And Val(FGrid.TextMatrix(I, PQty)) <> 0 Then
            GCn.Execute ("Update Part Set PurDocId = '" & txt(TxtDocID) & "',PurDate = " & ConvertDate(txt(VDate)) & ",PurRate=" & Val(FGrid.TextMatrix(I, NDP)) & " where Part_No='" & FGrid.TextMatrix(I, PNo) & "' and Div_Code='" & PubDivCode & "'")
        End If
    Next
    'A/c Posting for Transfer Case
    If left(txt(ChlType).TEXT, 1) = "S" Then
        mNarr = "Through Stock Transfer Receipts"
        I = 0
        LedgAry(I).SubCode = txt(DrAcCode).Tag
        LedgAry(I).AmtDr = Val(txt(NetAmt))
        LedgAry(I).Narration = mNarr
        LedgAry(I).ContraSub = txt(Party).Tag
        I = I + 1
        LedgAry(I).SubCode = txt(Party).Tag
        LedgAry(I).AmtCr = Val(txt(NetAmt))
        LedgAry(I).Narration = mNarr
        LedgAry(I).ContraSub = txt(DrAcCode).Tag
        If Val(txt(NetAmt)) > 0 Then
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, txt(TxtDocID), CDate(txt(VDate)), mNarr & "[Common]")
            If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        End If
    End If
    'EOF Posting
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select DocID as SearchCode, DocID from Sp_Purch where left(DocId,1) = '" & PubDivCode & "' and v_type in ('" & ChalVType & "','" & TrfVType & "') and v_date<=" & ConvertDate(Format(PubEndDate, "dd-mm-yyyy")) & " And DocId = '" & txt(TxtDocID) & "' Order by V_Date Desc, DocId desc")
    End If
    Master.MoveFirst
'    rsTrans.Requery
    Master.FIND "DocId = '" & txt(TxtDocID) & "'"
    'lp 11-03-03
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        txt(VDate).Tag = txt(VDate).TEXT
       If Val(txt(SerialNo)) > DeCodeDocID(txt(TxtDocID).Tag, Document_No) Then
            MsgBox "Challan No." & Trim(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    'If TopCtrl1.TopText2.CAPTION = "Add" Then
        TopCtrl1_ePrn
   ' End If
    Exit Sub

errlbl:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
      Dim sitecond As String
      sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("sp_purch.docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "SELECT sp_purch.DocId as searchcode,Sp_Purch.DocId, sp_purch.V_Type, " & cCStr("sp_purch.V_No") & " As V_No,SP_Purch.Party_Doc_No as PDocNo, sp_purch.Site_Code, " & cDt("sp_purch.V_Date") & "  AS VoucherDate, SubGroup.Name as PartyName FROM sp_purch LEFT JOIN SubGroup ON sp_purch.Party_Code = SubGroup.Subcode where left(DocId,1) = '" & PubDivCode & "' " & sitecond & " and v_type in ('" & ChalVType & "','" & TrfVType & "') order by sp_purch.V_Date Desc,sp_purch.docId"
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
        Set Master = GCn.Execute("select DocID as SearchCode, DocID from Sp_Purch where left(DocId,1) = '" & PubDivCode & "' and v_type in ('" & ChalVType & "','" & TrfVType & "') and v_date<=" & ConvertDate(Format(PubEndDate, "dd-mm-yyyy")) & " And DocId = '" & MyValue & "' Order by V_Date Desc, DocId desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
If txt(ChlType).TEXT = "" And Index <> VDate Then txt(ChlType).SetFocus
TxtGrid(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case ChlType
        ListArray = Array("Purchase Receipt", "Stock Transfer")
        Set mListItem = ListView_Items(ListView, txt, ChlType, ListArray, 2)
    Case VType
        ListArray = Array("Cash", "Credit")
        Set mListItem = ListView_Items(ListView, txt, VType, ListArray, 2)
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
    Case DrAcCode
        Set DGDrAc.DataSource = RsDrAc
        If RsDrAc.RecordCount = 0 Or (RsDrAc.EOF = True Or RsDrAc.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsDrAc!Name Then
            RsDrAc.MoveFirst
            RsDrAc.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case PermitType
        Set DGForm.DataSource = rsForm31
        If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm31!Name Then
            rsForm31.MoveFirst
            rsForm31.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case FormType
        Set DGForm.DataSource = rsForm
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case SerialNo
        If IsValid(txt(ChlType), "Challan Type") = False Then Exit Sub
    Case Addition, Deduction, TaxAmt, NetAmt
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
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case LC
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case ChlType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case Party
        If txt(VType).TEXT = "Credit" Then
            DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        End If
    Case DrAcCode   'Transfer Case
        DGridTxtKeyDown DGDrAc, txt, Index, RsDrAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    Case Transport
        DGridTxtKeyDown_Mast DGTrans, txt, Transport, rsTrans, KeyCode, False, 0
    Case FormType
        DGridTxtKeyDown DGForm, txt, FormType, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
    Case PermitType
        DGridTxtKeyDown DGForm, txt, PermitType, rsForm31, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
End Select
If DGDrAc.Visible = False And FrmList.Visible = False And DGTrans.Visible = False And DGGod.Visible = False And DGParty.Visible = False And DGForm.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Deduction Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Deduction Then
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
    Case PermitType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, Index, rsForm31, KeyAscii, "Name"
    Case Party
        If txt(VType).TEXT = "Credit" Then
            If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "Name"
            lblGroup.Visible = True: lblGroup.BackColor = vbBlack: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
        End If
    Case DrAcCode
        If DGDrAc.Visible = True Then DGridTxtKeyPress txt, Index, RsDrAc, KeyAscii, "Name"
'    Case Transport
'        If DGTrans.Visible = True Then DGridTxtKeyPress txt, Index, rsTrans, KeyAscii, "Name"
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, Index, rsForm, KeyAscii, "Name"
    Case SerialNo, CaseNo
        Call NumPress(txt(Index), KeyAscii, 6, 0)
    Case Addition, Deduction, TaxAmt, NetAmt
        Call NumPress(txt(Index), KeyAscii, 8, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case ChlType
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case VType
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case LC
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case Transport
        If DGTrans.Visible = True Then DGridTxtKeyUp_Mast txt, Transport, rsTrans, KeyCode, "Name"
    Case Addition, Deduction, TaxAmt
        Amt_Cal
End Select
Amt_Cal
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, mDrAcFlag As Boolean
Select Case Index
    Case ChlType
        If txt(Index).TEXT <> "" Then
            txt(Index).TEXT = ListView.SelectedItem.TEXT
            If left(txt(Index).TEXT, 1) = "P" Then
                mVType = ChalVType
                mDrAcFlag = False
            ElseIf left(txt(Index).TEXT, 1) = "S" Then
                mVType = TrfVType
                mDrAcFlag = True
                txt(VType) = "Credit"
            End If
        End If
        txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        txt(TxtDocID).Tag = txt(TxtDocID)
        If IsValid(txt(ChlType), Label3(31)) = False Then Cancel = True:   Exit Sub
        Label3(6).Visible = mDrAcFlag
        'LblColon(0).Visible = mDrAcFlag
        txt(DrAcCode).Visible = mDrAcFlag
        txt(VType).Enabled = Not mDrAcFlag
    Case VType
        If IsValid(txt(VType), "Cash Credit") = False Then Cancel = True:   Exit Sub
        If txt(VType).Tag = txt(VType).TEXT Then Exit Sub
        If txt(VType).TEXT <> "" Then txt(VType).TEXT = ListView.SelectedItem.TEXT
        If txt(VType).TEXT = "Cash" Then
            txt(Party).TEXT = "Cash"
            txt(Party).Tag = PubSprCashAc
            mPartyType = 0
        Else
            txt(Party).TEXT = ""
            txt(Party).Tag = ""
        End If
        If TopCtrl1.TopText2 = "Add" Then
            txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        End If
        txt(TxtDocID).Tag = txt(TxtDocID)
        txt(VType).Tag = txt(VType).TEXT
    Case LC
        If txt(LC).TEXT <> "" Then txt(LC).TEXT = ListView.SelectedItem.TEXT
        If IsValid(txt(LC), "Purchase Type") = False Then Cancel = True:   Exit Sub
    Case Party
        If IsValid(txt(Index), Label3(3)) = False Then Cancel = True: Exit Sub
        If txt(ChlType) = "Purchase Receipt" Then
            If txt(VType).TEXT = "Cash" Then
                mPartyType = 0
                txt(Index).Tag = PubSprCashAc
                GSQL = "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, " & cTrim(cMID("OrderID", "8", "5")) & " + " & cCStr(cTrim("Right(OrderID,8)")) & " as OurDocNo,V_Date From SP_Order Where left(OrderId,1) = '" & PubDivCode & "' and left(Order_Type,4)='S_PO' and V_Date<=" & ConvertDate(Format(txt(VDate), "dd-mmm-yyyy")) & " and OrdClosDate is null Order By OrderID"
            ElseIf txt(VType).TEXT = "Credit" Then
                txt(Index).TEXT = RsParty!Name
                txt(Index).Tag = RsParty!Code
                mPartyType = RsParty!Party_Type
                GSQL = "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, " & cTrim(cMID("OrderID", "8", "5")) & " + " & cCStr(cTrim("Right(OrderID,8)")) & " as OurDocNo,V_Date From SP_Order Where left(OrderId,1) = '" & PubDivCode & "' and left(Order_Type,4)='S_PO' and Party_Code='" & txt(Party).Tag & "' and V_Date<=" & ConvertDate(Format(txt(VDate), "dd-mmm-yyyy")) & " and OrdClosDate is null Order By OrderID"
            End If
            Set rsPONo = New ADODB.Recordset
            rsPONo.CursorLocation = adUseClient
            rsPONo.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
            Set DGPONo.DataSource = rsPONo
        Else
'            Txt(Index).Tag = RsParty!Code
            mPartyType = RsParty!Party_Type
        End If
    Case DrAcCode
        If IsValid(txt(Index), Label3(6)) = False Then Cancel = True: Exit Sub
        txt(Index).TEXT = RsDrAc!Name
        txt(Index).Tag = RsDrAc!Code
    Case PermitType
        If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsForm31!Name
            txt(Index).Tag = rsForm31!Code
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsForm!Name
            txt(Index).Tag = rsForm!Code
        End If
    Case SuppChlNo
        If GCn.Execute("Select * from SP_Purch where Party_Code='" & txt(Party).Tag & "' and Party_Doc_No='" & txt(SuppChlNo).TEXT & "'").RecordCount > 0 And TopCtrl1.TopText2 = "Add" Then
            MsgBox " This Supplier Document No for the Party Already Exists.", vbInformation + vbOKOnly, "Validation": txt(SuppChlNo).SetFocus
            Cancel = True
            Exit Sub
        End If
    Case SuppChlDate, GrDate
        txt(Index).TEXT = RetDate(txt(Index))
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
             txt(VDate).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        Cancel = Not CheckFinYear(txt(Index))
        If Cancel = False Then
            If TopCtrl1.TopText2.CAPTION = "Add" Then
                txt(VType).SetFocus
            Else
                txt(Party).SetFocus
            End If
        End If
    Case SerialNo
        If IsValid(txt(SerialNo), "SerialNo") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag Then      ' Manual
            txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            txt(TxtDocID).Tag = txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select Docid From sp_purch Where docid='" & txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(SerialNo).SetFocus
            End If
        End If
    Case TaxAmt
        txt(Index) = Format(Val(txt(Index)), "0.00")
End Select
Set Rst = Nothing
End Sub

Private Sub DGPart_Click()
On Error GoTo ELoop
    If RsPart.RecordCount > 0 Then
        Select Case FGrid.Col
            Case PNo
                TxtGrid(0).TEXT = RsPart!Code
            Case PName
                TxtGrid(0).TEXT = RsPart!Name
            Case LName
                TxtGrid(0).TEXT = RsPart!LName
        End Select
    End If
    TxtGridValid_PNo
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGPart.Visible = False
Exit Sub
ELoop:
    CheckError
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

Private Sub DGTrans_Click()
    If rsTrans.RecordCount > 0 Then
        txt(Transport).TEXT = rsTrans!Name
    End If
    txt(Transport).SetFocus
    DGTrans.Visible = False
End Sub

Private Sub DGPONo_Click()
    If rsPONo.RecordCount > 0 Then
            TxtGrid(0).TEXT = rsPONo!Name
            FGrid.TextMatrix(FGrid.Row, PONo) = rsPONo!Name
            FGrid.TextMatrix(FGrid.Row, PONOCode) = rsPONo!Code
    End If
    TxtGrid(0).SetFocus
    DGPONo.Visible = False
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
        Case PONo, PQty, DQty, PartSrlNo
           FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case DisPer, DisRs
            FGrid.TextMatrix(FGrid.Row, DisRs) = "" '0.00"
            FGrid.TextMatrix(FGrid.Row, DisPer) = "" '0.00"
        Case DisOrd, DisOrdRs
            FGrid.TextMatrix(FGrid.Row, DisOrd) = "" '0.00"
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = "" '0.00"
        Case TaxPer, TaxAmt1
            FGrid.TextMatrix(FGrid.Row, TaxPer) = ""  '0.00"
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""  '0.00"
        Case SatPer, SatAmt
            FGrid.TextMatrix(FGrid.Row, SatPer) = ""  '0.00"
            FGrid.TextMatrix(FGrid.Row, SatAmt1) = ""  '0.00"
    End Select
    Amt_Cal1
    Amt_Cal
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case PNo, PName, LName
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
        Case PONo, Taxable, MRP, Godown, PQty, DQty, FRate, DisPer, DisOrd, DisRs, DisOrdRs, PartSrlNo, TaxPer, TaxAmt1, SatPer, SatAmt1
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
    Case PONo
        Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
    Case PNo, PName, LName
       Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
    Case Unit, Amt, ItemVal
        FGrid.Col = FGrid.Col + 1
        FGrid.SetFocus
    Case PartSrlNo, Godown, MRP, Taxable
        If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
           Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        End If
    Case PQty, DQty, FRate, DisPer, DisOrd, DisRs, DisOrdRs
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
    If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
    If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then Me.TopCtrl1.tDel = True
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select SubGroup.Name,SubGroup.Party_Type,SP_Purch.* from SP_Purch " _
        & " left join SubGroup on SP_Purch.Party_Code=SubGroup.SubCode " _
        & " where DocID='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
    If Master1!CancelYN = 1 Then
        LblCancel.Visible = True
        TopCtrl1.tEdit = False
    Else
        LblCancel.Visible = False
    End If
    If Master1!Invoice_DocID <> "" Then
        TopCtrl1.tEdit = False
        TopCtrl1.tDel = False
    End If
    If PubBackEnd = "A" Then
        mSatYn = IIf(VNull(Master1!SAT_YN) = 1, True, False)
    Else
        mSatYn = IIf(VNull(Master1!SAT_YN) = True, True, False)
    End If
    
    DispText_Vat
    
    txt(SatAmt) = Format(VNull(Master1!SatAmt), "0.00")
    txt(SFCAmt) = Format(VNull(Master1!SFCAmt), "0.00")
    txt(TxtDocID).TEXT = Master!SearchCode
    txt(TxtDocID).Tag = txt(TxtDocID)
    mVType = Master1!V_Type
    LblDiv.CAPTION = "Division : " & left(Master!DocID, 1)
    LblSite.CAPTION = "Site Code : " & Master1!Site_Code
    LblUser = IIf(Not IsNull(Master1!AddDate), "Add By : " & XNull(Master1!AddBy) & "  Dated : " & XNull(Master1!AddDate), "") & IIf(Not IsNull(Master1!ModifyDate), "     Modify By : " & XNull(Master1!ModifyBy) & "  Dated : " & XNull(Master1!ModifyDate), "")
    LblVPrefix.CAPTION = mID(Master!DocID, 9, 5)
    txt(SerialNo).TEXT = Master1!V_NO
    txt(VDate).TEXT = Master1!V_DATE
    mVType = Master1!V_Type
    If mVType = ChalVType Then
        txt(ChlType).TEXT = "Purchase Receipt"
        mDrAcFlag = False
    ElseIf mVType = TrfVType Then
        txt(ChlType).TEXT = "Stock Transfer"
        mDrAcFlag = True
        txt(DrAcCode) = GCn.Execute("Select Name from subgroup where SubCode='" & Master1!DrAc_Code & "'").Fields(0).Value
    End If
    Label3(6).Visible = mDrAcFlag
    LblColon(0).Visible = mDrAcFlag
    txt(DrAcCode).Visible = mDrAcFlag
    txt(DrAcCode).Tag = IIf(IsNull(Master1!DrAc_Code), "", Master1!DrAc_Code)
    txt(VType).TEXT = Master1!Cash_Credit
    txt(Party).Tag = Master1!Party_code
    If Master1!Cash_Credit = "Cash" Then
        txt(Party).TEXT = Master1!Party_Name
        mPartyType = 0
    Else
        txt(Party).TEXT = IIf(IsNull(Master1!Name), "", Master1!Name)
        mPartyType = VNull(Master1!Party_Type)
    End If
    txt(SuppChlNo).TEXT = IIf(IsNull(Master1!Party_Doc_No), "", Master1!Party_Doc_No)
    txt(SuppChlDate).TEXT = IIf(IsNull(Master1!Party_Doc_Date), "", Master1!Party_Doc_Date)
    txt(PermitType).Tag = IIf(IsNull(Master1!RoadPermit_FormCode), "", Master1!RoadPermit_FormCode)
    If txt(PermitType).Tag <> "" Then
        txt(PermitType).TEXT = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(PermitType).Tag & "'").Fields(0).Value
    Else
        txt(PermitType).TEXT = ""
    End If
    txt(FormNo).TEXT = IIf(IsNull(Master1!RoadPermit_No), "", Master1!RoadPermit_No)
    txt(GrNo).TEXT = IIf(IsNull(Master1!GR_RR_No), "", Master1!GR_RR_No)
    txt(GrDate).TEXT = IIf(IsNull(Master1!GR_RR_Date), "", Master1!GR_RR_Date)
    txt(LC).TEXT = IIf(Master1!L_C = "L", "Local", "Central")
    txt(FormType).Tag = IIf(IsNull(Master1!Form_Code), "", Master1!Form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType).TEXT = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType).TEXT = ""
    End If
    LblIVal.CAPTION = Format(IIf(IsNull(Master1!Tot_No_of_Items), 0, Master1!Tot_No_of_Items), "0")
    LblDQty.CAPTION = Format(IIf(IsNull(Master1!Tot_Doc_Qty), 0, Master1!Tot_Doc_Qty), "0.000")
    LblPQty.CAPTION = Format(IIf(IsNull(Master1!Tot_Phy_Qty), 0, Master1!Tot_Phy_Qty), "0.000")
    LblAmt.CAPTION = Format(IIf(IsNull(Master1!Tot_Amt), 0, Master1!Tot_Amt), "0.00")
    txt(TOTAmt).TEXT = Format(IIf(IsNull(Master1!Tot_Amt), 0, Master1!Tot_Amt), "0.00")
    txt(TotDis).TEXT = Format(IIf(IsNull(Master1!Tot_Disc_Amt), 0, Master1!Tot_Disc_Amt), "0.00")
    txt(TotOrdDis).TEXT = Format(IIf(IsNull(Master1!Tot_Ord_DiscAmt), 0, Master1!Tot_Ord_DiscAmt), "0.00")
    txt(TotGoods).TEXT = Format(IIf(IsNull(Master1!Tot_Goods_Value), 0, Master1!Tot_Goods_Value), "0.00")
    txt(TaxAmt).TEXT = Format(IIf(IsNull(Master1!Tax_Amt), 0, Master1!Tax_Amt), "0.00")
    txt(Addition).TEXT = Format(IIf(IsNull(Master1!Addition), 0, Master1!Addition), "0.00")
    txt(Deduction).TEXT = Format(IIf(IsNull(Master1!Deduction), 0, Master1!Deduction), "0.00")
    txt(NetAmt).TEXT = Format(IIf(IsNull(Master1!Net_Amt), 0, Master1!Net_Amt), "0.00")
    txt(CaseNo).TEXT = IIf(IsNull(Master1!Case_No), "", Master1!Case_No)
    txt(CaseMark).TEXT = IIf(IsNull(Master1!Case_Mark), "", Master1!Case_Mark)
    txt(Transport).TEXT = IIf(IsNull(Master1!Transport), "", Master1!Transport)
    txt(SupplyMode).TEXT = IIf(IsNull(Master1!Supply_Mode), "", Master1!Supply_Mode)
    
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT SPO.Order_Reg_No,P.Part_Name, " & cTrim(cMID("Stk.Order_DocID", "8", "5")) & " + " & cCStr(cTrim("Right(Stk.Order_DocID,8)")) & " As OrderIDDisp, " & _
            " P.Local_Name,P.UNIT, P.MRP, P.Cur_MRP_TBStk,P.Cur_MRP_TPStk,P.Cur_TB_Stk,P.Cur_TP_Stk, " & _
            " P.TP_SRate,P.Part_Grade, P.TB_SRate, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate, Stk.*, G.God_Name" & _
            " FROM ((Sp_Stock Stk LEFT JOIN Part P ON Stk.Part_No = P.PART_NO and P.Div_Code = left(STK.Docid,1)) LEFT JOIN Godown G ON Stk.Godown = G.God_Code ) " & _
            " Left Join SP_Order SPO on Stk.Order_DocID=SPO.OrderID " & _
            " where Stk.docId = '" & Master!DocID & "'")
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        I = 1
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = Rs!Srl_No
'                .TextMatrix(i, PONo) = IIf(IsNull(rs!OrderIDDisp), "", rs!OrderIDDisp)
                .TextMatrix(I, PONo) = IIf(IsNull(Rs!Order_Reg_No), "", Rs!Order_Reg_No)
                .TextMatrix(I, PNo) = Rs!Part_No
                .TextMatrix(I, PONOCode) = XNull(Rs!Order_DocId)
                .TextMatrix(I, POSrlNo) = XNull(Rs!Order_Srl_No)
                .TextMatrix(I, Unit) = IIf(IsNull(Rs!Unit), "", Rs!Unit)
                .TextMatrix(I, MRP) = IIf(Rs!MRP_YN = 1, "Yes", "No")
                .TextMatrix(I, Taxable) = IIf(Rs!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, DQty) = IIf(Rs!Qty_Doc = 0, "", Format(Rs!Qty_Doc, "0.000"))
                .TextMatrix(I, PQty) = IIf(Rs!Qty_Rec = 0, "", Format(Rs!Qty_Rec, "0.000"))
                .TextMatrix(I, FRate) = IIf(Rs!Rate = 0, "", Format(Rs!Rate, "0.0000"))
                .TextMatrix(I, Amt) = IIf(Rs!Amount = 0, "", Format(Rs!Amount, "0.00"))
                .TextMatrix(I, DisPer) = IIf(Rs!Disc_Per = 0, "", Format(Rs!Disc_Per, "0.00"))
                .TextMatrix(I, DisRs) = IIf(Rs!Disc_Amt = 0, "", Format(Rs!Disc_Amt, "0.00"))
                .TextMatrix(I, DisOrd) = IIf(Rs!ord_Discper = 0, "", Format(Rs!ord_Discper, "0.00"))
                .TextMatrix(I, DisOrdRs) = IIf(Rs!ord_Discamt = 0, "", Format(Rs!ord_Discamt, "0.00"))
                
                   .TextMatrix(I, SFCPer) = VNull(Rs!SFCPer)
                .TextMatrix(I, SFCAmt1) = Format(VNull(Rs!SFCAmt), "0.00")
                
                If PubVATYN = 1 Then
                    .TextMatrix(I, TaxPer) = VNull(Rs!TaxPer)
                    .TextMatrix(I, TaxAmt1) = Format(VNull(Rs!TaxAmt), "0.00")
                    If mSatYn Then
                        .TextMatrix(I, SatPer) = VNull(Rs!SatPer)
                        .TextMatrix(I, SatAmt1) = Format(VNull(Rs!SatAmt), "0.00")
                    End If
                End If
                .TextMatrix(I, NDP) = IIf(Rs!V_Rate = 0, "", Format(Rs!V_Rate, "0.00"))
                .TextMatrix(I, ItemVal) = IIf(Rs!Net_Amt = 0, "", Format(Rs!Net_Amt, "0.00"))
                .TextMatrix(I, God) = Rs!Godown
                .TextMatrix(I, Godown) = IIf(IsNull(Rs!God_Name), "", Rs!God_Name)
                .TextMatrix(I, PName) = IIf(IsNull(Rs!Part_Name), "", Rs!Part_Name)
                .TextMatrix(I, LName) = IIf(IsNull(Rs!Local_Name), "", Rs!Local_Name)
                .TextMatrix(I, MRPStkTB) = IIf(IsNull(Rs!Cur_MRP_TbStk), "", Rs!Cur_MRP_TbStk)
                .TextMatrix(I, MRPStkTP) = IIf(IsNull(Rs!Cur_MRP_TPStk), "", Rs!Cur_MRP_TPStk)
                .TextMatrix(I, MRPRate) = IIf(Rs!MRP = 0, "", Format(Rs!MRP, "0.00"))
                .TextMatrix(I, TBStk) = IIf(IsNull(Rs!Cur_TB_STk), "", Rs!Cur_TB_STk)
                .TextMatrix(I, TPStk) = IIf(IsNull(Rs!Cur_TP_Stk), "", Rs!Cur_TP_Stk)
                .TextMatrix(I, TBRate) = IIf(IsNull(Rs!TB_SRate), "", Rs!TB_SRate)
                .TextMatrix(I, TPRate) = IIf(IsNull(Rs!TP_SRate), "", Rs!TP_SRate)
                .TextMatrix(I, Bin) = IIf(IsNull(Rs!Bin_Loca), "", Rs!Bin_Loca)
                'modishekhar
                .TextMatrix(I, PartGrade) = IIf(IsNull(Rs!Part_Grade), "", Rs!Part_Grade)
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
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
Set Rs = Nothing
Set Master1 = Nothing
Grid_Hide
Call Amt_Cal
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte
' |Part No.1|Part Name2|Unit 3|PO No 4|Taxable 5|MRP6|Qty(Doc)7|Qty(Phy)8|NDP 9 |Amount 10
' |Dis %11|Ord Dis %12|Amount 13|Loal Name 14|Curr Stk Qty 15|MRP Qty 16 |Taxable Qty 17|TaxPaid Qty 18|Taxable Rate 19|TaxPaid Rate 20|Bin Location 21|Last Purch Rate 22|High Purch Rate 23|Low Purch Rate 24

'    FGrid.FormatString = "SrNo.|Part No.            |Part Name             |Unit |Godown          |PO No.         |Tax Y/N|MRP Y/N| Qty(Doc)|Qty(Phy)|Rate     |Amount    |Dis %    |Dis Rs   |Ord Dis %  |Ord Dis Rs  |NDP     |ItemValue   |Local Name|Curr Stk Qty|MRP Qty|Taxable Qty|TaxPaid Qty|Taxable Rate|TaxPaid Rate|Bin Location|Last Purch Rate|High Purch Rate|Low Purch Rate"
'    SrNo.1|Part No.2|Part Name3|Unit 4|Godown5|PO No.6|Tax Y/N 7|MRP Y/N8| Qty(Doc)9|Qty(Phy)10|Rate 11|Amount12|Dis %13|Dis Rs14|Ord Dis %15|Ord Dis Rs16|NDP 17|ItemValue 18|Local Name19|Curr Stk Qty20|MRP Qty21|Taxable Qty22|TaxPaid Qty23|Taxable Rate24|TaxPaid Rate25|Bin Location26|Last Purch Rate27|High Purch Rate28|Low Purch Rate29"
    With FGrid
        .Cols = 42
        .RowHeightMin = PubGridRowHeight
        
        .TextMatrix(0, 0) = "S.No."
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, PONo) = "Order No"
        .ColAlignment(PONo) = flexAlignLeftCenter
        .ColWidth(PONo) = 1395
        
        .TextMatrix(0, POSrlNo) = "Order SrlNo"
        .ColAlignment(POSrlNo) = flexAlignLeftCenter
        .ColWidth(POSrlNo) = 0
        
        .TextMatrix(0, PNo) = "Part No"
        .ColAlignment(PNo) = flexAlignLeftCenter
        .ColWidth(PNo) = 1500
        
        .TextMatrix(0, Unit) = "Unit"
        .ColAlignment(Unit) = flexAlignLeftCenter
        .ColWidth(Unit) = 550

        .TextMatrix(0, MRP) = "MRP"
        .ColAlignment(MRP) = flexAlignLeftCenter
        .ColWidth(MRP) = 450

        .TextMatrix(0, Taxable) = "Tax"
        .ColAlignment(Taxable) = flexAlignLeftCenter
        .ColWidth(Taxable) = 420
        
        .TextMatrix(0, DQty) = "Qty(Doc)"
        .ColAlignmentFixed(DQty) = flexAlignRightCenter
        .ColWidth(DQty) = 960

        .TextMatrix(0, PQty) = "Qty(Phy)"
        .ColAlignmentFixed(PQty) = flexAlignRightCenter
        .ColWidth(PQty) = 960

        .TextMatrix(0, FRate) = "Rate" 'NDP"
        .ColAlignmentFixed(FRate) = flexAlignRightCenter
        .ColWidth(FRate) = 870

        .TextMatrix(0, Amt) = "Amount"
        .ColAlignmentFixed(Amt) = flexAlignRightCenter
        .ColWidth(Amt) = 1065

        .TextMatrix(0, DisPer) = "Disc%"
        .ColAlignmentFixed(DisPer) = flexAlignRightCenter
        .ColWidth(DisPer) = 555

        .TextMatrix(0, DisRs) = "Disc.Amt"
        .ColAlignmentFixed(DisRs) = flexAlignRightCenter
        .ColWidth(DisRs) = 840
        
        .TextMatrix(0, DisOrd) = "ODis%"
        .ColAlignmentFixed(DisOrd) = flexAlignRightCenter
        .ColWidth(DisOrd) = 555

        .TextMatrix(0, DisOrdRs) = "OrdDisc"
        .ColAlignmentFixed(DisOrdRs) = flexAlignRightCenter
        .ColWidth(DisOrdRs) = 840
        
              .TextMatrix(0, SFCPer) = "SFCPer"
            .ColAlignmentFixed(SFCPer) = flexAlignRightCenter
            .ColWidth(SFCPer) = 840
            
            .TextMatrix(0, SFCAmt1) = "SFCAmt"
            .ColAlignmentFixed(SFCAmt1) = flexAlignRightCenter
            .ColWidth(SFCAmt1) = 840
            
            
        If PubVATYN = 1 Then
        
            .TextMatrix(0, TaxPer) = "TaxPer"
            .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
            .ColWidth(TaxPer) = 840
            
            .TextMatrix(0, TaxAmt1) = "TaxAmt"
            .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
            .ColWidth(TaxAmt1) = 840
            
            If PubSatYn = 1 Then
                .TextMatrix(0, SatPer) = "SAT %"
                .ColAlignmentFixed(SatPer) = flexAlignRightCenter
                .ColWidth(SatPer) = 840
                
                .TextMatrix(0, SatAmt1) = "SAT Amt"
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
        
        .TextMatrix(0, NDP) = "NDP" 'Rate"
        .ColAlignmentFixed(NDP) = flexAlignRightCenter
        .ColWidth(NDP) = 870
    End With
    For I = 21 To 35
        FGrid.ColWidth(I) = 0
    Next
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    FrmDetail.width = 6285: FrmDetail.left = Me.width - (FrmDetail.width + mRtScale): FrmDetail.top = mTopScale: FrmDetail.height = 2130
    FGrid.left = Me.left: FGrid.width = Me.width - 90: FGrid.top = 2500 ': FGrid.height = 2895
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    DGPONo.left = FGrid.left: DGPONo.top = DGPart.top: DGPONo.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    DGGod.left = Me.width - (DGGod.width + mRtScale): DGGod.top = mTopScale
    FrmPrn.left = (Me.width - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
    DGVno.left = 5145: DGVno.top = mTopScale
    
'    DGParty.width = 5130:   DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale '390
 '   DGParty.height = 4935
    
    DGParty.width = 9700:  DGParty.left = 1000: DGParty.top = mTopScale   '390
    DGParty.height = 4935
    DGDrAc.width = 5130:   DGDrAc.left = Me.width - (DGDrAc.width + mRtScale): DGDrAc.top = mTopScale '390
    DGDrAc.height = 4935
    DGTrans.width = DGParty.width: DGTrans.left = DGParty.left: DGTrans.top = DGParty.top: DGTrans.height = DGParty.height
    DGForm.width = DGParty.width: DGForm.left = DGParty.left: DGForm.top = DGParty.top: DGForm.height = DGParty.height
    DGOrdPart.left = (Me.width - DGOrdPart.width) / 2: DGOrdPart.top = FGrid.top + FGrid.height: DGOrdPart.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
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
    txt(ChlType).Enabled = False
End If
txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub

Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGPONo.Visible = True Then DGPONo.Visible = False
    If DGTrans.Visible = True Then DGTrans.Visible = False
    If DGGod.Visible = True Then DGGod.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub

Private Sub Amt_Cal1()
    Dim mAmount As Double, DisAmt As Double, OrdDisAmt1 As Double
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
    If PubVATYN = 1 Then
        If FGrid.TextMatrix(FGrid.Row, TaxPer) <> "" Then
            mAmount = Val(FGrid.TextMatrix(FGrid.Row, Amt))
            DisAmt = Val(FGrid.TextMatrix(FGrid.Row, DisRs))
            OrdDisAmt1 = Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs))
            If FGrid.TextMatrix(FGrid.Row, MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes" Then
                If StrCmp(left(PubComp_Name, 3), "JMK") Then
                
                   FGrid.TextMatrix(FGrid.Row, SFCAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, SFCPer)) / 100, "0.00")
                   ' FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer) + Val(FGrid.TextMatrix(FGrid.Row, SFCAmt1))) / 100, "0.00")
                    FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1) + Val(FGrid.TextMatrix(FGrid.Row, SFCAmt1))) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                    If mSatYn Then
                        FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                    End If
                Else
                    If mSatYn Then
                        mTaxableAmt = Format((mAmount - (DisAmt + OrdDisAmt1)) * 100 / (100 + Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) + Val(FGrid.TextMatrix(FGrid.Row, SatPer))), "0.00")
                        FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                        FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                    Else
                        FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / (100 + Val(FGrid.TextMatrix(FGrid.Row, TaxPer))), "0.00")
                    End If
                    FGrid.TextMatrix(FGrid.Row, ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, ItemVal)) - Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) - Val(FGrid.TextMatrix(FGrid.Row, SatAmt1)), "0.00")
                End If
            ElseIf FGrid.TextMatrix(FGrid.Row, MRP) = "No" And FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes" Then
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                If mSatYn Then
                    FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                End If
            Else
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
                FGrid.TextMatrix(FGrid.Row, SatAmt1) = ""
            End If
        End If
    End If
End Sub
 Private Sub Amt_Cal()
 Dim I As Integer
 Dim IQty As Double
 Dim DQty1 As Double
 Dim ICnt As Integer
 Dim IGAmt As Double
 Dim NSFCAmt As Double
 Dim IDic As Double
 Dim IOrdDic As Double
 Dim IAmt As Double
 Dim TaxPer As Double
 Dim mTaxPer As Double
 Dim TaxAmount As Double, TaxAmountMRP As Double
 Dim SatAmount As Double
 Dim SurPer As Double
 For I = 1 To FGrid.Rows - 1
    If FGrid.TextMatrix(I, PNo) <> "" Then
        IQty = IQty + Val(FGrid.TextMatrix(I, PQty))
        DQty1 = DQty1 + Val(FGrid.TextMatrix(I, DQty))
        IAmt = IAmt + Val(FGrid.TextMatrix(I, Amt))
        IDic = IDic + Val(FGrid.TextMatrix(I, DisRs))
        IOrdDic = IOrdDic + Val(FGrid.TextMatrix(I, DisOrdRs))
        IGAmt = IGAmt + Val(FGrid.TextMatrix(I, ItemVal)) + Val(FGrid.TextMatrix(I, SFCAmt1))
        NSFCAmt = NSFCAmt + Val(FGrid.TextMatrix(I, SFCAmt1))
        
        If PubVATYN = 1 Then
            If FGrid.TextMatrix(I, MRP) = "Yes" Then
                TaxAmountMRP = TaxAmountMRP + Val(FGrid.TextMatrix(I, TaxAmt1))
                TaxAmount = TaxAmount + Val(FGrid.TextMatrix(I, TaxAmt1))
            Else
                TaxAmount = TaxAmount + Val(FGrid.TextMatrix(I, TaxAmt1))
            End If
            SatAmount = SatAmount + Val(FGrid.TextMatrix(I, SatAmt1))
        Else
            If txt(FormType) <> "" And FGrid.TextMatrix(I, Taxable) = "Yes" Then
                TaxPer = GCn.Execute("Select Tax_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'").Fields(0).Value
                SurPer = GCn.Execute("Select Tax_Sur_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'").Fields(0).Value
                mTaxPer = TaxPer + (TaxPer * SurPer / 100)
                    If FGrid.TextMatrix(I, MRP) = "Yes" Then
                        If UCase(left(PubComp_Name, 3)) <> "JMK" Then
                            TaxAmountMRP = TaxAmountMRP + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / (100 + mTaxPer), 2)
                            TaxAmount = TaxAmount + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / (100 + mTaxPer), 2)
                        Else
                            TaxAmountMRP = TaxAmountMRP + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / 100, 2)
                            TaxAmount = TaxAmount + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / 100, 2)
                        End If
                    Else
                        TaxAmount = TaxAmount + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / 100, 2)
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
    txt(SFCAmt).TEXT = Format(NSFCAmt, "0.00")
    txt(TOTAmt).TEXT = Format(IAmt, "0.00")
    txt(TotDis).TEXT = Format(IDic, "0.00")
    txt(TotOrdDis).TEXT = Format(IOrdDic, "0.00")
    txt(TotGoods).TEXT = Format(IGAmt, "0.00")
    txt(TaxAmt) = Format(TaxAmount, "0.00")
    txt(SatAmt) = Format(SatAmount, "0.00")
    
    'If PubVATYN = 1 Then
    '    Txt(NetAmt).TEXT = Format((IGAmt + (Val(Txt(TaxAmt).TEXT)) + Val(Txt(Addition).TEXT) - Val(Txt(Deduction).TEXT)), "0.00")
    'Else
        If UCase(left(PubComp_Name, 3)) <> "JMK" Then
            txt(NetAmt).TEXT = Format((IGAmt + (Val(txt(TaxAmt).TEXT)) + (Val(txt(SatAmt).TEXT)) + Val(txt(Addition).TEXT) - Val(txt(Deduction).TEXT)), "0.00")
        Else
            txt(NetAmt).TEXT = Format((IGAmt + (Val(txt(TaxAmt).TEXT)) + (Val(txt(SatAmt).TEXT)) + Val(txt(Addition).TEXT) - Val(txt(Deduction).TEXT)), "0.00")
        End If
    'End If
    
 End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Grid_Hide
If FrmDetail.Visible = False Then FrmDetail.Visible = True
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
TxtGrid(Index).MaxLength = 0
Select Case FGrid.Col
    Case PONo
        If txt(ChlType) <> "Purchase Receipt" Or rsPONo.RecordCount = 0 _
            Or (rsPONo.EOF = True Or rsPONo.BOF = True) Or FGrid.TextMatrix(FGrid.Row, PONo) = "" Then Exit Sub
        If FGrid.TextMatrix(FGrid.Row, PONOCode) <> rsPONo!Code Then
            rsPONo.MoveFirst
            rsPONo.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, PONOCode) & "'"
        End If
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
        If FGrid.TextMatrix(FGrid.Row, LName) <> "" Then
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
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then TxtGrid(0).TEXT = TxtGrid(0).Tag: Exit Sub
Select Case FGrid.Col
    Case PONo   '3
        If txt(ChlType) = "Purchase Receipt" Then
            DGridTxtKeyDown DGPONo, TxtGrid, Index, rsPONo, KeyCode, True, 1
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
            End If
        End If
    Case PNo    '1
        If DGPart.Visible = False Then DGridColSwap DGPart, 0
        DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, 1
        End If
    Case Godown
        DGridTxtKeyDown DGGod, TxtGrid, 0, rsGod, KeyCode, True, 1, frmGodown, "frmGodown"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, PONo
        End If
    Case MRP, Taxable, DQty, PQty, FRate, DisRs, DisOrdRs, PartSrlNo, TaxAmt1, SatAmt1
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
        End If
        
    Case TaxPer, SatPer
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, , Godown
        End If
        
        
    Case DisPer
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, , DisOrd
        End If
        
    Case DisOrd
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, , Godown
        End If
        
    Case PName
        If DGPart.Visible = False Then DGridColSwap DGPart, 1
        DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 1, frmPartMast, "frmPartMast"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
        End If
    Case LName   '3
        If DGPart.Visible = False Then DGridColSwap DGPart, 2
        DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave Then GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case PONo And txt(ChlType) = "Purchase Receipt"
        If DGPONo.Visible = True Then DGridTxtKeyPress TxtGrid, Index, rsPONo, KeyAscii, "name"
    Case PNo
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "CODE"
    Case PName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "name"
    Case LName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "Lname"
    Case PQty, DQty
        NumPress TxtGrid(Index), KeyAscii, 8, 3
    Case DisPer, DisOrd
        NumPress TxtGrid(Index), KeyAscii, 2, 2
    Case DisRs, DisOrdRs
        NumPress TxtGrid(Index), KeyAscii, 8, 2
    Case FRate
        NumPress TxtGrid(Index), KeyAscii, 8, 4
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
    Case PONo And txt(ChlType) = "Purchase Receipt"
        If KeyCode <> 13 And DGPONo.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, rsPONo, KeyCode, "name", True
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
    Case Taxable
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            TxtGrid(Index) = ""
        ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Or Trim(TxtGrid(Index)) = "" Then
            TxtGrid(Index) = "Yes"
        Else
            TxtGrid(Index) = "No"
        End If
    Case MRP
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            TxtGrid(Index) = ""
        ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
            TxtGrid(Index) = "Yes"
        Else
            TxtGrid(Index) = "No"
        End If
        
    Case FRate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.0000")
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
    Case PQty, DQty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.000")
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
        If FGrid.Col = DQty Then: FGrid.TextMatrix(FGrid.Row, PQty) = FGrid.TextMatrix(FGrid.Row, DQty)
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
        
    Case TaxPer
        If Val(TxtGrid(Index)) <> 0 Then
            FGrid.TextMatrix(FGrid.Row, TaxPer) = Format(TxtGrid(Index).TEXT, "0.00")
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
        Else
             FGrid.TextMatrix(FGrid.Row, TaxPer) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
        End If
        
    Case SatPer
        If Val(TxtGrid(Index)) <> 0 Then
            FGrid.TextMatrix(FGrid.Row, SatPer) = Format(TxtGrid(Index).TEXT, "0.00")
            FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, SatPer) = ""
            FGrid.TextMatrix(FGrid.Row, SatAmt1) = ""
        End If
        
    Case PartSrlNo
        FGrid.TextMatrix(FGrid.Row, PartSrlNo) = TxtGrid(Index)
End Select
Amt_Cal1
Amt_Cal
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
'Dim j As Integer
Dim TmpRst As ADODB.Recordset
Select Case FGrid.Col
    Case PONo
        If ChkDuplicate = False Then Exit Function
        TxtGridValid_PONo
    Case PNo, PName, LName
        If ChkDuplicate = False Then Exit Function
        TxtGridValid_PNo
    
    Case MRP, Taxable
        If ChkDuplicate = False Then Exit Function
        TxtGridValid_TaxMRP
        'Nra Modi for NDP Stock transfer
        If txt(ChlType) = "Stock Transfer" Then
        Set TmpRst = GCn.Execute("Select PurcDisc_Per from Part_DiscFactor Left Join Part on Part.Disc_Factor=Part_DiscFactor.DiscFac_Catg where Part.Part_No='" & RsPart!Code & "'")
        If TmpRst.RecordCount > 0 Then
              If VNull(TmpRst!PurcDisc_Per) > 0 Then
                   FGrid.TextMatrix(FGrid.Row, FRate) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) - (Val(FGrid.TextMatrix(FGrid.Row, FRate)) * VNull(TmpRst!PurcDisc_Per) / 100), "0.00")
              End If
        End If
        End If
        '***************************
        Amt_Cal1
        Amt_Cal
    
    Case Godown
        TxtGridValid_Godown
        
    Case FRate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(Index) = "", "", Format(Val(TxtGrid(Index).TEXT), "0.0000"))
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
        Amt_Cal1
        Amt_Cal
        
    Case DQty, PQty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(Index) = "", "", Format(Val(TxtGrid(Index).TEXT), "0.000"))
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
        Amt_Cal1
        Amt_Cal
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
        Amt_Cal
    Case DisRs
        FGrid.TextMatrix(FGrid.Row, DisRs) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisRs)) + Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) > Val(FGrid.TextMatrix(FGrid.Row, Amt)) Then
            TxtGridLeave = False: Exit Function
        End If
        Amt_Cal1
        Amt_Cal
    Case DisOrd
        FGrid.TextMatrix(FGrid.Row, DisOrd) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) = 0 Then FGrid.TextMatrix(FGrid.Row, DisOrdRs) = 0
        Amt_Cal1
        Amt_Cal
    Case DisOrdRs
        FGrid.TextMatrix(FGrid.Row, DisOrdRs) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisRs)) + Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) > Val(FGrid.TextMatrix(FGrid.Row, Amt)) Then
'            TxtGridLeave = False: Exit Function
            Exit Function
        End If
        Amt_Cal1
        Amt_Cal
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
        Col4 = FGrid.Col
        Col1 = PONo
        Col2 = Taxable
        Col3 = MRP
    Case MRP
        Col1 = PNo
        Col2 = PONo
        Col3 = Taxable
        Col4 = MRP
    Case Taxable
        Col1 = PNo
        Col2 = PONo
        Col4 = Taxable
        Col3 = MRP
    Case PONo
        Col1 = PNo
        Col4 = PONo
        Col2 = Taxable
        Col3 = MRP
    End Select
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col2))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col3))) + CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))) + CStr(Trim(FGrid.TextMatrix(I, Col3))) + CStr(Trim(FGrid.TextMatrix(I, Col4))))
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
Rst.Open "Select Order_DocID,Order_Srl_No,Qty_Iss From SP_Stock Where DocId='" & txt(TxtDocID).TEXT & "'", GCn, adOpenDynamic, adLockOptimistic
If Rst.RecordCount > 0 Then
    While Not Rst.EOF
        If Rst!Order_DocId <> "" Then
            GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty-" & Rst!Qty_Iss & " Where OrderId='" & Rst!Order_DocId & "' and Srl_No=" & VNull(Rst!Order_Srl_No) & "" 'Part_No='" & Rst!Part_No & "'"
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

Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case FromVno, ToVno
        RsVno.Close
        RsVno.Open "Select V_No as code from Sp_Purch where left(DocId,1) = '" & PubDivCode & "' and Sp_Purch.V_Type='" & ChalVType & "'", GCn, adOpenDynamic, adLockOptimistic
        Set DGVno.DataSource = RsVno
        If txtPrint(Index).TEXT <> RsVno!Code Then
            RsVno.MoveFirst
            RsVno.FIND "code ='" & txtPrint(Index).TEXT & "'"
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
    Case FromVno, ToVno
        DGridTxtKeyDown DGVno, txtPrint, Index, RsVno, KeyCode, False, 0
End Select
If DGVno.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> FromVno Then Ctrl_UpKeyDown KeyCode, Shift
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
End Select
End Sub


Private Sub TxtGridValid_PONo()
If rsPONo.RecordCount = 0 Or (rsPONo.EOF = True Or rsPONo.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, PONo) = ""
    FGrid.TextMatrix(FGrid.Row, PONOCode) = ""
Else
    FGrid.TextMatrix(FGrid.Row, PONo) = rsPONo!Name
    FGrid.TextMatrix(FGrid.Row, PONOCode) = rsPONo!Code
End If
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
'******************** For Tax in Line File *************************
If PubVATYN = 1 Then
   If txt(FormType).Tag <> "" Then
        Set rsTaxPer = GCn.Execute("Select Tax_Per, AddTaxPer, L_C,SFCPER from TaxForms where Form_Code='" & txt(FormType).Tag & "'")
         If rsTaxPer.RecordCount > 0 Then
            FGrid.TextMatrix(FGrid.Row, TaxPer) = rsTaxPer!Tax_Per
            FGrid.TextMatrix(FGrid.Row, SatPer) = VNull(rsTaxPer!AddTaxPer)
                 FGrid.TextMatrix(FGrid.Row, SFCPer) = VNull(rsTaxPer!SFCPer)
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
'*******************************************************************
End If
If FGrid.TextMatrix(FGrid.Row, PONOCode) <> "" Then
    GSQL = "Select s1.Srl_No From SP_Order1 S1 Where OrderID='" & FGrid.TextMatrix(FGrid.Row, PONOCode) & "' and Part_No='" & FGrid.TextMatrix(FGrid.Row, PNo) & "'"
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    If GRs.RecordCount > 1 Or GRs.RecordCount <= 0 Then
        Set GRs = New ADODB.Recordset
        GRs.CursorLocation = adUseClient
        GRs.Open "Select s1.Srl_No,s1.PART_NO,P.Part_Name,s1.QTY,s1.Sup_Qty,(s1.Qty-s1.Sup_Qty) As PendQty From SP_Order1 S1 Left Join Part P on S1.Part_no=P.Part_No and P.Div_Code = left(S1.Orderid,1) Where OrderID='" & FGrid.TextMatrix(FGrid.Row, PONOCode) & "'", GCn, adOpenStatic, adLockReadOnly
        Set DGOrdPart.DataSource = GRs
        GRs.FIND ("Part_No='" & FGrid.TextMatrix(FGrid.Row, PNo) & "'")
        If GRs.EOF Then
            GRs.MoveFirst
        End If
        DGOrdPart.Visible = True
        DGOrdPart.ZOrder 0
        DGOrdPart.SetFocus
    Else
        FGrid.TextMatrix(FGrid.Row, POSrlNo) = GRs!Srl_No
        Set GRs = Nothing
        FGrid.SetFocus
        DGOrdPart.Visible = False
    End If
End If
If FGrid.TextMatrix(FGrid.Rows - 1, PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

Private Sub TxtGridValid_Godown()
    If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or TxtGrid(0).TEXT = "" Then
        FGrid.TextMatrix(FGrid.Row, Godown) = ""
        FGrid.TextMatrix(FGrid.Row, God) = ""
    Else
        FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
        FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
    End If
End Sub

Private Sub TxtGridValid_TaxMRP()
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
        FGrid.TextMatrix(FGrid.Row, FRate) = Format(GetRate(mPartyType, FGrid, CDate(txt(VDate)), FGrid.TextMatrix(FGrid.Row, PNo), MRP, Val(FGrid.TextMatrix(FGrid.Row, MRPRate)), Taxable, Val(FGrid.TextMatrix(FGrid.Row, TBRate)), Val(FGrid.TextMatrix(FGrid.Row, TPRate)), EffectDate, MRPRate), "0.0000")
    End If
'    Amt_Cal
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
    mQry = "SELECT SP.DocID,SP.V_Type,SP.V_No,SP.V_Date,SP.Cash_Credit,SP.L_C,SP.Party_Code,SG.NamePrefix,SP.Party_Name," & _
        "SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone,SP.Form_Code,SP.Party_Doc_No,SP.Party_Doc_Date," & _
        "SP.GR_RR_No,SP.GR_RR_Date,SP.RoadPermit_No,SP.Tot_No_of_Items,SP.Tot_Doc_Qty,SP.Tot_Phy_Qty," & _
        "SP.OilAmt,SP.SprAmt,SP.Tot_Amt,SP.Tot_Disc_Amt,SP.Tot_Ord_DiscAmt,SP.Tot_Goods_Value,SP.Tax_Amt," & _
        "SP.Addition,SP.Deduction,SP.Net_Amt,SP.Case_No,SP.Case_Mark,SP.Transport," & _
        "SP.Supply_Mode,SP.Invoice_DocID,SP.Printed_YN,SP.CancelYN,SP.U_Name,SP.U_EntDt," & _
        "Stk.Srl_No,Part.Part_Name,Stk.Part_No,Stk.Qty_Rec,Stk.Qty_Iss,Stk.Qty_Ret,Stk.Rate," & _
        "Stk.Disc_Per,Stk.Disc_Amt,Stk.Ord_DiscAmt,Stk.Net_Amt as INetAmt,Stk.TaxPer,Stk.TaxAmt, Stk.SatPer, Stk.SatAmt, Sp.SatAmt As SatAmt_H,Stk.SFCPer,Stk.SFCAMt " & _
    "FROM ((SP_Purch SP LEFT JOIN SP_Stock Stk ON SP.DocID = Stk.DocID) " & _
        "LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1)) " & _
        "LEFT JOIN (SubGroup SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON SP.Party_Code = SG.SubCode " & _
    "where SP.docid = '" & Master!SearchCode & "'"
Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "SpMtrlRect", "SpMtrlRect")
        Call WindowsPrint(Index, mQry)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(mQry)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "SpMtrlRect", "SpMtrlRect")
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
Dim Rst As ADODB.Recordset
Dim I As Integer, RST1 As ADODB.Recordset, Rst2 As ADODB.Recordset
On Error GoTo ERRORHANDLER
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
        CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
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
                    rpt.FormulaFields(I).TEXT = "'" & ChalVType & "'"
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
    Dim I As Integer, j As Integer
    Dim PrintStr As String
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstPurChl As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mQty As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    
    Set RstPurChl = GCn.Execute(mQry)
    If RstPurChl.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 0
    
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 17
    mFooter = mFooter + FooterCnt
    
      
    'Sale Bill Header
      
    mDocStr = IIf(mVType = ChalVType, "MATERIAL RECEIPT", "STOCK TRANSFER")
    mDupStr = IIf(RstPurChl!Printed_YN = 1, "(DUPLICATE)", "")
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")

    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!S_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
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
    Print #1, PSTR(RstPurChl!NamePrefix & " " & RstPurChl!Party_Name, 40) & Space(1) & PSTR("MRN No.", 13) & " : " & PrinID(RstPurChl!DocID) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstPurChl!Add1), 40) & Space(1) & mEmph & PSTR("MRN Date", 13) & " : " & PSTR(STR(RstPurChl!V_DATE), 14) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstPurChl!Add2), 40) & Space(7) & PSTR("Party Document No.", 20) & " : " & RstPurChl!Party_Doc_No
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstPurChl!Add3), 40) & Space(7) & PSTR("Party Document Date", 20) & " : " & IIf(IsNull(RstPurChl!Party_Doc_Date), " ", RstPurChl!Party_Doc_Date)
    mHeader = mHeader + 1
    Print #1, XNull(RstPurChl!CityName)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
    mHeader = mHeader + 1
    If StrCmp(left(PubComp_Name, 3), "jmk") Then
        Print #1, mChr17 & PSTR("SRL.No", 7) & PSTR("PART NO.", 21) & PSTR("DESCRIPTION", 33) & PSTR("QUANTITY", 10, , AlignRight) & PSTR("RATE", 9, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & Space(1) & PSTR("DISC %", 6, , AlignRight) & PSTR("DISC.AMT", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & mChr18
    Else
        Print #1, mChr17 & PSTR("SRL.No", 7) & PSTR("PART NO.", 16) & PSTR("DESCRIPTION", 28) & PSTR("QUANTITY", 10, , AlignRight) & PSTR("RATE", 9, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & PSTR("TAXAMT", 10, , AlignRight) & Space(1) & PSTR("DISC %", 6, , AlignRight) & PSTR("DISC.AMT", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & mChr18
    End If
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    If RstPurChl.RecordCount > 0 Then
        Do Until RstPurChl.EOF
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
                Print #1, PSTR("M/s " & RstPurChl!Party_Name, 40) & Space(7) & PSTR("Challan No.", 13) & " : " & PSTR(STR(RstPurChl!V_NO), 14) & mEmph1
                mHeader = mHeader + 1
                Print #1, PSTR(XNull(RstPurChl!Add1), 40) & Space(7) & mEmph & PSTR("Challan Date", 13) & " : " & PSTR(STR(RstPurChl!V_DATE), 14) & mEmph1
                mHeader = mHeader + 1
                Print #1, PSTR(XNull(RstPurChl!Add2), 40) & Space(7) & PSTR("Party Document No.", 20) & " : " & XNull(RstPurChl!Party_Doc_No)
                mHeader = mHeader + 1
                Print #1, PSTR(XNull(RstPurChl!Add3), 40) & Space(7) & PSTR("Party Document Date", 20) & " : " & IIf(IsNull(RstPurChl!Party_Doc_Date), "", RstPurChl!Party_Doc_Date)
                mHeader = mHeader + 1
                Print #1, XNull(RstPurChl!CityName)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                If StrCmp(left(PubComp_Name, 3), "jmk") Then
                    Print #1, mChr17 & PSTR("SRL.No", 7) & PSTR("PART NO.", 15) & PSTR("DESCRIPTION", 33) & PSTR("QUANTITY", 10, , AlignRight) & PSTR("RATE", 9, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & Space(1) & PSTR("DISC %", 6, , AlignRight) & PSTR("DICS.AMT", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & mChr18
                Else
                    Print #1, mChr17 & PSTR("SRL.No", 7) & PSTR("PART NO.", 16) & PSTR("DESCRIPTION", 28) & PSTR("QUANTITY", 10, , AlignRight) & PSTR("RATE", 9, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & PSTR("TAXAMT", 10, , AlignRight) & Space(1) & PSTR("DISC %", 6, , AlignRight) & PSTR("DICS.AMT", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & mChr18
                End If
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            If StrCmp(left(PubComp_Name, 3), "jmk") Then
                PrintStr = mChr17 & PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstPurChl!Part_No, 21, , AlignLeft) & PSTR(RstPurChl!Part_Name, 33) & PSTR(RstPurChl!Qty_Rec, 10, 3) & PSTR(RstPurChl!Rate, 9, 2) & PSTR(RstPurChl!TaxPer, 6, 2) & Space(1) & PSTR(RstPurChl!Disc_Per, 6, 2) & PSTR(RstPurChl!Disc_Amt, 10, 2) & PSTR(RstPurChl!INetAmt, 10, 2) & mChr18
            Else
                PrintStr = mChr17 & PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstPurChl!Part_No, 16, , AlignLeft) & PSTR(RstPurChl!Part_Name, 28) & PSTR(RstPurChl!Qty_Rec, 10, 3) & PSTR(RstPurChl!Rate, 9, 2) & PSTR(RstPurChl!TaxPer, 6, 2) & PSTR(RstPurChl!TaxAmt, 10, 2) & Space(1) & PSTR(RstPurChl!Disc_Per, 6, 2) & PSTR(RstPurChl!Disc_Amt, 10, 2) & PSTR(RstPurChl!INetAmt, 10, 2) & mChr18
            End If
            mQty = mQty + RstPurChl!Qty_Rec: mAmount = mAmount + RstPurChl!INetAmt
            Print #1, PrintStr
            RstPurChl.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    RstPurChl.MoveFirst
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, PSTR("Total  > > ", 51, , AlignRight) & PSTR(mQty, 10, 3) & Space(9) & PSTR(mAmount, 10, 2)
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, PSTR("GR No", 16) & ": " & PSTR(XNull(RstPurChl!GR_RR_No), 20) & " | " & PSTR("Total Goods Value", 22) & " : " & PSTR(RstPurChl!Tot_Goods_Value, 13, 2)
    Print #1, PSTR("GR Date", 16) & ": " & PSTR(XNull(RstPurChl!GR_RR_Date), 20) & " | " & PSTR("Discount", 22) & " : " & PSTR(RstPurChl!Tot_Disc_Amt, 13, 2)
    Print #1, PSTR("Case Mark No.", 16) & ": " & PSTR(XNull(RstPurChl!Case_Mark), 20) & " | " & PSTR("Order Discount Amount", 22) & " : " & PSTR(RstPurChl!Tot_Ord_DiscAmt, 13, 2)
    Print #1, PSTR("Road Permit No", 16) & ": " & PSTR(XNull(RstPurChl!RoadPermit_No), 20) & " | " & IIf(PubVATYN = 1 And Not StrCmp(left(PubComp_Name, 3), "jmk"), PSTR("V A T", 22), PSTR("Tax Amount", 22)) & " : " & PSTR(RstPurChl!Tax_Amt, 13, 2)
    Print #1, Space(36) & " | " & PSTR("S A T", 22) & " : " & PSTR(RstPurChl!SatAmt_H, 13, 2)
    Print #1, PSTR("TransPort", 16) & ": " & PSTR(XNull(RstPurChl!Transport), 20) & " | " & PSTR("Addition", 22) & " : " & PSTR(RstPurChl!Addition, 13, 2)
    Print #1, PSTR("Mode Of Dispatch", 16) & ": " & PSTR(XNull(RstPurChl!Supply_Mode), 20) & " | " & PSTR("Deduction", 22) & " : " & PSTR(RstPurChl!Deduction, 13, 2)
    Print #1, Space(38) & " | " & mEmph & PSTR("Net Payble Rs.", 22) & " : " & PSTR(RstPurChl!Net_Amt, 13, 2) & mEmph1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(RstPurChl!Net_Amt, "Rupees", "Paise") & mDoub1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, ""
    Print #1, "Goods Receipt  By" & Space(15) & "Checked By" & Space(15) & PSTR("Store Incharge", PageWidth - 56, , AlignRight) & mChr17
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
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
        GCn.Execute "update Sp_purch set Printed_YN = 1 where Sp_purch.docid='" & Master!SearchCode & "'"
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
                .TextMatrix(0, SatPer) = "SAT %"
                .ColAlignmentFixed(SatPer) = flexAlignRightCenter
                .ColWidth(SatPer) = 840
                
                .TextMatrix(0, SatAmt1) = "SAT Amt"
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
