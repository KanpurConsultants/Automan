VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmEstimateQuot 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Estimate/Quotation Entry"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12300
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
   ScaleHeight     =   7230
   ScaleWidth      =   12300
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
      Height          =   1890
      Left            =   2475
      TabIndex        =   180
      Top             =   2175
      Visible         =   0   'False
      Width           =   5505
      Begin VB.ComboBox CmboPLNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2295
         TabIndex        =   209
         Text            =   "Combo1"
         Top             =   630
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox TxtPerformaNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Height          =   315
         Left            =   3900
         TabIndex        =   207
         Top             =   300
         Width           =   1545
      End
      Begin VB.CheckBox ChkMerg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAECF0&
         Caption         =   "Merge Labour Performa"
         Height          =   240
         Left            =   90
         TabIndex        =   198
         Top             =   360
         Width           =   2610
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
         Left            =   5175
         MousePointer    =   99  'Custom
         Picture         =   "frmEstimateQuot.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   190
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
         Picture         =   "frmEstimateQuot.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   189
         ToolTipText     =   "Screen"
         Top             =   1545
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmEstimateQuot.frx":0678
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
         Left            =   4050
         MaskColor       =   &H00FFC0FF&
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Printer "
         Top             =   1545
         UseMaskColor    =   -1  'True
         Width           =   1425
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmEstimateQuot.frx":0982
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
         Left            =   4050
         MaskColor       =   &H00EFD5B8&
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Screen"
         Top             =   1215
         Width           =   1425
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmEstimateQuot.frx":0C8C
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
         Left            =   4050
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "Printer "
         Top             =   885
         Width           =   1425
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
         TabIndex        =   185
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
         TabIndex        =   184
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
         TabIndex        =   183
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
         Left            =   1440
         TabIndex        =   182
         Top             =   1290
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
         Left            =   15
         TabIndex        =   181
         Top             =   1290
         Width           =   750
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Performa No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   49
         Left            =   1140
         TabIndex        =   210
         Top             =   645
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Performa No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   47
         Left            =   2745
         TabIndex        =   206
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Line Line8 
         X1              =   1185
         X2              =   1185
         Y1              =   1080
         Y2              =   1170
      End
      Begin VB.Line Line7 
         X1              =   2565
         X2              =   2565
         Y1              =   1200
         Y2              =   1305
      End
      Begin VB.Line Line5 
         X1              =   75
         X2              =   75
         Y1              =   1200
         Y2              =   1305
      End
      Begin VB.Line Line6 
         X1              =   2535
         X2              =   60
         Y1              =   1185
         Y2              =   1185
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
         Left            =   -435
         TabIndex        =   193
         Top             =   885
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
         TabIndex        =   192
         Top             =   1545
         Width           =   5145
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
         Left            =   30
         TabIndex        =   191
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.TextBox Txt 
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
      Index           =   59
      Left            =   10245
      TabIndex        =   211
      Text            =   "3"
      Top             =   4800
      Width           =   1215
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
      Index           =   58
      Left            =   7440
      MaxLength       =   20
      TabIndex        =   18
      Top             =   1620
      Visible         =   0   'False
      Width           =   720
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
      Height          =   225
      Index           =   57
      Left            =   7980
      MaxLength       =   50
      TabIndex        =   20
      Top             =   1935
      Width           =   3660
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
      Index           =   56
      Left            =   6225
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1620
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   55
      Left            =   6360
      TabIndex        =   201
      Text            =   "00.00"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   270
      Index           =   54
      Left            =   6225
      MaxLength       =   14
      TabIndex        =   14
      Top             =   765
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   2730
      Left            =   7995
      Negotiate       =   -1  'True
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   6915
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "name"
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
            DividerStyle    =   3
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGHist 
      Height          =   2520
      Left            =   6645
      Negotiate       =   -1  'True
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   7110
      Visible         =   0   'False
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   4445
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "RegNo"
         Caption         =   "Reg. No."
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
         DataField       =   "Chassis"
         Caption         =   "Chassis No."
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
         DataField       =   "Name"
         Caption         =   "Owner Name"
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
         DataField       =   "PhoneOff"
         Caption         =   "Phone (O)"
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
         DataField       =   "Govt"
         Caption         =   "Govt"
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
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPartyType 
      Height          =   2295
      Left            =   -2280
      Negotiate       =   -1  'True
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4048
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
         Caption         =   "Party Type"
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
            DividerStyle    =   3
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
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
      Height          =   270
      Index           =   51
      Left            =   1140
      TabIndex        =   13
      Top             =   1905
      Width           =   2040
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   -3510
      Negotiate       =   -1  'True
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   2520
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
         Caption         =   "Voucher No."
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
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   1590
      Negotiate       =   -1  'True
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   7170
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
      Left            =   -6045
      TabIndex        =   143
      Top             =   3705
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
         TabIndex        =   174
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         TabIndex        =   170
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
         TabIndex        =   169
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
         TabIndex        =   168
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
         TabIndex        =   167
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
         TabIndex        =   166
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
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
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
         TabIndex        =   161
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   144
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
   Begin MSDataGridLib.DataGrid DGJCNo 
      Height          =   3330
      Left            =   9270
      Negotiate       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7050
      Visible         =   0   'False
      Width           =   8250
      _ExtentX        =   14552
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Job Card No"
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
         DataField       =   "Reg_No"
         Caption         =   "Registration No"
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
         DataField       =   "ChassisNo"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "Job_Date"
         Caption         =   "Open Date"
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
         DataField       =   "Party"
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
      BeginProperty Column06 
         DataField       =   "Address1"
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
      BeginProperty Column07 
         DataField       =   "Address2"
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
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2294.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
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
      Height          =   270
      Index           =   3
      Left            =   10770
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1590
      Width           =   825
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
      Height          =   270
      Index           =   8
      Left            =   6225
      MaxLength       =   20
      TabIndex        =   15
      Text            =   "012345678901234"
      Top             =   1050
      Width           =   1935
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
      Height          =   270
      Index           =   2
      Left            =   10125
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1305
      Width           =   1470
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
      Height          =   270
      Index           =   9
      Left            =   6225
      MaxLength       =   25
      TabIndex        =   16
      Top             =   1335
      Width           =   1935
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
      Height          =   270
      Index           =   7
      Left            =   6225
      MaxLength       =   14
      TabIndex        =   8
      Top             =   480
      Width           =   1515
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
      Height          =   270
      Index           =   6
      Left            =   3690
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   480
      Width           =   1275
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
      Height          =   270
      Index           =   5
      Left            =   1305
      TabIndex        =   7
      Top             =   480
      Width           =   975
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
      Height          =   270
      Index           =   12
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1335
      Width           =   3945
   End
   Begin VB.TextBox Txt 
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
      Height          =   270
      Index           =   1
      Left            =   10125
      MaxLength       =   8
      TabIndex        =   2
      ToolTipText     =   "Press S-> Stores or W-> Workshop"
      Top             =   1020
      Width           =   1470
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
      Height          =   270
      Index           =   10
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   9
      Top             =   765
      Width           =   3945
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
      Height          =   270
      Index           =   0
      Left            =   9285
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2310
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
      Height          =   270
      Index           =   11
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1050
      Width           =   3945
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   42
      Left            =   5475
      TabIndex        =   62
      Top             =   6195
      Width           =   1440
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   31
      Left            =   10245
      TabIndex        =   42
      Top             =   4260
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   48
      Left            =   5715
      TabIndex        =   61
      Text            =   "999999.99"
      Top             =   6855
      Width           =   1020
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   41
      Left            =   10245
      TabIndex        =   54
      Top             =   6420
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   16
      Left            =   5115
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4530
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   15
      Left            =   2985
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4530
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   14
      Left            =   5115
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1215
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
      Left            =   2445
      TabIndex        =   23
      Top             =   2700
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Txt 
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
      Index           =   46
      Left            =   4245
      TabIndex        =   59
      Top             =   6585
      Width           =   1020
   End
   Begin VB.TextBox Txt 
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
      Index           =   45
      Left            =   3585
      TabIndex        =   58
      ToolTipText     =   "Service Tax %"
      Top             =   6600
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Index           =   44
      Left            =   1170
      TabIndex        =   57
      Top             =   6855
      Width           =   1020
   End
   Begin VB.TextBox Txt 
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
      Index           =   43
      Left            =   1170
      TabIndex        =   56
      Text            =   "999999.99"
      Top             =   6585
      Width           =   1020
   End
   Begin VB.TextBox Txt 
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
      Index           =   40
      Left            =   10245
      TabIndex        =   53
      Top             =   6150
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   39
      Left            =   10245
      TabIndex        =   50
      Top             =   5610
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   38
      Left            =   9570
      TabIndex        =   49
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5610
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Index           =   47
      Left            =   4245
      TabIndex        =   60
      Top             =   6855
      Width           =   1020
   End
   Begin VB.TextBox Txt 
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
      Index           =   37
      Left            =   10245
      TabIndex        =   48
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   35
      Left            =   11655
      TabIndex        =   46
      Text            =   "3"
      Top             =   6750
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   34
      Left            =   10980
      TabIndex        =   45
      Text            =   "33"
      ToolTipText     =   "Surcharge % on Local Sales Tax"
      Top             =   6750
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Index           =   33
      Left            =   10245
      TabIndex        =   44
      Text            =   "32"
      Top             =   4530
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   32
      Left            =   9570
      TabIndex        =   43
      ToolTipText     =   "Local Sales Tax %"
      Top             =   4530
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Index           =   29
      Left            =   2985
      TabIndex        =   41
      Top             =   6150
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Left            =   2985
      TabIndex        =   40
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   27
      Left            =   2295
      TabIndex        =   39
      ToolTipText     =   "General Surcharge %"
      Top             =   5880
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Index           =   13
      Left            =   2985
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   36
      Left            =   10245
      TabIndex        =   47
      Top             =   5070
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   30
      Left            =   7650
      TabIndex        =   55
      Top             =   6795
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   26
      Left            =   5115
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5610
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   25
      Left            =   2985
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "99999999.99"
      Top             =   5610
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   24
      Left            =   5115
      TabIndex        =   36
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   23
      Left            =   4425
      TabIndex        =   35
      ToolTipText     =   "Discount % Taxpaid"
      Top             =   5340
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Index           =   22
      Left            =   2985
      TabIndex        =   34
      Text            =   "99999999.99"
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   21
      Left            =   2295
      TabIndex        =   33
      Text            =   "99.99"
      ToolTipText     =   "Discount % Taxable"
      Top             =   5340
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Index           =   20
      Left            =   5115
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Left            =   2985
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   18
      Left            =   5115
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   17
      Left            =   2985
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1785
      Left            =   15
      TabIndex        =   24
      Top             =   2235
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   3149
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   30
      BackColorFixed  =   14940925
      ForeColorFixed  =   8388608
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   14737632
      BackColorUnpopulated=   16777215
      GridColor       =   12640511
      GridColorFixed  =   0
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "hhh"
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
      _Band(0).Cols   =   30
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
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
      Index           =   50
      Left            =   10245
      TabIndex        =   52
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   49
      Left            =   9570
      TabIndex        =   51
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5880
      Width           =   570
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   4
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   6
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   165
      Visible         =   0   'False
      Width           =   555
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
      Height          =   270
      Index           =   52
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1620
      Width           =   2685
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
      Height          =   270
      Index           =   53
      Left            =   4230
      MaxLength       =   40
      TabIndex        =   19
      Top             =   1890
      Width           =   2670
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax"
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
      Index           =   50
      Left            =   7575
      TabIndex        =   212
      Top             =   4800
      Width           =   1140
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Req.No."
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
      Index           =   48
      Left            =   6810
      TabIndex        =   208
      Top             =   1635
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
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
      Index           =   46
      Left            =   7005
      TabIndex        =   205
      Top             =   1935
      Width           =   855
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
      Height          =   270
      Index           =   34
      Left            =   6135
      TabIndex        =   204
      Top             =   1620
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplimentary"
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
      Left            =   4785
      TabIndex        =   203
      Top             =   1605
      Width           =   1185
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out Side Lab"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   44
      Left            =   6540
      TabIndex        =   202
      Top             =   4905
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
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
      Index           =   43
      Left            =   5130
      TabIndex        =   200
      Top             =   765
      Width           =   495
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
      Height          =   270
      Index           =   33
      Left            =   6135
      TabIndex        =   199
      Top             =   765
      Width           =   195
   End
   Begin VB.Line Line9 
      X1              =   765
      X2              =   7065
      Y1              =   6540
      Y2              =   6540
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print Title :"
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
      Index           =   42
      Left            =   3240
      TabIndex        =   197
      Top             =   1935
      Width           =   900
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City           :"
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
      Index           =   40
      Left            =   120
      TabIndex        =   194
      Top             =   1650
      Width           =   840
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Type"
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
      Height          =   270
      Index           =   38
      Left            =   120
      TabIndex        =   178
      Top             =   1905
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   16
      Left            =   1035
      TabIndex        =   177
      Top             =   1905
      Width           =   195
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   2
      Left            =   8535
      TabIndex        =   141
      Top             =   1590
      Width           =   810
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   8535
      TabIndex        =   140
      Top             =   1305
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   93
      Left            =   10005
      TabIndex        =   139
      Top             =   1305
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   92
      Left            =   10005
      TabIndex        =   138
      Top             =   1590
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   90
      Left            =   10005
      TabIndex        =   137
      Top             =   1020
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
      Height          =   270
      Index           =   1
      Left            =   6135
      TabIndex        =   136
      Top             =   1050
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   7
      Left            =   5100
      TabIndex        =   135
      Top             =   1050
      Width           =   1035
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
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
      Height          =   270
      Index           =   8
      Left            =   5100
      TabIndex        =   134
      Top             =   1335
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
      Height          =   270
      Index           =   2
      Left            =   6135
      TabIndex        =   133
      Top             =   1335
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   5
      Left            =   6135
      TabIndex        =   132
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. No."
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
      Left            =   5160
      TabIndex        =   131
      Top             =   480
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   4
      Left            =   3600
      TabIndex        =   130
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Date"
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
      Height          =   270
      Index           =   5
      Left            =   2655
      TabIndex        =   129
      Top             =   480
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   31
      Left            =   900
      TabIndex        =   128
      Top             =   765
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   1215
      TabIndex        =   127
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   6
      Left            =   120
      TabIndex        =   126
      Top             =   480
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   3
      Left            =   1335
      TabIndex        =   125
      Top             =   240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Card Y/N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   4
      Left            =   180
      TabIndex        =   124
      Top             =   240
      Visible         =   0   'False
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   25
      Left            =   9180
      TabIndex        =   123
      Top             =   510
      Width           =   195
   End
   Begin VB.Label Lbl 
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
      Index           =   34
      Left            =   8535
      TabIndex        =   122
      Top             =   510
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   23
      Left            =   900
      TabIndex        =   121
      Top             =   1050
      Width           =   195
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   10
      Left            =   120
      TabIndex        =   120
      Top             =   1050
      Width           =   690
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party"
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
      Index           =   9
      Left            =   120
      TabIndex        =   119
      Top             =   765
      Width           =   450
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Originated From"
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
      Left            =   8535
      TabIndex        =   118
      Top             =   1020
      Width           =   1365
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
      Left            =   10530
      TabIndex        =   117
      Top             =   765
      Width           =   810
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
      Left            =   8535
      TabIndex        =   116
      Top             =   765
      Width           =   660
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
      Left            =   10125
      TabIndex        =   4
      Top             =   1590
      Width           =   600
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   1485
      Left            =   8430
      Shape           =   4  'Rounded Rectangle
      Top             =   435
      Width           =   3300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2D5C0&
      Caption         =   "Labour"
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
      Index           =   26
      Left            =   180
      TabIndex        =   115
      Top             =   6405
      Width           =   600
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spares Amount"
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
      Index           =   11
      Left            =   180
      TabIndex        =   114
      Top             =   4800
      Width           =   1320
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item-wise Disc Total"
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
      Left            =   180
      TabIndex        =   113
      Top             =   4260
      Width           =   1680
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MRP Item's Amount"
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
      Index           =   33
      Left            =   180
      TabIndex        =   112
      Top             =   4530
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   600
      Left            =   5295
      Shape           =   4  'Rounded Rectangle
      Top             =   5910
      Width           =   1770
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET PAYABLE AMT"
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
      Height          =   210
      Index           =   39
      Left            =   5385
      TabIndex        =   96
      Top             =   5910
      Width           =   1590
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAXABLE TOTAL"
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
      Index           =   36
      Left            =   7575
      TabIndex        =   111
      Top             =   4260
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   32
      Left            =   9105
      TabIndex        =   110
      Top             =   4260
      Width           =   210
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET LABOUR AMT"
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
      Left            =   5445
      TabIndex        =   109
      Top             =   6570
      Width           =   1515
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   30
      Left            =   2385
      TabIndex        =   108
      Top             =   6870
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   22
      Left            =   3420
      TabIndex        =   107
      Top             =   6870
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET SPARE AMT"
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
      Index           =   32
      Left            =   7575
      TabIndex        =   106
      Top             =   6420
      Width           =   1380
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
      Index           =   26
      Left            =   9105
      TabIndex        =   105
      Top             =   6420
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
      Index           =   30
      Left            =   1980
      TabIndex        =   104
      Top             =   4530
      Width           =   210
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
      Index           =   29
      Left            =   8340
      TabIndex        =   103
      Top             =   4020
      Width           =   45
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
      Height          =   255
      Index           =   7
      Left            =   7050
      TabIndex        =   102
      Top             =   4020
      Width           =   1470
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
      Height          =   255
      Left            =   11100
      TabIndex        =   101
      Top             =   4020
      Width           =   465
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
      Height          =   255
      Left            =   9135
      TabIndex        =   100
      Top             =   4020
      Width           =   105
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   25
      Left            =   9405
      TabIndex        =   99
      Top             =   4020
      Width           =   1170
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
      Index           =   28
      Left            =   8625
      TabIndex        =   98
      Top             =   4020
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
      Index           =   27
      Left            =   10680
      TabIndex        =   97
      Top             =   4020
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
      Index           =   21
      Left            =   3420
      TabIndex        =   95
      Top             =   6600
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax"
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
      Index           =   29
      Left            =   2385
      TabIndex        =   94
      Top             =   6600
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
      Height          =   255
      Index           =   20
      Left            =   1020
      TabIndex        =   93
      Top             =   6855
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   28
      Left            =   180
      TabIndex        =   92
      Top             =   6855
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   19
      Left            =   1020
      TabIndex        =   91
      Top             =   6585
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Index           =   27
      Left            =   180
      TabIndex        =   90
      Top             =   6585
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   18
      Left            =   9105
      TabIndex        =   89
      Top             =   6150
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Round Off"
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
      Index           =   37
      Left            =   7575
      TabIndex        =   88
      Top             =   6150
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   17
      Left            =   9105
      TabIndex        =   87
      Top             =   5610
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOT on Sub Total (B)"
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
      Left            =   7575
      TabIndex        =   86
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5610
      Width           =   1710
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (B) TB+TP"
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
      Index           =   24
      Left            =   7575
      TabIndex        =   85
      Top             =   5340
      Width           =   1680
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
      Index           =   15
      Left            =   10515
      TabIndex        =   84
      Top             =   6750
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surcharge on Tax"
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
      Left            =   8985
      TabIndex        =   83
      Top             =   6750
      Visible         =   0   'False
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   9105
      TabIndex        =   82
      Top             =   4530
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
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
      Left            =   7575
      TabIndex        =   81
      Top             =   4530
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   13
      Left            =   1980
      TabIndex        =   80
      Top             =   6150
      Width           =   210
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   21
      Left            =   180
      TabIndex        =   79
      Top             =   6150
      Width           =   1200
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
      Index           =   12
      Left            =   1980
      TabIndex        =   78
      Top             =   5880
      Width           =   210
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Surcharge"
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
      Left            =   180
      TabIndex        =   77
      Top             =   5880
      Width           =   1560
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
      Index           =   11
      Left            =   1980
      TabIndex        =   76
      Top             =   4260
      Width           =   210
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
      Index           =   10
      Left            =   9105
      TabIndex        =   75
      Top             =   5070
      Width           =   210
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   18
      Left            =   7575
      TabIndex        =   74
      Top             =   5070
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   9
      Left            =   8310
      TabIndex        =   73
      Top             =   6690
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Lbl 
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   17
      Left            =   6795
      TabIndex        =   72
      Top             =   6810
      Visible         =   0   'False
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   1980
      TabIndex        =   71
      Top             =   5610
      Width           =   210
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (A)"
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
      Index           =   16
      Left            =   180
      TabIndex        =   70
      Top             =   5610
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   7
      Left            =   1980
      TabIndex        =   69
      Top             =   5340
      Width           =   210
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   68
      Top             =   5340
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   6
      Left            =   1980
      TabIndex        =   67
      Top             =   5070
      Width           =   210
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oil Amount"
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
      Index           =   14
      Left            =   180
      TabIndex        =   66
      Top             =   5070
      Width           =   930
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Paid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   13
      Left            =   5565
      TabIndex        =   65
      Top             =   4020
      Width           =   765
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   12
      Left            =   3525
      TabIndex        =   64
      Top             =   4020
      Width           =   675
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
      Index           =   24
      Left            =   1980
      TabIndex        =   63
      Top             =   4800
      Width           =   210
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReSale Tax             :"
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
      Index           =   35
      Left            =   7575
      TabIndex        =   176
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5895
      Width           =   1575
   End
End
Attribute VB_Name = "frmEstimateQuot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMRevDisTBPer As Double, mMRevDisTPPer As Double
Dim mTBDisAmtMRP As Double, mTPDisAmtMRP As Double
Dim mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double
Dim mMRPLubeTB As Double, mMRPLubeTP  As Double
Dim mVatYn As Byte
Private Const mSP2 As String = " "
Private FirstPrint As Boolean
Public PubEstimateType$ 'used in MDI
Dim TAddMode As Boolean

Dim mCheckNegetiveStockSiteWise As Boolean

Dim GridKey As Integer
Dim RsCity As ADODB.Recordset
Dim RsHist As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim RsJob As ADODB.Recordset
Dim rsPartyType As ADODB.Recordset
'Dim RsPart As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mDocId As String * 21
Dim mVType As String
Dim mVPrefix As String
Dim mSearchCode As String
'grid color scheme
'Private Const CellBackColLeave As String = &HECE4D7    '&HECE4D7   '&HEDF7FE
'Private Const GridBackColorBkg As String = &HE2D5C0
'Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation
Dim vPrefix As String
Dim OldJCNo As String                           ' Used For Job Card No

Private Const DocID As Byte = 0                 ' Doc.ID
Private Const OrgFrom As Byte = 1               ' Originated From
Private Const VDate As Byte = 2                 ' Date
Private Const SerialNo As Byte = 3              ' Serial No.
Private Const JCYN As Byte = 4                  ' Job Card Required
Private Const JCNo As Byte = 5                  ' Job Card No.
Private Const OpDate As Byte = 6                ' Job CardOpen Date
Private Const RegNo As Byte = 7                 ' Reg. No
Private Const ChassisNo As Byte = 8             ' Chassis No.
Private Const Engno As Byte = 9                 ' Engine No.
Private Const Party As Byte = 10                ' Party Name
Private Const Address1 As Byte = 11             ' Address1
Private Const Address2 As Byte = 12             ' Address2
Private Const IWDiscTotTB As Byte = 13          ' Item-wise Disc Total Taxable
Private Const IWDiscTotTP As Byte = 14          ' Item-wise Disc Total Taxpaid
Private Const MRPAmtTB As Byte = 15             ' MRP Item's Amount Taxable
Private Const MRPAmtTP As Byte = 16             ' MRP Item's Amount Taxpaid
Private Const SprAmtTB As Byte = 17             ' Spares Amount Taxable
Private Const SprAmtTP As Byte = 18             ' Spares Amount Taxpaid
Private Const OilAmtTB As Byte = 19             ' Oil Amount Taxable
Private Const OilAmtTP As Byte = 20             ' Oil Amount Taxpaid
Private Const DiscPerTB As Byte = 21            '
Private Const DiscAmtTB As Byte = 22            '
Private Const DiscPerTP As Byte = 23            '
Private Const DiscAmtTP As Byte = 24            '
Private Const STotATB As Byte = 25              '
Private Const STotATP As Byte = 26              '
Private Const GenSurPer As Byte = 27           '
Private Const GenSurAmt As Byte = 28           '
Private Const TransAmt As Byte = 29             '
Private Const Addition As Byte = 30            '
Private Const TaxableTot As Byte = 31           ' Taxable Total
Private Const STaxPer As Byte = 32              '
Private Const STaxAmt As Byte = 33              '
Private Const TaxSurPer As Byte = 34            '
Private Const TaxSurAmt As Byte = 35            '
Private Const PackCrg As Byte = 36              '
Private Const STotB As Byte = 37                '
Private Const TurnOverPer As Byte = 38          '
Private Const TurnOverAmt As Byte = 39          '
Private Const SROff As Byte = 40                '
Private Const NetSprAmt As Byte = 41            '
Private Const NetAmt As Byte = 42              '
Private Const LabAmt As Byte = 43               '
Private Const LabDisc As Byte = 44              '
Private Const ServTaxPer As Byte = 45         '
Private Const ServTaxAmt As Byte = 46         '
Private Const LabROff As Byte = 47              '
Private Const NetLabAmt As Byte = 48            '
Private Const ReSalTaxPer As Byte = 49          '
Private Const ReSalTaxAmt As Byte = 50          '
Private Const PartyType = 51                    ' Party Type
Private Const City = 52
Private Const PrintTitle = 53
Private Const Model = 54
Private Const OutSideLabAmt = 55
Private Const Suppli = 56
Private Const Remarks = 57
Private Const ReqNo = 58
Private Const SatAmt = 59

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0
Private Const Col_ReqNo As Byte = 1             'Requisition No
Private Const Col_PNo As Byte = 2               ' Part No
Private Const Col_Unit As Byte = 3              ' Unit
Private Const Col_MRP As Byte = 4               ' MRP Yes/No
Private Const Col_Taxable As Byte = 5           ' Taxable Yes/No
Private Const Col_Qty As Byte = 6               ' Qty
Private Const Col_Rate As Byte = 7              ' Rate
Private Const Col_MRPRate As Byte = 8           ' MRP Rate
Private Const Col_Amt As Byte = 9              ' Amt
Private Const Col_DiscPer As Byte = 10          ' Disc. %
Private Const Col_DiscAmt As Byte = 11          ' Disc. Amt.
Private Const Col_ItemVal As Byte = 12          ' Item Value
Private Const Col_PName As Byte = 13            ' Part Name
Private Const Col_LName As Byte = 14            ' Local Name
Private Const Col_MRPStkTP As Byte = 15         ' MRP Qty TB 'Current Stk Qty
Private Const Col_MRPStkTB As Byte = 16         ' MRP Qty TB
Private Const Col_TBStk As Byte = 17            ' Taxbale Qty
Private Const Col_TPStk As Byte = 18            ' Tax Paid Qty
Private Const Col_TBRate As Byte = 19           ' Taxbale Rate
Private Const Col_TPRate As Byte = 20           ' Tax Paid Rate
Private Const Col_Bin As Byte = 21              ' Bin
Private Const Col_LastRate As Byte = 22         ' Last Purchase Rate
Private Const Col_HPRate As Byte = 23           ' High Purchase Rate
Private Const Col_LPRate As Byte = 24           ' Low Purchase Rate
Private Const Col_PartGrade As Byte = 25        ' Part Grade (Used for Oil Item)
Private Const Col_EffectDate As Byte = 26       ' MRP Effective Date/TB Effective Date
Private Const Col_Purpose As Byte = 27  '' New
Private Const Col_TaxPer As Byte = 28
Private Const Col_TaxAmt As Byte = 29
Private Const Col_SatPer As Byte = 30
Private Const Col_SatAmt As Byte = 31


Private Const FromVno As Byte = 0
Private Const ToVno As Byte = 1
Private Const VType1 As Byte = 2

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim rstForm As ADODB.Recordset
Dim Syctrl As ADODB.Recordset

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To txt.Count - 1
        'modishekhar
        If I = DocID Or I = OpDate _
            Or I = IWDiscTotTB Or I = IWDiscTotTP Or I = MRPAmtTB Or I = MRPAmtTP _
            Or I = SprAmtTB Or I = SprAmtTP Or I = OilAmtTB Or I = OilAmtTP Or I = STotATB Or I = STotATP _
            Or I = TaxableTot Or I = NetSprAmt Or I = SROff Or I = LabROff Or I = NetLabAmt _
            Or I = STotB Or I = NetAmt Or I = NetLabAmt Or I = LabROff Then
        Else
            txt(I).Enabled = Enb
        End If
    Next
    
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select DocId As SearchCode From Estimate where left(DocId,1)='" & PubDivCode & "' and V_Type in ('S_QU','W_EST','S_INV') And DocId = '" & MyValue & "'  Order by V_Date desc,V_Type,DocID desc")
    End If
    If Master.EOF = True Then Exit Sub
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
End Sub

'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
    Next I
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
    FrmPrn.left = (Me.width - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
'Serial No  | Part No | Part Name |Unit | MRP Yes/No | Taxable Yes/No  | Qty | Rate | Amt | Disc. % | Disc. Amt. | Item Value | Local Name
    With FGrid
        .left = Me.left '+ 60
        .width = Me.width - 90
        .top = 2190 '1980
'        .BackColor = CellBackColLeave
'        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 32

        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450
        
        .TextMatrix(0, Col_ReqNo) = "Req. No"
        .ColAlignment(Col_ReqNo) = flexAlignRightCenter
        If UCase(left(PubComp_Name, 3)) = "JMK" Then
            .ColWidth(Col_ReqNo) = 1000
        Else
            .ColWidth(Col_ReqNo) = 0
        End If
        
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

        .TextMatrix(0, Col_Qty) = "Qty"
        .ColAlignmentFixed(Col_Qty) = flexAlignRightCenter
        .ColWidth(Col_Qty) = 960

        .TextMatrix(0, Col_Rate) = "Rate"
        .ColAlignmentFixed(Col_Rate) = flexAlignRightCenter
        .ColWidth(Col_Rate) = 870

        .TextMatrix(0, Col_MRPRate) = "MRP Rate"
        .ColAlignmentFixed(Col_MRPRate) = flexAlignRightCenter
        .ColWidth(Col_MRPRate) = 870

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
        .ColWidth(Col_DiscPer) = 555

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
        .ColWidth(Col_DiscAmt) = 840

        .TextMatrix(0, Col_TaxPer) = "Tax %"
        .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
        .ColWidth(Col_TaxPer) = 555

        .TextMatrix(0, Col_TaxAmt) = "Tax Amt"
        .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
        .ColWidth(Col_TaxAmt) = 840

        .TextMatrix(0, Col_SatPer) = "SAT %"
        .ColAlignmentFixed(Col_SatPer) = flexAlignRightCenter
        .ColWidth(Col_SatPer) = 555

        .TextMatrix(0, Col_SatAmt) = "SAT Amt"
        .ColAlignmentFixed(Col_SatAmt) = flexAlignRightCenter
        .ColWidth(Col_SatAmt) = 840

        .TextMatrix(0, Col_ItemVal) = "Item Value"
        .ColAlignmentFixed(Col_ItemVal) = flexAlignRightCenter
        .ColWidth(Col_ItemVal) = 1095

        .TextMatrix(0, Col_LName) = "Local Name"
        .ColAlignment(Col_LName) = flexAlignLeftCenter
        .ColWidth(Col_LName) = 2000

        .TextMatrix(0, Col_MRPStkTP) = "MRP Qty TP"
        .ColAlignmentFixed(Col_MRPStkTP) = flexAlignRightCenter
        .ColWidth(Col_MRPStkTP) = 0

        .TextMatrix(0, Col_MRPStkTB) = "MRP Qty TB"
        .ColAlignmentFixed(Col_MRPStkTB) = flexAlignRightCenter
        .ColWidth(Col_MRPStkTB) = 0

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
    DGPartyType.left = Me.width - (DGPartyType.width + mRtScale): DGPartyType.top = mTopScale
  
    FrmDetail.width = 6285: FrmDetail.left = 5595: FrmDetail.top = FGrid.top + FGrid.height: FrmDetail.height = 2130
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (DGPart.top + mBotScale)
    DGHist.width = FGrid.width: DGHist.left = FGrid.left: DGHist.top = FGrid.top + FGrid.height: DGHist.height = Me.height - (DGHist.top + mBotScale)
    DgCity.width = 4100: DgCity.left = Me.width - (DgCity.width + mRtScale): DgCity.top = mTopScale: DgCity.height = 2865
    With DGJCNo
        .left = 4300: DGJCNo.top = mTopScale
        .Columns(0).width = 1335.118
        .Columns(1).width = 1560.189
        .Columns(2).width = 1964.976
        .Columns(3).width = 2009.764
    End With
    
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

Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGJCNo.Visible = True Then DGJCNo.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
    If DGPartyType.Visible = True Then DGPartyType.Visible = False
    If TopCtrl1.TopText2 = "Browse" Then FrmDetail.Visible = False
End Sub

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
            If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Sub CountItem()
Dim I As Integer, TotItems As Integer, TotQty As Double
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            TotQty = TotQty + Val(FGrid.TextMatrix(I, Col_Qty))
            TotItems = TotItems + 1
        End If
    Next I
    LblIVal.CAPTION = Format(TotItems, "0")
    LblQty.CAPTION = Format(TotQty, "0.00")
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
Dim Rst As ADODB.Recordset, I As Integer
Dim mItemDiscTotTB As Double, mItemDiscTotTP As Double
On Error GoTo ELoop
    FrmDetail.Visible = False
    mMRevDisTBPer = 0
    mMRevDisTPPer = 0
    mTBDisAmtMRP = 0
    mTPDisAmtMRP = 0
    mMRPTax = 0
    mMRPTaxSur = 0
    mMRPTOT = 0
    mMRPReSales = 0
    mMRPLubeTB = 0
    mMRPLubeTP = 0

    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        With Master1
            .CursorLocation = adUseClient
            .Open "Select Estimate.*,SGType.Description As PartyTypeDesc,City.CityName " _
                & "From ((Estimate left Join SubGroupType SGType on Estimate.Party_Type=SGType.Party_Type) " _
                & "Left Join City on Estimate.CityCode=City.CityCode) " _
                & "where Estimate.DocID='" & Master!SearchCode & "' Order by V_Date,V_Type", GCn, adOpenStatic, adLockReadOnly
        End With
    
        txt(DocID).TEXT = Master1!DocID
        mSearchCode = txt(DocID)
        LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        mVType = Master1!V_Type
        If mVType = "W_EST" Then
            txt(OrgFrom).TEXT = "Workshop"
            Lbl(0).CAPTION = "Estimate Date"
            Lbl(2).CAPTION = "Estimate Sr.No"
        ElseIf mVType = "S_QU" Then
            txt(OrgFrom).TEXT = "Stores"
            Lbl(0).CAPTION = "Quotation Date"
            Lbl(2).CAPTION = "Quotation Sr.No"
        ElseIf mVType = "S_INV" Then
            txt(OrgFrom).TEXT = "Invoice"
            Lbl(0).CAPTION = "Invoice Date"
            Lbl(2).CAPTION = "Invoice No"
        End If
        txt(VDate).TEXT = Master1!V_DATE
        
        
        mVatYn = PubVATYN
        If CDate(Master1!V_DATE) < CDate("1/Jan/2008") And StrCmp(left(PubComp_Name, 3), "Jmk") Then
            mVatYn = 0
        End If
        
        
        
        LblVPrefix.CAPTION = mID(Master1!DocID, 9, 5)
        txt(SerialNo).TEXT = Master1!V_NO
        txt(PrintTitle) = IIf(IsNull(Master1!PrintTitle), "", Master1!PrintTitle)
        If IsNull(Master1!job_docid) Or Trim(Master1!job_docid) = "" Then
            txt(JCYN).TEXT = "No"
            txt(JCNo).TEXT = ""
            txt(RegNo).TEXT = XNull(Master1!RegNo)
            txt(Model) = XNull(Master1!Model)
            txt(ChassisNo).TEXT = XNull(Master1!Chassis)
            txt(Engno).TEXT = XNull(Master1!Engine)
        Else
            txt(JCYN).TEXT = "Yes"
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select J.Job_No,J.Job_Date, " & xIsNull("H.RegNo", "") & " As Reg_No," & xIsNull("H.Chassis", "''") & " As ChassisNo," & xIsNull("H.Engine", "''") & " As EngineNo From Job_Card J Left Join HisCard H on J.CardNo=H.CardNo Where J.DocID='" & Master1!job_docid & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                txt(JCNo).TEXT = Rst!Job_No
                txt(JCNo).Tag = Master1!job_docid
                OldJCNo = Rst!Job_No
                txt(OpDate).TEXT = Rst!Job_Date
                txt(RegNo).TEXT = Rst!Reg_No
                txt(Model) = Master1!Model
                txt(ChassisNo).TEXT = Rst!ChassisNo
                txt(Engno).TEXT = Rst!EngineNo
            Else
                txt(JCNo).Tag = ""
                OldJCNo = ""
                txt(OpDate).TEXT = ""
                txt(RegNo).TEXT = ""
                txt(Model) = ""
                txt(ChassisNo).TEXT = ""
                txt(Engno).TEXT = ""
            End If
        End If
        txt(Party).TEXT = XNull(Master1!Party_Name)
        txt(Address1).TEXT = XNull(Master1!Address)
        txt(Address2).TEXT = XNull(Master1!Address2)
        txt(City).TEXT = XNull(Master1!CityName)
        txt(City).Tag = XNull(Master1!CityCode)
        txt(PartyType) = IIf(IsNull(Master1!PartyTypeDesc), "General", Master1!PartyTypeDesc)
        txt(PartyType).Tag = IIf(IsNull(Master1!Party_Type), 0, Master1!Party_Type)
        txt(PrintTitle) = IIf(IsNull(Master1!PrintTitle), "", Master1!PrintTitle)
        txt(Suppli).TEXT = IIf(Master1!Suppl_YN = 1, "Yes", "No")
        txt(Remarks).TEXT = XNull(Master1!Remarks)
        txt(MRPAmtTB).TEXT = Format(Master1!SprAmt_MRP_TB, "0.00")
        txt(MRPAmtTP).TEXT = Format(Master1!SprAmt_MRP_TP, "0.00")
        mMRPLubeTB = Master1!OilAmt_MRP_TB
        mMRPLubeTP = Master1!OilAmt_MRP_TP
        txt(SprAmtTB).TEXT = Format(Master1!SprAmt_TB, "0.00")
        txt(SprAmtTP).TEXT = Format(Master1!SprAmt_TP, "0.00")
        txt(OilAmtTB).TEXT = Format(Master1!OilAmt_TB, "0.00")
        txt(OilAmtTP).TEXT = Format(Master1!OilAmt_TP, "0.00")
        txt(DiscPerTB).TEXT = Format(Master1!D_Per_TB, "0.00")
        txt(DiscAmtTB).TEXT = Format(Master1!D_Amt_TB, "0.00")
        txt(DiscPerTP).TEXT = Format(Master1!D_Per_TP, "0.00")
        txt(DiscAmtTP).TEXT = Format(Master1!D_Amt_TP, "0.00")
        
        txt(STotATB).TEXT = Format((Master1!SprAmt_MRP_TB + Master1!SprAmt_TB + Master1!OilAmt_TB) - Master1!D_Amt_TB, "0.00")
        txt(STotATP).TEXT = Format((Master1!SprAmt_MRP_TP + Master1!SprAmt_TP + Master1!OilAmt_TP) - Master1!D_Amt_TP, "0.00")
'        Txt(Addition).Text = Format(Master1!Addition, "0.00")
        txt(GenSurPer).TEXT = Format(Master1!Gen_Sur_Per, "0.00")
        txt(GenSurAmt).TEXT = Format(Master1!Gen_Sur_Amt, "0.00")
        txt(TransAmt).TEXT = Format(Master1!Trans_Amt, "0.00")

'        Txt(TaxableTot) = Format(Val(Txt(STotATB)) + Val(Txt(Addition)) + Val(Txt(PackCrg)) + Val(Txt(GenSurAmt)) + Val(Txt(TransAmt)), "0.00")
        txt(TaxableTot) = Format(Val(txt(STotATB)) + Val(txt(GenSurAmt)) + Val(txt(TransAmt)), "0.00")
        txt(STaxPer).TEXT = Format(Master1!Tax_Per, "0.00")
        txt(STaxAmt).TEXT = Format(Master1!Tax_Amt, "0.00")
        txt(SatAmt).TEXT = Format(Master1!SatAmt, "0.00")
        txt(TaxSurPer).TEXT = Format(Master1!Tax_Sur_Per, "0.00")
        txt(TaxSurAmt).TEXT = Format(Master1!Tax_Sur_Amt, "0.00")
        
        
        txt(PackCrg).TEXT = Format(Master1!Packing, "0.00")
'       Txt(STotB) = Format(Val(Txt(TaxableTot)) + Val(Txt(STaxAmt)) + Val(Txt(TaxSurAmt)), "0.00")
        txt(STotB) = Format(Val(txt(STotATP)) + Val(txt(TaxableTot)) + Val(txt(PackCrg)) + Val(txt(STaxAmt)) + Val(txt(SatAmt)) + Val(txt(TaxSurAmt)), "0.00")
        txt(TurnOverPer).TEXT = Format(Master1!TOT_Per, "0.00")
        txt(TurnOverAmt).TEXT = Format(Master1!Tot_Amt, "0.00")
        txt(ReSalTaxPer).TEXT = Format(Master1!ReSalTax_Per, "0.00")
        txt(ReSalTaxAmt).TEXT = Format(Master1!ReSalTax_Amt, "0.00")
        txt(SROff).TEXT = Format(Master1!Rounded, "0.00")
'        Txt(NetSprAmt) = Format(Val(Txt(STotB)) + Val(Txt(STotATP)) + Val(Txt(TurnOverAmt)) + Val(Txt(SROff)), "0.00")
        txt(NetSprAmt) = Format(Val(txt(STotB)) + Val(txt(TurnOverAmt)) + Val(txt(ReSalTaxAmt)) + Val(txt(SROff)), "0.00")
        
        mTBDisAmtMRP = Master1!D_Amt_MRP_TB
        mTPDisAmtMRP = Master1!D_Amt_MRP_TP
        mMRPTax = IIf(IsNull(Master1!Tax_AmtMRP), 0, Master1!Tax_AmtMRP)
        mMRPTaxSur = IIf(IsNull(Master1!TaxSur_AmtMRP), 0, Master1!TaxSur_AmtMRP)
        mMRPTOT = IIf(IsNull(Master1!Tot_AmtMrp), 0, Master1!Tot_AmtMrp)
        
        txt(LabAmt).TEXT = Format(Master1!Lab_Amt, "0.00")
        txt(LabDisc).TEXT = Format(Master1!Lab_D_Amt, "0.00")
        txt(ServTaxPer).TEXT = Format(Master1!Lab_TaxPer, "0.00")
        txt(ServTaxAmt).TEXT = Format(Master1!Lab_TaxAmt, "0.00")
        txt(NetLabAmt).TEXT = Format(Master1!Lab_Total_Amt, "0.00")
        txt(LabROff) = Format(Master1!Lab_Rounded, "0.00")
        
        txt(NetAmt).TEXT = Format(Master1!Total_Amt + Master1!Lab_Total_Amt, "0.00")
        
        FGrid.Rows = 1
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select P.Part_Name ,P.Local_Name ,P.Unit ,P.MRP ,P.MRP_Effect_Dt ,P.TB_SRate ,P.TP_SRate ,P.TB_Effect_Dt ,P.Part_Grade ,P.Cur_MRP_TBStk, P.Cur_MRP_TPStk,P.Cur_TB_Stk ,P.Cur_TP_Stk ,P.Bin_Loca ,P.High_Pur_Rate ,P.Low_Pur_Rate,Estimate1.* From Estimate1 Left Join Part P On Estimate1.Part_No=P.Part_No and P.Div_Code = left(Estimate1.Docid,1) Where Estimate1.DocID='" & Master1!DocID & "'", GCn, adOpenStatic, adLockReadOnly
        
        If Rst.RecordCount > 0 Then
            I = 1
            Do Until Rst.EOF
                            '|0 Col_SrNo |1 Col_PNo             |2 Col_Unit          |3 Col_MRP         |4 Col_Taxable                |5 Col_Qty                          |6 Col_Rate                       |7 Col_MRPRate                              |8 Col_Amt                                      |9 Col_DiscPer                              |10 Col_DiscAmt                   |11 Col_ItemVal                      |12 Col_PName            |13 Col_LName            |14 Col_MRPStkTP      |15 Col_MRPStkTB          |16 Col_TBStk            |17 Col_TPStk          |18 Col_TBRate          |19 Col_TPRate            |20 Col_Bin    |21 Col_LastRate          |22 Col_HPRate              |23 Col_LPRate            |24 Col_PartGrade         |25 Col_EffectDate
'                FGrid.AddItem i & Chr(9) & Rst!Part_No & Chr(9) & Rst!Unit & Chr(9) & Rst!MRPYN & Chr(9) & Rst!TaxYN & Chr(9) & Format(Rst!Qty, "0.000") & Chr(9) & Format(Rst!Rate, "0.00") & Chr(9) & Format(Rst!MRP_Rate, "0.00") & Chr(9) & Format((Rst!Qty * Rst!Rate), "0.00") & Chr(9) & Format(Rst!Disc_Per, "0.00") & Chr(9) & Format(Rst!Disc_Amt, "0.00") & Chr(9) & Format(Rst!AMOUNT, "0.00") & Chr(9) & Rst!Part_Name & Chr(9) & Rst!Local_Name & Chr(9) & Rst!Curstk & Chr(9) & Rst!MRPQty & Chr(9) & Rst!Cur_TB_Stk & Chr(9) & Rst!Cur_TP_Stk & Chr(9) & Rst!TB_SRate & Chr(9) & Rst!TP_SRate & Chr(9) & Rst!Bin_Loca & Chr(9) & " " & Chr(9) & Rst!high_pur_rate & Chr(9) & Rst!low_pur_rate & Chr(9) & Rst!Part_Grade & Chr(9) & Format(IIf(Rst!MRPYN = "Yes", Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
                             '0                 1                     2                     3                   4                            5                                   6                                  7                                            8                                             9                                         10                              11                                     12                     13                      14                        15                  16                       17                          18                   19                      20                         21                       22                    23                      24                    25
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Qty) = Format(Rst!Qty, "0.00")
                    .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                    .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")
                    If Rst!MRP_YN = 1 Then
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty * Rst!MRP_Rate), "0.00")
                    Else
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty * Rst!Rate), "0.00")
                    End If
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                    .TextMatrix(I, Col_TaxPer) = Format(Rst!TaxPer, "0.00")
                    .TextMatrix(I, Col_TaxAmt) = Format(Rst!TaxAmt, "0.00")
                    .TextMatrix(I, Col_SatPer) = Format(Rst!SatPer, "0.00")
                    .TextMatrix(I, Col_SatAmt) = Format(Rst!SatAmt, "0.00")
                    
                    If mVatYn = 1 Then
                        .TextMatrix(I, Col_ItemVal) = Format(Rst!Amount - Rst!Disc_Amt - Rst!TaxAmt, "0.00")
                    Else
                        .TextMatrix(I, Col_ItemVal) = Format(Rst!Amount - Rst!Disc_Amt, "0.00")
                    End If
                    
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
                    .TextMatrix(I, Col_Purpose) = "Charge"
                    .TextMatrix(I, Col_ReqNo) = IIf(VNull(Rst!ReqNo) = 0, "", Rst!ReqNo)
                End With

                If Rst!Tax_YN = 1 Then
                    mItemDiscTotTB = mItemDiscTotTB + Rst!Disc_Amt
                Else
                    mItemDiscTotTP = mItemDiscTotTP + Rst!Disc_Amt
                End If
                Rst.MoveNext
                I = I + 1
            Loop
'            Txt(IWDiscTotTB).Text = Format(mItemDiscTotTB, "0.00")
'            Txt(IWDiscTotTP).Text = Format(mItemDiscTotTP, "0.00")
            FGrid.FixedRows = 1
            CountItem
        Else
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
        End If
        txt(IWDiscTotTB).TEXT = Format(mItemDiscTotTB, "0.00")
        txt(IWDiscTotTP).TEXT = Format(mItemDiscTotTP, "0.00")
    Else
        BlankText
    End If
    Grid_Hide
    FGrid_GotFocus
Set Rst = Nothing
Set Master1 = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
    Select Case FGrid.Col
        Case Col_PNo, Col_PName, Col_LName
            If RsPart.RecordCount = 0 Then TxtGridLeave = False: Exit Function
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_PNo
            
        Case Col_Taxable, Col_MRP
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_TaxMRP
            
        Case Col_Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            FGrid.TextMatrix(FGrid.Row, Col_MRPRate) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_DiscPer, Col_TaxPer, Col_SatPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_MRPRate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
            
        Case Col_DiscAmt
            If Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) < Val(TxtGrid(0)) Then
                MsgBox "Item-wsie Disc. Amount is greater than Item Value", vbOKOnly, "Item-wise Disc. Checking"
                TxtGridLeave = False: Exit Function
            End If
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_ReqNo
            FGrid.TextMatrix(FGrid.Row, Col_ReqNo) = Val(TxtGrid(0).TEXT)
            
            
    End Select
    TxtGridLeave = True
    'Important at the time of validating  a control if you are making the visibility of
    'control false forcefully it will generate error
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
    'Eof Section
End Function

'* Used for Calculate the Amount
Private Sub Amt_Cal()
    Dim mAmount As Double
    Dim DisAmt As Double
    Dim I As Integer
    
    Dim mTaxableAmt As Double
    
    
    For I = 0 To FGrid.Rows - 1
        If UCase(FGrid.TextMatrix(I, Col_MRP)) = "YES" Then
            FGrid.TextMatrix(I, Col_Amt) = Format((Val(FGrid.TextMatrix(I, Col_MRPRate)) * Val(FGrid.TextMatrix(I, Col_Qty))), "0.00")
        Else
            FGrid.TextMatrix(I, Col_Amt) = Format((Val(FGrid.TextMatrix(I, Col_Rate)) * Val(FGrid.TextMatrix(I, Col_Qty))), "0.00")
        End If
        FGrid.TextMatrix(I, Col_DiscAmt) = Format(((Val(FGrid.TextMatrix(I, Col_Amt)) * Val(FGrid.TextMatrix(I, Col_DiscPer))) / 100), "0.00")
        FGrid.TextMatrix(I, Col_ItemVal) = Format((Val(FGrid.TextMatrix(I, Col_Amt)) - Val(FGrid.TextMatrix(I, Col_DiscAmt))), "0.00")
        
        
        If mVatYn = 1 Then
            If FGrid.TextMatrix(I, Col_TaxPer) <> "" Then
                mAmount = Val(FGrid.TextMatrix(I, Col_Amt))
                DisAmt = Val(FGrid.TextMatrix(I, Col_DiscAmt))
                If FGrid.TextMatrix(I, Col_MRP) = "Yes" And FGrid.TextMatrix(I, Col_Taxable) = "Yes" Then
                    mTaxableAmt = Format((mAmount - DisAmt) * 100 / (100 + Val(FGrid.TextMatrix(I, Col_TaxPer)) + Val(FGrid.TextMatrix(I, Col_SatPer))), "0.00")
                    FGrid.TextMatrix(I, Col_TaxAmt) = Format(mTaxableAmt * Val(FGrid.TextMatrix(I, Col_TaxPer)) / 100, "0.00")
                    FGrid.TextMatrix(I, Col_SatAmt) = Format(mTaxableAmt * Val(FGrid.TextMatrix(I, Col_SatPer)) / 100, "0.00")
                    FGrid.TextMatrix(I, Col_ItemVal) = Format(Val(FGrid.TextMatrix(I, Col_ItemVal)) - Val(FGrid.TextMatrix(I, Col_TaxAmt)) - Val(FGrid.TextMatrix(I, Col_SatAmt)), "0.00")
                ElseIf FGrid.TextMatrix(I, Col_MRP) = "No" And FGrid.TextMatrix(I, Col_Taxable) = "Yes" Then
                    FGrid.TextMatrix(I, Col_TaxAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(I, Col_TaxPer)) / 100, "0.00")
                    FGrid.TextMatrix(I, Col_SatAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(I, Col_SatPer)) / 100, "0.00")
                Else
                    FGrid.TextMatrix(I, Col_TaxAmt) = ""
                    FGrid.TextMatrix(I, Col_SatAmt) = ""
                End If
            End If
        End If
    Next


    txt(ServTaxAmt) = Format((Val(txt(LabAmt)) * Val(txt(ServTaxPer)) / 100), "0.00")
    txt(NetLabAmt) = Format(Val(txt(LabAmt)) + Val(txt(ServTaxAmt)), "0.00")
    txt(LabROff) = Format(txt(NetLabAmt) - Round(txt(NetLabAmt), 0), "0.00")
    txt(NetLabAmt) = Format(Val(txt(NetLabAmt)) - Val(txt(LabROff)), "0.00")

    MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
            Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
            
    If mVatYn = 1 Then
       MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt)
    Else
        MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, 0, False ', _
            'Txt(LabAmt), Txt(LabDisc), Txt(ServTaxPer), Txt(ServTaxAmt), Txt(LabROff), Txt(NetLabAmt), Txt(OutSideLabAmt)
    End If
        
    If mVatYn = 0 Then
        Set rstForm = GCn.Execute("Select * from taxforms")
        Set Syctrl = GCn.Execute("Select * from syctrl")
        GSQL = "Select Form_Desc from TaxForms where Form_Code='" & Syctrl!LocalTaxFormSpr & "'"
        rstForm.MoveFirst
        rstForm.FIND ("Form_Code ='" & Syctrl!LocalTaxFormSpr & "'")
        If rstForm.EOF = False Then
            txt(STaxPer) = rstForm!Tax_Per
            txt(TaxSurPer) = rstForm!Tax_Sur_Per
        Else
            txt(STaxPer) = ""
            txt(TaxSurPer) = ""
        End If
    End If
    txt(NetAmt) = Format(Round(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), 2), "0.00")
End Sub

Private Sub Check1_Click()

End Sub

Private Sub ChkMerg_Click()
Dim TempRs As ADODB.Recordset
' If Txt(JCNo) = "" Then
    If ChkMerg.Value = 1 Then
       ' Lbl(47).Visible = True: TxtPerformaNo.Visible = True: TxtPerformaNo.TEXT = "": TxtPerformaNo.SetFocus
        Lbl(49).Visible = True: CmboPLNo.Visible = True
        Set TempRs = GCn.Execute("Select V_No from Estimate where RegNo='" & txt(RegNo) & "'  and V_Type='W_PL'")
        CmboPLNo.Clear
        If TempRs.RecordCount > 0 Then
            Do Until TempRs.EOF
                If IsNull(TempRs!V_NO) = False Then CmboPLNo.AddItem TempRs!V_NO
                TempRs.MoveNext
            Loop
            CmboPLNo = CmboPLNo.List(1)
        Else
            CmboPLNo = ""
        End If
    Else
        'Lbl(47).Visible = False: TxtPerformaNo.Visible = False
        Lbl(49).Visible = False: CmboPLNo.Visible = False
    End If
' End If
Set TempRs = Nothing
End Sub

Private Sub DGCity_Click()
If RsCity.RecordCount > 0 Then
    txt(City).Tag = RsCity!Code
    txt(City).TEXT = RsCity!Name
End If
DgCity.Visible = False
txt(City).SetFocus
End Sub

Private Sub DGHist_Click()
If RsHist.RecordCount > 0 Then
    txt(RegNo).TEXT = XNull(RsHist!RegNo)
    txt(Model).TEXT = XNull(RsHist!Model)
    txt(ChassisNo).TEXT = XNull(RsHist!Chassis)
    txt(Engno).TEXT = XNull(RsHist!Engine)
    txt(Party).TEXT = XNull(RsHist!Name)
    txt(Address1).TEXT = XNull(RsHist!Add1)
    txt(Address2).TEXT = XNull(RsHist!Add2)
    txt(City).Tag = XNull(RsHist!CityCode)
    txt(City).TEXT = XNull(RsHist!CityName)
End If
DGHist.Visible = False
txt(RegNo).SetFocus
End Sub

Private Sub DGJCNo_Click()
    If RsJob.RecordCount > 0 Then
        txt(JCNo).TEXT = RsJob!Name
        txt(JCNo).Tag = RsJob!Code
        txt(RegNo).TEXT = RsJob!Reg_No
        txt(OpDate).TEXT = RsJob!Job_Date
        txt(ChassisNo).TEXT = RsJob!ChassisNo
        txt(Engno).TEXT = RsJob!EngineNo
        txt(Party).TEXT = RsJob!Party
        txt(Address1).TEXT = RsJob!Address1
        txt(Address2).TEXT = RsJob!Address2
    End If
    DGJCNo.Visible = False
    txt(JCNo).SetFocus
End Sub

Private Sub DGPartyType_Click()
On Error GoTo ELoop
    If rsPartyType.RecordCount > 0 Then
        txt(PartyType).TEXT = rsPartyType!Name
        txt(PartyType).Tag = rsPartyType!Code
    End If
    txt(PartyType).SetFocus
    DGPartyType.Visible = False
Exit Sub
ELoop:
    CheckError
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
    End If
    TxtGridValid_PNo
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGPart.Visible = False
Exit Sub
ELoop:
    CheckError
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
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
Call Ini_Pub
    If mVatYn = 1 Then
        Lbl(22) = "V A T"
    End If
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next
    txt(VDate).Tag = PubLoginDate
    Lbl(25) = pubTOTCaption
    
    '***********Nra Modi for SDT ********
    If PubSDTYN = 1 Then
        Lbl(25).CAPTION = XNull(GCn.Execute("Select TOTCaption from Syctrl").Fields(0).Value)
    End If
    '************************************
    
    
    
    If PubReSaleTaxPer = 0 Then
        Lbl(35).Visible = False
        txt(ReSalTaxPer).Visible = False
        txt(ReSalTaxAmt).Visible = False
    End If
    If PubWCompCode = "" Then
        For I = 3 To 8
            Lbl(I).Visible = False
            LblColon(I - 3).Visible = False
            txt(I + 1).Visible = False
            Lbl(I + 23).Visible = False
            txt(I + 40).Visible = False
        Next
        Line9.Visible = False
        For I = 19 To 22
            LblColon(I).Visible = False
        Next
        mVType = "S_QU"
    Else
        mVType = "W_EST"
    End If
    PubOutSideLabDisc = GCn.Execute("select " & vIsNull("OutSideLabDisc", "0") & " as OutSideLabDisc from Syctrl").Fields(0).Value
    PubSrvTaxOnOutSideLab = GCn.Execute("select " & vIsNull("SrvTaxOnOutSideLab", "0") & " as SrvTaxOnOutSideLab from Syctrl").Fields(0).Value

    Set DGPart.DataSource = RsPart
    Set rsPartyType = New ADODB.Recordset
    rsPartyType.CursorLocation = adUseClient
    rsPartyType.Open "Select Party_Type As Code,Description As Name From SubGroupType Order by Description", GCn, adOpenDynamic, adLockOptimistic
    Set DGPartyType.DataSource = rsPartyType

    Dim sitecond As String
    
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("j.DocId", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    

    Set RsJob = New ADODB.Recordset
    RsJob.CursorLocation = adUseClient
    RsJob.Open "Select J.DocID as Code, " & cCStr("J.Job_No") & " as Name,J.Job_No,J.Job_Date,J.CardNo, " & xIsNull("H.RegNo", "") & " As Reg_No, " & xIsNull("H.Chassis", "") & " As ChassisNo," & xIsNull("H.Engine", "") & " As EngineNo, " & xIsNull("H.Name", "") & " As Party, " & xIsNull("H.Add1", "") & " As Address1, " & xIsNull("H.Add2", "") & " As Address2 From Job_Card J Left Join HisCard H on J.CardNo=H.CardNo Where Left(J.DocID,1)='" & PubDivCode & "' " & sitecond & " and (Right(j.DocId_InvSpr,8) <> 'Cancelld' OR J.DocId_InvSpr Is Null) Order By J.Job_No", GCn, adOpenDynamic, adLockOptimistic
    Set DGJCNo.DataSource = RsJob
    

    sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    

    Set Master = New ADODB.Recordset
    With Master
        .CursorLocation = adUseClient
        If PubMoveRecYn Then
            .Open "Select DocId As SearchCode From Estimate where left(DocId,1)='" & PubDivCode & "' and V_Type in ('S_QU','W_EST','S_INV') " & sitecond & " Order by V_Date desc,V_Type,DocID desc", GCn, adOpenDynamic, adLockOptimistic
        Else
            .Open "Select Top 1 DocId As SearchCode From Estimate where left(DocId,1)='" & PubDivCode & "' and V_Type in ('S_QU','W_EST','S_INV') " & sitecond & " Order by V_Date desc,V_Type,DocID desc", GCn, adOpenDynamic, adLockOptimistic
        End If
    End With
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct V_No as code from Estimate where V_Type in ('S_QU','W_EST') order by V_No", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno
    
    'modi lps 02.09.03
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "Select CityCode as code,CityName as name FROM City Order by CityName", GCn, adOpenDynamic, adLockOptimistic
    Set DgCity.DataSource = RsCity
    RsCity.Sort = "Name"
    
    Set RsHist = New ADODB.Recordset
    RsHist.CursorLocation = adUseClient
    'Modify SQL for speed
    RsHist.Open "Select " & xIsNull("RegNo", "") & " as Code,Chassis,RegNo,Model,Name,Engine,GOVT_YN," & _
            " " & cIIF("GOVT_YN=0", "'No'", "'Yes'") & " as Govt, Add1,Add2,PhoneOff,PhoneResi,Mobile," & _
            " VehSerialNo,HISCARD.CityCode,City.CityName " & _
            " FROM (Hiscard " & _
            " left join city on Hiscard.CityCode=City.CityCode) " & _
            " Where HISCARD.Div_Code='" & PubDivCode & "' Order by Regno", GCn, adOpenDynamic, adLockOptimistic
    Set DGHist.DataSource = RsHist
    RsHist.Sort = "Code"
    'eof modi
    
    FrmPrn.left = (Me.width - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
    DGVno.left = 5145: DGVno.top = mTopScale
    
    MoveRec
    Disp_Text SETS("INI", Me, Master)
Exit Sub
ELoop:
    CheckError
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
    Set RsJob = Nothing
'    Set RsPart = Nothing
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
    txt(VDate).TEXT = txt(VDate).Tag
    If PubEstimateType = "Spare" Then
        txt(OrgFrom).TEXT = "Stores"
        Lbl(0).CAPTION = "Quotation Date"
        Lbl(2).CAPTION = "Quotation Sr.No"
        txt(JCYN).TEXT = "No"
        txt(Suppli).TEXT = "No"
    Else
        txt(OrgFrom).TEXT = "Workshop"
        Lbl(0).CAPTION = "Estimate Date"
        Lbl(2).CAPTION = "Estimate Sr.No"
        txt(JCYN).TEXT = "Yes"
        txt(Suppli).TEXT = "No"
    End If
    txt(ReSalTaxPer) = IIf(PubReSaleTaxPer = 0, "", Format(PubReSaleTaxPer, "0.00"))
    'txt(DocId) = GetDocID(GCnFaS, mVType, txt(Vdate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
    'MDOCID = txt(DocId)
    txt(PartyType).Tag = 0
    txt(PartyType) = "General"
    If PubWCompCode = "" Then
        txt(OrgFrom).Enabled = False
        txt(VDate).SetFocus
    Else
        txt(OrgFrom).Enabled = True
        txt(OrgFrom).SetFocus
    End If
    txt(TurnOverPer) = MainLib.TOTCal()
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    txt(OrgFrom).Enabled = False
    txt(SerialNo).Enabled = False
    FGrid.AddItem FGrid.Rows
    'Enable / Disable Text Box if values zero
    DisableEnableFooter txt(MRPAmtTB), txt(MRPAmtTP), txt(SprAmtTB), txt(SprAmtTP), _
            txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), txt(DiscPerTP), _
            txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), _
            txt(GenSurPer), txt(GenSurAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt)
    'EOF enable / disable section
    
    If PubWCompCode <> "" Then
        'Txt(JCYN).SetFocus
        txt(JCNo).SetFocus
    Else
        txt(Party).SetFocus
    End If
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        If pubUName = "SA" Then
            txt(VDate).Enabled = True
            txt(VDate).SetFocus
        Else
            txt(VDate).Enabled = False
        End If
    Else
        txt(VDate).Enabled = False
        
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
                GCn.Execute ("Delete From Estimate1 Where DocID='" & txt(DocID) & "'")
                GCn.Execute ("Delete From Estimate Where DocID='" & txt(DocID) & "'")
            GCn.CommitTrans
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
Exit Sub
ELoop:
    GCn.RollbackTrans
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
    'GSQL = "Select Estimate.DocId As SearchCode,Estimate.Site_Code,Switch(Estimate.V_Type='S_QU','Stores',Estimate.V_Type='W_EST','Workshop') As VType, Estimate.V_No, Estimate.V_Date AS VDate, Estimate.Party_Name,Estimate.RegNo FROM Estimate Order by Estimate.V_Date,Estimate.V_Type"
    
     Dim sitecond As String
     sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("Estimate.DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    
    If PubBackEnd = "A" Then
        GSQL = "Select Estimate.DocId As SearchCode,Estimate.Site_Code,Switch(Estimate.V_Type='S_QU','Stores',Estimate.V_Type='W_EST','Workshop') As VType, " & cCStr("Estimate.V_No") & " As V_No, " & cDt("Estimate.V_Date") & " AS VDate, Estimate.Party_Name,Estimate.RegNo FROM Estimate where left(DocId,1)='" & PubDivCode & "' and V_Type in ('S_QU','W_EST','S_INV') " & sitecond & " Order by Estimate.V_Date,Estimate.V_Type"
    ElseIf PubBackEnd = "S" Then
        GSQL = "Select Estimate.DocId As SearchCode,Estimate.Site_Code,Case Estimate.V_Type When 'S_QU' Then 'Stores' When 'W_EST' Then 'Workshop' End As VType, " & cCStr("Estimate.V_No") & " As V_No, " & cDt("Estimate.V_Date") & " AS VDate, Estimate.Party_Name,Estimate.RegNo FROM Estimate where left(DocId,1)='" & PubDivCode & "' and V_Type in ('S_QU','W_EST','S_INV') " & sitecond & "  Order by Estimate.V_Date,Estimate.V_Type"
    End If
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
ChkMerg.Value = False
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
Lbl(47).Visible = False: TxtPerformaNo.Visible = False
End Sub

Private Sub TopCtrl1_eRef()
    RsJob.Requery
    RsPart.Requery
    Master.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean
Dim DocIdHlp As String, mGridFilled As Boolean
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    If IsValid(txt(OrgFrom), "Originated From") = False Then Exit Sub
    If IsValid(txt(VDate), "Date") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Serial Number") = False Then Exit Sub
    If IsValid(txt(PrintTitle), "Print Title") = False Then Exit Sub
    If txt(OrgFrom) = "Workshop" Then
'        If IsValid(txt(JCYN), "Job Card (Yes/No)") = False Then Exit Sub
'        If txt(JCYN) = "Yes" Then
'            If IsValid(txt(JCNo), "Job Card No") = False Then Exit Sub
'        End If
        If IsValid(txt(RegNo), "Reg. No.") = False Then Exit Sub
    Else
        If IsValid(txt(Party), "Party Name") = False Then Exit Sub
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Please Specify MRP Yes/No in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Please Specify Taxable Yes/No in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Please Specify Quantity in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Rate)) = 0 Then
'                If PubULabel <> "Y" Then
                    MsgBox "Please Specify Rate in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Rate: FGrid.SetFocus: Exit Sub
'                End If
            End If
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_PNo: FGrid.SetFocus: Exit Sub
    'Amount Calculation
    If UCase(left(PubComp_Name, 3)) <> "JMK" Then
        MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
            Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
    End If
    
    If mVatYn = 1 Then
       MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt)
    Else
        MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, 0, False
            ', _
            Txt (LabAmt), Txt(LabDisc), Txt(ServTaxPer), Txt(ServTaxAmt), Txt(LabROff), Txt(NetLabAmt), Txt(OutSideLabAmt)
    End If
    
'    MainLib.LabCalc txt(LabAmtTB), txt(LabAmtTP), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt), mLabDiscAmtTB
'    If UCase(left(PubComp_Name, 3)) = "JMK" Then
'        Txt(TurnOverAmt) = Format((Val(Txt(TurnOverPer)) * Val(Txt(TaxableTot)) / 100), "0.00")
'        Txt(NetSprAmt) = Format(Txt(NetSprAmt) + Round(Txt(TurnOverAmt), 2), "0.00")
'    End If
    'Txt(NetAmt) = Format(Val(Txt(NetSprAmt)) + Val(Txt(NetLabAmt)), "0.00")
    'EOF Amount Calculation
    
    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Add" Then
            'lp 11-03-03
            mDocId = txt(DocID)
            If GCn.Execute("Select Count(*) From Estimate Where DocID='" & txt(DocID) & "'").Fields(0) > 0 Then
                If VoucherEditFlag Then
                    MsgBox "Serial No. " & txt(SerialNo) & " already exists, Retry", vbCritical, "Validation Error"
                    txt(SerialNo).SetFocus
                    GoTo ELoop
                Else
                    txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                    If Val(txt(SerialNo)) <= Val(DeCodeDocID(mDocId, Document_No)) Then
                        MsgBox "Serial No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                        GoTo ELoop
                    End If
                End If
            End If
            DocIdHlp = UCase(Replace(txt(DocID), " ", ""))
            '**********
            GCn.Execute "Insert Into Estimate(" _
                & "DocID,DocIDHelp,V_Type,V_No,Site_Code," _
                & "V_Date,Stores_Works,Party_Name,Address,Address2,Party_Type," _
                & "Job_DocID,SprAmt_MRP_TB,SprAmt_MRP_TP,OilAmt_MRP_TB,OilAmt_MRP_TP," _
                & "SprAmt_TB,SprAmt_TP,OilAmt_TB,OilAmt_TP," _
                & "D_Per_TB,D_Amt_TB,D_Per_TP,D_Amt_TP,Addition," _
                & "Packing,Gen_Sur_Per,Gen_Sur_Amt,Trans_Amt,Tax_Per," _
                & "Tax_Amt,Tax_Sur_Per,Tax_Sur_Amt,TOT_Per,TOT_Amt," _
                & "ReSalTax_Per, ReSalTax_Amt,Rounded,Total_Amt," _
                & "Lab_Amt, Lab_D_Amt, Lab_TaxPer, Lab_TaxAmt, " _
                & "Lab_Rounded,Lab_Total_Amt,Det_Tax,U_Name,U_EntDt,U_AE, " _
                & "D_Per_MRP_TB,D_Amt_MRP_TB,D_Per_MRP_TP,D_Amt_MRP_TP,Tax_AmtMRP,TaxSur_AmtMRP,TOT_AmtMRP," _
                & "CityCode,RegNo,Chassis,Engine,PrintTitle,Model,Suppl_YN,Remarks, SatAmt) Values(" _
                & "'" & txt(DocID) & "','" & DocIdHlp & "','" & mVType & "'," & txt(SerialNo) & ",'" & PubSiteCode & PubSiteCode & _
                "'," & ConvertDate(Format(txt(VDate), "dd/MMM/yyyy")) & ",'" & txt(OrgFrom) & "','" & txt(Party) & "','" & txt(Address1) & "','" & txt(Address2) & "'," & Val(txt(PartyType).Tag) & _
                ",'" & txt(JCNo).Tag & "'," & Val(txt(MRPAmtTB)) & "," & Val(txt(MRPAmtTP)) & "," & mMRPLubeTB & "," & mMRPLubeTP & _
                "," & Val(txt(SprAmtTB)) & "," & Val(txt(SprAmtTP)) & "," & Val(txt(OilAmtTB)) & "," & Val(txt(OilAmtTP)) & _
                "," & Val(txt(DiscPerTB)) & "," & Val(txt(DiscAmtTB)) & "," & Val(txt(DiscPerTP)) & "," & Val(txt(DiscAmtTP)) & "," & Val(txt(Addition)) & _
                "," & Val(txt(PackCrg)) & "," & Val(txt(GenSurPer)) & "," & Val(txt(GenSurAmt)) & "," & Val(txt(TransAmt)) & "," & Val(txt(STaxPer)) & _
                "," & Val(txt(STaxAmt)) & "," & Val(txt(TaxSurPer)) & "," & Val(txt(TaxSurAmt)) & "," & Val(txt(TurnOverPer)) & "," & Val(txt(TurnOverAmt)) & _
                "," & Val(txt(ReSalTaxPer)) & "," & Val(txt(ReSalTaxAmt)) & "," & Val(txt(SROff)) & "," & Val(txt(NetSprAmt)) & _
                "," & Val(txt(LabAmt)) & "," & Val(txt(LabDisc)) & "," & Val(txt(ServTaxPer)) & "," & Val(txt(ServTaxAmt)) & _
                "," & Val(txt(LabROff)) & "," & Val(txt(NetLabAmt)) & ",'" & PubTaxDetOnSprInv & "','" & pubUName & "'," & ConvertDate(PubServerDate) & _
                ",'A'," & mMRevDisTBPer & "," & mMRevDisTPPer & "," & mTBDisAmtMRP & "," & mTPDisAmtMRP & "," & mMRPTax & "," & mMRPTaxSur & ", " & mMRPTOT & _
                ",'" & txt(City).Tag & "','" & txt(RegNo) & "','" & txt(ChassisNo) & "','" & txt(Engno) & "','" & txt(PrintTitle) & "','" & txt(Model) & "'," & IIf(txt(Suppli) = "Yes", "1", "0") & ",'" & txt(Remarks) & "', " & Val(txt(SatAmt)) & ")"
            'update Table only when DocSrlNo >Table.SerialNo
            If txt(OrgFrom) <> "Invoice" Then
                UpdVouSrlNo GCnFaS, txt(DocID), txt(VDate)
            End If
            
            
        Else
            GCn.Execute ("Delete From Estimate1 Where DocID='" & txt(DocID) & "'")
            GCn.Execute "Update Estimate Set " _
                & "V_Date=" & ConvertDate(txt(VDate)) & ",V_No=" & txt(SerialNo) & ", Party_Name='" & txt(Party) & "',Address='" & txt(Address1) & "',Address2='" & txt(Address2) & _
                "',CityCode='" & txt(City).Tag & "',Party_Type=" & Val(txt(PartyType).Tag) & ",Job_DocID='" & txt(JCNo).Tag & _
                "',SprAmt_MRP_TB=" & Val(txt(MRPAmtTB)) & ",SprAmt_MRP_TP=" & Val(txt(MRPAmtTP)) & _
                ",OilAmt_MRP_TB=" & mMRPLubeTB & ",OilAmt_MRP_TP=" & mMRPLubeTP & ",SprAmt_TB=" & Val(txt(SprAmtTB)) & ",SprAmt_TP=" & Val(txt(SprAmtTP)) & _
                ",OilAmt_TB=" & Val(txt(OilAmtTB)) & ",OilAmt_TP=" & Val(txt(OilAmtTP)) & _
                ",D_Per_TB=" & Val(txt(DiscPerTB)) & ",D_Amt_TB=" & Val(txt(DiscAmtTB)) & _
                ",D_Per_TP=" & Val(txt(DiscPerTP)) & ",D_Amt_TP=" & Val(txt(DiscAmtTP)) & _
                ",Addition=" & Val(txt(Addition)) & ",Packing=" & Val(txt(PackCrg)) & _
                ",Gen_Sur_Per=" & Val(txt(GenSurPer)) & ",Gen_Sur_Amt=" & Val(txt(GenSurAmt)) & _
                ",Trans_Amt=" & Val(txt(TransAmt)) & ",Tax_Per=" & Val(txt(STaxPer)) & _
                ",Tax_Amt=" & Val(txt(STaxAmt)) & ",Tax_Sur_Per=" & Val(txt(TaxSurPer)) & _
                ",Tax_Sur_Amt=" & Val(txt(TaxSurAmt)) & ",TOT_Per=" & Val(txt(TurnOverPer)) & _
                ",TOT_Amt=" & Val(txt(TurnOverAmt)) & ",Rounded=" & Val(txt(SROff)) & _
                ",Total_Amt=" & Val(txt(NetSprAmt)) & ",Lab_Amt=" & Val(txt(LabAmt)) & ",Lab_D_Amt=" & Val(txt(LabDisc)) & _
                ",Lab_TaxPer=" & Val(txt(ServTaxPer)) & ",Lab_TaxAmt=" & Val(txt(ServTaxAmt)) & _
                ",Lab_Rounded=" & Val(txt(LabROff)) & ",Lab_Total_Amt=" & Val(txt(NetLabAmt)) & ",Det_Tax=" & PubTaxDetOnSprInv & _
                ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E',D_Per_MRP_TB=" & mMRevDisTBPer & _
                ",D_Amt_MRP_TB=" & mMRevDisTPPer & ", D_Per_MRP_TP =" & mTBDisAmtMRP & ", D_Amt_MRP_TP=" & mTPDisAmtMRP & " ,Tax_AmtMRP=" & mMRPTax & _
                ",TaxSur_AmtMRP= " & mMRPTaxSur & ", TOT_AmtMRP= " & mMRPTOT & ", ReSalTax_Per=" & Val(txt(ReSalTaxPer)) & ", ReSalTax_Amt=" & Val(txt(ReSalTaxAmt)) & _
                ",RegNo='" & txt(RegNo) & "',Chassis='" & txt(ChassisNo) & "',Engine='" & txt(Engno) & _
                "',Suppl_YN=" & IIf(txt(Suppli) = "Yes", 1, 0) & _
                ",Remarks='" & txt(Remarks) & "'" & _
                ",PrintTitle='" & txt(PrintTitle) & "',Model='" & txt(Model) & "', SatAmt = " & Val(txt(SatAmt)) & " Where DocID='" & txt(DocID) & "'"
                
        End If

        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" Then
                GCn.Execute "Insert Into Estimate1(" _
                    & "DocID,Sr_No,V_Type,Site_Code,Part_No," _
                    & "Qty,Tax_YN,MRP_YN,Rate,MRP_Rate," _
                    & "Disc_Per,Disc_Amt,Amount, TaxPer,  TaxAmt, SatPer, SatAmt," _
                    & "Lab_Code,Lab_Charges,Lab_Desc,ReqNo," _
                    & "U_Name,U_EntDt,U_AE,Item_Value) " _
                    & "Values(" _
                    & "'" & txt(DocID) & "'," & I & ",'" & mVType & "','" & PubSiteCode & PubSiteCode & "','" & FGrid.TextMatrix(I, Col_PNo) & "'," _
                    & "" & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & "," & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," _
                    & "" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & ", " & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ", " & Val(FGrid.TextMatrix(I, Col_TaxAmt)) & ", " & Val(FGrid.TextMatrix(I, Col_SatPer)) & ", " & Val(FGrid.TextMatrix(I, Col_SatAmt)) & " " & _
                    ",'',0,''," & Val(FGrid.TextMatrix(I, Col_ReqNo)) & "," _
                    & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "'," & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ")"
            End If
        Next
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(DocID)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select DocId As SearchCode From Estimate where left(DocId,1)='" & PubDivCode & "' and V_Type in ('S_QU','W_EST','S_INV') And DocId = '" & mSearchCode & "'  Order by V_Date desc,V_Type,DocID desc")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    'lp 11-03-03
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > DeCodeDocID(mDocId, Document_No) Then
            MsgBox "Serial No." & Trim(DeCodeDocID(mDocId, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    TopCtrl1_ePrn
Exit Sub
ELoop:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Master.FIND "SearchCode='" & mSearchCode & "'"
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To txt.Count - 1
            txt(I).BackColor = CtrlBColOrg
            txt(I).ForeColor = CtrlFColOrg
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
On Error GoTo ELoop
Ctrl_GetFocus txt(Index)
TxtGrid(0).Visible = False
Grid_Hide
Select Case Index
    Case JCNo
        OldJCNo = txt(Index).TEXT
        If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsJob!Name Then
            RsJob.MoveFirst
            RsJob.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case RegNo
        If RsHist.RecordCount = 0 Or txt(Index) = "" Then Exit Sub
        If UCase(txt(Index)) <> UCase(XNull(RsHist!RegNo)) Then
            RsHist.MoveFirst
            RsHist.FIND "RegNo ='" & txt(Index) & "'"
        End If
    Case City
        If RsCity.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsCity!Name Then
            RsCity.MoveFirst
            RsCity.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case PartyType
        If rsPartyType.RecordCount = 0 Or (rsPartyType.EOF = True Or rsPartyType.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsPartyType!Name Then
            rsPartyType.MoveFirst
            rsPartyType.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case OrgFrom
        txt(Index).ToolTipText = " (W) Workshop | (S) Store | (I) Invoice "
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case RegNo
            DGridTxtKeyDown DGHist, txt, Index, RsHist, KeyCode, False, 2
        Case City
            DGridTxtKeyDown DgCity, txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
        Case JCNo
            NumDown txt(Index), KeyCode, 8, 0
            If RsJob.RecordCount > 0 Then
                DGridTxtKeyDown DGJCNo, txt, JCNo, RsJob, KeyCode, False, 1
            Else
                Txt_Validate Index, True
            End If
        Case PartyType
            DGridTxtKeyDown DGPartyType, txt, Index, rsPartyType, KeyCode, False, 1
        Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt
            NumDown txt(Index), KeyCode, 8, 2
        Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ServTaxPer
            NumDown txt(Index), KeyCode, 2, 2
    End Select
    If DgCity.Visible = False And DGHist.Visible = False And DGPartyType.Visible = False And DGJCNo.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And _
            ((txt(OrgFrom) = "Workshop" And Index = ServTaxAmt) Or _
             (txt(OrgFrom) = "Stores" And ((PubReSaleTaxPer = 0 And Index = TurnOverAmt) Or (PubReSaleTaxPer <> 0 And Index = ReSalTaxAmt)))) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        Else
            If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If PubWCompCode = "" Then
                If Index <> VDate And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            Else
                If Index <> OrgFrom And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If PubWCompCode = "" Then
                If Index <> Party And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            Else
                If Index <> JCYN And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
Call CheckQuote(KeyAscii)
    Select Case Index
        Case RegNo, City, Party
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select

Select Case Index
    Case City
        DGridTxtKeyPress txt, Index, RsCity, KeyAscii, "name"
    Case RegNo
        If DGHist.Visible = True Then DGridTxtKeyPress txt, Index, RsHist, KeyAscii, "RegNo"
    Case OrgFrom
        
            If KeyAscii = 83 Or KeyAscii = 115 Then         ' S/s
                txt(Index).TEXT = "Stores"
                Lbl(0).CAPTION = "Quotation Date"
                Lbl(2).CAPTION = "Quotation Sr.No"
                txt(JCYN).TEXT = "No"
                KeyAscii = 0
                'modishekhar
                Call CtrlEnable(False)
            ElseIf KeyAscii = 87 Or KeyAscii = 119 Then   ' W/w
                txt(Index).TEXT = "Workshop"
                Lbl(0).CAPTION = "Estimate Date"
                Lbl(2).CAPTION = "Estimate Sr.No"
                txt(JCYN).TEXT = "Yes"
                KeyAscii = 0
                'modishekhar
                Call CtrlEnable(True)
            ElseIf Asc("I") = KeyAscii Or Asc("i") = KeyAscii Then    'I/i
                If pubUName <> "SA" Then
                    MsgBox "Permission Denied !"
                    Exit Sub
                End If
                txt(Index).TEXT = "Invoice"
                Lbl(0).CAPTION = "Invoice Date"
                Lbl(2).CAPTION = "Invoice No"
                KeyAscii = 0
            End If
        
    Case JCYN
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                txt(Index).TEXT = "Yes"
                txt(JCNo).Enabled = True
                txt(Party).Enabled = False
                txt(Address1).Enabled = False
                txt(Address2).Enabled = False
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                txt(Index).TEXT = "No"
                txt(JCNo).Enabled = False
                txt(Party).Enabled = True
                txt(Address1).Enabled = True
                txt(Address2).Enabled = True
                KeyAscii = 0
            End If
            txt(JCNo).TEXT = ""
            txt(OpDate).TEXT = ""
            txt(RegNo).TEXT = ""
            txt(Party).TEXT = ""
            txt(Address1).TEXT = ""
            txt(Address2).TEXT = ""
            txt(ChassisNo).TEXT = ""
            txt(Engno).TEXT = ""
        Else
            KeyAscii = 0
        End If
        Case Suppli
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                txt(Index).TEXT = "Yes"
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
       Else
            KeyAscii = 0
        End If
    Case JCNo
        If DGJCNo.Visible = True Then DGridTxtKeyPress txt, JCNo, RsJob, KeyAscii, "Name"
    Case PartyType
        If DGPartyType.Visible = True Then DGridTxtKeyPress txt, Index, rsPartyType, KeyAscii, "Name"
    Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt, ReSalTaxAmt
        NumPress txt(Index), KeyAscii, 8, 2
    Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer, ServTaxPer
        NumPress txt(Index), KeyAscii, 2, 2
    Case ReqNo, SerialNo
        NumPress txt(Index), KeyAscii, 8, 0
        
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo ELoop
    Select Case Index
        Case DiscPerTB
            If Val(txt(DiscPerTB)) = 0 Then txt(DiscAmtTB) = ""
        Case DiscPerTP
            If Val(txt(DiscPerTP)) = 0 Then txt(DiscAmtTP) = ""
        Case GenSurPer
            If Val(txt(GenSurPer)) = 0 Then txt(GenSurAmt) = ""
'        Case STaxPer
'            If Val(txt(STaxPer)) = 0 Then txt(STaxAmt) = ""
'        Case TaxSurPer
'            If Val(txt(TaxSurPer)) = 0 Then txt(TaxSurAmt) = ""
'        Case TurnOverPer
'            If Val(Txt(TurnOverPer)) = 0 Then Txt(TurnOverAmt) = ""
    End Select
    Select Case Index
        Case DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, PackCrg, TurnOverPer, TurnOverAmt, ReSalTaxAmt, ReSalTaxPer
            If Val(txt(MRPAmtTB)) + Val(txt(MRPAmtTP)) <> 0 Then
                MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                    Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                    Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
                    Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
            End If
            If mVatYn = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt)
            Else
                MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, 0, False
                    ', _
                    txt (LabAmt), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt)
            End If
        Case ServTaxPer
            txt(ServTaxAmt) = Format((Val(txt(LabAmt)) * Val(txt(ServTaxPer)) / 100), "0.00")
            txt(NetLabAmt) = Format(Val(txt(LabAmt)) + Val(txt(ServTaxAmt)), "0.00")
            txt(LabROff) = Format(txt(NetLabAmt) - Round(txt(NetLabAmt), 0), "0.00")
            txt(NetLabAmt) = Format(Val(txt(NetLabAmt)) - Val(txt(LabROff)), "0.00")
            Amt_Cal
        Case ServTaxAmt
            'Txt(NetLabAmt) = Format(Val(Txt(LabAmt)) + Val(Txt(ServTaxAmt)), "0.00")
            'Amt_Cal
    End Select
    If UCase(left(PubComp_Name, 3)) = "JMK" And mVatYn = 0 Then
        txt(TurnOverAmt) = Format((Val(txt(TurnOverPer)) * Val(txt(TaxableTot)) / 100), "0.00")
        txt(NetSprAmt) = Format(Val(txt(NetSprAmt)) + Round(Val(txt(TurnOverAmt)), 2), "0.00")
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
       
    Select Case Index
        Case OrgFrom
            If Not (Trim(txt(Index).TEXT) <> "Stores" Or Trim(txt(Index).TEXT) <> "Workshop" Or Trim(txt(Index).TEXT) <> "Invoice") Then
                txt(Index).TEXT = "Stores"
            End If
            If Trim(txt(Index).TEXT) = "Workshop" Then
                Lbl(0).CAPTION = "Estimate Date"
                Lbl(2).CAPTION = "Estimate Sr.No"
                txt(JCYN).TEXT = "Yes"
                mVType = "W_EST"
                txt(PrintTitle) = "WORKSHOP ESTIMATE"
                
            txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            mDocId = txt(DocID)
            ElseIf Trim(txt(Index).TEXT) = "Stores" Then
                Lbl(0).CAPTION = "Quotation Date"
                Lbl(2).CAPTION = "Quotation Sr.No"
                txt(JCYN).TEXT = "No"
                mVType = "S_QU"
                txt(PrintTitle) = "SPARE QUOTATION"
            txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            mDocId = txt(DocID)
            ElseIf Trim(txt(Index).TEXT) = "Invoice" Then
                Lbl(0).CAPTION = "Invoice Date"
                Lbl(2).CAPTION = "Invoice Sr.No"
                txt(JCYN).TEXT = "No"
                mVType = "S_INV"
                txt(PrintTitle) = "WORKSHOP SPARE INVOICE (CREDIT)"
                txt(PrintTitle).Locked = True
                txt(SerialNo).Enabled = True
            End If
            
        Case VDate
            txt(Index).TEXT = RetDate(txt(Index))
            Cancel = Not CheckFinYear(txt(Index))
            If Cancel = False Then
                If Trim(txt(OrgFrom).TEXT) <> "Invoice" And TopCtrl1.TopText2 = "Add" Then
                    txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                    mDocId = txt(DocID)
                End If
            End If
        Case SerialNo
        If Trim(txt(OrgFrom).TEXT) = "Invoice" Then
            txt(DocID) = PubDivCode + PubSiteCode + PubSiteCode + mVType + mVType + Space(8 - Len(txt(Index))) + txt(Index)
            mDocId = txt(DocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select V_No From Estimate Where DocID='" & txt(DocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(SerialNo).SetFocus
            End If
        Else
            If VoucherEditFlag Then      ' Manual
                txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                mDocId = txt(DocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select V_No From Estimate Where DocID='" & txt(DocID) & "'", GCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    txt(SerialNo).SetFocus
                End If
            End If
        End If
            
'        Case JCYN
'            If Not Trim(Txt(Index).Text) <> "Yes" Or Trim(Txt(Index).Text) <> "No" Then
'                Txt(Index).Text = "Yes"
'            End If
'            If Trim(Txt(Index).Text) = "Yes" Then
'                Txt(JCNo).Enabled = True
'                Txt(Party).Enabled = False
'                Txt(Address1).Enabled = False
'                Txt(Address2).Enabled = False
'            ElseIf Trim(Txt(Index).Text) = "No" Then
'                Txt(JCNo).Enabled = False
'                Txt(Party).Enabled = True
'                Txt(Address1).Enabled = True
'                Txt(Address2).Enabled = True
'            End If
        Case JCNo
                If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then
                    txt(Index).TEXT = ""
                    txt(Index).Tag = ""
                    If txt(OrgFrom).TEXT = "Workshop" Then
                        txt(RegNo).Enabled = True
                        txt(ChassisNo).Enabled = True
                        txt(Engno).Enabled = True
                    End If
                Else
                    If txt(Index).TEXT = OldJCNo Then Exit Sub
                    If DGJCNo.Visible = True Then txt(JCNo).TEXT = RsJob!Name
                    txt(JCNo).Tag = RsJob!Code
                    txt(RegNo).TEXT = RsJob!Reg_No
                    txt(OpDate).TEXT = RsJob!Job_Date
                    txt(ChassisNo).TEXT = RsJob!ChassisNo
                    txt(Engno).TEXT = RsJob!EngineNo
                    txt(Party).TEXT = RsJob!Party
                    txt(Address1).TEXT = RsJob!Address1
                    txt(Address2).TEXT = RsJob!Address2
                    
                    If UCase(left(PubComp_Name, 3)) <> "JMK" Then
                        If TopCtrl1.TopText2.CAPTION = "Add" Then
                            FGrid.Rows = 1
                            Set Rst = New ADODB.Recordset
                            Rst.CursorLocation = adUseClient
                            Rst.Open "Select S.Part_No,P.Part_Name,P.Local_Name ,P.Unit ,P.MRP ,P.MRP_Effect_Dt ,P.TB_SRate ,P.TP_SRate ,P.TB_Effect_Dt ,P.Part_Grade ,(P.Cur_MRP_TBStk+P.Cur_MRP_TPStk) as MRPQty ,(P.Cur_MRP_TBStk+P.Cur_MRP_TPStk+P.Cur_TB_Stk+P.Cur_TP_Stk) As CurStk, P.Cur_TB_Stk ,P.Cur_TP_Stk ,P.Bin_Loca ,P.High_Pur_Rate ,P.Low_Pur_Rate,S.TAX_YN,S.MRP_YN,S.Qty_Iss,S.Qty_Ret,S.Rate,S.MRP_Rate,S.Disc_Per,S.Disc_Amt, s.TaxPer,S.TaxAmt, S.SatPer, S.SatAmt, S.Net_Amt " & _
                                " From SP_Stock S Left Join Part P On S.Part_No=P.Part_No and P.Div_Code = left(S.Docid,1) " & _
                                " Where S.Job_DocID='" & txt(JCNo).Tag & "' and (S.Qty_Iss - S.Qty_Ret)  > 0 and S.Purpose='C'", GCn, adOpenStatic, adLockReadOnly
                            If Rst.RecordCount > 0 Then
                                I = 1
                                Do Until Rst.EOF
            '                            '|0 Col_SrNo |1 Col_PNo             |2 Col_Unit          |3 Col_MRP         |4 Col_Taxable                |5 Col_Qty                          |6 Col_Rate                       |7 Col_MRPRate                              |8 Col_Amt                                      |9 Col_DiscPer                      |10 Col_DiscAmt  |11 Col_FlatDiscPer |12 Col_FlatDiscAmt|13 Col_TaxPer|14 Col_TaxAmt|15 Col_ItemVal |16 Col_PName            |17 Col_LName       |18 Col_MRPStkTP      |19 Col_MRPStkTB             |20 Col_TBStk             |21 Col_TPStk          |22 Col_TBRate          |23 Col_TPRate             |24 Col_Bin    |25 Col_LastRate          |26 Col_HPRate              |27 Col_LPRate         |28 Col_PartGrade                                                     |29 Col_EffectDate
        '                            FGrid.AddItem i & Chr(9) & Rst!Part_No & Chr(9) & Rst!Unit & Chr(9) & Rst!MRPYN & Chr(9) & Rst!TaxYN & Chr(9) & Rst!Qty_iss & Chr(9) & Format(Rst!Rate, "0.00") & Chr(9) & Format(Rst!MRP_Rate, "0.00") & Chr(9) & Format((Rst!Qty_iss * Rst!Rate), "0.00") & Chr(9) & Format(Rst!Disc_Per, "0.00") & Chr(9) & Format(Rst!Disc_Amt, "0.00") & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & Rst!Part_Name & Chr(9) & Rst!Local_Name & Chr(9) & Rst!Curstk & Chr(9) & Rst!MRPQty & Chr(9) & Rst!Cur_TB_Stk & Chr(9) & Rst!Cur_TP_Stk & Chr(9) & Rst!TB_SRate & Chr(9) & Rst!TP_SRate & Chr(9) & Rst!Bin_Loca & Chr(9) & " " & Chr(9) & Rst!high_pur_rate & Chr(9) & Rst!low_pur_rate & Chr(9) & Rst!Part_Grade & Chr(9) & Format(IIf(Rst!MRPYN = "Yes", Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
        '                                           0                 1                     2                     3                   4                  5                                6                                  7                                            8                                               9                                         10                      11            12           13          14           15                   16                     17                         18                   19                      20                         21                       22                    23                      24                 25                     26                          27                        28                                                               29
                                    FGrid.AddItem ""
                                    With FGrid
                                        .TextMatrix(I, Col_SrNo) = I
                                        .TextMatrix(I, Col_PNo) = Rst!Part_No
                                        .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                                        .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                                        .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                                        .TextMatrix(I, Col_Qty) = Format(Rst!Qty_Iss - Rst!Qty_Ret, "0.00")
                                        .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                                        .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")
                                        If Rst!MRP_YN = 1 Then
                                            .TextMatrix(I, Col_Amt) = Format(((Rst!Qty_Iss - Rst!Qty_Ret) * Rst!MRP_Rate), "0.00")
                                        Else
                                            .TextMatrix(I, Col_Amt) = Format(((Rst!Qty_Iss - Rst!Qty_Ret) * Rst!Rate), "0.00")
                                        End If
                                        .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                                        .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                                        .TextMatrix(I, Col_TaxPer) = Format(VNull(Rst!TaxPer), "0.00")
                                        .TextMatrix(I, Col_TaxAmt) = Format(VNull(Rst!TaxAmt), "0.00")
                                        .TextMatrix(I, Col_SatPer) = Format(VNull(Rst!SatPer), "0.00")
                                        .TextMatrix(I, Col_SatAmt) = Format(VNull(Rst!SatAmt), "0.00")
                                        .TextMatrix(I, Col_ItemVal) = Format(VNull(Rst!Net_Amt), "0.00") 'Format(.TextMatrix(I, Col_Amt), "0.00")
                                        .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                                        .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                                        .TextMatrix(I, Col_MRPStkTP) = IIf(IsNull(Rst!Curstk), "", Rst!Curstk)
                                        .TextMatrix(I, Col_MRPStkTB) = IIf(IsNull(Rst!MRPQty), "", Rst!MRPQty)
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
                                        .TextMatrix(I, Col_Purpose) = "Charge"
                                    End With
                                    Rst.MoveNext
                                    I = I + 1
                                Loop
                                FooterValue
                                FGrid.FixedRows = 1
                            Else
                                FGrid.AddItem FGrid.Rows
                                FGrid.FixedRows = 1
                            End If
                            
                            
                            
                            Set Rst = GCn.Execute("Select Sum(LabourAmt) As LabAmt  " & _
                                                 "From Job_Lab Where Job_DocID='" & txt(JCNo).Tag & "'")
                            If Rst.RecordCount > 0 Then
                                txt(LabAmt) = Format(VNull(Rst!LabAmt), "0.00")
                                txt(ServTaxPer) = GCn.Execute("Select " & vIsNull("Service_Tax", "0") & " From Syctrl").Fields(0).Value
                                Amt_Cal
                            End If
                        End If
                    End If
                End If
        Case RegNo
            If RsHist.EOF = False Or RsHist.BOF = False Then
                If txt(Index).TEXT <> "" Then
                    txt(RegNo).TEXT = XNull(RsHist!RegNo)
                    txt(Model).TEXT = XNull(RsHist!Model)
                    txt(ChassisNo).TEXT = XNull(RsHist!Chassis)
                    txt(Engno).TEXT = XNull(RsHist!Engine)
                    txt(Party).TEXT = XNull(RsHist!Name)
                    txt(Address1).TEXT = XNull(RsHist!Add1)
                    txt(Address2).TEXT = XNull(RsHist!Add2)
                    txt(City).Tag = XNull(RsHist!CityCode)
                    txt(City).TEXT = XNull(RsHist!CityName)
                End If
            End If
        Case City
            If RsCity.EOF = False And RsCity.BOF = False Then
                If txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = RsCity!Name
                    txt(Index).Tag = RsCity!Code
                End If
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Case PartyType
            If rsPartyType.RecordCount > 0 Or (rsPartyType.EOF = False Or rsPartyType.BOF = False) Then
                If txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = rsPartyType!Name
                    txt(Index).Tag = rsPartyType!Code
                Else
                    txt(Index).TEXT = ""
                    txt(Index).Tag = ""
                End If
            End If
        Case DiscAmtTB
            If (Val(txt(MRPAmtTB)) + Val(txt(SprAmtTB)) + Val(txt(OilAmtTB))) < Val(txt(DiscAmtTB)) Then
                MsgBox "Discount Amount is greater than Goods Value!", vbOKOnly, "Discount Value Check"
                Cancel = False
            End If
            If Val(txt(Index).TEXT) = 0 Then
                txt(Index).TEXT = ""
            Else
                txt(Index).TEXT = Format(txt(Index), "0.00")
            End If
            
        Case DiscAmtTP
            If (Val(txt(MRPAmtTP)) + Val(txt(SprAmtTP)) + Val(txt(OilAmtTP))) < Val(txt(DiscAmtTP)) Then
                MsgBox "Discount Amount is greater than Goods Value!", vbOKOnly, "Discount Value Check"
                Cancel = True
            End If
            If Val(txt(Index).TEXT) = 0 Then
                txt(Index).TEXT = ""
            Else
                txt(Index).TEXT = Format(txt(Index), "0.00")
            End If
            
        Case DiscPerTB, DiscPerTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, SROff, LabAmt, LabDisc, ServTaxPer, ServTaxAmt
            If Val(txt(Index).TEXT) = 0 Then
                txt(Index).TEXT = ""
            Else
                txt(Index).TEXT = Format(txt(Index), "0.00")
            End If
        Case Remarks
            If FGrid.ColWidth(FGrid.Col) = 0 Then FGrid.Col = 2
    End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
    Ctrl_GetFocus TxtGrid(Index)
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
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                End If
            End If
        Case Col_Qty
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                End If
            End If
        Case Col_Rate, Col_MRPRate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, 2
                End If
            End If
        Case Col_DiscPer, Col_TaxPer, Col_SatPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     'GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PNo, 1
                End If
            End If
        Case Col_DiscAmt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, 5
                End If
            End If
        Case Col_ReqNo
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PNo
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
    Case Col_DiscPer, Col_TaxPer, Col_SatPer
        NumPress TxtGrid(Index), KeyAscii, 2, 2
    Case Col_Rate, Col_DiscAmt
        NumPress TxtGrid(Index), KeyAscii, 8, 2
    Case Col_ReqNo
        NumPress TxtGrid(Index), KeyAscii, 8, 0
        
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
                If KeyCode <> 13 And DGPart.Visible = False Then
                    TxtGrid_KeyDown Index, GridKey, 0
                    DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Code", True
                End If
            Case Col_PName
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Name", True
            Case Col_LName
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "LName", True
            Case Col_MRP
                If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                    TxtGrid(Index) = ""
                ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
                    TxtGrid(Index) = "Yes"
                Else
                    TxtGrid(Index) = "No"
                End If
            Case Col_Taxable
                If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                    TxtGrid(Index) = ""
                ElseIf UCase(left$(TxtGrid(Index), 1)) = "N" Then
                    TxtGrid(Index) = "No"
                Else
                    TxtGrid(Index) = "Yes"
                End If
                
            Case Col_DiscPer, Col_DiscAmt, Col_Rate, Col_TaxPer, Col_SatPer
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                Amt_Cal
            Case Col_Qty
                FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0.00")
            Case Col_ReqNo
                FGrid.TextMatrix(FGrid.Row, Col_ReqNo) = Val(TxtGrid(Index).TEXT)
            
                
'                CountItem
'                Amt_Cal
        End Select
'        Amt_Cal
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
    TxtGrid(0).Visible = False
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
            TAddMode = False
        Case Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_MRPRate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                GridDblClick Me, FGrid, TxtGrid, 0
                TAddMode = False
            End If
    End Select
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
'        FGrid.CellBackColor = CellBackColLeave
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
'        FGrid.CellBackColor = CellBackColLeave
        SendKeysA vbKeyTab, True
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid.Col
            Case Col_MRP, Col_Taxable
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            Case Col_Qty, Col_Rate, Col_Amt, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                Amt_Cal
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case FGrid.Col
            Case Col_PNo, Col_PName, Col_LName
                GridDblClick Me, FGrid, TxtGrid, 0
                TAddMode = False
            Case Col_Taxable, Col_MRP, Col_Qty, Col_Rate, Col_MRPRate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    GridDblClick Me, FGrid, TxtGrid, 0
                    TAddMode = False
                End If
            Case Col_Amt
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    FGrid_LeaveCell
                    GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_LName, , Col_DiscPer
                End If
            Case Col_ItemVal
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    FGrid_LeaveCell
                    GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_LName, , Col_PName
                End If
        End Select
        TAddMode = False
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
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            End If
        Case Col_Qty, Col_Rate, Col_MRPRate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            End If
        Case Col_ReqNo
            Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
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
                FGrid.TextMatrix(I, Col_SrNo) = I
            Next
            CountItem
            MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
                FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
                Col_MRPStkTB, Col_MRPStkTP, _
                Col_TBStk, Col_TPStk, _
                Col_MRPRate, Col_TBRate, _
                Col_TPRate, Col_Bin, _
                Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
            If mVatYn = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt)
            Else
                MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, 0, False
                    ', _
                    txt (LabAmt), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt)
            End If
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
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
        If TxtGrid(0).Visible = False Then
            MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                    Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                    Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
                    Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
            If mVatYn = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt)
            Else
                MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, 0, False
                    ', _
                    txt (LabAmt), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt)
            End If
        End If
    End If
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
Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case FromVno, ToVno
            If IsValid(txt(VType1), "Organised from") = False Then Exit Sub
            RsVno.Close
            RsVno.Open "Select V_no as code from Estimate where Estimate.V_Type ='" & txtPrint(VType1).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
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
    Case FromVno, ToVno
        DGridTxtKeyDown DGVno, txtPrint, Index, RsVno, KeyCode, False, 0
End Select
If DGVno.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> VType1 Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case VType1
        If KeyAscii = 83 Or KeyAscii = 115 Or KeyAscii = 87 Or KeyAscii = 119 Then
            If KeyAscii = 83 Or KeyAscii = 115 Then         ' S/s
                txtPrint(Index).TEXT = "Stores"
                KeyAscii = 0
            ElseIf KeyAscii = 87 Or KeyAscii = 119 Then     ' W/w
                txtPrint(Index).TEXT = "Workshop"
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
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
    Case VType1
        If IsValid(txt(VType1), "Voucher Type") = False Then Cancel = True:   Exit Sub
        If txtPrint(VType1).TEXT = "Stores" Then
            txtPrint(VType1).Tag = "S_QU"
        ElseIf txtPrint(VType1).TEXT = "Workshop" Then
            txtPrint(VType1).Tag = "W_EST"
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
Dim rsTaxPer As ADODB.Recordset

Dim OldPNo$, mRate As Double
    If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Or TxtGrid(0).TEXT = "" Then
        FGrid.TextMatrix(FGrid.Row, Col_PNo) = ""
        FGrid.TextMatrix(FGrid.Row, Col_PName) = ""
        FGrid.TextMatrix(FGrid.Row, Col_LName) = ""
        
            MainLib.Fill_Data Val(txt(PartyType).Tag), LblFrm, FGrid, _
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
        
        
            MainLib.Fill_Data Val(txt(PartyType).Tag), LblFrm, FGrid, _
                RsPart!Code, RsPart!Name, RsPart!LName, _
                Col_Unit, Col_MRP, Col_Taxable, Col_MRPStkTB, Col_MRPStkTP, _
                Col_TBStk, Col_TPStk, _
                Col_MRPRate, Col_TBRate, _
                Col_TPRate, Col_Bin, _
                Col_HPRate, Col_LPRate, _
                Col_LastRate, Col_PartGrade, _
                Col_EffectDate, Col_DiscPer, mCheckNegetiveStockSiteWise

        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> OldPNo Then
                mRate = GetRate(Val(txt(PartyType).Tag), FGrid, CDate(txt(VDate)), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
                FGrid.TextMatrix(FGrid.Row, Col_Rate) = Format(mRate, "0.00")
'                FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
            End If
        End If
    End If
    
    If mVatYn = 1 Then
        GSQL = "Select TAX_Per, AddTaxPer from TaxForms where Form_Code=(Select LocalTaxFormSpr from syctrl)"
        Set rsTaxPer = GCn.Execute(GSQL)
         If rsTaxPer.RecordCount > 0 Then
              FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = rsTaxPer!Tax_Per
              FGrid.TextMatrix(FGrid.Row, Col_SatPer) = rsTaxPer!AddTaxPer
                                            
                Set rsTaxPer = GCn.Execute("Select VatPer, AddTaxPer From Part_Grade Where PartGrade_Code='" & FGrid.TextMatrix(FGrid.Row, Col_PartGrade) & "'")
                If rsTaxPer.RecordCount > 0 Then
                    If VNull(rsTaxPer!VatPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = Format(rsTaxPer!VatPer, "0.00")
                    If VNull(rsTaxPer!AddTaxPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_SatPer) = Format(rsTaxPer!AddTaxPer, "0.00")
                End If
                
         End If
    End If
    
    If FGrid.TextMatrix(FGrid.Rows - 1, Col_PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

Private Sub TxtGridValid_TaxMRP()
Dim mRate As Double
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'        If TopCtrl1.TopText2 = "Add" Or _
            TopCtrl1.TopText2 = "Edit" And Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) = 0 Then
            mRate = GetRate(Val(txt(PartyType).Tag), FGrid, CDate(txt(VDate)), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
            FGrid.TextMatrix(FGrid.Row, Col_Rate) = Format(mRate, "0.00")
'        End If
    End If
    Amt_Cal
End Sub
Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmPrn.Visible = False
        If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
            If TopCtrl1.TopText2.CAPTION = "Add" Then
                txt(VDate).Tag = txt(VDate).TEXT
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
        mRepName = IIf(mVatYn = 1, "SpQuotVat", "SpQuot")
        If UCase(left(PubComp_Name, 6)) = "RASHMI" Then
            mRepName = "SpQuot_Rashmi"
        End If
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
    Case PDos
        If UCase(left(PubComp_Name, 3)) <> "JMK" Then
            Call SpeedPrint(Optpre.Value, ChkMerg.Value)
            FrmPrn.Visible = False
        Else
            Call SpeedPrintJMK(Optpre.Value, ChkMerg.Value)
            FrmPrn.Visible = False
        End If
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "SpQuot", "SpQuot")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        txt(VDate).Tag = txt(VDate).TEXT
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
'Previous writen code is commented at the last of form
Dim RST1 As ADODB.Recordset
Dim RstRep As ADODB.Recordset
Dim mQry$, mQryLab$, mLabDocId$
Dim I As Integer, j As Integer
'On Error GoTo ERRORHANDLER
   
'        mQRY = "SELECT Estimate.*, Estimate1.*, City.CityName, syctrl.EstiInvFooter, Part.Part_Name " & _
        "FROM (((Estimate LEFT JOIN Estimate1 ON Estimate.DocID = Estimate1.DocID) LEFT JOIN City ON Estimate.CityCode = City.CityCode) LEFT JOIN Part ON Estimate1.Part_No = Part.PART_NO and Part.Div_Code = left(Estimate1.Docid,1)) LEFT JOIN syctrl ON syctrl.LinkTable  >= Estimate.U_AE " & _
        "where Estimate.DocId='" & Master!searchcode & "'"
        
        
    If ChkMerg.Value = 1 Then
        If txt(JCNo).Tag & txt(RegNo) <> "" Then
            If txt(JCNo).Tag <> "" Then
                GSQL = "Select top 1 DocID from Estimate where Job_DocID='" & txt(JCNo).Tag & "' and V_Type='W_PL'"
            Else
                If CmboPLNo.TEXT = "" Then
                    MsgBox "Enter the Performa No to Merze.", vbOKOnly + vbInformation, "Validation Message"
                    CmboPLNo.SetFocus: Exit Sub
                    'TxtPerformaNo.SetFocus: Exit Sub
                Else
                    GSQL = "Select top 1 DocID from Estimate where RegNo='" & txt(RegNo) & "' and V_No=" & Val(CmboPLNo.TEXT) & " and V_Type='W_PL'"
                End If
            End If
            If GCn.Execute(GSQL).RecordCount > 0 Then
                mLabDocId = GCn.Execute(GSQL).Fields(0).Value
            Else
                MsgBox "Labour Performa Not Feeded." & vbCrLf & "Merge Printing not possible!", vbCritical, "Validation": Exit Sub
            End If
        Else
            MsgBox "Job No./Reg. No. Not Feeded." & vbCrLf & "Merge Printing not possible!", vbCritical, "Validation": Exit Sub
        End If
    End If
        
    If ChkMerg.Value = 1 Then
        GSQL = "SELECT '1' as Orig," & cMID("E1.DocID", "4", "5") & " as V_Type,E1.DocID," & vIsNull("E1.Sr_No", "0") & " as Sr_No,E.Job_DocID,E.v_Date,E.Model,E.RegNo,E.Chassis,E.Engine,E.Party_Code,E.NamePrefix,E.Party_Name,E.Address,E.Address2,E.Address3,City.CityName,E.PhoneNo," & _
            "E.SprAmt_MRP_TB,E.SprAmt_MRP_TP,E.OilAmt_MRP_TB,E.OilAmt_MRP_TP,E.SprAmt_TB,E.SprAmt_TP,E.OilAmt_TB, E.OilAmt_TP,E.D_Per_TB, E.D_Amt_TB, E.D_Per_TP,E.D_Amt_TP,E.D_Per_MRP_TB,E.D_Amt_MRP_TB," & _
            "E.D_Per_MRP_TP,E.D_Amt_MRP_TP,E.Addition,E.Gen_Sur_Per,E.Gen_Sur_Amt,E.Trans_Amt,E.Tax_Per, E.Tax_Amt, E.Tax_AmtMRP, E.Tax_Sur_Per,E.Tax_Sur_Amt,E.TaxSur_AmtMRP,E.Packing, E.TOT_Per, E.Tot_Amt, E.TOT_AmtMRP,E.ReSalTax_Per,E.ReSalTax_Amt,E.Total_Amt," & _
            "E.Rounded,E.Det_Tax,E.Printed_YN,E.U_Name,E.U_EntDt,0 as Lab_Amt, 0 as Lab_D_Amt, 0 as Lab_TaxPer, 0 as Lab_TaxAmt,0 as Lab_RoundOff,0 as Lab_Total_Amt," & _
            "E1.Part_No as Part_No,P.Part_Name," & vIsNull("E1.Qty", "0") & " as Qty," & _
            " " & vIsNull("E1.Tax_YN", "0") & " as Tax_YN," & vIsNull("E1.MRP_YN", "0") & " as MRP_YN," & vIsNull("E1.Rate", "0") & " as Rate," & _
            " " & vIsNull("E1.MRP_Rate", "0") & " as MRP_Rate," & _
            " " & vIsNull("E1.Disc_Per", "0") & " as Disc_Per," & vIsNull("E1.Disc_Amt", "0") & " as Disc_Amt," & vIsNull("E1.Amount", "0") & " as Amount," & _
            "Syctrl.EstiInvFooter, E.Remarks, E1.TaxPer, E1.TaxAmt, convert(VARCHAR,E1.ReqNo) As RegNo, E.PrintTitle " & _
        "FROM (((((Estimate as E left JOIN Estimate1 as E1 ON E.DocID = E1.DocId) " & _
            "left JOIN Part as P ON E1.Part_No = P.Part_No and P.Div_Code = left(E1.Docid,1)) " & _
            "LEFT JOIN SubGroup as SG ON E.Party_Code = SG.SubCode) " & _
            "left join City ON E.CityCode = City.CityCode) " & _
            "LEFT JOIN Job_Card JC on E.Job_DocID = JC.DocId) " & _
            "LEFT JOIN Syctrl ON Syctrl.LinkTable<>" & xIsNull("E.U_AE", "") & " " & _
        "where E.DocId='" & Master!SearchCode & "' "
    Else
        GSQL = "SELECT '1' as Orig," & cMID("E1.DocID", "4", "5") & " as V_Type,E1.DocID," & vIsNull("E1.Sr_No", "0") & " as Sr_No,E.Job_DocID,E.v_Date,E.Model,E.RegNo,E.Chassis,E.Engine,E.Party_Code,E.NamePrefix,E.Party_Name,E.Address,E.Address2,E.Address3,City.CityName,E.PhoneNo," & _
            "E.SprAmt_MRP_TB,E.SprAmt_MRP_TP,E.OilAmt_MRP_TB,E.OilAmt_MRP_TP,E.SprAmt_TB,E.SprAmt_TP,E.OilAmt_TB, E.OilAmt_TP,E.D_Per_TB, E.D_Amt_TB, E.D_Per_TP,E.D_Amt_TP,E.D_Per_MRP_TB,E.D_Amt_MRP_TB," & _
            "E.D_Per_MRP_TP,E.D_Amt_MRP_TP,E.Addition,E.Gen_Sur_Per,E.Gen_Sur_Amt,E.Trans_Amt,E.Tax_Per, E.Tax_Amt, E.Tax_AmtMRP, E.Tax_Sur_Per,E.Tax_Sur_Amt,E.TaxSur_AmtMRP,E.Packing, E.TOT_Per, E.Tot_Amt, E.TOT_AmtMRP,E.ReSalTax_Per,E.ReSalTax_Amt,E.Total_Amt," & _
            "E.Rounded,E.Det_Tax,E.Printed_YN,E.U_Name,E.U_EntDt,E.Lab_Amt,E.Lab_D_Amt,E.Lab_TaxPer,E.Lab_TaxAmt,E.Lab_Rounded As Lab_RoundOff,E.Lab_Total_Amt," & _
            "E1.Part_No,P.Part_Name," & vIsNull("E1.Qty", "0") & " as Qty," & _
            " " & vIsNull("E1.Tax_YN", "0") & " as Tax_YN," & vIsNull("E1.MRP_YN", "0") & " as MRP_YN," & vIsNull("E1.Rate", "0") & " as Rate," & _
            " " & vIsNull("E1.MRP_Rate", "0") & " as MRP_Rate," & _
            " " & vIsNull("E1.Disc_Per", "0") & " as Disc_Per," & vIsNull("E1.Disc_Amt", "0") & " as Disc_Amt," & vIsNull("E1.Amount", "0") & " as Amount," & _
            "Syctrl.EstiInvFooter, E.Remarks, E1.TaxPer, E1.TaxAmt, E1.ReqNo,E.PrintTitle,E.SatAmt " & _
        "FROM (((((Estimate as E left JOIN Estimate1 as E1 ON E.DocID = E1.DocId) " & _
            "left JOIN Part as P ON E1.Part_No = P.Part_No and P.Div_Code = left(E1.Docid,1)) " & _
            "LEFT JOIN SubGroup as SG ON E.Party_Code = SG.SubCode) " & _
            "left join City ON E.CityCode = City.CityCode) " & _
            "LEFT JOIN Job_Card JC on E.Job_DocID = JC.DocId) " & _
            "LEFT JOIN Syctrl ON Syctrl.LinkTable<>" & xIsNull("E.U_AE", "") & " " & _
            "where E.DocId='" & Master!SearchCode & "' order by E1.Part_No"
    End If
        
    mQryLab = "SELECT '2' as Orig," & cMID("E.DocID", "4", "5") & " as V_Type,E.DocID," & vIsNull("E1.Sr_No", "0") & " as Sr_No,E.Job_DocID,E.v_Date,E.Model,E.RegNo,E.Chassis,E.Engine,E.Party_Code,E.NamePrefix,E.Party_Name,E.Address,E.Address2,E.Address3,City.CityName,E.PhoneNo," & _
        "0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
        "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
        "0 as Rounded,0 as Det_Tax,E.Printed_YN,E.U_Name,E.U_EntDt,E.Lab_Amt,E.Lab_D_Amt,E.Lab_TaxPer,E.Lab_TaxAmt,E.Lab_Rounded as Lab_RoundOff,E.Lab_Total_Amt, " & _
        "E1.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name," & _
        "E1.Hrs_Taken as Qty,E1.Tax_YN as Tax_YN,E1.MRP_YN," & vIsNull("E1.Rate", "0") & " as Rate," & _
        " " & vIsNull("E1.MRP_Rate", "0") & "  as MRP_Rate," & _
        "0 as Disc_Per,0 as Disc_Amt,E1.Lab_Charges as Amount," & _
        "Syctrl.EstiInvFooter,E.remarks, 0 As TaxPer, 0 As TaxAmt, '' As ReqNo,E.PrintTitle " & _
    "FROM (((((Estimate as E LEFT JOIN Estimate1 as E1 ON E.DocId = E1.DocID) " & _
        "left join SubGroup as SG ON E.Party_Code = SG.SubCode) " & _
        "LEFT JOIN City ON E.CityCode = City.CityCode) " & _
        "LEFT JOIN Labour ON E1.Lab_Code = Labour.Lab_Code) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>" & xIsNull("E.U_AE", "") & ") " & _
    "Where E.DocId='" & mLabDocId & "'"
    
    If ChkMerg.Value = 1 Then GSQL = GSQL & " Union All " & mQryLab & " Order By 1,2,3,Part_No"
        
        
    Set RstRep = GCn.Execute(GSQL)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    If mVType <> "S_QU" Then
        Set RST1 = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    Else
        Set RST1 = GCn.Execute("select S_SecSpeciality AS W_SecSpeciality,S_SecLST AS W_SecLST,S_SecLST_Date AS W_SecLST_Date,S_SecCST AS W_SecCST,S_SecCST_Date AS W_SecCST_Date,S_SecPhone AS W_SecPhone,S_SecFax AS W_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    End If
    
    CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".TTX", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("LST")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecFax & "'"
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecSpeciality & "'"
            Case UCase("TOTCaption")
                rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
            Case UCase("Title")
                rpt.FormulaFields(I).TEXT = "'" & txt(PrintTitle) & "'"
        End Select
    Next
    rpt.Database.SetDataSource RstRep
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
'                Case UCase("Title")
'                    rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            End Select
            Next
            rpt.PrintOut False
            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                GCn.Execute "update Estimate set Printed_YN = 1  where Estimate.docid='" & Master!SearchCode & "' "
            End If
        Case 1  'screen
            Call Report_View(rpt, txt(PrintTitle), , True)
    End Select
    CmdPrint(PSetUp).Tag = ""
    Set RST1 = Nothing
    Set RstRep = Nothing
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

Private Sub SpeedPrint(PrePrinted, Merge)
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
    Dim RstCompDet As ADODB.Recordset, RstEstimate As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject, mQryLab$
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double
    Dim MRPTaxStr$, mTPAmtStr$, mTBAmtStr$
    Dim mSprCaption As Boolean, mLabCaption As Boolean, mLabDiscAmtStr$
    Dim mLabDocId$
    Dim mLabDiscAmt As Double, mLabAmt As Double, mNetLabAmt As Double
    Dim mServTaxPer As Double, mServTaxAmt As Double, mLabROffAmt As Double, mNetAmt As Double
    Dim mRtPart1$, mRtPart2$, mRtPart3$, mRtPart4$, mRtPart5$, mRtPart6$
    
    If ChkMerg.Value = 1 Then
        If txt(JCNo).Tag & txt(RegNo) <> "" Then
            If txt(JCNo).Tag <> "" Then
                GSQL = "Select top 1 DocID from Estimate where Job_DocID='" & txt(JCNo).Tag & "' and V_Type='W_PL'"
            Else
                If CmboPLNo.TEXT = "" Then
                    MsgBox "Enter the Performa No to Merze.", vbOKOnly + vbInformation, "Validation Message"
                    CmboPLNo.SetFocus: Exit Sub
                    'TxtPerformaNo.SetFocus: Exit Sub
                Else
                    GSQL = "Select top 1 DocID from Estimate where RegNo='" & txt(RegNo) & "' and V_No=" & Val(CmboPLNo.TEXT) & " and V_Type='W_PL'"
                End If
            End If
            If GCn.Execute(GSQL).RecordCount > 0 Then
                mLabDocId = GCn.Execute(GSQL).Fields(0).Value
            Else
                MsgBox "Labour Performa Not Feeded." & vbCrLf & "Merge Printing not possible!", vbCritical, "Validation": Exit Sub
            End If
        Else
            MsgBox "Job No./Reg. No. Not Feeded." & vbCrLf & "Merge Printing not possible!", vbCritical, "Validation": Exit Sub
        End If
    End If
    
    GSQL = "SELECT '1' as Orig," & cMID("E.DocID", "4", "5") & " as V_Type,E.DocID," & vIsNull("E1.Sr_No", "0") & " as Sr_No,E.Job_DocID,E.v_Date,E.Model,E.RegNo,E.Chassis,E.Engine,E.Party_Code,E.NamePrefix,E.Party_Name,E.Address,E.Address2,E.Address3,City.CityName,E.PhoneNo," & _
        "E.SprAmt_MRP_TB,E.SprAmt_MRP_TP,E.OilAmt_MRP_TB,E.OilAmt_MRP_TP,E.SprAmt_TB,E.SprAmt_TP,E.OilAmt_TB, E.OilAmt_TP,E.D_Per_TB, E.D_Amt_TB, E.D_Per_TP,E.D_Amt_TP,E.D_Per_MRP_TB,E.D_Amt_MRP_TB," & _
        "E.D_Per_MRP_TP,E.D_Amt_MRP_TP,E.Addition,E.Gen_Sur_Per,E.Gen_Sur_Amt,E.Trans_Amt,E.Tax_Per, E.Tax_Amt, E.Tax_AmtMRP, E.Tax_Sur_Per,E.Tax_Sur_Amt,E.TaxSur_AmtMRP,E.Packing, E.TOT_Per, E.Tot_Amt, E.TOT_AmtMRP,E.ReSalTax_Per,E.ReSalTax_Amt,E.Total_Amt," & _
        "E.Rounded,E.Det_Tax,E.Printed_YN,E.U_Name,E.U_EntDt,0 as Lab_Amt, 0 as Lab_D_Amt, 0 as Lab_TaxPer, 0 as Lab_TaxAmt,0 as Lab_RoundOff,0 as Lab_Total_Amt," & _
        "E1.Part_No as Part_No,P.Part_Name," & vIsNull("E1.Qty", "0") & " as Qty," & _
        " " & vIsNull("E1.Tax_YN", "0") & " as Tax_YN," & vIsNull("E1.MRP_YN", "0") & " as MRP_YN," & vIsNull("E1.Rate", "0") & " as Rate," & _
        " " & vIsNull("E1.MRP_Rate", "0") & " as MRP_Rate," & _
        " " & vIsNull("E1.Disc_Per", "0") & " as Disc_Per," & vIsNull("E1.Disc_Amt", "0") & " as Disc_Amt, " & vIsNull("E1.Amount", "0") & " as Amount, " & _
        "Syctrl.EstiInvFooter,E.Remarks,E1.SatPer, E1.SatAmt, E.SatAmt As SatAmt_H " & _
    "FROM (((((Estimate as E left JOIN Estimate1 as E1 ON E.DocID = E1.DocId) " & _
        "left JOIN Part as P ON E1.Part_No = P.Part_No and P.Div_Code = left(E1.Docid,1)) " & _
        "LEFT JOIN SubGroup as SG ON E.Party_Code = SG.SubCode) " & _
        "left join City ON E.CityCode = City.CityCode) " & _
        "LEFT JOIN Job_Card JC on E.Job_DocID = JC.DocId) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>E.U_AE " & _
    "where E.DocId='" & Master!SearchCode & "'"

    mQryLab = "SELECT '2' as Orig," & cMID("E.DocID", "4", "5") & " as V_Type,E.DocID," & vIsNull("E1.Sr_No", "0") & " as Sr_No,E.Job_DocID,E.v_Date,E.Model,E.RegNo,E.Chassis,E.Engine,E.Party_Code,E.NamePrefix,E.Party_Name,E.Address,E.Address2,E.Address3,City.CityName,E.PhoneNo," & _
        "0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
        "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
        "0 as Rounded,0 as Det_Tax,E.Printed_YN,E.U_Name,E.U_EntDt,E.Lab_Amt,E.Lab_D_Amt,E.Lab_TaxPer,E.Lab_TaxAmt,E.Lab_Rounded as Lab_RoundOff,E.Lab_Total_Amt, " & _
        "E1.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name," & _
        "E1.Hrs_Taken as Qty,E1.Tax_YN as Tax_YN,E1.MRP_YN," & vIsNull("E1.Rate", "0") & " as Rate," & _
        " " & vIsNull("E1.MRP_Rate", "0") & "  as MRP_Rate," & _
        "0 as Disc_Per,0 as Disc_Amt,E1.Lab_Charges as Amount," & _
        "Syctrl.EstiInvFooter,E.remarks,0 As SatPer, 0 As SatAmt, 0 As SatAmt_H " & _
    "FROM (((((Estimate as E LEFT JOIN Estimate1 as E1 ON E.DocId = E1.DocID) " & _
        "left join SubGroup as SG ON E.Party_Code = SG.SubCode) " & _
        "LEFT JOIN City ON E.CityCode = City.CityCode) " & _
        "LEFT JOIN Labour ON E1.Lab_Code = Labour.Lab_Code) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>E.U_AE) " & _
    "Where E.DocId='" & mLabDocId & "'"
    
    If ChkMerg.Value = 1 Then
        GSQL = GSQL & " Union All " & mQryLab & " Order By 1,2,3,Part_No"
        
        mQryLab = "SELECT E.Lab_Amt,E.Lab_D_Amt,E.Lab_TaxPer,E.Lab_TaxAmt,E.Lab_Rounded as Lab_RoundOff,E.Lab_Total_Amt " & _
            "FROM Estimate as E Where E.DocId='" & mLabDocId & "'"

        Set GRs = New Recordset
        Set GRs = GCn.Execute(mQryLab)
        If GRs.RecordCount > 0 Then
            mLabAmt = IIf(IsNull(GRs!Lab_Amt), 0, GRs!Lab_Amt)
            mLabDiscAmt = IIf(IsNull(GRs!Lab_D_Amt), 0, GRs!Lab_D_Amt)
            mServTaxPer = IIf(IsNull(GRs!Lab_TaxPer), 0, GRs!Lab_TaxPer)
            mServTaxAmt = IIf(IsNull(GRs!Lab_TaxAmt), 0, GRs!Lab_TaxAmt)
            mLabROffAmt = IIf(IsNull(GRs!Lab_RoundOff), 0, GRs!Lab_RoundOff)
            mNetLabAmt = IIf(IsNull(GRs!Lab_Total_Amt), 0, GRs!Lab_Total_Amt)
        End If
        Set GRs = Nothing
    Else
        GSQL = GSQL & " Order By 1,2,3,E1.Part_No"
        mLabAmt = Val(txt(LabAmt))
        mLabDiscAmt = Val(txt(LabDisc))
        mServTaxPer = Val(txt(ServTaxPer))
        mServTaxAmt = Val(txt(ServTaxAmt))
        mLabROffAmt = Val(txt(LabROff))
        mNetLabAmt = Val(txt(NetLabAmt))
    End If
    mNetAmt = Val(txt(NetSprAmt)) + mNetLabAmt
    
    Set RstEstimate = New Recordset
    Set RstEstimate = GCn.Execute(GSQL)
    
    If RstEstimate.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select EstiInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
 
    PageLength = PubPageLength
    PageWidth = 80 - Len(mSP2)
    mHeader = 0   'Ideal 17
    mFooter = IIf(mVType = "S_QU", 18, 21)
    mFooter = mFooter + FooterCnt
   'mFooter = IIf(mVType = SalCrVType And RstEstimate!Printed_YN = 0, mFooter + mGatePass, mFooter)
    'Sale Bill Header
    If txt(PrintTitle) <> "" Then
        mDocStr = Trim(txt(PrintTitle))
    Else
        mDocStr = IIf(mVType = "S_QU", "SPARE QUOTATION", "WORKSHOP ESTIMATE")
    End If
    mDupStr = IIf(RstEstimate!Printed_YN = 1, "(DUPLICATE)", "")
    If (mMRPTax + mMRPTaxSur + mMRPTOT) > 0 Then
        MRPTaxStr = "* Note:"
        If (mMRPTax + mMRPTaxSur) > 0 Then
            MRPTaxStr = MRPTaxStr & "Sales Tax Rs." & mMRPTax & ",Surcharge Rs." & mMRPTaxSur
        End If
        If (mMRPTOT) > 0 Then
            MRPTaxStr = MRPTaxStr & " Turn Over Tax " & mMRPTOT
        End If
        MRPTaxStr = MRPTaxStr & " already added in MRP *'"
    End If
    
    If mVType <> "S_QU" Then
        Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
        mRtPart1 = mEmph & PSTR("Document No.", 12) & " : " & PrinID(RstEstimate!DocID) & mEmph1
        mRtPart2 = PSTR("DATE", 12, , AlignRight) & " : " & STR(RstEstimate!V_DATE)
        If IsNull(RstEstimate!job_docid) Or RstEstimate!job_docid = "" Then
            mRtPart3 = PSTR("Job Card No.", 12) & " : "
        Else
            mRtPart3 = PSTR("Job Card No.", 12) & " : " & PrinID(RstEstimate!job_docid)
        End If
        mRtPart4 = PSTR("Chassis No.", 12) & " : " & XNull(RstEstimate!Chassis)
        mRtPart5 = PSTR("Vehicle No.", 12) & " : " & XNull(RstEstimate!RegNo)
        mRtPart6 = PSTR("Model", 12) & " : " & XNull(RstEstimate!Model)
    Else
        Set RstCompDet = GCn.Execute("select S_SecSpeciality AS W_SecSpeciality,S_SecLST AS W_SecLST,S_SecLST_Date AS W_SecLST_Date,S_SecCST AS W_SecCST,S_SecCST_Date AS W_SecCST_Date,S_SecPhone AS W_SecPhone,S_SecFax AS W_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
        mRtPart2 = mEmph & PSTR("Document No.", 12) & " : " & PrinID(RstEstimate!DocID) & mEmph1
        mRtPart3 = PSTR("DATE", 12, , AlignRight) & " : " & STR(RstEstimate!V_DATE)
    End If
        '*********
    If PrePrinted Then
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        mHeader = 8
    Else
        Print #1, mSP2 & PRN_TIT(PubComp_Name, "A", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCompDet!W_SecSpeciality) <> "" Then
            Print #1, mSP2 & PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, mSP2 & PRN_TIT(PubComp_Add, "C", PageWidth)
        mHeader = mHeader + 1
        If PubComp_Add2 <> "" Then
            Print #1, mSP2 & PRN_TIT(PubComp_Add2, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If PubComp_City <> "" Then
            Print #1, mSP2 & PRN_TIT(PubComp_City, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
    End If
    Print #1, mSP2 & PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, mSP2 & mChr18 & Space(46) & mEmph & mRtPart1 & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP2 & PSTR("To,", 46) & mRtPart2 & mEmph
    mHeader = mHeader + 1
    Print #1, mSP2 & PSTR(XNull(RstEstimate!NamePrefix) & " " & RstEstimate!Party_Name, 44) & mEmph1 & Space(2) & mRtPart3
    mHeader = mHeader + 1
    Print #1, mSP2 & PSTR(XNull(RstEstimate!Address), 40) & Space(6) & mEmph & mRtPart4 & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP2 & PSTR(XNull(RstEstimate!Address2), 40) & Space(6) & mRtPart5
    mHeader = mHeader + 1
    Print #1, mSP2 & PSTR(XNull(RstEstimate!Address3) & IIf(XNull(RstEstimate!Address3) <> "" And XNull(RstEstimate!CityName) <> "", ",", "") & XNull(RstEstimate!CityName), 44) _
    & Space(2) & mRtPart6
    mHeader = mHeader + 1
    Print #1, mSP2 & "Remarks : " & PSTR(XNull(RstEstimate!Remarks), 50)
    mHeader = mHeader + 1
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
    mHeader = mHeader + 1
    If RstEstimate!Det_Tax = 1 Then
        Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
        mHeader = mHeader + 1
        Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
        mHeader = mHeader + 1
    Else
        Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 27) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18 '& mDoub1
        mHeader = mHeader + 1
    End If
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    If RstEstimate!orig = "1" Then
        Print #1, mSP2 & mEmph & "*Spare Details*" & mEmph1 & mChr17
        mHeader = mHeader + 1
        mSprCaption = True
    ElseIf RstEstimate!orig = "2" Then
        Print #1, mSP2 & mEmph & "*Labour Details*" & mEmph1 & mChr17
        mHeader = mHeader + 1
        mLabCaption = True
    End If
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    LAdd = VNull(RstEstimate!Gen_Sur_Amt) + VNull(RstEstimate!Trans_Amt) + VNull(RstEstimate!Tax_Amt) + VNull(RstEstimate!Tax_Sur_Amt) + VNull(RstEstimate!Packing) + VNull(RstEstimate!ReSalTax_Amt) + VNull(RstEstimate!Tot_Amt)
    SubTot = RstEstimate!SprAmt_TB + RstEstimate!SprAmt_TP + RstEstimate!SprAmt_MRP_TB + RstEstimate!SprAmt_MRP_TP _
        + RstEstimate!OilAmt_TB + RstEstimate!OilAmt_TP + Val(txt(IWDiscTotTP).TEXT) + Val(txt(IWDiscTotTB).TEXT)
    If RstEstimate.RecordCount > 0 Then
        I = 1
        Do Until RstEstimate.EOF
            If mLine > mFix Then
                Page = Page + 1
                Print #1, mChr18 & mSP2 & Replace(Space(PageWidth), " ", "-")
                Print #1, mSP2 & Space((PageWidth) - Len("Contd. on next page.." + STR(Page))) & "Contd. on next page.." & STR(Page)
                Do Until mLine >= mFix + mFooter - 2
                     Print #1, ""
                    mLine = mLine + 1
                Loop
                Print #1, mEject
                
                'Header On Second Page
                mHeader = 0
                If Not FirstPrint Then
                    Print #1, ""
                    FirstPrint = True
                End If
                If PrePrinted Then
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    mHeader = 8
                Else
                    Print #1, mSP2 & PRN_TIT(PubComp_Name, "A", PageWidth)
                    mHeader = mHeader + 1
                    If XNull(RstCompDet!W_SecSpeciality) <> "" Then
                        Print #1, mSP2 & PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
                        mHeader = mHeader + 1
                    End If
                End If
                Print #1, mSP2 & PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, mSP2 & mChr18 & Space(46) & mEmph & mRtPart1 & mEmph1
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR("To,", 46) & mRtPart2 & mEmph
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR(XNull(RstEstimate!NamePrefix) & " " & RstEstimate!Party_Name, 44) & mEmph1 & Space(2) & mRtPart3
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR(XNull(RstEstimate!Address), 40) & Space(6) & mEmph & mRtPart4 & mEmph1
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR(XNull(RstEstimate!Address2), 40) & Space(6) & mRtPart5
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR(XNull(RstEstimate!Address3) & IIf(XNull(RstEstimate!Address3) <> "" And XNull(RstEstimate!CityName) <> "", ",", "") & XNull(RstEstimate!CityName), 44) _
                & Space(2) & mRtPart6
                mHeader = mHeader + 1
                Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
                mHeader = mHeader + 1
                If RstEstimate!Det_Tax = 1 Then
                    Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                    mHeader = mHeader + 1
                    Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
                    mHeader = mHeader + 1
                Else
                    Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 27) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18 '& mDoub1
                    mHeader = mHeader + 1
                End If
                Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            If mLabCaption = False Then
                If RstEstimate!orig = "2" Then
                    Print #1, mSP2 & mChr18 & mEmph & "*Labour Details*" & mEmph1 & mChr17
                    mHeader = mHeader + 1
                    mLabCaption = True
                End If
            End If
            mRate = IIf(RstEstimate!MRP_YN = 1, RstEstimate!MRP_Rate, RstEstimate!Rate)
            If RstEstimate!orig = "1" Then
                If RstEstimate!Det_Tax = 1 Then
                    mTPAmtStr = PSTR(0, 12, 2)
                    mTBAmtStr = PSTR(0, 12, 2)
                    If RstEstimate!Tax_YN = 0 Then
                        mTPAmtStr = PSTR(RstEstimate!Amount, 12, 2)
                        mTBAmtStr = PSTR(0, 12, 2)
                    Else
                        mTPAmtStr = PSTR(0, 12, 2)
                        mTBAmtStr = PSTR(RstEstimate!Amount, 12, 2)
                    End If
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstEstimate!Part_No, 22, , AlignLeft) & PSTR(RstEstimate!Part_Name, 34) & PSTR(RstEstimate!Qty, 12, 2)
                    PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstEstimate!MRP_YN = 1, "M", IIf(RstEstimate!MRP_YN = 0, "L", "")) & _
                    PSTR(RstEstimate!Disc_Per, 8, 2) & " %" & PSTR(RstEstimate!Disc_Amt, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                Else
                    LAmtItem = RstEstimate!Amount + RstEstimate!Disc_Amt
                    LDAmt = LAmtItem + (LAmtItem * (LAdd / IIf(SubTot = 0, 1, SubTot)))
                    LAmtVal = LAmtVal + (LAmtItem * (LAdd / IIf(SubTot = 0, 1, SubTot)))
                    LdRate = LDAmt / IIf(RstEstimate!Qty = 0, 1, RstEstimate!Qty)
                    If I = RstEstimate.RecordCount Then
                        If LAmtVal <> LAdd Then LDAmt = LDAmt + (LAdd - LAmtVal)
                        LdRate = LDAmt / IIf(RstEstimate!Qty = 0, 1, RstEstimate!Qty)
                    End If
                    mGrossAmt = mGrossAmt + (LDAmt - RstEstimate!Disc_Amt)
                    I = I + 1
                    mAmount = Round(RstEstimate!Qty * RstEstimate!Rate, 2) - RstEstimate!Disc_Amt
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstEstimate!Part_No, 26, , AlignLeft) & PSTR(RstEstimate!Part_Name, 40) & PSTR(RstEstimate!Qty, 12, 2)
                    PrintStr = PrintStr & PSTR(LdRate, 11, 2) & " " & IIf(RstEstimate!MRP_YN = 1, "M", "L") & _
                    PSTR(RstEstimate!Disc_Per, 8, 2) & " %" & PSTR(RstEstimate!Disc_Amt, 10, 2) & _
                    PSTR(LDAmt - RstEstimate!Disc_Amt, 12, 2)
                End If
            Else    'Labour
                If Val(txt(ServTaxAmt)) <= 0 Then
                    mTPAmtStr = PSTR(RstEstimate!Amount, 12, 2)
                    mTBAmtStr = PSTR(0, 12, 2)
                Else
                    mTPAmtStr = PSTR(0, 12, 2)
                    mTBAmtStr = PSTR(RstEstimate!Amount, 12, 2)
                End If
                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstEstimate!Part_No, 22, , AlignLeft) & PSTR(RstEstimate!Part_Name, 34) & PSTR(Round(RstEstimate!Qty, 2), 12, 2)
                PrintStr = PrintStr & PSTR(mRate, 11, 2) & "  " & _
                PSTR(RstEstimate!Disc_Per, 8, 2) & " %" & PSTR(RstEstimate!Disc_Amt, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
            End If
            Print #1, mSP2 & PrintStr
            mSlNo = mSlNo + 1
            mLine = mLine + 1
NXT:
            RstEstimate.MoveNext
'            mSlNo = mSlNo + 1
'            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop

    Print #1, mChr18 & mSP2 & "Customer's Signature"
' SALE FOOTER
    '22 space maintain between heading and :
    RstEstimate.MoveFirst
    If RstEstimate!Det_Tax = 1 Then
        Print #1, mSP2 & Replace(Space(20), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
    
        Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
        ; " | " & PSTR(IIf(mVatYn, "VAT", "Sales Tax "), 10, 0) & PSTR(RstEstimate!Tax_Per, 5, 2) & "%" & PSTR(RstEstimate!Tax_Amt, 12, 2) & mDoub
        
        Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstEstimate!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstEstimate!SprAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & IIf(VNull(RstEstimate!SatAmt_H) > 0, PSTR("S A T ", 16, 0) & PSTR(RstEstimate!SatAmt_H, 12, 2), PSTR("Tax Surc. ", 10, 0) & PSTR(RstEstimate!Tax_Sur_Per, 5, 2) & "%" & PSTR(RstEstimate!Tax_Sur_Amt, 12, 2)) & mDoub
      
        Print #1, mSP2 & PSTR("Spares Amount", 16) & PSTR(RstEstimate!SprAmt_TP, 11, 2) & Space(8) & PSTR(RstEstimate!SprAmt_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Misc. Charges", 16) & PSTR(RstEstimate!Packing, 12, 2) & mDoub
        Print #1, mSP2 & PSTR("Oil Amount ", 16) & PSTR(RstEstimate!OilAmt_TP, 11, 2) & Space(8) & PSTR(RstEstimate!OilAmt_TB, 12, 2) & mDoub1 _
        ; " | " & mEmph & PSTR("Sub Total[TP+TB]", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
        
        Print #1, mSP2 & PSTR("Discount ", 10, 0) & PSTR(RstEstimate!D_Per_TP, 5, 2) & "%" & PSTR(RstEstimate!D_Amt_TP, 11, 2) & PSTR(RstEstimate!D_Per_TB, 7, 2) & "%" & PSTR(RstEstimate!D_Amt_TB, 12, 2) _
        ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstEstimate!TOT_Per, 5, 2) & "%" & PSTR(RstEstimate!Tot_Amt, 12, 2) & mEmph
        
        Print #1, mSP2 & PSTR("Sub Total [A]", 16) & PSTR(Val(txt(STotATP)), 11, 2) & Space(8) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
        ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstEstimate!ReSalTax_Per, 5, 2) & "%" & PSTR(RstEstimate!ReSalTax_Amt, 12, 2)
        
        Print #1, mSP2 & PSTR("Gen Surch ", 10, 0) & PSTR(RstEstimate!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 11, 2) & PSTR(RstEstimate!Gen_Sur_Amt, 20, 2) _
        ; " | " & PSTR("Round Off", 16) & PSTR(Round(RstEstimate!Rounded, 2), 12, 2)
       
        Print #1, mSP2 & PSTR("Transportation", 16) & PSTR(0, 11, 2) & PSTR(RstEstimate!Trans_Amt, 20, 2) _
        ; " | " & mEmph & PSTR("Net Spare Rs.", 16) & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
    Else
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mDoub
        Print #1, mSP2 & Space(44) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
        If RstEstimate!D_Amt_TP + RstEstimate!D_Amt_TB > 0 Then
            Print #1, mSP2 & Space(44) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstEstimate!D_Amt_TP + RstEstimate!D_Amt_TB, 12, 2)
        Else
            Print #1, ""
        End If
'        Print #1, mSP2 & Space(44) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Val(Txt(NetSprAmt)) - (mGrossAmt - (RstEstimate!D_Amt_TP + RstEstimate!D_Amt_TB)), 12, 2) & mEmph
        Print #1, mSP2 & Space(44) & PSTR("Round Off  ", 20) & " : " & PSTR(Round(RstEstimate!Rounded, 2), 12, 2) & mEmph
        Print #1, mSP2 & Space(44) & PSTR("Net Spare Rs.", 20) & " : " & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
    End If
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    If mVType = "W_EST" Then
        If mLabDiscAmt > 0 Then
            mLabDiscAmtStr = "Discount : " & PSTR(mLabDiscAmt, 8, 2)
        Else
            mLabDiscAmtStr = Space(19)
        End If
        PrintStr = "Total Labour      : " & PSTR(mLabAmt, 8, 2)
        PrintStr = PrintStr & " | " & mLabDiscAmtStr & " | " & mEmph & "Net Labour Rs.: " & PSTR(mNetLabAmt, 9, 2) & mEmph1
        Print #1, mSP2 & PrintStr
        PrintStr = "Service Tax @" & PSTR(mServTaxPer, 5, 2) & ": " & PSTR(mServTaxAmt, 8, 2)
        PrintStr = PrintStr & " | " & "Round Off: " & PSTR(mLabROffAmt, 8, 2) & " | " & mEmph & "Net Amount Rs.: " & PSTR(mNetAmt, 9, 2) & mEmph1
        Print #1, mSP2 & PrintStr
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    End If
    Print #1, mSP2 & mDoub & ntow(mNetAmt, "Rupees", "Paise") & mDoub1
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    Print #1, mSP2 & mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
    Print #1, mSP2 & PSTR(mTaxdesc, 25) & Space((PageWidth) - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, ""
    Print #1, mSP2 & mDoub & "Terms & Condition:" & mDoub1 & Replace(Space((PageWidth) - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
       If mID(Footer, I, 1) = vbLf Then
           Print #1, mSP2 & RTrim(mID(Footer, j, I - j))
           j = I + 1
       End If
    Next
    Print #1, mSP2 & Space((((PageWidth) * 1.7) - Len("* a dataman software *" & "   " & pubUName & "   " & PubServerDate)) / 2) & "* a dataman software *" & "   " & pubUName & "   " & PubServerDate & mChr18
    Print #1, mEject
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
        GCn.Execute "update Estimate set Printed_YN = 1  where Estimate.docid='" & Master!SearchCode & "' "
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

'modishekhar
Private Sub CtrlEnable(Enb As Boolean)
Dim I As Integer
For I = 43 To 48
   txt(I).Enabled = Enb
   If Enb = False Then txt(I).TEXT = ""
Next
txt(47).Enabled = False
txt(48).Enabled = False
End Sub
Private Sub FooterValue()
    MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
            Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
    If mVatYn = 1 Then
       MainLib.SprCalcVAT WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt)
    Else
        MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_Purpose, True
            ', _
            txt (LabAmt), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt)
    End If
End Sub
Private Sub SpeedPrintJMK(PrePrinted, Merge)
'On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per Page 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer, K As Integer
    Dim PrintStr As String
    Dim tmprs As ADODB.Recordset, HlpLineNo$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double, SrvTaxNo$
    Dim SrvGatePassOn$, Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double
    Dim MRPTaxStr$, mTPAmtStr$, mTBAmtStr$
    Dim mSprCaption As Boolean, mLabCaption As Boolean, mLabDiscAmtStr$, mQryLab$
    Dim mTotRow, mTotRowTemp As Integer
    Dim mLabDocId$
    Dim mNetTotal As Double
    Dim mLabAmt, mLabDiscAmt, mServTaxPer, mServTaxAmt, mLabROffAmt, mNetLabAmt As Double
   
    If ChkMerg.Value = 1 Then
        If txt(JCNo).Tag & txt(RegNo) <> "" Then
            If txt(JCNo).Tag <> "" Then
                'GSQL = "Select DocID from Estimate where Job_DocID='" & Txt(JCNo).Tag & "' and V_No=" & Val(CmboPLNo) & " and V_Type='W_PL'"
                GSQL = "Select DocID from Estimate where RegNo='" & txt(RegNo) & "' and V_No=" & Val(CmboPLNo) & " and V_Type='W_PL'"
            Else
                If CmboPLNo = "" Then
                    MsgBox "Enter the Performa No to Merze.", vbOKOnly + vbInformation, "Validation Message"
                    CmboPLNo.SetFocus: Exit Sub
                Else
                    GSQL = "Select DocID from Estimate where RegNo='" & txt(RegNo) & "' and V_No=" & Val(CmboPLNo) & " and V_Type='W_PL'"
                End If
            End If
            If GCn.Execute(GSQL).RecordCount > 0 Then
                mLabDocId = GCn.Execute(GSQL).Fields(0).Value
            Else
                MsgBox "Labour Performa Not Feeded." & vbCrLf & "Merge Printing not possible!", vbCritical, "Validation": Exit Sub
            End If
        Else
            MsgBox "Job No./Reg. No. Not Feeded." & vbCrLf & "Merge Printing not possible!", vbCritical, "Validation": Exit Sub
        End If
    End If
    
    GSQL = "SELECT '1' as Orig, " & cMID("E.DocID", "4", "5") & " as V_Type,E.DocID, " & vIsNull("E1.Sr_No", "0") & " as Sr_No,E.Job_DocID,E.v_Date,E.Model,E.RegNo,E.Chassis,E.Engine,E.Party_Code,E.NamePrefix,E.Party_Name,E.Address,E.Address2,E.Address3,City.CityName,E.PhoneNo," & _
        "E.SprAmt_MRP_TB,E.SprAmt_MRP_TP,E.OilAmt_MRP_TB,E.OilAmt_MRP_TP,E.SprAmt_TB,E.SprAmt_TP,E.OilAmt_TB, E.OilAmt_TP,E.D_Per_TB, E.D_Amt_TB, E.D_Per_TP,E.D_Amt_TP,E.D_Per_MRP_TB,E.D_Amt_MRP_TB," & _
        "E.D_Per_MRP_TP,E.D_Amt_MRP_TP,E.Addition,E.Gen_Sur_Per,E.Gen_Sur_Amt,E.Trans_Amt,E.Tax_Per, E.Tax_Amt, E.Tax_AmtMRP, E.Tax_Sur_Per,E.Tax_Sur_Amt,E.TaxSur_AmtMRP,E.Packing, E.TOT_Per, E.Tot_Amt, E.TOT_AmtMRP,E.ReSalTax_Per,E.ReSalTax_Amt,E.Total_Amt," & _
        "E.Rounded,E.Det_Tax,E.Printed_YN,E.U_Name,E.U_EntDt,0 as Lab_Amt, 0 as Lab_D_Amt, 0 as Lab_TaxPer, 0 as Lab_TaxAmt,0 as Lab_RoundOff,0 as Lab_Total_Amt," & _
        "E1.Part_No,P.Part_Name, " & vIsNull("E1.Qty", "0") & " as Qty," & _
        "" & vIsNull("E1.Tax_YN", "0") & " as Tax_YN, " & vIsNull("E1.MRP_YN", "0") & " as MRP_YN, " & vIsNull("E1.Rate", "0") & " as Rate," & _
        "" & vIsNull("E1.MRP_Rate", "0") & " as MRP_Rate," & _
        "" & vIsNull("E1.Disc_Per", "0") & " as Disc_Per, " & vIsNull("E1.Disc_Amt", "0") & " as Disc_Amt, " & vIsNull("E1.Amount", "0") & " as Amount," & _
        "Syctrl.EstiInvFooter,E.Remarks,E1.Amount as Net_Amt2,E1.ReqNo as ReqNo,JC.AtKmsHrs as KMS, E1.TaxPer, E1.TaxAmt, E1.SatPer, E1.SatAmt, E.SatAmt as SatAmt_H " & _
    "FROM (((((Estimate as E left JOIN Estimate1 as E1 ON E.DocID = E1.DocId) " & _
        "left JOIN Part as P ON E1.Part_No = P.Part_No and P.Div_Code = left(E1.Docid,1)) " & _
        "LEFT JOIN SubGroup as SG ON E.Party_Code = SG.SubCode) " & _
        "left join City ON E.CityCode = City.CityCode) " & _
        "LEFT JOIN Job_Card JC on E.Job_DocID = JC.DocId) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>E.U_AE " & _
    "where E.DocId='" & Master!SearchCode & "'"

    mQryLab = "SELECT '2' as Orig, " & cMID("E.DocID", "4", "5") & " as V_Type,E.DocID, " & vIsNull("E1.Sr_No", "0") & " as Sr_No,E.Job_DocID,E.v_Date,E.Model,E.RegNo,E.Chassis,E.Engine,E.Party_Code,E.NamePrefix,E.Party_Name,E.Address,E.Address2,E.Address3,City.CityName,E.PhoneNo," & _
        "0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
        "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
        "0 as Rounded,0 as Det_Tax,E.Printed_YN,E.U_Name,E.U_EntDt,E.Lab_Amt,E.Lab_D_Amt,E.Lab_TaxPer,E.Lab_TaxAmt,E.Lab_Rounded as Lab_RoundOff,E.Lab_Total_Amt, " & _
        "E1.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name," & _
        "E1.Hrs_Taken as Qty,E1.Tax_YN as Tax_YN,E1.MRP_YN, " & vIsNull("E1.Rate", "0") & " as Rate," & _
        "" & vIsNull("E1.MRP_Rate", "0") & " as MRP_Rate," & _
        "0 as Disc_Per,0 as Disc_Amt,E1.Lab_Charges as Amount," & _
        "Syctrl.EstiInvFooter,E.remarks,E1.Lab_Charges as Net_Amt2,0 as ReqNo,'' as KMS, 0 As TaxPer, 0 As TaxAmt, 0 As SatPer, 0 As SatAmt, 0 As SatAmt_H " & _
    "FROM (((((Estimate as E LEFT JOIN Estimate1 as E1 ON E.DocId = E1.DocID) " & _
        "left join SubGroup as SG ON E.Party_Code = SG.SubCode) " & _
        "LEFT JOIN City ON E.CityCode = City.CityCode) " & _
        "LEFT JOIN Labour ON E1.Lab_Code = Labour.Lab_Code) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>E.U_AE) " & _
    "Where E.DocId='" & mLabDocId & "'"
        
    
    GSQL = GSQL & " Union All " & mQryLab & " Order By 3,2,1,Part_No"
    
        
        mQryLab = "SELECT E.Lab_Amt,E.Lab_D_Amt,E.Lab_TaxPer,E.Lab_TaxAmt,E.Lab_Rounded as Lab_RoundOff,E.Lab_Total_Amt " & _
            "FROM Estimate as E Where E.DocId='" & mLabDocId & "'"

        Set GRs = New Recordset
        Set GRs = GCn.Execute(mQryLab)
        If GRs.RecordCount > 0 Then
            
            mLabAmt = IIf(IsNull(GRs!Lab_Amt), 0, GRs!Lab_Amt)
            mLabDiscAmt = IIf(IsNull(GRs!Lab_D_Amt), 0, GRs!Lab_D_Amt)
            mServTaxPer = IIf(IsNull(GRs!Lab_TaxPer), 0, GRs!Lab_TaxPer)
            mServTaxAmt = IIf(IsNull(GRs!Lab_TaxAmt), 0, GRs!Lab_TaxAmt)
            mLabROffAmt = IIf(IsNull(GRs!Lab_RoundOff), 0, GRs!Lab_RoundOff)
            mNetLabAmt = IIf(IsNull(GRs!Lab_Total_Amt), 0, GRs!Lab_Total_Amt)
        End If
        Set GRs = Nothing
                
        
    Set RstJob = GCn.Execute(GSQL)
    If RstJob.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select WorkShopInvFooter from Syctrl").Fields(0).Value)
    SrvGatePassOn = XNull(GCn.Execute("select SrvGatePass_On from Syctrl").Fields(0).Value)

    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
 
    PageLength = PubPageLength
    PageWidth = 80 - Len(mSP2) '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
    mHeader = 0   'Ideal 17
    mFooter = 12    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
    mGatePass = 9
    mDetTax = 15
    mFooter = IIf(RstJob!Det_Tax = 1, mFooter, mDetTax)
    mFooter = mFooter + FooterCnt
    'modi lps 03-04-2003
'    mFooter = IIf(RstJob!Printed_yn = 0, mFooter + mGatePass, mFooter)
    If RstJob!Printed_YN = 0 Then   'Not Printed
        If PubSrvGatePass = 1 And SrvGatePassOn = "S" Then  'GatePass on Spare Bill Required
            mFooter = mFooter + mGatePass
        End If
    End If
    mFooter = 35
    'eof modi
    'Sale Bill Header
'    If Not Provisional Then
'        If mVatYn = 1 Then
'           mDocStr = "RETAIL INVOICE"
'        Else
'            If RstJob!CrMemo = 0 Then
'                mDocStr = "CASH MEMO"
'            Else
'                mDocStr = "INVOICE"
'            End If
'        End If
'    Else
'        mDocStr = "PROVISIONAL BILL "
'    End If
'
'    If Not Provisional Then
'        mDupStr = IIf(RstJob!Printed_YN = 1, "(DUPLICATE)", "")
'    Else
'        mDupStr = ""
'    End If

    Set tmprs = GCn.Execute("Select HelpLineNo from Syctrl")
    If tmprs.RecordCount > 0 Then
        HlpLineNo = IIf(IsNull(tmprs!HelpLineNo), "", Trim(tmprs!HelpLineNo))
        Set tmprs = Nothing
    End If


    If (mMRPTax + mMRPTaxSur + mMRPTOT) > 0 Then
        MRPTaxStr = "* Note:"
        If (mMRPTax + mMRPTaxSur) > 0 Then
            MRPTaxStr = MRPTaxStr & IIf(mVatYn = 1, "VAT Rs. ", "Sales Tax Rs.") & mMRPTax & ",Surcharge Rs." & mMRPTaxSur
        End If
        If (mMRPTOT) > 0 Then
            'MRPTaxStr = MRPTaxStr & pubTOTCaption & mMRPTOT
        End If
        MRPTaxStr = MRPTaxStr & " already added in MRP *'"
    End If
'    If GCn.Execute("select Printing_Desc from TaxForms where Form_Code = '" & RstJob!Form_Code & "'").RecordCount > 0 Then
'        mTaxdesc = GCn.Execute("select Printing_Desc from TaxForms where Form_Code = '" & RstJob!Form_Code & "'").Fields(0).Value
'    End If
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    If UCase(pubUName) <> "SA" And txt(OrgFrom) = "Invoice" Then
        mDupStr = IIf(RstJob!Printed_YN = 1, "(DUPLICATE)", "")
        If mDupStr <> "" Then
            MsgBox "Second Printing Can Be done Only By SA."
            Exit Sub
        End If
    End If
    If PrePrinted Then
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        mHeader = 8
    Else
        Print #1, mSP2 & PRN_TIT(PubComp_Name, "A", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCompDet!W_SecSpeciality) <> "" Then
            Print #1, mSP2 & PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, mSP2 & PRN_TIT(PubComp_Add, "C", PageWidth)
        mHeader = mHeader + 1
        If PubComp_Add2 <> "" Then
            Print #1, mSP2 & PRN_TIT(PubComp_Add2, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If PubComp_City <> "" Then
            Print #1, mSP2 & PRN_TIT(PubComp_City, "C", PageWidth)
            mHeader = mHeader + 1
        End If
    End If
        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        'Service tax No Printing............
        SrvTaxNo = XNull(GCn.Execute("Select SrvTaxNo from Syctrl").Fields(0).Value)
        Print #1, mSP2 & PSTR("Serv.Tax No.  : " & SrvTaxNo, 50, , AlignLeft) & PSTR("HelpLine No :" & HlpLineNo, 30, , AlignRight)
        mHeader = mHeader + 1
        '....................................
        If mVatYn = 1 Then
            Print #1, PRN_TIT("* " & txt(PrintTitle) & " *", "B", PageWidth)
            mHeader = mHeader + 1
        Else
            Print #1, PRN_TIT("*" & txt(PrintTitle) & "*", "B", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, ""
        mHeader = mHeader + 1
        'Print #1, mSP2 & mChr18 & "TO," & Space(33) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " : " & Right(RstJob!DocId, 6) & mEmph1
        Print #1, mSP2 & mChr18 & "TO," & Space(39) & mEmph & PSTR(Lbl(2).CAPTION, 22, , AlignRight) & " : " & Right(txt(DocID), 7) & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP2 & mEmph & PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(1) & PSTR("DATE", 12, , AlignRight) & "        : " & Format(txt(VDate), "dd/MMM/yyyy")
        mHeader = mHeader + 1
        
        Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(13) & PSTR("Job Card No.", 12) & ":" & Right(RstJob!job_docid, 6)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(RstJob!Address3) & IIf(XNull(RstJob!Address3) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44)
        mHeader = mHeader + 1
        Print #1, mSP2 & "Phone : " & PSTR(XNull(RstJob!PhoneNO), 20)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR("Chass.No.", 11) & PSTR(XNull(RstJob!Chassis), 15) & Space(1) & Space(6) & PSTR(XNull(RstJob!Model), 13) & Space(1) & PSTR("Reg.", 5) & PSTR(XNull(RstJob!RegNo), 14) & " Kms:" & PSTR(XNull(RstJob!Kms), 6)
        mHeader = mHeader + 1
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
        mHeader = mHeader + 1
        If mVatYn = 1 Then
            Print #1, mSP2 & PSTR("SrNo", 5) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 25) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP Rate", 11, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR(" DISC%", 6, , AlignRight) & PSTR("DISC. AMT", 12, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("Tax Amt", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18    '& mDoub1
            mHeader = mHeader + 1
        Else
            If RstJob!Det_Tax = 1 Then
                Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 27) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(6) & PSTR("RATE", 11, , AlignRight) & Space(5) & "<---------AMOUNT--------- >"
                mHeader = mHeader + 1
                Print #1, mSP2 & Space(113) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 10, , AlignRight) & mChr18     '& mDoub1
                mHeader = mHeader + 1
            Else
                Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 27) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(6) & PSTR("RATE", 11, , AlignRight) & Space(5) & "<---------AMOUNT--------- >"
                mHeader = mHeader + 1
                Print #1, mSP2 & Space(113) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 10, , AlignRight) & mChr18     '& mDoub1
                mHeader = mHeader + 1
            End If
        End If
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1

        Print #1, mSP2 & mChr18 & mEmph & "*Labour Details*" & Replace(Space(PageWidth - 16), " ", "-") & mEmph1 & mChr17
        mHeader = mHeader + 1
        
'        If RstJob!Orig = "1" Then
'            Print #1, mSP2 & mEmph & "*Spare Details*" & mEmph1 & mChr17
'            mHeader = mHeader + 1
'            mSprCaption = True
'        ElseIf RstJob!Orig = "2" Then
'            Print #1, mSP2 & mEmph & "*Labour Details*" & Replace(Space(PageWidth - 16), " ", "-") & mEmph1 & mChr17
'            mHeader = mHeader + 1
'            mLabCaption = True
'        End If
        mFix = PageLength - (mHeader)
        Page = 1
        mLine = 1
        mSlNo = 1
        LAdd = VNull(RstJob!Gen_Sur_Amt) + VNull(RstJob!Trans_Amt) + VNull(RstJob!Tax_Amt) + VNull(RstJob!Tax_Sur_Amt) + VNull(RstJob!Packing) + VNull(RstJob!ReSalTax_Amt) + VNull(RstJob!Tot_Amt)
        SubTot = RstJob!SprAmt_TB + RstJob!SprAmt_TP + RstJob!SprAmt_MRP_TB + RstJob!SprAmt_MRP_TP _
        + RstJob!OilAmt_TB + RstJob!OilAmt_TP + Val(txt(IWDiscTotTP).TEXT) + Val(txt(IWDiscTotTB).TEXT)
        mTotRow = RstJob.RecordCount
        mTotRowTemp = RstJob.RecordCount
        If RstJob.RecordCount > 0 Then
            I = 1
            Do Until RstJob.EOF = True
                If mTotRow > 30 Then
                    mFix = 20
                ElseIf mTotRow >= 15 And mTotRow <= 30 Then
                    mFix = 20
                Else
                    mFix = 20
                End If
                
                If mLine > mFix Then
                    Page = Page + 1
                    mTotRow = mTotRow - 30
                    Print #1, mChr18 & mSP2 & Replace(Space(PageWidth), " ", "-")
                    Print #1, mSP2 & Space((PageWidth) - Len("Contd. on next page.." + STR(Page))) & "Contd. on next page.." & STR(Page)
                    Do Until mLine >= (mFix + mFooter - 20)
                         Print #1, ""
                        mLine = mLine + 1
                    Loop
                    Print #1, mEject
                    
                    'Header On Second Page
                    mHeader = 0
                    If Not FirstPrint Then
                        Print #1, ""
                        FirstPrint = True
                    End If
                    If PrePrinted Then
                        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                        mHeader = 8
                    Else
                        Print #1, mSP2 & PRN_TIT(PubComp_Name, "A", PageWidth)
                        mHeader = mHeader + 1
                        If XNull(RstCompDet!W_SecSpeciality) <> "" Then
                            Print #1, mSP2 & PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
                            mHeader = mHeader + 1
                        End If
                        Print #1, mSP2 & PRN_TIT(PubComp_Add, "C", PageWidth)
                        mHeader = mHeader + 1
                        If PubComp_Add2 <> "" Then
                            Print #1, mSP2 & PRN_TIT(PubComp_Add2, "C", PageWidth)
                            mHeader = mHeader + 1
                        End If
                        If PubComp_City <> "" Then
                            Print #1, mSP2 & PRN_TIT(PubComp_City, "C", PageWidth)
                            mHeader = mHeader + 1
                        End If
                    End If
                    Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 27, , AlignRight, " ")
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
                    mHeader = mHeader + 1
                    'Service tax No Printing............
                    SrvTaxNo = XNull(GCn.Execute("Select SrvTaxNo from Syctrl").Fields(0).Value)
        
                    Print #1, mSP2 & PSTR("Serv.Tax No.  : " & SrvTaxNo, 40, , AlignLeft)
                    mHeader = mHeader + 1
                    '....................................
                    If mVatYn = 1 Then
                        Print #1, PRN_TIT("** " & txt(PrintTitle) & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                    Else
                        Print #1, PRN_TIT("**" & txt(PrintTitle) & "**", "B", PageWidth)
                        mHeader = mHeader + 1
                    End If
                Print #1, ""
                mHeader = mHeader + 1
                'Print #1, mSP2 & mChr18 & "TO," & Space(33) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " : " & Right(RstJob!DocId, 6) & mEmph1
                Print #1, mSP2 & mChr18 & "TO," & Space(39) & mEmph & PSTR(Lbl(2).CAPTION, 22, , AlignRight) & " : " & Right(txt(DocID), 6) & mEmph1
                mHeader = mHeader + 1
                Print #1, mSP2 & mEmph & PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(1) & PSTR("DATE", 12, , AlignRight) & "        : " & Format(txt(VDate), "dd/MMM/yyyy")
                mHeader = mHeader + 1
                
                Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(13) & PSTR("Job Card No.", 12) & ":" & Right(RstJob!job_docid, 6)
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40)
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR(XNull(RstJob!Address3) & IIf(XNull(RstJob!Address3) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44)
                mHeader = mHeader + 1
                Print #1, mSP2 & "Phone : " & PSTR(XNull(RstJob!PhoneNO), 20)
                mHeader = mHeader + 1
                Print #1, mSP2 & PSTR("Chass.No.", 11) & PSTR(XNull(RstJob!Chassis), 15) & Space(1) & Space(6) & PSTR(XNull(RstJob!Model), 13) & Space(1) & PSTR("Reg.", 5) & PSTR(XNull(RstJob!RegNo), 14) & " Kms:" & PSTR(XNull(RstJob!Kms), 6)
                mHeader = mHeader + 1
                Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
                mHeader = mHeader + 1
                If mVatYn = 1 Then
                    Print #1, mSP2 & PSTR("SrNo", 5) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 25) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP Rate", 11, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR(" DISC%", 6, , AlignRight) & PSTR("DISC. AMT", 12, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("Tax Amt", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18   '& mDoub1
                    mHeader = mHeader + 1
                Else
                    If RstJob!Det_Tax = 1 Then
                        Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 27) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(6) & PSTR("RATE", 11, , AlignRight) & Space(5) & "<---------AMOUNT--------- >"
                        mHeader = mHeader + 1
                        Print #1, mSP2 & Space(113) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 10, , AlignRight) & mChr18     '& mDoub1
                        mHeader = mHeader + 1
                    Else
                        Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 27) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(6) & PSTR("RATE", 11, , AlignRight) & Space(5) & "<---------AMOUNT--------- >"
                        mHeader = mHeader + 1
                        Print #1, mSP2 & Space(113) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 10, , AlignRight) & mChr18     '& mDoub1
                        mHeader = mHeader + 1
                    End If
                    
                End If
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
                    mFix = PageLength - (mHeader + mFooter)
                    mLine = 1
                End If
                If mLabCaption = False Then
                    If RstJob!orig = "1" Then
                        Print #1, mSP2 & mChr18 & mEmph & "*Spare Details*" & Replace(Space(PageWidth - 16), " ", "-") & mEmph1 & mChr17
                        mHeader = mHeader + 1
                        mLabCaption = True
                    End If
                End If
                mRate = IIf(RstJob!MRP_YN = 1, RstJob!MRP_Rate, RstJob!Rate)
                If RstJob!orig = "1" Then
                    If RstJob!Det_Tax = 1 Then
                        mTPAmtStr = PSTR(0, 12, 2)
                        mTBAmtStr = PSTR(0, 12, 2)
'                    If RstJob!Purpose = "W" Then
'                        mTBAmtStr = "*Warranty*"
'                    ElseIf RstJob!Purpose = "P" Then
'                        mTBAmtStr = "*PDI*"
'                    ElseIf RstJob!Purpose = "F" Then
'                        mTBAmtStr = "*Free*"
'                    ElseIf RstJob!Purpose = "L" Then
'                        mTBAmtStr = "*Compliment*"
'                    ElseIf RstJob!Purpose = "O" Then
'                        mTBAmtStr = "*Company*"
'                    Else
                    If RstJob!Tax_YN = 0 Then
                        mTPAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                        mTBAmtStr = PSTR(0, 12, 2)
                    Else
                        mTPAmtStr = PSTR(0, 12, 2)
                        mTBAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                    End If
                        
                    If mVatYn = 1 Then
                        PrintStr = PSTR(Trim(STR(mSlNo)), 5) & PSTR(IIf(RstJob!ReqNo = 0, "--", CStr(RstJob!ReqNo)), 7, , AlignLeft) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 25) & PSTR(RstJob!Qty, 12, 3)
                        PrintStr = PrintStr & PSTR(VNull(RstJob!MRP_Rate), 11, 2) & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                        PSTR(RstJob!Disc_Per, 5, 2) & "%" & PSTR(RstJob!Disc_Amt, 10, 2) & PSTR(VNull(RstJob!TaxPer), 6, 2) & PSTR(VNull(RstJob!TaxAmt), 10, 2) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                    Else
                        PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(IIf(RstJob!ReqNo = 0, "--", CStr(RstJob!ReqNo)), 7, , AlignLeft) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 27) & PSTR(RstJob!Qty, 12, 3)
                        PrintStr = PrintStr & IIf(RstJob!MRP_YN = 1, PSTR(mRate, 11, 2), PSTR("--", 11, 2, AlignRight)) & Space(6) & PSTR(mRate, 11, 2) & _
                        Space(8) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                    End If

                Else
                    LAmtItem = RstJob!Net_Amt2 + RstJob!Disc_Amt2
                    LDAmt = LAmtItem + (LAmtItem * (LAdd / IIf(SubTot = 0, 1, SubTot)))
                    LAmtVal = LAmtVal + (LAmtItem * (LAdd / IIf(SubTot = 0, 1, SubTot)))
                    LdRate = LDAmt / IIf(RstJob!Qty_Iss = 0, 1, RstJob!Qty_Iss)
                    If I = RstJob.RecordCount Then
                        If LAmtVal <> LAdd Then LDAmt = LDAmt + (LAdd - LAmtVal)
                        LdRate = LDAmt / IIf(RstJob!Qty_Iss = 0, 1, RstJob!Qty_Iss)
                    End If
                    mGrossAmt = mGrossAmt + (LDAmt - RstJob!Disc_Amt2)
                    I = I + 1
                    mAmount = Round(RstJob!Qty_Iss * RstJob!Rate, 2) - RstJob!Disc_Amt2
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 26, , AlignLeft) & PSTR(RstJob!Part_Name, 40) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(LdRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", "L") & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & _
                    PSTR(LDAmt - RstJob!Disc_Amt2, 12, 2)
                End If
            Else    'Labour
                If mServTaxAmt <= 0 Then
                    mTPAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                    mTBAmtStr = PSTR(0, 12, 2)
                Else
                    mTPAmtStr = PSTR(0, 12, 2)
                    mTBAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                End If
                If mVatYn = 1 Then
                    PrintStr = PSTR(Trim(STR(mSlNo)), 5) & Space(7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 25) & PSTR(Format(RstJob!Qty, "0.000"), 12, 3, AlignRight)
                    PrintStr = PrintStr & PSTR("--", 11, 2) & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                    PSTR(RstJob!Disc_Per, 5, 2) & "%" & PSTR(RstJob!Disc_Amt, 10, 2) & Space(6) & Space(10) & PSTR(IIf(Val(mTBAmtStr) > 0, mTBAmtStr, mTPAmtStr), 12, 2, AlignRight)
                Else
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 34)
                    PrintStr = PrintStr & Space(48) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                End If
            End If
            If PrintStr <> "" Then
                Print #1, mChr17 & mSP2 & PrintStr & mChr18
                mLine = mLine + 1
                mHeader = mHeader + 1
            End If
            mSlNo = mSlNo + 1
NXT:
            RstJob.MoveNext
            If mLine >= mFix Then
                If RstJob.EOF = True And (mTotRow > 15 And mTotRow <= 30) Then
                       RstJob.MovePrevious
                       Page = Page + 1
                       Do Until mTotRow >= 30
                             Print #1, ""
                            mTotRow = mTotRow + 1
                        Loop
                        Print #1, mChr18 & mSP2 & Replace(Space(PageWidth), " ", "-")
                        Print #1, mSP2 & Space((PageWidth) - Len("Contd. on next page.." + STR(Page))) & "Contd. on next page.." & STR(Page)
                        Print #1, mEject
                        'Header On Second Page
                        mHeader = 0
                        If Not FirstPrint Then
                            Print #1, ""
                            FirstPrint = True
                        End If
                        If PrePrinted Then
                            Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                            Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                            mHeader = 8
                        Else
                            Print #1, mSP2 & PRN_TIT(PubComp_Name, "A", PageWidth)
                            mHeader = mHeader + 1
                            If XNull(RstCompDet!W_SecSpeciality) <> "" Then
                                Print #1, mSP2 & PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
                                mHeader = mHeader + 1
                            End If
                            Print #1, mSP2 & PRN_TIT(PubComp_Add, "C", PageWidth)
                            mHeader = mHeader + 1
                            If PubComp_Add2 <> "" Then
                                Print #1, mSP2 & PRN_TIT(PubComp_Add2, "C", PageWidth)
                                mHeader = mHeader + 1
                            End If
                            If PubComp_City <> "" Then
                                Print #1, mSP2 & PRN_TIT(PubComp_City, "C", PageWidth)
                                mHeader = mHeader + 1
                            End If
                        End If
                        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 27, , AlignRight, " ")
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
                        mHeader = mHeader + 1
                        'Service tax No Printing............
                        SrvTaxNo = XNull(GCn.Execute("Select SrvTaxNo from Syctrl").Fields(0).Value)
            
                        Print #1, mSP2 & PSTR("Serv.Tax No.  : " & SrvTaxNo, 40, , AlignLeft)
                        mHeader = mHeader + 1
                        '....................................
                        If mVatYn = 1 Then
                            Print #1, PRN_TIT("** " & txt(PrintTitle) & " **", "B", PageWidth)
                            mHeader = mHeader + 1
                        Else
                            Print #1, PRN_TIT("**" & txt(PrintTitle) & "**", "B", PageWidth)
                            mHeader = mHeader + 1
                        End If
                        Print #1, mSP2 & mChr18 & Space(46) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocID) & mEmph1
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!job_docid)
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(13) & PSTR("Job Card No.", 12) & ":" & Right(RstJob!job_docid, 6)
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40)
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR(XNull(RstJob!Address3) & IIf(XNull(RstJob!Address3) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44)
                        mHeader = mHeader + 1
                        Print #1, mSP2 & "Phone : " & PSTR(XNull(RstJob!PhoneNO), 20)
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR("Chass.No.", 11) & PSTR(XNull(RstJob!Chassis), 15) & Space(1) & Space(6) & PSTR(XNull(RstJob!Model), 13) & Space(1) & PSTR("Reg.", 5) & PSTR(XNull(RstJob!RegNo), 14) & " Kms:" & PSTR(XNull(RstJob!Kms), 6)
                        mHeader = mHeader + 1
                        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                        mHeader = mHeader + 1
                        Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
                        mHeader = mHeader + 1
                        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17
                        mHeader = mHeader + 1
                        mFix = PageLength - (mHeader + mFooter)
                        mLine = 1
                        Do Until mLine >= 15
                            Print #1, ""
                            mLine = mLine + 1
                        Loop
                        RstJob.MoveNext
                End If
            End If
'            mSlNo = mSlNo + 1
'            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop

    Print #1, mChr18 & mSP2 & "Customer's Signature"
    


    
' SALE FOOTER
    '22 space maintain between heading and :
    RstJob.MoveFirst
    'If mTotRow <= 15 Then
    'If RstJob!Det_Tax = 1 Then
        If mVatYn = 1 Then
            Print #1, mSP2 & Replace(Space(35), " ", "-") & "Taxable Amt" & Replace(Space(33), " ", "-")

            Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & Space(19) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("V A T     ", 10, 0) & Space(6) & PSTR(Val(txt(STaxAmt)), 12, 2) & mDoub
            
            Print #1, mSP2 & Space(47) _
            ; " | " & PSTR("S A T     ", 10, 0) & Space(6) & PSTR(Val(txt(SatAmt)), 12, 2) & mDoub
            
            Print #1, mSP2 & PSTR("MRP Items Amt", 16) & Space(19) & PSTR(Val(txt(MRPAmtTB)), 12, 2) & mDoub1 _
            ; " | " & PSTR("Misc. Charges", 16) & PSTR(Val(txt(PackCrg)), 12, 2) & mDoub
            
            Print #1, mSP2 & PSTR("Spares Amount", 16) & Space(19) & PSTR(Val(txt(SprAmtTB)), 12, 2) & mDoub1 _
            ; " | " & mEmph & PSTR("Sub Total", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
    
            Print #1, mSP2 & PSTR("Oil Amount ", 16) & Space(19) & PSTR(Val(txt(OilAmtTB)), 12, 2) & mDoub1 _
            ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(Val(txt(TurnOverPer)), 5, 2) & "%" & PSTR(Val(txt(TurnOverAmt)), 12, 2)
            
            Print #1, mSP2 & PSTR("Discount ", 16, 0) & Space(11) & PSTR(Val(txt(DiscPerTB)), 7, 2) & "%" & PSTR(Val(txt(DiscAmtTB)), 12, 2) _
            ; " | " & PSTR("Round Off", 16) & PSTR(Round(Val(txt(SROff)), 2), 12, 2)
            
            Print #1, mSP2 & PSTR("Sub Total [A]", 16) & Space(19) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
            ; " | " & mEmph & PSTR("Net Spare + Lub. Rs.", 16) & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
        Else
        
            Print #1, mSP2 & Replace(Space(20), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
            
            Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(Val(txt(STaxPer)), 5, 2) & "%" & PSTR(Val(txt(STaxAmt)), 12, 2) & mDoub
            
            Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(Val(txt(MRPAmtTP)), 11, 2) & Space(8) & PSTR(Val(txt(MRPAmtTB)), 12, 2); " | " & mDoub1 _
            & PSTR("Misc. Charges", 16) & PSTR(Val(txt(PackCrg)), 12, 2) & mDoub

                
            Print #1, mSP2 & PSTR("Spares Amount", 16) & PSTR(Val(txt(SprAmtTP)), 11, 2) & Space(8) & PSTR(Val(txt(SprAmtTB)), 12, 2) & mDoub1 _
            ; " | " & mEmph & PSTR("Sub Total[TP+TB]", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
    
            Print #1, mSP2 & PSTR("Oil Amount ", 16) & PSTR(Val(txt(OilAmtTP)), 11, 2) & Space(8) & PSTR(Val(txt(OilAmtTB)), 12, 2) & mDoub1 _
            ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(Val(txt(TurnOverPer)), 5, 2) & "%" & PSTR(Val(txt(TurnOverAmt)), 12, 2)
            
            Print #1, mSP2 & PSTR("Discount ", 10, 0) & PSTR(Val(txt(DiscPerTP)), 5, 2) & "%" & PSTR(Val(txt(DiscAmtTP)), 11, 2) & PSTR(Val(txt(DiscPerTB)), 7, 2) & "%" & PSTR(Val(txt(DiscAmtTB)), 12, 2) _
            ; " | " & PSTR("Round Off", 16) & PSTR(Round(Val(txt(SROff)), 2), 12, 2)
            
            Print #1, mSP2 & PSTR("Sub Total [A]", 16) & PSTR(Val(txt(STotATP)), 11, 2) & Space(8) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
            ; " | " & mEmph & PSTR("Net Spare + Lub. Rs.", 16) & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
        End If
    

    
    If Val(txt(LabDisc)) > 0 Then
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
        mLabDiscAmtStr = "Discount  : " & PSTR(Val(txt(LabDisc)), 8, 2)
    Else
        mLabDiscAmtStr = Space(19)
    End If
    
    
    PrintStr = mLabDiscAmtStr
    'PrintStr = PrintStr & " |" & "Round Off      :  " & PSTR(Val(txt(LabROff)), 7, 2) & " |" & "Net Payble Amt Rs.: " & PSTR(Val(txt(NetAmt)), 9, 2) & mEmph1
    Print #1, mSP2 & PrintStr
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    


        If mLabDiscAmt > 0 Then
            mLabDiscAmtStr = "Discount  : " & PSTR(mLabDiscAmt, 8, 2)
        Else
            mLabDiscAmtStr = Space(19)
        End If

        PrintStr = mEmph & "Total Lab.: " & PSTR(mLabAmt, 8, 2)
        PrintStr = PrintStr & " |Serv.Tax @ " & PSTR(mServTaxPer, 4, 2) & ":" & PSTR(mServTaxAmt, 9, 2) & "|" & "Net Labour Rs.    : " & PSTR(mNetLabAmt, 9, 2)
        Print #1, mSP2 & PrintStr
        PrintStr = mLabDiscAmtStr
        mNetTotal = Val(txt(NetSprAmt)) + mNetLabAmt
        'PrintStr = PrintStr & "  |" & "Round Off      :  " & PSTR(mLabROffAmt, 7, 2) & " |" & "Net Payble Amt Rs.: " & PSTR(mNetLabAmt, 9, 2) & mEmph1
        PrintStr = PrintStr & "  |" & "Round Off      :  " & PSTR(mLabROffAmt, 7, 2) & " |" & "Net Payble Amt Rs.: " & PSTR(mNetTotal, 9, 2) & mEmph1

        Print #1, mSP2 & PrintStr
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    Print #1, mSP2 & mDoub & ntow(mNetTotal, "Rupees", "Paise") & mDoub1
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & "The service tax  amount claimed on this invoice will be paid to govt. as per statutory provision" & mChr18
    If mVatYn = 1 Then
        Print #1, ""
    Else
        'If UCase(left(PubComp_Name, 3)) <> "JMK" Then
            Print #1, mSP2 & mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
        'End If
    End If
    Print #1, mSP2 & PSTR(mTaxdesc, 25) & Space(PageWidth - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, ""
    Print #1, mSP2 & mDoub & "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
       If mID(Footer, I, 1) = vbLf Then
           Print #1, mSP2 & RTrim(mID(Footer, j, I - j))
           j = I + 1
       End If
    Next
    Print #1, mSP2 & Space((((PageWidth) * 1.7) - Len("* a dataman software *" & "   " & pubUName & "   " & PubServerDate)) / 2) & "* a dataman software *" & "   " & pubUName & "   " & PubServerDate & mChr18
'Gate Pass Footer()
    'If RstJob!Printed_YN = 0 Then
'        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "S" Then
'            Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
'            Print #1, mSP2 & PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", (PageWidth)) & mEmph
'            'Print #1, mSP2 & "GATE PASS No. & DATE : " & XNull(RstJob!GP_No) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(1) & "Job Card No.: " & Right(RstJob!Job_DocID, 6)
'            Print #1, mSP2 & "Vehicle No. : " & XNull(RstJob!RegNo) & Space(5) & "Chassis No. : " & XNull(RstJob!Chassis) _
'            & Space(5) & mChr17 & "Model : " & XNull(RstJob!Model) & mChr18
'            Print #1, ""
'            Print #1, mSP2 & "Vehicle has been received from workshop & work done as per  my satisfaction."
'            Print #1, ""
'            Print #1, mSP2 & "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
'            Print #1, mSP2 & mChr17 & Space((((PageWidth) * 1.7) - Len("* a dataman software *" & "   " & pubUName & "   " & PubServerDate)) / 2) & "* a dataman software *" & "   " & pubUName & "   " & PubServerDate & mChr18
'        End If
    'End If
    'End If
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    FirstPrint = IIf(FirstPrint, FirstPrint, True)
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
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        GCn.Execute "update Estimate set Printed_YN = 1  where Estimate.docid='" & Master!SearchCode & "' "
    Else
        If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
            GCn.Execute "update Estimate set Printed_YN = 1  where Estimate.docid='" & Master!SearchCode & "' "
        End If
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
    
    mVatYn = PubVATYN
End Sub

