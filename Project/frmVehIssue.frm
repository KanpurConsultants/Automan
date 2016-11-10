VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehIssue 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Issue Entry"
   ClientHeight    =   7230
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
   ScaleHeight     =   7230
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
      Height          =   1605
      Left            =   1050
      TabIndex        =   19
      Top             =   5160
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
         Picture         =   "frmVehIssue.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "frmVehIssue.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Print"
         DisabledPicture =   "frmVehIssue.frx":0678
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
         TabIndex        =   27
         ToolTipText     =   "Printer "
         Top             =   825
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmVehIssue.frx":0982
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
         TabIndex        =   26
         ToolTipText     =   "Screen"
         Top             =   495
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmVehIssue.frx":0C8C
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
         TabIndex        =   25
         ToolTipText     =   "Printer "
         Top             =   285
         Visible         =   0   'False
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         Left            =   -75
         TabIndex        =   32
         Top             =   315
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   6705
      Left            =   6120
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   11827
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "CODE"
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
      BeginProperty Column01 
         DataField       =   "EngineNo"
         Caption         =   "Engine No."
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
            ColumnWidth     =   2250.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2160
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2910
      Left            =   60
      Negotiate       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4140
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5133
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
      Caption         =   "Party Help"
      ColumnCount     =   1
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4424.882
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
      Index           =   3
      Left            =   2250
      MaxLength       =   30
      TabIndex        =   1
      Top             =   435
      Width           =   3525
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
      Left            =   8385
      MaxLength       =   8
      TabIndex        =   2
      Top             =   450
      Width           =   1680
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
      Left            =   5760
      TabIndex        =   5
      Top             =   4530
      Visible         =   0   'False
      Width           =   690
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
      Left            =   10815
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   675
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
      Index           =   0
      Left            =   8385
      MaxLength       =   12
      TabIndex        =   3
      Top             =   720
      Width           =   1665
   End
   Begin MSDataGridLib.DataGrid DGCol 
      Height          =   2130
      Left            =   2250
      Negotiate       =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   3757
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
      Caption         =   "Colour Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Colors"
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
            ColumnWidth     =   4380.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGod 
      Height          =   2130
      Left            =   6180
      Negotiate       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3757
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2745
      Left            =   -360
      TabIndex        =   6
      Top             =   1335
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   4842
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   14
      BackColorFixed  =   12632319
      ForeColorFixed  =   16384
      BackColorSel    =   15196124
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   12632319
      GridColorFixed  =   32896
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "SrNo.|Model | Taxable |Colour | Chassis No |Engine No |SDM/STM |Service Book No |Godown|Received from |Remark|Colcode|partycode"
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
      _Band(0).Cols   =   14
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
      Index           =   0
      Left            =   2010
      TabIndex        =   18
      Top             =   450
      Width           =   45
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
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   0
      Left            =   915
      TabIndex        =   17
      Top             =   420
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
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
      Index           =   1
      Left            =   7170
      TabIndex        =   13
      Top             =   450
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   92
      Left            =   8145
      TabIndex        =   12
      Top             =   450
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
      Index           =   8
      Left            =   10575
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RSO Purchase Y/N"
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
      Index           =   27
      Left            =   8850
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   91
      Left            =   8130
      TabIndex        =   8
      Top             =   720
      Width           =   45
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   2
      Left            =   7530
      TabIndex        =   7
      Top             =   705
      Width           =   390
   End
End
Attribute VB_Name = "frmVehIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGod As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim RsMod  As ADODB.Recordset
Dim RsCol  As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsChassis As ADODB.Recordset
Dim GridKey As Integer

Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const VDate As Byte = 0
Private Const SerialNo As Byte = 1
Private Const RSO_WORK As Byte = 2
Private Const PName As Byte = 3
' Col Declaration

'Sal_VNO,RSO_WORK,MODEL ,TAX_YN,Colour_Code,ChassisNo, EngineNo,SDM_STM_NO,Srv_BookNo,Godown,PartyCode,Remarks

Private Const ChassisNo As Byte = 1
Private Const EngineNo As Byte = 2
Private Const Model As Byte = 3
Private Const Colours As Byte = 4
Private Const Taxable As Byte = 5
Private Const InDate As Byte = 6
Private Const SDM_STM_NO As Byte = 7
Private Const Srv_BookNo  As Byte = 8
Private Const Godown As Byte = 9
Private Const PartyName  As Byte = 10
Private Const Remarks  As Byte = 11
Private Const ColCode  As Byte = 12
Private Const PartyCode  As Byte = 13
Private Const God As Byte = 14
Private Const ChassisNoOld As Byte = 15
Private Const mVType As String = "V_TRF"
Private Const mVPerfix As String = "V_TRF"

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "VehIssue", "VehIssue")
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
        
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
'If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
'    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
'    Disp_Text SETS("INI", Me, Master)
'    Call MoveRec
'End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub DgChassis_Click()
    DgChassis.Visible = False
    If RsChassis.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsChassis!Code
    End If
    TxtGrid(0).SetFocus
End Sub
Private Sub DGCol_Click()
    DGCol.Visible = False
    If RsCol.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsCol!Name
         FGrid.TextMatrix(FGrid.Row, Colours) = RsCol!Name
         FGrid.TextMatrix(FGrid.Row, ColCode) = RsCol!Code
    End If
   TxtGrid(0).SetFocus
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
'Dim i As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
'    Master.Open "SELECT DISTINCT veh_stock.Sal_VNO AS searchcode, veh_stock.Sal_VNO, veh_stock.Sal_VDate, veh_stock.RSO_WORK FROM veh_stock where Chassis_RctSiteCode  = '" & PubSiteCode & "' and left(Sal_DocId,1)  = '" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic

   Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(veh_stock.Sal_Site_Code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If


    If PubMoveRecYn Then
        Master.Open "SELECT DISTINCT veh_stock.Sal_VNO AS searchcode, veh_stock.Sal_VNO, veh_stock.Sal_VDate,veh_stock.RSO_WORK ,Veh_Stock.TrfParty,SG.Name as PName FROM veh_stock Left Join SubGroup SG on Veh_Stock.TrfParty=SG.SubCode where left(Sal_DocId,1)  = '" & PubDivCode & "' and " & cMID("Sal_DocId", "4", "5") & "  = '" & mVType & "' " & sitecond & " ", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "SELECT DISTINCT Top 1 veh_stock.Sal_VNO AS searchcode, veh_stock.Sal_VNO, veh_stock.Sal_VDate,veh_stock.RSO_WORK ,Veh_Stock.TrfParty,SG.Name as PName FROM veh_stock Left Join SubGroup SG on Veh_Stock.TrfParty=SG.SubCode where left(Sal_DocId,1)  = '" & PubDivCode & "' and " & cMID("Sal_DocId", "4", "5") & "  = '" & mVType & "' " & sitecond & " ", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    Set RsCol = New ADODB.Recordset
    RsCol.CursorLocation = adUseClient
    RsCol.Open "select Col_code as code,col_Desc  as name from colmast order by col_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGCol.DataSource = RsCol
    
    Set rsGod = New ADODB.Recordset
    rsGod.CursorLocation = adUseClient
    rsGod.Open "select god_code as code,god_name as name from godown where appli_for = 1 order by god_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGod.DataSource = rsGod
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME from SubGroup Where firmCode = '" & PubFirmCode & "' and Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME from SubGroup " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        " order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
   
    Set RsChassis = New ADODB.Recordset
    RsChassis.CursorLocation = adUseClient
    RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code, Veh_Stock.EngineNo,Veh_Stock.Chassis_RctDocNo, Veh_Stock.INDATE, Veh_Stock.Srv_BookNo, Veh_Stock.Mfg_Month, Veh_Stock.Mfg_Yr, Veh_Stock.SDM_STM_NO, Veh_Stock.TAX_YN, Godown.God_Name, ColMast.Col_Desc, Veh_Stock.Colour_Code, Veh_Stock.Godown ,Veh_Stock.Model" & _
        " FROM (Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code) " & _
        "LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
        "where (len(Veh_Stock.Sal_DocId) = 0 Or Veh_stock.Sal_DocId Is Null)"), GCn, adOpenDynamic, adLockOptimistic
    Set DgChassis.DataSource = RsChassis
    
    SetDGHelp DGParty, Txt, PName, Me.height, leftSide
    SetDGHelp DgChassis, TxtGrid, 0, Me.height, leftSide
    
   
    Call MoveRec
    
    Ini_Grid
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
Set rsGod = Nothing
Set RsParty = Nothing
Set RsMod = Nothing
Set RsCol = Nothing
Set Master = Nothing
Set RsChassis = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Txt(SerialNo).TEXT = VNull(GCn.Execute("select MAX(Sal_VNo) from veh_stock where Sal_Site_Code  = '" & PubSiteCode & "' and left(Sal_DocId,1)  = '" & PubDivCode & "' and " & cMID("Sal_DocId", "4", "5") & "= '" & mVType & "'").Fields(0).Value) + 1
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    Txt(PName).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer, mDocId As String

If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    GCn.BeginTrans
        mDocId = PubDivCode & PubSiteCode & PubSiteCode & mVType & mVPerfix & Txt(SerialNo)
        GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_VDate = null,Sal_Site_Code ='',Sal_VType='',Sal_VNo=0,TrfParty='',Remarks=''  where Sal_DocId  = '" & mDocId & "'"
        GCn.CommitTrans
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
 Dim I As Integer
    Disp_Text SETS("EDIT", Me, Master)
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
End Sub

Private Sub TopCtrl1_eRef()
    rsGod.Requery
    RsParty.Requery
    RsCol.Requery
    RsMod.Requery
    RsChassis.Requery
End Sub
Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim CardNo As String
    On Error GoTo errlbl
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    If IsValid(Txt(VDate), "Issue Date") = False Then Exit Sub
    If IsValid(Txt(SerialNo), "Serial Number") = False Then Exit Sub
    If IsValid(Txt(PName), "Party Name") = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ChassisNo) <> "" Then
            If FGrid.TextMatrix(I, ChassisNo) = "" Then MsgBox "Fill Chassis No  in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = ChassisNo: FGrid.SetFocus: Exit Sub
        End If
    Next
    GCn.BeginTrans
    mTrans = True
    Dim mDocId As String
    mDocId = PubDivCode & PubSiteCode & PubSiteCode & mVType & mVPerfix & Txt(SerialNo)

    If TopCtrl1.TopText2.CAPTION = "Add" Then
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ChassisNo) <> "" Then
                GCn.Execute "Update Veh_Stock set Sal_DocId = '" & mDocId & "',Sal_VDate=" & ConvertDate(Txt(VDate)) & ",Sal_Site_Code ='" & PubSiteCode & "',Sal_VType='" & mVType & "',Sal_VNo='" & Val(Txt(SerialNo)) & "',TrfParty='" & Txt(PName).Tag & "',Remarks='" & FGrid.TextMatrix(I, Remarks) & "'  where ChassisNo  = '" & FGrid.TextMatrix(I, ChassisNo) & "' AND (Sal_DocId  = '' or Sal_DocId is Null)"
            End If
        Next
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
        GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_VDate = null,Sal_Site_Code ='',Sal_VType='',Sal_VNo=0,TrfParty='',Remarks=''  where Sal_DocId  = '" & mDocId & "'"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ChassisNo) <> "" Then
                GCn.Execute "Update Veh_Stock set Sal_DocId = '" & mDocId & "',Sal_VDate=" & ConvertDate(Txt(VDate)) & ",Sal_Site_Code ='" & PubSiteCode & "',Sal_VType='" & mVType & "',Sal_VNo='" & Val(Txt(SerialNo)) & "',TrfParty='" & Txt(PName).Tag & "',Remarks='" & FGrid.TextMatrix(I, Remarks) & "'  where ChassisNo  = '" & FGrid.TextMatrix(I, ChassisNo) & "'"
            End If
        Next
    End If
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("SELECT DISTINCT veh_stock.Sal_VNO AS searchcode, veh_stock.Sal_VNO, veh_stock.Sal_VDate,veh_stock.RSO_WORK ,Veh_Stock.TrfParty,SG.Name as PName FROM veh_stock Left Join SubGroup SG on Veh_Stock.TrfParty=SG.SubCode where left(Sal_DocId,1)  = '" & PubDivCode & "' and " & cMID("Sal_DocId", "4", "5") & "  = '" & mVType & "' And veh_stock.Sal_VNO = " & Val(Txt(SerialNo)) & " ")
    End If
    RsParty.Requery
    Master.FIND "Sal_VNO = " & Val(Txt(SerialNo)) & ""
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(veh_stock.Sal_Site_Code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    GSQL = "SELECT DISTINCT Sal_VNO AS SearchCode, Sal_VNO as SerialNo, Sal_VDate as V_Date, ChassisNo as Chassis_No,EngineNo as Engine_No FROM veh_stock where left(Sal_DocId,1)  = '" & PubDivCode & "' and " & cMID("Sal_DocId", "4", "5") & "  = '" & mVType & "' " & sitecond & " order by ChassisNo"
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
        Master.FIND "searchcode=" & MyValue
    Else
        Set Master = GCn.Execute("SELECT DISTINCT veh_stock.Sal_VNO AS searchcode, veh_stock.Sal_VNO, veh_stock.Sal_VDate,veh_stock.RSO_WORK ,Veh_Stock.TrfParty,SG.Name as PName FROM veh_stock Left Join SubGroup SG on Veh_Stock.TrfParty=SG.SubCode where left(Sal_DocId,1)  = '" & PubDivCode & "' and " & cMID("Sal_DocId", "4", "5") & "  = '" & mVType & "' And veh_stock.Sal_VNO = " & MyValue & " ")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Select Case Index
    Case PName
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(PName) = "" Then Exit Sub
            If Txt(PName) <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.FIND "name ='" & Txt(PName) & "'"
            End If
            
End Select
TxtGrid(0).Visible = False
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
Select Case Index
     Case PName
            DGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, True, 1, frmSubGroup, "frmSubGroup"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave(PName) = True Then
                   DGridTxtKeyDown_Mast DGParty, Txt, Index, RsParty, KeyCode, False
                End If
            End If
End Select
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
If DGParty.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> VDate Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> RSO_WORK Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case Index
    Case SerialNo
        Call NumPress(Txt(Index), KeyAscii, 6, 0)
    Case RSO_WORK
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
    Case PName
        If DGParty.Visible = True Then DGridTxtKeyPress Txt, Index, RsParty, KeyAscii, "Name"

End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Select Case Index
    Case VDate
        If Len(Trim(Txt(VDate).TEXT)) = 0 Then
             Txt(VDate).TEXT = PubLoginDate
        Else
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
End Select
Set Rst = Nothing
End Sub

Private Sub DGGod_Click()
    DGGod.Visible = False
    If rsGod.RecordCount > 0 Then
        TxtGrid(0).TEXT = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
    End If
   TxtGrid(0).SetFocus
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        Txt(PName).TEXT = RsParty!Name
        Txt(PName).Tag = RsParty!Code
    End If
    DGParty.Visible = False
    Txt(PName).SetFocus
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case ChassisNo, Remarks
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case ChassisNo, Remarks
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case ChassisNo, Remarks
        Call GridDblClick(Me, FGrid, TxtGrid, 0)
End Select
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                FGrid.RemoveItem (FGrid.Row)
        End If
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
   
FGrid.SetFocus
End If
Exit Sub
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
Select Case FGrid.Col
    Case ChassisNo, Remarks
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
Next I
End Sub

Private Sub MoveRec()
Dim Rs As ADODB.Recordset, I As Integer
On Error GoTo error1
'TopCtrl1.tPrn = False
If Master.RecordCount > 0 Then
    Txt(SerialNo).TEXT = XNull(Master!Sal_VNO)
    Txt(VDate).TEXT = Master!Sal_VDate
    Txt(PName) = XNull(Master!PName)
    Txt(PName).Tag = XNull(Master!TrfParty)
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Veh_Stock.*, ColMast.Col_Desc, SubGroup.Name AS party, Godown.God_Name " & _
    "FROM ((Veh_Stock LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN SubGroup ON Veh_Stock.PartyCode = SubGroup.Subcode) LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code " & _
    "where left(Veh_Stock.Sal_DocId,1)  = '" & PubDivCode & "' and Veh_Stock.Sal_VNO = " & Master!Sal_VNO & "")
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        I = 1
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = I
                .TextMatrix(I, Model) = Rs!Model
                .TextMatrix(I, Taxable) = IIf(Rs!Tax_YN = 0, "No", "Yes")
                .TextMatrix(I, Colours) = XNull(Rs!Col_Desc)
                .TextMatrix(I, ChassisNo) = Rs!ChassisNo
                .TextMatrix(I, EngineNo) = Rs!EngineNo
                .TextMatrix(I, InDate) = XNull(Rs!InDate)
                .TextMatrix(I, SDM_STM_NO) = XNull(Rs!SDM_STM_NO)
                .TextMatrix(I, Srv_BookNo) = Rs!Srv_BookNo
                .TextMatrix(I, Godown) = XNull(Rs!God_Name)
                .TextMatrix(I, Remarks) = Rs!Remarks
                .TextMatrix(I, ColCode) = Rs!Colour_Code
                .TextMatrix(I, PartyCode) = Rs!PartyCode
                .TextMatrix(I, God) = Rs!Godown
                .TextMatrix(I, ChassisNoOld) = Rs!ChassisNo
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
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
'Dim i As Byte
'SrNo.0|Model1|RSO/Work2|Tax3|Quantiy4|Rate5|Tax%6|TaxAmt7|Surch%8|SurchAmt9|Amount10

'Model 1| Taxable 2|Colour 3| Chassis No 4|Engine No 5|SDM/STM 6|Service Book No 7|Chassis Godown 8|Received from 8 |Remark 10
    With FGrid
        .Cols = 16
        .left = Me.left '+45
        .width = Me.width - 90
        .top = 1590
        .RowHeightMin = PubGridRowHeight
        
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, Model) = "Model"
        .ColAlignment(Model) = flexAlignLeftCenter
        .ColWidth(Model) = 1500

        .TextMatrix(0, Taxable) = "Tax"
        .ColAlignment(Taxable) = flexAlignLeftCenter
        .ColWidth(Taxable) = 360

        .TextMatrix(0, Colours) = "Colours"
        .ColAlignment(Colours) = flexAlignLeftCenter
        .ColWidth(Colours) = 1200

        .TextMatrix(0, ChassisNo) = "Chassis No"
        .ColAlignment(ChassisNo) = flexAlignLeftCenter
        .ColWidth(ChassisNo) = 1650

        .TextMatrix(0, EngineNo) = "Engine No"
        .ColAlignment(EngineNo) = flexAlignLeftCenter
        .ColWidth(EngineNo) = 1650
        
        .TextMatrix(0, InDate) = "InDate"
        .ColAlignment(InDate) = flexAlignLeftCenter
        .ColWidth(InDate) = 1080

        .TextMatrix(0, SDM_STM_NO) = "SDM/STM No"
        .ColAlignment(SDM_STM_NO) = flexAlignLeftCenter
        .ColWidth(SDM_STM_NO) = 0
        .TextMatrix(0, Srv_BookNo) = "SrvBookNo"
        .ColAlignment(Srv_BookNo) = flexAlignLeftCenter
        .ColWidth(Srv_BookNo) = 0

        .TextMatrix(0, Godown) = "Godown"
        .ColAlignment(Godown) = flexAlignLeftCenter
        .ColWidth(Godown) = 0

        .TextMatrix(0, PartyName) = "Party"
        .ColAlignment(PartyName) = flexAlignLeftCenter
        .ColWidth(PartyName) = 0

        .TextMatrix(0, Remarks) = "Remarks"
        .ColAlignment(Remarks) = flexAlignLeftCenter
        .ColWidth(Remarks) = 3500
        
        .ColWidth(ColCode) = 0
        .ColWidth(PartyCode) = 0
        .ColWidth(God) = 0
        .ColWidth(ChassisNoOld) = 0
End With
BackColorSelLeave = FGrid.BackColorSel
ForeColorSelEnter = FGrid.ForeColorSel
'DGParty.left = Me.left + 45: DGParty.top = FGrid.top + FGrid.height + 50
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    Txt(I).ForeColor = CtrlFColOrg
Next
If TopCtrl1.TopText2 = "Edit" Then
    Txt(VDate).Enabled = False
    Txt(SerialNo).Enabled = False
End If
txtDisabled_Color Me
TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
    If DgChassis.Visible = True Then DgChassis.Visible = False
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
       
       
       Case ChassisNo
            TxtGrid(0).MaxLength = 20
       Case EngineNo
            TxtGrid(0).MaxLength = 25
       Case Remarks
            TxtGrid(0).MaxLength = 40
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TxtGrid(0) = TxtGrid(0).Tag: Exit Sub
            Select Case FGrid.Col
               
                Case Remarks
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 10
                        End If
                    End If
                Case ChassisNo
                    If DgChassis.Visible = False Then SetDGHelp DgChassis, TxtGrid, 0, Me.height, leftSide
                    DGridTxtKeyDown_Mast DgChassis, TxtGrid, Index, RsChassis, KeyCode, True, 0
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                            GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 10
                        End If
                    End If
            End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case ChassisNo
        If KeyCode <> 13 And DgChassis.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
        DGridTxtKeyUp_Mast TxtGrid, Index, RsChassis, KeyCode, "code"
    
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
Dim j As Integer
Select Case Index
    Case PName
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index) = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index) = RsParty!Name
               Txt(Index).Tag = RsParty!Code
            End If
            Exit Function
            
End Select
Select Case FGrid.Col
        Case ChassisNo
            If ChkDul_Chassis = True Then TxtGridLeave = False: Exit Function
'**********modi end
            'MODISHEKHAR
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
                If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                  MsgBox "Chassis Sold" & vbCrLf & "Editing Denied", vbInformation, "Editing Denied": FGrid.SetFocus: TxtGridLeave = False: Exit Function
                End If
            End If
            'END MODI
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = UCase(TxtGrid(0).TEXT)
            FGrid.TextMatrix(FGrid.Row, EngineNo) = XNull(RsChassis!EngineNo)
            FGrid.TextMatrix(FGrid.Row, Model) = XNull(RsChassis!Model)
            FGrid.TextMatrix(FGrid.Row, Taxable) = IIf(VNull(RsChassis!Tax_YN) = 1, "Yes", "No")
            FGrid.TextMatrix(FGrid.Row, Color) = XNull(RsChassis!Col_Desc)
            FGrid.TextMatrix(FGrid.Row, InDate) = XNull(RsChassis!InDate)
            'FGrid.Col = PartyName
            FGrid.SetFocus
        
        Case Remarks
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function

Private Function ChkDuplicate() As Boolean
Dim I As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte
    Select Case FGrid.Col
    Case Model
        Col2 = Model
        Col1 = ChassisNo
    Case ChassisNo
        Col1 = Model
        Col2 = ChassisNo
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
Private Function ChkDul_Chassis() As Boolean
Dim I As Integer
If TxtGrid(0).TEXT = FGrid.TextMatrix(FGrid.Row, ChassisNo) Then
    ChkDul_Chassis = False
    Exit Function
End If
For I = 1 To FGrid.Rows - 1
    If I <> FGrid.Row Then
        If FGrid.TextMatrix(I, ChassisNo) = TxtGrid(0).TEXT Then
            MsgBox "Same Chassis No already taken ", vbInformation, "Duplicate Chassis"
            ChkDul_Chassis = True
            Exit Function
        End If
    End If
Next
ChkDul_Chassis = False
End Function

Private Sub WindowsPrint(Index As Integer)
Dim Rst As ADODB.Recordset, mQry As String
Dim RstSub1 As ADODB.Recordset
Dim I As Integer
Dim Rst2 As ADODB.Recordset
Dim mCurrBal, mCrLimit As Double
On Error GoTo ERRORHANDLER
    
    mQry = " SELECT Veh_Stock.*, ColMast.Col_Desc, SubGroup.Name AS party, Godown.God_Name,Subgroup.Add1,SubGroup.Add2,SubGroup.Phone " & _
            " FROM ((Veh_Stock LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code) " & _
            " LEFT JOIN SubGroup ON Veh_Stock.TrfParty = SubGroup.Subcode) " & _
            " LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code " & _
            " where left(Veh_Stock.Sal_DocId,1)  = '" & PubDivCode & "' and Veh_Stock.Sal_VNO = " & Master!Sal_VNO & ""

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
        
    
        
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
               
    'rpt.Database.SetDataSource Rst
    

    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
            
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("SubTitle")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecSpeciality & "'"
                Case UCase("LST")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST & "'"
                Case UCase("LSTDate")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST_Date & "'"
                Case UCase("CST")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST & "'"
                Case UCase("CSTDate")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST_Date & "'"
                Case UCase("Phone")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecPhone & "'"
                Case UCase("Fax")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecFax & "'"
                Case UCase("Gram")
                    rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecGram & "'"
                Case UCase("TitleType")
                        rpt.FormulaFields(I).TEXT = "'Vehicle Transfer Note'"
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
                    Case UCase("TitleType")
                        rpt.FormulaFields(I).TEXT = "'Vehicle Transfer Note'"
                End Select
                Next
                rpt.PrintOut False
        Case PScreen  'screen
                Call Report_View(rpt, Me.CAPTION, 0, True)
End Select
CmdPrint(PSetUp).Tag = ""
Set Rst = Nothing
Set Rst2 = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

