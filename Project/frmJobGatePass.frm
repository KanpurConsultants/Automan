VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmJobGatePass 
   BackColor       =   &H00CFE0E0&
   Caption         =   "GatePass Entry"
   ClientHeight    =   7530
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
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   11820
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   1650
      Left            =   780
      Negotiate       =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5490
      Visible         =   0   'False
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   2910
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
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Job_No"
         Caption         =   "Job No."
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
         DataField       =   "RegNo"
         Caption         =   "Reg. No"
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
      BeginProperty Column04 
         DataField       =   "VehSerialNo"
         Caption         =   "Veh.Srl No."
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
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
      Height          =   1095
      Index           =   33
      Left            =   7440
      MaxLength       =   150
      TabIndex        =   27
      Top             =   6000
      Width           =   4365
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
      Height          =   1050
      Index           =   32
      Left            =   7410
      MaxLength       =   150
      TabIndex        =   26
      Top             =   4560
      Width           =   4365
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
      Left            =   2175
      MaxLength       =   245
      TabIndex        =   94
      Top             =   5115
      Visible         =   0   'False
      Width           =   1275
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2460
      Width           =   5310
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
      Height          =   210
      Index           =   28
      Left            =   10245
      MaxLength       =   12
      TabIndex        =   22
      Top             =   2790
      Width           =   1200
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
      Height          =   480
      Index           =   31
      Left            =   8265
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   3510
      Width           =   3180
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
      Height          =   210
      Index           =   30
      Left            =   10245
      MaxLength       =   12
      TabIndex        =   24
      Top             =   3270
      Width           =   1200
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
      Height          =   210
      Index           =   29
      Left            =   10245
      TabIndex        =   23
      Top             =   3030
      Width           =   1200
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   27
      Left            =   1560
      TabIndex        =   5
      Text            =   "Help"
      Top             =   1185
      Width           =   2790
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
      Left            =   2640
      TabIndex        =   73
      Top             =   5280
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
         Picture         =   "frmJobGatePass.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   83
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
         Picture         =   "frmJobGatePass.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmJobGatePass.frx":0678
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
         TabIndex        =   81
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmJobGatePass.frx":0982
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
         TabIndex        =   80
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmJobGatePass.frx":0C8C
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
         Top             =   0
         Width           =   4695
      End
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
      Index           =   23
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   675
      Width           =   1200
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   24
      Left            =   5670
      MaxLength       =   25
      TabIndex        =   2
      Top             =   675
      Width           =   1200
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   26
      Left            =   1560
      MaxLength       =   90
      TabIndex        =   21
      Top             =   3735
      Width           =   5310
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   25
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   3
      Text            =   "Help"
      Top             =   930
      Width           =   2790
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   5670
      MaxLength       =   12
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "10-MAy-2003"
      Top             =   1185
      Width           =   1200
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   13
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3225
      Width           =   2940
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
      Index           =   22
      Left            =   10545
      TabIndex        =   32
      Top             =   1965
      Width           =   1080
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2205
      Width           =   5310
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2715
      Width           =   5310
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
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2970
      Width           =   5310
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
      Index           =   3
      Left            =   1560
      MaxLength       =   14
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1590
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
      Index           =   1
      Left            =   5670
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "Help"
      Top             =   930
      Width           =   1200
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
      Index           =   21
      Left            =   8400
      MaxLength       =   40
      TabIndex        =   33
      Top             =   2220
      Width           =   3225
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
      Index           =   4
      Left            =   4815
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2055
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   16
      Left            =   5295
      MaxLength       =   10
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1575
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   15
      Left            =   3540
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1395
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
      Index           =   7
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1590
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
      Index           =   19
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   30
      Top             =   1710
      Width           =   3225
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   6
      Left            =   4815
      MaxLength       =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1695
      Width           =   2055
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
      Index           =   17
      Left            =   8400
      MaxLength       =   8
      TabIndex        =   28
      Top             =   1455
      Width           =   990
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
      Index           =   18
      Left            =   10545
      MaxLength       =   25
      TabIndex        =   29
      Top             =   1455
      Width           =   1080
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   4815
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2055
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   14
      Left            =   1875
      MaxLength       =   25
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1260
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
      Index           =   20
      Left            =   8400
      MaxLength       =   8
      TabIndex        =   31
      Top             =   1965
      Width           =   705
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
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1590
   End
   Begin MSDataGridLib.DataGrid DGStaff 
      Height          =   2865
      Left            =   9270
      Negotiate       =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   6045
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   5054
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
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Caption         =   "Staff Name"
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
            ColumnWidth     =   3495.118
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCont 
      Height          =   2910
      Left            =   1470
      Negotiate       =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   5145
      Visible         =   0   'False
      Width           =   4740
      _ExtentX        =   8361
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
         Caption         =   "Contractor Name"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2625
      Left            =   195
      TabIndex        =   93
      Top             =   4620
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   4630
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      BackColorFixed  =   12632319
      ForeColorFixed  =   128
      BackColorSel    =   13166810
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   12632319
      GridColorFixed  =   33023
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "MW"
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work Instructions :"
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
      Height          =   225
      Index           =   25
      Left            =   7395
      TabIndex        =   98
      Top             =   5700
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Complaints :"
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
      Height          =   225
      Index           =   24
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Complaints :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   22
      Left            =   7380
      TabIndex        =   96
      Top             =   4290
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parts Enclosed :"
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
      Left            =   180
      TabIndex        =   95
      Top             =   4275
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "External Job Recd. Date                :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   17
      Left            =   7365
      TabIndex        =   92
      Top             =   2805
      Width           =   3090
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   21
      Left            =   7365
      TabIndex        =   91
      Top             =   3525
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contractor Bill No.                           :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   20
      Left            =   7365
      TabIndex        =   90
      Top             =   3045
      Width           =   3270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contractor Charges Rs.                :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   18
      Left            =   7365
      TabIndex        =   89
      Top             =   3285
      Width           =   3060
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contractor"
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
      Index           =   14
      Left            =   60
      TabIndex        =   87
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gate Pass No.*"
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
      Index           =   13
      Left            =   60
      TabIndex        =   72
      Top             =   690
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gate Pass Dt.*"
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
      Index           =   11
      Left            =   4380
      TabIndex        =   71
      Top             =   690
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose"
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
      Index           =   9
      Left            =   75
      TabIndex        =   70
      Top             =   3735
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Name*"
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
      Left            =   60
      TabIndex        =   69
      Top             =   945
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JC Open Dt.*"
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
      Left            =   4380
      TabIndex        =   68
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   60
      TabIndex        =   67
      Top             =   2460
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
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
      Left            =   60
      TabIndex        =   66
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name"
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
      Left            =   60
      TabIndex        =   65
      Top             =   2205
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Index           =   10
      Left            =   60
      TabIndex        =   64
      Top             =   3240
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of Vehicle on Floor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   7
      Left            =   7125
      TabIndex        =   63
      Top             =   1215
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Service"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   27
      Left            =   7125
      TabIndex        =   62
      Top             =   1710
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   36
      Left            =   7125
      TabIndex        =   61
      Top             =   1455
      Width           =   1035
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division            :"
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
      Left            =   7125
      TabIndex        =   60
      Top             =   735
      Width           =   1470
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job DocID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   7125
      TabIndex        =   59
      Top             =   975
      Width           =   960
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   30
      Left            =   10455
      TabIndex        =   57
      Top             =   1965
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "History Srl No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   15
      Left            =   9255
      TabIndex        =   56
      Top             =   1965
      Width           =   1245
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   5
      Left            =   8310
      TabIndex        =   55
      Top             =   2220
      Width           =   75
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Mechanic"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   16
      Left            =   7125
      TabIndex        =   54
      Top             =   2220
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No.*"
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
      Index           =   12
      Left            =   4380
      TabIndex        =   53
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
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
      Index           =   8
      Left            =   3510
      TabIndex        =   52
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(M)"
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
      Left            =   4965
      TabIndex        =   51
      Top             =   3480
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
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
      Left            =   3180
      TabIndex        =   50
      Top             =   3480
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(O)"
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
      Left            =   1530
      TabIndex        =   49
      Top             =   3480
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Serial No."
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
      Left            =   60
      TabIndex        =   48
      Top             =   1950
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No."
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
      Left            =   60
      TabIndex        =   47
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   1785
      Left            =   7035
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code      :"
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
      Left            =   8910
      TabIndex        =   46
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type"
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
      Left            =   3510
      TabIndex        =   45
      Top             =   1980
      Width           =   1125
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   16
      Left            =   8310
      TabIndex        =   44
      Top             =   1455
      Width           =   75
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   12
      Left            =   8310
      TabIndex        =   43
      Top             =   1710
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last KMs "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   32
      Left            =   7125
      TabIndex        =   42
      Top             =   1965
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Left            =   60
      TabIndex        =   41
      Top             =   1695
      Width           =   495
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   14
      Left            =   10455
      TabIndex        =   40
      Top             =   1455
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job Dt."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   34
      Left            =   9450
      TabIndex        =   39
      Top             =   1455
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
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
      Index           =   33
      Left            =   3510
      TabIndex        =   38
      Top             =   1695
      Width           =   915
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   9
      Left            =   8310
      TabIndex        =   37
      Top             =   1965
      Width           =   75
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   9675
      TabIndex        =   36
      Top             =   1230
      Width           =   75
   End
   Begin VB.Label LblTotVeh 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   11355
      TabIndex        =   35
      Top             =   1230
      Width           =   105
   End
End
Attribute VB_Name = "frmJobGatePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ForSiteCode$
Dim ADDFLAG$
Dim TAddMode As Boolean

Dim MyIndex As Byte
Dim Rst As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim rsCont As ADODB.Recordset
Dim RsJob As ADODB.Recordset
Dim RsStaff As ADODB.Recordset

'Text Box (Form)
Private Const JobNo As Byte = 1
Private Const JobDt As Byte = 2
Private Const VehRegNo As Byte = 3
Private Const Chassis As Byte = 4
Private Const Model As Byte = 5
Private Const Engine As Byte = 6
Private Const VehSrlNo As Byte = 7
Private Const SrvType As Byte = 8
Private Const OwnerName As Byte = 9
Private Const Address1 As Byte = 10
Private Const Address2 As Byte = 11
Private Const Address3 As Byte = 12
Private Const City As Byte = 13
Private Const PhoneOff As Byte = 14
Private Const PhoneResi As Byte = 15
Private Const Mobile As Byte = 16
Private Const LastJobNo As Byte = 17
Private Const LastJobDt As Byte = 18
Private Const LastSrv As Byte = 19
Private Const LastKMS As Byte = 20
Private Const LastMech As Byte = 21
Private Const HistNo As Byte = 22
Private Const GPNo As Byte = 23
Private Const GPDt As Byte = 24
Private Const MechName As Byte = 25
Private Const Purpose As Byte = 26
Private Const ContName As Byte = 27
Private Const RecdDate As Byte = 28
Private Const ContBillNo As Byte = 29
Private Const ContAmt As Byte = 30
Private Const Remarks As Byte = 31
Private Const Complaints As Byte = 32
Private Const Instructions As Byte = 33

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
'Grid Initializations
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_PartName As Byte = 1          ' Part Name
Private Const Col_Qty As Byte = 2               ' Quantity
Private Const Col_Recieved_YN As Byte = 3       ' Parts Recieved Yes/No
Private Const Col_TestReport_YN As Byte = 4     ' test reports Yes/No
Private Const Col_Complain As Byte = 5           'Complain if any
'Grid Color declaration
Private Const CellBackColLeave As String = &HC8E8DA
'Private Const CellForeColLeave As String = &H0&
'Private Const CellBackColEnter As String = &HC0E0FF
Private Const GridBackColorBkg As String = &HBAD3C9
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Dim mRepName As String

Private Sub DGCont_Click()
If rsCont.RecordCount > 0 Then
    Txt(MyIndex).TEXT = rsCont!Name
    Txt(MyIndex).Tag = rsCont!Code
End If
Txt(MyIndex).SetFocus
DGCont.Visible = False
End Sub

Private Sub DGJob_Click()
If Master.RecordCount > 0 Then
    Call History_Field
End If
Txt(MyIndex).SetFocus
DGJob.Visible = False
End Sub

Private Sub DGStaff_Click()
If RsStaff.RecordCount > 0 Then
    Txt(MyIndex).TEXT = RsStaff!Name
    Txt(MyIndex).Tag = RsStaff!Code
End If
Txt(MyIndex).SetFocus
DGStaff.Visible = False
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
Dim SrNo As Integer
    '' pending points :
    '' No Provision found for Incoming Time of Vehicle -- SKIP
    '' No Provision for Outgoing Time for Vehicle -- SKIP
    
    TopCtrl1.Tag = PubUParam: WinSetting Me:     Ini_Grid
    ForSiteCode = PubSiteCode
    Call BlankText
    
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  left(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select GP.GatePassNo+GP.Site_Code as SearchCode from Job_GatePass as GP where left(GatePassNo,1)='" & PubDivCode & "' " & sitecond & " Order by GP.GatePassDate Desc,GP.GatePassNo Desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 GP.GatePassNo+GP.Site_Code as SearchCode from Job_GatePass as GP where left(GatePassNo,1)='" & PubDivCode & "' " & sitecond & " Order by GP.GatePassDate Desc,GP.GatePassNo Desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
     sitecond = "and  " & cMID("J.DocId", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    Set RsJob = New ADODB.Recordset
    With RsJob
        .CursorLocation = adUseClient
        .Open "select  J.DocId AS CODE," & cCStr("J.Job_No") & " As FindJobNo,J.Job_No, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, J.DocId,J.Govt_YN, J.Job_Date, J.JobCloseDate,j.cardno, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName from ((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode  where left(j.DocId,1)='" & PubDivCode & "' " & sitecond & " Order by J.docID", GCn, adOpenDynamic, adLockOptimistic
    End With
    RsJob.Sort = "Code"
    Set DGJob.DataSource = RsJob
    
    Set rsCont = New ADODB.Recordset
    rsCont.CursorLocation = adUseClient
    rsCont.Open "Select FinCode as code,FinName as name FROM ContractFinance where FinCatg=1 Order by FinName", GCn, adOpenDynamic, adLockOptimistic
    Set DGCont.DataSource = rsCont
    rsCont.Sort = "Code"
    rsCont.Sort = "Name"
    
    Set RsStaff = New ADODB.Recordset
    RsStaff.CursorLocation = adUseClient
    RsStaff.Open "Select Emp_Code as code,Emp_Name as name FROM Emp_Mast Order by Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGStaff.DataSource = RsStaff
    RsStaff.Sort = "Name"
    Txt(GPDt).Tag = PubLoginDate
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ADDFLAG <> "B" Then
        If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsJob = Nothing
    Set RsStaff = Nothing
    Set rsCont = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer, mGPNo As Single
    RsJob.Filter = ("Jobclosedate = Null")
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    'mGPNo = GCn.Execute("select iif(isnull(max(right(GatePassNo,,4))),0,max(right(GatePassNo,4)))+1 from Job_GatePass where left(GatePassNo,1)='" & PubDivCode & "' and mid(GatePassNo,2,1)='" & PubSiteCode & "'").Fields(0)
    mGPNo = GCn.Execute("select " & vIsNull("max(" & cMID("GatePassNo", "3", "6") & ")", "0") & "+1 from Job_GatePass where left(GatePassNo,1)='" & PubDivCode & "' and " & cMID("GatePassNo", "2", "1") & "='" & PubSiteCode & "'").Fields(0)
    Txt(GPNo).TEXT = PubDivCode & PubSiteCode & mGPNo
    Txt(GPDt) = Txt(GPDt).Tag
    Txt(MechName).SetFocus
'   Txt(JobNo).SetFocus
    FGrid.Visible = True
    FGrid.Clear
    Ini_Grid
'    FGrid.Cols = 4
'    FGrid.ColWidth(2) = FGrid.width - (FGrid.ColWidth(0) + FGrid.ColWidth(1) + FGrid.ColWidth(3))
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
If GCn.Execute("select ExtJobGatePassNo from Job_Lab where ExtJobGatePassNo=''").RecordCount > 0 Then
    MsgBox "Delete Denied ! " & vbCrLf & "Gate Pass used in Labour", vbCritical, "Validation"
    Exit Sub
End If
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
                    
        GCn.Execute "Delete from Job_Lab  where job_Docid='" & lblDocId.CAPTION & "'"
    
        GCn.CommitTrans
        
        Master.Requery
        Call UpdRequery
        
        If Master.RecordCount > 0 Then
            Call MoveRec
        Else
            Call BlankText
        End If
        BUTTONS True, Me, Master, 0
    End If
    Exit Sub
eloop1:
    GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message"
End Sub

Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
    If Txt(JobNo) <> "" Then
        If Not IsNull(RsJob!JobCloseDate) Then
            MsgBox "JobCard is Closed,Editing not allowed", vbInformation, "Validation"
            Exit Sub
        End If
    End If
    Disp_Text SETS("EDIT", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    
    Txt(JobNo).Enabled = False
    Txt(Chassis).Enabled = False
    Txt(RecdDate).SetFocus
    FGrid.Visible = True
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Grid_Hide
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
'Select Case FGrid.Col
'    Case Col_PartNo
'        RsPart.Sort = "Code"
'        If FGrid.TextMatrix(FGrid.Row, Col_PartNo) <> "" Then
'            RsPart.MoveFirst
'            RsPart.FIND "Code='" & FGrid.TextMatrix(FGrid.Row, Col_PartNo) & "'"
'            If RsPart.EOF = True Then RsPart.MoveFirst
'        End If
'    Case Col_PartName
'        If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
'        RsPart.Sort = "PartName"
'        If FGrid.TextMatrix(FGrid.Row, Col_PartName) <> "" Then
'            RsPart.MoveFirst
'            RsPart.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_PartName) & "'"
'            If RsPart.EOF = True Then RsPart.MoveFirst
'        End If
'End Select
'Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then TxtGrid(0).TEXT = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case Col_PartName
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Complain
                End If
            End If
        Case Col_Qty
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown)) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Complain
                Else
                    TxtGrid(0).SetFocus
                End If
            End If
        Case Col_Recieved_YN
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown)) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Complain
                Else
                    TxtGrid(0).SetFocus
                End If
            End If
        Case Col_TestReport_YN
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown)) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Complain
                Else
                    TxtGrid(0).SetFocus
                End If
            End If
        Case Col_Complain
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown)) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Complain - 1
                Else
                    TxtGrid(0).SetFocus
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = vbKeyEscape Then Exit Sub
CheckQuote KeyAscii
Select Case FGrid.Col
'    Case Col_PartNo
'        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Code"
'    Case Col_PartName
'        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Name"
    Case Col_Qty
        NumPress TxtGrid(Index), KeyAscii, 6, 2
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
'        Case Col_PartNo
'            If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, KeyCode, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Code", True
'        Case Col_PartName
'            If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, KeyCode, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Name", True
        Case Col_Recieved_YN, Col_TestReport_YN
            If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                TxtGrid(Index) = ""
            ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
                TxtGrid(Index) = "Yes"
            Else
                TxtGrid(Index) = "No"
            End If
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0.00")
            CountItem
    End Select

End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)

Exit Sub
ELoop:
    CheckError

End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case Col_PartName
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
    Case Col_Qty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
    Case Col_Recieved_YN
        FGrid.TextMatrix(FGrid.Row, Col_Recieved_YN) = TxtGrid(0).TEXT
    Case Col_TestReport_YN
        FGrid.TextMatrix(FGrid.Row, Col_TestReport_YN) = TxtGrid(0).TEXT
    Case Col_Complain
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    End Select
    TxtGridLeave = True
    'Important at the time of validating  a control if you are making the visibility of
    'control false forcefully it will generate error
    If ValidateCall = False Then
        FGrid.SetFocus
            TxtGrid(0).Visible = False
    End If
End Function

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid_KeyPress vbKeyReturn
TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
On Error GoTo ELoop
SetMaxLength
    Select Case FGrid.Col
        Case Col_SrNo, Col_PartName, Col_Qty, Col_Recieved_YN, Col_TestReport_YN, Col_Complain
            Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case Col_Recieved_YN, Col_TestReport_YN, Col_Complain
            Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
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
                For I = 1 To FGrid.Rows - 1
                   FGrid.TextMatrix(I, Col_SrNo) = I
                Next
                CountItem
             End If
                'Recalculate Footer Values
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
    
End Sub

Private Sub TxtGridValid_PNo()
'Called from TxtGrid_Validate & TxtGridLeave procedures
End Sub
Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub TopCtrl1_eExit()
Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  left(gp.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "select GP.GatePassNo+GP.Site_Code as SearchCode,GP.GatePassNo,GP.Site_Code,GP.GatePassDate,GP.Purpose, Emp_Mast.Emp_Name as MechName " & _
        "from Job_GatePass as GP Left Join Emp_Mast on GP.Mech_Code=Emp_Mast.Emp_Code " & _
        "where left(Job_DocId,1)='" & PubDivCode & "' " & sitecond & " order by gp.GatePassNo"
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
        Set Master = GCn.Execute("select GP.GatePassNo+GP.Site_Code as SearchCode from Job_GatePass as GP where left(GatePassNo,1)='" & PubDivCode & "' and Site_Code='" & PubSiteCode & "' And GP.GatePassNo+GP.Site_Code  = '" & MyValue & "' Order by GP.GatePassDate Desc,GP.GatePassNo Desc")
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
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        RsJob.Filter = ""
        FGrid.Clear
        FGrid.Cols = 7
        Call Ini_Grid
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
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else
CmdPrint(PWindows).SetFocus
End Sub

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub
Private Sub TopCtrl1_eSave()
    Dim mTrans As Boolean, mGPNo As Single, I As Integer
    
'    On Error GoTo errlbl
    Grid_Hide
    If IsValid(Txt(GPNo), "GatePass No.") = False Then Exit Sub
    If IsValid(Txt(GPDt), "GatePass Dt.") = False Then Exit Sub
    If IsValid(Txt(MechName), "Staff name") = False Then Exit Sub
    If IsValid(Txt(Purpose), "Purpose") = False Then Exit Sub
    If ADDFLAG = "A" And Txt(JobNo) = "" Then
        If MsgBox("Job Card No. Empty, Continue ?", vbYesNo + vbCritical + vbDefaultButton2, "Job No. checking !") = vbNo Then
            Txt(JobNo).SetFocus
            Exit Sub
        End If
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PartName) = "" Then
            If FGrid.TextMatrix(I, Col_Qty) <> "" Then
                MsgBox "Please fill PartNo in Grid", vbInformation
                FGrid.Row = I: FGrid.Col = Col_PartName: FGrid.SetFocus: Exit Sub
            End If
        ElseIf FGrid.TextMatrix(I, Col_Qty) = "" Then
            If FGrid.TextMatrix(I, Col_PartName) <> "" Then
                MsgBox "Please fill Qty in Grid", vbInformation
                FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
            End If
        End If
    Next
    
    GCn.BeginTrans
    mTrans = True
    
    If ADDFLAG = "A" Then
        '' Get gate pass serial no
        'mGPNo = GCn.Execute("select iif(isnull(max(right(GatePassNo,4))),0,max(right(GatePassNo,4)))+1 from Job_GatePass where left(GatePassNo,1)='" & PubDivCode & "' and mid(GatePassNo,2,1)='" & PubSiteCode & "'").Fields(0)
        mGPNo = GCn.Execute("select " & vIsNull("max(" & cMID("GatePassNo", "3", "6") & ")", "0") & "+1 from Job_GatePass where left(GatePassNo,1)='" & PubDivCode & "' and " & cMID("GatePassNo", "2", "1") & "='" & PubSiteCode & "'").Fields(0)
        Txt(GPNo).TEXT = PubDivCode & PubSiteCode & mGPNo
        
        GSQL = "insert into Job_GatePass(" _
            & "GatePassNo,Site_Code,GatePassDate,Job_DocId,mech_code," _
            & "Purpose, U_Name, U_EntDt, U_AE,ContractCode,Complaints,Instructions) " _
            & " values(" _
            & "'" & Txt(GPNo) & "','" & PubSiteCode & "'," & ConvertDate(Txt(GPDt).TEXT) & ",'" & lblDocId.CAPTION & "','" & Txt(MechName).Tag & "'," _
            & "'" & Txt(Purpose).TEXT & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & ADDFLAG & "','" & Txt(ContName).Tag & "','" & Txt(Complaints) & "','" & Txt(Instructions) & "')"
        
    ElseIf ADDFLAG = "E" Then
        GSQL = "Update Job_GatePass set GatePassDate = " & ConvertDate(Txt(GPDt).TEXT) & ",mech_code='" & Txt(MechName).Tag & "',Purpose='" & Txt(Purpose).TEXT & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & ADDFLAG & _
            "',ContractCode='" & Txt(ContName).Tag & "', ContractRecdDate=" & ConvertDate(Txt(RecdDate)) & ",ContractorBillNo='" & Txt(ContBillNo) & "',ContractAmt=" & Val(Txt(ContAmt)) & ",Remarks='" & Txt(Remarks) & "', Complaints='" & Txt(Complaints) & "', Instructions = '" & Txt(Instructions) & "'" & _
            " Where GatePassNo='" & Txt(GPNo) & "' and Site_Code='" & PubSiteCode & "'"
                
        GCn.Execute ("Delete from Job_GatePass1 where GatePassNo='" & Txt(GPNo) & "'")
    End If
    For I = 1 To FGrid.Rows - 1
      If FGrid.TextMatrix(I, Col_PartName) <> "" And Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
          GCn.Execute ("insert into Job_GatePass1 (GatePassNo,Site_Code,Part_Name,Qty,Part_Rec,Test_Report,Complaint, U_Name, U_EntDt, U_AE) " & _
                " values('" & Txt(GPNo) & "','" & PubSiteCode & "','" & FGrid.TextMatrix(I, Col_PartName) & "'," & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & IIf(FGrid.TextMatrix(I, Col_Recieved_YN) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_TestReport_YN) = "Yes", 1, 0) & ",'" & FGrid.TextMatrix(I, Col_Complain) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & ADDFLAG & "')")
      End If
    Next
    
    GCn.Execute GSQL
    GCn.CommitTrans
    mTrans = False
    Txt(GPDt).Tag = Txt(GPDt)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select GP.GatePassNo+GP.Site_Code as SearchCode from Job_GatePass as GP where left(GatePassNo,1)='" & PubDivCode & "' and Site_Code='" & PubSiteCode & "' And GP.GatePassNo+GP.Site_Code  = '" & Txt(GPNo) & PubSiteCode & "' Order by GP.GatePassDate Desc,GP.GatePassNo Desc")
    End If
    Call UpdRequery
    RsJob.Filter = ""
    Master.FIND "SearchCode = '" & Txt(GPNo) & PubSiteCode & "'"
    If ADDFLAG = "A" Then TopCtrl1_ePrn
    Disp_Text SETS("INI", Me, Master)
    Ini_Grid
    If Txt(JobNo) <> "" Then
        Call MoveRec
    End If
    Exit Sub

errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case JobNo
            RsJob.Filter = ("Job_Date<=" & ConvertDate(Txt(GPDt)) & " and Jobclosedate = Null")
            If RsJob.RecordCount <= 0 Then GoTo lblEndSub
            DGridColSwap DGJob, 0
            RsJob.Sort = "JOB_NO"
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("JOB_NO='" & Txt(Index).TEXT & "'")
            End If
        Case Chassis
            RsJob.Filter = ("Job_Date<=" & ConvertDate(Txt(GPDt)) & " and Jobclosedate = Null")
            If RsJob.RecordCount <= 0 Then GoTo lblEndSub
            DGridColSwap DGJob, 1
            RsJob.Sort = "CHASSIS"
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("CHASSIS='" & Txt(Index).TEXT & "'")
            End If
        Case MechName
            If RsStaff.RecordCount <= 0 Then GoTo lblEndSub
            DGridColSwap DGStaff, 1
            RsStaff.Sort = "name"
            If Txt(Index).TEXT <> "" And Txt(Index).Tag <> RsStaff!Code Then
                RsStaff.FIND ("name='" & Txt(Index).TEXT & "'")
            End If
        Case ContName
            If rsCont.RecordCount <= 0 Then GoTo lblEndSub ' = True Or RsCont.BOF = True Or Txt(Index).Text = "" Then Exit Sub
            rsCont.Sort = "name"
            rsCont.MoveFirst
            rsCont.FIND "name='" & Txt(Index) & "'"
    End Select
lblEndSub:
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobNo
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 1
        Case Chassis
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 4
        Case MechName
            DGridTxtKeyDown DGStaff, Txt, Index, RsStaff, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case ContName
            DGridColSwap DGCont, 1
            DGridTxtKeyDown DGCont, Txt, Index, rsCont, KeyCode, False, 1, frmContract, "frmContract"
    End Select
    If DGJob.Visible = False And DGStaff.Visible = False And DGCont.Visible = False Then
        '' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index <> Instructions) Or (ADDFLAG = "E" And Index <> Instructions)) Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index = Instructions) Or (ADDFLAG = "E" And Index = Instructions)) Then
            FGrid.Row = 1: FGrid.Col = 1: FGrid.SetFocus
            'If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        
        ' KEY UP
        If ADDFLAG = "A" Then
            If Index <> GPNo Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf ADDFLAG = "E" Then
            If Index <> GPDt Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
'        Case GPNo
'            Call NumPress(Txt(Index), KeyAscii, 8, 0)
        Case JobNo
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "Findjobno"
        Case Chassis
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "chassis"
        Case MechName
            DGridTxtKeyPress Txt, Index, RsStaff, KeyAscii, "name"
        Case ContName
            DGridTxtKeyPress Txt, Index, rsCont, KeyAscii, "Name"
        Case ContAmt
            Call NumPress(Txt(Index), KeyAscii, 6, 2)
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case JobNo, Chassis ', VehRegNo, Model, VehSrlNo, OwnerName
'            If Txt(Index).Tag <> "" Then
            If Txt(Index) <> "" Then
'                RsJob.Sort = "CODE"
                RsJob.MoveFirst
                RsJob.FIND ("CODE='" & Txt(Index).Tag & "'")
                If RsJob.BOF = True Or RsJob.EOF = True Then Exit Sub
                LblDiv.CAPTION = "Division : " & left(RsJob!DocID, 1)
                LblSite.CAPTION = "Site Code : " & DeCodeDocID(RsJob!DocID, Current_Site)
                lblDocId = RsJob!DocID
                Call History_Field
            Else
                Call History_Field(True)
            End If
        Case MechName
            If RsStaff.EOF Or RsStaff.BOF Or Txt(Index) = "" Then
                Txt(MechName).Tag = ""
                Txt(MechName).TEXT = ""
            Else
                Txt(MechName).Tag = RsStaff!Code
                Txt(MechName).TEXT = RsStaff!Name
            End If
        Case ContName
            If rsCont.EOF Or rsCont.BOF Or Txt(Index) = "" Then
                Txt(ContName).Tag = ""
                Txt(ContName).TEXT = ""
            Else
                Txt(ContName).Tag = rsCont!Code
                Txt(ContName).TEXT = rsCont!Name
            End If
        Case GPDt
            Txt(GPDt).TEXT = RetDate(Txt(GPDt))
        Case RecdDate
            Txt(RecdDate).TEXT = RetDate(Txt(RecdDate))
            If Txt(RecdDate) <> "" Then
                If CDate(Txt(RecdDate)) < CDate(Txt(GPDt)) Then
                    MsgBox "External Job recd. Date is less than Gate Pass Date", vbCritical, "Date Validation"
                    Cancel = True
                End If
            End If
    End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 1 To Txt.Count
        Txt(I).TEXT = ""
        If I <> GPDt Then Txt(I).Tag = ""
    Next I
    lblDocId.CAPTION = ""
    lblDocId.Refresh
End Sub

Private Sub MoveRec()
Dim Rs As Recordset, Master1 As Recordset, rs1 As Recordset
Dim mVor As String
Dim I As Integer
On Error GoTo error1
    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select GP.*, Emp_Mast.Emp_Name as MechName,ContractFinance.FinName " & _
            " from (Job_GatePass as GP Left Join Emp_Mast on GP.Mech_Code=Emp_Mast.Emp_Code) " & _
            " Left Join  ContractFinance on GP.ContractCode=ContractFinance.FinCode " & _
            " Where GP.GatePassNo+GP.Site_Code='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
            
        Set rs1 = New Recordset
        rs1.CursorLocation = adUseClient
        rs1.Open "Select * from Job_GatePass1 where GatePassNo ='" & Master1!GatePassNo & "'", GCn, adOpenStatic, adLockReadOnly
            
        LblDiv.CAPTION = "Division : " & left(Master1!job_docid, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        lblDocId.CAPTION = Master1!job_docid
        
        Txt(GPNo).TEXT = XNull(Master1!GatePassNo)
        Txt(GPDt).TEXT = XNull(Master1!GatePassDate)
        Txt(Purpose).TEXT = XNull(Master1!Purpose)
        Txt(MechName).TEXT = XNull(Master1!MechName)
        Txt(MechName).Tag = XNull(Master1!mech_code)
        Txt(ContName).TEXT = XNull(Master1!FinName)
        Txt(ContName).Tag = XNull(Master1!ContractCode)
        Txt(RecdDate).TEXT = IIf(IsNull(Master1!ContractRecdDate), "", Master1!ContractRecdDate)
        Txt(ContBillNo).TEXT = IIf(IsNull(Master1!ContractorBillNo), "", Master1!ContractorBillNo)
        Txt(ContAmt).TEXT = IIf(Master1!ContractAmt = 0, "", Format(Master1!ContractAmt, "0.00"))
        Txt(Remarks).TEXT = IIf(IsNull(Master1!Remarks), "", Master1!Remarks)
        Txt(Complaints).TEXT = IIf(IsNull(Master1!Complaints), "", Master1!Complaints)
        Txt(Instructions).TEXT = IIf(IsNull(Master1!Instructions), "", Master1!Instructions)
        RsJob.Sort = "code"
        RsJob.FIND ("Code='" & Master1!job_docid & "'")
'        If RsJob.EOF Or RsJob.BOF Then Exit Sub
        FGrid.Clear
        If rs1.RecordCount <> 0 Then
            Ini_Grid
            FGrid.Rows = rs1.RecordCount + 1
            For I = 1 To rs1.RecordCount
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PartName) = XNull(rs1.Fields("Part_Name"))
                    .TextMatrix(I, Col_Qty) = XNull(rs1.Fields("Qty"))
                    .TextMatrix(I, Col_Recieved_YN) = IIf(VNull(rs1.Fields("Part_Rec")) = "1", "Yes", "No")
                    .TextMatrix(I, Col_TestReport_YN) = IIf(VNull(rs1.Fields("Test_Report")) = "1", "Yes", "No")
                    .TextMatrix(I, Col_Complain) = XNull(rs1.Fields("Complaint"))
                End With
                rs1.MoveNext
            Next
        Else
            Ini_Grid
        End If
        Call History_Field(RsJob.EOF)
        Call veh_count
    Else
        Call BlankText
    End If
    Grid_Hide
    Set Rs = Nothing
    Set Master1 = Nothing
    Exit Sub
error1:
    CheckError
End Sub
Private Sub Ini_Grid()
  With FGrid
        .Rows = 2
        .Cols = 6
        .left = Me.left + 60
        .width = 7200
        .top = Txt(Purpose).top + Txt(Purpose).height + 400
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        

        .TextMatrix(0, Col_SrNo) = "S.No"
        .TextMatrix(1, Col_SrNo) = 1
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 420

        .TextMatrix(0, Col_PartName) = "Part Name"
        .ColAlignment(Col_PartName) = flexAlignLeftCenter
        .ColWidth(Col_PartName) = 2500

        .TextMatrix(0, Col_Qty) = "Qty"
        .ColAlignment(Col_Qty) = flexAlignLeftCenter
        .ColWidth(Col_Qty) = 620
        
        .TextMatrix(0, Col_Recieved_YN) = "Recieved"
        .ColAlignment(Col_Recieved_YN) = flexAlignLeftCenter
        .ColWidth(Col_Recieved_YN) = 400

        .TextMatrix(0, Col_TestReport_YN) = "Test Report"
        .ColAlignment(Col_TestReport_YN) = flexAlignLeftCenter
        .ColWidth(Col_TestReport_YN) = 400
        
        .TextMatrix(0, Col_Complain) = "Complaint"
        .ColAlignment(Col_Complain) = flexAlignLeftCenter
        .ColWidth(Col_Complain) = 2500
        BackColorSelLeave = FGrid.BackColorSel
        ForeColorSelEnter = FGrid.ForeColorSel
    
End With
    DGJob.left = Me.left: DGJob.width = Me.width - 90: DGJob.top = Txt(Purpose).top + Txt(Purpose).height: DGJob.height = 3000
    DGStaff.width = 4740: DGStaff.left = Shape2.left: DGStaff.top = mTopScale: DGStaff.height = 5000
    DGCont.width = 5000: DGCont.left = Me.width - (DGCont.width + mRtScale): DGCont.top = mTopScale: DGCont.height = 5000
End Sub

Private Sub CountItem()
Dim I As Integer, TotItems As Integer, TotQty As Double
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PartName) <> "" Then
            TotQty = TotQty + Val(FGrid.TextMatrix(I, Col_Qty))
            TotItems = TotItems + 1
        End If
    Next I

End Sub

Private Function ChkDuplicate() As Boolean
Dim I As Integer, X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte, Col4 As Byte
Select Case FGrid.Col
    Case Col_PartName
        Col2 = FGrid.Col
End Select
    X = UCase(CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col2))))
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

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    'New Testing for Speed purpose
    ADDFLAG = left(TopCtrl1.TopText2, 1)
    'eof New Testing
    For I = 1 To Txt.Count
        Txt(I).Enabled = Enb
    Next
    
    For I = 1 To Txt.Count
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    
    Txt(JobDt).Enabled = False
    Txt(GPNo).Enabled = False
    Txt(Engine).Enabled = False
    Txt(SrvType).Enabled = False
    
    Txt(Address1).Enabled = False
    Txt(Address2).Enabled = False
    Txt(Address3).Enabled = False
    Txt(City).Enabled = False
    Txt(PhoneOff).Enabled = False
    Txt(PhoneResi).Enabled = False
    Txt(Mobile).Enabled = False
    
    Txt(LastJobDt).Enabled = False
    Txt(LastJobNo).Enabled = False
    Txt(LastSrv).Enabled = False
    Txt(LastKMS).Enabled = False
    Txt(LastMech).Enabled = False
    Txt(HistNo).Enabled = False
    If ADDFLAG = "A" Then
        Txt(RecdDate).Enabled = False
        Txt(ContAmt).Enabled = False
        Txt(ContBillNo).Enabled = False
        Txt(Remarks).Enabled = False
    End If
End Sub

Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
    If DGStaff.Visible = True Then DGStaff.Visible = False
End Sub

Private Sub veh_count()
    If Txt(JobDt).TEXT <> "" Then
        LblTotVeh.CAPTION = GCn.Execute("select count(*) from job_Card where JobCloseDate = " & ConvertDate("01/Jan/1900") & " or JobCloseDate Is Null  and left(Docid,1)='" & PubDivCode & "' ").Fields(0)
    End If
End Sub

Private Sub UpdRequery()
    rsCont.Requery
    RsJob.Requery
    RsStaff.Requery
End Sub

Private Sub History_Field(Optional MakeBlank As Boolean)
If MakeBlank Then
    Txt(HistNo).Tag = ""
    Txt(HistNo).TEXT = ""
    
    Txt(VehRegNo).Tag = ""
    Txt(Chassis).Tag = ""
    Txt(Model).Tag = ""
    Txt(VehSrlNo).Tag = ""
    Txt(OwnerName).Tag = ""
    Txt(JobNo).Tag = ""
    
    Txt(JobNo).TEXT = ""
    Txt(JobDt).TEXT = ""
    Txt(SrvType).TEXT = ""
    Txt(VehRegNo).TEXT = ""
    Txt(Chassis).TEXT = ""
    Txt(Model).TEXT = ""
    Txt(Engine).TEXT = ""
    Txt(VehSrlNo).TEXT = ""
    Txt(OwnerName).TEXT = ""
    Txt(Address1).TEXT = ""
    Txt(Address2).TEXT = ""
    Txt(Address3).TEXT = ""
    Txt(City).TEXT = ""
    Txt(PhoneOff).TEXT = ""
    Txt(PhoneResi).TEXT = ""
    Txt(Mobile).TEXT = ""
Else
    Txt(HistNo).Tag = RsJob!CardNo
    Txt(HistNo).TEXT = RsJob!CardNo
    
    Txt(VehRegNo).Tag = XNull(RsJob!Code)
    Txt(Chassis).Tag = XNull(RsJob!Code)
    Txt(Model).Tag = XNull(RsJob!Code)
    Txt(VehSrlNo).Tag = XNull(RsJob!Code)
    Txt(OwnerName).Tag = XNull(RsJob!Code)
    Txt(JobNo).Tag = XNull(RsJob!Code)  'additional
    Txt(JobNo).TEXT = XNull(RsJob!Job_No)
    Txt(JobDt).TEXT = RsJob!Job_Date
    Txt(SrvType).TEXT = XNull(RsJob!Serv_Desc)
    Txt(VehRegNo).TEXT = XNull(RsJob!RegNo)
    Txt(Chassis).TEXT = XNull(RsJob!Chassis)
    Txt(Model).TEXT = XNull(RsJob!Model)
    Txt(Engine).TEXT = XNull(RsJob!Engine)
    Txt(VehSrlNo).TEXT = XNull(RsJob!VehSerialNo)
    Txt(OwnerName).TEXT = XNull(RsJob!Name)
    Txt(Address1).TEXT = XNull(RsJob!Add1)
    Txt(Address2).TEXT = XNull(RsJob!Add2)
    Txt(Address3).TEXT = XNull(RsJob!Add3)
    Txt(City).TEXT = XNull(RsJob!CityName)
    Txt(PhoneOff).TEXT = XNull(RsJob!PhoneOff)
    Txt(PhoneResi).TEXT = XNull(RsJob!PhoneResi)
    Txt(Mobile).TEXT = XNull(RsJob!Mobile)
End If
Call UpdLastJC
End Sub

Private Sub UpdLastJC()
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open "SELECT Top 1 JOB_NO,JOB_DATE,AtKMsHrs,Srv.Serv_SrlNo,Srv.Serv_Type,Srv.SERV_DESC AS SrvDesc,EMP_MAST.EMP_NAME AS MECH_NAME " & _
            " FROM (JOB_CARD LEFT JOIN Service_Type Srv ON JOB_CARD.SERV_TYPE=Srv.SERV_TYPE) " & _
            " LEFT JOIN EMP_MAST ON JOB_CARD.RECBY_MECHANIC=EMP_MAST.EMP_CODE " & _
            " WHERE CARDNO='" & Txt(HistNo).TEXT & _
            "' and Job_Date< " & ConvertDate(Txt(JobDt)) & _
            " ORDER BY JOB_DATE Desc ", GCn, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        Txt(LastJobNo).TEXT = XNull(RsTemp!Job_No)
        Txt(LastJobDt).TEXT = RsTemp!Job_Date
        Txt(LastKMS).TEXT = VNull(RsTemp!AtKMsHrs)
        Txt(LastSrv).TEXT = XNull(RsTemp!SrvDesc)
        Txt(LastSrv).Tag = VNull(RsTemp!Serv_SrlNo)
        Txt(LastMech).TEXT = XNull(RsTemp!MECH_NAME)
    Else
        Txt(LastJobNo).TEXT = "":           Txt(LastJobDt).TEXT = ""
        Txt(LastKMS).TEXT = "":             Txt(LastSrv).TEXT = ""
        Txt(LastMech).TEXT = "":             Txt(LastSrv).Tag = ""
    End If
    Set RsTemp = Nothing
End Sub
Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmPrn.Visible = False
        If Index <> PSetUp And ADDFLAG <> "B" Then
            If ADDFLAG = "A" Then TopCtrl1_eAdd: Exit Sub
            Disp_Text SETS("INI", Me, Master)
            Call MoveRec
        End If
    End If
End Sub
Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
GSQL = "SELECT JG.Job_DocID,JC.Job_Date,JG.GatePassNo,JG.GatePassDate,JG.Mech_Code,JG.Purpose,JG.Job_DocID,JG.Complaints,JG.Instructions,HC.MODEL,HC.RegNo,HC.RegDate,HC.Name,HC.Add1 as CAdd1,HC.Add2 as CAdd2,HC.Add3 as CAdd3,HC.Chassis,HC.Engine,City.CityName,Emp_Mast.Emp_Name,CF.FinName,CF.Add1,CF.Add2,City1.CityName as ContCity " & _
    " FROM (((((Job_GatePass as JG LEFT JOIN JOB_CARD as JC ON JG.Job_DocID=JC.DocId) " & _
    " LEFT JOIN HISCARD as HC ON JC.CardNo=HC.CardNo) " & _
    " LEFT JOIN EMP_MAST ON JG.MECH_CODE=EMP_MAST.EMP_CODE) " & _
    " Left Join ContractFinance as CF on JG.ContractCode=CF.FinCode) " & _
    " Left Join City on HC.CityCode=City.CityCode) " & _
    " Left Join City as City1 on CF.City=City1.CityCode WHERE JG.GatePassNo='" & Txt(GPNo) & "'"
Select Case Index
    Case PScreen, PWindows
        If Txt(JobNo) <> "" Then
            mRepName = IIf(OptPlain.Value = True, "GatePass", "GatePass")
        Else
            mRepName = IIf(OptPlain.Value = True, "NonJobGatePass", "NonJobGatePass")
        End If
        Call WindowsPrint(GSQL, Index)
        FrmPrn.Visible = False
    Case PDos
        If Txt(JobNo) = "" Then
            Call SpeedPrint(GSQL)
            FrmPrn.Visible = False
        Else
            Call SpeedPrint1(GSQL)
            FrmPrn.Visible = False
        End If
        
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "GatePass", "GatePass")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And ADDFLAG <> "B" Then
    If ADDFLAG = "A" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrint(mQry As String, Index As Integer)
On Error GoTo ERRORHANDLER
Dim Rst As ADODB.Recordset
Dim RST1 As ADODB.Recordset
Dim mReportCount As Integer, I As Integer
 
Set RST1 = GCn.Execute(mQry)
Set Rst = GCn.Execute("Select GatePassNo,Part_Name,Site_Code,Qty,Part_Rec,Test_Report,Complaint " & _
    " from Job_GatePass1 where GatePassNo='" & RST1!GatePassNo & "' Order By GatePassNo,Part_Name")

CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("TITLE1")
            rpt.FormulaFields(I).TEXT = "'** GATE PASS **'"
    End Select
Next
     
rpt.Database.SetDataSource Rst
rpt.ReadRecords
Set Rst = Nothing

Select Case Index
    Case PWindows, PScreen  'Printer
        If UCase(mRepName) = "GATEPASS" Then
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case UCase("Title")
                        rpt.FormulaFields(I).TEXT = "'** GATE PASS **'"
                    Case UCase("GatePassNo")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!GatePassNo & "'"
                    Case UCase("Contractor")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!FinName & "'"
                    Case UCase("Cont_Add1")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Add1 & "'"
                    Case UCase("Cont_Add2")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Add2 & "'"
                    Case UCase("Cont_City")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!ContCity & "'"
                    Case UCase("Job_No")
                         rpt.FormulaFields(I).TEXT = "'" & PrinID(RST1!job_docid) & "'"
                    Case UCase("Job_Date")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Job_Date & "'"
                    Case UCase("Cust_Name")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Name & "'"
                    Case UCase("Cust_add1")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!CAdd1 & "'"
                    Case UCase("Cust_add2")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!CAdd2 & "'"
                    Case UCase("Cust_Add3")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!CAdd3 & "'"
                    Case UCase("Cust_City")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!CityName & "'"
                    Case UCase("Reg_No")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!RegNo & "'"
                    Case UCase("Chassis_No")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Chassis & "'"
                    Case UCase("Engine_No")
                        rpt.FormulaFields(I).TEXT = "' " & RST1!Engine & " '"
                    Case UCase("Complaints")
                        rpt.FormulaFields(I).TEXT = "'" & Trim(RST1!Complaints) & "'"
                    Case UCase("Instructions")
                        rpt.FormulaFields(I).TEXT = "'" & Trim(RST1!Instructions) & "'"
                End Select
            Next
        Else
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case UCase("Title")
                        rpt.FormulaFields(I).TEXT = "'** GATE PASS **'"
                    Case UCase("GatePassNo")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!GatePassNo & "'"
                    Case UCase("Contractor")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!FinName & "'"
                    Case UCase("Cont_Add1")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Add1 & "'"
                    Case UCase("Cont_Add2")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Add2 & "'"
                    Case UCase("Cont_City")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!ContCity & "'"
                    Case UCase("StaffName")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Emp_Name & "'"
                    Case UCase("Purpose")
                        rpt.FormulaFields(I).TEXT = "'" & RST1!Purpose & "'"
                End Select
            Next
        End If
End Select

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
            End Select
        Next
        rpt.PrintOut False
    Case PScreen  'screen
        Call Report_View(rpt, "** GATE PASS **", , True)
End Select

CmdPrint(PSetUp).Tag = ""
Set rpt = Nothing
Set RST1 = Nothing
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

Private Sub SpeedPrint(mQry$)
On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Purpose 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstGate As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mLftMargin$
    mLftMargin = "    "
    Set RstGate = GCn.Execute(mQry)
    If RstGate.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 0
     
    PageLength = PubPageLengthHalf
    PageWidth = 80   '137 for chr15
    
    mHeader = 0   'Ideal 17
    mFooter = 7 + 4
          
    'Header
    mDocStr = "** GATE PASS (Non Job) **"
    Print #1, Chr(27) + Chr(67) + Chr(PageLength) ' instead of Print #1,meject
    mHeader = mHeader + 1
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    If PubComp_Add2 <> "" Then
        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(mDocStr, "A", PageWidth) & mChr18 '& mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & mEmph & PSTR("Gate Pass No.", 16) & " : " & RstGate!GatePassNo & Space(10) & PSTR("Gate Pass Date", 16) & " : " & CDate(RstGate!GatePassDate) & mEmph1
    mHeader = mHeader + 1
    Print #1, mLftMargin & "To,"
    mHeader = mHeader + 1
    Print #1, mLftMargin & "Security,"
    mHeader = mHeader + 1
    Print #1, mLftMargin & "Vehicle/Goods as per details permitted to leave workshop:"
    mHeader = mHeader + 1
    Print #1, mLftMargin & PSTR("Contractor Name", 15) & " : " & XNull(RstGate!FinName)
    mHeader = mHeader + 1
    Print #1, mLftMargin & PSTR("Address", 15) & " : " & PSTR(XNull(RstGate!Add1), 40)
    mHeader = mHeader + 1
    If XNull(RstGate!Add2) <> "" Then
        Print #1, mLftMargin & Space(18) & XNull(RstGate!Add2)
        mHeader = mHeader + 1
    End If
    If XNull(RstGate!ContCity) <> "" Then
        Print #1, mLftMargin & Space(18) & XNull(RstGate!ContCity)
        mHeader = mHeader + 1
    End If
    Print #1, mLftMargin & PSTR("Staff", 15) & " : " & RstGate!Emp_Name
    mHeader = mHeader + 1
    If Len(RstGate!Purpose) <= 45 Then
        Print #1, mLftMargin & PSTR("Purpose", 15) & " : " & XNull(RstGate!Purpose)
        mHeader = mHeader + 1
    ElseIf Len(RstGate!Purpose) > 45 Then
        Print #1, mLftMargin & PSTR("Purpose", 15) & "   " & XNull(left(RstGate!Purpose, 45))
        mHeader = mHeader + 1
        Print #1, mLftMargin & Space(15) & "   " & XNull(mID(RstGate!Purpose, 46, 44))
        mHeader = mHeader + 1
    End If
    
    Do Until mHeader >= PageLength - mFooter
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    Print #1, ""
    Print #1, mLftMargin & "Customer" & Space(15) & "Auth. Signatory" & Space(15) & "Security"
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("")) / 2) & "" & mChr18
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
Private Sub SpeedPrint1(mQry$)
On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Purpose 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer
    Dim PrintStr$, PrintDate$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstGate As ADODB.Recordset, RstGate1 As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mLftMargin$
    mLftMargin = "    "
    PrintDate = date
    Set RstGate = GCn.Execute(mQry)
    Set RstGate1 = GCn.Execute("Select GatePassNo,Part_Name,Site_Code,Qty,Part_Rec,Test_Report,Complaint " & _
    " from Job_GatePass1 where GatePassNo='" & RstGate!GatePassNo & "' Order By GatePassNo,Part_Name")
    If RstGate.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 0
     
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    
    mHeader = 0   'Ideal 17
    mFooter = 13
    
    
          
    'Header
    mDocStr = "** GATE PASS **"
    Print #1, Chr(27) + Chr(67) + Chr(PageLength) ' instead of Print #1,meject
    mHeader = mHeader + 1
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    If PubComp_Add2 <> "" Then
        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(mDocStr, "A", PageWidth) & mChr18 '& mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & mEmph & PSTR("Gate Pass No.", 16) & " : " & RstGate!GatePassNo & Space(25) & PSTR("Gate Pass Date", 16) & " :" & CDate(RstGate!GatePassDate) & mEmph1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & "To," & Space(55) & PSTR("Print Date :", 12, , AlignRight) & PSTR(PrintDate, 12, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & "M/S " & PSTR(RstGate!FinName, 54, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & IIf(RstGate!Add1 <> "", RstGate!Add1 & ",", "") & IIf(RstGate!Add2 <> "", RstGate!Add2, "")
    mHeader = mHeader + 1
    Print #1, mLftMargin & IIf(RstGate!ContCity <> "", RstGate!ContCity, "")
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & "Dear Sir,"
    mHeader = mHeader + 1
    Print #1, mLftMargin + "      " & " We are enclosing here with the following components/assembly along with our "
    mHeader = mHeader + 1
    Print #1, mLftMargin & "observations.Please carry out the necessary repair/service"
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & "Vehicele Details :"
    Print #1, ""
    mHeader = mHeader + 1

    Print #1, mLftMargin & mEmph & PSTR("Job No.       :", 15, , AlignLeft) & mEmph1 & PSTR(PrinID(RstGate!job_docid), 30, , AlignLeft) & mEmph & PSTR("Registration No.:", 18, , AlignRight) & mEmph1 & PSTR(RstGate!RegNo, 20, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & mEmph & PSTR("Customer Name :", 15, , AlignLeft) & mEmph1 & PSTR(RstGate!Name, 30, , AlignLeft) & mEmph & PSTR("Chassis No.     :", 18, , AlignRight) & mEmph1 & PSTR(RstGate!Chassis, 20, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & mEmph & PSTR("Address       :", 15, , AlignLeft) & mEmph1 & PSTR(RstGate!CAdd1, 30, , AlignLeft) & mEmph & PSTR("Job Date        :", 18, , AlignRight) & mEmph1 & RstGate!Job_Date
    mHeader = mHeader + 1
    Print #1, mLftMargin & "           " & Space(4) & PSTR(RstGate!CAdd2, 31, , AlignLeft) & mEmph & PSTR("Engine No.      :", 17, , AlignLeft) & mEmph1 & PSTR(RstGate!Engine, 20, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & "           " & Space(4) & PSTR(RstGate!CAdd3, 32, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & "           " & Space(4) & PSTR(RstGate!CityName, 32, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & mEmph & "Complaint Reported By Customer" & mEmph1
    mHeader = mHeader + 1
    Print #1, mLftMargin & PSTR(left(RstGate!Complaints, 60), 60, , AlignLeft)
    mHeader = mHeader + 1
    If Len(RstGate!Complaints) > 60 And Len(RstGate!Complaints) < 120 Then
        Print #1, mLftMargin & PSTR(mID(RstGate!Complaints, 60, 60), 60, , AlignLeft)
        mHeader = mHeader + 1
    ElseIf Len(RstGate!Complaints) > 120 And Len(RstGate!Complaints) < 150 Then
        Print #1, mLftMargin & PSTR(Right(RstGate!Complaints, 60), 60, , AlignLeft)
        mHeader = mHeader + 1
    End If
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & mEmph & "Work Instructions" & mEmph1
    mHeader = mHeader + 1
    Print #1, mLftMargin & PSTR(left(RstGate!Instructions, 60), 60, , AlignLeft)
    mHeader = mHeader + 1
    If Len(RstGate!Instructions) > 60 And Len(RstGate!Instructions) < 120 Then
        Print #1, mLftMargin & PSTR(mID(RstGate!Instructions, 60, 60), 60, , AlignLeft)
        mHeader = mHeader + 1
    ElseIf Len(RstGate!Instructions) > 120 And Len(RstGate!Instructions) < 150 Then
        Print #1, mLftMargin & PSTR(Right(RstGate!Instructions, 60), 60, , AlignLeft)
        mHeader = mHeader + 1
    End If
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & "Encloser Parts Detail"
    mHeader = mHeader + 1
    
'*******************************************************

    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & PSTR("Part Name", 30, , AlignLeft) & PSTR("Part Sr.No.(If Any)", 19, , AlignLeft) & PSTR(" Qty", 7, , AlignRight) & "  " & PSTR("Complaints", 25, , AlignLeft) & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
    
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter + 5)
    Page = 1
    mLine = 1
    mSlNo = 1
    
    If RstGate1.RecordCount > 0 Then
        I = 1
        Do While Not RstGate1.EOF = True
        
            If mLine > mFix Then
                Page = Page + 1
                Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-")
                Print #1, Space(PageWidth - Len("Contd. on next page.." + STR(Page))) & "Contd. on next page.." & STR(Page)
                Do Until mLine >= mFix + mFooter - 2
                    Print #1, ""
                    mLine = mLine + 1
                Loop
               'Header On Second Page
                mHeader = 0
'               Print #1, Chr(27) + Chr(67) + Chr(PageLength) ' instead of Print #1,meject
'                mHeader = mHeader + 1
                                
                Print #1, Chr(27) + Chr(67) + Chr(PageLength) ' instead of Print #1,meject
                mHeader = mHeader + 1
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
                If PubComp_Add2 <> "" Then
                    Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
                    mHeader = mHeader + 1
                End If
                If PubComp_City <> "" Then
                    Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
                    mHeader = mHeader + 1
                End If
                Print #1, PRN_TIT(mDocStr, "A", PageWidth) & mChr18 '& mEmph
                mHeader = mHeader + 1
                Print #1, mLftMargin & ""
                mHeader = mHeader + 1
                Print #1, mLftMargin & mEmph & PSTR("Gate Pass No.", 16) & " : " & RstGate!GatePassNo & Space(25) & PSTR("Gate Pass Date", 16) & " :" & CDate(RstGate!GatePassDate) & mEmph1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, mLftMargin & "To," & Space(55) & PSTR("Print Date :", 12, , AlignRight) & PSTR(PrintDate, 12, , AlignLeft)
                mHeader = mHeader + 1
                Print #1, mLftMargin & "M/S " & PSTR(RstGate!FinName, 54, , AlignLeft)
                mHeader = mHeader + 1
                Print #1, mLftMargin & IIf(RstGate!Add1 <> "", RstGate!Add1 & ",", "") & IIf(RstGate!Add2 <> "", RstGate!Add2, "")
                mHeader = mHeader + 1
                Print #1, mLftMargin & IIf(RstGate!ContCity <> "", RstGate!ContCity, "")
                mHeader = mHeader + 1
                Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
                mHeader = mHeader + 1
                Print #1, mLftMargin & PSTR("Part Name", 30, , AlignLeft) & PSTR("Part Sr.No.(If Any)", 19, , AlignLeft) & PSTR(" Qty", 7, , AlignRight) & "  " & PSTR("Complaints", 25, , AlignLeft) & mEmph
                mHeader = mHeader + 1
                Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                Print #1, PrintStr
                mLine = 1
             End If
            Print #1, mLftMargin & PSTR(RstGate1!Part_Name, 30, , AlignLeft) & "                   " & PSTR(Format(RstGate1!Qty, "00.00"), 7, , AlignRight) & "   " & PSTR(RstGate1!Complaint, 25, , AlignLeft)
            mHeader = mHeader + 1
            RstGate1.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    
    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & "|" & Space(30) & "|" & Space(8) & "To be filled after recieving the part  |" & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & "|         |  Part Dispatched   |  Part Recieved   | Test Report  |  Complaint  |" & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & "|  Date   |                    |                  |              |             |" & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & "|  Time   |                    |                  |              |             |" & mEmph
    mHeader = mHeader + 1
    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-") & mEmph
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mLftMargin & PSTR("Signature", 40, , AlignLeft) & PSTR("Signature", 30, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, mLftMargin & PSTR("Tata Motors Dealer/TASS", 40, , AlignLeft) & PSTR("Ancillary Authoriesed Setup", 30, , AlignLeft)
    Print #1, mLftMargin & Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("")) / 2) & "" & mChr18
    'Print #1, mEject
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
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub


Private Sub SetMaxLength()
    Select Case FGrid.Col
        Case Col_PartName
            TxtGrid(0).MaxLength = 40
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Col_Qty
            TxtGrid(0).MaxLength = 10
            TxtGrid(0).Alignment = 1   '0-Left Align
        Case Col_Recieved_YN, Col_TestReport_YN
            TxtGrid(0).MaxLength = 3
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Col_Complain
            TxtGrid(0).MaxLength = 100
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Else
            TxtGrid(0).MaxLength = 0
    End Select
End Sub

