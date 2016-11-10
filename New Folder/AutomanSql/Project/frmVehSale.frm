VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehSale 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Sale Bill"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13305
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
   ScaleHeight     =   11010
   ScaleWidth      =   13305
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   2445
      Left            =   30
      Negotiate       =   -1  'True
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   8235
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   4313
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
      Caption         =   "Chassis Help"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Chassis No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
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
         DataField       =   "God_Name"
         Caption         =   "Godown"
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
         DataField       =   "Model_Desc"
         Caption         =   "Model Desc"
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
         DataField       =   "Col_desc"
         Caption         =   "Colour"
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
         DataField       =   "PBill_No"
         Caption         =   "TelcoBillNo"
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
         DataField       =   "PBill_Date"
         Caption         =   "TelcoDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "PurVNo"
         Caption         =   "Purch No."
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
      BeginProperty Column09 
         DataField       =   "Pur_VDate"
         Caption         =   "PurchDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Rate"
         Caption         =   "Amount"
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
      BeginProperty Column11 
         DataField       =   "Al_Name"
         Caption         =   "Alloted"
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
            ColumnWidth     =   2415.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2039.811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "99,99,999.99"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Height          =   210
      Index           =   73
      Left            =   8535
      TabIndex        =   193
      Text            =   "99999999.99"
      Top             =   6840
      Width           =   1320
   End
   Begin VB.CheckBox ChkNewFinancer 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8295
      TabIndex        =   191
      Top             =   5610
      Width           =   195
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
      Height          =   210
      Index           =   72
      Left            =   8535
      TabIndex        =   58
      Top             =   5265
      Width           =   3195
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
      Height          =   210
      Index           =   71
      Left            =   8535
      TabIndex        =   57
      Top             =   5040
      Width           =   3195
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
      Height          =   210
      Index           =   70
      Left            =   4905
      TabIndex        =   186
      Top             =   4620
      Width           =   510
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
      Height          =   210
      Index           =   69
      Left            =   5475
      TabIndex        =   184
      Top             =   4605
      Width           =   1335
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
      Height          =   3000
      Left            =   540
      TabIndex        =   119
      Top             =   5220
      Visible         =   0   'False
      Width           =   8400
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
         Index           =   9
         Left            =   4455
         TabIndex        =   131
         Top             =   2280
         Width           =   2010
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
         Index           =   8
         Left            =   5580
         TabIndex        =   122
         Top             =   375
         Width           =   2325
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
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   123
         Text            =   "12-MAR-2003"
         Top             =   945
         Width           =   1305
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
         Index           =   6
         Left            =   4170
         MaxLength       =   15
         TabIndex        =   129
         Top             =   2010
         Width           =   2325
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
         Left            =   5955
         TabIndex        =   124
         Top             =   945
         Width           =   540
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
         Index           =   7
         Left            =   2490
         MaxLength       =   15
         TabIndex        =   130
         Top             =   2280
         Width           =   1305
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
         Index           =   3
         Left            =   4170
         MaxLength       =   150
         TabIndex        =   126
         Top             =   1215
         Width           =   2325
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
         Left            =   1950
         TabIndex        =   125
         Top             =   1215
         Width           =   480
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
         Index           =   5
         Left            =   2250
         TabIndex        =   128
         Top             =   2010
         Width           =   480
      End
      Begin VB.TextBox txtPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   1950
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   127
         Top             =   1485
         Width           =   4545
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmVehSale.frx":0000
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
         Left            =   6675
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Printer "
         Top             =   1950
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmVehSale.frx":030A
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
         Left            =   6675
         Style           =   1  'Graphical
         TabIndex        =   133
         ToolTipText     =   "Screen"
         Top             =   2280
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmVehSale.frx":0614
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
         Left            =   6675
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Printer "
         Top             =   2610
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
         Picture         =   "frmVehSale.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Screen"
         Top             =   2640
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
         Left            =   8085
         MousePointer    =   99  'Custom
         Picture         =   "frmVehSale.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Delete Current Record"
         Top             =   15
         Width           =   315
      End
      Begin VB.OptionButton OptPlain 
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
         Left            =   2145
         TabIndex        =   120
         Top             =   315
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Optpre 
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
         Left            =   2145
         TabIndex        =   121
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.C.No."
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
         Index           =   58
         Left            =   3840
         TabIndex        =   166
         Top             =   2295
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Option"
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
         Index           =   50
         Left            =   4440
         TabIndex        =   161
         Top             =   390
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Index           =   48
         Left            =   330
         TabIndex        =   160
         Top             =   960
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         Height          =   1695
         Left            =   225
         Top             =   900
         Width           =   6330
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
         Left            =   45
         TabIndex        =   146
         Top             =   15
         Width           =   8085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTO Name"
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
         Index           =   49
         Left            =   3075
         TabIndex        =   145
         Top             =   2025
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temp. Sale Certificate Y/N"
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
         Index           =   51
         Left            =   3735
         TabIndex        =   144
         Top             =   960
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Certificate Print Date"
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
         Index           =   52
         Left            =   315
         TabIndex        =   143
         Top             =   2295
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Body"
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
         Index           =   53
         Left            =   3105
         TabIndex        =   142
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seating Capacity"
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
         Index           =   54
         Left            =   330
         TabIndex        =   141
         Top             =   1230
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight In Printing Y/N"
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
         Index           =   55
         Left            =   330
         TabIndex        =   140
         Top             =   2025
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CertificateNarration"
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
         Index           =   56
         Left            =   345
         TabIndex        =   139
         Top             =   1500
         Width           =   1590
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
         Left            =   345
         TabIndex        =   138
         Top             =   2640
         Width           =   6315
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   41
         Left            =   420
         TabIndex        =   137
         Top             =   465
         Width           =   825
      End
      Begin VB.Line Line6 
         X1              =   1695
         X2              =   1395
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line8 
         X1              =   1710
         X2              =   1710
         Y1              =   720
         Y2              =   420
      End
      Begin VB.Line Line2 
         X1              =   1710
         X2              =   2040
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line3 
         X1              =   1710
         X2              =   2055
         Y1              =   735
         Y2              =   735
      End
   End
   Begin MSDataGridLib.DataGrid DgBodyBuilder 
      Height          =   2100
      Left            =   1725
      Negotiate       =   -1  'True
      TabIndex        =   183
      TabStop         =   0   'False
      Top             =   8580
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   3704
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
            ColumnWidth     =   2684.977
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
      Height          =   210
      Index           =   68
      Left            =   1185
      MaxLength       =   25
      TabIndex        =   27
      Top             =   2865
      Width           =   2205
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
      Height          =   210
      Index           =   67
      Left            =   5475
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   5955
      Width           =   1335
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
      Height          =   210
      Index           =   66
      Left            =   5475
      TabIndex        =   178
      Top             =   6180
      Width           =   1335
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
      Height          =   210
      Index           =   65
      Left            =   1815
      TabIndex        =   176
      Top             =   6435
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DgSubvention 
      Height          =   2175
      Left            =   -1635
      Negotiate       =   -1  'True
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   9210
      Visible         =   0   'False
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   3836
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "SchemeNo"
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
         DataField       =   "FromDate"
         Caption         =   "From Date"
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
         DataField       =   "ToDate"
         Caption         =   "To Date"
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
         DataField       =   "ModelGrp_Name"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "TotalSubvention"
         Caption         =   "Subvention"
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
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2204.788
         EndProperty
      EndProperty
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
      Height          =   210
      Index           =   64
      Left            =   5115
      TabIndex        =   26
      Top             =   2640
      Width           =   2295
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
      Index           =   63
      Left            =   5115
      MaxLength       =   8
      TabIndex        =   15
      Top             =   1515
      Width           =   2295
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
      Index           =   62
      Left            =   1815
      MaxLength       =   8
      TabIndex        =   171
      TabStop         =   0   'False
      Top             =   4635
      Width           =   1335
   End
   Begin VB.CommandButton CmdTransPost 
      Caption         =   "Post Trans."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10485
      TabIndex        =   169
      Top             =   15
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post A/c"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9255
      TabIndex        =   167
      Top             =   15
      Width           =   1245
   End
   Begin VB.CommandButton CancelBill 
      Caption         =   "Cancel Bill"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   15
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sale Vs Stock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6570
      TabIndex        =   168
      Top             =   15
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
      Height          =   210
      Index           =   61
      Left            =   1185
      MaxLength       =   20
      TabIndex        =   16
      Top             =   1740
      Width           =   2205
   End
   Begin MSDataGridLib.DataGrid DGCol 
      Height          =   3510
      Left            =   2595
      Negotiate       =   -1  'True
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   8910
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   6191
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
      Caption         =   "Color Help"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Col_Desc"
         Caption         =   "Color"
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
         Caption         =   "Col_Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
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
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
         EndProperty
      EndProperty
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   60
      Left            =   10260
      TabIndex        =   164
      Top             =   4335
      Visible         =   0   'False
      Width           =   1335
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
      Height          =   210
      Index           =   59
      Left            =   5475
      TabIndex        =   51
      Top             =   5730
      Width           =   1335
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
      Height          =   210
      Index           =   58
      Left            =   4905
      TabIndex        =   50
      Top             =   5730
      Width           =   510
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
      Index           =   56
      Left            =   10395
      MaxLength       =   8
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   1845
      Width           =   1335
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
      Index           =   57
      Left            =   10395
      MaxLength       =   12
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1335
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
      Height          =   210
      Index           =   54
      Left            =   1185
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   615
      Width           =   360
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
      Height          =   210
      Index           =   55
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   10
      Top             =   615
      Width           =   5850
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
      Height          =   210
      Index           =   53
      Left            =   1185
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   390
      Width           =   360
   End
   Begin MSDataGridLib.DataGrid DGFin 
      Height          =   3885
      Left            =   2745
      Negotiate       =   -1  'True
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   9210
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   6853
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
      Caption         =   "Financier Help"
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
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2865
      Left            =   -2025
      Negotiate       =   -1  'True
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   9240
      Visible         =   0   'False
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   5054
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
         DataField       =   "Chas_Type"
         Caption         =   "Chassis Type"
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
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6075.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1349.858
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
      Height          =   210
      Index           =   50
      Left            =   9210
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "01234"
      Top             =   1335
      Width           =   585
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   49
      Left            =   8535
      MaxLength       =   20
      TabIndex        =   63
      Text            =   "01234567890123456789"
      Top             =   6390
      Width           =   2265
   End
   Begin MSDataGridLib.DataGrid DGBook 
      Height          =   2175
      Left            =   1545
      Negotiate       =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   8445
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3836
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
      Caption         =   "Booking Help"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "code"
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
      BeginProperty Column01 
         DataField       =   "ord_date"
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
      BeginProperty Column02 
         DataField       =   "Name"
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
      BeginProperty Column03 
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3089.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2580.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2355.024
         EndProperty
      EndProperty
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
      Height          =   210
      Index           =   47
      Left            =   5475
      TabIndex        =   53
      Top             =   6630
      Width           =   1335
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   10170
      TabIndex        =   85
      Top             =   8805
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   120
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   105
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
            Name            =   "Verdana"
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
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   -660
      Negotiate       =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   9135
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   4935
      Left            =   -1530
      Negotiate       =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   8805
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      Caption         =   "Tax Form Help "
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
      Height          =   210
      Index           =   21
      Left            =   10395
      MaxLength       =   10
      TabIndex        =   28
      Text            =   "0123456789"
      Top             =   2295
      Width           =   1335
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
      Height          =   210
      Index           =   22
      Left            =   10395
      MaxLength       =   12
      TabIndex        =   29
      Top             =   2520
      Width           =   1335
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
      Height          =   210
      Index           =   20
      Left            =   5115
      TabIndex        =   23
      Top             =   2415
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   23
      Left            =   10395
      TabIndex        =   30
      Top             =   2745
      Width           =   1335
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
      Height          =   210
      Index           =   41
      Left            =   5475
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6405
      Width           =   1335
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   42
      Left            =   8535
      TabIndex        =   54
      Top             =   4365
      Width           =   1335
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
      Height          =   210
      Index           =   38
      Left            =   5475
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5055
      Width           =   1335
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
      Height          =   210
      Index           =   10
      Left            =   6930
      TabIndex        =   24
      Top             =   2415
      Width           =   480
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
      Height          =   210
      Index           =   19
      Left            =   8535
      MaxLength       =   10
      TabIndex        =   62
      Top             =   6165
      Width           =   1335
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
      Height          =   210
      Index           =   18
      Left            =   8535
      MaxLength       =   30
      TabIndex        =   64
      Text            =   "012345678901234"
      Top             =   6615
      Width           =   2265
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
      ForeColor       =   &H00FF00FF&
      Height          =   210
      Index           =   46
      Left            =   8535
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1335
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
      Height          =   210
      Index           =   44
      Left            =   8535
      TabIndex        =   56
      Top             =   4815
      Width           =   3195
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
      Height          =   210
      Index           =   43
      Left            =   8535
      TabIndex        =   59
      Top             =   5490
      Width           =   2265
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   45
      Left            =   8535
      TabIndex        =   55
      Top             =   4590
      Width           =   1335
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
      Height          =   210
      Index           =   17
      Left            =   1185
      MaxLength       =   25
      TabIndex        =   25
      Top             =   2640
      Width           =   2205
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
      Height          =   210
      Index           =   40
      Left            =   5475
      TabIndex        =   49
      Top             =   5505
      Width           =   1335
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
      Height          =   210
      Index           =   13
      Left            =   5115
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2295
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
      Height          =   210
      Index           =   14
      Left            =   1185
      MaxLength       =   15
      TabIndex        =   20
      Top             =   2190
      Width           =   2205
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
      Height          =   210
      Index           =   12
      Left            =   1185
      MaxLength       =   20
      TabIndex        =   18
      Top             =   1965
      Width           =   2205
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
      Height          =   210
      Index           =   11
      Left            =   5115
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1740
      Width           =   2295
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
      Height          =   210
      Index           =   4
      Left            =   9210
      TabIndex        =   6
      Top             =   1560
      Width           =   2490
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
      Height          =   210
      Index           =   9
      Left            =   1185
      MaxLength       =   40
      TabIndex        =   14
      Text            =   " "
      Top             =   1515
      Width           =   2205
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
      Height          =   210
      Index           =   8
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   13
      Text            =   " 0123456789012345678901234567890123456789"
      Top             =   1290
      Width           =   6225
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
      Height          =   210
      Index           =   7
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   12
      Text            =   " "
      Top             =   1065
      Width           =   6225
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
      Height          =   210
      Index           =   6
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   11
      Top             =   840
      Width           =   6225
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   1590
      Negotiate       =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   8910
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
      Caption         =   "Site Help"
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
      Height          =   210
      Index           =   28
      Left            =   1815
      TabIndex        =   37
      Top             =   5535
      Width           =   1335
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
      Height          =   210
      Index           =   33
      Left            =   4905
      TabIndex        =   42
      Top             =   4380
      Width           =   510
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
      Height          =   210
      Index           =   35
      Left            =   10050
      TabIndex        =   44
      Top             =   7095
      Visible         =   0   'False
      Width           =   510
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
      Height          =   210
      Index           =   32
      Left            =   1815
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6660
      Width           =   1335
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
      Height          =   210
      Index           =   36
      Left            =   10605
      TabIndex        =   45
      Top             =   7095
      Visible         =   0   'False
      Width           =   1335
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
      Height          =   210
      Index           =   37
      Left            =   5475
      TabIndex        =   46
      Top             =   4830
      Width           =   1335
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
      Height          =   210
      Index           =   1
      Left            =   9210
      MaxLength       =   20
      TabIndex        =   2
      Top             =   885
      Width           =   2505
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
      Height          =   210
      Index           =   27
      Left            =   1815
      TabIndex        =   36
      Top             =   5310
      Width           =   1335
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
      Height          =   210
      Index           =   26
      Left            =   1815
      TabIndex        =   35
      Top             =   5085
      Width           =   1335
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
      Height          =   210
      Index           =   34
      Left            =   5475
      TabIndex        =   43
      Top             =   4380
      Width           =   1335
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8D8FE&
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
      Left            =   5250
      TabIndex        =   31
      Top             =   3870
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSDataGridLib.DataGrid DGADItem 
      Height          =   4935
      Left            =   720
      Negotiate       =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   9045
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      Caption         =   "Addition Fitments Help"
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
      Height          =   210
      Index           =   16
      Left            =   5115
      MaxLength       =   25
      TabIndex        =   21
      Top             =   2190
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   210
      Index           =   24
      Left            =   1815
      TabIndex        =   33
      Top             =   4410
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "99,99,999.99"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Height          =   210
      Index           =   25
      Left            =   1815
      TabIndex        =   34
      Text            =   "99999999.99"
      Top             =   4860
      Width           =   1335
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
      Height          =   210
      Index           =   3
      Left            =   9825
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1335
      Width           =   1875
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13305
      _ExtentX        =   23469
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
      Height          =   210
      Index           =   5
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   8
      Top             =   390
      Width           =   5850
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
      Index           =   0
      Left            =   9195
      MaxLength       =   21
      TabIndex        =   1
      Top             =   420
      Width           =   2520
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
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   48
      Left            =   8535
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5940
      Width           =   1335
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
      Height          =   210
      Index           =   31
      Left            =   1815
      TabIndex        =   40
      Top             =   6210
      Width           =   1335
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
      Height          =   210
      Index           =   30
      Left            =   1815
      TabIndex        =   39
      Top             =   5985
      Width           =   1335
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
      Height          =   210
      Index           =   39
      Left            =   5475
      TabIndex        =   48
      Top             =   5280
      Width           =   1335
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
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   29
      Left            =   1815
      TabIndex        =   38
      Top             =   5760
      Width           =   1335
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
      Height          =   210
      Index           =   15
      Left            =   1185
      MaxLength       =   25
      TabIndex        =   22
      Text            =   "0123456789012345678901234"
      Top             =   2415
      Width           =   2205
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
      Height          =   210
      Index           =   2
      Left            =   9210
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1110
      Width           =   2490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1230
      Left            =   60
      TabIndex        =   32
      Top             =   3105
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   2170
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   11
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   16703741
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   8421631
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   "SrNo.|Add/Del Item |Type      |Qty|Rate  |Amount|Itemcode"
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
      _Band(0).Cols   =   11
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
      Height          =   210
      Index           =   52
      Left            =   8430
      Locked          =   -1  'True
      TabIndex        =   151
      TabStop         =   0   'False
      Text            =   "VFa"
      Top             =   2790
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
      Height          =   210
      Index           =   51
      Left            =   8430
      Locked          =   -1  'True
      TabIndex        =   152
      TabStop         =   0   'False
      Text            =   "0123456789"
      Top             =   2550
      Width           =   1275
   End
   Begin VB.Label LblSpDiscount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sp. Discount.................."
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
      Left            =   6885
      TabIndex        =   194
      Top             =   6855
      Width           =   2160
   End
   Begin VB.Label LblFinName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin Address..............Fin Address.............."
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
      Left            =   8940
      TabIndex        =   192
      Top             =   7920
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label LblFinGrp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin Address..............Fin Address.............."
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
      Left            =   6930
      TabIndex        =   190
      Top             =   5265
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier"
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
      Index           =   69
      Left            =   0
      TabIndex        =   189
      Top             =   15
      Width           =   765
   End
   Begin VB.Label LblFinancerGroup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin Address..............Fin Address.............."
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
      Left            =   6930
      TabIndex        =   188
      Top             =   5055
      Width           =   3660
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
      TabIndex        =   187
      Top             =   7170
      Width           =   45
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
      Index           =   67
      Left            =   3300
      TabIndex        =   185
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body Builder"
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
      Index           =   66
      Left            =   45
      TabIndex        =   182
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R.T.O. ..........................."
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
      Index           =   65
      Left            =   3300
      TabIndex        =   181
      Top             =   5970
      Width           =   2220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance......................."
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
      Index           =   64
      Left            =   3300
      TabIndex        =   180
      Top             =   6195
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Handling Charges...."
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
      Index           =   63
      Left            =   165
      TabIndex        =   177
      Top             =   6450
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subvention Scheme"
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
      Index           =   62
      Left            =   3405
      TabIndex        =   174
      Top             =   2655
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery From"
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
      Index           =   61
      Left            =   3405
      TabIndex        =   173
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subvention............"
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
      Index           =   60
      Left            =   165
      TabIndex        =   172
      Top             =   4635
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Desc*"
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
      Index           =   59
      Left            =   60
      TabIndex        =   170
      Top             =   1755
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Over Tax  @"
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
      Index           =   57
      Left            =   3300
      TabIndex        =   162
      Top             =   5745
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Index           =   47
      Left            =   9255
      TabIndex        =   159
      Top             =   1845
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
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
      Left            =   60
      TabIndex        =   156
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label LblAcPostBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Posting"
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
      Left            =   7425
      TabIndex        =   154
      Top             =   2565
      Width           =   960
   End
   Begin VB.Label LblAcPostDt 
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7425
      TabIndex        =   153
      Top             =   2805
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Axle No. ........"
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
      Left            =   6900
      TabIndex        =   149
      Top             =   6405
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Fuel Amount.............."
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
      Left            =   3300
      TabIndex        =   147
      Top             =   6645
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No*"
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
      Index           =   25
      Left            =   60
      TabIndex        =   118
      Top             =   1980
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
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
      TabIndex        =   117
      Top             =   2205
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Round Off......................."
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
      Index           =   46
      Left            =   3300
      TabIndex        =   114
      Top             =   6420
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Invoice Value..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   20
      Left            =   6930
      TabIndex        =   113
      Top             =   4380
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sob Total [B]...................."
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
      Index           =   19
      Left            =   3300
      TabIndex        =   112
      Top             =   5070
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt Y/N"
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
      Left            =   5790
      TabIndex        =   111
      Top             =   2430
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Book No.*"
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
      Index           =   45
      Left            =   6900
      TabIndex        =   110
      Top             =   6180
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. Information......"
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
      Left            =   6900
      TabIndex        =   109
      Top             =   6630
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Total Adv......."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Index           =   35
      Left            =   6930
      TabIndex        =   108
      Top             =   5730
      Width           =   1665
   End
   Begin VB.Label LblFinancer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financer"
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
      Left            =   6930
      TabIndex        =   107
      Top             =   4830
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source of Fund....."
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
      Left            =   6930
      TabIndex        =   106
      Top             =   5505
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financed Amount..."
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
      Left            =   6930
      TabIndex        =   105
      Top             =   4605
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RTO Office"
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
      Left            =   45
      TabIndex        =   104
      Top             =   2655
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax on Other Fitment......."
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
      Left            =   3300
      TabIndex        =   103
      Top             =   5520
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No"
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
      Left            =   3405
      TabIndex        =   102
      Top             =   1980
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model "
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
      Left            =   3405
      TabIndex        =   101
      Top             =   1755
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No.*"
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
      Index           =   23
      Left            =   8040
      TabIndex        =   100
      Top             =   1575
      Width           =   1140
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
      Index           =   21
      Left            =   60
      TabIndex        =   99
      Top             =   1485
      Width           =   345
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
      Index           =   14
      Left            =   60
      TabIndex        =   98
      Top             =   855
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mfg. Bill No."
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
      Left            =   9255
      TabIndex        =   97
      Top             =   2310
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
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
      Index           =   7
      Left            =   9780
      TabIndex        =   95
      Top             =   2505
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Octroi....................."
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
      Left            =   165
      TabIndex        =   94
      Top             =   5325
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NDP"
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
      Left            =   9825
      TabIndex        =   93
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SobTotal [A]............"
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
      Index           =   18
      Left            =   165
      TabIndex        =   92
      Top             =   6690
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surch. on Tax  @"
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
      Index           =   17
      Left            =   8445
      TabIndex        =   91
      Top             =   7110
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. Charges..................."
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
      Index           =   16
      Left            =   3300
      TabIndex        =   90
      Top             =   4845
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name*"
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
      Index           =   15
      Left            =   8040
      TabIndex        =   89
      Top             =   900
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incidental charges"
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
      Left            =   165
      TabIndex        =   88
      Top             =   5100
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax                 @"
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
      Left            =   3300
      TabIndex        =   87
      Top             =   4395
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type*"
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
      Left            =   45
      TabIndex        =   83
      Top             =   2430
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Ded/Short*"
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
      Left            =   3405
      TabIndex        =   82
      Top             =   2220
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Rate................"
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
      Left            =   165
      TabIndex        =   81
      Top             =   4425
      Width           =   1785
   End
   Begin VB.Label LblRebate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rebate.................."
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
      Left            =   165
      TabIndex        =   80
      Top             =   4875
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   1440
      Left            =   7935
      Top             =   375
      Width           =   3795
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
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
      Left            =   7215
      TabIndex        =   79
      Top             =   450
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill  No.*"
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
      Left            =   8040
      TabIndex        =   78
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division           "
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
      Left            =   8040
      TabIndex        =   77
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    "
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
      Left            =   10140
      TabIndex        =   76
      Top             =   660
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Y/N"
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
      Left            =   3405
      TabIndex        =   75
      Top             =   2430
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
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
      Left            =   8040
      TabIndex        =   74
      Top             =   435
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net OutStanding......"
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
      Index           =   41
      Left            =   6930
      TabIndex        =   72
      Top             =   5955
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation.........."
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
      Left            =   165
      TabIndex        =   71
      Top             =   6225
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MVT........................"
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
      Left            =   165
      TabIndex        =   70
      Top             =   6015
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Fitment Amount......."
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
      Left            =   3300
      TabIndex        =   69
      Top             =   5295
      Width           =   2310
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transit Insurance....."
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
      Left            =   165
      TabIndex        =   68
      Top             =   5775
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Temp. Registration"
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
      Index           =   22
      Left            =   165
      TabIndex        =   67
      Top             =   5550
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name*"
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
      TabIndex        =   66
      Top             =   405
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date*"
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
      Left            =   8040
      TabIndex        =   65
      Top             =   1125
      Width           =   825
   End
End
Attribute VB_Name = "frmVehSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mSP5 As String = "     "
Dim mInvPrefixHt As Integer
Dim RsChassis As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim RsMod As ADODB.Recordset
Dim RsSite As ADODB.Recordset
Dim rsFin As ADODB.Recordset
Dim RsCol As ADODB.Recordset
Dim RsBodyBuilder As ADODB.Recordset
Dim RsSubvention As ADODB.Recordset
Dim RsADItem As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim RSBook As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim GridKey As Integer
Dim DocID As String * 21
Public mVType As String
Dim VoucherEditFlag As Boolean
Dim FinAcCode As String
Dim vPrefix As String
Dim CancelBillY_N As Boolean
Dim mDealerContribution As Double
Dim mTataContribution As Double
Dim mPartyLstNo$

'Grid color scheme
'Private Const CellBackColLeave As String = &HBAD3C9
'Private Const CellForeColLeave As String = &H0&
'Private Const CellBackColEnter As String = &HC0E0FF
'Private Const GridBackColorBkg As String = &HCAECF0
Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Dim SubTot, SubTott As Double



Private Const TxtDocID      As Byte = 0
Private Const SiteCode      As Byte = 1
Private Const VDate         As Byte = 2
Private Const SerialNo      As Byte = 3
Private Const BookNo        As Byte = 4
Private Const Party         As Byte = 5
Private Const Add1          As Byte = 6
Private Const Add2          As Byte = 7
Private Const Add3          As Byte = 8
Private Const City          As Byte = 9
Private Const Govt_YN       As Byte = 10
Private Const Model         As Byte = 11
Private Const ChassisNo     As Byte = 12
Private Const EngineNo      As Byte = 13
Private Const Colours       As Byte = 14
Private Const FormType      As Byte = 15
Private Const ADType        As Byte = 16
Private Const RTO           As Byte = 17
Private Const SpclInfo      As Byte = 18
Private Const SrvBookNo     As Byte = 19
Private Const Taxable       As Byte = 20
Private Const TelcoInvNo    As Byte = 21
Private Const TelcoInvDate  As Byte = 22
Private Const NDP           As Byte = 23
Private Const SaleRate      As Byte = 60
Private Const Rebate        As Byte = 25
Private Const IncCharge     As Byte = 26
Private Const Octroi        As Byte = 27
Private Const TempReg       As Byte = 28
Private Const TransIns      As Byte = 29
Private Const MVT           As Byte = 30
Private Const Transportation As Byte = 31
Private Const SubTotA       As Byte = 32
Private Const TaxPer        As Byte = 33
Private Const TaxAmt        As Byte = 34
Private Const TaxSurPer     As Byte = 35
Private Const TaxSurch      As Byte = 36
Private Const MisCharge     As Byte = 37
Private Const SubTotB       As Byte = 38
Private Const OthFitAmt     As Byte = 39
Private Const OthFitTax     As Byte = 40
Private Const ROff          As Byte = 41
Private Const GTotAmt       As Byte = 42
Private Const FundSource    As Byte = 43
Private Const FB_Code       As Byte = 44
Private Const FinAmt        As Byte = 45
Private Const AdvAmt        As Byte = 46
Private Const FuelAmt       As Byte = 47
Private Const NetOStng      As Byte = 48
Private Const TransAxlNo    As Byte = 49
Private Const InvPrefix     As Byte = 50    'Invoice Prefix used in DocID 12-04-03
Private Const AcPostByName  As Byte = 51
Private Const AcPostDate    As Byte = 52
Private Const NamePrefix    As Byte = 53
Private Const FNamePrefix   As Byte = 54
Private Const fname         As Byte = 55
Private Const DelChNo       As Byte = 56
Private Const DelChDate     As Byte = 57
Private Const TOTPer        As Byte = 58
Private Const TOTAmt        As Byte = 59
Private Const SubAmt        As Byte = 24
Private Const ModelDesc     As Byte = 61
Private Const Subvention    As Byte = 62
Private Const DeliveryFrom  As Byte = 63
Private Const SubventionScheme  As Byte = 64
Private Const HandlingCharges   As Byte = 65
Private Const Insurance     As Byte = 66
Private Const RTOfee        As Byte = 67
Private Const SatAmt        As Byte = 69
Private Const SatPer        As Byte = 70
Private Const FinAdd1       As Byte = 71
Private Const FinAdd2       As Byte = 72
Private Const SpecialDiscount       As Byte = 73


Private Const ADItem As Byte = 1
Private Const Qty As Byte = 2
Private Const Rate As Byte = 3
Private Const Amt As Byte = 4
Private Const TaxPer1 As Byte = 5
Private Const TaxAmt1 As Byte = 6
Private Const TaxSurPer1 As Byte = 7
Private Const TaxSurAmt1 As Byte = 8
Private Const FinalAmt As Byte = 9
Private Const ADItemCode  As Byte = 10

Private Const TempInvDate As Byte = 0
Private Const CertiTempYN As Byte = 1
Private Const Seet As Byte = 2
Private Const Body As Byte = 3
Private Const Narr As Byte = 4
Private Const WtPrn As Byte = 5
Private Const RTOName As Byte = 6
Private Const CertiPrnDate As Byte = 7
Private Const DocType As Byte = 8

Private Const DocInv As Byte = 0
Private Const DocSaleCert As Byte = 1
Private Const DocForm22 As Byte = 2
Private Const DocForm22A As Byte = 3

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName$, mRepNameCert$, mRepName22$, mRepName22A$, mOldChasis$

Private Sub CancelBill_Click()
    Dim I As Integer, Rst As ADODB.Recordset
    CancelBillY_N = True
    If TopCtrl1.TopText2 <> "Browse" Then
        MsgBox "Cancellation Denied in this mode !", vbInformation
        CancelBillY_N = False
        Exit Sub
    End If
          
    'OfftakeIncentiveSrlNo
    'SubventionSrlNo
    GSQL = "Select OfftakeIncentiveSrlNo,SubventionSrlNo from veh_stock where Sal_DocId='" & txt(TxtDocID) & "' and ChassisNo  = '" & txt(ChassisNo).TEXT & "'"
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        If Rst!OfftakeIncentiveSrlNo <> "" Or Rst!SubventionSrlNo <> "" Then
            MsgBox "Offtake Incentive Claim / Subvention Letter made." & vbCrLf & "Deletion denied!", vbCritical, "Deletion Denied"
            Set Rst = Nothing
            Exit Sub
        End If
    End If
    Set Rst = Nothing
    If GCn.Execute("Select DelCh_DocId from  veh_order where Inv_DocId = '" & Master!SearchCode & "'").Fields(0).Value <> "" Then
        MsgBox "Delivery has been made against this Invoice", vbInformation, "Deletion Denied": Set Rst = Nothing: Exit Sub
    End If
    
    
If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub
    If MsgBox(" Are You Sure to Cancel the Invoice ? ", vbInformation + vbYesNo, "Cancel Bill Message") = vbYes Then
        Dim mCancelDate As String
        mCancelDate = InputBox("Please Enter Cancellation Date.", "Date of Cancellation", txt(VDate))
        If mCancelDate = "" Then
            Exit Sub
        Else
            mCancelDate = RetDate(mCancelDate)
        End If
        
        
        GCn.BeginTrans
        GCnFaV.BeginTrans

        
        Dim mFields$

        mFields = "OrdDocId,OrdDocIDHelp,Ord_SiteCode,Ord_VType,Ord_No,Ord_Date,Quot_SiteCode,Quot_DocId,QuotSrl_No,PartyCode, " & _
    "AREA,REF_CODE,REP_CODE,Profession,  PURPOSE,MODEL,EXP_DATE,QTY,RATE,INTD_USE,Permit_N_Z,GOVT_YN,Fund_Source, " & _
    "AddVeri_YN,PermitReq_YN,FIN_YN,FB_CODE,FIN_AcCode,FIN_AMT,P_AMOUNT,P_DATE,Other_Facilities,Book_UName, " & _
    "Book_UEntDt,Book_UAE,Chas_Type,Chassis,VehSerialNo,Srv_BookNo,Colour_Code,DelCh_DocId,DelCh_DocIDHelp, " & _
    "DelCh_SiteCode,DelCh_VType,DelCh_No,DelCh_DT,DelChPrn_YN,DelCh_UName,DelCh_UEntDt,DelCh_UAE,Inv_DocId, " & _
    "Inv_DocIDHelp,Inv_SiteCode,Inv_VType,Inv_No,Inv_Date,Form_Code,FormNo,FormIssRecDate,TrnType_Prn,VRATE, " & _
    "MARGINE,REBATE,InciChrg,Octroi,RegTemp,TransitInsu,Transport,MVT,TAX_Per,TAX_Amt,Surcharge_Per,Surcharge_Amt, " & _
    "OtherChrg,FIT_AMT,FIT_TAX,Round_off,DieselAmt,Interest_YN,RebDays,InterestPer,Interest,TDS_YN,TDS_Per,TDS_Amt, " & _
    "TDS_CNO,TDS_CDATE,TDS_IDATE,TDS_SIGN,TDS_DESIG,TDS_BankName,TDS_ChalNo,TDS_ChalDate,Certi,CertiPrn_YN,TCertiPrn_YN, " & _
    "BillPrn_YN,REG_FEE,REG_NO,DETAILS_YN,INS_FEE,INS_NOTE,S_CHARGE,STAMP_DUTY,RoundOff_YN,Net_Amount,RTO,REMARK, " & _
    "REMARK1,REMARK2,REMARK3,FirstVeh_YN,MISC_INFO,AdditionalSrv,TInv_YN,Inv_UName,Inv_UEntDt,Inv_UAE,Trf_Date, " & _
    "DelCh_AcPostByUName,DelCh_AcPostByUEntDt,Inv_Prefix,Inv_AcPostByUName,Inv_AcPostByUEntDt,TOT_Per, " & _
    "TOT_Amt,SubTot,RegBy,AdvEMI,AccCode,AccQty,OffTake,InsComm,FinPayOut,FinInc,EBTA,Retail, " & _
    "SPInc,Brokrage,Subvention,AmtRecd,SiebelOrderNo,SiebelInvoiceNo,SubventionScheme,DeliveryFrom, " & _
    "RTOFee,Insurance,DealerContribution,TataContribution,HandlingCharges,Book_AddBy,Book_AddDate, " & _
    "Book_ModifyBy,Book_ModifyDate,DelCh_AddBy,DelCh_AddDate,DelCh_ModifyBy,DelCh_ModifyDate,Inv_AddBy, " & _
    "Inv_AddDate , Inv_ModifyBy, Inv_ModifyDate, SAT_YN, SatPer, SatAmt"

        DocID = txt(TxtDocID)
        GCn.Execute ("Insert into Veh_Order1 (" & mFields & ") " & _
        "select " & mFields & " from Veh_Order where Inv_DocId='" & DocID & "'")
        GCn.Execute ("Update Veh_Order1 set Inv_UEntDt=" & ConvertDate(date) & " where Inv_DocId='" & DocID & "'")
        
        GCn.CommitTrans
        GCnFaV.CommitTrans
'*******START POSTING
        Dim MsgStr$, rsCtrlAc As ADODB.Recordset, RsTemp As ADODB.Recordset, mPostFinAmt As Byte
        Dim mGTotAmt As Double, mTOT_Ac_Code$, mCommNarr$
        
        Set rsCtrlAc = New ADODB.Recordset
        rsCtrlAc.CursorLocation = adUseClient
        rsCtrlAc.Open "Select Fitment_Ac,SpecialDiscountAc,Fuel_Ac,VehROff_Ac From AcControls", GCnFaV, adOpenStatic, adLockReadOnly
        If rsCtrlAc.RecordCount <= 0 Then
            MsgStr = "Please Add Records in A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        If IsNull(rsCtrlAc!Fitment_Ac) Or rsCtrlAc!Fitment_Ac = "" Or _
            IsNull(rsCtrlAc!Fuel_Ac) Or rsCtrlAc!Fuel_Ac = "" Or _
            IsNull(rsCtrlAc!VehROff_Ac) Or rsCtrlAc!VehROff_Ac = "" Then
            MsgStr = "Please define Fitment,Fuel and Round Off A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        rsForm.MoveFirst        'Vehicle Sale A/c Code, Tax A/c Code, Surcharge A/c Code
        rsForm.FIND "Name ='" & txt(FormType) & "'"
        If IsNull(rsForm!PurSal_Ac_Code) Or rsForm!PurSal_Ac_Code = "" Then
            MsgStr = "Please Define Sale A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        'Tax A/c Code Checking
        If Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)) <> 0 Then
            If IsNull(rsForm!Tax_Ac_Code) Or rsForm!Sur_Ac_Code = "" Then
                MsgStr = "Please Define Tax A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
                'CancelBillLedgerPost = False
                GoTo lblExit
            End If
        End If
        'Tax A/c Code Checking
        If Val(txt(SatAmt)) <> 0 Then
            If IsNull(rsForm!AddTaxAc) Or rsForm!AddTaxAc = "" Then
                MsgStr = "Please Define Additional Tax A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
                'CancelBillLedgerPost = False
                GoTo lblExit
            End If
        End If
        
        
        'Financier A/c Checking
        mTOT_Ac_Code = G_FaCn.Execute("select " & xIsNull("totax_ac", "") & " as TOT_Ac from AcControls where Div_Code='" & PubDivCode & "'").Fields(0).Value
        If Val(txt(TOTAmt)) <> 0 And mTOT_Ac_Code = "" Then
            MsgStr = "Please define TOT A/c Code in Vehicle Controls" & vbCrLf & "A/c Posting Aborted !"
            'CancelBillLedgerPost = False
            GoTo lblExit
        End If
        mPostFinAmt = GCn.Execute("select " & vIsNull("PostFinAmt", "0") & " as PostFinAmt from Syctrl").Fields(0).Value
        If mPostFinAmt = 1 And Val(txt(FinAmt)) <> 0 Then
            If txt(FundSource) = "Hypothecation" Or txt(FundSource) = "Hire Purchase" Then
                Set RsTemp = New ADODB.Recordset
                RsTemp.CursorLocation = adUseClient
                RsTemp.Open "Select switch(Ac_YN='1','Y',Ac_YN<>'1','N') as ACYN,AcCode From ContractFinance where FinCode='" & txt(FB_Code).Tag & "' ", GCn, adOpenStatic, adLockReadOnly
                If RsTemp!AcYN = "Y" Then
                    If RsTemp!AcCode = "" Or IsNull(RsTemp!AcCode) Then
                        MsgStr = "Please define A/c Code in Financier Master" & vbCrLf & "A/c Posting Aborted !"
                        GoTo lblExit
                    End If
                End If
            End If
        End If
       ' If CheckCtrls Then 'Control setting found Ok
            'CancelBillLedgerPost = True: Exit Function
        'End If
        
        'A/c Posting related declarations
        Dim mBookDocID$
        Dim LedgAry(7) As LedgRec, mResult As Byte, mNarr$
        
        'Sale Party A/c
        GCn.BeginTrans
        GCnFaV.BeginTrans
        
        mBookDocID = GCn.Execute("select OrdDocId from Veh_Order where Inv_DocId='" & txt(TxtDocID) & "'").Fields(0).Value
        mNarr = "By Cancelled Sales Invoice No." & txt(InvPrefix) & txt(SerialNo) & " Dt. " & txt(VDate) & " Chassis " & txt(ChassisNo)
        mCommNarr = mNarr & "[Common]"
        I = 0
        LedgAry(I).SubCode = txt(Party).Tag
        mGTotAmt = Val(txt(GTotAmt))
        If mPostFinAmt = 0 Then
            mGTotAmt = Val(txt(GTotAmt)) + Val(txt(FinAmt))
        End If
        LedgAry(I).AmtCr = Round(Val(txt(GTotAmt)), 2)
        LedgAry(I).Narration = mNarr
        'Vehicle Sale A/c
        'Modi LPS 05.12.2003
        If Val(txt(SubTotA)) + Val(txt(MisCharge)) - Val(txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsForm!PurSal_Ac_Code
            LedgAry(I).AmtDr = Round(Val(txt(SubTotA)) + Val(txt(MisCharge)), 2)
            LedgAry(I).Narration = mNarr
        End If
        'eof Modi
        'Fitment Amount
        If Val(txt(OthFitAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fitment_Ac
            LedgAry(I).AmtDr = Round(Val(txt(OthFitAmt)), 2)
            LedgAry(I).Narration = mNarr & " Additional Fitments on Vehicle Sale Bill"
        End If
        'Tax Amt
        If Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)) <> 0 Then
            If rsForm!Tax_Ac_Code <> "" And rsForm!Sur_Ac_Code <> "" _
                 And rsForm!Tax_Ac_Code <> rsForm!Sur_Ac_Code Then
                If Val(txt(TaxAmt)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                    LedgAry(I).AmtDr = Round(Val(txt(TaxAmt)) + Val(txt(OthFitTax)), 2)
                    LedgAry(I).Narration = mNarr & " Sale Tax"
                End If
                If Val(txt(TaxSurch)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Sur_Ac_Code
                    LedgAry(I).AmtDr = Round(Val(txt(TaxSurch)), 2)
                    LedgAry(I).Narration = mNarr & " Surcharge on Sales Tax"
                End If
            Else
                I = I + 1
                LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                LedgAry(I).AmtDr = Round(Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)), 2)
                LedgAry(I).Narration = mNarr & " Sales Tax & Surcharge"
            End If
        End If
        
        'Tax Amt
        If Val(txt(SatAmt)) <> 0 Then
            If XNull(rsForm!AddTaxAc) <> "" Then
                If Val(txt(SatAmt)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!AddTaxAc
                    LedgAry(I).AmtDr = Round(Val(txt(SatAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Additional Tax"
                End If
            End If
        End If
        
        If Val(txt(TOTAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = mTOT_Ac_Code
            LedgAry(I).AmtDr = Val(txt(TOTAmt))
            LedgAry(I).Narration = mNarr & " TOT Amt"
        End If
        
        If Val(txt(ROff)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!VehROff_Ac
            If Val(txt(ROff)) > 0 Then
                LedgAry(I).AmtDr = Round(Val(txt(ROff)), 2)
            Else
                LedgAry(I).AmtCr = Round(Abs(Val(txt(ROff))), 2)
            End If
            LedgAry(I).Narration = mNarr & " Round Off"
        End If
        
        
        If Val(txt(SpecialDiscount)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!SpecialDiscountAc
            LedgAry(I).ContraSub = txt(Party).Tag
            LedgAry(I).AmtCr = Round(Val(txt(SpecialDiscount)), 2)
            LedgAry(I).Narration = mNarr & " Special Discount on Vehicle Sale Bill"
            
            
            I = I + 1
            LedgAry(I).SubCode = txt(Party).Tag
            LedgAry(I).ContraSub = rsCtrlAc!SpecialDiscountAc
            LedgAry(I).AmtDr = Round(Val(txt(SpecialDiscount)), 2)
            LedgAry(I).Narration = mNarr & " Special Discount on Vehicle Sale Bill"
            
        End If
        
        
        
        'Fuel Amount
        If Val(txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fuel_Ac
            LedgAry(I).AmtCr = Round(Val(txt(FuelAmt)), 2)
            LedgAry(I).Narration = mNarr & " Fuel Amount"
        End If
        
        If mPostFinAmt = 1 And Val(txt(FinAmt)) <> 0 Then
            If txt(FundSource) = "Hypothecation" Or txt(FundSource) = "Hire Purchase" Then
                If RsTemp!AcCode = "" Or IsNull(RsTemp!AcCode) Then
                Else
                    I = I + 1
                    LedgAry(I).SubCode = RsTemp!AcCode
                    LedgAry(I).AmtCr = Round(Val(txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amt."
                    I = I + 1
                    LedgAry(I).SubCode = txt(Party).Tag
                    LedgAry(I).AmtDr = Round(Val(txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amount."
                End If
            End If
        End If
        DocID = left(txt(TxtDocID), 8) & "Cancl" & Right(txt(TxtDocID), 8)
        mResult = LedgerPost("C", LedgAry, GCnFaV, DocID, CDate(mCancelDate), mCommNarr)
        If mResult <> 1 Then
            MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
           ' ProcAcPost = False
        Else
            'ProcAcPost = True
        End If
        GCn.CommitTrans
        GCnFaV.CommitTrans
lblExit:
If MsgStr <> "" Then
    MsgBox MsgStr, vbCritical, "A/c Posting"
ElseIf err.NUMBER > 0 Then
    MsgBox err.Description, vbCritical, "A/c Posting"
End If
Set rsCtrlAc = Nothing
Set RsTemp = Nothing
If MsgBox("Print Cancelled bill Copy ?", vbInformation + vbYesNo, "Print Information") = vbYes Then
    Call CmdPrint_Click(5)
End If
'****************END POSTING
        'Unposting of Ledger completed
        ' creation stopped by LP Singh at Udaipur 15-04-2003
        'GCn.Execute ("update hiscard set Dealer_Code='', CouponNo='', TransAxelNo='', Supplier_BillNo='', Supplier_BillDate=Null, Name='" & PubComp_Name & "',Add1='',Add2='',Add3='',CityCode='',Govt_YN =  0 " & _
            " where Chassis ='" & txt(ChassisNo) & "'")
        GCn.BeginTrans
        GCnFaV.BeginTrans
        
            GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_VDate=Null, Srv_BookNo = '', TransAxlNo='' " & _
                "where ChassisNo  = '" & txt(ChassisNo) & "' and Sal_DocId = '" & txt(TxtDocID) & "'"
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, ADItem) <> "" Then
                    GCn.Execute ("delete from veh_purch2 where DocId='" & txt(TxtDocID) & "'")
                End If
            Next
            GCn.Execute ("update veh_order  " & _
                "set Inv_DocId='',Inv_DocIDHelp='' ,Inv_SiteCode='',Inv_VType='',Inv_No=Null ,Inv_Date=null,Form_Code='', " & _
                "TAX_Per=0,TAX_Amt=0,Surcharge_Per=0,Surcharge_Amt=0,MARGINE=0,VRATE=0,REBATE=0, Subvention=0, " & _
                "InciChrg=0,Octroi=0,RegTemp=0,TransitInsu=0,Transport=0,MVT=0,OtherChrg=0,FIT_AMT=0,FIT_TAX=0, " & _
                "DieselAmt=0,MISC_INFO='',RTO='',Round_off=0, " & _
                "FB_Code='' , FIN_AcCode='', FIN_AMT=0, " & _
                "TrnType_Prn=0,Fund_Source=0,Chassis='' , Srv_BookNo='', " & _
                "Inv_UName='', Inv_UEntDt=null, Inv_UAE= '',Inv_AcPostByUName='',Inv_AcPostByUEntDt=Null " & _
                "where Inv_DocId='" & txt(TxtDocID) & "'")
        GCnFaV.CommitTrans
        GCn.CommitTrans
        Master.Requery
        RSBook.Requery
        Call MoveRec
        BUTTONS True, Me, Master, 0
        CancelBillY_N = False
End If
Exit Sub

End Sub

Private Sub ChkNewFinancer_Click()
 If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
    If TopCtrl1.TopText2 = "Add" Or TopCtrl1.TopText2 = "Edit" Then
        If ChkNewFinancer.Value = 1 Then
            txt(FinAdd1).Enabled = True
            txt(FinAdd2).Enabled = True
            txt(FinAdd1).BackColor = CtrlBColOrg
            txt(FinAdd2).BackColor = CtrlBColOrg
        Else
            txt(FinAdd1).Enabled = False
            txt(FinAdd2).Enabled = False
            txt(FinAdd1).BackColor = CtrlBColDisabled
            txt(FinAdd2).BackColor = CtrlBColDisabled
        End If
    End If
  Else
    If TopCtrl1.TopText2 = "Add" Then
        If ChkNewFinancer.Value = 1 Then
            txt(FinAdd1).Enabled = True
            txt(FinAdd2).Enabled = True
            txt(FinAdd1).BackColor = CtrlBColOrg
            txt(FinAdd2).BackColor = CtrlBColOrg
        Else
            txt(FinAdd1).Enabled = False
            txt(FinAdd2).Enabled = False
            txt(FinAdd1).BackColor = CtrlBColDisabled
            txt(FinAdd2).BackColor = CtrlBColDisabled
        End If
    End If
End If
End Sub

Private Sub cmdPost_Click()
Dim I As Integer, mStartdate As String, mEndDate As String
    mStartdate = InputBox("Posting Required from which Date ?", "Start Date for Posting", PubLoginDate)
    mEndDate = InputBox("Posting Required upto which Date ?", "Last Date for Posting", PubLoginDate)
    
    If mStartdate = "" Or mEndDate = "" Then Exit Sub
    mStartdate = MakeDate(mStartdate)
    mEndDate = MakeDate(mEndDate)
    Master.MoveFirst
    Do Until Master.EOF
        If IsNull(Master!Inv_Date) Then GoTo MyNextRecord
        If Master!Inv_Date < CDate(mStartdate) Then GoTo MyNextRecord
        If Master!Inv_Date > CDate(mEndDate) Then GoTo MyNextRecord
        
        
        Call MoveRec
        
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        
        If AcPostAuthorisation(txt(AcPostByName)) = False Then GoTo MyNextRecord
        Disp_Text SETS("EDIT", Me, Master)
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            If CDate(txt(VDate).TEXT) >= PubStartDate And CDate(txt(VDate).TEXT) <= PubEndDate Then
                ProcAcPost
            End If
        End If
        Disp_Text SETS("INI", Me, Master)
        'Call MoveRec
MyNextRecord:
        Master.MoveNext
    Loop
End Sub

Private Sub CmdTransPost_Click()
Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        Disp_Text SETS("EDIT", Me, Master)
        txt(VDate).SetFocus
        FGrid.AddItem FGrid.Rows
        TopCtrl1_eSave
        FrmPrn.Visible = False
        Me.Refresh
MyNextRecord:
        Master.MoveNext
    Loop
End Sub

Private Sub Command1_Click()
Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
Dim MyStr As String

Set rs1 = GCn.Execute("Select * From Veh_Order where Inv_Date>=" & ConvertDate(PubStartDate) & " order By Left(orddocid,1),Inv_Date,Inv_No")
Set rs2 = GCn.Execute("Select * From Veh_Stock")

If rs1.RecordCount > 0 Then rs1.MoveFirst
Do Until rs1.EOF
    If rs2.RecordCount > 0 Then rs2.MoveFirst
    rs2.FIND ("ChassisNo='" & rs1!Chassis & "'")
    If IsNull(rs2!Sal_Docid) Or rs2!Sal_Docid = "" Then
        MyStr = MyStr & vbCrLf & left(rs1!OrdDocId, 1) & " " & rs1!Inv_No & " " & rs1!Inv_Date & " " & rs1!Chassis
    End If
    rs1.MoveNext
Loop
MsgBox MyStr, vbCritical, "Chassis Billed but not Updated in Stock File"
Set rs1 = Nothing
Set rs2 = Nothing
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub DGBook_Click()
    If RSBook.RecordCount > 0 Then
        txt(BookNo).TEXT = RSBook!Code
        txt(BookNo).Tag = RSBook!OrdDocId
        FillRecords
    End If
    txt(BookNo).SetFocus
    DGBook.Visible = False
End Sub

Private Sub DGFin_Click()
    If rsFin.RecordCount > 0 Then
        txt(FB_Code).TEXT = rsFin!Name
        txt(FB_Code).Tag = rsFin!Code
        FinAcCode = rsFin!Code
    End If
    txt(FB_Code).SetFocus
    DGFin.Visible = False
End Sub
Private Sub DGCol_Click()
    If RsCol.RecordCount > 0 Then
        txt(Colours).TEXT = RsCol!Col_Desc
        txt(Colours).Tag = RsCol!Col_Code
    End If
    txt(Colours).SetFocus
    DGCol.Visible = False
End Sub
Private Sub DGSite_Click()
    If RsSite.RecordCount > 0 Then
        txt(SiteCode).TEXT = RsSite!Name
        txt(SiteCode).Tag = RsSite!Code
    End If
    txt(SiteCode).SetFocus
    DGSite.Visible = False
End Sub

Private Sub DgSubvention_Click()
    If RsSubvention.RecordCount > 0 Then
        txt(SubventionScheme).TEXT = RsSubvention!Name
        txt(SubventionScheme).Tag = RsSubvention!Code
    End If
    txt(SubventionScheme).SetFocus
    DgSubvention.Visible = False
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
If pubUName <> "SA" And pubUName <> "SANJAY1" Then
    CancelBill.Visible = False
    cmdPost.Visible = False
End If
Else
If pubUName <> "SA" Then
    CancelBill.Visible = False
End If
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
    TopCtrl1.Tag = PubUParam: WinSetting Me:     Ini_Grid
    
    If PubVATYN = 1 Then
        Label3(12) = "V A T @"
    End If
    Label3(57) = pubTOTCaption
    mVType = "V_SB"
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
     Dim sitecond As String
     sitecond = " And Inv_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = sitecond & " and " & cMID("inv_DocId", "3", "1") & "='" & PubSiteCode & "'"
     End If


    If PubMoveRecYn Then
        Master.Open "select Inv_DocId as SearchCode,Veh_Order.* from Veh_Order where left(Inv_DocID,1)='" & PubDivCode & "' and " & cTrim(cMID("Inv_DocID", "4", "5")) & "= '" & mVType & "' " & sitecond & " Order by Inv_Date desc,Inv_No desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 Inv_DocId as SearchCode,Veh_Order.* from Veh_Order where left(Inv_DocID,1)='" & PubDivCode & "' and " & cTrim(cMID("Inv_DocID", "4", "5")) & "= '" & mVType & "' " & sitecond & " Order by Inv_Date desc,Inv_No desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
        
    
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
    Set RSBook = New ADODB.Recordset
    RSBook.CursorLocation = adUseClient
    RSBook.Open "Select distinct OrdDocID, " & cTrim(CStr("veh_order.Ord_No")) & " as code,veh_order.ord_date,subgroup.Name,SubGroup.Add1, SubGroup.LstNo,City.CityName, Right(Ord_SiteCode,1) as Site_Code from ((Veh_Order left join subgroup on subgroup.subcode = veh_order.partycode)Left Join City on City.CityCode=SubGroup.CityCode) where left(veh_order.OrdDocId,1)= '" & PubDivCode & "' and (veh_order.Inv_DocId='' Or veh_order.Inv_DocId Is Null) Order by  " & cTrim(CStr("veh_order.Ord_No")) & " ", GCn, adOpenDynamic, adLockOptimistic
    Set DGBook.DataSource = RSBook
    
    Set RsBodyBuilder = GCn.Execute("Select BodyBuilderCode As Code, BodyBuilderDesc As Name From BodyBuilder Order By BodyBuilderDesc ")
    Set DgBodyBuilder.DataSource = RsBodyBuilder
        
    Set rsFin = New ADODB.Recordset
    rsFin.CursorLocation = adUseClient
    rsFin.Open "select fincode as code,finname + ',' +  " & xIsNull("add1", "") & " +  ',' +  " & xIsNull("add2", "") & " + " & xIsNull("City.CityName", "") & " as name, Add1, Add2,AcCode,FinBankCode, UnderFinGrp, FinName from ContractFinance " & _
    "left join city on left(ContractFinance.City,4)=City.CityCode where fincatg = 0  order by finname", GCn, adOpenDynamic, adLockOptimistic
    Set DGFin.DataSource = rsFin
    
    Set RsCol = New ADODB.Recordset
    RsCol.CursorLocation = adUseClient
    RsCol.Open "select Col_Code,Col_desc from ColMast", GCn, adOpenDynamic, adLockOptimistic
    Set DGCol.DataSource = RsCol
  
    Set RsMod = New ADODB.Recordset
    RsMod.CursorLocation = adUseClient
    RsMod.Open "select Model as code,Model_Desc as NAME, Chas_Type from Model where (Div_Code='" & PubDivCode & "' or Div_Code='') order by Model", GCn, adOpenDynamic, adLockOptimistic
    Set DGMod.DataSource = RsMod
    
    
    Set RsSubvention = GCn.Execute("Select SchemeNo As Code, SchemeNo As Name, FromDate, ToDate, MG.ModelGrp_Name, Model, " & _
                                   "DealerContribution, TataContribution, TotalSubvention " & _
                                   "From Subvention S " & _
                                   "Left Join Model_Grp MG On S.ModelGroup=MG.ModelGrp_Code " & _
                                   "Order By SchemeNo")
    Set DgSubvention.DataSource = RsSubvention
    
    Set rsForm = New ADODB.Recordset
    With rsForm
        .CursorLocation = adUseClient
        .Open "SELECT T.Form_Code as Code,T.Form_Desc as Name,T.Tax_Sur_Per, T.AddTaxPer, T.Tax_Per,T1.Tax_Ac_Code,T1.Sur_Ac_Code,T1.PurSal_Ac_Code, T1.AddTaxAc " & _
            "FROM TaxForms as T left join TaxFormsAc as T1 on  T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code " & _
            "where T.Vehicle_YN = 1 and T.Trn_Type = 'Sale' Order by Form_Desc ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGForm.DataSource = rsForm
    
    Set RsADItem = New ADODB.Recordset
    With RsADItem
        .CursorLocation = adUseClient
        .Open "SELECT  Prod_Code as code,Prod_name as name,Rate  FROM veh_amdModel order by  veh_amdModel.Prod_name ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGADItem.DataSource = RsADItem
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
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
Set RsMod = Nothing
Set RsADItem = Nothing
Set RsSite = Nothing
Set rsForm = Nothing
Set RsVno = Nothing
Set RsChassis = Nothing
Set rsFin = Nothing
Set Master = Nothing
Set mListItem = Nothing
End Sub



Private Sub ListView_Click()
If FrmPrn.Visible = False Then
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txt(Val(ListView.Tag)).SetFocus
Else
    txtPrint(DocType).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txtPrint(DocType).SetFocus
End If
End Sub

Private Sub OptPlain_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub Optpre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        txt(SiteCode).Tag = PubSiteCode
        txt(SiteCode) = PubSiteName
        txt(VDate).SetFocus
    Else
        txt(SiteCode).Tag = PubSiteCode
        txt(SiteCode) = PubSiteName

        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
            txt(VDate).SetFocus
        Else
            txt(SiteCode).SetFocus
        End If
    End If
    
     txt(TOTPer) = MainLib.TOTCal()
Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer, Rst As ADODB.Recordset
Dim LedgAry(1) As LedgRec, mResult As Byte
    'OfftakeIncentiveSrlNo
    'SubventionSrlNo
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    
    GSQL = "Select OfftakeIncentiveSrlNo,SubventionSrlNo from veh_stock where Sal_DocId='" & txt(TxtDocID) & "' and ChassisNo  = '" & txt(ChassisNo).TEXT & "'"
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        If Rst!OfftakeIncentiveSrlNo <> "" Or Rst!SubventionSrlNo <> "" Then
            MsgBox "Offtake Incentive Claim / Subvention Letter made." & vbCrLf & "Deletion denied!", vbCritical, "Deletion Denied"
            Set Rst = Nothing
            Exit Sub
        End If
    End If
    Set Rst = Nothing
    If GCn.Execute("Select DelCh_DocId from  veh_order where Inv_DocId = '" & Master!SearchCode & "'").Fields(0).Value <> "" Then
        MsgBox "Delivery has been made against this Invoice", vbInformation, "Deletion Denied": Set Rst = Nothing: Exit Sub
    End If
    
If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub

If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    GCn.BeginTrans
    GCnFaV.BeginTrans
    
    CreateLog Me, Master!SearchCode, False
    'Unpost Ledger a/c
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, txt(TxtDocID))
    If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    'Unposting of Ledger completed
'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
'    GCn.Execute ("update hiscard set Dealer_Code='', CouponNo='', TransAxelNo='', Supplier_BillNo='', Supplier_BillDate=Null, Name='" & PubComp_Name & "',Add1='',Add2='',Add3='',CityCode='',Govt_YN =  0 " & _
        " where Chassis ='" & txt(ChassisNo) & "'")
    GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_DocIDHelp='',Sal_Site_Code='',Sal_VType='',Sal_VNo=null,FIN_AcCode='',Sal_Rate=0,Sal_VDate=Null,Ord_SiteCode='',Ord_DocId='', Srv_BookNo = '', TransAxlNo='' " & _
        "where ChassisNo  = '" & txt(ChassisNo) & "' and Sal_DocId = '" & txt(TxtDocID) & "'"
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ADItem) <> "" Then
            GCn.Execute ("delete from veh_purch2 where DocId='" & txt(TxtDocID) & "'")
        End If
    Next
    GCn.Execute ("update veh_order  " & _
        "set Inv_DocId='',Inv_DocIDHelp='' ,Inv_SiteCode='',Inv_VType='',Inv_No=Null ,Inv_Date=null,Form_Code='', " & _
        "TAX_Per=0,TAX_Amt=0,Surcharge_Per=0,Surcharge_Amt=0,MARGINE=0,VRATE=0,REBATE=0, Subvention=0, " & _
        "InciChrg=0,Octroi=0,RegTemp=0,TransitInsu=0,Transport=0,MVT=0,OtherChrg=0,FIT_AMT=0,FIT_TAX=0, " & _
        "DieselAmt=0,MISC_INFO='',RTO='',Round_off=0, " & _
        "FB_Code='' , FIN_AcCode='', FIN_AMT=0, " & _
        "TrnType_Prn=0,Fund_Source=0,Chassis='' , Srv_BookNo='', " & _
        "Inv_UName='', Inv_UEntDt=null, Inv_UAE= '',Inv_AcPostByUName='',Inv_AcPostByUEntDt=Null,SiebelInvoiceNo='' " & _
        "where Inv_DocId='" & txt(TxtDocID) & "'")
    GCnFaV.CommitTrans
    GCn.CommitTrans
    Master.Requery
    RSBook.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If err.NUMBER <> 0 Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1

    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
        
        
    If txt(DelChNo) <> "" Then MsgBox "Delivery Made, Edit denied !", vbInformation, "Validation": Exit Sub
'eof lp
    If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    txt(VDate).SetFocus
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
txtPrint(DocType) = "Sale Bill"
txtPrint(CertiTempYN) = ""
txtPrint(TempInvDate) = txt(VDate)
txtPrint(CertiPrnDate) = txt(VDate)
txtPrint(RTOName) = txt(RTO)
'txtPrint(Seet) = GCn.Execute("Select SEAT from Model where Model='" & Txt(Model) & "'").Fields(0).Value
'txtPrint(Body) = GCn.Execute("Select INTD_USE from Veh_Order where inv_Docid='" & Txt(TxtDocId) & "'").Fields(0).Value
'txtPrint(Narr) = "Signature of Manufacturer/Dealer or Officer of Defence Department"

mRepName = IIf(OptPlain.Value = True, "VehSale", "VehSale")
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
CmdPrint(PWindows).SetFocus
If PubSpeedPrint Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
End Sub

Private Sub TopCtrl1_eRef()
    RsMod.Requery
    RSBook.Requery
    RsSite.Requery
    rsForm.Requery
    RsADItem.Requery
'    RsChassis.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    Dim DocIdHlp$, sqlstr$, mDlrID$, mPBIllNo$, mPBIllDate$, CardNo$
    Dim mFundSource As Byte
    Dim mTrntypeprn As Byte, mQuotDocID$, mQuotDocIDSrlNo As Integer
    
'    On Error GoTo errlbl

If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
If TxtGrid(0).Visible = True Then
    If TxtGridLeave = False Then
        TxtGrid(0).SetFocus
        Exit Sub
    End If
End If
Grid_Hide

If IsValid(txt(SiteCode), "SiteCode") = False Then Exit Sub
If IsValid(txt(VDate), "Bill Date") = False Then Exit Sub
If IsValid(txt(SerialNo), "Bill Number") = False Then Exit Sub
If IsValid(txt(Party), "Party Name") = False Then Exit Sub
If IsValid(txt(BookNo), "Booking No.") = False Then Exit Sub
If IsValid(txt(FormType), "Form Type") = False Then Exit Sub
If IsValid(txt(Model), "Model") = False Then Exit Sub
If IsValid(txt(ChassisNo), "Chassis") = False Then Exit Sub

If Val(txt(FinAmt)) > 0 Then
    If IsValid(txt(FB_Code), "Financier") = False Then Exit Sub
    If IsValid(txt(FundSource), "Source of Fund") = False Then Exit Sub
    If txt(FundSource) <> "Hypothecation" And _
        txt(FundSource) <> "Hire Purchase" And _
        txt(FundSource) <> "Lease" And _
        txt(FundSource) <> "Agreement" And _
        txt(FundSource) <> "Lease & Agreement" And _
        txt(FundSource) <> "Loan Cum Hypt." Then
        
        MsgBox "Invalid Source of Fund !", vbCritical, "Fund Source Validation"
        txt(FundSource).SetFocus
        Exit Sub
    End If
Else
'    If txt(FundSource) <> "Own Fund" Then
'        MsgBox "Financed Amount is zero, Correct Source of Fund", vbCritical, "Fund Source Validation"
'        txt(FundSource).SetFocus
'        Exit Sub
'    End If
End If

GSQL = "Select Model,ChassisNo,Sal_DocID from Veh_Stock where ChassisNo<>'" & txt(ChassisNo) & "' and Srv_BookNo = '" & txt(SrvBookNo) & "'" ' and Model='" & Txt(Model) & "'"
sqlstr = "Select Model,ChassisNo,Sal_DocID from Veh_Stock where ChassisNo<>'" & txt(ChassisNo) & "' and TransAxlNo = '" & txt(TransAxlNo) & "'" ' and Model='" & Txt(Model) & "'"
'Service Book No. checking
Set Rst = New ADODB.Recordset
Rst.CursorLocation = adUseClient
Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
If Rst.RecordCount > 0 Then
'    MsgBox "Service Book No. " & Txt(SrvBookNo) & " is already allocated/issued for " & vbCrLf & "Model " & Rst!Model & " and Chassis No." & Rst!ChassisNo, vbCritical, "Duplicate Service Book No."
'    Txt(SrvBookNo).SetFocus
    Set Rst = Nothing
'    Exit Sub
End If
'TransAxlNo checking
If txt(TransAxlNo) <> "" Then
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open sqlstr, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        MsgBox "Trans Axle No. " & txt(TransAxlNo) & " is already allocated/issued for " & vbCrLf & "Model " & Rst!Model & " and Chassis No." & Rst!ChassisNo, vbCritical, "Duplicate TransAxle No."
        txt(TransAxlNo).SetFocus
        Set Rst = Nothing
        Exit Sub
    End If
End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ADItem) <> "" Then
            If Val(FGrid.TextMatrix(I, Qty)) = 0 Then MsgBox "Fill Quantity in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Qty: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
        End If
    Next
    Amt_Cal False
    Select Case txt(FundSource).TEXT
        Case "Hypothecation"
            mFundSource = 0
        Case "Hire Purchase"
            mFundSource = 1
        Case "Lease"
            mFundSource = 3
        Case "Agreement"
            mFundSource = 4
        Case "Lease & Agreement"
            mFundSource = 5
        Case "Loan Cum Hypt."
            mFundSource = 6
        Case Else
            mFundSource = 2 'Own Fund
    End Select
    Select Case txt(ADType).TEXT
        Case "No Detail"
            mTrntypeprn = 0
        Case "Name/Qty"
            mTrntypeprn = 1
        Case "Name/Qty/Amount"
            mTrntypeprn = 2
    End Select
    '********* cHECKING pOSTING cOTROLS
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        If ProcAcPost(True) = False Then Me.ActiveControl.SetFocus: Exit Sub
        txt(AcPostByName) = pubUName
        txt(AcPostDate) = PubServerDate
    End If
    '**********
    mDlrID = GCn.Execute("Select " & xIsNull("Dealer_ID", "") & " from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    Set Rst = New ADODB.Recordset
    Rst.Open "Select PBILL_NO,PBILL_DATE from Veh_Stock where ChassisNo = '" & txt(ChassisNo) & "'", GCn, adOpenStatic, adLockReadOnly
    mPBIllNo = IIf(IsNull(Rst!PBILL_NO), "", Rst!PBILL_NO)
    mPBIllDate = IIf(IsNull(Rst!PBILL_DATE), "", Rst!PBILL_DATE)
    Set Rst = Nothing
    
'    If TopCtrl1.TopText2.CAPTION = "Add" Then
'    '   lp 11-03-03
'        DocId = Txt(TxtDocId)
'        If GCn.Execute("select count(*) from veh_order where inv_DocID='" & Txt(TxtDocId) & "'").Fields(0)  > 0 Then
'            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
'                MsgBox "Bill No. already exists, Retry", vbCritical, "Validation Error"
'                Txt(SerialNo).SetFocus
'                Exit Sub
'            Else
'                Txt(TxtDocId) = GetDocIDVBill(GCnFaV, mVType, Txt(Vdate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
'                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocId, Document_No)) Then
'                    MsgBox "Bill No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
'                    Exit Sub
'                End If
'            End If
'        End If
'   End If
   DocIdHlp = Replace(txt(TxtDocID), " ", "")
    GCn.BeginTrans
    GCnFaV.BeginTrans
    mTrans = True
    
 If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
    If TopCtrl1.TopText2 = "Edit" Then
        DocID = txt(TxtDocID)
        
        
        Dim mFinBankCode1 As String
        If ChkNewFinancer.Value = 1 Then
            mFinBankCode1 = PubSiteCode & Format(VNull(GCn.Execute("Select (Max(" & cVal(cTrim(cMID("FinCode", "3", "4"))) & ") + 1) From ContractFinance Where " & cTrim(cMID("FinCode", "3", "4")) & "<>'' ").Fields(0).Value), "00000")
                          
            GSQL = "Insert Into ContractFinance (FinCode,Site_Code,FinCatg,FinBankCode,UnderFinGrp,FinName,Add1,Add2,AcCode,Ac_YN,U_Name,U_EntDt,U_AE) " & _
                " Values ('" & mFinBankCode1 & "','" & PubSiteCode & "', 0,'" & LblFinancer.Tag & "','" & LblFinancerGroup.Tag & "','" & LblFinName.Tag & "','" & txt(FinAdd1) & _
                "','" & txt(FinAdd2) & "','',0,'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            GCn.Execute GSQL
            txt(FB_Code).Tag = mFinBankCode1
         End If
 End If

End If

    If TopCtrl1.TopText2 = "Add" Then
        DocID = txt(TxtDocID)
        
        
        Dim mFinBankCode As String
        If ChkNewFinancer.Value = 1 Then
            mFinBankCode = PubSiteCode & Format(VNull(GCn.Execute("Select (Max(" & cVal(cTrim(cMID("FinCode", "3", "4"))) & ") + 1) From ContractFinance Where " & cTrim(cMID("FinCode", "3", "4")) & "<>'' ").Fields(0).Value), "00000")
                          
            GSQL = "Insert Into ContractFinance (FinCode,Site_Code,FinCatg,FinBankCode,UnderFinGrp,FinName,Add1,Add2,AcCode,Ac_YN,U_Name,U_EntDt,U_AE) " & _
                " Values ('" & mFinBankCode & "','" & PubSiteCode & "', 0,'" & LblFinancer.Tag & "','" & LblFinancerGroup.Tag & "','" & LblFinName.Tag & "','" & txt(FinAdd1) & _
                "','" & txt(FinAdd2) & "','',0,'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            GCn.Execute GSQL
            txt(FB_Code).Tag = mFinBankCode
        End If
        
        
'Stopped Because of Duplicating Bill No. if Site is Different
'        If GCn.Execute("select count(*) from veh_order where inv_DocID='" & Txt(TxtDocID) & "'").Fields(0) > 0 Then
'            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
'                MsgBox "Bill No. already exists, Retry", vbCritical, "Validation Error"
'                Txt(SerialNo).SetFocus
'                GoTo errlbl
'            Else
'                Txt(TxtDocID) = GetDocIDVBill(GCnFaV, mVtype, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix, Txt(SiteCode).Tag)
'                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
'                    MsgBox "Bill No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
'                    GoTo errlbl
'                End If
'            End If
'        End If
        If GCn.Execute("select count(*) from veh_order where Left(Inv_DocID,1)='" & PubDivCode & "' And Inv_VType = '" & mVType & "' AND " & cTrim(cMID("Inv_DocId", "9", "5")) & " ='" & LblVPrefix & "'  And Inv_No=" & Val(txt(SerialNo)) & "").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
                MsgBox "Bill No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                GoTo errlbl
            Else
                If mPartyLstNo <> "" Then
                    If PubVehTaxInvPrefix <> "" Then
                        LblVPrefix = PubVehTaxInvPrefix
                        txt(InvPrefix) = PubVehTaxInvPrefix
                    End If
                End If
            
                txt(TxtDocID) = GetDocIDVBill(GCnFaV, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix, txt(SiteCode).Tag)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    MsgBox "Bill No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If

        DocIdHlp = Replace(txt(TxtDocID), " ", "")
        '** eof
'        GCn.Execute ("update veh_order" & _
'            " set Inv_DocId='" & Txt(TxtDocID) & "',Inv_DocIDHelp='" & DocIdHlp & "' ,Inv_SiteCode='" & PubSiteCode & Txt(SiteCode).Tag & "',Inv_VType='" & mVType & "',Inv_No=" & Val(Txt(SerialNo).TEXT) & " ,Inv_Date=" & ConvertDate(Txt(Vdate)) & " ,Form_Code='" & Txt(FormType).Tag & _
'            "',TAX_Per=" & Val(Txt(TaxPer)) & ",model = '" & Txt(Model).TEXT & "', TAX_Amt=" & Val(Txt(TaxAmt)) & ",Surcharge_Per=" & Val(Txt(TaxSurPer)) & ",Surcharge_Amt=" & Val(Txt(TaxSurch)) & ",MARGINE=" & (Val(Txt(SaleRate)) - Val(Txt(NDP))) & ",VRATE=" & Val(Txt(NDP)) & ",REBATE=" & Val(Txt(Rebate)) & _
'            " ,InciChrg=" & Val(Txt(IncCharge)) & ",Octroi=" & Val(Txt(Octroi)) & " ,RegTemp=" & Val(Txt(TempReg)) & ",TransitInsu=" & Val(Txt(TransIns)) & ",Transport=" & Val(Txt(Transportation)) & ",MVT=" & Val(Txt(MVT)) & ",OtherChrg=" & Val(Txt(MisCharge)) & ",FIT_AMT=" & Val(Txt(OthFitAmt)) & ",FIT_TAX=" & Val(Txt(OthFitTax)) & _
'            " ,TOT_Per=" & Val(Txt(TOTPer)) & ",TOT_Amt=" & Val(Txt(TOTAmt)) & ",DieselAmt=" & Val(Txt(FuelAmt)) & ",MISC_INFO='" & Txt(SpclInfo) & "',RTO='" & Txt(RTO) & "',Round_off=" & Val(Txt(ROff)) & _
'            " ,FB_Code='" & Txt(FB_Code).Tag & "' , FIN_AcCode='" & FinAcCode & "', FIN_AMT=" & Val(Txt(FinAmt)) & _
'            " ,Net_Amount = " & Val(Txt(GTotAmt)) & ", TrnType_Prn=" & mTrntypeprn & ",Fund_Source=" & mFundSource & ",Chassis='" & Txt(ChassisNo) & _
'            "',Inv_UName='" & pubUName & "', Inv_UEntDt=#" & PubServerDate & "#, Inv_UAE= 'A' " & _
'            " ,Inv_AcPostByUName='" & Txt(AcPostByName) & "',Inv_AcPostByUEntDt=" & ConvertDate(Txt(AcPostDate)) & _
'            " ,SubTot=" & Val(Txt(SubAmt)) & ", Colour_Code='" & Txt(Colours).Tag & "',EngineNo='" & Txt(EngineNo) & "' where OrdDocId = '" & Txt(BookNo).Tag & "'")
            
        GCn.Execute ("update veh_order" & _
            " set Inv_DocId='" & txt(TxtDocID) & "',Inv_DocIDHelp='" & DocIdHlp & "' ,Inv_SiteCode='" & PubSiteCode & txt(SiteCode).Tag & "',Inv_VType='" & mVType & "',Inv_No=" & Val(txt(SerialNo).TEXT) & " ,Inv_Date=" & ConvertDate(txt(VDate)) & " ,Form_Code='" & txt(FormType).Tag & _
            "',TAX_Per=" & Val(txt(TaxPer)) & ",model = '" & txt(Model).TEXT & "', TAX_Amt=" & Val(txt(TaxAmt)) & ", SatPer = " & Val(txt(SatPer)) & ", SatAmt=" & Val(txt(SatAmt)) & ", Surcharge_Per=" & Val(txt(TaxSurPer)) & ",Surcharge_Amt=" & Val(txt(TaxSurch)) & ",MARGINE=" & (Val(txt(SaleRate)) - Val(txt(NDP))) & ",VRATE=" & Val(txt(NDP)) & ", Subvention=" & Val(txt(Subvention)) & ",REBATE=" & Val(txt(Rebate)) & _
            " ,InciChrg=" & Val(txt(IncCharge)) & ",Octroi=" & Val(txt(Octroi)) & " ,RegTemp=" & Val(txt(TempReg)) & ",TransitInsu=" & Val(txt(TransIns)) & ",Transport=" & Val(txt(Transportation)) & ", HandlingCharges=" & Val(txt(HandlingCharges)) & ",MVT=" & Val(txt(MVT)) & ",OtherChrg=" & Val(txt(MisCharge)) & ", RTOFee= " & Val(txt(RTOfee)) & ", Insurance = " & Val(txt(Insurance)) & ",FIT_AMT=" & Val(txt(OthFitAmt)) & ",FIT_TAX=" & Val(txt(OthFitTax)) & _
            " ,TOT_Per=" & Val(txt(TOTPer)) & ",TOT_Amt=" & Val(txt(TOTAmt)) & ",DieselAmt=" & Val(txt(FuelAmt)) & ",MISC_INFO='" & txt(SpclInfo) & "',RTO='" & txt(RTO) & "', SpecialDiscount=" & Val(txt(SpecialDiscount)) & ",Round_off=" & Val(txt(ROff)) & _
            " ,FB_Code='" & txt(FB_Code).Tag & "' , FIN_AcCode='" & FinAcCode & "', FIN_AMT=" & Val(txt(FinAmt)) & _
            " ,Net_Amount = " & Val(txt(GTotAmt)) & ", TrnType_Prn=" & mTrntypeprn & ",Fund_Source=" & mFundSource & ",Chassis='" & txt(ChassisNo) & _
            "',Inv_UName='" & pubUName & "', Inv_UEntDt=" & ConvertDate(PubServerDate) & ", Inv_UAE= 'A' " & _
            " ,Inv_AcPostByUName='" & txt(AcPostByName) & "',Inv_AcPostByUEntDt=" & ConvertDate(txt(AcPostDate)) & _
            " , Inv_AddBy ='" & pubUName & "', Inv_AddDate = " & ConvertDateTime(PubServerDate) & " ,SubTot=" & Val(txt(SubAmt)) & ", Colour_Code='" & txt(Colours).Tag & "', SubventionScheme='" & txt(SubventionScheme).Tag & "', DealerContribution=" & mDealerContribution & ", TataContribution=" & mTataContribution & ", DeliveryFrom='" & txt(DeliveryFrom) & "' where OrdDocId = '" & txt(BookNo).Tag & "'")
            
        GCn.Execute "Update Veh_Stock set Sal_DocId = '" & txt(TxtDocID) & "',Sal_VDate=" & ConvertDate(txt(VDate)) & ", Srv_BookNo = '" & txt(SrvBookNo) & "', TransAxlNo='" & txt(TransAxlNo) & "' where ChassisNo  = '" & txt(ChassisNo) & "' and Model='" & txt(Model) & "'"
    'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
        If GCn.Execute("select Chassis from Hiscard where Chassis='" & txt(ChassisNo) & "'").RecordCount = 0 Then
            CardNo = PubSiteCode + Right("000000" & GCn.Execute("select max(" & cVal(cMID("cardno", "3", "len(cardno)-1")) & ")+1 from hiscard").Fields(0), 6)
        
               GCn.Execute ("insert into hiscard (CardNo,CardDate,Div_Code,Site_Code,Model,Chassis,Engine, " & _
                " CouponNo,TransAxelNo,Supplier_BillNo,Supplier_BillDate,Name,Add1,Add2,Add3,CityCode,Govt_YN, " & _
                " U_Name, U_EntDt, U_AE,Delivery_Date)" & _
                " values ('" & CardNo & "'," & ConvertDate(txt(VDate)) & ",'" & PubDivCode & _
                "','" & PubSiteCode & "','" & txt(Model).TEXT & "','" & txt(ChassisNo).TEXT & _
                "','" & txt(EngineNo).TEXT & "','" & txt(SrvBookNo) & _
                "','" & txt(TransAxlNo) & "','" & mPBIllNo & "'," & ConvertDate(mPBIllDate) & _
                ",'" & txt(Party) & "','" & txt(Add1) & "','" & txt(Add2) & "','" & txt(Add3) & _
                "','" & txt(City).Tag & "'," & IIf(txt(Govt_YN) = "Yes", 1, 0) & ",'" & pubUName & _
                "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2.CAPTION, 1) & "'," & ConvertDate(txt(VDate)) & ")")

        End If
        mQuotDocID = GCn.Execute("select " & xIsNull("Quot_DocID", "") & " as QuotID from Veh_Order where OrdDocID = '" & txt(BookNo).Tag & "'").Fields(0).Value
        mQuotDocIDSrlNo = GCn.Execute("select " & xIsNull("QuotSrl_No", "") & " as QuotSrlNo from Veh_Order where OrdDocID = '" & txt(BookNo).Tag & "'").Fields(0).Value
        If mQuotDocID <> "" Then
            GCn.Execute "Update Veh_SubGroupQuot set Got_Lost='Got',GotLost_Date=" & ConvertDate(txt(VDate)) & " where QuotDocId='" & mQuotDocID & "' and QuotSrl_No=" & mQuotDocIDSrlNo & ""
        End If
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaV, txt(TxtDocID), txt(VDate)
    Else
        CreateLog Me, Master!SearchCode, False
        
        If txt(ChassisNo) <> mOldChasis Then
            'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
                '   GCn.Execute ("update hiscard set Dealer_Code='', CouponNo='', TransAxelNo='', Supplier_BillNo='', Supplier_BillDate=Null, Name='" & PubComp_Name & "',Add1='',Add2='',Add3='',CityCode='',Govt_YN =  0 " & _
                '        " where Chassis ='" & Txt(ChassisNo).Tag & "' and Model='" & Txt(Model).Tag & "'")
                GCn.Execute "Update Veh_Stock set Sal_DocId = '',Sal_VDate=Null, Srv_BookNo = '', TransAxlNo='' " & _
                        "where ChassisNo  = '" & mOldChasis & "' and Sal_DocId = '" & txt(TxtDocID) & "'"
        
         End If
        GCn.Execute ("update veh_order  " & _
            "set Form_Code='" & txt(FormType).Tag & "', Inv_Date=" & ConvertDate(txt(VDate)) & ", " & _
            "TAX_Per=" & Val(txt(TaxPer)) & ",model = '" & txt(Model) & "',TAX_Amt=" & Val(txt(TaxAmt)) & ", SatPer = " & Val(txt(SatPer)) & ",SatAmt=" & Val(txt(SatAmt)) & ",Surcharge_Per=" & Val(txt(TaxSurPer)) & ",Surcharge_Amt=" & Val(txt(TaxSurch)) & ",MARGINE=" & (Val(txt(SaleRate)) - Val(txt(NDP))) & ",VRATE=" & Val(txt(NDP)) & ", Subvention=" & Val(txt(Subvention)) & ",REBATE=" & Val(txt(Rebate)) & ", " & _
            "InciChrg=" & Val(txt(IncCharge)) & ",Octroi=" & Val(txt(Octroi)) & ",RegTemp=" & Val(txt(TempReg)) & ",TransitInsu=" & Val(txt(TransIns)) & ",Transport=" & Val(txt(Transportation)) & ", HandlingCharges=" & Val(txt(HandlingCharges)) & ",MVT=" & Val(txt(MVT)) & ",OtherChrg=" & Val(txt(MisCharge)) & ", RTOFee=" & Val(txt(RTOfee)) & ", Insurance=" & Val(txt(Insurance)) & ",FIT_AMT=" & Val(txt(OthFitAmt)) & ",FIT_TAX=" & Val(txt(OthFitTax)) & ", " & _
            "TOT_Per=" & Val(txt(TOTPer)) & ",TOT_Amt=" & Val(txt(TOTAmt)) & ",DieselAmt=" & Val(txt(FuelAmt)) & ",MISC_INFO='" & txt(SpclInfo) & "',RTO='" & txt(RTO) & "',Round_off=" & Val(txt(ROff)) & ", " & _
            "FB_Code='" & txt(FB_Code).Tag & "' , FIN_AcCode='" & FinAcCode & "', FIN_AMT=" & Val(txt(FinAmt)) & ", " & _
            "Net_Amount = " & Val(txt(GTotAmt)) & ",TrnType_Prn=" & mTrntypeprn & ",Fund_Source=" & mFundSource & ",Chassis='" & txt(ChassisNo) & "', Srv_BookNo='" & txt(SrvBookNo) & "', " & _
            "Inv_UName='" & pubUName & "', Inv_UEntDt=" & ConvertDate(PubServerDate) & ", Inv_UAE= 'E', " & _
            "Inv_AcPostByUName='" & txt(AcPostByName) & "',Inv_AcPostByUEntDt=" & ConvertDate(txt(AcPostDate)) & _
            ", Inv_ModifyBy = '" & pubUName & "', Inv_ModifyDate = " & ConvertDateTime(PubServerDate) & ", SpecialDiscount=" & Val(txt(SpecialDiscount)) & " ,SubTot=" & Val(txt(SubAmt)) & ", Colour_Code='" & txt(Colours).Tag & "', SubventionScheme='" & txt(SubventionScheme).Tag & "', DealerContribution=" & mDealerContribution & ", TataContribution=" & mTataContribution & ", DeliveryFrom='" & txt(DeliveryFrom) & "'   where Inv_DocId='" & txt(TxtDocID) & "'")
            
            
            
        GCn.Execute "Update Veh_Stock set Sal_DocId = '" & txt(TxtDocID) & "',Sal_VDate=" & ConvertDate(txt(VDate)) & ", Srv_BookNo = '" & txt(SrvBookNo) & "', TransAxlNo='" & txt(TransAxlNo) & _
            "' where ChassisNo  = '" & txt(ChassisNo) & "' and Model='" & txt(Model) & "'"
    'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
'        GCn.Execute ("update hiscard set Dealer_Code='" & mDlrID & "', CouponNo='" & Txt(SrvBookNo) & "', TransAxelNo='" & Txt(TransAxlNo) & "', Supplier_BillNo='" & mPBIllNo & "', Supplier_BillDate=" & ConvertDate(mPBIllDate) & _
'            ", Name='" & Txt(Party) & "',Add1='" & Txt(Add1) & "',Add2='" & Txt(Add2) & "',Add3='" & Txt(Add3) & "',CityCode='" & Txt(City).Tag & "',Govt_YN = " & IIf(Txt(Govt_YN) = "Yes", 1, 0) & _
'            " where Chassis ='" & Txt(ChassisNo) & "' and Model='" & Txt(Model) & "'")
    End If
    
    GCn.Execute ("update hiscard set Dealer_Code='" & mDlrID & "', CouponNo='" & txt(SrvBookNo) & "', TransAxelNo='" & txt(TransAxlNo) & "', Supplier_BillNo='" & mPBIllNo & "', Supplier_BillDate=" & ConvertDate(mPBIllDate) & _
        ", Name='" & txt(Party) & "',Add1='" & txt(Add1) & "',Add2='" & txt(Add2) & "',Add3='" & txt(Add3) & "',CityCode='" & txt(City).Tag & "',Govt_YN = " & IIf(txt(Govt_YN) = "Yes", 1, 0) & _
        " where Chassis ='" & txt(ChassisNo) & "' and Model='" & txt(Model) & "'")
    
    GCn.Execute ("delete from veh_purch2 where DocId='" & txt(TxtDocID) & "'")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ADItem) <> "" And Val(FGrid.TextMatrix(I, Qty)) <> 0 Then
            GCn.Execute ("insert into veh_purch2(DocId,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE,TAX_PER,TAX_AMT,TaxSur_Per,TaxSur_AMT, U_Name, U_EntDt, U_AE) " & _
                "values('" & txt(TxtDocID).TEXT & "'," & I & ",'" & PubSiteCode & txt(SiteCode).Tag & "','" & mVType & "','" & txt(SerialNo).TEXT & "', " & _
                "'" & FGrid.TextMatrix(I, ADItemCode) & "','A'," & Val(FGrid.TextMatrix(I, Qty)) & "," & Val(FGrid.TextMatrix(I, Rate)) & "," & Val(FGrid.TextMatrix(I, TaxPer1)) & ", " & _
                "" & Val(FGrid.TextMatrix(I, TaxAmt1)) & "," & Val(FGrid.TextMatrix(I, TaxSurPer1)) & "," & Val(FGrid.TextMatrix(I, TaxSurAmt1)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2.CAPTION, 1) & "')")
        End If
    Next
    'A/c Posting
    
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        If StrCmp(left(PubComp_Name, 5), "Ujwal") Then
            If CDate(RetDate(txt(VDate))) < CDate("01/OCT/2007") Then
                ProcAcPost
            End If
        Else
            ProcAcPost
        End If
    End If
    'EOF of A/c Posting Section
GCnFaV.CommitTrans
GCn.CommitTrans
Set Rst = Nothing
mTrans = False
    RSBook.Requery
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select Inv_DocId as SearchCode,Veh_Order.* from Veh_Order where left(Inv_DocID,1)='" & PubDivCode & "' and " & cTrim(cMID("Inv_DocID", "4", "5")) & "= '" & mVType & "' And Inv_DocId ='" & txt(TxtDocID) & "' Order by Inv_Date desc,Inv_No desc")
    End If
    Master.FIND "Inv_DocId = '" & txt(TxtDocID) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(DocID, Document_No)) Then
            MsgBox "Bill No." & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop

     Dim sitecond As String
     sitecond = " And Inv_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = sitecond & " and " & cMID("Veh_Order.inv_DocId", "3", "1") & "='" & PubSiteCode & "'"
     End If
    
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select Inv_DocId as SearchCode, " & cCStr("Inv_No", 10) & " As Invoice_No, " & cDt("Inv_Date") & " As Inv_Date, " & cCStr("Ord_No", 10) & " As Order_No,SG.Name,Model,Chassis,Inv_DocId " & _
        " from Veh_Order left join SubGroup SG on Veh_Order.PartyCode=SG.Subcode where left(Inv_DocID,1)='" & PubDivCode & "' and " & cTrim(cMID("Inv_DocID", "4", "5")) & " = '" & mVType & "' " & sitecond & " Order By Inv_Date Desc"
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
        Set Master = GCn.Execute("select Inv_DocId as SearchCode,Veh_Order.* from Veh_Order where left(Inv_DocID,1)='" & PubDivCode & "' and " & cTrim(cMID("Inv_DocID", "4", "5")) & "= '" & mVType & "' And Inv_DocId ='" & MyValue & "' Order by Inv_Date desc,Inv_No desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus txt(Index)
Grid_Hide
Dim I As Integer, XXA() As String  'Rst As ADODB.Recordset,
Select Case Index
    Case InvPrefix
        Set GRs = New ADODB.Recordset
        With GRs
             .CursorLocation = adUseClient
             .Open "SELECT Prefix from VehBill_Counter where Div_Code='" & PubDivCode & "' ", GCnFaV, adOpenDynamic, adLockOptimistic
        End With
        Do While Not GRs.EOF
            I = I
            ReDim Preserve XXA(I)
            XXA(I) = GRs!Prefix
            I = I + 1
            GRs.MoveNext
        Loop
        Set mListItem = ListView_Items(ListView, txt, InvPrefix, XXA, GRs.RecordCount)
        mInvPrefixHt = GRs.RecordCount * 260
        Set GRs = Nothing
        txt(Index) = ListView.SelectedItem.TEXT
    Case ADType
        ListArray = Array("No Detail", "Name/Qty", "Name/Qty/Amount")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
    Case DeliveryFrom
        ListArray = Array("ShowRoom", "Godown")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    Case FundSource
        ListArray = Array("Hypothecation", "Hire Purchase", "Own Fund", "Lease", "Agreement", "Lease & Agreement", "Loan Cum Hypt.")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 7)
    Case FB_Code
        rsFin.Requery
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or txt(FB_Code).TEXT = "" Then Exit Sub
        If txt(FB_Code).Tag <> rsFin!Code Then
            rsFin.MoveFirst
            rsFin.FIND "Code ='" & txt(FB_Code).TEXT & "'"
        End If
        
    Case BookNo
        RSBook.Filter = adFilterNone
        
        If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
            If StrCmp(txt(InvPrefix), PubVehTaxInvPrefix) Then
                RSBook.Filter = " LstNo <> '' And Site_Code = '" & txt(SiteCode).Tag & "' "
            Else
                RSBook.Filter = " LstNo = ''  And Site_Code = '" & txt(SiteCode).Tag & "' "
            End If
        Else
            If StrCmp(txt(InvPrefix), PubVehTaxInvPrefix) Then
                RSBook.Filter = " LstNo <> '' "
            Else
                RSBook.Filter = " LstNo = '' "
            End If
        End If
    
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or txt(BookNo).TEXT = "" Then Exit Sub
        If txt(BookNo).TEXT <> RSBook!Code Then
            RSBook.MoveFirst
            RSBook.FIND "code ='" & txt(BookNo).TEXT & "'"
        End If
        If StrCmp(left(PubComp_Name, 6), "Prayag") Then RSBook.Sort = "Name": Set DGBook.DataSource = RSBook
    Case ChassisNo
        If txt(Model) = "" Then MsgBox "Select Model First", vbInformation, "Validation": txt(Model).SetFocus: Exit Sub
        '14-05-03 lps
        'kunal
        If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
            Set RsChassis = GCn.Execute("SELECT VStk.ChassisNo as code, VStk.EngineNo, Model.Model as MODEL, VStk.Srv_BookNo, VStk.VRATE, VStk.Colour_Code, " & _
                " ColMast.Col_Desc, VStk.PBILL_NO, VStk.PBILL_DATE, " & cMID("VStk.Pur_DocId", "14", "8") & " as PurVNo, VStk.Pur_VDate, VStk.AL_Name, VStk.tax_yn, VStk.RSO_WORK, VStk.INDATE, VStk.BodyBuilder_IssDate, Godown.God_Name,Model.Model_Desc " & _
                " FROM (Veh_Stock VStk LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
                " left join Model on VStk.Model=Model.Model " & _
                " left join Godown on VStk.Godown=Godown.God_Code " & _
                " where Model.Div_Code='" & PubDivCode & "' and (VStk.Sal_DocId='" & txt(TxtDocID) & "' or VStk.Sal_DocId= '' or VStk.Sal_DocId Is Null) " & _
                " and (Vstk.Pur_VDate<=" & ConvertDate(txt(VDate)) & " or VStk.Pur_VDate is null)")
            'eof
            Set DgChassis.DataSource = RsChassis
        Else
            Set RsChassis = GCn.Execute("SELECT VStk.ChassisNo as code, VStk.EngineNo, VStk.MODEL, VStk.Srv_BookNo, VStk.VRATE, VStk.Colour_Code, " & _
                " ColMast.Col_Desc, VStk.PBILL_NO, VStk.PBILL_DATE, " & cMID("VStk.Pur_DocId", "14", "8") & " as PurVNo, VStk.Pur_VDate, VStk.AL_Name, VStk.tax_yn, VStk.RSO_WORK, VStk.INDATE, VStk.BodyBuilder_IssDate, Godown.God_Name " & _
                " FROM (Veh_Stock VStk LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
                " left join Model on VStk.Model=Model.Model " & _
                " left join Godown on VStk.Godown=Godown.God_Code " & _
                " where Model.Div_Code='" & PubDivCode & "' and (VStk.Sal_DocId='" & txt(TxtDocID) & "' or VStk.Sal_DocId= '' or VStk.Sal_DocId Is Null) " & _
                " and (Vstk.Pur_VDate<=" & ConvertDate(txt(VDate)) & " or VStk.Pur_VDate is null)")
            'eof
            Set DgChassis.DataSource = RsChassis
            DgChassis.Columns(3).Visible = False
        End If
    Case Colours
        If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsCol!Col_Desc Then
            RsCol.MoveFirst
            RsCol.FIND "Col_desc='" & txt(Index).TEXT & "'"
        End If
    Case SiteCode
        Set DGSite.DataSource = RsSite
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
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case SubventionScheme
        
        If txt(Model).Tag <> "" Then
            Set RsSubvention = GCn.Execute("Select SchemeNo As Code, SchemeNo As Name, FromDate, ToDate, MG.ModelGrp_Name, Model, " & _
                                           "DealerContribution, TataContribution, TotalSubvention " & _
                                           "From Subvention S " & _
                                           "Left Join Model_Grp MG On S.ModelGroup=MG.ModelGrp_Code " & _
                                           "Where ModelGroup = (Select Grp_Code From Model Where Model='" & txt(Model).Tag & "') " & _
                                           "Order By SchemeNo")
            Set DgSubvention.DataSource = RsSubvention
        End If
        If RsSubvention.RecordCount = 0 Or (RsSubvention.EOF = True Or RsSubvention.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsSubvention!Name Then
            RsSubvention.MoveFirst
            RsSubvention.FIND "name ='" & txt(Index).TEXT & "'"
        End If
        
'    Case SerialNo, TaxAmt, TaxSurch, TaxPer, TaxSurPer, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation
'        SendKeys "{HOME}+{END}"
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
    Case InvPrefix
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, mInvPrefixHt
    Case BookNo
        If StrCmp(left(PubComp_Name, 6), "Prayag") Then
            DGridTxtKeyDown DGBook, txt, Index, RSBook, KeyCode, False, 3
        Else
            DGridTxtKeyDown DGBook, txt, Index, RSBook, KeyCode, False, 0
        End If
        
    Case ADType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 900
    Case DeliveryFrom
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case FundSource
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width + 50, 1600
    Case SiteCode
        DGridTxtKeyDown DGSite, txt, Index, RsSite, KeyCode, False, 1
'    Case Model
'        DGridTxtKeyDown DGMod, txt, Index, RsMod, KeyCode, False, 0, frmModel
'        If DGMod.Visible = True Then txt(ChassisNo).Text = ""
    Case ChassisNo
        DGridTxtKeyDown DgChassis, txt, Index, RsChassis, KeyCode, False, 0
    Case FormType
        DGridTxtKeyDown DGForm, txt, Index, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
    Case SubventionScheme
        DGridTxtKeyDown DgSubvention, txt, Index, RsSubvention, KeyCode, False, 1, FrmSubventionMast, "FrmSubventionMast"
    Case FB_Code
        DGridTxtKeyDown DGFin, txt, Index, rsFin, KeyCode, False, 1, frmFinMast, "frmFinMast"
    Case Colours
        DGridTxtKeyDown DGCol, txt, Index, RsCol, KeyCode, False, 1, frmColor, "frmColor"
    Case Model
        DGridTxtKeyDown DGMod, txt, Index, RsMod, KeyCode, False, 0, frmModel, "frmModel"
    Case ChassisNo
        DGridTxtKeyDown DgChassis, txt, Index, RsChassis, KeyCode, False, 0
End Select
If FrmList.Visible = False And DgSubvention.Visible = False And DGBook.Visible = False And DGSite.Visible = False And DGFin.Visible = False _
    And DgChassis.Visible = False And DGMod.Visible = False And DGForm.Visible = False And DGCol.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VDate Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> SpclInfo Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = SpclInfo Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SiteCode Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> BookNo Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case BookNo
        If StrCmp(left(PubComp_Name, 6), "Prayag") Then
            If DGBook.Visible = True Then DGridTxtKeyPress txt, Index, RSBook, KeyAscii, "Name"
        Else
            If DGBook.Visible = True Then DGridTxtKeyPress txt, Index, RSBook, KeyAscii, "Code"
        End If
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress txt, Index, RsSite, KeyAscii, "Name"
'    Case Model
'        If DGMod.Visible = True Then DGridTxtKeyPress txt, Index, RsMod, KeyAscii, "code"
    Case ChassisNo
        If DgChassis.Visible = True Then DGridTxtKeyPress txt, Index, RsChassis, KeyAscii, "code", False
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, Index, rsForm, KeyAscii, "Name"
    Case SubventionScheme
        If DgSubvention.Visible = True Then DGridTxtKeyPress txt, Index, RsSubvention, KeyAscii, "Name"
    Case FB_Code
        If DGFin.Visible = True Then DGridTxtKeyPress txt, Index, rsFin, KeyAscii, "Name"
    Case Colours
        If DGCol.Visible = True Then DGridTxtKeyPress txt, Index, RsCol, KeyAscii, "Col_Desc"
    Case SerialNo
        Call NumPress(txt(Index), KeyAscii, 7, 0)
    Case SaleRate, Rebate, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation, HandlingCharges, SubAmt, Subvention
        Call NumPress(txt(Index), KeyAscii, 8, 2)
    Case TaxAmt, TaxSurch, MisCharge, FinAmt, TOTAmt, RTOfee, Insurance, SatAmt
        Call NumPress(txt(Index), KeyAscii, 7, 2)
    Case FuelAmt
        Call NumPress(txt(Index), KeyAscii, 6, 2)
    Case TaxPer, TaxSurPer, TOTPer, SatPer
        Call NumPress(txt(Index), KeyAscii, 2, 2)
    Case SpecialDiscount
        Call NumPress(txt(Index), KeyAscii, 8, 2)
End Select
'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case InvPrefix, FundSource, ADType, DeliveryFrom
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case FormType
        If DGForm.Visible = True Then
            If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            txt(TaxPer).TEXT = IIf(IsNull(rsForm!Tax_Per), 0, rsForm!Tax_Per)
            txt(SatPer).TEXT = IIf(IsNull(rsForm!AddTaxPer), 0, rsForm!AddTaxPer)
            txt(TaxAmt).TEXT = Val(txt(SubTotA).TEXT) * Val(txt(TaxPer).TEXT) / 100
            txt(SatAmt).TEXT = Val(txt(SubTotA).TEXT) * Val(txt(SatPer).TEXT) / 100
            txt(TaxSurPer).TEXT = IIf(IsNull(rsForm!Tax_Sur_Per), 0, rsForm!Tax_Sur_Per)
            txt(TaxSurch).TEXT = Val(txt(TaxSurPer).TEXT) * Val(txt(TaxAmt).TEXT) / 100
             Amt_Cal False
        End If
    Case TaxPer, TaxSurPer, TOTPer, SatPer
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
         Amt_Cal False
    Case SubAmt
          Amt_Cal False
            SubTot = Val(txt(Index))
            'Txt(SubAmt) = Val(Txt(Index))
    Case Rebate, Subvention
         Amt_Cal False
         'SubTot = SubTot
         'Txt(SubAmt) = SubTott
     Case IncCharge, Octroi, TempReg, TransIns, MVT, Transportation, HandlingCharges
          Amt_Cal False
    Case MisCharge, FinAmt, AdvAmt, FuelAmt, Insurance, RTOfee
          Amt_Cal False
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Integer
Select Case Index
    Case DeliveryFrom
        If txt(Index) <> "" Then txt(Index) = ListView.SelectedItem.TEXT
        Amt_Cal False
    Case FundSource, ADType
        If txt(Index) <> "" Then txt(Index) = ListView.SelectedItem.TEXT
    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RSBook!Code
            txt(Index).Tag = RSBook!OrdDocId
            If XNull(RSBook!Ord_Date) <> "" Then
                If DateDiff("D", XNull(RSBook!Ord_Date), txt(VDate)) < 0 Then
                    MsgBox "Bill Date Can't be Less than Booking Date!"
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
        Cancel = Not FillRecords '= False Then  = True
'    Case Model
'        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or txt(Index).Text = "" Then
'            txt(Index).Text = ""
'            txt(Index).Tag = ""
'        Else
'            txt(Index).Text = RsMod!Code
'            txt(Index).Tag = RsMod!Code
'        End If
'        If IsValid(txt(Index), "Model") = False Then Cancel = True: GoTo lblExitSub
'        txt(ChassisNo).SetFocus
    
    Case FB_Code
        rsFin.Requery
        rsFin.MoveFirst
        rsFin.FIND "Code ='" & txt(FB_Code).Tag & "'"
        
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            LblFinancer.Tag = ""
            FinAcCode = ""
            txt(FinAdd1) = ""
            txt(FinAdd2) = ""
        Else
            txt(Index).TEXT = rsFin!Name
            txt(Index).Tag = rsFin!Code
            LblFinancer.Tag = XNull(rsFin!FinbankCode)
            FinAcCode = rsFin!AcCode
            LblFinName.Tag = rsFin!FinName
            LblFinGrp.Tag = rsFin!UnderFinGrp
            txt(FinAdd1) = XNull(rsFin!Add1)
            txt(FinAdd2) = XNull(rsFin!Add2)
        End If
        
    Case ChassisNo
        IsValid txt(Index), "Chassis No.", True
        If RsChassis.RecordCount = 0 Or (RsChassis.EOF = True Or RsChassis.BOF = True) Or txt(Index).TEXT = "" Then
            txt(ChassisNo) = ""
            Cancel = Not Fill_Data(False)
        Else
            txt(ChassisNo) = RsChassis!Code
            If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
                If XNull(RsChassis!BodyBuilder_IssDate) <> "" Then
                    MsgBox "Chassis was issued to body builder"
                End If
            End If
            If UCase(left(PubComp_Name, 4)) <> "ENAR" Then
                txt(Colours).Tag = IIf(IsNull(RsChassis!Colour_Code), "", RsChassis!Colour_Code)
                If txt(Colours).Tag <> "" Then
                    txt(Colours).TEXT = GCn.Execute("select col_desc from colmast where col_code = '" & txt(Colours).Tag & "'").Fields(0).Value
                End If
            End If
            Cancel = Not Fill_Data(True)
        End If
'        If Txt(ChassisNo).Text = "" And RsChassis.RecordCount  > 0 Then
'            MsgBox "chassis no is required", vbInformation, "Validation  Check"
'            Cancel = True: GoTo lblExitSub
'        End If
    Case SiteCode
        If IsValid(txt(Index), "Site Code") = False Then Cancel = True: GoTo lblExitSub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsSite!Name
            txt(Index).Tag = RsSite!Code
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsForm!Name
            txt(Index).Tag = rsForm!Code
        End If
    Case SubventionScheme
        If RsSubvention.RecordCount = 0 Or (RsSubvention.EOF = True Or RsSubvention.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            txt(Subvention) = "0.00"
        Else
            txt(Index).TEXT = RsSubvention!Name
            txt(Index).Tag = RsSubvention!Code
            txt(Subvention) = Format(RsSubvention!TotalSubvention, "0.00")
            mDealerContribution = VNull(RsSubvention!DealerContribution)
            mTataContribution = VNull(RsSubvention!TataContribution)
            Amt_Cal False
            If Val(txt(NDP)) > Val(txt(SubTotA)) Then
                MsgBox "By Using this Scheme SubTotal(A) Is Less Than N.D.P.!"
'                txt(Index).TEXT = ""
'                txt(Index).Tag = ""
'                txt(Subvention) = "0.00"
'                mDealerContribution = 0
'                mTataContribution = 0
            End If
        End If
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
            txt(VDate).TEXT = PubLoginDate
        Else
0
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        If CheckFinYear(txt(Index)) Then
'            txt(TxtDocId) = GetDocIDVBill(GCnFaV, mVType, txt(Vdate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
'            DocId = txt(TxtDocId)
        Else
            Cancel = True
        End If
    Case InvPrefix
        If txt(Index) <> "" Then txt(Index) = ListView.SelectedItem.TEXT
        LblVPrefix = txt(Index)
        txt(TxtDocID) = GetDocIDVBill(GCnFaV, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix, txt(SiteCode).Tag)
        DocID = txt(TxtDocID)
    Case SerialNo
        If IsValid(txt(SerialNo), "Serial No.") = False Then Cancel = True:  GoTo lblExitSub
        'If VoucherEditFlag Then      ' Manual
            txt(TxtDocID) = GetDocIDVBill(GCnFaV, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix, txt(SiteCode).Tag)
            DocID = txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select Inv_DocId From veh_Order Where Inv_DocID='" & txt(TxtDocID) & "'", GCn, adOpenStatic, adLockReadOnly
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Set Rst = Nothing
                Cancel = True
                txt(SerialNo).SetFocus
         'End If
        End If
    Case Rebate, Subvention
         Amt_Cal True
         txt(Index).TEXT = IIf(Val(txt(Index)) <> 0, Format(txt(Index), "0.00"), "")
         If Val(txt(Index)) > Val(txt(SaleRate)) Then
            MsgBox "Rebate Rs." & txt(Rebate) & " is greater than Sale Rate Rs." & txt(SaleRate), vbOKOnly, "Validation"
            Cancel = True
        End If
        If PubVehRateIncTaxYn = 1 Then AmtCal1
    Case TaxSurPer, TaxSurch, SaleRate, Rebate, Subvention, IncCharge, Octroi, TempReg, TransIns, MVT, Transportation, HandlingCharges, MisCharge, RTOfee, Insurance, OthFitAmt, OthFitTax, FinAmt, AdvAmt, FuelAmt
         txt(Index).TEXT = IIf(Val(txt(Index)) <> 0, Format(txt(Index), "0.00"), "")
         Amt_Cal True
End Select
lblExitSub:
Set Rst = Nothing
End Sub

Private Sub DGADItem_Click()
    If RsADItem.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsADItem!Name
         FGrid.TextMatrix(FGrid.Row, ADItem) = RsADItem!Name
         FGrid.TextMatrix(FGrid.Row, ADItemCode) = RsADItem!Code
    End If
    TxtGrid(0).SetFocus
    DGADItem.Visible = False
End Sub

Private Sub DgChassis_Click()
    If RsChassis.RecordCount > 0 Then
        txt(ChassisNo).TEXT = RsChassis!Code
        Fill_Data True
    End If
    txt(ChassisNo).SetFocus
    DgChassis.Visible = False
End Sub

Private Sub DGMod_Click()
If RsMod.RecordCount > 0 Then
    txt(Model) = RsMod!Code
End If
txt(Model).SetFocus
DGMod.Visible = False
End Sub

Private Sub DGForm_Click()
    If rsForm.RecordCount > 0 Then
        txt(FormType).TEXT = rsForm!Name
        txt(FormType).Tag = rsForm!Code
    End If
    txt(FormType).SetFocus
    DGForm.Visible = False
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
txt(Model).Tag = ""
txt(ChassisNo).Tag = ""
LblFinancer.Tag = ""
LblFinancerGroup.Tag = ""
LblFinGrp.Tag = ""
LblFinName.Tag = ""
ChkNewFinancer.Value = 0
End Sub

Private Sub MoveRec()
Dim RsTemp As ADODB.Recordset
Dim Rst As Recordset
Dim I As Integer
On Error GoTo error1
Grid_Hide

    If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
        If AllowEditDel(pubUName, Master!Inv_Date, PubLoginDate) = False Then
            TopCtrl1.tDel = False
            TopCtrl1.tEdit = False
        Else
            TopCtrl1.tDel = True
            TopCtrl1.tEdit = True
        End If
    End If

If Master.RecordCount > 0 Then
    DocID = Master!Inv_DocId
    txt(TxtDocID) = Master!Inv_DocId
    LblDiv.CAPTION = "Division : " & left(Master!Inv_DocId, 1)
    LblSite.CAPTION = "Site Code : " & mID(Master!Inv_SiteCode, 1, 1)
    txt(SiteCode).Tag = mID(Master!Inv_SiteCode, 2, 1)
    txt(SiteCode).TEXT = GCn.Execute("select site_desc from site where site_code = '" & txt(SiteCode).Tag & "'").Fields(0).Value
    LblUser = IIf(Not IsNull(Master!Inv_AddDate), "Add By : " & XNull(Master!Inv_AddBy) & "  Dated : " & XNull(Master!Inv_AddDate), "") & IIf(Not IsNull(Master!Inv_ModifyDate), "     Modify By : " & XNull(Master!Inv_ModifyBy) & "  Dated : " & XNull(Master!Inv_ModifyDate), "")
    LblVPrefix.CAPTION = mID(Master!Inv_DocId, 8, 5)
    txt(InvPrefix) = DeCodeDocID(Master!Inv_DocId, Document_Prefix)
    txt(SerialNo).TEXT = Master!Inv_No
    txt(VDate).TEXT = XNull(Master!Inv_Date)
    txt(BookNo).TEXT = Master!Ord_No
    txt(BookNo).Tag = Master!OrdDocId
    txt(DelChNo) = DeCodeDocID(XNull(Master!DelCh_DocId), Document_No)
    txt(DelChDate) = IIf(IsNull(Master!DelCh_DT), "", Master!DelCh_DT)
    '*** A/c Posting Status
    txt(AcPostByName) = IIf(IsNull(Master!Inv_AcPostByUName), "", Master!Inv_AcPostByUName)
    txt(AcPostDate) = IIf(IsNull(Master!Inv_AcPostByUEntDt), "", Master!Inv_AcPostByUEntDt)
    '***
    If Not IsNull(Master!Fund_Source) Then
        Select Case Master!Fund_Source
            Case 0 '0 Hypothecation ,1 Hire purchase ,2 Own Fund,3 Lease
                txt(FundSource).TEXT = "Hypothecation"
            Case 1
                txt(FundSource).TEXT = "Hire Purchase"
'            Case 2
'                txt(FundSource).Text = "Own Fund"
            Case 3
                txt(FundSource).TEXT = "Lease"
            Case 4
                txt(FundSource).TEXT = "Agreement"
            Case 5
                txt(FundSource).TEXT = "Lease & Agreement"
            Case 6
                txt(FundSource).TEXT = "Loan Cum Hypt."
            Case Else
                txt(FundSource).TEXT = "Own Fund"
        End Select
    Else
        txt(FundSource).TEXT = ""
    End If
    If Not IsNull(Master!TrnType_Prn) Then
        Select Case Master!TrnType_Prn
            Case 0
                txt(ADType).TEXT = "No Detail"
            Case 1
                txt(ADType).TEXT = "Name/Qty"
            Case 2
                txt(ADType).TEXT = "Name/Qty/Amount"
        End Select
    Else
        txt(ADType).TEXT = ""
    End If
    
    txt(Party).Tag = IIf(IsNull(Master!PartyCode), "", Master!PartyCode)
    If txt(Party).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NamePrefix,name,FPrefix,FName,add1,add2,add3,CityCode, LstNo from SubGroup where Subcode = '" & txt(Party).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        txt(NamePrefix).TEXT = IIf(IsNull(Rst!NamePrefix), "", Rst!NamePrefix)
        txt(Party).TEXT = Rst!Name
        txt(FNamePrefix).TEXT = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
        txt(fname).TEXT = IIf(IsNull(Rst!fname), "", Rst!fname)
        txt(Add1).TEXT = IIf(IsNull(Rst!Add1), "", Rst!Add1)
        txt(Add2).TEXT = IIf(IsNull(Rst!Add2), "", Rst!Add2)
        txt(Add3).TEXT = IIf(IsNull(Rst!Add3), "", Rst!Add3)
        txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
        mPartyLstNo = XNull(Rst!LstNo)
        If txt(City).Tag <> "" Then
            txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & txt(City).Tag & "'").Fields(0).Value
        End If
    End If
    txt(Model).TEXT = Master!Model
    txt(Model).Tag = Master!Model
    If txt(Model).Tag <> "" Then
        Set RsTemp = GCn.Execute("Select " & xIsNull("Model_Desc", "") & " From Model Where Model='" & txt(Model).Tag & "'  ")
        If RsTemp.RecordCount > 0 Then
            txt(ModelDesc) = RsTemp(0)
        End If
    End If
    txt(Govt_YN).TEXT = IIf(Master!Govt_YN = 1, "Yes", "No")
    txt(Colours).Tag = IIf(IsNull(Master!Colour_Code), "", Master!Colour_Code)
    If txt(Colours).Tag <> "" Then
        txt(Colours).TEXT = GCn.Execute("select col_desc from colmast where col_code = '" & txt(Colours).Tag & "'").Fields(0).Value
    End If
    txt(FormType).Tag = IIf(IsNull(Master!Form_Code), "", Master!Form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType).TEXT = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType).TEXT = ""
    End If
    txt(NDP).TEXT = IIf(IsNull(Master!vrate) Or Master!vrate = 0, "", Format(Master!vrate, "0.00"))
    
    If PubVehRateIncTaxYn <> 1 Or VNull(Master!SubTot) = 0 Then
        txt(SaleRate).TEXT = Format(Val(txt(NDP)) + IIf(IsNull(Master!Margine), 0, Master!Margine), "0.00")
    Else
       txt(SaleRate).TEXT = VNull(Format(Master!SubTot, "0.00"))
    End If
    txt(DeliveryFrom) = XNull(Master!DeliveryFrom)
    txt(Rebate).TEXT = IIf(IsNull(Master!Rebate) Or Master!Rebate = 0, "", Format(Master!Rebate, "0.00"))
    mDealerContribution = VNull(Master!DealerContribution)
    mTataContribution = VNull(Master!TataContribution)
    txt(Subvention).TEXT = IIf(IsNull(Master!Subvention) Or Master!Subvention = 0, "", Format(Master!Subvention, "0.00"))
    txt(IncCharge).TEXT = IIf(IsNull(Master!InciChrg) Or Master!InciChrg = 0, "", Format(Master!InciChrg, "0.00"))
    txt(Octroi).TEXT = IIf(IsNull(Master!Octroi) Or Master!Octroi = 0, "", Format(Master!Octroi, "0.00"))
    txt(TempReg).TEXT = IIf(IsNull(Master!RegTemp) Or Master!RegTemp = 0, "", Format(Master!RegTemp, "0.00"))
    txt(TransIns).TEXT = IIf(IsNull(Master!TransitInsu) Or Master!TransitInsu = 0, "", Format(Master!TransitInsu, "0.00"))
    txt(MVT).TEXT = IIf(IsNull(Master!MVT) Or Master!MVT = 0, "", Format(Master!MVT, "0.00"))
    txt(Transportation).TEXT = IIf(IsNull(Master!Transport) Or Master!Transport = 0, "", Format(Master!Transport, "0.00"))
    txt(HandlingCharges).TEXT = IIf(IsNull(Master!HandlingCharges) Or Master!HandlingCharges = 0, "", Format(Master!HandlingCharges, "0.00"))
    txt(SubTotA) = Format((Val(txt(SaleRate)) - Val(txt(Rebate)) - Val(txt(Subvention)) + Val(txt(IncCharge)) + Val(txt(Octroi)) + Val(txt(TempReg)) + Val(txt(TransIns)) + Val(txt(MVT)) + Val(txt(Transportation)) + Val(txt(HandlingCharges))), "0.00")
    txt(SpecialDiscount) = Format(VNull(Master!SpecialDiscount), "0.00")
    
    txt(TaxPer).TEXT = IIf(IsNull(Master!Tax_Per) Or Master!Tax_Per = 0, "", Format(Master!Tax_Per, "0.00"))
    txt(TaxAmt).TEXT = IIf(IsNull(Master!Tax_Amt) Or Master!Tax_Amt = 0, "", Format(Master!Tax_Amt, "0.00"))
    txt(SatPer).TEXT = IIf(IsNull(Master!SatPer) Or Master!SatPer = 0, "", Format(Master!SatPer, "0.00"))
    txt(SatAmt).TEXT = IIf(IsNull(Master!SatAmt) Or Master!SatAmt = 0, "", Format(Master!SatAmt, "0.00"))
    
    txt(TaxSurPer).TEXT = IIf(IsNull(Master!surcharge_per) Or Master!surcharge_per = 0, "", Format(Master!surcharge_per, "0.00"))
    txt(TaxSurch).TEXT = IIf(IsNull(Master!Surcharge_Amt) Or Master!Surcharge_Amt = 0, "", Format(Master!Surcharge_Amt, "0.00"))
    txt(MisCharge).TEXT = IIf(IsNull(Master!OtherChrg) Or Master!OtherChrg = 0, "", Format(Master!OtherChrg, "0.00"))
    txt(RTOfee).TEXT = IIf(IsNull(Master!RTOfee) Or Master!RTOfee = 0, "", Format(Master!RTOfee, "0.00"))
    txt(Insurance).TEXT = IIf(IsNull(Master!Insurance) Or Master!Insurance = 0, "", Format(Master!Insurance, "0.00"))
    txt(SubTotB) = Format((Val(txt(SubTotA)) + Val(txt(TaxAmt)) + Val(txt(SatAmt)) + Val(txt(TaxSurch)) + Val(txt(MisCharge))), "0.00")
        
    txt(OthFitAmt).TEXT = IIf(IsNull(Master!Fit_Amt) Or Master!Fit_Amt = 0, "", Format(Master!Fit_Amt, "0.00"))
    txt(OthFitTax).TEXT = IIf(IsNull(Master!Fit_Tax) Or Master!Fit_Tax = 0, "", Format(Master!Fit_Tax, "0.00"))
    txt(TOTPer) = IIf(IsNull(Master!TOT_Per) Or Master!TOT_Per = 0, "", Format(Master!TOT_Per, "0.00"))
    txt(TOTAmt) = IIf(IsNull(Master!Tot_Amt) Or Master!Tot_Amt = 0, "", Format(Master!Tot_Amt, "0.00"))
    txt(FuelAmt).TEXT = IIf(IsNull(Master!DieselAmt) Or Master!DieselAmt = 0, "", Format(Master!DieselAmt, "0.00"))
    txt(ROff).TEXT = IIf(IsNull(Master!Round_off) Or Master!Round_off = 0, "", Format(Master!Round_off, "0.00"))
    'Modi LPS 05.12.2003
    txt(GTotAmt) = Format((Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) + Val(txt(TOTAmt)) - Val(txt(FuelAmt)) + Val(txt(ROff))), "0.00")
    'eof Modi
    
    'modified for docid / invdate by lps
    txt(AdvAmt) = IIf(PartyAdvance(Master!OrdDocId, txt(VDate)) <> 0, Format(PartyAdvance(Master!OrdDocId, txt(VDate)), "0.00"), "")
  ' Txt(AdvAmt) = Format(IIf(IsNull(Master!P_Amount), 0, Master!P_Amount), "0.00")
  ' end modi
    txt(NetOStng) = Format((Val(txt(GTotAmt)) - Val(txt(AdvAmt))), "0.00")
    txt(FinAmt).TEXT = IIf(IsNull(Master!Fin_Amt) Or Master!Fin_Amt = 0, "", Format(Master!Fin_Amt, "0.00"))
    txt(FB_Code).Tag = IIf(IsNull(Master!FB_Code), "", Master!FB_Code)
    
    If txt(FB_Code).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select fincode as code,finname + ',' +  " & xIsNull("add1", "") & " +  ',' +  " & xIsNull("add2", "") & " + " & xIsNull("City.CityName", "") & " as name, FinName, Add1, Add2,AcCode,FinBankCode, UnderFinGrp from ContractFinance " & _
                 "left join city on left(ContractFinance.City,4)=City.CityCode where fincatg = 0  And FinCode = '" & txt(FB_Code).Tag & "' order by finname ", GCn, adOpenDynamic, adLockOptimistic

'        Rst.Open "select fincode as code,finname + ',' + " & xIsNull("City.CityName", "") & " as name,AcCode " & _
'        " from ContractFinance left join city on left(ContractFinance.City,4)=City.CityCode " & _
'        " where fincatg = 0 and  fincode = '" & txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        
        txt(FB_Code).TEXT = Rst!Name
        txt(FB_Code).Tag = Rst!Code
        LblFinancer.Tag = XNull(Rst!FinbankCode)
        FinAcCode = Rst!AcCode
        LblFinName.Tag = Rst!FinName
        LblFinGrp.Tag = Rst!UnderFinGrp
        txt(FinAdd1) = XNull(Rst!Add1)
        txt(FinAdd2) = XNull(Rst!Add2)
        FinAcCode = IIf(IsNull(Rst!AcCode), "", Rst!AcCode)
    Else
        txt(FB_Code).TEXT = ""
        FinAcCode = ""
    End If
    
    txt(SubventionScheme) = XNull(Master!SubventionScheme)
    txt(SubventionScheme).Tag = XNull(Master!SubventionScheme)
    txt(SpclInfo).TEXT = IIf(IsNull(Master!MISC_INFO), "", Master!MISC_INFO)
    txt(RTO).TEXT = IIf(IsNull(Master!RTO), "", Master!RTO)
    txt(ChassisNo).TEXT = IIf(IsNull(Master!Chassis), "", Master!Chassis)
    txt(ChassisNo).Tag = IIf(IsNull(Master!Chassis), "", Master!Chassis)
    txt(SubAmt) = txt(SaleRate).TEXT
    SubTot = txt(SubAmt)
    
    mOldChasis = txt(ChassisNo).Tag
    Set Rst = New Recordset
    Rst.Open "SELECT Veh_Stock.TransAxlNo,Veh_Stock.Srv_BookNo,Veh_Stock.EngineNo,Veh_Stock.VehSerialNo, " & _
                    "Veh_Stock.tax_yn,Veh_Stock.PBILL_NO,Veh_Stock.PBILL_DATE " & _
             "FROM Veh_Stock " & _
             "where Veh_Stock.MODEL  = '" & txt(Model) & "' and Veh_Stock.ChassisNo = '" & txt(ChassisNo) & "' " & _
             "and Veh_Stock.Sal_DocId= '" & Master!Inv_DocId & "'", GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        txt(TransAxlNo).TEXT = IIf(IsNull(Rst!TransAxlNo), "", Rst!TransAxlNo)
        txt(SrvBookNo).TEXT = IIf(IsNull(Rst!Srv_BookNo), "", Rst!Srv_BookNo)
        txt(EngineNo).TEXT = IIf(IsNull(Rst!EngineNo), "", Rst!EngineNo)
        txt(TelcoInvNo).TEXT = IIf(IsNull(Rst!PBILL_NO), "", Rst!PBILL_NO)
        txt(TelcoInvDate).TEXT = IIf(IsNull(Rst!PBILL_DATE), "", Rst!PBILL_DATE)
        txt(Taxable).TEXT = IIf(Rst!Tax_YN = 1, "Yes", "No")
    Else
        txt(TransAxlNo).TEXT = ""
        txt(SrvBookNo).TEXT = ""
        txt(EngineNo).TEXT = ""
        txt(TelcoInvNo).TEXT = ""
        txt(TelcoInvDate).TEXT = ""
        txt(Taxable).TEXT = ""
    End If
    Set Rst = New Recordset
    Set Rst = GCn.Execute("SELECT Veh_AMDModel.Prod_Name, Veh_Purch2.Srl_No, Veh_Purch2.PROD_CODE, Veh_Purch2.QTY, Veh_Purch2.RATE,Veh_Purch2.TAX_PER,Veh_Purch2.TAX_AMT,Veh_Purch2.TaxSur_Per,Veh_Purch2.TaxSur_AMT " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_purch2.DocId = '" & Master!Inv_DocId & "'")
    
    
    
    FGrid.Rows = 1
    I = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            With FGrid
                .AddItem ""
                .TextMatrix(I, 0) = Rst!Srl_No
                .TextMatrix(I, ADItem) = Rst!Prod_Name
                .TextMatrix(I, Qty) = Format(IIf(IsNull(Rst!Qty), "", Rst!Qty), "0")
                .TextMatrix(I, Rate) = Format(IIf(IsNull(Rst!Rate), "", Rst!Rate), "0.00")
                .TextMatrix(I, Amt) = Format(.TextMatrix(I, Qty) * .TextMatrix(I, Rate), "0.00")
                .TextMatrix(I, TaxPer1) = Format(IIf(IsNull(Rst!Tax_Per), "", Rst!Tax_Per), "0.00")
                .TextMatrix(I, TaxAmt1) = Format(IIf(IsNull(Rst!Tax_Amt), "", Rst!Tax_Amt), "0.00")
                .TextMatrix(I, TaxSurPer1) = Format(IIf(IsNull(Rst!TaxSur_Per), "", Rst!TaxSur_Per), "0.00")
                .TextMatrix(I, TaxSurAmt1) = Format(IIf(IsNull(Rst!TaxSur_Amt), "", Rst!TaxSur_Amt), "0.00")
                .TextMatrix(I, FinalAmt) = Format((Val(.TextMatrix(I, Amt)) + Val(.TextMatrix(I, TaxAmt1)) + Val(.TextMatrix(I, TaxSurAmt1))), "0.00")
                .TextMatrix(I, ADItemCode) = Rst!Prod_Code
            End With
            Rst.MoveNext
           I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Amt_Cal True
    Set Rst = Nothing
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
'lp 10-03-03
Amt_Cal False
Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte
    
    With FGrid
        .left = Me.left '+45
        .top = 3120
        .Cols = 11
'        .BackColor = CellBackColLeave
'        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight

        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, ADItem) = "Accessories / Additional Fitments"
        .ColAlignment(ADItem) = flexAlignLeftCenter
        .ColWidth(ADItem) = 3720
        
       
        .TextMatrix(0, Qty) = "Qty"
        .ColAlignmentFixed(Qty) = flexAlignRightCenter
        .ColWidth(Qty) = 540

        .TextMatrix(0, Rate) = "Rate"
        .ColAlignmentFixed(Rate) = flexAlignRightCenter
        .ColWidth(Rate) = 855
        
        .TextMatrix(0, Amt) = "Amount"
        .ColAlignmentFixed(Amt) = flexAlignRightCenter
        .ColWidth(Amt) = 1065
        
        .TextMatrix(0, TaxPer1) = "Tax%"
        .ColAlignmentFixed(TaxPer1) = flexAlignRightCenter
        .ColWidth(TaxPer1) = 690
        
        .TextMatrix(0, TaxAmt1) = "TaxAmt"
        .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
        .ColWidth(TaxAmt1) = 990
        
        .TextMatrix(0, TaxSurPer1) = "Surch%"
        .ColAlignmentFixed(TaxSurPer1) = flexAlignRightCenter
        .ColWidth(TaxSurPer1) = 720
        
        .TextMatrix(0, TaxSurAmt1) = "SurchAmt"
        .ColAlignmentFixed(TaxSurAmt1) = flexAlignRightCenter
        .ColWidth(TaxSurAmt1) = 990
  
        .TextMatrix(0, FinalAmt) = "NetAmt"
        .ColAlignmentFixed(FinalAmt) = flexAlignRightCenter
        .ColWidth(FinalAmt) = 1065
        .ColWidth(ADItemCode) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel

DGSite.left = 4260: DGSite.top = mTopScale
DGForm.left = Me.width - (DGForm.width + mRtScale): DGForm.top = mTopScale
DGADItem.left = Me.width - (DGADItem.width + mRtScale): DGADItem.top = mTopScale
DGBook.left = 0: DGBook.top = FGrid.top: DGBook.width = Me.width - 90: DGBook.height = Me.height - (DGBook.top + mBotScale)
DGMod.left = 0: DGMod.width = Me.width - 90: DGMod.top = FGrid.top: DGMod.height = Me.height - (DGMod.top + mBotScale)
DgChassis.left = 0: DgChassis.width = Me.width - 90: DgChassis.top = FGrid.top: DgChassis.height = Me.height - (DgChassis.top + mBotScale)
DgSubvention.left = 0: DgSubvention.width = Me.width - 90: DgSubvention.top = FGrid.top: DgSubvention.height = Me.height - (DgSubvention.top + mBotScale)
With DgChassis
    .Columns(0).width = 1769.953
    .Columns(1).width = 2055.118
    .Columns(2).width = 1709.858
    .Columns(3).width = 1019.906
    .Columns(4).width = 1184.882
    .Columns(5).width = 1275.024
    .Columns(6).width = 929.7639
    .Columns(7).width = 1230.236
    .Columns(8).width = 0
    .Columns(9).width = 2039.811
End With
DGFin.left = Me.width - (DGFin.width + mRtScale): DGFin.top = mTopScale
DGCol.left = Me.width - (DGCol.width + mRtScale): DGCol.top = mTopScale

End Sub
Private Sub Disp_Text(Enb As Boolean)
On Error GoTo errlbl
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next



If UCase(left(PubComp_Name, 4)) = "ENAR" Then txt(SiteCode).Enabled = False

If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
    LblRebate = "Sp.Discount"
    LblSpDiscount = "Rebate"
End If



If TopCtrl1.TopText2 <> "Browse" Then
    If PubSiteWiseDisplayYn = 1 And PubSiteType <> "H" Then
        txt(SiteCode).Enabled = False
    Else
        txt(SiteCode).Enabled = True
    End If
End If
If TopCtrl1.TopText2 = "Edit" Then
    txt(SiteCode).Enabled = False
    'txt(VDate).Enabled = True
    txt(SerialNo).Enabled = False
    txt(InvPrefix).Enabled = False
    txt(BookNo).Enabled = False
    If GCn.Execute("select DelCh_DocId from veh_stock where chassisNo ='" & txt(ChassisNo) & "'").Fields(0).Value = "" Then
'        Txt(Model).Enabled = True
If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
        txt(ChassisNo).Enabled = False
      Else
        txt(ChassisNo).Enabled = True
      End If
    Else
'        Txt(Model).Enabled = False
'        Txt(ChassisNo).Enabled = False
    End If
End If
If TopCtrl1.TopText2 = "Add" Then
    txt(SerialNo).Enabled = True
    txt(InvPrefix).Enabled = True
    txt(Colours).Enabled = True
End If

txt(TxtDocID).Enabled = False
txt(Taxable).Enabled = False
txt(NamePrefix).Enabled = False
txt(Party).Enabled = False
txt(FNamePrefix).Enabled = False
txt(fname).Enabled = False
txt(Add1).Enabled = False
txt(Add2).Enabled = False
txt(Add3).Enabled = False
txt(City).Enabled = False
txt(Govt_YN).Enabled = False
txt(TelcoInvNo).Enabled = False
txt(TelcoInvDate).Enabled = False
txt(Model).Enabled = False
txt(EngineNo).Enabled = False

txt(NDP).Enabled = False
txt(SubTotA).Enabled = False
txt(OthFitAmt).Enabled = False
txt(OthFitTax).Enabled = False
txt(SubTotB).Enabled = False
'Txt(TaxPer).Enabled = False: Txt(TaxAmt).Enabled = False
txt(TaxSurPer).Enabled = False: txt(TaxSurch).Enabled = False
txt(ROff).Enabled = False
txt(GTotAmt).Enabled = False
txt(NetOStng).Enabled = False
txt(DelChNo).Enabled = False
txt(DelChDate).Enabled = False
txt(ModelDesc).Enabled = False

txt(FinAdd1).Enabled = False
txt(FinAdd2).Enabled = False

txtDisabled_Color Me


If GCn.Execute("Select " & vIsNull("RtoInsInBill", "0") & " From Syctrl").Fields(0) = 1 Then
    txt(RTOfee).Enabled = True
    txt(Insurance).Enabled = True
Else
    txt(RTOfee).Enabled = False
    txt(Insurance).Enabled = False
End If

If StrCmp(left(PubComp_Name, 4), "Yash") Then
    txtPrint(TempInvDate).Enabled = False
End If

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol

If PubSiebelActiveYn = 1 And pubUName = "SA" Or pubUName = "SANJAY1" Then
    cmdPost.Visible = True
Else
    cmdPost.Visible = False
End If
errlbl:
End Sub
Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGBook.Visible = True Then DGBook.Visible = False
    If DGFin.Visible = True Then DGFin.Visible = False
    If DGCol.Visible = True Then DGCol.Visible = False
    If DGMod.Visible = True Then DGMod.Visible = False
    If DgChassis.Visible = True Then DgChassis.Visible = False
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGADItem.Visible = True Then DGADItem.Visible = False
    If DGSite.Visible = True Then DGSite.Visible = False
    If DgSubvention.Visible = True Then DgSubvention.Visible = False
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ADItemCode) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, ADItem) <> RsADItem!Code Then
                RsADItem.MoveFirst
                RsADItem.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, ADItemCode) & "'"
            End If
        Case Qty, TaxPer1, TaxAmt1, TaxSurPer1, TaxSurAmt1
'            SendKeys "{HOME}+{END}"
     End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TxtGrid(0) = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case ADItem    '1
            DGridTxtKeyDown DGADItem, TxtGrid, Index, RsADItem, KeyCode, True, 1, frmVehAMDMast, "frmVehAMDMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, TaxSurAmt1
                End If
            End If
        Case Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, TaxSurAmt1
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
    Case ADItem
        If DGADItem.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsADItem, KeyAscii, "name"
    Case Rate, TaxAmt1, TaxSurAmt1
        Call NumPress(TxtGrid(Index), KeyAscii, 8, 2)
    Case TaxPer1, TaxSurPer1
        Call NumPress(TxtGrid(Index), KeyAscii, 2, 2)
    Case Qty
        Call NumPress(TxtGrid(Index), KeyAscii, 6, 0)
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
        Select Case FGrid.Col
            Case ADItem
                If KeyCode <> 13 And DGADItem.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsADItem, KeyCode, "name", True
            Case Qty
                FGrid.TextMatrix(FGrid.Row, Qty) = Format(Val(TxtGrid(Index).TEXT), "0")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxSurPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal False
            Case Rate
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                FGrid.TextMatrix(FGrid.Row, Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Qty))), "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxSurPer1)) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal False
            Case TaxAmt1
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, Amt)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                Else
                    FGrid.TextMatrix(FGrid.Row, TaxPer1) = Format((100 * Val(TxtGrid(Index).TEXT)) / Val(FGrid.TextMatrix(FGrid.Row, Amt)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal False
            Case TaxPer1
                FGrid.TextMatrix(FGrid.Row, TaxPer1) = TxtGrid(Index).TEXT
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = "0.00"
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal False
            Case TaxSurAmt1
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                If Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) = 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = "0.00"
                Else
                   FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = Format((100 * Val(TxtGrid(Index).TEXT)) / Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)), "0.00")
                End If
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal False
            Case TaxSurPer1
                FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = TxtGrid(Index).TEXT
                FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = Format(Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
                FGrid.TextMatrix(FGrid.Row, FinalAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) + Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) + Val(FGrid.TextMatrix(FGrid.Row, TaxSurAmt1))), "0.00")
                Amt_Cal False
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
Select Case FGrid.Col
    Case ADItem
        If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, ADItem) = ""
            FGrid.TextMatrix(FGrid.Row, ADItemCode) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, ADItemCode) = RsADItem!Code
            FGrid.TextMatrix(FGrid.Row, ADItem) = RsADItem!Name
            FGrid.TextMatrix(FGrid.Row, Rate) = Format(IIf(IsNull(RsADItem!Rate), 0, RsADItem!Rate), "0.00")
        End If
        If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
    Case Rate, TaxPer1, TaxAmt1, TaxSurPer1, TaxSurAmt1
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
        Amt_Cal False
    Case Qty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0")
        Amt_Cal False
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
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
'    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
    Grid_Hide
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
        Case ADItem
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case Qty
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Amt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, FinalAmt) = ""
        Case Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Amt) = ""
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, FinalAmt) = ""
        Case TaxPer1, TaxAmt1
            FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxPer1) = ""
        Case TaxSurPer1, TaxSurAmt1
            FGrid.TextMatrix(FGrid.Row, TaxSurPer1) = ""
            FGrid.TextMatrix(FGrid.Row, TaxSurAmt1) = ""
    End Select
    Amt_Cal False
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case ADItem, Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
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
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case ADItem
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case Amt
        FGrid.Col = FGrid.Col + 1
    Case Qty, Rate, TaxSurPer1, TaxSurAmt1, TaxPer1, TaxAmt1
       Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub Amt_Cal(YesNo As Boolean)

Dim I As Byte
Dim Tottax As Double
Dim TotAdd As Double, SubTotC As Double, TurnOverTax As Double
Dim STax, SubTotA1 As Double
Dim SAmt As Double
If UCase(txt(DeliveryFrom)) = "GODOWN" Then
    txt(Octroi) = "0.00"
End If
If PubVehRateIncTaxYn <> 1 Then
    txt(SaleRate) = txt(SubAmt)
    For I = 1 To FGrid.Rows - 1
       If FGrid.TextMatrix(I, ADItem) <> "" Then
            TotAdd = TotAdd + Val(FGrid.TextMatrix(I, Amt))
            Tottax = Tottax + Val(FGrid.TextMatrix(I, TaxAmt1)) + Val(FGrid.TextMatrix(I, TaxSurAmt1))
       End If
    Next
    txt(OthFitAmt) = IIf(TotAdd <> 0, Format(TotAdd, "0.00"), "")
    txt(OthFitTax) = IIf(Tottax <> 0, Format(Tottax, "0.00"), "")
    txt(SubTotA) = Format((Val(txt(SaleRate)) - Val(txt(Rebate)) - Val(txt(Subvention)) + Val(txt(IncCharge)) + Val(txt(Octroi)) + Val(txt(TempReg)) + Val(txt(TransIns)) + Val(txt(MVT)) + Val(txt(Transportation)) + Val(txt(HandlingCharges))), "0.00")
    If Val(txt(TaxPer)) > 0 Then
        txt(TaxAmt) = Format(Val(txt(SubTotA)) * Val(txt(TaxPer)) / 100, "0.00")
        txt(TaxSurch) = Format(Val(txt(TaxSurPer)) * Val(txt(TaxAmt)) / 100, "0.00")
    End If
    If Val(txt(SatPer)) > 0 Then
        txt(SatAmt) = Format(Val(txt(SubTotA)) * Val(txt(SatPer)) / 100, "0.00")
    End If
    txt(SubTotB) = Format((Val(txt(SubTotA)) + Val(txt(TaxAmt)) + Val(txt(SatAmt)) + Val(txt(TaxSurch)) + Val(txt(MisCharge))), "0.00")
    SubTotC = Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax))
    If PubSDTYN = 1 Then
        txt(TOTAmt) = Format(Val(txt(SubTotA)) * Val(txt(TOTPer)) / 100, "0.00")
    Else
        txt(TOTAmt) = Format(SubTotC * Val(txt(TOTPer)) / 100, "0.00")
    End If
    'txt(ROff) = dmRoundOff(Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)))
    txt(ROff) = dmRoundOff(SubTotC + Val(txt(TOTAmt)) + Val(txt(RTOfee)) + Val(txt(Insurance)))
    'txt(GTotAmt) = Format((Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)) + Val(txt(ROff))), "0.00")
    txt(GTotAmt) = Format((SubTotC + Val(txt(TOTAmt)) - Val(txt(FuelAmt)) + Val(txt(RTOfee)) + Val(txt(Insurance)) + Val(txt(ROff))), "0.00")
    txt(NetOStng) = Format((Val(txt(GTotAmt)) - Val(txt(AdvAmt))), "0.00")
    
Else
    If YesNo = True Then
        STax = Val(txt(TaxPer)) + Val(txt(TOTPer)) + Val(txt(SatPer)) + (Round(Val(txt(TaxPer)) * Val(txt(TaxSurPer)) / 100, 2))
        For I = 1 To FGrid.Rows - 1
           If FGrid.TextMatrix(I, ADItem) <> "" Then
                TotAdd = TotAdd + Val(FGrid.TextMatrix(I, Amt))
                Tottax = Tottax + Val(FGrid.TextMatrix(I, TaxAmt1)) + Val(FGrid.TextMatrix(I, TaxSurAmt1))
           End If
        Next
    
        txt(OthFitAmt) = IIf(TotAdd <> 0, Format(TotAdd, "0.00"), "")
        txt(OthFitTax) = IIf(Tottax <> 0, Format(Tottax, "0.00"), "")
    
        txt(SubTotA) = Format((Val(txt(SubAmt)) - Val(txt(Rebate)) - Val(txt(Subvention)) + Val(txt(IncCharge)) + Val(txt(Octroi)) + Val(txt(TempReg)) + Val(txt(TransIns)) + Val(txt(MVT)) + Val(txt(Transportation)) + Val(txt(HandlingCharges))), "0.00")
      '  SubTotA1 = SubTot - Val(Txt(Rebate))
        txt(SubTotA) = Format(Val(txt(SubTotA)) - (Val(txt(SubTotA)) * STax / (100 + STax)), "0.00")
        
        txt(SaleRate) = Format(Val(txt(SubTotA)) + Val(txt(Rebate)) + Val(txt(Subvention)) - Val(txt(IncCharge)) + Val(txt(Octroi)) + Val(txt(TempReg)) + Val(txt(TransIns)) + Val(txt(MVT)) + Val(txt(Transportation)) + Val(txt(HandlingCharges)), "0.00")
        'SAmt = Format(Val(Txt(SubTotA)) + Val(Txt(Rebate)) - Val(Txt(IncCharge)) + Val(Txt(Octroi)) + Val(Txt(TempReg)) + Val(Txt(TransIns)) + Val(Txt(MVT)) + Val(Txt(Transportation)), "0.00")
        
    
        If Val(txt(TaxPer)) > 0 Then
            txt(TaxAmt) = Format(Val(txt(SubTotA)) * Val(txt(TaxPer)) / 100, "0.00")
            txt(TaxSurch) = Format(Val(txt(TaxSurPer)) * Val(txt(TaxAmt)) / 100, "0.00")
        End If
        If Val(txt(SatPer)) > 0 Then
            txt(SatAmt) = Format(Val(txt(SubTotA)) * Val(txt(SatPer)) / 100, "0.00")
        End If
    
        txt(SubTotB) = Format((Val(txt(SubTotA)) + Val(txt(TaxAmt)) + Val(txt(SatAmt)) + Val(txt(TaxSurch)) + Val(txt(MisCharge))), "0.00")
        SubTotC = Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax))
        If PubSDTYN = 1 Then
            txt(TOTAmt) = Format(Val(txt(SubTotA)) * Val(txt(TOTPer)) / 100, "0.00")
        Else
            txt(TOTAmt) = Format(SubTotC * Val(txt(TOTPer)) / 100, "0.00")
        End If
        'txt(ROff) = dmRoundOff(Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)))
        txt(ROff) = dmRoundOff(SubTotC + Val(txt(TOTAmt)) + Val(txt(RTOfee)) + Val(txt(Insurance)))
        'txt(GTotAmt) = Format((Val(txt(SubTotB)) + Val(txt(OthFitAmt)) + Val(txt(OthFitTax)) - Val(txt(FuelAmt)) + Val(txt(ROff))), "0.00")
        txt(GTotAmt) = Format((SubTotC + Val(txt(TOTAmt)) - Val(txt(FuelAmt)) + Val(txt(RTOfee)) + Val(txt(Insurance)) + Val(txt(ROff))), "0.00")
        txt(NetOStng) = Format((Val(txt(GTotAmt)) - Val(txt(AdvAmt))), "0.00")
    End If


End If

End Sub

Private Function Fill_Data(Enb As Boolean) As Boolean
Dim Rst As ADODB.Recordset
Dim Margin As Double
If Enb Then
    If txt(Model) <> RsChassis!Model Then
        If MsgBox("Model changed Continue Yes/No ? ", vbYesNo + vbCritical + vbDefaultButton2, "Check") = vbNo Then
            GoTo NXT
        Else
            txt(Model) = RsChassis!Model
            txt(Model).Tag = RsChassis!Model
            If txt(Model).Tag <> "" Then
                txt(ModelDesc) = GCn.Execute("Select " & xIsNull("Sales_Desc", "") & " From Model Where Model='" & txt(Model).Tag & "'  ").Fields(0).Value
            End If
        End If
    End If
    If IsNull(RsChassis!InDate) Then
        If MsgBox("Vehicle In Transit Continue Yes/No ? ", vbYesNo + vbCritical + vbDefaultButton2, "Check") = vbNo Then
            GoTo NXT
        End If
    End If
    txt(EngineNo) = IIf(IsNull(RsChassis!EngineNo), "", RsChassis!EngineNo)
    txt(SrvBookNo) = IIf(IsNull(RsChassis!Srv_BookNo), "", RsChassis!Srv_BookNo)
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        If txt(Colours) = "" Then
            txt(Colours) = IIf(IsNull(RsChassis!Col_Desc), "", RsChassis!Col_Desc)
            txt(Colours).Tag = IIf(IsNull(RsChassis!Colour_Code), "", RsChassis!Colour_Code)
        End If
    Else
        txt(Colours) = IIf(IsNull(RsChassis!Col_Desc), "", RsChassis!Col_Desc)
        txt(Colours).Tag = IIf(IsNull(RsChassis!Colour_Code), "", RsChassis!Colour_Code)
    End If
    txt(TelcoInvDate).TEXT = IIf(IsNull(RsChassis!PBILL_DATE), "", RsChassis!PBILL_DATE)
    txt(TelcoInvNo).TEXT = IIf(IsNull(RsChassis!PBILL_NO), "", RsChassis!PBILL_NO)
    txt(Taxable).TEXT = IIf(RsChassis!Tax_YN = 1, "Yes", "No")
    txt(NDP).TEXT = IIf(IsNull(RsChassis!vrate) Or RsChassis!vrate = 0, "", Format(RsChassis!vrate, "0.00"))
    Set Rst = New Recordset
    Rst.Open "Select P_RATE,s_rate,INCI_CHRG,OCTROI,REG_TEMP,INS_TRN,TRANSPORT,MVT,REG_FEE,INS_FEE, Margine, HandlingCharges, Reg_Fee,Reg_FeeCom, Ins_Fee from veh_rate where model = '" & txt(Model).TEXT & "' and Effective_Date <= " & ConvertDate(txt(VDate)) & " and RSO_WORK = " & VNull(RsChassis!RSO_WORK) & " and TAXABLE_YN = " & RsChassis!Tax_YN & " order by Effective_Date DESC", GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        If UCase(left(PubComp_Name, 3)) = "LMP" Then
            txt(SubAmt) = VNull(Rst!p_rate) + VNull(Rst!Margine)
            If Val(txt(SubAmt)) = 0 Then
                MsgBox "Rates Not Defined in Vehicle Rate Declration!"
            End If
        Else
            If Val(txt(NDP).TEXT) = 0 Then
               txt(NDP).TEXT = IIf(IsNull(Rst!p_rate) Or Rst!p_rate = 0, "", Format(Rst!p_rate, "0.00"))
            End If
            Margin = IIf(IsNull(Rst!S_Rate), 0, Rst!S_Rate) - IIf(IsNull(Rst!p_rate), 0, Rst!p_rate)
            'Txt(SaleRate).TEXT = Format((Val(Txt(NDP)) + Margin), "0.00")
            txt(SaleRate).TEXT = Format(IIf(IsNull(Rst!S_Rate), 0, Rst!S_Rate), "0.00")
        End If
    Else
        If UCase(left(PubComp_Name, 3)) = "LMP" Then
            MsgBox "Rates Not Defined in Vehicle Rate Declation Master For Given Criteria "
        Else
            Margin = 0
            'Txt(SaleRate).TEXT = Format((Val(Txt(NDP)) + Margin), "0.00")
            txt(SubAmt).TEXT = Format((Val(txt(NDP)) + Margin), "0.00")
        End If
    End If
    If txt(Model) <> "" And (Val(txt(SubAmt)) = 0 Or left(UCase(PubComp_Name), 6) = "J.M.A.") Then
        txt(SubAmt) = Format(VNull(GCn.Execute("SELECT Sale_Rate from Model where model='" & txt(Model) & "' and (div_code='" & PubDivCode & "' or Div_Code='')").Fields(0).Value), "0.00")
    End If
    
    If txt(Model) <> "" And (Val(txt(SubAmt)) = 0 Or left(UCase(PubComp_Name), 4) = "ENAR") Then
        txt(SubAmt) = Format(VNull(GCn.Execute("SELECT Sale_Rate from Model where model='" & txt(Model) & "' and (div_code='" & PubDivCode & "' or Div_Code='')").Fields(0).Value), "0.00")
    End If
    If Rst.RecordCount > 0 Then
        txt(IncCharge) = IIf(IsNull(Rst!INCI_CHRG) Or Rst!INCI_CHRG = 0, "", Format(Rst!INCI_CHRG, "0.00"))
        txt(Octroi) = IIf(IsNull(Rst!Octroi) Or Rst!Octroi = 0, "", Format(Rst!Octroi, "0.00"))
        txt(TempReg) = IIf(IsNull(Rst!REG_TEMP) Or Rst!REG_TEMP = 0, "", Format(Rst!REG_TEMP, "0.00"))
        txt(TransIns) = IIf(IsNull(Rst!INS_TRN) Or Rst!INS_TRN = 0, "", Format(Rst!INS_TRN, "0.00"))
        txt(MVT) = IIf(IsNull(Rst!MVT) Or Rst!MVT = 0, "", Format(Rst!MVT, "0.00"))
        txt(Transportation) = IIf(IsNull(Rst!Transport) Or Rst!Transport = 0, "", Format(Rst!Transport, "0.00"))
        txt(HandlingCharges) = IIf(IsNull(Rst!HandlingCharges) Or Rst!HandlingCharges = 0, "", Format(Rst!HandlingCharges, "0.00"))
        If GCn.Execute("Select " & vIsNull("Purpose", "0") & " From Veh_Order Where OrdDocId='" & txt(BookNo).Tag & "'").Fields(0).Value = 3 Then
            txt(RTOfee) = IIf(IsNull(Rst!REG_FEECom) Or Rst!REG_FEECom = 0, "", Format(Rst!REG_FEECom, "0.00"))
        Else
            txt(RTOfee) = IIf(IsNull(Rst!REG_FEE) Or Rst!REG_FEE = 0, "", Format(Rst!REG_FEE, "0.00"))
        End If
        
        txt(Insurance) = IIf(IsNull(Rst!INS_FEE) Or Rst!INS_FEE = 0, "", Format(Rst!INS_FEE, "0.00"))
    Else
        txt(IncCharge) = ""
        txt(Octroi) = ""
        txt(TempReg) = ""
        txt(TransIns) = ""
        txt(MVT) = ""
        txt(Transportation) = ""
        txt(HandlingCharges) = ""
        txt(RTOfee) = ""
        txt(Insurance) = ""
    End If
    Amt_Cal False
    Fill_Data = True
    Exit Function
End If
Fill_Data = True
NXT:
    txt(EngineNo) = ""
    txt(SrvBookNo) = ""
    txt(Colours) = ""
    txt(Colours).Tag = ""
    txt(TelcoInvDate).TEXT = ""
    txt(TelcoInvNo).TEXT = ""
    txt(Taxable).TEXT = ""
    txt(NDP).TEXT = ""
    txt(SaleRate).TEXT = ""
    Amt_Cal False
End Function

Private Function FillRecords() As Boolean
On Error GoTo error1
Dim Rst As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim RsBooking  As ADODB.Recordset
    Set RsBooking = New Recordset
    RsBooking.CursorLocation = adUseClient
    RsBooking.Open "SELECT Veh_Order.OrdDocID,Veh_Order.Inv_DocId,Veh_Order.PartyCode, Veh_Order.Ord_Date , " & _
    "Veh_Order.GOVT_YN, Veh_Order.MODEL, Veh_Order.Chassis, Veh_Order.Srv_BookNo, Veh_Order.RATE, Veh_Order.Fund_Source, Veh_Order.FB_CODE, Veh_Order.FIN_AcCode, Veh_Order.FIN_AcCode, Veh_Order.Colour_Code, Veh_Order.FIN_AMT " & _
    "FROM Veh_Order " & _
    "where OrdDocid = '" & txt(BookNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
    
    If RsBooking.RecordCount = 0 Then
        MsgBox "Booking No. Not Exist", vbInformation, "Booking Not Found"
        txt(NamePrefix).TEXT = ""
        txt(Party).TEXT = ""
        txt(Party).Tag = ""
        txt(FNamePrefix).TEXT = ""
        txt(fname).TEXT = ""
        txt(Add1).TEXT = ""
        txt(Add2).TEXT = ""
        txt(Add3).TEXT = ""
        txt(City).TEXT = ""
        txt(Model).TEXT = ""
        txt(ModelDesc) = ""
        txt(Govt_YN).TEXT = ""
        txt(ChassisNo).TEXT = ""
        txt(Colours).Tag = ""
        txt(Colours).TEXT = ""
        txt(SrvBookNo).TEXT = ""
        txt(NDP).TEXT = ""
        txt(FundSource).TEXT = ""
        txt(FB_Code).TEXT = ""
        txt(FB_Code).Tag = ""
        FinAcCode = ""
        txt(FinAmt).TEXT = ""
        txt(BookNo).Tag = ""
        txt(BookNo).SetFocus
        Set RsBooking = Nothing
        FillRecords = False
        Exit Function
    Else
        If RsBooking!Inv_DocId <> Null Or RsBooking!Inv_DocId <> "" Then
            MsgBox "Invoice Exist Against Booking No", vbInformation, "Validation Check"
            txt(NamePrefix).TEXT = ""
            txt(Party).TEXT = ""
            txt(Party).Tag = ""
            txt(FNamePrefix).TEXT = ""
            txt(fname).TEXT = ""
            txt(Add1).TEXT = ""
            txt(Add2).TEXT = ""
            txt(Add3).TEXT = ""
            txt(City).TEXT = ""
            txt(Model).TEXT = ""
            txt(ModelDesc) = ""
            txt(Govt_YN).TEXT = ""
            txt(ChassisNo).TEXT = ""
            txt(Colours).Tag = ""
            txt(Colours).TEXT = ""
            txt(SrvBookNo).TEXT = ""
            txt(NDP).TEXT = ""
            txt(FundSource).TEXT = ""
            txt(FB_Code).TEXT = ""
            txt(FB_Code).Tag = ""
            FinAcCode = ""
            txt(FinAmt).TEXT = ""
            txt(BookNo).Tag = ""
            txt(BookNo).SetFocus
            Set RsBooking = Nothing
            FillRecords = False
            Exit Function
        End If
        txt(AdvAmt) = Format(PartyAdvance(RsBooking!OrdDocId, txt(VDate)), "0.00")
        txt(Party).Tag = IIf(IsNull(RsBooking!PartyCode), "", RsBooking!PartyCode)
        If txt(Party).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select NamePrefix,name,FPrefix,FName,add1,add2,add3,CityCode, LstNo from SubGroup where Subcode = '" & txt(Party).Tag & "'", GCn, adOpenStatic, adLockReadOnly
            txt(NamePrefix).TEXT = IIf(IsNull(Rst!NamePrefix), "", Rst!NamePrefix)
            txt(Party).TEXT = Rst!Name
            txt(FNamePrefix).TEXT = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
            txt(fname).TEXT = IIf(IsNull(Rst!fname), "", Rst!fname)
            txt(Add1).TEXT = IIf(IsNull(Rst!Add1), "", Rst!Add1)
            txt(Add2).TEXT = IIf(IsNull(Rst!Add2), "", Rst!Add2)
            txt(Add3).TEXT = IIf(IsNull(Rst!Add3), "", Rst!Add3)
            txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
            If txt(City).Tag <> "" Then
                txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & txt(City).Tag & "'").Fields(0).Value
            End If
            txt(RTO).TEXT = txt(City).TEXT
            
'''''            mPartyLstNo = XNull(Rst!LstNo)
'''''            If mPartyLstNo <> "" Then
'''''                If PubVehTaxInvPrefix <> "" Then
'''''                    LblVPrefix = PubVehTaxInvPrefix
'''''                    txt(InvPrefix) = PubVehTaxInvPrefix
'''''                End If
'''''            Else
'''''                Set RsTemp = G_FaCn.Execute("Select Prefix From VehBill_Counter Where V_Type = '" & mVType & "' And Prefix<>'" & PubVehTaxInvPrefix & "' And Div_Code='" & PubDivCode & "' And Date_From<= " & ConvertDate(txt(VDate)) & " ")
'''''                If RsTemp.RecordCount > 0 Then
'''''                    LblVPrefix = XNull(RsTemp(0))
'''''                    txt(InvPrefix) = XNull(RsTemp(0))
'''''                End If
'''''            End If
'''''            txt(SerialNo) = ""
'''''            txt(TxtDocID) = GetDocIDVBill(GCnFaV, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix, txt(SiteCode).Tag)
'''''            DocID = txt(TxtDocID)
            
        End If
        txt(Model).TEXT = RsBooking!Model
        txt(Model).Tag = RsBooking!Model
        If txt(Model).Tag <> "" Then
            txt(ModelDesc) = GCn.Execute("Select " & xIsNull("Sales_Desc", "") & " From Model Where Model='" & txt(Model).Tag & "'  ").Fields(0).Value
        End If

        txt(Govt_YN).TEXT = IIf(IsNull(RsBooking!Govt_YN), "", RsBooking!Govt_YN)
        txt(Colours).Tag = IIf(IsNull(RsBooking!Colour_Code), "", RsBooking!Colour_Code)
        If txt(Colours).Tag <> "" Then
            txt(Colours).TEXT = GCn.Execute("select col_desc from colmast where col_code = '" & txt(Colours).Tag & "'").Fields(0).Value
        End If
        Select Case RsBooking!Fund_Source
            Case 0 '0 Hypothecation ,1 Hire purchase ,2 Own Fund,3 Lease
                txt(FundSource).TEXT = "Hypothecation"
            Case 1
                txt(FundSource).TEXT = "Hire Purchase"
            Case 3
                txt(FundSource).TEXT = "Lease"
            Case 4
                txt(FundSource).TEXT = "Agreement"
            Case 5
                txt(FundSource).TEXT = "Lease & Agreement"
            Case 6
                txt(FundSource).TEXT = "Loan Cum Hypt."
            Case Else
                txt(FundSource).TEXT = "Own Fund"
        End Select
        txt(FB_Code).Tag = IIf(IsNull(RsBooking!FB_Code), "", RsBooking!FB_Code)
        If txt(FB_Code).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select fincode as code,finname + ',' + " & xIsNull("City.CityName", "") & " as name,AcCode,FinBankCode " & _
                    " from ContractFinance left join city on left(ContractFinance.City,4)=City.CityCode " & _
                    " where fincatg = 0 and  fincode = '" & txt(FB_Code).Tag & "'", GCn, adOpenStatic, adLockReadOnly
            txt(FB_Code).TEXT = Rst!Name
            FinAcCode = IIf(IsNull(Rst!AcCode), "", Rst!AcCode)
        Else
            txt(FB_Code).TEXT = ""
            FinAcCode = ""
        End If
        txt(FinAmt).TEXT = IIf(IsNull(RsBooking!Fin_Amt), "", RsBooking!Fin_Amt)
        
    End If
Set Rst = Nothing
FillRecords = True
error1:
    CheckError

End Function
'************************ PRINTING CODE ******************

Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case DocType
        ListArray = Array("Sale Bill", "Sale Certificate", "Form22", "Form22A", "Declaration")
        Set mListItem = ListView_Items(ListView, txtPrint, Index, ListArray, 5)
End Select
End Sub

Private Sub TxtPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case DocType
        ListView_KeyDown FrmList, ListView, txtPrint, Index, KeyCode, Shift, FrmPrn.left + txtPrint(Index).left, (FrmPrn.top + txtPrint(Index).top + txtPrint(Index).height), txtPrint(Index).width, 1200
End Select
If FrmList.Visible = False And DGSite.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> TempInvDate Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case CertiTempYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            txtPrint(Index) = "Yes"
            If StrCmp(left(PubComp_Name, 4), "Yash") Then
                txtPrint(RTOName) = PubComp_City
            End If
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txtPrint(Index) = "No"
            If StrCmp(left(PubComp_Name, 4), "Yash") Then
                txtPrint(RTOName) = txt(RTO)
            End If
            
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txtPrint(Index) = ""
        End If
        KeyAscii = 0
        FldEnabled1 (IIf(txtPrint(Index) = "Yes", True, False))
    Case WtPrn
        If UCase(Chr(KeyAscii)) = "Y" Then
            txtPrint(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txtPrint(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txtPrint(Index) = ""
        End If
        KeyAscii = 0
    Case Seet
       Call NumPress(txtPrint(Index), KeyAscii, 2, 0)
End Select
'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub txtPrint_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case DocType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txtPrint, Index, KeyCode, mListItem
End Select
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
  Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case DocType
        If txtPrint(Index).TEXT <> "" Then txtPrint(Index).TEXT = ListView.SelectedItem.TEXT
        If txtPrint(Index).TEXT = "Sale Certificate" Then FldEnabled True Else FldEnabled False
    Case TempInvDate, CertiPrnDate
        txtPrint(Index).TEXT = RetDate(txtPrint(Index))
End Select
End Sub

Private Sub FldEnabled(Enb As Boolean)
    txtPrint(RTOName).Enabled = Enb
    txtPrint(CertiPrnDate).Enabled = Enb
    txtPrint(CertiTempYN).Enabled = Enb
    txtPrint(Seet).Enabled = Enb
    txtPrint(Body).Enabled = Enb
    txtPrint(Narr).Enabled = Enb
    txtPrint(WtPrn).Enabled = Enb
    If Enb = False Then
        txtPrint(CertiPrnDate).TEXT = ""
        txtPrint(CertiTempYN).TEXT = ""
        txtPrint(Seet).TEXT = ""
        txtPrint(Body).TEXT = ""
        txtPrint(Narr).TEXT = ""
    End If
End Sub
Private Sub FldEnabled1(Enb As Boolean)
    'txtPrint(Seet).Enabled = Enb
    txtPrint(Body).Enabled = Enb
    txtPrint(Narr).Enabled = Enb
    txtPrint(WtPrn).Enabled = Enb
    txtPrint(RTOName).Enabled = Enb
    If Enb = False Then
        'txtPrint(Seet).TEXT = ""
        txtPrint(Body).TEXT = ""
        txtPrint(Narr).TEXT = ""
        txtPrint(WtPrn).TEXT = ""
    End If
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
'*****For Calcelled bill Dos printing
If CancelBillY_N = True Then
    txtPrint(DocType) = "Sale Bill"
    Index = 2
End If
'*******
If IsValid(txtPrint(DocType), "Print Document") = False Then Exit Sub
'"Sale Bill", "Sale Certificate", "Form22", "Form22A", "Declaration"
Select Case txtPrint(DocType)
    Case "Sale Bill"
        'mRepName = IIf(OptPlain.Value = True, "VehSale", "VehSale")
        'modi lps 15-04-2003
        mRepName = GCn.Execute("Select VBilRptName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
        GSQL = "SELECT VO.OrdDocID,Ord_No,Ord_Date,VO.Model,VO.PartyCode,VO.RATE,VO.Fund_Source,VO.FIN_YN,VO.FB_CODE,VO.FIN_AMT,VO.Inv_DocId,VO.Inv_VType,VO.Inv_No,VO.Inv_Date," & _
            " VO.Form_Code,VO.TrnType_Prn,VO.VRATE,VO.MARGINE,VO.REBATE,VO.InciChrg,VO.Octroi,VO.RegTemp,VO.TransitInsu,VO.Transport,VO.MVT,VO.TAX_Per,VO.TAX_Amt,VO.Surcharge_Per," & _
            " VO.Surcharge_Amt,VO.TOT_Per,VO.TOT_Amt,VO.OtherChrg,VO.FIT_AMT,VO.FIT_TAX,VO.Round_off,VO.DieselAmt,VO.BillPrn_YN,VO.DETAILS_YN,VO.INS_FEE,VO.INS_NOTE,VO.S_CHARGE,VO.RoundOff_YN,VO.Net_Amount," & _
            " VO.Inv_UName,VO.Inv_UEntDt,Veh_Purch1.gate,CF.finname,City_1.CityName as fincity, CF.Add1 as finadd1,CF.Add2 as finadd2,FinBank.FinBankName,site.site_desc,VStk.Pur_DocId," & _
            " VStk.Sal_DocId,VStk.ChassisNo, VStk.EngineNo, VStk.PBILL_NO, VStk.PBILL_DATE, " & _
            " M.Model_Desc,M.Model_Desc1,M.Sales_Desc, ColMast.Col_Desc,SG.NamePrefix, SG.Name, " & _
            " " & cIIF("SG.Add1 Is Null or SG.Add1=''", "SG.TAdd1", "SG.Add1") & " as PAdd1, " & _
            " " & cIIF("SG.Add2 Is Null or SG.Add1=''", "SG.TAdd2", "SG.Add2") & " as PAdd2, " & _
            " " & cIIF("SG.Add3 Is Null or SG.Add1=''", "SG.TAdd3", "SG.Add3") & " as PAdd3, " & _
            " " & cIIF("City.CityName Is Null or City.CityName=''", "TCity.CityName", "City.CityName") & " as PCityName, " & _
            " SG.FPrefix,SG.FName,M.WHEELBASE,M.RIMS,M.TYRES,M.TyreDetails,M.GearBoxNo,TaxForms.Printing_Desc,SG.PANNo, " & ConvertDate(txtPrint(TempInvDate)) & " as InvDate,VO.TOT_Per,VO.TOT_Amt,TaxForms.Printing_desc as FormName,M.Model as Modl,M.SALES_DESC,Emp_Mast.Emp_Name,VO.Misc_Info, VO.Subvention, VO.HandlingCharges, VO.RTOFee, VO.Insurance, VP.OBNo As OthDlrBillNo, VP.OBDate As OthDlrBillDate, VO.DeliveryFrom, SG.Phone, SG.Mobile, VO.Inv_Date As SaleBillDate, Sg.LstNo As PartyLstNo, TaxForms.L_C, VO.SatPer, VO.SatAmt,SG.PIN as PinNo " & _
            " FROM (((((((((((((Veh_Order as VO LEFT JOIN Veh_Stock as VStk ON VO.Inv_DocId = VStk.Sal_DocId) " & _
            " LEFT JOIN TaxForms ON VO.Form_Code = TaxForms.Form_Code) " & _
            " LEFT JOIN ColMast ON VO.Colour_Code = ColMast.Col_Code) " & _
            " LEFT JOIN Model as M ON VO.MODEL = M.MODEL) " & _
            " LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
            " LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
            " LEFT JOIN City TCity ON SG.TCityCode = TCity.CityCode) " & _
            " LEFT JOIN ContractFinance as CF ON VO.FB_CODE = CF.FinCode) " & _
            " LEFT JOIN Site ON right(VO.Inv_SiteCode,1) = Site.Site_Code) Left Join Veh_Purch1 VP On VStk.Pur_DocId=VP.DocId ) " & _
            " LEFT JOIN FinBank ON CF.FinBankCode = FinBank.FinBankCode) " & _
            " LEFT JOIN City AS City_1 ON CF.City = City_1.CityCode) " & _
            " LEFT JOIN Veh_Purch1 ON VStk.Pur_DocId = Veh_Purch1.DocID)" & _
            " LEFT JOIN Emp_Mast ON VO.Rep_Code = Emp_Mast.Emp_Code  " & _
            " where VO.Inv_DocId = '" & Master!SearchCode & "'"
    Case "Sale Certificate"
        'Check For GN Motors Kota For Preprinted Stationary Printing
        If left(PubComp_Name, 10) = "GANGANAGAR" Then
            mRepName = IIf(OptPlain.Value = True, "VehSaleCertKOTA", "VehSaleCertKOTA")
        Else
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert", "VehSaleCert")
        End If
        GSQL = "SELECT VO.INTD_USE, VP1.Tot_Amount,VStk.vrate, VStk.Mfg_Month ,VStk.Mfg_Yr, " & _
            "VO.CertiPrn_YN, VO.TCertiPrn_YN,SG.FPrefix, SG.FName, city_1.cityname as TCity,SG.TAdd1, " & _
            "SG.TAdd2, SG.TAdd3,SG.TPIN, VO.Inv_DocId, " & ConvertDate(IIf(txtPrint(CertiPrnDate) <> "", txtPrint(CertiPrnDate), txtPrint(TempInvDate))) & " as Inv_Date, VO.Fund_Source, VO.P_AMOUNT, " & _
            "VO.DelCh_DT, '" & txtPrint(RTOName) & "' as RTO, Model_Grp.ModelGrp_Name, city.CityName, Fincity.cityname as FinCity," & _
            "SG.Name, ColMast.Col_Desc, M.MODEL, M.Vehicle_Type, " & _
            "M.Model_Desc, M.Model_Desc1, M.Model_Desc2,M.Wheel_Catg, " & _
            "M.TYRES, M.TYRE_F, M.TYRE_M, M.TYRE_R,M.TYRE_FS, M.TYRE_MS, M.TYRE_RS, " & _
            "M.TyreDetails, M.SEAT, M.RLW, M.HORSEPOWER, M.FRONT_A_WT, M.REAR_A_WT," & _
            "M.UNLADEN_WT, M.GROSS_WT, M.WHEELBASE, M.CYLINDER, M.FUEL, M.TRADE_NO, M.Manufacturer, " & _
            "VStk.ChassisNo, VStk.EngineNo,Finbank.FinBankName,CF.FinName, CF.Add1 as FAdd1," & _
            "CF.add2 as Fadd2, CF.PinCode as FPin , SG.Add1, SG.Add2, SG.Add3, SG.PIN ,vo.Inv_UName,vo.Inv_UEntDt,M.Model as Modl,M.Sales_Desc,purpose.Purposename,'" & txtPrint(9) & "' as TCNO,M.GearBoxNo, VO.HandlingCharges,M.BodyType as Body,SG.Phone AS PartyPhone ,SG.Mobile AS PartyMobile " & _
            "FROM (((((((((((Veh_Order VO " & _
            "LEFT JOIN veh_Stock VStk ON VO.Inv_DocId = VStk.Sal_DocId) " & _
            "LEFT JOIN ColMast ON Vo.Colour_Code = ColMast.Col_Code) " & _
            "LEFT JOIN Veh_Purch1 VP1 ON VStk.Pur_DocId = VP1.DocID) " & _
            "LEFT JOIN ContractFinance CF ON VO.FB_CODE = CF.FinCode) " & _
            "LEFT JOIN Subgroup SG ON VO.PartyCode = SG.SubCode) " & _
            "LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
            "LEFT JOIN Model_Grp ON M.Grp_Code = Model_Grp.ModelGrp_Code) " & _
            "LEFT JOIN city AS fincity ON CF.City = fincity.CityCode) " & _
            "LEFT JOIN Finbank ON CF.FinBankCode = Finbank.FinBankCode) " & _
            "LEFT JOIN city ON SG.CityCode = city.CityCode) " & _
            "LEFT JOIN City AS city_1 ON SG.TCityCode = city_1.CityCode) " & _
            "LEFT JOIN Purpose ON VO.purpose = purpose.purposecode " & _
            "where VO.Inv_DocId = '" & Master!SearchCode & "' "
            If StrCmp(left(PubComp_Name, 4), "Yash") Then
                GSQL = GSQL & " and (VO.DelCh_docid <> Null Or VO.DelCh_DocId<>'') "
            End If
    Case "Form22A", "Form22"
        If txtPrint(DocType) = "Form22" Then
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert22", "VehSaleCert22")
        Else
            mRepName = IIf(OptPlain.Value = True, "VehSaleCert22A", "VehSaleCert22A")
        End If
        GSQL = "SELECT M.Manufacturer, D.MfgAdd1,D.MfgAdd2,D.MfgAdd3," & _
            "M.MODEL,M.Chas_Type,M.Vehicle_Type, M.Sales_Desc," & _
            "M.Model_Desc, M.Model_Desc1, M.Model_Desc2, " & _
            "M.WHEELBASE,M.Fuel, VStk.ChassisNo , VStk.EngineNo,vo.Inv_UName,vo.Inv_UEntDt " & _
            "FROM ((veh_order as VO LEFT JOIN veh_Stock as VStk ON VO.Inv_DocId = VStk.Sal_DocId) " & _
            "LEFT JOIN Model as M ON VO.MODEL = M.MODEL) " & _
            "Left Join Division as D on D.Div_Code=left(VO.Inv_DocId,1) " & _
            " where VO.Inv_DocId = '" & Master!SearchCode & "' " 'and (VO.DelCh_docid <> Null Or VO.DelCh_DocId<>'')"
End Select
Select Case Index
    Case PScreen, PWindows
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        If txtPrint(DocType) = "Sale Certificate" Then
            SpeedPrintCerti GSQL
        ElseIf txtPrint(DocType) = "Form22A" Then
            SpeedPrint22A GSQL
        ElseIf txtPrint(DocType) = "Form22" Then
            SpeedPrint22 GSQL
        ElseIf txtPrint(DocType) = "Sale Bill" Then
            If UCase(left(PubComp_Name, 5)) = "SANYA" Then
                SpeedPrintInvSanya GSQL
            ElseIf UCase(left(PubComp_Name, 7)) = "SHANKAR" Or UCase(left(PubComp_Name, 6)) = "MAURYA" Then
                SpeedPrintInvSHANKAR GSQL
            ElseIf UCase(left(PubComp_Name, 5)) = "SOCIE" Then
                SpeedPrintInvSOCIETY GSQL
            ElseIf UCase(left(PubComp_Name, 3)) = "JMK" Then
                SpeedPrintInvJMK GSQL
            ElseIf UCase(left(PubComp_Name, 6)) = "J.M.A." Then
                SpeedPrintInvJMA GSQL
            Else
                SpeedPrintInv GSQL
            End If
        Else
            SpeedPrintDeclar
        End If
        FrmPrn.Visible = False
        
    Case PSetUp
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
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset, RstSub2 As ADODB.Recordset
Dim I As Integer, Cnt As Integer, Foot1 As String, Foot2 As String, Foot3 As String, Foot4 As String
Dim Foot5 As String, Foot6 As String, Foot7 As String, Foot8 As String, Foot9 As String
Dim RstCompDet As ADODB.Recordset, j As Integer, Footer As String
Dim Rst2 As ADODB.Recordset
Dim tmprs As ADODB.Recordset
Dim PrnTitle As String, HlpLineNo$
On Error GoTo ERRORHANDLER



HlpLineNo = ""
Set tmprs = GCn.Execute("Select HelpLineNo from Syctrl")
If tmprs.RecordCount > 0 Then
    HlpLineNo = IIf(IsNull(tmprs!HelpLineNo), "", Trim(tmprs!HelpLineNo))
    Set tmprs = Nothing
End If




If txtPrint(DocType) = "Sale Certificate" Then
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    Set RstCompDet = GCn.Execute("select V_SecPAN_No,V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("CompPanNo")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecPAN_No & "'"
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecSpeciality & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecGram & "'"
            Case UCase("RTOName")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(RTOName) & "'"
            Case UCase("PrnDate")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(CertiPrnDate) & "'"
            Case UCase("TempYN")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(CertiTempYN) & "'"
            Case UCase("Seet")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(Seet) & "'"
            Case UCase("Body")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(Body) & "'"
            Case UCase("Narr")
               ' rpt.FormulaFields(I).TEXT = "'" & txtPrint(Narr) & "'"
            Case UCase("WtPrn")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(WtPrn) & "'"
            
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Select Case Index
        Case PWindows
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
                If txtPrint(CertiTempYN) = "Yes" Then
                    GCn.Execute "update veh_order set TCertiPrn_YN = 1  where where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
                Else
                    GCn.Execute "update veh_order set CertiPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "' "
                End If
            End If
            Set Rst = Nothing
            Set RstCompDet = Nothing
            Set rpt = Nothing
        Case PScreen 'screen
            Call Report_View(rpt, Me.CAPTION, , True)
            Set Rst = Nothing
            Set RstCompDet = Nothing
    End Select
ElseIf txtPrint(DocType) = "Form22A" Or txtPrint(DocType) = "Form22" Then
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
'            Case UCase("PrnDate")
'                rpt.FormulaFields(i).Text = "'" & txtPrint(PrnDate) & "'"
            Case UCase("Narr")
                rpt.FormulaFields(I).TEXT = "'" & txtPrint(Narr) & "'"
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Select Case Index
        Case PWindows
            rpt.PrintOut False
            Set Rst = Nothing
'            Set Rst1 = Nothing
            Set rpt = Nothing
        Case PScreen 'screen
            Call Report_View(rpt, Me.CAPTION, , True)
            Set Rst = Nothing
'            Set Rst1 = Nothing
    End Select
Else    'Sale Bill
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    'Recordset is made for subreport1
    mQry = "SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
    "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    "where Veh_Purch2.DocId = '" & Master!SearchCode & "'"
    
    Set RstSub1 = New Recordset
    RstSub1.CursorLocation = adUseClient
    RstSub1.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

   'Recordset is made for subreport2
   
    mQry = "SELECT Veh_Purch2.Trn_Type, Veh_Purch2.DocID, Veh_Purch2.QTY, Veh_Purch2.RATE, Veh_AMDModel.Prod_Name " & _
    "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code where veh_stock.Chassisno = '" & txt(ChassisNo) & "'"
        
    Set RstSub2 = New Recordset
    RstSub2.CursorLocation = adUseClient
    RstSub2.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
'modi LPS 15-04-2003
'    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
'    CreateFieldDefFile RstSub1, PubRepoPath + "\" & mRepName & "1.ttx", True
'    CreateFieldDefFile RstSub2, PubRepoPath + "\" & mRepName & "2.ttx", True
    CreateFieldDefFile Rst, PubRepoPath + "\VehSale.ttx", True
    CreateFieldDefFile RstSub1, PubRepoPath + "\VehSale1.ttx", True
    CreateFieldDefFile RstSub2, PubRepoPath + "\VehSale2.ttx", True
    
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    Set RstCompDet = New ADODB.Recordset
    RstCompDet.CursorLocation = adUseClient
    RstCompDet.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
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
            Case 8
                Foot8 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            Case 9
                Foot9 = left(RTrim(mID(Footer, j, I - j - 1)), 130)
            End Select
            Cnt = Cnt + 1
            j = I + 1
        End If
    Next
    
    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select SupInvOnVehSaleInv , TaxDetOnVehInv from Syctrl", GCn, adOpenDynamic, adLockOptimistic
        For I = 1 To rpt.ParameterFields.Count
            Select Case UCase(rpt.ParameterFields(I).ParameterFieldName)
                Case UCase("PrePrinted")
                    rpt.ParameterFields(I).AddCurrentValue (IIf(OptPlain.Value, False, True))
            End Select
        Next
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecSpeciality & "'"
            Case UCase("AmtPrefix")
                rpt.FormulaFields(I).TEXT = "'" & PubAmountPrefix & "'"
            Case UCase("TelcoInvYN")
                rpt.FormulaFields(I).TEXT = "" & Rst2!SupInvOnVehSaleInv & ""
            Case UCase("TaxDetYN")
                rpt.FormulaFields(I).TEXT = "" & Rst2!TaxDetOnVehInv & ""
'            Case UCase("InvPrefix")
'                rpt.FormulaFields(i).Text = "'" & Rst2!VehSaleInv_Prefix & "'"
            Case UCase("LST")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecGram & "'"
            Case UCase("SubRep1")
                rpt.FormulaFields(I).TEXT = "" & IIf(RstSub1.RecordCount = 0, 0, 1) & ""
            Case UCase("SubRep2")
                rpt.FormulaFields(I).TEXT = "" & IIf(RstSub2.RecordCount = 0, 0, 1) & ""
           Case UCase("AddDet")
                rpt.FormulaFields(I).TEXT = "" & IIf(txt(ADType) = "No Detail", 0, IIf(txt(ADType) = "Name/Qty", 1, 2)) & ""
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
            Case UCase("Foot8")
                rpt.FormulaFields(I).TEXT = "'" & Foot8 & "'"
            Case UCase("Foot9")
                rpt.FormulaFields(I).TEXT = "'" & Foot9 & "'"
            Case UCase("TOTCaption")
                rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
            Case UCase("HelpLineNo")
                rpt.FormulaFields(I).TEXT = "'" & HlpLineNo & "'"
                
        End Select
    Next
    For I = 1 To rpt.OpenSubreport("SubRep2").FormulaFields.Count
        Select Case UCase(rpt.OpenSubreport("SubRep2").FormulaFields(I).FormulaFieldName)
            Case UCase("AddDet")
            rpt.OpenSubreport("SubRep2").FormulaFields(I).TEXT = "" & IIf(txt(ADType) = "No Detail", 0, IIf(txt(ADType) = "Name/Qty", 1, 2)) & ""
        End Select
    Next
    rpt.Database.SetDataSource Rst
    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstSub1
    rpt.OpenSubreport("SubRep2").Database.SetDataSource RstSub2
    rpt.ReadRecords
    If PubVATYN = 1 Then
        Set tmprs = GCn.Execute("Select Description from SubGroupType Left join Subgroup on Subgroup.Party_Type=SubgroupType.Party_Type where Subgroup.SubCode='" & txt(Party).Tag & "'")
        If tmprs.RecordCount > 0 Then
            If tmprs!Description = "Dealer" Or UCase(XNull(tmprs!Description)) = "SUB DEALER" Then
                PrnTitle = "Tax Invoice"
            Else
                PrnTitle = "Retail Invoice"
            End If
        Else
            PrnTitle = "Retail Invoice"
        End If
    Else
        PrnTitle = "Sale Invoice"
    End If
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
                        rpt.FormulaFields(I).TEXT = "'" & PrnTitle & "'"
                End Select
            Next
            rpt.PrintOut False
            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
            End If
            Set Rst = Nothing
            Set RstCompDet = Nothing
            Set rpt = Nothing
        Case PScreen  'screen
            Call Report_View(rpt, PrnTitle, , True)
            Set Rst = Nothing
            Set RstCompDet = Nothing
    End Select
End If
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
Private Sub SpeedPrint22A(mQry$)
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
    Dim RstCert As ADODB.Recordset
    Dim PageWidth As Byte, PageLength As Integer
    Dim mHeader As Byte, mFooter As Byte
    Dim fob As New FileSystemObject

    Set RstCert = GCn.Execute(mQry)

    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
 
    PageLength = PubPageLength
    PageWidth = 80 '34
    mHeader = 0   'Ideal 17
    mFooter = 2
        
    Print #1, Chr(27) + Chr(67) + Chr(36) & PRN_TIT(Trim(RstCert!Manufacturer), "B", PageWidth) 'small paper size
'        Print #1, PRN_TIT(Trim(RstCert!Manufacturer), "C", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCert!MfgAdd1) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd1)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd2) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd2)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd3) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd3)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PRN_TIT("F O R M - 22-A", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[See Rule 47 (g),124,126a and 127]", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("Part 1", "B", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("(Issued By The Manufacturer)", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "Certified of compliance and pollution standards/safety of components"
        mHeader = mHeader + 1
        Print #1, mSP5 & "Road Worthiness (for vehicle whose Body is fabricated separately)"
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "Certified that Tata " & mChr17 & RstCert!FUEL & mChr18 & " Vehicle Model " & mEmph & RstCert!Model_Desc & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP5 & RstCert!WHEELBASE & "MM Wheel Base (Brand name of the vehicle) Truck/Bus/Car bearing "
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Chassis Number :" & RstCert!ChassisNo
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Engine Number  :" & RstCert!EngineNo
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "complies with the provisions of  the  Motor Vehicles Act, 1988 and the rule"
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "made there under."
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("For " & RstCert!Manufacturer, PageWidth, , AlignRight) & mEmph1
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, PSTR("Signature of the manufacturer", PageWidth, , AlignRight)
        Do Until mHeader >= PageLength - mFooter - 6
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        Print #1, mSP5 & Replace(Space(PageWidth), " ", "-")
'        Print #1, mSP5 & mChr17 & RstCert!Inv_UName & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(RstCert!Inv_UName)) / 2) & "* a dataman software *" & mChr18
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

Private Sub SpeedPrint22(mQry As String)
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
    Dim RstCert As ADODB.Recordset
    Dim PageWidth As Byte, PageLength As Integer
    Dim mHeader As Byte, mFooter As Byte
    Dim fob As New FileSystemObject

    Set RstCert = GCn.Execute(mQry)

    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1

    PageLength = PubPageLength
    PageWidth = 80 '34
    mHeader = 0   'Ideal 17
    mFooter = 2
    Print #1, Chr(27) + Chr(67) + Chr(36) & PRN_TIT(Trim(RstCert!Manufacturer), "B", PageWidth) 'small paper size
        
'        Print #1, PRN_TIT(Trim(RstCert!Manufacturer), "B", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCert!MfgAdd1) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd1)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd2) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd2)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If XNull(RstCert!MfgAdd3) <> "" Then
            Print #1, PRN_TIT(Trim(XNull(RstCert!MfgAdd3)), "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PRN_TIT("F O R M - 22", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[See Rule 47 (g), and 127]", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PRN_TIT("Initial certificate of Road Worthiness", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("(To be Issued By The Manufacturer)", "C", PageWidth)
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "Certified that Tata " & mChr17 & RstCert!FUEL & mChr18 & " Vehicle Model " & mEmph & RstCert!Model_Desc & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP5 & RstCert!WHEELBASE & "MM Wheel Base (Brand name of the vehicle) Truck/Bus/Car bearing "
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Chassis Number :" & RstCert!ChassisNo
        mHeader = mHeader + 1
        Print #1, mSP5 & "   Engine Number  :" & RstCert!EngineNo
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP5 & "complies with the provisions of  the  Motor Vehicles Act, 1988 and the rule"
        mHeader = mHeader + 1
        Print #1, mSP5 & "made there under."
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("For " & RstCert!Manufacturer, PageWidth, , AlignRight) & mEmph1
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, " "
        mHeader = mHeader + 1
        Print #1, PSTR("Signature of the manufacturer", PageWidth, , AlignRight)
        
        Do Until mHeader >= PageLength - mFooter - 6
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        Print #1, mSP5 & Replace(Space(PageWidth), " ", "-")
'        Print #1, mChr17 & RstCert!Inv_UName & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(RstCert!Inv_UName)) / 2) & "* a dataman software *" & mChr18
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

Private Sub SpeedPrintCerti(mQry$)
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
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstCert As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer1$, Footer2$, Footer3$, Footer4$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim NetAmt As Double
    
    Dim mPAdd1$, mPAdd2$, mPAdd3$, mPCity$, mPPin$
    Dim mTAdd1$, mTAdd2$, mTAdd3$, mTCITY$, mTPin$
    Dim mComp_Add$, mComp_Add2$, mComp_City$, mPhone$, mFax$

    Set RstCert = GCn.Execute(mQry)
    If RstCert.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Set RstCert = Nothing: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
 
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 4
    
    mComp_Add = Trim(PubComp_Add)
    mComp_Add2 = Trim(PubComp_Add2)
    mComp_City = Trim(PubComp_City)
'    If XNull(RstCompDet!V_SecPhone) <> "" Then
'        mPhone = "PHONE : " & RstCompDet!V_SecPhone
'    End If
'    If XNull(RstCompDet!V_SecFax) <> "" Then
'        mFax = "  Fax :" & RstCompDet!V_SecFax
'    End If
    
    'Header
    'Form22A,RTOName,PrnDate,CertiTempYN,Seet,Body,Narr,WtPrn,TempRto
    
    If txtPrint(CertiTempYN) = "Yes" Then
        mDocStr = "Temporary Sale Certificate"
        If RstCert!TCertiPrn_YN = 1 Then
            mDupStr = " (Duplicate)"
        End If
        mPAdd1 = XNull(RstCert!Add1)
        mPAdd2 = XNull(RstCert!Add2)
        mPAdd3 = XNull(RstCert!Add3)
        mPCity = XNull(RstCert!CityName)
        mPPin = XNull(RstCert!Pin)
        
        mTAdd1 = XNull(RstCert!tAdd1)
        mTAdd2 = XNull(RstCert!tAdd2)
        mTAdd3 = XNull(RstCert!TAdd3)
        mTCITY = XNull(RstCert!tCity)
        mTPin = XNull(RstCert!TPIN)
    Else
        mDocStr = "Sale Certificate"
        If RstCert!CertiPrn_YN = 1 Then
            mDupStr = " (Duplicate)"
        End If
        mPAdd1 = XNull(RstCert!Add1)
        mPAdd2 = XNull(RstCert!Add2)
        mPAdd3 = XNull(RstCert!Add3)
        mPCity = XNull(RstCert!CityName)
        mPPin = XNull(RstCert!Pin)
        
        mTAdd1 = XNull(RstCert!tAdd1)
        mTAdd2 = XNull(RstCert!tAdd2)
        mTAdd3 = XNull(RstCert!TAdd3)
        mTCITY = XNull(RstCert!tCity)
        mTPin = XNull(RstCert!TPIN)
    End If
 '   mDocStr = "Sale Certificate"
    
    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")
  If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
  Print #1, ""
  Print #1, ""
  Print #1, ""
Print #1, ""
  Else
      Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    End If
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
    Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", "  Fax : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
    mHeader = mHeader + 1
      If UCase(left(PubComp_Name, 6)) <> "J.M.A." Then Print #1, ""
    mHeader = mHeader + 1
    Print #1, PRN_TIT("Form-21", "B", PageWidth)
    mHeader = mHeader + 1
    Print #1, PRN_TIT("[See Rule 47(a) and (d)]", "C", PageWidth)
    mHeader = mHeader + 1
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "C", PageWidth) & mChr18 & mEmph
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("The Registration Authority,", 40) & mEmph & "     Invoice No.  : " & PSTR(Trim(mID(RstCert!Inv_DocId, 9, 5)) & "-" & Trim(mID(RstCert!Inv_DocId, 14, 8)), 14, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR(txtPrint(RTOName), 40) & mEmph & "     Invoice Date : " & PSTR(txtPrint(TempInvDate), 20, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
'    Print #1, mSP5 & "Ex. Factory Price Rs.: " & Format(RstCert!VRATE, "0.00")
'    mHeader = mHeader + 1
    
    Print #1, mSP5 & "(To be issued by the manufacturer, Dealer or Officer or defence (in Case of "
    mHeader = mHeader + 1
    Print #1, mSP5 & "Military auctioned vehicles)for presentation along with the application For"
    mHeader = mHeader + 1
    Print #1, mSP5 & "registration of a motor vehicle.)"
    mHeader = mHeader + 1
    
    Print #1, mSP5 & "Certified that - " & mEmph & "One " & RstCert!Model_Desc & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP5 & "Has been delivered by us on " & mEmph & RstCert!DelCh_DT & mEmph1 & " to :- "
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Name of Buyer", 28) & " : " & RstCert!Name
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Son/Wife/Daughter of ", 28) & " : " & XNull(RstCert!FPrefix) & " " & XNull(RstCert!fname)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Address(Permanent)", 28) & " : " & mPAdd1
    mHeader = mHeader + 1
        If mPAdd2 <> "" Then
            Print #1, mSP5 & Space(31) & mPAdd2
            mHeader = mHeader + 1
        End If
        If mPAdd3 <> "" Then
            Print #1, mSP5 & Space(31) & mPAdd3
            mHeader = mHeader + 1
        End If
    Print #1, mSP5 & Space(31) & mPCity & " " & mPPin
    mHeader = mHeader + 1
    If UCase(left(PubComp_Name, 6)) <> "J.M.A." Then Print #1, ""
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("Address(Temporary)", 28) & " : " & mTAdd1
    mHeader = mHeader + 1
    If mTAdd2 <> "" Then
        Print #1, mSP5 & Space(31) & mTAdd2
        mHeader = mHeader + 1
    End If
    If mTAdd3 <> "" Then
        Print #1, mSP5 & Space(31) & mTAdd3
        mHeader = mHeader + 1
    End If
    Print #1, mSP5 & Space(31) & mTCITY & " " & mTPin
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    If RstCert!Fund_Source = 0 Then
        Print #1, mSP5 & "The vehicle is held under agreement of Hypothecation with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 1 Then
        Print #1, mSP5 & "The vehicle is held under agreement of Hire purchase with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 3 Then
        Print #1, mSP5 & "The vehicle is held under agreement of Lease with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 4 Then
        Print #1, mSP5 & "The vehicle is held under Hire purchase finance agreement with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 5 Then
        Print #1, mSP5 & "The vehicle is held under Hire purchase finance Lease&agreement with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    ElseIf RstCert!Fund_Source = 6 Then
        Print #1, mSP5 & "The vehicle is held under Loan Cum Hypothecation Agreement with "
        mHeader = mHeader + 1
        Print #1, mSP5 & mEmph & RstCert!FinName & mEmph1
        mHeader = mHeader + 1
    Else
        Print #1, ""
        mHeader = mHeader + 1
    End If
    
    If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
        Print #1, mSP5 & XNull(RstCert!FAdd1) & " " & XNull(RstCert!FAdd2)
        mHeader = mHeader + 1
        Print #1, mSP5 & XNull(RstCert!FinCity) & " " & XNull(RstCert!FPin)
        mHeader = mHeader + 1
    End If
    
    Print #1, mSP5 & "The details of the vehicle are given below :  "
    mHeader = mHeader + 1
    
    If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
        Print #1, mSP5 & PSTR("1. Class of Vehicle", 32) & " : " & XNull(RstCert!Vehicle_Type)
        mHeader = mHeader + 1
    Else
        Print #1, mSP5 & PSTR("1. Class of Vehicle", 32) & " : Four Wheeler "
        mHeader = mHeader + 1
    End If
    
    Print #1, mSP5 & PSTR("2. Maker's Name", 32) & " : " & XNull(RstCert!Manufacturer)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("3. Chassis No.", 32) & " : " & XNull(RstCert!ChassisNo)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("4. Engine No.", 32) & " : " & XNull(RstCert!EngineNo)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("5. Horse Power/Cubic Capacity", 32) & " : " & XNull(RstCert!HorsePower)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("6. Fuel Used", 32) & " : " & XNull(RstCert!FUEL)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("7. No. of Cylinders ", 32) & " : " & XNull(RstCert!Cylinder)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("8. Month & Year of Mfg.", 32) & " : " & XNull(RstCert!Mfg_Month) & " " & XNull(RstCert!Mfg_Yr)
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("9. Seating Capacity(Incld. Driver)", 32) & " : " & RstCert!Seat
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("10.Unleaden Weight", 32) & " : " & RstCert!Unladen_Wt '& "Kg."
    mHeader = mHeader + 1
    Print #1, mSP5 & PSTR("11.Purpose", 32) & " : " & RstCert!PurposeName
    mHeader = mHeader + 1
    
    
    If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
        Print #1, mSP5 & PSTR("12.Maximum Axle Weight and number and Description of tyres", 60) & " : "
        mHeader = mHeader + 1
        Print #1, mSP5 & PSTR("   (in Case of Transport Vehicle)", 32)
        mHeader = mHeader + 1
        Print #1, mSP5 & "    " & PSTR("Front Axle", 28) & " : " & RstCert!Front_A_Wt ' & "Kg."
        mHeader = mHeader + 1
        Print #1, mSP5 & "    " & PSTR("Rear Axle", 28) & " : " & RstCert!Rear_A_Wt '& "Kg."
        mHeader = mHeader + 1
        Print #1, mSP5 & "    " & PSTR("Any Other Axle", 28) & " : "
        mHeader = mHeader + 1
        Print #1, mSP5 & "   " & Space(32) & PSTR("Front", 16, , AlignLeft) & PSTR("Rear", 16, , AlignLeft)
        mHeader = mHeader + 1
        Print #1, mSP5 & "    " & PSTR("No. of Tyres", 28) & " : " & PSTR(CStr(XNull(RstCert!Tyre_F)), 16, , AlignLeft) & PSTR(CStr(XNull(RstCert!Tyre_R)), 16, , AlignLeft)
        mHeader = mHeader + 1
        Print #1, mSP5 & "    " & PSTR("Size of Tyres", 28) & " : " & PSTR(RstCert!Tyre_FS, 16, , AlignLeft) & PSTR(RstCert!Tyre_RS, 16, , AlignLeft)
        mHeader = mHeader + 1
          If UCase(left(PubComp_Name, 6)) <> "J.M.A." Then
        Print #1, mSP5 & "    " & PSTR("Other Details", 28) & " : " & RstCert!TyreDetails
        mHeader = mHeader + 1
        End If
        Print #1, mSP5 & PSTR("13.Colour of Body", 32) & " : " & RstCert!Col_Desc
        mHeader = mHeader + 1
        Print #1, mSP5 & PSTR("14.Gross Vehicle Weight", 32) & " : " & RstCert!Gross_Wt '& "Kg."
        mHeader = mHeader + 1
        Print #1, mSP5 & PSTR("15.Type of Body", 32) & " : " & txtPrint(Body)
        mHeader = mHeader + 1
        Print #1, mSP5 & PSTR("16.Trade Ceri.No.", 32) & " : " & RstCert!TCNO
        mHeader = mHeader + 1
        
        Print #1, mSP5 & PSTR("17.WheelBase", 32) & " : " & RstCert!WHEELBASE & " MM"
        mHeader = mHeader + 1
        
        Print #1, mSP5 & PSTR("18.Gear Box", 32) & " : " & RstCert!GearBoxNo
        mHeader = mHeader + 1
        
    Else
        Print #1, mSP5 & PSTR("12.Colour of Body", 32) & " : " & RstCert!Col_Desc
        mHeader = mHeader + 1
        Print #1, mSP5 & PSTR("13.Trade No", 32) & " : " & RstCert!Trade_NO
        mHeader = mHeader + 1
    End If
    
    Print #1, ""
    mHeader = mHeader + 1
    'Print #1, mSP5 & mChr17 & "*Strike out whichever is inapplicable" & mChr18
    'mHeader = mHeader + 1
    Do Until mHeader >= PageLength - (mFooter + 6)
        Print #1, ""
        mHeader = mHeader + 1
    Loop
      If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
            Print #1, ""
        Print #1, ""
          Print #1, ""
    '    If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
'          Print #1, ""
        
      Else
    Print #1, mEmph & PSTR("For " & PubComp_Name, PageWidth, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, mSP5 & mChr17 & txtPrint(Narr) & mChr18
'    If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
    Print #1, mSP5 & Space(55) & "Authorised Signatory"
    End If
    Print #1, "* a dataman software *"
        'End Of Page 1 For SAle Certificate
 '   Else
 '       Print #1, ""
 '       Print #1, ""
 '   End If
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
        If txtPrint(CertiTempYN) = "Yes" Then
            GCn.Execute "update veh_order set TCertiPrn_YN = 1  where where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
        Else
            GCn.Execute "update veh_order set CertiPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
        End If
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub



Private Sub SpeedPrintInv(mQry$)
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
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim mSaleRate As Double, mNetAmt As Single, mInv_No$

     Set Rstsale = GCn.Execute(mQry)
    
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 25
    mFooter = mFooter + FooterCnt
    
    ' Header
    If CancelBillY_N = True Then
        mDocStr = "Sale Invoice (Credit Note)"
    Else
        Dim tmprs As ADODB.Recordset
        Set tmprs = GCn.Execute("Select Description from SubGroupType Left join Subgroup on Subgroup.Party_Type=SubgroupType.Party_Type where Subgroup.SubCode='" & Rstsale!PartyCode & "'")
        If tmprs.RecordCount > 0 Then
            If tmprs!Description = "Dealer" Then
                mDocStr = "Tax Invoice"
            Else
                mDocStr = "Retail Invoice"
            End If
        Else
            mDocStr = "Retail Invoice"
        End If
        
    End If
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        mDupStr = ""
    Else
        mDupStr = IIf(Rstsale!BillPrn_YN = 1, " (Duplicate)", "")
    End If
 '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease, 4-Agreement, 5-Lease & Agreement

    If Rstsale!Fund_Source = 0 Then   'Hypothecation
        Left1 = "To,"
        If txt(NamePrefix) <> "" Then
            Left2 = txt(NamePrefix) & " " & XNull(Rstsale!Name)
        Else
            Left2 = XNull(Rstsale!Name)
        End If
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Hypothecation to  "
        Right2 = XNull(Rstsale!FinName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = ""
        
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        If txt(NamePrefix) <> "" Then
            Left2 = txt(NamePrefix) & " " & XNull(Rstsale!Name)
        Else
            Left2 = XNull(Rstsale!Name)
        End If
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        If UCase(left(PubComp_Name, 4)) = "ENAR" Then
            Right1 = "A/C  "
        Else
            Right1 = "Leaser  "
        End If
        Right2 = XNull(Rstsale!FinName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = ""
        
    ElseIf Rstsale!Fund_Source = 6 Then
        Left1 = "To,"
        If txt(NamePrefix) <> "" Then
            Left2 = txt(NamePrefix) & " " & XNull(Rstsale!Name)
        Else
            Left2 = XNull(Rstsale!Name)
        End If
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & ", " & XNull(Rstsale!PCityName)
        
        Right1 = "Under Loan Cum Hypt. Agreement with  "
        Right2 = XNull(Rstsale!FinName)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = ""
    
        
    ElseIf Rstsale!Fund_Source = 1 Or _
           Rstsale!Fund_Source = 4 Or _
           Rstsale!Fund_Source = 5 Then
        
        Left1 = "Sold to under HPA with, "      '1-Hire Purchase
        If Rstsale!Fund_Source = 4 Then         '4-Agreement
            Left1 = "Hire Purchase Finance Agreement with, "
        ElseIf Rstsale!Fund_Source = 5 Then     '5-Lease & Agreement
            Left1 = "Hire Purchase Finance Lease&Agreement with, "
        
        End If
        Left2 = " U/F " & XNull(Rstsale!FinName)
        Left3 = XNull(Rstsale!FinAdd1)
        Left4 = XNull(Rstsale!FinAdd2)
        Left5 = XNull(Rstsale!FinCity)
        Left6 = ""
        
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!PAdd1)
        Right5 = XNull(Rstsale!PAdd2)
        Right6 = XNull(Rstsale!PAdd3) & ", " & XNull(Rstsale!PCityName)
        
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
    End If
    

    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
    mInv_No = Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_Prefix)) & " - " & Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_No))
    
    '************
   If UCase(left(PubComp_Name, 4)) = "COMM" Then
        mSaleRate = Format(Rstsale!vrate + Rstsale!Margine - Rstsale!Subvention + Rstsale!InciChrg _
                     + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
                     + Rstsale!MVT + Rstsale!Transport + Rstsale!Round_off, "0.00")
    Else
        mSaleRate = Format(Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate - Rstsale!Subvention + Rstsale!InciChrg _
                    + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
                    + Rstsale!MVT + Rstsale!Transport + Rstsale!Round_off, "0.00")

    End If
    mNetAmt = mSaleRate + Rstsale!Tax_Amt + VNull(Rstsale!SatAmt) + Rstsale!Surcharge_Amt _
        + Rstsale!Tot_Amt + Rstsale!OtherChrg + Rstsale!Fit_Amt _
        + Rstsale!Fit_Tax - Rstsale!DieselAmt - IIf(StrCmp(left(PubComp_Name, 4), "comm"), VNull(Rstsale!Rebate), 0)
    '***********
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
    'Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
    If PubComp_Contact <> "" Then
        Print #1, PRN_TIT(PubComp_Contact, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40)
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mChr18 & mEmph & PSTR(Left1, 40) & Space(10) & PSTR(Right1, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(Left2 = "", "--", Left2), 40) & Space(10) & PSTR(Right2, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(Left3 = "", "--", Left3), 40) & Space(10) & PSTR(Right3, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(Left4 = "", "--", Left4), 40) & Space(10) & PSTR(Right4, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(Left5 = "", "--", Left5), 40) & Space(10) & PSTR(Right5, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(Left6 = "", "--", Left6), 40) & Space(10) & PSTR(Right6, 40)
    mHeader = mHeader + 1
    
    If Rstsale!Fund_Source = 0 Or Rstsale!Fund_Source = 2 Or Rstsale!Fund_Source = 3 Or Rstsale!Fund_Source = 6 Then
        Print #1, PSTR("Phone : " & XNull(Rstsale!Phone) & "," & XNull(Rstsale!Mobile), 40)
        mHeader = mHeader + 1
    End If
    
    Print #1, ""
    Print #1, IIf(XNull(Rstsale!PartyLstNo) = "", Space(50), PSTR("Tin/Lst No : " & XNull(Rstsale!PartyLstNo), 50)) & mEmph & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    'Print #1, Space(50) & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
    'mHeader = mHeader + 1
    Print #1, IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Mfg. Invoice No.: " & XNull(Rstsale!PBILL_NO) & " " & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 50), Space(50)) & mEmph & "Invoice Date : " & STR(Rstsale!Inv_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("Booking No. : " & STR(Rstsale!Ord_No), 40) & Space(10) & PSTR("Booking Date :" & IIf(IsNull(Rstsale!Ord_Date), "", (STR(Rstsale!Ord_Date))), 40)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, " P A R T I C U L A R S "
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    'If Rstsale!Model = Rstsale!Sales_Desc Then     '' Modification made due to VC No. System
    '    Print #1, mEmph & PSTR("Model : " & Trim(Rstsale!ModelGrp_Name) & " " & Rstsale!Model_Desc, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight) & mEmph1
    'Else
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        Print #1, mEmph & PSTR("Model : " & Trim(Rstsale!Model) & ", " & Trim(Rstsale!Sales_Desc), 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight) & mEmph1
    Else
        Print #1, mEmph & PSTR("Model : " & Rstsale!Model, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight) & mEmph1
    End If
    
    
    'End If
    mHeader = mHeader + 1
    Print #1, mEmph & PSTR("        " & Rstsale!Model_Desc, 45) & mEmph1
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PSTR("Colour      : " & Rstsale!Col_Desc, 40) & IIf(VNull(Rstsale!surcharge_per) = 0, "", PSTR("Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %", 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2))
    mHeader = mHeader + 1
    Print #1, PSTR("Chassis No. : " & Rstsale!ChassisNo, 45) & IIf(Rstsale!TOT_Per = 0, "", PSTR("TOT @ " & Format(Rstsale!TOT_Per, "0.00") & " %", 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2))
    mHeader = mHeader + 1
    Print #1, PSTR("Engine No.  : " & Rstsale!EngineNo, 40)
    mHeader = mHeader + 1
    If StrCmp(left(PubComp_Name, 4), "comm") Then
        Print #1, Space(45) & IIf(Rstsale!Rebate = 0, "", PSTR("Rebate ", 22, , AlignRight) & ": " & PSTR(Rstsale!Rebate, 11, 2))
        mHeader = mHeader + 1
    End If
    
    Print #1, Space(45) & IIf(Rstsale!Tax_Amt = 0, "", PSTR(IIf(left(UCase(PubComp_Name), 4) = "ENAR", XNull(Rstsale!Printing_Desc), XNull(Rstsale!Printing_Desc)), 22, , AlignRight) & ": " & PSTR(Rstsale!Tax_Amt, 11, 2))
    mHeader = mHeader + 1
    If VNull(Rstsale!SatAmt) > 0 Then
        Print #1, Space(45) & PSTR("S A T  @ " & VNull(Rstsale!SatPer) & "%  ", 22, , AlignRight) & ": " & PSTR(Rstsale!SatAmt, 11, 2)
        mHeader = mHeader + 1
    End If
                
    'Print #1, "" & mEmph1
    'mHeader = mHeader + 1
    
    'Print #1, ""
    'mHeader = mHeader + 1
    'Print #1, PSTR("Sr", 3) & PSTR("Item Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, 2, AlignRight) & " " & PSTR("<----Tax---- >", 13) & " " & PSTR("<-Sur.On Tax- >", 14) & " " & PSTR("Amount", 9, , AlignRight)
    'mHeader = mHeader + 1
    'Print #1, "No." & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
    'mHeader = mHeader + 1
    'Print #1, Replace(Space(PageWidth), " ", "-")
    'mHeader = mHeader + 1
        
    'Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
    '    "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    '    "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")
    'cnt = 1
    'If Rst.RecordCount > 0 Then
    '    Do Until Rst.EOF
    '        Print #1, mChr17 & STR(cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) + Rst!Tax_Amt + Rst!TaxSur_Amt), 10, 2)
    '        mHeader = mHeader + 1
    '        cnt = cnt + 1
    '        Rst.MoveNext
    '    Loop
    'End If
    
   ' Print #1, Replace(Space(PageWidth), " ", "-")
   ' mHeader = mHeader + 1
    
    'Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
    '    "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    '    "where veh_stock.Chassisno = '" & txt(ChassisNo) & "' " & _
    '    "Group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
    'If Rst.RecordCount > 0 Then
     '   Print #1, mDoub & PSTR("Addition/Deletion/Shortage Detail", 52) & PSTR("Qty", 13, , AlignRight) & PSTR("Amount", 15, , AlignRight) & mDoub1
     '   mHeader = mHeader + 1
     '   Do Until Rst.EOF
     '       Print #1, PSTR(IIf(Rst!Trn_Type = "A", "Addition", IIf(Rst!Trn_Type = "D", "Deletion", "Shortage")), 52) & PSTR(Rst!TotQty, 13, 2) & PSTR(Rst!Amt, 15, 2)
     '       mHeader = mHeader + 1
     '       Rst.MoveNext
     '   Loop
     '   Print #1, Replace(Space(PageWidth), " ", "-")
     '   mHeader = mHeader + 1
    'End If
        
    Do Until mHeader >= PageLength - mFooter
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    If VNull(Rstsale!Round_off) <> 0 Then
        Print #1, PSTR(IIf(VNull(Rstsale!Round_off) = 0, "", "Round Off"), 65, , AlignRight) & " : " & PSTR(Rstsale!Round_off, 12, 2)
    End If
   ' Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
    Print #1, mEmph & PSTR("Total Bill Amount", 65, , AlignRight) & " : " & PSTR(Amount_Fill((mNetAmt), PubAmountPrefix), 12, 2, AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, ntow(mNetAmt, "Rupees", "Paise") & mEmph1
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, "Complete With Tools and equipment as supplied by the manufacturer including "
    Print #1, "excise duty,Sales tax / VAT & delivery & handing charges."
    Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, XNull(Rstsale!MISC_INFO)
    
    Print #1, "Delevered By :" & XNull(Rstsale!Emp_Name)
    Print #1, ""
    Print #1, PSTR("Customer Signature               Accountant                Authorised Signatory", PageWidth, , AlignLeft)
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")

    Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
        
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    
    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
           
    Print #1, mChr17 & Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
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
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If

    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Sub SpeedPrintInvSanya(mQry$)
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
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim mSaleRate As Single, mNetAmt As Single, mInv_No$

     Set Rstsale = GCn.Execute(mQry)
    
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 25
    mFooter = mFooter + FooterCnt
    
    ' Header
    If CancelBillY_N = True Then
        mDocStr = "Sale Invoice (Credit Note)"
    Else
        mDocStr = "Sale Invoice"
    End If
    mDupStr = IIf(Rstsale!BillPrn_YN = 0, "", " (Duplicate)")
 '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease, 4-Agreement, 5-Lease & Agreement

    If Rstsale!Fund_Source = 0 Then   'Hypothecation
        Left1 = "To,"
        If txt(NamePrefix) <> "" Then
            Left2 = txt(NamePrefix) & " " & XNull(Rstsale!Name)
        Else
            Left2 = XNull(Rstsale!Name)
        End If
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Hypothecation to  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = ""
        
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        If txt(NamePrefix) <> "" Then
            Left2 = txt(NamePrefix) & " " & XNull(Rstsale!Name)
        Else
            Left2 = XNull(Rstsale!Name)
        End If
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Leaser  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = ""
        
    ElseIf Rstsale!Fund_Source = 6 Then
        Left1 = "To,"
        If txt(NamePrefix) <> "" Then
            Left2 = txt(NamePrefix) & " " & XNull(Rstsale!Name)
        Else
            Left2 = XNull(Rstsale!Name)
        End If
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Loan Cum Hypt. Agreement with  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = ""
    
        
    ElseIf Rstsale!Fund_Source = 1 Or _
           Rstsale!Fund_Source = 4 Or _
           Rstsale!Fund_Source = 5 Then
        
        Left1 = "Sold to under HPA with, "      '1-Hire Purchase
        If Rstsale!Fund_Source = 4 Then         '4-Agreement
            Left1 = "Hire Purchase Finance Agreement with, "
        ElseIf Rstsale!Fund_Source = 5 Then     '5-Lease & Agreement
            Left1 = "Hire Purchase Finance Lease&Agreement with, "
        
        End If
        Left2 = " U/F " & XNull(Rstsale!finbankname)
        Left3 = XNull(Rstsale!FinAdd1)
        Left4 = XNull(Rstsale!FinAdd2)
        Left5 = XNull(Rstsale!FinCity)
        Left6 = ""
        
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!PAdd1)
        Right5 = XNull(Rstsale!PAdd2)
        Right6 = XNull(Rstsale!PAdd3) & XNull(Rstsale!PCityName)
        
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
    End If
    

    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
    mInv_No = Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_Prefix)) & " - " & Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_No))
    
    '************
    mSaleRate = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate - Rstsale!Subvention + Rstsale!InciChrg _
         + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
         + Rstsale!MVT + Rstsale!Transport
         
    mNetAmt = mSaleRate + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!Tot_Amt + Rstsale!OtherChrg + Rstsale!Fit_Amt _
        + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off
    '***********
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
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40)
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, mChr18 & mEmph & PSTR(Left1, 40) & Space(10) & PSTR(Right1, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(Left2, 40) & Space(10) & PSTR(Right2, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left3, 40) & Space(10) & PSTR(Right3, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left4, 40) & Space(10) & PSTR(Right4, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left5, 40) & Space(10) & PSTR(Right5, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left6, 40) & Space(10) & PSTR(Right6, 40)
    mHeader = mHeader + 1
    Print #1, ""
    Print #1, Space(50) & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, Space(50) & mEmph & "Invoice Date : " & STR(Rstsale!Inv_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("Booking No. : " & STR(Rstsale!Ord_No), 40) & Space(10) & PSTR("Booking Date :" & IIf(IsNull(Rstsale!Ord_Date), "", (STR(Rstsale!Ord_Date))), 40)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, " P A R T I C U L A R S "
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, mEmph & PSTR("Model : " & Rstsale!Model_Desc, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight) & mEmph1
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PSTR("Colour      : " & Rstsale!Col_Desc, 40) & IIf(VNull(Rstsale!surcharge_per) = 0, "", PSTR("Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %", 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2))
    mHeader = mHeader + 1
    Print #1, PSTR("Chassis No. : " & Rstsale!ChassisNo, 45) & IIf(Rstsale!TOT_Per = 0, "", PSTR("TOT @ " & Format(Rstsale!TOT_Per, "0.00") & " %", 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2))
    mHeader = mHeader + 1
    Print #1, PSTR("Engine No.  : " & Rstsale!EngineNo, 40)
    mHeader = mHeader + 1
    Print #1, Space(45) & IIf(Rstsale!Tax_Amt = 0, "", PSTR("V A T ", 22, , AlignRight) & ": " & PSTR(Rstsale!Tax_Amt, 11, 2))
    mHeader = mHeader + 1
                
    'Print #1, "" & mEmph1
    'mHeader = mHeader + 1
    
    'Print #1, ""
    'mHeader = mHeader + 1
    'Print #1, PSTR("Sr", 3) & PSTR("Item Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, 2, AlignRight) & " " & PSTR("<----Tax---- >", 13) & " " & PSTR("<-Sur.On Tax- >", 14) & " " & PSTR("Amount", 9, , AlignRight)
    'mHeader = mHeader + 1
    'Print #1, "No." & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
    'mHeader = mHeader + 1
    'Print #1, Replace(Space(PageWidth), " ", "-")
    'mHeader = mHeader + 1
        
    'Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
    '    "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    '    "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")
    'cnt = 1
    'If Rst.RecordCount > 0 Then
    '    Do Until Rst.EOF
    '        Print #1, mChr17 & STR(cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) + Rst!Tax_Amt + Rst!TaxSur_Amt), 10, 2)
    '        mHeader = mHeader + 1
    '        cnt = cnt + 1
    '        Rst.MoveNext
    '    Loop
    'End If
    
   ' Print #1, Replace(Space(PageWidth), " ", "-")
   ' mHeader = mHeader + 1
    
    'Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
    '    "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
    '    "where veh_stock.Chassisno = '" & txt(ChassisNo) & "' " & _
    '    "Group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
    'If Rst.RecordCount > 0 Then
     '   Print #1, mDoub & PSTR("Addition/Deletion/Shortage Detail", 52) & PSTR("Qty", 13, , AlignRight) & PSTR("Amount", 15, , AlignRight) & mDoub1
     '   mHeader = mHeader + 1
     '   Do Until Rst.EOF
     '       Print #1, PSTR(IIf(Rst!Trn_Type = "A", "Addition", IIf(Rst!Trn_Type = "D", "Deletion", "Shortage")), 52) & PSTR(Rst!TotQty, 13, 2) & PSTR(Rst!Amt, 15, 2)
     '       mHeader = mHeader + 1
     '       Rst.MoveNext
     '   Loop
     '   Print #1, Replace(Space(PageWidth), " ", "-")
     '   mHeader = mHeader + 1
    'End If
        
    Do Until mHeader >= PageLength - mFooter
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    If VNull(Rstsale!Round_off) <> 0 Then
        Print #1, PSTR(IIf(VNull(Rstsale!Round_off) = 0, "", "Round Off"), 65, , AlignRight) & " : " & PSTR(Rstsale!Round_off, 12, 2)
    End If
   ' Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
    Print #1, mEmph & PSTR("Total Bill Amount", 65, , AlignRight) & " : " & PSTR(Amount_Fill((mNetAmt), PubAmountPrefix), 12, 2, AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, ntow(mNetAmt, "Rupees", "Paise") & mEmph1
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, "Complete With Tools and equipment as supplied by the manufacturer including "
    Print #1, "excise duty,Sales tax & delivery & handing charges."
    Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, "Delevered By :" & XNull(Rstsale!Emp_Name)
    Print #1, ""
    Print #1, PSTR("Customer Signature               Accountant                Authorised Signatory", PageWidth, , AlignLeft)
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")

    Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
        
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    
    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
           
    Print #1, mChr17 & Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
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
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If

    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Sub SpeedPrintInvSOCIETY(mQry$)
On Error GoTo ELoop
Dim TotSDT As String
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
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim mSaleRate As Double, mNetAmt As Single, mInv_No$

     Set Rstsale = GCn.Execute(mQry)
    
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 19
    mFooter = mFooter + FooterCnt
    
    ' Header
    If CancelBillY_N = True Then
        mDocStr = "Sale Invoice (Credit Note)"
    Else
        mDocStr = "Sale Invoice"
    End If
    mDupStr = IIf(Rstsale!BillPrn_YN = 0, "", " (Duplicate)")
 '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease, 4-Agreement, 5-Lease & Agreement

    If Rstsale!Fund_Source = 0 Then   'Hypothecation
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Hypothecation to  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = ""
        Right4 = ""
        Right5 = ""
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Leaser  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = ""
        Right4 = ""
        Right5 = ""
        Right6 = "Lease Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 6 Then
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Loan Cum Hypt. Agreement with  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = ""
        Right4 = ""
        Right5 = ""
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
    
        
    ElseIf Rstsale!Fund_Source = 1 Or _
           Rstsale!Fund_Source = 4 Or _
           Rstsale!Fund_Source = 5 Then
        
        Left1 = "Sold to under HPA with, "      '1-Hire Purchase
        If Rstsale!Fund_Source = 4 Then         '4-Agreement
            Left1 = "Hire Purchase Finance Agreement with, "
        ElseIf Rstsale!Fund_Source = 5 Then     '5-Lease & Agreement
            Left1 = "Hire Purchase Finance Lease&Agreement with, "
        
        End If
        Left2 = " U/F " & XNull(Rstsale!finbankname)
        Left3 = ""
        Left4 = ""
        Left5 = ""
        Left6 = ""
        
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!PAdd1)
        Right5 = XNull(Rstsale!PAdd2)
        Right6 = XNull(Rstsale!PAdd3) & XNull(Rstsale!PCityName)
        
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
    End If
    

    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
    mInv_No = Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_Prefix)) & " - " & Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_No))
    
    '************
    mSaleRate = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate - Rstsale!Subvention + Rstsale!InciChrg _
         + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
         + Rstsale!MVT + Rstsale!Transport
         
    mNetAmt = mSaleRate + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!Tot_Amt + Rstsale!OtherChrg + Rstsale!Fit_Amt _
        + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off
    '***********
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

    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, mChr18 & mEmph & PSTR(Left1, 40) & PSTR(Right1, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(Left2, 40) & PSTR(Right2, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left3, 40) & PSTR(Right3, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left4, 40) & PSTR(Right4, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left5, 40) & PSTR(Right5, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left6, 40) & PSTR(Right6, 40)
    mHeader = mHeader + 1
        
    Print #1, IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Mfg. Invoice No.: " & XNull(Rstsale!PBILL_NO) & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 40), Space(40)) & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("Sale Type. : " & XNull(Rstsale!FormName), 40) & mEmph & "Invoice Date : " & STR(Rstsale!Inv_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, "Booking No. & Date  : " & STR(Rstsale!Ord_No) & "    " & IIf(IsNull(Rstsale!Ord_Date), "", (STR(Rstsale!Ord_Date)))
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, PSTR("Model : " & Rstsale!Modl, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight)
    mHeader = mHeader + 1
    Print #1, PSTR(Rstsale!Model_Desc, 45) & PSTR(IIf(Rstsale!Tax_Per = 0, "", "Tax @ " & Format(Rstsale!Tax_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tax_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Colour      : " & Rstsale!Col_Desc, 40) & PSTR(IIf(Rstsale!surcharge_per = 0, "", "Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %"), 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2)
    mHeader = mHeader + 1
    If PubSDTYN = 1 Then
        Print #1, PSTR("Chassis No. : " & Rstsale!ChassisNo, 45) & PSTR(IIf(Rstsale!TOT_Per = 0, "", "SDT @ " & Format(Rstsale!TOT_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2)
        mHeader = mHeader + 1
    Else
        Print #1, PSTR("Chassis No. : " & Rstsale!ChassisNo, 45) & PSTR(IIf(Rstsale!TOT_Per = 0, "", "TOT @ " & Format(Rstsale!TOT_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2)
        mHeader = mHeader + 1
    End If
        
    Print #1, PSTR("Engine No.  : " & Rstsale!EngineNo, 40) & PSTR("Other Charges", 27, , AlignRight) & ": " & PSTR(Rstsale!OtherChrg, 11, 2)
    mHeader = mHeader + 1
    
    Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")
  
    If Rst.RecordCount > 0 Then
        Print #1, "Other Fitments Details : " & mEmph1
        mHeader = mHeader + 1
    
        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        mHeader = mHeader + 1
        Print #1, PSTR("Sr", 3) & PSTR("Item Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, 2, AlignRight) & " " & PSTR("<----Tax---- >", 13) & " " & PSTR("<-Sur.On Tax- >", 14) & " " & PSTR("Amount", 9, , AlignRight)
        mHeader = mHeader + 1
        Print #1, "No." & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        
        Cnt = 1
        Do Until Rst.EOF
            Print #1, mChr17 & STR(Cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) + Rst!Tax_Amt + Rst!TaxSur_Amt), 10, 2)
            mHeader = mHeader + 1
            Cnt = Cnt + 1
            Rst.MoveNext
        Loop
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
    End If
    
'    Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
'        "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
'        "where veh_stock.Chassisno = '" & Txt(ChassisNo) & "' " & _
'        "Group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
'    If Rst.RecordCount > 0 Then
'        Print #1, mDoub & PSTR("Addition/Deletion/Shortage Detail", 52) & PSTR("Qty", 13, , AlignRight) & PSTR("Amount", 15, , AlignRight) & mDoub1
'        mHeader = mHeader + 1
'        Do Until Rst.EOF
'            Print #1, PSTR(IIf(Rst!Trn_Type = "A", "Addition", IIf(Rst!Trn_Type = "D", "Deletion", "Shortage")), 52) & PSTR(Rst!TotQty, 13, 2) & PSTR(Rst!Amt, 15, 2)
'            mHeader = mHeader + 1
'            Rst.MoveNext
'        Loop
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
'    End If
    Do Until mHeader >= PageLength - (mFooter + 5)
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    If Rstsale!Round_off <> 0 Then
        Print #1, PSTR(IIf(Rstsale!Round_off = 0, "", "Round Off"), 65, , AlignRight) & " : " & PSTR(Rstsale!Round_off, 12, 2)
    Else
        Print #1, ""
    End If
    If Rstsale!DieselAmt <> 0 Then
        Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
    Else
        Print #1, ""
    End If
    Print #1, mEmph & PSTR("Bill Amount", 65, , AlignRight) & " : " & PSTR(Amount_Fill((mNetAmt), PubAmountPrefix), 12, 2, AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, ntow(mNetAmt, "Rupees", "Paise") & mEmph1
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, "Complete With Tools and equipment as supplied by the manufacturer including "
    Print #1, "excise duty,Sales tax & delivery & handing charges."
    Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, ""
    Print #1, "Accountant" & PSTR("Authorised Signatory", PageWidth - 10, , AlignRight)
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")

    Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
        
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    
    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
           
    Print #1, mChr17 & Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
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
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If

    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintDeclar()
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
    Dim I As Integer, j As Integer, mQry As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim Cnt As Byte, NetAmt As Double, PrnStr As String, PrnStr1 As String, mRegCert As String
    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set Rstsale = GCn.Execute("SELECT veh_order.*,City.CityName,  " & _
        " Veh_Stock.ChassisNo, Veh_Stock.EngineNo,Model.Model_Desc,Model.Model_Desc1, " & _
        " SubGroup.Name, SubGroup.Add1,SubGroup.Add2,SubGroup.Add3,SubGroup.Tadd1,SubGroup.Tadd2,SubGroup.Tadd3 FROM  " & _
        "(((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode where veh_order.Inv_DocId = '" & Master!SearchCode & "'")

    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1


    FooterCnt = 1
    Footer = ""

'    For i = 1 To Len(Footer)
'        If Mid(Footer, i, 1) = vbLf Then
'            FooterCnt = FooterCnt + 1
'        End If
'    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 6
    mFooter = mFooter + FooterCnt

    ' Header

    mDocStr = "DECLARATION"

        Print #1, PRN_TIT(mDocStr, "A", PageWidth)
        mHeader = mHeader + 1
        Print #1, PRN_TIT("[Under rule 2148 (1)]", "C", PageWidth) & mChr18 & mEmph
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "Declartion No. : " & Space(10) & "Date : " & mEmph1
        mHeader = mHeader + 1
        Print #1, "I/We declare that the following consignment of notified comodity is "
        mHeader = mHeader + 1
        Print #1, "Despatched from a place within West Bengal : "
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "1. " & "Name & address of the consignor : " & mEmph & PubComp_Name & mEmph1
        Print #1, Space(25) & " : " & mEmph & PubComp_Add & mEmph1
        mHeader = mHeader + 1
        Print #1, Space(25) & " : " & mEmph & PubComp_Add2 & mEmph1
        mHeader = mHeader + 1
        Print #1, Space(25) & " : " & mEmph & PubComp_City & mEmph1
        mHeader = mHeader + 1
        Print #1, Space(25) & " : " & IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax)
        mHeader = mHeader + 1
        If txtPrint(CertiTempYN) = "Yes" Then
            Print #1, PSTR("2. a) " & "Name and address ", 25) & " : " & mEmph & XNull(Rstsale!Name) & mEmph1
            mHeader = mHeader + 1
            Print #1, PSTR("of the Consignee", 25) & " : " & mEmph & XNull(Rstsale!tAdd1) & mEmph1
            mHeader = mHeader + 1
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!tAdd2) & XNull(Rstsale!TAdd3) & mEmph1
            mHeader = mHeader + 1
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, PSTR("   b) " & "Temporary Address", 25) & " : " & XNull(Rstsale!Add1)
            mHeader = mHeader + 1
            If XNull(Rstsale!Add2) <> "" Then
                Print #1, Space(25) & " : " & Rstsale!Add2
                mHeader = mHeader + 1
            End If
            If XNull(Rstsale!Add3) <> "" Then
                Print #1, Space(23) & " : " & Rstsale!Add3
                mHeader = mHeader + 1
            End If
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!CityName) & mEmph1
            mHeader = mHeader + 1
        Else
            Print #1, PSTR("2. a) " & "Name and address ", 25) & " : " & mEmph & XNull(Rstsale!Name) & mEmph1
            mHeader = mHeader + 1
            Print #1, PSTR("of the Consignee", 25) & " : " & mEmph & XNull(Rstsale!Add1) & mEmph1
            mHeader = mHeader + 1
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!Add2) & XNull(Rstsale!Add3) & mEmph1
            mHeader = mHeader + 1
            Print #1, Space(25) & " : " & mEmph & XNull(Rstsale!CityName) & mEmph1
            mHeader = mHeader + 1
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, PSTR("   b) " & "Temporary Address", 25) & " : " & XNull(Rstsale!tAdd1)
            mHeader = mHeader + 1
            If XNull(Rstsale!tAdd2) <> "" Then
                Print #1, Space(25) & " : " & Rstsale!tAdd2
                mHeader = mHeader + 1
            End If
            If XNull(Rstsale!TAdd3) <> "" Then
                Print #1, Space(23) & " : " & Rstsale!TAdd3
                mHeader = mHeader + 1
            End If
            
        End If
        
        Print #1, "   c) " & "Registration certificate No. of the consignee [if registered under "
        mHeader = mHeader + 1
        Print #1, "the West Bengal Sales Tax Act. 1994(West Ben. Act. XLIX of 1994)/ "
        mHeader = mHeader + 1
        mRegCert = XNull(GCn.Execute("select RegCertNo from Syctrl").Fields(0).Value)
        Print #1, "the central Sals Tax Act. 1956(74 of 1956) ] : " & mEmph & "Nil" & mEmph1
        mHeader = mHeader + 1

        Print #1, "3. " & "Place Of Dispatch : " & mEmph & PubComp_City & mEmph1
        mHeader = mHeader + 1
        Print #1, "4. " & "Destination : "
        mHeader = mHeader + 1
        Print #1, "5. " & "Description of consignment  : " & mEmph & Rstsale!Model_Desc & mEmph1
        mHeader = mHeader + 1
        Print #1, "6. " & PSTR("Quantity  : ", 15) & mEmph & "1 No. (One)" & mEmph1
        mHeader = mHeader + 1
        Print #1, "7. " & PSTR(" Weight : ", 15)
        mHeader = mHeader + 1
        NetAmt = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate - Rstsale!Subvention + Rstsale!InciChrg _
        + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
        + Rstsale!MVT + Rstsale!Transport + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!OtherChrg + Rstsale!Fit_Amt + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off

        Print #1, "8. " & PSTR("Value  : ", 15) & mEmph & NetAmt & mEmph1
        mHeader = mHeader + 1
        Print #1, "9. " & "Consignor Bill/Cash Memo/Other"
        mHeader = mHeader + 1
        Print #1, "   " & "Document(Specify) No. and date :" & mEmph & Rstsale!Inv_No & " Dt. " & Rstsale!Inv_Date & mEmph1
        mHeader = mHeader + 1
        Print #1, "   " & "Consignment or deleivery note No. and Date : " & mEmph & "delivery Receipt Dt.____________ "
        mHeader = mHeader + 1
        Print #1, "Chassis Deliverd at " & PubComp_City & " on ___________ and now "
        mHeader = mHeader + 1
        Print #1, "transported by the customer by his/her/their own mode" & mEmph1
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "I/We declare that I/we hold the registration certificate No " & mRegCert
        mHeader = mHeader + 1
        Print #1, "Under the West Bengal Sales Tax Act. 1994.(West Ben. Act XLIX of 1994)."
        mHeader = mHeader + 1
        Print #1, "We have not manufactured the comodity in West Bengal/not transported the "
        mHeader = mHeader + 1
        Print #1, "commodities from outside of West Bengal"
        mHeader = mHeader + 1
        Print #1, "The Above Statement are true to the best of my/our knowledge and belief"
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mEmph & "Chassis No :" & Rstsale!ChassisNo
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "Engine No :" & Rstsale!EngineNo & mEmph1
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, "Temp. Regn No. ______________________ :"
        mHeader = mHeader + 1
        Print #1, Space(40) & "Signature ______________________ :"
        mHeader = mHeader + 1
        Print #1, Space(40) & "Status of the declarent  ______________________ :" & mEmph1
        mHeader = mHeader + 1
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
Private Function ProcAcPost(Optional CheckCtrls As Boolean) As Boolean
Dim mPostOctraiSaperateYN As Boolean
Dim mSubventionTata As Double
On Error GoTo lblExit
        Dim MsgStr$, rsCtrlAc As ADODB.Recordset, RsTemp As ADODB.Recordset, mPostFinAmt As Byte
        Dim mGTotAmt As Double, mTOT_Ac_Code$, mCommNarr$
        
        mPostOctraiSaperateYN = VNull(GCn.Execute("Select PostOctraiSaperatelyYN From Syctrl").Fields(0))
        
        Set rsCtrlAc = New ADODB.Recordset
        rsCtrlAc.CursorLocation = adUseClient
        rsCtrlAc.Open "Select Fitment_Ac,Fuel_Ac,VehROff_Ac, OctraiAc, IndirectExpAc, SubventionAc, SubventionClaimAc, RegnFeeAc, InsuranceFeeAc, SpecialDiscountAc From AcControls Where Div_Code='" & PubDivCode & "'", GCnFaV, adOpenStatic, adLockReadOnly
        If rsCtrlAc.RecordCount <= 0 Then
            MsgStr = "Please Add Records in A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        
        If IsNull(rsCtrlAc!Fitment_Ac) Or rsCtrlAc!Fitment_Ac = "" Or _
            IsNull(rsCtrlAc!Fuel_Ac) Or rsCtrlAc!Fuel_Ac = "" Or _
            IsNull(rsCtrlAc!VehROff_Ac) Or rsCtrlAc!VehROff_Ac = "" Then
            MsgStr = "Please define Fitment,Fuel and Round Off A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        
        If Val(txt(SpecialDiscount)) > 0 Then
            If XNull(rsCtrlAc!SpecialDiscountAc) = "" Then
                MsgBox "Please define Special Discount A/c in Vehicle A/c Controls."
                ProcAcPost = False
                GoTo lblExit
            End If
        End If
        
        
        If UCase(left(PubComp_Name, 3)) = "LMP" Then
            If IsNull(rsCtrlAc!IndirectExpAc) Or rsCtrlAc!IndirectExpAc = "" Or _
                IsNull(rsCtrlAc!SubventionClaimAc) Or rsCtrlAc!SubventionClaimAc = "" Or _
                IsNull(rsCtrlAc!SubventionAc) Or rsCtrlAc!SubventionAc = "" Then
                MsgStr = "Please define Indirect Expences, Subvention Ac in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
                ProcAcPost = False
                GoTo lblExit
            End If
        End If
        
        
        If UCase(left(PubComp_Name, 3)) = "LMP" Then
            If IsNull(rsCtrlAc!RegnFeeAc) Or rsCtrlAc!RegnFeeAc = "" Or _
            IsNull(rsCtrlAc!InsuranceFeeAc) Or rsCtrlAc!InsuranceFeeAc = "" Then
                MsgBox "Please Define Registration Fee Ac, Insurance Fee Ac In Vehicle Controls!"
                ProcAcPost = False
                GoTo lblExit
            End If
        End If
        
        

        If IsNull(rsCtrlAc!OctraiAc) Or rsCtrlAc!OctraiAc = "" Then
            MsgStr = "Please define Octrai A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If

        
        rsForm.MoveFirst        'Vehicle Sale A/c Code, Tax A/c Code, Surcharge A/c Code
        rsForm.FIND "Name ='" & txt(FormType) & "'"
        If IsNull(rsForm!PurSal_Ac_Code) Or rsForm!PurSal_Ac_Code = "" Then
            MsgStr = "Please Define Sale A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        
        'Tax A/c Code Checking
        If Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)) <> 0 Then
            If IsNull(rsForm!Tax_Ac_Code) Or rsForm!Sur_Ac_Code = "" Then
                MsgStr = "Please Define Tax A/c in Tax Forms" & vbCrLf & "A/c Posting Aborted !"
                ProcAcPost = False
                GoTo lblExit
            End If
        End If
        
        'Financier A/c Checking
        mTOT_Ac_Code = G_FaCn.Execute("select " & xIsNull("totax_ac", "") & " as TOT_Ac from AcControls where Div_Code='" & PubDivCode & "'").Fields(0).Value
        If Val(txt(TOTAmt)) <> 0 And mTOT_Ac_Code = "" Then
            MsgStr = "Please define TOT A/c Code in Vehicle Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        
        mPostFinAmt = GCn.Execute("select " & vIsNull("PostFinAmt", "0") & " as PostFinAmt from Syctrl").Fields(0).Value
        If mPostFinAmt = 1 And Val(txt(FinAmt)) <> 0 Then
            If txt(FundSource) = "Hypothecation" Or txt(FundSource) = "Hire Purchase" Then
                Set RsTemp = New ADODB.Recordset
                RsTemp.CursorLocation = adUseClient
                If PubBackEnd = "A" Then
                    RsTemp.Open "Select switch(Ac_YN='1','Y',Ac_YN<>'1','N') as ACYN,AcCode From ContractFinance where FinCode='" & txt(FB_Code).Tag & "' ", GCn, adOpenStatic, adLockReadOnly
                Else
                    RsTemp.Open "Select (Case When Ac_YN='1' Then 'Y' When Ac_YN<>'1' Then 'N' End) as ACYN,AcCode From ContractFinance where FinCode='" & txt(FB_Code).Tag & "' ", GCn, adOpenStatic, adLockReadOnly
                End If
                If RsTemp!AcYN = "Y" Then
                    If RsTemp!AcCode = "" Or IsNull(RsTemp!AcCode) Then
                        MsgStr = "Please define A/c Code in Financier Master" & vbCrLf & "A/c Posting Aborted !"
                        GoTo lblExit
                    End If
                End If
            End If
        End If
        If CheckCtrls Then 'Control setting found Ok
            ProcAcPost = True: Exit Function
        End If
        
        'A/c Posting related declarations
        Dim I As Integer, mBookDocID$
        Dim LedgAry(9) As LedgRec, mResult As Byte, mNarr$
        
        'Sale Party A/c
        mBookDocID = GCn.Execute("select OrdDocId from Veh_Order where Inv_DocId='" & txt(TxtDocID) & "'").Fields(0).Value
        mNarr = "By Sales Invoice No." & txt(InvPrefix) & txt(SerialNo) & " Dt. " & txt(VDate) & " Chassis " & txt(ChassisNo)
        mCommNarr = mNarr & "[Common]"
        I = 0
        LedgAry(I).SubCode = txt(Party).Tag
        mGTotAmt = Val(txt(GTotAmt))
        If mPostFinAmt = 0 Then
            mGTotAmt = Val(txt(GTotAmt)) + Val(txt(FinAmt))
        End If
        LedgAry(I).AmtDr = Round(Val(txt(GTotAmt)), 2)
        LedgAry(I).Narration = mNarr
        'Vehicle Sale A/c
        'Modi LPS 05.12.2003
        
        If Val(txt(SubTotA)) + Val(txt(MisCharge)) - Val(txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsForm!PurSal_Ac_Code
            LedgAry(I).AmtCr = Round(Val(txt(SubTotA)) + Val(txt(MisCharge)) - Val(txt(Octroi)), 2)
            LedgAry(I).Narration = mNarr
        End If
        'eof Modi
        'Fitment Amount
        If Val(txt(OthFitAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fitment_Ac
            LedgAry(I).AmtCr = Round(Val(txt(OthFitAmt)), 2)
            LedgAry(I).Narration = mNarr & " Additional Fitments on Vehicle Sale Bill"
        End If
        'Tax Amt
        If Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)) <> 0 Then
            If rsForm!Tax_Ac_Code <> "" And rsForm!Sur_Ac_Code <> "" _
                 And rsForm!Tax_Ac_Code <> rsForm!Sur_Ac_Code Then
                If Val(txt(TaxAmt)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                    LedgAry(I).AmtCr = Round(Val(txt(TaxAmt)) + Val(txt(OthFitTax)), 2)
                    LedgAry(I).Narration = mNarr & " Sale Tax"
                End If
                If Val(txt(TaxSurch)) <> 0 Then
                    I = I + 1
                    LedgAry(I).SubCode = rsForm!Sur_Ac_Code
                    LedgAry(I).AmtCr = Round(Val(txt(TaxSurch)), 2)
                    LedgAry(I).Narration = mNarr & " Surcharge on Sales Tax"
                End If
            Else
                I = I + 1
                LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                LedgAry(I).AmtCr = Round(Val(txt(TaxAmt)) + Val(txt(TaxSurch)) + Val(txt(OthFitTax)), 2)
                LedgAry(I).Narration = mNarr & " Sales Tax & Surcharge"
            End If
        End If
        
        If Val(txt(SatAmt)) <> 0 Then
            If XNull(rsForm!AddTaxAc) <> "" Then
                I = I + 1
                LedgAry(I).SubCode = rsForm!AddTaxAc
                LedgAry(I).AmtCr = Round(Val(txt(SatAmt)), 2)
                LedgAry(I).Narration = mNarr & " Additional Tax"
            End If
        End If
        
        If Val(txt(TOTAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = mTOT_Ac_Code
            LedgAry(I).AmtCr = Val(txt(TOTAmt))
            LedgAry(I).Narration = mNarr & " TOT Amt"
        End If
        If Val(txt(ROff)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!VehROff_Ac
            If Val(txt(ROff)) > 0 Then
                LedgAry(I).AmtCr = Round(Val(txt(ROff)), 2)
            Else
                LedgAry(I).AmtDr = Round(Abs(Val(txt(ROff))), 2)
            End If
            LedgAry(I).Narration = mNarr & " Round Off"
        End If
        'Fuel Amount
        If Val(txt(FuelAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!Fuel_Ac
            LedgAry(I).AmtDr = Round(Val(txt(FuelAmt)), 2)
            LedgAry(I).Narration = mNarr & " Fuel Amount"
        End If
        
        If mPostFinAmt = 1 And Val(txt(FinAmt)) <> 0 Then
            If txt(FundSource) = "Hypothecation" Or txt(FundSource) = "Hire Purchase" Then
                If RsTemp!AcCode = "" Or IsNull(RsTemp!AcCode) Then
                Else
                    I = I + 1
                    LedgAry(I).SubCode = RsTemp!AcCode
                    LedgAry(I).AmtDr = Round(Val(txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amt."
                    I = I + 1
                    LedgAry(I).SubCode = txt(Party).Tag
                    LedgAry(I).AmtCr = Round(Val(txt(FinAmt)), 2)
                    LedgAry(I).Narration = mNarr & " Finance Amount."
                End If
            End If
        End If
        
        If Val(txt(Octroi)) <> 0 Then
            If rsCtrlAc!OctraiAc = "" Or IsNull(rsCtrlAc!OctraiAc) Then
            Else
                I = I + 1
                LedgAry(I).SubCode = rsCtrlAc!OctraiAc
                LedgAry(I).ContraSub = txt(Party).Tag
                LedgAry(I).AmtCr = Round(Val(txt(Octroi)), 2)
                LedgAry(I).Narration = mNarr & " Octroi"
            End If
        End If
                
                
        If Val(txt(RTOfee)) <> 0 Then
            If rsCtrlAc!RegnFeeAc = "" Or IsNull(rsCtrlAc!RegnFeeAc) Then
            Else
                I = I + 1
                LedgAry(I).SubCode = rsCtrlAc!RegnFeeAc
                LedgAry(I).ContraSub = txt(Party).Tag
                LedgAry(I).AmtCr = Round(Val(txt(RTOfee)), 2)
                LedgAry(I).Narration = mNarr & " Registration Fee"
            End If
        End If
                
        If Val(txt(Insurance)) <> 0 Then
            If rsCtrlAc!InsuranceFeeAc = "" Or IsNull(rsCtrlAc!InsuranceFeeAc) Then
            Else
                I = I + 1
                LedgAry(I).SubCode = rsCtrlAc!InsuranceFeeAc
                LedgAry(I).ContraSub = txt(Party).Tag
                LedgAry(I).AmtCr = Round(Val(txt(Insurance)), 2)
                LedgAry(I).Narration = mNarr & " Insurance Fee"
            End If
        End If
                
        
        If Val(txt(SpecialDiscount)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!SpecialDiscountAc
            LedgAry(I).ContraSub = txt(Party).Tag
            LedgAry(I).AmtDr = Round(Val(txt(SpecialDiscount)), 2)
            LedgAry(I).Narration = mNarr & " Special Discount on Vehicle Sale Bill"
            
            
            I = I + 1
            LedgAry(I).SubCode = txt(Party).Tag
            LedgAry(I).ContraSub = rsCtrlAc!SpecialDiscountAc
            LedgAry(I).AmtCr = Round(Val(txt(SpecialDiscount)), 2)
            LedgAry(I).Narration = mNarr & " Special Discount on Vehicle Sale Bill"
            
        End If
        
        
        
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, txt(TxtDocID), CDate(txt(VDate)), mCommNarr)
        If mResult <> 1 Then
            MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
            ProcAcPost = False
        Else
            ProcAcPost = True
        End If
lblExit:
If MsgStr <> "" Then
    MsgBox MsgStr, vbCritical, "A/c Posting"
ElseIf err.NUMBER > 0 Then
    MsgBox err.Description, vbCritical, "A/c Posting"
End If
Set rsCtrlAc = Nothing
Set RsTemp = Nothing
End Function

Public Function GetDocIDVBill(FACn As ADODB.Connection, ByVal VType As String, ByVal VDate As String, _
    ByRef VoucherEditFlag As Boolean, ByRef TxtSrlNo As Object, _
    ByRef lblPrefix As Object, Optional ForSiteCode As String) As String
'FACn As ADODB.Connection,
Dim Rst As ADODB.Recordset, VNo As Long, NotExists As Boolean
Dim TEMPSQL$, DivBaseNumber As Boolean, FaVoucher As Boolean
'12-04-03
'Voucher_Prefix replaced with VehBill_Counter table
'Change in connection CGN to FACn
    If FACn.Execute("Select distinct Category,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'").RecordCount <= 0 Then
        MsgBox "Please Add Record in Voucher Type Table in FA Data" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error": GetDocIDVBill = "": Exit Function
        GetDocIDVBill = ""
        GoTo errlbl
    Else
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Set Rst = FACn.Execute("Select distinct " & cIIF("Category='FA'", cBoolean(True), cBoolean(False)) & " as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
    End If
    FaVoucher = Rst!FaVoucher
    DivBaseNumber = IIf(Rst!DivBaseNumber = 0, False, True)
    If Rst.RecordCount <= 0 Then
        MsgBox "Please Define Document Numbering System  " & vbCrLf & " in Voucher Controls under Utility Menu", vbCritical, "System Configuration"
        GetDocIDVBill = ""
        GoTo errlbl
    End If
    
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    'No Division Base No. in FA /(divison base no introduced by lps at udaipur
    If FaVoucher Then MsgBox "Please Category in Voucher Type Table changed" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error": GetDocIDVBill = "": GoTo errlbl
        'Voucher No's other than FA (Division Base No possible, Voucher No. table from FAData)
        'Voucher No. From FA Data as per connection passed
        TEMPSQL = "Select Top 1 VT.Number_Method,VP.Prefix,VP.Start_Srl_No+1 as Start_Srl_No from Voucher_Type VT Left Join VehBill_Counter VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "' And VP.Prefix='" & lblPrefix & "'"
        If DivBaseNumber Then
            TEMPSQL = TEMPSQL & " and VP.Div_Code='" & PubDivCode & "'"
        End If
        TEMPSQL = TEMPSQL & " and VP.Date_From<=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " Order By VP.Date_From DESC"
        If FACn.Execute(TEMPSQL).RecordCount > 0 Then
            Rst.Open TEMPSQL, FACn, adOpenStatic, adLockReadOnly
        Else
            'Applicable for No Records in Prefix Table & Manual Only
            'Rst.Open "Select VT.Number_Method,VT.SerialNo_From_Table,VT.V_Type From Voucher_Type VT ", FACn, adOpenDynamic, adLockOptimistic
            MsgBox "Please Add Record in Vehicle Bill Counter table " & vbCrLf & "Vehicle Sale Invoice No. Creation failed!", vbCritical, "Fatal Error": GetDocIDVBill = "": Exit Function
            GetDocIDVBill = ""
            GoTo errlbl
        End If
        '*---------
'        lblPrefix = Rst!Prefix
        If IsMissing(ForSiteCode) Then
            ForSiteCode = PubSiteCode
        ElseIf ForSiteCode = "" Then
            ForSiteCode = PubSiteCode
        End If
        If Rst!Number_Method = "Manual" Then
            VoucherEditFlag = True
            TxtSrlNo.Enabled = True
            If Val(TxtSrlNo) > 0 Then
                VNo = Val(TxtSrlNo)
            Else
                VNo = Rst!start_srl_no
            End If
        Else    'Automatic No.
            VoucherEditFlag = False
            If TopCtrl1.TopText2 = "Add" Then
                 TxtSrlNo.Enabled = True
                 VNo = Rst!start_srl_no
            Else
                TxtSrlNo.Enabled = False
                VNo = Val(txt(SerialNo))
            End If
           
        End If
    If Val(txt(SerialNo)) > 0 And Val(txt(SerialNo)) < VNo Then
       VNo = Val(txt(SerialNo))
       TxtSrlNo = VNo
    Else
        TxtSrlNo = VNo
    End If
    GetDocIDVBill = PubDivCode + PubSiteCode + ForSiteCode + Space(5 - Len(CStr(VType))) + VType + Space(5 - Len(CStr(Rst!Prefix))) + Rst!Prefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
errlbl:
    Set Rst = Nothing
End Function

Public Sub AmtCal1()
'Dim SubTot, STax, TaxVal As Double
'    If PubVehRateIncTaxYn = 1 Then
'        SubTot = Val(Txt(SubTotA))
'        STax = Val(Txt(TaxPer)) + Val(Txt(TaxSurPer)) + (Val(Txt(TaxPer)) * Val(Txt(TaxSurPer)) / 100)
'        TaxVal = Val(SubTot) - (Val(SubTot) * Val(STax) / 100 + Val(STax))
'        Txt(SaleRate) = Format(Val(TaxVal) - Val(Txt(Rebate)), "0.00")
'        Txt(SubTotA) = Format(TaxVal, "0.00")
'    End If
End Sub
Private Sub SpeedPrintInvSHANKAR(mQry$)
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
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim mSaleRate As Single, mNetAmt As Single, mInv_No$

     Set Rstsale = GCn.Execute(mQry)
    
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 19
    mFooter = mFooter + FooterCnt
    
    ' Header
    If CancelBillY_N = True Then
        mDocStr = "Sale Invoice (Credit Note)"
    Else
        Dim tmprs As ADODB.Recordset
        Set tmprs = GCn.Execute("Select Description from SubGroupType Left join Subgroup on Subgroup.Party_Type=SubgroupType.Party_Type where Subgroup.SubCode='" & Rstsale!PartyCode & "'")
        If tmprs.RecordCount > 0 Then
            If tmprs!Description = "Dealer" Then
                mDocStr = "Tax Invoice"
            Else
                mDocStr = "Retail Invoice"
            End If
        Else
            mDocStr = "Retail Invoice"
        End If
        
    End If
    mDupStr = IIf(Rstsale!BillPrn_YN = 0, "", " (Duplicate)")
 '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease, 4-Agreement, 5-Lease & Agreement

    If Rstsale!Fund_Source = 0 Then   'Hypothecation
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Hypothecation to  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Leaser  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Lease Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 6 Then
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Loan Cum Hypt. Agreement with  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
    
        
    ElseIf Rstsale!Fund_Source = 1 Or _
           Rstsale!Fund_Source = 4 Or _
           Rstsale!Fund_Source = 5 Then
        
        Left1 = "Sold to under HPA with, "      '1-Hire Purchase
        If Rstsale!Fund_Source = 4 Then         '4-Agreement
            Left1 = "Hire Purchase Finance Agreement with, "
        ElseIf Rstsale!Fund_Source = 5 Then     '5-Lease & Agreement
            Left1 = "Hire Purchase Finance Lease&Agreement with, "
        
        End If
        Left2 = " U/F " & XNull(Rstsale!finbankname)
        Left3 = XNull(Rstsale!FinAdd1)
        Left4 = XNull(Rstsale!FinAdd2)
        Left5 = XNull(Rstsale!FinCity)
        Left6 = ""
        
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!PAdd1)
        Right5 = XNull(Rstsale!PAdd2)
        Right6 = XNull(Rstsale!PAdd3) & XNull(Rstsale!PCityName)
        
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
    End If
    

    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
    mInv_No = Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_No))
    
    '************
    mSaleRate = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate - Rstsale!Subvention + Rstsale!InciChrg _
         + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
         + Rstsale!MVT + Rstsale!Transport
         
    mNetAmt = mSaleRate + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!Tot_Amt + Rstsale!OtherChrg + Rstsale!Fit_Amt _
        + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off
    '***********
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
    Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignRight)
    mHeader = mHeader + 1

    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, mChr18 & mEmph & PSTR(Left1, 40) & PSTR(Right1, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(Left2, 40) & PSTR(Right2, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left3, 40) & PSTR(Right3, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left4, 40) & PSTR(Right4, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left5, 40) & PSTR(Right5, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left6, 40) & PSTR(Right6, 40)
    mHeader = mHeader + 1
        
    Print #1, IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Mfg. Invoice No.: " & XNull(Rstsale!PBILL_NO) & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 40), Space(40)) & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, Space(40) & mEmph & "Invoice Date : " & STR(Rstsale!Inv_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, "Booking No. & Date  : " & STR(Rstsale!Ord_No) & "    " & IIf(IsNull(Rstsale!Ord_Date), "", (STR(Rstsale!Ord_Date)))
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, PSTR("Model : " & Rstsale!Model_Desc, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight)
    mHeader = mHeader + 1
    Print #1, PSTR(Rstsale!Model_Desc1, 45) & PSTR(IIf(Rstsale!Tax_Per = 0, "", "VAT @ " & Format(Rstsale!Tax_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tax_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Colour      : " & Rstsale!Col_Desc, 40) & PSTR(IIf(Rstsale!surcharge_per = 0, "", "Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %"), 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Chassis No. : " & Rstsale!ChassisNo, 45) & PSTR(IIf(Rstsale!TOT_Per = 0, "", "TOT @ " & Format(Rstsale!TOT_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Engine No.  : " & Rstsale!EngineNo, 40) & PSTR("Other Charges", 27, , AlignRight) & ": " & PSTR(Rstsale!OtherChrg, 11, 2)
    mHeader = mHeader + 1
   
    Print #1, "Other Fitments Details : " & mEmph1
    mHeader = mHeader + 1

    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("Sr", 3) & PSTR("Item Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, 2, AlignRight) & " " & PSTR("<----Tax---- >", 13) & " " & PSTR("<-Sur.On Tax- >", 14) & " " & PSTR("Amount", 9, , AlignRight)
    mHeader = mHeader + 1
    Print #1, "No." & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1

    Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")
    Cnt = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            Print #1, mChr17 & STR(Cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) + Rst!Tax_Amt + Rst!TaxSur_Amt), 10, 2)
            mHeader = mHeader + 1
            Cnt = Cnt + 1
            Rst.MoveNext
        Loop
    End If

    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1

'    Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
'        "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
'        "where veh_stock.Chassisno = '" & Txt(ChassisNo) & "' " & _
'        "Group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
'    If Rst.RecordCount > 0 Then
'        Print #1, mDoub & PSTR("Addition/Deletion/Shortage Detail", 52) & PSTR("Qty", 13, , AlignRight) & PSTR("Amount", 15, , AlignRight) & mDoub1
'        mHeader = mHeader + 1
'        Do Until Rst.EOF
'            Print #1, PSTR(IIf(Rst!Trn_Type = "A", "Addition", IIf(Rst!Trn_Type = "D", "Deletion", "Shortage")), 52) & PSTR(Rst!TotQty, 13, 2) & PSTR(Rst!Amt, 15, 2)
'            mHeader = mHeader + 1
'            Rst.MoveNext
'        Loop
'        Print #1, Replace(Space(PageWidth), " ", "-")
'        mHeader = mHeader + 1
'    End If
    Do Until mHeader >= PageLength - mFooter
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    Print #1, PSTR(IIf(Rstsale!Round_off = 0, "", "Round Off"), 65, , AlignRight) & " : " & PSTR(Rstsale!Round_off, 12, 2)
    Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
    Print #1, mEmph & PSTR("Bill Amount", 65, , AlignRight) & " : " & PSTR(Amount_Fill((mNetAmt), PubAmountPrefix), 12, 2, AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, ntow(mNetAmt, "Rupees", "Paise") & mEmph1
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, "Complete With Tools and equipment as supplied by the manufacturer including "
    Print #1, "excise duty,Sales tax & delivery & handing charges."
    Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, ""
    Print #1, "Accountant" & PSTR("Authorised Signatory", PageWidth - 10, , AlignRight)
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")

    Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
        
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    
    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
           
    Print #1, mChr17 & Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
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
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If

    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub


Private Sub SpeedPrintInvJMK(mQry$)
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
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim mSaleRate As Single, mNetAmt As Single, mInv_No$

     Set Rstsale = GCn.Execute(mQry)
    
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 19
    mFooter = mFooter + FooterCnt
    
    ' Header
    If CancelBillY_N = True Then
        mDocStr = "Sale Invoice (Credit Note)"
    Else
        mDocStr = "Sale Invoice"
    End If
    mDupStr = IIf(Rstsale!BillPrn_YN = 0, "", " (Duplicate)")
 '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease, 4-Agreement, 5-Lease & Agreement

    If Rstsale!Fund_Source = 0 Then   'Hypothecation
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Hypothecation to  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Leaser  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Lease Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 6 Then
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        Right1 = "Under Loan Cum Hypt. Agreement with  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
    
        
    ElseIf Rstsale!Fund_Source = 1 Or _
           Rstsale!Fund_Source = 4 Or _
           Rstsale!Fund_Source = 5 Then
        
        Left1 = "Sold to under HPA with, "      '1-Hire Purchase
        If Rstsale!Fund_Source = 4 Then         '4-Agreement
            Left1 = "Hire Purchase Finance Agreement with, "
        ElseIf Rstsale!Fund_Source = 5 Then     '5-Lease & Agreement
            Left1 = "Hire Purchase Finance Lease&Agreement with, "
        
        End If
        Left2 = " U/F " & XNull(Rstsale!finbankname)
        Left3 = XNull(Rstsale!FinAdd1)
        Left4 = XNull(Rstsale!FinAdd2)
        Left5 = XNull(Rstsale!FinCity)
        Left6 = ""
        
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!PAdd1)
        Right5 = XNull(Rstsale!PAdd2)
        Right6 = XNull(Rstsale!PAdd3) & XNull(Rstsale!PCityName)
        
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
    End If
    

    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
    mInv_No = Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_Prefix)) & " - " & Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_No))
    
    '************
    mSaleRate = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate - Rstsale!Subvention + Rstsale!InciChrg _
         + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
         + Rstsale!MVT + Rstsale!Transport
         
    mNetAmt = mSaleRate + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!Tot_Amt + Rstsale!OtherChrg + Rstsale!Fit_Amt _
        + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off
    '***********
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
    Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignRight)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, mChr18 & mEmph & PSTR(Left1, 40) & PSTR(Right1, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, mEmph & PSTR(Left2, 40) & PSTR(Right2, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(Left3, 40) & PSTR(Right3, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left4, 40) & PSTR(Right4, 40)
    mHeader = mHeader + 1
    Print #1, PSTR(Left5, 40) & PSTR(Right5, 40)
    mHeader = mHeader + 1
    'Print #1, PSTR(Left6, 40) & PSTR(Right6, 40)
    'mHeader = mHeader + 1
        
    Print #1, IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Mfg. Invoice No.: " & XNull(Rstsale!PBILL_NO) & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 40), Space(40)) & mEmph & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
    mHeader = mHeader + 1
    Print #1, Space(40) & mEmph & "Invoice Date : " & STR(Rstsale!Inv_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, Space(40) & mEmph & "Booking No. & Date:" & STR(Rstsale!Ord_No) & " " & IIf(IsNull(Rstsale!Ord_Date), "", (STR(Rstsale!Ord_Date))) & mEmph1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, PSTR("Model:" & Rstsale!Model_Desc, 45) & PSTR("Sale Rate", 22, , AlignRight) & ": " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight)
    mHeader = mHeader + 1
    Print #1, PSTR(Rstsale!Model_Desc1, 45) & PSTR(IIf(Rstsale!Tax_Per = 0, "", "Tax @ " & Format(Rstsale!Tax_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tax_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, PSTR("Colour      : " & Rstsale!Col_Desc, 40) & PSTR(IIf(Rstsale!surcharge_per = 0, "", "Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %"), 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, mEmph & PSTR("Chassis No. : " & Rstsale!ChassisNo, 45) & mEmph1 & PSTR(IIf(Rstsale!TOT_Per = 0, "", "SDT @ " & Format(Rstsale!TOT_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2)
    mHeader = mHeader + 1
    Print #1, mEmph & PSTR("Engine No.  : " & Rstsale!EngineNo, 40) & mEmph1 & PSTR("Other Charges", 27, , AlignRight) & ": " & PSTR(Rstsale!OtherChrg, 11, 2)
    mHeader = mHeader + 1
    
    Print #1, "Other Fitments Details : " & mEmph1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("Sr", 3) & PSTR("Item Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, 2, AlignRight) & " " & PSTR("<----Tax---- >", 13) & " " & PSTR("<-Sur.On Tax- >", 14) & " " & PSTR("Amount", 9, , AlignRight)
    mHeader = mHeader + 1
    Print #1, "No." & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1

    Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")
    Cnt = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            Print #1, mChr17 & STR(Cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) + Rst!Tax_Amt + Rst!TaxSur_Amt), 10, 2)
            mHeader = mHeader + 1
            Cnt = Cnt + 1
            Rst.MoveNext
        Loop
    End If

    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1

    Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
        "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where veh_stock.Chassisno = '" & txt(ChassisNo) & "' " & _
        "Group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
    If Rst.RecordCount > 0 Then
        Print #1, mDoub & PSTR("Addition/Deletion/Shortage Detail", 52) & PSTR("Qty", 13, , AlignRight) & PSTR("Amount", 15, , AlignRight) & mDoub1
        mHeader = mHeader + 1
        Do Until Rst.EOF
            Print #1, PSTR(IIf(Rst!Trn_Type = "A", "Addition", IIf(Rst!Trn_Type = "D", "Deletion", "Shortage")), 52) & PSTR(Rst!TotQty, 13, 2) & PSTR(Rst!Amt, 15, 2)
            mHeader = mHeader + 1
            Rst.MoveNext
        Loop
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
    End If
    Do Until mHeader >= PageLength - mFooter - 5
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    Print #1, PSTR(IIf(Rstsale!Round_off = 0, "", "Round Off"), 65, , AlignRight) & " : " & PSTR(Rstsale!Round_off, 12, 2)
    'Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
    Print #1, mEmph & PSTR("Bill Amount", 65, , AlignRight) & " : " & PSTR(Amount_Fill((Round(mNetAmt, 0)), PubAmountPrefix), 12, 2, AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, ntow(mNetAmt, "Rupees", "Paise") & mEmph1
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, "Complete With Tools and equipment as supplied by the manufacturer including "
    Print #1, "excise duty,Sales tax & delivery & handling charges."
    Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
    Print #1, ""
    Print #1, ""
    Print #1, "Accountant               Customer" & PSTR("Authorised Signatory", PageWidth - 33, , AlignRight)
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")

    Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
        
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    
    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
           
    Print #1, mChr17 & Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
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
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If

    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Sub SpeedPrintInvJMA(mQry$)
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
    Dim PrintStr$
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, Rstsale As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity$
    Dim Cnt As Byte, mAmt As Double, PrnStr$, PrnStr1$
    Dim Left1$, Left2$, Left3$
    Dim Left4$, Left5$, Left6$, Left7$
    Dim Right1$, Right2$, Right3$
    Dim Right4$, Right5$, Right6$, Right7$
    Dim mSaleRate As Single, mNetAmt As Single, mInv_No$
Dim mManeyReceipt As ADODB.Recordset
     Set Rstsale = GCn.Execute(mQry)
    
    If Rstsale.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    mJuriCity = XNull(GCn.Execute("Select Juri_city from syctrl").Fields(0).Value)
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehSaleInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 19
    mFooter = mFooter + FooterCnt
    
    ' Header
    If CancelBillY_N = True Then
        mDocStr = "Sale Invoice (Credit Note)"
    Else
'        mDocStr = "Sale Invoice"
    End If
'    mDupStr = IIf(Rstsale!BillPrn_YN = 0, "", " (Duplicate)")
 '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease, 4-Agreement, 5-Lease & Agreement

    If Rstsale!Fund_Source = 0 Then   'Hypothecation
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        'Right1 = "Undr Hypothecation to  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        'Right1 = "Leaser  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Lease Amount :" & Format(Rstsale!Fin_Amt, "0.00")
        
    ElseIf Rstsale!Fund_Source = 6 Then
        Left1 = "To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
        
        'Right1 = "Under Loan Cum Hypt. Agreement with  "
        Right2 = XNull(Rstsale!finbankname)
        Right3 = XNull(Rstsale!FinAdd1)
        Right4 = XNull(Rstsale!FinAdd2)
        Right5 = XNull(Rstsale!FinCity)
        Right6 = "Finance Amount :" & Format(Rstsale!Fin_Amt, "0.00")
    
        
    ElseIf Rstsale!Fund_Source = 1 Or _
           Rstsale!Fund_Source = 4 Or _
           Rstsale!Fund_Source = 5 Then
        
        Left1 = "Sold to under HPA with, "      '1-Hire Purchase
        If Rstsale!Fund_Source = 4 Then         '4-Agreement
            Left1 = "Hire Purchase Finance Agreement with, "
        ElseIf Rstsale!Fund_Source = 5 Then     '5-Lease & Agreement
            Left1 = "Hire Purchase Finance Lease&Agreement with, "
        
        End If
        Left2 = " U/F " & XNull(Rstsale!finbankname)
        Left3 = XNull(Rstsale!FinAdd1)
        Left4 = XNull(Rstsale!FinAdd2)
        Left5 = XNull(Rstsale!FinCity)
        Left6 = ""
        
        'Right1 = "Delivered to Hirer, "
        Right2 = XNull(Rstsale!Name)
        Right3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Right4 = XNull(Rstsale!PAdd1)
        Right5 = XNull(Rstsale!PAdd2)
        Right6 = XNull(Rstsale!PAdd3) & XNull(Rstsale!PCityName)
        
    Else
        Left1 = "Sold To,"
        Left2 = XNull(Rstsale!Name)
        Left3 = XNull(Rstsale!FPrefix) & " " & XNull(Rstsale!fname)
        Left4 = XNull(Rstsale!PAdd1)
        Left5 = XNull(Rstsale!PAdd2)
        Left6 = XNull(Rstsale!PAdd3) & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & IIf(XNull(Rstsale!PCityName) = "" Or XNull(Rstsale!PAdd3) = "", "", ",") & XNull(Rstsale!PCityName)
    End If
    

    Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

    Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
    mInv_No = Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_Prefix)) & " - " & Trim(DeCodeDocID(Rstsale!Inv_DocId, Document_No))
    
    '************
    mSaleRate = Rstsale!vrate + Rstsale!Margine - Rstsale!Rebate - Rstsale!Subvention + Rstsale!InciChrg _
         + Rstsale!Octroi + Rstsale!RegTemp + Rstsale!TransitInsu _
         + Rstsale!MVT + Rstsale!Transport
         
    mNetAmt = mSaleRate + Rstsale!Tax_Amt + Rstsale!Surcharge_Amt _
        + Rstsale!Tot_Amt + Rstsale!OtherChrg + Rstsale!Fit_Amt _
        + Rstsale!Fit_Tax - Rstsale!DieselAmt + Rstsale!Round_off
    '***********
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
   ' Print #1, ""
    mHeader = mHeader + 9
'    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!V_SecSpeciality) <> "" Then
        Print #1, "             " & PRN_TIT(RstCompDet!V_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, "       " & PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
         
    If PubComp_Add2 <> "" Or PubComp_City <> "" Then
        Print #1, "        " & PRN_TIT(PubComp_Add2 & IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, "         " & PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
    mHeader = mHeader + 1
    Print #1, "       " & PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 36) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignRight)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    
    Print #1, "      " & IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Mfg. Invoice No.: " & XNull(Rstsale!PBILL_NO) & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 40), Space(10)) & mEmph & " " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1 & Space(36) & mEmph & " " & STR(Rstsale!Inv_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
     Print #1, ""
    mHeader = mHeader + 1
    
    Print #1, mChr18 & "       " & mEmph & PSTR(Left1, 40) & PSTR(Right1, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, mEmph & "        " & PSTR(Left2, 40) & PSTR(Right2, 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, "       " & PSTR(Left3, 40) & PSTR(Right3, 40)
    mHeader = mHeader + 1
    Print #1, "       " & PSTR(Left4, 40) & PSTR(Right4, 40)
    mHeader = "       " & mHeader + 1
    Print #1, "       " & PSTR(Left5, 40) & PSTR(Right5, 40)
    mHeader = mHeader + 1
        
'    Print #1, IIf(RstInvDet!SupInvOnVehSaleInv = 1, PSTR("Mfg. Invoice No.: " & XNull(Rstsale!PBILL_NO) & IIf(IsNull(Rstsale!PBILL_DATE), "", Rstsale!PBILL_DATE), 40), Space(40)) & mEmph & "Invoice No.  : " & PSTR(mInv_No, 17, , AlignLeft) & mEmph1
'    mHeader = mHeader + 1
'    Print #1, Space(40) & mEmph & "Invoice Date : " & STR(Rstsale!Inv_Date) & mEmph1
'    mHeader = mHeader + 1
''    Print #1, Space(40) & mEmph & "Booking No. & Date:" & STR(Rstsale!Ord_No) & " " & IIf(IsNull(Rstsale!Ord_Date), "", (STR(Rstsale!Ord_Date))) & mEmph1
'    mHeader = mHeader + 1
    'Print #1, Replace(Space(PageWidth), " ", "-")
    'mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
       
    
'    Print #1, "Other Fitments Details : " & mEmph1
'    mHeader = mHeader + 1
'    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
'    mHeader = mHeader + 1
'    Print #1, PSTR("Sr", 3) & PSTR("Item Name", 22) & " " & PSTR("Qty", 3, , AlignRight) & " " & PSTR("Rate", 11, 2, AlignRight) & " " & PSTR("<----Tax---- >", 13) & " " & PSTR("<-Sur.On Tax- >", 14) & " " & PSTR("Amount", 9, , AlignRight)
'    mHeader = mHeader + 1
'    Print #1, "No." & Space(39) & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & " " & PSTR("%", 5, , AlignRight) & " " & PSTR("Amount", 7, , AlignRight) & mDoub1
'    mHeader = mHeader + 1
'    Print #1, Replace(Space(PageWidth), " ", "-")
'    mHeader = mHeader + 1

    Set Rst = GCn.Execute("SELECT Veh_Purch2.TAX_PER, Veh_Purch2.TAX_AMT,Veh_Purch2.docid,Veh_Purch2.rate,Veh_Purch2.QTY, Veh_Purch2.TaxSur_Per, Veh_Purch2.TaxSur_AMT, Veh_AMDModel.Prod_Name " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_Purch2.DocId = '" & Master!SearchCode & "'")
    Cnt = 1
'    If Rst.RecordCount > 0 Then
'        Do Until Rst.EOF
'            Print #1, mChr17 & STR(Cnt) & ". " & PSTR(Rst!Prod_Name, 40) & mChr18 & " " & PSTR(Rst!Qty, 3) & " " & PSTR(Rst!Rate, 11, 2) & " " & PSTR(Rst!Tax_Per, 5, 2) & " " & PSTR(Rst!Tax_Amt, 7, 2) & " " & PSTR(Rst!TaxSur_Per, 5, 2) & " " & PSTR(Rst!TaxSur_Amt, 7, 2) & " " & PSTR(((Rst!Rate * Rst!Qty) + Rst!Tax_Amt + Rst!TaxSur_Amt), 10, 2)
'            mHeader = mHeader + 1
'            Cnt = Cnt + 1
'            Rst.MoveNext
'        Loop
'    End If

     Print #1, ""
     
     mHeader = mHeader + 1
     Print #1, ""
     mHeader = mHeader + 1
'     Print #1, ""
    Print #1, "       " & PSTR("Model:" & Rstsale!Model_Desc, 35) & PSTR("Sale Rate       ", 22, , AlignRight) & ":        " & PSTR(Format(mSaleRate, "0.00"), 11, 2, AlignRight)
    mHeader = mHeader + 1
    Print #1, "       " & PSTR(Rstsale!Model_Desc1, 35) & PSTR(IIf(Rstsale!Tax_Per = 0, "", "Tax @" & Format(Rstsale!Tax_Per, "0.00") & "%" & "     "), 22, , AlignRight) & ":        " & PSTR(Rstsale!Tax_Amt, 11, 2)
    mHeader = mHeader + 1
    'Print #1, "       " & PSTR("Colour       : " & Rstsale!Col_Desc, 30) & PSTR(IIf(Rstsale!surcharge_per = 0, "", "Tax On Surch. @ " & Format(Rstsale!surcharge_per, "0.00") & " %"), 27, , AlignRight) & ": " & PSTR(Rstsale!Surcharge_Amt, 11, 2)
    Print #1, "       " & PSTR("Colour       : " & Rstsale!Col_Desc, 30)
    mHeader = mHeader + 1
    Print #1, "       " & "Fitted With " & Rstsale!Tyres & " Tyre" & "& " & Rstsale!Rims & "Rims  "
     mHeader = mHeader + 1
        Print #1, "       " & Rstsale!TyreDetails & " Price Including Excice duty"
     mHeader = mHeader + 1
    
    'Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1

'    Set Rst = GCn.Execute("SELECT Veh_Purch2.Trn_Type,  sum(Veh_Purch2.QTY) as totqty, sum(Veh_Purch2.QTY * Veh_Purch2.RATE) as amt , Veh_AMDModel.Prod_Name " & _
'        "FROM (Veh_Purch2 LEFT JOIN Veh_Stock ON Veh_Purch2.DocID = Veh_Stock.Pur_DocId) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
'        "where veh_stock.Chassisno = '" & txt(ChassisNo) & "' " & _
'        "Group by Veh_Purch2.Trn_Type,Veh_AMDModel.Prod_Name")
'    If Rst.RecordCount > 0 Then
'        Print #1, "  " & mDoub & PSTR("Addition/Deletion/Shortage Detail", 52) & PSTR("Qty", 13, , AlignRight) & PSTR("Amount", 15, , AlignRight) & mDoub1
'        mHeader = mHeader + 1
'        Do Until Rst.EOF
'            Print #1, "       " & PSTR(IIf(Rst!Trn_Type = "A", "Addition", IIf(Rst!Trn_Type = "D", "Deletion", "Shortage")), 52) & PSTR(Rst!TotQty, 13, 2) & PSTR(Rst!Amt, 15, 2)
'            mHeader = mHeader + 1
'            Rst.MoveNext
'        Loop
'        Print #1, ""
'        mHeader = mHeader + 1
'    End If
    Dim m As Integer
    Dim mBAL As Double
    m = 0
        mBAL = 0
    Set mManeyReceipt = GCn.Execute("SELECT v_date, V_Type, V_No, Amount FROM Rect WHERE Ord_DocId='" & txt(BookNo).Tag & "' ")
    Do Until mHeader >= PageLength - mFooter + 6
          If mManeyReceipt.RecordCount > 0 And m = 7 Then

           Print #1, "        " & PSTR("M.R.No", 8) & Space(5) & PSTR("Date", 10) & PSTR("Amount", 15, , AlignRight)
           mHeader = mHeader + 1
           Do Until mManeyReceipt.EOF
           Print #1, "        " & PSTR(mManeyReceipt!V_NO, 8) & Space(5) & mManeyReceipt!V_DATE & PSTR(mManeyReceipt!Amount, 15, 2)
             mHeader = mHeader + 1
             m = m + 1
             mBAL = mBAL + Val(mManeyReceipt!Amount)
            mManeyReceipt.MoveNext
           Loop
          Else
            m = m + 1
            Print #1, ""
            mHeader = mHeader + 1
            End If
    Loop
If mManeyReceipt.RecordCount > 0 Then
    If mBAL - mNetAmt > 0 Then Print #1, "     " & PSTR("Refundable Balance ", 20, , AlignRight) & " : " & PSTR(mBAL - mNetAmt, 12, 2, AlignRight)
    If mBAL - mNetAmt < 0 Then Print #1, "     " & PSTR("Balance ", 20, , AlignRight) & " : " & PSTR(mNetAmt - mBAL, 12, 2, AlignRight)
    mHeader = mHeader + 1
Else
       Print #1, ""
       mHeader = mHeader + 1
End If
Print #1, ""
mHeader = mHeader + 1
'    Print #1, "     " & PSTR(IIf(Rstsale!TOT_Per = 0, "", "SDT @ " & Format(Rstsale!TOT_Per, "0.00") & " %"), 22, , AlignRight) & ": " & PSTR(Rstsale!Tot_Amt, 11, 2)
    Print #1, "     " & PSTR("Other Charges", 55, , AlignRight) & " :       " & PSTR(Rstsale!OtherChrg, 12, 2)
    mHeader = mHeader + 1
    
    Print #1, " " & PSTR(IIf(Rstsale!Round_off = 0, "", "Round Off"), 55, , AlignRight) & "     :        " & PSTR(Rstsale!Round_off, 12, 2)
    mHeader = mHeader + 1
    'Print #1, PSTR("Less  Fuel Amount", 65, , AlignRight) & " : " & PSTR(Rstsale!DieselAmt, 12, 2)
   
    Print #1, "   " & mEmph & PSTR("Bill Amount", 55, , AlignRight) & "   :        " & PSTR(Amount_Fill((Round(mNetAmt, 0)), PubAmountPrefix), 12, 2, AlignRight)
    mHeader = mHeader + 1
'    Print #1, Replace(Space(PageWidth), " ", "-")
Print #1, ""
    mHeader = mHeader + 1
'        Print #1, "     " & ntow(mNetAmt, "Rupees", "Paise")
'mHeader = mHeader + 1
    'Print #1, Replace(Space(PageWidth), " ", "-")
'    Print #1, ""
    
    
     Print #1, mEmph & "               " & PSTR("                            " & Rstsale!ChassisNo, 45) & "                " & left(ntow(Round(mNetAmt, 0), "Rupees", "Paise"), 36) & mEmph1
    mHeader = mHeader + 1
    Print #1, mEmph & "               " & Space(33) & " " & Trim(mID(ntow(Round(mNetAmt, 0), "Rupees", "Paise"), 37, 30)) & mEmph1
    mHeader = mHeader + 1
    Print #1, mEmph & "               " & PSTR("                            " & Rstsale!EngineNo, 45) & "                " & Trim(mID(ntow(Round(mNetAmt, 0), "Rupees", "Paise"), 67, 30)) & mEmph1
    mHeader = mHeader + 1
    
'    Print #1, "Complete With Tools and equipment as supplied by the manufacturer including "
'    Print #1, "excise duty,Sales tax & delivery & handling charges."
'    Print #1, "E. & OE." & mEmph & PSTR("For " & PubComp_Name, PageWidth - 8, , AlignRight) & mEmph1
'    Print #1, ""
'    Print #1, ""
'    Print #1, "Accountant               Customer" & PSTR("Authorised Signatory", PageWidth - 33, , AlignRight)
'    Print #1, ""
'    Print #1, Replace(Space(PageWidth), " ", "-")
'
'    Print #1, mEmph & "Terms & Condition :" & mEmph1 & mChr17
'
'    Footer = Footer & vbLf
'    j = 1
'    For I = 1 To Len(Footer)
'        If mID(Footer, I, 1) = vbLf Then
'            Print #1, RTrim(mID(Footer, j, I - j))
'            j = I + 1
'        End If
'    Next
    Print #1, ""
    Print #1, ""
    Print #1, ""
    mHeader = mHeader + 3
    'Print #1, mChr18 & Replace(Space(PageWidth), " ", "-") & mChr17
    Print #1, mChr18 & " " & mChr17
           mHeader = mHeader + 1
'    Print #1, mChr17 & Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt) & Space(((PageWidth * 1.7) - Len("* a dataman software *") - Len(Rstsale!Inv_UName & " " & STR(Rstsale!Inv_UEntDt))) / 2) & "* a dataman software *" & mChr18
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
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If

    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub





