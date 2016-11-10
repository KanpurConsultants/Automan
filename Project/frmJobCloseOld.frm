VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmJobClose 
   BackColor       =   &H00D7C6C8&
   Caption         =   "Job Close/Unclose Entry"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11865
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
      Left            =   75
      TabIndex        =   249
      Top             =   2310
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton CmdPBill 
      Caption         =   "Print Privisional Bill"
      Height          =   315
      Left            =   7455
      TabIndex        =   248
      Top             =   45
      Width           =   2700
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
      Height          =   255
      Index           =   84
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   245
      TabStop         =   0   'False
      Text            =   "9999999"
      Top             =   2370
      Width           =   870
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
      Height          =   1530
      Left            =   1110
      TabIndex        =   222
      Top             =   4710
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
         Picture         =   "frmJobClose.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   232
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
         Picture         =   "frmJobClose.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   228
         ToolTipText     =   "Screen"
         Top             =   1170
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmJobClose.frx":0678
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
         Index           =   0
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   231
         ToolTipText     =   "Printer "
         Top             =   885
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmJobClose.frx":0982
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
         Index           =   1
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   230
         ToolTipText     =   "Screen"
         Top             =   585
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmJobClose.frx":0C8C
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
         Index           =   2
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   229
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
         TabIndex        =   235
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
         TabIndex        =   234
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
         TabIndex        =   233
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
         Left            =   120
         TabIndex        =   224
         Top             =   705
         Width           =   1260
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
         Left            =   105
         TabIndex        =   223
         Top             =   345
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.CheckBox ChkRep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Spares"
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
         Index           =   0
         Left            =   1680
         TabIndex        =   225
         Top             =   360
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox ChkRep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Labour"
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
         Index           =   1
         Left            =   1710
         TabIndex        =   226
         Top             =   600
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox ChkRep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Both"
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
         Index           =   2
         Left            =   1710
         TabIndex        =   227
         Top             =   840
         Visible         =   0   'False
         Width           =   1485
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
         TabIndex        =   237
         Top             =   1185
         Width           =   4635
      End
      Begin VB.Label Label2 
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
         Height          =   270
         Left            =   0
         TabIndex        =   236
         Top             =   15
         Width           =   4680
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
      Index           =   67
      Left            =   8130
      Locked          =   -1  'True
      TabIndex        =   243
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
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
      Index           =   83
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   240
      TabStop         =   0   'False
      Top             =   4650
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   1650
      Left            =   -8430
      Negotiate       =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   5130
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   3
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3869.858
         EndProperty
      EndProperty
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
      Index           =   81
      Left            =   10155
      TabIndex        =   69
      ToolTipText     =   "Turn Over Tax %"
      Top             =   4920
      Width           =   540
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   80
      Left            =   8340
      MaxLength       =   5
      TabIndex        =   34
      Text            =   "WithDrawn"
      Top             =   3525
      Width           =   540
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
      Index           =   79
      Left            =   8145
      TabIndex        =   65
      Text            =   "2"
      Top             =   5745
      Width           =   1050
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
      Index           =   78
      Left            =   7560
      TabIndex        =   64
      Text            =   "2"
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5745
      Width           =   540
   End
   Begin VB.TextBox Txt 
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
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   215
      TabStop         =   0   'False
      Top             =   4110
      Width           =   975
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
      Height          =   255
      Index           =   1
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   205
      TabStop         =   0   'False
      Top             =   3180
      Width           =   3885
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
      Index           =   75
      Left            =   5025
      MaxLength       =   40
      TabIndex        =   77
      Top             =   6570
      Width           =   4170
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
      Index           =   74
      Left            =   5025
      MaxLength       =   4
      TabIndex        =   76
      Text            =   "yes"
      Top             =   6300
      Width           =   450
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
      Index           =   68
      Left            =   10770
      TabIndex        =   70
      Text            =   "999999.99"
      Top             =   4920
      Width           =   975
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
      Index           =   82
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   4380
      Width           =   975
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
      Index           =   69
      Left            =   10155
      TabIndex        =   71
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5190
      Width           =   540
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
      Index           =   70
      Left            =   10770
      TabIndex        =   72
      Top             =   5190
      Width           =   975
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   72
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5730
      Width           =   975
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
      Index           =   56
      Left            =   2685
      TabIndex        =   54
      Top             =   6285
      Width           =   1050
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
      Index           =   58
      Left            =   7560
      TabIndex        =   56
      ToolTipText     =   "Local Sales Tax %"
      Top             =   4395
      Width           =   540
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
      Index           =   59
      Left            =   8145
      TabIndex        =   57
      Top             =   4395
      Width           =   1050
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
      Index           =   60
      Left            =   7560
      TabIndex        =   58
      ToolTipText     =   "Surcharge % on Local Sales Tax"
      Top             =   4665
      Width           =   540
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
      Index           =   61
      Left            =   8145
      TabIndex        =   59
      Top             =   4665
      Width           =   1050
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
      Index           =   62
      Left            =   8145
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5205
      Width           =   1050
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
      Index           =   63
      Left            =   7560
      TabIndex        =   62
      Text            =   "99.99"
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5475
      Width           =   540
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
      Index           =   64
      Left            =   8145
      TabIndex        =   63
      Top             =   5475
      Width           =   1050
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
      Index           =   65
      Left            =   8145
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1050
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   66
      Left            =   8145
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Text            =   "9999999.99"
      Top             =   6285
      Width           =   1050
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
      ForeColor       =   &H00C000C0&
      Height          =   255
      Index           =   57
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   6555
      Width           =   1050
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
      Index           =   42
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4935
      Width           =   1050
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
      Index           =   43
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4935
      Width           =   1050
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
      Index           =   44
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5205
      Width           =   1050
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
      Index           =   45
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5205
      Width           =   1050
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
      Left            =   2085
      TabIndex        =   46
      Text            =   "99.99"
      ToolTipText     =   "Discount % Taxable"
      Top             =   5475
      Width           =   540
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
      Index           =   47
      Left            =   2685
      TabIndex        =   47
      Text            =   "99999999.99"
      Top             =   5475
      Width           =   1050
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
      Index           =   48
      Left            =   3870
      TabIndex        =   48
      ToolTipText     =   "Discount % Taxpaid"
      Top             =   5475
      Width           =   540
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
      Left            =   4440
      TabIndex        =   49
      Top             =   5475
      Width           =   1050
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
      Index           =   50
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "99999999.99"
      Top             =   5745
      Width           =   1050
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
      Index           =   51
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5745
      Width           =   1050
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
      Index           =   53
      Left            =   8145
      TabIndex        =   60
      Top             =   4935
      Width           =   1050
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
      Index           =   38
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4395
      Width           =   1050
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
      Index           =   54
      Left            =   2085
      TabIndex        =   52
      ToolTipText     =   "General Surcharge %"
      Top             =   6015
      Width           =   540
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
      Index           =   55
      Left            =   2685
      TabIndex        =   53
      Top             =   6015
      Width           =   1050
   End
   Begin VB.TextBox Txt 
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
      Index           =   39
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4395
      Width           =   1050
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
      Index           =   40
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1050
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
      Index           =   41
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1050
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
      ForeColor       =   &H00800080&
      Height          =   270
      Index           =   73
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Text            =   "999999.99"
      Top             =   6007
      Width           =   975
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
      Index           =   32
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   31
      Text            =   "Help"
      Top             =   3555
      Width           =   3195
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
      Height          =   255
      Index           =   33
      Left            =   1800
      MaxLength       =   40
      TabIndex        =   32
      Text            =   "Help"
      Top             =   3825
      Width           =   2565
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
      Index           =   36
      Left            =   6345
      MaxLength       =   40
      TabIndex        =   36
      Top             =   3795
      Width           =   2535
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
      Height          =   255
      Index           =   34
      Left            =   6345
      TabIndex        =   33
      Top             =   3525
      Width           =   1350
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
      Index           =   35
      Left            =   10890
      MaxLength       =   5
      TabIndex        =   210
      Text            =   "Extra"
      Top             =   525
      Visible         =   0   'False
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
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   10620
      TabIndex        =   37
      Top             =   3525
      Width           =   1110
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
      Height          =   255
      Index           =   28
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1350
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
      Index           =   29
      Left            =   10410
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1290
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
      Height          =   255
      Index           =   26
      Left            =   10545
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "29/Oct/2003"
      Top             =   2100
      Width           =   1155
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
      Index           =   27
      Left            =   10830
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2370
      Width           =   870
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
      Height          =   255
      Index           =   24
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2100
      Width           =   1350
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
      Index           =   25
      Left            =   10830
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1830
      Width           =   870
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
      Height          =   255
      Index           =   23
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "9999999"
      Top             =   1830
      Width           =   870
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
      Height          =   255
      Index           =   21
      Left            =   10830
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1290
      Width           =   870
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
      Index           =   22
      Left            =   10830
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1560
      Width           =   870
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
      Index           =   20
      Left            =   7815
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1170
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
      Height          =   255
      Index           =   19
      Left            =   8115
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1290
      Width           =   870
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
      Height          =   255
      Index           =   4
      Left            =   5985
      TabIndex        =   3
      Top             =   510
      Width           =   1275
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
      Index           =   30
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2910
      Width           =   3885
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
      Index           =   31
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "Help"
      Top             =   3270
      Width           =   3195
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
      Height          =   255
      Index           =   3
      Left            =   3795
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "22-APR-2002"
      Top             =   510
      Width           =   1275
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
      Index           =   15
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2940
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
      Index           =   11
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Help"
      Top             =   1590
      Width           =   4845
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
      Height          =   255
      Index           =   12
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1860
      Width           =   4845
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
      Height          =   255
      Index           =   13
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2130
      Width           =   4845
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
      Height          =   255
      Index           =   14
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4845
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
      Height          =   255
      Index           =   5
      Left            =   1785
      MaxLength       =   14
      TabIndex        =   4
      Text            =   "Help"
      Top             =   780
      Width           =   1470
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
      Height          =   255
      Index           =   2
      Left            =   1350
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "Help"
      Top             =   510
      Width           =   1470
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   661
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
      Index           =   6
      Left            =   4650
      MaxLength       =   15
      TabIndex        =   5
      Text            =   "Help"
      Top             =   780
      Width           =   2145
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
      Index           =   18
      Left            =   4575
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1650
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
      Index           =   17
      Left            =   2925
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1395
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
      Height          =   255
      Index           =   9
      Left            =   1785
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
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
      Height          =   255
      Index           =   8
      Left            =   4650
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1050
      Width           =   2145
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
      Index           =   10
      Left            =   4650
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2145
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
      Index           =   16
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1260
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
      Height          =   255
      Index           =   7
      Left            =   1455
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1050
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid DGMech 
      Height          =   2865
      Left            =   7395
      Negotiate       =   -1  'True
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   7815
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
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3404.977
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGReason 
      Height          =   2865
      Left            =   7815
      Negotiate       =   -1  'True
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   7665
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
         Caption         =   "Reason For Delay"
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
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3404.977
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2865
      Left            =   8370
      Negotiate       =   -1  'True
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   7350
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
            DividerStyle    =   3
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3495.118
         EndProperty
      EndProperty
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
      Index           =   71
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   5460
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Fgrid1 
      Height          =   2745
      Left            =   45
      TabIndex        =   204
      Top             =   7740
      Visible         =   0   'False
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   4842
      _Version        =   393216
      BackColor       =   12243913
      Cols            =   15
      BackColorFixed  =   12648447
      ForeColorFixed  =   128
      BackColorSel    =   16711680
      BackColorBkg    =   12640511
      GridColor       =   16744703
      GridColorFixed  =   12632319
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
      FormatString    =   "hhhhh"
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   52
      Left            =   3030
      TabIndex        =   35
      Text            =   "WithDrawn"
      Top             =   7275
      Visible         =   0   'False
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
      Height          =   255
      Index           =   76
      Left            =   1575
      MaxLength       =   40
      TabIndex        =   78
      Text            =   "Help"
      Top             =   6840
      Width           =   4125
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
      Index           =   77
      Left            =   7620
      MaxLength       =   40
      TabIndex        =   79
      Text            =   "Help"
      Top             =   6840
      Width           =   4125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2745
      Left            =   2535
      TabIndex        =   203
      Top             =   7575
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4842
      _Version        =   393216
      BackColor       =   12632319
      Cols            =   15
      BackColorFixed  =   12640511
      ForeColorFixed  =   16512
      BackColorSel    =   16761024
      BackColorBkg    =   9944522
      GridColor       =   12640511
      GridColorFixed  =   8421631
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
      FormatString    =   "fff"
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      Left            =   7725
      TabIndex        =   247
      Top             =   2370
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon Value"
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
      Index           =   44
      Left            =   6510
      TabIndex        =   246
      Top             =   2385
      Width           =   1170
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Lab. Amount"
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
      Index           =   8
      Left            =   6690
      TabIndex        =   244
      Top             =   4095
      Visible         =   0   'False
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   10530
      TabIndex        =   242
      Top             =   4650
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Lab."
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
      Index           =   7
      Left            =   9330
      TabIndex        =   241
      Top             =   4665
      Width           =   1080
   End
   Begin VB.Line Line6 
      X1              =   11220
      X2              =   11220
      Y1              =   6375
      Y2              =   6825
   End
   Begin VB.Line Line5 
      X1              =   10245
      X2              =   10245
      Y1              =   6360
      Y2              =   6810
   End
   Begin VB.Label LblSprBill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   9420
      TabIndex        =   239
      Top             =   6585
      Width           =   630
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out Side Labour"
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
      Index           =   6
      Left            =   9300
      TabIndex        =   238
      Top             =   4140
      Width           =   1335
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReSale Tax               :"
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
      Index           =   17
      Left            =   5655
      TabIndex        =   221
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5760
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insp.Sheet No."
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
      Left            =   9495
      TabIndex        =   220
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. Vehicle"
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
      Index           =   19
      Left            =   9660
      TabIndex        =   219
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Time"
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
      Left            =   9750
      TabIndex        =   218
      Top             =   1845
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp.Del.Time"
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
      Index           =   25
      Left            =   9615
      TabIndex        =   217
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esti. Labour"
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
      Left            =   9285
      TabIndex        =   216
      Top             =   2655
      Width           =   990
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      X1              =   3900
      X2              =   5490
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line3 
      X1              =   135
      X2              =   6210
      Y1              =   3240
      Y2              =   3240
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
      Index           =   52
      Left            =   10080
      TabIndex        =   214
      Top             =   5190
      Width           =   180
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
      Height          =   225
      Index           =   5
      Left            =   0
      TabIndex        =   213
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   0
      Width           =   1710
   End
   Begin VB.Label lblLabourPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Bill Prefix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   255
      TabIndex        =   212
      Top             =   7545
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblSparePrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Bill Prefix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   255
      TabIndex        =   211
      Top             =   7305
      Visible         =   0   'False
      Width           =   1290
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
      Index           =   60
      Left            =   10545
      TabIndex        =   209
      Top             =   5460
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded Off"
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
      Left            =   9330
      TabIndex        =   208
      Top             =   5475
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body Damage"
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
      Index           =   15
      Left            =   6510
      TabIndex        =   207
      Top             =   3195
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   7710
      TabIndex        =   206
      Top             =   3180
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard DocID :"
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
      Left            =   8220
      TabIndex        =   200
      Top             =   705
      Width           =   1350
   End
   Begin VB.Label lblGatePass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GP No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   11220
      TabIndex        =   199
      Top             =   6585
      Width           =   660
   End
   Begin VB.Label lblLabourBill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   10395
      TabIndex        =   198
      Top             =   6585
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GP No."
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
      Index           =   43
      Left            =   11265
      TabIndex        =   197
      Top             =   6300
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spr Inv No."
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
      Left            =   9315
      TabIndex        =   196
      Top             =   6300
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lab Inv No."
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
      Index           =   41
      Left            =   10290
      TabIndex        =   195
      Top             =   6300
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr A/c (Labour) : "
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
      Height          =   225
      Index           =   21
      Left            =   6210
      TabIndex        =   194
      Top             =   6855
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Party : "
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
      Height          =   225
      Index           =   13
      Left            =   3945
      TabIndex        =   193
      Top             =   6585
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash (Y/N) : "
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
      Height          =   225
      Index           =   11
      Left            =   3945
      TabIndex        =   192
      Top             =   6315
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   15
      Left            =   1965
      TabIndex        =   191
      Top             =   4665
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
      Index           =   17
      Left            =   1965
      TabIndex        =   190
      Top             =   6015
      Width           =   180
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
      Index           =   36
      Left            =   1965
      TabIndex        =   189
      Top             =   4395
      Width           =   180
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
      Index           =   37
      Left            =   7275
      TabIndex        =   188
      Top             =   4935
      Width           =   180
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   38
      Left            =   2940
      TabIndex        =   187
      Top             =   7305
      Visible         =   0   'False
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   39
      Left            =   1965
      TabIndex        =   186
      Top             =   5745
      Width           =   180
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
      Index           =   40
      Left            =   1965
      TabIndex        =   185
      Top             =   5475
      Width           =   180
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   41
      Left            =   1965
      TabIndex        =   184
      Top             =   5205
      Width           =   180
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   42
      Left            =   1965
      TabIndex        =   183
      Top             =   4935
      Width           =   180
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
      ForeColor       =   &H00C000C0&
      Height          =   255
      Index           =   43
      Left            =   1980
      TabIndex        =   182
      Top             =   6555
      Width           =   180
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
      Index           =   50
      Left            =   1980
      TabIndex        =   181
      Top             =   6285
      Width           =   180
   End
   Begin VB.Label lblLabGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Show Labour"
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
      Left            =   10575
      TabIndex        =   180
      ToolTipText     =   "Show Labour (Alt+B)"
      Top             =   3795
      UseMnemonic     =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblSprGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Show Spares"
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
      Left            =   9030
      TabIndex        =   179
      ToolTipText     =   "Show Spares (Alt+P)"
      Top             =   3795
      Width           =   1155
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   3
      Left            =   9330
      TabIndex        =   178
      Top             =   4935
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   55
      Left            =   10095
      TabIndex        =   177
      Top             =   4920
      Width           =   180
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   54
      Left            =   7500
      TabIndex        =   176
      Top             =   6285
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nontaxable Lab."
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
      Index           =   2
      Left            =   9330
      TabIndex        =   175
      Top             =   4395
      Width           =   1365
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Srv. Tax"
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
      Left            =   9330
      TabIndex        =   174
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5205
      Width           =   630
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET LABOUR "
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
      Height          =   225
      Index           =   0
      Left            =   9330
      TabIndex        =   173
      Top             =   5760
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   51
      Left            =   10560
      TabIndex        =   172
      Top             =   5760
      Width           =   45
   End
   Begin VB.Line Line2 
      BorderStyle     =   6  'Inside Solid
      X1              =   135
      X2              =   11730
      Y1              =   4095
      Y2              =   4095
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   21
      Left            =   195
      TabIndex        =   171
      Top             =   6285
      Width           =   1200
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local Sales Tax"
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
      Left            =   5655
      TabIndex        =   170
      Top             =   4410
      Width           =   1305
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
      Index           =   49
      Left            =   7275
      TabIndex        =   169
      Top             =   4395
      Width           =   180
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   23
      Left            =   5655
      TabIndex        =   168
      Top             =   4665
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   48
      Left            =   7275
      TabIndex        =   167
      Top             =   4665
      Width           =   180
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   24
      Left            =   5655
      TabIndex        =   166
      Top             =   5205
      Width           =   1680
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   25
      Left            =   5655
      TabIndex        =   165
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5490
      Width           =   1710
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded Off"
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
      Index           =   26
      Left            =   5655
      TabIndex        =   164
      Top             =   6030
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   45
      Left            =   7275
      TabIndex        =   163
      Top             =   6015
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET PAYABLE :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   32
      Left            =   9330
      TabIndex        =   162
      Top             =   6030
      Width           =   1260
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
      ForeColor       =   &H00C000C0&
      Height          =   225
      Index           =   36
      Left            =   195
      TabIndex        =   161
      Top             =   6570
      Width           =   1365
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable (TB)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   12
      Left            =   2685
      TabIndex        =   160
      Top             =   4140
      Width           =   1035
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Paid (TP)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   13
      Left            =   4455
      TabIndex        =   159
      Top             =   4140
      Width           =   1095
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   14
      Left            =   195
      TabIndex        =   158
      Top             =   5205
      Width           =   930
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   15
      Left            =   195
      TabIndex        =   157
      Top             =   5475
      Width           =   735
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   16
      Left            =   195
      TabIndex        =   156
      Top             =   5745
      Width           =   1080
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   42
      Left            =   2145
      TabIndex        =   155
      Top             =   7305
      Visible         =   0   'False
      Width           =   660
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   18
      Left            =   5655
      TabIndex        =   154
      Top             =   4950
      Width           =   1185
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   20
      Left            =   195
      TabIndex        =   153
      Top             =   6015
      Width           =   1560
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET SPARE/LUB AMT"
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
      Height          =   225
      Index           =   30
      Left            =   5655
      TabIndex        =   152
      Top             =   6300
      Width           =   1770
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   33
      Left            =   195
      TabIndex        =   151
      Top             =   4665
      Width           =   1665
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   19
      Left            =   195
      TabIndex        =   150
      Top             =   4395
      Width           =   1680
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   11
      Left            =   195
      TabIndex        =   149
      Top             =   4935
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor"
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
      Index           =   40
      Left            =   165
      TabIndex        =   148
      Top             =   3555
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
      Height          =   225
      Index           =   35
      Left            =   1095
      TabIndex        =   147
      Top             =   3555
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Remarks"
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
      Index           =   36
      Left            =   4920
      TabIndex        =   146
      Top             =   3795
      Width           =   1305
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
      Index           =   34
      Left            =   6255
      TabIndex        =   145
      Top             =   3795
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Delay Reason"
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
      Index           =   35
      Left            =   165
      TabIndex        =   144
      Top             =   3840
      Width           =   1515
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
      Index           =   33
      Left            =   1725
      TabIndex        =   143
      Top             =   3825
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Index           =   34
      Left            =   7755
      TabIndex        =   142
      Top             =   3525
      Width           =   420
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
      Index           =   32
      Left            =   8205
      TabIndex        =   141
      Top             =   3525
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Completion Date"
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
      Index           =   32
      Left            =   4485
      TabIndex        =   140
      Top             =   3525
      Width           =   1740
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
      Index           =   31
      Left            =   6255
      TabIndex        =   139
      Top             =   3525
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr A/c (Spares) : "
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
      Height          =   225
      Index           =   30
      Left            =   195
      TabIndex        =   138
      Top             =   6855
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Service Date"
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
      Index           =   29
      Left            =   9060
      TabIndex        =   137
      Top             =   3525
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   29
      Left            =   10530
      TabIndex        =   136
      Top             =   3525
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
      Index           =   27
      Left            =   10320
      TabIndex        =   135
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esti. Spares"
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
      Left            =   6675
      TabIndex        =   134
      Top             =   2655
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
      Height          =   255
      Index           =   25
      Left            =   7725
      TabIndex        =   133
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   24
      Left            =   10740
      TabIndex        =   132
      Top             =   2370
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp.Del.Date"
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
      Left            =   9360
      TabIndex        =   131
      Top             =   2115
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   23
      Left            =   10470
      TabIndex        =   130
      Top             =   2100
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
      Left            =   10740
      TabIndex        =   129
      Top             =   1830
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon No."
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
      Left            =   6690
      TabIndex        =   128
      Top             =   2115
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   21
      Left            =   7725
      TabIndex        =   127
      Top             =   2100
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current KMS"
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
      Left            =   6645
      TabIndex        =   126
      Top             =   1845
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
      Index           =   19
      Left            =   7725
      TabIndex        =   125
      Top             =   1830
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
      Index           =   16
      Left            =   10740
      TabIndex        =   124
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   10740
      TabIndex        =   123
      Top             =   1290
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   17
      Left            =   6975
      TabIndex        =   122
      Top             =   1305
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
      Height          =   255
      Index           =   12
      Left            =   8025
      TabIndex        =   121
      Top             =   1290
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   16
      Left            =   7290
      TabIndex        =   120
      Top             =   1560
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
      Height          =   255
      Index           =   10
      Left            =   7725
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   118
      Top             =   510
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Dt."
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
      Left            =   5100
      TabIndex        =   117
      Top             =   510
      Width           =   750
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
      Left            =   7725
      TabIndex        =   116
      Top             =   2910
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Remarks"
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
      Left            =   6420
      TabIndex        =   115
      Top             =   2925
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
      Height          =   255
      Index           =   7
      Left            =   2055
      TabIndex        =   114
      Top             =   3270
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivered by Mechanic"
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
      Index           =   2
      Left            =   165
      TabIndex        =   113
      Top             =   3270
      Width           =   1830
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Dt."
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
      Left            =   2925
      TabIndex        =   112
      Top             =   510
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   3690
      TabIndex        =   111
      Top             =   510
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   26
      Left            =   165
      TabIndex        =   110
      Top             =   1860
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
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
      Index           =   31
      Left            =   165
      TabIndex        =   109
      Top             =   2940
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name"
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
      Index           =   39
      Left            =   165
      TabIndex        =   108
      Top             =   1590
      Width           =   1095
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   10
      Left            =   165
      TabIndex        =   107
      Top             =   2670
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total UnClosed Jobs :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Index           =   7
      Left            =   8220
      TabIndex        =   106
      Top             =   945
      Width           =   1830
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division :"
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
      Left            =   8220
      TabIndex        =   105
      Top             =   480
      Width           =   750
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job DocID:"
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
      Left            =   9630
      TabIndex        =   104
      Top             =   705
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   1305
      TabIndex        =   102
      Top             =   2670
      Width           =   45
   End
   Begin VB.Line Line1 
      X1              =   6210
      X2              =   11745
      Y1              =   3480
      Y2              =   3480
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
      Left            =   1245
      TabIndex        =   101
      Top             =   510
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No."
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
      Index           =   12
      Left            =   165
      TabIndex        =   100
      Top             =   510
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
      Index           =   1
      Left            =   4545
      TabIndex        =   99
      Top             =   780
      Width           =   45
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
      Height          =   255
      Index           =   8
      Left            =   3495
      TabIndex        =   98
      Top             =   780
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(M)"
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
      Index           =   6
      Left            =   4320
      TabIndex        =   97
      Top             =   2940
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
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
      Index           =   5
      Left            =   2655
      TabIndex        =   96
      Top             =   2940
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(O)"
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
      Index           =   4
      Left            =   1095
      TabIndex        =   95
      Top             =   2940
      Width           =   255
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
      Index           =   0
      Left            =   1680
      TabIndex        =   94
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Serial No."
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
      Index           =   0
      Left            =   165
      TabIndex        =   93
      Top             =   1320
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   90
      Left            =   1680
      TabIndex        =   92
      Top             =   780
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No."
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
      Index           =   3
      Left            =   165
      TabIndex        =   91
      Top             =   780
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   855
      Left            =   8160
      Top             =   390
      Width           =   3375
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code :"
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
      TabIndex        =   90
      Top             =   465
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type"
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
      Left            =   3495
      TabIndex        =   89
      Top             =   1320
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
      Index           =   18
      Left            =   1350
      TabIndex        =   88
      Top             =   1050
      Width           =   45
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
      Height          =   255
      Index           =   38
      Left            =   165
      TabIndex        =   87
      Top             =   1050
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
      Height          =   255
      Index           =   26
      Left            =   1305
      TabIndex        =   86
      Top             =   1590
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
      Index           =   8
      Left            =   4545
      TabIndex        =   85
      Top             =   1320
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
      Index           =   11
      Left            =   4545
      TabIndex        =   84
      Top             =   1050
      Width           =   45
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Index           =   33
      Left            =   3495
      TabIndex        =   83
      Top             =   1050
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
      Height          =   255
      Index           =   6
      Left            =   1305
      TabIndex        =   82
      Top             =   1860
      Width           =   45
   End
   Begin VB.Label LblTotVeh 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   10140
      TabIndex        =   81
      Top             =   960
      Width           =   795
   End
End
Attribute VB_Name = "frmJobClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMRevDisTBPer As Double, mMRevDisTPPer As Double
Dim mTBDisAmtMRP As Double, mTPDisAmtMRP As Double
Dim mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double
Dim mMRPLubeTB As Double, mMRPLubeTP  As Double, mLabDiscAmtTB As Single
Private Const mSP2 As String = " "

Private FirstPrint As Boolean
Dim mCardNo$
Public mVType As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function
Dim ForSiteCode As String
Dim SepLabPost As Boolean
Dim MyIndex As Byte
Dim Rst As ADODB.Recordset
Dim mFormCode$
Dim SpareDocID As String
Dim LabourDocID As String
Dim SpareVtype As String
Dim LabourVtype As String
Dim VoucherEditFlag As Boolean
Dim mAddFlag$
Dim Provisional As Boolean

Dim Master As ADODB.Recordset
Dim RsJob As ADODB.Recordset
Dim RsMech As ADODB.Recordset
Dim RsSuper As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim RsReason As ADODB.Recordset
Dim rsCtrlAc As ADODB.Recordset, rsCtrlAcLab As ADODB.Recordset

Private Const BodyDamage As Byte = 1
Private Const JobNo As Byte = 2
Private Const JobDt As Byte = 3
Private Const JobCDt As Byte = 4
Private Const VehRegNo As Byte = 5
Private Const Chassis As Byte = 6
Private Const Model As Byte = 7
Private Const Engine As Byte = 8
Private Const VehSrlNo As Byte = 9
Private Const SrvType As Byte = 10
Private Const OwnerName As Byte = 11
Private Const Address1 As Byte = 12
Private Const Address2 As Byte = 13
Private Const Address3 As Byte = 14
Private Const City As Byte = 15
Private Const PhoneOff As Byte = 16
Private Const PhoneResi As Byte = 17
Private Const Mobile As Byte = 18
Private Const BookNo As Byte = 19
Private Const BookDt As Byte = 20
Private Const InspNo As Byte = 21
Private Const GovtYn As Byte = 22
Private Const CurrentKMS As Byte = 23
Private Const CouponNo As Byte = 24
Private Const ArrTime As Byte = 25
Private Const DelDate As Byte = 26
Private Const DelTime As Byte = 27
Private Const EstSpare As Byte = 28
Private Const EstLabour As Byte = 29
Private Const OpenRemark As Byte = 30
Private Const MechName As Byte = 31
Private Const SuperName As Byte = 32
Private Const JobDelay As Byte = 33
Private Const JobCompDt As Byte = 34
Private Const JobCompTm As Byte = 80
Private Const ExtraField As Byte = 35
Private Const CloseRemark As Byte = 36
Private Const NextSrv As Byte = 37
Private Const IWDiscTotTB As Byte = 38          ' Item-wise Disc Total Taxabl
Private Const IWDiscTotTP As Byte = 39          ' Item-wise Disc Total Taxpaid
Private Const MRPAmtTB As Byte = 40         ' MRP Item's Amount Taxable
Private Const MRPAmtTP As Byte = 41         ' MRP Item's Amount Taxpaid
Private Const SprAmtTB As Byte = 42             ' Spares Amount Taxable
Private Const SprAmtTP As Byte = 43             ' Spares Amount Taxpaid
Private Const OilAmtTB As Byte = 44             ' Oil Amount Taxable
Private Const OilAmtTP As Byte = 45             ' Oil Amount Taxpaid
Private Const DiscPerTB As Byte = 46            '
Private Const DiscAmtTB As Byte = 47            '
Private Const DiscPerTP As Byte = 48            '
Private Const DiscAmtTP As Byte = 49            '
Private Const STotATB As Byte = 50            '
Private Const STotATP As Byte = 51              '
Private Const Addition As Byte = 52            'Witdrawn
Private Const PackCrg As Byte = 53              '
Private Const GenSurPer As Byte = 54           '
Private Const GenSurAmt As Byte = 55           '
Private Const TransAmt As Byte = 56             '
Private Const TaxableTot As Byte = 57           '
Private Const STaxPer As Byte = 58              '
Private Const STaxAmt As Byte = 59              '
Private Const TaxSurPer As Byte = 60            '
Private Const TaxSurAmt As Byte = 61            '
Private Const STotB As Byte = 62                '
Private Const TurnOverPer As Byte = 63          '
Private Const TurnOverAmt As Byte = 64          '
Private Const SROff As Byte = 65                '
Private Const NetSprAmt As Byte = 66
Private Const LabAmt As Byte = 67                '
Private Const LabDisc As Byte = 68          '
Private Const ServTaxPer As Byte = 69          '
Private Const ServTaxAmt As Byte = 70                '
Private Const LabROff As Byte = 71                '
Private Const NetLabAmt As Byte = 72
Private Const NetAmt As Byte = 73               '
Private Const CashBill As Byte = 74          '
Private Const CashParty As Byte = 75          '
Private Const SpareParty As Byte = 76                '
Private Const LabourParty As Byte = 77          '
Private Const ReSalTaxPer As Byte = 78          '
Private Const ReSalTaxAmt As Byte = 79          '
'Private Const LessAdv As Byte = 80          '
Private Const LabDisPer As Byte = 81          '
Private Const OutSideLabAmt As Byte = 0
Private Const LabAmtTP As Byte = 82
Private Const LabAmtTB As Byte = 83
Private Const CouponVal As Byte = 84

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0
Private Const Col_PNo As Byte = 1
Private Const Col_ReqNoDocId As Byte = 2
Private Const Col_ReqDate As Byte = 3
Private Const Col_ReqNo As Byte = 4
Private Const Col_ReqSrNo As Byte = 5
Private Const Col_MRP As Byte = 6
Private Const Col_Taxable As Byte = 7
Private Const Col_Qty As Byte = 8
Private Const Col_Unit As Byte = 9
Private Const Col_Rate As Byte = 10
Private Const Col_MRPRate As Byte = 11
Private Const Col_Amt As Byte = 12
Private Const Col_DiscPer As Byte = 13
Private Const Col_DiscAmt As Byte = 14
Private Const Col_ItemVal As Byte = 15
Private Const Col_Purpose As Byte = 16
Private Const Col_PName As Byte = 17
Private Const Col_LName As Byte = 18
Private Const Col_ClaimNo As Byte = 19
Private Const Col_CompYN As Byte = 20
Private Const Col_PartGrade As Byte = 21

'FGrid1 Columns
Private Const C_LabCode As Byte = 1
Private Const C_LabName As Byte = 2
Private Const C_TaxYN As Byte = 3
Private Const C_PaidBy As Byte = 4  'New
Private Const C_ChrgType As Byte = 5    'New
Private Const C_Hrs As Byte = 6    'New
Private Const C_Rate As Byte = 7    'New
Private Const C_Amt As Byte = 8     'New
Private Const C_External As Byte = 9    'New
Private Const C_GPNo As Byte = 10
Private Const C_Remarks As Byte = 11
Private Const C_ContName As Byte = 12
Private Const C_WIssueDt As Byte = 13
Private Const C_WRecdDt As Byte = 14
Private Const C_ContAmt As Byte = 15

Private Const FromVno As Byte = 0
Private Const ToVno As Byte = 1
Private Const VType1 As Byte = 2
Private Const ChkSprInv As Byte = 0
Private Const ChkLabInv As Byte = 1
Private Const ChkSprBoth As Byte = 2
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim mRepName1 As String
Dim mRepName2 As String

Private Sub CmdPBill_Click()
    If Trim(lblDocId.CAPTION) = "" Or TopCtrl1.TopText2 = "Browse" Then
        MsgBox "No bill selected ! Please select the bill for printing", vbInformation, "Validation Error"
        Exit Sub
    Else
        Provisional = True
        FrmPrn.top = 2220
        FrmPrn.left = (Me.width - FrmPrn.width) / 2
        FrmPrn.Visible = True
        FrmPrn.ZOrder 0
        OptPlain.Value = True
        LblPrinter.CAPTION = Printer.DeviceName
        ChkRep(0).Visible = False
        ChkRep(1).Visible = False
        CmdPrint(0).Visible = False
        CmdPrint(1).Visible = False
    End If

End Sub

Private Sub DGJob_Click()
If Master.RecordCount > 0 Then
    Call History_Field
End If
DGJob.Visible = False
Txt(MyIndex).SetFocus
End Sub

Private Sub DGMech_Click()
If DGMech.Columns(0).CAPTION = "Mechanic Name" Then
    If RsMech.RecordCount > 0 Then
        Txt(MyIndex).TEXT = RsMech!Name
        Txt(MyIndex).Tag = RsMech!Code
    End If
ElseIf DGMech.Columns(0).CAPTION = "WorkShop Staff" Then
    If RsSuper.RecordCount > 0 Then
        Txt(MyIndex).TEXT = RsSuper!Name
        Txt(MyIndex).Tag = RsSuper!Code
    End If
End If
DGMech.Visible = False
Txt(MyIndex).SetFocus
End Sub

Private Sub DGParty_Click()
If RsParty.RecordCount > 0 Then
    Txt(MyIndex).TEXT = RsParty!Name
    Txt(MyIndex).Tag = RsParty!Code
End If
DGParty.Visible = False
lblGroup.Visible = False
Txt(MyIndex).SetFocus
End Sub

Private Sub DGReason_Click()
If RsReason.RecordCount > 0 Then
    Txt(MyIndex).TEXT = RsReason!Name
    Txt(MyIndex).Tag = RsReason!Code
End If
DGReason.Visible = False

Txt(MyIndex).SetFocus
End Sub

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$, key$
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
    If rsCtrlAc.RecordCount <= 0 Then
        MsgStr = "No Records in Spare A/c Controls"
        UnLoadFrm = True
    End If
    If rsCtrlAc!SprSalTP_Ac = "" Or _
        rsCtrlAc!OilSalTP_Ac = "" Or rsCtrlAc!OilSalTP_Ac = "" Or _
        rsCtrlAc!SprCash_Ac = "" Or rsCtrlAc!SprGenSur_Ac = "" Or _
        rsCtrlAc!Transportation_Ac = "" Or rsCtrlAc!ReSaleTax_Ac = "" Or _
        rsCtrlAc!MiscChrg_Ac = "" Or rsCtrlAc!TOTax_Ac = "" Or rsCtrlAc!SprROff_Ac = "" Then
        MsgStr = "Please Fill Spare"
        UnLoadFrm = True
    End If
    'EOF Spare A/c control checking
    'Checking Labour A/c Controls
    If rsCtrlAcLab.RecordCount <= 0 Then
        MsgStr = "No Records in Labour"
        UnLoadFrm = True
    End If
    If rsCtrlAcLab!SrvCash_Ac = "" Or rsCtrlAcLab!SrvLabourTB_Ac = "" Or _
        rsCtrlAcLab!SrvLabour_Ac = "" Or rsCtrlAcLab!SrvTax_Ac = "" Or rsCtrlAcLab!SrvROff_Ac = "" Then
        MsgStr = "Please Fill Labour A/c Controls"
        UnLoadFrm = True
    End If
    If UnLoadFrm Then
        MsgBox "Jobclose Loading Aborted !" & vbCrLf & MsgStr & " A/c Controls through Utility Menu", vbInformation, "Validation"
        Unload Me
    End If
'    key = InputBox("Enter the Key word to Re-Post ", "Key Word")
'    If key = "2827272" Then
'        Do Until Master.EOF
'            Disp_Text SETS("INI", Me, Master)
'            MoveRec
'            TopCtrl1_eEdit
'            TopCtrl1_eSave
'            Master.MoveNext
'        Loop
'    Else
'        MsgBox "Invalid key.Please Contact Dataman for valid Key"
'    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
If Shift = 4 And KeyCode = vbKeyP Then
    Call lblSprGrid_Click
ElseIf Shift = 4 And KeyCode = vbKeyB Then
    Call lblLabGrid_Click
End If
Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim SrNo As Integer
    WinSetting Me:     Ini_Grid
    TopCtrl1.Tag = PubUParam
    ForSiteCode = PubSiteCode
    Txt(JobCDt).Tag = PubLoginDate
    mVType = "W_SIC"
    If pubTOT_On = 1 Then
        Lbl(25) = "TOT on SubTot(BefTax)"
    End If
    If PubReSaleTaxPer = 0 Then
        Lbl(17).Visible = False
        Txt(ReSalTaxPer).Visible = False
        Txt(ReSalTaxAmt).Visible = False
    End If
        
    If UCase(PubSFADataPath) <> UCase(PubWFADataPath) Then
        SepLabPost = True
    End If
    Call BlankText
    lblSprGrid.Tag = 0
    lblLabGrid.Tag = 0
    
    PubOutSideLabDisc = GCn.Execute("select iif(isnull(OutSideLabDisc),0,OutSideLabDisc) as OutSideLabDisc from Syctrl").Fields(0).Value
    PubSrvTaxOnOutSideLab = GCn.Execute("select iif(isnull(SrvTaxOnOutSideLab),0,SrvTaxOnOutSideLab) as SrvTaxOnOutSideLab from Syctrl").Fields(0).Value
    
    'Checking Spare a/c Controls
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
    rsCtrlAc.Open "Select SprGenSur_Ac,ReSaleTax_Ac,SprSalTP_Ac,OilSalTB_Ac,OilSalTP_Ac,SprCash_Ac,SprDiscTB_Ac,Transportation_Ac,MiscChrg_Ac,TOTax_Ac,SprROff_Ac From AcControls Where Div_Code='" & PubDivCode & "'", GCnFaS, adOpenStatic, adLockOptimistic
    'EOF Spare A/c control checking
    'Checking Labour A/c Controls
    Set rsCtrlAcLab = New ADODB.Recordset
    rsCtrlAcLab.CursorLocation = adUseClient
    rsCtrlAcLab.Open "Select SrvCash_Ac,SrvLabourTB_Ac,SrvLabour_Ac,SrvTax_Ac,SrvROff_Ac From AcControls Where Div_Code='" & PubDivCode & "'", GCnFaW, adOpenStatic, adLockOptimistic
    'EOF Labour A/c control checking
    
    ''Pending Points from Syctrl if Set= true : REQ
    ''1. Checking for Bill Serial No. Generation if -
    ''- Labour is not feeded and labour Amount is Zero
    ''- Chargeable Spares not feeded and Total Spares Value is Zero
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Jc.DocId AS CODE " _
            & "from job_card as JC where left(JC.DocId,1)='" & PubDivCode & "' and JC.JobCloseDate<>Null  Order by JC.JobCloseDate desc,JC.DocID desc", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsJob = New ADODB.Recordset
    RsJob.CursorLocation = adUseClient
    RsJob.Open "select Jc.DocId AS CODE,cstr(JC.Job_No) as FindJobNo,JC.Job_No,HC.Model,HC.RegNo,HC.Chassis,HC.Engine,HC.VehSerialNo,HC.Name " & _
                " from (job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) " & _
                " " & _
                " " & _
                " where left(JC.DocID,1)='" & PubDivCode & "' and isnull( JobCloseDate) order by JC.DocID", GCn, adOpenDynamic, adLockOptimistic
    RsJob.Sort = "Code"
    Set DGJob.DataSource = RsJob
    
    Set RsMech = New ADODB.Recordset
    RsMech.CursorLocation = adUseClient
    RsMech.Open "Select Emp_Code as code,Emp_Name as name FROM Emp_Mast where Div_Code='" & PubDivCode & "' And Designation  in (" & pubWrkDesigRest & ") Order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsSuper = New ADODB.Recordset
    RsSuper.CursorLocation = adUseClient
    RsSuper.Open "Select Emp_Code as code,Emp_Name as name FROM Emp_Mast where Div_Code='" & PubDivCode & "' And Designation in ('" & pubWrkDesigSuper & "') Order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsReason = New ADODB.Recordset
    RsReason.CursorLocation = adUseClient
    RsReason.Open "Select Code,R_Desc as name FROM Job_Delay Order by R_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGReason.DataSource = RsReason
    
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup " & _
        "left join [" & PubSFADataPath & "].AcGroup on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where FirmCode = '" & PubFirmCode & _
        "' and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    RsParty.Sort = "Name"
    
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
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
    Set RsMech = Nothing
    Set RsSuper = Nothing
    Set RsReason = Nothing
    Set RsParty = Nothing
    Set rsCtrlAc = Nothing
    Set rsCtrlAcLab = Nothing
End Sub

Private Sub lblSprGrid_Click()
    If Val(lblSprGrid.Tag) = 1 Then
        FGrid.Visible = False       '' spare detail grid
        lblSprGrid.Tag = 0
        lblSprGrid.CAPTION = "Show Spares"
    Else        '' 0
        FGrid1.Visible = False      '' labour detail grid
        lblLabGrid.Tag = 0
        lblLabGrid.CAPTION = "Show Labour"
        
        FGrid.Visible = True       '' spare detail grid
        FGrid.ZOrder 0
        lblSprGrid.Tag = 1
        lblSprGrid.CAPTION = "Hide Spares"
    End If
End Sub

Private Sub lblSprGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSprGrid.BackColor = &HCFE0E0
    lblSprGrid.ForeColor = &H800080
    lblSprGrid.BorderStyle = 1
    lblSprGrid.BackStyle = 1
    lblSprGrid.Font.Bold = True
End Sub

Private Sub lblLabGrid_Click()
    If Val(lblLabGrid.Tag) = 1 Then
        FGrid1.Visible = False       '' Labour detail grid
        lblLabGrid.Tag = 0
        lblLabGrid.CAPTION = "Show Labour"
    Else        '' 0
        FGrid.Visible = False      '' Spare detail grid
        lblSprGrid.Tag = 0
        lblSprGrid.CAPTION = "Show Spares"
        
        FGrid1.Visible = True       '' Labour detail grid
        FGrid1.ZOrder 0
        lblLabGrid.Tag = 1
        lblLabGrid.CAPTION = "Hide Labour"
    End If
End Sub

Private Sub lblLabGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblLabGrid.BackColor = &HCFE0E0
        lblLabGrid.ForeColor = &H800080
        lblLabGrid.BorderStyle = 1
        lblLabGrid.BackStyle = 1
        lblLabGrid.Font.Bold = True
End Sub


Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Call Fgrid_Ini
    mAddFlag = "A"
    Txt(JobCDt).TEXT = Txt(JobCDt).Tag 'Format(Date, "dd/MMM/yyyy")
    Txt(JobCompTm).TEXT = Format(Time, "hh:mm")
    Txt(JobNo).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
Dim I As Integer
Dim mTrans As Boolean, vBook As Variant
Dim LedgAry(1) As LedgRec, mResult As Byte  ', LedgAryLab(1) As LedgRec

On Error GoTo eloop1
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
        GCnFaS.BeginTrans
        GCnFaW.BeginTrans
        mTrans = True
        For I = 1 To FGrid.Rows - 1
            GCn.Execute ("Update SP_Stock Set " _
                        & "Invoice_DocId ='',V_Date2=Null," _
                        & "Rate2=0,MRP_Rate2=0,Disc_Per2=0,Disc_Amt2=0,Amount2=0,Net_Amt2=0 " _
                        & "Where DocID='" & FGrid.TextMatrix(I, Col_ReqNoDocId) & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ReqSrNo)))
        Next
        GCn.Execute "Delete from Sp_Sale where DocID='" & SpareDocID & "'"
        GSQL = "Update Job_Card set JobCloseDate=Null,JobComp_Dt_Time=Null,CrMemo=0,BillingName='',DelBy='',NextSrvDate=Null, " _
            & "docId_invspr='',DocId_InvLab='',gp_no='',DrLab_AcCode='',DrSpr_AcCode='',LabAmt_TB=0,Lab_TaxPer=0,Lab_TaxAmt= 0," _
            & "Lab_D_Amt= 0,Lab_RoundOff= 0,NetLab_Amt= 0,Remark='',DelayReason ='',ClosedU_Name='',ClosedU_EntDt=Null,ClosedU_AE='',LabBillPrinted=0 " _
            & "where DocId='" & Txt(JobNo).Tag & "'"
        GCn.Execute GSQL
        'Unpost Ledger a/c
        If Txt(CashBill).TEXT = "Yes" Then
            ProcAcPost rsCtrlAc, rsCtrlAcLab
        Else
            'to avoid errors of Old System
            LedgerUnPost GCnFaS, Txt(JobNo).Tag
            If SepLabPost Then
                LedgerUnPost GCnFaW, Txt(JobNo).Tag
            End If
            'eof of Old System
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, SpareDocID)
            If mResult <> 1 Then MsgBox "Error in Ledger Un-Posting", vbOKOnly, "Validation"
            If SepLabPost Then
                mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaW, LabourDocID)
                If mResult <> 1 Then MsgBox "Error in Ledger Labour Un-Posting", vbOKOnly, "Validation"
            End If
        End If
        'Unposting of Ledger completed
        
        GCnFaW.CommitTrans
        GCnFaS.CommitTrans
        GCn.CommitTrans
        mTrans = False
        Master.Requery
        Call UpdRequery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        Else
            Call BlankText
        End If
        Call MoveRec
        BUTTONS True, Me, Master, 0
    End If
    Exit Sub
eloop1:
    If mTrans = True Then GCnFaW.RollbackTrans: GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
    Call Fgrid_Ini
    If Master.EOF = True Or Master.BOF = True Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    mAddFlag = "E"
    For I = 1 To 30
        Txt(I).Enabled = False
    Next I
    Txt(CashBill).Enabled = False
    If Txt(CashBill).TEXT = "Yes" Then
        Txt(SpareParty).Enabled = False
        Txt(LabourParty).Enabled = False
        Txt(CashParty).Enabled = True
    Else
        Txt(SpareParty).Enabled = True
        Txt(LabourParty).Enabled = True
        Txt(CashParty).Enabled = False
    End If
    Call txtDisabled_Color
    Txt(MechName).SetFocus
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
    Call Fgrid_Ini
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select Jc.DocId AS SearchCode,JC.Job_No,trim(right(JC.DocId_InvSpr,8)) as Inv_No,trim(mid(Jc.DocId,9,5)) as Prefix,JC.Site_Code,JC.Govt_YN, JC.Job_Date, JC.JobCloseDate,HC.Model,HC.RegNo, HC.Chassis, HC.Engine,HC.VehSerialNo,HC.Name,HC.Add1, HC.Add2, HC.Add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName,jc.OpenRemarks,jc.Body_Damage,Jc.Job_BookNo,Jc.Job_InspDocID,jc.Coupon,jc.ExpDelDate,Jc.DelBy,Jc.RecBy_Supervisor,Jc.DelayReason,Jc.JobComp_Dt_Time,JC.Remark,EM.EMP_NAME AS Mechanic,EMP.Emp_Name as Supervisor," _
                & "jc.CRMemo,jc.BillingName,jc.NetLab_Amt,JC.DocId_InvSpr,Jc.DocId_InvLab,JC.GP_No " _
                & "from (((((job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) left Join Service_Type as ST on JC.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode) left join Emp_Mast as EM on JC.Delby=EM.Emp_Code) left join Emp_Mast as EMP on Jc.RecBy_Supervisor=Emp.Emp_Code) Left Join Job_Delay as JD on JC.DelayReason=JD.Code where left(JC.DocId,1)='" & PubDivCode & "' and jobclosedate<>Null  order by JC.docID"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("Code='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    Call Fgrid_Ini
    Call MoveRec
End Sub

Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    Call Fgrid_Ini
    Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    Call Fgrid_Ini
    Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    Call Fgrid_Ini
    Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
        Call MoveRec
        Disp_Text SETS("INI", Me, Master)
    Else
        Me.ActiveControl.SetFocus
    End If
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Provisional = False
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
ChkRep(0).Visible = True
ChkRep(1).Visible = True
CmdPrint(0).Visible = True
CmdPrint(1).Visible = True
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus

End Sub

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer
Dim mTrans As Boolean, MyGPNo$
'A/c Posting related declarations
Dim rsForm As ADODB.Recordset, mLabAmtTB As Double, mLabAmtTP As Double
Dim mTotSprAmt As Double, mTotOilAmt As Double
Dim mDaySprCash As Double, mDayLabCash As Double, DivBaseNumber As Boolean
'On Error GoTo errlbl
Dim mCurrBal As Double, mEditValue As Double, mCrLimit As Double

    Grid_Hide
    'Check Mechanic in Labour
    GSQL = "Select S_No from Job_Lab where Job_DocID='" & Txt(JobNo).Tag & "' and Job_Lab.S_No not in(Select S_No from Job_Lab2 where Job_DocID='" & Txt(JobNo).Tag & "')"
    If GCn.Execute(GSQL).RecordCount > 0 Then
        MsgBox "Please Enter Mechanic in Labour Done Entry", vbCritical, "Mechanic Name"
        Exit Sub
    End If
    If Val(Txt(CouponVal)) <> 0 Then
        GSQL = "Select count(Lab_Code) from Job_Lab where LabourAmt=" & Val(Txt(CouponVal)) & " and  Chrg_Type='F'"
        If GCn.Execute(GSQL).RecordCount <= 0 Then
            MsgBox "Please Enter Free Service Labour in Labour Done Entry", vbCritical, "Free Service Labour"
            Exit Sub
        End If
    End If
    'eof
    If IsValid(Txt(JobNo), "Job Card No.") = False Then Exit Sub
    If IsValid(Txt(JobCompDt), "Job Completion Date") = False Then Exit Sub
    If Txt(JobCompTm) = "" Or Txt(JobCompTm) = "00:00" Then MsgBox "Job Completion Time is required", vbOKOnly, "Validation": Txt(JobCompTm).SetFocus: Exit Sub
    If IsValid(Txt(MechName), "Mechanic name") = False Then Exit Sub
    If IsValid(Txt(SuperName), "Supervisor name") = False Then Exit Sub
    If Txt(JobDelay).Enabled = True Then
        If IsValid(Txt(JobDelay), "Reason for Job Delay") = False Then Exit Sub
    End If
    If IsValid(Txt(NextSrv), "Next Service Date") = False Then Exit Sub
    If Txt(NextSrv) <> "" Then
        If CDate(Txt(JobCDt)) > CDate(Txt(NextSrv)) Then
            MsgBox "Next Service Date is less than Job close Date", vbOKOnly, "Validation"
            Txt(NextSrv).SetFocus: Exit Sub
        End If
    End If
    If Txt(CashBill).TEXT = "Yes" Then
        If IsValid(Txt(CashParty), "Cash Party Name") = False Then Exit Sub
        Txt(SpareParty).Tag = PubSprCashAc  ' mSprTempAc
        Txt(LabourParty).Tag = PubSrvCashAc ' mSrvTempAc
    Else
        If IsValid(Txt(SpareParty), "Debit A/c Spare Party Name") = False Then Exit Sub
        If IsValid(Txt(LabourParty), "Debit A/c Labour Party Name") = False Then Exit Sub
    End If
    'Check Cr Limit for Challans
    'Temporary disabled at Kota
'    If PubCrLimitCheck = 1 And Txt(CashBill) <> "Yes" Then
'        mCurrBal = 0
'        mEditValue = 0
'        mCurrBal = G_FaCn.Execute("Select Curr_Bal from SubGroup where SubCode='" & Txt(SpareParty).Tag & "'").Fields(0).Value
'        mCrLimit = GCn.Execute("Select CreditLimit from SubGroup where SubCode='" & Txt(SpareParty).Tag & "'").Fields(0).Value
'        If mAddFlag <> "A" Then
'            mEditValue = GCn.Execute("Select Total_Amt from SP_Sale S Where S.DocID = '" & SpareDocID & "'").Fields(0).Value
'            mEditValue = mEditValue + GCn.Execute("Select NetLab_Amt from Job_Card J Where J.DocID = '" & Txt(JobNo).Tag & "'").Fields(0).Value
'        End If
'        mCurrBal = mCurrBal - mEditValue + Val(Txt(NetAmt))
'        If mCurrBal > 0 Then    'Dr Balance
'            If mCurrBal > mCrLimit Then
'                MsgBox "Cr Limit Rs." & mCrLimit & " Exceeds by Rs." & mCurrBal - mCrLimit & vbCrLf & "Add/Edit Denied !", vbInformation, "Cr Limit Checking"
''                Me.ActiveControl.SetFocus: Exit Sub
'            End If
'        End If
'    End If
    'EOF Cr Limit Checking
    
    'Check If Job Closed by another User
    If mAddFlag = "A" Then
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select JobCloseDate,ClosedU_Name,ClosedU_EntDt from Job_Card where Job_Card.DocId='" & Txt(JobNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        If Not IsNull(Rst!JobCloseDate) Then 'Job Closed
            MsgBox "Job Already Closed by User " & Rst!ClosedU_Name & " Dt." & Rst!ClosedU_EntDt
            GoTo lblExit
        End If
    End If
    
    'Add records
    GCn.BeginTrans
    GCnFaS.BeginTrans
    GCnFaW.BeginTrans
    mTrans = True
    mLabAmtTB = Val(Txt(LabAmtTB))
    mLabAmtTP = Val(Txt(LabAmtTP))
    If mAddFlag = "A" Then
        'Creating Bill Numbers
        '' Note: Manual Numbring System for Spares/Labour Bill/Gate pass is Not maintained
        '**********
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Set Rst = GCnFaS.Execute("Select DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & SpareVtype & "'")
        
        If Rst.RecordCount <= 0 Then
            MsgBox "Please Add Record in Voucher_Type Table in FA Data" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error"
            Me.ActiveControl.SetFocus
            GoTo errlbl
        End If
        DivBaseNumber = IIf(Rst!DivBaseNumber = 0, False, True)
        '**********
        '' for Spare Bill Duplicate Check
        GSQL = "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No " & _
            "From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type " & _
            "Where VP.V_Type='" & SpareVtype & "'"
        If DivBaseNumber Then
            GSQL = GSQL & " and VP.Div_Code='" & PubDivCode & "'"
        End If
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open GSQL & " Order By VP.Div_Code,VP.Date_From DESC", GCnFaS, adOpenDynamic, adLockOptimistic
        If Val(Rst!start_srl_no) >= Val(LblSprBill.CAPTION) Then
            SpareDocID = GetDocID(GCnFaS, SpareVtype, Txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
        End If
        If Rst.RecordCount > 0 Then
            GSQL = "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_tYPE & "' "
            If DivBaseNumber Then
                GSQL = GSQL & " and Div_Code ='" & PubDivCode & "'"
            End If
            GSQL = GSQL & " and Date_From=#" & Format(Rst!Date_From, "dd/MMM/yyyy") & "#"
            GCnFaS.Execute GSQL
        End If
        
        '' for Labour Bill Duplicate Check
        GSQL = "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No " & _
            "From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type " & _
            "Where VP.V_Type='" & LabourVtype & "'"
        If DivBaseNumber Then
            GSQL = GSQL & " and VP.Div_Code='" & PubDivCode & "'"
        End If
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open GSQL & " Order By VP.Date_From DESC", GCnFaW, adOpenDynamic, adLockOptimistic
        If Val(Rst!start_srl_no) >= Val(LblSprBill.CAPTION) Then
            LabourDocID = GetDocID(GCnFaW, SpareVtype, Txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
        End If
        If Rst.RecordCount > 0 Then
            GSQL = "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_tYPE & "' "
            If DivBaseNumber Then
                GSQL = GSQL & " and Div_Code ='" & PubDivCode & "'"
            End If
            GSQL = GSQL & " and Date_From=#" & Format(Rst!Date_From, "dd/MMM/yyyy") & "#"
            GCnFaS.Execute GSQL
        End If
        '' gp_no => DivCode(1)+SiteCode(2)+GpNo(5)
        'Creating Gate Pass No.
        MyGPNo = "00000" & GCn.Execute("select iif(isnull(max(val(right(gp_no,5)))),0,max(val(right(gp_no,5))))+1 from job_card where left(gp_no,1)='" & PubDivCode & "' AND mid(gp_no,2,1)='" & PubSiteCode & "'").Fields(0).Value
        MyGPNo = PubDivCode & PubSiteCode & ForSiteCode & Right(MyGPNo, 5)
        lblGatePass = MyGPNo
        '' Job_card Table
        GSQL = "Update Job_Card set JobCloseDate=" & ConvertDate(Txt(JobCDt)) & ",JobComp_Dt_Time=#" & Format(Txt(JobCompDt) & " " & Txt(JobCompTm), "dd/MMM/yyyy hh:mm") & _
            "#,CrMemo=" & IIf(Txt(CashBill) = "Yes", 0, 1) & ",BillingName='" & Txt(CashParty) & "',DelBy='" & Txt(MechName).Tag & "',RecBy_Supervisor='" & Txt(SuperName).Tag & "',NextSrvDate=#" & Txt(NextSrv) & _
            "#,DocId_InvSpr='" & SpareDocID & "',DocId_InvLab='" & LabourDocID & "',GP_NO='" & MyGPNo & "',DrSpr_AcCode='" & Txt(SpareParty).Tag & "',DrLab_AcCode='" & Txt(LabourParty).Tag & _
            "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & Val(Txt(LabDisc)) & ",Lab_TaxPer=" & Val(Txt(ServTaxPer)) & ",Lab_TaxAmt= " & Val(Txt(ServTaxAmt)) & _
            ",Lab_RoundOff= " & Val(Txt(LabROff)) & ",NetLab_Amt= " & Val(Txt(NetLabAmt)) & ",Remark='" & Txt(CloseRemark) & "',DelayReason ='" & Txt(JobDelay).Tag & _
            "',ClosedU_Name='" & pubUName & "',ClosedU_EntDt=#" & PubServerDate & "#,ClosedU_AE='" & left(TopCtrl1.TopText2, 1) & _
            "',LabAmt_Out=" & Val(Txt(OutSideLabAmt)) & " where Job_Card.DocId='" & Txt(JobNo).Tag & "'"
        GCn.Execute GSQL
        
        '' SP_Sale Table
        ' Pending Fields -> LineFileTaxSum,GP_No,GP_Date
        
        GSQL = "Insert Into SP_Sale(" _
            & "DocID ,DocIDHelp ,V_Type ,V_No ,Site_Code ," _
            & "V_Date,Cash_Credit ,Party_Code ,Party_Name ,Job_DocId," _
            & "L_C,Form_Code,CrAc,AcPosting_Yn,Det_Tax,SprAmt_MRP_TB ," _
            & "SprAmt_MRP_TP,OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB,SprAmt_TP ,OilAmt_TB ,OilAmt_TP ," _
            & "D_Per_TB ,D_Amt_TB ,D_Per_TP ,D_Amt_TP ,Addition ," _
            & "Packing ,Gen_Sur_Per ,Gen_Sur_Amt ,Trans_Amt ,Tax_Per ," _
            & "Tax_Amt ,Tax_Sur_Per ,Tax_Sur_Amt ,TOT_Per ,TOT_Amt ," _
            & "ReSalTax_Per, ReSalTax_Amt,Rounded ,Total_Amt, U_Name,U_EntDt,U_AE,D_Per_MRP_TB, " _
            & "D_Amt_MRP_TB, D_Per_MRP_TP, D_Amt_MRP_TP, Tax_AmtMRP, TaxSur_AmtMRP, TOT_AmtMRP, GP_NO, GP_DATE) " _
            & "Values(" _
            & "'" & SpareDocID & "','" & SpareDocID & "','" & SpareVtype & "'," & Val(LblSprBill.CAPTION) & ",'" & PubSiteCode & PubSiteCode & _
            "', " & ConvertDate(Txt(JobCDt)) & ",'" & IIf(Txt(CashBill) = "Yes", "Cash", "Credit") & "','" & Txt(SpareParty).Tag & "','" & IIf(Txt(CashBill) = "Yes", Txt(CashParty), Txt(SpareParty)) & "','" & Txt(JobNo).Tag & _
            "', 'L','" & mFormCode & "','CrAc',1,'" & PubTaxDetOnSprInv & "'," & Val(Txt(MRPAmtTB)) - mMRPLubeTB & _
            " ," & Val(Txt(MRPAmtTP)) - mMRPLubeTP & "," & mMRPLubeTB & "," & mMRPLubeTP & "," & Val(Txt(SprAmtTB)) & "," & Val(Txt(SprAmtTP)) & "," & Val(Txt(OilAmtTB)) & "," & Val(Txt(OilAmtTP)) & _
            " ," & Val(Txt(DiscPerTB)) & "," & Val(Txt(DiscAmtTB)) & "," & Val(Txt(DiscPerTP)) & "," & Val(Txt(DiscAmtTP)) & "," & Val(Txt(Addition)) & _
            " ," & Val(Txt(PackCrg)) & "," & Val(Txt(GenSurPer)) & "," & Val(Txt(GenSurAmt)) & "," & Val(Txt(TransAmt)) & "," & Val(Txt(STaxPer)) & _
            " ," & Val(Txt(STaxAmt)) & "," & Val(Txt(TaxSurPer)) & "," & Val(Txt(TaxSurAmt)) & "," & Val(Txt(TurnOverPer)) & "," & Val(Txt(TurnOverAmt)) & _
            " ," & Val(Txt(ReSalTaxPer)) & "," & Val(Txt(ReSalTaxAmt)) & "," & Val(Txt(SROff)) & "," & Val(Txt(NetSprAmt)) & ",'" & pubUName & "',#" & PubServerDate & "#,'A'," & mMRevDisTBPer & _
            " , " & mTBDisAmtMRP & "," & mMRevDisTPPer & "," & mTPDisAmtMRP & "," & mMRPTax & "," & mMRPTaxSur & ", " & mMRPTOT & ",'" & MyGPNo & "',#" & Format(Txt(JobCDt), "dd/MMM/yyyy") & "#)"
        GCn.Execute GSQL
        '' Sp_Stock Updation
        For I = 1 To FGrid.Rows - 1
            GCn.Execute ("Update SP_Stock Set " _
                & "Invoice_DocId ='" & SpareDocID & "',V_Date2=" & ConvertDate(Txt(JobCDt).TEXT) & "," _
                & "Rate2=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ",MRP_Rate2=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," _
                & "Disc_Per2=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," _
                & "Amount2=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Net_Amt2=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & "," _
                & "U_Name='" & pubUName & "',U_EntDt=#" & PubServerDate & "#,U_AE='E' " _
                & "Where DocID='" & FGrid.TextMatrix(I, Col_ReqNoDocId) & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ReqSrNo)))
        Next
        
    ElseIf mAddFlag = "E" Then
        '' Job_card Table
        GSQL = "Update Job_Card set JobComp_Dt_Time=#" & Format(Txt(JobCompDt) & " " & Txt(JobCompTm), "dd/MMM/yyyy hh:mm") & _
            "#,CrMemo=" & IIf(Txt(CashBill) = "Yes", 0, 1) & ",BillingName='" & Txt(CashParty) & "',DelBy='" & Txt(MechName).Tag & "',RecBy_Supervisor='" & Txt(SuperName).Tag & "',NextSrvDate=#" & Txt(NextSrv) & _
            "#, DrSpr_AcCode='" & Txt(SpareParty).Tag & "',DrLab_AcCode='" & Txt(LabourParty).Tag & _
            "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & Val(Txt(LabDisc)) & ",Lab_TaxPer=" & Val(Txt(ServTaxPer)) & ",Lab_TaxAmt= " & Val(Txt(ServTaxAmt)) & _
            ",Lab_RoundOff= " & Val(Txt(LabROff)) & ",NetLab_Amt= " & Val(Txt(NetLabAmt)) & ",Remark='" & Txt(CloseRemark) & "',DelayReason ='" & Txt(JobDelay).Tag & _
            "',ClosedU_Name='" & pubUName & "',ClosedU_EntDt=#" & PubServerDate & "#,ClosedU_AE='" & left(TopCtrl1.TopText2, 1) & _
            "',LabAmt_Out=" & Val(Txt(OutSideLabAmt)) & "  where Job_Card.DocId='" & Txt(JobNo).Tag & "'"
        GCn.Execute GSQL
        
        GSQL = "update SP_Sale set Det_Tax='" & PubTaxDetOnSprInv & "', Form_Code='" & mFormCode & "', Party_Code='" & Txt(SpareParty).Tag & "',Party_Name='" & IIf(Txt(CashBill) = "Yes", Txt(CashParty), Txt(SpareParty)) & _
            "', SprAmt_MRP_TB=" & Val(Txt(MRPAmtTB)) - mMRPLubeTB & ",SprAmt_MRP_TP=" & Val(Txt(MRPAmtTP)) - mMRPLubeTP & _
            " ,OilAmt_MRP_TB=" & mMRPLubeTB & ",OilAmt_MRP_TP=" & mMRPLubeTP & ",SprAmt_TB=" & Val(Txt(SprAmtTB)) & ",SprAmt_TP=" & Val(Txt(SprAmtTP)) & ",OilAmt_TB=" & Val(Txt(OilAmtTB)) & ",OilAmt_TP=" & Val(Txt(OilAmtTP)) & _
            " ,D_Per_TB=" & Val(Txt(DiscPerTB)) & ",D_Amt_TB=" & Val(Txt(DiscAmtTB)) & ",D_Per_TP=" & Val(Txt(DiscPerTP)) & ",D_Amt_TP=" & Val(Txt(DiscAmtTP)) & ",Addition=" & Val(Txt(Addition)) & _
            " ,Packing=" & Val(Txt(PackCrg)) & ", Gen_Sur_Per=" & Val(Txt(GenSurPer)) & ",Gen_Sur_Amt=" & Val(Txt(GenSurAmt)) & ",Trans_Amt=" & Val(Txt(TransAmt)) & ",Tax_Per=" & Val(Txt(STaxPer)) & _
            " ,Tax_Amt=" & Val(Txt(STaxAmt)) & ", Tax_Sur_Per=" & Val(Txt(TaxSurPer)) & ",Tax_Sur_Amt=" & Val(Txt(TaxSurAmt)) & ",TOT_Per=" & Val(Txt(TurnOverPer)) & ",TOT_Amt=" & Val(Txt(TurnOverAmt)) & _
            " ,ReSalTax_Per=" & Val(Txt(ReSalTaxPer)) & ", ReSalTax_Amt=" & Val(Txt(ReSalTaxAmt)) & ",Rounded=" & Val(Txt(SROff)) & _
            " ,Total_Amt=" & Val(Txt(NetSprAmt)) & ",U_Name='" & pubUName & "',U_EntDt=#" & PubServerDate & "#,U_AE='E'" & _
            " ,D_Per_MRP_TB=" & mMRevDisTBPer & ",D_Amt_MRP_TB=" & mTBDisAmtMRP & ", D_Per_MRP_TP =" & mMRevDisTPPer & ", D_Amt_MRP_TP=" & mTPDisAmtMRP & _
            " ,Tax_AmtMRP=" & mMRPTax & ",TaxSur_AmtMRP= " & mMRPTaxSur & ", TOT_AmtMRP= " & mMRPTOT & _
            " where SP_Sale.DocID='" & SpareDocID & "'"
        GCn.Execute GSQL
        '' Note : Updation of SP_Stock Not required in Edit Mode
        '' Sp_Stock Updation
        For I = 1 To FGrid.Rows - 1
            GCn.Execute ("Update SP_Stock Set " _
                & "Rate2=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ",MRP_Rate2=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," _
                & "Disc_Per2=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," _
                & "Amount2=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Net_Amt2=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & "," _
                & "U_Name='" & pubUName & "',U_EntDt=#" & PubServerDate & "#,U_AE='E' " _
                & "Where DocID='" & FGrid.TextMatrix(I, Col_ReqNoDocId) & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ReqSrNo)))
        Next
    End If
    GCn.Execute ("Update Hiscard set Locked_Text='" & Txt(CloseRemark) & "',LJob_DocId='" & Txt(JobNo).Tag & "',LJob_Date= " & ConvertDate(Txt(JobCDt)) & _
        " where CardNo='" & mCardNo & "'")
    'A/c Posting
    If PubDealerID = "1109800" Then
        If Txt(CashBill) = "Yes" And CDate(Txt(JobCDt)) <= CDate(pubLockDate) Then
'            MsgBox "Job Close Date " & Txt(JobCDt) & " is less than Lock Date " & pubLockDate, vbInformation, "Works Cash Posting Locked"
            GoTo lblExit2
        End If
        ProcAcPost rsCtrlAc, rsCtrlAcLab
    Else
        ProcAcPost rsCtrlAc, rsCtrlAcLab
    End If
    'EOF of A/c Posting Section
lblExit2:
    GCnFaW.CommitTrans
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
lblExit:
    Set Rst = Nothing
    Master.Requery
    Call UpdRequery
    If mAddFlag = "A" Then Txt(JobCDt).Tag = Txt(JobCDt)
    Master.FIND "Code = '" & Txt(JobNo).Tag & "'"
    TopCtrl1_ePrn
    Exit Sub

errlbl:
    If mTrans Then GCnFaS.RollbackTrans: GCnFaW.RollbackTrans: GCn.RollbackTrans
    Set Rst = Nothing
    CheckError
End Sub
Private Sub ProcAcPost(rsCtrlAc As ADODB.Recordset, rsCtrlAcLab As ADODB.Recordset)
'On Error GoTo lblExit
Dim xMRPSprTp As Double, xMRPOilTp As Double
Dim xSprTp As Double, xOilTp As Double
Dim mShare As Single, mShareAmt As Double, mShare2Amt As Double
Dim xNetAmt As Double, xRoundAmt As Double, xSprAmtMRPTB As Double, xSprAmtMRPTP As Double
Dim xOilAmtMRPTB As Double, xOilAmtMRPTP As Double
Dim xSprAmtTB  As Double, xSprAmtTP As Double, xOilAmtTB As Double, xOilAmtTP As Double
Dim xDisAmtTB As Double, xDisAmtTP As Double, xDisAmtMRPTB As Double, xDisAmtMRPTP As Double
Dim xGenSurAmt As Double, xTrans As Double, xTaxAmt As Double, xTaxAmtMRP As Double, xPack As Double
Dim xTurnOver As Double, xReSaleTaxAmt As Double, mFADocidSpr$, mFADocidLab$, mQRY$
Dim xNetLabAmt As Double, xLabAmtTB As Double, xLabAmtTP As Double, xLabDisc As Double
Dim xServTaxAmt As Double, xLabROff As Single
Dim rsTemp As ADODB.Recordset, rsTemp1 As ADODB.Recordset
'A/c Posting related declarations
Dim LedgAry() As LedgRec, LedgAryLab() As LedgRec, mCommNarr$, mLabSQL$
Dim mResult As Byte, mNarr$, TaxSQL$, I As Integer, J As Integer
Dim mSprAmtMRPTB As Double, mSprAmtTB As Double
Dim mOilAmtMRPTB As Double, mOilAmtTB As Double
Dim mTotMRPOilTB As Double, mTotOilTB As Double, mTotShareAmt As Double
Dim mShareSpr As Single, mShareAmtSpr As Double, mShare2AmtSpr As Double
Dim mTot1ShareAmt As Double, mTot2ShareAmt As Double, mTot3ShareAmt As Double
Dim PartyCode$, PartyCodeLab$

    xNetLabAmt = 0
    xLabAmtTB = 0
    xLabAmtTP = 0
    xLabDisc = 0
    xServTaxAmt = 0
    xLabROff = 0
    
    mOilAmtMRPTB = 0
    mSprAmtMRPTB = 0
    mSprAmtTB = 0
    mOilAmtTB = 0


    TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code,sum(Tax_Amt+Tax_AmtMRP) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt " & _
        " from SP_Sale left join TaxFormsAc as TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code"
    'to avoid errors of Old System
'    LedgerUnPost GCnFaS, Txt(JobNo).Tag
'    If SepLabPost Then
'        LedgerUnPost GCnFaW, Txt(JobNo).Tag
'    End If
    'eof of Old System
    If Txt(CashBill) = "Yes" Then
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where V_Date=#" & CDate(Txt(JobCDt)) & "# and left(docid,8)='" & left(SpareDocID, 8) & _
            "' group by TF.PurSal_Ac_Code"
            
        mQRY = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where V_Date= #" & CDate(Txt(JobCDt)) & "# and left(docid,8)='" & left(SpareDocID, 8) & "'"
        'for tax
        TaxSQL = TaxSQL & " where  V_Date=#" & CDate(Txt(JobCDt)) & "# and left(docid,8)='" & left(SpareDocID, 8) & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code"
        '**Labour
        mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
            ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Lab_RoundOff) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
            "from Job_Card where JobCloseDate=#" & CDate(Txt(JobCDt)) & _
            "# and left(DocId_InvLab,8)='" & left(LabourDocID, 8) & "'"
        '***********
        mNarr = "Workshop Cash Sale (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        mFADocidSpr = left(SpareDocID, 8) & "YYYYY" & "  " & Format(Txt(JobCDt), "yymmdd")
        mFADocidLab = left(LabourDocID, 8) & "ZZZZZ" & "  " & Format(Txt(JobCDt), "yymmdd")
        PartyCode = PubSprCashAc
        PartyCodeLab = PubSrvCashAc
    Else
        mFADocidSpr = SpareDocID
        mFADocidLab = LabourDocID
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where docid='" & SpareDocID & _
            "' group by TF.PurSal_Ac_Code"

        mQRY = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "left join TaxFormsAc as TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where docid='" & SpareDocID & "'"
        'for tax
        TaxSQL = TaxSQL & " where docid='" & SpareDocID & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code"
        '**Labour
        mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
            ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Lab_RoundOff) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
            "from Job_Card where format(JobCloseDate,'dd/MMM/yyyy')=" & ConvertDate(Txt(JobCDt)) & _
            " and DocId_InvLab='" & LabourDocID & "'"
        '****
'        mNarr = "Works Cr Spare Bill No. " & Right(SpareDocID, 13) & " Dt." & Txt(JobCDt)
'        If xNetLabAmt <> 0 Then
'            If lblLabourBill <> "" Then
'                mNarr = mNarr & " Labour Bill No. " & lblLabourBill
'            End If
'        End If
'        mCommNarr = mNarr & " [Common]"
        PartyCode = Txt(SpareParty).Tag
        PartyCodeLab = Txt(LabourParty).Tag
    End If
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open mQRY, GCn, adOpenStatic, adLockReadOnly
    
    'for tax purpose
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open TaxSQL, GCn, adOpenStatic, adLockReadOnly

'        1.MRP Spr TB = SprAmtMRPTB-OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP
'        3.MRP Oil TB = OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP

'        2.MRP Spr TP = SprAmtMRPTP-OilAmtMRPTP - part of DisAmtMRPTP
'        4.MRP Oil TP = OilAmtMRPTP - part of DisAmtMRPTP

'        1.Spr TB = SprAmtTB - part of DisAmtTB
'        3.Oil TB = OilAmtTB - part of DisAmtTB
'        2.Spr TP = SprAmtTP - part of DisAmtTP
'        4.Oil TP = OilAmtTP - part of DisAmtTP
    xNetAmt = IIf(IsNull(rsTemp1!NetAmt), 0, rsTemp1!NetAmt): xRoundAmt = IIf(IsNull(rsTemp1!RoundAmt), 0, rsTemp1!RoundAmt)
    xSprAmtMRPTB = IIf(IsNull(rsTemp1!SprAmtMrpTB), 0, rsTemp1!SprAmtMrpTB)
    xOilAmtMRPTB = IIf(IsNull(rsTemp1!OilAmtMrpTB), 0, rsTemp1!OilAmtMrpTB)
    xSprAmtMRPTP = IIf(IsNull(rsTemp1!SprAmtMrpTP), 0, rsTemp1!SprAmtMrpTP)
    xOilAmtMRPTP = IIf(IsNull(rsTemp1!OilAmtMrpTP), 0, rsTemp1!OilAmtMrpTP)
    xSprAmtTB = IIf(IsNull(rsTemp1!SprAmtTB), 0, rsTemp1!SprAmtTB)
    xOilAmtTB = IIf(IsNull(rsTemp1!OilAmtTB), 0, rsTemp1!OilAmtTB)
    xSprAmtTP = IIf(IsNull(rsTemp1!SprAmtTP), 0, rsTemp1!SprAmtTP)
    xOilAmtTP = IIf(IsNull(rsTemp1!OilAmtTP), 0, rsTemp1!OilAmtTP)
    xDisAmtTB = IIf(IsNull(rsTemp1!DisAmtTB), 0, rsTemp1!DisAmtTB)
    xDisAmtTP = IIf(IsNull(rsTemp1!DisAmtTP), 0, rsTemp1!DisAmtTP)
    xDisAmtMRPTB = IIf(IsNull(rsTemp1!DisAmtMRPTB), 0, rsTemp1!DisAmtMRPTB)
    xDisAmtMRPTP = IIf(IsNull(rsTemp1!DisAmtMRPTP), 0, rsTemp1!DisAmtMRPTP)
    xGenSurAmt = IIf(IsNull(rsTemp1!GenSurAmt), 0, rsTemp1!GenSurAmt)
    xTrans = IIf(IsNull(rsTemp1!Trans), 0, rsTemp1!Trans)
    xTaxAmt = IIf(IsNull(rsTemp1!TaxAmt), 0, rsTemp1!TaxAmt)
    xTaxAmtMRP = IIf(IsNull(rsTemp1!TaxAmtMRP), 0, rsTemp1!TaxAmtMRP)
    xPack = IIf(IsNull(rsTemp1!Pack), 0, rsTemp1!Pack)
    xTurnOver = IIf(IsNull(rsTemp1!TurnOver), 0, rsTemp1!TurnOver)
    xReSaleTaxAmt = IIf(IsNull(rsTemp1!ReSaleTaxAmt), 0, rsTemp1!ReSaleTaxAmt)
    '*** Sale Amount Row
    I = 1
    ReDim Preserve LedgAry(1)
    '**Taxable Spr / Oil Calculation
     Do While GRs.EOF = False
        mOilAmtMRPTB = IIf(IsNull(GRs!OilAmtMrpTB), 0, GRs!OilAmtMrpTB)
        mSprAmtMRPTB = IIf(IsNull(GRs!SprAmtMrpTB), 0, GRs!SprAmtMrpTB) ' - mOilAmtMRPTB
        mSprAmtTB = IIf(IsNull(GRs!SprAmtTB), 0, GRs!SprAmtTB)
        mOilAmtTB = IIf(IsNull(GRs!OilAmtTB), 0, GRs!OilAmtTB)
        'Allocate values in their proportions
        If (mSprAmtMRPTB + mOilAmtMRPTB) <> 0 Then
            mShare = Round((mSprAmtMRPTB + mOilAmtMRPTB) * 100 / (xSprAmtMRPTB + xOilAmtMRPTB), 2)
            mShareAmt = Round(xDisAmtMRPTB * mShare / 100, 2)
            mShare2Amt = Round(xTaxAmtMRP * mShare / 100, 2)
            mShareSpr = Round(mSprAmtMRPTB * 100 / (mSprAmtMRPTB + mOilAmtMRPTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mShare2AmtSpr = Round(mShare2Amt * mShareSpr / 100, 2)
            mTot1ShareAmt = mTot1ShareAmt + mShareAmt
            mTot2ShareAmt = mTot2ShareAmt + mShare2Amt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + (xDisAmtMRPTB - mTot1ShareAmt)
                mShare2Amt = mShare2Amt + (xTaxAmtMRP - mTot2ShareAmt)
            End If
            mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr)
            mOilAmtMRPTB = mOilAmtMRPTB - ((mShareAmt + mShare2Amt) - (mShareAmtSpr + mShare2AmtSpr))
        End If
        '*****
        If (mSprAmtTB + mOilAmtTB) <> 0 Then
            mShare = Round((mSprAmtTB + mOilAmtTB) * 100 / (xSprAmtTB + xOilAmtTB), 2)
            mShareAmt = Round((xDisAmtTB - xDisAmtMRPTB) * mShare / 100, 2)
            mShareSpr = Round(mSprAmtTB * 100 / (mSprAmtTB + mOilAmtTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mTot3ShareAmt = mTot3ShareAmt + mShareAmt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + ((xDisAmtTB - xDisAmtMRPTB) - mTot3ShareAmt)
            End If
            mSprAmtTB = mSprAmtTB - (mShareAmtSpr)
            mOilAmtTB = mOilAmtTB - (mShareAmt - mShareAmtSpr)
        End If
        'Spare Sale A/c Taxable
        mTotMRPOilTB = mTotMRPOilTB + mOilAmtMRPTB
        mTotOilTB = mTotOilTB + mOilAmtTB
        '*****
        If mSprAmtMRPTB + mSprAmtTB <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = GRs!PurSal_Ac_Code
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(mSprAmtMRPTB + mSprAmtTB, 2)
            LedgAry(I).Narration = mNarr ' & " Spare"
        End If
        GRs.MoveNext
    Loop
    If (xSprAmtMRPTP + xOilAmtMRPTP) <> 0 Then
        xMRPSprTp = xSprAmtMRPTP '- xOilAmtMRPTP
        xMRPOilTp = xOilAmtMRPTP
        mShare = Round(xMRPSprTp * 100 / (xSprAmtMRPTP + xOilAmtMRPTP), 2)
        mShareAmt = Round(xDisAmtMRPTP * mShare / 100, 2)
        xMRPSprTp = xMRPSprTp - (mShareAmt)
        xMRPOilTp = xMRPOilTp - (xDisAmtMRPTP - (mShareAmt))
    End If
    If (xSprAmtTP + xOilAmtTP) <> 0 Then
        mShare = Round(xSprAmtTP * 100 / (xSprAmtTP + xOilAmtTP), 2)
        mShareAmt = Round((xDisAmtTP - xDisAmtMRPTP) * mShare / 100, 2)
        xSprTp = xSprAmtTP - (mShareAmt)
        xOilTp = xOilAmtTP - ((xDisAmtTP - xDisAmtMRPTP) - (mShareAmt))
    End If
    '*Labour
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open mLabSQL, GCn, adOpenStatic, adLockReadOnly
    If rsTemp1.RecordCount > 0 Then
        xNetLabAmt = IIf(IsNull(rsTemp1!NetLab_Amt), 0, rsTemp1!NetLab_Amt)
        xLabAmtTB = IIf(IsNull(rsTemp1!LabAmt_TB), 0, rsTemp1!LabAmt_TB)
        xLabAmtTP = IIf(IsNull(rsTemp1!LabAmt_TP), 0, rsTemp1!LabAmt_TP)
        xLabDisc = IIf(IsNull(rsTemp1!Lab_D_Amt), 0, rsTemp1!Lab_D_Amt)
        xServTaxAmt = IIf(IsNull(rsTemp1!Lab_TaxAmt), 0, rsTemp1!Lab_TaxAmt)
        xLabROff = IIf(IsNull(rsTemp1!Lab_RoundOff), 0, rsTemp1!Lab_RoundOff)
    End If
    Set rsTemp1 = Nothing
    If (xLabAmtTB + xLabAmtTP) <> 0 Then
        mShare = Round(xLabAmtTB * 100 / (xLabAmtTB + xLabAmtTP), 2)
        mShareAmt = Round(xLabDisc * mShare / 100, 2)
        xLabAmtTB = xLabAmtTB - (mShareAmt)
        xLabAmtTP = xLabAmtTP - (xLabDisc - mShareAmt)
    End If
    '**
'0.Party A/c or Cash A/c
'1.Taxable Spr = MRP Spr TB + SPR TB
'2.Taxpaid Spr = MRP Spr TP + SPR TP
'3.Taxable Oil = MRP Oil TB + Oil TB
'4.Taxable Oil = MRP Oil TP + Oil TP
'5.xGenSurAmt
'6.xPack
'7.xTurnOver
'8.xReSaleTaxAmt
    '*******
    'Sale Party A/c
    'I = 0
    If Txt(CashBill) <> "Yes" Then
        mNarr = "Works Job No. " & PrinID(Txt(JobNo).Tag) & " Cr Spare Bill No. " & PrinID(SpareDocID) & " Dt." & Txt(JobCDt) & " Rs." & Format(xNetAmt, "0.00")
        If xNetLabAmt <> 0 Then
            If lblLabourBill <> "" Then
                mNarr = mNarr & " Labour Bill No. " & PrinID(LabourDocID) & " Rs." & Format(xNetLabAmt, "0.00")
            End If
        End If
        mCommNarr = mNarr & " [Common]"
    End If
    LedgAry(0).SubCode = PartyCode
    LedgAry(0).AmtDr = IIf(SepLabPost, xNetAmt, xNetAmt + xNetLabAmt)
    LedgAry(0).AmtCr = 0
    LedgAry(0).Narration = mNarr
    'Spare Sale A/c Taxpaid
    If xMRPSprTp + xSprTp <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprSalTP_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(xMRPSprTp + xSprTp, 2)
        LedgAry(I).Narration = mNarr ' & " Spare"
    End If
    'Oil Sale A/c Taxable
    If mTotMRPOilTB + mTotOilTB <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!OilSalTB_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mTotMRPOilTB + mTotOilTB, 2)
        LedgAry(I).Narration = mNarr ' & " Spare"
    End If
     'Oil Sale A/c Taxpaid
     If xMRPOilTp + xOilTp <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!OilSalTP_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(xMRPOilTp + xOilTp, 2)
        LedgAry(I).Narration = mNarr ' & " Spare"
     End If
      'GenSurAmt
     If xGenSurAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprGenSur_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(xGenSurAmt, 2)
        LedgAry(I).Narration = mNarr ' & " Sale Tax"
     End If
    'Transportation
     If xTrans <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!Transportation_Ac
        If xTrans > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xTrans, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xTrans), 2)
           LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " Transportation"
     End If
     If rsTemp.RecordCount > 0 Then
         Do While rsTemp.EOF = False
            If rsTemp!TaxAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = rsTemp!Tax_Ac_Code
                If rsTemp!TaxAmt > 0 Then
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(rsTemp!TaxAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(rsTemp!TaxAmt), 2)
                    LedgAry(I).AmtCr = 0
                End If
                LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
            If rsTemp!TaxSurAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = rsTemp!Sur_Ac_Code
                If rsTemp!TaxSurAmt > 0 Then
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(rsTemp!TaxSurAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(rsTemp!TaxSurAmt), 2)
                    LedgAry(I).AmtCr = 0
                End If
                 LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
             End If
             rsTemp.MoveNext
         Loop
     End If
    'Misc / Packing Chrg
    If xPack <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!MiscChrg_Ac
        If Val(xPack) > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xPack, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xPack), 2)
            LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " Misc Charges"
    End If
    If xTurnOver <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!TOTax_Ac
        If xTurnOver > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xTurnOver, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xTurnOver), 2)
            LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " TurnOver Amt"
    End If
    If xReSaleTaxAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!ReSaleTax_Ac
        If xReSaleTaxAmt > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xReSaleTaxAmt, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xReSaleTaxAmt), 2)
            LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " ReSale Tax Amount"
    End If
    '********
    If SepLabPost Then  'Separate Posting for Spr & Labour
        If xRoundAmt <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = rsCtrlAc!SprROff_Ac
            If xRoundAmt > 0 Then
                LedgAry(I).AmtDr = 0
                LedgAry(I).AmtCr = Round(xRoundAmt, 2)
            Else
                LedgAry(I).AmtDr = Round(Abs(xRoundAmt), 2)
                LedgAry(I).AmtCr = 0
            End If
            LedgAry(I).Narration = mNarr '& " Round Off"
        End If
        'Second LedgAryLab
        'Labour Amt
        If xNetLabAmt <> 0 Then
            ReDim Preserve LedgAryLab(0)
            I = 0
            LedgAryLab(I).SubCode = IIf(Txt(CashBill) = "Yes", PubSrvCashAc, Txt(LabourParty).Tag)
            LedgAryLab(I).AmtDr = xNetLabAmt
            LedgAryLab(I).Narration = mNarr & "Labour charges"
            If xLabAmtTB <> 0 Then
                I = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(I)
                LedgAryLab(I).SubCode = rsCtrlAcLab!SrvLabourTB_Ac    'Labour A/c Code
                LedgAryLab(I).AmtCr = xLabAmtTB
                LedgAryLab(I).Narration = mNarr & "Labour charges"
            End If
            If xLabAmtTP <> 0 Then
                I = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(I)
                LedgAryLab(I).SubCode = rsCtrlAcLab!SrvLabour_Ac    'Labour A/c Code
                LedgAryLab(I).AmtCr = xLabAmtTP
                LedgAryLab(I).Narration = mNarr & "Labour charges"
            End If
            'Service Tax
            If xServTaxAmt <> 0 Then
                I = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(I)
                LedgAryLab(I).SubCode = rsCtrlAcLab!SrvTax_Ac    'Service Tax A/c Code
                LedgAryLab(I).AmtCr = xServTaxAmt
                LedgAryLab(I).Narration = mNarr & " Service Tax on Labour charges"
            End If
            'Labour Round Off
            If xLabROff <> 0 Then
                I = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(I)
                LedgAryLab(I).SubCode = rsCtrlAcLab!SrvROff_Ac
                If xLabROff > 0 Then
                    LedgAryLab(I).AmtCr = xLabROff
                Else
                    LedgAryLab(I).AmtDr = Abs(xLabROff)
                End If
                LedgAryLab(I).Narration = mNarr & " Labour Round Diff."
            End If
        End If
    Else    'Combined Posting for Spr & Labour
        'Net Posting Amt = Spr + Labour Amt
        'Labour Taxable
        If xLabAmtTB <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = rsCtrlAcLab!SrvLabourTB_Ac    'Taxable Labour A/c Code
            LedgAry(I).AmtCr = xLabAmtTB
            LedgAry(I).Narration = mNarr & "Labour charges"
        End If
        'Labour Taxpaid
        If xLabAmtTP <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = rsCtrlAcLab!SrvLabour_Ac    'Taxpaid Labour A/c Code
            LedgAry(I).AmtCr = xLabAmtTP
            LedgAry(I).Narration = mNarr & "Labour charges"
        End If
        'Service Tax
        If xServTaxAmt <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = rsCtrlAcLab!SrvTax_Ac    'Service Tax A/c Code
            LedgAry(I).AmtCr = xServTaxAmt
            LedgAry(I).Narration = mNarr & " Service Tax on Labour charges"
        End If
        'Round Off = Spare Round Off + Labour round Off
        If xRoundAmt + xLabROff <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = rsCtrlAc!SprROff_Ac
            If xRoundAmt + xLabROff > 0 Then
                LedgAry(I).AmtCr = xRoundAmt + xLabROff
            Else
                LedgAry(I).AmtDr = Abs(xRoundAmt + xLabROff)
            End If
            LedgAry(I).Narration = mNarr & " Round Diff. Spare+Labour"
        End If
    End If
    mResult = LedgerPost(mAddFlag, LedgAry, GCnFaS, mFADocidSpr, CDate(Txt(JobCDt)), mCommNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    If SepLabPost Then
        mResult = LedgerPost(mAddFlag, LedgAryLab, GCnFaW, mFADocidLab, CDate(Txt(JobCDt)), mCommNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    End If
lblExit:
    Set GRs = Nothing
    Set rsTemp = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description, vbCritical, "Ledger Posting Failed!'"
End Sub

Private Sub Txt_GotFocus(Index As Integer)
'On Error GoTo lblExit
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case JobNo
            DGridColSwap DGJob, 0
            RsJob.Sort = "JOB_NO"   'FindJobNo" '
            If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
'                RsJob.MoveFirst
                RsJob.FIND ("JOB_NO='" & Txt(Index).TEXT & "'")
            End If
        Case Chassis
            DGridColSwap DGJob, 1
            RsJob.Sort = "CHASSIS"
            If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
'                RsJob.MoveFirst
                RsJob.FIND ("CHASSIS='" & Txt(Index).TEXT & "'")
            End If
        Case VehRegNo
            DGridColSwap DGJob, 2
            RsJob.Sort = "REGNO"
            If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("REGNO='" & Txt(Index).TEXT & "'")
            End If
'        Case OwnerName
'            DGridColSwap DGJob, 5
'            RsJob.Sort = "name"
'            If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
'            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
'                RsJob.FIND ("NAME='" & Txt(Index).Text & "'")
'            End If
'NRA Update
        Case MechName
            DGMech.Columns(0).CAPTION = "Mechanic Name"
            Set DGMech.DataSource = RsMech
            DGridColSwap DGMech, 1
            RsMech.Sort = "name"
            If Txt(Index).TEXT <> "" And Txt(Index).Tag <> RsMech!Code Then
                RsMech.FIND ("name='" & Txt(Index).TEXT & "'")
            End If
        Case SuperName
            DGMech.Columns(0).CAPTION = "WorkShop Staff"
            Set DGMech.DataSource = RsSuper
            DGridColSwap DGMech, 1
            RsSuper.Sort = "name"
            If Txt(Index).TEXT <> "" And Txt(Index).Tag <> RsSuper!Code Then
                RsSuper.FIND ("name='" & Txt(Index).TEXT & "'")
            End If
        Case JobDelay
            DGridColSwap DGReason, 1
            RsReason.Sort = "name"
            If Txt(Index).TEXT <> "" And Txt(Index).Tag <> RsReason!Code Then
                RsReason.FIND ("name='" & Txt(Index).TEXT & "'")
            End If
        Case SpareParty
            DGridColSwap DGParty, 1
            RsParty.Sort = "name"
            If Txt(Index).TEXT <> "" And Txt(Index).Tag <> RsParty!Code Then
                RsParty.FIND ("name='" & Txt(Index).TEXT & "'")
            End If
        Case LabourParty
            DGridColSwap DGParty, 1
            RsParty.Sort = "name"
            If Txt(Index).TEXT <> "" And Txt(Index).Tag <> RsParty!Code Then
                RsParty.FIND ("name='" & Txt(Index).TEXT & "'")
            End If
    End Select
lblExit:
    If err.NUMBER <> 0 Then MsgBox err.Description, vbCritical, "Ledger Posting Failed!'"
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobNo
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 1
        Case VehRegNo
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 3
        Case Chassis
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 4
'        Case OwnerName
'            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 7
        Case MechName
            DGridTxtKeyDown DGMech, Txt, Index, RsMech, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case SuperName
            DGridTxtKeyDown DGMech, Txt, Index, RsSuper, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case JobDelay
            DGridTxtKeyDown DGReason, Txt, Index, RsReason, KeyCode, False, 1, frmJobDelay, "frmJobDelay"
        Case SpareParty
            DGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        Case LabourParty
            DGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt
            NumDown Txt(Index), KeyCode, 8, 2
        Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, LabDisPer
            NumDown Txt(Index), KeyCode, 2, 2
    End Select
    If DGJob.Visible = False And DGMech.Visible = False And DGReason.Visible = False And DGParty.Visible = False Then
        '' KEY DOWN
        If KeyCode = vbKeyReturn And lblGroup.Visible = True Then
            lblGroup.Visible = False
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If (Txt(LabourParty).Enabled = True And Index <> LabourParty) Or (Txt(LabourParty).Enabled = False And Index <> CashParty) Then
                Ctrl_DownKeyDown KeyCode, Shift
            End If
            If (Txt(LabourParty).Enabled = True And Index = LabourParty) Or (Txt(LabourParty).Enabled = False And Index = CashParty) Then
                 If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
        End If
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If Index <> JobNo Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index <> MechName Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
        Case JobCompTm
            Call NumPress(Txt(Index), KeyAscii, 2, 2)
        Case JobNo
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "FindJobNo"
        Case VehRegNo
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "regno"
        Case Chassis
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "chassis"
'        Case OwnerName
'            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "name"
        Case MechName, SuperName
            DGridTxtKeyPress Txt, Index, RsMech, KeyAscii, "name"
        Case SuperName
            DGridTxtKeyPress Txt, Index, RsSuper, KeyAscii, "name"
        Case JobDelay
            DGridTxtKeyPress Txt, Index, RsReason, KeyAscii, "name"
        Case SpareParty
             If DGParty.Visible = True Then DGridTxtKeyPress Txt, Index, RsParty, KeyAscii, "name":   lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
        Case LabourParty
            If DGParty.Visible = True Then DGridTxtKeyPress Txt, Index, RsParty, KeyAscii, "name":   lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
            
        Case DiscAmtTB, DiscAmtTP, Addition, PackCrg, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, TurnOverAmt, LabAmt, SROff, LabDisc, ServTaxAmt, NetLabAmt
            NumPress Txt(Index), KeyAscii, 8, 2
        Case LabDisPer, DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ServTaxPer
            NumPress Txt(Index), KeyAscii, 2, 2
        Case CashBill
            If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
                If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                    Txt(Index).TEXT = "Yes"
                    KeyAscii = 0
                    mVType = "W_SIC"
                    LabourVtype = "W_LIC"
                ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                    Txt(Index).TEXT = "No"
                    KeyAscii = 0
                    mVType = "W_SIR"
                    LabourVtype = "W_LIR"
                End If
'                Call Generate_Prefix
                Call Txt_Validate(CashBill, False)
'                Call txtDisabled_Color
            Else
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, SROff, ReSalTaxPer, ReSalTaxAmt
            If Val(Txt(MRPAmtTB)) + Val(Txt(MRPAmtTP)) <> 0 Then
                MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                        Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                        Val(Txt(DiscPerTB)), Val(Txt(DiscPerTP)), _
                        Val(Txt(STaxPer)), Val(Txt(TaxSurPer)), Val(Txt(TurnOverPer))
            End If
            MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_Purpose, True
                ', _
                Txt (LabAmt), Txt(LabDisc), Txt(ServTaxPer), Txt(ServTaxAmt), Txt(LabROff), Txt(NetLabAmt), Txt(OutSideLabAmt)
            'Nra updation
            If Val(Txt(LabAmtTB)) <> 0 Then
                Txt(ServTaxPer) = MainLib.Serv_Tax
            Else
                Txt(ServTaxPer) = Format(0, "0.00")
                Txt(ServTaxAmt) = Format(0, "0.00")
            End If
            'Nra end updation
            
            MainLib.LabCalc Txt(LabAmtTB), Txt(LabAmtTP), Txt(LabDisc), Txt(ServTaxPer), Txt(ServTaxAmt), Txt(LabROff), Txt(NetLabAmt), Txt(OutSideLabAmt), mLabDiscAmtTB
            Txt(NetAmt) = Format(Val(Txt(NetSprAmt)) + Val(Txt(NetLabAmt)), "0.00")
            
        Case LabDisPer, LabAmt, LabDisc, ServTaxPer, ServTaxAmt
            If Index = LabDisPer Then
                If PubOutSideLabDisc = 0 Then   'No
                    Txt(LabDisc) = Round((Val(Txt(LabAmtTB)) + Val(Txt(LabAmtTP)) - Val(Txt(OutSideLabAmt))) * Val(Txt(LabDisPer)) / 100, 2)
                Else
                    Txt(LabDisc) = Round((Val(Txt(LabAmtTB)) + Val(Txt(LabAmtTP))) * Val(Txt(LabDisPer)) / 100, 2)
                    mLabDiscAmtTB = Round(Val(Txt(LabAmtTB)) * Val(Txt(LabDisPer)) / 100, 2)
                End If
            End If
            'Nra updation
            If Val(Txt(LabAmtTB)) <> 0 Then
                Txt(ServTaxPer) = MainLib.Serv_Tax
            Else
                Txt(ServTaxPer) = Format(0, "0.00")
                Txt(ServTaxAmt) = Format(0, "0.00")
            End If
            'Nra end updatio
            MainLib.LabCalc Txt(LabAmtTB), Txt(LabAmtTP), Txt(LabDisc), Txt(ServTaxPer), Txt(ServTaxAmt), Txt(LabROff), Txt(NetLabAmt), Txt(OutSideLabAmt), mLabDiscAmtTB
            Txt(NetAmt) = Format(Val(Txt(NetSprAmt)) + Val(Txt(NetLabAmt)), "0.00")
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case JobCompTm
            Txt(Index) = Format(Txt(Index), "hh:mm")
        Case JobNo, VehRegNo, Chassis ', OwnerName
            If Txt(Index).Tag <> "" Then
                RsJob.Sort = "CODE"
                RsJob.FIND ("CODE='" & Txt(Index).Tag & "'")
            Else
                Cancel = True: Exit Sub
            End If
            If RsJob.BOF = True Or RsJob.EOF = True Then Exit Sub
            
            'External Job Recd Checking
            GSQL = "Select Job_Docid From Job_GatePass as GP " & _
                "Where GP.Job_DocId='" & Txt(JobNo).Tag & "' and ContractRecdDate is null"
            If GCn.Execute(GSQL).RecordCount > 0 Then
                MsgBox "External Job Not Recd !", vbCritical, "Extrenal Job !"
                Cancel = True: Exit Sub
            End If
            'External Job Entry Checking
            GSQL = "Select GatePassNo From Job_GatePass as GP " & _
                "Where GP.Job_DocId='" & Txt(JobNo).Tag & "' and GP.GatePassNo not in (Select distinct ExtJobGatePassNo from Job_Lab as JL where JL.Job_DocId='" & Txt(JobNo).Tag & "')"
            If GCn.Execute(GSQL).RecordCount > 0 Then
                MsgBox "External Job Entry Pending !", vbCritical, "Labour Entry!"
                Cancel = True: Exit Sub
            End If
            'Labour Checking
            GSQL = "Select Job_Docid From Job_Lab JL " & _
                "Where JL.Job_DocId='" & Txt(JobNo).Tag & "'"
            If GCn.Execute(GSQL).RecordCount <= 0 Then
                If MsgBox("Labour Not Feeded !" & vbCrLf & "Continue ?", vbYesNo + vbCritical + vbDefaultButton2, "Labour Not Feeded !") = vbNo Then
                    Cancel = True: Exit Sub
                End If
            End If
            lblDocId = RsJob!Code
            Call History_Field
            Call Fill_Grid
                        '**
            If Index = JobNo Then
                If Txt(Index) <> "" Then
                    Txt(VehRegNo).TabStop = False
                    Txt(Chassis).TabStop = False
                Else
                    Txt(VehRegNo).TabStop = False
                    Txt(Chassis).TabStop = False
                End If
            End If
            '**

        Case MechName
            If RsMech.EOF = True Or RsMech.BOF = True Or Txt(Index).TEXT = "" Then
                Txt(Index).Tag = ""
                Txt(Index).TEXT = ""
            Else
                Txt(Index).Tag = RsMech!Code
                Txt(Index).TEXT = RsMech!Name
            End If
        Case SuperName
            If RsSuper.EOF = True Or RsSuper.BOF = True Or Txt(Index).TEXT = "" Then
                Txt(Index).Tag = ""
                Txt(Index).TEXT = ""
            Else
                Txt(Index).Tag = RsSuper!Code
                Txt(Index).TEXT = RsSuper!Name
            End If
        Case JobDelay
            If RsReason.EOF = True Or RsReason.BOF = True Or Txt(Index).TEXT = "" Then
                Txt(Index).Tag = ""
                Txt(Index).TEXT = ""
            Else
                Txt(Index).Tag = RsReason!Code
                Txt(Index).TEXT = RsReason!Name
            End If
        'Modi LPS 01-04
        Case SpareParty
            If RsParty.EOF = True Or RsParty.BOF = True Or Txt(Index).TEXT = "" Then
                Txt(SpareParty).Tag = ""
                Txt(SpareParty).TEXT = ""
                Txt(LabourParty).Tag = ""
                Txt(LabourParty).TEXT = ""
            Else
                Txt(SpareParty).Tag = RsParty!Code
                Txt(SpareParty).TEXT = RsParty!Name
                Txt(LabourParty).Tag = RsParty!Code
                Txt(LabourParty).TEXT = RsParty!Name
            End If
        Case LabourParty
            If RsParty.EOF = True Or RsParty.BOF = True Or Txt(Index).TEXT = "" Then
                Txt(Index).Tag = ""
                Txt(Index).TEXT = ""
            Else
                Txt(Index).Tag = RsParty!Code
                Txt(Index).TEXT = RsParty!Name
            End If
        'eof modi
        Case JobCDt
            Txt(Index).TEXT = RetDate(Txt(Index))
            Cancel = Not CheckFinYear(Txt(Index))
            If Cancel Then Exit Sub
            GSQL = "Select top 1 v_date from Sp_Stock where Job_Docid='" & Txt(JobNo).Tag & "' and V_Date>#" & Format(Txt(JobCDt), "dd/mmm/yyyy") & "#"
            If GCn.Execute(GSQL).RecordCount > 0 Then
                MsgBox "Job Close Date is Less than Part Issue Date", vbCritical, "Date Checking!"
                Cancel = True: Exit Sub
            End If
            If CDate(Format(Txt(JobCDt), "dd/mm/yyyy")) <= CDate(Format(Txt(DelDate), "dd/mm/yyyy")) Then
                Txt(JobDelay).Enabled = False
                Txt(JobDelay).TEXT = ""
                Txt(JobDelay).Tag = ""
            Else
                Txt(JobDelay).Enabled = True
            End If
            If Txt(NextSrv) = "" Then
                Txt(NextSrv) = CDate(Txt(JobCDt)) + PubNextSrvDays
            End If
            Txt(JobCompDt).TEXT = Format(Txt(JobCDt), "dd/MMM/yyyy")
        Case JobCompDt
            Txt(Index).TEXT = RetDate(Txt(Index))
        Case NextSrv
            Txt(Index).TEXT = RetDate(Txt(Index))
        Case LabDisPer, DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, SROff, LabAmt, LabDisc, ServTaxPer, ServTaxAmt
            If Val(Txt(Index).TEXT) = 0 Then
                Txt(Index).TEXT = ""
            Else
                Txt(Index).TEXT = Format(Txt(Index), "0.00")
            End If
        Case CashBill
            If Txt(Index).TEXT = "Yes" Then
                Txt(CashParty) = Txt(OwnerName)
            End If
            Call Generate_Prefix
            Call txtDisabled_Color
        Case CashParty
            If TopCtrl1.TopText2 = "Add" Then
                Call Generate_Prefix
                Call txtDisabled_Color
            End If
    End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
        If I <> JobCDt Then
            Txt(I).Tag = ""
        End If
    Next I
    mCardNo = ""
    mLabDiscAmtTB = 0
    lblDocId.CAPTION = ""
    lblDocId.Refresh
    
    LblSprBill = ""
    lblLabourBill = ""
    lblGatePass = ""
    
End Sub

Private Sub MoveRec()
On Error GoTo errlbl
Dim Master1 As ADODB.Recordset ',rs As Recordset
'Dim mVor As String
'Dim i As Integer
mMRPReSales = 0
mMRPLubeTB = 0
mMRPLubeTP = 0
mAddFlag = "I"
mCardNo = ""
    If Master.RecordCount > 0 Then
        Set Master1 = GCn.Execute("select JC.Job_No,JC.Site_Code,JC.Govt_YN, JC.Job_Date, JC.JobCloseDate,jc.cardno,jc.OpenRemarks,jc.Body_Damage,jc.ObservBy_Eng,Jc.Job_BookNo,Jc.Job_InspDocID,Jc.AtKMsHrs,jc.Coupon,jc.Coupon_Value,jc.ArrivalTime,jc.ExpDelDate,jc.Est_SpCost,jc.Est_LabCost,Jc.DelBy,Jc.RecBy_Supervisor,Jc.DelayReason,Jc.JobComp_Dt_Time,JC.Remark,jc.NextSrvDate,EM.EMP_NAME AS Mechanic,EMP.Emp_Name as Supervisor,JD.R_Desc as ReasonName," _
            & "jc.CRMemo,jc.BillingName,JC.DrSpr_AcCode,jc.DrLab_AcCode,jc.labamt_tb,jc.labamt_tp,jc.lab_d_amt,jc.lab_taxper,jc.lab_taxamt,jc.lab_roundoff,jc.netlab_amt,JC.DocId_InvSpr,Jc.DocId_InvLab,JC.GP_No,HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name,HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName,jc.LabAmt_Out " _
            & "from (((((job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) left Join Service_Type as ST on JC.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode) left join Emp_Mast as EM on JC.Delby=EM.Emp_Code) left join Emp_Mast as EMP on Jc.RecBy_Supervisor=Emp.Emp_Code) Left Join Job_Delay as JD on JC.DelayReason=JD.Code where JC.DocId='" & Master!Code & "'")
        
        LblDiv.CAPTION = "Division : " & left(Master!Code, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        lblDocId.CAPTION = Master!Code
        
        SpareDocID = Master1!DocId_InvSpr
        LabourDocID = Master1!DocID_InvLab
        
        lblGatePass = Right(Master1!GP_No, 5)
        LblSprBill = DeCodeDocID(Master1!DocId_InvSpr, Document_No) ' 14, 8)
        lblLabourBill = DeCodeDocID(Master1!DocID_InvLab, Document_No) ', 14, 8)
        
        lblSparePrefix = DeCodeDocID(Master1!DocId_InvSpr, Document_Prefix) ', 9, 5)
        lblLabourPrefix = DeCodeDocID(Master1!DocID_InvLab, Document_Prefix) ', 9, 5)
        If Master1!Govt_YN = 0 Then   'Govt = No
            mFormCode = pubLocalTaxFormSpr
        Else
            mFormCode = pubGovtTaxFormSpr
        End If
        Txt(JobNo).Tag = Master!Code
        Txt(BodyDamage).TEXT = IIf(IsNull(Master1!body_damage), "", Master1!body_damage)
        Txt(JobNo).TEXT = Master1!Job_No
        Txt(JobDt).TEXT = Format(Master1!Job_Date, "dd/MMM/yyyy")

        Txt(JobCDt).TEXT = Format(Master1!JobCloseDate, "dd/MMM/yyyy")
        Txt(VehRegNo).TEXT = XNull(Master1!RegNo)
        Txt(Chassis).TEXT = XNull(Master1!Chassis)
        mCardNo = Master1!CardNo
        
        Txt(Model).TEXT = XNull(Master1!Model)
        Txt(Engine).TEXT = XNull(Master1!Engine)
        Txt(VehSrlNo).TEXT = XNull(Master1!VehSerialNo)
        Txt(SrvType).TEXT = XNull(Master1!Serv_Desc)
        Txt(OwnerName).TEXT = XNull(Master1!Name)
        Txt(Address1).TEXT = XNull(Master1!Add1)
        Txt(Address2).TEXT = XNull(Master1!Add2)
        Txt(Address3).TEXT = XNull(Master1!Add3)
        Txt(City).TEXT = XNull(Master1!CityName)
        Txt(PhoneOff).TEXT = XNull(Master1!PhoneOff)
        Txt(PhoneResi).TEXT = XNull(Master1!PhoneResi)
        Txt(Mobile).TEXT = XNull(Master1!Mobile)
        Txt(BookNo).TEXT = XNull(Master1!Job_BookNo)
        Txt(BookDt).TEXT = ""
        Txt(InspNo).TEXT = Trim(DeCodeDocID(XNull(Master1!Job_Inspdocid), Document_No))
        Txt(GovtYn).TEXT = IIf(Master1!Govt_YN = 0, "No ", "Yes")
        Txt(CurrentKMS).TEXT = IIf(IsNull(Master1!AtKMsHrs), "", Master1!AtKMsHrs)
        Txt(CouponNo).TEXT = IIf(IsNull(Master1!Coupon), "", Master1!Coupon)
        Txt(CouponVal).TEXT = IIf(IsNull(Master1!Coupon_Value), "", Master1!Coupon_Value)
        Txt(ArrTime).TEXT = Format(Master1!ArrivalTime, "hh:mm")
        Txt(DelDate).TEXT = Format(Master1!ExpDelDate, "dd/MMM/yyyy")
        Txt(DelTime).TEXT = Format(Master1!ExpDelDate, "hh:mm")
        Txt(EstSpare).TEXT = IIf(IsNull(Master1!Est_SpCost), "", Master1!Est_SpCost)
        Txt(EstLabour).TEXT = IIf(IsNull(Master1!Est_LabCost), "", Master1!Est_LabCost)
        Txt(OpenRemark).TEXT = XNull(Master1!OpenRemarks)
'special case
        Txt(MechName).TEXT = XNull(Master1!Mechanic)
        Txt(SuperName).TEXT = XNull(Master1!Supervisor)
        Txt(JobDelay).TEXT = XNull(Master1!ReasonName)
        Txt(JobCompDt).TEXT = Format(Master1!JobComp_Dt_Time, "dd/MMM/yyyy")
        Txt(JobCompTm) = Format(XNull(Master1!JobComp_Dt_Time), "hh:mm")
        Txt(CloseRemark).TEXT = XNull(Master1!Remark)
        Txt(NextSrv).TEXT = XNull(Master1!NextSrvDate)
        Txt(MechName).Tag = XNull(Master1!DelBy)
        Txt(SuperName).Tag = XNull(Master1!RecBy_Supervisor)
        Txt(JobDelay).Tag = XNull(Master1!DelayReason)
        '****
        Txt(OutSideLabAmt) = IIf(IsNull(Master1!LabAmt_Out), "", Format(Master1!LabAmt_Out, "0.00"))
        Txt(LabAmt) = Format(Master1!LabAmt_TB + Master1!LabAmt_TP - Master1!LabAmt_Out, "0.00")
        Txt(LabAmtTB) = Format(Master1!LabAmt_TB, "0.00")
        Txt(LabAmtTP) = Format(Master1!LabAmt_TP, "0.00")
        '****
        Txt(LabDisc) = Format(Master1!Lab_D_Amt, "0.00")
        Txt(ServTaxPer) = Format(Master1!Lab_TaxPer, "0.00")
        Txt(ServTaxAmt) = Format(Master1!Lab_TaxAmt, "0.00")
        Txt(LabROff) = Format(Master1!Lab_RoundOff, "0.00")
        Txt(NetLabAmt) = Format(Master1!NetLab_Amt, "0.00")
        
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select D_Per_TB,D_Per_TP,D_Amt_TB,D_Amt_TP,Addition,Packing,Gen_Sur_Per,Gen_Sur_Amt,Trans_Amt,Tax_Per,Tax_Amt, " _
                & "Tax_Sur_Per,Tax_Sur_Amt,TOT_Per,Tot_Amt,Rounded,ReSalTax_Per,ReSalTax_Amt,Total_Amt,OilAmt_MRP_TB,OilAmt_MRP_TP from Sp_Sale where DocId='" & Master1!DocId_InvSpr & "'", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            Txt(DiscPerTB) = IIf(IsNull(Rst!D_Per_TB), "", Format(Rst!D_Per_TB, "0.00"))
            Txt(DiscPerTP) = IIf(IsNull(Rst!D_Per_TP), "", Format(Rst!D_Per_TP, "0.00"))
            Txt(DiscAmtTB) = IIf(IsNull(Rst!D_Amt_TB), "", Format(Rst!D_Amt_TB, "0.00"))
            Txt(DiscAmtTP) = IIf(IsNull(Rst!D_Amt_TP), "", Format(Rst!D_Amt_TP, "0.00"))
            Txt(Addition) = IIf(IsNull(Rst!Addition), "", Format(Rst!Addition, "0.00"))
            Txt(PackCrg) = IIf(IsNull(Rst!Packing), "", Format(Rst!Packing, "0.00"))
            Txt(GenSurPer) = IIf(IsNull(Rst!Gen_Sur_Per), "", Format(Rst!Gen_Sur_Per, "0.00"))
            Txt(GenSurAmt) = IIf(IsNull(Rst!Gen_Sur_Amt), "", Format(Rst!Gen_Sur_Amt, "0.00"))
            Txt(TransAmt) = IIf(IsNull(Rst!Trans_Amt), "", Format(Rst!Trans_Amt, "0.00"))
            Txt(STaxPer) = IIf(IsNull(Rst!Tax_Per), "", Format(Rst!Tax_Per, "0.00"))
            Txt(STaxAmt) = IIf(IsNull(Rst!Tax_Amt), "", Format(Rst!Tax_Amt, "0.00"))
            Txt(TaxSurPer) = IIf(IsNull(Rst!Tax_Sur_Per), "", Format(Rst!Tax_Sur_Per, "0.00"))
            Txt(TaxSurAmt) = IIf(IsNull(Rst!Tax_Sur_Amt), "", Format(Rst!Tax_Sur_Amt, "0.00"))
            Txt(TurnOverPer) = IIf(IsNull(Rst!TOT_Per), "", Format(Rst!TOT_Per, "0.00"))
            Txt(TurnOverAmt) = IIf(IsNull(Rst!Tot_Amt), "", Format(Rst!Tot_Amt, "0.00"))
            Txt(SROff) = IIf(IsNull(Rst!Rounded), "", Format(Rst!Rounded, "0.00"))
            Txt(ReSalTaxPer) = IIf(IsNull(Rst!ReSalTax_Per), "", Format(Rst!ReSalTax_Per, "0.00"))
            Txt(ReSalTaxAmt) = IIf(IsNull(Rst!ReSalTax_Amt), "", Format(Rst!ReSalTax_Amt, "0.00"))
            Txt(NetSprAmt) = IIf(IsNull(Rst!Total_Amt), "", Format(Rst!Total_Amt, "0.00"))
            mMRPLubeTB = IIf(IsNull(Rst!OilAmt_MRP_TB), 0, Rst!OilAmt_MRP_TB)
            mMRPLubeTP = IIf(IsNull(Rst!OilAmt_MRP_TP), 0, Rst!OilAmt_MRP_TP)
        Else
            Txt(DiscPerTB) = ""
            Txt(DiscPerTP) = ""
            Txt(DiscAmtTB) = ""
            Txt(DiscAmtTP) = ""
            Txt(Addition) = ""
            Txt(PackCrg) = ""
            Txt(GenSurPer) = ""
            Txt(GenSurAmt) = ""
            Txt(TransAmt) = ""
            Txt(STaxPer) = ""
            Txt(STaxAmt) = ""
            Txt(TaxSurAmt) = ""
            Txt(TaxSurPer) = ""
            Txt(TurnOverPer) = ""
            Txt(TurnOverAmt) = ""
            Txt(SROff) = ""
            Txt(ReSalTaxPer) = ""
            Txt(ReSalTaxAmt) = ""
            Txt(NetSprAmt) = ""
        End If
        Set Rst = Nothing
        Call Fill_Grid

        Txt(CashBill) = IIf(Master1!CrMemo = 0, "Yes", "No")
        mVType = IIf(Master1!CrMemo = 0, "W_SIC", "W_SIR")
        
        Txt(CashParty) = IIf(IsNull(Master1!BillingName), "", Master1!BillingName)
        Txt(SpareParty).Tag = IIf(IsNull(Master1!DrSpr_AcCode), "", Master1!DrSpr_AcCode)
        Txt(LabourParty).Tag = IIf(IsNull(Master1!DrLab_AcCode), "", Master1!DrLab_AcCode)
        If Txt(CashBill) <> "Yes" Then
            If GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrSpr_AcCode & "'").RecordCount > 0 Then
                Txt(SpareParty).TEXT = GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrSpr_AcCode & "'").Fields(0).Value
            End If
            If GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrLab_AcCode & "'").RecordCount > 0 Then
                Txt(LabourParty).TEXT = GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrLab_AcCode & "'").Fields(0).Value
            End If
        Else
'            Txt(SpareParty).Tag = ""
'            Txt(LabourParty).Tag = ""
            Txt(SpareParty) = ""
            Txt(LabourParty) = ""
        End If
        Call veh_count
        Txt(DiscPerTB).Enabled = False
        Txt(DiscPerTP).Enabled = False
        Txt(DiscAmtTB).Enabled = False
        Txt(DiscAmtTP).Enabled = False
        Txt(Addition).Enabled = False
        Txt(PackCrg).Enabled = False
        Txt(GenSurPer).Enabled = False
        Txt(GenSurAmt).Enabled = False
        Txt(TransAmt).Enabled = False
        Txt(STaxPer).Enabled = False
        Txt(STaxAmt).Enabled = False
        Txt(TaxSurAmt).Enabled = False
        Txt(TaxSurPer).Enabled = False
    Else
        Call BlankText
    End If
    Grid_Hide
    
errlbl:
    Set Master1 = Nothing
    Set Rst = Nothing
    If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub Ini_Grid()
Dim MeLeft As Long, MeWidth As Long
MeLeft = Me.left
MeWidth = Me.width

    With FGrid
        .left = MeLeft ' + 45
        .width = MeWidth - 90
        .top = Txt(Model).top   '2685
        .height = 2745
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 22

        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450

        .TextMatrix(0, Col_PNo) = "Part No"
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 2430

        .TextMatrix(0, Col_PName) = "Part Name"
        .ColAlignment(Col_PName) = flexAlignLeftCenter
        .ColWidth(Col_PName) = 2500

        .TextMatrix(0, Col_ReqNoDocId) = "Req.Slip DocID"
        .ColAlignment(Col_ReqNoDocId) = flexAlignLeftCenter
        .ColWidth(Col_ReqNoDocId) = 0
        
        .TextMatrix(0, Col_ReqDate) = "Requi.Date"
        .ColAlignment(Col_ReqDate) = flexAlignLeftCenter
        .ColWidth(Col_ReqDate) = 1035
        
        .TextMatrix(0, Col_ReqNo) = "Requi.No."
        .ColAlignment(Col_ReqNo) = flexAlignLeftCenter
        .ColWidth(Col_ReqNo) = 780
        
        .TextMatrix(0, Col_Purpose) = "Purpose"
        .ColAlignment(Col_Purpose) = flexAlignLeftCenter
        .ColWidth(Col_Purpose) = 690

        .TextMatrix(0, Col_ReqSrNo) = "Req.Slip Serial No."
        .ColAlignment(Col_ReqSrNo) = flexAlignLeftCenter
        .ColWidth(Col_ReqSrNo) = 0

        .TextMatrix(0, Col_Unit) = "Unit"
        .ColAlignment(Col_Unit) = flexAlignLeftCenter
        .ColWidth(Col_Unit) = 435

        .TextMatrix(0, Col_MRP) = "MRP"
        .ColAlignment(Col_MRP) = flexAlignLeftCenter
        .ColWidth(Col_MRP) = 450

        .TextMatrix(0, Col_Taxable) = "Tax"
        .ColAlignment(Col_Taxable) = flexAlignLeftCenter
        .ColWidth(Col_Taxable) = 360

        .TextMatrix(0, Col_Qty) = "Qty"
        .ColAlignmentFixed(Col_Qty) = flexAlignRightCenter
        .ColWidth(Col_Qty) = 720

        .TextMatrix(0, Col_Rate) = "Rate"
        .ColAlignmentFixed(Col_Rate) = flexAlignRightCenter
        .ColWidth(Col_Rate) = 870

        .TextMatrix(0, Col_MRPRate) = "MRP Rate"
        .ColAlignmentFixed(Col_MRPRate) = flexAlignRightCenter
        .ColWidth(Col_MRPRate) = 0

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
        .ColWidth(Col_DiscPer) = 510

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
        .ColWidth(Col_DiscAmt) = 840

        .TextMatrix(0, Col_ItemVal) = "Item Value"
        .ColAlignmentFixed(Col_ItemVal) = flexAlignRightCenter
        .ColWidth(Col_ItemVal) = 1095

        .TextMatrix(0, Col_LName) = "Local Name"
        .ColAlignment(Col_LName) = flexAlignLeftCenter
        .ColWidth(Col_LName) = 2000
    
        .TextMatrix(0, Col_ClaimNo) = "Claim No."
        .ColAlignmentFixed(Col_ClaimNo) = flexAlignRightCenter
        .ColWidth(Col_ClaimNo) = 1095
        
        .TextMatrix(0, Col_CompYN) = "Stores"
        .ColAlignmentFixed(Col_CompYN) = flexAlignRightCenter
        .ColWidth(Col_CompYN) = 500
    End With
    
    With FGrid1
        .width = FGrid.width
        .left = FGrid.left
        .top = FGrid.top
        .height = FGrid.height
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 16
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        .TextMatrix(0, C_LabCode) = "Lab.Code"
        .ColAlignment(C_LabCode) = flexAlignLeftCenter
        .ColWidth(C_LabCode) = 780
        
        .TextMatrix(0, C_LabName) = "Labour Description"
        .ColAlignment(C_LabName) = flexAlignLeftCenter
        .ColWidth(C_LabName) = 4185

        .TextMatrix(0, C_TaxYN) = "Tax YN"
        .ColAlignment(C_TaxYN) = flexAlignCenterCenter
        .ColWidth(C_TaxYN) = 675
        
        .TextMatrix(0, C_PaidBy) = "PaidBy"
        .ColAlignment(C_PaidBy) = flexAlignCenterCenter
        .ColWidth(C_PaidBy) = 675
        
        .TextMatrix(0, C_ChrgType) = "Type"
        .ColAlignmentFixed(C_ChrgType) = flexAlignCenterCenter
        .ColAlignment(C_ChrgType) = flexAlignLeftCenter
        .ColWidth(C_ChrgType) = 885
        
        .TextMatrix(0, C_Hrs) = "Hrs."
        .ColAlignmentFixed(C_Hrs) = flexAlignCenterCenter
        .ColAlignment(C_Hrs) = flexAlignRightCenter
        .ColWidth(C_Hrs) = 600
        
        .TextMatrix(0, C_Rate) = "Rate" 'Ch.Amt." '
        .ColAlignmentFixed(C_Rate) = flexAlignCenterCenter
        .ColAlignment(C_Rate) = flexAlignRightCenter
        .ColWidth(C_Rate) = 600

        .TextMatrix(0, C_Amt) = "Amount"
        .ColAlignmentFixed(C_Amt) = flexAlignCenterCenter
        .ColAlignment(C_Amt) = flexAlignRightCenter
        .ColWidth(C_Amt) = 795

'        .TextMatrix(0, C_WarHrs) = "Wr.Hrs."
'        .ColAlignment(C_WarHrs) = flexAlignRightCenter
'        .ColWidth(C_WarHrs) = 600
'
'        .TextMatrix(0, C_WarAmt) = "Wr.Amt."
'        .ColAlignmentFixed(C_WarAmt) = flexAlignCenterCenter
'        .ColAlignment(C_WarAmt) = flexAlignRightCenter
'        .ColWidth(C_WarAmt) = 795

'        .TextMatrix(0, C_MechName) = "Mechanic Name"
'        .ColAlignment(C_MechName) = flexAlignLeftCenter
'        .ColWidth(C_MechName) = 2500

        .TextMatrix(0, C_External) = "Extl"
        .ColAlignment(C_External) = flexAlignLeftCenter
        .ColWidth(C_External) = 400

        .TextMatrix(0, C_GPNo) = "GP No."
        .ColAlignment(C_GPNo) = flexAlignLeftCenter
        .ColWidth(C_GPNo) = 800

        .TextMatrix(0, C_Remarks) = "Remarks"
        .ColAlignment(C_Remarks) = flexAlignLeftCenter
        .ColWidth(C_Remarks) = 1815

        .TextMatrix(0, C_ContName) = "Contractor Name"
        .ColAlignment(C_ContName) = flexAlignLeftCenter
        .ColWidth(C_ContName) = 2280

        .TextMatrix(0, C_WIssueDt) = "Issue Dt."
        .ColAlignment(C_WIssueDt) = flexAlignRightCenter
        .ColWidth(C_WIssueDt) = 1100

        .TextMatrix(0, C_WRecdDt) = "Recd Dt."
        .ColAlignment(C_WRecdDt) = flexAlignRightCenter
        .ColWidth(C_WRecdDt) = 1100

        .TextMatrix(0, C_ContAmt) = "Cont. Amt."
        .ColAlignment(C_ContAmt) = flexAlignRightCenter
        .ColWidth(C_ContAmt) = 950
    End With
    
    DGJob.width = MeWidth - 90: DGJob.left = MeLeft: DGJob.top = Line2.Y1: DGJob.height = 3000
    DGMech.width = 4740: DGMech.left = MeWidth - (DGMech.width + mRtScale): DGMech.top = mTopScale: DGMech.height = 5000
    DGReason.width = 4740: DGReason.left = MeWidth - (DGReason.width + mRtScale): DGReason.top = mTopScale: DGReason.height = 5000
    DGParty.width = 4740: DGParty.left = MeWidth - (DGParty.width + mRtScale): DGParty.top = mTopScale: DGParty.height = 5000
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
End Sub

Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
    If DGMech.Visible = True Then DGMech.Visible = False
    If DGReason.Visible = True Then DGReason.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
End Sub

Private Sub veh_count()
    If Txt(JobDt).TEXT <> "" Then
        LblTotVeh.CAPTION = GCn.Execute("select count(*) from job_Card where JobCloseDate=Null or isnull(JobCloseDate) and left(Docid,1)='" & PubDivCode & "'").Fields(0)
    End If
End Sub

Private Sub UpdRequery()
    RsJob.Requery
    RsMech.Requery
    RsSuper.Requery
    RsReason.Requery
    RsParty.Requery
End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub
Private Sub History_Field()
Dim rsForm As ADODB.Recordset, rsJob2 As ADODB.Recordset
    Set rsJob2 = GCn.Execute("select J.CardNo,J.Job_Date,J.Govt_YN,City.CityName,ST.Serv_Desc,HC.Add1,HC.Add2,HC.Add3,HC.PhoneOff,HC.PhoneResi,HC.Mobile, " _
        & "J.RecBy_Mechanic,J.RecBy_Supervisor,J.DelBy,J.Job_BookNo,J.Job_Inspdocid,J.ATKMSHRS,J.Coupon,J.ArrivalTime,J.ExpDelDate," _
        & "J.Est_SpCost,J.Est_LabCost,J.OpenRemarks,J.Body_Damage,E.Emp_Name as Mechanic,E1.Emp_Name as Supervisor,E2.Emp_Name as DelByMechanic " _
        & "from ((((((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) " _
        & "left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) " _
        & "Left Join City on HC.CityCode=City.CityCode) " _
        & "Left Join Emp_Mast as E on J.RecBy_Mechanic=E.Emp_Code) " _
        & "Left Join Emp_Mast as E1 on J.RecBy_Supervisor=E1.Emp_Code) " _
        & "Left Join Emp_Mast as E2 on J.DelBy=E2.Emp_Code) " _
        & "where J.DocId='" & RsJob!Code & "'")
    
    Txt(VehRegNo).Tag = RsJob!Code
    Txt(Chassis).Tag = RsJob!Code
    Txt(OwnerName).Tag = RsJob!Code
    Txt(JobNo).Tag = RsJob!Code
    Txt(JobNo).TEXT = IIf(IsNull(RsJob!Job_No), "", RsJob!Job_No)
    Txt(JobDt).TEXT = rsJob2!Job_Date
    Txt(VehRegNo).TEXT = IIf(IsNull(RsJob!RegNo), "", RsJob!RegNo)
    Txt(Chassis).TEXT = IIf(IsNull(RsJob!Chassis), "", RsJob!Chassis)
    Txt(Model).TEXT = IIf(IsNull(RsJob!Model), "", RsJob!Model)
    Txt(Engine).TEXT = IIf(IsNull(RsJob!Engine), "", RsJob!Engine)
    Txt(VehSrlNo).TEXT = IIf(IsNull(RsJob!VehSerialNo), "", RsJob!VehSerialNo)
    Txt(OwnerName).TEXT = IIf(IsNull(RsJob!Name), "", RsJob!Name)
    Txt(Address1).TEXT = IIf(IsNull(rsJob2!Add1), "", rsJob2!Add1)
    Txt(Address2).TEXT = IIf(IsNull(rsJob2!Add2), "", rsJob2!Add2)
    Txt(Address3).TEXT = IIf(IsNull(rsJob2!Add3), "", rsJob2!Add3)
    Txt(City).TEXT = IIf(IsNull(rsJob2!CityName), "", rsJob2!CityName)
    Txt(PhoneOff).TEXT = IIf(IsNull(rsJob2!PhoneOff), "", rsJob2!PhoneOff)
    Txt(PhoneResi).TEXT = IIf(IsNull(rsJob2!PhoneResi), "", rsJob2!PhoneResi)
    Txt(Mobile).TEXT = IIf(IsNull(rsJob2!Mobile), "", rsJob2!Mobile)
    Txt(MechName).Tag = rsJob2!RecBy_Mechanic
    Txt(MechName) = IIf(IsNull(rsJob2!Mechanic), "", rsJob2!Mechanic)
    Txt(SuperName).Tag = rsJob2!RecBy_Supervisor
    Txt(SuperName) = IIf(IsNull(rsJob2!Supervisor), "", rsJob2!Supervisor)
'special case
'    Txt(MechName).Tag = rsJob2!DelBy
'    Txt(MechName) = IIf(IsNull(rsJob2!DelByMechanic), "", rsJob2!DelByMechanic)
    Txt(SrvType).TEXT = IIf(IsNull(rsJob2!Serv_Desc), "", rsJob2!Serv_Desc)
    Txt(BookNo).TEXT = IIf(IsNull(rsJob2!Job_BookNo), "", rsJob2!Job_BookNo)
    Txt(BookDt).TEXT = ""
    Txt(InspNo).TEXT = Trim(DeCodeDocID(XNull(rsJob2!Job_Inspdocid), Document_No))
    Txt(GovtYn).TEXT = IIf(rsJob2!Govt_YN = 0, "No ", "Yes")
    Txt(CurrentKMS).TEXT = rsJob2!AtKMsHrs
    Txt(CouponNo).TEXT = rsJob2!Coupon
    Txt(ArrTime).TEXT = Format(rsJob2!ArrivalTime, "hh:mm")
    Txt(DelDate).TEXT = Format(rsJob2!ExpDelDate, "dd/MMM/yyyy")
    Txt(DelTime).TEXT = Format(rsJob2!ExpDelDate, "hh:mm")
    Txt(EstSpare).TEXT = rsJob2!Est_SpCost
    Txt(EstLabour).TEXT = rsJob2!Est_LabCost
    Txt(OpenRemark).TEXT = IIf(IsNull(rsJob2!OpenRemarks), "", rsJob2!OpenRemarks)
    Txt(BodyDamage).TEXT = rsJob2!body_damage
    
    Txt(GenSurPer) = Format(PubGenSurChrgOnSpr, "0.00")
    Txt(TurnOverPer) = Format(PubTOT_Rate, "0.00")
    
    Txt(CashBill) = "Yes"
    If rsJob2!Govt_YN = 0 Then   'Govt = No
        mFormCode = pubLocalTaxFormSpr
    Else
        mFormCode = pubGovtTaxFormSpr
    End If
    GSQL = "Select Tax_Per,Tax_Sur_Per FROM TaxForms where Form_Code='" & mFormCode & "'"
    
    Set rsForm = New ADODB.Recordset
    rsForm.CursorLocation = adUseClient
    rsForm.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If rsForm.RecordCount > 0 Then
        Txt(STaxPer).TEXT = IIf(IsNull(rsForm!Tax_Per), "", Format(rsForm!Tax_Per, "0.00"))
        Txt(TaxSurPer).TEXT = IIf(IsNull(rsForm!Tax_Sur_Per), "", Format(rsForm!Tax_Sur_Per, "0.00"))
    Else
        MsgBox "Please Add/Define Local/Govt Tax Form in " & vbCrLf & " Tax Forms/System Controls", vbOKOnly, "Control Validation"
    End If
    Set rsForm = Nothing
    Set rsJob2 = Nothing
End Sub

Private Sub txtDisabled_Color()
Dim I As Integer
    For I = 0 To Txt.Count - 1
        If Txt(I).Enabled = False Or Txt(I).Locked = True Then
            Txt(I).BackColor = &HEBF0F1
        Else
            Txt(I).BackColor = CtrlBColOrg
        End If
    Next I
End Sub

Private Sub Fill_Grid()
Dim I As Integer, TmpStr$
    If Txt(JobNo).Tag = "" Then Exit Sub
    
    '' Spares Details
    FGrid.Rows = 1
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    GSQL = "Select SPS.Job_DocId,SPS.DocId,SPS.V_Date,SPS.V_type, SPS.V_no, Sps.Srl_no, SPS.Part_No,SPS.Lub_Category,(SPS.Qty_iss-SPS.Qty_ret) as ReqQty, SPS.Tax_yn,SPS.Mrp_YN,SPS.Rate,SPS.MRP_Rate,SPS.Amount,SPS.Disc_per,SPS.Disc_Amt,SPS.Net_Amt,SPS.Purpose,SPS.TrnComplete_YN,SPS.Claim_No,Part.Part_Name,Part.Local_Name,part.Unit,Part.Part_Grade " & _
            "FROM SP_Stock AS SPS " & _
            "left join part on SpS.part_no=part.Part_No and Part.Div_Code = left(SPS.DocID,1) " & _
            "Where SPS.Job_DocId='" & Txt(JobNo).Tag & "' Order By SPS.V_No,SPS.Srl_No"
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    I = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, Col_SrNo) = I
                .TextMatrix(I, Col_PNo) = Rst!Part_No
                .TextMatrix(I, Col_ReqNoDocId) = XNull(Rst!DocId)
                .TextMatrix(I, Col_ReqNo) = Rst!V_NO
                .TextMatrix(I, Col_ReqDate) = Rst!V_DATE
                .TextMatrix(I, Col_ReqSrNo) = Rst!Srl_No
                .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, Col_Qty) = Format(Rst!ReqQty, "0.00")
                .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_RATE, "0.00")
                If Rst!Purpose = "P" Then
                    TmpStr = "PDI"
                ElseIf Rst!Purpose = "F" Then
                    TmpStr = "Free Service"
                ElseIf Rst!Purpose = "C" Then
                    TmpStr = "Charge"
                    .TextMatrix(I, Col_Amt) = Format(Rst!AMOUNT, "0.00")
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                    .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt, "0.00")
                ElseIf Rst!Purpose = "W" Then
                    TmpStr = "Warranty"
                ElseIf Rst!Purpose = "O" Then
                    TmpStr = "Company Vehicle"
                ElseIf Rst!Purpose = "L" Then
                    TmpStr = "Complementary"
                Else
                    TmpStr = ""
                End If
                .TextMatrix(I, Col_Purpose) = TmpStr
                .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                .TextMatrix(I, Col_ClaimNo) = IIf(IsNull(Rst!claim_no), "", Rst!claim_no)
                .TextMatrix(I, Col_CompYN) = IIf(Rst!TrnCompLete_YN = 1, "OK", "Pend")
                .TextMatrix(I, Col_PartGrade) = IIf(IsNull(Rst!Part_Grade), "", Rst!Part_Grade)
            End With
            I = I + 1
            Rst.MoveNext
        Loop
        If FGrid.Rows <= 1 Then FGrid.AddItem ""
        FGrid.FixedRows = 1
    Else
        FGrid.Rows = FGrid.Rows
        FGrid.AddItem ""
        FGrid.FixedRows = 1
    End If
    
    '' Labour Details
    Rst.Close
    FGrid1.Rows = 1
    'GSQL = "Select JL.*, Labour.Lab_Desc AS LabName,CF.FinName AS ContName " & _
            "From (Job_Lab as JL left join labour on JL.Lab_Code=Labour.Lab_Code) " & _
            "LEFT JOIN ContractFinance CF ON JL.ContractCode=CF.FinCode " & _
            "Where JL.Job_DocId='" & Txt(JobNo).Tag & "' order by JL.S_No"
            
    GSQL = "Select JL.*, L.Lab_Desc AS LabName,CF.FinName AS ContName,GP.GatePassDate,GP.ContractRecdDate,GP.ContractAmt,GP.ContractCode " & _
        " From ((((Job_Lab as JL left join labour as L on JL.Lab_Code=L.Lab_Code) " & _
        " Left Join Labour_Model LM on JL.Lab_Code=LM.Lab_Code) " & _
        " left join Job_GatePass as GP on JL.ExtJobGatePassNo=GP.GatePassNo) " & _
        " Left Join ContractFinance as CF ON GP.ContractCode=CF.FinCode) " & _
        " Where JL.Job_DocId='" & Txt(JobNo).Tag & "' Order by JL.S_No"
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    I = 1
    Txt(OutSideLabAmt) = ""
    Txt(LabAmt) = ""
    Txt(LabAmtTB) = ""
    Txt(LabAmtTP) = ""
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_LabCode) = Rst!Lab_Code
                .TextMatrix(I, C_LabName) = XNull(Rst!LabName)
                .TextMatrix(I, C_TaxYN) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, C_PaidBy) = XNull(Rst!CHRG_FROM)
'                If Rst!Hrs_Taken + Rst!Lab_Rate > 0 Then
                If Rst!CHRG_FROM = "M" Or Rst!CHRG_FROM = "O" Then
                    If Rst!Chrg_Type = "W" Then 'Warranty
                        .TextMatrix(I, C_ChrgType) = "Warranty"
                        .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_War = 0, "", Format(Rst!Hrs_War, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(Rst!War_Lab_Rate = 0, "", Format(Rst!War_Lab_Rate, "0.00"))
                        
                    Else    'Free Service
                        .TextMatrix(I, C_ChrgType) = "Free Service"
                        .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_Taken = 0, "", Format(Rst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(Rst!Lab_Rate = 0, "", Format(Rst!Lab_Rate, "0.00"))
                        
                    End If
                Else
                    .TextMatrix(I, C_ChrgType) = "Chargeable"
                    .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_Taken = 0, "", Format(Rst!Hrs_Taken, "0.00"))
                    .TextMatrix(I, C_Rate) = IIf(Rst!Lab_Rate = 0, "", Format(Rst!Lab_Rate, "0.00"))
                    .TextMatrix(I, C_Amt) = IIf(Rst!LabourAmt = 0, "", Format(Rst!LabourAmt, "0.00"))
                End If
                
                .TextMatrix(I, C_External) = IIf(Rst!External_yn = "1", "Yes", "No")
                .TextMatrix(I, C_GPNo) = XNull(Rst!ExtJobGatePassNo)
                .TextMatrix(I, C_ContName) = XNull(Rst!ContName)
                .TextMatrix(I, C_WIssueDt) = IIf(IsNull(Rst!GatePassDate), "", Rst!GatePassDate)
                .TextMatrix(I, C_WRecdDt) = IIf(IsNull(Rst!ContractRecdDate), "", Rst!ContractRecdDate)
                .TextMatrix(I, C_ContAmt) = IIf(Rst!ContractAmt = 0, "", Format(Rst!ContractAmt, "0.00"))
                .TextMatrix(I, C_Remarks) = XNull(Rst!Contract_Remarks)
'                .TextMatrix(i, C_ContCode) = XNull(Rst!ContractCode)
                If Rst!CHRG_FROM = "C" Then
                    If Rst!External_yn = "1" Then
                        Txt(OutSideLabAmt) = Format(Val(Txt(OutSideLabAmt)) + Rst!LabourAmt, "0.00")
                    Else
                        Txt(LabAmt) = Format(Val(Txt(LabAmt)) + Rst!LabourAmt, "0.00")
                    End If
                    If Rst!Tax_YN = 1 Then
                        Txt(LabAmtTB) = Format(Val(Txt(LabAmtTB)) + Rst!LabourAmt, "0.00")
                    Else
                        Txt(LabAmtTP) = Format(Val(Txt(LabAmtTP)) + Rst!LabourAmt, "0.00")
                    End If
                End If
            End With
            I = I + 1
            Rst.MoveNext
        Loop
    End If
    If FGrid1.Rows <= 1 Then FGrid1.AddItem ""
    FGrid1.FixedRows = 1
    Set Rst = Nothing
    MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(Txt(DiscPerTB)), Val(Txt(DiscPerTP)), _
            Val(Txt(STaxPer)), Val(Txt(TaxSurPer)), Val(Txt(TurnOverPer))
    MainLib.SprCalc WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
        Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
        Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
        Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
        Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
        Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
        Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
        Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
        Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_Purpose, True
    MainLib.LabCalc Txt(LabAmtTB), Txt(LabAmtTP), Txt(LabDisc), Txt(ServTaxPer), Txt(ServTaxAmt), Txt(LabROff), Txt(NetLabAmt), Txt(OutSideLabAmt), mLabDiscAmtTB
    Txt(NetAmt) = Format(Val(Txt(NetSprAmt)) + Val(Txt(NetLabAmt)), "0.00")
End Sub
Private Sub Fgrid_Ini()
    FGrid.Visible = False
    FGrid1.Visible = False
    
    lblSprGrid.CAPTION = "Show Spares"
    lblLabGrid.CAPTION = "Show Labour"
    
    lblSprGrid.Tag = 0
    lblLabGrid.Tag = 0
End Sub

Private Sub Generate_Prefix()
    If Txt(CashBill).TEXT = "Yes" Then
        Txt(CashParty).Enabled = True
        Txt(SpareParty).Enabled = False
        Txt(LabourParty).Enabled = False
        Txt(SpareParty).TEXT = ""
        Txt(LabourParty).TEXT = ""
        LabourVtype = "W_LIC"
        mVType = "W_SIC"
        Txt(SpareParty).Tag = PubSprCashAc
        Txt(LabourParty).Tag = PubSrvLabAc
    Else
        Txt(CashParty).Enabled = False
        Txt(CashParty).TEXT = ""
        Txt(CashParty).Tag = ""
        Txt(SpareParty).Enabled = True
        Txt(LabourParty).Enabled = True
        LabourVtype = "W_LIR"
        mVType = "W_SIR"
    End If
    SpareVtype = mVType
    SpareDocID = GetDocID(GCnFaS, SpareVtype, Txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
    LabourDocID = GetDocID(GCnFaW, SpareVtype, Txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
    LblSprBill.Refresh
    lblLabourBill.Refresh
End Sub

Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmPrn.Visible = False
        If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
            If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
            Call MoveRec
            Disp_Text SETS("INI", Me, Master)
        End If
    End If
End Sub

Private Sub Cmdprint_Click(Index As Integer)
'On Error GoTo ERRORHANDLER
If ChkRep(ChkSprInv).Value = 0 And ChkRep(ChkLabInv).Value = 0 Then
    MsgBox "Please Select Spare / Labour Invoice Option", vbCritical, "Validation": Exit Sub
End If
Dim mQryLab$, mSepLabInv As Byte
'PurposeStr = ""
Set GRs = GCn.Execute("Select PrintCompanyIssue,PrintComplIssue,SepLabourInv from Syctrl")
'If GRs!PrintCompanyIssue = 0 Then 'Not print
'    PurposeStr = "'O'"
'End If
'If GRs!PrintComplIssue = 0 Then 'Not print
'    PurposeStr = IIf(PurposeStr = "", "", PurposeStr & ",") & "'L'"
'End If
mSepLabInv = GRs!SepLabourInv
Set GRs = Nothing
If mSepLabInv = 0 Then  'No, Merge Invoice=Spare + Labour
    If Provisional Then
        Dim mFormCode$, mPrintDesc$
        If Txt(GovtYn) = "No" Then  'Govt = No
            mFormCode = pubLocalTaxFormSpr
        Else
            mFormCode = pubGovtTaxFormSpr
        End If
        mPrintDesc = GCn.Execute("select Printing_desc from TaxForms where Form_Code='" & mFormCode & "'").Fields(0).Value
        GSQL = "SELECT '1' as Orig,JC.AtKMsHrs,JC.Lab_D_Amt,SPStk.DocID as ReqDocID,iif(isnull(SPStk.Srl_No),0,SPStk.Srl_No) as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab, #" & date & "#  as v_Date,'' AS Party_Code,' " & Txt(OwnerName) & "' as Party_Name,'' as Address,'' as NamePrefix,' " & Txt(OwnerName) & "' as Name,' " & Txt(Address1) & "' as Add1,' " & Txt(Address2) & "' as Add2,' " & Txt(Address3) & "' as Add3,' " & Txt(City) & "' as CityName,'' as PIN,' " & Txt(PhoneResi) & "' as Phone," & _
            "'' as CSTNo,'' as L_C,'" & mFormCode & "' as Form_Code,'" & mPrintDesc & "' as Printing_Desc,'' as Remarks, " & Val(Txt(MRPAmtTB)) & " as SprAmt_MRP_TB, " & Val(Txt(MRPAmtTP)) & " as SprAmt_MRP_TP," & mMRPLubeTB & " as OilAmt_MRP_TB," & mMRPLubeTP & " as OilAmt_MRP_TP, " & Val(Txt(SprAmtTB)) & " as SprAmt_TB, " & Val(Txt(SprAmtTP)) & " as SprAmt_TP, " & Val(Txt(OilAmtTB)) & " as OilAmt_TB,  " & Val(Txt(OilAmtTP)) & " as OilAmt_TP, " & Val(Txt(DiscPerTB)) & " as D_Per_TB,  " & Val(Txt(DiscAmtTB)) & " as D_Amt_TB,  " & Val(Txt(DiscPerTP)) & " as D_Per_TP, " & Val(Txt(DiscAmtTP)) & " as D_Amt_TP,0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP,0 as D_Amt_MRP_TP, " & Val(Txt(Addition)) & " as Addition, " & Val(Txt(GenSurPer)) & " as Gen_Sur_Per, " & Val(Txt(GenSurAmt)) & " as Gen_Sur_Amt, " & Val(Txt(TransAmt)) & " as Trans_Amt, " & Val(Txt(STaxPer)) & " as Tax_Per,  " & Val(Txt(STaxAmt)) & " as Tax_Amt, 0 as Tax_AmtMRP,  " & Val(Txt(TaxSurPer)) & " as Tax_Sur_Per, " & Val(Txt(TaxSurAmt)) & " as Tax_Sur_Amt,0 as TaxSur_AmtMRP, " & Val(Txt(PackCrg)) & " as Packing,  " & Val(Txt(TurnOverPer)) & " as TOT_Per,  " & Val(Txt(TurnOverAmt)) & " as Tot_Amt, 0 as TOT_AmtMRP,0 as ReSalTax_Per,0 as ReSalTax_Amt, " & Val(Txt(STotB)) & " as Total_Amt," & _
            " " & Val(Txt(SROff)) & " as Rounded, " & PubTaxDetOnSprInv & " as Det_Tax,0 as GP_No,'' as GP_Date,1 as Printed_YN, ' " & pubUName & "' as U_Name, ' " & date$ & "' as U_EntDt,0 as CancelYN,0 as LabAmt_TB, 0 as LabAmt_TP, 0 as Lab_TaxPer, 0 as Lab_TaxAmt, 0 as Lab_D_Amt,0 as Lab_RoundOff,0 as NetLab_Amt," & _
            "SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown,iif(isnull(SPStk.Qty_Doc),0,SPStk.Qty_Doc) as Qty_Doc,iif(isnull(SPStk.Qty_Rec),0,SPStk.Qty_Rec) as Qty_Rec," & _
            "iif(isnull(SPStk.Qty_Iss),0,SPStk.Qty_Iss) as Qty_Iss,iif(isnull(SPStk.Qty_Ret),0,SPStk.Qty_Ret) as Qty_Ret,iif(isnull(SPStk.Tax_YN),0,SPStk.Tax_YN) as Tax_YN,iif(isnull(SPStk.MRP_YN),0,SPStk.MRP_YN) as MRP_YN,iif(isnull(SPStk.Rate),0,SPStk.Rate) as Rate," & _
            "iif(isnull(SPStk.MRP_Rate),0,SPStk.MRP_Rate) as MRP_Rate,iif(isnull(SPStk.Purpose),'',SPStk.Purpose) as Purpose,SPStk.Part_SrlNo,iif(isnull(SPStk.Rate2),0,SPStk.Rate) as Rate2,iif(isnull(SPStk.MRP_Rate),0,SPStk.MRP_Rate) as MRP_Rate2," & _
            "iif(isnull(SPStk.Disc_Per2),0,SPStk.Disc_Per2) as Disc_Per2,iif(isnull(SPStk.Disc_Amt2),0,SPStk.Disc_Amt2) as Disc_Amt2,iif(isnull(SPStk.Amount),0,SPStk.Amount) as Amount2,iif(isnull(SPStk.Net_Amt),0,SPStk.Net_Amt) as Net_Amt2,'' as Chrg_From,'' as External_YN, " & _
            "Syctrl.WorkShopInvFooter,iif(isnull(Syctrl.SrvGatePass_On),0,Syctrl.SrvGatePass_On) as SrvGatePass_On,iif(isnull(Syctrl.SrvGatePass),0,Syctrl.SrvGatePass) as SrvGatePass " & _
        "FROM ((SP_Stock as SPStk left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
            "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON SPStk.Job_DocID = JC.DocId) " & _
            "LEFT JOIN Syctrl ON Syctrl.LinkTable<>SPStk.U_AE " & _
            "where SPStk.Job_DocId='" & lblDocId & "'"  'modi lps  and (SPStk.Qty_Iss -SPStk.Qty_Ret)>0 "
    'Modi LPS at Cuttack 31.08.03
    '    If PurposeStr <> "" Then
    '        GSQL = GSQL & " and SPStk.Purpose not in (" & PurposeStr & ")"
    '    End If
        'GSQL = GSQL & "Order by SpStk.Docid,SpStk.Srl_No"
        mQryLab = "SELECT '2' as Orig,0 as AtKMsHrs,0 as Lab_D_Amt,'                     ' as ReqDocID,JL.S_No as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab,#" & date & "#  as v_Date,JC.DrLab_AcCode as Party_Code,JC.BillingName as Party_Name,'' as Address,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
            "SG.CSTNo,'' as L_C,'' as Form_Code,'' as Printing_Desc,'' as Remarks,0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
            "0 as Rounded,0 as Det_Tax,JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted as Printed_YN,JC.U_Name,JC.U_EntDt,0 as CancelYN,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
            "JL.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name,'' as Lub_Category, '' as Godown,0 as Qty_Doc, 0 as Qty_Rec, " & _
            "iif(isnull(Hrs_Taken),0,Hrs_Taken) as Qty_Iss,0 as Qty_Ret,iif(isnull(JL.Tax_YN),0,JL.Tax_YN) as Tax_YN, 0 as MRP_YN,0 as Rate," & _
            "0 as MRP_Rate,'' as Purpose,'' as Part_SrlNo,iif(isnull(JL.Lab_Rate),0,JL.Lab_Rate) as Rate2,0 as MRP_Rate2," & _
            "0 as Disc_Per2,0 as Disc_Amt2,0 as Amount2,iif(JL.Chrg_From = 'C',JL.LabourAmt,0) as Net_Amt2,JL.Chrg_From,JL.External_YN," & _
            "Syctrl.WorkShopInvFooter,iif(isnull(Syctrl.SrvGatePass_On),0,Syctrl.SrvGatePass_On) as SrvGatePass_On," & _
            "iif(isnull(Syctrl.SrvGatePass),0,Syctrl.SrvGatePass) as SrvGatePass " & _
        "FROM ((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
            "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
            "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
            "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
            "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
            "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code) " & _
        "Where JC.DocId='" & lblDocId & "'" ' Order By JL.JobDocID, JL.S_No"
     
        GSQL = GSQL & " Union All " & mQryLab & " Order By 1,2,3"
    Else

        GSQL = "SELECT '1' as Orig,JC.AtKMsHrs,JC.Lab_D_Amt,SPStk.DocID as ReqDocID,iif(isnull(SPStk.Srl_No),0,SPStk.Srl_No) as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvSpr,JC.DocID_InvLab,JC.JobCloseDate as v_Date,s.Party_Code,s.Party_Name,s.Address,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
            "SG.CSTNo,s.L_C,s.Form_Code,TF.Printing_Desc,s.Remarks,s.SprAmt_MRP_TB,s.SprAmt_MRP_TP,s.OilAmt_MRP_TB,s.OilAmt_MRP_TP,s.SprAmt_TB,s.SprAmt_TP,s.OilAmt_TB, s.OilAmt_TP,s.D_Per_TB, s.D_Amt_TB, s.D_Per_TP,s.D_Amt_TP,s.D_Per_MRP_TB,s.D_Amt_MRP_TB," & _
            "s.D_Per_MRP_TP,s.D_Amt_MRP_TP,s.Addition,s.Gen_Sur_Per,s.Gen_Sur_Amt,s.Trans_Amt,s.Tax_Per, s.Tax_Amt, s.Tax_AmtMRP, s.Tax_Sur_Per,s.Tax_Sur_Amt,s.TaxSur_AmtMRP,s.Packing, s.TOT_Per, s.Tot_Amt, s.TOT_AmtMRP,s.ReSalTax_Per,s.ReSalTax_Amt,s.Total_Amt," & _
            "s.Rounded,s.Det_Tax,s.GP_No,s.GP_Date,s.Printed_YN,s.U_Name, s.U_EntDt,S.CancelYN,0 as LabAmt_TB, 0 as LabAmt_TP, 0 as Lab_TaxPer, 0 as Lab_TaxAmt, 0 as Lab_D_Amt,0 as Lab_RoundOff,0 as NetLab_Amt," & _
            "SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown,iif(isnull(SPStk.Qty_Doc),0,SPStk.Qty_Doc) as Qty_Doc,iif(isnull(SPStk.Qty_Rec),0,SPStk.Qty_Rec) as Qty_Rec," & _
            "iif(isnull(SPStk.Qty_Iss),0,SPStk.Qty_Iss) as Qty_Iss,iif(isnull(SPStk.Qty_Ret),0,SPStk.Qty_Ret) as Qty_Ret,iif(isnull(SPStk.Tax_YN),0,SPStk.Tax_YN) as Tax_YN,iif(isnull(SPStk.MRP_YN),0,SPStk.MRP_YN) as MRP_YN,iif(isnull(SPStk.Rate),0,SPStk.Rate) as Rate," & _
            "iif(isnull(SPStk.MRP_Rate),0,SPStk.MRP_Rate) as MRP_Rate,iif(isnull(SPStk.Purpose),'',SPStk.Purpose) as Purpose,SPStk.Part_SrlNo,iif(isnull(SPStk.Rate2),0,SPStk.Rate2) as Rate2,iif(isnull(SPStk.MRP_Rate2),0,SPStk.MRP_Rate2) as MRP_Rate2," & _
            "iif(isnull(SPStk.Disc_Per2),0,SPStk.Disc_Per2) as Disc_Per2,iif(isnull(SPStk.Disc_Amt2),0,SPStk.Disc_Amt2) as Disc_Amt2,iif(isnull(SPStk.Amount2),0,SPStk.Amount2) as Amount2,iif(isnull(SPStk.Net_Amt2),0,SPStk.Net_Amt2) as Net_Amt2,'' as Chrg_From,'' as External_YN, " & _
            "Syctrl.WorkShopInvFooter,iif(isnull(Syctrl.SrvGatePass_On),0,Syctrl.SrvGatePass_On) as SrvGatePass_On,iif(isnull(Syctrl.SrvGatePass),0,Syctrl.SrvGatePass) as SrvGatePass " & _
        "FROM (((((SP_Sale as S left JOIN SP_Stock as SPStk ON S.DocID = SPStk.Invoice_DocId) " & _
            "left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
            "LEFT JOIN (SubGroup as SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON S.Party_Code = SG.SubCode) " & _
            "Left Join TaxForms TF on S.Form_Code=TF.Form_Code) " & _
            "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON S.Job_DocID = JC.DocId) " & _
            "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
            "where S.Job_DocId='" & Master!Code & "'" 'modi lps  and (SPStk.Qty_Iss -SPStk.Qty_Ret)>0 "
    'Modi LPS at Cuttack 31.08.03
    '    If PurposeStr <> "" Then
    '        GSQL = GSQL & " and SPStk.Purpose not in (" & PurposeStr & ")"
    '    End If
        'GSQL = GSQL & "Order by SpStk.Docid,SpStk.Srl_No"
        mQryLab = "SELECT '2' as Orig,0 as AtKMsHrs,0 as Lab_D_Amt,JL.Job_DocID as ReqDocID,JL.S_No as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvSpr,JC.DocID_InvLab,JC.JobCloseDate as v_Date,JC.DrLab_AcCode as Party_Code,JC.BillingName as Party_Name,'' as Address,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
            "SG.CSTNo,'' as L_C,'' as Form_Code,'' as Printing_Desc,'' as Remarks,0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
            "0 as Rounded,0 as Det_Tax,JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted as Printed_YN,JC.U_Name,JC.U_EntDt,0 as CancelYN,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
            "JL.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name,'' as Lub_Category, '' as Godown,0 as Qty_Doc, 0 as Qty_Rec, " & _
            "iif(isnull(Hrs_Taken),0,Hrs_Taken) as Qty_Iss,0 as Qty_Ret,iif(isnull(JL.Tax_YN),0,JL.Tax_YN) as Tax_YN, 0 as MRP_YN,0 as Rate," & _
            "0 as MRP_Rate,'' as Purpose,'' as Part_SrlNo,iif(isnull(JL.Lab_Rate),0,JL.Lab_Rate) as Rate2,0 as MRP_Rate2," & _
            "0 as Disc_Per2,0 as Disc_Amt2,0 as Amount2,iif(JL.Chrg_From = 'C',JL.LabourAmt,0) as Net_Amt2,JL.Chrg_From,JL.External_YN," & _
            "Syctrl.WorkShopInvFooter,iif(isnull(Syctrl.SrvGatePass_On),0,Syctrl.SrvGatePass_On) as SrvGatePass_On," & _
            "iif(isnull(Syctrl.SrvGatePass),0,Syctrl.SrvGatePass) as SrvGatePass " & _
        "FROM ((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
            "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
            "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
            "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
            "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
            "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code) " & _
        "Where JC.DocId='" & Master!Code & "'" ' Order By JL.JobDocID, JL.S_No"
     
        GSQL = GSQL & " Union All " & mQryLab & " order by 1,2,3,Part_No"
    End If
Else
    GSQL = "SELECT JC.AtKMsHrs,JC.DocID as JobDocID,JC.CrMemo,JC.Lab_D_Amt,H.Model,H.RegNo,H.Chassis,JC.DocId_InvSpr,JC.JobCloseDate as v_Date," & _
        "s.Party_Code,s.Party_Name,s.Address,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
        "SG.CSTNo,s.L_C,s.Form_Code,TF.Printing_Desc,s.Remarks," & _
        "s.SprAmt_MRP_TB,s.SprAmt_MRP_TP,s.OilAmt_MRP_TB,s.OilAmt_MRP_TP,s.SprAmt_TB,s.SprAmt_TP,s.OilAmt_TB, s.OilAmt_TP," & _
        "s.D_Per_TB, s.D_Amt_TB, s.D_Per_TP,s.D_Amt_TP,s.D_Per_MRP_TB,s.D_Amt_MRP_TB,s.D_Per_MRP_TP,s.D_Amt_MRP_TP,s.Addition,s.Gen_Sur_Per,s.Gen_Sur_Amt,s.Trans_Amt," & _
        "s.Tax_Per, s.Tax_Amt, s.Tax_AmtMRP, s.Tax_Sur_Per,s.Tax_Sur_Amt,s.TaxSur_AmtMRP,s.Packing, s.TOT_Per, s.Tot_Amt, s.TOT_AmtMRP,s.ReSalTax_Per,s.ReSalTax_Amt,s.Total_Amt,s.Rounded," & _
        "s.Det_Tax,s.GP_No,s.GP_Date,s.Printed_YN,s.U_Name, s.U_EntDt,S.CancelYN," & _
        "0 as LabAmt_TB, 0 as LabAmt_TP, 0 as Lab_TaxPer, 0 as Lab_TaxAmt, 0 as Lab_D_Amt,0 as Lab_RoundOff,0 as NetLab_Amt, " & _
        "iif(isnull(SPStk.Srl_No),0,SPStk.Srl_No) as Srl_No,SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown," & _
        "iif(isnull(SPStk.Qty_Doc),0,SPStk.Qty_Doc) as Qty_Doc,iif(isnull(SPStk.Qty_Rec),0,SPStk.Qty_Rec) as Qty_Rec,iif(isnull(SPStk.Qty_Iss),0,SPStk.Qty_Iss) as Qty_Iss," & _
        "iif(isnull(SPStk.Qty_Ret),0,SPStk.Qty_Ret) as Qty_Ret,iif(isnull(SPStk.Tax_YN),0,SPStk.Tax_YN) as Tax_YN,iif(isnull(SPStk.MRP_YN),0,SPStk.MRP_YN) as MRP_YN," & _
        "iif(isnull(SPStk.Rate),0,SPStk.Rate) as Rate,iif(isnull(SPStk.MRP_Rate),0,SPStk.MRP_Rate) as MRP_Rate," & _
        "iif(isnull(SPStk.Purpose),'',SPStk.Purpose) as Purpose,SPStk.Part_SrlNo,iif(isnull(SPStk.Rate2),0,SPStk.Rate2) as Rate2,iif(isnull(SPStk.MRP_Rate2),0,SPStk.MRP_Rate2) as MRP_Rate2," & _
        "iif(isnull(SPStk.Disc_Per2),0,SPStk.Disc_Per2) as Disc_Per2,iif(isnull(SPStk.Disc_Amt2),0,SPStk.Disc_Amt2) as Disc_Amt2,iif(isnull(SPStk.Amount2),0,SPStk.Amount2) as Amount2," & _
        "iif(isnull(SPStk.Net_Amt2),0,SPStk.Net_Amt2) as Net_Amt2," & _
        "Syctrl.WorkShopInvFooter,iif(isnull(Syctrl.SrvGatePass_On),0,Syctrl.SrvGatePass_On) as SrvGatePass_On," & _
        "iif(isnull(Syctrl.SrvGatePass),0,Syctrl.SrvGatePass) as SrvGatePass " & _
    "FROM (((((SP_Sale as S left JOIN SP_Stock as SPStk ON S.DocID = SPStk.Invoice_DocId) " & _
        "left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
        "LEFT JOIN (SubGroup as SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON S.Party_Code = SG.SubCode) " & _
        "Left Join TaxForms TF on S.Form_Code=TF.Form_Code) " & _
        "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON S.Job_DocID = JC.DocId) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
    "where S.Job_DocId='" & Master!Code & "' and (SPStk.Qty_Iss -SPStk.Qty_Ret)>0 "
'Modi LPS at Cuttack 31.08.03
'    If PurposeStr <> "" Then
'        GSQL = GSQL & " and SPStk.Purpose not in (" & PurposeStr & ")"
'    End If
    GSQL = GSQL & "Order By SPStk.Part_No"
    
    'Labour SQL
    mQryLab = "SELECT 0 as AtKMsHrs,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvLab,JC.JobCloseDate as v_Date," & _
        "JC.DrLab_AcCode as Party_Code,JC.BillingName as Party_Name,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
        "JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted,JC.U_Name,JC.U_EntDt," & _
        "JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
        "JL.S_No,JL.Lab_Code,Labour.Lab_Desc as LabName,JL.Tax_YN,JL.Hrs_Taken,JL.Hrs_War,JL.Lab_Rate," & _
        "JL.LabourAmt,JL.Chrg_From,JL.External_YN," & _
        "Syctrl.LabInvFooter,iif(isnull(Syctrl.SrvGatePass_On),0,Syctrl.SrvGatePass_On) as SrvGatePass_On," & _
        "iif(isnull(Syctrl.SrvGatePass),0,Syctrl.SrvGatePass) as SrvGatePass " & _
    "FROM ((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
        "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
        "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
        "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
        "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
        "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code) " & _
    "Where JC.DocId='" & Master!Code & _
    "' Order By JL.Lab_Code"
End If

Select Case Index
    Case PScreen, PWindows
        If OptPlain.Value = True Then
            If mSepLabInv = 0 Then ' Merge Invoice
                mRepName = "WorkShopBill"
            Else    'Separate Invoice for Spare & Labour
                If ChkRep(ChkSprInv).Value = 1 Then
                    mRepName = "SprJobBill"
                End If
                If ChkRep(ChkLabInv).Value = 1 Then
                    mRepName1 = "LabBill"
                End If
            End If
        Else
            If mSepLabInv = 0 Then ' Merge Invoice
                mRepName = "WorkShopBill"
            Else    'Separate Invoice for Spare & Labour
                If ChkRep(ChkSprInv).Value = 1 Then
                    mRepName = "SprJobBill"
                End If
                If ChkRep(ChkLabInv).Value = 1 Then
                    mRepName1 = "LabBill"
                End If
            End If
        End If
        If mSepLabInv = 0 Then
'            Call WindowsPrintBoth(Index, GSQL, mQryLab)
        Else
            If ChkRep(ChkSprInv).Value = 1 Then Call WindowsPrintSpr(Index, GSQL)
            If ChkRep(ChkLabInv).Value = 1 Then Call WindowsPrintLab(Index, mQryLab)
        End If
        FrmPrn.Visible = False
    Case PDos
        If mSepLabInv = 0 Then
            Call SpeedPrintBoth(GSQL, Optpre.Value)
        Else
            If ChkRep(ChkSprInv).Value = 1 Then Call SpeedPrintSpr(GSQL, Optpre.Value)
            If ChkRep(ChkLabInv).Value = 1 Then Call SpeedPrintLab(mQryLab)
        End If
        FrmPrn.Visible = False
    Case PSetUp
'        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Not Provisional Then
    If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
        If MsgBox("Print Service Letter? ", vbYesNo + vbCritical + vbDefaultButton2, "Service Letter Printing") = vbYes Then
            SpeedPrintSrvLet
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
        Call MoveRec
        Disp_Text SETS("INI", Me, Master)
    End If
Else
    Provisional = False
End If
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

Private Sub SpeedPrintSpr(mQRY$, PrePrinted)
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
    Dim I As Integer, J As Integer
    Dim PrintStr As String
    Dim rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim SrvGatePassOn$, Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double
    Dim MRPTaxStr$, mTPAmtStr$, mTBAmtStr$
    
    Set RstJob = GCn.Execute(mQRY)
    If RstJob.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
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
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 -> X * 0.56
    'chr 10 to chr 17 -> X * 1.7
        
    mHeader = 0   'Ideal 17
    mFooter = 19    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
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
    'eof modi
      
    'Sale Bill Header
    If RstJob!CrMemo = 0 Then
        mDocStr = "CASH MEMO"
    Else
        mDocStr = "INVOICE"
    End If
    mDupStr = IIf(RstJob!Printed_YN = 1, "(DUPLICATE)", "")
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
    mTaxdesc = GCn.Execute("select Printing_Desc from TaxForms where Form_Code = '" & RstJob!Form_Code & "'").Fields(0).Value
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    If PrePrinted Then
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        mHeader = 8
    Else
        Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCompDet!W_SecSpeciality) <> "" Then
            Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
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
        Print #1, PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
    End If
        Print #1, PRN_TIT("** WORKSHOP SPARE " & mDocStr & mDupStr & " **", "B", PageWidth)
        mHeader = mHeader + 1
        Print #1, mChr18 & Space(48) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
        mHeader = mHeader + 1
        Print #1, PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
        & Space(4) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
        mHeader = mHeader + 1
        
        Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
        mHeader = mHeader + 1
        If RstJob!Det_Tax = 1 Then
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 35) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----->" & "<---------AMOUNT--------->"
            mHeader = mHeader + 1
            Print #1, Space(89) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
            mHeader = mHeader + 1
        Else
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 28) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
            mHeader = mHeader + 1
        End If
        Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
        mHeader = mHeader + 1
        mFix = PageLength - (mHeader + mFooter)
        Page = 1
        mLine = 1
        mSlNo = 1
        LAdd = VNull(RstJob!Gen_Sur_Amt) + VNull(RstJob!Trans_Amt) + VNull(RstJob!Tax_Amt) + VNull(RstJob!Tax_Sur_Amt) + VNull(RstJob!Packing) + VNull(RstJob!ReSalTax_Amt) + VNull(RstJob!Tot_Amt)
        SubTot = RstJob!SprAmt_TB + RstJob!SprAmt_TP + RstJob!SprAmt_MRP_TB + RstJob!SprAmt_MRP_TP _
        + RstJob!OilAmt_TB + RstJob!OilAmt_TP + Val(Txt(IWDiscTotTP).TEXT) + Val(Txt(IWDiscTotTB).TEXT)
        If RstJob.RecordCount > 0 Then
            I = 1
            Do Until RstJob.EOF
                If mLine > mFix Then
                    Page = Page + 1
                    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-")
                    Print #1, Space(PageWidth - Len("Contd. on next page.." + str(Page))) & "Contd. on next page.." & str(Page)
                    Do Until mLine >= mFix + mFooter - 2
                        Print #1, ""
                        mLine = mLine + 1
                    Loop
                    Print #1, mEject
                    
                    'Header On Second Page
                    mHeader = 0
                    If PrePrinted Then
                        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                        mHeader = 8
                    Else
                        Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                        mHeader = mHeader + 1
                        If XNull(RstCompDet!W_SecSpeciality) <> "" Then
                            Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
                            mHeader = mHeader + 1
                        End If
                    End If
                     
                    Print #1, PRN_TIT("** WORKSHOP SPARE " & mDocStr & mDupStr & " **", "B", PageWidth)
                    mHeader = mHeader + 1
                    mHeader = mHeader + 1
                    Print #1, mChr18 & Space(48) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                    mHeader = mHeader + 1
                    Print #1, PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
                    mHeader = mHeader + 1
                    Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
                    mHeader = mHeader + 1
'                    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
'                    & Space(4) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
'                    mHeader = mHeader + 1
                   
                    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
                    mHeader = mHeader + 1
                    If RstJob!Det_Tax = 1 Then
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 35) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----->" & "<---------AMOUNT--------->"
                        mHeader = mHeader + 1
                        Print #1, Space(89) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
                        mHeader = mHeader + 1
                    Else
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 28) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
                        mHeader = mHeader + 1
                    End If
                    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
                    mHeader = mHeader + 1
                    mFix = PageLength - (mHeader + mFooter)
                    mLine = 1
                End If
                mRate = IIf(RstJob!MRP_YN = 1, RstJob!MRP_Rate2, RstJob!Rate2)
                If RstJob!Det_Tax = 1 Then
                    mTPAmtStr = PSTR(0, 12, 2)
                    mTBAmtStr = PSTR(0, 12, 2)
                    If RstJob!Purpose = "W" Then
                        mTBAmtStr = "*Warranty*"
                    ElseIf RstJob!Purpose = "P" Then
                        mTBAmtStr = "*PDI*"
                    ElseIf RstJob!Purpose = "F" Then
                        mTBAmtStr = "*Free*"
                    ElseIf RstJob!Purpose = "L" Then
                        mTBAmtStr = "*Compliment*"
                    ElseIf RstJob!Purpose = "O" Then
                        mTBAmtStr = "*Company*"
                    Else
                        If RstJob!Tax_YN = 0 Then
                            mTPAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                            mTBAmtStr = PSTR(0, 12, 2)
                        Else
                            mTPAmtStr = PSTR(0, 12, 2)
                            mTBAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                        End If
                    End If
                    'modishekhar
                    If RstJob!Purpose = "W" And GCn.Execute("select PrnWarrSpr from syctrl").Fields(0).Value = 0 Then GoTo NXT
                    'modi lps at Cuttack 31.08.03
                    If RstJob!Purpose = "L" And GCn.Execute("select PrintComplIssue from syctrl").Fields(0).Value = 0 Then GoTo NXT
                    If RstJob!Purpose = "O" And GCn.Execute("select PrintCompanyIssue from syctrl").Fields(0).Value = 0 Then GoTo NXT

                    PrintStr = PSTR(Trim(str(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 35) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                        PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
  
'                       IIf(RstJob!Purpose = "W", "*Warranty*", & _
                        IIf(RstJob!Tax_YN = 0, PSTR(RstJob!Net_Amt2, 12, 2) & PSTR(0, 12, 2), PSTR(0, 12, 2) & PSTR(RstJob!Net_Amt2, 12, 2)))
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
                    mAmount = Round(RstJob!Qty_Iss * RstJob!Rate, 2) - RstJob!Disc_Amt
                    PrintStr = PSTR(Trim(str(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 28, , AlignLeft) & PSTR(RstJob!Part_Name, 40) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(LdRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", "L") & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & _
                    PSTR(LDAmt - RstJob!Disc_Amt2, 12, 2)
                End If
                Print #1, PrintStr
                'modi lps at Cuttack 31.08.03
                mSlNo = mSlNo + 1
                mLine = mLine + 1
NXT:
                RstJob.MoveNext
'                mSlNo = mSlNo + 1
'                mLine = mLine + 1
            Loop
        End If
        Do Until mLine >= mFix
            Print #1, ""
            mLine = mLine + 1
        Loop
    
        Print #1, mChr18 & "Customer's Signature"
' SALE FOOTER
    '22 space maintain between heading and :
    RstJob.MoveFirst
    If RstJob!Det_Tax = 1 Then
        Print #1, Replace(Space(21), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
    
        Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(Txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
        ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
        
        Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Tax Surc. ", 10, 0) & PSTR(RstJob!Tax_Sur_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Sur_Amt, 12, 2) & mDoub
      
        Print #1, PSTR("Spares Amount", 16) & PSTR(RstJob!SprAmt_TP, 12, 2) & Space(8) & PSTR(RstJob!SprAmt_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Misc. Charges", 16) & PSTR(RstJob!Packing, 12, 2) & mDoub

'"Itemwise Dis.Amt 01234567.12 00.00% 01234567.12 | Itemwise Dis.Amt 01234567.12"
'col1(16) col2(28) col3(35) col4(47) ,col5(50) ,col6(66) ,col7(78)
'col1(16) col2(12) col3(7) col4(12) ,col5(3) ,col6(16) ,col7(12)

        Print #1, PSTR("Oil Amount ", 16) & PSTR(RstJob!OilAmt_TP + RstJob!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstJob!OilAmt_TB + RstJob!OilAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & mEmph & PSTR("Sub Total[TP+TB]", 16) & PSTR(Val(Txt(STotB)), 12, 2) & mEmph1
        
        Print #1, PSTR("Discount ", 10, 0) & PSTR(RstJob!D_Per_TP, 5, 2) & "%" & PSTR(RstJob!D_Amt_TP, 12, 2) & PSTR(RstJob!D_Per_TB, 7, 2) & "%" & PSTR(RstJob!D_Amt_TB, 12, 2) _
        ; " | " & PSTR("TO Tax ", 10, 0) & PSTR(RstJob!TOT_Per, 5, 2) & "%" & PSTR(RstJob!Tot_Amt, 12, 2) & mEmph
        
        Print #1, PSTR("Sub Total [A]", 16) & PSTR(Val(Txt(STotATP)), 12, 2) & Space(8) & PSTR(Val(Txt(STotATB)), 12, 2) & mEmph1 _
        ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstJob!ReSalTax_Per, 5, 2) & "%" & PSTR(RstJob!ReSalTax_Amt, 12, 2)
        
        Print #1, PSTR("Gen Surch ", 10, 0) & PSTR(RstJob!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 12, 2) & PSTR(RstJob!Gen_Sur_Amt, 20, 2) _
        ; " | " & PSTR("Round Off", 16) & PSTR(RstJob!Rounded, 12, 2)
       
        Print #1, PSTR("Transportation", 16) & PSTR(0, 12, 2) & PSTR(RstJob!Trans_Amt, 20, 2) _
        ; " | " & mEmph & PSTR("Net Payble Rs.", 16) & PSTR(Val(Txt(NetSprAmt)), 12, 2) & mEmph1
    Else
        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        Print #1, Space(45) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
        If RstJob!D_Amt_TP + RstJob!D_Amt_TB > 0 Then
            Print #1, Space(45) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstJob!D_Amt_TP + RstJob!D_Amt_TB, 12, 2)
        Else
            Print #1, ""
        End If
        Print #1, Space(45) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Val(Txt(NetSprAmt)) - (mGrossAmt - (RstJob!D_Amt_TP + RstJob!D_Amt_TB)), 12, 2) & mEmph
        Print #1, Space(45) & PSTR("Net Payble Rs.", 20) & " : " & PSTR(Val(Txt(NetSprAmt)), 12, 2) & mEmph1
    End If
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(Val(Txt(NetSprAmt)), "Rupees", "Paise") & mDoub1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
    Print #1, PSTR(mTaxdesc, 25) & Space(PageWidth - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1 & mDoub
    Print #1, ""
    Print #1, "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    J = 1
    For I = 1 To Len(Footer)
       If mID(Footer, I, 1) = vbLf Then
           Print #1, RTrim(mID(Footer, J, I - J))
           J = I + 1
       End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
'Gate Pass Footer()
    If RstJob!Printed_YN = 0 Then
        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "S" Then
            Print #1, Replace(Space(PageWidth), " ", "-")
            Print #1, PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", 80) & mEmph
            Print #1, "GATE PASS No. & DATE : " & XNull(RstJob!GP_No) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(10) & "Job Card No. : " & PrinID(RstJob!JobDocID)
            Print #1, "Vehicle No. : " & XNull(RstJob!RegNo) & Space(5) & "Chassis No. : " & XNull(RstJob!Chassis) _
            & Space(5) & mChr17 & "Model : " & XNull(RstJob!Model) & mChr18
            Print #1,
            Print #1, "Vehicle has been received from workshop & work done as per  my satisfaction."
            Print #1, ""
            Print #1, "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
            Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
        End If
    End If
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
'        'Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'        Print #1, "Type C:\RepPrint.Txt> Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Spare Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & Txt(JobNo).Tag & "'"
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintLab(mQRY$)
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
    Dim I As Integer, J As Integer
    Dim PrintStr As String
    Dim rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double, mLabourAmt As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double
    
    Set RstJob = GCn.Execute(mQRY)
'    If RstJob.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(RstJob!LabInvFooter)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
 
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 -> X * 0.56
    'chr 10 to chr 17 -> X * 1.7
        
    mHeader = 0   'Ideal 17
    mFooter = 15
    mFooter = mFooter + FooterCnt
    mGatePass = 8
    'modi lps 03-04-2003
'    mFooter = IIf(RstJob!LabBillPrinted = 0, mFooter + mGatePass, mFooter)
    If RstJob!LabBillPrinted = 0 Then   'Not Printed
        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "L" Then    'GatePass on Labour Bill Required
            mFooter = mFooter + mGatePass
        End If
    End If
    'eof lps
    
    'Header
    If RstJob!CrMemo = 0 Then
        mDocStr = "CASH MEMO"
    Else
        mDocStr = "INVOICE"
    End If
    mDupStr = IIf(RstJob!LabBillPrinted = 1, "(DUPLICATE)", "")
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!W_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
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
    Print #1, PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 27, , AlignRight, " ")
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
    mHeader = mHeader + 1
   
    Print #1, PRN_TIT("** LABOUR " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, mChr18 & Space(48) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocID_InvLab) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
    mHeader = mHeader + 1
'    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(2) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
    & Space(4) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") '& mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("Srl.", 4) & "<-------------Labour Detail-------------->" & " " & PSTR("Hrs", 10, , AlignRight) & PSTR("Rate", 10, , AlignRight) & PSTR("Amount", 12, , AlignRight)
    mHeader = mHeader + 1
    Print #1, PSTR("No.", 4) & PSTR("Code", 7) & PSTR("Description", 35) '& mDoub1 & mChr18
    mHeader = mHeader + 1
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    
    If RstJob.RecordCount > 0 Then
        I = 1
        Do Until RstJob.EOF
            If mLine > mFix Then
                Page = Page + 1
                Print #1, mChr18 & Replace(Space(PageWidth), " ", "-")
                Print #1, Space(PageWidth - Len("Contd. on next page.." + str(Page))) & "Contd. on next page.." & str(Page)
                Do Until mLine >= mFix + mFooter - 2
                    Print #1, ""
                    mLine = mLine + 1
                Loop
                Print #1, mEject
                'Header On Second Page
                mHeader = 0
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                If XNull(RstCompDet!W_SecSpeciality) <> "" Then
                    Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
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
                Print #1, PRN_TIT("** LABOUR " & mDocStr & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, mChr18 & Space(48) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocID_InvLab) & mEmph1
                mHeader = mHeader + 1
                Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                mHeader = mHeader + 1
                Print #1, PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
                mHeader = mHeader + 1
                Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
                mHeader = mHeader + 1
'                Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
'                mHeader = mHeader + 1
'                Print #1, PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
'                & Space(4) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
'                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-") '& mDoub
                mHeader = mHeader + 1
                Print #1, PSTR("Srl.", 4) & "<-------------Labour Detail-------------->" & " " & PSTR("Hrs", 10, , AlignRight) & PSTR("Rate", 10, , AlignRight) & PSTR("Amount", 12, , AlignRight)
                mHeader = mHeader + 1
                Print #1, PSTR("No.", 4) & PSTR("Code", 7) & PSTR("Description", 35) '& mDoub1 & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            If RstJob!CHRG_FROM = "C" Then
                mLabourAmt = RstJob!LabourAmt  'Lab_Rate
            Else
                mLabourAmt = 0
            End If
            PrintStr = PSTR(Trim(str(mSlNo)) & ".", 4) & PSTR(RstJob!Lab_Code, 6, , AlignLeft) & " " & PSTR(RstJob!LabName, 35) & " " & PSTR(RstJob!Hrs_Taken, 10, 2) & PSTR(RstJob!Lab_Rate, 10, 2) & PSTR(mLabourAmt, 12, 2)
            Print #1, PrintStr
            RstJob.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    Else
        Print #1, PRN_TIT("** No Labour **", "A", PageWidth)
        mLine = mLine + 1
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    Print #1, mChr18 & "Customer's Signature"
' SALE FOOTER
    RstJob.MoveFirst
    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
    Print #1, Space(45) & PSTR("TOTAL AMOUNT", 20) & " : " & PSTR(RstJob!LabAmt_TB + RstJob!LabAmt_TP, 12, 2) & mDoub1
    
    If RstJob!Lab_D_Amt > 0 Then
        Print #1, Space(45) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstJob!Lab_D_Amt, 12, 2)
    Else
        Print #1, ""
    End If
    Print #1, Space(45) & PSTR("SERVICE TAX @" & str(RstJob!Lab_TaxPer), 20) & " : " & PSTR(RstJob!Lab_TaxAmt, 12, 2)
    Print #1, Space(45) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(RstJob!Lab_RoundOff, 12, 2) & mEmph
    Print #1, Space(45) & PSTR("Net Payble Rs.", 20) & " : " & PSTR(RstJob!NetLab_Amt, 12, 2) & mEmph1

    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(RstJob!NetLab_Amt, "Rupees", "Paise") & mDoub1
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & "E & O.E." & mChr18 & Space(PageWidth - (Len("For " & PubComp_Name) + 6)) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, "" & mDoub
    Print #1, "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    J = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, J, I - J))
            J = I + 1
        End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
'Gate Pass Footer()
    If RstJob!LabBillPrinted = 0 Then
        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "L" Then
            Print #1, Replace(Space(PageWidth), " ", "-")
            Print #1, PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", 80) & mEmph
            Print #1, "GATE PASS No. & DATE : " & XNull(RstJob!GP_No) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(10) & "Job Card No. : " & PrinID(RstJob!JobDocID)
            Print #1, "Vehicle No. : " & XNull(RstJob!RegNo) & Space(5) & "Chassis No. : " & XNull(RstJob!Chassis) _
            & Space(5) & mChr17 & "Model : " & XNull(RstJob!Model) & mChr18
            Print #1,
            Print #1, "Vehicle has been received from workshop & work done as per  my satisfaction."
            Print #1, ""
            Print #1, "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
            Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
        End If
    End If
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Labour Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update Job_Card set LabBillPrinted = 1 where Job_Card.DocId='" & Txt(JobNo).Tag & "'"
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintSrvLet()
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
    Dim I As Integer, J As Integer
    Dim PrintStr As String
    Dim rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim SrvGatePassOn$, Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim fob As New FileSystemObject
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 -> X * 0.56
    'chr 10 to chr 17 -> X * 1.7
        
    mHeader = 0   'Ideal 17
    mFooter = 15
    mFooter = mFooter + FooterCnt
    
    mDocStr = "Next Service Letter"
    
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!W_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
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
    Print #1, PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "Phone: " & XNull(RstCompDet!W_SecPhone)), 27, , AlignLeft, " ") & Space(8) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "Fax: " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "To,"
    mHeader = mHeader + 1
    Print #1, mSP2 & mEmph & PSTR(Txt(OwnerName), 40) & mEmph1
    mHeader = mHeader + 1
    If Txt(Address1) <> "" Then
        Print #1, mSP2 & PSTR(Txt(Address1), 40) & mEmph1
        mHeader = mHeader + 1
    End If
    If Txt(Address2) <> "" Then
        Print #1, mSP2 & PSTR(Txt(Address2), 40) & mEmph1
        mHeader = mHeader + 1
    End If
    If Txt(Address3) <> "" Then
        Print #1, mSP2 & PSTR(Txt(Address3), 40) & mEmph1
        mHeader = mHeader + 1
    End If
    Print #1, mSP2 & PSTR(Txt(City), 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "Dear Vehicle Owner,"
    mHeader = mHeader + 1
    Print #1, mSP2 & "We are pleased to attend your " & Txt(Model) & " Vehicle at our workshop and hope"
    mHeader = mHeader + 1
    Print #1, mSP2 & "you are satisfied with our working & services."
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "Please remember,next service of your vehicle "
    mHeader = mHeader + 1
    Print #1, mSP2 & Space(10) & "Reg.No.    :" & Txt(VehRegNo)
    mHeader = mHeader + 1
    Print #1, mSP2 & Space(10) & "Chassis No.:" & Txt(Chassis)
    mHeader = mHeader + 1
    Print #1, mSP2 & " should be carried out before dt." & Txt(NextSrv) & "."
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "We shall put in all efforts in providing services in all times."
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "Truely your's"
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "Works Manager"
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, Space((PageWidth - Len("*a dataman software*")) / 2) & mChr17 & "*a dataman software*" & mChr18
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\WinNt") Then
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Sub SpeedPrintBoth(mQRY$, PrePrinted As Boolean)
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
    Dim I As Integer, J As Integer, K As Integer
    Dim PrintStr As String
    Dim rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim SrvGatePassOn$, Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double
    Dim MRPTaxStr$, mTPAmtStr$, mTBAmtStr$
    Dim mSprCaption As Boolean, mLabCaption As Boolean, mLabDiscAmtStr$
    Dim mTotRow, mTotRowTemp As Integer
    
    Set RstJob = GCn.Execute(mQRY)
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
    'chr 17 to chr 10 -> X * 0.56
    'chr 10 to chr 17 -> X * 1.7
    mHeader = 0   'Ideal 17
    mFooter = 22    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
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
    'eof modi

    'Sale Bill Header
    If Not Provisional Then
        If RstJob!CrMemo = 0 Then
            mDocStr = "CASH MEMO"
        Else
            mDocStr = "INVOICE"
        End If
    Else
        mDocStr = "PROVISIONAL BILL "
    End If
    If Not Provisional Then
        mDupStr = IIf(RstJob!Printed_YN = 1, "(DUPLICATE)", "")
    Else
        mDupStr = ""
    End If
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
    mTaxdesc = GCn.Execute("select Printing_Desc from TaxForms where Form_Code = '" & RstJob!Form_Code & "'").Fields(0).Value
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    
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
        Print #1, mSP2 & PRN_TIT("** WORKSHOP " & mDocStr & mDupStr & " **", "B", PageWidth)
        mHeader = mHeader + 1
        Print #1, mSP2 & mChr18 & Space(36) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(RstJob!Add1), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(RstJob!Add2), 40) & Space(6) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
        & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
        mHeader = mHeader + 1
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
        mHeader = mHeader + 1
        If RstJob!Det_Tax = 1 Then
            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----->" & "<---------AMOUNT--------->"
            mHeader = mHeader + 1
            Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
            mHeader = mHeader + 1
        Else
            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 27) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18 '& mDoub1
            mHeader = mHeader + 1
        End If
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        If RstJob!Orig = "1" Then
            Print #1, mSP2 & mEmph & "*Spare Details*" & mEmph1 & mChr17
            mHeader = mHeader + 1
            mSprCaption = True
        ElseIf RstJob!Orig = "2" Then
            Print #1, mSP2 & mEmph & "*Labour Details*" & mEmph1 & mChr17
            mHeader = mHeader + 1
            mLabCaption = True
        End If
        mFix = PageLength - (mHeader + mFooter)
        Page = 1
        mLine = 1
        mSlNo = 1
        LAdd = VNull(RstJob!Gen_Sur_Amt) + VNull(RstJob!Trans_Amt) + VNull(RstJob!Tax_Amt) + VNull(RstJob!Tax_Sur_Amt) + VNull(RstJob!Packing) + VNull(RstJob!ReSalTax_Amt) + VNull(RstJob!Tot_Amt)
        SubTot = RstJob!SprAmt_TB + RstJob!SprAmt_TP + RstJob!SprAmt_MRP_TB + RstJob!SprAmt_MRP_TP _
        + RstJob!OilAmt_TB + RstJob!OilAmt_TP + Val(Txt(IWDiscTotTP).TEXT) + Val(Txt(IWDiscTotTB).TEXT)
        mTotRow = RstJob.RecordCount
        mTotRowTemp = RstJob.RecordCount
        If RstJob.RecordCount > 0 Then
            I = 1
            Do Until RstJob.EOF = True
                If mTotRow > 30 Then
                    mFix = 30
                ElseIf mTotRow > 15 And mTotRow <= 30 Then
                    mFix = 30
                Else
                    mFix = (PageLength - (mHeader + mFooter))
                End If
                If mLine > mFix Then
                    Page = Page + 1
                    mTotRow = mTotRow - 30
                    Print #1, mChr18 & mSP2 & Replace(Space(PageWidth), " ", "-")
                    Print #1, mSP2 & Space((PageWidth) - Len("Contd. on next page.." + str(Page))) & "Contd. on next page.." & str(Page)
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
                    Print #1, mSP2 & PRN_TIT("** WORKSHOP " & mDocStr & mDupStr & " **", "B", PageWidth)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & mChr18 & Space(46) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(RstJob!Add1), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(RstJob!Add2), 40) & Space(6) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
                    & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
                    mHeader = mHeader + 1
                    If RstJob!Det_Tax = 1 Then
                        Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----->" & "<---------AMOUNT--------->"
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
                    If RstJob!Orig = "2" Then
                        Print #1, mSP2 & mChr18 & mEmph & "*Labour Details*" & mEmph1 & mChr17
                        mHeader = mHeader + 1
                        mLabCaption = True
                    End If
                End If
                mRate = IIf(RstJob!MRP_YN = 1, RstJob!MRP_Rate2, RstJob!Rate2)
                If RstJob!Orig = "1" Then
                    If RstJob!Det_Tax = 1 Then
                        mTPAmtStr = PSTR(0, 12, 2)
                        mTBAmtStr = PSTR(0, 12, 2)
                    If RstJob!Purpose = "W" Then
                        mTBAmtStr = "*Warranty*"
                    ElseIf RstJob!Purpose = "P" Then
                        mTBAmtStr = "*PDI*"
                    ElseIf RstJob!Purpose = "F" Then
                        mTBAmtStr = "*Free*"
                    ElseIf RstJob!Purpose = "L" Then
                        mTBAmtStr = "*Compliment*"
                    ElseIf RstJob!Purpose = "O" Then
                        mTBAmtStr = "*Company*"
                    Else
                        If RstJob!Tax_YN = 0 Then
                            mTPAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                            mTBAmtStr = PSTR(0, 12, 2)
                        Else
                            mTPAmtStr = PSTR(0, 12, 2)
                            mTBAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                        End If
                    End If
                    'modishekhar
                    If RstJob!Purpose = "W" And GCn.Execute("select PrnWarrSpr from syctrl").Fields(0).Value = 0 Then GoTo NXT
                    'modi lps at Cuttack 31.08.03
                    If RstJob!Purpose = "L" And GCn.Execute("select PrintComplIssue from syctrl").Fields(0).Value = 0 Then GoTo NXT
                    If RstJob!Purpose = "O" And GCn.Execute("select PrintCompanyIssue from syctrl").Fields(0).Value = 0 Then GoTo NXT
                    If RstJob!Qty_Iss - RstJob!Qty_Ret <= 0 Then GoTo NXT
                    'eof modi
                    PrintStr = PSTR(Trim(str(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 34) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
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
                    PrintStr = PSTR(Trim(str(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 26, , AlignLeft) & PSTR(RstJob!Part_Name, 40) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(LdRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", "L") & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & _
                    PSTR(LDAmt - RstJob!Disc_Amt2, 12, 2)
                End If
            Else    'Labour
                If Val(Txt(ServTaxAmt)) <= 0 Then
                    mTPAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                    mTBAmtStr = PSTR(0, 12, 2)
                Else
                    mTPAmtStr = PSTR(0, 12, 2)
                    mTBAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                End If
                PrintStr = PSTR(Trim(str(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 34) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
            End If
            Print #1, mSP2 & PrintStr
            mSlNo = mSlNo + 1
            mLine = mLine + 1
NXT:
            RstJob.MoveNext
            If RstJob.EOF = True And (mTotRow > 15 And mTotRow <= 30) Then
                   RstJob.MovePrevious
                   Page = Page + 1
                   Do Until mTotRow >= 30
                         Print #1, ""
                        mTotRow = mTotRow + 1
                    Loop
                    Print #1, mChr18 & mSP2 & Replace(Space(PageWidth), " ", "-")
                    Print #1, mSP2 & Space((PageWidth) - Len("Contd. on next page.." + str(Page))) & "Contd. on next page.." & str(Page)
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
                    Print #1, mSP2 & PRN_TIT("** WORKSHOP " & mDocStr & mDupStr & " **", "B", PageWidth)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & mChr18 & Space(46) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(RstJob!Add1), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(RstJob!Add2), 40) & Space(6) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
                    & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----->" & "<---------AMOUNT--------->"
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
    If RstJob!Det_Tax = 1 Then
        Print #1, mSP2 & Replace(Space(20), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
    
        Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(Txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
        ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
        
        Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Tax Surc. ", 10, 0) & PSTR(RstJob!Tax_Sur_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Sur_Amt, 12, 2) & mDoub
      
        Print #1, mSP2 & PSTR("Spares Amount", 16) & PSTR(RstJob!SprAmt_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Misc. Charges", 16) & PSTR(RstJob!Packing, 12, 2) & mDoub

        Print #1, mSP2 & PSTR("Oil Amount ", 16) & PSTR(RstJob!OilAmt_TP + RstJob!OilAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!OilAmt_TB + RstJob!OilAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & mEmph & PSTR("Sub Total[TP+TB]", 16) & PSTR(Val(Txt(STotB)), 12, 2) & mEmph1
        
        Print #1, mSP2 & PSTR("Discount ", 10, 0) & PSTR(RstJob!D_Per_TP, 5, 2) & "%" & PSTR(RstJob!D_Amt_TP, 11, 2) & PSTR(RstJob!D_Per_TB, 7, 2) & "%" & PSTR(RstJob!D_Amt_TB, 12, 2) _
        ; " | " & PSTR("TO Tax ", 10, 0) & PSTR(RstJob!TOT_Per, 5, 2) & "%" & PSTR(RstJob!Tot_Amt, 12, 2) & mEmph
        
        Print #1, mSP2 & PSTR("Sub Total [A]", 16) & PSTR(Val(Txt(STotATP)), 11, 2) & Space(8) & PSTR(Val(Txt(STotATB)), 12, 2) & mEmph1 _
        ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstJob!ReSalTax_Per, 5, 2) & "%" & PSTR(RstJob!ReSalTax_Amt, 12, 2)
        
        Print #1, mSP2 & PSTR("Gen Surch ", 10, 0) & PSTR(RstJob!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 11, 2) & PSTR(RstJob!Gen_Sur_Amt, 20, 2) _
        ; " | " & PSTR("Round Off", 16) & PSTR(Round(RstJob!Rounded, 2), 12, 2)
       
        Print #1, mSP2 & PSTR("Transportation", 16) & PSTR(0, 11, 2) & PSTR(RstJob!Trans_Amt, 20, 2) _
        ; " | " & mEmph & PSTR("Net Spare + Lub. Rs.", 16) & PSTR(Val(Txt(NetSprAmt)), 12, 2) & mEmph1
    Else
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mDoub
        Print #1, mSP2 & Space(44) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
        If RstJob!D_Amt_TP + RstJob!D_Amt_TB > 0 Then
            Print #1, mSP2 & Space(44) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstJob!D_Amt_TP + RstJob!D_Amt_TB, 12, 2)
        Else
            Print #1, ""
        End If
        Print #1, mSP2 & Space(44) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Val(Txt(NetSprAmt)) - (mGrossAmt - (RstJob!D_Amt_TP + RstJob!D_Amt_TB)), 12, 2) & mEmph
        Print #1, mSP2 & Space(44) & PSTR("Net Spare + Lub. Rs.", 20) & " : " & PSTR(Val(Txt(NetSprAmt)), 12, 2) & mEmph1
    End If

    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    If Val(Txt(LabDisc)) > 0 Then
        mLabDiscAmtStr = "Discount  : " & PSTR(Val(Txt(LabDisc)), 8, 2)
    Else
        mLabDiscAmtStr = Space(19)
    End If
    PrintStr = "Total Lab.: " & PSTR(Val(Txt(LabAmt)), 8, 2)
    PrintStr = PrintStr & " |Serv.Tax @ " & PSTR(Val(Txt(ServTaxPer)), 4, 2) & ":" & PSTR(Val(Txt(ServTaxAmt)), 9, 2) & " |" & mEmph & "Net Labour Rs.    : " & PSTR(Val(Txt(NetLabAmt)), 9, 2) & mEmph1
    Print #1, mSP2 & PrintStr
    PrintStr = mLabDiscAmtStr
    PrintStr = PrintStr & " |" & "Round Off      :  " & PSTR(Val(Txt(LabROff)), 7, 2) & " |" & mEmph & "Net Payble Amt Rs.: " & PSTR(Val(Txt(NetAmt)), 9, 2) & mEmph1
    Print #1, mSP2 & PrintStr
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    Print #1, mSP2 & mDoub & ntow(Txt(NetAmt), "Rupees", "Paise") & mDoub1
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    Print #1, mSP2 & mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
    Print #1, mSP2 & PSTR(mTaxdesc, 25) & Space(PageWidth - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, ""
    Print #1, mSP2 & mDoub & "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    J = 1
    For I = 1 To Len(Footer)
       If mID(Footer, I, 1) = vbLf Then
           Print #1, mSP2 & RTrim(mID(Footer, J, I - J))
           J = I + 1
       End If
    Next
    Print #1, mSP2 & Space((((PageWidth) * 1.7) - Len("* a dataman software *" & "   " & pubUName & "   " & PubServerDate)) / 2) & "* a dataman software *" & "   " & pubUName & "   " & PubServerDate & mChr18
'Gate Pass Footer()
    If RstJob!Printed_YN = 0 Then
        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "S" Then
            Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
            Print #1, mSP2 & PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", (PageWidth)) & mEmph
            Print #1, mSP2 & "GATE PASS No. & DATE : " & XNull(RstJob!GP_No) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(10) & "Job Card No. : " & PrinID(RstJob!JobDocID)
            Print #1, mSP2 & "Vehicle No. : " & XNull(RstJob!RegNo) & Space(5) & "Chassis No. : " & XNull(RstJob!Chassis) _
            & Space(5) & mChr17 & "Model : " & XNull(RstJob!Model) & mChr18
            Print #1, ""
            Print #1, mSP2 & "Vehicle has been received from workshop & work done as per  my satisfaction."
            Print #1, ""
            Print #1, mSP2 & "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
            Print #1, mSP2 & mChr17 & Space((((PageWidth) * 1.7) - Len("* a dataman software *" & "   " & pubUName & "   " & PubServerDate)) / 2) & "* a dataman software *" & "   " & pubUName & "   " & PubServerDate & mChr18
        End If
    End If
    'End If
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    FirstPrint = IIf(FirstPrint, FirstPrint, True)
    If fob.FolderExists("c:\WinNt") Then
        'Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'        Print #1, "Type C:\RepPrint.Txt> Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt>" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If Provisional Then
        MsgBox "Provisional Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !"
    Else
        If MsgBox("Spare Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
            GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & Txt(JobNo).Tag & "'"
        End If
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub WindowsPrintSpr(Index As Integer, mQRY$)
Dim I As Integer, RstJob As ADODB.Recordset, RST1 As ADODB.Recordset, mDocStr$
Dim mPrintGatePass As Byte
On Error GoTo ERRORHANDLER
Set RstJob = GCn.Execute(mQRY)
If RstJob.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

CreateFieldDefFile RstJob, PubRepoPath + "\" & mRepName & ".ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath & "\" & mRepName & ".RPT")
rpt.Database.SetDataSource RstJob
rpt.ReadRecords
RstJob.MoveFirst
Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
mDocStr = IIf(RstJob!CrMemo = 0, "CASH MEMO", "INVOICE")
If RstJob!Printed_YN = 0 Then
    If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "S" Then
        mPrintGatePass = 1
    End If
End If

        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("Speciality")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecSpeciality & "'"
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
                    'modishekhar
                Case UCase("PrnWarrSpr")
                    rpt.FormulaFields(I).TEXT = "" & GCn.Execute("select PrnWarrSpr from syctrl").Fields(0).Value & ""
                Case UCase("HeaderTitle")
                    rpt.FormulaFields(I).TEXT = "'WORKSHOP SPARE " & mDocStr & "'"
                Case UCase("PrintGatePass")
                    rpt.FormulaFields(I).TEXT = mPrintGatePass
            End Select
        Next
Select Case Index
    Case PScreen  'screen
         Call Report_View(rpt, mDocStr, , True)
    Case PWindows 'Printer
        'Report_DocHeader rpt
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
                    rpt.FormulaFields(I).TEXT = "'" & mDocStr & "'"
            End Select
        Next
        rpt.PrintOut False
        If MsgBox("Spare Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
            GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & Txt(JobNo).Tag & "'"
        End If
End Select
CmdPrint(0).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrintLab(Index As Integer, mQRY$)
Dim I As Integer, Rst As ADODB.Recordset, RST1 As ADODB.Recordset, mDocStr$
Dim mPrintGatePass As Byte
On Error GoTo ERRORHANDLER

Set Rst = GCn.Execute(mQRY)
'If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName1 & ".ttx", True
    Set rpt = rdApp.OpenReport(PubRepoPath & "\" & mRepName1 & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Rst.MoveFirst
    If ChkRep(ChkSprBoth).Value = Checked Then Exit Sub
    Set RST1 = New Recordset
    RST1.CursorLocation = adUseClient
    RST1.Open "select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    mDocStr = IIf(Rst!CrMemo = 0, "CASH MEMO", "INVOICE")
    If Rst!LabBillPrinted = 0 Then
        If Rst!SrvGatePass = 1 And Rst!SrvGatePass_On <> "S" Then
            mPrintGatePass = 1
        End If
    End If
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("Speciality")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecSpeciality & "'"
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecSpeciality & "'"
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
            Case UCase("HeaderTitle")
                rpt.FormulaFields(I).TEXT = "'LABOUR " & mDocStr & "'"
            Case UCase("PrintGatePass")
                rpt.FormulaFields(I).TEXT = mPrintGatePass
        End Select
    Next
Select Case Index
    Case PScreen  'screen
         Call Report_View(rpt, mDocStr, , True)
    Case PWindows 'Printer
'        Report_DocHeader rpt
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
                    rpt.FormulaFields(I).TEXT = "'" & mDocStr & "'"
            End Select
        Next
        rpt.PrintOut False
        If MsgBox("Labour Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
            GCn.Execute "update Job_Card set LabBillPrinted = 1 where Job_Card.DocId='" & Txt(JobNo).Tag & "'"
        End If
End Select
CmdPrint(0).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrint2(Index As Integer)
Dim I As Integer, Rst As ADODB.Recordset, RST1 As ADODB.Recordset
On Error GoTo ERRORHANDLER
    Set Rst = GCn.Execute("SELECT City.CityName, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, SubGroup.Phone, Syctrl.SprInvFooter, Syctrl.WorkShopInvFooter, Part.Part_Name,Job_Card.Job_No, HisCard.Model, HisCard.RegNo, HisCard.Chassis,Sp_Stock.*, Sp_Sale.* " & _
    "FROM (((((Sp_sale LEFT JOIN Sp_Stock as Sp_Stock ON Sp_Sale.DocID = Sp_Stock.Invoice_DocId) LEFT JOIN Part ON Sp_Stock.Part_No = Part.Part_No and Part.Div_Code = left(SP_Stock.DocID,1)) " & _
    "LEFT JOIN SubGroup ON Sp_Sale.Party_Code = SubGroup.SubCode) LEFT JOIN (Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) ON Sp_Sale.Job_DocID = Job_Card.DocId) LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN Syctrl ON Syctrl.LinkTable >= Sp_Sale.U_AE " & _
    "where Sp_Sale.Job_DocID='" & Master!Code & "' and Sp_Stock.Qty_Iss-Sp_Stock.Qty_Ret<>0")
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

        CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName2 & ".ttx", True
        Set rpt = rdApp.OpenReport(PubRepoPath & "\" & mRepName2 & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        Set RST1 = New Recordset
        RST1.CursorLocation = adUseClient
        RST1.Open "select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic

        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("SubTitle")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecSpeciality & "'"
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
                Case UCase("Title")
                    rpt.FormulaFields(I).TEXT = "'SPARES INVOICE'"
            End Select
        Next
Select Case Index
    Case PScreen  'screen
         Call Report_View(rpt, "SPARES INVOICE", , True)
    Case PWindows 'Printer
        rpt.PrintOut False
        If MsgBox("Spare Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
            GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & Txt(JobNo).Tag & "'"
        End If
End Select
CmdPrint(0).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

