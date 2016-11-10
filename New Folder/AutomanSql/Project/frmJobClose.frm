VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmJobClose 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Job Close/Unclose Entry"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14925
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
   ScaleHeight     =   9300
   ScaleWidth      =   14925
   Begin VB.TextBox Txt 
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
      Index           =   95
      Left            =   5625
      TabIndex        =   228
      Top             =   7200
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox Txt 
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
      Index           =   94
      Left            =   5715
      TabIndex        =   227
      Top             =   7125
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton Pwinprint 
      Caption         =   "Windows Provisional Bill"
      Height          =   300
      Left            =   10800
      TabIndex        =   226
      Top             =   60
      Visible         =   0   'False
      Width           =   2310
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
      Height          =   1845
      Left            =   7785
      TabIndex        =   170
      Top             =   2865
      Visible         =   0   'False
      Width           =   5025
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
         TabIndex        =   175
         Top             =   840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Depreciation Bill"
         DisabledPicture =   "frmJobClose.frx":0000
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
         Index           =   6
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   222
         ToolTipText     =   "Screen"
         Top             =   1485
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00F8D7FD&
         Caption         =   "Print Oth.Dlr.Bill"
         DisabledPicture =   "frmJobClose.frx":030A
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
         Index           =   5
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   196
         ToolTipText     =   "Printer "
         Top             =   1185
         Width           =   1590
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
         Picture         =   "frmJobClose.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   180
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
         Picture         =   "frmJobClose.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Screen"
         Top             =   1470
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
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
         Index           =   0
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   179
         ToolTipText     =   "Printer "
         Top             =   885
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmJobClose.frx":0F96
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
         TabIndex        =   178
         ToolTipText     =   "Screen"
         Top             =   585
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmJobClose.frx":12A0
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
         TabIndex        =   177
         ToolTipText     =   "Printer "
         Top             =   300
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
         TabIndex        =   183
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
         TabIndex        =   182
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
         TabIndex        =   181
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         Left            =   1710
         TabIndex        =   173
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
         TabIndex        =   174
         Top             =   600
         Value           =   1  'Checked
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
         TabIndex        =   185
         Top             =   1485
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
         TabIndex        =   184
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
      Index           =   97
      Left            =   1845
      TabIndex        =   224
      Text            =   "99999999.99"
      Top             =   6630
      Visible         =   0   'False
      Width           =   1080
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   96
      Left            =   12555
      TabIndex        =   223
      Text            =   "99999999.99"
      Top             =   810
      Visible         =   0   'False
      Width           =   1080
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   93
      Left            =   7125
      TabIndex        =   200
      Top             =   4155
      Width           =   1695
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   92
      Left            =   3510
      TabIndex        =   40
      ToolTipText     =   "General Surcharge %"
      Top             =   3930
      Width           =   510
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   91
      Left            =   1830
      TabIndex        =   38
      ToolTipText     =   "General Surcharge %"
      Top             =   3930
      Width           =   510
   End
   Begin VB.TextBox Txt 
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
      Index           =   90
      Left            =   7590
      MaxLength       =   40
      TabIndex        =   75
      Text            =   "012345678901234567890123456789"
      Top             =   5955
      Width           =   1230
   End
   Begin VB.TextBox Txt 
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
      Index           =   89
      Left            =   5085
      MaxLength       =   20
      TabIndex        =   74
      Text            =   "012345678901234567890123456789"
      Top             =   5955
      Width           =   1995
   End
   Begin VB.TextBox Txt 
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
      Index           =   88
      Left            =   1830
      MaxLength       =   20
      TabIndex        =   73
      Text            =   "012345678901234567890123456789"
      Top             =   5955
      Width           =   2115
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   87
      Left            =   1920
      TabIndex        =   215
      Top             =   7740
      Visible         =   0   'False
      Width           =   1290
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
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   212
      TabStop         =   0   'False
      Top             =   4155
      Width           =   1620
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
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   4155
      Width           =   1620
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
      Left            =   9555
      TabIndex        =   210
      Top             =   7380
      Width           =   1170
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   60
      Left            =   9030
      TabIndex        =   209
      ToolTipText     =   "Surcharge % on Local Sales Tax"
      Top             =   7380
      Width           =   495
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   82
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   208
      TabStop         =   0   'False
      Top             =   4155
      Width           =   1290
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   86
      Left            =   10425
      TabIndex        =   205
      Top             =   5055
      Width           =   1290
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   85
      Left            =   9870
      TabIndex        =   204
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5055
      Width           =   510
   End
   Begin VB.TextBox Txt 
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
      Index           =   76
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   76
      Text            =   "012345678901234567890123456789"
      Top             =   6180
      Width           =   4035
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
      Left            =   7650
      TabIndex        =   199
      Top             =   3930
      Width           =   1170
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
      Left            =   7125
      TabIndex        =   198
      Text            =   "99.99"
      ToolTipText     =   "Local Sales Tax %"
      Top             =   3930
      Width           =   495
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   315
      Left            =   9600
      TabIndex        =   197
      Top             =   30
      Width           =   1185
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2850
      Left            =   2550
      Negotiate       =   -1  'True
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   8400
      Visible         =   0   'False
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   5027
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
      ColumnCount     =   4
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
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
      Left            =   10500
      TabIndex        =   195
      Top             =   8550
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton CmdPBill 
      Caption         =   "Print Provisional Bill"
      Height          =   315
      Left            =   6885
      TabIndex        =   194
      Top             =   30
      Width           =   2700
   End
   Begin VB.TextBox Txt 
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
      Index           =   84
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   192
      TabStop         =   0   'False
      Text            =   "9999999"
      Top             =   2145
      Width           =   1350
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
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   3690
      Visible         =   0   'False
      Width           =   1695
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   83
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1290
   End
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   1650
      Left            =   -4335
      Negotiate       =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   10740
      Visible         =   0   'False
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   2910
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   81
      Left            =   9870
      TabIndex        =   64
      Text            =   "99.99"
      ToolTipText     =   "Turn Over Tax %"
      Top             =   4605
      Width           =   510
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
      Index           =   80
      Left            =   7515
      MaxLength       =   5
      TabIndex        =   34
      Text            =   "WithDrawn"
      Top             =   3330
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   79
      Left            =   7650
      TabIndex        =   61
      Text            =   "2"
      Top             =   5055
      Width           =   1170
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   78
      Left            =   7125
      TabIndex        =   60
      Text            =   "2"
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5055
      Width           =   495
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
      Index           =   0
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   3930
      Width           =   1290
   End
   Begin VB.TextBox Txt 
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
      Index           =   1
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   2820
      Width           =   3885
   End
   Begin VB.TextBox Txt 
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
      Index           =   75
      Left            =   4575
      MaxLength       =   50
      TabIndex        =   72
      Top             =   5730
      Width           =   4245
   End
   Begin VB.TextBox Txt 
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
      Index           =   74
      Left            =   4575
      MaxLength       =   4
      TabIndex        =   71
      Text            =   "yes"
      Top             =   5505
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
      Left            =   10425
      TabIndex        =   65
      Text            =   "999999.99"
      Top             =   4605
      Width           =   1290
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
      Left            =   9870
      TabIndex        =   66
      ToolTipText     =   "Turn Over Tax %"
      Top             =   4830
      Width           =   510
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
      Left            =   10425
      TabIndex        =   67
      Top             =   4830
      Width           =   1290
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   72
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   5505
      Width           =   1290
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
      Left            =   1830
      TabIndex        =   54
      Top             =   5505
      Width           =   1620
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
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1695
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
      Left            =   7125
      TabIndex        =   58
      Text            =   "99.99"
      ToolTipText     =   "Turn Over Tax %"
      Top             =   4830
      Width           =   495
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   64
      Left            =   7650
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4830
      Width           =   1170
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
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1695
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   66
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Text            =   "9999999.99"
      Top             =   5505
      Width           =   1695
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   210
      Index           =   57
      Left            =   1830
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5730
      Width           =   1620
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   42
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1620
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
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1620
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
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1620
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   45
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1620
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   46
      Left            =   1830
      TabIndex        =   46
      Text            =   "99.99"
      ToolTipText     =   "Discount % Taxable"
      Top             =   4830
      Width           =   510
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
      Left            =   2370
      TabIndex        =   47
      Text            =   "99999999.99"
      Top             =   4830
      Width           =   1080
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   48
      Left            =   3510
      TabIndex        =   48
      Text            =   "99.99"
      ToolTipText     =   "Discount % Taxpaid"
      Top             =   4830
      Width           =   510
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   49
      Left            =   4050
      TabIndex        =   49
      Top             =   4830
      Width           =   1080
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
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "99999999.99"
      Top             =   5055
      Width           =   1620
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
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5055
      Width           =   1620
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
      Left            =   7125
      TabIndex        =   56
      Top             =   4380
      Width           =   1695
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
      Index           =   38
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3930
      Width           =   1080
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
      Left            =   1830
      TabIndex        =   52
      ToolTipText     =   "General Surcharge %"
      Top             =   5280
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   55
      Left            =   2400
      TabIndex        =   53
      Top             =   5280
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
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3930
      Width           =   1080
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
      Index           =   73
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Text            =   "999999.99"
      Top             =   5730
      Width           =   1290
   End
   Begin VB.TextBox Txt 
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
      Index           =   32
      Left            =   1065
      MaxLength       =   40
      TabIndex        =   31
      Text            =   "Help"
      Top             =   3330
      Width           =   2865
   End
   Begin VB.TextBox Txt 
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
      Left            =   5490
      MaxLength       =   40
      TabIndex        =   32
      Text            =   "Help"
      Top             =   3105
      Width           =   2565
   End
   Begin VB.TextBox Txt 
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
      Index           =   36
      Left            =   9420
      MaxLength       =   40
      TabIndex        =   36
      Top             =   3105
      Width           =   2310
   End
   Begin VB.TextBox Txt 
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
      Left            =   5490
      TabIndex        =   33
      Top             =   3330
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
      TabIndex        =   159
      Text            =   "Extra"
      Top             =   525
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox Txt 
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
      Left            =   10620
      TabIndex        =   37
      Top             =   3330
      Width           =   1110
   End
   Begin VB.TextBox Txt 
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
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1350
   End
   Begin VB.TextBox Txt 
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
      Left            =   10545
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1155
   End
   Begin VB.TextBox Txt 
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
      Left            =   10545
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "29/Oct/2003"
      Top             =   1920
      Width           =   1155
   End
   Begin VB.TextBox Txt 
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
      Index           =   27
      Left            =   10545
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2145
      Width           =   1155
   End
   Begin VB.TextBox Txt 
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
      Index           =   24
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1350
   End
   Begin VB.TextBox Txt 
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
      Index           =   25
      Left            =   10545
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1155
   End
   Begin VB.TextBox Txt 
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
      Index           =   23
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "9999999"
      Top             =   1695
      Width           =   1350
   End
   Begin VB.TextBox Txt 
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
      Left            =   10545
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1245
      Width           =   1155
   End
   Begin VB.TextBox Txt 
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
      Index           =   22
      Left            =   10545
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1470
      Width           =   1155
   End
   Begin VB.TextBox Txt 
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
      Index           =   20
      Left            =   7815
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1470
      Width           =   1350
   End
   Begin VB.TextBox Txt 
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
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1245
      Width           =   1350
   End
   Begin VB.TextBox Txt 
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
      Index           =   4
      Left            =   6840
      TabIndex        =   3
      Top             =   510
      Width           =   1275
   End
   Begin VB.TextBox Txt 
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
      Left            =   7815
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2595
      Width           =   3885
   End
   Begin VB.TextBox Txt 
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
      Index           =   31
      Left            =   1065
      MaxLength       =   40
      TabIndex        =   30
      Text            =   "Help"
      Top             =   3105
      Width           =   2865
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   3
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "22-APR-2002"
      Top             =   525
      Width           =   1230
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   15
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2325
      Width           =   5040
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   11
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Help"
      Top             =   1425
      Width           =   5040
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   12
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1650
      Width           =   5040
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   13
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1875
      Width           =   5040
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   14
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2100
      Width           =   5040
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   5
      Left            =   1260
      MaxLength       =   14
      TabIndex        =   4
      Text            =   "Help"
      Top             =   750
      Width           =   1740
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   2
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "Help"
      Top             =   525
      Width           =   1740
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   6
      Left            =   4200
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "Help"
      Top             =   750
      Width           =   3915
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   18
      Left            =   4785
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1515
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   17
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1395
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   9
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1740
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   8
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   975
      Width           =   3915
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   10
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2100
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   16
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1260
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   7
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   975
      Width           =   1740
   End
   Begin MSDataGridLib.DataGrid DGMech 
      Height          =   2865
      Left            =   615
      Negotiate       =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   10665
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
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
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3404.977
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGReason 
      Height          =   2865
      Left            =   5715
      Negotiate       =   -1  'True
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   9765
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   5054
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
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
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3404.977
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
      Left            =   10425
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1290
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Fgrid1 
      Height          =   2745
      Left            =   6135
      TabIndex        =   155
      Top             =   10455
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
      Top             =   7935
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox Txt 
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
      Index           =   77
      Left            =   1830
      MaxLength       =   40
      TabIndex        =   77
      Text            =   "Help"
      Top             =   6405
      Width           =   4035
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2745
      Left            =   7455
      TabIndex        =   154
      Top             =   10575
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
   Begin VB.Label LblExcise 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excise : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   30
      TabIndex        =   225
      Top             =   6660
      Visible         =   0   'False
      Width           =   735
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
      Left            =   105
      TabIndex        =   221
      Top             =   6900
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax Amt."
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
      Left            =   5220
      TabIndex        =   220
      Top             =   4170
      Width           =   1680
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
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   47
      Left            =   7125
      TabIndex        =   219
      Top             =   5955
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chq/DD No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   46
      Left            =   4035
      TabIndex        =   218
      Top             =   5970
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   45
      Left            =   30
      TabIndex        =   217
      Top             =   5955
      Width           =   1350
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Free/Warr Labour"
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
      Left            =   5295
      TabIndex        =   216
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   8010
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   5985
      X2              =   11760
      Y1              =   6570
      Y2              =   6555
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   5955
      X2              =   11730
      Y1              =   6195
      Y2              =   6195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surcharge on Tax"
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
      Left            =   7125
      TabIndex        =   214
      Top             =   7395
      Width           =   1530
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nontaxable Lab."
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
      Left            =   8925
      TabIndex        =   213
      Top             =   4170
      Width           =   1380
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET PAYBLE"
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
      Index           =   10
      Left            =   8925
      TabIndex        =   207
      Top             =   5745
      Width           =   1035
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cess"
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
      Left            =   8940
      TabIndex        =   206
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5070
      Width           =   420
   End
   Begin VB.Label LblCurrBal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr Bal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11055
      TabIndex        =   203
      Top             =   7770
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LblCurrBal1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr Bal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11055
      TabIndex        =   202
      Top             =   8025
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local Sales Tax"
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
      Left            =   5220
      TabIndex        =   201
      Top             =   3930
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon Value"
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
      Left            =   6420
      TabIndex        =   193
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Lab. Amount"
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
      Left            =   5145
      TabIndex        =   191
      Top             =   3705
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Lab."
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
      Left            =   8925
      TabIndex        =   189
      Top             =   4395
      Width           =   1095
   End
   Begin VB.Label LblSprBill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7080
      TabIndex        =   187
      Top             =   6285
      Width           =   615
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out Side Labour"
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
      Left            =   8895
      TabIndex        =   186
      Top             =   3945
      Width           =   1380
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReSale Tax               :"
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
      Left            =   5220
      TabIndex        =   169
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5070
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insp.Sheet No."
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
      Index           =   18
      Left            =   9255
      TabIndex        =   168
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. Vehicle"
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
      Index           =   19
      Left            =   9255
      TabIndex        =   167
      Top             =   1470
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Time"
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
      Left            =   9255
      TabIndex        =   166
      Top             =   1710
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp.Del.Time"
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
      Left            =   9255
      TabIndex        =   165
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esti. Labour"
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
      Left            =   9255
      TabIndex        =   164
      Top             =   2385
      Width           =   1005
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      X1              =   3525
      X2              =   5115
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Line Line3 
      X1              =   30
      X2              =   11715
      Y1              =   3060
      Y2              =   3060
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
      TabIndex        =   162
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
      TabIndex        =   161
      Top             =   8205
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
      TabIndex        =   160
      Top             =   7965
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded Off"
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
      Left            =   8925
      TabIndex        =   158
      Top             =   5295
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Dealer"
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
      Left            =   6420
      TabIndex        =   157
      Top             =   2835
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobDocID :"
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
      Left            =   8220
      TabIndex        =   151
      Top             =   705
      Width           =   960
   End
   Begin VB.Label lblGatePass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GP No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   11010
      TabIndex        =   150
      Top             =   6285
      Width           =   630
   End
   Begin VB.Label lblLabourBill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   9300
      TabIndex        =   149
      Top             =   6285
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GP No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   43
      Left            =   10335
      TabIndex        =   148
      Top             =   6285
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spr Inv No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   42
      Left            =   5970
      TabIndex        =   147
      Top             =   6285
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lab Inv No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   41
      Left            =   8175
      TabIndex        =   146
      Top             =   6285
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr A/c (Labour) : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   21
      Left            =   30
      TabIndex        =   145
      Top             =   6405
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Party : "
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
      Index           =   13
      Left            =   3495
      TabIndex        =   144
      Top             =   5745
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash (Y/N) : "
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
      Index           =   11
      Left            =   3495
      TabIndex        =   143
      Top             =   5520
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   38
      Left            =   2940
      TabIndex        =   142
      Top             =   7965
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblLabGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Show Labour"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   10575
      TabIndex        =   141
      ToolTipText     =   "Show Labour (Alt+B)"
      Top             =   3660
      UseMnemonic     =   0   'False
      Width           =   1110
   End
   Begin VB.Label lblSprGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Show Spares"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   9030
      TabIndex        =   140
      ToolTipText     =   "Show Spares (Alt+P)"
      Top             =   3660
      Width           =   1125
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   8925
      TabIndex        =   139
      Top             =   4620
      Width           =   735
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Srv. Tax"
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
      Left            =   8925
      TabIndex        =   138
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   4845
      Width           =   735
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET LABOUR "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   8925
      TabIndex        =   137
      Top             =   5520
      Width           =   1155
   End
   Begin VB.Line Line2 
      BorderStyle     =   6  'Inside Solid
      X1              =   15
      X2              =   11700
      Y1              =   3615
      Y2              =   3615
   End
   Begin VB.Label Lbl 
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
      Index           =   21
      Left            =   45
      TabIndex        =   136
      Top             =   5520
      Width           =   1245
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (B) TB+TP"
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
      Left            =   5220
      TabIndex        =   135
      Top             =   4620
      Width           =   1770
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOT on Sub Total (B)"
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
      Left            =   5220
      TabIndex        =   134
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   4845
      Width           =   1815
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded Off"
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
      Left            =   5220
      TabIndex        =   133
      Top             =   5295
      Width           =   1065
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAXABLE TOTAL"
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
      Index           =   36
      Left            =   45
      TabIndex        =   132
      Top             =   5745
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable (TB)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   12
      Left            =   2325
      TabIndex        =   131
      Top             =   3660
      Width           =   1110
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Paid (TP)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   13
      Left            =   3960
      TabIndex        =   130
      Top             =   3660
      Width           =   1155
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oil Amount"
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
      Left            =   45
      TabIndex        =   129
      Top             =   4620
      Width           =   945
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   45
      TabIndex        =   128
      Top             =   4845
      Width           =   735
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (A)"
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
      Left            =   45
      TabIndex        =   127
      Top             =   5070
      Width           =   1140
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
      TabIndex        =   126
      Top             =   7965
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. Charges"
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
      Index           =   18
      Left            =   5220
      TabIndex        =   125
      Top             =   4395
      Width           =   1200
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Surcharge"
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
      Index           =   20
      Left            =   45
      TabIndex        =   124
      Top             =   5295
      Width           =   1620
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET SPARE/LUB AMT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   30
      Left            =   5220
      TabIndex        =   123
      Top             =   5520
      Width           =   1785
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MRP Item's Amount"
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
      Left            =   45
      TabIndex        =   122
      Top             =   4170
      Width           =   1680
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item-wise Disc Total"
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
      Index           =   19
      Left            =   45
      TabIndex        =   121
      Top             =   3945
      Width           =   1755
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spares Amount"
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
      Left            =   45
      TabIndex        =   120
      Top             =   4395
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor*"
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
      Left            =   45
      TabIndex        =   119
      Top             =   3330
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Remarks"
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
      Left            =   8085
      TabIndex        =   118
      Top             =   3135
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Delay Reason"
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
      Left            =   3945
      TabIndex        =   117
      Top             =   3120
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Left            =   6915
      TabIndex        =   116
      Top             =   3330
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Comp. Date*"
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
      Left            =   3945
      TabIndex        =   115
      Top             =   3330
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr A/c (Spares) : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   30
      Left            =   30
      TabIndex        =   114
      Top             =   6180
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Service Date"
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
      Left            =   8100
      TabIndex        =   113
      Top             =   3360
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esti. Spares"
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
      Left            =   6420
      TabIndex        =   112
      Top             =   2385
      Width           =   1020
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   24
      Left            =   10740
      TabIndex        =   111
      Top             =   2145
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp.Del.Date"
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
      Left            =   9255
      TabIndex        =   110
      Top             =   1935
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   22
      Left            =   10740
      TabIndex        =   109
      Top             =   1695
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon No."
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
      Left            =   6420
      TabIndex        =   108
      Top             =   1935
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current KMS"
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
      Index           =   20
      Left            =   6420
      TabIndex        =   107
      Top             =   1710
      Width           =   1095
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   16
      Left            =   10740
      TabIndex        =   106
      Top             =   1470
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   14
      Left            =   10740
      TabIndex        =   105
      Top             =   1245
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No."
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
      Left            =   6420
      TabIndex        =   104
      Top             =   1260
      Width           =   1035
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
      Index           =   16
      Left            =   6420
      TabIndex        =   103
      Top             =   1470
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Dt."
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
      Left            =   5940
      TabIndex        =   102
      Top             =   510
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Remarks"
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
      Left            =   6420
      TabIndex        =   101
      Top             =   2610
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic*"
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
      Left            =   45
      TabIndex        =   100
      Top             =   3105
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Dt."
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
      Left            =   3060
      TabIndex        =   99
      Top             =   540
      Width           =   765
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
      Left            =   45
      TabIndex        =   98
      Top             =   1650
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
      Left            =   45
      TabIndex        =   97
      Top             =   2550
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name*"
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
      Left            =   45
      TabIndex        =   96
      Top             =   1425
      Width           =   1215
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
      Left            =   45
      TabIndex        =   95
      Top             =   2325
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total UnClosed Jobs :"
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
      Index           =   7
      Left            =   8220
      TabIndex        =   94
      Top             =   945
      Width           =   1860
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division :"
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
      Left            =   8220
      TabIndex        =   93
      Top             =   480
      Width           =   810
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   9180
      TabIndex        =   92
      Top             =   705
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No."
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
      Left            =   45
      TabIndex        =   90
      Top             =   525
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No.*"
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
      Left            =   3030
      TabIndex        =   89
      Top             =   750
      Width           =   1110
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
      Left            =   4485
      TabIndex        =   88
      Top             =   2535
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
      Left            =   2775
      TabIndex        =   87
      Top             =   2535
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
      Left            =   1155
      TabIndex        =   86
      Top             =   2535
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Srl No"
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
      Left            =   45
      TabIndex        =   85
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Regn. No."
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
      Left            =   45
      TabIndex        =   84
      Top             =   750
      Width           =   840
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   825
      Left            =   8160
      Top             =   390
      Width           =   3540
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code :"
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
      Left            =   9795
      TabIndex        =   83
      Top             =   465
      Width           =   975
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
      Left            =   3030
      TabIndex        =   82
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model*"
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
      Left            =   45
      TabIndex        =   81
      Top             =   975
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No.*"
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
      Left            =   3030
      TabIndex        =   80
      Top             =   975
      Width           =   1020
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
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   10830
      TabIndex        =   79
      Top             =   960
      Width           =   105
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
Dim mMRPLubeTB As Double, mMRPLubeTP  As Double, mLabDiscAmtTB As Single, JobValue As Double
Private Const mSP2 As String = " "
Dim FreeLabForTax As Double
Dim mCmdPostCounter As Integer
Dim mReposting As Boolean
Dim mLabDiscAfterTaxYn As Byte

Dim mVatYn As Byte

Private FirstPrint As Boolean
Dim mCardNo$
Public mVType As String
Dim mPartyType As Byte
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

Private Const Excise_Amt As Byte = 97

Private Const LabDisPer     As Byte = 81          '
Private Const OutSideLabAmt As Byte = 0
Private Const LabAmtTP      As Byte = 82
Private Const LabAmtTB      As Byte = 83
Private Const CouponVal     As Byte = 84
Private Const eCessPer      As Byte = 85
Private Const eCessAmt      As Byte = 86
Private Const FreeWarrLabAmt As Byte = 87
Private Const CreditCardNo  As Byte = 88
Private Const ChqNo         As Byte = 89
Private Const ChqDate       As Byte = 90
Private Const IWDiscPerTB   As Byte = 91
Private Const IWDiscPerTP   As Byte = 92
Private Const SatAmt   As Byte = 93



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
Private Const Col_TaxPer As Byte = 15
Private Const Col_TaxAmt As Byte = 16
Private Const Col_SatPer As Byte = 17
Private Const Col_SatAmt As Byte = 18

Private Const Col_ItemVal As Byte = 19
Private Const Col_Purpose As Byte = 20
Private Const Col_PName As Byte = 21
Private Const Col_LName As Byte = 22
Private Const Col_ClaimNo As Byte = 23
Private Const Col_CompYN As Byte = 24
Private Const Col_PartGrade As Byte = 25

'FGrid1 Columns
Private Const C_LabCode As Byte = 1
Private Const C_LabName As Byte = 2
Private Const C_TaxYN As Byte = 3
Private Const C_PaidBy As Byte = 4
Private Const C_ChrgType As Byte = 5
Private Const C_Hrs As Byte = 6
Private Const C_Rate As Byte = 7
Private Const C_Amt As Byte = 8
Private Const C_External As Byte = 9
Private Const C_GPNo As Byte = 10
Private Const C_Remarks As Byte = 11
Private Const C_ContName As Byte = 12
Private Const C_WIssueDt As Byte = 13
Private Const C_WRecdDt As Byte = 14
Private Const C_ContAmt As Byte = 15
Private Const C_ContAcCode As Byte = 16

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
Private Const DEP_Bill As Byte = 6
Dim mRepName As String
Dim mRepName1 As String
Dim mRepName2 As String
Dim prnGatePass As Boolean


Dim mServiceTaxPer_Saperate As Double
Dim mServiceTaxAmt_Saperate As Double
Dim mECessPer As Double
Dim mECessAmt As Double
Dim mHECessPer As Double
Dim mHECessAmt As Double



Private Sub CmdOk_Click()

Dim I As Integer, mStartdate As String, mEndDate As String
Dim DupMaster As ADODB.Recordset
Dim mCurrAbs As Long
On Error GoTo DispErr

    mStartdate = InputBox("Posting Required from which Date ?", "Start Date for Posting", PubLoginDate)
    mEndDate = InputBox("Posting Required upto which Date ?", "Last Date for Posting", PubLoginDate)
    If mStartdate = "" Or mEndDate = "" Then Exit Sub
    
    mStartdate = MakeDate(mStartdate)
    mEndDate = MakeDate(mEndDate)
    
    mReposting = True
    
'    If Master.RecordCount > 0 Then TopCtrl1_eLast
'    Do Until Master.BOF
'        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) >= CDate(mStartdate) And CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) <= CDate(mEndDate) Then
'            TopCtrl1_eEdit
'            Call Txt_Validate(STaxAmt, False)
'            mCmdPostCounter = 1
'            TopCtrl1_eSave
'            mCmdPostCounter = 0
'        End If
'        TopCtrl1_ePrev
'    Loop
'
'
'    If Master.RecordCount > 0 Then TopCtrl1_eLast
'    Do Until Master.BOF
'        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) >= CDate(mStartdate) And CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) <= CDate(mEndDate) Then
'            TopCtrl1_eEdit
'            Call Txt_Validate(STaxAmt, False)
'            TopCtrl1_eSave
'        End If
'        TopCtrl1_ePrev
'    Loop
    
    
    If Master.RecordCount > 0 Then Master.MoveLast

    Do Until Master.BOF
        Debug.Print Master(0)
        If IsNull(Master!JobCloseDate) Then GoTo MyNextRecord
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) < CDate(mStartdate) Then GoTo MyNextRecord
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) > CDate(mEndDate) Then GoTo MyNextRecord



'        If Master!ClosedU_EntDt < CDate(mStartdate) Then GoTo MyNextRecord
'        If Master!ClosedU_EntDt > CDate(mEndDate) Then GoTo MyNextRecord

        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        lblSparePrefix.Refresh
        LblSprBill.Refresh
        lblLabourBill.Refresh
        lblGatePass.Refresh
        Call TopCtrl1_eEdit
        
        Amt_Cal

        Call Txt_Validate(STaxAmt, False)
        'ProcAcPost rsCtrlAc, rsCtrlAcLab
        mCmdPostCounter = 1
        Call TopCtrl1_eSave
        mCmdPostCounter = 0
MyNextRecord:
        Disp_Text SETS("INI", Me, Master)
        Master.MovePrevious
    Loop

    
    If Master.RecordCount > 0 Then Master.MoveLast
    Do Until Master.BOF
        If IsNull(Master!ClosedU_EntDt) Then GoTo MyNextRecord2
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) < CDate(mStartdate) Then GoTo MyNextRecord2
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) > CDate(mEndDate) Then GoTo MyNextRecord2

        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        lblSparePrefix.Refresh
        LblSprBill.Refresh
        lblLabourBill.Refresh
        lblGatePass.Refresh
        Call TopCtrl1_eEdit
        Amt_Cal
        Call Txt_Validate(STaxAmt, False)
        'ProcAcPost rsCtrlAc, rsCtrlAcLab
        Call TopCtrl1_eSave
MyNextRecord2:
        Disp_Text SETS("INI", Me, Master)
        Master.MovePrevious
    Loop
    
    mReposting = False
    
    Disp_Text SETS("INI", Me, Master)
    MsgBox "Reposting Completed"
    Unload Me
Exit Sub
DispErr:
    MsgBox err.Description



End Sub


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
        If UCase(left(PubComp_Name, 6)) = "RASHMI" Then Optpre.Value = True Else OptPlain.Value = True
        LblPrinter.CAPTION = Printer.DeviceName
        ChkRep(0).Visible = False
        ChkRep(1).Visible = False
'        CmdPrint(0).Visible = False
'        CmdPrint(1).Visible = False
         CmdPrint(0).Visible = True
         CmdPrint(1).Visible = True
    End If
End Sub

Private Sub cmdPost_Click()
Dim I As Integer, mStartdate As String, mEndDate As String
Dim DupMaster As ADODB.Recordset
Dim mCurrAbs As Long
On Error GoTo DispErr

    mStartdate = InputBox("Posting Required from which Date ?", "Start Date for Posting", PubLoginDate)
    mEndDate = InputBox("Posting Required upto which Date ?", "Last Date for Posting", PubLoginDate)
    If mStartdate = "" Or mEndDate = "" Then Exit Sub
    
    mStartdate = MakeDate(mStartdate)
    mEndDate = MakeDate(mEndDate)
    
    mReposting = True
    
'    If Master.RecordCount > 0 Then TopCtrl1_eLast
'    Do Until Master.BOF
'        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) >= CDate(mStartdate) And CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) <= CDate(mEndDate) Then
'            TopCtrl1_eEdit
'            Call Txt_Validate(STaxAmt, False)
'            mCmdPostCounter = 1
'            TopCtrl1_eSave
'            mCmdPostCounter = 0
'        End If
'        TopCtrl1_ePrev
'    Loop
'
'
'    If Master.RecordCount > 0 Then TopCtrl1_eLast
'    Do Until Master.BOF
'        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) >= CDate(mStartdate) And CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) <= CDate(mEndDate) Then
'            TopCtrl1_eEdit
'            Call Txt_Validate(STaxAmt, False)
'            TopCtrl1_eSave
'        End If
'        TopCtrl1_ePrev
'    Loop
    
    
    If Master.RecordCount > 0 Then Master.MoveLast

    Do Until Master.BOF
        Debug.Print Master(0)
        If IsNull(Master!JobCloseDate) Then GoTo MyNextRecord
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) < CDate(mStartdate) Then GoTo MyNextRecord
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) > CDate(mEndDate) Then GoTo MyNextRecord



'        If Master!ClosedU_EntDt < CDate(mStartdate) Then GoTo MyNextRecord
'        If Master!ClosedU_EntDt > CDate(mEndDate) Then GoTo MyNextRecord

        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        lblSparePrefix.Refresh
        LblSprBill.Refresh
        lblLabourBill.Refresh
        lblGatePass.Refresh
        Call TopCtrl1_eEdit
        Call Txt_Validate(STaxAmt, False)
        'ProcAcPost rsCtrlAc, rsCtrlAcLab
        mCmdPostCounter = 1
        Call TopCtrl1_eSave
        mCmdPostCounter = 0
MyNextRecord:
        Disp_Text SETS("INI", Me, Master)
        Master.MovePrevious
    Loop


    If Master.RecordCount > 0 Then Master.MoveLast
    Do Until Master.BOF
        If IsNull(Master!ClosedU_EntDt) Then GoTo MyNextRecord2
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) < CDate(mStartdate) Then GoTo MyNextRecord2
        If CDate(Format(Master!JobCloseDate, "DD/MMM/YYYY")) > CDate(mEndDate) Then GoTo MyNextRecord2

        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        lblSparePrefix.Refresh
        LblSprBill.Refresh
        lblLabourBill.Refresh
        lblGatePass.Refresh
        Call TopCtrl1_eEdit
        Call Txt_Validate(STaxAmt, False)
        'ProcAcPost rsCtrlAc, rsCtrlAcLab
        Call TopCtrl1_eSave
MyNextRecord2:
        Disp_Text SETS("INI", Me, Master)
        Master.MovePrevious
    Loop
    
    mReposting = False
    
    Disp_Text SETS("INI", Me, Master)
    MsgBox "Reposting Completed"
    Unload Me
Exit Sub
DispErr:
    MsgBox err.Description
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DGJob_Click()
If Master.RecordCount > 0 Then
    Call History_Field
End If
DGJob.Visible = False
txt(MyIndex).SetFocus
End Sub
Private Sub DGMech_Click()
If DGMech.Columns(0).CAPTION = "Mechanic Name" Then
    If RsMech.RecordCount > 0 Then
        txt(MyIndex).TEXT = RsMech!Name
        txt(MyIndex).Tag = RsMech!Code
    End If
ElseIf DGMech.Columns(0).CAPTION = "WorkShop Staff" Then
    If RsSuper.RecordCount > 0 Then
        txt(MyIndex).TEXT = RsSuper!Name
        txt(MyIndex).Tag = RsSuper!Code
    End If
End If
DGMech.Visible = False
txt(MyIndex).SetFocus
End Sub
Private Sub DGParty_Click()
If RsParty.RecordCount > 0 Then
    txt(MyIndex).TEXT = RsParty!Name
    txt(MyIndex).Tag = RsParty!Code
End If
DGParty.Visible = False
lblGroup.Visible = False
txt(MyIndex).SetFocus
End Sub
Private Sub DGReason_Click()
If RsReason.RecordCount > 0 Then
    txt(MyIndex).TEXT = RsReason!Name
    txt(MyIndex).Tag = RsReason!Code
End If
DGReason.Visible = False

txt(MyIndex).SetFocus
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
    If mVatYn = 1 Then
        Lbl(22).CAPTION = "V A T  "
        txt(58).Visible = False
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
ElseIf Shift = 4 And KeyCode = vbKeyV Then
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
    txt(JobCDt).Tag = PubLoginDate
    mVatYn = PubVATYN
    mLabDiscAfterTaxYn = VNull(GCn.Execute("Select LabDiscAfterTaxYn From Syctrl").Fields(0).Value)
    
    mVType = "W_SIC"
    If pubTOT_On = 1 Then
        Lbl(25) = "TOT on SubTot(BefTax)"
    End If
    If PubSDTYN = 1 Then
        Lbl(25) = pubTOTCaption
    End If
    If PubReSaleTaxPer = 0 Then
        Lbl(17).Visible = False
        txt(ReSalTaxPer).Visible = False
        txt(ReSalTaxAmt).Visible = False
    End If
        
    If UCase(PubSFADataPath) <> UCase(PubWFADataPath) Then
        SepLabPost = True
    End If
    Call BlankText
    lblSprGrid.Tag = 0
    lblLabGrid.Tag = 0
    
    PubOutSideLabDisc = GCn.Execute("select " & vIsNull("OutSideLabDisc", "0") & " as OutSideLabDisc from Syctrl").Fields(0).Value
    PubSrvTaxOnOutSideLab = GCn.Execute("select " & vIsNull("SrvTaxOnOutSideLab", "0") & " as SrvTaxOnOutSideLab from Syctrl").Fields(0).Value
    
    'Checking Spare a/c Controls
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
    rsCtrlAc.Open "Select SprGenSur_Ac,ReSaleTax_Ac,SprSalTP_Ac,OilSalTB_Ac,OilSalTP_Ac,SprCash_Ac,SprDiscTB_Ac,Transportation_Ac,MiscChrg_Ac,TOTax_Ac,SprROff_Ac ,FSBCrAc From AcControls Where Div_Code='" & PubDivCode & "'", GCnFaS, adOpenStatic, adLockOptimistic
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
    
    
     Dim sitecond As String
     sitecond = " And  JobCloseDate Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("jc.Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If


    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select Jc.DocId AS CODE,ClosedU_EntDt, JobCloseDate " _
                & "from job_card as JC where left(JC.DocId,1)='" & PubDivCode & "' " & sitecond & " and JC.JobCloseDate Is Not Null   Order by JC.JobCloseDate desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 Jc.DocId AS CODE,ClosedU_EntDt, JobCloseDate " _
                & "from job_card as JC where left(JC.DocId,1)='" & PubDivCode & "' " & sitecond & " and JC.JobCloseDate Is Not Null   Order by JC.JobCloseDate desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("jc.Docid", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    Set RsJob = New ADODB.Recordset
    RsJob.CursorLocation = adUseClient
    RsJob.Open "select Jc.DocId AS CODE," & cCStr("JC.Job_No") & " as FindJobNo,JC.Job_No,HC.Model,HC.RegNo,HC.Chassis,HC.Engine,HC.VehSerialNo,HC.Name " & _
                " from (job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) " & _
                " " & _
                " " & _
                " where left(JC.DocID,1)='" & PubDivCode & "'  " & sitecond & "  and JobCloseDate IS NULL order by JC.DocID", GCn, adOpenDynamic, adLockOptimistic
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
        
    GSQL = "select SubGroup.SubCode as code, SubGroup.NAME as Name, Curr_Bal, SubGroup.Add1 ,Party_Type,City.CityName from ((SubGroup " & _
        "left Join City on City.Citycode=SubGroup.Citycode )" & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode )" & _
        "Where " & _
        " SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    RsParty.Sort = "Name"
    
    
    If UCase(left(PubComp_Name, 3)) = "NAC" Then
    LblExcise.Visible = True
    txt(Excise_Amt).Visible = True
    End If
    
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    
    DGParty.Columns(1).Visible = False
    
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


Private Sub Pwinprint_Click()
Dim Index As Integer
Call WindowsPrint(Index)
End Sub

Private Sub WindowsPrint(Index As Integer)
Dim mQry As String, RepTitle$
Dim Condstr$, mDocStr$
Dim RST1 As ADODB.Recordset
Dim Speciality$
Dim Rst As ADODB.Recordset
Dim I As Integer
Dim mQryLab As String

On Error GoTo ERRORHANDLER
      GSQL = "SELECT '1' as Orig,JC.AtKMsHrs,JC.Lab_D_Amt,SPStk.DocID as ReqDocID," & vIsNull("SPStk.Srl_No", "0") & " as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab, " & ConvertDate(date) & "  as v_Date,'' AS Party_Code,' " & txt(OwnerName) & "' as Party_Name,'' as Address,'' as NamePrefix,' " & txt(OwnerName) & "' as Name,' " & txt(Address1) & "' as Add1,' " & txt(Address2) & "' as Add2,' " & txt(Address3) & "' as Add3,' " & txt(City) & "' as CityName,'' as PIN,' " & txt(PhoneResi) & "' as Phone," & _
            "'' as CSTNo,'' as L_C,'" & mFormCode & "' as Form_Code,'' as Remarks, " & Val(txt(MRPAmtTB)) & " as SprAmt_MRP_TB, " & Val(txt(MRPAmtTP)) & " as SprAmt_MRP_TP," & mMRPLubeTB & " as OilAmt_MRP_TB," & mMRPLubeTP & " as OilAmt_MRP_TP, " & Val(txt(SprAmtTB)) & " as SprAmt_TB, " & Val(txt(SprAmtTP)) & " as SprAmt_TP, " & Val(txt(OilAmtTB)) & " as OilAmt_TB,  " & Val(txt(OilAmtTP)) & " as OilAmt_TP, " & Val(txt(DiscPerTB)) & " as D_Per_TB,  " & Val(txt(DiscAmtTB)) & " as D_Amt_TB,  " & Val(txt(DiscPerTP)) & " as D_Per_TP, " & Val(txt(DiscAmtTP)) & " as D_Amt_TP,0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP,0 as D_Amt_MRP_TP, " & Val(txt(Addition)) & " as Addition, " & Val(txt(GenSurPer)) & " as Gen_Sur_Per, " & Val(txt(GenSurAmt)) & " as Gen_Sur_Amt, " & Val(txt(TransAmt)) & " as Trans_Amt, " & Val(txt(STaxPer)) & " as Tax_Per,  " & Val(txt(STaxAmt)) & " as Tax_Amt, 0 as Tax_AmtMRP,  " & Val(txt(TaxSurPer)) & " as Tax_Sur_Per, " & Val(txt(TaxSurAmt)) & " as Tax_Sur_Amt,0 as TaxSur_AmtMRP, " & Val(txt(PackCrg)) & " as Packing,  " & Val(txt(TurnOverPer)) & " as TOT_Per,  " & Val(txt(TurnOverAmt)) & " as Tot_Amt, 0 as TOT_AmtMRP,0 as ReSalTax_Per,0 as ReSalTax_Amt, " & Val(txt(STotB)) & " as Total_Amt," & _
            " " & Val(txt(SROff)) & " as Rounded, " & PubTaxDetOnSprInv & " as Det_Tax,'' as GP_No,'' as GP_Date,1 as Printed_YN, ' " & pubUName & "' as U_Name, ' " & date$ & "' as U_EntDt,0 as CancelYN,0 as LabAmt_TB, 0 as LabAmt_TP, 0 as Lab_TaxPer, 0 as Lab_TaxAmt, 0 as Lab_D_Amt,0 as Lab_RoundOff,0 as NetLab_Amt," & _
            "SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown," & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc, " & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec," & _
            "" & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret," & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN," & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & vIsNull("SPStk.Rate", "0") & " as Rate," & _
            "" & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate," & xIsNull("SPStk.Purpose", "") & " as Purpose,SPStk.Part_SrlNo," & vIsNull("SPStk.Rate2", "0") & " as Rate2," & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate2," & _
            "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2," & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2," & vIsNull("SPStk.Amount", "0") & " as Amount2," & vIsNull("SPStk.Net_Amt", "0") & " as Net_Amt2,'' as Chrg_From,0 as External_YN, " & _
            "Syctrl.WorkShopInvFooter, " & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,SPStk.TaxPer,SPStk.TaxAmt,Jc.Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, SpStk.SatPer, SpStk.SatAmt, " & Val(txt(SatAmt)) & " As SatAmt_H," & cCStr(xIsNull("SpStk.v_No", "")) & " as ReqNo ,SPStk.Disc_Per,SPStk.Disc_Amt " & _
            " FROM (((SP_Stock as SPStk left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
            "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON SPStk.Job_DocID = JC.DocId) " & _
            "LEFT JOIN Syctrl ON Syctrl.LinkTable<>SPStk.U_AE)" & _
            "where SPStk.Job_DocId='" & lblDocId & "'"
        If UCase(left(PubComp_Name, 5)) = "NAWAL" Then
            GSQL = GSQL & " AND " & xIsNull("SPStk.Purpose", "") & " = 'C'"
        End If

    mQryLab = "SELECT '2' as Orig,0 as AtKMsHrs,0 as Lab_D_Amt,' ' as ReqDocID,JL.S_No as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab," & ConvertDate(date) & "  as v_Date,JC.DrLab_AcCode as Party_Code,JC.BillingName as Party_Name,'' as Address,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
                "SG.CSTNo,'' as L_C,'' as Form_Code,'' as Remarks,0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
                "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
                "0 as Rounded," & PubTaxDetOnSprInv & " as Det_Tax,JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted ,JC.U_Name,JC.U_EntDt,0 as CancelYN,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
                "JL.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name,'' as Lub_Category, '' as Godown,0 as Qty_Doc, 0 as Qty_Rec, " & _
                "" & vIsNull("Hrs_Taken", "0") & " as Qty_Iss,0 as Qty_Ret," & vIsNull("JL.Tax_YN", "0") & " as Tax_YN, 0 as MRP_YN,0 as Rate," & _
                "0 as MRP_Rate,'' as Purpose,'' as Part_SrlNo," & vIsNull("JL.Lab_Rate", "0") & " as Rate2,0 as MRP_Rate2," & _
                "0 as Disc_Per2,0 as Disc_Amt2,0 as Amount2," & cIIF("JL.Chrg_From = 'C'", "JL.LabourAmt", "0") & " as Net_Amt2,JL.Chrg_From,JL.External_YN," & _
                "Syctrl.WorkShopInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
                "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass ,0 as TaxPer,0 as TaxAmt,'' AS Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, 0 As SatPer, 0 As SatAmt, 0 As SatAmt_H,' ' as ReqNo ,0 as Disc_Per,0 as Disc_Amt  " & _
                " FROM ((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
                "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
                "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
                "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
                "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
                "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code)" & _
                "Where JC.DocId='" & lblDocId & "'"
    
                 GSQL = GSQL & " Union All " & mQryLab
Set Rst = New Recordset
Rst.CursorLocation = adUseClient
Rst.Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
mRepName = "PBill"

CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")

Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select Div_SName,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    mDocStr = "** Job Provisional Bill" & IIf(RST1!Div_SName = "", "", " (" & RST1!Div_SName & ")") & "**"

For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SubTitle")
            rpt.FormulaFields(I).TEXT = "'" & Speciality & "'"
        Case UCase("Comp_Phone")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPhone & "'"
        Case UCase("Comp_Fax")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecFax & "'"
        Case UCase("RepTitle")
            rpt.FormulaFields(I).TEXT = "'" & mDocStr & "'"
    End Select
Next
           
rpt.Database.SetDataSource Rst
rpt.ReadRecords

Select Case Index
 Case PWindows 'Printer
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
  
        Call Report_View(rpt, "Job Provisional Bill", , True)
   
    End Select
 
Set Rst = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub


Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim RsTemp As ADODB.Recordset

Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Call Fgrid_Ini
    mAddFlag = "A"
    txt(JobCDt).TEXT = txt(JobCDt).Tag 'Format(Date, "dd/MMM/yyyy")
    txt(JobCompTm).TEXT = Format(time, "hh:mm")
    txt(JobNo).SetFocus
    txt(TurnOverPer) = MainLib.TOTCal()
    txt(ServTaxPer) = VNull(GCn.Execute("Select Service_tax from Syctrl").Fields(0).Value)
    txt(eCessPer) = GCn.Execute("Select " & vIsNull("eCessPer", "0") & " From Syctrl").Fields(0)
    
    Set RsTemp = GCn.Execute("Select Service_Tax, ServiceTaxPer_Saperate, ECessPer, HECessPer From Syctrl")
    If RsTemp.RecordCount > 0 Then
        txt(ServTaxPer) = VNull(RsTemp!Service_Tax)
        mServiceTaxPer_Saperate = VNull(RsTemp!ServiceTaxPer_Saperate)
        mECessPer = VNull(RsTemp!eCessPer)
        mHECessPer = VNull(RsTemp!HECessPer)
    End If
    Pwinprint.Visible = True
    
    Set RsTemp = Nothing
    
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
Dim I As Integer
Dim mTrans As Boolean, vBook As Variant
Dim LedgAry(1) As LedgRec, mResult As Byte  ', LedgAryLab(1) As LedgRec

On Error GoTo eloop1
    If IsEditable(RetDate(txt(JobCDt))) = False Then Exit Sub
    
    ApplyConsolidatedPosting CDate(txt(JobCDt))
    
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
        GCnFaS.BeginTrans
        GCnFaW.BeginTrans
        mTrans = True
        CreateLog Me, Master!Code, mReposting
        For I = 1 To FGrid.Rows - 1
            GCn.Execute ("Update SP_Stock Set " _
                        & "Invoice_DocId ='',V_Date2=Null," _
                        & "Rate2=0,MRP_Rate2=0,Disc_Per2=0,Disc_Amt2=0,Amount2=0,Net_Amt2=0 " _
                        & "Where DocID='" & FGrid.TextMatrix(I, Col_ReqNoDocId) & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ReqSrNo)))
        Next
        GCn.Execute "Delete from Sp_Sale where DocID='" & SpareDocID & "'"
        
        If UCase(left(PubComp_Name, 3)) = "JMK" Or RSOJPR = True Then
            GCn.Execute ("update job_Card set LastInvDocId=docId_invspr,LastLabInvDocId=DocId_InvLab,LastInvNoSuff=" & cIIF("len(LastInvNoSuff)=0 or LastInvNoSuff is null", "1", "LastInvNoSuff+1") & ",LastLabInvNoSuff=" & cIIF("len(LastLabInvNoSuff)=0 or LastLabInvNoSuff is null", "1", "LastLabInvNoSuff+1") & " where DocId='" & txt(JobNo).Tag & "'")
            GSQL = "Update Job_Card set JobCloseDate=Null,JobComp_Dt_Time=Null,CrMemo=0,BillingName='',DelBy='',NextSrvDate=Null, " _
                & "docId_invspr='',DocId_InvLab='',gp_no='',DrLab_AcCode='',DrSpr_AcCode='',LabAmt_TB=0,Lab_TaxPer=0,Lab_TaxAmt= 0," _
                & "Lab_D_Amt= 0,Lab_RoundOff= 0,NetLab_Amt= 0,Remark='',DelayReason ='',ClosedU_Name='',ClosedU_EntDt=Null,ClosedU_AE='',LabBillPrinted=0 " _
               & "where DocId='" & txt(JobNo).Tag & "'"
        Else
            GSQL = "Update Job_Card set JobCloseDate=Null,JobComp_Dt_Time=Null,CrMemo=0,BillingName='',DelBy='',NextSrvDate=Null, " _
                & "docId_invspr='',DocId_InvLab='',gp_no='',DrLab_AcCode='',DrSpr_AcCode='',LabAmt_TB=0,Lab_TaxPer=0,Lab_TaxAmt= 0," _
                & "Lab_D_Amt= 0,Lab_RoundOff= 0,NetLab_Amt= 0,Remark='',DelayReason ='',ClosedU_Name='',ClosedU_EntDt=Null,ClosedU_AE='',LabBillPrinted=0 " _
               & "where DocId='" & txt(JobNo).Tag & "'"
            
        End If
        GCn.Execute GSQL
        
'        GCn.Execute ("Insert into Deletelog Values('" & SpareDocID & "',1," & Val(Txt(NetAmt)) & ",'" & pubUName & "'," & ConvertDate(date$) & ",'" & Time$ & "')")
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, lblDocId)
        If mResult <> 1 Then MsgBox "Error in Ledger Un-Posting", vbOKOnly, "Validation"

        'Unpost Ledger a/c
        If txt(CashBill).TEXT = "Yes" And IsConsolidatedPosting Then
            ProcAcPost rsCtrlAc, rsCtrlAcLab
            If UCase(left(PubComp_Name, 6)) = "TRUPTI" Then
                mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, lblDocId)
                If mResult <> 1 Then MsgBox "Error in Ledger Un-Posting", vbOKOnly, "Validation"
            End If
        Else
            'to avoid errors of Old System
            LedgerUnPost GCnFaS, txt(JobNo).Tag
            If SepLabPost Then
                LedgerUnPost GCnFaW, txt(JobNo).Tag
            End If
            'eof of Old System
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, SpareDocID)
            If mResult <> 1 Then MsgBox "Error in Ledger Un-Posting", vbOKOnly, "Validation"
            If SepLabPost Then
                mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaW, LabourDocID)
                If mResult <> 1 Then MsgBox "Error in Ledger Labour Un-Posting", vbOKOnly, "Validation"
            End If
           ' If UCase(left(PubComp_Name, 6)) = "TRUPTI" Then
                mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, lblDocId)
                If mResult <> 1 Then MsgBox "Error in Ledger Un-Posting", vbOKOnly, "Validation"
           ' End If
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
   '  Pwinprint.Visible = True
    If IsEditable(RetDate(txt(JobCDt))) = False Then Exit Sub

    Call Fgrid_Ini
    If Master.EOF = True Or Master.BOF = True Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    mAddFlag = "E"
    For I = 1 To 30
        txt(I).Enabled = False
    Next I
    If pubUName = "SA" And UCase(left(PubComp_Name, 4)) = "ENAR" Then
        txt(JobCDt).Enabled = True
    End If
    txt(CashBill).Enabled = False
    If txt(CashBill).TEXT = "Yes" Then
        txt(SpareParty).Enabled = False
        txt(LabourParty).Enabled = False
        txt(CashParty).Enabled = True
    Else
        txt(SpareParty).Enabled = True
        txt(LabourParty).Enabled = True
        txt(CashParty).Enabled = False
    End If
    Call txtDisabled_Color
    txt(MechName).SetFocus
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
    Dim sitecond As String
    sitecond = " And  JobCloseDate Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("jc.Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "select Jc.DocId AS SearchCode,JC.Job_No," & cTrim("right(JC.DocId_InvSpr,8)") & " as Inv_No," & cTrim(cMID("Jc.DocId", "9", "5")) & " as Prefix,JC.Site_Code,JC.Govt_YN, JC.Job_Date, JC.JobCloseDate,HC.Model,HC.RegNo, HC.Chassis, HC.Engine,HC.VehSerialNo,HC.Name,HC.Add1, HC.Add2, HC.Add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName,jc.OpenRemarks,jc.Body_Damage,Jc.Job_BookNo,Jc.Job_InspDocID,jc.Coupon,jc.ExpDelDate,Jc.DelBy,Jc.RecBy_Supervisor,Jc.DelayReason,Jc.JobComp_Dt_Time,JC.Remark,EM.EMP_NAME AS Mechanic,EMP.Emp_Name as Supervisor," _
                & "jc.CRMemo,jc.BillingName,jc.NetLab_Amt,JC.DocId_InvSpr,Jc.DocId_InvLab,JC.GP_No " _
                & "from (((((job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) left Join Service_Type as ST on JC.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode) left join Emp_Mast as EM on JC.Delby=EM.Emp_Code) left join Emp_Mast as EMP on Jc.RecBy_Supervisor=Emp.Emp_Code) Left Join Job_Delay as JD on JC.DelayReason=JD.Code where left(JC.DocId,1)='" & PubDivCode & "' " & sitecond & " and JC.JobCloseDate Is Not Null order by JC.docID"
    Set SearchForm = Me
    FIND2.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("Code='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("select Jc.DocId AS CODE, ClosedU_EntDt, JobCloseDate " _
                & "from job_card as JC where left(JC.DocId,1)='" & PubDivCode & "' and JC.JobCloseDate Is Not Null And JC.DocId = '" & MyValue & "' Order by JC.JobCloseDate desc")
    End If
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
      Pwinprint.Visible = False
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
If UCase(left(PubComp_Name, 6)) = "RASHMI" Then Optpre.Value = True Else OptPlain.Value = True
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
Dim mTrans As Boolean, MyGPNo$, RsTemp As ADODB.Recordset
'A/c Posting related declarations
Dim rsForm As ADODB.Recordset, mLabAmtTB As Double, mLabAmtTP As Double
Dim mTotSprAmt As Double, mTotOilAmt As Double
Dim mDaySprCash As Double, mDayLabCash As Double, DivBaseNumber As Boolean
'On Error GoTo errlbl
Dim mCurrBal As Double, mEditValue As Double, mCrLimit As Double, SrvCat As String


    If TopCtrl1.TopText2 = "Add" Then
        If StrCmp(left(PubComp_Name, 7), "Vandana") Then
            If MsgBox("Is Servicing" + vbCrLf + "W.A." + vbCrLf + "W.B." + vbCrLf + "Tyre Rotation" + vbCrLf + "Battery Charging" + vbCrLf + "Billed ?", vbYesNo) = vbNo Then Exit Sub
        End If
    End If


    If IsEditable(RetDate(txt(JobCDt))) = False Then Exit Sub
    
    Grid_Hide
    
    Amt_Cal
    
    'Check Mechanic in Labour
    GSQL = "Select S_No from Job_Lab where Job_DocID='" & txt(JobNo).Tag & "' and Job_Lab.S_No not in(Select S_No from Job_Lab2 where Job_DocID='" & txt(JobNo).Tag & "')"
    If GCn.Execute(GSQL).RecordCount > 0 Then
        MsgBox "Please Enter Mechanic in Labour Done Entry", vbCritical, "Mechanic Name"
        Exit Sub
    End If
    If Val(txt(CouponVal)) <> 0 Then
        GSQL = "Select count(Lab_Code) from Job_Lab where LabourAmt=" & Val(txt(CouponVal)) & " and  Chrg_Type='F'"
        If GCn.Execute(GSQL).RecordCount <= 0 Then
            MsgBox "Please Enter Free Service Labour in Labour Done Entry", vbCritical, "Free Service Labour"
            Exit Sub
        End If
    End If
    
    
     If VNull(G_FaCn.Execute("Select FSBOnlinePost from AcControls where Div_Code='" & PubDivCode & "'").Fields(0).Value) = 1 Then
            SrvCat = GCn.Execute("Select Serv_Catg from Service_Type Left Join Job_Card on Service_type.Serv_type=Job_Card.Serv_Type where Job_card.Job_No=" & txt(JobNo) & "").Fields(0).Value
            If (SrvCat = "F") Then
                Set RsTemp = New ADODB.Recordset
                RsTemp.CursorLocation = adUseClient
                RsTemp.Open "Select AMd_Dealer.D_Name from (Amd_Dealer Left Join HisCard on HisCard.Dealer_code=Amd_Dealer.D_Code) Left Join Job_card on job_Card.CardNO=HisCard.CardNo  where Job_Card.Job_No =" & txt(JobNo) & "", GCn, adOpenStatic, adLockReadOnly
                If RsTemp.RecordCount = 0 Then
                    MsgBox " Dealer Name is Empty.FSB Posting can not be made."
                    Exit Sub
                End If
            End If
     End If
     
    'eof
    If IsValid(txt(JobNo), "Job Card No.") = False Then Exit Sub
    If IsValid(txt(JobCompDt), "Job Completion Date") = False Then Exit Sub
    
    If GCn.Execute("Select Count(*) From Job_Card Where SiebelDocId Is Not Null And DocId='" & txt(JobNo).Tag & "'").Fields(0).Value = 0 And mReposting = False Then
        If txt(JobCompTm) = "" Or txt(JobCompTm) = "00:00" Then MsgBox "Job Completion Time is required", vbOKOnly, "Validation": txt(JobCompTm).SetFocus: Exit Sub
        If IsValid(txt(MechName), "Mechanic name") = False Then Exit Sub
        If IsValid(txt(SuperName), "Supervisor name") = False Then Exit Sub
        If txt(JobDelay).Enabled = True Then
            If IsValid(txt(JobDelay), "Reason for Job Delay") = False Then Exit Sub
        End If
        If IsValid(txt(NextSrv), "Next Service Date") = False Then Exit Sub
        If txt(NextSrv) <> "" Then
            If CDate(txt(JobCDt)) > CDate(txt(NextSrv)) Then
                MsgBox "Next Service Date is less than Job close Date", vbOKOnly, "Validation"
                txt(NextSrv).SetFocus: Exit Sub
            End If
        End If
        If txt(CashBill).TEXT = "Yes" Then
            If IsValid(txt(CashParty), "Cash Party Name") = False Then Exit Sub
        End If
    End If
    If txt(CashBill).TEXT = "Yes" Then
        txt(SpareParty).Tag = PubSprCashAc  ' mSprTempAc
        txt(LabourParty).Tag = PubSrvCashAc ' mSrvTempAc
    Else
        If IsValid(txt(SpareParty), "Debit A/c Spare Party Name") = False Then Exit Sub
        If IsValid(txt(LabourParty), "Debit A/c Labour Party Name") = False Then Exit Sub
    End If
'    Check Cr Limit for Challans
    If PubCrLimitCheck = 1 And txt(CashBill) <> "Yes" Then
        mCurrBal = 0
        mEditValue = 0
        mCurrBal = VNull(G_FaCn.Execute("Select Sum(AmtDr)-Sum(AmtCr) from Ledger where SubCode='" & txt(SpareParty).Tag & "'").Fields(0).Value)
        mCrLimit = VNull(GCn.Execute("Select CreditLimit from SubGroup where SubCode='" & txt(SpareParty).Tag & "'").Fields(0).Value)
        If mAddFlag <> "A" Then
            mEditValue = VNull(GCn.Execute("Select Total_Amt from SP_Sale S Where S.DocID = '" & SpareDocID & "'").Fields(0).Value)
            mEditValue = mEditValue + VNull(GCn.Execute("Select NetLab_Amt from Job_Card J Where J.DocID = '" & txt(JobNo).Tag & "'").Fields(0).Value)
        End If
        mCurrBal = (mCurrBal - mEditValue) + Val(txt(NetAmt))
        If mCurrBal > 0 Then     'Dr Balance
            If mCurrBal > mCrLimit And mCrLimit > 0 Then
                MsgBox "Cr Limit Rs." & mCrLimit & " Exceeds by Rs." & mCurrBal - mCrLimit & vbCrLf & "Add/Edit Denied !", vbInformation, "Cr Limit Checking"
                Me.ActiveControl.SetFocus: Exit Sub
            End If
        End If
    End If
 '   EOF Cr Limit Checking
    
    'Check If Job Closed by another User
    If mAddFlag = "A" Then
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select JobCloseDate,ClosedU_Name,ClosedU_EntDt from Job_Card where Job_Card.DocId='" & txt(JobNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        If Not IsNull(Rst!JobCloseDate) Then 'Job Closed
            MsgBox "Job Already Closed by User " & Rst!ClosedU_Name & " Dt." & Rst!ClosedU_EntDt
            GoTo lblExit
        End If
    End If
    
    'Add records
    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    mLabAmtTB = Val(txt(LabAmtTB))
    mLabAmtTP = Val(txt(LabAmtTP))
    If mVatYn = 1 Then
        mMRPTax = 0
    End If
    
    'If StrCmp(left(PubComp_Name, 3), "Jmk") Then
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" Then
                GCn.Execute "Update Sp_Stock Set Rate = " & Val(FGrid.TextMatrix(I, Col_Rate)) & ",Mrp_Rate = " & Val(FGrid.TextMatrix(I, Col_MRPRate)) & ", Amount = " & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Disc_Per=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ", Disc_Amt=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & ", " & _
                            "Net_Amt=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ", TaxPer=" & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ", TaxAmt=" & Val(FGrid.TextMatrix(I, Col_TaxAmt)) & ", SatPer=" & Val(FGrid.TextMatrix(I, Col_SatPer)) & ", SatAmt=" & Val(FGrid.TextMatrix(I, Col_SatAmt)) & ", " & _
                            "Rate2=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ",MRP_Rate2=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & _
                            "Disc_Per2=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & _
                            "Amount2=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Net_Amt2=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ",Excise_Amt=" & Val(txt(Excise_Amt)) & " " & _
                            "Where Sp_Stock.DocId = '" & FGrid.TextMatrix(I, Col_ReqNoDocId) & "' And Srl_No = " & Val(FGrid.TextMatrix(I, Col_ReqSrNo)) & " "
            End If
        Next I
        
        
        For I = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(I, C_LabCode) <> "" Then
                GCn.Execute "Update Job_Lab Set Lab_Rate = " & Val(FGrid1.TextMatrix(I, C_Rate)) & ", LabourAmt = " & Val(FGrid1.TextMatrix(I, C_Amt)) & " Where Job_DocId = '" & txt(JobNo).Tag & "' And Lab_Code = '" & FGrid1.TextMatrix(I, C_LabCode) & "' "
            End If
        Next I
    'End If
    
    
    If mAddFlag = "A" Then
        'Creating Bill Numbers
        '' Note: Manual Numbring System for Spares/Labour Bill/Gate pass is Not maintained
        '**********
        Set Rst = G_FaCn.Execute("Select DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & SpareVtype & "'")
        
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
        Set Rst = G_FaCn.Execute(GSQL & " Order By VP.Div_Code,VP.Date_From DESC")
        If Val(Rst!start_srl_no) >= Val(LblSprBill.CAPTION) And SpareDocID = "" Then
            SpareDocID = GetDocID(G_FaCn, SpareVtype, txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
        End If
        '*************For cancel Bill management ***************************
        Dim tmpVal, tmpVal1 As String
        
        tmpVal = XNull(GCn.Execute("Select LastInvDocId From job_card where docid='" & txt(JobNo).Tag & "'").Fields(0).Value)
        If tmpVal = "" Then
            If Rst.RecordCount > 0 Then
                GSQL = "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No + 1 Where V_Type='" & Rst!V_Type & "' "
                If DivBaseNumber Then
                    GSQL = GSQL & " and Div_Code ='" & PubDivCode & "'"
                End If
                GSQL = GSQL & " and Date_From=" & ConvertDate(Format(Rst!Date_From, "dd/MMM/yyyy")) & ""
                GCnFaW.BeginTrans
                    GCnFaW.Execute GSQL
                GCnFaW.CommitTrans
            End If
        End If
        '*********************************************************************
        '' for Labour Bill Duplicate Check
        GSQL = "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No " & _
            "From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type " & _
            "Where VP.V_Type='" & LabourVtype & "'"
        If DivBaseNumber Then
            GSQL = GSQL & " and VP.Div_Code='" & PubDivCode & "'"
        End If
        Set Rst = G_FaCn.Execute(GSQL & " Order By VP.Date_From DESC")
        If Val(Rst!start_srl_no) >= Val(LblSprBill.CAPTION) And LabourDocID = "" Then
            LabourDocID = GetDocID(G_FaCn, SpareVtype, txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
        End If
        tmpVal1 = XNull(GCn.Execute("Select LastLabInvDocId from job_card where docid='" & txt(JobNo).Tag & "'").Fields(0).Value)
        If tmpVal1 = "" Then
            If Rst.RecordCount > 0 Then
                GSQL = "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_Type & "' "
                If DivBaseNumber Then
                    GSQL = GSQL & " and Div_Code ='" & PubDivCode & "'"
                End If
                GSQL = GSQL & " and Date_From=" & ConvertDate(Format(Rst!Date_From, "dd/MMM/yyyy")) & ""
                GCnFaW.BeginTrans
                    GCnFaW.Execute GSQL
                GCnFaW.CommitTrans
            End If
        End If
        MyGPNo = "00000" & GCn.Execute("select " & vIsNull("Max(" & cVal("right(gp_no,5)") & ")", "0") & "+1 from job_card where left(gp_no,1)='" & PubDivCode & "' AND " & cMID("gp_no", "2", "1") & "='" & PubSiteCode & "'").Fields(0).Value
        MyGPNo = PubDivCode & PubSiteCode & ForSiteCode & Right(MyGPNo, 5)
        lblGatePass = MyGPNo
        '' Job_card Table
        If PubBackEnd = "A" Then
            GSQL = "Update Job_Card set JobCloseDate=" & ConvertDate(txt(JobCDt)) & ",JobComp_Dt_Time=#" & Format(txt(JobCompDt) & " " & txt(JobCompTm), "dd/MMM/yyyy hh:mm") & _
                "#,CrMemo=" & IIf(txt(CashBill) = "Yes", 0, 1) & ",BillingName='" & txt(CashParty) & "',DelBy='" & txt(MechName).Tag & "',RecBy_Supervisor='" & txt(SuperName).Tag & "',NextSrvDate=" & ConvertDate(txt(NextSrv)) & _
                ",DocId_InvSpr='" & SpareDocID & "',DocId_InvLab='" & LabourDocID & "',GP_NO='" & MyGPNo & "',DrSpr_AcCode='" & txt(SpareParty).Tag & "',DrLab_AcCode='" & txt(LabourParty).Tag & _
                "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & Val(txt(LabDisc)) & ",LabD_pER= " & Val(txt(LabDisPer)) & ",Lab_TaxPer=" & Val(txt(ServTaxPer)) & ",Lab_TaxAmt= " & Val(txt(ServTaxAmt)) & _
                ",Lab_RoundOff= " & Val(txt(LabROff)) & ",NetLab_Amt= " & Val(txt(NetLabAmt)) & ",Remark='" & txt(CloseRemark) & "',DelayReason ='" & txt(JobDelay).Tag & _
                "',ClosedU_Name='" & pubUName & "',ClosedU_EntDt=" & ConvertDate(PubServerDate) & ",ClosedU_AE='" & left(TopCtrl1.TopText2, 1) & _
                "',Closed_AddBy = '" & pubUName & "', Closed_AddDate = " & ConvertDateTime(PubServerDate) & ",LabAmt_Out=" & Val(txt(OutSideLabAmt)) & ", ServiceTaxPer_Saperate=" & mServiceTaxPer_Saperate & ", ServiceTaxAmt_Saperate=" & mServiceTaxAmt_Saperate & ", eCessPer=" & mECessPer & ", eCessAmt=" & mECessAmt & ", " & _
                "HeCessPer=" & mHECessPer & ", HeCessAmt=" & mHECessAmt & ", CreditCardNo='" & txt(CreditCardNo) & "', ChqNo='" & txt(ChqNo) & "', ChqDate=" & ConvertDate(txt(ChqDate)) & ", Variation_Spare = 0, Variation_Labour = 0, FreeWarrLabAmt=" & Val(txt(FreeWarrLabAmt)) & " " & _
                "Where Job_Card.DocId='" & txt(JobNo).Tag & "'"
        ElseIf PubBackEnd = "S" Then
            GSQL = "Update Job_Card set JobCloseDate=" & ConvertDate(txt(JobCDt)) & ",JobComp_Dt_Time='" & Format(txt(JobCompDt) & " " & txt(JobCompTm), "dd/MMM/yyyy hh:mm") & _
                "',CrMemo=" & IIf(txt(CashBill) = "Yes", 0, 1) & ",BillingName='" & txt(CashParty) & "',DelBy='" & txt(MechName).Tag & "',RecBy_Supervisor='" & txt(SuperName).Tag & "',NextSrvDate=" & ConvertDate(txt(NextSrv)) & _
                ",DocId_InvSpr='" & SpareDocID & "',DocId_InvLab='" & LabourDocID & "',GP_NO='" & MyGPNo & "',DrSpr_AcCode='" & txt(SpareParty).Tag & "',DrLab_AcCode='" & txt(LabourParty).Tag & _
                "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & Val(txt(LabDisc)) & ",LabD_pER= " & Val(txt(LabDisPer)) & ",Lab_TaxPer=" & Val(txt(ServTaxPer)) & ",Lab_TaxAmt= " & Val(txt(ServTaxAmt)) & _
                ",Lab_RoundOff= " & Val(txt(LabROff)) & ",NetLab_Amt= " & Val(txt(NetLabAmt)) & ",Remark='" & txt(CloseRemark) & "',DelayReason ='" & txt(JobDelay).Tag & _
                "',ClosedU_Name='" & pubUName & "',ClosedU_EntDt=" & ConvertDate(PubServerDate) & ",ClosedU_AE='" & left(TopCtrl1.TopText2, 1) & _
                "',Closed_AddBy = '" & pubUName & "', Closed_AddDate = " & ConvertDateTime(PubServerDate) & ",LabAmt_Out=" & Val(txt(OutSideLabAmt)) & ", ServiceTaxPer_Saperate=" & mServiceTaxPer_Saperate & ", ServiceTaxAmt_Saperate=" & mServiceTaxAmt_Saperate & ", eCessPer=" & mECessPer & ", eCessAmt=" & mECessAmt & ", Variation_Spare = 0, Variation_Labour = 0, " & _
                "HeCessPer=" & mHECessPer & ", HeCessAmt=" & mHECessAmt & ", CreditCardNo='" & txt(CreditCardNo) & "', ChqNo='" & txt(ChqNo) & "', ChqDate=" & ConvertDate(txt(ChqDate)) & ",FreeWarrLabAmt=" & Val(txt(FreeWarrLabAmt)) & "   " & _
                "where Job_Card.DocId='" & txt(JobNo).Tag & "'"
        End If
        GCn.Execute GSQL
        
        '' SP_Sale Table
        ' Pending Fields - > LineFileTaxSum,GP_No,GP_Date
        GSQL = "Delete from sp_sale where docID = '" & SpareDocID & "'"
        GCn.Execute (GSQL)
        GSQL = "Insert Into SP_Sale(" _
            & "DocID ,DocIDHelp ,V_Type ,V_No ,Site_Code ," _
            & "V_Date,Cash_Credit ,Party_Code ,Party_Name ,Job_DocId," _
            & "L_C,Form_Code,CrAc,AcPosting_Yn,Det_Tax,SprAmt_MRP_TB ," _
            & "SprAmt_MRP_TP,OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB,SprAmt_TP ,OilAmt_TB ,OilAmt_TP ," _
            & "D_Per_TB ,D_Amt_TB ,D_Per_TP ,D_Amt_TP ,Addition ," _
            & "Packing ,Gen_Sur_Per ,Gen_Sur_Amt ,Trans_Amt ,Tax_Per ," _
            & "Tax_Amt ,Tax_Sur_Per ,Tax_Sur_Amt, SatAmt,TOT_Per ,TOT_Amt ," _
            & "ReSalTax_Per, ReSalTax_Amt,Rounded ,Total_Amt, U_Name,U_EntDt,U_AE, AddBy, AddDate,D_Per_MRP_TB, " _
            & "D_Amt_MRP_TB, D_Per_MRP_TP, D_Amt_MRP_TP, Tax_AmtMRP, TaxSur_AmtMRP, TOT_AmtMRP, GP_NO, GP_DATE,Excise_Amt) " _
            & "Values(" _
            & "'" & SpareDocID & "','" & SpareDocID & "','" & SpareVtype & "'," & Val(LblSprBill.CAPTION) & ",'" & PubSiteCode & PubSiteCode & _
            "', " & ConvertDate(txt(JobCDt)) & ",'" & IIf(txt(CashBill) = "Yes", "Cash", "Credit") & "','" & txt(SpareParty).Tag & "','" & IIf(txt(CashBill) = "Yes", txt(CashParty), txt(SpareParty)) & "','" & txt(JobNo).Tag & _
            "', 'L','" & mFormCode & "','CrAc',1,'" & PubTaxDetOnSprInv & "'," & Val(txt(MRPAmtTB)) - mMRPLubeTB & _
            " ," & Val(txt(MRPAmtTP)) - mMRPLubeTP & "," & mMRPLubeTB & "," & mMRPLubeTP & "," & Val(txt(SprAmtTB)) & "," & Val(txt(SprAmtTP)) & "," & Val(txt(OilAmtTB)) & "," & Val(txt(OilAmtTP)) & _
            " ," & Val(txt(DiscPerTB)) & "," & Val(txt(DiscAmtTB)) & "," & Val(txt(DiscPerTP)) & "," & Val(txt(DiscAmtTP)) & "," & Val(txt(Addition)) & _
            " ," & Val(txt(PackCrg)) & "," & Val(txt(GenSurPer)) & "," & Val(txt(GenSurAmt)) & "," & Val(txt(TransAmt)) & "," & Val(txt(STaxPer)) & _
            " ," & Val(txt(STaxAmt)) & "," & Val(txt(TaxSurPer)) & "," & Val(txt(TaxSurAmt)) & ", " & Val(txt(SatAmt)) & "," & Val(txt(TurnOverPer)) & "," & Val(txt(TurnOverAmt)) & _
            " ," & Val(txt(ReSalTaxPer)) & "," & Val(txt(ReSalTaxAmt)) & "," & Val(txt(SROff)) & "," & Val(txt(NetSprAmt)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & "," & mMRevDisTBPer & _
            " , " & mTBDisAmtMRP & "," & mMRevDisTPPer & "," & mTPDisAmtMRP & "," & mMRPTax & "," & mMRPTaxSur & ", " & mMRPTOT & ",'" & MyGPNo & "'," & ConvertDate(Format(txt(JobCDt), "dd/MMM/yyyy")) & "," & Val(txt(Excise_Amt)) & " )"
        GCn.Execute GSQL
        '' Sp_Stock Updation
        For I = 1 To FGrid.Rows - 1
            GCn.Execute ("Update SP_Stock Set " _
                & "Invoice_DocId ='" & SpareDocID & "',V_Date2=" & ConvertDate(txt(JobCDt).TEXT) & "," _
                & "Rate2=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ",MRP_Rate2=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," _
                & "Disc_Per2=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," _
                & "Amount2=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Net_Amt2=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & "," _
                & "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " _
                & "Where DocID='" & FGrid.TextMatrix(I, Col_ReqNoDocId) & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ReqSrNo)))
        Next
        
    ElseIf mAddFlag = "E" Then
        '' Job_card Table
        CreateLog Me, Master!Code, mReposting
        If PubBackEnd = "A" Then
            GSQL = "Update Job_Card set JobCloseDate=" & ConvertDate(txt(JobCDt)) & ",JobComp_Dt_Time=#" & Format(txt(JobCompDt) & " " & txt(JobCompTm), "dd/MMM/yyyy hh:mm") & _
                "#,CrMemo=" & IIf(txt(CashBill) = "Yes", 0, 1) & ",BillingName='" & txt(CashParty) & "',DelBy='" & txt(MechName).Tag & "',RecBy_Supervisor='" & txt(SuperName).Tag & "',NextSrvDate=" & ConvertDate(txt(NextSrv)) & _
                ", DrSpr_AcCode='" & txt(SpareParty).Tag & "',DrLab_AcCode='" & txt(LabourParty).Tag & _
                "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & Val(txt(LabDisc)) & ",LabD_Per= " & Val(txt(LabDisPer)) & ",Lab_TaxPer=" & Val(txt(ServTaxPer)) & ",Lab_TaxAmt= " & Val(txt(ServTaxAmt)) & _
                ",Lab_RoundOff= " & Val(txt(LabROff)) & ",NetLab_Amt= " & Val(txt(NetLabAmt)) & ",Remark='" & txt(CloseRemark) & "',DelayReason ='" & txt(JobDelay).Tag & _
                "',ClosedU_Name='" & pubUName & "',ClosedU_EntDt=" & ConvertDate(PubServerDate) & ", Variation_Spare = 0, Variation_Labour = 0,ClosedU_AE='" & left(TopCtrl1.TopText2, 1) & _
                "',Closed_ModifyBy='" & pubUName & "',Closed_ModifyDate=" & ConvertDateTime(PubServerDate) & ",LabAmt_Out=" & Val(txt(OutSideLabAmt)) & ", ServiceTaxPer_Saperate=" & mServiceTaxPer_Saperate & ", ServiceTaxAmt_Saperate=" & mServiceTaxAmt_Saperate & ", eCessPer=" & mECessPer & ", eCessAmt=" & mECessAmt & ", HEcessPer = " & mHECessPer & ", HECessAmt = " & mHECessAmt & ", CreditCardNo='" & txt(CreditCardNo) & "', ChqNo='" & txt(ChqNo) & "', ChqDate=" & ConvertDate(txt(ChqDate)) & ", FreeWarrLabAmt= " & Val(txt(FreeWarrLabAmt)) & "    where Job_Card.DocId='" & txt(JobNo).Tag & "'"
        ElseIf PubBackEnd = "S" Then
            GSQL = "Update Job_Card set JobCloseDate=" & ConvertDate(txt(JobCDt)) & ",JobComp_Dt_Time='" & Format(txt(JobCompDt) & " " & txt(JobCompTm), "dd/MMM/yyyy hh:mm") & _
                "',CrMemo=" & IIf(txt(CashBill) = "Yes", 0, 1) & ",BillingName='" & txt(CashParty) & "',DelBy='" & txt(MechName).Tag & "',RecBy_Supervisor='" & txt(SuperName).Tag & "',NextSrvDate=" & ConvertDate(txt(NextSrv)) & _
                ", DrSpr_AcCode='" & txt(SpareParty).Tag & "',DrLab_AcCode='" & txt(LabourParty).Tag & _
                "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & Val(txt(LabDisc)) & ",LabD_Per= " & Val(txt(LabDisPer)) & ",Lab_TaxPer=" & Val(txt(ServTaxPer)) & ",Lab_TaxAmt= " & Val(txt(ServTaxAmt)) & _
                ",Lab_RoundOff= " & Val(txt(LabROff)) & ",NetLab_Amt= " & Val(txt(NetLabAmt)) & ",Remark='" & txt(CloseRemark) & "',DelayReason ='" & txt(JobDelay).Tag & _
                "',ClosedU_Name='" & pubUName & "',ClosedU_EntDt=" & ConvertDate(PubServerDate) & ", Variation_Spare = 0, Variation_Labour = 0,ClosedU_AE='" & left(TopCtrl1.TopText2, 1) & _
                "',Closed_ModifyBy='" & pubUName & "',Closed_ModifyDate=" & ConvertDateTime(PubServerDate) & ",LabAmt_Out=" & Val(txt(OutSideLabAmt)) & ", ServiceTaxPer_Saperate=" & mServiceTaxPer_Saperate & ", ServiceTaxAmt_Saperate=" & mServiceTaxAmt_Saperate & ", eCessPer=" & mECessPer & ", eCessAmt=" & mECessAmt & ", HEcessPer = " & mHECessPer & ", HECessAmt = " & mHECessAmt & ", CreditCardNo='" & txt(CreditCardNo) & "', ChqNo='" & txt(ChqNo) & "', ChqDate=" & ConvertDate(txt(ChqDate)) & ", FreeWarrLabAmt=" & Val(txt(FreeWarrLabAmt)) & "    where Job_Card.DocId='" & txt(JobNo).Tag & "'"
        End If
            GCn.Execute GSQL
        
        GSQL = "update SP_Sale set V_Date=" & ConvertDate(txt(JobCDt)) & ",Det_Tax='" & PubTaxDetOnSprInv & "', Form_Code='" & mFormCode & "', Party_Code='" & txt(SpareParty).Tag & "',Party_Name='" & IIf(txt(CashBill) = "Yes", txt(CashParty), txt(SpareParty)) & _
            "', SprAmt_MRP_TB=" & Val(txt(MRPAmtTB)) - mMRPLubeTB & ",SprAmt_MRP_TP=" & Val(txt(MRPAmtTP)) - mMRPLubeTP & _
            " ,OilAmt_MRP_TB=" & mMRPLubeTB & ",OilAmt_MRP_TP=" & mMRPLubeTP & ",SprAmt_TB=" & Val(txt(SprAmtTB)) & ",SprAmt_TP=" & Val(txt(SprAmtTP)) & ",OilAmt_TB=" & Val(txt(OilAmtTB)) & ",OilAmt_TP=" & Val(txt(OilAmtTP)) & _
            " ,D_Per_TB=" & Val(txt(DiscPerTB)) & ",D_Amt_TB=" & Val(txt(DiscAmtTB)) & ",Excise_Amt=" & Val(txt(Excise_Amt)) & " ,D_Per_TP=" & Val(txt(DiscPerTP)) & ",D_Amt_TP=" & Val(txt(DiscAmtTP)) & ",Addition=" & Val(txt(Addition)) & _
            " ,Packing=" & Val(txt(PackCrg)) & ", Gen_Sur_Per=" & Val(txt(GenSurPer)) & ",Gen_Sur_Amt=" & Val(txt(GenSurAmt)) & ",Trans_Amt=" & Val(txt(TransAmt)) & ",Tax_Per=" & Val(txt(STaxPer)) & _
            " ,Tax_Amt=" & Val(txt(STaxAmt)) & ", Tax_Sur_Per=" & Val(txt(TaxSurPer)) & ", SatAmt = " & Val(txt(SatAmt)) & ",Tax_Sur_Amt=" & Val(txt(TaxSurAmt)) & ",TOT_Per=" & Val(txt(TurnOverPer)) & ",TOT_Amt=" & Val(txt(TurnOverAmt)) & _
            " ,ReSalTax_Per=" & Val(txt(ReSalTaxPer)) & ", ReSalTax_Amt=" & Val(txt(ReSalTaxAmt)) & ",Rounded=" & Val(txt(SROff)) & _
            " ,Total_Amt=" & Val(txt(NetSprAmt)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & _
            " , ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ",D_Per_MRP_TB=" & mMRevDisTBPer & ",D_Amt_MRP_TB=" & mTBDisAmtMRP & ", D_Per_MRP_TP =" & mMRevDisTPPer & ", D_Amt_MRP_TP=" & mTPDisAmtMRP & _
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
                & "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " _
                & "Where DocID='" & FGrid.TextMatrix(I, Col_ReqNoDocId) & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ReqSrNo)))
        Next
    End If
    GCn.Execute ("Update Hiscard set Locked_Text='" & txt(CloseRemark) & "',LJob_DocId='" & txt(JobNo).Tag & "',LJob_Date= " & ConvertDate(txt(JobCDt)) & _
        " where CardNo='" & mCardNo & "'")
    'A/c Posting
    If PubDealerID = "1109800" Then
        If txt(CashBill) = "Yes" And CDate(txt(JobCDt)) <= CDate(pubLockDate) Then
'           MsgBox "Job Close Date " & Txt(JobCDt) & " is less than Lock Date " & pubLockDate, vbInformation, "Works Cash Posting Locked"
            GoTo lblExit2
        End If
        ProcAcPost rsCtrlAc, rsCtrlAcLab
    Else
        If mCmdPostCounter = 0 Then
            ProcAcPost rsCtrlAc, rsCtrlAcLab
        End If
    End If
    'EOF of A/c Posting Section
lblExit2:
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
lblExit:
    Set Rst = Nothing
    If mReposting = False Then UpdRequery
    If mAddFlag = "A" Then
        If PubMoveRecYn Then
            Master.Requery
        Else
            Set Master = GCn.Execute("select Jc.DocId AS CODE, ClosedU_EntDt, JobCloseDate " _
                    & "from job_card as JC where left(JC.DocId,1)='" & PubDivCode & "' and JC.JobCloseDate Is Not Null And JC.DocId = '" & txt(JobNo).Tag & "' Order by JC.JobCloseDate desc")
        End If
        txt(JobCDt).Tag = txt(JobCDt)
    End If
    
    Master.FIND "Code = '" & txt(JobNo).Tag & "'"
    If mReposting = False Then
        TopCtrl1_ePrn
    End If
    Exit Sub

errlbl:
    If mTrans Then G_FaCn.RollbackTrans:  GCn.RollbackTrans
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
Dim xTurnOver, xTurnOverMrp As Double, xReSaleTaxAmt As Double, mFADocidSpr$, mFADocidLab$, mQry$
Dim xNetLabAmt As Double, xLabAmtTB As Double, xLabAmtTP As Double, xLabDisc As Double
Dim xServTaxAmt As Double, xLabROff, xsprRoff As Single
Dim RsTemp As ADODB.Recordset, rsTemp1 As ADODB.Recordset
'A/c Posting related declarations
Dim LedgAry() As LedgRec, LedgAryLab() As LedgRec, mCommNarr$, mLabSQL$
Dim mResult As Byte, mNarr$, TaxSQL$, I As Integer, j As Integer
Dim mSprAmtMRPTB As Double, mSprAmtTB As Double
Dim mOilAmtMRPTB As Double, mOilAmtTB As Double
Dim mTotMRPOilTB As Double, mTotOilTB As Double, mTotShareAmt As Double
Dim mShareSpr As Single, mShareAmtSpr As Double, mShare2AmtSpr As Double
Dim mTot1ShareAmt As Double, mTot2ShareAmt As Double, mTot3ShareAmt As Double
Dim PartyCode$, PartyCodeLab$
Dim mDrAmt As Double, mCrAmt As Double
Dim mDiff As Double
Dim TmpSQL$, SrvCat$, mFreeServCode$, mDName$, dSubCode$, FSBVal As Double, mServTaxPer As Double, mServTaxAmt As Double

    ApplyConsolidatedPosting CDate(txt(JobCDt))

    xNetLabAmt = 0
    xLabAmtTB = 0
    xLabAmtTP = 0
    xLabDisc = 0
    xServTaxAmt = 0
    xLabROff = 0
    xsprRoff = 0
    
    mOilAmtMRPTB = 0
    mSprAmtMRPTB = 0
    mSprAmtTB = 0
    mOilAmtTB = 0
    
    'If mVatYn = 1 Then
    '    TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code,sum(Tax_Amt) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt " & _
    '        " from SP_Sale left join TaxFormsAc as TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code"
    'Else
        TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code, TF.AddTaxAc, sum(Tax_Amt+Tax_AmtMRP) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt, Sum(SatAmt) As SatAmt " & _
            " from SP_Sale left join TaxFormsAc as TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code"
   
    'End If
    'to avoid errors of Old System
'    LedgerUnPost GCnFaS, Txt(JobNo).Tag
'    If SepLabPost Then
'        LedgerUnPost GCnFaW, Txt(JobNo).Tag
'    End If
    'eof of Old System
    If txt(CashBill) = "Yes" And IsConsolidatedPosting Then
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where V_Date=" & ConvertDate(CDate(txt(JobCDt))) & " and left(docid,8)='" & left(SpareDocID, 8) & _
            "' group by TF.PurSal_Ac_Code"
            mQry = "select " & _
                "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
                "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
                "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
                "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
                "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
                "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
                "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
                "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
                "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
                "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, " & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "Sum(TOT_Amt)", "Sum(TOT_Amt+TOT_AmtMrp)") & " as TurnOver, sum(TOT_AmtMrp) as  Tot_AmtMrp, " & _
                "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
                "from SP_Sale " & _
                "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
                "where V_Date= " & ConvertDate(CDate(txt(JobCDt))) & " and left(docid,8)='" & left(SpareDocID, 8) & "'"
        'for tax
        TaxSQL = TaxSQL & " where  V_Date=" & ConvertDate(CDate(txt(JobCDt))) & " and left(docid,8)='" & left(SpareDocID, 8) & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code, Tf.AddTaxAc"
        '**Labour
        mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
            ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Round(Lab_RoundOff,2)) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
            "from Job_Card where JobCloseDate=" & ConvertDate(CDate(txt(JobCDt))) & _
            " and left(DocId_InvLab,8)='" & left(LabourDocID, 8) & "'"
        '***********
        mNarr = "Workshop Cash Sale (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        mFADocidSpr = left(SpareDocID, 8) & "YYYYY" & "  " & Format(txt(JobCDt), "yymmdd")
        mFADocidLab = left(LabourDocID, 8) & "ZZZZZ" & "  " & Format(txt(JobCDt), "yymmdd")
        PartyCode = PubSprCashAc
        PartyCodeLab = PubSrvCashAc
    Else
        mFADocidSpr = SpareDocID
        mFADocidLab = LabourDocID
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where docid='" & SpareDocID & _
            "' group by TF.PurSal_Ac_Code"

        mQry = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, " & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "Sum(TOT_Amt)", "Sum(TOT_Amt+TOT_AmtMrp)") & " as TurnOver, sum(TOT_AmtMrp) as  Tot_AmtMrp, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "left join TaxFormsAc as TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where docid='" & SpareDocID & "'"
        'for tax
        TaxSQL = TaxSQL & " where docid='" & SpareDocID & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code, TF.AddTaxAc"
        '**Labour
'        mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
'            ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Lab_RoundOff) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
'            "from Job_Card where format(JobCloseDate,'dd/MMM/yyyy')=" & ConvertDate(Txt(JobCDt)) & _
'            " and DocId_InvLab='" & LabourDocID & "'"
        If PubBackEnd = "A" Then
            mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
                ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Round(Lab_RoundOff,2)) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
                "from Job_Card where format(JobCloseDate,'dd/MMM/yyyy')=" & ConvertDate(txt(JobCDt)) & _
                " and DocId='" & txt(JobNo).Tag & "'"
        Else
            mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
                ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Lab_RoundOff) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
                "from Job_Card where JobCloseDate=" & ConvertDate(txt(JobCDt)) & _
                " and DocId='" & txt(JobNo).Tag & "'"
        End If
        '****
'        mNarr = "Works Cr Spare Bill No. " & Right(SpareDocID, 13) & " Dt." & Txt(JobCDt)
'        If xNetLabAmt <> 0 Then
'            If lblLabourBill <> "" Then
'                mNarr = mNarr & " Labour Bill No. " & lblLabourBill
'            End If
'        End If
'        mCommNarr = mNarr & " [Common]"
        PartyCode = txt(SpareParty).Tag
        PartyCodeLab = txt(LabourParty).Tag
    End If
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open mQry, GCn, adOpenStatic, adLockReadOnly
    
    'for tax purpose
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open TaxSQL, GCn, adOpenStatic, adLockReadOnly

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
    If UCase(left(PubComp_Name, 3)) <> "JMK" Then
        xTurnOverMrp = IIf(IsNull(rsTemp1!Tot_AmtMrp), 0, rsTemp1!Tot_AmtMrp)
    End If
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
            mShareAmt = Round((xDisAmtMRPTB + xDisAmtTB) * mShare / 100, 2)
            mShare2Amt = Round(xTaxAmtMRP * mShare / 100, 2)
            mShareSpr = Round(mSprAmtMRPTB * 100 / (mSprAmtMRPTB + mOilAmtMRPTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mShare2AmtSpr = Round(mShare2Amt * mShareSpr / 100, 2)
            mTot1ShareAmt = mTot1ShareAmt + mShareAmt
            mTot2ShareAmt = mTot2ShareAmt + mShare2Amt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                If Val(txt(DiscPerTB)) = 0 Then
                    mShareAmt = mShareAmt + ((xDisAmtMRPTB + xDisAmtTB) - mTot1ShareAmt)
                Else
                    mShareAmt = mShareAmt + (xDisAmtMRPTB - mTot1ShareAmt)
                End If
                mShare2Amt = mShare2Amt + (xTaxAmtMRP - mTot2ShareAmt)
            End If
            If UCase(left(PubComp_Name, 3)) = "JMK" Then
                mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr)
            Else
                
                'If mVatYn = 1 Then
                '    mSprAmtMRPTB = mSprAmtMRPTB '- (mShareAmtSpr + mShare2AmtSpr + xTurnOverMrp)
                'Else
                    mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr + xTurnOverMrp)
                'End If
            End If
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
    'If Txt(CashBill) <> "Yes" Then
        mNarr = "Works Job No. " & PrinID(txt(JobNo).Tag) & "  Spare Bill No. " & PrinID(SpareDocID) & " Dt." & txt(JobCDt) & " Rs." & Format(xNetAmt, "0.00") & " Veh.No." & txt(VehRegNo).TEXT
        If xNetLabAmt <> 0 Then
            If lblLabourBill <> "" Then
                mNarr = mNarr & " Labour Bill No. " & PrinID(LabourDocID) & " Rs." & Format(xNetLabAmt, "0.00")
            End If
        End If
        mCommNarr = mNarr & " "
    'End If
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
     If RsTemp.RecordCount > 0 Then
         Do While RsTemp.EOF = False
            If RsTemp!TaxAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!Tax_Ac_Code
                If RsTemp!TaxAmt > 0 Then
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(RsTemp!TaxAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(RsTemp!TaxAmt), 2)
                    LedgAry(I).AmtCr = 0
                End If
                LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
            If RsTemp!SatAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!AddTaxAc
                If RsTemp!TaxAmt > 0 Then
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(RsTemp!SatAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(RsTemp!SatAmt), 2)
                    LedgAry(I).AmtCr = 0
                End If
                LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
            
            If RsTemp!TaxSurAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!Sur_Ac_Code
                If RsTemp!TaxSurAmt > 0 Then
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(RsTemp!TaxSurAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(RsTemp!TaxSurAmt), 2)
                    LedgAry(I).AmtCr = 0
                End If
                 LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
             End If
             RsTemp.MoveNext
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
            LedgAryLab(I).SubCode = IIf(txt(CashBill) = "Yes", PubSrvCashAc, txt(LabourParty).Tag)
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
    
        '**********************************
        'Round Off = Spare Round Off + Labour round Off
        If xRoundAmt + xLabROff <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = rsCtrlAc!SprROff_Ac
            If xRoundAmt + xLabROff > 0 Then
                LedgAry(I).AmtCr = Round(xRoundAmt, 2) + Round(xLabROff, 2)
            Else
                LedgAry(I).AmtDr = Abs(Round(xRoundAmt, 2) + Round(xLabROff, 2))
            End If
            LedgAry(I).Narration = mNarr & " Round Diff. Spare+Labour"
        End If
    End If
    
    If PubSiebelActiveYn = 1 Then
        mDrAmt = 0
        mCrAmt = 0
        For j = 0 To UBound(LedgAry)
            mDrAmt = mDrAmt + IIf(IsNull(LedgAry(j).AmtDr), 0, LedgAry(j).AmtDr)
            mCrAmt = mCrAmt + IIf(IsNull(LedgAry(j).AmtCr), 0, LedgAry(j).AmtCr)
        Next
        mDiff = Round((Format(mDrAmt, "0.00") - Format(mCrAmt, "0.00")), 2)
        
        If mDiff > 0 Then
            If mDiff <= 0.05 Then
                If LedgAry(I).AmtDr > 0 Then
                    LedgAry(I).AmtDr = Val(LedgAry(I).AmtDr) - mDiff
                Else
                    LedgAry(I).AmtCr = Val(LedgAry(I).AmtCr) + mDiff
                End If
            End If
        Else
            If mDiff >= -0.05 Then
                If LedgAry(I).AmtDr > 0 Then
                    LedgAry(I).AmtDr = Val(LedgAry(I).AmtDr) - mDiff
                Else
                    LedgAry(I).AmtCr = Val(LedgAry(I).AmtCr) + mDiff
                End If
            End If
        End If
    End If
    
    mResult = LedgerPost(mAddFlag, LedgAry, GCnFaS, mFADocidSpr, CDate(txt(JobCDt)), mCommNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    If SepLabPost Then
        mResult = LedgerPost(mAddFlag, LedgAryLab, GCnFaW, mFADocidLab, CDate(txt(JobCDt)), mCommNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    End If
   
        'Free service coupon value posting.
        If VNull(G_FaCn.Execute("Select FSBOnlinePost from AcControls where Div_Code='" & PubDivCode & "'").Fields(0).Value) = 1 And TopCtrl1.TopText2 <> "Browse" Then
            I = 0: xServTaxAmt = 0
            Dim mGroupCode As String, mDCity As String, mDCode As String
            ReDim LedgAry(I)
            SrvCat = GCn.Execute("Select Serv_Catg from Service_Type Left Join Job_Card on Service_type.Serv_type=Job_Card.Serv_Type where Job_card.Job_No=" & txt(JobNo) & "").Fields(0).Value
            If (SrvCat = "F") Then
                Set RsTemp = New ADODB.Recordset
                RsTemp.CursorLocation = adUseClient
                RsTemp.Open "Select AMd_Dealer.D_Code, AMd_Dealer.D_Name,AMd_Dealer.D_City from (Amd_Dealer Left Join HisCard on HisCard.Dealer_code=Amd_Dealer.D_Code) Left Join Job_card on job_Card.CardNO=HisCard.CardNo  where Job_Card.Job_No =" & txt(JobNo) & "", GCn, adOpenStatic, adLockReadOnly
                If RsTemp.RecordCount > 0 Then
                    mDName = XNull(RsTemp!D_Name)
                    mDCity = XNull(RsTemp!D_City)
                    mDCode = XNull(RsTemp!D_Code)
                Else
                    MsgBox "Selling Dealer Information having some problem", vbCritical, "Error Message for Free/PDI Service"
                    GoTo SkipSellingDealer
                End If
                Set RsTemp = New ADODB.Recordset
                RsTemp.CursorLocation = adUseClient
                RsTemp.Open "Select SubGroup.subCode,SubGroup.Name from Subgroup  With (NoLock) where subgroup.name ='" & left(mDName, 40) & "'", G_FaCn, adOpenStatic, adLockReadOnly
                If RsTemp.RecordCount > 0 Then
                    dSubCode = XNull(RsTemp!SubCode)
                Else
                        dSubCode = PubSiteCode & IIf(PubFirmCode = "", "0", PubFirmCode) & Format(G_CompCn.Execute("Select SubGroupAcCode From SubGroupCounter  With (NoLock)").Fields(0).Value, "000000")
                        mGroupCode = XNull(G_FaCn.Execute("Select OthDealerGrp from AcControls  With (NoLock)").Fields(0).Value)
                        TmpSQL = SubGroupUpdate("Add", "SubGroup", dSubCode, dSubCode, left(mDName, 40), "", "N", "A", mGroupCode)
                        If PubBackEnd = "A" Then G_FaCn.Execute (TmpSQL)
                        GCn.Execute (TmpSQL)

                        TmpSQL = SubGroupUpdate("Add", "SubGroupAlias", dSubCode, dSubCode, left(mDName, 40), "", "N", "A", mGroupCode)
                        If PubBackEnd = "A" Then G_FaCn.Execute (TmpSQL)
                        GCn.Execute (TmpSQL)
                        
                        G_CompCn.Execute ("Update SubGroupCounter set SubGroupAcCode=SubGroupAcCode+1 ")
                   'End If
                End If
                If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
                        FSBVal = FreeLabForTax
                        mServTaxPer = VNull(GCn.Execute("Select Service_tax from Syctrl").Fields(0).Value)
                        If UCase(mDCity) = UCase(PubComp_City) Or (mDCode <> PubDealerID And PubDealerID <> "") Then
                            FSBVal = FSBVal + (FSBVal * 20 / 100)
                        End If
                        If FSBVal <> 0 Then
                        I = UBound(LedgAry) + 1
                        ReDim Preserve LedgAry(I)
                        LedgAry(I).SubCode = dSubCode
                        LedgAry(I).AmtDr = FSBVal + (FSBVal * Format(mServTaxPer, "0.00")) / 100
                        LedgAry(I).Narration = mNarr & " Free Service Coupon Value "

                        I = UBound(LedgAry) + 1
                        ReDim Preserve LedgAry(I)
                        LedgAry(I).SubCode = rsCtrlAc!FSBCrAc
                        LedgAry(I).AmtCr = FSBVal + (FSBVal * Format(mServTaxPer, "0.00")) / 100
                        LedgAry(I).Narration = mNarr & " Free Service Coupon Value "

                    End If
                Else
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.CursorLocation = adUseClient
                    RsTemp.Open "Select Coupon_Value from Job_Card With (NoLock) where Job_No =" & txt(JobNo) & "", GCn, adOpenStatic, adLockReadOnly
                    FSBVal = VNull(RsTemp!Coupon_Value)
                    mServTaxPer = VNull(GCn.Execute("Select Service_tax from Syctrl  With (NoLock) ").Fields(0).Value)
'                    If UCase(mDCity) = UCase(PubComp_City) Or (mDCode <> PubDealerID And PubDealerID <> "") Then
'                       FSBVal = FSBVal + (FSBVal * 20 / 100)
'                    End If
                    xServTaxAmt = (FSBVal * Format(mServTaxPer, "0.00")) / 100
                    If FSBVal <> 0 Then
                        I = UBound(LedgAry) + 1
                        ReDim Preserve LedgAry(I)
                        LedgAry(I).SubCode = dSubCode
                        LedgAry(I).AmtDr = FSBVal + xServTaxAmt
                        LedgAry(I).Narration = mNarr & " Free Service Coupon Value "

                        I = UBound(LedgAry) + 1
                        ReDim Preserve LedgAry(I)
                        LedgAry(I).SubCode = rsCtrlAc!FSBCrAc
                        LedgAry(I).AmtCr = FSBVal
                        LedgAry(I).Narration = " Free Service Coupon Value "

                    End If
                 End If
            End If
            
            If xServTaxAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = rsCtrlAcLab!SrvTax_Ac    'Service Tax A/c Code
                LedgAry(I).AmtCr = xServTaxAmt
                LedgAry(I).Narration = " Service Tax on Free Service Coupon Value"
            End If
            PubImportData = True
                mResult = LedgerPost(mAddFlag, LedgAry, GCnFaS, lblDocId, CDate(txt(JobCDt)), "Against Free Service Coupon Value")
                If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
            PubImportData = False
            
        End If
SkipSellingDealer:
    If UCase(left(PubComp_Name, 6)) = "TRUPTI" Then
            'Nra modi for contractor amount posting For trupti Motors
             Dim T As Integer, mContAmt As Double, mJobContAc As String
             I = 0
             ReDim LedgAry(I)
             mContAmt = 0
             mJobContAc = XNull(G_FaCn.Execute("Select JobContractor From AcControls Where Div_Code= '" & PubDivCode & "'").Fields(0).Value)
             For T = 1 To FGrid1.Rows - 1
                 If FGrid1.TextMatrix(T, C_External) = "Yes" Then
                     If FGrid1.TextMatrix(T, C_ContAcCode) = "" Then
                         MsgBox "Please define Contractor A/C in Contractor/OEM Master For Contractor " & FGrid1.TextMatrix(T, C_ContName) & " .Posting Aborted !", vbInformation
                         Exit Sub
                     End If
                     If mJobContAc = "" Then
                         MsgBox "Please define Job Contractor A/C in System Controls.Posting Aborted !", vbInformation
                         Exit Sub
                     End If
                     If Val(FGrid1.TextMatrix(T, C_ContAmt)) <> 0 Then
                         I = UBound(LedgAry) + 1
                         ReDim Preserve LedgAry(I)
                         LedgAry(I).SubCode = FGrid1.TextMatrix(T, C_ContAcCode)
                         LedgAry(I).AmtCr = Val(FGrid1.TextMatrix(T, C_ContAmt))
                         mContAmt = mContAmt + Val(FGrid1.TextMatrix(T, C_ContAmt))
                         LedgAry(I).Narration = mNarr & "Contractor charges"
                     End If
                 End If
             Next
             If mContAmt <> 0 Then
                 I = UBound(LedgAry) + 1
                 ReDim Preserve LedgAry(I)
                 LedgAry(I).SubCode = mJobContAc
                 LedgAry(I).AmtDr = mContAmt
                 LedgAry(I).Narration = mNarr & "Contractor charges"
             End If
             
             mResult = LedgerPost(mAddFlag, LedgAry, GCnFaS, lblDocId, CDate(txt(JobCDt)), mCommNarr)
             If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
             
             '***********************end modi **********************
    End If
lblExit:
    Set GRs = Nothing
    Set RsTemp = Nothing
    Set rsTemp1 = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description, vbCritical, "Ledger Posting Failed!'"
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo lblExit
    Ctrl_GetFocus txt(Index)
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case JobNo
            DGridColSwap DGJob, 0
            RsJob.Sort = "JOB_NO"   'FindJobNo" '
            If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
            If txt(Index).Tag <> "" And txt(Index).Tag <> RsJob!Code Then
'                RsJob.MoveFirst
                RsJob.FIND ("JOB_NO='" & txt(Index).TEXT & "'")
            End If
        Case Chassis
            DGridColSwap DGJob, 1
            RsJob.Sort = "CHASSIS"
            If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
            If txt(Index).Tag <> "" And txt(Index).Tag <> RsJob!Code Then
'                RsJob.MoveFirst
                RsJob.FIND ("CHASSIS='" & txt(Index).TEXT & "'")
            End If
        Case VehRegNo
            DGridColSwap DGJob, 2
            RsJob.Sort = "REGNO"
            If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
            If txt(Index).Tag <> "" And txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("REGNO='" & txt(Index).TEXT & "'")
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
            If RsMech.RecordCount = 0 Then Exit Sub
                RsMech.MoveFirst
                If txt(Index).TEXT <> "" And txt(Index).Tag <> RsMech!Code Then
                    RsMech.FIND ("name='" & txt(Index).TEXT & "'")
                End If
        Case SuperName
            DGMech.Columns(0).CAPTION = "WorkShop Staff"
            Set DGMech.DataSource = RsSuper
            DGridColSwap DGMech, 1
            RsSuper.Sort = "name"
            If txt(Index).TEXT <> "" And txt(Index).Tag <> RsSuper!Code Then
                RsSuper.FIND ("name='" & txt(Index).TEXT & "'")
            End If
        Case JobDelay
            DGridColSwap DGReason, 1
            RsReason.Sort = "name"
            If RsReason.RecordCount > 0 Then
                If txt(Index).TEXT <> "" And txt(Index).Tag <> RsReason!Code Then
                    RsReason.FIND ("name='" & txt(Index).TEXT & "'")
                End If
            End If
        Case SpareParty
            DGridColSwap DGParty, 1
            RsParty.Sort = "name"
            If txt(Index).TEXT <> "" And txt(Index).Tag <> RsParty!Code Then
                RsParty.FIND ("name='" & txt(Index).TEXT & "'")
            End If
        Case LabourParty
            DGridColSwap DGParty, 1
            RsParty.Sort = "name"
            If txt(Index).TEXT <> "" And txt(Index).Tag <> RsParty!Code Then
                RsParty.FIND ("name='" & txt(Index).TEXT & "'")
            End If
    End Select
Exit Sub
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
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 1
        Case VehRegNo
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 3
        Case Chassis
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 4
'        Case OwnerName
'            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 7
        Case MechName
            DGridTxtKeyDown DGMech, txt, Index, RsMech, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case SuperName
            DGridTxtKeyDown DGMech, txt, Index, RsSuper, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case JobDelay
            DGridTxtKeyDown DGReason, txt, Index, RsReason, KeyCode, False, 1, frmJobDelay, "frmJobDelay"
        Case SpareParty
            DGridTxtKeyDown DGParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            If KeyCode = 13 Then
                If RsParty.BOF = False And RsParty.EOF = False Then
                    LblCurrBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                    LblCurrBal = LblCurrBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
                End If
            End If
        Case LabourParty
            DGridTxtKeyDown DGParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            If KeyCode = 13 Then
                If RsParty.BOF = False And RsParty.EOF = False Then
                    LblCurrBal1 = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                    LblCurrBal1 = LblCurrBal1 & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
                End If
            End If
        Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, TurnOverAmt, Excise_Amt
            NumDown txt(Index), KeyCode, 8, 2
        Case GenSurPer, STaxPer, TaxSurPer, TurnOverPer
            NumDown txt(Index), KeyCode, 2, 2
        Case DiscPerTB, DiscPerTP, LabDisPer
            NumDown txt(Index), KeyCode, 2, 4
    End Select
    If DGJob.Visible = False And DGMech.Visible = False And DGReason.Visible = False And DGParty.Visible = False Then
        '' KEY DOWN
        If KeyCode = vbKeyReturn And lblGroup.Visible = True Then
            lblGroup.Visible = False
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If (txt(LabourParty).Enabled = True And Index <> LabourParty) Or (txt(LabourParty).Enabled = False And Index <> CashParty) Then
                Ctrl_DownKeyDown KeyCode, Shift
            End If
            
            If (txt(LabourParty).Enabled = True And Index = LabourParty) Or (txt(LabourParty).Enabled = False And Index = CashParty) Or (Index = CreditCardNo And txt(CreditCardNo) <> "") Or (Index = ChqDate And txt(ChqNo) <> "") Then
                 Txt_Validate Index, False
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

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
        Case JobCompTm
            Call NumPress(txt(Index), KeyAscii, 2, 2)
        Case JobNo
            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "FindJobNo"
        Case VehRegNo
            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "regno"
        Case Chassis
            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "chassis"
'        Case OwnerName
'            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "name"
        Case MechName, SuperName
            DGridTxtKeyPress txt, Index, RsMech, KeyAscii, "name"
        Case SuperName
            DGridTxtKeyPress txt, Index, RsSuper, KeyAscii, "name"
        Case JobDelay
            DGridTxtKeyPress txt, Index, RsReason, KeyAscii, "name"
        Case SpareParty
             If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "name":   lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
        Case LabourParty
            If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "name":   lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
        Case DiscAmtTB, DiscAmtTP, Addition, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, TurnOverAmt, LabAmt, SROff, ServTaxAmt, NetLabAmt, Excise_Amt
            NumPress txt(Index), KeyAscii, 8, 2
        Case GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ServTaxPer
            NumPress txt(Index), KeyAscii, 2, 2
        Case LabDisPer
            NumPress txt(Index), KeyAscii, 2, 4
        Case IWDiscPerTB, IWDiscPerTP
           If StrCmp(left(PubComp_Name, 3), "jmk") Then
                If Val(txt(DiscPerTB)) > 0 Or Val(txt(DiscPerTP)) > 0 Then
                    KeyAscii = 0
                End If
            End If
            NumPress txt(Index), KeyAscii, 2, 4
        Case DiscPerTB, DiscPerTP
            If StrCmp(left(PubComp_Name, 3), "jmk") Then
                If Val(txt(IWDiscTotTB)) > 0 Or Val(txt(IWDiscTotTP)) > 0 Then
                    KeyAscii = 0
                End If
            End If
            NumPress txt(Index), KeyAscii, 2, 4
        Case CashBill
            If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
                If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                    txt(Index).TEXT = "Yes"
                    KeyAscii = 0
                    mVType = "W_SIC"
                    LabourVtype = "W_LIC"
                ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                    txt(Index).TEXT = "No"
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
        Case DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, SROff, ReSalTaxPer, ReSalTaxAmt, IWDiscPerTB, IWDiscPerTP, Excise_Amt
            If Index = IWDiscPerTB Or Index = IWDiscPerTP Then Amt_Cal
                            
            If Val(txt(MRPAmtTB)) + Val(txt(MRPAmtTP)) <> 0 Then
                MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                        Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                        Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
                        Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
            
            End If
            '*******************************************************************
            If mVatYn = 1 Then
                
                
                MainLib.SprCalcVAT WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                        Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                        Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                        txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                        txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                        txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                        txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                        txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                        txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt), Col_Purpose, True
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
            End If
                ', _
                Txt (LabAmt), Txt(LabDisc), Txt(ServTaxPer), Txt(ServTaxAmt), Txt(LabROff), Txt(NetLabAmt), Txt(OutSideLabAmt)
            'Nra updation
            If Val(txt(LabAmtTB)) <> 0 Then
                txt(ServTaxPer) = MainLib.Serv_Tax
            Else
                txt(ServTaxPer) = Format(0, "0.00")
                txt(ServTaxAmt) = Format(0, "0.00")
            End If
            'Nra end updation
            
            MainLib.LabCalc txt(LabAmtTB), txt(LabAmtTP), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt), mLabDiscAmtTB, mECessPer, mECessAmt, txt(FreeWarrLabAmt), mServiceTaxPer_Saperate, mServiceTaxAmt_Saperate, mHECessPer, mHECessAmt
            If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
                txt(ServTaxPer) = MainLib.Serv_Tax
                txt(ServTaxAmt).TEXT = Format((Val(txt(LabAmtTB).TEXT) + Val(FreeLabForTax) - Val(txt(LabDisc).TEXT)) * Val(txt(ServTaxPer)) / 100, "0.00")
                txt(NetLabAmt).TEXT = Val(txt(LabAmtTB).TEXT) + Val(txt(LabAmtTP).TEXT) + Val(txt(ServTaxAmt).TEXT) - Val(txt(LabDisc).TEXT)
                txt(LabROff).TEXT = Format(dmRoundOff(txt(NetLabAmt)), "0.00")
                txt(NetLabAmt).TEXT = Format(Val(txt(NetLabAmt)) + Val(txt(LabROff)), "0.00")
            End If
            
            If UCase(left(PubComp_Name, 3)) = "JMK" Then
                    txt(TurnOverAmt) = Format((Val(txt(STotATB)) + Val(txt(STaxAmt))) * Val(txt(TurnOverPer)) / 100, "0.00")
                    txt(NetSprAmt).TEXT = Format(Val(txt(STotB)) + Val(txt(TurnOverAmt)), "0.00")
                    'txt(SROff).TEXT = Format(Val(txt(NetSprAmt).TEXT) - Round(Val(txt(STotB).TEXT) + Val(txt(TurnOverAmt)), 0), "0.00")
                    txt(NetSprAmt).TEXT = Format(Round(txt(NetSprAmt).TEXT, 0), "0.00")
                    txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
            Else
                txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
            End If

            
            
        Case LabDisPer, LabAmt, LabDisc, ServTaxPer, ServTaxAmt
            If Index = LabDisPer Then
                If PubOutSideLabDisc = 0 Then   'No
                    txt(LabDisc) = Round((Val(txt(LabAmtTB)) + Val(txt(LabAmtTP)) - Val(txt(OutSideLabAmt))) * Val(txt(LabDisPer)) / 100, 2)
                Else
                    txt(LabDisc) = Round((Val(txt(LabAmtTB)) + Val(txt(LabAmtTP))) * Val(txt(LabDisPer)) / 100, 2)
                    mLabDiscAmtTB = Round(Val(txt(LabAmtTB)) * Val(txt(LabDisPer)) / 100, 2)
                End If
            End If
            'Nra updation
            If Val(txt(LabAmtTB)) <> 0 Then
                txt(ServTaxPer) = MainLib.Serv_Tax
            Else
                txt(ServTaxPer) = Format(0, "0.00")
                txt(ServTaxAmt) = Format(0, "0.00")
            End If
            'Nra end updatio
            MainLib.LabCalc txt(LabAmtTB), txt(LabAmtTP), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt), mLabDiscAmtTB, mECessPer, mECessAmt, txt(FreeWarrLabAmt), mServiceTaxPer_Saperate, mServiceTaxAmt_Saperate, mHECessPer, mHECessAmt
            
            If mLabDiscAfterTaxYn = 1 Then
                If Index = LabDisPer Then
                    If PubOutSideLabDisc = 0 Then   'No
                        txt(LabDisc) = Round((Val(txt(LabAmtTB)) + Val(txt(LabAmtTP)) - Val(txt(ServTaxAmt)) - Val(txt(OutSideLabAmt))) * Val(txt(LabDisPer)) / 100, 2)
                    Else
                        txt(LabDisc) = Round((Val(txt(LabAmtTB)) + Val(txt(LabAmtTP)) - Val(txt(ServTaxAmt))) * Val(txt(LabDisPer)) / 100, 2)
                        mLabDiscAmtTB = Round((Val(txt(LabAmtTB)) - Val(txt(ServTaxAmt))) * Val(txt(LabDisPer)) / 100, 2)
                    End If
                End If
            End If
            
            
            If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
                txt(ServTaxPer) = MainLib.Serv_Tax
                txt(ServTaxAmt).TEXT = Format((Val(txt(LabAmtTB).TEXT) + Val(FreeLabForTax) - Val(txt(LabDisc))) * Val(txt(ServTaxPer)) / 100, "0.00")
                txt(NetLabAmt).TEXT = Val(txt(LabAmtTB).TEXT) + Val(txt(LabAmtTP).TEXT) + Val(txt(ServTaxAmt).TEXT) - Val(txt(LabDisc).TEXT)
                txt(LabROff).TEXT = Format(dmRoundOff(txt(NetLabAmt)), "0.00")
                txt(NetLabAmt).TEXT = Format(Val(txt(NetLabAmt)) + Val(txt(LabROff)), "0.00")
            End If
            txt(NetAmt) = Format(Round(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), 0), "0.00")
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Dim RsTemp As Recordset
    Select Case Index
        Case JobCompTm
            txt(Index) = Format(txt(Index), "hh:mm")
        Case JobNo, VehRegNo, Chassis ', OwnerName
            If txt(Index).Tag <> "" Then
                RsJob.Sort = "CODE"
                RsJob.FIND ("CODE='" & txt(Index).Tag & "'")
            Else
                Cancel = True: Exit Sub
            End If
            If RsJob.BOF = True Or RsJob.EOF = True Then Exit Sub
            
            
            If Index = JobNo And (UCase(left(PubComp_Name, 3)) = "JMK" Or UCase(left(PubComp_Name, 3)) = "LMP") Then
                Set RsTemp = GCn.Execute("Select " & xIsNull("Locked_Text", "") & " From HisCard Where CardNo = (Select CardNo From Job_Card Where DocId = '" & RsJob!Code & "')")
                If RsTemp(0) <> "" Then
                    MsgBox RsTemp(0)
                End If
            End If
            
            'External Job Recd Checking
            GSQL = "Select Job_Docid From Job_GatePass as GP " & _
                "Where GP.Job_DocId='" & txt(JobNo).Tag & "' and ContractRecdDate is null"
            If GCn.Execute(GSQL).RecordCount > 0 Then
                MsgBox "External Job Not Recd !", vbCritical, "Extrenal Job !"
                Cancel = True: Exit Sub
            End If
            'External Job Entry Checking
'            GSQL = "Select GatePassNo From Job_GatePass as GP " & _
'                "Where GP.Job_DocId='" & Txt(JobNo).Tag & "' and GP.GatePassNo not in (Select distinct ExtJobGatePassNo from Job_Lab as JL where JL.Job_DocId='" & Txt(JobNo).Tag & "')"
'            If GCn.Execute(GSQL).RecordCount  > 0 Then
'                MsgBox "External Job Entry Pending !", vbCritical, "Labour Entry!"
'                Cancel = True: Exit Sub
'            End If
            'Labour Checking
            GSQL = "Select Job_Docid From Job_Lab JL " & _
                "Where JL.Job_DocId='" & txt(JobNo).Tag & "'"
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
                If txt(Index) <> "" Then
                    txt(VehRegNo).TabStop = False
                    txt(Chassis).TabStop = False
                Else
                    txt(VehRegNo).TabStop = False
                    txt(Chassis).TabStop = False
                End If
            End If
            '**

        Case MechName
            If RsMech.EOF = True Or RsMech.BOF = True Or txt(Index).TEXT = "" Then
                txt(Index).Tag = ""
                txt(Index).TEXT = ""
            Else
                txt(Index).Tag = RsMech!Code
                txt(Index).TEXT = RsMech!Name
            End If
        Case SuperName
            If RsSuper.EOF = True Or RsSuper.BOF = True Or txt(Index).TEXT = "" Then
                txt(Index).Tag = ""
                txt(Index).TEXT = ""
            Else
                txt(Index).Tag = RsSuper!Code
                txt(Index).TEXT = RsSuper!Name
            End If
        Case JobDelay
            If RsReason.EOF = True Or RsReason.BOF = True Or txt(Index).TEXT = "" Then
                txt(Index).Tag = ""
                txt(Index).TEXT = ""
            Else
                txt(Index).Tag = RsReason!Code
                txt(Index).TEXT = RsReason!Name
            End If
        'Modi LPS 01-04
        Case SpareParty
            If RsParty.EOF = True Or RsParty.BOF = True Or txt(Index).TEXT = "" Then
                txt(SpareParty).Tag = ""
                txt(SpareParty).TEXT = ""
                txt(LabourParty).Tag = ""
                txt(LabourParty).TEXT = ""
                LblCurrBal = ""
                LblCurrBal1 = ""
            Else
                txt(SpareParty).Tag = RsParty!Code
                txt(SpareParty).TEXT = RsParty!Name
                txt(LabourParty).Tag = RsParty!Code
                txt(LabourParty).TEXT = RsParty!Name
                LblCurrBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                LblCurrBal = LblCurrBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
                LblCurrBal1 = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                LblCurrBal1 = LblCurrBal1 & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
                
            End If
        Case LabourParty
            If RsParty.EOF = True Or RsParty.BOF = True Or txt(Index).TEXT = "" Then
                txt(Index).Tag = ""
                txt(Index).TEXT = ""
                LblCurrBal1 = ""
            Else
                txt(Index).Tag = RsParty!Code
                txt(Index).TEXT = RsParty!Name
                LblCurrBal1 = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                LblCurrBal1 = LblCurrBal1 & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
            End If
        'eof modi
        Case JobCDt
            txt(Index).TEXT = RetDate(txt(Index))
            Cancel = Not CheckFinYear(txt(Index))
            If Cancel Then Exit Sub
            GSQL = "Select top 1 v_date from Sp_Stock where Job_Docid='" & txt(JobNo).Tag & "' and V_Date >" & ConvertDate(Format(txt(JobCDt), "dd/mmm/yyyy")) & ""
            If GCn.Execute(GSQL).RecordCount > 0 Then
                MsgBox "Job Close Date is Less than Part Issue Date", vbCritical, "Date Checking!"
                Cancel = True: Exit Sub
            End If
            If RetDate(txt(JobCDt)) <= RetDate(txt(DelDate)) Then
                txt(JobDelay).Enabled = False
                txt(JobDelay).TEXT = ""
                txt(JobDelay).Tag = ""
            Else
                txt(JobDelay).Enabled = True
            End If
            Dim mNextSrvDays As Integer
            If txt(NextSrv) = "" Then
                mNextSrvDays = GCn.Execute("Select " & vIsNull("Max(Days)", "0") & " From Service_Type Where Serv_Type='" & txt(SrvType).Tag & "'").Fields(0).Value
                txt(NextSrv) = CDate(txt(JobCDt)) + IIf(mNextSrvDays = 0, PubNextSrvDays, mNextSrvDays)
            End If
            txt(JobCompDt).TEXT = Format(txt(JobCDt), "dd/MMM/yyyy")
        Case JobCompDt, ChqDate
            txt(Index).TEXT = RetDate(txt(Index))
        Case NextSrv
            txt(Index).TEXT = RetDate(txt(Index))
        Case DiscAmtTB, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, SROff, LabAmt, LabDisc, ServTaxPer, ServTaxAmt, Excise_Amt
            If Val(txt(Index).TEXT) = 0 Then
                txt(Index).TEXT = ""
            Else
                txt(Index).TEXT = Format(txt(Index), "0.00")
            End If
        Case LabDisPer, DiscPerTB, DiscPerTP
            If Val(txt(Index).TEXT) = 0 Then
                txt(Index).TEXT = ""
            Else
                txt(Index).TEXT = Format(txt(Index), "0.0000")
            End If
        Case CashBill
            If txt(Index).TEXT = "Yes" Then
                txt(CashParty) = txt(OwnerName)
            End If
            Call Generate_Prefix
            Call txtDisabled_Color
        Case CashParty
            If TopCtrl1.TopText2 = "Add" Then
                Call Generate_Prefix
                Call txtDisabled_Color
            End If
        Case CreditCardNo
            If Trim(txt(Index)) <> "" Then
                txt(SpareParty).Tag = PubCreditCardAc
                txt(SpareParty) = GCn.Execute("Select Name From SubGroup Where SubCode='" & txt(SpareParty).Tag & "'").Fields(0)
                txt(LabourParty).Tag = PubCreditCardAc
                txt(LabourParty) = GCn.Execute("Select Name From SubGroup Where SubCode='" & txt(LabourParty).Tag & "'").Fields(0)
                txt(LabourParty).Enabled = False
                txt(SpareParty).Enabled = False
                Generate_Prefix
            End If
        Case ChqNo
            If Trim(txt(Index)) <> "" Then
                txt(SpareParty).Tag = PubChqClrAc
                txt(SpareParty) = GCn.Execute("Select Name From SubGroup Where SubCode='" & txt(SpareParty).Tag & "'").Fields(0)
                txt(LabourParty).Tag = PubChqClrAc
                txt(LabourParty) = GCn.Execute("Select Name From SubGroup Where SubCode='" & txt(LabourParty).Tag & "'").Fields(0)
                txt(LabourParty).Enabled = False
                txt(SpareParty).Enabled = False
                Generate_Prefix
            End If
        
    End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
        If I <> JobCDt Then
            txt(I).Tag = ""
        End If
    Next I
    mCardNo = ""
    mLabDiscAmtTB = 0
    lblDocId.CAPTION = ""
    lblDocId.Refresh
    
    LblSprBill = ""
    lblLabourBill = ""
    lblGatePass = ""
    LblCurrBal = ""
    LblCurrBal1 = ""
    
    mMRevDisTBPer = 0
    mMRevDisTPPer = 0
    mTBDisAmtMRP = 0
    mTPDisAmtMRP = 0
    
End Sub

Private Sub MoveRec()
On Error GoTo errlbl
Dim CurrBal As Double
Dim Master1 As ADODB.Recordset ',rs As Recordset
'Dim mVor As String
'Dim i As Integer
mMRPReSales = 0
mMRPLubeTB = 0
mMRPLubeTP = 0
mAddFlag = "I"
mCardNo = ""

    If Master.RecordCount > 0 Then
        Set Master1 = GCn.Execute("select JC.Job_No,JC.Site_Code,JC.Govt_YN, JC.Job_Date, JC.JobCloseDate,jc.cardno,jc.OpenRemarks, " _
                                & "jc.Body_Damage,jc.ObservBy_Eng,Jc.Job_BookNo,Jc.Job_InspDocID,Jc.AtKMsHrs,jc.Coupon,jc.Coupon_Value, " _
                                & "jc.ArrivalTime,jc.ExpDelDate,jc.Est_SpCost,jc.Est_LabCost,Jc.DelBy,Jc.RecBy_Supervisor,Jc.DelayReason, " _
                                & "Jc.JobComp_Dt_Time,JC.Remark,jc.NextSrvDate,EM.EMP_NAME AS Mechanic,EMP.Emp_Name as Supervisor, " _
                                & "JD.R_Desc as ReasonName,jc.CRMemo,jc.BillingName,JC.DrSpr_AcCode,jc.DrLab_AcCode,jc.labamt_tb, " _
                                & "jc.labamt_tp,jc.lab_d_amt,jc.lab_taxper,jc.lab_taxamt,jc.lab_roundoff,jc.netlab_amt,JC.DocId_InvSpr, " _
                                & "Jc.DocId_InvLab,JC.GP_No, JC.eCessPer, JC.eCessAmt, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name,HC.Add1, " _
                                & "HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName,jc.LabAmt_Out, " _
                                & "Amd_Dealer.D_Name,JC.LabD_Per, CreditCardNo, ChqNo, ChqDate, FreeWarrLabAmt, Closed_AddBy, Closed_AddDate, Closed_ModifyBy, Closed_ModifyDate, JC.Variation_Spare, JC.Variation_Labour " _
                        & "from ((((((job_card as JC Left Join Hiscard as HC on JC.CardNo=HC.CardNo) " _
                                & "Left Join Service_Type as ST on JC.Serv_Type=ST.Serv_Type) " _
                                & "Left Join City on HC.CityCode=City.CityCode) " _
                                & "left join Emp_Mast as EM on JC.Delby=EM.Emp_Code) " _
                                & "left join Emp_Mast as EMP on Jc.RecBy_Supervisor=Emp.Emp_Code) " _
                                & "Left Join Job_Delay as JD on JC.DelayReason=JD.Code) " _
                                & "Left Join Amd_Dealer on HC.Dealer_Code=Amd_Dealer.D_Code " _
                                & "where JC.DocId='" & Master!Code & "'")

    If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
        If AllowEditDel(pubUName, Master1!JobCloseDate, PubLoginDate) = False Then
            TopCtrl1.tDel = False
            TopCtrl1.tEdit = False
        Else
            TopCtrl1.tDel = True
            TopCtrl1.tEdit = True
        End If
    End If
        
        LblDiv.CAPTION = "Division : " & left(Master!Code, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        lblDocId.CAPTION = Master!Code
        
        SpareDocID = XNull(Master1!DocId_InvSpr)
        LabourDocID = XNull(Master1!DocID_InvLab)
        
        lblGatePass = XNull(Right(Master1!gp_no, 5))
        LblSprBill = DeCodeDocID(XNull(Master1!DocId_InvSpr), Document_No) ' 14, 8)
        lblLabourBill = DeCodeDocID(XNull(Master1!DocID_InvLab), Document_No)         ', 14, 8)
        
        lblSparePrefix = XNull(DeCodeDocID(XNull(Master1!DocId_InvSpr), Document_Prefix)) ', 9, 5)
        lblLabourPrefix = XNull(DeCodeDocID(XNull(Master1!DocID_InvLab), Document_Prefix)) ', 9, 5)
        If Master1!Govt_YN = 0 Then   'Govt = No
            mFormCode = pubLocalTaxFormSpr
        Else
            mFormCode = pubGovtTaxFormSpr
        End If
        
        LblUser = IIf(Not IsNull(Master1!Closed_AddDate), "Add By : " & XNull(Master1!Closed_AddBy) & "  Dated : " & XNull(Master1!Closed_AddDate), "") & IIf(Not IsNull(Master1!Closed_ModifyDate), "     Modify By : " & XNull(Master1!Closed_ModifyBy) & "  Dated : " & XNull(Master1!Closed_ModifyDate), "")
        txt(JobNo).Tag = Master!Code
        'txt(BodyDamage).TEXT = IIf(IsNull(Master1!body_damage), "", Master1!body_damage)
        txt(JobNo).TEXT = Master1!Job_No
        txt(JobDt).TEXT = Format(Master1!Job_Date, "dd/MMM/yyyy")

        txt(JobCDt).TEXT = Format(Master1!JobCloseDate, "dd/MMM/yyyy")
        
        
        
        
        mVatYn = PubVATYN
        If CDate(Master1!JobCloseDate) < CDate("01/Jan/2008") And StrCmp(left(PubComp_Name, 3), "Jmk") Then
            mVatYn = 0
        End If
        
        With FGrid
            If mVatYn = 1 Then
                .TextMatrix(0, Col_TaxPer) = "TaxPer"
                .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
                .ColWidth(Col_TaxPer) = 840
                
                .TextMatrix(0, Col_TaxAmt) = "TaxAmt"
                .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
                .ColWidth(Col_TaxAmt) = 840
                
                If PubSatYn = 1 Then
                    .TextMatrix(0, Col_SatPer) = "SAT %"
                    .ColAlignmentFixed(Col_SatPer) = flexAlignRightCenter
                    .ColWidth(Col_SatPer) = 840
                    
                    .TextMatrix(0, Col_SatAmt) = "SAT Amt"
                    .ColAlignmentFixed(Col_SatAmt) = flexAlignRightCenter
                    .ColWidth(Col_SatAmt) = 840
                Else
                    .ColWidth(Col_SatPer) = 0
                    .ColWidth(Col_SatAmt) = 0
                End If
            Else
                .TextMatrix(0, Col_TaxPer) = ""
                .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
                .ColWidth(Col_TaxPer) = 0
                
                .TextMatrix(0, Col_TaxAmt) = ""
                .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
                .ColWidth(Col_TaxAmt) = 0
                
                .ColWidth(Col_SatPer) = 0
                .ColWidth(Col_SatAmt) = 0
            End If
        End With
        
        
        txt(VehRegNo).TEXT = XNull(Master1!RegNo)
        txt(Chassis).TEXT = XNull(Master1!Chassis)
        mCardNo = Master1!CardNo
        
        txt(Model).TEXT = XNull(Master1!Model)
        txt(Engine).TEXT = XNull(Master1!Engine)
        txt(VehSrlNo).TEXT = XNull(Master1!VehSerialNo)
        txt(SrvType).TEXT = XNull(Master1!Serv_Desc)
        txt(OwnerName).TEXT = XNull(Master1!Name)
        txt(Address1).TEXT = XNull(Master1!Add1)
        txt(Address2).TEXT = XNull(Master1!Add2)
        txt(Address3).TEXT = XNull(Master1!Add3)
        txt(City).TEXT = XNull(Master1!CityName)
        txt(PhoneOff).TEXT = XNull(Master1!PhoneOff)
        txt(PhoneResi).TEXT = XNull(Master1!PhoneResi)
        txt(Mobile).TEXT = XNull(Master1!Mobile)
        txt(BookNo).TEXT = XNull(Master1!Job_BookNo)
        txt(BookDt).TEXT = ""
        txt(InspNo).TEXT = Trim(DeCodeDocID(XNull(Master1!Job_Inspdocid), Document_No))
        txt(GovtYn).TEXT = IIf(Master1!Govt_YN = 0, "No ", "Yes")
        txt(CurrentKMS).TEXT = IIf(IsNull(Master1!AtKMsHrs), "", Master1!AtKMsHrs)
        txt(CouponNo).TEXT = IIf(IsNull(Master1!Coupon), "", Master1!Coupon)
        txt(CouponVal).TEXT = IIf(IsNull(Master1!Coupon_Value), "", Master1!Coupon_Value)
        txt(ArrTime).TEXT = Format(Master1!ArrivalTime, "hh:mm")
        txt(DelDate).TEXT = Format(Master1!ExpDelDate, "dd/MMM/yyyy")
        txt(DelTime).TEXT = Format(Master1!ExpDelDate, "hh:mm")
        txt(EstSpare).TEXT = IIf(IsNull(Master1!Est_SpCost), "", Master1!Est_SpCost)
        txt(EstLabour).TEXT = IIf(IsNull(Master1!Est_LabCost), "", Master1!Est_LabCost)
        txt(OpenRemark).TEXT = XNull(Master1!OpenRemarks)
        txt(BodyDamage).TEXT = XNull(Master1!D_Name)
'special case
        txt(MechName).TEXT = XNull(Master1!Mechanic)
        txt(SuperName).TEXT = XNull(Master1!Supervisor)
        txt(JobDelay).TEXT = XNull(Master1!ReasonName)
        txt(JobCompDt).TEXT = Format(Master1!JobComp_Dt_Time, "dd/MMM/yyyy")
        txt(JobCompTm) = Format(XNull(Master1!JobComp_Dt_Time), "hh:mm")
        txt(CloseRemark).TEXT = XNull(Master1!Remark)
        txt(NextSrv).TEXT = XNull(Master1!NextSrvDate)
        txt(MechName).Tag = XNull(Master1!DelBy)
        txt(SuperName).Tag = XNull(Master1!RecBy_Supervisor)
        txt(JobDelay).Tag = XNull(Master1!DelayReason)
        txt(CreditCardNo) = XNull(Master1!CreditCardNo)
        txt(ChqNo) = XNull(Master1!ChqNo)
        txt(ChqDate) = XNull(Master1!ChqDate)
        '****
        txt(OutSideLabAmt) = IIf(IsNull(Master1!LabAmt_Out), "", Format(Master1!LabAmt_Out, "0.00"))
        txt(LabAmt) = Format(Master1!LabAmt_TB + Master1!LabAmt_TP - Master1!LabAmt_Out, "0.00")
        txt(LabAmtTB) = Format(Master1!LabAmt_TB, "0.00")
        txt(LabAmtTP) = Format(Master1!LabAmt_TP, "0.00")
        txt(FreeWarrLabAmt) = Format(VNull(Master1!FreeWarrLabAmt), "0.00")
        '****
        txt(LabDisc) = Format(Master1!Lab_D_Amt, "0.00")
        txt(LabDisPer) = Format(Master1!LabD_Per, "0.0000")
        txt(ServTaxPer) = Format(Master1!Lab_TaxPer, "0.00")
        txt(ServTaxAmt) = Format(Master1!Lab_TaxAmt, "0.00")
        txt(eCessPer) = Format(Master1!eCessPer, "0.00")
        txt(eCessAmt) = Format(Master1!eCessAmt, "0.00")
        
        txt(LabROff) = Format(Master1!Lab_RoundOff, "0.00")
        txt(NetLabAmt) = Format(Master1!NetLab_Amt, "0.00")
        
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select D_Per_TB,D_Per_TP,D_Amt_TB,D_Amt_TP,Addition,Packing,Gen_Sur_Per,Gen_Sur_Amt,Trans_Amt,Tax_Per,Tax_Amt, " _
                & "Tax_Sur_Per,Tax_Sur_Amt,TOT_Per,Tot_Amt,Rounded,ReSalTax_Per,ReSalTax_Amt,Total_Amt,OilAmt_MRP_TB,OilAmt_MRP_TP,Excise_Amt from Sp_Sale where DocId='" & Master1!DocId_InvSpr & "'", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            txt(DiscPerTB) = IIf(IsNull(Rst!D_Per_TB), "", Format(Rst!D_Per_TB, "0.0000"))
            txt(DiscPerTP) = IIf(IsNull(Rst!D_Per_TP), "", Format(Rst!D_Per_TP, "0.0000"))
            txt(DiscAmtTB) = IIf(IsNull(Rst!D_Amt_TB), "", Format(Rst!D_Amt_TB, "0.00"))
            txt(Excise_Amt) = IIf(IsNull(Rst!Excise_Amt), "", Format(Rst!Excise_Amt, "0.00"))
            
            txt(DiscAmtTP) = IIf(IsNull(Rst!D_Amt_TP), "", Format(Rst!D_Amt_TP, "0.00"))
            txt(Addition) = IIf(IsNull(Rst!Addition), "", Format(Rst!Addition, "0.00"))
            txt(PackCrg) = IIf(IsNull(Rst!Packing), "", Format(Rst!Packing, "0.00"))
            txt(GenSurPer) = IIf(IsNull(Rst!Gen_Sur_Per), "", Format(Rst!Gen_Sur_Per, "0.00"))
            txt(GenSurAmt) = IIf(IsNull(Rst!Gen_Sur_Amt), "", Format(Rst!Gen_Sur_Amt, "0.00"))
            txt(TransAmt) = IIf(IsNull(Rst!Trans_Amt), "", Format(Rst!Trans_Amt, "0.00"))
            txt(STaxPer) = IIf(IsNull(Rst!Tax_Per), "", Format(Rst!Tax_Per, "0.00"))
            txt(STaxAmt) = IIf(IsNull(Rst!Tax_Amt), "", Format(Rst!Tax_Amt, "0.00"))
            txt(TaxSurPer) = IIf(IsNull(Rst!Tax_Sur_Per), "", Format(Rst!Tax_Sur_Per, "0.00"))
            txt(TaxSurAmt) = IIf(IsNull(Rst!Tax_Sur_Amt), "", Format(Rst!Tax_Sur_Amt, "0.00"))
            txt(TurnOverPer) = IIf(IsNull(Rst!TOT_Per), "", Format(Rst!TOT_Per, "0.00"))
            txt(TurnOverAmt) = IIf(IsNull(Rst!Tot_Amt), "", Format(Rst!Tot_Amt, "0.00"))
            txt(SROff) = IIf(IsNull(Rst!Rounded), "", Format(Rst!Rounded, "0.00"))
            txt(ReSalTaxPer) = IIf(IsNull(Rst!ReSalTax_Per), "", Format(Rst!ReSalTax_Per, "0.00"))
            txt(ReSalTaxAmt) = IIf(IsNull(Rst!ReSalTax_Amt), "", Format(Rst!ReSalTax_Amt, "0.00"))
            txt(NetSprAmt) = IIf(IsNull(Rst!Total_Amt), "", Format(Rst!Total_Amt, "0.00"))
            mMRPLubeTB = IIf(IsNull(Rst!OilAmt_MRP_TB), 0, Rst!OilAmt_MRP_TB)
            mMRPLubeTP = IIf(IsNull(Rst!OilAmt_MRP_TP), 0, Rst!OilAmt_MRP_TP)
        Else
            txt(DiscPerTB) = ""
            txt(DiscPerTP) = ""
            txt(DiscAmtTB) = ""
            txt(Excise_Amt) = ""
            txt(DiscAmtTP) = ""
            txt(Addition) = ""
            txt(PackCrg) = ""
            txt(GenSurPer) = ""
            txt(GenSurAmt) = ""
            txt(TransAmt) = ""
            txt(STaxPer) = ""
            txt(STaxAmt) = ""
            txt(TaxSurAmt) = ""
            txt(TaxSurPer) = ""
            txt(TurnOverPer) = ""
            txt(TurnOverAmt) = ""
            txt(SROff) = ""
            txt(ReSalTaxPer) = ""
            txt(ReSalTaxAmt) = ""
            txt(NetSprAmt) = ""
        End If
        Set Rst = Nothing
        Call Fill_Grid
        txt(IWDiscPerTB) = ""
        txt(IWDiscPerTP) = ""

        txt(CashBill) = IIf(Master1!CrMemo = 0, "Yes", "No")
        mVType = IIf(Master1!CrMemo = 0, "W_SIC", "W_SIR")
        
        txt(CashParty) = IIf(IsNull(Master1!BillingName), "", Master1!BillingName)
        txt(SpareParty).Tag = IIf(IsNull(Master1!DrSpr_AcCode), "", Master1!DrSpr_AcCode)
        txt(LabourParty).Tag = IIf(IsNull(Master1!DrLab_AcCode), "", Master1!DrLab_AcCode)
        If txt(CashBill) <> "Yes" Then
            If GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrSpr_AcCode & "'").RecordCount > 0 Then
                txt(SpareParty).TEXT = GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrSpr_AcCode & "'").Fields(0).Value
                CurrBal = GCn.Execute("select " & vIsNull("Curr_Bal", "0") & " From Subgroup where subcode='" & Master1!DrSpr_AcCode & "'").Fields(0).Value
                LblCurrBal = "Bal. " & Format(Abs(CurrBal), "0.00")
                LblCurrBal = LblCurrBal & IIf(CurrBal > 0, " Cr", IIf(CurrBal < 0, " Dr", ""))
            End If
            If GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrLab_AcCode & "'").RecordCount > 0 Then
                txt(LabourParty).TEXT = GCn.Execute("select Name From Subgroup where subcode='" & Master1!DrLab_AcCode & "'").Fields(0).Value
                CurrBal = GCn.Execute("select " & vIsNull("Curr_Bal", "0") & " From Subgroup where subcode='" & Master1!DrLab_AcCode & "'").Fields(0).Value
                LblCurrBal1 = "Bal. " & Format(Abs(CurrBal), "0.00")
                LblCurrBal1 = LblCurrBal1 & IIf(CurrBal > 0, " Cr", IIf(CurrBal < 0, " Dr", ""))
            End If
        Else
            txt(SpareParty).Tag = ""
            txt(LabourParty).Tag = ""
            txt(SpareParty) = ""
            txt(LabourParty) = ""
        End If
        Call veh_count
        txt(DiscPerTB).Enabled = False
        txt(DiscPerTP).Enabled = False
        txt(DiscAmtTB).Enabled = False
        txt(DiscAmtTP).Enabled = False
        txt(Addition).Enabled = False
        txt(PackCrg).Enabled = False
        txt(GenSurPer).Enabled = False
        txt(GenSurAmt).Enabled = False
        txt(TransAmt).Enabled = False
        txt(STaxPer).Enabled = False
        txt(STaxAmt).Enabled = False
        txt(TaxSurAmt).Enabled = False
        txt(TaxSurPer).Enabled = False
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
        .top = txt(Model).top   '2685
        .height = 2745
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 26

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
        .ColWidth(Col_MRPRate) = 850

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
        .ColWidth(Col_DiscPer) = 510

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
        .ColWidth(Col_DiscAmt) = 840

        If mVatYn = 1 Then
            .TextMatrix(0, Col_TaxPer) = "TaxPer"
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 840
            
            .TextMatrix(0, Col_TaxAmt) = "TaxAmt"
            .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt) = 840
            
            If PubSatYn = 1 Then
                .TextMatrix(0, Col_SatPer) = "SAT %"
                .ColAlignmentFixed(Col_SatPer) = flexAlignRightCenter
                .ColWidth(Col_SatPer) = 840
                
                .TextMatrix(0, Col_SatAmt) = "SAT Amt"
                .ColAlignmentFixed(Col_SatAmt) = flexAlignRightCenter
                .ColWidth(Col_SatAmt) = 840
            Else
                .ColWidth(Col_SatPer) = 0
                .ColWidth(Col_SatAmt) = 0
            End If
        Else
            
            .TextMatrix(0, Col_TaxPer) = ""
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 0
            
            .TextMatrix(0, Col_TaxAmt) = ""
            .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt) = 0
            
            .ColWidth(Col_SatPer) = 0
            .ColWidth(Col_SatAmt) = 0
        End If
        
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
        .Cols = 17
        
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
        
        .TextMatrix(0, C_ContAcCode) = ""
        .ColAlignment(C_ContAcCode) = flexAlignLeftCenter
        .ColWidth(C_ContAcCode) = 0

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
    'DGParty.width = 4740: DGParty.left = MeWidth - (DGParty.width + mRtScale): DGParty.top = mTopScale: DGParty.height = 5000
    DGParty.width = 8700: DGParty.left = 1000: DGParty.top = mTopScale: DGParty.height = 4000
    'DGParty.Columns(1).width = 0
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next
    If Not StrCmp(left(PubComp_Name, 3), "JMK") And Not StrCmp(left(PubComp_Name, 7), "Singhal") Then
        txt(IWDiscPerTB).Visible = False
        txt(IWDiscPerTP).Visible = False
    End If
    If PubSiebelActiveYn = 1 And pubUName = "SA" Then
        cmdPost.Visible = True
    Else
        cmdPost.Visible = False
    End If
            
    
    
End Sub

Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
    If DGMech.Visible = True Then DGMech.Visible = False
    If DGReason.Visible = True Then DGReason.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
End Sub

Private Sub veh_count()
    If txt(JobDt).TEXT <> "" Then
        LblTotVeh.CAPTION = GCn.Execute("select count(*) from job_Card where JobCloseDate=Null or JobCloseDate Is Null and left(Docid,1)='" & PubDivCode & "'").Fields(0)
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
Dim lblstring As String, Bal As Double
If DGParty.Row >= 0 Then
    lblstring = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    Bal = Abs(VNull(G_FaCn.Execute("Select Sum(AmtDr)-Sum(AmtCr) from Ledger where SubCode='" & RsParty!Code & "'").Fields(0).Value))
    If Bal > 0 Then
        lblGroup.TEXT = lblstring & "  |  " & Format(Bal, "0.00") & " Dr.  "
    ElseIf Bal < 0 Then
        lblGroup.TEXT = lblstring & "  |  " & Format(Bal, "0.00") & " Cr.  "
    End If
    lblGroup.Refresh
End If

End Sub
Private Sub History_Field()
Dim rsForm As ADODB.Recordset, rsJob2 As ADODB.Recordset
    Set rsJob2 = GCn.Execute("select J.CardNo,J.Job_Date,J.Govt_YN,City.CityName,ST.Serv_Desc,HC.Add1,HC.Add2,HC.Add3,HC.PhoneOff,HC.PhoneResi,HC.Mobile, " _
        & "J.RecBy_Mechanic,J.RecBy_Supervisor,J.DelBy,J.Job_BookNo,J.Job_Inspdocid,J.ATKMSHRS,J.Coupon,J.ArrivalTime,J.ExpDelDate," _
        & "J.Est_SpCost,J.Est_LabCost,J.OpenRemarks,J.Body_Damage,E.Emp_Name as Mechanic,E1.Emp_Name as Supervisor,E2.Emp_Name as DelByMechanic,Amd_Dealer.D_Name, J.Serv_Type " _
        & "from (((((((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) " _
        & "left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) " _
        & "Left Join City on HC.CityCode=City.CityCode) " _
        & "Left Join Emp_Mast as E on J.RecBy_Mechanic=E.Emp_Code) " _
        & "Left Join Emp_Mast as E1 on J.RecBy_Supervisor=E1.Emp_Code) " _
        & "Left Join Emp_Mast as E2 on J.DelBy=E2.Emp_Code) " _
        & "Left Join Amd_Dealer  on Amd_Dealer.D_Code=HC.Dealer_Code) " _
        & "where J.DocId='" & RsJob!Code & "'")
    
    txt(VehRegNo).Tag = RsJob!Code
    txt(Chassis).Tag = RsJob!Code
    txt(OwnerName).Tag = RsJob!Code
    txt(JobNo).Tag = RsJob!Code
    txt(JobNo).TEXT = IIf(IsNull(RsJob!Job_No), "", RsJob!Job_No)
    txt(JobDt).TEXT = rsJob2!Job_Date
    txt(VehRegNo).TEXT = IIf(IsNull(RsJob!RegNo), "", RsJob!RegNo)
    txt(Chassis).TEXT = IIf(IsNull(RsJob!Chassis), "", RsJob!Chassis)
    txt(Model).TEXT = IIf(IsNull(RsJob!Model), "", RsJob!Model)
    txt(Engine).TEXT = IIf(IsNull(RsJob!Engine), "", RsJob!Engine)
    txt(VehSrlNo).TEXT = IIf(IsNull(RsJob!VehSerialNo), "", RsJob!VehSerialNo)
    txt(OwnerName).TEXT = IIf(IsNull(RsJob!Name), "", RsJob!Name)
    txt(Address1).TEXT = IIf(IsNull(rsJob2!Add1), "", rsJob2!Add1)
    txt(Address2).TEXT = IIf(IsNull(rsJob2!Add2), "", rsJob2!Add2)
    txt(Address3).TEXT = IIf(IsNull(rsJob2!Add3), "", rsJob2!Add3)
    txt(City).TEXT = IIf(IsNull(rsJob2!CityName), "", rsJob2!CityName)
    txt(PhoneOff).TEXT = IIf(IsNull(rsJob2!PhoneOff), "", rsJob2!PhoneOff)
    txt(PhoneResi).TEXT = IIf(IsNull(rsJob2!PhoneResi), "", rsJob2!PhoneResi)
    txt(Mobile).TEXT = IIf(IsNull(rsJob2!Mobile), "", rsJob2!Mobile)
    txt(MechName).Tag = rsJob2!RecBy_Mechanic
    txt(MechName) = IIf(IsNull(rsJob2!Mechanic), "", rsJob2!Mechanic)
    txt(SuperName).Tag = rsJob2!RecBy_Supervisor
    txt(SuperName) = IIf(IsNull(rsJob2!Supervisor), "", rsJob2!Supervisor)
'special case
'    Txt(MechName).Tag = rsJob2!DelBy
'    Txt(MechName) = IIf(IsNull(rsJob2!DelByMechanic), "", rsJob2!DelByMechanic)
    txt(SrvType).TEXT = IIf(IsNull(rsJob2!Serv_Desc), "", rsJob2!Serv_Desc)
    txt(SrvType).Tag = XNull(rsJob2!Serv_Type)
    txt(BookNo).TEXT = IIf(IsNull(rsJob2!Job_BookNo), "", rsJob2!Job_BookNo)
    txt(BookDt).TEXT = ""
    txt(InspNo).TEXT = Trim(DeCodeDocID(XNull(rsJob2!Job_Inspdocid), Document_No))
    txt(GovtYn).TEXT = IIf(rsJob2!Govt_YN = 0, "No ", "Yes")
    txt(CurrentKMS).TEXT = rsJob2!AtKMsHrs
    txt(CouponNo).TEXT = XNull(rsJob2!Coupon)
    txt(ArrTime).TEXT = Format(rsJob2!ArrivalTime, "hh:mm")
    txt(DelDate).TEXT = Format(rsJob2!ExpDelDate, "dd/MMM/yyyy")
    txt(DelTime).TEXT = Format(rsJob2!ExpDelDate, "hh:mm")
    txt(EstSpare).TEXT = rsJob2!Est_SpCost
    txt(EstLabour).TEXT = rsJob2!Est_LabCost
    txt(OpenRemark).TEXT = IIf(IsNull(rsJob2!OpenRemarks), "", rsJob2!OpenRemarks)
    txt(BodyDamage).TEXT = XNull(rsJob2!D_Name)
    
    txt(GenSurPer) = Format(PubGenSurChrgOnSpr, "0.00")
    txt(TurnOverPer) = Format(PubTOT_Rate, "0.00")
    
    txt(CashBill) = "Yes"
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
        txt(STaxPer).TEXT = IIf(IsNull(rsForm!Tax_Per), "", Format(rsForm!Tax_Per, "0.00"))
        txt(TaxSurPer).TEXT = IIf(IsNull(rsForm!Tax_Sur_Per), "", Format(rsForm!Tax_Sur_Per, "0.00"))
    Else
        MsgBox "Please Add/Define Local/Govt Tax Form in " & vbCrLf & " Tax Forms/System Controls", vbOKOnly, "Control Validation"
    End If
    Set rsForm = Nothing
    Set rsJob2 = Nothing
End Sub

Private Sub txtDisabled_Color()
Dim I As Integer
    For I = 0 To txt.Count - 1
        If txt(I).Enabled = False Or txt(I).Locked = True Then
            txt(I).BackColor = &HEBF0F1
        Else
            txt(I).BackColor = CtrlBColOrg
        End If
    Next I
End Sub
Private Sub Fill_Grid()
Dim mQry$
Dim I As Integer, TmpStr$, RsTemp As ADODB.Recordset
    If txt(JobNo).Tag = "" Then Exit Sub
    FreeLabForTax = 0
    mTBDisAmtMRP = 0
    mTPDisAmtMRP = 0
'' Spares Details
    FGrid.Rows = 1
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    GSQL = "Select SPS.Job_DocId,SPS.DocId,SPS.V_Date,SPS.V_type, SPS.V_no, Sps.Srl_no, SPS.Part_No,SPS.Lub_Category,(SPS.Qty_iss-SPS.Qty_ret) as ReqQty, SPS.Tax_yn,SPS.Mrp_YN,SPS.Rate,SPS.MRP_Rate,SPS.Amount,SPS.Disc_per,SPS.Disc_Amt,SPS.Net_Amt,SPS.Purpose,SPS.TrnComplete_YN,SPS.Claim_No,Part.Part_Name,Part.Local_Name,part.Unit,Part.Part_Grade,SPS.TaxPer,SPS.TaxAmt,SPS.SatPer, SPS.SatAmt " & _
            "FROM SP_Stock AS SPS " & _
            "left join part on SpS.part_no=part.Part_No and Part.Div_Code = left(SPS.DocID,1) " & _
            "Where SPS.Job_DocId='" & txt(JobNo).Tag & "' Order By SPS.V_No,SPS.Srl_No"
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    I = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, Col_SrNo) = I
                .TextMatrix(I, Col_PNo) = Rst!Part_No
                .TextMatrix(I, Col_ReqNoDocId) = XNull(Rst!DocID)
                .TextMatrix(I, Col_ReqNo) = Rst!V_NO
                .TextMatrix(I, Col_ReqDate) = Rst!V_DATE
                .TextMatrix(I, Col_ReqSrNo) = Rst!Srl_No
                .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, Col_Qty) = Format(Rst!ReqQty, "0.00")
                .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")
                If Rst!Purpose = "P" Then
                    TmpStr = "PDI"
                ElseIf Rst!Purpose = "F" Then
                    TmpStr = "Free Service"
                ElseIf Rst!Purpose = "C" Or (Rst!Purpose = "A" And StrCmp(left(PubComp_Name, 4), "Enar")) Then
                    If StrCmp(left(PubComp_Name, 4), "Enar") Then
                        TmpStr = IIf(Rst!Purpose = "C", "Charge", "AMC")
                    Else
                        TmpStr = "Charge"
                    End If
                    .TextMatrix(I, Col_Amt) = Format(Rst!Amount, "0.00")
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.0000")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                    .TextMatrix(I, Col_TaxPer) = Format(Rst!TaxPer, "0.00")
                    .TextMatrix(I, Col_TaxAmt) = Format(Rst!TaxAmt, "0.00")
                    .TextMatrix(I, Col_SatPer) = Format(Rst!SatPer, "0.00")
                    .TextMatrix(I, Col_SatAmt) = Format(Rst!SatAmt, "0.00")
                    
                    .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt, "0.00")
                ElseIf Rst!Purpose = "W" Or (Rst!Purpose = "A" And Not StrCmp(left(PubComp_Name, 4), "Enar")) Then
                    If Not StrCmp(left(PubComp_Name, 4), "Enar") Then
                        TmpStr = IIf(Rst!Purpose = "A", "AMC", "Warranty")
                    Else
                        TmpStr = "Warranty"
                    End If
                    If UCase(left(PubComp_Name, 3)) = "LMP" Then
                        .TextMatrix(I, Col_TaxPer) = Format(Rst!TaxPer, "0.00")
                        .TextMatrix(I, Col_TaxAmt) = Format(Rst!TaxAmt, "0.00")
                        .TextMatrix(I, Col_SatPer) = Format(Rst!SatPer, "0.00")
                        .TextMatrix(I, Col_SatAmt) = Format(Rst!SatAmt, "0.00")
                        
                        .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt, "0.00")
                    End If
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
    
    mQry = "Select Distinct Disc_Per " & _
           "FROM SP_Stock AS SPS " & _
           "left join part on SpS.part_no=part.Part_No and Part.Div_Code = left(SPS.DocID,1) " & _
           "Where SPS.Job_DocId='" & txt(JobNo).Tag & "'"
    
    If GCn.Execute(mQry).RecordCount > 1 Then
        If StrCmp(left(PubComp_Name, 3), "JMK") Then
            txt(IWDiscPerTB).Enabled = False
            txt(IWDiscPerTP).Enabled = False
        End If
    End If
    
    '' Labour Details
    Rst.Close
    FGrid1.Rows = 1
                
    GSQL = "Select JL.*, L.Lab_Desc AS LabName,CF.FinName AS ContName,GP.GatePassDate,GP.ContractRecdDate,GP.ContractAmt,GP.ContractCode,CF.AcCode as ContractSubCode " & _
        " From ((((Job_Lab as JL left join labour as L on JL.Lab_Code=L.Lab_Code) " & _
        " Left Join Labour_Model LM on JL.Lab_Code=LM.Lab_Code) " & _
        " left join Job_GatePass as GP on JL.ExtJobGatePassNo=GP.GatePassNo) " & _
        " Left Join ContractFinance as CF ON GP.ContractCode=CF.FinCode) " & _
        " Where JL.Job_DocId='" & txt(JobNo).Tag & "' Order by JL.S_No"
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    I = 1
    txt(OutSideLabAmt) = ""
    txt(LabAmt) = ""
    txt(LabAmtTB) = ""
    txt(LabAmtTP) = ""
    txt(FreeWarrLabAmt) = ""
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_LabCode) = Rst!Lab_Code
                .TextMatrix(I, C_LabName) = XNull(Rst!LabName)
                .TextMatrix(I, C_TaxYN) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, C_PaidBy) = XNull(Rst!Chrg_From)
'                If Rst!Hrs_Taken + Rst!Lab_Rate  > 0 Then
                If Rst!Chrg_From = "M" Or Rst!Chrg_From = "O" Then
                    If Rst!Chrg_Type = "W" Then 'Warranty
                        .TextMatrix(I, C_ChrgType) = "Warranty"
                        .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_War = 0, "", Format(Rst!Hrs_War, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(Rst!War_Lab_Rate = 0, "", Format(Rst!War_Lab_Rate, "0.00"))
                        If PubTaxOnFreeLabYn = 1 Then
                            .TextMatrix(I, C_Amt) = IIf(Rst!LabourAmt = 0, "", Format(Rst!LabourAmt, "0.00"))
                        End If
                    ElseIf Rst!Chrg_Type = "P" Then 'PDI
                        .TextMatrix(I, C_ChrgType) = "PDI"
                        .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_Taken = 0, "", Format(Rst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(Rst!Lab_Rate = 0, "", Format(Rst!Lab_Rate, "0.00"))
                        If PubTaxOnFreeLabYn = 1 Then
                            .TextMatrix(I, C_Amt) = IIf(Rst!LabourAmt = 0, "", Format(Rst!LabourAmt, "0.00"))
                        End If
                        If UCase(left(PubComp_Name, 5)) = "SOCIE" And Rst!Tax_YN = 1 Then
                            FreeLabForTax = FreeLabForTax + Rst!LabourAmt
                        End If
                    ElseIf Rst!Chrg_Type = "A" And Not StrCmp(left(PubComp_Name, 4), "eNAR") Then 'Free Service
                        .TextMatrix(I, C_ChrgType) = "AMC"
                        .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_Taken = 0, "", Format(Rst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(Rst!Lab_Rate = 0, "", Format(Rst!Lab_Rate, "0.00"))
                        If PubTaxOnFreeLabYn = 1 Then
                            .TextMatrix(I, C_Amt) = IIf(Rst!LabourAmt = 0, "", Format(Rst!LabourAmt, "0.00"))
                        End If
                        If UCase(left(PubComp_Name, 5)) = "SOCIE" And Rst!Tax_YN = 1 Then
                                FreeLabForTax = FreeLabForTax + Rst!LabourAmt
                        End If
                        
                    Else    'Free Service
                        .TextMatrix(I, C_ChrgType) = "Free Service"
                        .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_Taken = 0, "", Format(Rst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(Rst!Lab_Rate = 0, "", Format(Rst!Lab_Rate, "0.00"))
                        If PubTaxOnFreeLabYn = 1 Then
                            .TextMatrix(I, C_Amt) = IIf(Rst!LabourAmt = 0, "", Format(Rst!LabourAmt, "0.00"))
                        End If
                        If UCase(left(PubComp_Name, 5)) = "SOCIE" And Rst!Tax_YN = 1 Then
                                FreeLabForTax = FreeLabForTax + Rst!LabourAmt
                        End If

                    End If
                Else
                    .TextMatrix(I, C_ChrgType) = "Chargeable"
                    .TextMatrix(I, C_Hrs) = IIf(Rst!Hrs_Taken = 0, "", Format(Rst!Hrs_Taken, "0.00"))
                    .TextMatrix(I, C_Rate) = IIf(Rst!Lab_Rate = 0, "", Format(Rst!Lab_Rate, "0.00"))
                    .TextMatrix(I, C_Amt) = IIf(Rst!LabourAmt = 0, "", Format(Rst!LabourAmt, "0.00"))
                     'FreeLabForTax = FreeLabForTax + Rst!LabourAmt

                End If
                
                .TextMatrix(I, C_External) = IIf(Rst!External_yn = "1", "Yes", "No")
                .TextMatrix(I, C_GPNo) = XNull(Rst!ExtJobGatePassNo)
                .TextMatrix(I, C_ContName) = XNull(Rst!ContName)
                .TextMatrix(I, C_ContAcCode) = XNull(Rst!ContractSubCode)
                .TextMatrix(I, C_WIssueDt) = IIf(IsNull(Rst!GatePassDate), "", Rst!GatePassDate)
                .TextMatrix(I, C_WRecdDt) = IIf(IsNull(Rst!ContractRecdDate), "", Rst!ContractRecdDate)
                .TextMatrix(I, C_ContAmt) = IIf(Rst!ContractAmt = 0, "", Format(Rst!ContractAmt, "0.00"))
                .TextMatrix(I, C_Remarks) = XNull(Rst!Contract_Remarks)
'                .TextMatrix(i, C_ContCode) = XNull(Rst!ContractCode)
                ' NRA MODI FOR TRUPTI MOTORS
'                If UCase(left(PubComp_Name, 6)) = "TRUPTI" Then
'                        If Rst!CHRG_FROM = "C" Then
'                            If Rst!External_yn = "1" Then
'                                Txt(OutSideLabAmt) = Format(Val(Txt(OutSideLabAmt)) + Rst!LabourAmt, "0.00")
'                            Else
'                                Txt(LabAmt) = Format(Val(Txt(LabAmt)) + Rst!LabourAmt, "0.00")
'                                If Rst!Tax_YN = 1 Then
'                                    Txt(LabAmtTB) = Format(Val(Txt(LabAmtTB)) + Rst!LabourAmt, "0.00")
'                                Else
'                                    Txt(LabAmtTP) = Format(Val(Txt(LabAmtTP)) + Rst!LabourAmt, "0.00")
'                                End If
'                           End If
'                        End If
'                Else
                    If Rst!Chrg_From = "C" Then
                        If Rst!External_yn = "1" Then
                            txt(OutSideLabAmt) = Format(Val(txt(OutSideLabAmt)) + Rst!LabourAmt, "0.00")
                        Else
                            txt(LabAmt) = Format(Val(txt(LabAmt)) + Rst!LabourAmt, "0.00")
                        End If
                        If Rst!Tax_YN = 1 Then
                            txt(LabAmtTB) = Format(Val(txt(LabAmtTB)) + Rst!LabourAmt, "0.00")
                        Else
                            txt(LabAmtTP) = Format(Val(txt(LabAmtTP)) + Rst!LabourAmt, "0.00")
                        End If
                    ElseIf PubTaxOnFreeLabYn = 1 Then
                        If Rst!Tax_YN = 1 Then
                            txt(FreeWarrLabAmt) = Format(Val(txt(FreeWarrLabAmt)) + Rst!LabourAmt, "0.00")
                        End If
                    End If
                    
'                End If
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
                txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt), Col_Purpose, True
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
    End If
    MainLib.LabCalc txt(LabAmtTB), txt(LabAmtTP), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt), mLabDiscAmtTB, mECessPer, mECessAmt, txt(FreeWarrLabAmt), mServiceTaxPer_Saperate, mServiceTaxAmt_Saperate, mHECessPer, mHECessAmt
    If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
        txt(ServTaxPer) = MainLib.Serv_Tax
        txt(ServTaxAmt).TEXT = Format((Val(txt(LabAmtTB).TEXT) + Val(FreeLabForTax) - Val(txt(LabDisc))) * Val(txt(ServTaxPer)) / 100, "0.00")
        txt(NetLabAmt).TEXT = Val(txt(LabAmtTB).TEXT) + Val(txt(LabAmtTP).TEXT) + Val(txt(ServTaxAmt).TEXT) - Val(txt(LabDisc).TEXT)
        txt(LabROff).TEXT = Format(dmRoundOff(txt(NetLabAmt)), "0.00")
        txt(NetLabAmt).TEXT = Format(Val(txt(NetLabAmt)) + Val(txt(LabROff)), "0.00")
        
    End If
'    If UCase(left(PubComp_Name, 3)) = "JMK" Then
'        txt(TurnOverAmt) = Val(txt(STotATB)) * Val(txt(TurnOverPer)) / 100
'        txt(NetSprAmt).TEXT = Format(Val(txt(NetSprAmt)) + Val(txt(TurnOverAmt)), "0.00")
'        txt(SROff).TEXT = Format(Val(txt(NetSprAmt).TEXT) - Round(txt(NetSprAmt).TEXT, 0), "0.00")
'        txt(NetSprAmt).TEXT = Format(Round(txt(NetSprAmt).TEXT, 0), "0.00")
'        txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
'    Else
'        txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
'    End If
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        txt(TurnOverAmt) = Format((Val(txt(STotATB)) + Val(txt(STaxAmt))) * Val(txt(TurnOverPer)) / 100, "0.00")
        txt(NetSprAmt).TEXT = Format(Val(txt(STotB)) + Val(txt(TurnOverAmt)), "0.00")
        'txt(SROff).TEXT = Format(Val(txt(NetSprAmt).TEXT) - Round(Val(txt(STotB).TEXT) + Val(txt(TurnOverAmt)), 0), "0.00")
        txt(NetSprAmt).TEXT = Format(Round(txt(NetSprAmt).TEXT, 0), "0.00")
        txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
    Else
        txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
    End If

    
    txt(NetAmt) = Format(Round(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), 0), "0.00")
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
Dim ZeroBillStopYn As Integer, InvSuffix As Integer, InvLabSuffix As Integer
Dim InvdocId, LabDocId As String
Dim CrYn As Integer
Dim ServType As String
ZeroBillStopYn = VNull(GCn.Execute("Select ZeroBill from syctrl").Fields(0).Value)
If ZeroBillStopYn = 1 Then
    If txt(CashBill).TEXT = "Yes" Then
        txt(CashParty).Enabled = True
        txt(SpareParty).Enabled = False
        txt(LabourParty).Enabled = False
        txt(SpareParty).TEXT = ""
        txt(LabourParty).TEXT = ""
        txt(CreditCardNo) = ""
        txt(ChqNo) = ""
        txt(ChqDate) = ""
        txt(CreditCardNo).Enabled = False
        txt(ChqNo).Enabled = False
        txt(ChqDate).Enabled = False

        If Val(txt(NetSprAmt)) = 0 Then
            mVType = "W_WWC"
        Else
            mVType = "W_SIC"
        End If
        If Val(txt(NetLabAmt)) = 0 Then
            LabourVtype = "W_WLC"
        Else
            LabourVtype = "W_LIC"
        End If
        txt(SpareParty).Tag = PubSprCashAc
        txt(LabourParty).Tag = PubSrvLabAc
    Else
        txt(CashParty).Enabled = False
        txt(CashParty).TEXT = ""
        txt(CashParty).Tag = ""
        
        txt(CreditCardNo).Enabled = True
        txt(ChqNo).Enabled = True
        txt(ChqDate).Enabled = True
        
        If txt(CreditCardNo) = "" And txt(ChqNo) = "" Then
            txt(SpareParty).Enabled = True
            txt(LabourParty).Enabled = True
        Else
            txt(SpareParty).Enabled = False
            txt(LabourParty).Enabled = False
        End If
        
        If Val(txt(NetSprAmt)) = 0 Then
            mVType = "W_WWR"
        Else
            mVType = "W_SIR"
        End If
        
        If Val(txt(NetLabAmt)) = 0 Then
            LabourVtype = "W_WLR"
        Else
            LabourVtype = "W_LIR"
        End If
        
    End If
    SpareVtype = mVType
    If UCase(left(PubComp_Name, 3)) = "JMK" Or RSOJPR = True Then
            InvdocId = XNull(GCn.Execute("Select LastInvDocId from Job_Card where DocId='" & txt(JobNo).Tag & "'").Fields(0).Value)
            LabDocId = XNull(GCn.Execute("Select LastLabInvDocId from Job_Card where DocId='" & txt(JobNo).Tag & "'").Fields(0).Value)
            If InvdocId <> "" And LabDocId <> "" Then  'CrYn = IIf(txt(CashBill) = "Yes", 0, 1) Then
                SpareDocID = InvdocId
                LabourDocID = LabDocId
            Else
                SpareDocID = GetDocID(GCnFaS, SpareVtype, txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
                LabourDocID = GetDocID(GCnFaW, LabourVtype, txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
                LblSprBill.Refresh
                lblLabourBill.Refresh
            End If
    Else
        SpareDocID = GetDocID(GCnFaS, SpareVtype, txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
        LabourDocID = GetDocID(GCnFaW, LabourVtype, txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
        LblSprBill.Refresh
        lblLabourBill.Refresh
    End If
Else
    If txt(CashBill).TEXT = "Yes" Then
        txt(CashParty).Enabled = True
        txt(SpareParty).Enabled = False
        txt(LabourParty).Enabled = False
        txt(SpareParty).TEXT = ""
        txt(LabourParty).TEXT = ""
        txt(CreditCardNo) = ""
        txt(ChqNo) = ""
        txt(ChqDate) = ""
        txt(CreditCardNo).Enabled = False
        txt(ChqNo).Enabled = False
        txt(ChqDate).Enabled = False
        
        LabourVtype = "W_LIC"
        mVType = "W_SIC"
        txt(SpareParty).Tag = PubSprCashAc
        txt(LabourParty).Tag = PubSrvLabAc
    Else
        txt(CashParty).Enabled = False
        txt(CashParty).TEXT = ""
        txt(CashParty).Tag = ""
        txt(CreditCardNo).Enabled = True
        txt(ChqNo).Enabled = True
        txt(ChqDate).Enabled = True
        
        If txt(CreditCardNo) = "" And txt(ChqNo) = "" Then
            txt(SpareParty).Enabled = True
            txt(LabourParty).Enabled = True
        Else
            txt(SpareParty).Enabled = False
            txt(LabourParty).Enabled = False
        End If
        LabourVtype = "W_LIR"
        mVType = "W_SIR"
    End If
    SpareVtype = mVType
    If UCase(left(PubComp_Name, 3)) = "JMK" Or RSOJPR = True Then
            InvdocId = XNull(GCn.Execute("Select LastInvDocId from Job_Card where DocId='" & txt(JobNo).Tag & "'").Fields(0).Value)
            LabDocId = XNull(GCn.Execute("Select LastLabInvDocId from Job_Card where DocId='" & txt(JobNo).Tag & "'").Fields(0).Value)
            If InvdocId <> "" And LabDocId <> "" Then
                SpareDocID = InvdocId
                LabourDocID = LabDocId
            Else
                SpareDocID = GetDocID(GCnFaS, SpareVtype, txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
                LabourDocID = GetDocID(GCnFaW, LabourVtype, txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
                LblSprBill.Refresh
                lblLabourBill.Refresh
            End If
    Else
        SpareDocID = GetDocID(GCnFaS, SpareVtype, txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
        LabourDocID = GetDocID(GCnFaW, LabourVtype, txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
        LblSprBill.Refresh
        lblLabourBill.Refresh
    End If
    'Stopped By Arpit after Discussing Mr. Panda and Mr. Jitendra at KNP 10.04.07
''''    If UCase(left(PubComp_Name, 3)) = "JMK" Then
''''        ServType = XNull(GCn.Execute("Select Serv_Desc from Job_Card Left Join Service_type on Job_Card.Serv_type=Service_type.Serv_type where Job_Card.DocId='" & Txt(JobNo).Tag & "'").Fields(0).Value)
''''        If Txt(CashBill).TEXT = "Yes" Then
''''            If UCase(left(ServType, 3)) = "ACC" Then
''''                SpareVtype = "W_SAC"
''''                LabourVtype = "W_LAC"
''''                SpareDocID = GetDocID(GCnFaS, "W_SAC", Txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
''''                LabourDocID = GetDocID(GCnFaW, "W_LAC", Txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
''''                LblSprBill.Refresh
''''                lblLabourBill.Refresh
''''            End If
''''        Else
''''            If UCase(left(ServType, 3)) = "ACC" Then
''''                SpareVtype = "W_SAR"
''''                LabourVtype = "W_LAR"
''''                SpareDocID = GetDocID(GCnFaS, "W_SAR", Txt(JobCDt), VoucherEditFlag, LblSprBill, lblSparePrefix, ForSiteCode)
''''                LabourDocID = GetDocID(GCnFaW, "W_LAR", Txt(JobCDt), VoucherEditFlag, lblLabourBill, lblLabourPrefix, ForSiteCode)
''''                LblSprBill.Refresh
''''                lblLabourBill.Refresh
''''            End If
''''        End If
''''    End If
End If
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

Private Sub CmdPrint_Click(Index As Integer)
'On Error GoTo ERRORHANDLER
Dim Rst As ADODB.Recordset
Dim I As Integer
If Index = 3 Then FrmPrn.Visible = False
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

If StrCmp(left(PubComp_Name, 7), "Vandana") Or IsCompName("Enar") Then
    mSepLabInv = 0
End If

    If Provisional Then
        Dim mFormCode$, mPrintDesc$
        If txt(GovtYn) = "No" Then  'Govt = No
            mFormCode = pubLocalTaxFormSpr
        Else
            mFormCode = pubGovtTaxFormSpr
        End If
        mPrintDesc = GCn.Execute("select Printing_desc from TaxForms where Form_Code='" & mFormCode & "'").Fields(0).Value
        GSQL = "SELECT '1' as Orig,JC.AtKMsHrs,JC.Lab_D_Amt,SPStk.DocID as ReqDocID," & vIsNull("SPStk.Srl_No", "0") & " as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab, " & ConvertDate(date) & "  as v_Date,'' AS Party_Code,' " & txt(OwnerName) & "' as Party_Name,'' as Address,'' as NamePrefix,' " & txt(OwnerName) & "' as Name,' " & txt(Address1) & "' as Add1,' " & txt(Address2) & "' as Add2,' " & txt(Address3) & "' as Add3,' " & txt(City) & "' as CityName,'' as PIN,' " & txt(PhoneResi) & "' as Phone," & _
            "'' as CSTNo,'' as L_C,'" & mFormCode & "' as Form_Code,'" & mPrintDesc & "' as Printing_Desc,'' as Remarks, " & Val(txt(MRPAmtTB)) & " as SprAmt_MRP_TB, " & Val(txt(MRPAmtTP)) & " as SprAmt_MRP_TP," & mMRPLubeTB & " as OilAmt_MRP_TB," & mMRPLubeTP & " as OilAmt_MRP_TP, " & Val(txt(SprAmtTB)) & " as SprAmt_TB, " & Val(txt(SprAmtTP)) & " as SprAmt_TP, " & Val(txt(OilAmtTB)) & " as OilAmt_TB,  " & Val(txt(OilAmtTP)) & " as OilAmt_TP, " & Val(txt(DiscPerTB)) & " as D_Per_TB,  " & Val(txt(DiscAmtTB)) & " as D_Amt_TB,  " & Val(txt(DiscPerTP)) & " as D_Per_TP, " & Val(txt(DiscAmtTP)) & " as D_Amt_TP,0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP,0 as D_Amt_MRP_TP, " & Val(txt(Addition)) & " as Addition, " & Val(txt(GenSurPer)) & " as Gen_Sur_Per, " & Val(txt(GenSurAmt)) & " as Gen_Sur_Amt, " & Val(txt(TransAmt)) & " as Trans_Amt, " & Val(txt(STaxPer)) & " as Tax_Per,  " & Val(txt(STaxAmt)) & " as Tax_Amt, 0 as Tax_AmtMRP,  " & Val(txt(TaxSurPer)) & " as Tax_Sur_Per, " & Val(txt(TaxSurAmt)) & " as Tax_Sur_Amt,0 as TaxSur_AmtMRP, " & Val(txt(PackCrg)) & " as Packing,  " & Val(txt(TurnOverPer)) & " as TOT_Per,  " & Val(txt(TurnOverAmt)) & " as Tot_Amt, 0 as TOT_AmtMRP,0 as ReSalTax_Per,0 as ReSalTax_Amt, " & Val(txt(STotB)) & " as Total_Amt," & _
            " " & Val(txt(SROff)) & " as Rounded, " & PubTaxDetOnSprInv & " as Det_Tax,'' as GP_No,'' as GP_Date,1 as Printed_YN, ' " & pubUName & "' as U_Name, ' " & date$ & "' as U_EntDt,0 as CancelYN,0 as LabAmt_TB, 0 as LabAmt_TP, 0 as Lab_TaxPer, 0 as Lab_TaxAmt, 0 as Lab_D_Amt,0 as Lab_RoundOff,0 as NetLab_Amt," & _
            "SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown," & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc, " & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec," & _
            "" & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret," & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN," & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & vIsNull("SPStk.Rate", "0") & " as Rate," & _
            "" & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate," & xIsNull("SPStk.Purpose", "") & " as Purpose,SPStk.Part_SrlNo," & vIsNull("SPStk.Rate2", "0") & " as Rate2," & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate2," & _
            "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2," & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2," & vIsNull("SPStk.Amount", "0") & " as Amount2," & vIsNull("SPStk.Net_Amt", "0") & " as Net_Amt2,'' as Chrg_From,0 as External_YN, " & _
            "Syctrl.WorkShopInvFooter, " & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,SPStk.TaxPer,SPStk.TaxAmt,Jc.Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, SpStk.SatPer, SpStk.SatAmt, " & Val(txt(SatAmt)) & " As SatAmt_H," & cCStr(xIsNull("SpStk.v_No", "")) & " as ReqNo,  " & xIsNull("P.Unit", "Each") & " As Unit " & _
        " FROM (((SP_Stock as SPStk left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
            "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON SPStk.Job_DocID = JC.DocId) " & _
            "LEFT JOIN Syctrl ON Syctrl.LinkTable<>SPStk.U_AE)" & _
            "where SPStk.Job_DocId='" & lblDocId & "'"  'modi lps  and (SPStk.Qty_Iss -SPStk.Qty_Ret) >0 "
        If UCase(left(PubComp_Name, 5)) = "NAWAL" Then
            GSQL = GSQL & " AND " & xIsNull("SPStk.Purpose", "") & " = 'C'"
        End If
            
    'Modi LPS at Cuttack 31.08.03
    '    If PurposeStr <> "" Then
    '        GSQL = GSQL & " and SPStk.Purpose not in (" & PurposeStr & ")"
    '    End If
        'GSQL = GSQL & "Order by SpStk.Docid,SpStk.Srl_No"
        mQryLab = "SELECT '2' as Orig,0 as AtKMsHrs,0 as Lab_D_Amt,'                     ' as ReqDocID,JL.S_No as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab," & ConvertDate(date) & "  as v_Date,JC.DrLab_AcCode as Party_Code,JC.BillingName as Party_Name,'' as Address,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
            "SG.CSTNo,'' as L_C,'' as Form_Code,'' as Printing_Desc,'' as Remarks,0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
            "0 as Rounded," & PubTaxDetOnSprInv & " as Det_Tax,JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted ,JC.U_Name,JC.U_EntDt,0 as CancelYN,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
            "JL.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name,'' as Lub_Category, '' as Godown,0 as Qty_Doc, 0 as Qty_Rec, " & _
            "" & vIsNull("Hrs_Taken", "0") & " as Qty_Iss,0 as Qty_Ret," & vIsNull("JL.Tax_YN", "0") & " as Tax_YN, 0 as MRP_YN,0 as Rate," & _
            "0 as MRP_Rate,'' as Purpose,'' as Part_SrlNo," & vIsNull("JL.Lab_Rate", "0") & " as Rate2,0 as MRP_Rate2," & _
            "0 as Disc_Per2,0 as Disc_Amt2,0 as Amount2," & cIIF("JL.Chrg_From = 'C'", "JL.LabourAmt", "0") & " as Net_Amt2,JL.Chrg_From,JL.External_YN," & _
            "Syctrl.WorkShopInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
            "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass ,0 as TaxPer,0 as TaxAmt,'' AS Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, 0 As SatPer, 0 As SatAmt, 0 As SatAmt_H,' ' as ReqNo, '' as Unit   " & _
        " FROM ((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
            "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
            "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
            "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
            "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
            "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code)" & _
        "Where JC.DocId='" & lblDocId & "'" ' Order By JL.JobDocID, JL.S_No"
     
        GSQL = GSQL & " Union All " & mQryLab '& " Order By 1,2,3"
        
    Else

        If mSepLabInv = 0 Then  'No, Merge Invoice=Spare + Labour
            GSQL = "SELECT '1' as Orig,JC.AtKMsHrs,JC.Lab_D_Amt,JC.ArrivalTime,JC.JobComp_Dt_Time,SPStk.DocID as ReqDocID, " & vIsNull("SPStk.Srl_No", "0") & " as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvSpr,JC.DocID_InvLab,JC.JobCloseDate as v_Date,s.Party_Code,s.Party_Name,s.Address,SG.NamePrefix,H.Name,H.Add1,H.Add2,H.Add3,City.CityName,SG.PIN,SG.Phone," & _
                "SG.CSTNo,s.L_C,s.Form_Code,TF.Printing_Desc,s.Remarks,s.SprAmt_MRP_TB,s.SprAmt_MRP_TP,s.OilAmt_MRP_TB,s.OilAmt_MRP_TP,s.SprAmt_TB,s.SprAmt_TP,s.OilAmt_TB, s.OilAmt_TP,s.D_Per_TB, s.D_Amt_TB, s.D_Per_TP,s.D_Amt_TP,s.D_Per_MRP_TB,s.D_Amt_MRP_TB," & _
                "s.D_Per_MRP_TP,s.D_Amt_MRP_TP,s.Addition,s.Gen_Sur_Per,s.Gen_Sur_Amt,s.Trans_Amt," & vIsNull("s.Tax_Per", "0") & " As Tax_Per, " & vIsNull("s.Tax_Amt", "0") & " As Tax_Amt, s.Tax_AmtMRP, s.Tax_Sur_Per,s.Tax_Sur_Amt,s.TaxSur_AmtMRP,s.Packing, s.TOT_Per, s.Tot_Amt, s.TOT_AmtMRP,s.ReSalTax_Per,s.ReSalTax_Amt,s.Total_Amt," & _
                "s.Rounded," & PubTaxDetOnSprInv & " As Det_Tax,s.GP_No,s.GP_Date,s.Printed_YN,s.U_Name, s.U_EntDt,S.CancelYN,0 as LabAmt_TB, 0 as LabAmt_TP, 0 as Lab_TaxPer, 0 as Lab_TaxAmt, 0 as Lab_D_Amt,JC.Lab_RoundOff,0 as NetLab_Amt," & _
                "SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown," & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc," & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec," & _
                "" & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret," & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN," & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & vIsNull("SPStk.Rate", "0") & " as Rate," & _
                "" & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate," & xIsNull("SPStk.Purpose", "") & " as Purpose,SPStk.Part_SrlNo," & vIsNull("SPStk.Rate2", "0") & " as Rate2," & vIsNull("SPStk.MRP_Rate2", "0") & " as MRP_Rate2," & _
                "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2, " & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2," & vIsNull("SPStk.Amount2", "0") & " as Amount2, " & vIsNull("SPStk.Net_Amt2", "0") & " as Net_Amt2,'' as Chrg_From,'' as External_YN, " & _
                "Syctrl.WorkShopInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,SPStk.TaxPer," & cIIF("SPStk.Purpose='C' Or SPStk.Purpose='W'", "SPStk.TaxAmt", "0") & " As TaxAmt,Jc.Job_No, " & cCStr(xIsNull("SpStk.v_No", "")) & " as ReqNo,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, JC.Job_Date, H.PhoneOff, H.PhoneResi, H.Mobile, " & xIsNull("P.Unit", "Each") & " As Unit, 0 as eCessPer, 0 as eCessAmt, (Select SrvTaxNo From Syctrl) As SrvTaxNo, SG.LstNo As LstNoParty, H.PhoneOff, H.PhoneResi, H.Mobile, " & vIsNull("MG.ModelGrp_Name", "M.Model") & "  As ModelGrp_Name, SG.SiebelCode, SpStk.SatPer, SpStk.SatAmt, " & Val(txt(SatAmt)) & " as SatAmt_H,P.Unit,Syctrl.HelpLineNo " & _
            "FROM ((((((((SP_Sale as S left JOIN SP_Stock as SPStk ON S.Job_DocId = SPStk.Job_DocId) " & _
                "left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
                "LEFT JOIN SubGroup as SG  ON S.Party_Code = SG.SubCode) " & _
                "Left Join TaxForms TF on S.Form_Code=TF.Form_Code) " & _
                "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON S.Job_DocID = JC.DocId) " & _
                "LEFT JOIN City ON H.CityCode = City.CityCode)  " & _
                "Left Join Model M On M.Model=H.Model) " & _
                "Left Join Model_Grp MG On MG.ModelGrp_Code=M.Grp_Code) " & _
                "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
                "where S.Job_DocId='" & Master!Code & "'  " 'modi lps  and (SPStk.Qty_Iss -SPStk.Qty_Ret) >0 "
            If UCase(left(PubComp_Name, 5)) = "NAWAL" Then
                GSQL = GSQL & " AND " & xIsNull("SPStk.Purpose", "") & " = 'C'"
            End If
            GSQL = GSQL '& " ORDER BY SPStk.V_no"
        'Modi LPS at Cuttack 31.08.03
        '    If PurposeStr <> "" Then
        '        GSQL = GSQL & " and SPStk.Purpose not in (" & PurposeStr & ")"
        '    End If
            'GSQL = GSQL & "Order by SpStk.Docid,SpStk.Srl_No"
            mQryLab = "SELECT '2' as Orig,0 as AtKMsHrs,JC.Lab_D_Amt,JC.ArrivalTime,JC.JobComp_Dt_Time,JL.Job_DocID as ReqDocID,JL.S_No as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvSpr,JC.DocID_InvLab,JC.JobCloseDate as v_Date,JC.DrLab_AcCode as Party_Code,JC.BillingName as Party_Name,'' as Address,SG.NamePrefix,H.Name,H.Add1,H.Add2,H.Add3,City.CityName,SG.PIN,SG.Phone," & _
                "SG.CSTNo,'' as L_C,'' as Form_Code,'' as Printing_Desc,'' as Remarks,0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
                "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, (Select Tax_Amt From Sp_Sale Where Job_DocId=JC.DocId) as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,(Select Total_Amt From Sp_Sale Where Job_DocId=JC.DocId) as Total_Amt," & _
                "(Select Rounded From Sp_Sale Where Job_DocId=JC.DocId) as Rounded,0 as Det_Tax,JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted as Printed_YN,JC.U_Name,JC.U_EntDt,0 as CancelYN,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
                "JL.Lab_Code as Part_No,Labour.Lab_Desc  " & IIf(UCase(left(PubComp_Name, 3)) = "LMP" Or UCase(left(PubComp_Name, 7)) = "VANDANA" Or IsCompName("Enar"), "+ ' ' + JL.Mech_Voice", "") & " as Part_Name,'' as Lub_Category, '' as Godown,0 as Qty_Doc, 0 as Qty_Rec, " & _
                "" & vIsNull("Hrs_Taken", "0") & " as Qty_Iss,0 as Qty_Ret," & vIsNull("JL.Tax_YN", "0") & " as Tax_YN, 0 as MRP_YN,0 as Rate," & _
                "0 as MRP_Rate,JL.Chrg_Type as Purpose,'' as Part_SrlNo," & vIsNull("JL.Lab_Rate", "0") & " as Rate2,0 as MRP_Rate2," & _
                "0 as Disc_Per2,0 as Disc_Amt2," & cIIF("JL.Chrg_From = 'C'", "JL.LabourAmt", "0") & "  as Amount2," & cIIF("JL.Chrg_From = 'C'", "JL.LabourAmt", "0") & " as Net_Amt2,JL.Chrg_From,JL.External_YN," & _
                "Syctrl.WorkShopInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & "  as SrvGatePass_On," & _
                "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,0 as TaxPer,0 as TaxAmt,Jc.Job_No, " & cCStr(xIsNull("JL.Lab_Code", "")) & "  as ReqNo,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, JC.Job_Date, H.PhoneOff, H.PhoneResi, H.Mobile, '' As Unit, JC.eCessPer, JC.eCessAmt, (Select SrvTaxNo From Syctrl) As SrvTaxNo, SG.LstNo As LstNoParty, H.PhoneOff, H.PhoneResi, H.Mobile, " & vIsNull("MG.ModelGrp_Name", "M.Model") & "  As ModelGrp_Name, SG.SiebelCode, 0 As SatPer, 0 As SatAmt, 0 As SatAmt_H,'' as Unit ,'' as HelpLineNo   " & _
            "FROM (((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
                "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
                "LEFT JOIN City ON H.CityCode = City.CityCode) " & _
                "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
                "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
                "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code)  Left Join Model M On M.Model=H.Model) Left Join Model_Grp MG On MG.ModelGrp_Code=M.Grp_Code " & _
            "Where JC.DocId='" & Master!Code & "'" ' Order By JL.JobDocID, JL.S_No"

            
            
            If UCase(left(PubComp_Name, 3)) = "JMK" Then
                If ChkRep(0).Value = vbUnchecked And ChkRep(1).Value = vbChecked Then
                    GSQL = mQryLab
                ElseIf ChkRep(1).Value = vbChecked And ChkRep(0).Value = vbUnchecked Then
                ElseIf ChkRep(1).Value = vbChecked And ChkRep(0).Value = vbChecked Then
                    GSQL = GSQL & " Union All " & mQryLab & " order by Orig Desc "
                End If
            Else
                GSQL = GSQL & " Union All " & mQryLab '& " order by 1,2,3,Part_No"
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
                "" & vIsNull("SPStk.Srl_No", "0") & " as Srl_No,SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown," & _
                "" & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc," & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec," & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & _
                "" & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret," & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN," & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & _
                "" & vIsNull("SPStk.Rate", "0") & " as Rate, " & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate," & _
                "" & xIsNull("SPStk.Purpose", "") & " as Purpose,SPStk.Part_SrlNo, " & vIsNull("SPStk.Rate2", "0") & " as Rate2, " & vIsNull("SPStk.MRP_Rate2", "0") & " as MRP_Rate2," & _
                "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2, " & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2, " & vIsNull("SPStk.Amount2", "0") & " as Amount2," & _
                "" & vIsNull("SPStk.Net_Amt2", "0") & " as Net_Amt2," & _
                "Syctrl.WorkShopInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
                "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,SPStk.TaxPer,SPStk.TaxAmt,Jc.Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, SpStk.SatPer, SpStk.SatAmt, S.SatAmt As SatAmt_H, JC.Remark As JobRemark,    SPStk.DepAmt,   SPStk.InsuranceAmt, SPStk.DiffPeried,DITM.ShortName as PartType,P.UNIT as Part_Unit,SG.LSTNo,S.Cash_Credit " & _
            "FROM ((((((SP_Sale as S left JOIN SP_Stock as SPStk ON S.Job_DocId = SPStk.Job_DocId) " & _
                "left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
                "LEFT JOIN (SubGroup as SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON S.Party_Code = SG.SubCode) " & _
                "Left Join TaxForms TF on S.Form_Code=TF.Form_Code) " & _
                "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON S.Job_DocID = JC.DocId)LEFT JOIN Deprecation_itemMaster DITM ON SPStk.Dep_Item=ditm.Code) " & _
                "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
            "where S.Job_DocId='" & txt(JobNo).Tag & "' and ((SPStk.Qty_Iss -SPStk.Qty_Ret) >0 Or (SPStk.Qty_Iss -SPStk.Qty_Ret) Is Null) "
        'Modi LPS at Cuttack 31.08.03
        '    If PurposeStr <> "" Then
        '        GSQL = GSQL & " and SPStk.Purpose not in (" & PurposeStr & ")"
        '    End If
            'GSQL = GSQL & "Order By SPStk.Part_No"
            
            'Labour SQL
            mQryLab = "SELECT 0 as AtKMsHrs,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvLab,JC.JobCloseDate as v_Date," & _
                "JC.DrLab_AcCode as Party_Code," & cIIF("CrMemo=0", "JC.BillingName", "Sg.Name") & " as Party_Name,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
                "JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted,JC.U_Name,JC.U_EntDt," & _
                "JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
                "JL.S_No,JL.Lab_Code,Labour.Lab_Desc as LabName,JL.Tax_YN,JL.Hrs_Taken,JL.Hrs_War,JL.Lab_Rate," & _
                "JL.LabourAmt,JL.Chrg_From,JL.External_YN," & _
                "Syctrl.LabInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
                "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,0 as TaxPer,0 as TaxAmt,Jc.Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, 0 As SatPer, 0 As SatAmt, 0 As SatAmt_H, JC.Remark As JobRemark,SG.LSTNo " & _
            "FROM ((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
                "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
                "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
                "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
                "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
                "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code) " & _
            "Where JC.DocId='" & txt(JobNo).Tag & _
            "' Order By JL.Lab_Code"
''        mQryLab = "SELECT 0 as AtKMsHrs,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvLab,JC.JobCloseDate as v_Date," & _
''                "JC.DrLab_AcCode as Party_Code," & cIIF("CrMemo=0", "JC.BillingName", "Sg.Name") & " as Party_Name,SG.NamePrefix,SG.Name," & cIIF("CrMemo=0", "H.Add1", "Sg.Add1") & " as Party_Add1," & cIIF("CrMemo=0", "H.Add2", "Sg.Add2") & " as Party_Add2," & cIIF("CrMemo=0", "H.Add3", "Sg.Add3") & " as Party_Add3,City.CityName,SG.PIN,SG.Phone," & _
''                "JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted,JC.U_Name,JC.U_EntDt," & _
''                "JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
''                "JL.S_No,JL.Lab_Code,Labour.Lab_Desc as LabName,JL.Tax_YN,JL.Hrs_Taken,JL.Hrs_War,JL.Lab_Rate," & _
''                "JL.LabourAmt,JL.Chrg_From,JL.External_YN," & _
''                "Syctrl.LabInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
''                "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,0 as TaxPer,0 as TaxAmt,Jc.Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, 0 As SatPer, 0 As SatAmt, 0 As SatAmt_H, JC.Remark As JobRemark " & _
''                "FROM ((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
''                "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
''                "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
''                "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
''                "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
''                "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code) " & _
''                "Where JC.DocId='" & txt(JobNo).Tag & _
''                "' Order By JL.Lab_Code"
        End If
    End If

Select Case Index
Case DEP_Bill
mRepName = "Dep_Bill"
'Call WindowsPrintSpr(Index, GSQL)


            
            
'              mPrintDesc = GCn.Execute("select Printing_desc from TaxForms where Form_Code='" & mFormCode & "'").Fields(0).Value
        GSQL = "SELECT '1' as Orig,JC.AtKMsHrs,JC.Lab_D_Amt,SPStk.DocID as ReqDocID," & vIsNull("SPStk.Srl_No", "0") & " as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab, " & ConvertDate(date) & "  as v_Date,'' AS Party_Code,' " & txt(OwnerName) & "' as Party_Name,'' as Address,'' as NamePrefix,' " & txt(OwnerName) & "' as Name,' " & txt(Address1) & "' as Add1,' " & txt(Address2) & "' as Add2,' " & txt(Address3) & "' as Add3,' " & txt(City) & "' as CityName,'' as PIN,' " & txt(PhoneResi) & "' as Phone," & _
            "'' as CSTNo,'' as L_C,'" & mFormCode & "' as Form_Code,'" & mPrintDesc & "' as Printing_Desc,'' as Remarks, " & Val(txt(MRPAmtTB)) & " as SprAmt_MRP_TB, " & Val(txt(MRPAmtTP)) & " as SprAmt_MRP_TP," & mMRPLubeTB & " as OilAmt_MRP_TB," & mMRPLubeTP & " as OilAmt_MRP_TP, " & Val(txt(SprAmtTB)) & " as SprAmt_TB, " & Val(txt(SprAmtTP)) & " as SprAmt_TP, " & Val(txt(OilAmtTB)) & " as OilAmt_TB,  " & Val(txt(OilAmtTP)) & " as OilAmt_TP, " & Val(txt(DiscPerTB)) & " as D_Per_TB,  " & Val(txt(DiscAmtTB)) & " as D_Amt_TB,  " & Val(txt(DiscPerTP)) & " as D_Per_TP, " & Val(txt(DiscAmtTP)) & " as D_Amt_TP,0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP,0 as D_Amt_MRP_TP, " & Val(txt(Addition)) & " as Addition, " & Val(txt(GenSurPer)) & " as Gen_Sur_Per, " & Val(txt(GenSurAmt)) & " as Gen_Sur_Amt, " & Val(txt(TransAmt)) & " as Trans_Amt, " & Val(txt(STaxPer)) & " as Tax_Per,  " & Val(txt(STaxAmt)) & " as Tax_Amt, 0 as Tax_AmtMRP,  " & Val(txt(TaxSurPer)) & " as Tax_Sur_Per, " & Val(txt(TaxSurAmt)) & " as Tax_Sur_Amt,0 as TaxSur_AmtMRP, " & Val(txt(PackCrg)) & " as Packing,  " & Val(txt(TurnOverPer)) & " as TOT_Per,  " & Val(txt(TurnOverAmt)) & " as Tot_Amt, 0 as TOT_AmtMRP,0 as ReSalTax_Per,0 as ReSalTax_Amt, " & Val(txt(STotB)) & " as Total_Amt," & _
            " " & Val(txt(SROff)) & " as Rounded, " & PubTaxDetOnSprInv & " as Det_Tax,'' as GP_No,'' as GP_Date,1 as Printed_YN, ' " & pubUName & "' as U_Name, ' " & date$ & "' as U_EntDt,0 as CancelYN,0 as LabAmt_TB, 0 as LabAmt_TP, 0 as Lab_TaxPer, 0 as Lab_TaxAmt, 0 as Lab_D_Amt,0 as Lab_RoundOff,0 as NetLab_Amt," & _
            "SPStk.Part_No,P.Part_Name,SPStk.Lub_Category, SPStk.Godown," & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc, " & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec," & _
            "" & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret," & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN," & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & vIsNull("SPStk.Rate", "0") & " as Rate," & _
            "" & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate," & xIsNull("SPStk.Purpose", "") & " as Purpose,SPStk.Part_SrlNo," & vIsNull("SPStk.Rate2", "0") & " as Rate2," & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate2," & _
            "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2," & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2," & vIsNull("SPStk.Amount", "0") & " as Amount2," & vIsNull("SPStk.Net_Amt", "0") & " as Net_Amt2,'' as Chrg_From,0 as External_YN, " & _
            "Syctrl.WorkShopInvFooter, " & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,SPStk.TaxPer,SPStk.TaxAmt,Jc.Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, SpStk.SatPer, SpStk.SatAmt, " & Val(txt(SatAmt)) & " As SatAmt_H," & cCStr(xIsNull("SpStk.v_No", "")) & " as ReqNo,    SPStk.DepAmt,   SPStk.InsuranceAmt, SPStk.DiffPeried,DITM.ShortName as PartType,i.Name AS InsuranceCompanyName,SPStk.excise_amt  " & _
        " FROM ((((SP_Stock as SPStk left JOIN Part as P ON SPStk.Part_No = P.Part_No and P.Div_Code = left(SPStk.Docid,1)) " & _
            "LEFT JOIN (Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) ON SPStk.Job_DocID = JC.DocId) LEFT JOIN Deprecation_itemMaster DITM ON SPStk.Dep_Item=ditm.Code ) " & _
            "LEFT JOIN Syctrl ON Syctrl.LinkTable<>SPStk.U_AE)LEFT JOIN Insurance i ON h.InsuranceCompany=i.Code " & _
            "where SPStk.Job_DocId='" & lblDocId & "'"  'modi lps  and (SPStk.Qty_Iss -SPStk.Qty_Ret) >0 "
      
            
          mQryLab = "SELECT '2' as Orig,0 as AtKMsHrs,0 as Lab_D_Amt,'                     ' as ReqDocID,JL.S_No as Srl_No,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocID as DocId_InvSpr,JC.DocID as DocID_InvLab," & ConvertDate(date) & "  as v_Date,JC.DrLab_AcCode as Party_Code,JC.BillingName as Party_Name,'' as Address,SG.NamePrefix,SG.Name,SG.Add1,SG.Add2,SG.Add3,City.CityName,SG.PIN,SG.Phone," & _
            "SG.CSTNo,'' as L_C,'' as Form_Code,'' as Printing_Desc,'' as Remarks,0 as SprAmt_MRP_TB, 0 as SprAmt_MRP_TP, 0 as OilAmt_MRP_TB, 0 as OilAmt_MRP_TP,0 as SprAmt_TB, 0 as SprAmt_TP, 0 as OilAmt_TB, 0 as OilAmt_TP,0 as D_Per_TB, 0 as D_Amt_TB, 0 as D_Per_TP, 0 as D_Amt_TP, 0 as D_Per_MRP_TB,0 as D_Amt_MRP_TB," & _
            "0 as D_Per_MRP_TP, 0 as D_Amt_MRP_TP, 0 as Addition, 0 as Gen_Sur_Per, 0 as Gen_Sur_Amt,0 as Trans_Amt,0 as Tax_Per, 0 as Tax_Amt, 0 as Tax_AmtMRP, 0 as Tax_Sur_Per, 0 as Tax_Sur_Amt, 0 as TaxSur_AmtMRP, 0 as Packing, 0 as TOT_Per, 0 as Tot_Amt,0 as TOT_AmtMRP, 0 as ReSalTax_Per, 0 as ReSalTax_Amt,0 as Total_Amt," & _
            "0 as Rounded," & PubTaxDetOnSprInv & " as Det_Tax,JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted as Printed_YN,JC.U_Name,JC.U_EntDt,0 as CancelYN,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
            "JL.Lab_Code as Part_No,Labour.Lab_Desc as Part_Name,'' as Lub_Category, '' as Godown,0 as Qty_Doc, 0 as Qty_Rec, " & _
            "" & vIsNull("Hrs_Taken", "0") & " as Qty_Iss,0 as Qty_Ret," & vIsNull("JL.Tax_YN", "0") & " as Tax_YN, 0 as MRP_YN,0 as Rate," & _
            "0 as MRP_Rate,'' as Purpose,'' as Part_SrlNo," & vIsNull("JL.Lab_Rate", "0") & " as Rate2,0 as MRP_Rate2," & _
            "0 as Disc_Per2,0 as Disc_Amt2,0 as Amount2," & cIIF("JL.Chrg_From = 'C'", "JL.LabourAmt", "0") & " as Net_Amt2,JL.Chrg_From,JL.External_YN," & _
            "Syctrl.WorkShopInvFooter," & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
            "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass ,0 as TaxPer,0 as TaxAmt,'' AS Job_No,Jc.LastInvNoSuff,JC.LastLabInvNoSuff, 0 As SatPer, 0 As SatAmt, 0 As SatAmt_H,' ' as ReqNo ,    jl.DepAmt,   jl.InsuranceAmt, jl.DiffPeried,DITM.ShortName as PartType,i.Name AS InsuranceCompanyName,0 as excise_amt  " & _
        " FROM (((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
            "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
            "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
            "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
            "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
            "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code)LEFT JOIN Deprecation_itemMaster DITM ON JL.Dep_Item=ditm.Code)LEFT JOIN Insurance i ON h.InsuranceCompany=i.Code  " & _
        "Where JC.DocId='" & lblDocId & "'" ' Order By JL.JobDocID, JL.S_No"
     
        GSQL = GSQL & " Union All " & mQryLab '& " Order By 1,2,3"
        
        
 Set Rst = GCn.Execute(GSQL)
                CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
                If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
                Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
                rpt.Database.SetDataSource Rst
                
                Set Rst = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram, LstNoW, LstDateW from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
                For I = 1 To rpt.FormulaFields.Count
                    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                        Case UCase("CstStr")
                            rpt.FormulaFields(I).TEXT = "'" & Rst!W_SecCST & " Dated " & Rst!W_SecCST_Date & "'"
                        Case UCase("LstStr")
                            rpt.FormulaFields(I).TEXT = "'" & Rst!W_SecLST & " Dated " & Rst!W_SecLST_Date & "'"
                        Case UCase("LstNo")
                            rpt.FormulaFields(I).TEXT = "'" & Rst!LstNoW & " Dated " & Rst!LstDateW & "'"
                    End Select
                Next
                rpt.ReadRecords
                'rpt.PrintOut
                
                Call Report_View(rpt, Me.CAPTION, 0, True)
                
                Set Rst = Nothing
                Exit Sub
                
    Case PScreen, PWindows
        If OptPlain.Value = True Then
            If mSepLabInv = 0 Or Provisional Then ' Merge Invoice
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
        If mSepLabInv = 0 Or Provisional Then
            'Call WindowsPrintBoth(Index, GSQL, mQryLab)
            If UCase(left(PubComp_Name, 3)) = "LMP" Or UCase(left(PubComp_Name, 5)) = "UJWAL" Or StrCmp(left(PubComp_Name, 3), "JMK") Or StrCmp(left(PubComp_Name, 7), "Vandana") Or IsCompName("Enar") Then
            
                If StrCmp(left(PubComp_Name, 7), "Vandana") Or IsCompName("Enar") Then
                    mRepName = "WorkShopBillVandana"
                Else
                    mRepName = "WorkShopBillSiebel"
                End If
                Set Rst = GCn.Execute(GSQL)
                CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
                If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
                Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
                rpt.Database.SetDataSource Rst
                
                Set Rst = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram, LstNoW, LstDateW from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
                For I = 1 To rpt.FormulaFields.Count
                    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                        Case UCase("CstStr")
                            rpt.FormulaFields(I).TEXT = "'" & Rst!W_SecCST & " Dated " & Rst!W_SecCST_Date & "'"
                        Case UCase("LstStr")
                            rpt.FormulaFields(I).TEXT = "'" & Rst!W_SecLST & " Dated " & Rst!W_SecLST_Date & "'"
                        Case UCase("LstNo")
                            rpt.FormulaFields(I).TEXT = "'" & Rst!LstNoW & " Dated " & Rst!LstDateW & "'"
                    End Select
                Next
                rpt.ReadRecords
                'rpt.PrintOut
                
                Call Report_View(rpt, Me.CAPTION, 0, True)
                
                Set Rst = Nothing
                Exit Sub
            End If
        Else
            If ChkRep(ChkSprInv).Value = 1 Then Call WindowsPrintSpr(Index, GSQL)
            If ChkRep(ChkLabInv).Value = 1 Then Call WindowsPrintLab(Index, mQryLab)
        End If
        FrmPrn.Visible = False
    Case PDos
        If mSepLabInv = 0 Or Provisional Then
            If UCase(left(PubComp_Name, 3)) = "JMK" Then
                Call SpeedPrintBothJMK(GSQL, Optpre.Value)
            Else
                Call SpeedPrintBoth(GSQL, Optpre.Value)
            
            End If
        Else
            If ChkRep(ChkSprInv).Value = 1 Then Call SpeedPrintSpr(GSQL, Optpre.Value)
            If ChkRep(ChkLabInv).Value = 1 Then Call SpeedPrintLab(mQryLab)
        End If
        FrmPrn.Visible = False
    Case 5
       'Labour SQL
''        mQryLab = "SELECT 0 as AtKMsHrs,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvLab,JC.JobCloseDate as v_Date," & _
''            "JC.DrLab_AcCode as Party_Code,AD.D_Name as Party_Name,SG.NamePrefix,Ad.D_Code,AD.D_Name as Name,AD.D_Add1 as Add1,AD.D_Add2 as Add2,AD.D_Add3 as Add3,AD.D_City as CityName,'' as PIN,'' as Phone," & _
''            "JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted,JC.U_Name,JC.U_EntDt," & _
''            "JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
''            "JL.S_No,JL.Lab_Code,Labour.Lab_Desc as LabName,JL.Tax_YN,JL.Hrs_Taken,JL.Hrs_War,JL.Lab_Rate," & _
''            "JL.LabourAmt,JL.Chrg_Type,JL.External_YN," & _
''            "Syctrl.LabInvFooter, " & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
''            "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,0 as TaxPer,0 as TaxAmt,JC.Coupon_Value " & _
''        "FROM (((((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
''            "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
''            "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
''            "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
''            "LEFT JOIN Job_Lab JL ON JC.DocId = JL.Job_DocID) " & _
''            "LEFT JOIN Labour ON JL.Lab_Code = Labour.Lab_Code) " & _
''            "LEFT JOIN Amd_Dealer AD ON H.Dealer_Code = AD.D_Code) " & _
''        "Where JC.DocId='" & Master!Code & _
''        "' Order By JL.Lab_Code"
        mQryLab = "SELECT 0 as AtKMsHrs,JC.DocID as JobDocID,JC.CrMemo,H.Model,H.RegNo,H.Chassis,JC.DocId_InvLab,JC.JobCloseDate as v_Date," & _
            "JC.DrLab_AcCode as Party_Code,AD.D_Name as Party_Name,SG.NamePrefix,Ad.D_Code,AD.D_Name as Name,AD.D_Add1 as Add1,AD.D_Add2 as Add2,AD.D_Add3 as Add3,AD.D_City as CityName,'' as PIN,'' as Phone," & _
            "JC.GP_No,JC.JobCloseDate as GP_Date,JC.LabBillPrinted,JC.U_Name,JC.U_EntDt," & _
            "JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.Lab_RoundOff,JC.NetLab_Amt, " & _
            "Syctrl.LabInvFooter, " & vIsNull("Syctrl.SrvGatePass_On", "0") & " as SrvGatePass_On," & _
            "" & vIsNull("Syctrl.SrvGatePass", "0") & " as SrvGatePass,0 as TaxPer,0 as TaxAmt,JC.Coupon_Value,jc.atkmshrs as KM " & _
        "FROM (((((Job_Card JC LEFT JOIN HisCard H ON JC.CardNo = H.CardNo) " & _
            "left join SubGroup SG ON JC.DrLab_AcCode = SG.SubCode) " & _
            "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
            "LEFT JOIN Syctrl ON JC.U_AE <> Syctrl.LinkTable) " & _
            "LEFT JOIN Amd_Dealer AD ON H.Dealer_Code = AD.D_Code) " & _
            "Where JC.DocId='" & Master!Code & " '"

        SpeedPrintOthDlr (mQryLab)

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
Private Sub SpeedPrintSpr(mQry$, PrePrinted)
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
    
    Set RstJob = GCn.Execute(mQry)
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
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
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
    If mVatYn = 1 Then
        
        If RSOJPR = True Or left(PubComp_Name, 10) = "GANGANAGAR" Then
            mDocStr = "VAT INVOICE"
        Else
            mDocStr = "RETAIL INVOICE"
        End If
        If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
               mDocStr = "TAX INVOICE"
        End If
    Else
        If RstJob!CrMemo = 0 Then
            mDocStr = "CASH MEMO"
        Else
            mDocStr = "INVOICE"
        End If
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
       
    End If
        Print #1, PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        'Service tax No Printing............
        SrvTaxNo = XNull(GCn.Execute("Select SrvTaxNo from Syctrl").Fields(0).Value)
        
        Print #1, mSP2 & PSTR("Serv.Tax No.  : " & SrvTaxNo, 40, , AlignLeft)
        mHeader = mHeader + 1
        '....................................
        If mVatYn = 1 Then
            Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "B", PageWidth)
            mHeader = mHeader + 1
        Else
            Print #1, PRN_TIT("** WORKSHOP SPARE " & mDocStr & mDupStr & " **", "B", PageWidth)
            mHeader = mHeader + 1
        End If
        If RSOJPR = True And VNull(RstJob!LastInvNoSuff) > 0 Then
            Print #1, mSP2 & mChr18 & Space(36) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " : " & PrinID(RstJob!DocId_InvSpr) & "-" & VNull(RstJob!LastInvNoSuff) & mEmph1
        Else
            Print #1, mSP2 & mChr18 & Space(36) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
        End If
        Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
        mHeader = mHeader + 1
        '**********************************
        'Print #1, PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
        'mHeader = mHeader + 1
        'Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
        'mHeader = mHeader + 1
        'Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
        'mHeader = mHeader + 1
        'Print #1, PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
        '& Space(4) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
        'mHeader = mHeader + 1
        'Print #1, mSP2 & "Phone : " & PSTR(XNull(RstJob!Phone), 20)
        'mHeader = mHeader + 1
        '***************************************
        
        Print #1, mSP2 & PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40) & Space(6) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address3)) & IIf(XNull(txt(Address3)) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44) _
        & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
        mHeader = mHeader + 1
        Print #1, mSP2 & "Phone : " & PSTR(XNull(txt(PhoneOff)), 20)
        
        
        
        Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
        mHeader = mHeader + 1
        If mVatYn = 1 Then
            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR("DISC %", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 12, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("Tax Amt", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18    '& mDoub1
            mHeader = mHeader + 1
        Else
            If RstJob!Det_Tax = 1 Then
                Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                mHeader = mHeader + 1
                Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
                mHeader = mHeader + 1
            Else
                Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 27) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18 '& mDoub1
                mHeader = mHeader + 1
            End If
        End If
        Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
        mHeader = mHeader + 1
        mFix = PageLength - (mHeader + mFooter)
        Page = 1
        mLine = 1
        mSlNo = 1
        LAdd = VNull(RstJob!Gen_Sur_Amt) + VNull(RstJob!Trans_Amt) + VNull(RstJob!Tax_Amt) + VNull(RstJob!Tax_Sur_Amt) + VNull(RstJob!Packing) + VNull(RstJob!ReSalTax_Amt) + VNull(RstJob!Tot_Amt)
        SubTot = RstJob!SprAmt_TB + RstJob!SprAmt_TP + RstJob!SprAmt_MRP_TB + RstJob!SprAmt_MRP_TP _
        + RstJob!OilAmt_TB + RstJob!OilAmt_TP + Val(txt(IWDiscTotTP).TEXT) + Val(txt(IWDiscTotTB).TEXT)
        If RstJob.RecordCount > 0 Then
            I = 1
            Do Until RstJob.EOF
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
                     
                    If mVatYn = 1 Then
                            Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "B", PageWidth)
                            mHeader = mHeader + 1
                    Else
                            Print #1, PRN_TIT("** WORKSHOP SPARE " & mDocStr & mDupStr & " **", "B", PageWidth)
                            mHeader = mHeader + 1
                    End If
                    Print #1, mChr18 & Space(40) & mEmph & PSTR(mDocStr & " No.", 20) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                    mHeader = mHeader + 1
                    Print #1, PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
                    mHeader = mHeader + 1
                    Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
                    mHeader = mHeader + 1
                   
                    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
                    mHeader = mHeader + 1
                    If mVatYn = 1 Then
                           Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR("DISC %", 8, , AlignRight) & PSTR("DISC.AMT", 10, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("TaxAmt", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18   '& mDoub1
                           mHeader = mHeader + 1
                    Else
                        If RstJob!Det_Tax = 1 Then
                            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                            mHeader = mHeader + 1
                            Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
                            mHeader = mHeader + 1
                        Else
                            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 27) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18 '& mDoub1
                            mHeader = mHeader + 1
                        End If
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
                    ElseIf RstJob!Purpose = "A" And Not StrCmp(left(PubComp_Name, 4), "Enar") Then
                        mTBAmtStr = "*AMC*"
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

                    If mVatYn = 1 Then
                        'PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 30) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                        'PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                        'PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(Val(RstJob!TaxPer), 6, 2) & PSTR(Val(RstJob!TaxAmt), 10, 2) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                        If UCase(left(PubComp_Name, 7)) = "SHANKAR" Or UCase(left(PubComp_Name, 6)) = "MAURYA" Then
                            If RstJob!Purpose <> "W" And RstJob!Purpose <> "F" Then
                                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 30) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                                PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                                PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(Val(RstJob!TaxPer), 6, 2) & PSTR(Val(RstJob!TaxAmt), 10, 2) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                            ElseIf RstJob!Purpose = "W" Or RstJob!Purpose = "F" Then
                                PrintStr = "": mSlNo = mSlNo - 1
                            End If
                        Else
                            PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 30) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                            PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                            PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(Val(VNull(RstJob!TaxPer)), 6, 2) & PSTR(Val(VNull(RstJob!TaxAmt)), 10, 2) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                        End If
                    Else
                        PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 34) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                        PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                        PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                    End If
  
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
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 28, , AlignLeft) & PSTR(RstJob!Part_Name, 40) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(LdRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", "L") & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & _
                    PSTR(LDAmt - RstJob!Disc_Amt2, 12, 2)
                End If
                If PrintStr <> "" Then
                    Print #1, PrintStr
                    'modi lps at Cuttack 31.08.03
                    
                    mLine = mLine + 1
                End If
                mSlNo = mSlNo + 1
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
        JobValue = 0
        If RSOJPR = True Then
            JobValue = JobValue + VNull(GCn.Execute("Select sum((Qty_Iss-Qty_Ret)*Rate) from sp_stock where Job_docID='" & txt(JobNo).Tag & "' and purpose <> 'C'").Fields(0).Value)
            JobValue = JobValue + VNull(GCn.Execute("Select sum(LabourAmt) from Job_Lab where Job_docID='" & txt(JobNo).Tag & "' and Chrg_Type <> 'C'").Fields(0).Value)
            JobValue = JobValue + Val(txt(NetAmt))
            Print #1, mChr18 & mSP2 & "Customer's Signature         Job Value : " & Format(JobValue, "0.00")
        Else
            Print #1, mChr18 & mSP2 & "Customer's Signature "
        End If
    ' SALE FOOTER
    '22 space maintain between heading and :
    RstJob.MoveFirst
    If RstJob!Det_Tax = 1 Then
        Print #1, Replace(Space(21), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
    
'        Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(Txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
'        ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
'
'        Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
'        ; " | " & PSTR("Tax Surc. ", 10, 0) & PSTR(RstJob!Tax_Sur_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Sur_Amt, 12, 2) & mDoub
        If mVatYn = 1 Then
            Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("V A T     ", 10, 0) & Space(6) & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
            
            Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
            ; " | " & IIf(RstJob!SatAmt_H > 0, PSTR("S A T     ", 10, 0) & Space(6) & PSTR(RstJob!SatAmt_H, 12, 2), Space(10) & Space(6) & Space(12)) & mDoub
        Else
            If RSOJPR = True Then
                Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt + mMRPTax, 12, 2) & mDoub
                
                Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
                ; "" & mDoub
            Else
                Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
                
                Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
                ; "" & mDoub
            End If
        End If
      
        Print #1, PSTR("Spares Amount", 16) & PSTR(RstJob!SprAmt_TP, 12, 2) & Space(8) & PSTR(RstJob!SprAmt_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Misc. Charges", 16) & PSTR(RstJob!Packing, 12, 2) & mDoub

'"Itemwise Dis.Amt 01234567.12 00.00% 01234567.12 | Itemwise Dis.Amt 01234567.12"
'col1(16) col2(28) col3(35) col4(47) ,col5(50) ,col6(66) ,col7(78)
'col1(16) col2(12) col3(7) col4(12) ,col5(3) ,col6(16) ,col7(12)

        Print #1, PSTR("Oil Amount ", 16) & PSTR(RstJob!OilAmt_TP + RstJob!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstJob!OilAmt_TB + RstJob!OilAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & mEmph & PSTR("Sub Total[TP+TB]", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
        
        Print #1, PSTR("Discount ", 10, 0) & PSTR(RstJob!D_Per_TP, 5, 2) & "%" & PSTR(RstJob!D_Amt_TP, 12, 2) & PSTR(RstJob!D_Per_TB, 7, 2) & "%" & PSTR(RstJob!D_Amt_TB, 12, 2) _
        ; " | " & PSTR("TO Tax ", 10, 0) & PSTR(RstJob!TOT_Per, 5, 2) & "%" & PSTR(RstJob!Tot_Amt, 12, 2) & mEmph
        
        Print #1, PSTR("Sub Total [A]", 16) & PSTR(Val(txt(STotATP)), 12, 2) & Space(8) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
        ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstJob!ReSalTax_Per, 5, 2) & "%" & PSTR(RstJob!ReSalTax_Amt, 12, 2)
        
        Print #1, PSTR("Gen Surch ", 10, 0) & PSTR(RstJob!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 12, 2) & PSTR(RstJob!Gen_Sur_Amt, 20, 2) _
        ; " | " & PSTR("Round Off", 16) & PSTR(RstJob!Rounded, 12, 2)
       
        Print #1, PSTR("Transportation", 16) & PSTR(0, 12, 2) & PSTR(RstJob!Trans_Amt, 20, 2) _
        ; " | " & mEmph & PSTR("Net Payble Rs.", 16) & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
    Else
        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        Print #1, Space(45) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
        If RstJob!D_Amt_TP + RstJob!D_Amt_TB > 0 Then
            Print #1, Space(45) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstJob!D_Amt_TP + RstJob!D_Amt_TB, 12, 2)
        Else
            Print #1, ""
        End If
        Print #1, Space(45) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Val(txt(NetSprAmt)) - (mGrossAmt - (RstJob!D_Amt_TP + RstJob!D_Amt_TB)), 12, 2) & mEmph
        Print #1, Space(45) & PSTR("Net Payble Rs.", 20) & " : " & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
    End If
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(Val(txt(NetSprAmt)), "Rupees", "Paise") & mDoub1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & "The service tax  amount claimed on this invoice will be paid to govt. as per statutory provision" & mChr18
    Print #1, mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
    Print #1, PSTR(mTaxdesc, 25) & Space(PageWidth - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1 & mDoub
    Print #1, ""
    Print #1, "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
       If mID(Footer, I, 1) = vbLf Then
           Print #1, RTrim(mID(Footer, j, I - j))
           j = I + 1
       End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
'Gate Pass Footer()
    If RstJob!Printed_YN = 0 Then
        If (RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "S") Then
            Print #1, Replace(Space(PageWidth), " ", "-")
            Print #1, PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", 80) & mEmph
            Print #1, "GATE PASS No. & DATE : " & XNull(RstJob!gp_no) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(10) & "Job Card No. : " & PrinID(RstJob!JobDocID)
            Print #1, "Vehicle No. : " & XNull(RstJob!RegNo) & Space(5) & "Chassis No. : " & XNull(RstJob!Chassis) _
            & Space(5) & mChr17 & "Model : " & XNull(RstJob!Model) & mChr18
            Print #1,
            Print #1, "Vehicle has been received from workshop & work done as per  my satisfaction."
            Print #1, ""
            Print #1, "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
            Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
            prnGatePass = True
        End If
    End If
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
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
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Spare Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & txt(JobNo).Tag & "'"
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintLab(mQry$)
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
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double, SrvTaxNo$
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double, mLabourAmt As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double
    
    Set RstJob = GCn.Execute(mQry)
    If RstJob.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
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
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
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
   'Service tax No Printing............
        SrvTaxNo = XNull(GCn.Execute("Select SrvTaxNo from Syctrl").Fields(0).Value)
        
        Print #1, mSP2 & PSTR("Serv.Tax No.  : " & SrvTaxNo, 40, , AlignLeft)
        mHeader = mHeader + 1
        '....................................
    Print #1, PRN_TIT("** LABOUR " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, mChr18 & Space(48) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocID_InvLab) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR(txt(OwnerName), 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(txt(Address1)), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
    mHeader = mHeader + 1
'    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
    Print #1, PSTR(XNull(txt(Address2)), 40) & Space(2) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(txt(Address3)) & IIf(XNull(txt(Address3)) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44) _
    & Space(4) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
    mHeader = mHeader + 1
    Print #1, mSP2 & "Phone : " & PSTR(XNull(RstJob!Phone), 20)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") '& mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("Srl.", 4) & "<-------------Labour Detail-------------- >" & " " & PSTR("Hrs", 10, , AlignRight) & PSTR("Rate", 10, , AlignRight) & PSTR("Amount", 12, , AlignRight)
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
                Print #1, PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
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
                Print #1, PSTR("Srl.", 4) & "<-------------Labour Detail-------------- >" & " " & PSTR("Hrs", 10, , AlignRight) & PSTR("Rate", 10, , AlignRight) & PSTR("Amount", 12, , AlignRight)
                mHeader = mHeader + 1
                Print #1, PSTR("No.", 4) & PSTR("Code", 7) & PSTR("Description", 35) '& mDoub1 & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            If RstJob!Chrg_From = "C" Then
                mLabourAmt = RstJob!LabourAmt  'Lab_Rate
            Else
                mLabourAmt = 0
            End If
            PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 4) & PSTR(RstJob!Lab_Code, 6, , AlignLeft) & " " & PSTR(RstJob!LabName, 35) & " " & PSTR(RstJob!Hrs_Taken, 10, 2) & PSTR(RstJob!Lab_Rate, 10, 2) & PSTR(mLabourAmt, 12, 2)
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
    JobValue = 0
    If RSOJPR = True Then
        JobValue = JobValue + VNull(GCn.Execute("Select sum((Qty_Iss-Qty_Ret)*Rate) from sp_stock where Job_docID='" & txt(JobNo).Tag & "' and purpose <> 'C'").Fields(0).Value)
        JobValue = JobValue + VNull(GCn.Execute("Select sum(LabourAmt) from Job_Lab where Job_docID='" & txt(JobNo).Tag & "' and Chrg_Type <> 'C'").Fields(0).Value)
        JobValue = JobValue + Val(txt(NetAmt))
        Print #1, mChr18 & mSP2 & "Customer's Signature         Job Value : " & Format(JobValue, "0.00")
    Else
        Print #1, mChr18 & mSP2 & "Customer's Signature "
    End If
' SALE FOOTER
    RstJob.MoveFirst
    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
    Print #1, Space(45) & PSTR("TOTAL AMOUNT", 20) & " : " & PSTR(RstJob!LabAmt_TB + RstJob!LabAmt_TP, 12, 2) & mDoub1
    
    If RstJob!Lab_D_Amt > 0 Then
        Print #1, Space(45) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstJob!Lab_D_Amt, 12, 2)
    Else
        Print #1, ""
    End If
    Print #1, Space(45) & PSTR("SERVICE TAX @" & STR(RstJob!Lab_TaxPer), 20) & " : " & PSTR(RstJob!Lab_TaxAmt, 12, 2)
    Print #1, Space(45) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(RstJob!Lab_RoundOff, 12, 2) & mEmph
    Print #1, Space(45) & PSTR("Net Payble Rs.", 20) & " : " & PSTR(RstJob!NetLab_Amt, 12, 2) & mEmph1

    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(RstJob!NetLab_Amt, "Rupees", "Paise") & mDoub1
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & "The service tax  amount claimed on this invoice will be paid to govt. as per statutory provision" & mChr18
    Print #1, mChr17 & "E & O.E." & mChr18 & Space(PageWidth - (Len("For " & PubComp_Name) + 6)) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, "" & mDoub
    Print #1, "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
'Gate Pass Footer()
    'If RstJob!LabBillPrinted = 0 Then
        If (RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "L") Or prnGatePass = False Then
            Print #1, Replace(Space(PageWidth), " ", "-")
            Print #1, PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", 80) & mEmph
            Print #1, "GATE PASS No. & DATE : " & XNull(RstJob!gp_no) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(10) & "Job Card No. : " & PrinID(RstJob!JobDocID)
            Print #1, "Vehicle No. : " & XNull(RstJob!RegNo) & Space(5) & "Chassis No. : " & XNull(RstJob!Chassis) _
            & Space(5) & mChr17 & "Model : " & XNull(RstJob!Model) & mChr18
            Print #1,
            Print #1, "Vehicle has been received from workshop & work done as per  my satisfaction."
            Print #1, ""
            Print #1, "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
            Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
     '   End If
    End If
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
    If MsgBox("Labour Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update Job_Card set LabBillPrinted = 1 where Job_Card.DocId='" & txt(JobNo).Tag & "'"
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
    Dim I As Integer, j As Integer
    Dim PrintStr As String
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
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
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
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
    Print #1, mSP2 & mEmph & PSTR(txt(OwnerName), 40) & mEmph1
    mHeader = mHeader + 1
    If txt(Address1) <> "" Then
        Print #1, mSP2 & PSTR(txt(Address1), 40) & mEmph1
        mHeader = mHeader + 1
    End If
    If txt(Address2) <> "" Then
        Print #1, mSP2 & PSTR(txt(Address2), 40) & mEmph1
        mHeader = mHeader + 1
    End If
    If txt(Address3) <> "" Then
        Print #1, mSP2 & PSTR(txt(Address3), 40) & mEmph1
        mHeader = mHeader + 1
    End If
    Print #1, mSP2 & PSTR(txt(City), 40) & mEmph1
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "Dear Vehicle Owner,"
    mHeader = mHeader + 1
    Print #1, mSP2 & "We are pleased to attend your " & txt(Model) & " Vehicle at our workshop and hope"
    mHeader = mHeader + 1
    Print #1, mSP2 & "you are satisfied with our working & services."
    mHeader = mHeader + 1
    Print #1, mSP2 & ""
    mHeader = mHeader + 1
    Print #1, mSP2 & "Please remember,next service of your vehicle "
    mHeader = mHeader + 1
    Print #1, mSP2 & Space(10) & "Reg.No.    :" & txt(VehRegNo)
    mHeader = mHeader + 1
    Print #1, mSP2 & Space(10) & "Chassis No.:" & txt(Chassis)
    mHeader = mHeader + 1
    Print #1, mSP2 & " should be carried out before dt." & txt(NextSrv) & "."
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




Private Sub SpeedPrintBothJMK(mQry$, PrePrinted As Boolean)
'On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per Page 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18regno

    Dim Party As String
    
    Dim I As Integer, j As Integer, K As Integer
    Dim PrintStr As String
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double, SrvTaxNo$
    Dim SrvGatePassOn$, Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double
    Dim MRPTaxStr$, mTPAmtStr$, mTBAmtStr$, mStrLab$
    Dim mSprCaption As Boolean, mLabCaption As Boolean, mLabDiscAmtStr$
    Dim mTotRow, mTotRowTemp As Integer
    Dim tmprs As ADODB.Recordset
    Dim HlpLineNo As String

    
    Set RstJob = GCn.Execute(mQry)
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
    mFooter = 35
    'eof modi
    'Sale Bill Header
    If Not Provisional Then
        If mVatYn = 1 Then
            If StrCmp(txt(CashBill), "YES") Then
                mDocStr = "SALE INVOICE [CASH]"
            Else
                If GCn.Execute("Select * From Subgroup S Left Join SubGroupType St On S.Party_Type=St.Party_Type Where S.SubCode='" & txt(SpareParty).Tag & "' And St.Description='Dealer' ").RecordCount > 0 Then
                    mDocStr = "TAX INVOICE"
                Else
                    mDocStr = "SALE INVOICE [CREDIT]"
                End If
            End If
        Else
            If RstJob!CrMemo = 0 Then
                mDocStr = "CASH MEMO"
            Else
                mDocStr = "INVOICE  "
            End If
        End If
    Else
        mDocStr = "PROVISIONAL BILL "
    End If
    Party = IIf(RstJob!CrMemo = 1, txt(SpareParty).TEXT, txt(OwnerName).TEXT)
    
    If Not Provisional Then
        If UCase(pubUName) <> "SA" Then
            mDupStr = IIf(RstJob!Printed_YN = 1, "(DUPLICATE)", "")
            If mDupStr <> "" Then
                MsgBox "Second Printing Can Be done Only By SA."
                Exit Sub
            End If
        End If
    Else
        mDupStr = ""
    End If
    If (mMRPTax + mMRPTaxSur + mMRPTOT) > 0 Then
        MRPTaxStr = "* Note:"
        If (mMRPTax + mMRPTaxSur) > 0 Then
            MRPTaxStr = MRPTaxStr & "Sales Tax Rs." & mMRPTax & ",Surcharge Rs." & mMRPTaxSur
        End If
        If (mMRPTOT) > 0 Then
            'MRPTaxStr = MRPTaxStr & pubTOTCaption & mMRPTOT
        End If
        MRPTaxStr = MRPTaxStr & " already added in MRP *'"
    End If
    If GCn.Execute("select Printing_Desc from TaxForms where Form_Code = '" & RstJob!Form_Code & "'").RecordCount > 0 Then
        mTaxdesc = GCn.Execute("select Printing_Desc from TaxForms where Form_Code = '" & RstJob!Form_Code & "'").Fields(0).Value
    End If
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    
    
    Set tmprs = GCn.Execute("Select HelpLineNo from Syctrl")
    If tmprs.RecordCount > 0 Then
        HlpLineNo = IIf(IsNull(tmprs!HelpLineNo), "", Trim(tmprs!HelpLineNo))
        Set tmprs = Nothing
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
        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecLST) & IIf(XNull(RstCompDet!W_SecLST_Date) = "", "", " Dt. " & RstCompDet!W_SecLST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!W_SecPhone)), 28, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(RstCompDet!W_SecCST) & IIf(XNull(RstCompDet!W_SecCST_Date) = "", "", " Dt. " & RstCompDet!W_SecCST_Date), 45) & Space(7) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!W_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        'Service tax No Printing............
        SrvTaxNo = XNull(GCn.Execute("Select SrvTaxNo from Syctrl").Fields(0).Value)
        Print #1, mSP2 & PSTR("Serv.Tax No.:" & SrvTaxNo, 50, , AlignLeft) & PSTR("HelpLine No :" & HlpLineNo, 30, , AlignRight)
        mHeader = mHeader + 1
        '....................................
        If mVatYn = 1 Then
            Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "B", PageWidth)
            mHeader = mHeader + 1
        Else
            Print #1, PRN_TIT("** WORKSHOP SPARE/LABOUR " & mDocStr & mDupStr & " **", "B", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, mSP2 & mChr18 & "TO," & Space(39) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " :  " & Right(RstJob!DocId_InvSpr, 7) & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP2 & mEmph & PSTR(RstJob!NamePrefix & " " & Party, 43) & mEmph1 & PSTR("DATE", 12, , AlignRight) & "          : " & Format(RstJob!V_DATE, "dd/MMM/yyyy")
        mHeader = mHeader + 1
        
        Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(11) & PSTR("Job Card No.", 12) & "  :       " & (RstJob!Job_No)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address3)) & IIf(XNull(txt(Address3)) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44)
        mHeader = mHeader + 1
        Print #1, mSP2 & "Phone : " & PSTR(XNull(txt(PhoneOff)), 40) & "Mobile : " & PSTR(XNull(txt(Mobile)), 40)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR("Chass.No.", 11) & PSTR(XNull(RstJob!Chassis), 20) & Space(1) & Space(6) & PSTR(XNull(RstJob!Model), 13) & Space(1) & PSTR("Reg.", 5) & PSTR(XNull(RstJob!RegNo), 14) & " Kms:" & PSTR(XNull(txt(CurrentKMS)), 6)
        mHeader = mHeader + 1
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
        mHeader = mHeader + 1
        If mVatYn = 1 Then
            Print #1, mSP2 & PSTR("Srl", 4) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 24) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP Rate", 11, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR(" DISC%", 6, , AlignRight) & PSTR("DISC. AMT", 12, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("Tax Amt", 9, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18  '& mDoub1
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
        If RstJob!orig = "1" Then
            Print #1, mSP2 & mEmph & "*Spare Details*" & mEmph1 & mChr17
            mHeader = mHeader + 1
            mSprCaption = True
        ElseIf RstJob!orig = "2" Then
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
        + RstJob!OilAmt_TB + RstJob!OilAmt_TP + Val(txt(IWDiscTotTP).TEXT) + Val(txt(IWDiscTotTB).TEXT)
        mTotRow = RstJob.RecordCount
        mTotRowTemp = RstJob.RecordCount
        If RstJob.RecordCount > 0 Then
            I = 1
            Do Until RstJob.EOF = True
                If mTotRow > 30 Then
                    mFix = 30
                ElseIf mTotRow >= 15 And mTotRow <= 30 Then
                    mFix = 30
                Else
                    mFix = (PageLength - (mHeader + mFooter))
                End If
                'mFix = PageLength - (mHeader + mFooter)
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
        
                    Print #1, mSP2 & PSTR("Serv.Tax No.:" & SrvTaxNo, 40, , AlignLeft)
                    mHeader = mHeader + 1
                    If mVatYn = 1 Then
                        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                    Else
                        Print #1, PRN_TIT("** WORKSHOP SPARE/LABOUR " & mDocStr & mDupStr & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                    End If
                    Print #1, ""
                    mHeader = mHeader + 1
                    Print #1, mSP2 & mChr18 & "TO," & Space(39) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " :    " & Right(RstJob!DocId_InvSpr, 6) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & mEmph & PSTR(RstJob!NamePrefix & " " & Party, 43) & mEmph1 & PSTR("DATE", 12, , AlignRight) & "          : " & Format(RstJob!V_DATE, "dd/MMM/yyyy")
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(11) & PSTR("Job Card No.", 12) & "  :       " & (RstJob!Job_No)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(txt(Address3)) & IIf(XNull(txt(Address3)) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & "Phone : " & PSTR(XNull(txt(PhoneOff)), 40)
                    mHeader = mHeader + 1
                    'Print #1, mSP2 & PSTR("Chassis No.", 11) & PSTR(XNull(RstJob!Chassis), 15) & Space(1) & PSTR("Model.", 6) & PSTR(XNull(RstJob!Model), 15) & Space(1) & PSTR("Reg.No.", 7) & PSTR(XNull(RstJob!RegNo), 14) & " Kms:" & PSTR(XNull(RstJob!AtKMsHrs), 6)
                    'mHeader = mHeader + 1
                    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
                    mHeader = mHeader + 1
                    If mVatYn = 1 Then
                        Print #1, mSP2 & PSTR("Srl", 4) & PSTR("Req.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 24) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP Rate", 11, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR(" DISC%", 6, , AlignRight) & PSTR("DISC. AMT", 12, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("Tax Amt", 9, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18  '& mDoub1
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
                    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17
                    mHeader = mHeader + 1
                    '....................................
                    
                    mHeader = mHeader + 1
                    mFix = PageLength - (mHeader + mFooter)
                    mLine = 1
                    
                End If
                If mSprCaption = False Then
                    If RstJob!orig = "1" Then
                        Print #1, mSP2 & mChr18 & mEmph & "*Spare Details*" & Replace(Space(PageWidth - 16), " ", "-") & mEmph1 & mChr17
                        mHeader = mHeader + 1
                        mSprCaption = True
                    End If
                End If
                mRate = IIf(RstJob!MRP_YN = 1, RstJob!MRP_Rate2, RstJob!Rate2)
                If RstJob!orig = "1" Then
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
                    If mVatYn = 1 Then
                        PrintStr = PSTR(Trim(STR(mSlNo)), 4) & PSTR(RstJob!ReqNo, 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 24) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3) & PSTR(RstJob!Unit, 8, , AlignLeft)
                        PrintStr = PrintStr & PSTR(VNull(RstJob!MRP_Rate2), 7, 2) & PSTR(mRate, 8, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                        PSTR(RstJob!Disc_Per2, 5, 2) & "%" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(VNull(RstJob!TaxPer), 6, 2) & PSTR(VNull(RstJob!TaxAmt), 9, 2) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                    Else
                        PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!ReqNo, 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 27) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
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
                If RstJob!orig = "2" Then
                    If RstJob!Purpose = "W" Then
                        mTBAmtStr = "*Warranty*"
                    ElseIf RstJob!Purpose = "P" Then
                        mTBAmtStr = "*PDI*"
                    ElseIf RstJob!Purpose = "A" And Not StrCmp(left(PubComp_Name, 4), "Enar") Then
                        mTBAmtStr = "*AMC*"
                    ElseIf RstJob!Purpose = "F" Then
                        mTBAmtStr = "*Free*"
                    Else
                        If Val(txt(ServTaxAmt)) <= 0 Then
                            mTPAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                            mTBAmtStr = PSTR(0, 12, 2)
                        Else
                            mTPAmtStr = PSTR(0, 12, 2)
                            mTBAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                        End If
                    End If
                End If
                
                If mVatYn = 1 Then
                    PrintStr = PSTR(Trim(STR(mSlNo)), 4) & Space(7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 24) & PSTR(Format(RstJob!Qty_Iss - RstJob!Qty_Ret, "0.000"), 12, 3, AlignRight) & PSTR(RstJob!Unit, 8, , AlignLeft)
                    PrintStr = PrintStr & PSTR("--", 9, 2, AlignRight) & PSTR(mRate, 5, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                    PSTR(RstJob!Disc_Per2, 5, 2) & "%" & PSTR(RstJob!Disc_Amt2, 10, 2) & Space(6) & Space(9) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                Else
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & Space(7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 27)
                    PrintStr = PrintStr & Space(48) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                End If
            End If
            If PrintStr <> "" Then
                Print #1, mSP2 & PrintStr
                mLine = mLine + 1
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
                        Print #1, mSP2 & PRN_TIT("** WORKSHOP " & mDocStr & mDupStr & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                        Print #1, mSP2 & mChr18 & Space(46) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                        mHeader = mHeader + 1
                        Print #1, mSP2 & PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
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
    
    
    
    
    Dim mVat12 As Double
    Dim mVat4 As Double
    Dim mSaleVat12 As Double
    Dim mSaleVat4 As Double
    Dim mSatAmt As Double
    
    
' SALE FOOTER
    '22 space maintain between heading and :
    RstJob.MoveLast
    mVat12 = GCn.Execute("Select " & vIsNull("Sum(TaxAmt)", "0") & " From Sp_Stock Where TaxPer>=12.5 and Job_DocId='" & XNull(RstJob!JobDocID) & "'").Fields(0).Value
    mVat4 = GCn.Execute("Select " & vIsNull("Sum(TaxAmt)", "0") & " From Sp_Stock Where TaxPer<12 and Job_DocId='" & XNull(RstJob!JobDocID) & "'").Fields(0).Value
    mSaleVat12 = GCn.Execute("Select " & vIsNull("Sum(Net_Amt)", "0") & " From Sp_Stock Where TaxPer>=12.5 and Job_DocId='" & XNull(RstJob!JobDocID) & "'").Fields(0).Value
    mSaleVat4 = GCn.Execute("Select " & vIsNull("Sum(Net_Amt)", "0") & " From Sp_Stock Where TaxPer<12 and Job_DocId='" & XNull(RstJob!JobDocID) & "'").Fields(0).Value
    mSatAmt = GCn.Execute("Select " & vIsNull("Sum(SatAmt)", "0") & " From Sp_Stock Where Job_DocId='" & XNull(RstJob!JobDocID) & "'").Fields(0).Value
    
    'If mTotRow <= 15 Then
    If ChkRep(0).Value = vbChecked Then
        If RstJob!Det_Tax = 1 Then
            If mVatYn = 1 Then
                Print #1, mSP2 & Replace(Space(35), " ", "-") & "Taxable Amt" & Replace(Space(33), " ", "-")
                If mVatYn = 1 Then
'                    Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & Space(19) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
'                    ; " | " & PSTR("V A T     ", 10, 0) & Space(6) & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
                    
                    Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & Space(19) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                    ; " | " & PSTR("V A T 12.5% ", 11, 0) & Space(5) & PSTR(mVat12, 12, 2) & mDoub
                    
                    Print #1, mSP2 & PSTR("MRP Items Amt", 16) & Space(19) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
                    ; " | " & PSTR("V A T 4%", 16) & PSTR(mVat4, 12, 2) & mDoub
                    
                Else
                    Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & Space(19) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                    ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
                    
                    Print #1, mSP2 & PSTR("MRP Items Amt", 16) & Space(19) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
                    '; " | " & PSTR("Misc. Charges", 16) & PSTR(RstJob!Packing, 12, 2) & mDoub
                End If
                                                    
                
                Print #1, mSP2 & Space(47) & mDoub1 _
                ; " | " & PSTR("S A T", 16) & PSTR(mSatAmt, 12, 2) & mDoub
                '; " | " & mEmph & PSTR("Sub Total", 16) & PSTR(Val(Txt(STotB)), 12, 2) & mEmph1
                
                Print #1, mSP2 & PSTR("Spares Amount", 16) & Space(19) & PSTR(RstJob!SprAmt_TB, 12, 2) & mDoub1 _
                ; " | " & PSTR("Fuel Charges", 16) & PSTR(RstJob!Packing, 12, 2) & mDoub
                '; " | " & mEmph & PSTR("Sub Total", 16) & PSTR(Val(Txt(STotB)), 12, 2) & mEmph1
        
                Print #1, mSP2 & PSTR("Oil Amount ", 16) & Space(19) & PSTR(RstJob!OilAmt_TB + RstJob!OilAmt_MRP_TB, 12, 2) & mDoub1 _
                ; " | " & mEmph & PSTR("Sub Total", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
                '; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstJob!TOT_Per, 5, 2) & "%" & PSTR(RstJob!Tot_Amt, 12, 2)
                
                Print #1, mSP2 & PSTR("Discount ", 16) & Space(11) & PSTR(RstJob!D_Per_TB, 7, 2) & "%" & PSTR(RstJob!D_Amt_TB, 12, 2) _
                ; " | " & PSTR("Round Off", 16) & PSTR(Round(RstJob!Rounded, 2), 12, 2)
                
                Print #1, mSP2 & PSTR("Sub Total [A]", 16) & Space(19) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
                ; " | " & mEmph & PSTR("Net Spare + Lub. Rs.", 16) & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
            Else
                Print #1, mSP2 & Replace(Space(20), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
                If mVatYn = 1 Then
                    Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                    ; " | " & PSTR("Tax      ", 10, 0) & Space(6) & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
                    
                    Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
                    ; " | " & Space(10) & Space(6) & Space(12) & mDoub
                Else
                    Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                    ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
                    
                    Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2); " | " & mDoub1 _
                    & PSTR("Fuel Charges", 16) & PSTR(RstJob!Packing, 12, 2) & mDoub
                End If
                
                Print #1, mSP2 & PSTR("Spares Amount", 16) & PSTR(RstJob!SprAmt_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_TB, 12, 2) & mDoub1 _
                ; " | " & mEmph & PSTR("Sub Total[TP+TB]", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
        
                Print #1, mSP2 & PSTR("Oil Amount ", 16) & PSTR(RstJob!OilAmt_TP + RstJob!OilAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!OilAmt_TB + RstJob!OilAmt_MRP_TB, 12, 2) & mDoub1 _
                ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstJob!TOT_Per, 5, 2) & "%" & PSTR(RstJob!Tot_Amt, 12, 2)
                
                Print #1, mSP2 & PSTR("Discount ", 10, 0) & PSTR(RstJob!D_Per_TP, 5, 2) & "%" & PSTR(RstJob!D_Amt_TP, 11, 2) & PSTR(RstJob!D_Per_TB, 7, 2) & "%" & PSTR(RstJob!D_Amt_TB, 12, 2) _
                ; " | " & PSTR("Round Off", 16) & PSTR(Round(RstJob!Rounded, 2), 12, 2)
                
                Print #1, mSP2 & PSTR("Sub Total [A]", 16) & PSTR(Val(txt(STotATP)), 11, 2) & Space(8) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
                ; " | " & mEmph & PSTR("Net Spare + Lub. Rs.", 16) & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
            End If
        Else
            Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mDoub
            Print #1, mSP2 & Space(44) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
            If RstJob!D_Amt_TP + RstJob!D_Amt_TB > 0 Then
                Print #1, mSP2 & Space(44) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstJob!D_Amt_TP + RstJob!D_Amt_TB, 12, 2)
            Else
                Print #1, ""
            End If
            Print #1, mSP2 & Space(44) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Val(txt(NetSprAmt)) - (mGrossAmt - (RstJob!D_Amt_TP + RstJob!D_Amt_TB)), 12, 2) & mEmph
            Print #1, mSP2 & Space(44) & PSTR("Net Spare + Lub. Rs.", 20) & " : " & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
        End If
    End If

    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    
    If ChkRep(1).Value = vbChecked Then
    
        If Val(txt(LabDisc)) > 0 Then
            mLabDiscAmtStr = "Discount  : " & PSTR(Val(txt(LabDisc)), 8, 2)
        Else
            mLabDiscAmtStr = Space(19)
        End If
    
        PrintStr = mEmph & "Total Lab. :" & PSTR(Val(txt(LabAmt)), 8, 2)
        PrintStr = PrintStr & " |Serv.Tax @ " & PSTR(Val(txt(ServTaxPer)), 5, 2) & ":" & PSTR(Val(txt(ServTaxAmt)), 9, 2) & "|" & "Net Labour Rs.    : " & PSTR(Val(txt(NetLabAmt)), 9, 2)
        Print #1, mSP2 & PrintStr
        PrintStr = mLabDiscAmtStr
        If ChkRep(0).Value = vbChecked Then
            PrintStr = PrintStr & "  |" & "Round Off      :  " & PSTR(Val(txt(LabROff)), 7, 2) & " |" & "Net Payble Amt Rs.: " & PSTR(Val(txt(NetAmt)), 9, 2) & mEmph1
        Else
            PrintStr = PrintStr & "  |" & "Round Off      :  " & PSTR(Val(txt(LabROff)), 7, 2) & " |" & "Net Payble Amt Rs.: " & PSTR(Val(txt(NetLabAmt)), 9, 2) & mEmph1
        End If
        
        Print #1, mSP2 & PrintStr
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
        
    End If
    
    If ChkRep(0).Value = vbChecked And ChkRep(1).Value = vbChecked Then
        Print #1, mSP2 & mDoub & ntow(txt(NetAmt), "Rupees", "Paise") & mDoub1
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
        Print #1, mChr17 & "The service tax  amount claimed on this invoice will be paid to govt. as per statutory provision" & mChr18
    ElseIf ChkRep(0).Value = vbChecked And ChkRep(1).Value = vbUnchecked Then
        Print #1, mSP2 & mDoub & ntow(txt(NetSprAmt), "Rupees", "Paise") & mDoub1
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    ElseIf ChkRep(0).Value = vbUnchecked And ChkRep(1).Value = vbChecked Then
        Print #1, mSP2 & mDoub & ntow(txt(NetLabAmt), "Rupees", "Paise") & mDoub1
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
        Print #1, mChr17 & "The service tax  amount claimed on this invoice will be paid to govt. as per statutory provision" & mChr18
    End If
        
    If mVatYn = 1 Then
        Print #1, ""
    Else
        Print #1, mSP2 & mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
    End If
    Print #1, mSP2 & PSTR(mTaxdesc, 25) & Space(PageWidth - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1
    'Print #1, ""
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
    If Not Provisional Then
        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "S" Then
            Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
            Print #1, mSP2 & PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", (PageWidth)) & mEmph
            Print #1, mSP2 & "Vehicle No. : " & XNull(RstJob!RegNo) & Space(10) & "GATE PASS No. & DATE : " & XNull(RstJob!gp_no) & "  " & XNull(RstJob!GP_Date) & mEmph1
            Print #1, mSP2 & "Chassis No. : " & XNull(RstJob!Chassis) & Space(6) & "Job Card No..........: " & Right(RstJob!JobDocID, 6)
            Print #1, mSP2 & "Model       : " & XNull(RstJob!Model)
            Print #1, mSP2 & "Vehicle has been received from workshop & work done as per  my satisfaction."
            Print #1, ""
            Print #1, mSP2 & "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
            Print #1, mSP2 & mChr17 & Space((((PageWidth) * 1.7) - Len("* a dataman software *" & "   " & pubUName & "   " & PubServerDate)) / 2) & "* a dataman software *" & "   " & pubUName & "   " & PubServerDate & mChr18
        End If
    End If
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
    If Provisional Then
        MsgBox "Provisional Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !"
    Else
       GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & txt(JobNo).Tag & "'"
       GCn.Execute "update Job_Card set LabBillPrinted = 1 where DocID='" & txt(JobNo).Tag & "'"
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section

End Sub

Private Sub WindowsPrintSpr(Index As Integer, mQry$)
Dim I As Integer, RstJob As ADODB.Recordset, RST1 As ADODB.Recordset, mDocStr$
Dim mPrintGatePass As Byte
On Error GoTo ERRORHANDLER
Set RstJob = GCn.Execute(mQry)
If RstJob.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

CreateFieldDefFile RstJob, PubRepoPath + "\" & mRepName & ".ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath & "\" & mRepName & ".RPT")
rpt.Database.SetDataSource RstJob
rpt.ReadRecords
RstJob.MoveFirst
Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram,W_SecPAN_No from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
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
                Case UCase("CompPanNo")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPAN_No & "'"
                
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
            GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & txt(JobNo).Tag & "'"
        End If
End Select
CmdPrint(0).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrintLab(Index As Integer, mQry$)
Dim I As Integer, Rst As ADODB.Recordset, RST1 As ADODB.Recordset, mDocStr$
Dim mPrintGatePass As Byte
'On Error GoTo ERRORHANDLER

Set Rst = GCn.Execute(mQry)
'If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName1 & ".ttx", True
    Set rpt = rdApp.OpenReport(PubRepoPath & "\" & mRepName1 & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Rst.MoveFirst
    If ChkRep(ChkLabInv).Value = Unchecked Then Exit Sub
    Set RST1 = New Recordset
    RST1.CursorLocation = adUseClient
    RST1.Open "select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram,W_SecPAN_No from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
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
             Case UCase("CompPanNo")
                rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPAN_No & "'"
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
            GCn.Execute "update Job_Card set LabBillPrinted = 1 where Job_Card.DocId='" & txt(JobNo).Tag & "'"
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
    "LEFT JOIN SubGroup ON Sp_Sale.Party_Code = SubGroup.SubCode) LEFT JOIN (Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) ON Sp_Sale.Job_DocID = Job_Card.DocId) LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN Syctrl ON Syctrl.LinkTable  >= Sp_Sale.U_AE " & _
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
            GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & txt(JobNo).Tag & "'"
        End If
End Select
CmdPrint(0).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub SpeedPrintOthDlr(mQry$)
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
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double, SrvTaxNo$
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double, mLabourAmt As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mTotLab As Double
    Dim mTotAmt, mServPer, mServAmt As Double, AmtTmp As Double
    Set RstJob = GCn.Execute(mQry)
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
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
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
    mDocStr = "CREDIT INVOICE"
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
   'Service tax No Printing............
        SrvTaxNo = XNull(GCn.Execute("Select SrvTaxNo from Syctrl").Fields(0).Value)
        
        Print #1, mSP2 & PSTR("Serv.Tax No.  : " & SrvTaxNo, 40, , AlignLeft)
        mHeader = mHeader + 1
        '....................................
    Print #1, PRN_TIT("** LABOUR " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, mChr18 & Space(48) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocID_InvLab) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("To,", 48) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(8) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
    mHeader = mHeader + 1
'    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("Reg. No.", 12) & " : " & XNull(RstJob!RegNo)
  '  Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("  Kms:", 12) & "   : " & XNull(RstJob!AtKMsHrs)
    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(8) & PSTR("  Kms:", 12) & "   : " & XNull(RstJob!KM)
     
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
    & Space(4) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
    mHeader = mHeader + 1
    Print #1, mSP2 & "Phone : " & PSTR(XNull(RstJob!Phone), 20)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") '& mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("Srl.", 4) & "<-------------Labour Detail-------------- >" & " " & PSTR("Hrs", 10, , AlignRight) & PSTR("Rate", 10, , AlignRight) & PSTR("Amount", 12, , AlignRight)
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
                Print #1, PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(4) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
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
                Print #1, PSTR("Srl.", 4) & "<-------------Labour Detail-------------- >" & " " & PSTR("Hrs", 10, , AlignRight) & PSTR("Rate", 10, , AlignRight) & PSTR("Amount", 12, , AlignRight)
                mHeader = mHeader + 1
                Print #1, PSTR("No.", 4) & PSTR("Code", 7) & PSTR("Description", 35) '& mDoub1 & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
                If RstJob!Chrg_Type = "F" Then
                    mLabourAmt = RstJob!LabourAmt  'Lab_Rate
                    mServPer = VNull(RstJob!Lab_TaxPer)
                    mTotLab = mTotLab + mLabourAmt
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 4) & PSTR(RstJob!Lab_Code, 6, , AlignLeft) & " " & PSTR(RstJob!LabName, 35) & " " & PSTR(RstJob!Hrs_Taken, 10, 2) & PSTR(RstJob!Lab_Rate, 10, 2) & PSTR(mLabourAmt, 12, 2)
                    Print #1, PrintStr
                Else
                    mLabourAmt = 0
                End If
            Else
                mTotLab = VNull(RstJob!Coupon_Value)
                mServPer = VNull(GCn.Execute("Select Service_tax from Syctrl").Fields(0).Value)
                PrintStr = PSTR("Labour Charges", 65) & PSTR(mTotLab, 12, 2)
                Print #1, PrintStr
            End If
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
    Print #1, Space(45) & PSTR("TOTAL AMOUNT", 20) & " : " & PSTR(mTotLab, 12, 2) & mDoub1
'    If UCase(XNull(RstJob!CityName)) = UCase(PubComp_City) Then ' Or (XNull(RstJob!D_Code) <> PubDealerID And XNull(RstJob!D_Code) <> "") Then
'        AmtTmp = mTotLab * 20 / 100
'        Print #1, Space(40) & PSTR("SAME CITY DEBIT @ 20%", 25) & " : " & PSTR(AmtTmp, 12, 2) & mDoub1
'        mTotLab = mTotLab + AmtTmp
'    End If
    
    mServAmt = (mTotLab * mServPer / 100)
    Print #1, Space(45) & PSTR("SERVICE TAX @" & STR(mServPer), 20) & " : " & PSTR(mServAmt, 12, 2)
    'Print #1, Space(45) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(RstJob!Lab_RoundOff, 12, 2) & mEmph
    mTotAmt = Format(mTotLab + mServAmt, "0.00")
    Print #1, Space(45) & mDoub & PSTR("Net Payble Rs.", 20) & " : " & PSTR(mTotAmt, 12, 2, AlignRight) & mEmph1 & mDoub1

    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(mTotAmt, "Rupees", "Paise") & mDoub1
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & "The service tax  amount claimed on this invoice will be paid to govt. as per statutory provision" & mChr18
    Print #1, mChr17 & "E & O.E." & mChr18 & Space(PageWidth - (Len("For " & PubComp_Name) + 6)) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, "" & mDoub
    Print #1, "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
'Gate Pass Footer()
    If RstJob!LabBillPrinted = 0 Then
        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "L" Then
            Print #1, Replace(Space(PageWidth), " ", "-")
            Print #1, PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", 80) & mEmph
            Print #1, "GATE PASS No. & DATE : " & XNull(RstJob!gp_no) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(10) & "Job Card No. : " & PrinID(RstJob!JobDocID)
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
    If MsgBox("Labour Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update Job_Card set LabBillPrinted = 1 where Job_Card.DocId='" & txt(JobNo).Tag & "'"
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintBoth(mQry$, PrePrinted As Boolean)
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
    Dim mSprCaption As Boolean, mLabCaption As Boolean, mLabDiscAmtStr$
    Dim mTotRow, mTotRowTemp As Integer
    Dim RsTemp As ADODB.Recordset
    
    Set RstJob = GCn.Execute(mQry)
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
    mFooter = 22    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
    mGatePass = 9
    mDetTax = 15
    mFooter = IIf(RstJob!Det_Tax = 1, mFooter, mDetTax)
    mFooter = mFooter + FooterCnt + mGatePass
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
        If mVatYn = 1 Then
           
           If RSOJPR = True Or left(PubComp_Name, 10) = "GANGANAGAR" Then
                mDocStr = "VAT INVOICE"
           Else
                mDocStr = "RETAIL INVOICE"
           End If
           If UCase(left(PubComp_Name, 5)) = "UJWAL" Or UCase(left(PubComp_Name, 3)) = "LMP" Then
               mDocStr = "TAX INVOICE"
           End If
        Else
            If RstJob!CrMemo = 0 Then
                mDocStr = "CASH MEMO"
            Else
                mDocStr = "INVOICE"
            End If
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
            MRPTaxStr = MRPTaxStr & pubTOTCaption & mMRPTOT
        End If
        MRPTaxStr = MRPTaxStr & " already added in MRP *'"
    End If
    Set RsTemp = GCn.Execute("select Printing_Desc from TaxForms where Form_Code = '" & RstJob!Form_Code & "'")
    If RsTemp.RecordCount > 0 Then mTaxdesc = XNull(RsTemp!Printing_Desc)


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
            Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "B", PageWidth)
            mHeader = mHeader + 1
        Else
            Print #1, PRN_TIT("** WORKSHOP SPARE " & mDocStr & mDupStr & " **", "B", PageWidth)
            mHeader = mHeader + 1
        End If
        If Not Provisional Then
            If mVatYn = 1 Then
              
               If RSOJPR = True Or left(PubComp_Name, 10) = "GANGANAGAR" Then
                    mDocStr = "VAT INVOICE"
               Else
                    mDocStr = "RETAIL INVOICE"
               End If
                If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
                        mDocStr = "TAX INVOICE"
               End If
            Else
                If RstJob!CrMemo = 0 Then
                    mDocStr = "CASH MEMO"
                Else
                    mDocStr = "INVOICE"
                End If
            End If
        Else
            mDocStr = "PROVISIONAL BILL "
        End If
        
        If RSOJPR = True And VNull(RstJob!LastInvNoSuff) > 0 Then
            Print #1, mSP2 & mChr18 & Space(36) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " : " & PrinID(RstJob!DocId_InvSpr) & "-" & VNull(RstJob!LastInvNoSuff) & mEmph1
        Else
            Print #1, mSP2 & mChr18 & Space(36) & mEmph & PSTR(mDocStr & " No.", 22, , AlignRight) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
        End If
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
        mHeader = mHeader + 1
'***********************************
'        Print #1, mSP2 & PSTR(RstJob!NamePrefix & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
'        mHeader = mHeader + 1
'        Print #1, mSP2 & PSTR(XNull(RstJob!Add1), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
'        mHeader = mHeader + 1
'        Print #1, mSP2 & PSTR(XNull(RstJob!Add2), 40) & Space(6) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
'        mHeader = mHeader + 1
'        Print #1, mSP2 & PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
        & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
'       mHeader = mHeader + 1
'       Print #1, mSP2 & "Phone : " & PSTR(XNull(RstJob!Phone), 20)
'***********************************
        
        Print #1, mSP2 & PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40) & Space(6) & PSTR("Reg. No.", 8) & ": " & XNull(RstJob!RegNo) & "  Kms:" & RstJob!AtKMsHrs
        mHeader = mHeader + 1
        Print #1, mSP2 & PSTR(XNull(txt(Address3)) & IIf(XNull(txt(Address3)) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44) _
        & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
        mHeader = mHeader + 1
        Print #1, mSP2 & "Phone : " & PSTR(XNull(txt(PhoneOff)), 20)
        
        mHeader = mHeader + 1
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
        mHeader = mHeader + 1
        If mVatYn = 1 Then
            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR("DISC %", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 12, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("Tax Amt", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18    '& mDoub1
            mHeader = mHeader + 1
        Else
            If RstJob!Det_Tax = 1 Then
                Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                mHeader = mHeader + 1
                Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
                mHeader = mHeader + 1
            Else
                Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 27) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18 '& mDoub1
                mHeader = mHeader + 1


            End If
        End If
        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        If RstJob!orig = "1" Then
            Print #1, mSP2 & mEmph & "*Spare Details*" & mEmph1 & mChr17
            mHeader = mHeader + 1
            mSprCaption = True
        ElseIf RstJob!orig = "2" Then
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
        + RstJob!OilAmt_TB + RstJob!OilAmt_TP + Val(txt(IWDiscTotTP).TEXT) + Val(txt(IWDiscTotTB).TEXT)
        
        If UCase(left(PubSiteName, 8)) = "GWALTOLI" Then
            RstJob.MoveFirst
            Do While RstJob.EOF = False
               If RstJob!Purpose <> "F" Then mTotRow = mTotRow + 1
               RstJob.MoveNext
            Loop
        Else
            mTotRow = RstJob.RecordCount
        End If



        If mTotRow > 30 Then
            mFix = 30
        ElseIf mTotRow >= 15 And mTotRow <= 30 Then
            mFix = 30
        Else
            mFix = (PageLength - (mHeader + mFooter))
        End If
        
        'mTotRow = RstJob.RecordCount
        mTotRowTemp = RstJob.RecordCount
        If RstJob.RecordCount > 0 Then
            I = 1
            RstJob.MoveFirst
            Do Until RstJob.EOF = True
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
                        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                    Else
                        Print #1, PRN_TIT("** WORKSHOP SPARE " & mDocStr & mDupStr & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                    End If
                    Print #1, mSP2 & mChr18 & Space(40) & mEmph & PSTR(mDocStr & " No.", 20) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
                    mHeader = mHeader + 1
'*********************************
'                    Print #1, mSP2 & PSTR(XNull(RstJob!Add1), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
'                    mHeader = mHeader + 1
'                    Print #1, mSP2 & PSTR(XNull(RstJob!Add2), 40) & Space(6) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
'                    mHeader = mHeader + 1
'                    Print #1, mSP2 & PSTR(XNull(RstJob!Add3) & IIf(XNull(RstJob!Add3) <> "" And XNull(RstJob!CityName) <> "", ",", "") & XNull(RstJob!CityName), 44) _
'                    & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
'                    mHeader = mHeader + 1
'*********************************
                    Print #1, mSP2 & PSTR(XNull(txt(Address1)), 40) & Space(6) & mEmph & PSTR("Chassis No.", 12) & " : " & XNull(RstJob!Chassis) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(txt(Address2)), 40) & Space(6) & PSTR("Vehicle No.", 12) & " : " & XNull(RstJob!RegNo)
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(XNull(txt(Address3)) & IIf(XNull(txt(Address3)) <> "" And XNull(txt(City)) <> "", ",", "") & XNull(txt(City)), 44) _
                    & Space(2) & PSTR("Model", 12) & " : " & XNull(RstJob!Model)
                    
                    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17   ' & mDoub
                    mHeader = mHeader + 1
                    If mVatYn = 1 Then
                           Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & PSTR("DISC %", 8, , AlignRight) & PSTR("DISC.AMT", 10, , AlignRight) & PSTR("Tax %", 6, , AlignRight) & PSTR("TaxAmt", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18   '& mDoub1
                           mHeader = mHeader + 1
                    Else
                        If RstJob!Det_Tax = 1 Then
                            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 34) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                            mHeader = mHeader + 1
                            Print #1, mSP2 & Space(88) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mChr18    '& mDoub1
                            mHeader = mHeader + 1
                        Else
                            Print #1, mSP2 & PSTR("SRL.No", 7) & PSTR("PART NO.", 27) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mChr18 '& mDoub1
                            mHeader = mHeader + 1
                        End If
                    End If
                    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mChr17
                    mHeader = mHeader + 1
                    mFix = PageLength - (mHeader + mFooter)
                    mLine = 1
                    
                End If
                If mLabCaption = False Then
                    If RstJob!orig = "2" Then
                        Print #1, mSP2 & mChr18 & mEmph & "*Labour Details*" & mEmph1 & mChr17
                        mHeader = mHeader + 1
                        mLabCaption = True
                    End If
                End If
                
                mRate = IIf(RstJob!MRP_YN = 1, RstJob!MRP_Rate2, RstJob!Rate2)

                If RstJob!orig = "1" Then
                    If RstJob!Det_Tax = 1 Then
                        mTPAmtStr = PSTR(0, 12, 2)
                        mTBAmtStr = PSTR(0, 12, 2)
                    If RstJob!Purpose = "W" Then
                        mTBAmtStr = "*Warranty*"
                    ElseIf RstJob!Purpose = "P" Then
                        mTBAmtStr = "*PDI*"
                    ElseIf RstJob!Purpose = "A" And Not StrCmp(left(PubComp_Name, 4), "Enar") Then
                        mTBAmtStr = "*AMC*"
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
                    If mVatYn = 1 Then
                        If UCase(left(PubComp_Name, 7)) = "SHANKAR" Or UCase(left(PubComp_Name, 6)) = "MAURYA" Then
                            If RstJob!Purpose <> "W" And RstJob!Purpose <> "F" Then
                                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 30) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                                PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                                PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(VNull(RstJob!TaxPer), 6, 2) & PSTR(VNull(RstJob!TaxAmt), 10, 2) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                            ElseIf RstJob!Purpose = "W" Or RstJob!Purpose = "F" Then
                                PrintStr = "": mSlNo = mSlNo - 1
                            End If
                        Else
                            PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 30) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                            PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                            PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(VNull(RstJob!TaxPer), 6, 2) & PSTR(VNull(RstJob!TaxAmt), 10, 2) & PSTR(IIf(Val(mTBAmtStr) > 0 Or left(mTBAmtStr, 1) = "*", mTBAmtStr, mTPAmtStr), 12, 2, AlignRight)
                        End If
                    Else
                        If UCase(left(PubSiteName, 8)) = "GWALTOLI" Or UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
                            If RstJob!Purpose <> "F" Then
                                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 34) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                                PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                                PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                            ElseIf RstJob!Purpose = "W" Or RstJob!Purpose = "F" Then
                                PrintStr = ""
                                mSlNo = mSlNo - 1
                            End If
                        Else
                                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 34) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                                PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                                PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                        End If
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
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 40) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(LdRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", "L") & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & _
                    PSTR(LDAmt - RstJob!Disc_Amt2, 12, 2)
                End If
            Else    'Labour
            
                If UCase(left(PubComp_Name, 3)) = "LMP" Then mRate = 0
                If Val(txt(ServTaxAmt)) <= 0 Then
                    mTPAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                    mTBAmtStr = PSTR(0, 12, 2)
                Else
                    mTPAmtStr = PSTR(0, 12, 2)
                    mTBAmtStr = PSTR(RstJob!Net_Amt2, 12, 2)
                End If
                If mVatYn = 1 Then
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 30) & PSTR(RstJob!Qty_Iss - RstJob!Qty_Ret, 12, 3)
                    PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & Space(6) & Space(10) & PSTR(IIf(Val(mTBAmtStr) > 0, mTBAmtStr, mTPAmtStr), 12, 2, AlignRight)
                Else
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstJob!Part_No, 22, , AlignLeft) & PSTR(RstJob!Part_Name, 34) & PSTR(Format(RstJob!Qty_Iss - RstJob!Qty_Ret, "0.000"), 12, 3)
                    PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstJob!MRP_YN = 1, "M", IIf(RstJob!MRP_YN = 0, "L", "")) & _
                    PSTR(RstJob!Disc_Per2, 8, 2) & " %" & PSTR(RstJob!Disc_Amt2, 10, 2) & PSTR(mTPAmtStr, 12, 2, AlignRight) & PSTR(mTBAmtStr, 12, 2, AlignRight)
                End If
            End If
            If PrintStr <> "" Then
                Print #1, mSP2 & PrintStr
                mLine = mLine + 1
            End If
            mSlNo = mSlNo + 1
            
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
                        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                    Else
                        Print #1, PRN_TIT("** WORKSHOP SPARE " & mDocStr & mDupStr & " **", "B", PageWidth)
                        mHeader = mHeader + 1
                    End If
                    Print #1, mSP2 & mChr18 & Space(46) & mEmph & PSTR(mDocStr & " No.", 12) & " : " & PrinID(RstJob!DocId_InvSpr) & mEmph1
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR("To,", 46) & PSTR("DATE", 12, , AlignRight) & " : " & Format(RstJob!V_DATE, "dd/MMM/yyyy") & mEmph
                    mHeader = mHeader + 1
                    Print #1, mSP2 & PSTR(RstJob!NamePrefix & " " & RstJob!Party_Name, 44) & mEmph1 & Space(2) & PSTR("Job Card No.", 12) & " : " & PrinID(RstJob!JobDocID)
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


        Loop
    End If
    'If Page = 1 And mLine >= mTotRow Then mFix = 15
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    JobValue = 0
    If RSOJPR = True Then
        JobValue = JobValue + VNull(GCn.Execute("Select sum((Qty_Iss-Qty_Ret)*Rate) from sp_stock where Job_docID='" & txt(JobNo).Tag & "' and purpose <> 'C'").Fields(0).Value)
        JobValue = JobValue + VNull(GCn.Execute("Select sum(LabourAmt) from Job_Lab where Job_docID='" & txt(JobNo).Tag & "' and Chrg_Type <> 'C'").Fields(0).Value)
        JobValue = JobValue + Val(txt(NetAmt))
        Print #1, mChr18 & mSP2 & "Customer's Signature         Job Value : " & Format(JobValue, "0.00")
    Else
        Print #1, mChr18 & mSP2 & "Customer's Signature "
    End If
' SALE FOOTER
    '22 space maintain between heading and :
    RstJob.MoveFirst
    'If mTotRow <= 15 Then
    If RstJob!Det_Tax = 1 Then

        Print #1, mSP2 & Replace(Space(20), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
        If mVatYn = 1 Then
            Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("V A T     ", 10, 0) & Space(6) & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
            
            Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
            ; " | " & PSTR("S A T     ", 10, 0) & Space(6) & PSTR(RstJob!SatAmt_H, 12, 2) & mDoub
        Else
            If RSOJPR = True Then
                Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
                
                Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) _
                ; " | " & PSTR("MRP Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(mMRPTax, 12, 2) & mDoub

            Else
                Print #1, mSP2 & PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 11, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
                ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstJob!Tax_Per, 5, 2) & "%" & PSTR(RstJob!Tax_Amt, 12, 2) & mDoub
                
                Print #1, mSP2 & PSTR("MRP Items Amt", 16) & PSTR(RstJob!SprAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_MRP_TB, 12, 2) & mDoub1 _
                ; "" & mDoub
            End If
        End If
      
        Print #1, mSP2 & PSTR("Spares Amount", 16) & PSTR(RstJob!SprAmt_TP, 11, 2) & Space(8) & PSTR(RstJob!SprAmt_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Misc. Charges", 16) & PSTR(RstJob!Packing, 12, 2) & mDoub

        Print #1, mSP2 & PSTR("Oil Amount ", 16) & PSTR(RstJob!OilAmt_TP + RstJob!OilAmt_MRP_TP, 11, 2) & Space(8) & PSTR(RstJob!OilAmt_TB + RstJob!OilAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & mEmph & PSTR("Sub Total[TP+TB]", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
        
        Print #1, mSP2 & PSTR("Discount ", 10, 0) & PSTR(RstJob!D_Per_TP, 5, 2) & "%" & PSTR(RstJob!D_Amt_TP, 11, 2) & PSTR(RstJob!D_Per_TB, 7, 2) & "%" & PSTR(RstJob!D_Amt_TB, 12, 2) _
        ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstJob!TOT_Per, 5, 2) & "%" & PSTR(RstJob!Tot_Amt, 12, 2) & mEmph
        
        Print #1, mSP2 & PSTR("Sub Total [A]", 16) & PSTR(Val(txt(STotATP)), 11, 2) & Space(8) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
        ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstJob!ReSalTax_Per, 5, 2) & "%" & PSTR(RstJob!ReSalTax_Amt, 12, 2)
        
        Print #1, mSP2 & PSTR("Gen Surch ", 10, 0) & PSTR(RstJob!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 11, 2) & PSTR(RstJob!Gen_Sur_Amt, 20, 2) _
        ; " | " & PSTR("Round Off", 16) & PSTR(Round(RstJob!Rounded, 2), 12, 2)
       
        Print #1, mSP2 & PSTR("Transportation", 16) & PSTR(0, 11, 2) & PSTR(RstJob!Trans_Amt, 20, 2) _
        ; " | " & mEmph & PSTR("Net Spare + Lub. Rs.", 16) & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1
    Else

        Print #1, mSP2 & Replace(Space(PageWidth), " ", "-") & mDoub
        Print #1, mSP2 & Space(44) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
        If RstJob!D_Amt_TP + RstJob!D_Amt_TB > 0 Then
            Print #1, mSP2 & Space(44) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstJob!D_Amt_TP + RstJob!D_Amt_TB, 12, 2)
        Else
            Print #1, ""
        End If
        Print #1, mSP2 & Space(44) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Val(txt(NetSprAmt)) - (mGrossAmt - (RstJob!D_Amt_TP + RstJob!D_Amt_TB)), 12, 2) & mEmph
        Print #1, mSP2 & Space(44) & PSTR("Net Spare + Lub. Rs.", 20) & " : " & PSTR(Val(txt(NetSprAmt)), 12, 2) & mEmph1

    End If

    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    If Val(txt(LabDisc)) > 0 Then
        mLabDiscAmtStr = "Discount  : " & PSTR(Val(txt(LabDisc)), 8, 2)
    Else
        mLabDiscAmtStr = Space(19)
    End If
    PrintStr = "Total Lab.:" & PSTR(Val(txt(LabAmt)), 8, 2)
    PrintStr = PrintStr & " |Serv.Tax @ " & PSTR(Val(txt(ServTaxPer)), 5, 2) & ":" & PSTR(Val(txt(ServTaxAmt)), 9, 2) & " |" & mEmph & "Net Labour Rs.    : " & PSTR(Val(txt(NetLabAmt)), 9, 2) & mEmph1
    Print #1, mSP2 & PrintStr
    PrintStr = mLabDiscAmtStr
    PrintStr = PrintStr & " |" & "Round Off       :  " & PSTR(Val(txt(LabROff)), 7, 2) & " |" & mEmph & "Net Payble Amt Rs.: " & PSTR(Val(txt(NetAmt)), 9, 2) & mEmph1
    Print #1, mSP2 & PrintStr
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    Print #1, mSP2 & mDoub & ntow(txt(NetAmt), "Rupees", "Paise") & mDoub1
    Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
    Print #1, mChr17 & "The service tax  amount claimed on this invoice will be paid to govt. as per statutory provision" & mChr18
    If mVatYn = 1 Then
        Print #1, ""
    Else
        Print #1, mSP2 & mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
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
    If VNull(RstJob!Printed_YN) = 0 Then
        If RstJob!SrvGatePass = 1 And RstJob!SrvGatePass_On = "S" Then
            Print #1, mSP2 & Replace(Space(PageWidth), " ", "-")
            Print #1, mSP2 & PRN_TIT("* WORKSHOP SALE GATE PASS " & mDupStr & " *", "A", (PageWidth)) & mEmph
            Print #1, mSP2 & "GATE PASS No. & DATE : " & XNull(RstJob!gp_no) & "  " & XNull(RstJob!GP_Date) & mEmph1 & Space(10) & "Job Card No. : " & PrinID(RstJob!JobDocID)
            Print #1, mSP2 & "Vehicle No. : " & XNull(RstJob!RegNo) & Space(5) & "Chassis No. : " & XNull(RstJob!Chassis) _
            & Space(5) & mChr17 & "Model : " & XNull(RstJob!Model) & mChr18
            Print #1, mSP2 & PSTR("In DateTime  : " & RstJob!ArrivalTime, 40, , AlignLeft) & PSTR("Out DateTime : " & RstJob!JobComp_Dt_Time, 40, , AlignRight)
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
            
'    If fob.FolderExists("c:\WinNt") Then
        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'        Print #1, "Type C:\RepPrint.Txt > Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
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
    If Provisional Then
        MsgBox "Provisional Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !"
    Else
        If MsgBox("Spare Bill Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
            GCn.Execute "update Sp_Sale set Printed_YN = 1 where Sp_Sale.Job_DocID='" & txt(JobNo).Tag & "'"
        End If
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Function SubGroupUpdate(ByRef xType As String, ByRef xTableName As String, ByRef xAcID As String, ByRef xSubCode As String, ByRef xSubName As String, ByRef xSubNameBiLang As String, ByRef xAliasYN As String, xA_E As String, UnderGroup As String) As String
Dim Nature As String, GroupNature As String
Dim MyCurrBal As Double, xNature$
'    xNature = G_FACN.Execute("Select IsNull(Nature,'',Nature) N From AcGroup Where GroupCode='" & Txt(UnderGroup).Tag & "'").Fields(0).Value
'
'    '), "Other", G_FACN.Execute("Select Nature From AcGroup Where GroupCode='" & Txt(UnderGroup).Tag & "'").Fields(0).Value)

    Nature = IIf(IsNull(G_FaCn.Execute("Select Nature From AcGroup Where GroupCode='" & UnderGroup & "'").Fields(0).Value), "Other", G_FaCn.Execute("Select Nature From AcGroup Where GroupCode='" & UnderGroup & "'").Fields(0).Value)
    GroupNature = G_FaCn.Execute("Select GroupNature From AcGroup Where GroupCode='" & UnderGroup & "'").Fields(0).Value
    
     If xType = "Add" Then
        SubGroupUpdate = "Insert Into " & xTableName & "(" _
            & "AcID,Site_Code,SubCode,FirmCode,Name,AliasYN," _
            & "NameHelp,GroupCode,GroupNature,Nature,U_Name,U_AE,U_EntDt) " _
            & "Values ('" & xAcID & "','" & PubSiteCode & "','" & xSubCode & "','" & PubFirmCode & "','" & xSubName & "','" & xAliasYN & _
            "','" & FilterString(xSubName) & "','" & UnderGroup & "','" & GroupNature & "','" & Nature & "','" & pubUName & "','A'," & ConvertDate(PubLoginDate) & ")"
    Else
        SubGroupUpdate = ""
    End If
End Function


Sub Amt_Cal()
    Dim mAmount As Double, TaxAmt As Double, DisAmt As Double, OrdDisAmt1 As Double
    Dim TTaxAmt As Double, mTaxableAmt As Double
    Dim mQty As Double
    Dim I As Integer
    
    
    
    If StrCmp(left(PubComp_Name, 3), "Jmk") Or StrCmp(left(PubComp_Name, 7), "Singhal") Then
    
    
    
    
    

        With FGrid
            For I = 1 To .Rows - 1
                mQty = Val(.TextMatrix(I, Col_Qty))
            
                       
            
                If UCase(left(PubComp_Name, 3)) = "JMK" Then
                    .TextMatrix(I, Col_Amt) = Format((Val(.TextMatrix(I, Col_Rate)) * mQty), "0.00")
                Else
                    If UCase(.TextMatrix(I, Col_MRP)) = "YES" Then
                        .TextMatrix(I, Col_Amt) = Format((Val(.TextMatrix(I, Col_MRPRate)) * mQty), "0.00")
                    Else
                        .TextMatrix(I, Col_Amt) = Format((Val(.TextMatrix(I, Col_Rate)) * mQty), "0.00")
                    End If
                End If
                
                
                If StrCmp(FGrid.TextMatrix(I, Col_Taxable), "No") Then
                    If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                        FGrid.TextMatrix(I, Col_DiscPer) = "0"
                    Else
                        FGrid.TextMatrix(I, Col_DiscPer) = Format(Val(txt(IWDiscPerTP)))
                    End If
                Else
                    If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                        FGrid.TextMatrix(I, Col_DiscPer) = "0"
                    Else
                        If Val(txt(IWDiscPerTB)) > 0 Then
                            FGrid.TextMatrix(I, Col_DiscPer) = Format(Val(txt(IWDiscPerTB)))
                        End If
                    End If
                End If
                
                If .TextMatrix(I, Col_Purpose) = "Charge" Then
                    .TextMatrix(I, Col_DiscAmt) = Format(((Val(.TextMatrix(I, Col_Amt)) * Val(.TextMatrix(I, Col_DiscPer))) / 100), "0.00")
                    .TextMatrix(I, Col_ItemVal) = Format((Val(.TextMatrix(I, Col_Amt)) - Val(.TextMatrix(I, Col_DiscAmt))), "0.00")
                Else
                    .TextMatrix(I, Col_DiscAmt) = ""
                    .TextMatrix(I, Col_ItemVal) = ""
                End If
                 '******************** For Tax in Line File *************************
                If mVatYn = 1 Then
                    If .TextMatrix(I, Col_TaxPer) <> "" Then
                        mAmount = Val(.TextMatrix(I, Col_Amt))
                        DisAmt = Val(.TextMatrix(I, Col_DiscAmt))
                        If .TextMatrix(I, Col_MRP) = "Yes" And .TextMatrix(I, Col_Taxable) = "Yes" Then
                            If Val(.TextMatrix(I, Col_SatPer)) > 0 Then
                                mTaxableAmt = Format((mAmount - DisAmt) * 100 / (100 + Val(.TextMatrix(I, Col_TaxPer)) + Val(.TextMatrix(I, Col_SatPer))), "0.00")
                                .TextMatrix(I, Col_TaxAmt) = Format(mTaxableAmt * Val(.TextMatrix(I, Col_TaxPer)) / 100, "0.00")
                                .TextMatrix(I, Col_SatAmt) = Format(mTaxableAmt * Val(.TextMatrix(I, Col_SatPer)) / 100, "0.00")
                            Else
                                .TextMatrix(I, Col_TaxAmt) = Format((mAmount - DisAmt) * Val(.TextMatrix(I, Col_TaxPer)) / (100 + Val(.TextMatrix(I, Col_TaxPer))), "0.00")
                                .TextMatrix(I, Col_SatAmt) = 0
                            End If
                            If Val(.TextMatrix(I, Col_ItemVal)) > 0 Then
                                .TextMatrix(I, Col_ItemVal) = Format(Val(.TextMatrix(I, Col_ItemVal)) - Val(.TextMatrix(I, Col_TaxAmt)) - Val(.TextMatrix(I, Col_SatAmt)), "0.00")
                            End If
                        ElseIf .TextMatrix(I, Col_MRP) = "No" And .TextMatrix(I, Col_Taxable) = "Yes" Then
                            .TextMatrix(I, Col_TaxAmt) = Format((mAmount - DisAmt) * Val(.TextMatrix(I, Col_TaxPer)) / 100, "0.00")
                            .TextMatrix(I, Col_SatAmt) = Format((mAmount - DisAmt) * Val(.TextMatrix(I, Col_SatPer)) / 100, "0.00")
                        Else
                            .TextMatrix(I, Col_TaxAmt) = ""
                            .TextMatrix(I, Col_SatAmt) = ""
                        End If
                    End If
                End If
            Next I
        End With
        
        
        If Val(txt(MRPAmtTB)) + Val(txt(MRPAmtTP)) <> 0 Then
            MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                    Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                    Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
                    Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
        
        End If
        '*******************************************************************
        If mVatYn = 1 Then
            MainLib.SprCalcVAT WithLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, txt(SatAmt), Col_Purpose, True
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
        End If
            
        'Nra updation
        If Val(txt(LabAmtTB)) <> 0 Then
            txt(ServTaxPer) = MainLib.Serv_Tax
        Else
            txt(ServTaxPer) = Format(0, "0.00")
            txt(ServTaxAmt) = Format(0, "0.00")
        End If
        'Nra end updation
        
        MainLib.LabCalc txt(LabAmtTB), txt(LabAmtTP), txt(LabDisc), txt(ServTaxPer), txt(ServTaxAmt), txt(LabROff), txt(NetLabAmt), txt(OutSideLabAmt), mLabDiscAmtTB, mECessPer, mECessAmt, txt(FreeWarrLabAmt), mServiceTaxPer_Saperate, mServiceTaxAmt_Saperate, mHECessPer, mHECessAmt
        If UCase(left(PubComp_Name, 5)) = "SOCIE" Then
            txt(ServTaxPer) = MainLib.Serv_Tax
            txt(ServTaxAmt).TEXT = Format((Val(txt(LabAmtTB).TEXT) + Val(FreeLabForTax) - Val(txt(LabDisc).TEXT)) * Val(txt(ServTaxPer)) / 100, "0.00")
            txt(NetLabAmt).TEXT = Val(txt(LabAmtTB).TEXT) + Val(txt(LabAmtTP).TEXT) + Val(txt(ServTaxAmt).TEXT) - Val(txt(LabDisc).TEXT)
            txt(LabROff).TEXT = Format(dmRoundOff(txt(NetLabAmt)), "0.00")
            txt(NetLabAmt).TEXT = Format(Val(txt(NetLabAmt)) + Val(txt(LabROff)), "0.00")
        End If
        
        If UCase(left(PubComp_Name, 3)) = "JMK" Then
                txt(TurnOverAmt) = Format((Val(txt(STotATB)) + Val(txt(STaxAmt))) * Val(txt(TurnOverPer)) / 100, "0.00")
                txt(NetSprAmt).TEXT = Format(Val(txt(STotB)) + Val(txt(TurnOverAmt)), "0.00")
                'txt(SROff).TEXT = Format(Val(txt(NetSprAmt).TEXT) - Round(Val(txt(STotB).TEXT) + Val(txt(TurnOverAmt)), 0), "0.00")
                txt(NetSprAmt).TEXT = Format(Round(txt(NetSprAmt).TEXT, 0), "0.00")
                txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
        Else
            txt(NetAmt) = Format(Val(txt(NetSprAmt)) + Val(txt(NetLabAmt)), "0.00")
        End If
        
    End If
End Sub


