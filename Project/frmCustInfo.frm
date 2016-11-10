VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmCustInfo 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Customer Information"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11880
   Begin MSDataGridLib.DataGrid DGCont 
      Height          =   2910
      Left            =   7545
      Negotiate       =   -1  'True
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   5415
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
      BeginProperty Column01 
         DataField       =   "Cust_Name"
         Caption         =   "Customer Name"
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
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Invoce Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   2910
      TabIndex        =   123
      Top             =   2280
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   34
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1380
         Width           =   2865
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   645
         Index           =   33
         Left            =   2235
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   2865
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   32
         Left            =   2235
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2235
         Width           =   2865
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   31
         Left            =   2235
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1950
         Width           =   2865
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   30
         Left            =   2235
         MaxLength       =   14
         TabIndex        =   4
         Top             =   1665
         Width           =   2865
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   29
         Left            =   2235
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1095
         Width           =   2865
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   28
         Left            =   2235
         MaxLength       =   12
         TabIndex        =   1
         Top             =   810
         Width           =   2865
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   27
         Left            =   2235
         MaxLength       =   8
         TabIndex        =   0
         Top             =   525
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Varient"
         Height          =   195
         Index           =   7
         Left            =   1320
         TabIndex        =   132
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Club Eligibility"
         Height          =   195
         Index           =   6
         Left            =   870
         TabIndex        =   130
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. No."
         Height          =   195
         Index           =   5
         Left            =   1170
         TabIndex        =   129
         Top             =   1665
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Invoice Date"
         Height          =   195
         Index           =   4
         Left            =   390
         TabIndex        =   128
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chassis No."
         Height          =   195
         Index           =   3
         Left            =   975
         TabIndex        =   127
         Top             =   2235
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colour"
         Height          =   195
         Index           =   2
         Left            =   1365
         TabIndex        =   126
         Top             =   1950
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model No."
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   125
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Invoice No."
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   124
         Top             =   525
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdInvFill 
      Caption         =   "&Fill Inv. Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7455
      TabIndex        =   122
      Top             =   30
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
      Height          =   1605
      Left            =   5040
      TabIndex        =   108
      Top             =   4320
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmCustInfo.frx":0000
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
         TabIndex        =   113
         ToolTipText     =   "Printer "
         Top             =   285
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmCustInfo.frx":030A
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
         TabIndex        =   112
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmCustInfo.frx":0614
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
         TabIndex        =   111
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
         Picture         =   "frmCustInfo.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   110
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
         Picture         =   "frmCustInfo.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   109
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
         Index           =   18
         Left            =   0
         TabIndex        =   121
         Top             =   0
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
         TabIndex        =   120
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
         TabIndex        =   119
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
   Begin VB.CheckBox ChkDis 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "<100"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   4290
      TabIndex        =   47
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   5130
      MaxLength       =   20
      TabIndex        =   73
      Top             =   6630
      Width           =   1605
   End
   Begin VB.CheckBox ChkInt 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Other"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   4275
      TabIndex        =   72
      Top             =   6690
      Width           =   810
   End
   Begin VB.CheckBox ChkInt 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "CarRelated"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   1650
      TabIndex        =   71
      Top             =   6705
      Width           =   1155
   End
   Begin VB.CheckBox ChkInt 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Personal/Hobby Related"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   9285
      TabIndex        =   70
      Top             =   6225
      Width           =   2115
   End
   Begin VB.CheckBox ChkInt 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Entertainment"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   7065
      TabIndex        =   69
      Top             =   6255
      Width           =   1320
   End
   Begin VB.CheckBox ChkInt 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Lifestyle Related"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   4275
      TabIndex        =   68
      Top             =   6270
      Width           =   1605
   End
   Begin VB.CheckBox ChkInt 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Travel Related"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   1650
      TabIndex        =   67
      Top             =   6240
      Width           =   1380
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   66
      Top             =   5790
      Width           =   1620
   End
   Begin VB.CheckBox ChkFFM 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Other"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   6150
      TabIndex        =   65
      Top             =   5790
      Width           =   840
   End
   Begin VB.CheckBox ChkFFM 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Sahara"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   4290
      TabIndex        =   64
      Top             =   5790
      Width           =   840
   End
   Begin VB.CheckBox ChkFFM 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "JET"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   2985
      TabIndex        =   63
      Top             =   5790
      Width           =   810
   End
   Begin VB.CheckBox ChkFFM 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "IA"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   1665
      TabIndex        =   62
      Top             =   5790
      Width           =   735
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   7095
      MaxLength       =   20
      TabIndex        =   61
      Top             =   5250
      Width           =   1605
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Others"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   6150
      TabIndex        =   60
      Top             =   5250
      Width           =   810
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "BOB"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   8
      Left            =   4290
      TabIndex        =   59
      Top             =   5250
      Width           =   660
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Amex"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   2985
      TabIndex        =   58
      Top             =   5250
      Width           =   720
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Stanchart"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   6
      Left            =   1665
      TabIndex        =   57
      Top             =   5250
      Width           =   1005
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Diners"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   10215
      TabIndex        =   56
      Top             =   4800
      Width           =   750
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "SBI"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   8190
      TabIndex        =   55
      Top             =   4800
      Width           =   765
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "HSBC"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   6165
      TabIndex        =   54
      Top             =   4800
      Width           =   795
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "ICICI"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   4290
      TabIndex        =   53
      Top             =   4800
      Width           =   870
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "HDFC"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3000
      TabIndex        =   52
      Top             =   4800
      Width           =   870
   End
   Begin VB.CheckBox ChkCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Citibank"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   1665
      TabIndex        =   51
      Top             =   4800
      Width           =   930
   End
   Begin VB.CheckBox ChkSrv 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Other WorkShop"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   10200
      TabIndex        =   46
      Top             =   3840
      Width           =   1560
   End
   Begin VB.CheckBox ChkSrv 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Authorised WorkShop"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   7380
      TabIndex        =   45
      Top             =   3840
      Width           =   1920
   End
   Begin VB.CheckBox ChkCar 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Self/Driver"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   4290
      TabIndex        =   44
      Top             =   3840
      Width           =   1200
   End
   Begin VB.CheckBox ChkOcc 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "House Wife"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   10110
      TabIndex        =   37
      Top             =   2895
      Width           =   1335
   End
   Begin VB.CheckBox ChkOcc 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Student"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   8385
      TabIndex        =   36
      Top             =   2925
      Width           =   1110
   End
   Begin VB.CheckBox ChkOcc 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Retired"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   6045
      TabIndex        =   35
      Top             =   2895
      Width           =   1455
   End
   Begin VB.CheckBox ChkOcc 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Businessman"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3960
      TabIndex        =   34
      Top             =   2865
      Width           =   1515
   End
   Begin VB.CheckBox ChkOcc 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Salaried (govt/psu)"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   1680
      TabIndex        =   33
      Top             =   2865
      Width           =   2190
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   19
      Left            =   9315
      MaxLength       =   20
      TabIndex        =   32
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox ChkEdu 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Others"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   8385
      TabIndex        =   31
      Top             =   2535
      Width           =   960
   End
   Begin VB.CheckBox ChkEdu 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Post Graduate"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   6045
      TabIndex        =   30
      Top             =   2535
      Width           =   1635
   End
   Begin VB.CheckBox ChkEdu 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Graduate"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3960
      TabIndex        =   29
      Top             =   2535
      Width           =   1410
   End
   Begin VB.CheckBox ChkEdu 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Under Graduate"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   1710
      TabIndex        =   28
      Top             =   2535
      Width           =   1755
   End
   Begin VB.CheckBox ChkDis 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   ">1000"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   10200
      TabIndex        =   50
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox ChkDis 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "500-1000"
      ForeColor       =   &H80000008&
      Height          =   210
      HelpContextID   =   2
      Index           =   2
      Left            =   8190
      TabIndex        =   49
      Top             =   4320
      Width           =   1020
   End
   Begin VB.CheckBox ChkDis 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "100-500"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   6165
      TabIndex        =   48
      Top             =   4320
      Width           =   960
   End
   Begin VB.CheckBox ChkCar 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Driver"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3000
      TabIndex        =   43
      Top             =   3840
      Width           =   945
   End
   Begin VB.CheckBox ChkCar 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Self"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   1695
      TabIndex        =   42
      Top             =   3840
      Width           =   1125
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   20
      Left            =   2415
      MaxLength       =   20
      TabIndex        =   38
      Top             =   3300
      Width           =   2115
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   21
      Left            =   5640
      MaxLength       =   12
      TabIndex        =   39
      Top             =   3300
      Width           =   1020
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   23
      Left            =   10695
      MaxLength       =   12
      TabIndex        =   41
      Top             =   3300
      Width           =   1020
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   22
      Left            =   7440
      MaxLength       =   20
      TabIndex        =   40
      Top             =   3300
      Width           =   2055
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   9
      Top             =   630
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   10
      Top             =   930
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1230
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1530
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1830
      Width           =   885
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2790
      MaxLength       =   6
      TabIndex        =   14
      Top             =   1830
      Width           =   1185
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   15
      Top             =   2130
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   5115
      MaxLength       =   30
      TabIndex        =   16
      Top             =   630
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   5115
      MaxLength       =   30
      TabIndex        =   17
      Top             =   930
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   5115
      MaxLength       =   30
      TabIndex        =   18
      Top             =   1230
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   5115
      MaxLength       =   15
      TabIndex        =   19
      Top             =   1530
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   5115
      MaxLength       =   30
      TabIndex        =   20
      Top             =   1830
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   6465
      MaxLength       =   3
      TabIndex        =   21
      Top             =   2130
      Width           =   1080
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   9300
      MaxLength       =   6
      TabIndex        =   22
      Top             =   630
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   9300
      MaxLength       =   6
      TabIndex        =   23
      Top             =   930
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   9300
      MaxLength       =   2
      TabIndex        =   24
      Top             =   1230
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   9300
      MaxLength       =   20
      TabIndex        =   25
      Top             =   1530
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   9300
      MaxLength       =   20
      TabIndex        =   26
      Top             =   1830
      Width           =   2430
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   9300
      MaxLength       =   20
      TabIndex        =   27
      Top             =   2130
      Width           =   2430
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1515
      X2              =   11715
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Salaried (govt/psu)"
      Height          =   195
      Index           =   47
      Left            =   1935
      TabIndex        =   107
      Top             =   2865
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Interest Areas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   64
      Left            =   405
      TabIndex        =   106
      Top             =   6225
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Personal Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   63
      Left            =   330
      TabIndex        =   105
      Top             =   375
      Width           =   1305
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   1515
      X2              =   11715
      Y1              =   7020
      Y2              =   7020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Freq. Flyer Mem."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   58
      Left            =   210
      TabIndex        =   104
      Top             =   5790
      Width           =   1425
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1515
      X2              =   11715
      Y1              =   5625
      Y2              =   5625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Credit Cards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   53
      Left            =   585
      TabIndex        =   103
      Tag             =   "Cre"
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1515
      X2              =   11715
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1530
      X2              =   11730
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1530
      X2              =   11730
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Occupation Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   46
      Left            =   90
      TabIndex        =   102
      Top             =   2850
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Car Serviced at"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   41
      Left            =   5700
      TabIndex        =   101
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Average Distance Traveled in KM/Week"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   34
      Left            =   435
      TabIndex        =   100
      Top             =   4290
      Width           =   3465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Car Driven By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   450
      TabIndex        =   99
      Top             =   3855
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Other Car Owned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   165
      TabIndex        =   98
      Top             =   3300
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Model (1)"
      Height          =   195
      Index           =   21
      Left            =   1680
      TabIndex        =   97
      Top             =   3300
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Date of Purch."
      Height          =   195
      Index           =   22
      Left            =   4590
      TabIndex        =   96
      Top             =   3300
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Model (2)"
      Height          =   195
      Index           =   23
      Left            =   6705
      TabIndex        =   95
      Top             =   3300
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Educational Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   60
      TabIndex        =   94
      Top             =   2505
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Date of Purch."
      Height          =   195
      Index           =   24
      Left            =   9585
      TabIndex        =   93
      Top             =   3300
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Customer Name"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   92
      Top             =   630
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Contact Address"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   91
      Top             =   930
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "State"
      Height          =   195
      Index           =   3
      Left            =   975
      TabIndex        =   90
      Top             =   2190
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "City"
      Height          =   195
      Index           =   2
      Left            =   1095
      TabIndex        =   89
      Top             =   1860
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Pin"
      Height          =   195
      Index           =   4
      Left            =   2460
      TabIndex        =   88
      Top             =   1830
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Phone (O)"
      Height          =   195
      Index           =   5
      Left            =   4170
      TabIndex        =   87
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Phone (R)"
      Height          =   195
      Index           =   6
      Left            =   4170
      TabIndex        =   86
      Top             =   930
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Fax"
      Height          =   195
      Index           =   7
      Left            =   4635
      TabIndex        =   85
      Top             =   1290
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Mobile"
      Height          =   195
      Index           =   8
      Left            =   4425
      TabIndex        =   84
      Top             =   1530
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "E-Mail"
      Height          =   195
      Index           =   9
      Left            =   4455
      TabIndex        =   83
      Top             =   1830
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Personal Computer  (Y/N)"
      Height          =   195
      Index           =   10
      Left            =   4305
      TabIndex        =   82
      Top             =   2130
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Sex (M/F)"
      Height          =   195
      Index           =   11
      Left            =   8400
      TabIndex        =   81
      Top             =   630
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Marital Status (S/M)"
      Height          =   195
      Index           =   12
      Left            =   7695
      TabIndex        =   80
      Top             =   930
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "No of Children"
      Height          =   195
      Index           =   13
      Left            =   8100
      TabIndex        =   79
      Top             =   1230
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Birth Date"
      Height          =   195
      Index           =   14
      Left            =   8400
      TabIndex        =   78
      Top             =   1530
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Wedding Date"
      Height          =   195
      Index           =   15
      Left            =   8070
      TabIndex        =   77
      Top             =   1830
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Mother Tounge"
      Height          =   195
      Index           =   16
      Left            =   8010
      TabIndex        =   76
      Top             =   2130
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Month and year of Purchase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   26
      Left            =   4305
      TabIndex        =   75
      Top             =   30
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Model (1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   25
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "frmCustInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ADDFLAG$
Private Const CUST_NAME = 0
Private Const Add1 = 1
Private Const Add2 = 2
Private Const Add3 = 3
Private Const City = 4
Private Const Pin = 5
Private Const State = 6
Private Const PhoneO = 7
Private Const PhoneR = 8
Private Const FAx = 9
Private Const Mobile = 10
Private Const EMail = 11
Private Const PC_YN = 12
Private Const Sex = 13
Private Const MState = 14
Private Const NoofChild = 15
Private Const BDate = 16
Private Const WDate = 17
Private Const MTounge = 18
Private Const EduOther = 19
Private Const ModPur1 = 20
Private Const ModPur1Date = 21
Private Const ModPur2 = 22
Private Const ModPur2Date = 23
Private Const CardOther = 24
Private Const FFMOther = 25
Private Const InterestOther = 26
Dim ListArray As Variant
Dim mListItem As ListItem
Dim Master As ADODB.Recordset
Dim rsCont As ADODB.Recordset

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4

'For Invoice Detail
Private Const DInvNo = 27
Private Const DInvDate = 28
Private Const Model = 29
Private Const RegNo = 30
Private Const Color = 31
Private Const Chassis = 32
Private Const Eligibility = 33
Private Const VarientMod = 34

Dim mRepName As String
Private Sub Ini_Grid()
    DGCont.width = 5000: DGCont.left = Me.width - (DGCont.width + mRtScale): DGCont.top = mTopScale: DGCont.height = 5000
End Sub
Private Sub CmdInvFill_Click()
If TopCtrl1.TopText2 <> "Add" Then
    FrmInv.left = (Me.width / 2) - (FrmInv.width / 2)
    FrmInv.top = (Me.height / 2) - (FrmInv.height / 2)
    FrmInv.Visible = True
    If Txt(DInvNo).Enabled = True Then Txt(DInvNo).SetFocus
End If
End Sub

Private Sub Form_Load()
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    ForSiteCode = PubSiteCode
    Call BlankText
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Cust_Code as SearchCode from CustInfo Order By Cust_Code Desc ", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsCont = New ADODB.Recordset
    rsCont.CursorLocation = adUseClient
    rsCont.Open "Select right(VS.Sal_DocID,8) as Inv_No,SG.Name as Cust_Name,VS.Sal_VDate,VS.MODEL,VS.ChassisNo,C.Col_Desc from ((Veh_Stock VS Left Join ColMast C on VS.Colour_Code=C.Col_Code) Left Join Veh_Order VO on VS.Sal_DocId=VO.Inv_DocID) Left Join SubGroup SG On VO.PartyCode=SG.SubCode where Len(VS.Sal_DocID)  > 1", GCn, adOpenDynamic, adLockOptimistic
    Set DGCont.DataSource = rsCont
    rsCont.Sort = "Inv_No"
    rsCont.Sort = "Cust_Name"
    
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End Sub
Private Sub BlankText()
    Dim i As Byte
    For i = 0 To Txt.Count - 1
        Txt(i).TEXT = ""
    Next i
    For i = 0 To 3
        ChkEdu(i).Value = 0
    Next
    For i = 0 To 4
        ChkOcc(i).Value = 0
    Next
    For i = 0 To 2
        ChkCar(i).Value = 0
    Next
    For i = 0 To 1
        ChkSrv(i).Value = 0
    Next
    For i = 0 To 3
        ChkDis(i).Value = 0
    Next
    For i = 0 To 9
        ChkCard(i).Value = 0
    Next
    For i = 0 To 3
        ChkFFM(i).Value = 0
    Next
    For i = 0 To 5
        ChkInt(i).Value = 0
    Next
End Sub

Private Sub TopCtrl1_eCancel()
Dim i As Integer
On Error GoTo ErrorLoop
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        Call Ini_Grid
        Call MoveRec
    Else
        Me.ActiveControl.SetFocus
    End If
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eExit()
Unload Me
End Sub

Private Sub TopCtrl1_ePrn()
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
CmdPrint(PWindows).SetFocus
End Sub
Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
GSQL = "SELECT * from CustInfo WHERE Cust_Code=" & Txt(CUST_NAME).Tag & ""
Select Case Index
    Case PScreen, PWindows
        mRepName = "CustInfo"
        Call WindowsPrint(GSQL, Index)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = "CustInfo"
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
Private Sub WindowsPrint(mQRY As String, Index As Integer)
On Error GoTo ERRORHANDLER
Dim Rst As ADODB.Recordset
Dim Rst1 As ADODB.Recordset
Dim mReportCount As Integer, i As Integer
 
Set Rst = GCn.Execute(mQRY)

CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
For i = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("TITLE")
            rpt.FormulaFields(i).TEXT = "'CUSTOMER INFORMATION'"
    End Select
Next
     
rpt.Database.SetDataSource Rst
rpt.ReadRecords
Set Rst = Nothing
Select Case Index
    Case PWindows  'Printer
        For i = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
                Case UCase("comp_name")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_Name & "'"
                Case UCase("comp_add1")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_Add & "'"
                Case UCase("comp_add2")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_Add2 & "'"
                Case UCase("comp_city")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_City & "'"
            End Select
        Next
        rpt.PrintOut False
    Case PScreen  'screen
        Call Report_View(rpt, "CUSTOMER INFORMATION", , True)
End Select

CmdPrint(PSetUp).Tag = ""
Set rpt = Nothing
Set Rst1 = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eRef()
Call UpdRequery
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Grid_Hide
    Ctrl_GetFocus Txt(Index)
End Sub
Private Sub UpdRequery()
    Master.Requery
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
 End If
 Select Case Index
        Case DInvNo
            DGridTxtKeyDown DGCont, Txt, Index, rsCont, KeyCode, False, 0
            DGCont_Click
 End Select
 If DGCont.Visible = False Then
        '' KEY DOWN
        If FrmInv.Visible = True Then
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index = Eligibility) Or (ADDFLAG = "E" And Index = Eligibility)) Then
                If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
        Else
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index = InterestOther) Or (ADDFLAG = "E" And Index = InterestOther)) Then
                If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
        End If
        
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index <> InterestOther) Or (ADDFLAG = "E" And Index <> InterestOther)) Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        ' KEY UP
        If ADDFLAG = "A" Then
             If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf ADDFLAG = "E" Then
             If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
 End If
End Sub
Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
    Call CheckQuote(keyascii)
    Select Case Index
        Case Pin
            NumPress Txt(Index), keyascii, 6, 0
        Case NoofChild
            NumPress Txt(Index), keyascii, 2, 0
        Case PC_YN
            If Asc("Y") = keyascii Or Asc("y") = keyascii Then
                Txt(Index) = "Yes"
                keyascii = 0
            ElseIf Asc("N") = keyascii Or Asc("n") = keyascii Then
                Txt(Index) = "No"
                keyascii = 0
            Else
                keyascii = 0
            End If
        Case Sex
            If Asc("M") = keyascii Or Asc("m") = keyascii Then
                Txt(Index) = "Male"
                keyascii = 0
            ElseIf Asc("F") = keyascii Or Asc("f") = keyascii Then
                Txt(Index) = "Female"
                keyascii = 0
            Else
                keyascii = 0
            End If
        Case MState
            If Asc("S") = keyascii Or Asc("s") = keyascii Then
                Txt(Index) = "Single"
                keyascii = 0
            ElseIf Asc("M") = keyascii Or Asc("m") = keyascii Then
                Txt(Index) = "Married"
                keyascii = 0
            Else
                keyascii = 0
            End If
        Case DInvNo
            DGridTxtKeyPress Txt, Index, rsCont, keyascii, "Inv_No"
    End Select
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Txt(CUST_NAME).SetFocus
    Ini_Grid
    Exit Sub
ErrorLoop:
    CheckError
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
                    
        GCn.Execute "Delete from CustInfo  where Cust_Code=" & Txt(CUST_NAME).Tag & ""
    
        GCn.CommitTrans
        
        Master.Requery
        
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
Dim i As Integer
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    If FrmInv.Visible = True Then
        Txt(DInvNo).SetFocus
    Else
        Txt(CUST_NAME).SetFocus
    End If
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
If KeyCode = vbKeyEscape Then
    FrmInv.Visible = False
End If
Exit Sub
ELoop:
    CheckError
End Sub
Public Sub Disp_Text(Enb As Boolean)
    Dim i As Integer
    'New Testing for Speed purpose
    ADDFLAG = left(TopCtrl1.TopText2, 1)
    'eof New Testing
    For i = 0 To Txt.Count - 1
        Txt(i).Enabled = Enb
    Next
    For i = 0 To Txt.Count - 1
        Txt(i).BackColor = CtrlBColOrg
        Txt(i).ForeColor = CtrlFColOrg
    Next
    For i = 0 To 3
        ChkEdu(i).Enabled = Enb
    Next
    For i = 0 To 4
        ChkOcc(i).Enabled = Enb
    Next
    For i = 0 To 2
        ChkCar(i).Enabled = Enb
    Next
    For i = 0 To 1
        ChkSrv(i).Enabled = Enb
    Next
    For i = 0 To 3
        ChkDis(i).Enabled = Enb
    Next
    For i = 0 To 9
        ChkCard(i).Enabled = Enb
    Next
    For i = 0 To 3
        ChkFFM(i).Enabled = Enb
    Next
    For i = 0 To 5
        ChkInt(i).Enabled = Enb
    Next
End Sub
Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
  Select Case Index
    Case BDate
        Txt(Index) = RetDate(Txt(Index))
    Case WDate
        Txt(Index) = RetDate(Txt(Index))
    Case ModPur1Date, ModPur2Date
        If Txt(Index) <> "" Then
            Txt(Index) = RetDate(Txt(Index))
        End If
  End Select
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select Cust_Code as SearchCode,Cust_Name,City,State,bdate " & _
        "from CustInfo order by Cust_Name"
        Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub Grid_Hide()
    If DGCont.Visible = True Then DGCont.Visible = False
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode=" & MyValue & "")
    BUTTONS True, Me, Master, 0
    Call MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_eSave()
    Dim mTrans As Boolean, i As Integer, CustCode As Double
    On Error GoTo errlbl
    Grid_Hide
    If IsValid(Txt(CUST_NAME), "Customer Name") = False Then Exit Sub
    If IsValid(Txt(Add1), "Address") = False Then Exit Sub
    If IsValid(Txt(City), "City") = False Then Exit Sub
    If IsValid(Txt(State), "State") = False Then Exit Sub
    If IsValid(Txt(BDate), "Bdate") = False Then Exit Sub
    
    GCn.BeginTrans
    mTrans = True
    If ADDFLAG = "A" Then
        '' Get gate pass serial no
        CustCode = Format(VNull(GCn.Execute("Select max(Cust_Code) from CustInfo ").Fields(0)) + 1, "000000")
        GSQL = "insert into CustInfo(" _
            & "Cust_Code,Cust_Name,Add1,Add2,Add3,City,Pin,State,PhoneO,PhoneR,Mobile,Fax,EMail,PCY_N," _
            & "Sex,MState,NoOfChild,BDate,WDate,MTounge,Ugrd,Grd,Pgrd,OthEdu,Salaried,Businessman,Retired, " _
            & "Student,Housewife,OthMod1,OthMod2,ModPur1Dt,ModPur2Dt,Self,Driver, " _
            & "Both1,AuthWorkshop,OtherWorkshop,AvgDist1,AvgDist2,AvgDist3,AvgDist4,Citibank,HDFC,ICICI,HSBC,SBI,Diners,Stanchart,Amex, " _
            & "BOB,OtherBank,IA,Jet,Sahara,OtherAir,TravelRelated,LifeRelated,Entertainment,Parsonal,CarRelated, " _
            & "OtherInterest,U_Name,U_EntDt,U_AE)" _
            & " values(" _
            & "" & CustCode & ",'" & Txt(CUST_NAME) & "','" & Txt(Add1) & "','" & Txt(Add2) & "','" & Txt(Add3) & "'," _
            & "'" & Txt(City) & "'," & Txt(Pin) & ",'" & Txt(State) & "','" & Txt(PhoneO) & "','" & Txt(PhoneR) & "'," _
            & "'" & Txt(Mobile) & "','" & Txt(FAx) & "','" & Txt(EMail) & "'," & IIf(Txt(PC_YN) = "Yes", 1, 0) & ",'" & Txt(Sex) & "'," _
            & "'" & Txt(MState) & "'," & IIf(Txt(NoofChild) <> "", Txt(NoofChild), 0) & "," & ConvertDate(Txt(BDate)) & "," & ConvertDate(Txt(WDate)) & ",'" & Txt(MTounge) & "'," _
            & "" & IIf(ChkEdu(0).Value = 1, 1, 0) & "," & IIf(ChkEdu(1).Value = 1, 1, 0) & "," & IIf(ChkEdu(2).Value = 1, 1, 0) & "" _
            & ",'" & IIf(ChkEdu(3).Value = 1, Txt(EduOther), "") & "'," & IIf(ChkOcc(0).Value = 1, 1, 0) & "," & IIf(ChkOcc(1).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkOcc(2).Value = 1, 1, 0) & "," & IIf(ChkOcc(3).Value = 1, 1, 0) & "," & IIf(ChkOcc(4).Value = 1, 1, 0) & "" _
            & ",'" & Txt(ModPur1) & "','" & Txt(ModPur2) & "'," & ConvertDate(Txt(ModPur1Date)) & "," & ConvertDate(Txt(ModPur2Date)) & "," & IIf(ChkCar(0).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCar(1).Value = 1, 1, 0) & "," & IIf(ChkCar(2).Value = 1, 1, 0) & "," & IIf(ChkSrv(0).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkSrv(1).Value = 1, 1, 0) & "," & IIf(ChkDis(0).Value = 1, 1, 0) & "," & IIf(ChkDis(1).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkDis(2).Value = 1, 1, 0) & "," & IIf(ChkDis(3).Value = 1, 1, 0) & "," & IIf(ChkCard(0).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCard(1).Value = 1, 1, 0) & "," & IIf(ChkCard(2).Value = 1, 1, 0) & "," & IIf(ChkCard(3).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCard(4).Value = 1, 1, 0) & "," & IIf(ChkCard(5).Value = 1, 1, 0) & "," & IIf(ChkCard(6).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCard(7).Value = 1, 1, 0) & "," & IIf(ChkCard(8).Value = 1, 1, 0) & ",'" & IIf(ChkCard(9).Value = 1, Txt(CardOther), "") & "'" _
            & "," & IIf(ChkFFM(0).Value = 1, 1, 0) & "," & IIf(ChkFFM(1).Value = 1, 1, 0) & "," & IIf(ChkFFM(2).Value = 1, 1, 0) & ",'" & IIf(ChkFFM(3).Value = 1, Txt(FFMOther), "") & "'" _
            & "," & IIf(ChkInt(0).Value = 1, 1, 0) & "," & IIf(ChkInt(1).Value = 1, 1, 0) & "," & IIf(ChkInt(2).Value = 1, 1, 0) & "," & IIf(ChkInt(3).Value = 1, 1, 0) & "," & IIf(ChkInt(4).Value = 1, 1, 0) & ",'" & IIf(ChkInt(5).Value = 1, Txt(InterestOther), "") & "','" & pubUName & "',#" & PubServerDate & "#,'" & ADDFLAG & "')"
    ElseIf ADDFLAG = "E" Then
        GCn.Execute ("Delete from CustInfo where Cust_Code=" & Txt(CUST_NAME).Tag & "")
        GSQL = "insert into CustInfo(" _
            & "Cust_Code,Cust_Name,Add1,Add2,Add3,City,Pin,State,PhoneO,PhoneR,Mobile,Fax,EMail,PCY_N," _
            & "Sex,MState,NoOfChild,BDate,WDate,MTounge,Ugrd,Grd,Pgrd,OthEdu,Salaried,Businessman,Retired, " _
            & "Student,Housewife,OthMod1,OthMod2,ModPur1Dt,ModPur2Dt,Self,Driver, " _
            & "Both1,AuthWorkshop,OtherWorkshop,AvgDist1,AvgDist2,AvgDist3,AvgDist4,Citibank,HDFC,ICICI,HSBC,SBI,Diners,Stanchart,Amex, " _
            & "BOB,OtherBank,IA,Jet,Sahara,OtherAir,TravelRelated,LifeRelated,Entertainment,Parsonal,CarRelated, " _
            & "OtherInterest,DInv_No,DInv_Date,Model,Varient,RegNo,Color,Chassis,CEligibility,U_Name,U_EntDt,U_AE)" _
            & " values(" _
            & "" & Txt(CUST_NAME).Tag & ",'" & Txt(CUST_NAME) & "','" & Txt(Add1) & "','" & Txt(Add2) & "','" & Txt(Add3) & "'," _
            & "'" & Txt(City) & "'," & Txt(Pin) & ",'" & Txt(State) & "','" & Txt(PhoneO) & "','" & Txt(PhoneR) & "'," _
            & "'" & Txt(Mobile) & "','" & Txt(FAx) & "','" & Txt(EMail) & "'," & IIf(Txt(PC_YN) = "Yes", 1, 0) & ",'" & Txt(Sex) & "'," _
            & "'" & Txt(MState) & "'," & IIf(Txt(NoofChild) <> "", Txt(NoofChild), 0) & "," & ConvertDate(Txt(BDate)) & "," & ConvertDate(Txt(WDate)) & ",'" & Txt(MTounge) & "'," _
            & "" & IIf(ChkEdu(0).Value = 1, 1, 0) & "," & IIf(ChkEdu(1).Value = 1, 1, 0) & "," & IIf(ChkEdu(2).Value = 1, 1, 0) & "" _
            & ",'" & IIf(ChkEdu(3).Value = 1, Txt(EduOther), "") & "'," & IIf(ChkOcc(0).Value = 1, 1, 0) & "," & IIf(ChkOcc(1).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkOcc(2).Value = 1, 1, 0) & "," & IIf(ChkOcc(3).Value = 1, 1, 0) & "," & IIf(ChkOcc(4).Value = 1, 1, 0) & "" _
            & ",'" & Txt(ModPur1) & "','" & Txt(ModPur2) & "'," & ConvertDate(Txt(ModPur1Date)) & "," & ConvertDate(Txt(ModPur2Date)) & "," & IIf(ChkCar(0).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCar(1).Value = 1, 1, 0) & "," & IIf(ChkCar(2).Value = 1, 1, 0) & "," & IIf(ChkSrv(0).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkSrv(1).Value = 1, 1, 0) & "," & IIf(ChkDis(0).Value = 1, 1, 0) & "," & IIf(ChkDis(1).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkDis(2).Value = 1, 1, 0) & "," & IIf(ChkDis(3).Value = 1, 1, 0) & "," & IIf(ChkCard(0).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCard(1).Value = 1, 1, 0) & "," & IIf(ChkCard(2).Value = 1, 1, 0) & "," & IIf(ChkCard(3).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCard(4).Value = 1, 1, 0) & "," & IIf(ChkCard(5).Value = 1, 1, 0) & "," & IIf(ChkCard(6).Value = 1, 1, 0) & "" _
            & "," & IIf(ChkCard(7).Value = 1, 1, 0) & "," & IIf(ChkCard(8).Value = 1, 1, 0) & ",'" & IIf(ChkCard(9).Value = 1, Txt(CardOther), "") & "'" _
            & "," & IIf(ChkFFM(0).Value = 1, 1, 0) & "," & IIf(ChkFFM(1).Value = 1, 1, 0) & "," & IIf(ChkFFM(2).Value = 1, 1, 0) & ",'" & IIf(ChkFFM(3).Value = 1, Txt(FFMOther), "") & "'" _
            & "," & IIf(ChkInt(0).Value = 1, 1, 0) & "," & IIf(ChkInt(1).Value = 1, 1, 0) & "," & IIf(ChkInt(2).Value = 1, 1, 0) & "," & IIf(ChkInt(3).Value = 1, 1, 0) & "," & IIf(ChkInt(4).Value = 1, 1, 0) & ",'" & IIf(ChkInt(5).Value = 1, Txt(InterestOther), "") & "'" _
            & "," & Val(Txt(DInvNo)) & "," & ConvertDate(Txt(DInvDate)) & ",'" & Txt(Model) & "','" & Txt(VarientMod) & "','" & Txt(RegNo) & "','" & Txt(Color) & "','" & Txt(Chassis) & "','" & Txt(Eligibility) & "','" & pubUName & "',#" & PubServerDate & "#,'" & ADDFLAG & "')"
    End If
    GCn.Execute GSQL
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    'If ADDFLAG = "A" Then 'TopCtrl1_ePrn
    Disp_Text SETS("INI", Me, Master)
    Ini_Grid
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
Exit Sub
End Sub
Private Sub MoveRec()
Dim Master1 As Recordset, rs1 As Recordset
Dim mVor As String
Dim i As Integer
On Error GoTo error1
    If Master.RecordCount > 0 Then
    '   Master.Open "select GP.GatePassNo as SearchCode,GP.*, Emp_Mast.Emp_Name as MechName from Job_GatePass as GP Left Join Emp_Mast on GP.Mech_Code=Emp_Mast.Emp_Code  where left(Job_DocId,1)='" & PubDivCode & "' order by gp.GatePassNo", GCn, adOpenDynamic, adLockOptimistic
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select * from CustInfo Where Cust_Code=" & Master!SearchCode & "", GCn, adOpenStatic, adLockReadOnly
            
        
        Txt(CUST_NAME).TEXT = XNull(Master1!CUST_NAME)
        Txt(CUST_NAME).Tag = XNull(Master1!Cust_code)
        Txt(Add1).TEXT = XNull(Master1!Add1)
        Txt(Add2).TEXT = XNull(Master1!Add2)
        Txt(Add3).TEXT = XNull(Master1!Add3)
        Txt(City).TEXT = XNull(Master1!City)
        Txt(Pin).TEXT = XNull(Master1!Pin)
        Txt(State).TEXT = XNull(Master1!State)
        Txt(PhoneO).TEXT = XNull(Master1!PhoneO)
        Txt(PhoneR).TEXT = XNull(Master1!PhoneR)
        Txt(FAx).TEXT = XNull(Master1!FAx)
        Txt(Mobile).TEXT = XNull(Master1!Mobile)
        Txt(EMail).TEXT = XNull(Master1!EMail)
        Txt(PC_YN).TEXT = IIf(Master1!PCY_N = 1, "Yes", "No")
        Txt(Sex).TEXT = XNull(Master1!Sex)
        Txt(MState).TEXT = XNull(Master1!MState)
        Txt(NoofChild).TEXT = XNull(Master1!NoofChild)
        Txt(BDate).TEXT = XNull(Master1!BDate)
        Txt(WDate).TEXT = XNull(Master1!WDate)
        Txt(MTounge).TEXT = XNull(Master1!MTounge)
        
        If Master1!Ugrd = 1 Then ChkEdu(0).Value = 1 Else ChkEdu(0).Value = 0
        If Master1!Grd = 1 Then ChkEdu(1).Value = 1 Else ChkEdu(1).Value = 0
        If Master1!Pgrd = 1 Then ChkEdu(2).Value = 1 Else ChkEdu(2).Value = 0
        If Master1!OthEdu <> "" Then ChkEdu(3).Value = 1 Else ChkEdu(3).Value = 0: Txt(EduOther) = ""
        If Master1!OthEdu <> "" Then Txt(EduOther) = Master1!OthEdu Else ChkEdu(3).Value = 0: Txt(EduOther) = ""
        Txt(EduOther).Enabled = False
        
        If Master1!Salaried = 1 Then ChkOcc(0).Value = 1 Else ChkOcc(0).Value = 0
        If Master1!Businessman = 1 Then ChkOcc(1).Value = 1 Else ChkOcc(1).Value = 0
        If Master1!Retired = 1 Then ChkOcc(2).Value = 1 Else ChkOcc(2).Value = 0
        If Master1!Student = 1 Then ChkOcc(3).Value = 1 Else ChkOcc(3).Value = 0
        If Master1!Housewife = 1 Then ChkOcc(4).Value = 1 Else ChkOcc(4).Value = 0
        
        Txt(ModPur1).TEXT = XNull(Master1!OthMod1)
        Txt(ModPur2).TEXT = XNull(Master1!OthMod2)
        Txt(ModPur1Date).TEXT = XNull(Master1!ModPur1Dt)
        Txt(ModPur2Date).TEXT = XNull(Master1!ModPur2Dt)
        
        If Master1!Self = 1 Then ChkCar(0).Value = 1 Else ChkCar(0).Value = 0
        If Master1!Driver = 1 Then ChkCar(1).Value = 1 Else ChkCar(1).Value = 0
        If Master1!Both1 = 1 Then ChkCar(2).Value = 1 Else ChkCar(2).Value = 0
        If Master1!AuthWorkshop = 1 Then ChkSrv(0).Value = 1 Else ChkSrv(0).Value = 0
        If Master1!OtherWorkshop = 1 Then ChkSrv(1).Value = 1 Else ChkSrv(1).Value = 0
        
        If Master1!AvgDist1 = 1 Then ChkDis(0).Value = 1 Else ChkDis(0).Value = 0
        If Master1!AvgDist2 = 1 Then ChkDis(1).Value = 1 Else ChkDis(1).Value = 0
        If Master1!AvgDist3 = 1 Then ChkDis(2).Value = 1 Else ChkDis(2).Value = 0
        If Master1!AvgDist4 = 1 Then ChkDis(3).Value = 1 Else ChkDis(3).Value = 0
        
        If Master1!Citibank = 1 Then ChkCard(0).Value = 1 Else ChkCard(0).Value = 0
        If Master1!HDFC = 1 Then ChkCard(1).Value = 1 Else ChkCard(1).Value = 0
        If Master1!ICICI = 1 Then ChkCard(2).Value = 1 Else ChkCard(2).Value = 0
        If Master1!HSBC = 1 Then ChkCard(3).Value = 1 Else ChkCard(3).Value = 0
        If Master1!SBI = 1 Then ChkCard(4).Value = 1 Else ChkCard(4).Value = 0
        If Master1!Diners = 1 Then ChkCard(5).Value = 1 Else ChkCard(5).Value = 0
        If Master1!Stanchart = 1 Then ChkCard(6).Value = 1 Else ChkCard(6).Value = 0
        If Master1!Amex = 1 Then ChkCard(7).Value = 1 Else ChkCard(7).Value = 0
        If Master1!BOB = 1 Then ChkCard(8).Value = 1 Else ChkCard(8).Value = 0
        If Master1!OtherBank <> "" Then ChkCard(3).Value = 1 Else ChkCard(3).Value = 0
        If Master1!OtherBank <> "" Then Txt(CardOther) = Master1!OtherBank Else ChkCard(3).Value = 0
        
        If Master1!IA = 1 Then ChkFFM(0).Value = 1 Else ChkFFM(0).Value = 0
        If Master1!Jet = 1 Then ChkFFM(1).Value = 1 Else ChkFFM(1).Value = 0
        If Master1!Sahara = 1 Then ChkFFM(2).Value = 1 Else ChkFFM(2).Value = 0
        If Master1!OtherAir <> "" Then ChkFFM(3).Value = 1 Else ChkFFM(3).Value = 0
        If Master1!OtherAir <> "" Then Txt(FFMOther) = Master1!OtherAir Else ChkFFM(3).Value = 0
        
        If Master1!TravelRelated = 1 Then ChkInt(0).Value = 1 Else ChkInt(0).Value = 0
        If Master1!LifeRelated = 1 Then ChkInt(1).Value = 1 Else ChkInt(1).Value = 0
        If Master1!Entertainment = 1 Then ChkInt(2).Value = 1 Else ChkInt(2).Value = 0
        If Master1!Parsonal = 1 Then ChkInt(3).Value = 1 Else ChkInt(3).Value = 0
        If Master1!CarRelated = 1 Then ChkInt(4).Value = 1 Else ChkInt(4).Value = 0
        If Master1!OtherInterest <> "" Then ChkInt(5).Value = 1 Else ChkInt(5).Value = 0
        If Master1!OtherInterest <> "" Then Txt(InterestOther) = Master1!OtherInterest Else ChkInt(5).Value = 0
        
        Txt(DInvNo).TEXT = VNull(Master1!DInv_No)
        Txt(DInvDate).TEXT = VNull(Master1!DInv_Date)
        Txt(Model).TEXT = VNull(Master1!Model)
        Txt(VarientMod).TEXT = VNull(Master1!Varient)
        Txt(RegNo).TEXT = VNull(Master1!RegNo)
        Txt(Color).TEXT = VNull(Master1!Color)
        Txt(Chassis).TEXT = VNull(Master1!Chassis)
        Txt(Eligibility).TEXT = VNull(Master1!CEligibility)
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
Private Sub DGCont_Click()
If rsCont.RecordCount > 0 Then
    Txt(DInvNo).TEXT = rsCont!Inv_No
    Txt(DInvDate).TEXT = rsCont!Sal_VDate
    Txt(Model).TEXT = rsCont!Model
    Txt(Chassis).TEXT = rsCont!ChassisNo
    Txt(Color).TEXT = rsCont!Col_Desc
End If
'Txt(MyIndex).SetFocus
'DGCont.Visible = False
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
Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

