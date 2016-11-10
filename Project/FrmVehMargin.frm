VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmVehMargin 
   BackColor       =   &H00DAD9CF&
   Caption         =   "Vehicle Margine Statistics"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   FillColor       =   &H00400040&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
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
      Left            =   855
      TabIndex        =   120
      Top             =   2400
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "FrmVehMargin.frx":0000
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
         TabIndex        =   125
         ToolTipText     =   "Printer "
         Top             =   285
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "FrmVehMargin.frx":030A
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
         TabIndex        =   124
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "FrmVehMargin.frx":0614
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
         TabIndex        =   123
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
         Picture         =   "FrmVehMargin.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   122
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
         Picture         =   "FrmVehMargin.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   121
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
         Index           =   48
         Left            =   0
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
   Begin MSDataGridLib.DataGrid DGCont 
      Height          =   2910
      Left            =   3915
      Negotiate       =   -1  'True
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   4560
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
      BeginProperty Column01 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   66
      Left            =   10290
      TabIndex        =   63
      Top             =   6495
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   65
      Left            =   10290
      TabIndex        =   62
      Top             =   6225
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   64
      Left            =   10290
      TabIndex        =   61
      Top             =   5955
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   63
      Left            =   6825
      TabIndex        =   60
      Top             =   6795
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   62
      Left            =   6825
      TabIndex        =   59
      Top             =   6525
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   61
      Left            =   6825
      TabIndex        =   58
      Top             =   6255
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   60
      Left            =   6825
      TabIndex        =   57
      Top             =   5985
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   57
      Left            =   2970
      TabIndex        =   54
      Top             =   6240
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   59
      Left            =   2970
      TabIndex        =   56
      Top             =   6780
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   58
      Left            =   2970
      TabIndex        =   55
      Top             =   6510
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   56
      Left            =   2970
      TabIndex        =   53
      Top             =   5970
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   51
      Left            =   10290
      TabIndex        =   48
      Top             =   4845
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   49
      Left            =   10290
      TabIndex        =   46
      Top             =   4575
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   48
      Left            =   9030
      TabIndex        =   45
      Top             =   4575
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   50
      Left            =   9030
      TabIndex        =   47
      Top             =   4845
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   55
      Left            =   10290
      TabIndex        =   52
      Top             =   5505
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   53
      Left            =   10290
      TabIndex        =   50
      Top             =   5175
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   47
      Left            =   10290
      TabIndex        =   44
      Top             =   4305
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   54
      Left            =   9030
      TabIndex        =   51
      Top             =   5505
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   52
      Left            =   9030
      TabIndex        =   49
      Top             =   5175
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   45
      Left            =   10290
      TabIndex        =   42
      Top             =   4035
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   43
      Left            =   10290
      TabIndex        =   40
      Top             =   3765
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   44
      Left            =   9030
      TabIndex        =   41
      Top             =   4035
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   46
      Left            =   9030
      TabIndex        =   43
      Top             =   4305
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   42
      Left            =   9030
      TabIndex        =   39
      Top             =   3765
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   38
      Left            =   10290
      TabIndex        =   38
      Top             =   3060
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   36
      Left            =   10290
      TabIndex        =   36
      Top             =   2730
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   34
      Left            =   10290
      TabIndex        =   34
      Top             =   2415
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   32
      Left            =   6105
      MaxLength       =   50
      TabIndex        =   32
      Top             =   2415
      Width           =   2775
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   29
      Left            =   6120
      MaxLength       =   50
      TabIndex        =   29
      Top             =   2145
      Width           =   2775
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   26
      Left            =   6120
      MaxLength       =   50
      TabIndex        =   26
      Top             =   1875
      Width           =   2775
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   37
      Left            =   9030
      TabIndex        =   37
      Top             =   3060
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   35
      Left            =   9030
      TabIndex        =   35
      Top             =   2730
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   31
      Left            =   10290
      TabIndex        =   31
      Top             =   2145
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   28
      Left            =   10290
      TabIndex        =   28
      Top             =   1875
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   30
      Left            =   9030
      TabIndex        =   30
      Top             =   2145
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   33
      Left            =   9030
      TabIndex        =   33
      Top             =   2415
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   27
      Left            =   9030
      TabIndex        =   27
      Top             =   1875
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   25
      Left            =   4440
      TabIndex        =   25
      Top             =   4590
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   24
      Left            =   2970
      TabIndex        =   24
      Top             =   4590
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   23
      Left            =   4440
      TabIndex        =   23
      Top             =   4320
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   22
      Left            =   2970
      TabIndex        =   22
      Top             =   4320
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   21
      Left            =   4440
      TabIndex        =   21
      Top             =   3810
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   20
      Left            =   2970
      TabIndex        =   20
      Top             =   3810
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   18
      Left            =   2970
      TabIndex        =   18
      Top             =   3480
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   9
      Left            =   4440
      TabIndex        =   9
      Top             =   2130
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   19
      Left            =   4440
      TabIndex        =   19
      Top             =   3480
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   17
      Left            =   4440
      TabIndex        =   17
      Top             =   3210
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   15
      Left            =   4440
      TabIndex        =   15
      Top             =   2940
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   13
      Left            =   4440
      TabIndex        =   13
      Top             =   2670
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   11
      Top             =   2400
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   1860
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   8
      Left            =   2970
      TabIndex        =   8
      Top             =   2130
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   16
      Left            =   2970
      TabIndex        =   16
      Top             =   3210
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   14
      Left            =   2970
      TabIndex        =   14
      Top             =   2940
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   12
      Left            =   2970
      TabIndex        =   12
      Top             =   2670
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   10
      Left            =   2970
      TabIndex        =   10
      Top             =   2400
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   2970
      TabIndex        =   6
      Top             =   1860
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   9210
      TabIndex        =   5
      Top             =   720
      Width           =   2445
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   1905
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   2
      Top             =   450
      Width           =   1905
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   9210
      TabIndex        =   4
      Top             =   450
      Width           =   2445
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   1725
      TabIndex        =   1
      Top             =   720
      Width           =   2445
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   1725
      MaxLength       =   40
      TabIndex        =   0
      Top             =   450
      Width           =   2445
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   117
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   661
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   47
      Left            =   10455
      TabIndex        =   118
      Top             =   1470
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000.00"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   46
      Left            =   10260
      TabIndex        =   116
      Top             =   6840
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET MARGIN -->"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   45
      Left            =   8610
      TabIndex        =   115
      Top             =   6855
      Width           =   1470
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00008080&
      BorderWidth     =   3
      FillColor       =   &H0080C0FF&
      Height          =   330
      Left            =   8505
      Shape           =   4  'Rounded Rectangle
      Top             =   6780
      Width           =   3045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RTO"
      Height          =   195
      Index           =   44
      Left            =   8460
      TabIndex        =   114
      Top             =   6495
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount / SV"
      Height          =   195
      Index           =   43
      Left            =   8460
      TabIndex        =   113
      Top             =   6225
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc (Compli. Puja)"
      Height          =   195
      Index           =   42
      Left            =   8460
      TabIndex        =   112
      Top             =   5985
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Petrol / No Plate"
      Height          =   195
      Index           =   41
      Left            =   4470
      TabIndex        =   111
      Top             =   6795
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invent.Carrying Cost"
      Height          =   195
      Index           =   40
      Left            =   4470
      TabIndex        =   110
      Top             =   6525
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Incentive"
      Height          =   195
      Index           =   39
      Left            =   4470
      TabIndex        =   109
      Top             =   6255
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mgn on Spl No/Insu./Affid."
      Height          =   195
      Index           =   38
      Left            =   4470
      TabIndex        =   108
      Top             =   5985
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Margine on Regn."
      Height          =   195
      Index           =   37
      Left            =   960
      TabIndex        =   107
      Top             =   6240
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finance Incentive"
      Height          =   195
      Index           =   36
      Left            =   960
      TabIndex        =   106
      Top             =   6780
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Corp Inc(TELCO)"
      Height          =   195
      Index           =   35
      Left            =   960
      TabIndex        =   105
      Top             =   6510
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Margine of Vehicle"
      Height          =   195
      Index           =   34
      Left            =   960
      TabIndex        =   104
      Top             =   5970
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   103
      Top             =   6645
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   102
      Top             =   6885
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   101
      Top             =   6405
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   100
      Top             =   6090
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1290
      Left            =   180
      Top             =   5895
      Width           =   570
   End
   Begin VB.Shape Shape3 
      Height          =   1305
      Index           =   0
      Left            =   180
      Top             =   5895
      Width           =   11520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL -->"
      Height          =   195
      Index           =   33
      Left            =   7500
      TabIndex        =   99
      Top             =   5235
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DMA/Brokerage"
      Height          =   195
      Index           =   32
      Left            =   6210
      TabIndex        =   98
      Top             =   4845
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accessories Fitment"
      Height          =   195
      Index           =   31
      Left            =   6210
      TabIndex        =   97
      Top             =   4575
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment in Reg. Charges"
      Height          =   195
      Index           =   30
      Left            =   6225
      TabIndex        =   96
      Top             =   4035
      Width           =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Payment to Customer"
      Height          =   195
      Index           =   29
      Left            =   6210
      TabIndex        =   95
      Top             =   4305
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amt Paid"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   23
      Left            =   10470
      TabIndex        =   94
      Top             =   3405
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment  in Vehicle Cost"
      Height          =   195
      Index           =   28
      Left            =   6210
      TabIndex        =   93
      Top             =   3765
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bal. From/To Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   27
      Left            =   6180
      TabIndex        =   92
      Top             =   5520
      Width           =   2400
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   6060
      X2              =   11685
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   6060
      X2              =   11685
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   6060
      X2              =   11670
      Y1              =   3690
      Y2              =   3690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details of Dis./Subvention"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   24
      Left            =   6090
      TabIndex        =   91
      Top             =   3405
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amt.Adjusted"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   22
      Left            =   9075
      TabIndex        =   90
      Top             =   3405
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      Height          =   4815
      Index           =   1
      Left            =   6045
      Top             =   1020
      Width           =   5640
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   6045
      X2              =   11670
      Y1              =   3345
      Y2              =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bal. From/To Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   26
      Left            =   6180
      TabIndex        =   89
      Top             =   3075
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL -->"
      Height          =   195
      Index           =   25
      Left            =   7500
      TabIndex        =   88
      Top             =   2790
      Width           =   900
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   6060
      X2              =   11685
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   6060
      X2              =   11685
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6060
      X2              =   11670
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADJ.of Excess Amt.(Bill No/Date)"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   21
      Left            =   6090
      TabIndex        =   87
      Top             =   1470
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   20
      Left            =   10560
      TabIndex        =   86
      Top             =   1905
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   19
      Left            =   9210
      TabIndex        =   85
      Top             =   1470
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shortage -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   18
      Left            =   1245
      TabIndex        =   84
      Top             =   4590
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excess -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   17
      Left            =   1455
      TabIndex        =   83
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL -->"
      Height          =   195
      Index           =   16
      Left            =   1635
      TabIndex        =   82
      Top             =   3870
      Width           =   900
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   195
      X2              =   5820
      Y1              =   4155
      Y2              =   4155
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   195
      X2              =   5820
      Y1              =   3765
      Y2              =   3765
   End
   Begin VB.Shape Shape1 
      Height          =   4815
      Index           =   0
      Left            =   195
      Top             =   1035
      Width           =   5640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reg.Charges"
      Height          =   195
      Index           =   10
      Left            =   405
      TabIndex        =   81
      Top             =   2130
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EWP charges"
      Height          =   195
      Index           =   15
      Left            =   405
      TabIndex        =   80
      Top             =   3480
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accessories Charges"
      Height          =   195
      Index           =   14
      Left            =   405
      TabIndex        =   79
      Top             =   3210
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Charges"
      Height          =   195
      Index           =   13
      Left            =   405
      TabIndex        =   78
      Top             =   2940
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Affidavit Charges"
      Height          =   195
      Index           =   12
      Left            =   405
      TabIndex        =   77
      Top             =   2670
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special no. Charges"
      Height          =   195
      Index           =   11
      Left            =   405
      TabIndex        =   76
      Top             =   2400
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Price"
      Height          =   195
      Index           =   9
      Left            =   405
      TabIndex        =   75
      Top             =   1860
      Width           =   1140
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   210
      X2              =   5820
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   8
      Left            =   855
      TabIndex        =   74
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amt. Recieved"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   7
      Left            =   4395
      TabIndex        =   73
      Top             =   1485
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amt.Payable"
      ForeColor       =   &H00008080&
      Height          =   195
      Index           =   6
      Left            =   3135
      TabIndex        =   72
      Top             =   1485
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OUTFLOW"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   8325
      TabIndex        =   71
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "INFLOW"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   0
      Left            =   2595
      TabIndex        =   70
      Top             =   1110
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   1
      Left            =   6045
      Top             =   1035
      Width           =   5640
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   195
      Top             =   1050
      Width           =   5640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model / CLR"
      Height          =   195
      Index           =   5
      Left            =   7980
      TabIndex        =   69
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin. amount"
      Height          =   195
      Index           =   4
      Left            =   4245
      TabIndex        =   68
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin. Code"
      Height          =   195
      Index           =   3
      Left            =   4425
      TabIndex        =   67
      Top             =   450
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
      Height          =   195
      Index           =   2
      Left            =   390
      TabIndex        =   66
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seles Inv. No/Date"
      Height          =   195
      Index           =   1
      Left            =   7410
      TabIndex        =   65
      Top             =   450
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   64
      Top             =   450
      Width           =   1335
   End
End
Attribute VB_Name = "FrmVehMargin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Don't Change Tag Property of (Txt) Control as it is used in other activities
'FORM COLOR &H00C0FFFF&
Option Explicit
Public FormExit As Boolean
Dim ADDFLAG As Byte
Dim Master As ADODB.Recordset
Dim Master1 As ADODB.Recordset
Dim rsCont As ADODB.Recordset
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Const CustName = 0
Private Const SPerson = 1
Private Const FinCode = 2
Private Const FinAmt = 3
Private Const SalesInv = 4
Private Const Model = 5
'INFLOW CONTROLS
Private Const VehPriceAP = 6
Private Const VehPriceAR = 7
Private Const RegChrgAP = 8
Private Const RegChrgAR = 9
Private Const SPNoChrgAP = 10
Private Const SPNoChrgAR = 11
Private Const AffiChrgAP = 12
Private Const AffiChrgAR = 13
Private Const InsuChrgAP = 14
Private Const InsuChrgAR = 15
Private Const AccessoriesAP = 16
Private Const AccessoriesAR = 17
Private Const EWPChrgAP = 18
Private Const EWPChrgAR = 19
Private Const TotalAP = 20
Private Const TotalAR = 21
Private Const ExcessAP = 22
Private Const ExcessAR = 23
Private Const ShortageAP = 24
Private Const ShortageAR = 25
'OUTFLOW CONTROLS
Private Const AdjBill1 = 26
Private Const AdjBill1STot = 27
Private Const AdjBill1Tot = 28
Private Const AdjBill2 = 29
Private Const AdjBill2STot = 30
Private Const AdjBill2Tot = 31
Private Const AdjBill3 = 32
Private Const AdjBill3STot = 33
Private Const AdjBill3Tot = 34
Private Const AdjBillSTotal = 35
Private Const AdjBillTotal = 36
Private Const AdjBillBalSTotal = 37
Private Const AdjBillBalTotal = 38

Private Const AdjVehCostAA = 42
Private Const AdjVehCostAP = 43
Private Const AdjRegChrgAA = 44
Private Const AdjRegChrgAP = 45
Private Const CashPayAA = 46
Private Const CashPayAP = 47
Private Const FitmentAA = 48
Private Const FitmentAP = 49
Private Const DMAAA = 50
Private Const DMAAP = 51
Private Const DisTotalAA = 52
Private Const DisTotalAP = 53
Private Const DisBalAA = 54
Private Const DisBalAP = 55

Private Const MrgVeh = 56
Private Const MrgRegn = 57
Private Const CorpInc = 58
Private Const FinIncentive = 59
Private Const MrgSplNo = 60
Private Const PurIncentive = 61
Private Const InventoryCost = 62
Private Const Petrol = 63
Private Const Misc = 64
Private Const Discount = 65
Private Const RTO = 66
Dim i As Integer
Private Sub Form_Deactivate()
   Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
ELoop:
     CheckError
End Sub
Private Sub Form_Load()
Dim i As Byte
    WinSetting Me: Ini_Grid
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Sl_No as SearchCode From Veh_Margin order by Sl_No desc", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsCont = New ADODB.Recordset
    rsCont.CursorLocation = adUseClient
    rsCont.Open "Select right(VS.Sal_DocID,8) as Inv_No,SG.Name as Cust_Name,VS.Sal_VDate,VS.MODEL,VS.Sal_VNo,Emp_Mast.Emp_Name,VO.Fin_AcCode,VO.Fin_Amt from ((Veh_Stock VS Left Join Veh_Order VO on VS.Sal_DocId=VO.Inv_DocID) Left Join Emp_Mast on VO.Rep_Code=Emp_Mast.Emp_Code) Left Join SubGroup SG On VO.PartyCode=SG.SubCode where Len(VS.Sal_DocID)  > 1", GCn, adOpenDynamic, adLockOptimistic
    Set DGCont.DataSource = rsCont
    rsCont.Sort = "Inv_No"
    rsCont.Sort = "Cust_Name"
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
 End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    Disp_Text SETS("ADD", Me, Master)
    BlankText
    Txt(CustName).SetFocus
    Ini_Grid
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM, j As Byte, TmpSQL As String, mTran As Boolean
On Error GoTo ELoop
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        mTran = True
        GCn.Execute ("Delete From Veh_Margin Where Sl_No=" & Txt(CustName).Tag & "")
        GCn.CommitTrans
        mTran = False
        BUTTONS True, Me, Master, 0
        Master.Requery
        MoveRec
    End If
    Exit Sub
ELoop:
    If mTran = True Then GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message", App.Title
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records to Edit.", vbInformation, "Information": Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    Txt(CustName).SetFocus
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT Sl_No As SearchCode, CustName as CustomerName,SPerson as SalePerson,FinCode as FinencerCode, " & cTrim("SalInv") & " as SaleInvoice " & _
            "FROM VEH_Margin " & _
            "Order By Sl_No"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ELoop:
    CheckError
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eCancel()
Dim i As Byte
On Error GoTo ELoop
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        Call Ini_Grid
        MoveRec
    End If
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
    LblPrinter.CAPTION = Printer.DeviceName
    If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
End Sub
Private Sub TopCtrl1_eRef()
    Master.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, TmpSQL$, MaxCode%, i%, MyOldNo$
Dim SlNo As Double
On Error GoTo ELoop
    Grid_Hide
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        SlNo = VNull(GCn.Execute("Select max(Sl_No) from Veh_Margin").Fields(0).Value) + 1
        TmpSQL = "Insert Into Veh_Margin (Sl_No,CustName,SPerson,FinCode,FinAmt,SalInv,Model,VehPriceAP,VehPriceAR,RegChrgAP," _
                  & "RegChrgAR,SPNoChrgAP,SPNoChrgAR,AffiChrgAP,AffiChrgAR,InsuChrgAP,InsuChrgAR,EWPChrgAP,EWPChrgAR," _
                  & "AccessoriesAP,AccessoriesAR,TotalAP,TotalAR,ExcessAP,ExcessAR,ShortageAP,ShortageAR,AdjBill1,AdjBill2," _
                  & "AdjBill3,AdjBill1STot,AdjBill1Tot,AdjBill2STot,AdjBill2Tot,AdjBill3STot,AdjBill3Tot,AdjBillSTotal," _
                  & "AdjBillTotal,AdjBillBalSTot,AdjBillBalTot,AdjVehCostAA,AdjVehCostAP,AdjRegChrgAA,AdjRegChrgAP,CashPayAA," _
                  & "CashPayAP,FitmentAA,FitmentAP,DMAAA,DMAAP,DisTotalAA,DisTotalAP,DisBalAA,DisBalAP,MrgVeh,MrgRegn,CorpInc," _
                  & "FinIncentive,MrgSplNo,PurIncentive,InventCost,Petrol,Discount,Misc,RTO,NetMrg,U_Name,U_EntDt,U_AE) " _
                  & " Values (" & SlNo & ",'" & Txt(CustName) & "','" & Txt(SPerson) & "','" & Txt(FinCode) & "'," & Val(Txt(FinAmt)) & ",'" & Txt(SalesInv) & "'" & _
                   ",'" & Txt(Model) & "'," & Val(Txt(VehPriceAP)) & "," & Val(Txt(VehPriceAR)) & "," & Val(Txt(RegChrgAP)) & "," & Val(Txt(RegChrgAR)) & "" & _
                   "," & Val(Txt(SPNoChrgAP)) & "," & Val(Txt(SPNoChrgAR)) & "," & Val(Txt(AffiChrgAP)) & "," & Val(Txt(AffiChrgAR)) & "," & Val(Txt(InsuChrgAP)) & "" & _
                   "," & Val(Txt(InsuChrgAR)) & "," & Val(Txt(EWPChrgAP)) & "," & Val(Txt(EWPChrgAR)) & "," & Val(Txt(AccessoriesAP)) & "," & Val(Txt(AccessoriesAR)) & "" & _
                   "," & Val(Txt(TotalAP)) & "," & Val(Txt(TotalAR)) & "," & Val(Txt(ExcessAP)) & "," & Val(Txt(ExcessAR)) & "," & Val(Txt(ShortageAP)) & "," & Val(Txt(ShortageAR)) & "" & _
                   "," & Val(Txt(AdjBill1)) & "," & Val(Txt(AdjBill2)) & "," & Val(Txt(AdjBill3)) & "," & Val(Txt(AdjBill1STot)) & "," & Val(Txt(AdjBill1Tot)) & "" & _
                   "," & Val(Txt(AdjBill2STot)) & "," & Val(Txt(AdjBill2Tot)) & "," & Val(Txt(AdjBill3STot)) & "," & Val(Txt(AdjBill3Tot)) & "," & Val(Txt(AdjBillSTotal)) & "" & _
                   "," & Val(Txt(AdjBillTotal)) & "," & Val(Txt(AdjBillBalSTotal)) & "," & Val(Txt(AdjBillBalTotal)) & "," & Val(Txt(AdjVehCostAA)) & "," & Val(Txt(AdjVehCostAP)) & "" & _
                   "," & Val(Txt(AdjRegChrgAA)) & "," & Val(Txt(AdjRegChrgAP)) & "," & Val(Txt(CashPayAA)) & "," & Val(Txt(CashPayAP)) & "," & Val(Txt(FitmentAA)) & "," & Val(Txt(FitmentAP)) & "" & _
                   "," & Val(Txt(DMAAA)) & "," & Val(Txt(DMAAP)) & "," & Val(Txt(DisTotalAA)) & "," & Val(Txt(DisTotalAP)) & "," & Val(Txt(DisBalAA)) & "," & Val(Txt(DisBalAP)) & "" & _
                   "," & Val(Txt(MrgVeh)) & "," & Val(Txt(MrgRegn)) & "," & Val(Txt(CorpInc)) & "," & Val(Txt(FinIncentive)) & "," & Val(Txt(MrgSplNo)) & "," & Val(Txt(PurIncentive)) & "" & _
                   "," & Val(Txt(InventoryCost)) & "," & Val(Txt(Petrol)) & "," & Val(Txt(Discount)) & "," & Val(Txt(Misc)) & "," & Val(Txt(RTO)) & "," & Val(Label1(46).CAPTION) & ",'" & pubUName & "',#" & PubServerDate & "#,'" & IIf(ADDFLAG = 1, "A", "E") & "')"
    Else
            GCn.Execute ("Delete from Veh_Margin where Sl_No = " & Txt(CustName).Tag & "")
            TmpSQL = "Insert Into Veh_Margin (Sl_No,CustName,SPerson,FinCode,FinAmt,SalInv,Model,VehPriceAP,VehPriceAR,RegChrgAP," _
                  & "RegChrgAR,SPNoChrgAP,SPNoChrgAR,AffiChrgAP,AffiChrgAR,InsuChrgAP,InsuChrgAR,EWPChrgAP,EWPChrgAR," _
                  & "AccessoriesAP,AccessoriesAR,TotalAP,TotalAR,ExcessAP,ExcessAR,ShortageAP,ShortageAR,AdjBill1,AdjBill2," _
                  & "AdjBill3,AdjBill1STot,AdjBill1Tot,AdjBill2STot,AdjBill2Tot,AdjBill3STot,AdjBill3Tot,AdjBillSTotal," _
                  & "AdjBillTotal,AdjBillBalSTot,AdjBillBalTot,AdjVehCostAA,AdjVehCostAP,AdjRegChrgAA,AdjRegChrgAP,CashPayAA," _
                  & "CashPayAP,FitmentAA,FitmentAP,DMAAA,DMAAP,DisTotalAA,DisTotalAP,DisBalAA,DisBalAP,MrgVeh,MrgRegn,CorpInc," _
                  & "FinIncentive,MrgSplNo,PurIncentive,InventCost,Petrol,Discount,Misc,RTO,NetMrg,U_Name,U_EntDt,U_AE) " _
                  & " Values (" & Val(Txt(CustName).Tag) & ",'" & Txt(CustName) & "','" & Txt(SPerson) & "','" & Txt(FinCode) & "'," & Val(Txt(FinAmt)) & ",'" & Txt(SalesInv) & "'" & _
                   ",'" & Txt(Model) & "'," & Val(Txt(VehPriceAP)) & "," & Val(Txt(VehPriceAR)) & "," & Val(Txt(RegChrgAP)) & "," & Val(Txt(RegChrgAR)) & "" & _
                   "," & Val(Txt(SPNoChrgAP)) & "," & Val(Txt(SPNoChrgAR)) & "," & Val(Txt(AffiChrgAP)) & "," & Val(Txt(AffiChrgAR)) & "," & Val(Txt(InsuChrgAP)) & "" & _
                   "," & Val(Txt(InsuChrgAR)) & "," & Val(Txt(EWPChrgAP)) & "," & Val(Txt(EWPChrgAR)) & "," & Val(Txt(AccessoriesAP)) & "," & Val(Txt(AccessoriesAR)) & "" & _
                   "," & Val(Txt(TotalAP)) & "," & Val(Txt(TotalAR)) & "," & Val(Txt(ExcessAP)) & "," & Val(Txt(ExcessAR)) & "," & Val(Txt(ShortageAP)) & "," & Val(Txt(ShortageAR)) & "" & _
                   "," & Val(Txt(AdjBill1)) & "," & Val(Txt(AdjBill2)) & "," & Val(Txt(AdjBill3)) & "," & Val(Txt(AdjBill1STot)) & "," & Val(Txt(AdjBill1Tot)) & "" & _
                   "," & Val(Txt(AdjBill2STot)) & "," & Val(Txt(AdjBill2Tot)) & "," & Val(Txt(AdjBill3STot)) & "," & Val(Txt(AdjBill3Tot)) & "," & Val(Txt(AdjBillSTotal)) & "" & _
                   "," & Val(Txt(AdjBillTotal)) & "," & Val(Txt(AdjBillBalSTotal)) & "," & Val(Txt(AdjBillBalTotal)) & "," & Val(Txt(AdjVehCostAA)) & "," & Val(Txt(AdjVehCostAP)) & "" & _
                   "," & Val(Txt(AdjRegChrgAA)) & "," & Val(Txt(AdjRegChrgAP)) & "," & Val(Txt(CashPayAA)) & "," & Val(Txt(CashPayAP)) & "," & Val(Txt(FitmentAA)) & "," & Val(Txt(FitmentAP)) & "" & _
                   "," & Val(Txt(DMAAA)) & "," & Val(Txt(DMAAP)) & "," & Val(Txt(DisTotalAA)) & "," & Val(Txt(DisTotalAP)) & "," & Val(Txt(DisBalAA)) & "," & Val(Txt(DisBalAP)) & "" & _
                   "," & Val(Txt(MrgVeh)) & "," & Val(Txt(MrgRegn)) & "," & Val(Txt(CorpInc)) & "," & Val(Txt(FinIncentive)) & "," & Val(Txt(MrgSplNo)) & "," & Val(Txt(PurIncentive)) & "" & _
                   "," & Val(Txt(InventoryCost)) & "," & Val(Txt(Petrol)) & "," & Val(Txt(Discount)) & "," & Val(Txt(Misc)) & "," & Val(Txt(RTO)) & "," & Val(Label1(46).CAPTION) & ",'" & pubUName & "',#" & PubServerDate & "#,'" & IIf(ADDFLAG = 1, "A", "E") & "')"
    End If
    
    GCn.BeginTrans
    mTrans = True
    GCn.Execute (TmpSQL)
    GCn.CommitTrans
    Master.Requery
    TopCtrl1_ePrn
    '.FIND ("SearchCode='" & Txt().TEXT & "'")
    'If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Ini_Grid
    MoveRec
    Exit Sub
ELoop:
    If mTrans = True Then
        GCn.RollbackTrans
    End If
    CheckError
    Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Errloop
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case CustName
            DGridTxtKeyDown DGCont, Txt, Index, rsCont, KeyCode, False, 1
            DGCont_Click
    End Select
    If DGCont.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> RTO Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = RTO Then
             If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbCritical, App.Title: Exit Sub
End Sub
Private Sub Grid_Hide()
    If DGCont.Visible = True Then DGCont.Visible = False
End Sub
Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
    CheckQuote keyascii
    Select Case Index
        Case VehPriceAP, VehPriceAR, RegChrgAP, RegChrgAR, SPNoChrgAP, SPNoChrgAR, AffiChrgAP, AffiChrgAR, _
             InsuChrgAP, InsuChrgAR, EWPChrgAP, EWPChrgAR
             
            Call NumPress(Txt(Index), keyascii, 6, 2)
            
        Case AdjBill1STot, AdjBill2STot, AdjBill3STot, AdjBill1Tot, AdjBill2Tot, AdjBill3Tot, AccessoriesAP, AccessoriesAR, AdjVehCostAA, AdjVehCostAP, _
            AdjRegChrgAA, AdjRegChrgAP, CashPayAA, CashPayAP, FitmentAA, FitmentAP, DMAAA, DMAAP, DisTotalAA, DisTotalAP, DisBalAA, DisBalAP
            
            Call NumPress(Txt(Index), keyascii, 6, 2)
            
        Case MrgVeh, MrgRegn, CorpInc, FinIncentive, MrgSplNo, PurIncentive, InventoryCost, Misc, Discount, RTO
        
            Call NumPress(Txt(Index), keyascii, 6, 2)
            
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case VehPriceAP, VehPriceAR, RegChrgAP, RegChrgAR, SPNoChrgAP, SPNoChrgAR, AffiChrgAP, AffiChrgAR, _
             InsuChrgAP, InsuChrgAR, EWPChrgAP, EWPChrgAR, AccessoriesAP, AccessoriesAR
             
            Txt(TotalAP) = Format(InflowAP(), "0.00"): Txt(TotalAR) = Format(InflowAR(), "0.00")
            Txt(ExcessAR) = Format(IIf(Val(Txt(TotalAR)) > Val(Txt(TotalAP)), Val(Txt(TotalAR)) - Val(Txt(TotalAP)), 0), "0.00")
            Txt(ShortageAR) = Format(IIf(Val(Txt(TotalAP)) > Val(Txt(TotalAR)), Val(Txt(TotalAP)) - Val(Txt(TotalAR)), 0), "0.00")
            ' IF BALANCE IN EXCESS THEN
            If Val(Txt(ExcessAR)) > 0 Then
                'ENABLE EXCESS BILL ADJUSTMENT AMOUNT
                Txt(AdjBill1).Enabled = True: Txt(AdjBill2).Enabled = True: Txt(AdjBill3).Enabled = True
                Txt(AdjBill1STot).Enabled = True: Txt(AdjBill2STot).Enabled = True: Txt(AdjBill3STot).Enabled = True
            ElseIf Val(Txt(ExcessAR)) <= 0 Then
                'DISABLE EXCESS BILL ADJUSTMENT AMOUNT
                For i = 26 To 38
                    Txt(i).Enabled = False
                Next
            End If
        Case AdjBill1STot, AdjBill2STot, AdjBill3STot
            'IF ADJUSTMENT GREATER THEN EXCESS AMOUNT
            If OutFlowStot > Val(Txt(ExcessAR)) Then MsgBox " Not a Valid Adjustment ! Please Check Amount.": Txt(Index) = 0
            ' IF BALANCE IN EXCESS THEN
            If Val(Txt(ExcessAR)) > 0 Then
                If Index > AdjBill1STot Then
                    Txt(Index + 1) = Format(Val(Txt(Index)) + Val(Txt(Index - 2)), "0.00")
                Else
                    Txt(Index + 1) = Format(Val(Txt(Index)), "0.00")
                End If
                Txt(AdjBillTotal) = Format(OutFlowStot, "0.00")
                Txt(AdjBillBalTotal) = Format(Val(Txt(ExcessAR)) - Val(Txt(AdjBillTotal)), "0.00")
            End If
        Case AdjVehCostAA, AdjVehCostAP, AdjRegChrgAA, AdjRegChrgAP, CashPayAA, CashPayAP, FitmentAA, FitmentAP, DMAAA, DMAAP
                If DisTotal_AA > Val(Txt(AdjBillBalTotal)) Then MsgBox " Not a Valid Adjustment ! Please Check Amount.": Txt(Index) = 0
                Txt(DisTotalAA) = Format(DisTotal_AA, "0.00")
                Txt(DisTotalAP) = Format(DisTotal_AR, "0.00")
                Txt(DisBalAP) = Format(Val(Txt(AdjBillBalTotal)) - (Val(Txt(DisTotalAA)) + Val(Txt(DisTotalAP))), "0.00")
        Case MrgVeh, MrgRegn, CorpInc, FinIncentive, MrgSplNo, PurIncentive, InventoryCost, Misc, Discount, RTO
                Label1(46).CAPTION = Format(CalcNetMrg() - Val(Txt(DisBalAP)), "0.00")
    End Select
End Sub
Private Sub DGCont_Click()
If rsCont.RecordCount > 0 Then
    Txt(CustName).TEXT = XNull(rsCont!CUST_NAME)
    Txt(SPerson).TEXT = XNull(rsCont!Emp_Name)
    Txt(FinCode).TEXT = XNull(rsCont!Fin_AcCode)
    Txt(FinAmt).TEXT = VNull(rsCont!Fin_Amt)
    Txt(SalesInv).TEXT = XNull(rsCont!Inv_No)
    Txt(Model).TEXT = XNull(rsCont!Model)
End If
End Sub
Private Sub Ini_Grid()
    DGCont.width = 5000: DGCont.left = Me.width - (DGCont.width + mRtScale): DGCont.top = mTopScale: DGCont.height = 5000
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
     Select Case Index
        Case VehPriceAP, VehPriceAR, RegChrgAP, RegChrgAR, SPNoChrgAP, SPNoChrgAR, AffiChrgAP, AffiChrgAR, _
             InsuChrgAP, InsuChrgAR, EWPChrgAP, EWPChrgAR
             
             Txt(Index) = Format(Txt(Index), "0.00")
             Txt(Index).Alignment = vbRightJustify
             
             
        Case AdjBill1STot, AdjBill2STot, AdjBill3STot, AdjBill1Tot, AdjBill2Tot, AdjBill3Tot, AccessoriesAP, AccessoriesAR, AdjVehCostAA, AdjVehCostAP, _
            AdjRegChrgAA, AdjRegChrgAP, CashPayAA, CashPayAP, FitmentAA, FitmentAP, DMAAA, DMAAP, DisTotalAA, DisTotalAP, DisBalAA, DisBalAP
            
            Txt(Index) = Format(Txt(Index), "0.00")
            Txt(Index).Alignment = vbRightJustify
            
        Case MrgVeh, MrgRegn, CorpInc, FinIncentive, MrgSplNo, PurIncentive, InventoryCost, Misc, Discount, RTO
        
            Txt(Index) = Format(Txt(Index), "0.00")
            Txt(Index).Alignment = vbRightJustify
            
    End Select
End Sub
'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
    For i = 0 To 38
        Txt(i).TEXT = ""
        Txt(i).Tag = ""
    Next i
    For i = 42 To 66
        Txt(i).TEXT = ""
        Txt(i).Tag = ""
    Next i
    Label1(46).CAPTION = ""
End Sub
Private Sub MoveRec()
Dim Rst As New ADODB.Recordset, i As Integer
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select * from Veh_Margin Where Sl_No=" & Master!SearchCode & "", GCn, adOpenStatic, adLockReadOnly
        
        Txt(CustName) = XNull(Master1!CustName)
        Txt(CustName).Tag = VNull(Master1!SL_NO)
        Txt(SPerson) = XNull(Master1!SPerson)
        Txt(FinCode) = XNull(Master1!FinCode)
        Txt(FinAmt) = Format(VNull(Master1!FinAmt), "0.00")
        Txt(SalesInv) = XNull(Master1!SalInv)
        Txt(Model) = XNull(Master1!Model)
        Txt(VehPriceAP) = Format(VNull(Master1!VehPriceAP), "0.00")
        Txt(VehPriceAR) = Format(VNull(Master1!VehPriceAR), "0.00")
        Txt(RegChrgAP) = Format(VNull(Master1!RegChrgAP), "0.00")
        Txt(RegChrgAR) = Format(VNull(Master1!RegChrgAR), "0.00")
        Txt(SPNoChrgAP) = Format(VNull(Master1!SPNoChrgAP), "0.00")
        Txt(SPNoChrgAR) = Format(VNull(Master1!SPNoChrgAR), "0.00")
        Txt(AffiChrgAP) = Format(VNull(Master1!AffiChrgAP), "0.00")
        Txt(AffiChrgAR) = Format(VNull(Master1!AffiChrgAR), "0.00")
        Txt(InsuChrgAP) = Format(VNull(Master1!InsuChrgAP), "0.00")
        Txt(InsuChrgAR) = Format(VNull(Master1!InsuChrgAR), "0.00")
        Txt(EWPChrgAP) = Format(VNull(Master1!EWPChrgAP), "0.00")
        Txt(EWPChrgAR) = Format(VNull(Master1!EWPChrgAR), "0.00")
        
        Txt(AccessoriesAP) = Format(VNull(Master1!AccessoriesAP), "0.00")
        Txt(AccessoriesAR) = Format(VNull(Master1!AccessoriesAR), "0.00")
        Txt(TotalAP) = Format(VNull(Master1!TotalAP), "0.00")
        Txt(TotalAR) = Format(VNull(Master1!TotalAR), "0.00")
        Txt(ExcessAP) = Format(VNull(Master1!ExcessAP), "0.00")
        Txt(ExcessAR) = Format(VNull(Master1!ExcessAR), "0.00")
        Txt(ShortageAP) = Format(VNull(Master1!ShortageAP), "0.00")
        Txt(ShortageAR) = Format(VNull(Master1!ShortageAR), "0.00")
        Txt(AdjBill1) = XNull(Master1!AdjBill1)
        Txt(AdjBill2) = XNull(Master1!AdjBill2)
        Txt(AdjBill3) = XNull(Master1!AdjBill3)
        Txt(AdjBill1STot) = Format(VNull(Master1!AdjBill1STot), "0.00")
        Txt(AdjBill1Tot) = Format(VNull(Master1!AdjBill1Tot), "0.00")
        Txt(AdjBill2STot) = Format(VNull(Master1!AdjBill2STot), "0.00")
        Txt(AdjBill2Tot) = Format(VNull(Master1!AdjBill2Tot), "0.00")
        Txt(AdjBill3STot) = Format(VNull(Master1!AdjBill3STot), "0.00")
        Txt(AdjBill3Tot) = Format(VNull(Master1!AdjBill3Tot), "0.00")
        Txt(AdjBillSTotal) = Format(VNull(Master1!AdjBillSTotal), "0.00")
        Txt(AdjBillTotal) = Format(VNull(Master1!AdjBillTotal), "0.00")
        Txt(AdjBillBalSTotal) = Format(VNull(Master1!AdjBillBalSTot), "0.00")
        Txt(AdjBillBalTotal) = Format(VNull(Master1!AdjBillBalTot), "0.00")
        Txt(AdjVehCostAA) = Format(VNull(Master1!AdjVehCostAA), "0.00")
        Txt(AdjVehCostAP) = Format(VNull(Master1!AdjVehCostAP), "0.00")
        Txt(AdjRegChrgAA) = Format(VNull(Master1!AdjRegChrgAA), "0.00")
        Txt(AdjRegChrgAP) = Format(VNull(Master1!AdjRegChrgAP), "0.00")
        Txt(CashPayAA) = Format(VNull(Master1!CashPayAA), "0.00")
        Txt(CashPayAP) = Format(VNull(Master1!CashPayAP), "0.00")
        Txt(FitmentAA) = Format(VNull(Master1!FitmentAA), "0.00")
        Txt(FitmentAP) = Format(VNull(Master1!FitmentAP), "0.00")
        Txt(DMAAA) = Format(VNull(Master1!DMAAA), "0.00")
        Txt(DMAAP) = Format(VNull(Master1!DMAAP), "0.00")
        Txt(DisTotalAA) = Format(VNull(Master1!DisTotalAA), "0.00")
        Txt(DisTotalAP) = Format(VNull(Master1!DisTotalAP), "0.00")
        Txt(DisBalAA) = Format(VNull(Master1!DisBalAA), "0.00")
        
        Txt(DisBalAP) = Format(VNull(Master1!DisBalAP), "0.00")
        Txt(MrgVeh) = Format(VNull(Master1!MrgVeh), "0.00")
        Txt(MrgRegn) = Format(VNull(Master1!MrgRegn), "0.00")
        Txt(CorpInc) = Format(VNull(Master1!CorpInc), "0.00")
        Txt(FinIncentive) = Format(VNull(Master1!FinIncentive), "0.00")
        Txt(MrgSplNo) = Format(VNull(Master1!MrgSplNo), "0.00")
        Txt(PurIncentive) = Format(VNull(Master1!PurIncentive), "0.00")
        Txt(InventoryCost) = Format(VNull(Master1!InventCost), "0.00")
        Txt(Petrol) = Format(VNull(Master1!Petrol), "0.00")
        Txt(Discount) = Format(VNull(Master1!Discount), "0.00")
        Txt(Misc) = Format(VNull(Master1!Misc), "0.00")
        Txt(RTO) = Format(VNull(Master1!RTO), "0.00")
        Label1(46) = Format(VNull(Master1!NetMrg), "0.00")
    Else
        BlankText
    End If
    Grid_Hide
    Set Master1 = Nothing
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim i As Byte
    For i = 0 To 38
        Txt(i).Enabled = Enb
    Next
    For i = 42 To 66
        Txt(i).Enabled = Enb
    Next
    'DISABLE EXCESS BILL ADJUSTMENT AMOUNT
    For i = 26 To 38
        Txt(i).Enabled = False
    Next
    Txt(SPerson).Enabled = False
    Txt(FinCode).Enabled = False
    Txt(FinAmt).Enabled = False
    Txt(SalesInv).Enabled = False
    Txt(Model).Enabled = False
    Txt(TotalAP).Enabled = False
    Txt(TotalAR).Enabled = False
    Txt(ExcessAP).Enabled = False
    Txt(ExcessAR).Enabled = False
    Txt(ShortageAP).Enabled = False
    Txt(ShortageAR).Enabled = False
    Txt(AdjBillSTotal).Enabled = False
    Txt(AdjBillTotal).Enabled = False
    Txt(AdjBillBalSTotal).Enabled = False
    Txt(AdjBillBalTotal).Enabled = False
    Txt(DisTotalAA).Enabled = False
    Txt(DisTotalAP).Enabled = False
    Txt(DisBalAA).Enabled = False
    Txt(DisBalAP).Enabled = False
    Txt(AdjBill1Tot).Enabled = False
    Txt(AdjBill2Tot).Enabled = False
    Txt(AdjBill3Tot).Enabled = False
End Sub
Private Sub SaveMsg(Index As Integer)
    Grid_Hide
    Me.ActiveControl.SetFocus
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
Select Case Index
    Case PScreen, PWindows
        mRepName = "VehMargin"
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = "VehMargin"
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
Private Sub WindowsPrint(Index As Integer)
Dim mQRY As String, RepTitle$
Dim Condstr$, mDocStr$
Dim Rst1 As ADODB.Recordset
Dim Speciality$
Dim Rst As ADODB.Recordset
Dim i As Integer

On Error GoTo ERRORHANDLER
mQRY = "Select * from Veh_Margin Where Sl_No = " & Txt(CustName).Tag
Set Rst = New Recordset
Rst.CursorLocation = adUseClient
Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic

If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.Database.SetDataSource Rst
rpt.ReadRecords
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
        Call Report_View(rpt, "Vehicle Margin Statistics", , True)
End Select
CmdPrint(PSetUp).Tag = ""
Set Rst = Nothing
Set Rst1 = Nothing
Set rpt = Nothing
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
Private Function InflowAP()
    InflowAP = Val(Txt(VehPriceAP)) + Val(Txt(RegChrgAP)) + Val(Txt(SPNoChrgAP)) + Val(Txt(AffiChrgAP)) + Val(Txt(InsuChrgAP)) + Val(Txt(EWPChrgAP))
End Function
Private Function InflowAR()
    InflowAR = Val(Txt(VehPriceAR)) + Val(Txt(RegChrgAR)) + Val(Txt(SPNoChrgAR)) + Val(Txt(AffiChrgAR)) + Val(Txt(InsuChrgAR)) + Val(Txt(EWPChrgAR))
End Function
Private Function OutFlowStot()
    OutFlowStot = Val(Txt(AdjBill1STot)) + Val(Txt(AdjBill2STot)) + Val(Txt(AdjBill3STot))
End Function
Private Function OutFlowTot()
    OutFlowTot = Val(Txt(AdjBill1Tot)) + Val(Txt(AdjBill2Tot)) + Val(Txt(AdjBill3Tot))
End Function
Private Function DisTotal_AA()
    DisTotal_AA = Val(Txt(AdjVehCostAA)) + Val(Txt(AdjRegChrgAA)) + Val(Txt(CashPayAA)) + Val(Txt(FitmentAA)) + Val(Txt(DMAAA))
End Function
Private Function DisTotal_AR()
    DisTotal_AR = Val(Txt(AdjVehCostAP)) + Val(Txt(AdjRegChrgAP)) + Val(Txt(CashPayAP)) + Val(Txt(FitmentAP)) + Val(Txt(DMAAP))
End Function
Private Function CalcNetMrg()
    CalcNetMrg = (Val(Txt(MrgVeh)) + Val(Txt(MrgRegn)) + Val(Txt(CorpInc)) + Val(Txt(FinIncentive)) + Val(Txt(MrgSplNo)) + Val(Txt(PurIncentive))) - (Val(Txt(InventoryCost)) + Val(Txt(Misc)) + Val(Txt(Discount)) + Val(Txt(RTO)))
End Function

