VERSION 5.00
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmCustConern 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Customer Concern"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   7920
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
      Left            =   2805
      TabIndex        =   33
      Top             =   7275
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "FrmCustConern.frx":0000
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
         Height          =   330
         Index           =   2
         Left            =   3420
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Printer "
         Top             =   285
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "FrmCustConern.frx":030A
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
         TabIndex        =   37
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "FrmCustConern.frx":0614
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
         TabIndex        =   36
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
         Picture         =   "FrmCustConern.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "FrmCustConern.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Index           =   1
         Left            =   0
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7365
      MaxLength       =   11
      TabIndex        =   2
      Top             =   630
      Width           =   1860
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1800
      Width           =   1545
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3945
      MaxLength       =   10
      TabIndex        =   1
      Top             =   630
      Width           =   1560
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   3
      Top             =   893
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1200
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1500
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3945
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1800
      Width           =   1245
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5895
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   12
      Top             =   4065
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   14
      Top             =   4755
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   15
      Top             =   5055
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   16
      Top             =   5475
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   17
      Top             =   5775
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   18
      Left            =   3930
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   6300
      Width           =   5295
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   10
      Left            =   3930
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2610
      Width           =   5280
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3930
      MaxLength       =   25
      TabIndex        =   11
      Top             =   3540
      Width           =   3885
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3945
      MaxLength       =   15
      TabIndex        =   9
      Top             =   2070
      Width           =   2790
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   3945
      MaxLength       =   50
      TabIndex        =   13
      Top             =   4365
      Width           =   5280
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   661
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SL.No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   2160
      TabIndex        =   32
      Top             =   630
      Width           =   705
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   6825
      TabIndex        =   31
      Top             =   660
      Width           =   570
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name "
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2160
      TabIndex        =   30
      Top             =   930
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2175
      TabIndex        =   29
      Top             =   1245
      Width           =   1290
   End
   Begin VB.Label Label3 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel - R"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2115
      TabIndex        =   28
      Top             =   1830
      Width           =   630
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel - O"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   5265
      TabIndex        =   27
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label Label5 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   7125
      TabIndex        =   26
      Top             =   1845
      Width           =   570
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   2145
      Top             =   2445
      Width           =   7080
   End
   Begin VB.Label Label6 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Concern"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2145
      TabIndex        =   25
      Top             =   2625
      Width           =   750
   End
   Begin VB.Label Label7 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Received By:"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2145
      TabIndex        =   24
      Top             =   3585
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      Height          =   15
      Left            =   2145
      Top             =   3960
      Width           =   7080
   End
   Begin VB.Label Label8 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Immediate Action Taken"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2130
      TabIndex        =   23
      Top             =   4110
      Width           =   1755
   End
   Begin VB.Shape Shape3 
      Height          =   15
      Left            =   2130
      Top             =   4680
      Width           =   7080
   End
   Begin VB.Label Label9 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Root Cause Analysis :"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2145
      TabIndex        =   22
      Top             =   4830
      Width           =   1755
   End
   Begin VB.Shape Shape4 
      Height          =   15
      Left            =   2145
      Top             =   5385
      Width           =   7080
   End
   Begin VB.Label Label10 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Preventive Action :"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2145
      TabIndex        =   21
      Top             =   5490
      Width           =   1665
   End
   Begin VB.Shape Shape5 
      Height          =   15
      Left            =   2160
      Top             =   6165
      Width           =   7080
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2145
      TabIndex        =   20
      Top             =   6375
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00DAD9CF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Model"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2130
      TabIndex        =   19
      Top             =   2115
      Width           =   1740
   End
End
Attribute VB_Name = "FrmCustConern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Don't Change Tag Property of (Txt) Control as it is used in other activities
'FORM COLOR &H00C0FFFF&
Option Explicit
Public RstMainFormExit As Boolean
'Private Const CtrlBColOrg = &HC2D5B9        'Orginal BackColour
'Private Const CtrlFColOrg = &H80000012      'Orginal ForeColour
'Private Const CtrlBCol = &H80000008         'Changed BackColour
'Private Const CtrlFCol = &H8000000E         'Changed ForeColour
Dim ADDFLAG As Byte
Dim Master As ADODB.Recordset
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstState As ADODB.Recordset, mFlag As Byte
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Const SL_NO = 1, _
        VDate = 2, _
        CUST_NAME = 3, _
        CUST_ADD1 = 4, _
        CUST_ADD2 = 5, _
        Tel_R = 6, _
        TEL_O = 7, _
        Mobile = 8, _
        VehModel = 9, _
        Concern = 10, _
        Received_by = 11, _
        Imm_Act_Taken1 = 12, _
        Imm_Act_Taken2 = 13, _
        Root_Cause_Analysis1 = 14, _
        Root_Cause_Analysis2 = 15, _
        Pre_Action1 = 16, _
        Pre_Action2 = 17, _
        Remarks = 18
        
Private Sub Form_Deactivate()
    If RstMainFormExit = True Then Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift, RstMainFormExit
ELoop:
     CheckError
End Sub
Private Sub Form_Load()
Dim I As Byte
On Error GoTo ELoop
    
'    If rsUserPerm.RecordCount  > 0 Then
'        rsUserPerm.MoveFirst
'        rsUserPerm.FIND ("FORM_NAME='" & Me.CAPTION & "'")
'        If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
'    End If
    WinSetting Me
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select JB.Div_Code+JB.Site_Code+right(space(8)+CStr(JB.Book_no),8) as SearchCode, JB.Div_Code,JB.Site_Code,JB.Book_No from Job_Booking as JB  order by JB.Book_Date desc, JB.Div_Code desc,JB.Site_Code desc,JB.Book_no desc", GCn, adOpenDynamic, adLockOptimistic

    For I = 1 To Txt.Count
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    Set RstMain = New ADODB.Recordset
    RstMain.CursorLocation = adUseClient
    RstMain.Open "Select Ltrim(Rtrim(VEH_CustConcn.Sl_No)) As SearchCode,VEH_CustConcn.* From VEH_CustConcn  Order By VEH_CustConcn.VDate", GCn, adOpenDynamic, adLockOptimistic
    
    Disp_Text SETS("INI", Me, RstMain)
    MoveRec
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" And RstMainFormExit = False Then Form_Unload (-1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set RstMain = Nothing
    Set Master = Nothing
 End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    Disp_Text SETS("ADD", Me, RstMain)
    BlankText
    Txt(SL_NO).TEXT = GCn.Execute("Select IIF(ISNULL(MAX(Sl_No)),1,MAX(SL_NO)+1) AS AA From VEH_CustConcn").Fields(0).Value
    Txt(VDate).TEXT = PubLoginDate
    Txt(VDate).SetFocus
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM, J As Byte, TmpSQL As String, mTran As Boolean
On Error GoTo ELoop
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        mTran = True
        GCn.Execute ("Delete From Veh_CustConcn Where Sl_no=" & Trim(Txt(SL_NO)))
        GCn.CommitTrans
        mTran = False
        RstMain.Requery
        If RstMain.EOF = False Then RstMain.MoveLast
        MoveRec
        BUTTONS True, Me, RstMain, 0
    End If
    Exit Sub
ELoop:
    If mTran = True Then GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message", App.Title
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records to Edit.", vbInformation, "Information": Exit Sub
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(SL_NO).Enabled = False
    Txt(VDate).SetFocus
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT " & _
                "Ltrim(Rtrim(VEH_CustConcn.Sl_No)) As SearchCode, VEH_CustConcn.VDate,VEH_CustConcn.Cust_Name, VEH_CustConcn.Cust_add1, VEH_CustConcn.Cust_add2, VEH_CustConcn.Tel_R, VEH_CustConcn.Tel_O, VEH_CustConcn.Mobile, VEH_CustConcn.Concern, VEH_CustConcn.VehModel, VEH_CustConcn.Remarks " & _
            "FROM VEH_CustConcn " & _
            "Order By  VEH_CustConcn.Sl_No, VEH_CustConcn.VDate"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ELoop:
    CheckError
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    RstMain.MoveFirst
    RstMain.FIND ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, RstMain, 1
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, RstMain, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, RstMain, 3
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, RstMain, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eCancel()
Dim I As Byte
On Error GoTo ELoop
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, RstMain)
        For I = 1 To Txt.Count
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
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
    'If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else
    CmdPrint(PWindows).SetFocus
End Sub

Private Sub TopCtrl1_eRef()
    RstMain.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, TmpSQL$, MaxCode%, I%, MyOldNo$
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        TmpSQL = "Insert Into Veh_CustConcn (sl_no,Vdate,Cust_name,Cust_Add1,Cust_add2,Tel_R,Tel_O,Mobile,VehModel,Concern,Received_by,Imm_Act_Taken1,Imm_Act_Taken2,Root_Cause_Analysis1,Root_Cause_Analysis2,Pre_Action1,Pre_Action2,Remarks,U_Name,U_EntDt,U_AE) Values ('" & Txt(SL_NO) & "'," & ConvertDate(Txt(VDate)) & ",'" & Txt(CUST_NAME) & "','" & Txt(CUST_ADD1) & "','" & Txt(CUST_ADD2) & "','" & Txt(Tel_R) & "','" & Txt(TEL_O) & "','" & Txt(Mobile) & "','" & Txt(VehModel) & "','" & Txt(Concern) & "','" & Txt(Received_by) & "','" & Txt(Imm_Act_Taken1) & "','" & Txt(Imm_Act_Taken2) & "','" & Txt(Root_Cause_Analysis1) & "','" & Txt(Root_Cause_Analysis2) & "','" & Txt(Pre_Action1) & "','" & Txt(Pre_Action2) & "','" & Txt(Remarks) & "','" & pubUName & "',#" & PubServerDate & "#,'" & IIf(ADDFLAG = 1, "A", "E") & "')"
    Else
        TmpSQL = "Update Veh_CustConcn Set SL_No='" & Txt(SL_NO) & "',vdate=" & ConvertDate(Txt(VDate)) & ",Cust_Name='" & Txt(CUST_NAME) & "',Cust_Add1='" & Txt(CUST_ADD1) & "',Cust_Add2='" & Txt(CUST_ADD2) & "',Tel_R='" & Txt(Tel_R) & "',Tel_O='" & Txt(TEL_O) & "',Mobile='" & Txt(Mobile) & "',VehModel='" & Txt(VehModel) & "',Received_by='" & Txt(Received_by) & "',Concern='" & Txt(Concern) & "',Imm_Act_Taken1='" & Txt(Imm_Act_Taken1) & "',Imm_Act_Taken2='" & Txt(Imm_Act_Taken2) & "',Root_Cause_AnaLysis1='" & Txt(Root_Cause_Analysis1) & "',Root_Cause_AnaLysis2='" & Txt(Root_Cause_Analysis2) & "',Pre_Action1='" & Txt(Pre_Action1) & "',Pre_Action2='" & Txt(Pre_Action2) & "',Remarks='" & Txt(Remarks) & "',U_Name='" & pubUName & "',U_EntDt=#" & PubServerDate & "#,U_AE='E'" & " Where Sl_No=" & Txt(SL_NO).TEXT
    End If
    GCn.BeginTrans
        mTrans = True
        GCn.Execute (TmpSQL)
    GCn.CommitTrans
    RstMain.Requery
    TopCtrl1_ePrn
    RstMain.FIND ("SearchCode='" & Txt(SL_NO).TEXT & "'")
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, RstMain)
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
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Remarks Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Remarks Then
         If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SL_NO Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> VDate Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbCritical, App.Title: Exit Sub
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    CheckQuote KeyAscii
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub
'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 1 To Txt.Count
        Txt(I).TEXT = ""
        Txt(I).Tag = ""
    Next I
End Sub
Private Sub MoveRec()
Dim Rst As New ADODB.Recordset, I As Integer
On Error GoTo ELoop
    If RstMain.RecordCount > 0 Then
        Txt(SL_NO).TEXT = RstMain!SL_NO
        Txt(VDate).TEXT = RstMain!VDate
        Txt(CUST_NAME).TEXT = RstMain!CUST_NAME
        Txt(CUST_ADD1).TEXT = RstMain!CUST_ADD1
        Txt(CUST_ADD2).TEXT = RstMain!CUST_ADD2
        Txt(Tel_R).TEXT = RstMain!Tel_R
        Txt(TEL_O).TEXT = RstMain!TEL_O
        Txt(Mobile).TEXT = RstMain!Mobile
        Txt(VehModel).TEXT = RstMain!VehModel
        Txt(Concern).TEXT = RstMain!Concern
        Txt(Received_by).TEXT = RstMain!Received_by
        Txt(Imm_Act_Taken1).TEXT = RstMain!Imm_Act_Taken1
        Txt(Imm_Act_Taken2).TEXT = RstMain!Imm_Act_Taken2
        Txt(Root_Cause_Analysis1).TEXT = RstMain!Root_Cause_Analysis1
        Txt(Root_Cause_Analysis2).TEXT = RstMain!Root_Cause_Analysis2
        Txt(Pre_Action1).TEXT = RstMain!Pre_Action1
        Txt(Pre_Action2).TEXT = RstMain!Pre_Action2
        Txt(Remarks).TEXT = RstMain!Remarks
    Else
        BlankText
    End If
    Grid_Hide
    Set Rst = Nothing
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 1 To Txt.Count
        Txt(I).Enabled = Enb
    Next
End Sub
Private Sub Grid_Hide()
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
        mRepName = IIf(OptPlain.Value = True, "CustConcn", "CustConcn")
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "CustConcn", "CustConcn")
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
Dim RST1 As ADODB.Recordset
Dim Speciality$
Dim Rst As ADODB.Recordset
Dim I As Integer

On Error GoTo ERRORHANDLER
mQRY = "SELECT STR(YEAR(CC.VDATE))& '/'& CC.Sl_No AS SLNO,CC.Vdate,CC.Cust_Name, CC.Cust_Add1, CC.Cust_Add2, CC.Tel_R, CC.Tel_O, CC.Mobile, CC.VehModel, CC.Concern, CC.Received_By, CC.Imm_Act_Taken1, CC.Imm_Act_Taken2,CC.Root_Cause_Analysis1, CC.Root_Cause_Analysis2,CC.Pre_Action1,CC.Pre_Action2,CC.Remarks " & _
        "FROM Veh_CustConcn CC WHERE CC.SL_NO = " & Txt(SL_NO).TEXT
Set Rst = New Recordset
Rst.CursorLocation = adUseClient
Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic

If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
 
CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")

Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select Div_SName,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    mDocStr = "** Cutomer Concern " & IIf(RST1!Div_SName = "", "", " (" & RST1!Div_SName & ")") & "**"

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
        Call Report_View(rpt, "Customer Concern", , True)
End Select
CmdPrint(PSetUp).Tag = ""
Set Rst = Nothing
Set RST1 = Nothing
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

Private Sub SpeedPrint()

End Sub


