VERSION 5.00
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form FrmFinWrkSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finance Work Sheet"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10230
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   6870
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4005
      Width           =   2160
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   6870
      MaxLength       =   5
      TabIndex        =   8
      Top             =   4545
      Width           =   1065
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   6870
      LinkTimeout     =   10
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3465
      Width           =   2145
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4740
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1410
      Width           =   4305
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   6900
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1890
      Width           =   2130
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6885
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2430
      Width           =   2145
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   6870
      LinkTimeout     =   10
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2955
      Width           =   2145
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   4710
      MaxLength       =   10
      TabIndex        =   7
      Top             =   4545
      Width           =   1860
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
      Left            =   2805
      TabIndex        =   9
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "FrmFinWrkSheet.frx":0000
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
         TabIndex        =   14
         ToolTipText     =   "Printer "
         Top             =   285
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "FrmFinWrkSheet.frx":030A
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
         TabIndex        =   13
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "FrmFinWrkSheet.frx":0614
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
         TabIndex        =   12
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
         Picture         =   "FrmFinWrkSheet.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "FrmFinWrkSheet.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   10
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   661
   End
   Begin VB.Shape Shape5 
      Height          =   4005
      Left            =   1710
      Top             =   1275
      Width           =   2625
   End
   Begin VB.Shape Shape4 
      Height          =   4005
      Left            =   1710
      Top             =   1275
      Width           =   7605
   End
   Begin VB.Label Label20 
      Caption         =   "Rs"
      Height          =   300
      Left            =   4395
      TabIndex        =   44
      Top             =   3045
      Width           =   420
   End
   Begin VB.Label Label15 
      Caption         =   "Rs"
      Height          =   270
      Index           =   1
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Rs"
      Height          =   270
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Balance "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1905
      TabIndex        =   41
      Top             =   4530
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Initial Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1890
      TabIndex        =   40
      Top             =   3975
      Width           =   1830
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Installment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   1920
      TabIndex        =   39
      Top             =   3480
      Width           =   1830
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Margin Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1860
      TabIndex        =   38
      Top             =   2985
      Width           =   1980
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Finance Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1890
      TabIndex        =   37
      Top             =   2400
      Width           =   2370
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Invoice Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1905
      TabIndex        =   36
      Top             =   1950
      Width           =   2205
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   1905
      TabIndex        =   35
      Top             =   1485
      Width           =   810
   End
   Begin VB.Label Label3 
      Caption         =   "Months"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8130
      TabIndex        =   34
      Top             =   4545
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   " (A+B)"
      Height          =   285
      Left            =   3690
      TabIndex        =   33
      Top             =   4035
      Width           =   585
   End
   Begin VB.Label Label5 
      Caption         =   "    ( B)"
      Height          =   255
      Left            =   3600
      TabIndex        =   32
      Top             =   3555
      Width           =   675
   End
   Begin VB.Label Label6 
      Caption         =   "(A)"
      Height          =   255
      Left            =   3885
      TabIndex        =   31
      Top             =   3030
      Width           =   600
   End
   Begin VB.Label Label12 
      Caption         =   "Rs"
      Height          =   315
      Left            =   4425
      TabIndex        =   30
      Top             =   1995
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Rs"
      Height          =   330
      Left            =   4395
      TabIndex        =   29
      Top             =   2490
      Width           =   1005
   End
   Begin VB.Label Label14 
      Caption         =   "Rs"
      Height          =   315
      Left            =   4260
      TabIndex        =   28
      Top             =   3075
      Width           =   780
   End
   Begin VB.Label Label15 
      Caption         =   "Rs"
      Height          =   270
      Index           =   0
      Left            =   4380
      TabIndex        =   27
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "Rs"
      Height          =   285
      Left            =   4395
      TabIndex        =   26
      Top             =   4035
      Width           =   945
   End
   Begin VB.Label Label17 
      Caption         =   "Rs"
      Height          =   225
      Left            =   4395
      TabIndex        =   25
      Top             =   4635
      Width           =   330
   End
   Begin VB.Label Label18 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6615
      TabIndex        =   24
      Top             =   4575
      Width           =   225
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   7740
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape3 
      Height          =   15
      Left            =   0
      Top             =   0
      Width           =   7080
   End
   Begin VB.Shape Shape2 
      Height          =   15
      Left            =   0
      Top             =   0
      Width           =   7080
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   0
      Top             =   0
      Width           =   7080
   End
   Begin VB.Label Label9 
      BackColor       =   &H00DAD9CF&
      Caption         =   "Initial Payment "
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   1755
   End
End
Attribute VB_Name = "FrmFinWrkSheet"
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

Private Const FinModel = 1, _
        Inv_Amt = 2, _
        Fin_Amt = 3, _
        Margin_Amt = 4, _
        Installment = 5, _
        Initial_Pmt = 6, _
        BalanceInstl = 7, _
        BalanceMth = 8
        
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
    RstMain.Open "select JB.Div_Code+JB.Site_Code+right(space(8)+CStr(JB.Book_no),8) as SearchCode, JB.Div_Code,JB.Site_Code,JB.Book_No from Job_Booking as JB  order by JB.Book_Date desc, JB.Div_Code desc,JB.Site_Code desc,JB.Book_no desc", GCn, adOpenDynamic, adLockOptimistic
    
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
    Txt(FinModel).SetFocus
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM, J As Byte, TmpSQL As String, mTran As Boolean
On Error GoTo ELoop
'    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
'        GCn.BeginTrans
'        mTran = True
'        GCn.Execute ("Delete From Veh_CustConcn Where Sl_no=" & Trim(TXT(SL_NO)))
'        GCn.CommitTrans
'        mTran = False
'        RstMain.Requery
'        If RstMain.EOF = False Then RstMain.MoveLast
'        MoveRec
'        BUTTONS True, Me, RstMain, 0
'    End If
    Exit Sub
ELoop:
    If mTran = True Then GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message", App.Title
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records to Edit.", vbInformation, "Information": Exit Sub
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(FinModel).SetFocus
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
'    If TopCtrl1.TopText2.CAPTION = "Add" Then
'        TmpSQL = "Insert Into Veh_CustConcn (sl_no,Vdate,Cust_name,Cust_Add1,Cust_add2,Tel_R,Tel_O,Mobile,VehModel,Concern,Received_by,Imm_Act_Taken1,Imm_Act_Taken2,Root_Cause_Analysis1,Root_Cause_Analysis2,Pre_Action1,Pre_Action2,Remarks,U_Name,U_EntDt,U_AE) Values ('" & Txt(SL_NO) & "'," & ConvertDate(Txt(Vdate)) & ",'" & Txt(CUST_NAME) & "','" & Txt(CUST_ADD1) & "','" & Txt(CUST_ADD2) & "','" & Txt(Tel_R) & "','" & Txt(TEL_O) & "','" & Txt(Mobile) & "','" & Txt(VehModel) & "','" & Txt(Concern) & "','" & Txt(Received_by) & "','" & Txt(Imm_Act_Taken1) & "','" & Txt(Imm_Act_Taken2) & "','" & Txt(Root_Cause_Analysis1) & "','" & Txt(Root_Cause_Analysis2) & "','" & Txt(Pre_Action1) & "','" & Txt(Pre_Action2) & "','" & Txt(Remarks) & "','" & pubUName & "',#" & PubServerDate & "#,'" & IIf(ADDFLAG = 1, "A", "E") & "')"
'    Else
'        TmpSQL = "Update Veh_CustConcn Set SL_No='" & Txt(SL_NO) & "',vdate=" & ConvertDate(Txt(Vdate)) & ",Cust_Name='" & Txt(CUST_NAME) & "',Cust_Add1='" & Txt(CUST_ADD1) & "',Cust_Add2='" & Txt(CUST_ADD2) & "',Tel_R='" & Txt(Tel_R) & "',Tel_O='" & Txt(TEL_O) & "',Mobile='" & Txt(Mobile) & "',VehModel='" & Txt(VehModel) & "',Received_by='" & Txt(Received_by) & "',Concern='" & Txt(Concern) & "',Imm_Act_Taken1='" & Txt(Imm_Act_Taken1) & "',Imm_Act_Taken2='" & Txt(Imm_Act_Taken2) & "',Root_Cause_AnaLysis1='" & Txt(Root_Cause_Analysis1) & "',Root_Cause_AnaLysis2='" & Txt(Root_Cause_Analysis2) & "',Pre_Action1='" & Txt(Pre_Action1) & "',Pre_Action2='" & Txt(Pre_Action2) & "',Remarks='" & Txt(Remarks) & "',U_Name='" & pubUName & "',U_EntDt=#" & PubServerDate & "#,U_AE='E'" & " Where Sl_No=" & Txt(SL_NO).TEXT
'    End If
'    GCn.BeginTrans
'        mTrans = True
'        GCn.Execute (TmpSQL)
'    GCn.CommitTrans
'    RstMain.Requery
    TopCtrl1_ePrn
'    RstMain.FIND ("SearchCode='" & Txt(SL_NO).TEXT & "'")
  '  If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
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
    If KeyCode = vbKeyReturn Then
        Select Case Index
    
        Case BalanceMth
                If Txt(Index) = " " Or Val(Txt(Index)) = 0 Then
                      Txt(Index).SetFocus
                      Exit Sub
                End If
                If (Val(Txt(Inv_Amt)) - Val(Txt(Initial_Pmt))) <> (Val(Txt(BalanceInstl)) * Val(Txt(BalanceMth))) Then
                     MsgBox ("Please check the Installment Amt and Months")
                     Txt(Index).SetFocus
                     Exit Sub
                End If
        End Select
    End If
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> BalanceMth Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = BalanceMth Then
        
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> FinModel Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> FinModel Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbCritical, App.Title: Exit Sub
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    CheckQuote KeyAscii
    Select Case Index
        Case Inv_Amt
            NumPress Txt(Index), KeyAscii, 10, 2
        Case Fin_Amt
            NumPress Txt(Index), KeyAscii, 10, 2
        Case Margin_Amt
            NumPress Txt(Index), KeyAscii, 10, 2
        Case Installment
            NumPress Txt(Index), KeyAscii, 10, 2
        Case Initial_Pmt
            NumPress Txt(Index), KeyAscii, 10, 2
        Case BalanceInstl
            NumPress Txt(Index), KeyAscii, 10, 2
        Case BalanceMth
            NumPress Txt(Index), KeyAscii, 4, 0
    End Select
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
Private Sub Cmdprint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "FinWrkSheet", "FinWrkSheet")
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "FinWrkSheet", "FinWrkSheet")
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
Dim I As Integer

On Error GoTo ERRORHANDLER
mQRY = "select '" & Txt(FinModel) & "' as Fin_Model," & Txt(Inv_Amt) & " as Inv_Amt," & Txt(Fin_Amt) & " as Fin_Amt," & Txt(Margin_Amt) & " as Margin," & Txt(Installment) & " as installment," & Txt(Initial_Pmt) & " as Initial_pmt, " & Txt(BalanceInstl) & " as BalanceInstl," & Txt(BalanceMth) & " as BalanceMth from veh_custconcn"

Set Rst = New Recordset
Rst.CursorLocation = adUseClient
Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic

If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
 
CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")

Set Rst1 = New Recordset
Rst1.CursorLocation = adUseClient
Rst1.Open "select Div_SName,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    mDocStr = "** Finance Work Sheet " & IIf(Rst1!Div_SName = "", "", " (" & Rst1!Div_SName & ")") & "**"

For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SubTitle")
            rpt.FormulaFields(I).TEXT = "'" & Speciality & "'"
        Case UCase("Comp_Phone")
            rpt.FormulaFields(I).TEXT = "'" & Rst1!W_SecPhone & "'"
        Case UCase("Comp_Fax")
            rpt.FormulaFields(I).TEXT = "'" & Rst1!W_SecFax & "'"
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
        Call Report_View(rpt, "Finance Work Sheet", , True)
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

Private Sub SpeedPrint()

End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
       Case Inv_Amt
            If Txt(Index) = " " Or Val(Txt(Index)) = 0 Then
                 Cancel = True
                 Exit Sub
            End If
       Case Fin_Amt
           If Txt(Index) <> " " Or Val(Txt(Index)) <> 0 Then
                Txt(Margin_Amt).TEXT = Val(Txt(Inv_Amt)) - Val(Txt(Fin_Amt))
           End If
       Case Installment
                If Txt(Margin_Amt).TEXT <> Txt(Inv_Amt).TEXT Then
                     If Txt(Index) = " " Or Val(Txt(Index)) = 0 Then
                         Cancel = True
                         Exit Sub
                     End If
                End If
                If Txt(Index) <> " " Or Val(Txt(Index)) <> 0 Then
                     Txt(Initial_Pmt) = Val(Txt(Margin_Amt)) + Val(Txt(Installment))
                     Txt(BalanceInstl) = Val(Txt(Installment))
                     Txt(BalanceMth) = (Txt(Inv_Amt) - Txt(Initial_Pmt)) / Val(Txt(Installment))
                End If
       
      End Select
End Sub

