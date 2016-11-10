VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FaMagic 
   BackColor       =   &H00CBBE9E&
   Caption         =   "Magic"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "FaMagic.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11130
   Begin VB.TextBox TxtSearch1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   8595
      TabIndex        =   35
      Top             =   1170
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00B7A4F0&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   4000
      TabIndex        =   31
      Top             =   6400
      Width           =   4560
      Begin VB.CommandButton BTNSITEOK 
         BackColor       =   &H8000000A&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Refresh"
         Top             =   1650
         Width           =   4560
      End
      Begin VB.CheckBox Check 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   32
         Top             =   45
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
         Height          =   1650
         Index           =   0
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   2910
         _Version        =   393216
         BackColor       =   15595518
         ForeColor       =   128
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8388608
         ForeColorFixed  =   65535
         BackColorSel    =   12632256
         ForeColorSel    =   128
         BackColorBkg    =   14347757
         GridColor       =   8438015
         GridColorFixed  =   192
         GridColorUnpopulated=   16711935
         Appearance      =   0
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D9ACCB&
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   4185
      TabIndex        =   8
      Top             =   2550
      Width           =   4290
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D9ACCB&
         Caption         =   "Show Only Debit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   15
         Top             =   1410
         Width           =   2100
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D9ACCB&
         Caption         =   "Show Only Credit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   16
         Top             =   1635
         Width           =   2100
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D9ACCB&
         Caption         =   "Group Wise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   14
         Top             =   1185
         Width           =   2100
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D9ACCB&
         Caption         =   "Show Zero Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   13
         Top             =   960
         Width           =   2100
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D9ACCB&
         Caption         =   "Detailed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   12
         Top             =   735
         Width           =   2100
      End
      Begin VB.CommandButton BtnOK 
         BackColor       =   &H8000000A&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3045
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Refresh"
         Top             =   1275
         Width           =   1185
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D9ACCB&
         Caption         =   "Op.Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   540
         Value           =   1  'Checked
         Width           =   2100
      End
      Begin VB.TextBox TXTE_DATE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2010
         TabIndex        =   10
         Top             =   270
         Width           =   1395
      End
      Begin VB.TextBox TXTS_DATE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2010
         TabIndex        =   9
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   19
         Top             =   30
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   18
         Top             =   270
         Width           =   555
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   1695
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7329
      _Version        =   393216
      BackColor       =   14875388
      ForeColor       =   0
      FixedCols       =   0
      BackColorFixed  =   13741296
      BackColorSel    =   16711680
      BackColorBkg    =   14875388
      GridColor       =   0
      GridColorFixed  =   0
      GridColorUnpopulated=   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   0
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      HideSelection   =   0   'False
      Left            =   7605
      TabIndex        =   22
      Top             =   1770
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10230
      Begin VB.CommandButton BtnSite 
         Caption         =   "C&hange Sites"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Change Site"
         Top             =   30
         Width           =   1725
      End
      Begin VB.CommandButton BtnPrint 
         Caption         =   "&WinPrint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton BtnPrint 
         Caption         =   "P&review"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1155
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton BTNEXIT 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton BTNRefresh 
         BackColor       =   &H8000000A&
         Caption         =   "Para&meters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Parameters"
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton BtnPrint 
         Caption         =   "&DosPrint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   3465
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Trial-Group (F5)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   2
         Left            =   2835
         TabIndex        =   29
         Tag             =   "vbKeyF6"
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Trial-Ledger (F6)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   3
         Left            =   4275
         TabIndex        =   28
         Tag             =   "vbKeyF7"
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Profit/Loss (F4)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   1
         Left            =   1530
         TabIndex        =   27
         Tag             =   "vbKeyF9"
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance sheet (F3)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   0
         Left            =   -15
         TabIndex        =   26
         Tag             =   "vbKeyF6"
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CashFlow (F7)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   4
         Left            =   5760
         TabIndex        =   25
         Tag             =   "vbKeyF7"
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FundFlow (F8)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   5
         Left            =   7065
         TabIndex        =   24
         Tag             =   "vbKeyF7"
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash/Bank Books (F9)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   6
         Left            =   8355
         TabIndex        =   23
         Tag             =   "vbKeyF7"
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   975
      Width           =   7110
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      BorderStyle     =   0  'None
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
      Height          =   240
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1215
      Width           =   7110
   End
   Begin MSFlexGridLib.MSFlexGrid FGrid2 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6060
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   450
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14875388
      ForeColor       =   16512
      BackColorFixed  =   10281447
      BackColorBkg    =   14875388
      GridColor       =   14875388
      GridColorFixed  =   14875388
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FaMagic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MagicStack As adodb.Recordset, X11, RstEnviro As adodb.Recordset, mREFRESH As Boolean
Dim RsGrid1 As adodb.Recordset
Dim GridRow1() As Integer, mGridStartRow As Integer, mGridEndRow As Integer
Private PubDatamanFa As New DMFa.ClsFa, SiteCodeStore As String

''''''''''''''''
'''''Private Const PubShowSiteWiseReport As Boolean = False
''''''''''''''''

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
Dim RST1 As adodb.Recordset
Set RST1 = FGrid1.DataSource
FaSelGridKeyPress TxtSearch, FGrid1, RST1, KeyAscii, RST1.Fields(FGrid1.Col).Name, FaCellBackColEnter1, FaCellBackColLeave1: KeyAscii = 0
Set RST1 = Nothing
End Sub
Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TxtSearch = ""
    If KeyCode = vbKeyReturn Then FGrid1.SetFocus
End Sub
Private Sub ColoChange(Index As Integer)
Dim I As Integer
For I = 0 To 6
    LblShort(I).FontBold = False
    LblShort(I).ForeColor = &HFFFF&
Next
LblShort(Index).ForeColor = &HFFFF00
LblShort(Index).FontBold = True
End Sub
Private Sub MagLedger(SubCode As String, Optional FRow As Integer, Optional Fcol As Integer, Optional magStartDate As Date, Optional magEndDate As Date)
Dim RstLedger As adodb.Recordset
Set RstLedger = PubDatamanFa.FaMagLedger(Me, SubCode, FRow, Fcol, magStartDate, magEndDate)
If RstLedger.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstLedger, PubFaReportPath + "\FaMagLedger.ttx", True)
Set RstLedger = Nothing
End Sub
Private Sub MagCashBook(SubCode As String, Optional FRow As Integer, Optional Fcol As Integer, Optional magStartDate As Date, Optional magEndDate As Date)
Dim RstLedger As adodb.Recordset
Set RstLedger = PubDatamanFa.FaMagCashBook(Me, SubCode, FRow, Fcol, magStartDate, magEndDate)
If RstLedger.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstLedger, PubFaReportPath + "\FaMagCashBook.ttx", True)
Set RstLedger = Nothing
End Sub
Private Sub monthSum(LedCode As String, Optional FRow As Integer, Optional Fcol As Integer)
Dim RstMonthSum As adodb.Recordset
Set RstMonthSum = PubDatamanFa.FaMonthSum(Me, LedCode, FRow, Fcol)
If RstMonthSum.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstMonthSum, PubFaReportPath + "\FaMonthSum.ttx", True)
Set RstMonthSum = Nothing
End Sub
Private Sub Subtrial(GroupCode As String, Optional FRow As Integer, Optional Fcol As Integer)
Dim RstSubTrial As adodb.Recordset
Set RstSubTrial = New adodb.Recordset
Set RstSubTrial = PubDatamanFa.FaSubtrial(Me, GroupCode, FRow, Fcol)
If RstSubTrial.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstSubTrial, PubFaReportPath + "\FaSubTrial.ttx", True)
Set RstSubTrial = Nothing
End Sub
Private Sub BalSheet(Optional FRow As Integer, Optional Fcol As Integer)
Dim RstBalSheet As adodb.Recordset
Set RstBalSheet = PubDatamanFa.FaBalSheet(Me, FRow, Fcol)
ColoChange 0
If RstBalSheet.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstBalSheet, PubFaReportPath + "\FaBalSheet.ttx", True)
Set RstBalSheet = Nothing
End Sub
Private Sub ProfLoss(Optional FRow As Integer, Optional Fcol As Integer)
Dim RstProfLoss As adodb.Recordset
Set RstProfLoss = PubDatamanFa.FaProfLoss(Me, FRow, Fcol)
ColoChange 1
If RstProfLoss.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstProfLoss, PubFaReportPath + "\FaProfLoss.ttx", True)
Set RstProfLoss = Nothing
End Sub
Private Sub GroupTrial(Optional FRow As Integer, Optional Fcol As Integer)
Dim RstGroupTrial As adodb.Recordset
'Set RstGroupTrial = PubDatamanFa.FaGroupTrial(Me, FRow, Fcol)
Set RstGroupTrial = Module1.GroupTrial(Me, FRow, Fcol)
ColoChange 2
If RstGroupTrial.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstGroupTrial, PubFaReportPath + "\FaGroupTrial.ttx", True)
Set RstGroupTrial = Nothing
End Sub
Private Sub LedTrial(Optional FRow As Integer, Optional Fcol As Integer)
Dim RstLedTrial As adodb.Recordset
'Set RstLedTrial = PubDatamanFa.FaLedTrial(Me, FRow, Fcol)
Set RstLedTrial = Module1.LedTrial(Me, FRow, Fcol)
ColoChange 3
If RstLedTrial.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstLedTrial, PubFaReportPath + "\FaLedTrial.ttx", True)
Set RstLedTrial = Nothing
End Sub
Private Sub CashFlow(Optional FRow As Integer, Optional Fcol As Integer)
Dim RstCashFundFlow As adodb.Recordset
Set RstCashFundFlow = PubDatamanFa.FaCashFlow(Me, FRow, Fcol)
ColoChange 4
If RstCashFundFlow.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstCashFundFlow, PubFaReportPath + "\FaCashFundFlow.ttx", True)
Set RstCashFundFlow = Nothing
End Sub
Private Sub FundFlow(Optional FRow As Integer, Optional Fcol As Integer)
Dim RstCashFundFlow As adodb.Recordset
Set RstCashFundFlow = PubDatamanFa.FaFundFlow(Me, FRow, Fcol)
ColoChange 5
If RstCashFundFlow.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstCashFundFlow, PubFaReportPath + "\FaCashFundFlow.ttx", True)
Set RstCashFundFlow = Nothing
End Sub
Private Sub CashBankSum(Optional FRow As Integer, Optional Fcol As Integer)
Dim RstCashBankSum As adodb.Recordset
Set RstCashBankSum = PubDatamanFa.FaCashBankSum(Me, FRow, Fcol)
ColoChange 6
If RstCashBankSum.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstCashBankSum, PubFaReportPath + "\FaCashBankSum.ttx", True)
Set RstCashBankSum = Nothing
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub Check4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub Check5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub Check6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub FGrid1_LeaveCell()
    FGrid1.CellBackColor = &HE2FAFC
End Sub
Private Sub Form_Deactivate()
    mREFRESH = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If TypeOf Me.ActiveControl Is TextBox Or TypeOf Me.ActiveControl Is CheckBox Then
    If Frame1.Visible = False And TxtSearch.Visible = False Then If KeyCode = 13 Then SendKeysA vbKeyTab, True
End If
Select Case KeyCode
    Case 114
        LblShort_Click 0
    Case 115
        LblShort_Click 1
    Case 116
        LblShort_Click 2
    Case 117
        LblShort_Click 3
    Case 118
        LblShort_Click 4
    Case 119
        LblShort_Click 5
    Case 120
        LblShort_Click 6
    Case 121
        btnexit_Click 0
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MagicStack = Nothing
    Set PubDatamanFa = Nothing
    Set RstEnviro = Nothing
    Set RsGrid1 = Nothing
End Sub
Private Sub TXTE_DATE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub TXTE_DATE_Validate(Cancel As Boolean)
    TXTE_DATE = RetDate(TXTE_DATE)
End Sub
Private Sub TXTS_DATE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub TXTS_DATE_Validate(Cancel As Boolean)
    TXTS_DATE = RetDate(TXTS_DATE)
End Sub
Private Sub btnexit_Click(Index As Integer)
    Unload Me
End Sub
Private Sub BTNRefresh_Click()
    If Frame1.Visible = True Or Frame3.Visible = True Then Exit Sub
    Frame1.Visible = True
    Frame1.ZOrder 0
    TXTS_DATE.SetFocus
End Sub
Private Sub Form_Load()
    Me.left = 0
    Me.top = 0
    Me.height = 7395
    Me.width = 11775
    TXTS_DATE = PubStartDate
    TXTE_DATE = PubLoginDate
    Text1 = ""
    Frame2.left = 0
    Frame2.top = 0
    Text2.top = Frame2.top + Frame2.height
    Text1.top = Text2.top + Text2.height
    FGrid1.left = 0
    FGrid1.top = Text1.top + Text1.height + 50
    FGrid1.width = 11650
    FGrid1.height = 5400
    FGrid2.left = 0
    FGrid2.top = FGrid1.top + FGrid1.height
    FGrid2.width = 11650
    Frame2.width = FGrid2.width
    MousePointer = vbHourglass
    MousePointer = vbDefault
    Frame1.Visible = False
    Frame1.left = 1500
    Frame1.top = 500
    Frame3.Visible = False
    Frame3.left = 1500
    Frame3.top = 500
    Frame3.Visible = False
    BtnSite.top = BTNRefresh.top
    BtnSite.left = Frame2.left + Frame2.width - BtnSite.width - 100
    ''''''''''''''''''''''
    PubDatamanFa.FaBackEnd = PubBackEnd
    PubDatamanFa.FaPubLoginDate = PubLoginDate
    PubDatamanFa.FaPubDivCode = PubDivCode
    PubDatamanFa.FaPubSiteCode = PubSiteCode
    PubDatamanFa.FapubUName = pubUName
    PubDatamanFa.FaPubSiteCodeDisplay = PubSiteCodeDisplay
    PubDatamanFa.FaPubSiteName = PubSiteName
    PubDatamanFa.FaRunPIF = PubRunPIF
    PubDatamanFa.FaDosPort = PubFaDosPort
    PubDatamanFa.FaPubSiteType = PubFaSiteType
    Set PubDatamanFa.SetG_FaCn = G_FaCn
    Set PubDatamanFa.SetG_CompCn = G_CompCn
    Set PubDatamanFa.SetrsUserPerm = rsUserPerm.Clone
    Set PubDatamanFa.SetMasterRst = FaMasterRst.Clone
    SiteCodeStore = PubSiteCode
    ''''''''''''''''''''''
    If PubSiteWiseDisplayYn = 1 Then
            If PubFaSiteType <> 0 Then
                Set RsGrid1 = New adodb.Recordset
                RsGrid1.CursorLocation = adUseClient
                RsGrid1.Open "SELECT '' as O,SITE_DESC AS Site,SITE_CODE FROM SITE ORDER BY SITE_dESC", G_FaCn, adOpenStatic, adLockReadOnly
                Set GridSel(0).DataSource = RsGrid1
                ReDim Preserve GridRow1(0)
                GridSel(0).width = 5200: GridSel(0).ColWidth(0) = 1000: GridSel(0).ColWidth(2) = 0: GridSel(0).ColWidth(1) = 3000
                GridRow1(0) = 0
                Check(0).top = GridSel(0).top + 20: Check(0).left = GridSel(0).left + 40
            Else
            
                Set RsGrid1 = New adodb.Recordset
                RsGrid1.CursorLocation = adUseClient
                RsGrid1.Open "SELECT '' as O,SITE_DESC AS Site,SITE_CODE FROM SITE where SITE_CODE='" & PubSiteCode & "' ORDER BY SITE_dESC", G_FaCn, adOpenStatic, adLockReadOnly
                Set GridSel(0).DataSource = RsGrid1
                ReDim Preserve GridRow1(0)
                GridSel(0).width = 5200: GridSel(0).ColWidth(0) = 1000: GridSel(0).ColWidth(2) = 0: GridSel(0).ColWidth(1) = 3000
                GridRow1(0) = 0
                Check(0).top = GridSel(0).top + 20: Check(0).left = GridSel(0).left + 40
                Dim I As Integer, GridRow As Integer
                    GridRow1(0) = GridSel(0).Row
                     GridSel(0).TextMatrix(1, 0) = "ü"
                     ' ac_str = ac_str + IIf(ac_str = "", GridSel(0).TextMatrix(1, 1), "," + GridSel(0).TextMatrix(1, 1))
                    
        
            SiteAsign
            
            
            End If
Else
            If PubFaSiteType <> 0 Then
                    Set RsGrid1 = New adodb.Recordset
                    RsGrid1.CursorLocation = adUseClient
                    RsGrid1.Open "SELECT '' as O,SITE_DESC AS Site,SITE_CODE FROM SITE ORDER BY SITE_dESC", G_FaCn, adOpenStatic, adLockReadOnly
                    Set GridSel(0).DataSource = RsGrid1
                    ReDim Preserve GridRow1(0)
                    GridSel(0).width = 5200: GridSel(0).ColWidth(0) = 1000: GridSel(0).ColWidth(2) = 0: GridSel(0).ColWidth(1) = 3000
                    GridRow1(0) = 0
                    Check(0).top = GridSel(0).top + 20: Check(0).left = GridSel(0).left + 40
                Else
                
                
                
                End If
End If
    Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
    If RstEnviro.RecordCount <= 0 Then MsgBox "Parameter Not Set": Exit Sub
    FaClosingBalanceCalculation
'''''    If PubULabel <> 1 Then
'''''        LblShort(0).Enabled = False
'''''        LblShort(1).Enabled = False
'''''    End If
End Sub
Private Sub Form_Activate()
Dim mROW As Integer, mCol As Integer
If FGrid1.Tag = "VOUCHER" Then
    If MagicStack.RecordCount > 1 Then
        MagicStack.MoveLast
        mROW = FaVNull(MagicStack!FRow)
        mCol = FaVNull(MagicStack!Fcol)
        MagicStack.Delete
        If MagicStack.RecordCount > 0 Then
            MagicStack.MoveLast
            Text1 = FaXNull(MagicStack!Name)
            If PubFaSiteType <> 0 Then SiteAsign
            MagLedger MagicStack!LedCode, mROW, mCol, CDate(TXTS_DATE), CDate(TXTE_DATE)
        End If
    End If
    PubSiteCode = SiteCodeStore
ElseIf FGrid1.Tag = "OPDIFF" Then
    If MagicStack.RecordCount > 1 Then
        MagicStack.MoveLast
        mROW = FaVNull(MagicStack!FRow)
        mCol = FaVNull(MagicStack!Fcol)
        MagicStack.Delete
        If MagicStack.RecordCount > 0 Then
            MagicStack.MoveLast
            Text1 = FaXNull(MagicStack!Name)
            If PubFaSiteType <> 0 Then SiteAsign
            OpDiff mROW, mCol, TXTS_DATE
        End If
    End If
    PubSiteCode = SiteCodeStore
End If
If mREFRESH = True Then
    Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
    If RstEnviro.RecordCount <= 0 Then MsgBox "Parameter Not Set": Exit Sub
    mREFRESH = False
End If
End Sub
Private Sub ShowSubtrial(mGroupCode As String)
Dim RST1 As adodb.Recordset
Set RST1 = G_FaCn.Execute("SELECT GROUPNAME FROM ACGROUP WHERE GROUPCODE=" & FaChk_Text(Trim(mGroupCode)))
If RST1.RecordCount > 0 Then
    MagicStack.AddNew
    MagicStack!TypeName = "SUBTRIAL"
    MagicStack!GroupCode = mGroupCode
    MagicStack!Name = RST1!GroupName
    MagicStack!FRow = Val(FGrid1.Row)
    MagicStack!Fcol = Val(FGrid1.Col)
    MagicStack!FS_DATE = TXTS_DATE
    MagicStack!FE_DATE = TXTE_DATE
    MagicStack.Update
    Text1 = RST1!GroupName
    Subtrial mGroupCode
End If
Set RST1 = Nothing
End Sub
Private Sub FGrid1_DblClick()
Dim mFromDate As Date, mToDate As Date
Check2.Value = 0
Select Case UCase(Trim(FGrid1.Tag))
    Case "PROFLOSS", "BALSHEET", "CASHFLOW", "FUNDFLOW"
        If Trim(FGrid1.TextMatrix(FGrid1.Row, 0)) = "*****" Then
            MagicStack.AddNew
            MagicStack!TypeName = "PROFLOSS"
            MagicStack!FS_DATE = TXTS_DATE
            MagicStack!FE_DATE = TXTE_DATE
            MagicStack!Check1 = Check1.Value
            MagicStack!Check2 = Check2.Value
            MagicStack!Check3 = Check3.Value
            MagicStack!Check4 = Check4.Value
            MagicStack!Check5 = Check5.Value
            MagicStack!Check6 = Check6.Value
            MagicStack.Update
            Text1 = ""
            ProfLoss
        ElseIf Trim(FGrid1.TextMatrix(FGrid1.Row, 0)) = "Diff" Then
            MagicStack.AddNew
            MagicStack!TypeName = "OPDIFF"
            MagicStack!FS_DATE = TXTS_DATE
            MagicStack!FE_DATE = TXTE_DATE
            MagicStack!Check1 = Check1.Value
            MagicStack!Check2 = Check2.Value
            MagicStack!Check3 = Check3.Value
            MagicStack!Check4 = Check4.Value
            MagicStack!Check5 = Check5.Value
            MagicStack!Check6 = Check6.Value
            MagicStack.Update
            Text1 = ""
            OpDiff , , TXTS_DATE
        Else
            Select Case FGrid1.Col
                Case 1, 2, 3
                    If Trim(FGrid1.TextMatrix(FGrid1.Row, 0)) <> "" Then ShowSubtrial Trim(FGrid1.TextMatrix(FGrid1.Row, 0))
                Case 6, 7, 8
                    If Trim(FGrid1.TextMatrix(FGrid1.Row, 5)) <> "" Then ShowSubtrial Trim(FGrid1.TextMatrix(FGrid1.Row, 5))
            End Select
        End If
    Case "OPDIFF"
        If Trim(FGrid1.TextMatrix(FGrid1.Row, 6)) = "Opening Balance" Or Trim(FGrid1.TextMatrix(FGrid1.Row, 1)) = "" Then Exit Sub
        MagicStack.AddNew
        MagicStack!TypeName = "OPDIFF"
        MagicStack!GroupCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 11))
        MagicStack!LedCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 12))
        MagicStack!Name = Trim(FGrid1.TextMatrix(FGrid1.Row, 2))
        MagicStack!FRow = Val(FGrid1.Row)
        MagicStack!Fcol = Val(FGrid1.Col)
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack!Check1 = Check1.Value
        MagicStack!Check2 = Check2.Value
        MagicStack!Check3 = Check3.Value
        MagicStack!Check4 = Check4.Value
        MagicStack!Check5 = Check5.Value
        MagicStack!Check6 = Check6.Value
        MagicStack.Update
        Text1 = ""
        OpeningDiff Trim(FGrid1.TextMatrix(FGrid1.Row, 8)), Val(FGrid1.TextMatrix(FGrid1.Row, 9)), Trim(FGrid1.TextMatrix(FGrid1.Row, 10)), FGrid1
    Case "LEDGER", "CASHBOOK"
       ' If Trim(FGrid1.TextMatrix(FGrid1.Row, 6)) = "Opening Balance" Or Trim(FGrid1.TextMatrix(FGrid1.Row, 1)) = "" Then Exit Sub
        MagicStack.AddNew
        MagicStack!TypeName = "VOUCHER"
        MagicStack!GroupCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 0))
        MagicStack!LedCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 1))
        MagicStack!Name = Trim(FGrid1.TextMatrix(FGrid1.Row, 2))
        MagicStack!FRow = Val(FGrid1.Row)
        MagicStack!Fcol = Val(FGrid1.Col)
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack!Check1 = Check1.Value
        MagicStack!Check2 = Check2.Value
        MagicStack!Check3 = Check3.Value
        MagicStack!Check4 = Check4.Value
        MagicStack!Check5 = Check5.Value
        MagicStack!Check6 = Check6.Value
        MagicStack.Update
        Text1 = ""
        Voucher Trim(FGrid1.TextMatrix(FGrid1.Row, 3)), Val(FGrid1.TextMatrix(FGrid1.Row, 4)), Trim(FGrid1.TextMatrix(FGrid1.Row, 5)), FGrid1
    Case "MONTHSUM"
        If Trim(FGrid1.TextMatrix(FGrid1.Row, 2)) = "Opening Balance" Or Trim(FGrid1.TextMatrix(FGrid1.Row, 2)) = "" Then Exit Sub
        mFromDate = CDate("01/" + Trim(FGrid1.TextMatrix(FGrid1.Row, 2)) + "/" + Trim(FGrid1.TextMatrix(FGrid1.Row, 1)))
        TXTS_DATE = mFromDate
        If UCase(Trim(FGrid1.TextMatrix(FGrid1.Row, 2))) = "DECEMBER" Then
            mToDate = CDate("31/" + Trim(FGrid1.TextMatrix(FGrid1.Row, 2)) + "/" + Trim(FGrid1.TextMatrix(FGrid1.Row, 1)))
        Else
            mToDate = CDate("01/" + Trim(FGrid1.TextMatrix(FGrid1.Row, 2)) + "/" + Trim(FGrid1.TextMatrix(FGrid1.Row, 1)))
            mToDate = DateAdd("M", 1, mToDate)
            mToDate = mToDate - 1
        End If
        TXTE_DATE = mToDate
        MagicStack.AddNew
        MagicStack!TypeName = "LEDGER"
        MagicStack!LedCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 0))
        MagicStack!Name = Trim(FGrid1.TextMatrix(FGrid1.Row, 6))
        MagicStack!FRow = Val(FGrid1.Row)
        MagicStack!Fcol = Val(FGrid1.Col)
        MagicStack!FS_DATE = mFromDate
        MagicStack!FE_DATE = mToDate
        MagicStack!Check1 = Check1.Value
        MagicStack!Check2 = Check2.Value
        MagicStack!Check3 = Check3.Value
        MagicStack!Check4 = Check4.Value
        MagicStack!Check5 = Check5.Value
        MagicStack!Check6 = Check6.Value
        MagicStack.Update
        Text1 = Trim(FGrid1.TextMatrix(FGrid1.Row, 6))
        MagLedger Trim(FGrid1.TextMatrix(FGrid1.Row, 0)), , , CDate(mFromDate), CDate(mToDate)
    Case "CASHBANKSUM"
        MagicStack.AddNew
        MagicStack!TypeName = "CASHBOOK"
        MagicStack!LedCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 0))
        MagicStack!Name = Trim(FGrid1.TextMatrix(FGrid1.Row, 1))
        MagicStack!FRow = Val(FGrid1.Row)
        MagicStack!Fcol = Val(FGrid1.Col)
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack!Check1 = Check1.Value
        MagicStack!Check2 = Check2.Value
        MagicStack!Check3 = Check3.Value
        MagicStack!Check4 = Check4.Value
        MagicStack!Check5 = Check5.Value
        MagicStack!Check6 = Check6.Value
        MagicStack.Update
        Text1 = Trim(FGrid1.TextMatrix(FGrid1.Row, 1))
        MagCashBook Trim(FGrid1.TextMatrix(FGrid1.Row, 0)), , , CDate(TXTS_DATE), CDate(TXTE_DATE)
    Case "GROUPTRIAL"
        If Trim(FGrid1.TextMatrix(FGrid1.Row, 2)) = "# Difference in Opening Balance" Then
            MagicStack.AddNew
            MagicStack!TypeName = "OPDIFF"
            MagicStack!FS_DATE = TXTS_DATE
            MagicStack!FE_DATE = TXTE_DATE
            MagicStack!FRow = Val(FGrid1.Row)
            MagicStack!Fcol = Val(FGrid1.Col)
            MagicStack!Check1 = Check1.Value
            MagicStack!Check2 = Check2.Value
            MagicStack!Check3 = Check3.Value
            MagicStack!Check4 = Check4.Value
            MagicStack!Check5 = Check5.Value
            MagicStack!Check6 = Check6.Value
            MagicStack.Update
            Text1 = ""
            OpDiff , , TXTS_DATE
        Else
            ShowSubtrial Trim(FGrid1.TextMatrix(FGrid1.Row, 0))
        End If
    Case "LEDTRIAL", "SUBTRIAL"
        If Trim(FGrid1.TextMatrix(FGrid1.Row, 2)) = "# Difference in Opening Balance" Then
            MagicStack.AddNew
            MagicStack!TypeName = "OPDIFF"
            MagicStack!FS_DATE = TXTS_DATE
            MagicStack!FE_DATE = TXTE_DATE
            MagicStack!FRow = Val(FGrid1.Row)
            MagicStack!Fcol = Val(FGrid1.Col)
            MagicStack!Check1 = Check1.Value
            MagicStack!Check2 = Check2.Value
            MagicStack!Check3 = Check3.Value
            MagicStack!Check4 = Check4.Value
            MagicStack!Check5 = Check5.Value
            MagicStack!Check6 = Check6.Value
            MagicStack.Update
            Text1 = ""
            OpDiff , , TXTS_DATE
        ElseIf Trim(FGrid1.TextMatrix(FGrid1.Row, 0)) <> "" And Trim(FGrid1.TextMatrix(FGrid1.Row, 1)) = "" Then
            ShowSubtrial Trim(FGrid1.TextMatrix(FGrid1.Row, 0))
        ElseIf Trim(FGrid1.TextMatrix(FGrid1.Row, 1)) <> "" And RstEnviro!MonthTotal = "Yes" Then
            MagicStack.AddNew
            MagicStack!TypeName = "MONTHSUM"
            MagicStack!GroupCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 0))
            MagicStack!LedCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 1))
            MagicStack!Name = Trim(FGrid1.TextMatrix(FGrid1.Row, 2))
            MagicStack!FRow = Val(FGrid1.Row)
            MagicStack!Fcol = Val(FGrid1.Col)
            MagicStack!FS_DATE = TXTS_DATE
            MagicStack!FE_DATE = TXTE_DATE
            MagicStack!Check1 = Check1.Value
            MagicStack!Check2 = Check2.Value
            MagicStack!Check3 = Check3.Value
            MagicStack!Check4 = Check4.Value
            MagicStack!Check5 = Check5.Value
            MagicStack!Check6 = Check6.Value
            MagicStack.Update
            Text1 = Trim(FGrid1.TextMatrix(FGrid1.Row, 2))
            monthSum Trim(FGrid1.TextMatrix(FGrid1.Row, 1))
        ElseIf Trim(FGrid1.TextMatrix(FGrid1.Row, 1)) <> "" Then
            MagicStack.AddNew
            MagicStack!TypeName = "LEDGER"
            MagicStack!LedCode = Trim(FGrid1.TextMatrix(FGrid1.Row, 1))
            MagicStack!Name = Trim(FGrid1.TextMatrix(FGrid1.Row, 2))
            MagicStack!FRow = Val(FGrid1.Row)
            MagicStack!Fcol = Val(FGrid1.Col)
            MagicStack!FS_DATE = TXTS_DATE
            MagicStack!FE_DATE = TXTE_DATE
            MagicStack!Check1 = Check1.Value
            MagicStack!Check2 = Check2.Value
            MagicStack!Check3 = Check3.Value
            MagicStack!Check4 = Check4.Value
            MagicStack!Check5 = Check5.Value
            MagicStack!Check6 = Check6.Value
            MagicStack.Update
            Text1 = Trim(FGrid1.TextMatrix(FGrid1.Row, 2))
            MagLedger Trim(FGrid1.TextMatrix(FGrid1.Row, 1)), , , CDate(TXTS_DATE), CDate(TXTE_DATE)
        End If
End Select
If UCase(Trim(FGrid1.Tag)) <> "VOUCHER" And UCase(Trim(FGrid1.Tag)) <> "OPDIFF" Then FGrid1.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim mROW As Integer, mCol As Integer
If KeyAscii = 27 Then
    If Frame1.Visible = True Then
        Frame1.Visible = False
        KeyAscii = 0
        Exit Sub
    End If
    If MagicStack.RecordCount = 1 Then Exit Sub
    If MagicStack.RecordCount > 1 Then
        MagicStack.MoveLast
        mROW = FaVNull(MagicStack!FRow)
        mCol = FaVNull(MagicStack!Fcol)
        MagicStack.Delete
    End If
    If MagicStack.RecordCount > 0 Then
        MagicStack.MoveLast
        If Not IsNull(MagicStack!FS_DATE) Then
            TXTS_DATE = MagicStack!FS_DATE
        End If
        If Not IsNull(MagicStack!FE_DATE) Then
            TXTE_DATE = MagicStack!FE_DATE
        End If
        If Not IsNull(MagicStack!Check1) Then
            Check1.Value = MagicStack!Check1
        Else
            Check1.Value = 1
        End If
        If Not IsNull(MagicStack!Check2) Then
            Check2.Value = MagicStack!Check2
        Else
            Check2.Value = 0
        End If
        If Not IsNull(MagicStack!Check3) Then
            Check3.Value = MagicStack!Check3
        Else
            Check3.Value = 0
        End If
        If Not IsNull(MagicStack!Check4) Then
            Check4.Value = MagicStack!Check4
        Else
            Check4.Value = 0
        End If
        If Not IsNull(MagicStack!Check5) Then
            Check5.Value = MagicStack!Check5
        Else
            Check5.Value = 0
        End If
        If Not IsNull(MagicStack!Check6) Then
            Check6.Value = MagicStack!Check6
        Else
            Check6.Value = 0
        End If
        Select Case UCase(Trim(MagicStack!TypeName))
            Case "PROFLOSS"
                Text1 = ""
                ProfLoss mROW, mCol
            Case "BALSHEET"
                Text1 = ""
                BalSheet mROW, mCol
            Case "SUBTRIAL"
                Text1 = MagicStack!Name
                Subtrial MagicStack!GroupCode, mROW, mCol
            Case "GROUPTRIAL"
                Text1 = ""
                GroupTrial mROW, mCol
            Case "LEDTRIAL"
                Text1 = ""
                LedTrial mROW, mCol
            Case "LEDGER"
                Text1 = MagicStack!Name
                MagLedger MagicStack!LedCode, mROW, mCol, TXTS_DATE, TXTE_DATE
            Case "CASHBOOK"
                Text1 = MagicStack!Name
                MagLedger MagicStack!LedCode, mROW, mCol
            Case "MONTHSUM"
                Text1 = MagicStack!Name
                monthSum MagicStack!LedCode, mROW, mCol
            Case "CASHBANKSUM"
                Text1 = ""
                CashBankSum mROW, mCol
            Case "CASHFLOW"
                Text1 = ""
                CashFlow mROW, mCol
            Case "FUNDFLOW"
                Text1 = ""
                FundFlow mROW, mCol
        End Select
    End If
ElseIf KeyAscii = 13 Then
    If Frame1.Visible = False Then FGrid1_DblClick
End If
If Frame1.Visible = False And TxtSearch.Visible = False Then
    If UCase(Trim(FGrid1.Tag)) <> "VOUCHER" Then FGrid1.SetFocus
End If
End Sub
Private Sub btnok_Click()
MagicStack.MoveLast

If CDate(TXTE_DATE) < CDate(TXTS_DATE) Then MsgBox "To Date Can't be less then From Date": TXTE_DATE.SetFocus: Exit Sub
Frame1.Visible = False
MagicStack!FS_DATE = TXTS_DATE
MagicStack!FE_DATE = TXTE_DATE
MagicStack!Check1 = Check1.Value
MagicStack!Check2 = Check2.Value
MagicStack!Check3 = Check3.Value
MagicStack!Check4 = Check4.Value
MagicStack!Check5 = Check5.Value
MagicStack!Check6 = Check6.Value
MagicStack.Update
Select Case UCase(Trim(MagicStack!TypeName))
    Case "PROFLOSS"
        Text1 = ""
        ProfLoss
    Case "BALSHEET"
        Text1 = ""
        BalSheet
    Case "SUBTRIAL"
        Text1 = MagicStack!Name
        Subtrial MagicStack!GroupCode
    Case "LEDGER"
        Text1 = MagicStack!Name
        MagLedger MagicStack!LedCode, , , CDate(TXTS_DATE), CDate(TXTE_DATE)
    Case "CASHBOOK"
        Text1 = MagicStack!Name
        MagCashBook MagicStack!LedCode, , , CDate(TXTS_DATE), CDate(TXTE_DATE)
    Case "GROUPTRIAL"
        Text1 = ""
        GroupTrial
    Case "LEDTRIAL"
        Text1 = ""
        LedTrial
    Case "MONTHSUM"
        Text1 = MagicStack!Name
        monthSum MagicStack!LedCode
    Case "CASHBANKSUM"
        Text1 = FaXNull(MagicStack!Name)
        CashBankSum
    Case "CASHFLOW"
        Text1 = ""
        CashFlow
    Case "FUNDFLOW"
        Text1 = ""
        FundFlow
End Select
FGrid1.SetFocus
End Sub
Private Sub LblShort_Click(Index As Integer)
ColoChange Index
Select Case Index
    Case 0        'BALANCE SHEET
        If FGrid1.Tag <> "BALSHEET" Then
            Set MagicStack = Nothing
            Me.OpenType = "BALSHEET"
        End If
    Case 1        'PROFIT/LOSS
        If FGrid1.Tag <> "PROFLOSS" Then
            Set MagicStack = Nothing
            Me.OpenType = "PROFLOSS"
        End If
    Case 2        'TRIAL (GROUP)
        If FGrid1.Tag <> "GROUPTRIAL" Then
            Set MagicStack = Nothing
            Me.OpenType = "GROUPTRIAL"
        End If
    Case 3        'TRIAL (LEDGER)
        If FGrid1.Tag <> "LEDTRIAL" Then
            Set MagicStack = Nothing
            Me.OpenType = "LEDTRIAL"
        End If
    Case 4        'CASHFLOW
        If FGrid1.Tag <> "CASHFLOW" Then
            Set MagicStack = Nothing
            Me.OpenType = "CASHFLOW"
        End If
    Case 5        'FUNDFLOW
        If FGrid1.Tag <> "FUNDFLOW" Then
            Set MagicStack = Nothing
            Me.OpenType = "FUNDFLOW"
        End If
    Case 6        'CASHBANKSUM
        If FGrid1.Tag <> "CASHBANKSUM" Then
            Set MagicStack = Nothing
            Me.OpenType = "CASHBANKSUM"
        End If
End Select
End Sub
Private Sub FrmHotKey(ByRef KeyCode As Integer, ByRef Index As Integer)
Dim I As Integer
For I = 0 To 6
    LblShort(I).FontBold = False
    LblShort(I).ForeColor = &HFFFF&
Next
LblShort(Index).ForeColor = &HFFFF00
LblShort(Index).FontBold = True
End Sub
Public Property Let OpenType(ByVal TITL As String)
Set MagicStack = New adodb.Recordset
Set MagicStack = PubDatamanFa.FaStackRst(MagicStack)
Select Case UCase(TITL)
    Case "BALSHEET"
        MagicStack.AddNew
        MagicStack!TypeName = "BALSHEET"
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack.Update
        BalSheet
    Case "PROFLOSS"
        MagicStack.AddNew
        MagicStack!TypeName = "PROFLOSS"
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack.Update
        ProfLoss
    Case "GROUPTRIAL"
        MagicStack.AddNew
        MagicStack!TypeName = "GROUPTRIAL"
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack.Update
        GroupTrial
    Case "LEDTRIAL"
        MagicStack.AddNew
        MagicStack!TypeName = "LEDTRIAL"
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack.Update
        LedTrial
    Case "CASHBANKSUM"
        MagicStack.AddNew
        MagicStack!TypeName = "CASHBANKSUM"
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack.Update
        CashBankSum
    Case "CASHFLOW"
        MagicStack.AddNew
        MagicStack!TypeName = "CASHFLOW"
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack.Update
        CashFlow
    Case "FUNDFLOW"
        MagicStack.AddNew
        MagicStack!TypeName = "FUNDFLOW"
        MagicStack!FS_DATE = TXTS_DATE
        MagicStack!FE_DATE = TXTE_DATE
        MagicStack.Update
        FundFlow
End Select
End Property
Private Sub BTNPRINT_Click(Index As Integer)
Dim I As Integer, tit As Integer, X1
If Frame1.Visible = True Or Frame3.Visible = True Then Exit Sub
Select Case Index
    Case 1
        Select Case UCase(Trim(FGrid1.Tag))
            Case "BALSHEET", "PROFLOSS"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                If UCase(Trim(FGrid1.Tag)) = "PROFLOSS" And RstEnviro!ShowQty = "Yes" Then
                    tit = PubDatamanFa.FaProfitLossQtyDosPrinting(Me, UCase(Trim(FGrid1.Tag)), RstEnviro!VerticleBalanceSheet): Exit Sub
                Else
                    tit = PubDatamanFa.FaBalanceSheetDosPrinting(Me, UCase(Trim(FGrid1.Tag)), RstEnviro!VerticleBalanceSheet): Exit Sub
                End If
            Case "GROUPTRIAL", "LEDTRIAL", "SUBTRIAL"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                tit = PubDatamanFa.FaLedTrialDosPrinting(Me, UCase(Trim(FGrid1.Tag))): Exit Sub
            Case "MONTHSUM"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                tit = PubDatamanFa.FaMonthSumDosPrinting(Me, UCase(Trim(FGrid1.Tag))): Exit Sub
            Case "CASHFLOW", "FUNDFLOW"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                tit = PubDatamanFa.FaCashFlowDosPrinting(Me, UCase(Trim(FGrid1.Tag))): Exit Sub
            Case "CASHBANKSUM"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                tit = PubDatamanFa.FaCashBankSumDosPrinting(Me, UCase(Trim(FGrid1.Tag))): Exit Sub
            Case "LEDGER"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                tit = PubDatamanFa.FaLedgerDosPrinting(Me, UCase(Trim(FGrid1.Tag))): Exit Sub
                Exit Sub
            Case "CASHBOOK"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                tit = PubDatamanFa.FaCashBookDosPrinting(Me, UCase(Trim(FGrid1.Tag))): Exit Sub
            Case "OPDIFF"
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                tit = PubDatamanFa.FaOpDiffDosPrinting(Me): Exit Sub
        End Select
    Case Else
        Select Case UCase(Trim(FGrid1.Tag))
            Case "BALSHEET", "PROFLOSS"
                If UCase(Trim(FGrid1.Tag)) = "PROFLOSS" And RstEnviro!ShowQty = "Yes" Then
                    Set rpt = PubDatamanFa.FaProfLossQtyRpt
                Else
                    Set rpt = PubDatamanFa.FaProfLossRpt
                End If
'                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaProfLoss.ttx", True)
            Case "GROUPTRIAL", "LEDTRIAL", "SUBTRIAL"
                Set rpt = PubDatamanFa.FaLedTrialRpt
                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaLedTrial.ttx", True)
            Case "MONTHSUM"
                Set rpt = PubDatamanFa.FaMonthSumRpt
'                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaMonthSum.ttx", True)
            Case "CASHFLOW", "FUNDFLOW"
                Set rpt = PubDatamanFa.FaCashflRpt
'                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaCashfl.ttx", True)
            Case "CASHBANKSUM"
                Set rpt = PubDatamanFa.FaCBCBRpt
'                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaCBCB.ttx", True)
            Case "CASHBOOK"
                Set rpt = PubDatamanFa.FaMagCashBookRpt
'                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaMagCashBook.ttx", True)
            Case "LEDGER"
                Set rpt = PubDatamanFa.FaMagLedgerRpt
'                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaMagLedger.ttx", True)
            Case "OPDIFF"
                Set rpt = PubDatamanFa.FaOpDiffRpt
'                X1 = CreateFieldDefFile(FGrid1.DataSource, PubFaReportPath + "\FaOpDiff.ttx", True)
        End Select
End Select
rpt.Database.SetDataSource FGrid1.DataSource
For I = 1 To rpt.FormulaFields.Count
    Select Case rpt.FormulaFields(I).FormulaFieldName
        Case "title", "TITLE"
            rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
        Case "DATE", "date", "FROMDATE", "DT"
            If UCase(Trim(FGrid1.Tag)) = "BALSHEET" Then
                rpt.FormulaFields(I).TEXT = "'As On " & TXTE_DATE & "'"
            Else
                rpt.FormulaFields(I).TEXT = "'From " & TXTS_DATE & " To " & TXTE_DATE & "'"
            End If
        Case "ForParty"
            rpt.FormulaFields(I).TEXT = IIf(Trim(Text1) = "", "''", "'For       : " & Text1 & "'")
        Case "Lia"
            Select Case UCase(Trim(FGrid1.Tag))
                Case "BALSHEET"
                    If RstEnviro!VerticleBalanceSheet = "No" Then
                        rpt.FormulaFields(I).TEXT = "'Liabilities'"
                    Else
                        rpt.FormulaFields(I).TEXT = "'Particulars'"
                    End If
                Case "PROFLOSS"
                    rpt.FormulaFields(I).TEXT = "'Particulars'"
            End Select
        Case "ASS"
            Select Case UCase(Trim(FGrid1.Tag))
                Case "BALSHEET"
                    If RstEnviro!VerticleBalanceSheet = "No" Then
                        rpt.FormulaFields(I).TEXT = "'Assets'"
                    Else
                        rpt.FormulaFields(I).TEXT = "''"
                    End If
                Case "PROFLOSS"
                    rpt.FormulaFields(I).TEXT = "'Particulars'"
            End Select
        Case "Amt"
            Select Case UCase(Trim(FGrid1.Tag))
                Case "BALSHEET"
                    If RstEnviro!VerticleBalanceSheet = "No" Then
                        rpt.FormulaFields(I).TEXT = "'Amount'"
                    Else
                        rpt.FormulaFields(I).TEXT = "''"
                    End If
                Case "PROFLOSS"
                    rpt.FormulaFields(I).TEXT = "'Amount'"
            End Select
        Case "DosPrint"
            rpt.FormulaFields(I).TEXT = IIf(Index = 1, "'Y'", "'N'")
        Case "ShowQty"
            rpt.FormulaFields(I).TEXT = "'" & RstEnviro!ShowQty & "'"
        Case "PageNo"
            rpt.FormulaFields(I).TEXT = "'" & RstEnviro!pagenofill & "'"
        Case "DT1"
             rpt.FormulaFields(I).TEXT = "'" & RstEnviro!daterfill & "'"
    End Select
Next
rpt.ReadRecords
FaReport_View rpt, Index, Me.CAPTION, True
Set rpt = Nothing
End Sub
Public Property Let ShowSiteWise(ByVal YesNo As Boolean)
    PubDatamanFa.FaSiteWise = YesNo
    If YesNo = True Then
        BtnSite.Visible = True
    Else
        BtnSite.Visible = False
    End If
End Property
Private Sub BtnSite_Click()
If Frame1.Visible = True Or Frame3.Visible = True Then Exit Sub
    Frame3.Visible = True
    Frame3.ZOrder 0
End Sub
Private Sub BTNSITEOK_Click()
    SiteAsign
    btnok_Click
End Sub
Private Sub SiteAsign()
Dim ac_str As String, ac_str1 As String, I As Integer, GridRow As Integer, FormulaString As String
    Frame3.Visible = False
    ac_str = ""
    ac_str1 = ""
    If Check(0).Value = 1 Then
        RsGrid1.MoveFirst
        Do Until RsGrid1.EOF
            ac_str = ac_str + IIf(ac_str = "", "'" + RsGrid1!Site_Code + "'", "," + "'" + RsGrid1!Site_Code + "'")
            ac_str1 = ac_str1 + IIf(ac_str1 = "", Trim(RsGrid1!Site), "," + Trim(RsGrid1!Site))
            RsGrid1.MoveNext
        Loop
        If ac_str <> "" Then
            PubDatamanFa.FaPubSiteCodeDisplay = "(" + ac_str + ")"
            PubSiteCodeDisplay = "(" + ac_str + ")"
        End If
        If ac_str1 <> "" Then
            PubDatamanFa.FaPubSiteName = ac_str1
            PubSiteName = ac_str1
        End If
    Else
        ac_str = ""
        For I = 0 To UBound(GridRow1)
            If GridRow1(I) = 0 Then GoTo NXT:
            GridRow = GridRow1(I)
            If GridSel(0).TextMatrix(GridRow, 0) = "ü" Then
                ac_str = ac_str + IIf(ac_str = "", "'" + GridSel(0).TextMatrix(GridRow, 2) + "'", "," + "'" + GridSel(0).TextMatrix(GridRow, 2) + "'")
            End If
NXT:
        Next
        If ac_str <> "" Then
            PubDatamanFa.FaPubSiteCodeDisplay = "(" + ac_str + ")"
            PubSiteCodeDisplay = "(" + ac_str + ")"
        End If
        ac_str = ""
        For I = 0 To UBound(GridRow1)
            If GridRow1(I) = 0 Then GoTo nxt1:
            GridRow = GridRow1(I)
            If GridSel(0).TextMatrix(GridRow, 0) = "ü" Then
                ac_str = ac_str + IIf(ac_str = "", GridSel(0).TextMatrix(GridRow, 1), "," + GridSel(0).TextMatrix(GridRow, 1))
            End If
nxt1:
        Next
        If ac_str <> "" Then
            PubDatamanFa.FaPubSiteName = ac_str
            PubSiteName = ac_str
        End If
    End If
End Sub
Private Sub GridSel_EnterCell(Index As Integer)
    GridSel(Index).CellBackColor = FaCellBackColEnter1
End Sub
Private Sub GridSel_GotFocus(Index As Integer)
    GridSel(Index).CellBackColor = FaCellBackColEnter1
End Sub
Private Sub GridSel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
    If GridSel(Index).Rows < 1 Then Exit Sub
    If KeyCode = vbKeySpace And GridSel(Index).Col = 0 Then
        GridSel(Index).CellFontName = "WINGDINGS"
        GridSel(Index).CellFontSize = 14
        If GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü" Then
            GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = " "
            Select Case Index
                Case 0
                    For I = 0 To UBound(GridRow1)
                        If GridRow1(I) = GridSel(Index).Row Then GridRow1(I) = 0
                    Next
            End Select
        Else
            GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü"
            Select Case Index
                Case 0
                    I = UBound(GridRow1) + 1
                    ReDim Preserve GridRow1(I)
                    GridRow1(I) = GridSel(Index).Row
            End Select
        End If
    End If
End Sub
Private Sub GridSel_KeyPress(Index As Integer, KeyAscii As Integer)
    If GridSel(Index).Col = 0 Or GridSel(Index).Row = 0 Then Exit Sub
    Select Case Index
        Case 0
           FaSelGridKeyPress TxtSearch1, GridSel(Index), RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Index).Col).Name, FaCellBackColEnter1, FaCellBackColLeave1: KeyAscii = 0
    End Select
    TxtSearch.Tag = Index
End Sub
Private Sub GridSel_LeaveCell(Index As Integer)
    GridSel(Index).CellBackColor = FaCellBackColLeave1
End Sub
Private Sub GridSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GridSel(Index).Col <> 0 Then Exit Sub
    mGridStartRow = GridSel(Index).Row
End Sub
Private Sub GridSel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer, j As Integer
    If GridSel(Index).Col <> 0 Or mGridStartRow = 0 Then Exit Sub
    mGridEndRow = GridSel(Index).RowSel
    For j = mGridStartRow To mGridEndRow
        GridSel(Index).Row = j
        GridSel(Index).Col = 0
        GridSel(Index).CellFontName = "WINGDINGS"
        GridSel(Index).CellFontSize = 14
        If GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü" Then
            GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = " "
            Select Case Index
                Case 0
                    For I = 0 To UBound(GridRow1)
                        If GridRow1(I) = GridSel(Index).Row Then GridRow1(I) = 0
                    Next
            End Select
        Else
            GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü"
            Select Case Index
                Case 0
                    I = UBound(GridRow1) + 1
                    ReDim Preserve GridRow1(I)
                    GridRow1(I) = GridSel(Index).Row
            End Select
        End If
    Next
    mGridStartRow = 0
End Sub
Private Sub GridSel_Validate(Index As Integer, Cancel As Boolean)
    GridSel(Index).CellBackColor = FaCellBackColLeave1
End Sub
Private Sub Check_Click(Index As Integer)
    If Check(Index).Value = Unchecked Then
        GridSel(Index).Enabled = True
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 1: GridSel(Index).Col = 0
        End If
    Else
        GridSel(Index).Enabled = False
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 0: GridSel(Index).Col = 0
            GridSel(Index).RowSel = GridSel(Index).Rows - 1
        End If
    End If
End Sub
Private Sub Check_GotFocus(Index As Integer)
    Check(Index).BackColor = &HFF&
End Sub
Private Sub Check_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub
Private Sub TxtSearch_Click()
    TxtSearch.TEXT = ""
    FGrid1.SetFocus
    TxtSearch.Visible = False
End Sub
Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If FaNavigationKey(KeyCode) = True Then FGrid1.SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
End Sub
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
    FGrid1_KeyPress KeyAscii
End Sub
Private Sub TxtSearch_LostFocus()
    TxtSearch.TEXT = ""
    FGrid1.SetFocus
    TxtSearch.Visible = False
End Sub
Private Sub TxtSearch1_Click()
    TxtSearch1.TEXT = ""
    GridSel(Val(TxtSearch1.Tag)).SetFocus
    TxtSearch1.Visible = False
End Sub
Private Sub TxtSearch1_KeyDown(KeyCode As Integer, Shift As Integer)
    If FaNavigationKey(KeyCode) = True Then GridSel(Val(TxtSearch1.Tag)).SetFocus: TxtSearch1.Visible = False
    If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then GridSel(Val(TxtSearch1.Tag)).SetFocus: TxtSearch1.Visible = False
End Sub
Private Sub TxtSearch1_KeyPress(KeyAscii As Integer)
    Select Case TxtSearch1.Tag
        Case 0
           FaSelGridKeyPress TxtSearch1, GridSel(Val(TxtSearch1.Tag)), RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Val(TxtSearch1.Tag)).Col).Name, FaCellBackColEnter1, FaCellBackColLeave1
    End Select
End Sub
Private Sub TxtSearch1_LostFocus()
    TxtSearch1.TEXT = ""
    GridSel(Val(TxtSearch1.Tag)).SetFocus
    TxtSearch1.Visible = False
End Sub
Private Sub OpDiff(Optional FRow As Integer, Optional Fcol As Integer, Optional magStartDate As Date)
Dim RstOpDiff As adodb.Recordset
Set RstOpDiff = PubDatamanFa.FaOpenDiff(Me, FRow, Fcol)
ColoChange 0
If RstOpDiff.RecordCount <= 0 Then FGrid1.ClearStructure
'X11 = CreateFieldDefFile(RstBalSheet, PubFaReportPath + "\FaBalSheet.ttx", True)
Set RstOpDiff = Nothing
End Sub

