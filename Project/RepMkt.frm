VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form RepMkt 
   BackColor       =   &H00C8E8DA&
   Caption         =   "ReprtForm"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11820
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   11820
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Print"
      DownPicture     =   "RepMkt.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   6075
      Width           =   1290
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
      DownPicture     =   "RepMkt.frx":3132
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exit Form"
      Top             =   6075
      Width           =   1290
   End
   Begin VB.PictureBox Pic 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11820
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6045
      Width           =   11820
      Begin VB.Label LblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "LblTitle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   7230
         TabIndex        =   17
         Top             =   0
         Width           =   4470
      End
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7290
      TabIndex        =   14
      Top             =   -675
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   405
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
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
         BackColor       =   16777152
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   -135
      TabIndex        =   12
      Top             =   1470
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
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
      Index           =   4
      Left            =   5040
      TabIndex        =   7
      Top             =   3900
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
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
      Index           =   3
      Left            =   75
      TabIndex        =   5
      Top             =   3885
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
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
      Index           =   2
      Left            =   5685
      TabIndex        =   3
      Top             =   1830
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
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
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1860
      Visible         =   0   'False
      Width           =   915
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   90
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   285
      Visible         =   0   'False
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1785
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
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
      BackColorBkg    =   14873572
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   2
      Left            =   4965
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
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
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   4
      Left            =   4995
      TabIndex        =   8
      Top             =   3825
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
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
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   3
      Left            =   165
      TabIndex        =   6
      Top             =   3825
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
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
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1560
      Left            =   1575
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2752
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16512
      Rows            =   5
      Cols            =   3
      FixedRows       =   0
      BackColorFixed  =   13166810
      ForeColorFixed  =   16384
      BackColorSel    =   16711680
      ForeColorSel    =   12648447
      BackColorBkg    =   13166810
      GridColor       =   13166810
      GridColorFixed  =   13166810
      GridColorUnpopulated=   12648447
      GridLinesFixed  =   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
End
Attribute VB_Name = "RepMkt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CellBackColLeave As String = &HFFFFFF
Private Const CellBackColEnter As String = &HFFFFC0
Private Const CellBackColLeave1 As String = &HEDF7FE
Private Const CellBackColEnter1 As String = &HFFFFC0
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim RsGrid3 As ADODB.Recordset
Dim RsGrid4 As ADODB.Recordset
Dim RepTitle As String, RepName As String
Dim RepPrint As Boolean
Dim RstRep As ADODB.Recordset
Dim RstRep1 As ADODB.Recordset
Dim SubRep1 As Boolean
Private Const GridRowHeight As Integer = 270
'////////********VEHICLE***********////////////////////*****
Private Const DailyActivity As Byte = 1
Private Const AppointMent As Byte = 2
Private Const CallStatus As Byte = 3
Private Const CaseAnalysis As Byte = 4
Private Const ActivityMissing As Byte = 5
Private Const AppointmentNotKept As Byte = 6
Private Const GotLostRep As Byte = 7
Private Const PipeLineRep As Byte = 10
Private Const ProfPurRep As Byte = 11
Private Const DailySaleRep As Byte = 12
Private Const SalesTracRep As Byte = 13
Private Const FinTracRep As Byte = 14

Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const List1 As Byte = 2
Private Const List2 As Byte = 3
Private Const List3 As Byte = 4

Private Const Cat1 As Byte = 5
Private Const Cat2 As Byte = 6
Private Const Cat3 As Byte = 7
Private Const Cat4 As Byte = 8
Private Const Cat5 As Byte = 9

Public GRepFormName As String
Dim mLastRow As Integer
Dim mFirstRow As Integer
Dim mHelpGridNo
Dim GridKey As Integer
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim GridString1 As String
Dim GridString2 As String
Dim GridString3 As String
Dim GridString4 As String
Dim GridRow1() As Integer
Dim GridRow2() As Integer
Dim GridRow3() As Integer
Dim GridRow4() As Integer
Dim mGridStartRow As Integer
Dim mGridEndRow As Integer

Private Const SprMrRct$ = "SXGR"           'Material Receipt
Private Const SprMrTrf$ = "SXGRT"          'Material Rectipt Transfer
Private Const SprSlChal$ = "SYSC"           'Sale Challan       LPS 24-09
Private Const SprTrfChal$ = "SYSCT"         'Transfer Issue     LPS 24-09
Private Const SprSlCsh$ = "SYSIC"          'Cash Sale
Private Const SprSlCre$ = "SYSIR"          'Credit Sale
Private Const WksSlCsh$ = "W_SIC"          'Cash Sale
Private Const WksSlCre$ = "W_SIR"          'Credit Sale
Private Const SprSlRetCsh$ = "SXSRC"       'Cash Sale Return
Private Const SprSlRetCre$ = "SXSRR"       'Credit Sale Return
Private Const SprSlTrfRet$ = "SXSRT"       'Transfer Issue Return
Private Const SprPurCsh$ = "SXPIC"         'Cash Purchase
Private Const SprPurCre$ = "SXPIR"         'Credit Purchase
Private Const SprPrRetCsh$ = "SYPRC"       'Purchase Return Cash
Private Const SprPrRetCre$ = "SYPRR"       'Purchase Return Credit
Private Const SprPrTrfRet$ = "SYPRT"       'Transfer Receipt Return
Private Const SprQuotation$ = "S_QU"       'Spare Quotation
Private Const WksEst$ = "W_EST"       'Workshop Estimation
Private Const WksPro$ = "W_PL"       'Workshop Proforma Labour
Private Const WksGenReq$ = "W_RG"       'Workshop General Reqisition
Private Const WksReqWrt$ = "W_RW"       'Workshop Warranti Reqisition
Dim mListItem As ListItem

Private Sub btnexit_Click()
    Unload Me
End Sub

Private Sub BTNPRINT_Click()
On Error GoTo ERRORHANDLER
SubRep1 = False
RepPrint = True
Select Case GRepFormName
    Case DailyActivity
         If DailyActivityProc = False Then Exit Sub
    Case AppointMent
         If AppointmentProc = False Then Exit Sub
    Case CallStatus
         If CallStatusProc = False Then Exit Sub
    Case CaseAnalysis
         If CaseAnalysisProc = False Then Exit Sub
    Case PipeLineRep
         If PipeLineRepProc = False Then Exit Sub
    Case ProfPurRep
         If ProfPurRepProc = False Then Exit Sub
    Case DailySaleRep
        If DailySaleProc = False Then Exit Sub
    Case SalesTracRep, FinTracRep
        If SalesTracRepProc = False Then Exit Sub
End Select

CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If GRepFormName = DailySaleRep Then
    CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & "1.ttx", True
    CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & "2.ttx", True
End If
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")
rpt.Database.SetDataSource RstRep
If GRepFormName = DailySaleRep Then
   rpt.OpenSubreport("SubRep1").Database.SetDataSource RstRep
   rpt.OpenSubreport("SubRep2").Database.SetDataSource RstRep
Else
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
End If
rpt.ReadRecords
Set RstRep = Nothing
'Set rpt = Nothing 'Auto done by report_view function

Call Formulas
Call Report_View(rpt, RepTitle, , False)
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub Check1_Click(Index As Integer)
    If Check1(Index).Value = Unchecked Then
        GridSel(Index).Enabled = True
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 1: GridSel(Index).Col = 1
        End If
    Else
        GridSel(Index).Enabled = False
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 0: GridSel(Index).Col = 0
            GridSel(Index).RowSel = GridSel(Index).Rows - 1
        End If
    End If
End Sub

Private Sub Check1_GotFocus(Index As Integer)
    Check1(Index).BackColor = &HFF&
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub Check1_Validate(Index As Integer, Cancel As Boolean)
Check1(Index).BackColor = &H800000
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
WinSetting Me  ', 6885, 11500
   Global_Grid
   TopCtrl1.TopText2 = "Add"
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    WinSetting Me ', 6885, 11500
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If GridSel(4).Visible = True Then Set RsGrid1 = Nothing
If GridSel(1).Visible = True Then Set RsGrid2 = Nothing
If GridSel(2).Visible = True Then Set RsGrid3 = Nothing
If GridSel(3).Visible = True Then Set RsGrid4 = Nothing
Set RstRep = Nothing
Set mListItem = Nothing
Set rpt = Nothing
End Sub


Private Sub GridSel_EnterCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColEnter1
End Sub

Private Sub GridSel_GotFocus(Index As Integer)
GridSel(Index).CellBackColor = CellBackColEnter1
End Sub

Private Sub GridSel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = 13 Then SendKeysA vbKeyTab, True
If GridSel(Index).Rows < 1 Then Exit Sub
If KeyCode = vbKeySpace And GridSel(Index).Col = 0 Then
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü", " ", "ü")
    Select Case Index
        Case 1
            I = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(I)
            GridRow1(I) = GridSel(Index).Row
        Case 2
            I = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(I)
            GridRow2(I) = GridSel(Index).Row
        Case 3
            I = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(I)
            GridRow3(I) = GridSel(Index).Row
        Case 4
            I = UBound(GridRow4) + 1
            ReDim Preserve GridRow4(I)
            GridRow4(I) = GridSel(Index).Row
    End Select
End If
End Sub

Private Sub GridSel_KeyPress(Index As Integer, KeyAscii As Integer)
If GridSel(Index).Col = 0 Or GridSel(Index).Row = 0 Then Exit Sub
Select Case Index
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid4, KeyAscii, RsGrid4.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
End Select
TxtSearch.Tag = Index
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
FGrid.TextMatrix(Cat2, 1) = Val(FGrid.TextMatrix(Cat1, 1)) + 1
End Sub

Private Sub TxtSearch_Click()
TxtSearch.TEXT = "": GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If NavigationKey(KeyCode) = True Then GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
Select Case TxtSearch.Tag
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid4, KeyAscii, RsGrid4.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
End Select
End Sub

Private Sub TxtSearch_LostFocus()
    TxtSearch.TEXT = "": GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
End Sub

Private Sub GridSel_LeaveCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub

Private Sub GridSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If GridSel(Index).Col <> 0 Then Exit Sub
mGridStartRow = GridSel(Index).Row
End Sub

Private Sub GridSel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
Dim j As Integer
If GridSel(Index).Col <> 0 Or mGridStartRow = 0 Then Exit Sub
mGridEndRow = GridSel(Index).RowSel
For j = mGridStartRow To mGridEndRow
    GridSel(Index).Row = j
    GridSel(Index).Col = 0
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(j, 0) = IIf(GridSel(Index).TextMatrix(j, 0) = "ü", " ", "ü")
    Select Case Index
        Case 1
            I = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(I)
            GridRow1(I) = GridSel(Index).Row
        Case 2
            I = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(I)
            GridRow2(I) = GridSel(Index).Row
        Case 3
            I = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(I)
            GridRow3(I) = GridSel(Index).Row
        Case 4
            I = UBound(GridRow4) + 1
            ReDim Preserve GridRow4(I)
            GridRow4(I) = GridSel(Index).Row
    End Select
Next
mGridStartRow = 0
End Sub

Private Sub GridSel_Validate(Index As Integer, Cancel As Boolean)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub

Private Sub ListView_Click()
    TxtGrid(0).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    TxtGrid(0).SetFocus
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Select Case FGrid.Row
    Case List1
       Select Case GRepFormName
           Case CallStatus
                ListArray = Array("Cold", "Warm", "Hot", "Nill")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case CaseAnalysis
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
       End Select
    Case List2
       Select Case GRepFormName
           Case CaseAnalysis
               ListArray = Array("Yes", "No")
               Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
       End Select
    Case List3
'       Select Case GRepFormName
'           Case VehSalereg
'               ListArray = Array("PartyWise", "CityWise", "FinancierGrp", "FinancierName", "FormType", "All")
'               Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 6)
'       End Select
    Case Cat2
'        Select Case GRepFormName
'            Case ModFWiseMicro
'
'                FGrid.TextMatrix(Cat2, 1) = Val(FGrid.TextMatrix(Cat1, 1)) + 1
'                TxtGridLeave
'        End Select
End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
    Exit Sub
End If
Select Case FGrid.Row
    Case List1, List2, List3
        ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
    Case Date1, Date2, Cat1, Cat3, Cat4, Cat5
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Dim KeyCode As Integer
Call CheckQuote(KeyAscii)
Select Case FGrid.Row
    Case Cat1
'        Select Case GRepFormName
'            Case ModFWiseMicro
'                NumPress TxtGrid(Index), KeyAscii, 4, 0
'        End Select
   
        'KeyCode = 0
'           TxtGrid(0).Enabled = False

'    Case Cat3
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
'    Case Cat4
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
'    Case Cat5
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
    Select Case FGrid.Row
'        Case Cat1, Cat2
'             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
'
'        Case Cat2
'            'If Val(FGrid.TextMatrix(Cat2, 1))  > Val(FGrid.TextMatrix(Cat2, 1)) Then
            
        Case List1, List2, List3
            If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
            ListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
    End Select
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Dim KeyCode As Integer
Select Case FGrid.Row
    Case Cat3, Cat4, Cat5
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
    Case Cat1
'        TxtGrid(0).Text = GetYear
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
        If TxtGrid(0).TEXT <= (Format(PubStartDate, "YYYY") - 3) Or TxtGrid(0).TEXT >= (Format(PubEndDate, "YYYY")) Then
            MsgBox "Invalid Year Selection!"
            TxtGridLeave = False: Exit Function
        End If
               
    Case Cat2
        FGrid.TextMatrix(Cat2, 1) = Val(FGrid.TextMatrix(Cat1, 1)) + 1
'        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
    Case List1
        If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
    Case List2
        If TxtGrid(0).TEXT <> "" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            
            If GRepFormName = CaseAnalysis Then
            If FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No" Then
                Grid2Sql = "Select ' ' as O, PSG.Name + PSG.NSuffix + City.CityName as ProspectiveCustomerName,PSG.Cust_Code as Code " & _
                " from ProspectiveCust PSG left join City on PSG.CityCode=City.CityCode " & _
                " where Cust_code in (Select distinct Party_Code from Visits where ProspectiveCust_SubGroup=0) " & _
                " Order By PSG.Name"
            Else
                Grid2Sql = " Select ' ' as O, SG.Name + ' ' + City1.CityName as PartyName, SG.SubCode as Code " & _
                " from SubGroup SG left join City City1 on SG.CityCode=City1.CityCode " & _
                " where SubCode in (Select distinct Party_Code from Visits where ProspectiveCust_SubGroup=1) " & _
                " Order By SG.Name"
            End If
            GridInitialise 2, Grid2Sql
            End If
        End If
    Case List3
        If TxtGrid(0).TEXT <> "" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
'            Select Case TxtGrid(0).Text
'                Case "PartyWise"
'                    Grid3Sql = "select '' as O,Name as Party_Name,SubCode  as code from Subgroup order by Name"
'                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
'                 Case "CityWise"
'                    Grid3Sql = "select '' as O,CityName as City_Name,CityCode  as code from City order by CityName"
'                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
'                 Case "FinancierGrp"
'                    Grid3Sql = "select '' as O,FinGrpName as FinGrp_Name,FinGrpCode  as code from FinGroup order by FinGrpName"
'                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
'                 Case "FinancierName"
'                    Grid3Sql = "select '' as O,FinName as Financer_Name,FinCode  as code from ContractFinance order by FinName"
'                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
'                 Case "FormType"
'                    Grid3Sql = "select '' as O,Form_Desc as City_Name,Form_Code  as code from TaxForms order by Form_Desc"
'                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
'                Case "All"
'                    GridSel(3).Visible = False: Check1(3).Visible = False
'            End Select
        End If
    Case Date1, Date2
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function

'******* Fuctions **********

Private Sub Global_Grid()
Dim I As Integer, Cnt As Integer

Pic.top = Me.top - Pic.width - 10
BTNPRINT.left = (Pic.width - (BTNPRINT.width + BTNEXIT.width)) / 2: BTNPRINT.top = Pic.top + 10

BTNEXIT.left = BTNPRINT.left + BTNPRINT.width: BTNEXIT.top = Pic.top + 10

FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75

FGrid.Rows = 10  '5
FGrid.Cols = 3
FGrid.FixedCols = 1
FGrid.ColWidth(0) = 2200
FGrid.ColWidth(1) = 2000
FGrid.ColWidth(2) = 0
FGrid.ColAlignment(1) = flexAlignLeftCenter
For I = 0 To FGrid.Rows - 1
    FGrid.RowHeight(I) = 0
Next
Ini_Grid
For I = 1 To 4
    If GridSel(I).Visible = True Then Cnt = Cnt + 1
Next
'FGrid.Height = (((mLastRow - mFirstRow) + 1) * PubGridRowHeight) + 500
FGrid.height = (((mLastRow + 1) - mFirstRow) * PubGridRowHeight) + 500
Select Case mHelpGridNo
Case 0
    FGrid.top = 1000
Case 1
    GridSel(1).left = (Me.width - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    GridSel(1).height = Me.height - FGrid.height - Pic.height - 1200
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
Case 2
    GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    GridSel(1).height = Me.height - FGrid.height - Pic.height - 1200
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
    
    GridSel(2).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(2).top = FGrid.top + FGrid.height + 500
    GridSel(2).height = Me.height - FGrid.height - Pic.height - 1200
    Check1(2).top = GridSel(2).top + 20: Check1(2).left = GridSel(2).left + 40
    
Case 3
    GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
   
    GridSel(3).left = GridSel(1).left
    GridSel(3).top = GridSel(1).top + GridSel(1).height + 500
    Check1(3).top = GridSel(3).top + 20: Check1(3).left = GridSel(3).left + 40
    
    GridSel(2).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(2).top = FGrid.top + FGrid.height + 500
    GridSel(2).height = GridSel(1).height + GridSel(2).height + 500
    Check1(2).top = GridSel(2).top + 20: Check1(2).left = GridSel(2).left + 40
    
Case 4
    GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
    
    GridSel(2).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(2).top = FGrid.top + FGrid.height + 500
    Check1(2).top = GridSel(2).top + 20: Check1(2).left = GridSel(2).left + 40
    
    GridSel(3).left = GridSel(1).left
    GridSel(3).top = GridSel(1).top + GridSel(1).height + 500
    Check1(3).top = GridSel(3).top + 20: Check1(3).left = GridSel(3).left + 40
    
    GridSel(4).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(4).top = GridSel(1).top + GridSel(1).height + 500
    Check1(4).top = GridSel(4).top + 20: Check1(4).left = GridSel(4).left + 40

End Select
End Sub
Private Sub Grid_Hide()
If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub FGrid_DblClick()
    Select Case FGrid.Row
        Case Date1, Date2, List1, List2, List3, Cat1, Cat2, Cat3, Cat4, Cat5
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
Dim I As Integer
    Select Case FGrid.Row
        Case Cat1, Cat2, Cat3, Cat4, Cat5
            Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
        Case Date1, Date2, List1, List2, List3
            Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If KeyCode = vbKeyUp And Val(FGrid.Tag) = mFirstRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = mLastRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    FGrid.TextMatrix(FGrid.Row, 2) = ""
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Row
        Case Date1, Date2, List1, List2, List3, Cat1, Cat2, Cat3, Cat4, Cat5
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
   FGrid.CellBackColor = CellBackColEnter
   Grid_Hide
   TxtGrid(0).Visible = False
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Function FillString(GridArray As Variant, Gridindex As Integer, DataType As Byte) As String
On Error GoTo ELoop
Dim ac_str As String
Dim I As Integer
Dim GridRow As Integer
    ac_str = ""
    For I = 0 To UBound(GridArray)
        If GridArray(I) = 0 Then GoTo NXT:
        GridRow = GridArray(I)
        If GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü" Then
                If DataType = 0 Then
                   ac_str = ac_str + IIf(ac_str = "", GridSel(Gridindex).TextMatrix(GridRow, 2), "," + GridSel(Gridindex).TextMatrix(GridRow, 2))
                ElseIf DataType = 1 Then
                   ac_str = ac_str + IIf(ac_str = "", "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'", "," + "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'")
                End If
            GridSel(Gridindex).TextMatrix(GridRow, 0) = ""
        Else
            GridArray(I) = 0
        End If
NXT:
    Next
    For I = 0 To UBound(GridArray)
        GridRow = GridArray(I)
        If GridArray(I) <> 0 Then
            GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü"
        End If
    Next
'    Erase GridArray
'    ReDim Preserve GridArray(0)
'    GridArray(0) = 0
    If ac_str = "" Then
        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
        GridSel(Gridindex).SetFocus
'        RepPrint = False
        Exit Function
    End If
    FillString = ac_str
    Exit Function
ELoop:
    RepPrint = False
    MsgBox err.Description
End Function

Private Sub TxtKeyDown()
Dim I As Integer
    If FGrid.Row = mLastRow Then SendKeysA vbKeyTab, True: Exit Sub
    For I = FGrid.Row To FGrid.Rows - 1
         If FGrid.RowHeight(I + 1) <> 0 Then FGrid.Row = I + 1: Exit For
    Next
End Sub
Private Sub GridInitialise(Gridindex As Integer, GridSql As String)
Dim Index As Integer
Index = Gridindex
If Index = 1 Then
    Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
    RsGrid1.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid1
'    GridSel(Index).top = G1Top: GridSel(Index).left = G1left
    ReDim Preserve GridRow1(0)
    GridRow1(0) = 0
End If
If Index = 2 Then
    Set RsGrid2 = New ADODB.Recordset: RsGrid2.CursorLocation = adUseClient
    RsGrid2.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
'    GridSel(Index).top = G2Top: GridSel(Index).left = G2left
    ReDim Preserve GridRow2(0)
    GridRow2(0) = 0
End If
If Index = 3 Then
    Set RsGrid3 = New ADODB.Recordset: RsGrid3.CursorLocation = adUseClient
    RsGrid3.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid3
'    GridSel(Index).top = G3Top: GridSel(Index).left = G3left
        ReDim Preserve GridRow3(0)
        GridRow3(0) = 0
End If
If Index = 4 Then
    Set RsGrid4 = New ADODB.Recordset: RsGrid4.CursorLocation = adUseClient
    RsGrid4.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid4
'    GridSel(Index).top = G4Top: GridSel(Index).left = G4left
    ReDim Preserve GridRow4(0)
    GridRow4(0) = 0
End If
GridSel(Index).height = 1700
GridSel(Index).Visible = True: GridSel(Index).Enabled = False: Check1(Index).Visible = True
GridSel(Index).width = 5200: GridSel(Index).ColWidth(0) = 600: GridSel(Index).ColWidth(2) = 0: GridSel(Index).ColWidth(1) = 4000
'Check1(Index).top = GridSel(Index).top + 20: Check1(Index).left = GridSel(Index).left + 40
Check1(Index).width = 580: Check1(Index).height = GridSel(Index).RowHeight(0) + 20: Check1(Index).Value = Checked
End Sub

Private Sub Ini_Grid()
'Date1 , Date2, List1, List1, List2, List3
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
 Dim sitecond As String
 If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where  site_code='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
Select Case GRepFormName
    Case DailyActivity, AppointMent, DailySaleRep, SalesTracRep, FinTracRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubLoginDate - 1
            .TextMatrix(Date2, 1) = PubLoginDate - 1
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,E.Emp_Name As SalesExecutive,E.Emp_Code As Code from Emp_Mast E Where Emp_Type=0 Order by E.Emp_Name"
        GridInitialise 2, Grid2Sql
        
    Case CallStatus
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Call Status": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubLoginDate - 1
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Hot"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,E.Emp_Name As SalesExecutive,E.Emp_Code As Code from Emp_Mast E Where Emp_Type=0 Order by E.Emp_Name"
        GridInitialise 2, Grid2Sql
        
    Case CaseAnalysis
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Include Got/Lost Calls": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Customer ExistY/N": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubLoginDate - 1
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Yes"
            .TextMatrix(List2, 1) = "No"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,E.Emp_Name As SalesExecutive,E.Emp_Code As Code from Emp_Mast E Where Emp_Type=0 Order by E.Emp_Name"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "Select ' ' as O, PSG.Name + PSG.NSuffix + City.CityName as ProspectiveCustomerName,PSG.Cust_Code as Code " & _
            " from ProspectiveCust PSG left join City on PSG.CityCode=City.CityCode " & _
            " where Cust_code in (Select distinct Party_Code from Visits where ProspectiveCust_SubGroup=0) " & _
            " Order By PSG.Name"
          GridInitialise 2, Grid2Sql
        
    Case PipeLineRep
        With FGrid
            .TextMatrix(Date1, 0) = "As on Date ": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate - 1
        End With
        mFirstRow = Date1: mLastRow = Date1
        
        
    Case ProfPurRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubLoginDate - 1
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = List2 ': mHelpGridNo = 2
'        Grid1Sql = "select '' as O,E.Emp_Name As SalesExecutive,E.Emp_Code As Code from Emp_Mast E Where Emp_Type=0 Order by E.Emp_Name"
'        GridInitialise 1, Grid1Sql
'        Grid2Sql = "Select ' ' as O, PSG.Name & PSG.NSuffix & City.CityName as ProspectiveCustomerName,PSG.Cust_Code as Code " & _
            " from ProspectiveCust PSG left join City on PSG.CityCode=City.CityCode " & _
            " where Cust_code in (Select distinct Party_Code from Visits where ProspectiveCust_SubGroup=0) " & _
            " Order By PSG.Name"
'          GridInitialise 2, Grid2Sql

'Case VehStkHold    'vijay WKS 16/11/02
'        With FGrid
'            .TextMatrix(Date1, 0) = "As On Date": .RowHeight(Date1) = GridRowHeight
''            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
'            .TextMatrix(List1, 0) = "Stock All/Current": .RowHeight(List1) = GridRowHeight
''            .TextMatrix(List2, 0) = "Description Of Labour": .RowHeight(List2) = GridRowHeight
'
''            .TextMatrix(Date1, 1) = PubStartDate
'            .TextMatrix(Date1, 1) = PubLoginDate
'            .TextMatrix(List1, 1) = "All"
''            .TextMatrix(List2, 1) = "Standard"
'        End With
'        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
'        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
'        GridInitialise 1, Grid1Sql
'        Grid2Sql = "select '' as O,Subgroup.Name As Party_Name,Subgroup.SubCode As Code from SubGroup order by SubGroup.Name"
'          GridInitialise 2, Grid2Sql
'        Grid3Sql = "select '' as O,Model.Model_Desc As Model_Description,Model.Model As Code from Model order by Model.Model_Desc"
'          GridInitialise 3, Grid3Sql
'        Grid4Sql = "select '' as O,BMS.BMS_Name As Category,BMS.BMS_Code As Code from BMS order by BMS.BMS_Name"
'          GridInitialise 4, Grid4Sql
'
'
End Select
End Sub
Public Function IsNotBlank(FieldRow As Integer, FieldCaption As String) As Boolean
    If FGrid.TextMatrix(FieldRow, 1) = "" Then
        MsgBox FieldCaption & " Should not be Blank.", vbInformation, "Validation Check"
        FGrid.SetFocus
        FGrid.Row = FieldRow
        FGrid.Col = 1
        IsNotBlank = False
    Else
        IsNotBlank = True
    End If
End Function

Private Sub Formulas()
On Error GoTo ELoop
Dim I As Integer
Select Case GRepFormName
    Case DailyActivity, AppointMent, CaseAnalysis, ProfPurRep, SalesTracRep, FinTracRep
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
    Case PipeLineRep
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'As on :' + '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Function DailyActivityProc() As Boolean
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Function
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Function

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Function
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Function
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and V.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and V.Site_Code ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and E.Emp_Code in (" & GridString2 & ")"
    If PubBackEnd = "A" Then
        mQry = "Select V.VisitDate,V.Rep_Code,V.SrlNo,V.Div_Code,V.Site_Code,V.ProspectiveCust_SubGroup,V.Party_Code,V.NewEnquiry,V.Trf_YN," & _
            " V.TrfFrom_RepCode,V.Visit_Call,V.Meet_TimeFrom,V.Meet_TimeTo,V.REMARK1,V.OBJECTIVE,V.REMARK2,V.Call_Status,V.Call_Status," & _
            " V.EXPENCE,V.EXPREMARK,V.U_Name,V.U_EntDt,V.U_AE, E.Emp_Name, " & _
            " Switch(V.Call_Status=0,'Cold',V.Call_Status=1,'Warm',V.Call_Status=2,'Hot',V.Call_Status=3,'Nill') as CallStatus, " & _
            " Switch(V.Visit_Call=0,'Visit',V.Visit_Call=1,'Call') as VisitCall, " & _
            " Switch(V.ProspectiveCust_SubGroup=0,PSG.Name+PSG.NSuffix,V.ProspectiveCust_SubGroup=1,SG.Name) as PartyName, " & _
            " VObj.ObjDesc, E1.Emp_Name as TrfRepName " & _
            " FROM (((((Visits as V LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
            " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
            " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
            " Left JOIN Site ON V.Site_Code = Site.Site_Code) " & _
            " Left JOIN VisitObjective VObj ON V.OBJECTIVE = VObj.ObjCode) " & _
            " Left JOIN Emp_Mast E1 ON V.TrfFrom_RepCode = E1.Emp_Code " & _
            " WHERE (V.VisitDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
            " AND (V.VisitDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " ) "
    ElseIf PubBackEnd = "S" Then
        mQry = "Select V.VisitDate,V.Rep_Code,V.SrlNo,V.Div_Code,V.Site_Code,V.ProspectiveCust_SubGroup,V.Party_Code,V.NewEnquiry,V.Trf_YN," & _
            " V.TrfFrom_RepCode,V.Visit_Call,V.Meet_TimeFrom,V.Meet_TimeTo,V.REMARK1,V.OBJECTIVE,V.REMARK2,V.Call_Status,V.Call_Status," & _
            " V.EXPENCE,V.EXPREMARK,V.U_Name,V.U_EntDt,V.U_AE, E.Emp_Name, " & _
            " Case V.Call_Status When 0 Then 'Cold' When 1 Then 'Warm'When 2 Then 'Hot' When 3 Then 'Nill' End as CallStatus, " & _
            " Case V.Visit_Call When 0 Then 'Visit' When 1 Then 'Call' End as VisitCall, " & _
            " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Name+PSG.NSuffix When 1 Then SG.Name End as PartyName, " & _
            " VObj.ObjDesc, E1.Emp_Name as TrfRepName " & _
            " FROM (((((Visits as V LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
            " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
            " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
            " Left JOIN Site ON V.Site_Code = Site.Site_Code) " & _
            " Left JOIN VisitObjective VObj ON V.OBJECTIVE = VObj.ObjCode) " & _
            " Left JOIN Emp_Mast E1 ON V.TrfFrom_RepCode = E1.Emp_Code " & _
            " WHERE (V.VisitDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
            " AND (V.VisitDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " ) "
    End If
    mQry = mQry + Condstr
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
    RepName = "DailyActivity"
    RepTitle = UCase(Me.CAPTION)
    DailyActivityProc = True
    Exit Function
ELoop:
    DailyActivityProc = False
    MsgBox err.Description
End Function
Private Function DailySaleProc() As Boolean
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Function
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Function

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Function
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Function
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and V.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and V.Site_Code ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Emp_Mast.Emp_Code in (" & GridString2 & ")"
    
    mQry = "Select PC.Name,PC.PhoneResi,VS.Model,V.Call_Status,Emp_Mast.Emp_Name,V.Meet_TimeFrom,V.Remark1,VS.Got_Lost,VOL.Name as Lost_Cat,VS.OrdDocID,VS.PurchModel,V.Schemes,V.SalesNos,V.Prices,V.Pamphlets,V.Hoardings,V.Events,V.MediaAds,V.Misc from " & _
            " (((Visits V Left Join ProspectiveCust PC On V.Party_Code=PC.Cust_Code)" & _
            " Left Join Veh_SubgroupQuot VS on V.Party_Code=VS.PartyCode) " & _
            " Left Join Emp_Mast on V.Rep_Code=Emp_Mast.Emp_Code)" & _
            " Left Join Veh_OrdLostCatg VOL on VS.Lost_Cat=VOL.Code" & _
            " WHERE (V.VisitDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
            " AND (V.VisitDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " ) "
     mQry = mQry + Condstr
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
    RepName = "DailySaleRep"
    RepTitle = UCase(Me.CAPTION)
    DailySaleProc = True
    Exit Function
ELoop:
    DailySaleProc = False
    MsgBox err.Description
End Function
Private Function SalesTracRepProc() As Boolean
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Function
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Function

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Function
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Function
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and V.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and V.Site_Code ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Emp_Mast.Emp_Code in (" & GridString2 & ")"
    
    mQry = "Select PC.Name,PC.PhoneResi,Area.AreaName,VS.Model,Emp_Mast.Emp_Name,VS.Got_Lost,VS.OrdDocID," & _
           " ColMast.Col_Desc,VO.OrdDocId,VO.Ord_Date,VO.Inv_Date,CF.FinName from " & _
            " (((((((Visits V Left Join ProspectiveCust PC On V.Party_Code=PC.Cust_Code)" & _
            " Left Join Veh_SubgroupQuot VS on V.Party_Code=VS.PartyCode) " & _
            " Left Join Emp_Mast on V.Rep_Code=Emp_Mast.Emp_Code)" & _
            " Left Join Area on Pc.Area=Area.AreaCode)" & _
            " Left Join Veh_Order VO on VS.OrdDocId=Vo.OrdDocId)" & _
            " Left Join ColMast on VO.Colour_Code=ColMast.Col_Code)" & _
            " Left Join ContractFinance CF on VO.FB_Code=CF.FinCode)" & _
            " Left Join Veh_OrdLostCatg VOL on VS.Lost_Cat=VOL.Code" & _
            " WHERE (V.VisitDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
            " AND (V.VisitDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " ) "
     mQry = mQry + Condstr
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    If GRepFormName = SalesTracRep Then
        RepName = "SalesTracRep"
        RepTitle = UCase(Me.CAPTION)
    ElseIf GRepFormName = FinTracRep Then
        RepName = "FinTrackRep"
        RepTitle = UCase(Me.CAPTION)
    End If
    
    SalesTracRepProc = True
    Exit Function
ELoop:
    SalesTracRepProc = False
    MsgBox err.Description
End Function
Private Function AppointmentProc() As Boolean
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: GoTo ELoop
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: GoTo ELoop
'    If IsNotBlank(List1, FGrid.TextMatrix(List1, 1)) = False Then RepPrint = False: GoTo ELoop
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then GoTo ELoop
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then GoTo ELoop
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and V.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and V.Site_Code ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and E.Emp_Code in (" & GridString2 & ")"
    
'Sales Executive
'APT.DATE PARTY NAME & ADDRESS    PHONE   LAST DATE   STATUS  LAST VISIT REMARKS
'-------------------------------------------------------------------------------
    If PubBackEnd = "A" Then
        mQry = "Select V.VisitDate,V.Rep_Code,V.SrlNo,V.Div_Code,V.Site_Code,V.ProspectiveCust_SubGroup,V.Party_Code,V.NewEnquiry,V.Trf_YN," & _
             " V.TrfFrom_RepCode,V.Visit_Call,V.Meet_TimeFrom,V.Meet_TimeTo,V.REMARK1,V.OBJECTIVE,V.REMARK2,V.Call_Status,V.Call_Status," & _
             " V.EXPENCE,V.EXPREMARK,V.U_Name,V.U_EntDt,V.U_AE, E.Emp_Name, " & _
             " Switch(V.Call_Status=0,'Cold',V.Call_Status=1,'Warm',V.Call_Status=2,'Hot',V.Call_Status=3,'Nill') as CallStatus, " & _
             " Switch(V.Visit_Call=0,'Visit',V.Visit_Call=1,'Call') as VisitCall, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,PSG.Name+PSG.NSuffix,V.ProspectiveCust_SubGroup=1,PSG.Name) as PartyName, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,PSG.ConPerson,V.ProspectiveCust_SubGroup=1,SG.ConPerson) as PConPerson, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,PSG.Add1,V.ProspectiveCust_SubGroup=1,SG.Add1) as PAdd1, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,PSG.Add2,V.ProspectiveCust_SubGroup=1,SG.Add2) as PAdd2, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,PSG.Add3,V.ProspectiveCust_SubGroup=1,SG.Add3) as PAdd3, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,PSG.PhoneOff,V.ProspectiveCust_SubGroup=1,SG.Phone) as PPhone, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,PSG.Mobile,V.ProspectiveCust_SubGroup=1,SG.Mobile) as PMobile, " & _
             " Switch(V.ProspectiveCust_SubGroup=0,City.CityName,V.ProspectiveCust_SubGroup=1,City1.CityName) as CitName " & _
             " FROM (((((Visits as V LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
             " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
             " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
             " Left JOIN Site ON V.Site_Code = Site.Site_Code) " & _
             " Left JOIN City ON PSG.CityCode=City.CityCode) " & _
             " Left JOIN City City1 ON SG.CityCode=City1.CityCode " & _
             " WHERE (V.Next_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy hh:mm")) & ") " & _
             " AND (V.Next_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy 23:59")) & " ) "
    ElseIf PubBackEnd = "S" Then
        mQry = "Select V.VisitDate,V.Rep_Code,V.SrlNo,V.Div_Code,V.Site_Code,V.ProspectiveCust_SubGroup,V.Party_Code,V.NewEnquiry,V.Trf_YN," & _
             " V.TrfFrom_RepCode,V.Visit_Call,V.Meet_TimeFrom,V.Meet_TimeTo,V.REMARK1,V.OBJECTIVE,V.REMARK2,V.Call_Status,V.Call_Status," & _
             " V.EXPENCE,V.EXPREMARK,V.U_Name,V.U_EntDt,V.U_AE, E.Emp_Name, " & _
             " Case V.Call_Status When 0 Then 'Cold' When 1 Then 'Warm' When 2 Then 'Hot' When 3 Then 'Nill' End as CallStatus, " & _
             " Case V.Visit_Call When 0 Then 'Visit' When 1 Then 'Call' End as VisitCall, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Name+PSG.NSuffix When 1 Then PSG.Name End  as PartyName, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then PSG.ConPerson When 1 Then SG.ConPerson End as PConPerson, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Add1 When 1 Then SG.Add1 End as PAdd1, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Add2 When 1 then SG.Add2 End as PAdd2, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Add3 When 1 Then SG.Add3 End as PAdd3, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then PSG.PhoneOff When 1 Then SG.Phone End  as PPhone, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Mobile When 1 Then SG.Mobile End as PMobile, " & _
             " Case V.ProspectiveCust_SubGroup When 0 Then City.CityName When 1 Then City1.CityName End as CitName " & _
             " FROM (((((Visits as V LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
             " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
             " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
             " Left JOIN Site ON V.Site_Code = Site.Site_Code) " & _
             " Left JOIN City ON PSG.CityCode=City.CityCode) " & _
             " Left JOIN City City1 ON SG.CityCode=City1.CityCode " & _
             " WHERE (V.Next_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy hh:mm")) & ") " & _
             " AND (V.Next_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy 23:59")) & " ) "
    End If
    mQry = mQry + Condstr
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
    RepName = "Appointments"
    RepTitle = UCase(Me.CAPTION)
    AppointmentProc = True
    Exit Function
ELoop:
    AppointmentProc = False
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function

Private Function CallStatusProc() As Boolean
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then GoTo ELoop
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then GoTo ELoop
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 1)) = False Then GoTo ELoop

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then GoTo ELoop
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then GoTo ELoop
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and V.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and V.Site_Code ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and E.Emp_Code in (" & GridString2 & ")"
    Select Case FGrid.TextMatrix(List1, 1)
        Case "Cold"
            Condstr = Condstr & " and V.Call_Status=0"
        Case "Warm"
            Condstr = Condstr & " and V.Call_Status=1"
        Case "Hot"
            Condstr = Condstr & " and V.Call_Status=2"
        Case "Nill"
            Condstr = Condstr & " and V.Call_Status=3"
    End Select
'Sales Executive
'CUSTOMER   PHONE     VIS/CAL   LAST_DT   LAST Status REMARKS
'START_DT MODEL STATUS

'    mQRY = "Select V., E.Emp_Name, " & _
        " Switch(V.Call_Status=0,'Cold',V.Call_Status=1,'Warm',V.Call_Status=2,'Hot',V.Call_Status=3,'Nill') as CallStatus, " & _
        " Switch(V.Visit_Call=0,'Visit',V.Visit_Call=1,'Call') as VisitCall, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,PSG.Name+PSG.NSuffix,V.ProspectiveCust_SubGroup=1,SG.Name) as PartyName, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,PSG.PhoneOff,V.ProspectiveCust_SubGroup=1,SG.Phone) as PPhone, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,PSG.Mobile,V.ProspectiveCust_SubGroup=1,SG.Mobile) as PMobile " & _
        " FROM (((Visits LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
        " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
        " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
        " Left JOIN Site ON V.Site_Code = Site.Site_Code " & _
        " WHERE (V.Next_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#) " & _
        " AND (V.Next_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# ) "

'    mQRY = "Select VSGQ.* " & _
        " FROM Veh_SubGroupQuot VSGQ " & _
        " WHERE (VSGQ.Next_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#) " & _
        " AND (VSGQ.Next_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# ) "
'      modifiled by vikash /shakher
    If PubBackEnd = "A" Then
       mQry = "Select V.VisitDate,V.Rep_Code,V.SrlNo,V.Div_Code,V.Site_Code,V.ProspectiveCust_SubGroup,V.Party_Code,V.NewEnquiry,V.Trf_YN," & _
        " V.TrfFrom_RepCode,V.Visit_Call,V.Meet_TimeFrom,V.Meet_TimeTo,V.REMARK1,V.OBJECTIVE,V.REMARK2,V.Call_Status," & _
        " V.EXPENCE,V.EXPREMARK,V.U_Name,V.U_EntDt,V.U_AE, E.Emp_Name, " & _
        " Switch(V.Call_Status=0,'Cold',V.Call_Status=1,'Warm',V.Call_Status=2,'Hot',V.Call_Status=3,'Nill') as CallStatus, " & _
        " Switch(V.Visit_Call=0,'Visit',V.Visit_Call=1,'Call') as VisitCall, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,PSG.Name+PSG.NSuffix,V.ProspectiveCust_SubGroup=1,SG.Name) as PartyName, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,PSG.PhoneOff,V.ProspectiveCust_SubGroup=1,SG.Phone) as PPhone, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,PSG.Mobile,V.ProspectiveCust_SubGroup=1,SG.Mobile) as PMobile " & _
        " FROM (((Visits as V LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
        " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
        " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
        " Left JOIN Site ON V.Site_Code = Site.Site_Code " & _
        " WHERE (V.Next_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy hh:mm")) & ") " & _
        " AND (V.Next_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy 23:59")) & " ) "
    ElseIf PubBackEnd = "S" Then
       mQry = "Select V.VisitDate,V.Rep_Code,V.SrlNo,V.Div_Code,V.Site_Code,V.ProspectiveCust_SubGroup,V.Party_Code,V.NewEnquiry,V.Trf_YN," & _
        " V.TrfFrom_RepCode,V.Visit_Call,V.Meet_TimeFrom,V.Meet_TimeTo,V.REMARK1,V.OBJECTIVE,V.REMARK2,V.Call_Status," & _
        " V.EXPENCE,V.EXPREMARK,V.U_Name,V.U_EntDt,V.U_AE, E.Emp_Name, " & _
        " Case V.Call_Status When 0 Then 'Cold' When 1 Then 'Warm' When 2 Then 'Hot' When 3 Then 'Nill' End  as CallStatus, " & _
        " Case V.Visit_Call When 0 Then 'Visit' When 1 Then 'Call' End  as VisitCall, " & _
        " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Name+PSG.NSuffix When 1 Then SG.Name End as PartyName, " & _
        " Case V.ProspectiveCust_SubGroup When 0 Then PSG.PhoneOff When 1 Then SG.Phone End  as PPhone, " & _
        " Case V.ProspectiveCust_SubGroup When 0 Then PSG.Mobile When 1 Then SG.Mobile End  as PMobile " & _
        " FROM (((Visits as V LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
        " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
        " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
        " Left JOIN Site ON V.Site_Code = Site.Site_Code " & _
        " WHERE (V.Next_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy hh:mm")) & ") " & _
        " AND (V.Next_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy 23:59")) & " ) "
    End If

    mQry = mQry + Condstr
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
    RepName = "CallStatus"
    RepTitle = UCase(Me.CAPTION)
    CallStatusProc = True
    Exit Function
ELoop:
    CallStatusProc = False
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function
Private Function PipeLineRepProc() As Boolean
On Error GoTo ELoop
Dim mQryC0$, mQryC1$, mQryC2$, mQryC3$, Condstr$, Rst As ADODB.Recordset

'As on date :
'                                          <----LCV----- >                         <--M & HCV's- >                         <----207----- >
'Area Name RepName  Actual  New  FollowUp  c0  c1  c2  c3   Actual  New  FollowUp c0  c1  c2  c3   Actual  New  FollowUp c0  c1  c2  c3
'
'Area-wise Total
'Grand Total

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: GoTo ELoop
    
mQryC0 = "select distinct 'C0' as Org,V.NewEnquiry, 1 as C0, 0 as C1, 0 as C2, 0 as C3, V.VisitDate as Date1,V.Rep_Code,M.Vehicle_Type,E.Emp_Name,Area.AreaName " & _
    " from ((((Visits V left join Veh_SubGroupQuot VSQ on V.Rep_Code+V.Party_Code=VSQ.Rep_Code+VSQ.PartyCode)" & _
    " left join Model M on VSQ.Model=M.MODEL) " & _
    " left join Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
    " left join ProspectiveCust PC on V.Party_Code=PC.Cust_Code) " & _
    " left join Area on PC.AREA=Area.AreaCode " & _
    " where V.Div_Code='" & PubDivCode & "' and V.Visit_Call=1 and VSQ.Got_Lost='' " & _
    " and V.VisitDate = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " "

mQryC1 = "select distinct 'C1' as Org,0 as NewEnquiry, 0 as C0, 1 as C1, 0 as C2, 0 as C3,VQ.V_Date as Date1,VQ.Rep_Code,M.Vehicle_Type,E.Emp_Name,Area.AreaName " & _
    " from ((((Veh_Quot1 VQ1 left join Veh_Quot VQ on VQ1.DocId=VQ.DocID) " & _
    " left join Model M on VQ1.Model=M.MODEL) " & _
    " left join Emp_Mast E on VQ.Rep_Code=E.Emp_Code) " & _
    " left join ProspectiveCust PC on VQ.Party_Code=PC.Cust_Code) " & _
    " left join Area on PC.AREA=Area.AreaCode " & _
    " where left(VQ1.DocID,1)='" & PubDivCode & "' " & _
    " and VQ.V_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " "

mQryC2 = "select Distinct 'C2' as Org,0 as NewEnquiry, 0 as C0, 0 as C1, 1 as C2, 0 as C3,VO.Ord_Date as Date1,VO.Rep_Code,M.Vehicle_Type,E.Emp_Name,Area.AreaName " & _
    " from ((Veh_Order VO left join Model M on VO.Model=M.MODEL) " & _
    " left join Emp_Mast E on VO.Rep_Code=E.Emp_Code) " & _
    " left join Area on VO.AREA=Area.AreaCode " & _
    " where left(VO.OrdDocId,1)='" & PubDivCode & "' and " & _
    " (VO.Inv_DocID='' or VO.Inv_DocID Is Null) and " & _
    " (VO.OrdDocid in (Select distinct Ord_DocId from Rect where Vehicle_YN=1 and AMOUNT >0 and DrCr='C')) " & _
    " and VO.Ord_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " "

mQryC3 = "select distinct 'C3' as Org,0 as NewEnquiry, 0 as C0, 0 as C1, 0 as C2, 1 as C3,VO.Inv_Date as Date1,VO.Rep_Code,M.Vehicle_Type,E.Emp_Name,Area.AreaName " & _
    " from ((Veh_Order VO left join Model M on VO.Model=M.MODEL) " & _
    " left join Emp_Mast E on VO.Rep_Code=E.Emp_Code) " & _
    " left join Area on VO.AREA=Area.AreaCode " & _
    " where left(VO.Inv_DocId,1)='" & PubDivCode & "' and VO.Inv_Docid<>''" & _
    " and VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " "

GSQL = mQryC0 & " Union All " & mQryC1 & " Union All " & mQryC2 & " Union All " & mQryC3
    
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (GSQL), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
        'Create temp table
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "Date1", adDate, 7, adFldIsNullable
            .Fields.Append "AreaCode", adChar, 3, adFldIsNullable
            .Fields.Append "AreaName", adChar, 15, adFldIsNullable
            .Fields.Append "RepCode", adChar, 4, adFldIsNullable
            .Fields.Append "RepName", adChar, 40, adFldIsNullable
            .Fields.Append "Actual", adInteger, 4, adFldIsNullable
            .Fields.Append "LCV_New", adInteger, 4, adFldIsNullable
            .Fields.Append "LCV_FollowUp", adInteger, 4, adFldIsNullable
            .Fields.Append "LCV_C0", adInteger, 4, adFldIsNullable
            .Fields.Append "LCV_C1", adInteger, 4, adFldIsNullable
            .Fields.Append "LCV_C2", adInteger, 4, adFldIsNullable
            .Fields.Append "LCV_C3", adInteger, 4, adFldIsNullable
            .Fields.Append "MHCV_New", adInteger, 4, adFldIsNullable
            .Fields.Append "MHCV_FollowUp", adInteger, 4, adFldIsNullable
            .Fields.Append "MHCV_C0", adInteger, 4, adFldIsNullable
            .Fields.Append "MHCV_C1", adInteger, 4, adFldIsNullable
            .Fields.Append "MHCV_C2", adInteger, 4, adFldIsNullable
            .Fields.Append "MHCV_C3", adInteger, 4, adFldIsNullable
            .Fields.Append "207_New", adInteger, 4, adFldIsNullable
            .Fields.Append "207_FollowUp", adInteger, 4, adFldIsNullable
            .Fields.Append "207_C0", adInteger, 4, adFldIsNullable
            .Fields.Append "207_C1", adInteger, 4, adFldIsNullable
            .Fields.Append "207_C2", adInteger, 4, adFldIsNullable
            .Fields.Append "207_C3", adInteger, 4, adFldIsNullable
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
        End With
        'temp table created
        
        Do While Rst.EOF = False
            With RstRep
                .AddNew
                .Fields("Date1") = Rst!Date1
                .Fields("AreaName") = Rst!AreaName
                .Fields("RepName") = Rst!Emp_Name
                .Fields("Actual") = 0
                If Trim(UCase(Rst!Vehicle_Type)) = "LCV" Then
                    .Fields("LCV_New") = Rst!NewEnquiry
                    .Fields("LCV_FollowUp") = IIf(Rst!NewEnquiry = 0, 1, 0)
                    .Fields("LCV_C0") = Rst!c0
                    .Fields("LCV_C1") = Rst!c1
                    .Fields("LCV_C2") = Rst!c2
                    .Fields("LCV_C3") = Rst!c3
                    .Fields("207_New") = 0
                    .Fields("207_FollowUp") = 0
                    .Fields("207_C0") = 0
                    .Fields("207_C1") = 0
                    .Fields("207_C2") = 0
                    .Fields("207_C3") = 0
                    .Fields("MHCV_New") = 0
                    .Fields("MHCV_FollowUp") = 0
                    .Fields("MHCV_C0") = 0
                    .Fields("MHCV_C1") = 0
                    .Fields("MHCV_C2") = 0
                    .Fields("MHCV_C3") = 0

                ElseIf Trim(UCase(Rst!Vehicle_Type)) = "207" Then
                    .Fields("LCV_FollowUp") = 0
                    .Fields("LCV_C0") = 0
                    .Fields("LCV_C1") = 0
                    .Fields("LCV_C2") = 0
                    .Fields("LCV_C3") = 0
                    .Fields("207_New") = Rst!NewEnquiry
                    .Fields("207_FollowUp") = IIf(Rst!NewEnquiry = 0, 1, 0)
                    .Fields("207_C0") = Rst!c0
                    .Fields("207_C1") = Rst!c1
                    .Fields("207_C2") = Rst!c2
                    .Fields("207_C3") = Rst!c3
                    .Fields("LCV_New") = 0
                    .Fields("MHCV_New") = 0
                    .Fields("MHCV_FollowUp") = 0
                    .Fields("MHCV_C0") = 0
                    .Fields("MHCV_C1") = 0
                    .Fields("MHCV_C2") = 0
                    .Fields("MHCV_C3") = 0
                Else
                    .Fields("LCV_New") = 0
                    .Fields("LCV_FollowUp") = 0
                    .Fields("LCV_C0") = 0
                    .Fields("LCV_C1") = 0
                    .Fields("LCV_C2") = 0
                    .Fields("LCV_C3") = 0
                    .Fields("207_New") = 0
                    .Fields("207_FollowUp") = 0
                    .Fields("207_C0") = 0
                    .Fields("207_C1") = 0
                    .Fields("207_C2") = 0
                    .Fields("207_C3") = 0
                    .Fields("MHCV_New") = Rst!NewEnquiry
                    .Fields("MHCV_FollowUp") = IIf(Rst!NewEnquiry = 0, 1, 0)
                    .Fields("MHCV_C0") = Rst!c0
                    .Fields("MHCV_C1") = Rst!c1
                    .Fields("MHCV_C2") = Rst!c2
                    .Fields("MHCV_C3") = Rst!c3
                End If
                .Update
            End With
            Rst.MoveNext
        Loop
    Set Rst = Nothing
'    Set RstRep = New Recordset
'    RstRep.CursorLocation = adUseClient
'    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
    RepName = "PipeLineRep"
    RepTitle = UCase(Me.CAPTION)
    PipeLineRepProc = True
    Exit Function
ELoop:
    PipeLineRepProc = False
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function

Private Function ProfPurRepProc() As Boolean
On Error GoTo ELoop
Dim mQryC1$, mQryC2$, mQryC3$, Condstr$, Rst As ADODB.Recordset
Dim mAQryC1$, mAQryC2$, mAQryC3$, rstArea As ADODB.Recordset, mQryArea$
'Date From dd/mm/yyyy to dd/mm/yyyy
'                      <-Area-1- >   <-Area-2- >
'Profession            c1  c2  c3   c1  c2  c3
'   Purpose
'   Purpose-wise Total
'Profession-wise Total
'Grand Total

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: GoTo ELoop
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: GoTo ELoop

mQryC1 = "select distinct 'C1' as Org,ProfessionName,PurposeName,0 as NewEnquiry, 0 as C0, 1 as C1, 0 as C2, 0 as C3,VQ.V_Date as Date1,VQ.Rep_Code,M.Vehicle_Type,E.Emp_Name,Area.AreaName " & _
    " from ((((Veh_Quot1 VQ1 left join Veh_Quot VQ on VQ1.DocId=VQ.DocID) " & _
    " left join Model M on VQ1.Model=M.MODEL) " & _
    " left join Emp_Mast E on VQ.Rep_Code=E.Emp_Code) " & _
    " left join ProspectiveCust PC on VQ.Party_Code=PC.Cust_Code) " & _
    " left join Area on PC.AREA=Area.AreaCode " & _
    " left join Profession on PC.Profession=Profession.ProfessionCode) " & _
    " left join Purpose on VQ.Purpose=Purpose.PurposeCode) " & _
    " where left(VQ1.DocID,1)='" & PubDivCode & "' " & _
    " and VQ.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & _
    " and VQ.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

mQryC2 = "select Distinct 'C2' as Org,ProfessionName,PurposeName,0 as NewEnquiry, 0 as C0, 0 as C1, 1 as C2, 0 as C3,VO.Ord_Date as Date1,VO.Rep_Code,M.Vehicle_Type,E.Emp_Name,Area.AreaName " & _
    " from ((Veh_Order VO left join Model M on VO.Model=M.MODEL) " & _
    " left join Emp_Mast E on VO.Rep_Code=E.Emp_Code) " & _
    " left join Area on VO.AREA=Area.AreaCode " & _
    " left join Profession on VO.Profession=Profession.ProfessionCode) " & _
    " left join Purpose on VO.Purpose=Purpose.PurposeCode) " & _
    " where left(VO.OrdDocId,1)='" & PubDivCode & "' and " & _
    " (VO.Inv_DocID='' or isnull(VO.Inv_DocID)) and " & _
    " (VO.OrdDocid in (Select distinct Ord_DocId from Rect where Vehicle_YN=1 and AMOUNT >0 and DrCr='C')) " & _
    " and VO.Ord_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & _
    " and VO.Ord_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

mQryC3 = "select distinct 'C3' as Org,ProfessionName,PurposeName,0 as NewEnquiry, 0 as C0, 0 as C1, 0 as C2, 1 as C3,VO.Inv_Date as Date1,VO.Rep_Code,M.Vehicle_Type,E.Emp_Name,Area.AreaName " & _
    " from ((Veh_Order VO left join Model M on VO.Model=M.MODEL) " & _
    " left join Emp_Mast E on VO.Rep_Code=E.Emp_Code) " & _
    " left join Area on VO.AREA=Area.AreaCode " & _
    " left join Profession on VO.Profession=Profession.ProfessionCode) " & _
    " left join Purpose on VO.Purpose=Purpose.PurposeCode) " & _
    " where left(VO.Inv_DocId,1)='" & PubDivCode & "' and VO.Inv_Docid<>''" & _
    " and VO.Inv_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & _
    " and VO.Inv_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

GSQL = mQryC1 & " Union All " & mQryC2 & " Union All " & mQryC3

mAQryC1 = "select distinct Area.AreaName " & _
    " from ((((Veh_Quot1 VQ1 left join Veh_Quot VQ on VQ1.DocId=VQ.DocID) " & _
    " left join Model M on VQ1.Model=M.MODEL) " & _
    " left join Emp_Mast E on VQ.Rep_Code=E.Emp_Code) " & _
    " left join ProspectiveCust PC on VQ.Party_Code=PC.Cust_Code) " & _
    " left join Area on PC.AREA=Area.AreaCode " & _
    " left join Profession on PC.Profession=Profession.ProfessionCode) " & _
    " left join Purpose on VQ.Purpose=Purpose.PurposeCode) " & _
    " where left(VQ1.DocID,1)='" & PubDivCode & "' " & _
    " and VQ.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & _
    " and VQ.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "# "

mAQryC2 = "select Distinct Area.AreaName " & _
    " from ((Veh_Order VO left join Model M on VO.Model=M.MODEL) " & _
    " left join Emp_Mast E on VO.Rep_Code=E.Emp_Code) " & _
    " left join Area on VO.AREA=Area.AreaCode " & _
    " left join Profession on VO.Profession=Profession.ProfessionCode) " & _
    " left join Purpose on VO.Purpose=Purpose.PurposeCode) " & _
    " where left(VO.OrdDocId,1)='" & PubDivCode & "' and " & _
    " (VO.Inv_DocID='' or isnull(VO.Inv_DocID)) and " & _
    " (VO.OrdDocid in (Select distinct Ord_DocId from Rect where Vehicle_YN=1 and AMOUNT >0 and DrCr='C')) " & _
    " and VO.Ord_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & _
    " and VO.Ord_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

mAQryC3 = "select distinct Area.AreaName " & _
    " from ((Veh_Order VO left join Model M on VO.Model=M.MODEL) " & _
    " left join Emp_Mast E on VO.Rep_Code=E.Emp_Code) " & _
    " left join Area on VO.AREA=Area.AreaCode " & _
    " left join Profession on VO.Profession=Profession.ProfessionCode) " & _
    " left join Purpose on VO.Purpose=Purpose.PurposeCode) " & _
    " where left(VO.Inv_DocId,1)='" & PubDivCode & "' and VO.Inv_Docid<>''" & _
    " and VO.Inv_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & _
    " and VO.Inv_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
mQryArea = mAQryC1 & " Union " & mAQryC2 & " Union " & mAQryC3
    
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (GSQL), GCn, adOpenStatic, adLockReadOnly
'    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: RepPrint = False: Exit Function

    Set rstArea = New Recordset
    rstArea.CursorLocation = adUseClient
    rstArea.Open (mQryArea), GCn, adOpenStatic, adLockReadOnly
   
        'Create temp table
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "Date1", adDate, 7, adFldIsNullable
            .Fields.Append "Profession", adChar, 12, adFldIsNullable
            .Fields.Append "Purpose", adChar, 25, adFldIsNullable
            .Fields.Append "C1_A1", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A1", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A1", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A2", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A2", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A2", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A3", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A3", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A3", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A4", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A4", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A4", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A5", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A5", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A5", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A6", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A6", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A6", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A7", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A7", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A7", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A8", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A8", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A8", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A9", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A9", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A9", adInteger, 3, adFldIsNullable
            
            .Fields.Append "C1_A10", adInteger, 3, adFldIsNullable
            .Fields.Append "C2_A10", adInteger, 3, adFldIsNullable
            .Fields.Append "C3_A10", adInteger, 3, adFldIsNullable
            .Fields.Append "Head1", adChar, 10, adFldIsNullable
            .Fields.Append "Head2", adChar, 10, adFldIsNullable
            .Fields.Append "Head3", adChar, 10, adFldIsNullable
            .Fields.Append "Head4", adChar, 10, adFldIsNullable
            .Fields.Append "Head5", adChar, 10, adFldIsNullable
            .Fields.Append "Head6", adChar, 10, adFldIsNullable
            .Fields.Append "Head7", adChar, 10, adFldIsNullable
            .Fields.Append "Head8", adChar, 10, adFldIsNullable
            .Fields.Append "Head9", adChar, 10, adFldIsNullable
            .Fields.Append "Head10", adChar, 10, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
        End With
        'temp table created
        
        Do While Rst.EOF = False
            With RstRep
                .AddNew
                .Fields("Date1") = Rst!Date1
                .Fields("Profession") = Rst!ProfessionName
                .Fields("Purpose") = Rst!PurposeName
                'Area
                If rstArea.RecordCount > 0 Then
                    If IsNull(Rst!AreaName) Or Rst!AreaName = "" Then
                    Else
                        rstArea.MoveFirst
                        rstArea.FIND ("AreaName='" & Rst!AreaName & "'")
                        If rstArea.AbsolutePosition <= 9 Then
                            .Fields("Val" & rstArea.AbsolutePosition) = Rst!SalDay
                        Else
                            .Fields("Val10") = Rst!SalDay
                        End If
                    End If
                    rstArea.MoveFirst
                    Do While rstArea.EOF = False
                        If rstArea.AbsolutePosition <= 9 Then
                            .Fields("Head" & rstArea.AbsolutePosition) = left(rstArea!AreaName, 4)
                        Else
                            .Fields("Head10") = "OTH"
                        End If
                        rstArea.MoveNext
                    Loop
                End If


                If Trim(UCase(Rst!Vehicle_Type)) = "LCV" Then
                    .Fields("LCV_New") = Rst!NewEnquiry
                    .Fields("LCV_FollowUp") = IIf(Rst!NewEnquiry = 0, 1, 0)
                    .Fields("LCV_C0") = Rst!c0
                    .Fields("LCV_C1") = Rst!c1
                    .Fields("LCV_C2") = Rst!c2
                    .Fields("LCV_C3") = Rst!c3
                    .Fields("207_New") = 0
                    .Fields("207_FollowUp") = 0
                    .Fields("207_C0") = 0
                    .Fields("207_C1") = 0
                    .Fields("207_C2") = 0
                    .Fields("207_C3") = 0
                    .Fields("MHCV_New") = 0
                    .Fields("MHCV_FollowUp") = 0
                    .Fields("MHCV_C0") = 0
                    .Fields("MHCV_C1") = 0
                    .Fields("MHCV_C2") = 0
                    .Fields("MHCV_C3") = 0

                ElseIf Trim(UCase(Rst!Vehicle_Type)) = "207" Then
                    .Fields("LCV_FollowUp") = 0
                    .Fields("LCV_C0") = 0
                    .Fields("LCV_C1") = 0
                    .Fields("LCV_C2") = 0
                    .Fields("LCV_C3") = 0
                    .Fields("207_New") = Rst!NewEnquiry
                    .Fields("207_FollowUp") = IIf(Rst!NewEnquiry = 0, 1, 0)
                    .Fields("207_C0") = Rst!c0
                    .Fields("207_C1") = Rst!c1
                    .Fields("207_C2") = Rst!c2
                    .Fields("207_C3") = Rst!c3
                    .Fields("LCV_New") = 0
                    .Fields("MHCV_New") = 0
                    .Fields("MHCV_FollowUp") = 0
                    .Fields("MHCV_C0") = 0
                    .Fields("MHCV_C1") = 0
                    .Fields("MHCV_C2") = 0
                    .Fields("MHCV_C3") = 0
                Else
                    .Fields("LCV_New") = 0
                    .Fields("LCV_FollowUp") = 0
                    .Fields("LCV_C0") = 0
                    .Fields("LCV_C1") = 0
                    .Fields("LCV_C2") = 0
                    .Fields("LCV_C3") = 0
                    .Fields("207_New") = 0
                    .Fields("207_FollowUp") = 0
                    .Fields("207_C0") = 0
                    .Fields("207_C1") = 0
                    .Fields("207_C2") = 0
                    .Fields("207_C3") = 0
                    .Fields("MHCV_New") = Rst!NewEnquiry
                    .Fields("MHCV_FollowUp") = IIf(Rst!NewEnquiry = 0, 1, 0)
                    .Fields("MHCV_C0") = Rst!c0
                    .Fields("MHCV_C1") = Rst!c1
                    .Fields("MHCV_C2") = Rst!c2
                    .Fields("MHCV_C3") = Rst!c3
                End If
                .Update
            End With
            Rst.MoveNext
        Loop
    Set Rst = Nothing
'    Set RstRep = New Recordset
'    RstRep.CursorLocation = adUseClient
'    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
    RepName = "PipeLineRep"
    RepTitle = UCase(Me.CAPTION)
    ProfPurRepProc = True
    Exit Function
ELoop:
    ProfPurRepProc = False
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function
Private Function CaseAnalysisProc() As Boolean
On Error GoTo ELoop
Dim mQry As String, Condstr As String, mQRY1 As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: GoTo ELoop
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: GoTo ELoop
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 1)) = False Then RepPrint = False: GoTo ELoop
    If IsNotBlank(List1, FGrid.TextMatrix(List2, 1)) = False Then RepPrint = False: GoTo ELoop
    
    Condstr = " Where Visits.VisitDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Visits.VisitDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " "
       
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then GoTo ELoop
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then GoTo ELoop
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Visits.Rep_Code in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Visits.Party_Code in (" & GridString2 & ")"
    
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        Condstr = Condstr & " and Visits.ProspectiveCust_SubGroup =1 "
    Else
        Condstr = Condstr & " and Visits.ProspectiveCust_SubGroup = 0"
    End If
'PARTY:  VISHAL JAIN
'MODEL           START DT. STATUS CLOSE/LOST  GOT/LOST Date LOST REMARK (IF ANY)
'-----------------------------------------------------------------------
'REP NAME       VISIT DT  NEXT VISIT DT REMARKS   STATUS   EXPENCE
'-----------------------------------------------------------------------

'207/28 PICKUP   08/01/02  Hot    LOST                      OTHER BRAND
'207/28 PV       08/01/02  Hot
'407/31 BUSCH    08/01/02  Hot
'A.K. GANGULLY  08/01/02                dvdvf      Hot       0.00
'-----------------------------------  --------  ------------------------
'TOTAL :                                                     0.00

'    GSQL = "Select VSGQ.PartyCode&VSGQ.ProspectiveCust_SubGroup as PartyCode, " & _
        " VSGQ.Model,VSGQ.StartDate," & _
        " Switch(VSGQ.Call_Status=0,'Cold',VSGQ.Call_Status=1,'Warm',VSGQ.Call_Status=2,'Hot',VSGQ.Call_Status=3,'Nill') as CallStatus2, " & _
        " VSGQ.Got_Lost,VSGQ.GotLost_Date,VOLC.Name as LostCatName " & _
        " FROM ((((Visits V left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
        " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
        " Left JOIN Site ON V.Site_Code = Site.Site_Code) " & _
        " Left JOIN Veh_SubGroupQuot VSGQ ON V.Party_Code&V.ProspectiveCust_SubGroup=VSGQ.PartyCode&VSGQ.ProspectiveCust_SubGroup) " & _
        " Left JOIN Veh_OrdLostCatg VOLC ON VSGQ.Lost_Cat=VOLC.Code "
'        " WHERE (V.Next_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#) " & _
'        " AND (V.Next_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# ) "
'    mQRY = mQRY + CondStr
    
'    mQRY = "Select V.Party_Code&V.ProspectiveCust_SubGroup as PartyCode,E.Emp_Name,V.EXPENCE, V.VisitDate,V.Next_Date,V.Remark1,V.Remark2," & _
        " Switch(V.Call_Status=0,'Cold',V.Call_Status=1,'Warm',V.Call_Status=2,'Hot',V.Call_Status=3,'Nill') as CallStatus, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,PSG.Name+PSG.NSuffix,V.ProspectiveCust_SubGroup=1,SG.Name) as PartyName, " & _
        " Switch(V.ProspectiveCust_SubGroup=0,City.CityName,V.ProspectiveCust_SubGroup=1,City1.CityName) as CitName " & _
        " FROM ((((((Visits V LEFT JOIN Emp_Mast E on V.Rep_Code=E.Emp_Code) " & _
        " left join ProspectiveCust PSG on V.Party_Code=PSG.Cust_Code) " & _
        " LEFT JOIN SubGroup SG on V.Party_Code=SG.SubCode) " & _
        " Left JOIN Site ON V.Site_Code = Site.Site_Code) " & _
        " Left JOIN City ON SG.CityCode=City.CityCode) " & _
        " Left JOIN City City1 ON SG.CityCode=City1.CityCode) " & _
        " WHERE (V.Next_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy hh:mm") & "#) " & _
        " AND (V.Next_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy 23:59") & "# ) "
'    mQRY = mQRY + CondStr
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        mQry = "SELECT Visits.VisitDate,  Visits.EXPENCE, Visits.Call_Status, Visits.NEXT_DATE, Visits.REMARK1, Visits.Party_Code,Emp_Mast.Emp_Name,City.CityName, " & _
        "SubGroup.Name, SubGroup.Phone, SubGroup.Mobile " & _
        "FROM ((Visits LEFT JOIN SubGroup ON Visits.Party_Code = SubGroup.SubCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN Emp_Mast ON Visits.Rep_Code = Emp_Mast.Emp_Code"
    Else
        mQry = "SELECT Visits.VisitDate, Visits.EXPENCE, Visits.Call_Status, Visits.NEXT_DATE, Visits.REMARK1, Visits.Party_Code, Emp_Mast.Emp_Name, City.CityName," & _
        "ProspectiveCust.Name ,ProspectiveCust.PhoneOff + ProspectiveCust.PhoneResi as Phone, ProspectiveCust.Mobile " & _
        "FROM ((Visits LEFT JOIN Emp_Mast ON Visits.Rep_Code = Emp_Mast.Emp_Code) LEFT JOIN ProspectiveCust ON Visits.Party_Code = ProspectiveCust.Cust_Code) LEFT JOIN City ON ProspectiveCust.CityCode = City.CityCode"
    End If
    mQry = mQry + Condstr
    
    
    mQRY1 = "SELECT Veh_OrdLostCatg.NAME, Veh_SubGroupQuot.StartDate, Veh_SubGroupQuot.MODEL, Veh_SubGroupQuot.Call_Status, Veh_SubGroupQuot.Got_Lost, Veh_SubGroupQuot.PartyCode, Veh_SubGroupQuot.GotLost_Date " & _
    "FROM Veh_SubGroupQuot LEFT JOIN Veh_OrdLostCatg ON Veh_SubGroupQuot.Lost_Cat = Veh_OrdLostCatg.CODE"
    
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        Condstr = " where Veh_SubGroupQuot.ProspectiveCust_SubGroup = 1"
    Else
        Condstr = " where Veh_SubGroupQuot.ProspectiveCust_SubGroup = 0"
    End If
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Veh_OrdLostCatg.Rep_Code in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_OrdLostCatg.PartyCode in (" & GridString2 & ")"
    mQRY1 = mQRY1 + Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    
    Set RstRep1 = New Recordset
    RstRep1.CursorLocation = adUseClient
    RstRep1.Open (mQRY1), GCn, adOpenDynamic, adLockOptimistic
    SubRep1 = True
    
    RepName = "CaseAnalysis"
    RepTitle = UCase(Me.CAPTION)
    CaseAnalysisProc = True
    Exit Function
ELoop:
    CaseAnalysisProc = False
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function

Public Sub SelGridKeyPressLocal(txt As Object, SelGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If SelGrid(Index).Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then txt.TEXT = "": Exit Sub
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack Then
            If Len(txt.SelText) > 1 Then
                txt.SelLength = Len(txt.SelText) - 1
                FindStr = txt.SelText
            Else
                txt.TEXT = ""
                SelGrid(Index).SetFocus
                txt.Visible = False
                Exit Sub
            End If
        Else
            FindStr = txt.SelText + Chr(KeyAscii)
        End If
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
        Else    'character serach
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
        End If
        KeyAscii = 0
       If Rst.AbsolutePosition <> adPosEOF And Rst.AbsolutePosition <> adPosBOF Then
            SelGrid(Index).CellBackColor = CellBackColLeave
            SelGrid(Index).Row = Rst.AbsolutePosition
            SelGrid(Index).CellBackColor = CellBackColEnter
            txt.TEXT = Rst.Fields(FindFldName).Value
            txt.SelLength = Len(FindStr)
            txt.left = SelGrid(Index).CellLeft + SelGrid(Index).left
            txt.top = SelGrid(Index).CellTop + SelGrid(Index).top
            If txt.Visible = False Then
                txt.Visible = True: txt.ZOrder 0: txt.SetFocus: txt.BackColor = SelGrid(Index).CellBackColor
                 txt.ForeColor = SelGrid(Index).CellForeColor: txt.width = SelGrid(Index).CellWidth: txt.height = SelGrid(Index).CellHeight
            End If
       End If
End Sub

