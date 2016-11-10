VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form ReportWorkShop2 
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
   Moveable        =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11820
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Print"
      DownPicture     =   "ReportWorkShop2.frx":0000
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
      DownPicture     =   "ReportWorkShop2.frx":3132
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
      Left            =   5010
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
      FocusRect       =   0
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
      FocusRect       =   0
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
      FocusRect       =   0
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
      FocusRect       =   0
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
      FocusRect       =   0
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
Attribute VB_Name = "ReportWorkShop2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CellBackColLeave$ = &HFFFFFF
Private Const CellBackColEnter$ = &HFFFFC0
Private Const CellBackColLeave1$ = &HEDF7FE
Private Const CellBackColEnter1$ = &HFFFFC0
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim RsGrid3 As ADODB.Recordset
Dim RsGrid4 As ADODB.Recordset
Dim RepTitle$, RepName$
Dim RepPrint As Boolean
Dim RstRep As ADODB.Recordset
Dim RstRep1 As ADODB.Recordset
Dim SubRep1 As Boolean
'Modishekhar 17 mar
Dim FormulaStr1$, FormulaStr2$, FormulaStr3$, FormulaStr4$
Private Const GridRowHeight As Integer = 270


Private Const SprQuot As Byte = 1
Private Const WksEstimate As Byte = 2
Private Const WksPerforma As Byte = 3
Private Const WksReqReg As Byte = 4
Private Const WksVehDiary As Byte = 5
Private Const WksJobReg As Byte = 6
Private Const GatePassReg As Byte = 7
Private Const WksReqRegGrd As Byte = 8
Private Const WksRegOutLab As Byte = 9
Private Const OverTimeReg As Byte = 10
Private Const VehHisReg As Byte = 12
Private Const JobBookReg As Byte = 13
Private Const WorkshopVehicleReg As Byte = 14
Private Const InsuranceExpiryReg As Byte = 15


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

Public GRepFormName$
Dim mLastRow As Integer
Dim mFirstRow As Integer
Dim mHelpGridNo
Dim GridKey As Integer
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim GridString1$
Dim GridString2$
Dim GridString3$
Dim GridString4$
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
Dim TotalOpen As Double, TotalClosed As Double, TotalPending As Double
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub BTNPRINT_Click()
'On Error GoTo ERRORHANDLER
SubRep1 = False
RepPrint = True
Select Case GRepFormName
    Case OverTimeReg
        OverTimeRegProc
    Case WksRegOutLab
        WksOutLabRegProc
    Case WksReqRegGrd
        WksReqRegGrdProc
    Case GatePassReg
        GatePassRegProc
    Case WksVehDiary
        WksVehicleDiary
    Case SprQuot, WksEstimate, WksPerforma
        SprQuotReg
    Case WksJobReg
        WksJobRegister
    Case WksReqReg
        WksRequisition
    Case VehHisReg
        VehHisRegProc
    Case WorkshopVehicleReg
        ProcWorkshopVehicleRegister
    Case InsuranceExpiryReg
        ProcInsuranceExpiryRegister
    Case JobBookReg
        JobBookRegProc
End Select
If RepPrint = False Then Exit Sub
If GRepFormName = VehHisReg Then Exit Sub

CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")

rpt.Database.SetDataSource RstRep
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
rpt.ReadRecords
Set RstRep = Nothing

Call Formulas
Call Report_View(rpt, RepTitle, , False)
Set rpt = Nothing
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
    'If Mid(UserPermission(Me.Name), 4, 1) = "*" Then BTNPRINT.Enabled = False
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
                Case OverTimeReg
                    ListArray = Array("Detail", "Summary")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                Case SprQuot, WksEstimate, WksPerforma
                    ListArray = Array("Stores", "Workshop")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
               Case WksReqRegGrd, WksReqReg, WksJobReg
                    If RSOJPR = True Then
                        ListArray = Array("All", "Closed", "UnClosed", "Cancelled")
                        Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
                    Else
                        ListArray = Array("All", "Closed", "UnClosed")
                        Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
                    End If
                Case GatePassReg, WksVehDiary
                    ListArray = Array("Yes", "No")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                Case JobBookReg
                    ListArray = Array("Pending", "Done", "All")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
              End Select
        Case List2
            Select Case GRepFormName
                Case WksReqRegGrd, WksReqReg
                    ListArray = Array("All", "PDI", "Free Service", "Chargable", "Warranty", "General", "Company Vehicle", "Complementary", "AMC")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 9)
                Case WksEstimate, WksPerforma
                    ListArray = Array("Yes", "No")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                 Case WksVehDiary
                    ListArray = Array("Yes", "No", "All")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
        Case WksJobReg
                    ListArray = Array("All", "Regular", "On Site Repair", "Quick Repair")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
                    
            End Select
        Case List3
            Select Case GRepFormName
                Case WksReqRegGrd, WksReqReg
                    ListArray = Array("All", "Regular", "On Site Repair", "Quick Repair")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
                Case WksVehDiary
                    ListArray = Array("Job Open", "Job Close")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                    
            End Select
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
    Case Date1, Date2, Cat1, Cat2, Cat3, Cat4, Cat5
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
 Select Case FGrid.Row
    Case Cat1
        Select Case GRepFormName
        End Select
    Case Cat2
        Select Case GRepFormName
        End Select
    Case Cat3
        Select Case GRepFormName
        End Select
    Case Cat4
        Select Case GRepFormName
        End Select
    Case Cat5
        Select Case GRepFormName
        End Select
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
    Select Case FGrid.Row
'        Case Cat1, Cat2
'             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
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
Select Case FGrid.Row
    Case Cat1, Cat2, Cat3, Cat4, Cat5
         FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
    Case List1, List2, List3
        If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
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
Dim I As Integer, Cnt As Integer, GridHeight As Integer

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
'**
Ini_Grid
'**
For I = 1 To 4
    If GridSel(I).Visible = True Then Cnt = Cnt + 1
Next
For I = mFirstRow To mLastRow
    GridHeight = GridHeight + FGrid.RowHeight(I)
Next
FGrid.height = GridHeight + FGrid.RowHeight(mFirstRow)
'FGrid.Height = ((mLastRow - mFirstRow) * PubGridRowHeight) + 500
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
'    FGrid.CellBackColor = CellBackColLeave
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
Dim ac_str$
Dim I As Integer
Dim GridRow As Integer
Dim formulastr$   'Modishekhar 17 mar
formulastr = "" 'Modishekhar 17 mar
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
           'Modishekhar 17 mar
            If Len(formulastr + GridSel(Gridindex).TextMatrix(GridRow, 2)) < 255 Then
                formulastr = formulastr + IIf(formulastr = "", "For " & GridSel(Gridindex).TextMatrix(0, 1) & " : " & GridSel(Gridindex).TextMatrix(GridRow, 1), "," & GridSel(Gridindex).TextMatrix(GridRow, 1))
            End If
           'Modi End
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
'Modishekhar 17 mar
    Select Case Gridindex
        Case 1
            FormulaStr1 = mID(formulastr, 1, 254)
        Case 2
            FormulaStr2 = mID(formulastr, 1, 254)
        Case 3
            FormulaStr3 = mID(formulastr, 1, 254)
        Case 4
            FormulaStr4 = mID(formulastr, 1, 254)
    End Select
'modi end
    If ac_str = "" Then
        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
        GridSel(Gridindex).SetFocus
        RepPrint = False
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

Private Sub GridInitialise(Gridindex As Integer, GridSql$)
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
Dim Grid1Sql$, Grid2Sql$, Grid3Sql$, Grid4Sql$

 Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where site_code='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    
    
Select Case GRepFormName
    Case OverTimeReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Detail/Summary": .RowHeight(List1) = GridRowHeight
                       
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Detail"
            
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Emp_Name as EmployeeName,Emp_Code as code from Emp_Mast Order by Emp_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 3, Grid3Sql

    Case WksReqRegGrd
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Job Status": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Purpose Of Part": .RowHeight(List2) = GridRowHeight
            If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
                .TextMatrix(List3, 0) = "Job Type": .RowHeight(List3) = GridRowHeight
                .TextMatrix(List3, 1) = "All"
            End If
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List2, 1) = "All"
            
        End With
        If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
            mFirstRow = Date1: mLastRow = List3: mHelpGridNo = 4
        Else
            mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 4
        End If
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,PartGrade_Name as GradeName,PartGrade_Code as code from Part_Grade Order by PartGrade_Name"
        GridInitialise 4, Grid4Sql
    Case WksRegOutLab '9
        With FGrid
            .TextMatrix(Date1, 0) = "Job Open Date From": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "Job Open To Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql

    Case GatePassReg
        With FGrid
'            .BorderStyle = flexBorderNone
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Include Jobcard Y/N": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Yes"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case VehHisReg
        With FGrid
'            .BorderStyle = flexBorderNone
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,RegNo as RegNo,CardNo as code  from HisCard order by RegNo"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,Chassis,CardNo as code  from HisCard order by Chassis "
        GridInitialise 4, Grid4Sql

    Case WorkshopVehicleReg
    
        With FGrid
'            .BorderStyle = flexBorderNone
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
       'mHelpGridNo = 4
       mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
    
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Model ,Model as code  from Model order by Model"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,Chassis,CardNo as code  from HisCard order by Chassis "
        GridInitialise 4, Grid4Sql

    Case InsuranceExpiryReg
    
        With FGrid
            .BorderStyle = flexBorderNone
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
       mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 0
    
    Case WksVehDiary
        With FGrid
            .TextMatrix(Date1, 0) = "Open Date From": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "     Date To": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Print Complaints Y/N": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Closed Job Y/N ": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Date Filter": .RowHeight(List3) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "No"
            .TextMatrix(List2, 1) = "Yes"
            .TextMatrix(List3, 1) = "Job Open"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List3
        
    Case WksReqReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Job Status": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Purpose Of Part": .RowHeight(List2) = GridRowHeight
            If RSOJPR = True Then
                .TextMatrix(List3, 0) = "Job Type": .RowHeight(List3) = GridRowHeight
                .TextMatrix(List3, 1) = "All"
            End If
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List2, 1) = "All"
            
        End With
        If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
            mFirstRow = Date1: mLastRow = List3: mHelpGridNo = 3
        Else
            mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 3
        End If
'
'******* < By Rahul At U.N.Automobiles Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case WksJobReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Job Status": .RowHeight(List1) = GridRowHeight
            
            If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
                .TextMatrix(List2, 0) = "Job Type": .RowHeight(List2) = GridRowHeight
                .TextMatrix(List2, 1) = "All"
            End If
                .TextMatrix(Date1, 1) = PubStartDate
                .TextMatrix(Date2, 1) = PubLoginDate
                .TextMatrix(List1, 1) = "All"
        End With
        If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
            mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 4
        Else
            mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
        End If
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select '' as O,Emp_Name as ServiceAdvisor,Emp_Code  as code from Emp_Mast where Designation ='SUPERVISOR' order by Emp_Name"
        GridInitialise 4, Grid4Sql

    Case SprQuot
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Stores/Workshop": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Stores"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case JobBookReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Type": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from " & sitecond & " site order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
        
    Case WksEstimate, WksPerforma
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Stores/Workshop": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Details Y/N    ": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Workshop"
            .TextMatrix(List2, 1) = "Yes"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name" ' Where Div_Code='" & PubDivCode & "' order by Div_Name" BY VIKASH 31/10/2003
        GridInitialise 2, Grid2Sql
End Select
End Sub
Private Function IsNotBlank(FieldRow As Integer, FieldCaption$) As Boolean
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
Private Sub SprQuotReg()
On Error GoTo ELoop
Dim mQry$, Condstr$
'Date1,Date2,List1,List1,List1,List2,List1,List1
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where E.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and E.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and E.Stores_Works = '" & FGrid.TextMatrix(List1, 1) & "'"
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(E.site_code,1) in (" & GridString1 & ") and E.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and E.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and E.Stores_Works = '" & FGrid.TextMatrix(List1, 1) & "'"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(E.site_code,1) ='" & PubSiteCode & "' "
    End If
    
    mQry = "SELECT E.V_DATE,E.V_NO,E.DocID,E.Party_Name,(E.SprAmt_TB +E.SprAmt_TP + E.SprAmt_MRP_TB + E.SprAmt_MRP_TP - E.tax_amtmrp - E.taxsur_amtmrp) as SprAmt, (E.OilAmt_MRP_TB + E.OilAmt_MRP_TP + " & _
        "E.OilAmt_TB + E.OilAmt_TP) as OilAmt, (E.D_Amt_TB +  E.D_Amt_TP + E.D_Amt_MRP_TB + E.D_Amt_MRP_TP) as DisAmt, E.Total_Amt, E.Addition, E.Gen_Sur_Amt," & _
        "E.Trans_Amt, (E.Tax_Amt + E.Tax_AmtMRP + E.Tax_Sur_Amt + E.TaxSur_AmtMRP + E.TOT_Amt + E.TOT_AmtMRP + E.ReSalTax_Amt) as TaxAmt,E.Job_DocId,E.Model,E.RegNo," & _
        "E.Packing, E.Lab_Amt, E.Lab_D_Amt,E.Lab_TaxAmt,E.Suppl_YN  " & _
        "FROM Estimate E "
    
    mQry = mQry + Condstr + " order by E.v_No"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        RepName = "SprEstQuot"
    Else
        RepName = "SprEstSum"
    End If

    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub Formulas()
On Error GoTo ELoop
Dim I As Integer
'Modishekhar 17 mar
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("Formulastr1")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr1 & "'"
            Case UCase("Formulastr2")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr2 & "'"
            Case UCase("Formulastr3")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr3 & "'"
            Case UCase("Formulastr4")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr4 & "'"
        End Select
    Next
    FormulaStr1 = "": FormulaStr2 = "": FormulaStr3 = "": FormulaStr4 = ""
    'modi end
Select Case GRepFormName
    Case WksReqReg, WksReqRegGrd, WksRegOutLab, OverTimeReg, JobBookReg
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("OTRateStr")
                    rpt.FormulaFields(I).TEXT = "'Over Time Rate/Hr Rs.:'+ '" & Format(FGrid.TextMatrix(Cat1, 1), "0.00") & "'"
            End Select
        Next
   Case WksVehDiary
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'Open Date From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("complaintYN")
                    If FGrid.TextMatrix(List1, 1) = "Yes" Then
                        rpt.FormulaFields(I).TEXT = 1
                    Else
                        rpt.FormulaFields(I).TEXT = 0
                    End If
                Case UCase("TotalOpen")
                    rpt.FormulaFields(I).TEXT = "" & TotalOpen & ""
                Case UCase("TotalClosed")
                    rpt.FormulaFields(I).TEXT = "" & TotalClosed & ""
                Case UCase("TotalPending")
                    rpt.FormulaFields(I).TEXT = "" & TotalPending & ""
            End Select
        Next
    Case WksJobReg
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'Upto Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
    Case SprQuot, WksEstimate, WksPerforma
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("MinDiff")
                    rpt.FormulaFields(I).TEXT = "" & Val(FGrid.TextMatrix(Cat1, 1)) & ""
                Case UCase("SPRWORK")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "'"
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub
Private Sub WksRequisition()
On Error GoTo ELoop
Dim mQry$, Condstr$
Dim mPurpose$
Select Case FGrid.TextMatrix(List2, 1)
    Case "PDI"
        mPurpose = "P"
    Case "Free Service"
        mPurpose = "F"
    Case "Chargable"
        mPurpose = "C"
    Case "Warranty"
        mPurpose = "W"
    Case "Company Vehicle"
        mPurpose = "O"
    Case "Complementary"
        mPurpose = "L"
    Case "AMC"
        mPurpose = "A"
End Select
    'P- >PDI,F- >Free Service, C- >Chargable,W- >Warranty,O- >Company Vehicle,L- >Complementary

    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
        If IsNotBlank(List3, FGrid.TextMatrix(List3, 0)) = False Then RepPrint = False: Exit Sub
    End If
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If FGrid.TextMatrix(List1, 1) = "Closed" Then
        Condstr = " where Job_Card.JobCloseDate >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    Else
        Condstr = " where SP_Stock.v_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Stock.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    End If
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP_Stock.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP_Stock.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Job_Card.Serv_Type in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(SP_Stock.DocId,1) in (" & GridString3 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "UnClosed" Then Condstr = Condstr & " and Job_Card.JobCloseDate Is Null"
    If FGrid.TextMatrix(List1, 1) = "Closed" Then Condstr = Condstr & " and Job_Card.JobCloseDate Is Not Null "
    If FGrid.TextMatrix(List2, 1) <> "All" And FGrid.TextMatrix(List2, 1) <> "General" Then Condstr = Condstr & " and SP_Stock.Purpose = '" & mPurpose & "'"
    
    'Nra modi for general printing
    If FGrid.TextMatrix(List2, 1) = "General" Then
        Condstr = Condstr & " and left(SP_Stock.V_type,4) in ('" & WksGenReq & "')"
    Else
        'Condstr = Condstr & " and left(SP_Stock.V_type,4) in ('" & WksGenReq & "','" & WksReqWrt & "')"
        Condstr = Condstr & " and S.V_type in ('W_SIC','W_SIR')"
    End If
    'End Modi
    'Nra modi for Sorting
    If FGrid.TextMatrix(List3, 1) = "Regular" Then
        Condstr = Condstr & " and Job_Card.Jobtype='R'"
    ElseIf FGrid.TextMatrix(List3, 1) = "On Site Repair" Then
        Condstr = Condstr & " and Job_Card.Jobtype='O'"
    ElseIf FGrid.TextMatrix(List3, 1) = "Quick Repair" Then
        Condstr = Condstr & " and Job_Card.Jobtype='Q'"
    End If
    
    'End Modi
    'By Rahul U.N.Automobiles 11-04-2003
    
    mQry = "SELECT " & _
                "Part.Part_Name,SP_Stock.DocID, SP_Stock.V_No, SP_Stock.V_Date, Job_Card.DocID as JobDocID," & _
                "Job_Card.Job_Date, Job_Card.JobCloseDate,Job_Card.DocId_InvSpr, HisCard.RegNo, " & _
                "HisCard.Chassis, SP_Stock.Part_No, SP_Stock.Purpose, SP_Stock.Qty_Doc, SP_Stock.Qty_Iss," & _
                "SP_Stock.Qty_Ret, SP_Stock.Rate, SP_Stock.Amount, Part.Part_Grade, " & _
                "(SP_Stock.Claim_Div + SP_Stock.Claim_Site + SP_Stock.Claim_YearPrefix + SP_Stock.Claim_Type +  SP_Stock.Claim_No) as ClaimNo, SP_Stock.Claim_Date,SP_Stock.Remark,HisCard.Engine,Part_DiscFactor.PurcDisc_Per,SP_Stock.Tax_YN " & _
           "FROM ((((Sp_Sale S Left Join SP_Stock On S.DocId = Sp_Stock.Invoice_DocId) " & _
                "LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1)) " & _
                "LEFT JOIN Job_Card ON SP_Stock.Job_DocID = Job_Card.DocId) " & _
                "LEFT JOIN HisCard ON (Job_Card.CardNo = HisCard.CardNo)) " & _
                "LEFT JOIN Part_DiscFactor ON (Part.Disc_Factor = Part_DiscFactor.DiscFac_Catg)"
                
    mQry = mQry + Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksReqReg"
    RepTitle = UCase(Me.CAPTION) + " [" + FGrid.TextMatrix(List1, 1) + "]"
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub WksReqRegGrdProc()
On Error GoTo ELoop
Dim mQry$, Condstr$
Dim mPurpose$
Select Case FGrid.TextMatrix(List2, 1)
    Case "PDI"
        mPurpose = "P"
    Case "Free Service"
        mPurpose = "F"
    Case "Chargable"
        mPurpose = "C"
    Case "Warranty"
        mPurpose = "W"
    Case "Company Vehicle"
        mPurpose = "O"
    Case "Complementary"
        mPurpose = "L"
    Case "AMC"
        mPurpose = "A"
End Select
    'P- >PDI,F- >Free Service, C- >Chargable,W- >Warranty,O- >Company Vehicle,L- >Complementary

    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
        If IsNotBlank(List3, FGrid.TextMatrix(List3, 0)) = False Then RepPrint = False: Exit Sub
    End If
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where SPStk.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SPStk.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SPStk.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SPStk.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and JC.Serv_Type in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(SPStk.DocId,1) in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Part.Part_Grade in (" & GridString4 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "UnClosed" Then Condstr = Condstr & " and JC.JobCloseDate Is Null"
    If FGrid.TextMatrix(List1, 1) = "Closed" Then Condstr = Condstr & " and JC.JobCloseDate Is Not Null "
    If FGrid.TextMatrix(List2, 1) <> "All" And FGrid.TextMatrix(List2, 1) <> "General" Then Condstr = Condstr & " and SPStk.Purpose = '" & mPurpose & "'"
    'Nra modi for general printing
    If FGrid.TextMatrix(List2, 1) = "General" Then
        Condstr = Condstr & " and SPStk.V_type in ('" & WksGenReq & "')"
    Else
        'Condstr = Condstr & " and SPStk.V_type in ('" & WksGenReq & "','" & WksReqWrt & "')"
        Condstr = Condstr & " and S.V_type in ('W_SIC','W_SIR')"
    End If
    If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
        If FGrid.TextMatrix(List3, 1) = "Regular" Then
            Condstr = Condstr & " and JC.Jobtype='R'"
        ElseIf FGrid.TextMatrix(List3, 1) = "On Site Repair" Then
            Condstr = Condstr & " and JC.Jobtype='O'"
        ElseIf FGrid.TextMatrix(List3, 1) = "Quick Repair" Then
            Condstr = Condstr & " and JC.Jobtype='Q'"
        End If
    End If
    mQry = "SELECT " & _
                "PG.PartGrade_Name,Part.Part_Name,SPStk.docid, SPStk.V_No, SPStk.V_Date, " & _
                "JC.DocID as JobID,JC.Job_No,JC.Job_Date, JC.JobCloseDate,JC.DocId_InvSpr," & _
                "SPStk.Part_No, SPStk.Purpose, SPStk.Qty_Doc, SPStk.Qty_Iss," & _
                "SPStk.Qty_Ret, SPStk.Rate, SPStk.Amount " & _
           "FROM ((((Sp_Sale S Left Join SP_Stock as SPStk On S.DocId = SpStk.Invoice_DocId) " & _
                "LEFT JOIN Part ON SPStk.Part_No = Part.PART_NO and Part.Div_Code = left(SPStk.Docid,1)) " & _
                "LEFT JOIN Job_Card as JC ON SPStk.Job_DocID = JC.DocId)" & _
                "Left Join Part_Grade as PG on Part.Part_Grade=PG.PartGrade_Code)"
    mQry = mQry + Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksReqRegGrd"
    RepTitle = UCase(Me.CAPTION) + " [" + FGrid.TextMatrix(List1, 1) + "]"
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub WksVehicleDiary()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
   ' If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    
    If StrCmp(FGrid.TextMatrix(List3, 1), "Job Open") Then
        Condstr = " where Left(J.DocId,1)='" & PubDivCode & "' And J.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and J.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    Else
        Condstr = " where Left(J.DocId,1)='" & PubDivCode & "' And J.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and J.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    End If
    
    
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        Condstr = Condstr & " and Not J.JobCloseDate Is Null"
    ElseIf FGrid.TextMatrix(List2, 1) = "No" Then
        Condstr = Condstr & " and J.JobCloseDate Is Null"
    End If


    mQry = "SELECT " & _
        "J.DocId,J.Job_Date,J.Recp_time,H.RegNo,H.name,H.Chassis,H.Model,J.Job_No," & _
        "J.AtKMsHrs,ST.Serv_Catg,ST.Serv_Type,Emp_Mast.Emp_Name,H.PhoneOff,J.ExpDelDate,J.JobCloseDate," & _
        "" & vIsNull("J.JobCloseDate", "J.Job_Date") & " - J.Job_Date as Days,J.REMARK,JD.S_No,JD.Details,J.JobComp_Dt_Time," & cIIF("JD.S_No=1", "ST.FreeServCode", "0") & " as FreeServCode,H.Mobile  " & _
    "FROM (((Job_Card as J " & _
        "LEFT JOIN HisCard as H ON J.CardNo = H.CardNo) " & _
        "LEFT JOIN Emp_Mast ON J.RecBy_Supervisor = Emp_Mast.Emp_Code) " & _
        "LEFT JOIN Job_Demand as JD ON J.DocId = JD.Job_DocID) " & _
        "LEFT JOIN Service_Type as ST ON J.Serv_Type = ST.Serv_Type"
    mQry = mQry + Condstr + " Order by J.DocID"
    
    
    TotalOpen = VNull(GCn.Execute("Select Count(Docid) from Job_Card J " & Condstr & "").Fields(0).Value)
    TotalClosed = VNull(GCn.Execute("Select Count(Docid) from Job_Card J " & Condstr & " and Not J.JobCloseDate Is Null").Fields(0).Value)
    TotalPending = VNull(GCn.Execute("Select Count(Docid) from Job_Card J " & Condstr & " and J.JobCloseDate Is Null").Fields(0).Value)
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksVehDiary"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub GatePassRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    Condstr = " where GP.GatePassDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and GP.GatePassDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("GP.GatePassNo", "2", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("GP.GatePassNo", "2", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(GP.GatePassNo,1) in (" & GridString2 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "Yes" Then Condstr = Condstr & " and Not GP.Job_DocID Is Null"
    If FGrid.TextMatrix(List1, 1) = "No" Then Condstr = Condstr & " and (GP.Job_DocID Is Null or GP.Job_DocID='')"
    
    mQry = " SELECT GP.GatePassNo,GP.GatePassDate,GP.Job_DocID,GP.Mech_Code,GP.ContractCode," & _
        " GP.Purpose,GP.ContractRecdDate,GP.ContractAmt,GP.ContractorBillNo,GP.Remarks, " & _
        " E.Emp_Name,CF.FinName,H.Model,H.RegNo, H.Chassis,H.Name,H.Add1,H.Add2,H.Add3,City.CityName" & _
    " FROM (((((Job_GatePass as GP LEFT JOIN Job_Card on  GP.Job_DocID=Job_Card.DocID) " & _
        " Left join HisCard as H ON Job_Card.CardNo = H.CardNo)" & _
        " LEFT JOIN Emp_Mast as E ON GP.Mech_Code = E.Emp_Code) " & _
        " Left Join ContractFinance as CF on GP.ContractCode=CF.FinCode) " & _
        " Left Join City on H.CityCode=City.CityCode)"
    
    mQry = mQry & Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "GatePassReg"
    RepTitle = UCase(Me.CAPTION) + " [ " + FGrid.TextMatrix(List1, 1) + " ]"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VehHisRegProc()
On Error Resume Next
Dim mRepName$, mAddRec As Boolean, ExitLoop As Boolean, RecCountFix As Integer
Dim Rst As ADODB.Recordset, RstRep As ADODB.Recordset, RST1 As ADODB.Recordset
Dim Rst2 As ADODB.Recordset, RST3 As ADODB.Recordset, Rst4 As ADODB.Recordset
Dim mQry$, mQRY1$, Condstr$, CondStr1$, I As Integer


    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocID", "2", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocID", "2", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then CondStr1 = " and CardNo in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(JC.DocID,1) in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then CondStr1 = CondStr1 & " and CardNo in (" & GridString4 & ")"
      If Check1(1).Value = Unchecked Then CondStr1 = CondStr1 & " and " & cMID("Job_Card.DocID", "2", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then CondStr1 = CondStr1 & " and " & cMID("Job_Card.DocID", "2", "1") & "  ='" & PubSiteCode & "' "
    End If
    mQRY1 = "Select distinct CardNo as SearchCode from Job_Card where Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & CondStr1
            
            mRepName = "VehHisReg"
                        
                mQry = "select H.CardNo,H.Site_Code,H.CardDate,H.Model,H.RegNo,H.RegDate," & _
                    " H.Chas_Type,H.Chassis,H.Engine,H.VehSerialNo,H.Supplier_BillNo,H.Supplier_BillDate," & _
                    " H.Dealer_Code,H.Delivery_Date,H.CouponNo,H.ColourCode," & _
                    " H.Steer_Type,H.Steer_Make,H.Alternator,H.StarterMotor,H.Battery,H.GBoxNo,H.RAxelNo," & _
                    " H.Name,H.ConPerson,H.Add1,H.Add2,H.Add3,H.CityCode,H.PhoneOff,H.PhoneResi,H.Mobile,H.Mail_ID," & _
                    " H.DOB,H.DOM,H.OwnDrive,H.OwnerRemark,H.Next_JobDate,H.Ac_Code,H.Govt_YN,H.Inv_No,H.Locked_Text," & _
                    " H.LJob_DocId,H.LJob_Date,H.LJob_AtKMsHrs,M.Model_Desc,M.Chas_Type AS ModelChasType," & _
                    " col.Col_Desc, Amd_Dealer.D_Name,City.CityName, Emp.Emp_Name as LMechName, H.PhoneOff, H.PhoneResi, H.Mobile " & _
                    " from ((((((Hiscard H left join Model M on H.Model=M.Model) " & _
                    " left join Colmast Col on H.ColourCode=Col.Col_Code) " & _
                    " left join Amd_Dealer on H.Dealer_Code=Amd_Dealer.D_Code)" & _
                    " left join City on H.CityCode=City.CityCode) " & _
                    " Left Join Job_Card J on H.LJob_DocID=J.DocId) " & _
                    " Left Join Emp_Mast Emp on J.RecBy_Mechanic=Emp.Emp_Code) " & _
                    " where H.CardNo in (" & mQRY1 & ") order by H.CardNo"
                    
                Set Rst = New Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
                
                If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
                
                'Select Jobcard & other details
                mQry = "Select J.CardNo, J.DocId,J.Job_Date,J.JobCloseDate,J.AtKMsHrs,J.Serv_Type,J.RecBy_Mechanic as RecBy_MechCode, Emp.Emp_Name as RecBy_MechName," & _
                    " J.NetLab_Amt,S.Total_Amt as NetSpr_Amt,J.ObservBy_Super,J.ActionBy_Super" & _
                    " from (Job_Card J Left join SP_Sale S on J.DocId_InvSpr=S.DocID) " & _
                    " Left join Emp_Mast Emp on J.RecBy_Mechanic=Emp.Emp_Code " & _
                    " where J.CardNo in (" & mQRY1 & _
                    ") Order by J.Job_Date Desc, j.Docid Desc"
                Set RST1 = New Recordset
                RST1.CursorLocation = adUseClient
                RST1.Open (mQry), GCn, adOpenStatic, adLockReadOnly
                
                    'Create temp table
                    CreaTabRstRep RstRep, VehHisRstTmp
                    'temp table created
                
                If RST1.RecordCount > 0 Then
                    'Problem Reported Details
                    mQry = "Select J.CardNo,JD.Job_DocID,JD.S_No,JD.Code as Prob_Code,JD.Details as Prob_Reported " & _
                        " from (Job_Card J Left join Job_Demand JD on J.DocId=JD.Job_DocID) " & _
                        " where J.CardNo in (" & mQRY1 & ") and JD.Job_DocID=J.DocID " & _
                        " Order by JD.Job_DocID,JD.S_No"
                    Set Rst2 = New Recordset
                    Rst2.CursorLocation = adUseClient
                    Rst2.Open (mQry), GCn, adOpenStatic, adLockReadOnly
                    
                    'Labour Done
                    mQry = "Select J.CardNo,JL.Job_DocID,JL.Lab_Code,L.Lab_Desc as Lab_Done " & _
                        " from ((Job_Card J Left join Job_Lab JL on J.DocId=JL.Job_DocID) " & _
                        " Left join Labour L on JL.Lab_Code=L.Lab_Code) " & _
                        " where J.CardNo in (" & mQRY1 & ") and JL.Job_DocID=J.DocID " & _
                        " Order by JL.Job_DocId,JL.S_No"
                    Set RST3 = New Recordset
                    RST3.CursorLocation = adUseClient
                    RST3.Open (mQry), GCn, adOpenStatic, adLockReadOnly
                    'Requisition / Parts
                    mQry = "Select distinct J.CardNo, Stk.Job_DocID,Stk.DocId as DocIDReq, " & _
                        " Stk.Part_No,Part.Part_Name,Stk.Purpose, Stk.Qty_Iss - Qty_Ret as Qty, Stk.Rate, Stk.Amount " & _
                        " from ((Job_Card J Left join Sp_Stock as Stk on J.DocID=Stk.Job_DocId) " & _
                        " left join Part on Stk.Part_No=Part.Part_No ) " & _
                        " where J.CardNo in (" & mQRY1 & ") and Stk.Job_DocID=J.DocID " & _
                        " Order by Stk.Job_DocId,Stk.DocId"
                    Set Rst4 = New Recordset
                    Rst4.CursorLocation = adUseClient
                    Rst4.Open (mQry), GCn, adOpenStatic, adLockReadOnly
                    RST1.MoveFirst
                    For I = 1 To RST1.RecordCount
                        Rst.Filter = ("CardNo= '" & RST1!CardNo & "'")
                        With RstRep
                            .AddNew
                            If Rst.RecordCount > 0 Then
                            
                                .Fields("RegNo") = Rst!RegNo
                                .Fields("CustName") = Rst!Name
                                .Fields("Add1") = Rst!Add1
                                .Fields("Add2") = Rst!Add2
                                .Fields("Add3") = Rst!Add3
                                .Fields("DOSale") = Rst!Supplier_BillDate
                                .Fields("Chassis") = Rst!Chassis
                                .Fields("Engine") = Rst!Engine
                                .Fields("GBNo") = Rst!GBoxNo
                                .Fields("RANo") = Rst!RAxelNo
                                .Fields("PhoneOff") = XNull(Rst!PhoneOff)
                                .Fields("PhoneResi") = XNull(Rst!PhoneResi)
                                .Fields("Mobile") = XNull(Rst!Mobile)
                            End If
                            .Fields("CardNo") = RST1!CardNo
                            .Fields("JobDocID") = RST1!DocID
                            .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                            .Fields("Job_Date") = RST1!Job_Date
                            .Fields("JobCloseDate") = RST1!JobCloseDate
                            .Fields("AtKMsHrs") = RST1!AtKMsHrs
                            .Fields("Serv_Type") = RST1!Serv_Type
                            .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                            .Fields("RecBy_MechName") = RST1!RecBy_MechName
                            .Fields("NetLab_Amt") = RST1!NetLab_Amt
                            .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                            .Fields("ObservBy_Super") = RST1!ObservBy_Super
                            .Fields("ActionBy_Super") = RST1!ActionBy_Super
                            .Update
                        End With
                        'Problem Reported Details
                        If Rst2.RecordCount <= 0 Then GoTo lblLabDone
                        Rst2.MoveFirst
                        Rst2.FIND ("Job_DocID='" & RST1!DocID & "'")
                        If Rst2.EOF Then GoTo lblLabDone
                        ExitLoop = False
                        mAddRec = False
                        If I = 1 Then
                           RstRep.MoveFirst
                        Else
                            RstRep.Move (2 - (RstRep.RecordCount - RecCountFix))
                            If RstRep.EOF Then mAddRec = True
                        End If
                        Do While Not ExitLoop 'Rst2!Job_DocID = Rst1!DocId
                            If Rst2!job_docid = RST1!DocID Then
                                With RstRep
                                    If mAddRec Then
                                        .AddNew
                                        .Fields("CardNo") = RST1!CardNo
                                        .Fields("JobDocID") = RST1!DocID
                                        .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                                        .Fields("Job_Date") = RST1!Job_Date
                                        .Fields("JobCloseDate") = RST1!JobCloseDate
                                        .Fields("AtKMsHrs") = RST1!AtKMsHrs
                                        .Fields("Serv_Type") = RST1!Serv_Type
                                        .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                                        .Fields("RecBy_MechName") = RST1!RecBy_MechName
                                        .Fields("NetLab_Amt") = RST1!NetLab_Amt
                                        .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                                        .Fields("ObservBy_Super") = RST1!ObservBy_Super
                                        .Fields("ActionBy_Super") = RST1!ActionBy_Super
                                    End If
                                    .Fields("JobDocID") = RST1!DocID
                                    .Fields("Prob_Code") = Rst2!Prob_Code
                                    .Fields("Prob_Reported") = Rst2!Prob_Reported
                                    .Update
                                End With
                            End If
                            Rst2.MoveNext
                            If Rst2.EOF Then
                                ExitLoop = True
                            ElseIf Rst2!job_docid <> RST1!DocID Then
                                ExitLoop = True
                            Else
                                RstRep.MoveNext
                                If RstRep.EOF Then mAddRec = True
                            End If
                        Loop
                        
lblLabDone:
                        'Labour Done
                        If RST3.RecordCount <= 0 Then GoTo lblRequisition
                        RST3.MoveFirst
                        RST3.FIND ("Job_DocID='" & RST1!DocID & "'")
                        If RST3.EOF Then GoTo lblRequisition
                        ExitLoop = False
                        mAddRec = False
                        If I = 1 Then
                           RstRep.MoveFirst
                        Else
                            RstRep.Move (2 - (RstRep.RecordCount - RecCountFix))
                            If RstRep.EOF Then mAddRec = True
                        End If
                        Do While Not ExitLoop 'Rst3!Job_DocID = Rst1!DocId
                            If RST3!job_docid = RST1!DocID Then
                                With RstRep
                                    If mAddRec Then
                                        .AddNew
                                        .Fields("CardNo") = RST1!CardNo
                                        .Fields("JobDocID") = RST1!DocID
                                        .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                                        .Fields("Job_Date") = RST1!Job_Date
                                        .Fields("JobCloseDate") = RST1!JobCloseDate
                                        .Fields("AtKMsHrs") = RST1!AtKMsHrs
                                        .Fields("Serv_Type") = RST1!Serv_Type
                                        .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                                        .Fields("RecBy_MechName") = RST1!RecBy_MechName
                                        .Fields("NetLab_Amt") = RST1!NetLab_Amt
                                        .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                                        .Fields("ObservBy_Super") = RST1!ObservBy_Super
                                        .Fields("ActionBy_Super") = RST1!ActionBy_Super
                                    End If
                                    .Fields("JobDocID") = RST1!DocID
                                    .Fields("Lab_Code") = RST3!Lab_Code
                                    .Fields("Lab_Done") = RST3!Lab_Done
                                    .Update
                                End With
                            End If
                            RST3.MoveNext
                            If RST3.EOF Then
                                ExitLoop = True
                            ElseIf RST3!job_docid <> RST1!DocID Then
                                ExitLoop = True
                            Else
                                RstRep.MoveNext
                                If RstRep.EOF Then mAddRec = True
                            End If
                        Loop
                        
lblRequisition:
                        'Requisition / Parts
                        If Rst4.RecordCount <= 0 Then GoTo lblRst1MoveNext
                        Rst4.MoveFirst
                        Rst4.FIND ("Job_DocID='" & RST1!DocID & "'")
                        If Rst4.EOF Then GoTo lblRst1MoveNext
                        ExitLoop = False
                        mAddRec = False
                        If I = 1 Then
                           RstRep.MoveFirst
                        Else
                            RstRep.Move (2 - (RstRep.RecordCount - RecCountFix))
                            If RstRep.EOF Then mAddRec = True
                        End If
                        Do While Not ExitLoop 'Rst4!Job_DocID = Rst1!DocId
                            If Rst4!job_docid = RST1!DocID Then
                                With RstRep
                                    If mAddRec Then
                                        .AddNew
                                        .Fields("CardNo") = RST1!CardNo
                                        .Fields("JobDocID") = RST1!DocID
                                        .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                                        .Fields("Job_Date") = RST1!Job_Date
                                        .Fields("JobCloseDate") = RST1!JobCloseDate
                                        .Fields("AtKMsHrs") = RST1!AtKMsHrs
                                        .Fields("Serv_Type") = RST1!Serv_Type
                                        .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                                        .Fields("RecBy_MechName") = RST1!RecBy_MechName
                                        .Fields("NetLab_Amt") = RST1!NetLab_Amt
                                        .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                                        .Fields("ObservBy_Super") = RST1!ObservBy_Super
                                        .Fields("ActionBy_Super") = RST1!ActionBy_Super
                                    End If
                                    .Fields("JobDocID") = RST1!DocID
                                    .Fields("DocIDReq") = Rst4!DocIDReq
                                    .Fields("Part_No") = Rst4!Part_No
                                    .Fields("Part_Name") = Rst4!Part_Name
                                    .Fields("Purpose") = Rst4!Purpose
                                    .Fields("Rate") = Rst4!Rate
                                    .Fields("Qty") = Rst4!Qty
                                    .Fields("Amount") = Rst4!Amount
                                    .Update
                                End With
                            End If
                            Rst4.MoveNext
                            If Rst4.EOF Then
                                ExitLoop = True
                            ElseIf Rst4!job_docid <> RST1!DocID Then
                                ExitLoop = True
                            Else
                                RstRep.MoveNext
                                If RstRep.EOF Then mAddRec = True
                            End If
                        Loop
            
lblRst1MoveNext:
                        RecCountFix = RstRep.RecordCount
                        RST1.MoveNext
                    Next
                End If
     
        CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".TTX", True
 '      CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & "1.TTX", True
                
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
        rpt.Database.SetDataSource RstRep
        rpt.ReadRecords
 '       rpt.OpenSubreport("VehHisDet1").Database.SetDataSource RstRep
 '      rpt.OpenSubreport("VehHisDet1").ReadRecords
                
        Set Rst = Nothing
        Set RST1 = Nothing
        Set Rst2 = Nothing
        Set RST3 = Nothing
        Set Rst4 = Nothing
                
        Call Report_View(rpt, Me.CAPTION & "[History]", , False)
        Set rpt = Nothing
        Exit Sub
ELoop:
        MsgBox err.Description, vbCritical, Me.CAPTION
                
End Sub



Private Sub ProcWorkshopVehicleRegister()
On Error GoTo ELoop
Dim mRepName$, mAddRec As Boolean, ExitLoop As Boolean, RecCountFix As Integer
Dim Rst As ADODB.Recordset
Dim mQry$, mQRY1$, Condstr$, CondStr1$, I As Integer


    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and H.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and h.Site_Code ='" & PubSiteCode & "' "
    End If
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Model in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Div_Code in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and CardNo in (" & GridString4 & ")"
    
   mQRY1 = "Select distinct CardNo as SearchCode from HisCard where CardDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and CardDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & CondStr1
            
    RepName = "WorkshopVehicleRegister"
                        
                        
    mQry = "SELECT H.CardNo, H.Name, H.Add1, H.Add2, H.Add3, C.CityName, H.Mobile, " & _
            "H.PhoneOff, H.PhoneResi, H.RegNo , H.Chassis, H.Model " & _
            "FROM HisCard H " & _
            "LEFT JOIN City C ON H.CityCode =C.CityCode " & _
            "Where H.CardNo in (" & mQRY1 & ")" & Condstr & _
            "ORDER BY H.Name "
            
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
   
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
           
RepTitle = UCase(Me.CAPTION)
'    CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".TTX", True
'
'    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
'    rpt.Database.SetDataSource RstRep
'    rpt.ReadRecords
'
'    Set Rst = Nothing
'
'    Call Report_View(rpt, Me.CAPTION & "[History]", , False)
'    Set rpt = Nothing
    Exit Sub
    
ELoop:
    MsgBox err.Description, vbCritical, Me.CAPTION
End Sub



Private Sub WksJobRegister()
On Error GoTo ELoop
Dim mQry$, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
        If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    End If
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    If FGrid.TextMatrix(List1, 1) = "Closed" Then
        Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    ElseIf FGrid.TextMatrix(List1, 1) = "UnClosed" Then
        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and (len(JobCloseDate) = 0 or JC.JobCloseDate is null) "
    ElseIf FGrid.TextMatrix(List1, 1) = "Cancelled" Then
        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    Else
        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    End If
    If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
        If FGrid.TextMatrix(List2, 1) = "Regular" Then
            Condstr = Condstr & " and JC.Jobtype='R'"
        ElseIf FGrid.TextMatrix(List2, 1) = "On Site Repair" Then
            Condstr = Condstr & " and JC.Jobtype='O'"
        ElseIf FGrid.TextMatrix(List2, 1) = "Quick Repair" Then
            Condstr = Condstr & " and JC.Jobtype='Q'"
        End If
    End If
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and JC.Serv_Type in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(JC.DocId,1) in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and JC.RecBy_Supervisor in (" & GridString4 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "Closed" Then
        RepName = "WksJobClosReg"
        Condstr = Condstr & " and (len(JC.JobCloseDate) > 0 or JC.JobCloseDate is not null) and right(JC.DocId_InvSpr,8)<>'Cancelld' "
    ElseIf FGrid.TextMatrix(List1, 1) = "UnClosed" Then
        RepName = "WksJobReg"
        Condstr = Condstr & " and (len(JC.JobCloseDate) = 0 or JC.JobCloseDate is null)"
    ElseIf FGrid.TextMatrix(List1, 1) = "Cancelled" Then
        RepName = "WksJobClosReg"
        Condstr = Condstr & " and right(JC.DocId_InvSpr,8) = 'Cancelld' "
    Else
        RepName = "WksJobClosReg"
    End If
    'By Rahul U.N. Automobiles 11-04-2003
    mQry = " SELECT JC.DocId_InvSpr, JC.DocId_InvLab, JC.Job_Date, JC.Job_No, JC.DocId, JC.JobCloseDate, " & _
             " H.Name,JC.Serv_Type, ST.Serv_Desc, 0 as NetLab_Amt, H.RegNo,H.Model, H.Chassis, " & _
             " " & cIIF("JL.Chrg_From='M' and JL.Chrg_Type='P'", "JL.Labouramt", "0") & " as PDILabour, " & cIIF("JL.Chrg_From='M' and JL.Chrg_Type='F'", "JL.Labouramt", "0") & " as FreeLabour," & _
             " " & cIIF("JL.Chrg_From='C'", "JL.LabourAmt", "0") & " as ChrgLabour," & cIIF("JL.Chrg_From='M' and JL.Chrg_Type='W'", "JL.Labouramt", "0") & "" & _
             " as War_Lab_Rate,JL.LabourAmt, JL.Major_YN,JL.External_YN, 0 as Amount,0 as Net_Amt,0 as Total_Amt,'' as  Purpose,'' as Lub_Category,0 as V_No,0 as MisCharged,EM.Emp_Name " & _
        " FROM (((Job_Card as JC LEFT JOIN Hiscard as H ON JC.CardNo = H.CardNo)" & _
             " LEFT JOIN Job_Lab as JL ON JC.DocId = JL.Job_DocID) LEFT JOIN Emp_Mast as EM ON JC.RecBy_Supervisor = EM.Emp_Code)" & _
             " LEFT JOIN Service_Type as ST ON JC.Serv_Type = ST.Serv_Type " & Condstr & _
        " UNION ALL " & _
        " SELECT JC.DocId_InvSpr, JC.DocId_InvLab, JC.Job_Date, JC.Job_No," & _
             " JC.DocId, JC.JobCloseDate, H.Name,JC.Serv_Type, ST.Serv_Desc, 0 as NetLab_Amt," & _
             " H.RegNo,H.Model, H.Chassis,0 AS PDILabour, 0 AS FreeLabour, 0 AS ChrgLabour," & _
             " 0 AS War_Lab_Rate, 0 AS LabourAmt, 0 AS Major_YN,0 as External_YN,((SP_Stock.qty_iss - SP_Stock.qty_ret)*SP_Stock.Rate) as amount, SP_Stock.Net_Amt, 0 as Total_Amt, SP_Stock.Purpose,SP_Stock.Lub_Category, (SELECT Max(V_No) From SP_Stock Where DocId = JC.Docid_InvSpr) as V_No,0 as MisCharged ,EM.Emp_Name" & _
        " FROM (((Job_Card as JC LEFT JOIN Hiscard as H ON JC.CardNo = H.CardNo)" & _
             " LEFT JOIN Service_Type as ST ON JC.Serv_Type = ST.Serv_Type) LEFT JOIN Emp_Mast as EM ON JC.RecBy_Supervisor = EM.Emp_Code)" & _
             " LEFT JOIN SP_Stock ON JC.DocID = SP_Stock.Job_DocId " & Condstr & _
        " Union All " & _
        " SELECT JC.DocId_InvSpr, JC.DocId_InvLab, JC.Job_Date, JC.Job_No," & _
             " JC.DocId, JC.JobCloseDate,H.Name,JC.Serv_Type, ST.Serv_Desc, JC.NetLab_Amt," & _
             " H.RegNo,H.Model, H.Chassis,0 AS PDILabour, 0 AS FreeLabour, 0 AS ChrgLabour," & _
             " 0 AS War_Lab_Rate, 0 AS LabourAmt, 0 AS Major_YN,0 as External_YN,0 as amount, 0 as Net_Amt, SP_Sale.Total_Amt, '' as Purpose,'' as Lub_Category,SP_Sale.V_No, " & _
             "(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt+SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP+SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP+SP_Sale.Packing+SP_Sale.TOT_AmtMRP+SP_Sale.ReSalTax_Amt+SP_Sale.Rounded) as MisCharged,EM.Emp_Name" & _
        " FROM (((Job_Card as JC LEFT JOIN Hiscard as H ON JC.CardNo = H.CardNo) " & _
             " LEFT JOIN Service_Type as ST ON JC.Serv_Type = ST.Serv_Type)  LEFT JOIN Emp_Mast as EM ON JC.RecBy_Supervisor = EM.Emp_Code) " & _
             " LEFT JOIN SP_Sale ON JC.DocId_InvSpr = SP_Sale.DocID " & Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION) + " [ " + FGrid.TextMatrix(List1, 1) + " ]"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub WksOutLabRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    Condstr = " where JL.External_YN='1' and JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(JC.DocId,1) in (" & GridString2 & ")"
    
'    If FGrid.TextMatrix(List1, 1) = "UnClosed" Then CondStr = CondStr & " and JC.JobCloseDate Is Null"
'    If FGrid.TextMatrix(List1, 1) = "Closed" Then CondStr = CondStr & " and not JC.JobCloseDate Is Null"
    '
    mQry = " SELECT JC.DocId_InvSpr, JC.DocId_InvLab, JC.Job_Date, JC.Job_No, JC.DocId, JC.JobCloseDate,H.Name, ST.Serv_Desc, H.RegNo, H.Chassis," & _
                " JL.Lab_Code,L.Lab_Desc,JL.LabourAmt,JL.ExtJobGatePassNo,JGP.GatePassDate,JGP.ContractRecdDate,JGP.ContractAmt,JL.Contract_Remarks " & _
           " FROM (((((Job_Card as JC LEFT JOIN HisCard as H ON JC.CardNo = H.CardNo)" & _
                " LEFT JOIN Job_Lab as JL ON JC.DocId = JL.Job_DocID)" & _
                " LEFT JOIN Labour as L ON JL.Lab_Code = L.Lab_Code)" & _
                " LEFT JOIN Service_Type as ST ON JC.Serv_Type = ST.Serv_Type) " & _
                " LEFT JOIN Job_GatePass as JGP ON JL.ExtJobGatePassNo = JGP.GatePassNo) " & _
                Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksOutLabReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub OverTimeRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    Condstr = " where OT.OT_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and OT.OT_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("OT.Site_Code", "2", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("OT.Site_Code", "2", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and OT.Emp_Code in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and OT.Div_Code in (" & GridString3 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "Detail" Then RepName = "OverTimeRegDet"  'CondStr = CondStr & " and OT.JobCloseDate Is Null"
    If FGrid.TextMatrix(List1, 1) = "Summary" Then RepName = "OverTimeRegSum" 'CondStr = CondStr & " and not OT.JobCloseDate Is Null"
    
    mQry = " SELECT E.OT_Rate as OTRate,OT.Div_Code,OT.OT_Date,OT.Emp_Code,E.Emp_Name,OT.HrMinute,(" & cVal("left(OT.HrMinute,2))*60") & " + " & cVal("right(OT.HrMinute,2)") & " as OTHrMin," & _
            "OT.Site_Code,OT.Remarks,D.Designation " & _
           " FROM ((OverTime as OT LEFT JOIN Emp_Mast as E ON OT.Emp_Code=E.Emp_Code)" & _
            " LEFT JOIN Designation as D ON E.Designation=D.Designation)" & _
            Condstr
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'    RepName = "WksOutLabReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Public Sub SelGridKeyPressLocal(Txt As Object, SelGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName$, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    '$
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If SelGrid(Index).Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then Txt.TEXT = "": Exit Sub
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack Then
            If Len(Txt.SelText) > 1 Then
                Txt.SelLength = Len(Txt.SelText) - 1
                FindStr = Txt.SelText
            Else
                Txt.TEXT = ""
                SelGrid(Index).SetFocus
                Txt.Visible = False
                Exit Sub
            End If
        Else
            FindStr = Txt.SelText + Chr(KeyAscii)
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
            Txt.TEXT = Rst.Fields(FindFldName).Value
            Txt.SelLength = Len(FindStr)
            Txt.left = SelGrid(Index).CellLeft + SelGrid(Index).left
            Txt.top = SelGrid(Index).CellTop + SelGrid(Index).top
            If Txt.Visible = False Then
                Txt.Visible = True: Txt.ZOrder 0: Txt.SetFocus: Txt.BackColor = SelGrid(Index).CellBackColor
                 Txt.ForeColor = SelGrid(Index).CellForeColor: Txt.width = SelGrid(Index).CellWidth: Txt.height = SelGrid(Index).CellHeight
            End If
       End If
End Sub
Private Sub JobBookRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$

    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where ForServiceDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and ForServiceDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Site_code in (" & GridString1 & ") "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Div_code in (" & GridString2 & ") "
    
    If FGrid.TextMatrix(List1, 1) = "Pending" Then
        Condstr = Condstr & " and len(Job_Booking.Job_DocId)=0"
    ElseIf FGrid.TextMatrix(List1, 1) = "Done" Then
        Condstr = Condstr & " and len(Job_Booking.Job_DocId) > 0"
    End If
     
    mQry = "SELECT Job_Booking.Book_Date,Job_Booking.Name,Job_Booking.Add1,Job_Booking.Add2,Job_Booking.Add3,Job_Booking.PhoneOff," & _
            "Job_Booking.PhoneResi,Job_Booking.Mobile,Job_Booking.Model,Job_Booking.RegNo,Job_Booking.ForServiceDate,City.CityName," & _
            "Service_type.Serv_Desc From (Job_Booking Left Join City on Job_booking.CityCode=City.CityCode) " & _
            "Left Join Service_type on Job_Booking.Service_type=Service_type.Serv_type "
    
    mQry = mQry + Condstr + " order by ForServiceDate"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "JobBookReg"
    

    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ProcInsuranceExpiryRegister()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

mQry = "SELECT H.Name, H.RegNo, H.Model, H.Add1, H.add2, C.CityCode, H.PhoneOff, H.PhoneResi, H.Mobile, I.Name as InsuranceCompanyName, H.InsuranceExpiry as EndDate, H.Chassis " & _
       "FROM HisCard H " & _
       "LEFT JOIN Insurance I ON H.InsuranceCompany =I.Code " & _
       "LEFT JOIN City C ON C.CityCode = H.CityCode  " & _
       "Where H.InsuranceExpiry >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and H.InsuranceExpiry  <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & _
       "Order By H.Name "
           
           
    Set RstRep = New ADODB.Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    RepName = "InsuranceExpiryRegister"
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description

End Sub


