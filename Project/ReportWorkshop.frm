VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form ReportWorkShop 
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
      DownPicture     =   "ReportWorkshop.frx":0000
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
      DownPicture     =   "ReportWorkshop.frx":3132
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
Attribute VB_Name = "ReportWorkShop"
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
Dim FormulaStr1 As String, FormulaStr2 As String, FormulaStr3 As String, FormulaStr4 As String
Dim RstRep As ADODB.Recordset
Dim RstRep1 As ADODB.Recordset
Dim SubRep1 As Boolean
Private Const GridRowHeight As Integer = 270
'////////********WORKSHOP***********////////////////////*****
Private Const LabHrVar As Byte = 1
Private Const ServWiseJob As Byte = 3
Private Const ModWiseJob As Byte = 4
Private Const DemdVsSupp As Byte = 5 'Disable due merging of Works Requisition & Store Requisition
Private Const MechEarnRep As Byte = 6
Private Const MechEarnSum As Byte = 7
Private Const Dewisevehat As Byte = 8
Private Const MoWiseSprInv As Byte = 9
Private Const AgGpWiseInv As Byte = 10
Private Const JobWiseLabA As Byte = 11
Private Const DeWiseJobAna As Byte = 12
Private Const DelayResnAna As Byte = 13
Private Const WksRateVar As Byte = 14
Private Const CancelRep As Byte = 15
Private Const WksLabIncent As Byte = 16
Private Const VehGrdReg As Byte = 17
Private Const ModWiseSrvTax As Byte = 18
Private Const LabRevnueRep As Byte = 19
Private Const ServDueReg As Byte = 20
Private Const QuatRetReg As Byte = 21
Private Const PostServiceFollow As Byte = 22
Private Const STaxWSrvTax As Byte = 23
Private Const ReptJobAna As Byte = 24


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
    Case LabHrVar
        LabHrVarProc
        If RepPrint = False Then Exit Sub
    Case ServWiseJob
        ServWiseJobProc
        If RepPrint = False Then Exit Sub
    Case ModWiseSrvTax
        ModWiseSrvTaxProc
        If RepPrint = False Then Exit Sub
    Case ModWiseJob
        ModWiseJobProc
        If RepPrint = False Then Exit Sub
    Case DemdVsSupp
        DemdVsSuppProc
        If RepPrint = False Then Exit Sub
    Case MechEarnRep
        MechEarnRepProc
        If RepPrint = False Then Exit Sub
    Case MechEarnSum
        MechEarnSumProc
        If RepPrint = False Then Exit Sub
    Case Dewisevehat
        DeWiseVehAtProc
        If RepPrint = False Then Exit Sub
    Case MoWiseSprInv
        MoWiseSprInvProc
        If RepPrint = False Then Exit Sub
'    Case AgGpWiseInv
'        AgGpWiseInvProc
'        If RepPrint = False Then Exit Sub
    Case DeWiseJobAna
        DeWiseJobAnaProc
        If RepPrint = False Then Exit Sub
    Case DelayResnAna
        DelayResnAnaProc
        If RepPrint = False Then Exit Sub
    Case WksRateVar
        WksRateVarProc
        If RepPrint = False Then Exit Sub
    Case CancelRep
        CancelRepProc
        If RepPrint = False Then Exit Sub
    Case JobWiseLabA
        JobWiseLabAProc
        If RepPrint = False Then Exit Sub
    Case VehGrdReg
        VehGrdRegProc
        If RepPrint = False Then Exit Sub
    Case WksLabIncent
        WksLabIncentProc
        If RepPrint = False Then Exit Sub
    Case LabRevnueRep
        LabRevnueRepProc
        If RepPrint = False Then Exit Sub
    Case ServDueReg
        ServDueRegProc
        If RepPrint = False Then Exit Sub
    Case QuatRetReg
        QuatRetRegProc
        If RepPrint = False Then Exit Sub
    Case PostServiceFollow
        PostServiceFollowProc
        If RepPrint = False Then Exit Sub
    Case STaxWSrvTax
        STaxWSrvTaxRegProc
        If RepPrint = False Then Exit Sub
    Case ReptJobAna
        ReptJobAnaProc
        If RepPrint = False Then Exit Sub
        

End Select
'Nra Updation
If GRepFormName = MechEarnRep Or GRepFormName = MechEarnSum Then
    CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & ".ttx", True
Else
    CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
End If
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")
If GRepFormName = MechEarnRep Or GRepFormName = MechEarnSum Then
    rpt.Database.SetDataSource RstRep1
Else
    rpt.Database.SetDataSource RstRep
End If
'End Updation
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
rpt.ReadRecords
Set RstRep = Nothing
Set RstRep1 = Nothing
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
        Case WksRateVar
                ListArray = Array("High", "Low", "Both")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case LabHrVar  'vijay for work shop '16/11/02
              ListArray = Array("Hour", "Rate")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case DelayResnAna  'vijay for work shop '16/11/02
              ListArray = Array("Summary", "Detail")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case DemdVsSupp   'vijay dt 16/11/02
              ListArray = Array("All", "PDI", "Free Service", "Chargable", "Warranty", "Company Vehicle", "Complementary", "AMC")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 8)
            Case ReptJobAna
                ListArray = Array("Days", "KMS")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
          End Select
    Case List2
        Select Case GRepFormName
'        Case WksReqReg
'            ListArray = Array("All", "PDI", "Free Service", "Chargable", "Warranty", "Company Vehicle", "Complementary")
'            Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 7)
        End Select
'    Case List3
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
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
    Select Case FGrid.Row
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
'***
Ini_Grid
'***
For I = 1 To 4
    If GridSel(I).Visible = True Then Cnt = Cnt + 1
Next
For I = mFirstRow To mLastRow
    GridHeight = GridHeight + FGrid.RowHeight(I)
Next
FGrid.height = GridHeight + FGrid.RowHeight(mFirstRow)

'FGrid.Height = (((mLastRow + 1) - mFirstRow) * PubGridRowHeight) + 500
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
Dim formulastr As String   'Modishekhar 17 mar
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
            FormulaStr1 = mID(FormulaStr1, 1, 254)
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
    If GRepFormName = 15 Then
        RsGrid2.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
    Else
        RsGrid2.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
    End If
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
'Date1,Date2,List1,List1,List2,List3
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where  site_code ='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
Select Case GRepFormName
    Case WksLabIncent
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
'            .TextMatrix(List1, 0) = "Mechanic Share": .RowHeight(List1) = GridRowHeight
'            .TextMatrix(List2, 0) = "Helper Share": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
'            .TextMatrix(List1, 1) = "60%"
'            .TextMatrix(List2, 1) = "40%"
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Emp_Mast.Emp_Name as Employee_Name,Emp_Mast.Emp_Code as code from Emp_Mast Where Div_Code='" & PubDivCode & "' order by Emp_Mast.Emp_Name"
        GridInitialise 3, Grid3Sql
    Case WksRateVar
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "High/Low/Both": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Cat1, 0) = "Minimum Difference": .RowHeight(Cat1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Both"
            .TextMatrix(Cat1, 1) = "0.00"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Cat1

    Case MechEarnRep, MechEarnSum
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 3
        ' By Rahul U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Emp_Mast.Emp_Name as Mechanic,Emp_Mast.Emp_Code as code from Emp_Mast Where Designation='MECHANIC' Order by Emp_Mast.Emp_Name"
        GridInitialise 2, Grid2Sql
 
        Grid3Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 3, Grid3Sql
        
    Case LabRevnueRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
        ' By Rahul U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & "  order by site_desc"
        GridInitialise 2, Grid2Sql
    Case ServDueReg, PostServiceFollow, STaxWSrvTax
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 1
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
    Case QuatRetReg
        With FGrid
            .TextMatrix(Date1, 0) = "As On Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date1: mHelpGridNo = 0
    Case VehGrdReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
        ' By Rahul U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 2, Grid2Sql
  Case CancelRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
'       By Rahul At U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        
        Grid2Sql = "select distinct '' as O, category AS V_Category,category AS Code from voucher_type order by category "
        GridInitialise 2, Grid2Sql

  Case ServWiseJob, MoWiseSprInv
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2
 Case Dewisevehat, DeWiseJobAna ',MoWiseSprInv
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 3
'       By Rahul At U.N.Automobile Udaipur 11-04-2003

        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,AMD_Dealer.D_NAME as Dealer_Name,AMD_Dealer.D_Code as code from AMD_Dealer order by AMD_Dealer.D_Name"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Where Div_Code='" & PubDivCode & "'  order by Div_Name"
        GridInitialise 3, Grid3Sql


    Case ModWiseJob
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 3
        
'       By Rahul At U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Model_Desc as Model_Name,Model as code from Model order by Model_Desc"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 3, Grid3Sql
     Case ModWiseSrvTax, STaxWSrvTax
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case LabHrVar    'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Hour/Rate": .RowHeight(List1) = GridRowHeight
'            .TextMatrix(List2, 0) = "Description Of Labour": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Hour"
'            .TextMatrix(List2, 1) = "Standard"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 3
'       By Rahul At U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
        GridInitialise 2, Grid2Sql
     
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 3, Grid3Sql

    Case DelayResnAna    'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Detail/Summary": .RowHeight(List1) = GridRowHeight
'            .TextMatrix(List2, 0) = "Description Of Labour": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Summary"
'            .TextMatrix(List2, 1) = "Standard"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 3
'       By Rahul At U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        
        Grid2Sql = "select '' as O,R_Desc as Delay_Reason,Code from Job_Delay order by R_desc"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Where Div_Code='" & PubDivCode & "'  order by Div_Name"
        GridInitialise 3, Grid3Sql

     Case DemdVsSupp   'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Part Purpose": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 3
'       By Rahul At U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case JobWiseLabA   'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
'       By Rahul At U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Lab_Desc as LabourName,Lab_code  as code from Labour Where Div_Code='" & PubDivCode & "' order by Lab_desc"
        GridInitialise 2, Grid2Sql
                
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Where Div_Code='" & PubDivCode & "'  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        Grid2Sql = "select '' as O,Lab_Desc as LabourName,Lab_code  as code from Labour Where Div_Code='" & PubDivCode & "' order by Lab_desc"
        GridInitialise 2, Grid2Sql
        
        Grid4Sql = "select '' as O,Emp_Mast.Emp_Name as Mechanic,Emp_Mast.Emp_Code as code from Emp_Mast Where Designation='MECHANIC' Order by Emp_Mast.Emp_Name"
        GridInitialise 4, Grid4Sql
        
    Case ReptJobAna
    
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Rep.Period Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Cat1, 0) = "Repeat Period": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Days"
            .TextMatrix(Cat1, 1) = ""
            
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Cat1: mHelpGridNo = 0
        
        

End Select
End Sub
Private Function IsNotBlank(FieldRow As Integer, FieldCaption As String) As Boolean
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
    Case WksLabIncent
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
    Case ModWiseSrvTax, STaxWSrvTax, ReptJobAna
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
    Case WksRateVar
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

Case LabHrVar
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("HOURORRATE")
                rpt.FormulaFields(I).TEXT = "'Variation For Labour ' + '" & FGrid.TextMatrix(List1, 1) & "'"
    '        Case UCase("STDORBILL")
    '            rpt.FormulaFields(i).Text = "'For ' + '" & FGrid.TextMatrix(List2, 1) & "'"
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next
    
Case DelayResnAna
    For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("REASON")
            rpt.FormulaFields(I).TEXT = "'Reason With ' + '" & FGrid.TextMatrix(List1, 1) & "'"
'        Case UCase("STDORBILL")
'            rpt.FormulaFields(i).Text = "'For ' + '" & FGrid.TextMatrix(List2, 1) & "'"
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
    End Select
    Next
Case DemdVsSupp  'vijay 16/11/02
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("List1")
                rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Service'"
        End Select
    Next
Case ServWiseJob, ModWiseJob, MechEarnRep, MechEarnSum, Dewisevehat, JobWiseLabA, DeWiseJobAna, CancelRep, MoWiseSprInv 'vijay 16/11/02
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next
Case LabRevnueRep
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub LabHrVarProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    Condstr = " and Job_Card.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Job_Card.Serv_Type in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(Job_Card.DocId,1) in (" & GridString3 & ")"
    If FGrid.TextMatrix(List1, 1) = "Hour" Then 'CondStr = CondStr & " and isnull(Job_Card.JobCloseDate)"
        mQry = "SELECT Job_Card.DocId,Job_Card.Job_Date, Job_Card.Job_No, Job_Card.JobCloseDate, HisCard.Inv_No, HisCard.RegNo, HisCard.Model, HisCard.Chassis, Labour.Lab_Desc,Job_Lab.Hrs_Taken, Labour.Time_Req,Labour.Lab_Rate as LabRate,Job_Lab.Lab_Rate as JobRate " & _
               "FROM ((Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) LEFT JOIN Job_Lab ON Job_Card.docid = Job_Lab.Job_DocID) LEFT JOIN LABOUR ON Labour.LAb_code=job_lab.lab_code where Job_Lab.Hrs_Taken  > 0 "
    End If
    If FGrid.TextMatrix(List1, 1) = "Rate" Then ' CondStr = CondStr & " and isnotnull(Job_Card.JobCloseDate)"
        mQry = "SELECT Job_Card.DocId,Job_Card.Job_Date, Job_Card.Job_No, Job_Card.JobCloseDate, HisCard.Inv_No, HisCard.RegNo, HisCard.Model, HisCard.Chassis, Labour.Lab_Desc,Job_Lab.Lab_Rate, Labour.Lab_Rate " & _
               "FROM ((Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) LEFT JOIN Job_Lab ON Job_Card.docid = Job_Lab.Job_DocID) LEFT JOIN LABOUR ON Labour.LAb_code=job_lab.lab_code where Job_Lab.Lab_Rate  > 0 "
    End If
    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List1, 1) = "Hour" Then
        RepName = "LabHrVar"
    End If
    If FGrid.TextMatrix(List1, 1) = "Rate" Then
        RepName = "LabRateVar"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub CancelRepProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
  
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where System_Log.U_EntDt  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and System_Log.U_EntDt <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(System_Log.site_code,1) in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then CondStr1 = " AND " & cMID("System_Log.Old_DocumentID", "4", "5") & " IN (select V_Type from voucher_Type where category in (" & GridString2 & "))"
      'AND mid(Old_DocumentID,4,5) IN (select V_Type from voucher_Type where category in (" & GridString2 & "))
    
    mQry = "SELECT System_Log.WS_ID, System_Log.U_Name, System_Log.U_EntDt, System_Log.New_String, System_Log.Old_DocumentID, System_Log.Old_Amount, System_Log.Related_Detail" & _
           " From System_Log"
        
    mQry = mQry + Condstr + CondStr1

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "CancelRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub WksRateVarProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim CompOperater As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    Condstr = " and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

    If FGrid.TextMatrix(List1, 1) = "High" Then
        CompOperater = " >"
    ElseIf FGrid.TextMatrix(List1, 1) = "Low" Then
        CompOperater = "<"
    Else
        CompOperater = "<>"
    End If
'    If GRepFormName = SprCtrRateVari Then
    mQry = "SELECT SP_Stock.V_Type,SP_Stock.DocID,SP_Stock.V_No, SP_Stock.V_Date," & _
           "SP_Stock.Qty_Iss,SP_Stock.Rate,SP_Stock.Part_No, Part.Part_Name," & _
           "" & cIIF("SP_Stock.MRP_YN = 1", "Part.MRP", cIIF("SP_Stock.Tax_YN = 1", "Part.TB_SRate", "Part.TP_SRate")) & " as PartRate," & cIIF("SP_Stock.Tax_YN = 1", "Part.TB_Effect_Dt", "MRP_Effect_Dt") & " as EffDate " & _
            "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) " & _
            "where SP_Stock.Rate " & CompOperater & " " & cIIF("SP_Stock.MRP_YN = 1", "Part.MRP", cIIF("SP_Stock.Tax_YN = 1", "Part.TB_SRate", "Part.TP_SRate")) & " and sp_stock.v_type in ('W_RG', 'W_RW')" '('W_SIC','W_SIR')"
'    Else
'        mQRY = "SELECT SP_Stock.V_Type, SP_Stock.V_No, SP_Stock.V_Date, SP_Stock.DocID," & _
'        "SP_Stock.Qty_Rec as Qty_Iss,SP_Stock.Rate,SP_Stock.Part_No, Part.Part_Name," & _
'        "iif(SP_Stock.MRP_YN = 1,Part.MRP - Part_DiscFactor.PurcDisc_Per,iif(SP_Stock.Tax_YN = 1,Part.TB_SRate -Part_DiscFactor.PurcDisc_Per,Part.TP_SRate - Part_DiscFactor.PurcDisc_Per)) as PartRate,iif(SP_Stock.Tax_YN = 1,Part.TB_Effect_Dt,MRP_Effect_Dt) as EffDate " & _
'        "FROM (SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO) " & _
'        "LEFT JOIN Part_DiscFactor ON Part.Disc_Factor = Part_DiscFactor.DiscFac_Catg " & _
'        "where SP_Stock.Rate " & CompOperater & " iif(SP_Stock.MRP_YN = 1,Part.MRP - Part_DiscFactor.PurcDisc_Per,iif(SP_Stock.Tax_YN = 1,Part.TB_SRate - Part_DiscFactor.PurcDisc_Per,Part.TP_SRate - Part_DiscFactor.PurcDisc_Per)) and sp_stock.v_type in ('SXPIC','SXPIR')"
'    End If
    mQry = mQry + Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksRateVar"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub DelayResnAnaProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Job_Card.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    CondStr1 = " AND JobCloseDate Is Not Null AND " & cDt("Job_Card.JobCloseDate") & " <> " & cDt("Job_Card.ExpDelDate") & ""
    
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Job_Card.DelayReason in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(Job_Card.DocId,1) in (" & GridString3 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "Summary" Then 'CondStr = CondStr & " and isnull(Job_Card.JobCloseDate)"
    
    mQry = "SELECT Count(job_card.docid) AS NoOfVeh ,Job_Delay.R_Desc " & _
           "FROM (Job_Card LEFT JOIN Job_Delay ON Job_Card.DelayReason = Job_Delay.Code) LEFT JOIN Service_Type ON Job_Card.Serv_Type = Service_Type.Serv_Type " & Condstr & " " & CondStr1 & " " & _
           " Group By Job_Delay.R_desc"
    End If
    If FGrid.TextMatrix(List1, 1) = "Detail" Then ' CondStr = CondStr & " and isnotnull(Job_Card.JobCloseDate)"
    
    mQry = "SELECT Job_Card.Job_Date, Job_Card.Job_No, Service_Type.Serv_Desc, Job_Card.JobCloseDate, " & _
           "Job_Card.ExpDelDate, HisCard.Model, HisCard.RegNo, HisCard.Chassis" & _
           " FROM ((Job_Card LEFT JOIN Job_Delay ON Job_Card.DelayReason = Job_Delay.Code) " & _
           "LEFT JOIN Service_Type ON Job_Card.Serv_Type = Service_Type.Serv_Type) " & _
           "LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo " & Condstr & " " & CondStr1 & " "

    End If
    

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    If FGrid.TextMatrix(List1, 1) = "Summary" Then
    RepName = "DelayResnAnaSum"
    End If
    If FGrid.TextMatrix(List1, 1) = "Detail" Then
    RepName = "DelayResnAnaDet"
    End If
    
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub MechEarnRepProc()
On Error GoTo ELoop
' Nra Updation
Dim mQry$, Condstr$, CondStr1$, MySql1$, MySql$, NoVeh, NormalLabour As Long, WarLabour As Long, FreeLabour As Long, ContLabour As Long, FreeServLabour As Long, TotalLabour As Integer, I As Integer
Dim Rst As ADODB.Recordset, rstMec As ADODB.Recordset, RST1 As ADODB.Recordset, MyRst As ADODB.Recordset, RstMech As ADODB.Recordset
Set MyRst = New ADODB.Recordset
Set RstMech = New ADODB.Recordset
Set Rst = New ADODB.Recordset
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    'If Check1(1).Value = Unchecked Then CondStr1 = CondStr1 & " and left(J.DocId,1) in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then CondStr1 = CondStr1 & " and JL2.Mech_Code in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then CondStr1 = CondStr1 & " and " & cMID("Jl.job_Docid", "3", "1") & " in (" & GridString3 & ")"
     If Check1(3).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("Jl.job_Docid", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Total
'  1           2       3       4             5       6         7           8          9           10             11              12
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Emp_Code", adChar, 4, adFldIsNullable
        .Fields.Append "Emp_Name", adVarChar, 40, adFldIsNullable
        .Fields.Append "JobClosDate", adVarChar, 17, adFldIsNullable
        .Fields.Append "Job_DocID", adVarChar, 21, adFldIsNullable
        .Fields.Append "Serv_Type", adVarChar, 2, adFldIsNullable
        .Fields.Append "Serv_Detail", adVarChar, 20, adFldIsNullable
        .Fields.Append "Model", adVarChar, 15, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 15, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 14, adFldIsNullable
        .Fields.Append "Lab_Paid", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Warr", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_OthSelf", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Mfg", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Contract", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Disc", adDouble, 12, adFldIsNullable
        .Fields.Append "Net_Amt", adDouble, 12, adFldIsNullable
        .Fields.Append "S_No", adDouble, 4, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set RstRep1 = New ADODB.Recordset
    With RstRep1
        .Fields.Append "Emp_Code", adChar, 7, adFldIsNullable
        .Fields.Append "Emp_Name", adVarChar, 40, adFldIsNullable
        .Fields.Append "JobClosDate", adVarChar, 17, adFldIsNullable
        .Fields.Append "Job_DocID", adVarChar, 21, adFldIsNullable
        .Fields.Append "Serv_Type", adVarChar, 2, adFldIsNullable
        .Fields.Append "Serv_Detail", adVarChar, 20, adFldIsNullable
        .Fields.Append "Model", adVarChar, 15, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 15, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 14, adFldIsNullable
        .Fields.Append "Lab_Paid", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Warr", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_OthSelf", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Mfg", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Contract", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Disc", adDouble, 12, adFldIsNullable
        .Fields.Append "Net_Amt", adDouble, 12, adFldIsNullable
        .Fields.Append "S_No", adDouble, 4, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Discount Total
'  1           2       3       4             5       6         7           8          9           10             11              12      12

'No of Mechanic
        
'    mQRY = "Select DocID,JobClosDate,Serv_Type,Model,Chassis,RegNo,Emp_Code,Emp_Name," & _
'        "from ((Job_Card as J left Join Hiscard as H on J.CardNo=H.CardNo))" & _
'        "left Join Job_Lab2 as JL2 on J.DocID=JL2.Job_DocID " & _
'        "left Join Emp_Mast as E on JL2.Mech_Code=E.Emp_Code "


'Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour
On Error Resume Next
'TOTAL LABOUR DONE IN GIVEN CONDITION
    GSQL = "SELECT JL.Job_DocId,JL.s_No,JL.hrs_taken,JL.lab_rate," & _
        " JL.hrs_war,JL.war_lab_rate,JL.labourAmt,JL.Chrg_From,JL.Chrg_Type," & _
        " JL.ExtJobGatePassNo,JL.Job_DocID,JC.JobCloseDate" & _
        " From ((Job_Lab as JL left join Job_GatePass as GP on JL.ExtJobGatePassNo=GP.GatePassNo)left join Job_Card as JC on JL.Job_DocID=JC.DocID)" & Condstr & _
        " Order by JL.Job_DocID,JL.s_No,JC.JobCloseDate desc"
        ' JL.Job_DocId=GP.Job_DocId
        
    MyRst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    MyRst.MoveFirst
    Dim MyID As String
    Dim S_No As Integer
    MyID = IIf(IsNull(MyRst!job_docid), "", MyRst!job_docid)
    S_No = IIf(IsNull(MyRst!S_No), "", MyRst!S_No)
    RstRep.AddNew
    RstRep!Lab_warr = IIf(IsNull(RstRep!Lab_warr), 0, RstRep!Lab_warr)
    RstRep!Lab_mfg = IIf(IsNull(RstRep!Lab_mfg), 0, RstRep!Lab_mfg)
    RstRep!Lab_OthSelf = IIf(IsNull(RstRep!Lab_OthSelf), 0, RstRep!Lab_OthSelf)
    RstRep!Lab_Paid = IIf(IsNull(RstRep!Lab_Paid), 0, RstRep!Lab_Paid)
    RstRep!job_docid = MyRst!job_docid
    RstRep!JobClosDate = IIf(IsNull(MyRst!JobCloseDate), "", MyRst!JobCloseDate)
'TOTAL NO OF MECHANIC ON A PERTICULAR JOB AND SR NO.
    GSQL = "SELECT JL2.Job_DocID,JL2.Lab_Code as SearchCode,JL2.s_No,count(Mech_Code) as NoOfMech" & _
        " from ((Job_Lab2 as JL2 left join Job_Card as J on JL2.Job_DocID=J.Docid) " & _
        " LEFT JOIN Emp_Mast on Emp_Mast.Emp_Code=JL2.Mech_Code)" & CondStr1 & _
        " Group By JL2.Job_DocID,JL2.Lab_Code,JL2.s_No"

    RstMech.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    Dim onePart_Warr As Integer
    Dim onePart_mfg As Integer
    Dim onePart_OthSelf As Integer
    Dim onePart_LabPaid As Integer
    
    Dim lastPart_Warr As Integer
    Dim lastPart_mfg As Integer
    Dim lastPart_OthSelf As Integer
    Dim lastPart_LabPaid As Integer, Cnt As Integer
    
    For Cnt = 0 To MyRst.RecordCount
       If MyID <> IIf(IsNull(MyRst!job_docid), "", MyRst!job_docid) Or S_No <> IIf(IsNull(MyRst!S_No), "", MyRst!S_No) Then
           RstMech.MoveFirst
          While Not RstMech.EOF = True
             If MyID = RstMech!job_docid And S_No = RstMech!S_No Then
'EMPLOYEE NAME REGARDING PERTICULAR MECH. ID.
               GSQL = "SELECT JL2.Job_DocID,JL2.s_No,Emp_Mast.Emp_Name" & _
                      " from (Job_Lab2 as JL2 LEFT JOIN Emp_Mast on Emp_Mast.Emp_Code=JL2.Mech_Code)" & _
                      " where  JL2.Job_DocID= '" & RstMech!job_docid & "' and JL2.s_No=" & RstMech!S_No & ""
               
               Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
           ' Mechanic wise labour devidation
                    If RstMech!NoOfMech > 1 Then
                       onePart_Warr = IIf(Val(RstRep!Lab_warr) = 0, 0, Val(RstRep!Lab_warr) / RstMech!NoOfMech)
                       onePart_mfg = IIf(Val(RstRep!Lab_mfg) = 0, 0, Val(RstRep!Lab_mfg) / RstMech!NoOfMech)
                       onePart_OthSelf = IIf(Val(RstRep!Lab_OthSelf) = 0, 0, Val(RstRep!Lab_OthSelf) / RstMech!NoOfMech)
                       onePart_LabPaid = IIf(Val(RstRep!Lab_Paid) = 0, 0, Val(RstRep!Lab_Paid) / RstMech!NoOfMech)
                       
                       lastPart_Warr = IIf(Val(RstRep!Lab_warr) = 0, 0, Val(RstRep!Lab_warr) - (Val(onePart_Warr) * (RstMech!NoOfMech - 1)))
                       lastPart_mfg = IIf(Val(RstRep!Lab_mfg) = 0, 0, Val(RstRep!Lab_mfg) - (Val(onePart_mfg) * (RstMech!NoOfMech - 1)))
                       lastPart_OthSelf = IIf(Val(RstRep!Lab_OthSelf) = 0, 0, Val(RstRep!Lab_OthSelf) - (Val(onePart_OthSelf) * (RstMech!NoOfMech - 1)))
                       lastPart_LabPaid = IIf(Val(RstRep!Lab_Paid) = 0, 0, Val(RstRep!Lab_Paid) - (Val(onePart_LabPaid) * (RstMech!NoOfMech - 1)))
                       
                       For I = 1 To RstMech!NoOfMech
                            If I <= RstMech!NoOfMech - 1 Then
                                RstRep1.AddNew
                                RstRep1!Emp_Name = Rst!Emp_Name
                                RstRep1!JobClosDate = RstRep!JobClosDate
                                RstRep1!Lab_warr = onePart_Warr
                                RstRep1!Lab_mfg = onePart_mfg
                                RstRep1!Lab_OthSelf = onePart_OthSelf
                                RstRep1!Lab_Paid = onePart_LabPaid
                                RstRep!Lab_warr = 0
                                RstRep!Lab_mfg = 0
                                RstRep!Lab_OthSelf = 0
                                RstRep!Lab_Paid = 0
                            Else
                                RstRep1.AddNew
                                RstRep1!Emp_Name = Rst!Emp_Name
                                RstRep1!JobClosDate = RstRep!JobClosDate
                                RstRep1!Lab_warr = lastPart_Warr
                                RstRep1!Lab_mfg = lastPart_mfg
                                RstRep1!Lab_OthSelf = lastPart_OthSelf
                                RstRep1!Lab_Paid = lastPart_LabPaid
                                RstRep!Lab_warr = 0
                                RstRep!Lab_mfg = 0
                                RstRep!Lab_OthSelf = 0
                                RstRep!Lab_Paid = 0
                            End If
                            Rst.MoveNext
                       Next
                       Else
                                RstRep1.AddNew
                                RstRep1!Emp_Name = Rst!Emp_Name
                                RstRep1!JobClosDate = RstRep!JobClosDate
                                RstRep1!Lab_warr = RstRep!Lab_warr
                                RstRep1!Lab_mfg = RstRep!Lab_mfg
                                RstRep1!Lab_OthSelf = RstRep!Lab_OthSelf
                                RstRep1!Lab_Paid = RstRep!Lab_Paid
                                RstRep!Lab_warr = 0
                                RstRep!Lab_mfg = 0
                                RstRep!Lab_OthSelf = 0
                                RstRep!Lab_Paid = 0
                                Rst.MoveNext
                    End If
                End If
                RstMech.MoveNext
                If Rst.State = 1 Then: Rst.Close
            Wend
                       
           ' End devidation
           
        
            RstRep!Lab_warr = IIf(IsNull(RstRep!Lab_warr), 0, RstRep!Lab_warr)
            RstRep!Lab_mfg = IIf(IsNull(RstRep!Lab_mfg), 0, RstRep!Lab_mfg)
            RstRep!Lab_OthSelf = IIf(IsNull(RstRep!Lab_OthSelf), 0, RstRep!Lab_OthSelf)
            RstRep!Lab_Paid = IIf(IsNull(RstRep!Lab_Paid), 0, RstRep!Lab_Paid)
            RstRep!JobClosDate = IIf(IsNull(MyRst!JobCloseDate), "", MyRst!JobCloseDate)
            RstRep!job_docid = MyRst!job_docid
            RstRep!S_No = MyRst!S_No
            MyID = RstRep!job_docid
            S_No = RstRep!S_No
        End If
        ' Catagory wise Labour computing
        
        Select Case MyRst!Chrg_From
        Case "M"
            If MyRst!Chrg_Type = "W" Then
                RstRep!Lab_warr = RstRep!Lab_warr + MyRst!LabourAmt
            Else
                RstRep!Lab_mfg = RstRep!Lab_mfg + MyRst!LabourAmt
            End If
        Case "S", "O"
                RstRep!Lab_OthSelf = RstRep!Lab_OthSelf + MyRst!LabourAmt
        Case "C"
            RstRep!Lab_Paid = RstRep!Lab_Paid + MyRst!LabourAmt
        End Select
        
    
    MyRst.MoveNext
Next
    If MyRst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "MechEarnRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub LabRevnueRepProc()
On Error GoTo ELoop
Dim mQry$, Condstr$, Condstr2$, MySql1$, MySql$, NoVeh, NormalLabour As Long, WarLabour As Long, FreeLabour As Long, ContLabour As Long, FreeServLabour As Long, TotalLabour As Integer
Dim Rst As ADODB.Recordset, rstMec As ADODB.Recordset, RST1 As ADODB.Recordset, MyRst As ADODB.Recordset
Set MyRst = New ADODB.Recordset
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    'If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    
    Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    ' If Check1(1).Value = Unchecked Then Condstr = Condstr & " and JL.Mech_Code in (" & GridString2 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and " & cMID("jl.job_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(2).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("jl.job_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
   ' If Check1(2).Value = Unchecked Then Condstr = Condstr & " and JL.Mech_Code in (" & GridString2 & ")"
    
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Total
'  1           2       3       4             5       6         7           8          9           10             11              12
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Emp_Code", adChar, 4, adFldIsNullable
        .Fields.Append "Emp_Name", adVarChar, 40, adFldIsNullable
        .Fields.Append "JobClosDate", adDate, 7, adFldIsNullable
        .Fields.Append "Job_DocID", adVarChar, 21, adFldIsNullable
        .Fields.Append "Serv_Type", adVarChar, 2, adFldIsNullable
        .Fields.Append "Serv_Detail", adVarChar, 20, adFldIsNullable
        .Fields.Append "Model", adVarChar, 15, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 15, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 14, adFldIsNullable
        .Fields.Append "Lab_Paid", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Warr", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_OthSelf", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Mfg", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Contract", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Disc", adDouble, 12, adFldIsNullable
        .Fields.Append "Net_Amt", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Free", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Discount Total
'  1           2       3       4             5       6         7           8          9           10             11              12      12

On Error Resume Next
'Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour
'Nra Update
    GSQL = "SELECT JL.Job_DocId,JL.Site_Code,JL.Chrg_Type,JL.Chrg_From," & _
        " ((sum(JL.labourAmt)+JC.Lab_TaxAmt)-JC.Lab_D_Amt) as labourAmt,JC.Netlab_Amt,JL.major_yn,JL.external_yn,JL.Chrg_From,JL.Chrg_Type," & _
        " JL.ExtJobGatePassNo,JC.JobCloseDate,JC.Lab_D_Amt" & _
        " From ((Job_Card as JC Left Join Job_Lab as JL on JC.DocID=JL.Job_DocID) left join Job_GatePass as GP on JL.ExtJobGatePassNo=GP.GatePassNo)" & Condstr & _
        " Group By JL.Job_DocId,JL.Site_Code,JL.major_yn,JL.external_yn,JL.Chrg_From,JL.Chrg_Type," & _
        " JL.ExtJobGatePassNo,JC.JobCloseDate,JC.Lab_D_Amt,JL.Chrg_Type,JL.Chrg_From,JC.Lab_TaxAmt,JC.Netlab_Amt Order by JC.JobCloseDate"
           'JL.Job_DocId=GP.Job_DocId
        
    MyRst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    Dim myDate As String
    myDate = IIf(IsNull(MyRst!JobCloseDate), "", MyRst!JobCloseDate)
    RstRep.AddNew
    RstRep!Lab_warr = IIf(IsNull(RstRep!Lab_warr), 0, RstRep!Lab_warr)
    RstRep!Lab_mfg = IIf(IsNull(RstRep!Lab_mfg), 0, RstRep!Lab_mfg)
    RstRep!Lab_OthSelf = IIf(IsNull(RstRep!Lab_OthSelf), 0, RstRep!Lab_OthSelf)
    RstRep!Lab_Paid = IIf(IsNull(RstRep!Lab_Paid), 0, RstRep!Lab_Paid)
    RstRep!Lab_Contract = IIf(IsNull(RstRep!Lab_Contract), 0, RstRep!Lab_Contract)
    RstRep!JobClosDate = MyRst!JobCloseDate
    RstRep!Lab_Free = IIf(IsNull(RstRep!Lab_Free), 0, RstRep!Lab_Free)
    'RstRep!JobClosDate = IIf(IsNull(RstRep!JobClosDate), "", RstRep!JobClosDate)
    While Not MyRst.EOF = True
        If myDate <> IIf(IsNull(MyRst!JobCloseDate), "", MyRst!JobCloseDate) Then
            RstRep.AddNew
            RstRep!Lab_warr = IIf(IsNull(RstRep!Lab_warr), 0, RstRep!Lab_warr)
            RstRep!Lab_mfg = IIf(IsNull(RstRep!Lab_mfg), 0, RstRep!Lab_mfg)
            RstRep!Lab_OthSelf = IIf(IsNull(RstRep!Lab_OthSelf), 0, RstRep!Lab_OthSelf)
            RstRep!Lab_Paid = IIf(IsNull(RstRep!Lab_Paid), 0, RstRep!Lab_Paid)
            RstRep!Lab_Contract = IIf(IsNull(RstRep!Lab_Contract), 0, RstRep!Lab_Contract)
     '      RstRep!JobClosDate = IIf(IsNull(RstRep!JobClosDate), "", RstRep!JobClosDate)
            RstRep!JobClosDate = MyRst!JobCloseDate
            myDate = RstRep!JobClosDate
        End If
            If MyRst!Chrg_Type = "W" Then
                RstRep!Lab_warr = RstRep!Lab_warr + MyRst!LabourAmt
            End If
            If MyRst!Chrg_Type = "S" Or MyRst!Chrg_Type = "O" Then
                RstRep!Lab_OthSelf = RstRep!Lab_OthSelf + MyRst!LabourAmt
            End If
            If MyRst!Chrg_From = "C" And MyRst!Chrg_Type = "C" Then
                RstRep!Lab_Paid = RstRep!Lab_Paid + MyRst!NetLab_Amt
            End If
            If MyRst!Chrg_Type = "F" Then
                RstRep!Lab_Free = RstRep!Lab_Free + MyRst!LabourAmt
            End If
            If MyRst!External_yn = 1 Then
                RstRep!Lab_Contract = RstRep!Lab_Contract + MyRst!LabourAmt
            End If
    MyRst.MoveNext
    Wend
    If MyRst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "LabRevnueRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub VehGrdRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$, Grade$, PDI As Integer, Free As Integer, Warranty As Integer, Paid As Integer, Denting As Integer, Other As Integer, I As Integer, j As Integer
Dim MyRst As ADODB.Recordset
Dim rstmodel As ADODB.Recordset
Set MyRst = New ADODB.Recordset
Set rstmodel = New ADODB.Recordset
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    
    Condstr = " where H.CardDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and H.CardDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and H.Div_Code='" & PubDivCode & "'"
'      If Check1(1).Value = Unchecked Then Condstr = Condstr & " and JL.Mech_Code in (" & GridString1 & ")"
   
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(h.site_code,1) in (" & GridString2 & ")"
    If Check1(2).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(h.site_code,1) ='" & PubSiteCode & "' "
    End If
   
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Total
'  1           2       3       4             5       6         7           8          9           10             11              12
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Grade", adChar, 1, adFldIsNullable
        .Fields.Append "Owner_Name", adVarChar, 40, adFldIsNullable
        .Fields.Append "Model", adVarChar, 15, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 15, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 14, adFldIsNullable
        .Fields.Append "PDI", adDouble, 3, adFldIsNullable
        .Fields.Append "Free", adDouble, 3, adFldIsNullable
        .Fields.Append "Warranty", adDouble, 3, adFldIsNullable
        .Fields.Append "Paid", adDouble, 3, adFldIsNullable
        .Fields.Append "Denting", adDouble, 3, adFldIsNullable
        .Fields.Append "Other", adDouble, 3, adFldIsNullable
        .Fields.Append "Last_ServiceType", adVarChar, 35, adFldIsNullable
        .Fields.Append "Last_ServiceDate", adVarChar, 17, adFldIsNullable
        .Fields.Append "Days", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Discount Total
'  1           2       3       4             5       6         7           8          9           10             11              12      12

'Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour
'Nra Update
    GSQL = "SELECT H.CardDate, H.Model, H.RegNo, H.Chassis, H.Name,JC.CardNo, JC.Serv_Type," & _
           "ST.Serv_Desc,ST.Serv_Catg " & _
           "FROM ((HisCard AS H " & _
           "LEFT JOIN Job_Card AS JC ON H.CardNo=JC.CardNo)" & _
           "LEFT JOIN Service_Type ST ON ST.Serv_Type=JC.Serv_Type)" & Condstr & _
           " group by H.Model, H.CardNo, H.CardDate, H.RegNo, H.Chassis, H.Name,JC.Serv_Type,ST.Serv_Desc,ST.Serv_Catg,JC.CardNo"

    MyRst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
'Total Serviced Model
    GSQL = "Select distinct H.Model,CardNo,Chassis,RegNo from HisCard H" & Condstr
    rstmodel.Open GSQL, GCn, adOpenDynamic, adLockReadOnly
    For I = 1 To rstmodel.RecordCount
    If MyRst.RecordCount <= 0 Then: Exit For
        MyRst.MoveFirst
        For j = 1 To MyRst.RecordCount
        If rstmodel!CardNo = MyRst!CardNo Then
           Select Case MyRst!serv_catg
                Case "P"
                    PDI = PDI + 1
                Case "F"
                    If MyRst!Serv_Type = "WW" Then
                        Warranty = Warranty + 1
                    Else
                        Free = Free + 1
                    End If
                Case "C"
                    Paid = Paid + 1
                Case "D"
                    Denting = Denting + 1
                Case Else
                    Other = Other + 1
           End Select
        End If
        MyRst.MoveNext
        Next
        
'Grade Calculation
        If Free > 0 Then
            If Paid + Denting + Other > 0 Then
                Grade = "A"
            Else
                Grade = "C"
            End If
        Else
            If Paid + Denting > 0 Then
                Grade = "B"
            Else
                If Warranty > 0 Then
                    Grade = "D"
                Else
                    Grade = "E"
                End If
            End If
        End If
    
  ' Record addition
  On Error Resume Next
  With RstRep
        .AddNew
        !Grade = Grade
        !Model = rstmodel!Model
        !Chassis = rstmodel!Chassis
        !RegNo = rstmodel!RegNo
        !PDI = PDI
        !Free = Free
        !Warranty = Warranty
        !Paid = Paid
        !Denting = Denting
        !Other = Other
        GSQL = "Select max(JC.Job_Date) from Job_Card JC Where JC.CardNo='" & rstmodel!CardNo & "'"
        !Last_ServiceDate = GCn.Execute(GSQL).Fields(0).Value
         GSQL = "SELECT JC.serv_Type FROM Job_Card AS JC where JC.CardNo='" & rstmodel!CardNo & "' and JC.Job_Date=" & ConvertDate(!Last_ServiceDate) & ""
        !Last_ServiceType = IIf(IsNull(GCn.Execute(GSQL).Fields(0).Value), "", GCn.Execute(GSQL).Fields(0).Value)
        !DAYS = DateDiff("D", !Last_ServiceDate, date)
  End With
        PDI = 0: Paid = 0: Warranty = 0: Denting = 0: Other = 0: Free = 0
    rstmodel.MoveNext
Next
       

    If MyRst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "VehGrdRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub MechEarnSumProc()
On Error GoTo ELoop
' Nra Updation
Dim mQry$, Condstr$, CondStr1$, MySql1$, MySql$, NoVeh, NormalLabour As Long, WarLabour As Long, FreeLabour As Long, ContLabour As Long, FreeServLabour As Long, TotalLabour As Integer, I As Integer
Dim Rst As ADODB.Recordset, rstMec As ADODB.Recordset, RST1 As ADODB.Recordset, MyRst As ADODB.Recordset, RstMech As ADODB.Recordset
Set MyRst = New ADODB.Recordset
Set RstMech = New ADODB.Recordset
Set Rst = New ADODB.Recordset
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then CondStr1 = CondStr1 & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    If Check1(2).Value = Unchecked Then CondStr1 = CondStr1 & "  where JL2.Mech_Code in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then CondStr1 = CondStr1 & " where left(Job_Card.DocId,1) in (" & GridString3 & ")"
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Total
'  1           2       3       4             5       6         7           8          9           10             11              12
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Emp_Code", adChar, 4, adFldIsNullable
        .Fields.Append "Emp_Name", adVarChar, 40, adFldIsNullable
        .Fields.Append "JobClosDate", adVarChar, 17, adFldIsNullable
        .Fields.Append "Job_DocID", adVarChar, 21, adFldIsNullable
        .Fields.Append "Serv_Type", adVarChar, 2, adFldIsNullable
        .Fields.Append "Serv_Detail", adVarChar, 20, adFldIsNullable
        .Fields.Append "Model", adVarChar, 15, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 15, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 14, adFldIsNullable
        .Fields.Append "Lab_Paid", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Warr", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_OthSelf", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Mfg", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Contract", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Disc", adDouble, 12, adFldIsNullable
        .Fields.Append "Net_Amt", adDouble, 12, adFldIsNullable
        .Fields.Append "S_No", adDouble, 4, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set RstRep1 = New ADODB.Recordset
    With RstRep1
        .Fields.Append "Emp_Code", adChar, 7, adFldIsNullable
        .Fields.Append "Emp_Name", adVarChar, 40, adFldIsNullable
        .Fields.Append "JobClosDate", adVarChar, 17, adFldIsNullable
        .Fields.Append "Job_DocID", adVarChar, 21, adFldIsNullable
        .Fields.Append "Serv_Type", adVarChar, 2, adFldIsNullable
        .Fields.Append "Serv_Detail", adVarChar, 20, adFldIsNullable
        .Fields.Append "Model", adVarChar, 15, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 15, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 14, adFldIsNullable
        .Fields.Append "Lab_Paid", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Warr", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_OthSelf", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Mfg", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Contract", adDouble, 12, adFldIsNullable
        .Fields.Append "Lab_Disc", adDouble, 12, adFldIsNullable
        .Fields.Append "Net_Amt", adDouble, 12, adFldIsNullable
        .Fields.Append "S_No", adDouble, 4, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
'Close Date  Job No.  Model  Chassis No.  Reg No.  Service  Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour Discount Total
'  1           2       3       4             5       6         7           8          9           10             11              12      12

'No of Mechanic
        
'    mQRY = "Select DocID,JobClosDate,Serv_Type,Model,Chassis,RegNo,Emp_Code,Emp_Name," & _
'        "from ((Job_Card as J left Join Hiscard as H on J.CardNo=H.CardNo))" & _
'        "left Join Job_Lab2 as JL2 on J.DocID=JL2.Job_DocID " & _
'        "left Join Emp_Mast as E on JL2.Mech_Code=E.Emp_Code "


'Paid Labour  Warranty  Oth Dlr+Self  Free Service  Contract Labour
On Error Resume Next
'TOTAL LABOUR DONE IN GIVEN CONDITION
    GSQL = "SELECT JL.Job_DocId,JL.s_No,JL.hrs_taken,JL.lab_rate," & _
        " JL.hrs_war,JL.war_lab_rate,JL.labourAmt,JL.Chrg_From,JL.Chrg_Type," & _
        " JL.ExtJobGatePassNo,JL.Job_DocID,JC.JobCloseDate" & _
        " From ((Job_Lab as JL left join Job_GatePass as GP on JL.ExtJobGatePassNo=GP.GatePassNo)left join Job_Card as JC on JL.Job_DocID=JC.DocID)" & Condstr & _
        " Order by JL.Job_DocID,JL.s_No,JC.JobCloseDate desc"
        ' JL.Job_DocId=GP.Job_DocId
        
    MyRst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    MyRst.MoveFirst
    Dim MyID As String
    Dim S_No As Integer
    MyID = IIf(IsNull(MyRst!job_docid), "", MyRst!job_docid)
    S_No = IIf(IsNull(MyRst!S_No), "", MyRst!S_No)
    RstRep.AddNew
    RstRep!Lab_warr = IIf(IsNull(RstRep!Lab_warr), 0, RstRep!Lab_warr)
    RstRep!Lab_mfg = IIf(IsNull(RstRep!Lab_mfg), 0, RstRep!Lab_mfg)
    RstRep!Lab_OthSelf = IIf(IsNull(RstRep!Lab_OthSelf), 0, RstRep!Lab_OthSelf)
    RstRep!Lab_Paid = IIf(IsNull(RstRep!Lab_Paid), 0, RstRep!Lab_Paid)
    RstRep!job_docid = MyRst!job_docid
    RstRep!JobClosDate = IIf(IsNull(MyRst!JobCloseDate), "", MyRst!JobCloseDate)
'TOTAL NO OF MECHANIC ON A PERTICULAR JOB AND SR NO.
    GSQL = "SELECT JL2.Job_DocID,JL2.Lab_Code as SearchCode,JL2.s_No,count(Mech_Code) as NoOfMech" & _
        " from ((Job_Lab2 as JL2 left join Job_Card as J on JL2.Job_DocID=J.Docid) " & _
        " LEFT JOIN Emp_Mast on Emp_Mast.Emp_Code=JL2.Mech_Code)" & CondStr1 & _
        " Group By JL2.Job_DocID,JL2.Lab_Code,JL2.s_No"

    RstMech.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    Dim onePart_Warr As Integer
    Dim onePart_mfg As Integer
    Dim onePart_OthSelf As Integer
    Dim onePart_LabPaid As Integer
    
    Dim lastPart_Warr As Integer
    Dim lastPart_mfg As Integer
    Dim lastPart_OthSelf As Integer
    Dim lastPart_LabPaid As Integer, Cnt As Integer
    
    For Cnt = 0 To MyRst.RecordCount
       If MyID <> IIf(IsNull(MyRst!job_docid), "", MyRst!job_docid) Or S_No <> IIf(IsNull(MyRst!S_No), "", MyRst!S_No) Then
           RstMech.MoveFirst
          While Not RstMech.EOF = True
             If MyID = RstMech!job_docid And S_No = RstMech!S_No Then
'EMPLOYEE NAME REGARDING PERTICULAR MECH. ID.
               GSQL = "SELECT JL2.Job_DocID,JL2.s_No,Emp_Mast.Emp_Name" & _
                      " from (Job_Lab2 as JL2 LEFT JOIN Emp_Mast on Emp_Mast.Emp_Code=JL2.Mech_Code)" & _
                      " where  JL2.Job_DocID= '" & RstMech!job_docid & "' and JL2.s_No=" & RstMech!S_No & ""
               
               Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
           ' Mechanic wise labour devidation
                    If RstMech!NoOfMech > 1 Then
                       onePart_Warr = IIf(Val(RstRep!Lab_warr) = 0, 0, Val(RstRep!Lab_warr) / RstMech!NoOfMech)
                       onePart_mfg = IIf(Val(RstRep!Lab_mfg) = 0, 0, Val(RstRep!Lab_mfg) / RstMech!NoOfMech)
                       onePart_OthSelf = IIf(Val(RstRep!Lab_OthSelf) = 0, 0, Val(RstRep!Lab_OthSelf) / RstMech!NoOfMech)
                       onePart_LabPaid = IIf(Val(RstRep!Lab_Paid) = 0, 0, Val(RstRep!Lab_Paid) / RstMech!NoOfMech)
                       
                       lastPart_Warr = IIf(Val(RstRep!Lab_warr) = 0, 0, Val(RstRep!Lab_warr) - (Val(onePart_Warr) * (RstMech!NoOfMech - 1)))
                       lastPart_mfg = IIf(Val(RstRep!Lab_mfg) = 0, 0, Val(RstRep!Lab_mfg) - (Val(onePart_mfg) * (RstMech!NoOfMech - 1)))
                       lastPart_OthSelf = IIf(Val(RstRep!Lab_OthSelf) = 0, 0, Val(RstRep!Lab_OthSelf) - (Val(onePart_OthSelf) * (RstMech!NoOfMech - 1)))
                       lastPart_LabPaid = IIf(Val(RstRep!Lab_Paid) = 0, 0, Val(RstRep!Lab_Paid) - (Val(onePart_LabPaid) * (RstMech!NoOfMech - 1)))
                       
                       For I = 1 To RstMech!NoOfMech
                            If I <= RstMech!NoOfMech - 1 Then
                                RstRep1.AddNew
                                RstRep1!Emp_Name = Rst!Emp_Name
                                RstRep1!JobClosDate = RstRep!JobClosDate
                                RstRep1!Lab_warr = onePart_Warr
                                RstRep1!Lab_mfg = onePart_mfg
                                RstRep1!Lab_OthSelf = onePart_OthSelf
                                RstRep1!Lab_Paid = onePart_LabPaid
                                RstRep!Lab_warr = 0
                                RstRep!Lab_mfg = 0
                                RstRep!Lab_OthSelf = 0
                                RstRep!Lab_Paid = 0
                            Else
                                RstRep1.AddNew
                                RstRep1!Emp_Name = Rst!Emp_Name
                                RstRep1!JobClosDate = RstRep!JobClosDate
                                RstRep1!Lab_warr = lastPart_Warr
                                RstRep1!Lab_mfg = lastPart_mfg
                                RstRep1!Lab_OthSelf = lastPart_OthSelf
                                RstRep1!Lab_Paid = lastPart_LabPaid
                                RstRep!Lab_warr = 0
                                RstRep!Lab_mfg = 0
                                RstRep!Lab_OthSelf = 0
                                RstRep!Lab_Paid = 0
                            End If
                            Rst.MoveNext
                       Next
                       Else
                                RstRep1.AddNew
                                RstRep1!Emp_Name = Rst!Emp_Name
                                RstRep1!JobClosDate = RstRep!JobClosDate
                                RstRep1!Lab_warr = RstRep!Lab_warr
                                RstRep1!Lab_mfg = RstRep!Lab_mfg
                                RstRep1!Lab_OthSelf = RstRep!Lab_OthSelf
                                RstRep1!Lab_Paid = RstRep!Lab_Paid
                                RstRep!Lab_warr = 0
                                RstRep!Lab_mfg = 0
                                RstRep!Lab_OthSelf = 0
                                RstRep!Lab_Paid = 0
                                Rst.MoveNext
                    End If
                End If
                RstMech.MoveNext
                If Rst.State = 1 Then: Rst.Close
            Wend
                       
           ' End devidation
           
        
            RstRep!Lab_warr = IIf(IsNull(RstRep!Lab_warr), 0, RstRep!Lab_warr)
            RstRep!Lab_mfg = IIf(IsNull(RstRep!Lab_mfg), 0, RstRep!Lab_mfg)
            RstRep!Lab_OthSelf = IIf(IsNull(RstRep!Lab_OthSelf), 0, RstRep!Lab_OthSelf)
            RstRep!Lab_Paid = IIf(IsNull(RstRep!Lab_Paid), 0, RstRep!Lab_Paid)
            RstRep!JobClosDate = IIf(IsNull(MyRst!JobCloseDate), "", MyRst!JobCloseDate)
            RstRep!job_docid = MyRst!job_docid
            RstRep!S_No = MyRst!S_No
            MyID = RstRep!job_docid
            S_No = RstRep!S_No
        End If
        ' Catagory wise Labour computing
        
     Select Case MyRst!Chrg_From
        Case "M"
            If MyRst!Chrg_Type = "W" Then
                RstRep!Lab_warr = RstRep!Lab_warr + MyRst!LabourAmt
            Else
                RstRep!Lab_mfg = RstRep!Lab_mfg + MyRst!LabourAmt
            End If
        Case "S", "O"
                RstRep!Lab_OthSelf = RstRep!Lab_OthSelf + MyRst!LabourAmt
        Case "C"
            RstRep!Lab_Paid = RstRep!Lab_Paid + MyRst!LabourAmt
    End Select
    MyRst.MoveNext
Next
    If MyRst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "MechEarnSumm"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub DeWiseVehAtProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where jc.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and jc.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("jc.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("jc.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and H.Dealer_Code in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(jc.DocId,1) in (" & GridString3 & ")"
    mQry = " select jc.Serv_Type,count(JC.Serv_Type) as NoVeh,H.Dealer_Code,H.Model,AMD.D_NAME " & _
           " from ((Job_Card JC left join Service_Type ST on JC.Serv_Type=ST.Serv_Type) " & _
           " left join Hiscard H on JC.CardNo=H.CardNo) " & _
           " Left join Amd_Dealer AMD on H.Dealer_Code=AMD.D_Code "
    
    mQry = mQry + Condstr + " Group by H.Dealer_Code,H.Model,jc.Serv_Type,AMD.D_NAME "
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "Dlr-wiseVeh"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub DeWiseJobAnaProc()
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$
Dim I, j As Integer
Dim Rst As ADODB.Recordset
Dim RST1 As ADODB.Recordset
Set RstRep = New ADODB.Recordset
Dim rstEarn As ADODB.Recordset
Dim MySql, MySql1 As String, FillFirstRec As Boolean
On Error GoTo ELoop
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Job_Card.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.Job_Date <= " & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and H.Dealer_Code in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(Job_Card.DocId,1) in (" & GridString3 & ")"
   'Dealer Code  Dealer Name   Model      Service Types 1-15 Total
   'adVarChar    adVarChar     adVarchar  adInteger          adInteger
   'Header = 19
   'Values = 19
   
   With RstRep
        .Fields.Append "Dealer", adVarChar, 10, adFldIsNullable
        .Fields.Append "Chassiss-No", adVarChar, 40, adFldIsNullable
        .Fields.Append "Kms", adVarChar, 15, adFldIsNullable
        .Fields.Append "Job-No", adVarChar, 10, adFldIsNullable
        .Fields.Append "Job-Op-Dt", adVarChar, 10, adFldIsNullable
        .Fields.Append "Job-Clo-Dt", adVarChar, 10, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 10, adFldIsNullable
        
        .Fields.Append "Head1", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head2", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head3", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head4", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head5", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head6", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head7", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head8", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head9", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head10", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head11", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head12", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head13", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head14", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head15", adVarChar, 10, adFldIsNullable
        .Fields.Append "Head16", adVarChar, 10, adFldIsNullable  'TOTAL
        
        
        .Fields.Append "Cat1", adVarChar, 40, adFldIsNullable   '1
        .Fields.Append "Cat2", adVarChar, 15, adFldIsNullable   '2
        .Fields.Append "Cat3", adDouble, 10, adFldIsNullable    '3
        .Fields.Append "Cat4", adDouble, 8, adFldIsNullable     '4
        .Fields.Append "Cat5", adDate, 7, adFldIsNullable       '5
        .Fields.Append "Cat6", adDate, 7, adFldIsNullable       '6
        .Fields.Append "Cat7", adVarChar, 14, adFldIsNullable   '7
        
        .Fields.Append "Cat8", adDouble, 6, adFldIsNullable     '8
        .Fields.Append "Cat9", adDouble, 6, adFldIsNullable     '9
        .Fields.Append "Cat10", adDouble, 6, adFldIsNullable    '10
        .Fields.Append "Cat11", adDouble, 6, adFldIsNullable    '11
        .Fields.Append "Cat12", adDouble, 6, adFldIsNullable    '12
        .Fields.Append "Cat13", adDouble, 6, adFldIsNullable    '13
        .Fields.Append "Cat14", adDouble, 6, adFldIsNullable    '14
        .Fields.Append "Cat15", adDouble, 6, adFldIsNullable    '15
        .Fields.Append "Cat16", adDouble, 6, adFldIsNullable    '16
        .Fields.Append "Cat17", adDouble, 6, adFldIsNullable    '17
        .Fields.Append "Cat18", adDouble, 6, adFldIsNullable    '18
        .Fields.Append "Cat119", adDouble, 6, adFldIsNullable   '19
        .Fields.Append "Cat20", adDouble, 6, adFldIsNullable    '20
        .Fields.Append "Cat21", adDouble, 6, adFldIsNullable    '21
        .Fields.Append "Cat22", adDouble, 6, adFldIsNullable    '22
        .Fields.Append "Cat23", adDouble, 6, adFldIsNullable    '23
        
        
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
     End With
    
        MySql = "select serv_type from service_type order by Serv_Catg desc,FreeServCode"
        
        Set Rst = GCn.Execute(MySql)
        'Insert First Header Record
        With RstRep
            .AddNew
            .Fields(0).Value = "Dealer" 'rst1!Dealer_Code
            .Fields(1).Value = "Chassiss No" 'rst1!Emp_Name
            .Fields(2).Value = "Kms" 'rst1!Model
            .Fields(3).Value = "Job-No" 'rst1!Dealer_Code
            .Fields(4).Value = "Job-Op-Dt" 'rst1!Emp_Name
            .Fields(5).Value = "Job-Clo-Dt" 'rst1!Model
            .Fields(6).Value = "RegNo"
            Rst.MoveFirst
            I = 7
            While I <= 22 And Not Rst.EOF
                    RstRep.Fields(I).Value = Rst!Serv_Type
                I = I + 1
                Rst.MoveNext
            Wend
'            .Fields(i).Value = "Job-No"
'            .Fields(i + 1).Value = "Job-Op-Dt"
'            .Fields(i + 2).Value = "Job-Clo-Dt"
'            .Fields(i + 3).Value = "Owner"
            .Update
        End With
                
        MySql1 = "select count(JC.Serv_Type) as NoVeh, jc.Serv_Type,jc.AtKMsHrs As Kms,Jc.Job_No AS JobNo,H.RegNo AS VehRegNo,H.Chassis AS ChassisNo,Jc.Job_Date AS JobDate,jc.JobCloseDate AS JobCloDate,H.Name AS Owner,H.Model,H.Dealer_Code, " & xIsNull("AMD.D_NAME", "") & " as AmdName " & _
                " from ((Job_Card JC left join Service_Type ST on JC.Serv_Type=ST.Serv_Type) " & _
                " left join Hiscard H on JC.CardNo=H.CardNo) " & _
                " Left join Amd_Dealer AMD on H.Dealer_Code=AMD.D_Code " & _
                " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  " & Condstr2 & "" & _
                " group by jc.Serv_Type, H.Dealer_Code,H.Model,H.Dealer_Code,AMD.D_NAME,jc.AtKMsHrs,Jc.Job_No,H.Chassis ,Jc.Job_Date,jc.JobCloseDate,h.name,H.RegNo"
Dim ModelLoop As Boolean, DlrLoop As Boolean
        Set RST1 = GCn.Execute(MySql1)
        mDlrName = ""
        mModel = ""
        Do Until RST1.EOF = True
            mDlrName = RST1!AmdName
            DlrLoop = True
            Do While DlrLoop ' Rst1!AmDName = mDlrName        'Dealer
            
'                mModel = rst1!Model
'                ModelLoop = True
'                Do While ModelLoop 'Rst1!Model = mModel       'Model
                    With RstRep
                        
                        If FillFirstRec = True Then
                           RstRep.AddNew
                        End If
                        .Fields(23).Value = RST1!AmdName
                        .Fields(24).Value = RST1!ChassisNo
                        .Fields(25).Value = RST1!Kms
                        .Fields(26).Value = RST1!JobNo
                        .Fields(27).Value = RST1!JobDate
                        .Fields(28).Value = RST1!JobCloDate
                        .Fields(29).Value = RST1!VehRegNo
                        FillFirstRec = False
                        
                        j = CatFldName2(Rst, RST1!Serv_Type)
                        RstRep.Fields(j).Value = RST1!NoVeh
'                        RstRep.Fields(23 + i).Value = rst1!JobNo
'                        RstRep.Fields(24 + i).Value = rst1!JobDate
'                        RstRep.Fields(25 + i).Value = rst1!JobCloDate
'                        RstRep.Fields(26 + i).Value = rst1!Owner
                    End With
                    RST1.MoveNext
'                    If rst1.EOF Then
'                        ModelLoop = False
'                    ElseIf rst1!Model <> mModel Then
'                        ModelLoop = False
'                    End If
'                Loop

                If RST1.EOF Then
                    DlrLoop = False
                ElseIf RST1!AmdName <> mDlrName Then
                    DlrLoop = False
                End If
                FillFirstRec = True
            Loop

        Loop
        
'    mQRY = mQRY + CondStr

    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'    Set RstRep = New Recordset
'    RstRep.CursorLocation = adUseClient
'    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    RepName = "Dlr-wiseJob"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ModWiseJobProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where jc.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and jc.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("jc.DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("jc.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(jc.DocId,1) in (" & GridString3 & ")"
   
    mQry = "select count(JC.Serv_Type) as NoVeh, jc.Serv_Type,H.Model " & _
            " from Job_Card JC left join Hiscard H on JC.CardNo=H.CardNo "
    mQry = mQry + Condstr + " group by jc.Serv_Type, H.Model "
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "Model-wiseJob"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ModWiseSrvTaxProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
   
    
    Condstr = " where jc.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and jc.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and (Lab_TaxPer <> 0 or Lab_TaxAmt<>0) "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and JC.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and jc.Site_Code ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and H.Div_Code in (" & GridString2 & ")"
   
    mQry = "select JC.DocID,JC.JobCloseDate,JC.DocId_InvLab,JC.Lab_TaxAmt,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_D_Amt," & cIIF("JC.LabAmt_TB <> 0", "JC.LabAmt_Tb*JC.Lab_D_Amt/(JC.LabAmt_TB+JC.LabAmt_TP)", "0") & " as Lab_DiscAmtTB ,H.Model, JC.ServiceTaxAmt_Saperate, JC.ECessAmt, JC.HECessAmt " & _
            " from Job_Card JC left join Hiscard H on JC.CardNo=H.CardNo "
    mQry = mQry & Condstr & " Order By JC.JobCloseDate,JC.DocID_InvLab"
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "ServiceTaxReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub MoWiseSprInvProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset
            
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Model", adVarChar, 25, adFldIsNullable
        .Fields.Append "Prev", adInteger, 4, adFldIsNullable
        .Fields.Append "New", adInteger, 4, adFldIsNullable
        .Fields.Append "Close", adInteger, 4, adFldIsNullable
        .Fields.Append "TotalJobClose", adInteger, 4, adFldIsNullable
        .Fields.Append "Tgt", adInteger, 4, adFldIsNullable
        .Fields.Append "LabCahrged", adDouble, 12, adFldIsNullable
        .Fields.Append "LabPaid", adDouble, 12, adFldIsNullable
        .Fields.Append "SprCahrged", adDouble, 12, adFldIsNullable
        .Fields.Append "SprFree", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set Rst = GCn.Execute("SELECT distinct HisCard.Model FROM Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo where Job_Card.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And Job_Card.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  And left(Job_Card.DocId,1) in ('" & PubDivCode & "') order by HisCard.model")
    If Rst.RecordCount > 0 Then
    Do Until Rst.EOF
     
    With RstRep
       .AddNew
       .Fields("Model") = Rst!Model
       '.Fields("Tgt") = IIf(IsNull(Rst!serv_target), 0, Rst!serv_target)
       Set RST1 = GCn.Execute("select count(jc.DocId) as PrevJob from job_card Jc LEFT JOIN HisCard ON jc.CardNo = HisCard.CardNo where  jc.JobCloseDate Is Null And jc.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       .Fields("Prev") = IIf(RST1.RecordCount = 0, 0, RST1!PrevJob)
       Set RST1 = GCn.Execute("select count(jc.DocId) as NewJob from job_card Jc LEFT JOIN HisCard ON jc.CardNo = HisCard.CardNo where  job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       .Fields("New") = IIf(RST1.RecordCount = 0, 0, RST1!NewJob)
       Set RST1 = GCn.Execute("select count(jc.DocId) as CloseJob from job_card Jc LEFT JOIN HisCard ON jc.CardNo = HisCard.CardNo where  Not JobCloseDate Is Null And JobCloseDate  >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And JobCloseDate<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       .Fields("Close") = IIf(RST1.RecordCount = 0, 0, RST1!CloseJob)
       Set RST1 = GCn.Execute("select count(jc.DocId) as TotalJobClose from job_card Jc LEFT JOIN HisCard ON jc.CardNo = HisCard.CardNo where  Not JobCloseDate is null And JobCloseDate  >=" & ConvertDate(Format(PubStartDate, "dd/MMM/yyyy")) & " And JobCloseDate<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       .Fields("TotalJobClose") = IIf(RST1.RecordCount = 0, 0, RST1!TotalJobClose)
       Set RST1 = GCn.Execute("select sum(jc.NetLab_Amt) as LabCharged from job_card Jc LEFT JOIN HisCard ON jc.CardNo = HisCard.CardNo where job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       .Fields("LabCahrged") = IIf(RST1.RecordCount = 0, 0, RST1!LabCharged)
'       Set Rst1 = GCn.Execute("SELECT sum(Job_Lab.ContractAmt) as LabPaid FROM (Job_Card LEFT JOIN Job_Lab ON Job_Card.DocId = Job_Lab.Job_DocID) LEFT JOIN HisCard ON job_card.CardNo = HisCard.CardNo where Not IsNull(Job_Card.JobCloseDate) and Job_Card.job_date >=#" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# And Job_Card.job_date<=#" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#  and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       Set RST1 = GCn.Execute("SELECT sum(JG.ContractAmt) as LabPaid FROM ((Job_GatePass as JG LEFT JOIN Job_Card as JC on JG.Job_DocID=JC.DocID) Left Join Hiscard as H on JC.CardNo=H.Cardno) Where Left(JC.DocId,1)='" & PubDivCode & "' And " & cMID("JC.DocId", "3", "1") & "='" & PubSiteCode & "' And JC.JobCloseDate Is  Not Null and JC.JobCloseDate Is Not Null and JC.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And JC.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and h.model = '" & Rst!Model & "' group by h.model")
       .Fields("LabPaid") = IIf(RST1.RecordCount = 0, 0, RST1!LabPaid)
       Set RST1 = GCn.Execute("SELECT sum(SP_Stock.Net_Amt) as NetAmt FROM (Job_Card LEFT JOIN SP_Stock ON Job_Card.DocId = SP_Stock.Job_DocID) LEFT JOIN HisCard ON job_card.CardNo = HisCard.CardNo where Job_Card.JobCloseDate Is Not Null and Job_Card.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And Job_Card.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and SP_Stock.Purpose = 'C' and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       .Fields("SprCahrged") = IIf(RST1.RecordCount = 0, 0, RST1!NetAmt)
       Set RST1 = GCn.Execute("SELECT sum(SP_Stock.Net_Amt) as NetAmtFree FROM (Job_Card LEFT JOIN SP_Stock ON Job_Card.DocId = SP_Stock.Job_DocID) LEFT JOIN HisCard ON job_card.CardNo = HisCard.CardNo where Job_Card.JobCloseDate Is Not Null and Job_Card.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And Job_Card.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and  SP_Stock.Purpose <> 'C' and hiscard.model = '" & Rst!Model & "' group by hiscard.model")
       .Fields("SprFree") = IIf(RST1.RecordCount = 0, 0, RST1!NetAmtFree)
       .Update
    End With
    Rst.MoveNext
    Loop
    End If
    Set Rst = Nothing
    Set RST1 = Nothing
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "MoWiseSprInv"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub JobWiseLabAProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim mPurpose As String

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    Condstr = " where Job_Card.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
   
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Job_Lab.Lab_Code in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(Job_Card.DocId,1) in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Job_Lab2.Mech_Code in (" & GridString4 & ")"
           
    mQry = "SELECT Job_Card.DocId,Job_Card.Job_Date," & cIIF("Service_type.serv_Catg='C'", "Job_Lab.LabourAmt", "0") & " as Chargeable, " & _
           " " & cIIF("Service_type.serv_catg='P'", "Job_Lab.LabourAmt", "0") & " as PDI, " & cIIF("Service_Type.serv_Catg='F'", "Job_Lab.LabourAmt", "0") & " as Free, " & _
           "Job_Card.Job_No, Job_Card.JobCloseDate, HisCard.RegNo, HisCard.Chassis, LABOUR.Lab_Desc, " & _
           "" & cIIF("Job_Lab.Hrs_War =0", "0", "Job_Lab.War_Lab_Rate") & " as WarrLab, Job_Card.CardNo, " & _
           "Job_Card.Serv_Rate, Job_Card.Serv_Type, Job_Card.CRMemo, Job_Card.Lab_D_Amt,Job_Card.Lab_Paid, " & _
           "Job_Card.NetLab_Amt, Service_Type.Serv_Catg" & _
           " FROM ((((Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) " & _
           "LEFT JOIN Job_Lab ON Job_Card.DocId = Job_Lab.Job_DocID) " & _
           "LEFT JOIN LABOUR ON Job_Lab.Lab_Code = LABOUR.Lab_Code) " & _
           "LEFT JOIN Service_Type ON Job_Card.Serv_Type = Service_Type.Serv_Type) " & _
           "Left Join Job_Lab2 on Job_Lab.Job_DocId+Job_Lab.Lab_Code=Job_Lab2.Job_DocId+Job_Lab2.Lab_Code "

    mQry = mQry & Condstr
        
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "JobWiseLabA"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ServWiseJobProc()
'''*****************************//////////******************************************
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim mPurpose As String, Rst As ADODB.Recordset, RST1 As ADODB.Recordset

    'P- >PDI,F- >Free Service, C- >Chargable,W- >Warranty,O- >Company Vehicle,L- >Complementary
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "ServType", adVarChar, 2, adFldIsNullable
        .Fields.Append "ServDesc", adVarChar, 20, adFldIsNullable
        .Fields.Append "Prev", adInteger, 4, adFldIsNullable
        .Fields.Append "New", adInteger, 4, adFldIsNullable
        .Fields.Append "Close", adInteger, 4, adFldIsNullable
        .Fields.Append "TotalJobClose", adInteger, 4, adFldIsNullable
        .Fields.Append "Tgt", adInteger, 4, adFldIsNullable
        .Fields.Append "LabCahrged", adDouble, 12, adFldIsNullable
        .Fields.Append "LabPaid", adDouble, 12, adFldIsNullable
        .Fields.Append "SprCahrged", adDouble, 12, adFldIsNullable
        .Fields.Append "SprFree", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Set Rst = GCn.Execute("select  Serv_Target,Serv_Desc,Serv_Type,Site_Code from Service_Type order by Serv_Desc")
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            With RstRep
               .AddNew
               .Fields("ServType") = Rst!Serv_Type
               .Fields("ServDesc") = Rst!Serv_Desc
               .Fields("Tgt") = IIf(IsNull(Rst!Serv_Target), 0, Rst!Serv_Target)
               Set RST1 = GCn.Execute("select count(DocId) as PrevJob from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And  Left(DocId,1)='" & PubDivCode & "' And   isNull(JobCloseDate) And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("Prev") = IIf(RST1.RecordCount = 0, 0, RST1!PrevJob)
               Set RST1 = GCn.Execute("select count(DocId) as NewJob from job_card where  " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("New") = IIf(RST1.RecordCount = 0, 0, RST1!NewJob)
               Set RST1 = GCn.Execute("select count(DocId) as CloseJob from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And   Not IsNull(JobCloseDate) And JobCloseDate  >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And JobCloseDate<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("Close") = IIf(RST1.RecordCount = 0, 0, RST1!CloseJob)
               Set RST1 = GCn.Execute("select count(DocId) as TotalJobClose from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  Not IsNull(JobCloseDate) And JobCloseDate  >=" & ConvertDate(Format(PubStartDate, "dd/MMM/yyyy")) & " And JobCloseDate<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("TotalJobClose") = IIf(RST1.RecordCount = 0, 0, RST1!TotalJobClose)
               Set RST1 = GCn.Execute("select sum(NetLab_Amt) as LabCharged from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("LabCahrged") = IIf(RST1.RecordCount = 0, 0, RST1!LabCharged)
               'Set Rst1 = GCn.Execute("SELECT sum(Job_GatePass.ContractAmt) as LabPaid FROM Job_GatePass LEFT JOIN Job_Lab ON Job_Card.DocId = Job_Lab.Job_DocID where Mid(DocId,3,1)='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  Not IsNull(Job_Card.JobCloseDate) and Job_Card.job_date >=#" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# And Job_Card.job_date<=#" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#  and Job_Card.Serv_Type = '" & Rst!Serv_Type & "' group by Job_Card.Serv_Type")
               Set RST1 = GCn.Execute("SELECT sum(JG.ContractAmt) as LabPaid FROM (Job_GatePass as JG LEFT JOIN Job_Card as JC on JG.Job_DocID=JC.DocId) where " & cMID("JC.DocId", "3", "1") & "='" & PubSiteCode & "' And Left(JC.DocId,1)='" & PubDivCode & "' And  Not IsNull(JC.JobCloseDate) and JC.Job_Date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And JC.Job_Date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and JC.Serv_Type = '" & Rst!Serv_Type & "' group by JC.Serv_Type")
               .Fields("LabPaid") = IIf(RST1.RecordCount = 0, 0, RST1!LabPaid)
               Set RST1 = GCn.Execute("SELECT sum(SP_Stock.Net_Amt) as NetAmt FROM Job_Card LEFT JOIN SP_Stock ON Job_Card.DocId = SP_Stock.Job_DocID where " & cMID("Job_Card.DocId", "3", "1") & "='" & PubSiteCode & "' And Left(Job_Card.DocId,1)='" & PubDivCode & "' And Not IsNull(Job_Card.JobCloseDate) and Job_Card.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And Job_Card.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and Job_Card.Serv_Type = '" & Rst!Serv_Type & "' and SP_Stock.Purpose = 'C' group by Job_Card.Serv_Type")
               .Fields("SprCahrged") = IIf(RST1.RecordCount = 0, 0, RST1!NetAmt)
               Set RST1 = GCn.Execute("SELECT sum(SP_Stock.Net_Amt) as NetAmtFree FROM Job_Card LEFT JOIN SP_Stock ON Job_Card.DocId = SP_Stock.Job_DocID where " & cMID("Job_Card.DocId", "3", "1") & "='" & PubSiteCode & "' And Left(Job_Card.DocId,1)='" & PubDivCode & "' And  Not IsNull(Job_Card.JobCloseDate) and Job_Card.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And Job_Card.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and Job_Card.Serv_Type = '" & Rst!Serv_Type & "' and SP_Stock.Purpose <> 'C' group by Job_Card.Serv_Type")
               .Fields("SprFree") = IIf(RST1.RecordCount = 0, 0, RST1!NetAmtFree)
               .Update
            End With
        Rst.MoveNext
        Loop
    End If
    Set Rst = Nothing
    Set RST1 = Nothing
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "ServWiseJob"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
'****************************************************************
End Sub
Private Sub ServWiseSrvTaxProc()
'''*****************************//////////******************************************
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim mPurpose As String, Rst As ADODB.Recordset, RST1 As ADODB.Recordset

    'P- >PDI,F- >Free Service, C- >Chargable,W- >Warranty,O- >Company Vehicle,L- >Complementary
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "ServType", adVarChar, 2, adFldIsNullable
        .Fields.Append "ServDesc", adVarChar, 20, adFldIsNullable
        .Fields.Append "Prev", adInteger, 4, adFldIsNullable
        .Fields.Append "New", adInteger, 4, adFldIsNullable
        .Fields.Append "Close", adInteger, 4, adFldIsNullable
        .Fields.Append "TotalJobClose", adInteger, 4, adFldIsNullable
        .Fields.Append "Tgt", adInteger, 4, adFldIsNullable
        .Fields.Append "LabCahrged", adDouble, 12, adFldIsNullable
        .Fields.Append "LabPaid", adDouble, 12, adFldIsNullable
        .Fields.Append "SprCahrged", adDouble, 12, adFldIsNullable
        .Fields.Append "SprFree", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Set Rst = GCn.Execute("select  Serv_Target,Serv_Desc,Serv_Type,Site_Code from Service_Type order by Serv_Desc")
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            With RstRep
               .AddNew
               .Fields("ServType") = Rst!Serv_Type
               .Fields("ServDesc") = Rst!Serv_Desc
               .Fields("Tgt") = IIf(IsNull(Rst!Serv_Target), 0, Rst!Serv_Target)
               Set RST1 = GCn.Execute("select count(DocId) as PrevJob from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And  Left(DocId,1)='" & PubDivCode & "' And   isNull(JobCloseDate) And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("Prev") = IIf(RST1.RecordCount = 0, 0, RST1!PrevJob)
               Set RST1 = GCn.Execute("select count(DocId) as NewJob from job_card where  " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("New") = IIf(RST1.RecordCount = 0, 0, RST1!NewJob)
               Set RST1 = GCn.Execute("select count(DocId) as CloseJob from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And   Not IsNull(JobCloseDate) And JobCloseDate  >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And JobCloseDate<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("Close") = IIf(RST1.RecordCount = 0, 0, RST1!CloseJob)
               Set RST1 = GCn.Execute("select count(DocId) as TotalJobClose from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  Not IsNull(JobCloseDate) And JobCloseDate  >=" & ConvertDate(Format(PubStartDate, "dd/MMM/yyyy")) & " And JobCloseDate<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("TotalJobClose") = IIf(RST1.RecordCount = 0, 0, RST1!TotalJobClose)
               Set RST1 = GCn.Execute("select sum(NetLab_Amt) as LabCharged from job_card where " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and Serv_Type = '" & Rst!Serv_Type & "' group by Serv_Type")
               .Fields("LabCahrged") = IIf(RST1.RecordCount = 0, 0, RST1!LabCharged)
               'Set Rst1 = GCn.Execute("SELECT sum(Job_GatePass.ContractAmt) as LabPaid FROM Job_GatePass LEFT JOIN Job_Lab ON Job_Card.DocId = Job_Lab.Job_DocID where Mid(DocId,3,1)='" & PubSiteCode & "' And Left(DocId,1)='" & PubDivCode & "' And  Not IsNull(Job_Card.JobCloseDate) and Job_Card.job_date >=#" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# And Job_Card.job_date<=#" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#  and Job_Card.Serv_Type = '" & Rst!Serv_Type & "' group by Job_Card.Serv_Type")
               Set RST1 = GCn.Execute("SELECT sum(JG.ContractAmt) as LabPaid FROM (Job_GatePass as JG LEFT JOIN Job_Card as JC on JG.Job_DocID=JC.DocId) where " & cMID("JC.DocId", "3", "1") & "='" & PubSiteCode & "' And Left(JC.DocId,1)='" & PubDivCode & "' And  Not IsNull(JC.JobCloseDate) and JC.Job_Date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And JC.Job_Date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and JC.Serv_Type = '" & Rst!Serv_Type & "' group by JC.Serv_Type")
               .Fields("LabPaid") = IIf(RST1.RecordCount = 0, 0, RST1!LabPaid)
               Set RST1 = GCn.Execute("SELECT sum(SP_Stock.Net_Amt) as NetAmt FROM Job_Card LEFT JOIN SP_Stock ON Job_Card.DocId = SP_Stock.Job_DocID where " & cMID("Job_Card.DocId", "3", "1") & "='" & PubSiteCode & "' And Left(Job_Card.DocId,1)='" & PubDivCode & "' And Not IsNull(Job_Card.JobCloseDate) and Job_Card.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And Job_Card.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and Job_Card.Serv_Type = '" & Rst!Serv_Type & "' and SP_Stock.Purpose = 'C' group by Job_Card.Serv_Type")
               .Fields("SprCahrged") = IIf(RST1.RecordCount = 0, 0, RST1!NetAmt)
               Set RST1 = GCn.Execute("SELECT sum(SP_Stock.Net_Amt) as NetAmtFree FROM Job_Card LEFT JOIN SP_Stock ON Job_Card.DocId = SP_Stock.Job_DocID where " & cMID("Job_Card.DocId", "3", "1") & "='" & PubSiteCode & "' And Left(Job_Card.DocId,1)='" & PubDivCode & "' And  Not IsNull(Job_Card.JobCloseDate) and Job_Card.job_date >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And Job_Card.job_date<=" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and Job_Card.Serv_Type = '" & Rst!Serv_Type & "' and SP_Stock.Purpose <> 'C' group by Job_Card.Serv_Type")
               .Fields("SprFree") = IIf(RST1.RecordCount = 0, 0, RST1!NetAmtFree)
               .Update
            End With
        Rst.MoveNext
        Loop
    End If
    Set Rst = Nothing
    Set RST1 = Nothing
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "ServWiseJob"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
'****************************************************************
End Sub
Private Sub DemdVsSuppProc()
'''?????88888//////////******************************************
Dim mQry As String, Condstr As String
Dim mPurpose As String
On Error GoTo ELoop
    Select Case FGrid.TextMatrix(List1, 1)
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
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Job_Card.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP_Stock.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Job_Card.Serv_Type in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(SP_Stock.DocId,1) in (" & GridString3 & ")"
    
    If FGrid.TextMatrix(List1, 1) <> "All" Then Condstr = Condstr & " and SP_Stock.Purpose = '" & mPurpose & "'"
    Condstr = Condstr & " and SP_Stock.V_type in ('" & WksGenReq & "','" & WksReqWrt & "')"
            
    mQry = "SELECT Job_Card.Job_No, Job_Card.Job_Date, Job_Card.Serv_Type,HisCard.RegNo, HisCard.Chassis, " & _
           "Left(Sp_Stock.docid,13) AS RequisitionNo,SP_Stock.V_Date AS RequiDt, SP_Stock.Part_No, SP_Stock.Purpose,SP_Stock.Qty_Doc, SP_Stock.Qty_Rec, SP_Stock.Qty_Iss, SP_Stock.Qty_Ret,SP_Stock.Part_No " & _
           "FROM (sp_stock left join job_card ON job_card.docid =sp_stock.job_docid) LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo"

    mQry = mQry + Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "DemdVsSupp"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Public Sub SelGridKeyPressLocal(Txt As Object, SelGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
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
          
Private Function CatFldName(Rst As ADODB.Recordset, SrvType As String) As Byte
Rst.MoveFirst
Rst.FIND ("Serv_Type = '" & SrvType & "'")
CatFldName = 21 + Rst.AbsolutePosition
End Function

Private Function CatFldName1(Rst As ADODB.Recordset, SrvType As String) As Byte
Rst.MoveFirst
Rst.FIND ("Serv_Type = '" & SrvType & "'")
CatFldName1 = 17 + Rst.AbsolutePosition
End Function

Private Function CatFldName2(Rst As ADODB.Recordset, SrvType As String) As Byte
Rst.MoveFirst
Rst.FIND ("Serv_Type = '" & SrvType & "'")
CatFldName2 = 29 + Rst.AbsolutePosition
End Function

Private Sub WksLabIncentProc()
On Error GoTo ELoop
'2% on chargeable labour-discount(60%-$40%)
'1% on pdi and free
'2% on warranty (proportaniate to salary)
'penalty clause
Dim mQry As String, Condstr As String, CondStr1 As String
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    Condstr = " where Incentives.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    CondStr1 = " where SecondDt  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SecondDt <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Incentives.DocId", "2", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Unchecked Then CondStr1 = CondStr1 & " and " & cMID("Penalty.FirstDicId", "2", "1") & " in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Incentives.DocId,1) in (" & GridString2 & ")"
    If Check1(2).Value = Unchecked Then CondStr1 = CondStr1 & " and left(Penalty.FirstDicId,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Incentives.Mech_Code in (" & GridString3 & ")"
    If Check1(3).Value = Unchecked Then CondStr1 = CondStr1 & " and Penalty.Mech_Code in (" & GridString3 & ")"
'    mQRY = "SELECT " & _
'            " Job_Card.Docid, SP_Stock.V_No, SP_Stock.V_Date, Job_Card.Job_No," & _
'            "Job_Card.Job_Date, Job_Card.JobCloseDate,Job_Card.DocId_InvSpr, HisCard.RegNo, " & _
'            "HisCard.Chassis, SP_Stock.Part_No, SP_Stock.Purpose, SP_Stock.Qty_Doc, SP_Stock.Qty_Iss," & _
'            "SP_Stock.Qty_Ret, SP_Stock.Rate, SP_Stock.Amount, " & _
'            "(SP_Stock.Claim_Div + SP_Stock.Claim_Site + SP_Stock.Claim_YearPrefix + SP_Stock.Claim_Type +  SP_Stock.Claim_No) as ClaimNo, SP_Stock.Claim_Date " & _
'           "FROM ((SP_Stock " & _
'                "LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO) " & _
'                "LEFT JOIN Job_Card ON SP_Stock.Job_DocID = Job_Card.DocId)" & _
'                "LEFT JOIN HisCard ON (Job_Card.CardNo = HisCard.CardNo)"
'    mQRY = mQRY + CondStr

    mQry = "SELECT Incentives.Mech_Code, max(Incentives.Designation) as Des, max(Incentives.Emp_Name) as EName, sum(Incentives.Share) as Share,0 as Penal FROM Incentives  " & Condstr & " group by Incentives.Mech_Code "
    mQry = mQry & " Union All  " & _
                  " select   Penalty.Mech_Code, max(Penalty.Designation) as des, max(Penalty.Emp_Name) as EName,0 as Share, sum((" & vIsNull("Penalty.Share", "0") & " * 1.5)) as Penal FROM Penalty " & CondStr1 & " group by Penalty.Mech_Code "
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksIncPen"
    RepTitle = UCase(Me.CAPTION)
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description

End Sub
Private Sub ServDueRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    Condstr = " where jc.NextSrvDate >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and jc.NextSrvDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    mQry = " select H.Model,H.RegNo,H.Chassis,H.Engine,H.Name,H.Add1,H.Add2,H.Add3,H.PhoneOff,JC.NextSrvDate" & _
           " from Job_Card JC left join HisCard H on JC.CardNo=H.CardNo "
           
    mQry = mQry + Condstr + " Order by NextSrvDate"
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "ServDueReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub QuatRetRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        
    Condstr = " where  jc.Job_Date <= " & ConvertDate(DateAdd("m", -3, Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy"))) & " "
    
    mQry = " select distinct H.Model,H.RegNo,H.Chassis,H.Engine,H.Name,H.Add1,H.Add2,H.Add3,H.PhoneOff,max(JC.Job_Date),JC.Job_No,Service_Type.Serv_Desc" & _
           " from (Job_Card JC left join HisCard H on JC.CardNo=H.CardNo) Left Join Service_type on JC.Serv_Type=Service_Type.Serv_Type "
           
    mQry = mQry + Condstr + " Group by H.Model,H.RegNo,H.Chassis,H.Engine,H.Name,H.Add1,H.Add2,H.Add3,H.PhoneOff,JC.Job_Date,JC.Job_No,Service_Type.Serv_Desc"
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "QuatRetReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub PostServiceFollowProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    Condstr = " where jc.Job_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and jc.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    mQry = " select distinct H.Model,H.RegNo,H.Chassis,H.Engine,H.Name,H.Add1,H.Add2,H.Add3,H.PhoneOff" & _
           " from Job_Card JC left join HisCard H on JC.CardNo=H.CardNo "
           
    mQry = mQry + Condstr + " Order by H.Name"
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "PostServFollowReg"
    RepTitle = UCase(Me.CAPTION)
    SubRep1 = False
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub STaxWSrvTaxRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    Condstr = " where SP.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    Condstr2 = " where JC.JobCloseDate >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
           
           If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(1).Value = Unchecked Then Condstr2 = Condstr2 & " and " & cMID("jc.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr2 = Condstr2 & " and " & cMID("jc.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
'    mQRY = "Select SP.V_Date,iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),sum(SP.Total_Amt),0) as Total_Amt,iif(SP.V_Type in ('SYSRC','SYSRR'),sum(SP.Total_Amt),0) as Total_AmtRet, " & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),Sum(SP.OilAmt_MRP_TP+SP.OilAmt_TP),0) as Oil_AmtTP," & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),Sum(SP.OilAmt_MRP_TP+SP.OilAmt_TP),0) as Oil_AmtRetTP, " & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),Sum(SP.OilAmt_TB),0) as Oil_AmtTB," & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),Sum(SP.OilAmt_TB),0) as Oil_AmtRetTB, " & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),Sum(SP.OilAmt_MRP_TB),0) as Oil_AmtMRPTB," & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),Sum(SP.OilAmt_MRP_TB),0) as Oil_AmtRetMRPTB, " & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),(Sum(SP.SprAmt_MRP_TP+SP.SprAmt_TP)),0) as TP_Spr, " & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),(Sum(SP.SprAmt_MRP_TP+SP.SprAmt_TP)-sum(SP.D_Amt_MRP_TP+D_Amt_TP)),0) as TP_SprRet, " & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),(Sum(SP.SprAmt_TB)-sum(D_Amt_TB)),0) as TB_Spr," & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),(Sum(SP.SprAmt_TB)-sum(D_Amt_TB)),0) as TB_SprRet," & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),(Sum(SP.SprAmt_MRP_TB-SP.Tax_AmtMRP-SP.TOT_AmtMRP)-sum(SP.D_Amt_MRP_TB)),0) as TB_SprMRP," & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),(Sum(SP.SprAmt_MRP_TB-SP.Tax_AmtMRP-SP.TOT_AmtMRP)-sum(SP.D_Amt_MRP_TB)),0) as TB_SprRetMRP," & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),sum(SP.Tax_Amt+SP.Tax_AmtMRP),0) as Tax_Amt," & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),sum(SP.Tax_Amt+SP.Tax_AmtMRP),0) as Tax_AmtRet," & _
'           "iif(SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR'),sum(TOT_Amt+TOT_AmtMRP),0) as SdtAmt," & _
'           "iif(SP.V_Type in ('SXSRC','SXSRR'),sum(TOT_Amt+TOT_AmtMRP),0) as SdtAmtRet," & _
'           "sum(SP.Packing) as Packing,0 as TBLab,0 as TPLab,0 as FreeLab,0 as EXLab,0 as NetLab,0 as Lab_TaxAmt,sum(Rounded) as ROff,0 as Lab_D_Amt,0 as PDIWARR,sum(D_Amt_TP) as DiscAmtTP,sum(D_Amt_TB) as DiscAmtTB " & _
'           " From SP_Sale as SP " & Condstr
           
           
    mQry = "Select SP.V_Date," & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "sum(SP.Total_Amt)", "0") & " as Total_Amt, " & _
           "" & cIIF("SP.V_Type in ('SYSRC','SYSRR')", "sum(SP.Total_Amt)", "0") & " as Total_AmtRet, " & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "Sum(SP.OilAmt_MRP_TP+SP.OilAmt_TP)", "0") & " as Oil_AmtTP," & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "Sum(SP.OilAmt_MRP_TP+SP.OilAmt_TP)", "0") & " as Oil_AmtRetTP, " & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "Sum(SP.OilAmt_TB)", "0") & " as Oil_AmtTB," & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "Sum(SP.OilAmt_TB)", "0") & " as Oil_AmtRetTB, " & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "Sum(SP.OilAmt_MRP_TB)", "0") & " as Oil_AmtMRPTB," & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "Sum(SP.OilAmt_MRP_TB)", "0") & " as Oil_AmtRetMRPTB, " & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "(Sum(SP.SprAmt_MRP_TP+SP.SprAmt_TP))", "0") & " as TP_Spr, " & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(Sum(SP.SprAmt_MRP_TP+SP.SprAmt_TP)-sum(SP.D_Amt_MRP_TP+D_Amt_TP))", "0") & " as TP_SprRet, " & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "(Sum(SP.SprAmt_TB)-sum(D_Amt_TB))", "0") & " as TB_Spr," & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(Sum(SP.SprAmt_TB)-sum(D_Amt_TB))", "0") & " as TB_SprRet," & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "(Sum(SP.SprAmt_MRP_TB-SP.Tax_AmtMRP-SP.TOT_AmtMRP)-sum(SP.D_Amt_MRP_TB))", "0") & " as TB_SprMRP," & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(Sum(SP.SprAmt_MRP_TB-SP.Tax_AmtMRP-SP.TOT_AmtMRP)-sum(SP.D_Amt_MRP_TB))", "0") & " as TB_SprRetMRP," & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "sum(SP.Tax_Amt+SP.Tax_AmtMRP)", "0") & " as Tax_Amt," & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "sum(SP.Tax_Amt+SP.Tax_AmtMRP)", "0") & " as Tax_AmtRet," & _
           "" & cIIF("SP.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR')", "sum(TOT_Amt+TOT_AmtMRP)", "0") & " as SdtAmt," & _
           "" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "sum(TOT_Amt+TOT_AmtMRP)", "0") & " as SdtAmtRet," & _
           "sum(SP.Packing) as Packing,0 as TBLab,0 as TPLab,0 as FreeLab,0 as EXLab,0 as NetLab,0 as Lab_TaxAmt, " & _
           "sum(Rounded) as ROff,0 as Lab_D_Amt,0 as PDIWARR,sum(D_Amt_TP) as DiscAmtTP,sum(D_Amt_TB) as DiscAmtTB " & _
           " From SP_Sale as SP " & Condstr
           
           
    mQry = mQry + " and SP.V_Type in ('SYSIC','SYSIR','SXSRC','SXSRR','W_SIC','W_SIR') GROUP BY SP.V_Date,SP.V_Type"
    
    mQry = mQry + " Union All " + "Select JC.JobCloseDate,0 as Total_Amt,0 as Total_AmtRet,0 as Oil_AmtTP,0 as Oil_AmtRetTP,0 as Oil_AmtTB, " & _
           " 0 as Oil_AmtRetTB,0 as Oil_AmtMRPTB,0 as Oil_AmtRetMRPTB,0 as TP_Spr,0 as TP_SprRet " & _
           ",0 as TB_Spr,0 as TB_SprRet,0 as TB_SprMRP,0 as TB_SprRetMRP,0 as Tax_Amt,0 as Tax_AmtRet, " & _
           "0 as SDTAmt,0 as SDTAmtRet,0 as Packing, (JC.LabAmt_TB-JC.Lab_D_Amt) as TBLab,(JC.LabAmt_TP) as TPLab, " & _
           "Sum(" & cIIF("(JL.Chrg_From in ('M','O') and JL.Chrg_Type not in ('P','W'))", "JL.LabourAmt", "0") & ") as FreeLab, " & _
           "Sum(" & cIIF("(JL.External_YN='1' and JL.Chrg_Type <> 'F')", "JL.LabourAmt", "0") & ") as EXLab, " & _
           "Sum(" & cIIF("JL.Chrg_From in ('C','M','O') and JL.Chrg_Type not in ('P','W')", "JL.LabourAmt", "0") & ")-JC.Lab_D_Amt as NetLab " & _
           ",JC.Lab_TaxAmt,0 as ROff,JC.Lab_D_Amt,Sum(" & cIIF("JL.Chrg_Type in ('P','W')", "JL.LabourAmt", "0") & ") as PDIWarr, " & _
           "0 AS DiscAmtTP,0 AS DiscAmtTB " & _
           " From Job_Card JC Left Join Job_Lab JL on JC.Docid=JL.Job_DocId " & Condstr2

    mQry = mQry + " GROUP BY Jc.DocId,JC.JobCloseDate,JC.Lab_TaxAmt,JC.Lab_D_Amt,JC.LabAmt_TB,JC.LabAmt_TP,JC.Lab_TaxPer"
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "STaxWSrvTaxReg"
    RepTitle = UCase(Me.CAPTION)
    SubRep1 = False
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ReptJobAnaProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$
Dim TmpRst As ADODB.Recordset
Dim LastDt, CurDt As Date
Dim LstJobNo, CurJobNo, CurKms, LstKms  As Double
Dim CurMech, LstMech As String


    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    Condstr = " where Job_Card.Job_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    Set RstRep = New ADODB.Recordset
    With RstRep
    
        .Fields.Append "RegNo", adVarChar, 14, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 20, adFldIsNullable
        .Fields.Append "Model", adVarChar, 40, adFldIsNullable
        .Fields.Append "LstJobNo", adVarChar, 15, adFldIsNullable
        .Fields.Append "LstJobDt", adDate, 15, adFldIsNullable
        .Fields.Append "LstMech", adVarChar, 50, adFldIsNullable
        .Fields.Append "CurJobDt", adDate, 15, adFldIsNullable
        .Fields.Append "CurJobNo", adVarChar, 15, adFldIsNullable
        .Fields.Append "CurMech", adVarChar, 50, adFldIsNullable
        
        
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
        
    End With
    
    Set RstRep1 = GCn.Execute("Select Count(*),Job_Card.CardNo from Job_Card " & Condstr & " Group by Job_Card.CardNo Having Count(*)>=2")
    If RstRep1.RecordCount > 0 Then
        Do While RstRep1.EOF = False
            With RstRep
                Set TmpRst = GCn.Execute("Select * from Job_Card where CardNo='" & RstRep1!CardNo & "' Order By Job_date Desc ")
                    
                    CurDt = TmpRst!Job_Date
                    CurJobNo = TmpRst!Job_No
                    CurKms = TmpRst!AtKMsHrs
                    If GCn.Execute("Select Emp_Name from Emp_Mast where Emp_Code='" & TmpRst!DelBy & "' ").RecordCount > 0 Then
                        CurMech = GCn.Execute("Select Emp_Name from Emp_Mast where Emp_Code='" & TmpRst!DelBy & "' ").Fields(0).Value
                    End If
                    
                    TmpRst.MoveNext
                    LastDt = TmpRst!Job_Date
                    LstJobNo = TmpRst!Job_No
                    LstKms = TmpRst!AtKMsHrs
                    If GCn.Execute("Select Emp_Name from Emp_Mast where Emp_Code='" & TmpRst!DelBy & "' ").RecordCount > 0 Then
                        LstMech = GCn.Execute("Select Emp_Name from Emp_Mast where Emp_Code='" & TmpRst!DelBy & "' ").Fields(0).Value
                    End If
                    
                If FGrid.TextMatrix(List1, 1) = "Days" Then
                    If DateDiff("D", LastDt, CurDt) <= Val(FGrid.TextMatrix(Cat1, 1)) Then
                        .AddNew
                        .Fields("RegNo") = GCn.Execute("Select RegNo from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                        .Fields("Chassis") = GCn.Execute("Select Chassis from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                        .Fields("Model") = GCn.Execute("Select Model from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                        .Fields("LstJobDt") = LastDt
                        .Fields("CurJobDt") = CurDt
                        .Fields("LstJobNo") = LstJobNo
                        .Fields("CurJobNo") = CurJobNo
                        .Fields("CurMech") = CurMech
                        .Fields("LstMech") = LstMech
                        .Update
                    End If
                Else
                    If Val(CurKms - LstKms) <= Val(FGrid.TextMatrix(Cat1, 1)) Then
                        .AddNew
                        .Fields("RegNo") = GCn.Execute("Select RegNo from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                        .Fields("Chassis") = GCn.Execute("Select Chassis from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                        .Fields("Model") = GCn.Execute("Select Model from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                        .Fields("LstJobDt") = LastDt
                        .Fields("CurJobDt") = CurDt
                        .Fields("LstJobNo") = LstJobNo
                        .Fields("CurJobNo") = CurJobNo
                        .Fields("CurMech") = CurMech
                        .Fields("LstMech") = LstMech
                        .Update
                    End If
                End If
                
    End With
        
        RstRep1.MoveNext
        Loop
    End If
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "RepeatJobDet"
    RepTitle = UCase(Me.CAPTION)
    SubRep1 = False
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


