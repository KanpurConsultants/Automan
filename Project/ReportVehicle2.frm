VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form ReportVehicle2 
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
      DownPicture     =   "ReportVehicle2.frx":0000
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
      DownPicture     =   "ReportVehicle2.frx":3132
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
Attribute VB_Name = "ReportVehicle2"
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
Private Const FinCertificate As Byte = 1
Private Const InputTaxReg As Byte = 2
Private Const OutputTaxReg As Byte = 3
Private Const VATDiffReg As Byte = 4
Private Const Profitability As Byte = 5
Private Const DoPendingReport1 As Byte = 31
Private Const DoRecive1 As Byte = 32
'***********************************************************

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
    Case FinCertificate
        FinCertificateProc
        If RepPrint = False Then Exit Sub
    Case InputTaxReg
        InputTaxRegProc
        If RepPrint = False Then Exit Sub
    Case OutputTaxReg
        OutPutTaxRegProc
        If RepPrint = False Then Exit Sub
    Case VATDiffReg
        VATDiffRegProc
        If RepPrint = False Then Exit Sub
    Case Profitability
        ProcProfitability
        If RepPrint = False Then Exit Sub
    Case DoPendingReport1
        DoPendingReport
        If RepPrint = False Then Exit Sub
    Case DoRecive1
         DoRecive
         If RepPrint = False Then Exit Sub
End Select
       
CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
 If RepName <> "DoPending" And RepName <> "DoRecive" Then
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")

   rpt.Database.SetDataSource RstRep
   If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
    rpt.ReadRecords
 Else
 Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")
 rpt.Database.SetDataSource RstRep
    rpt.ReadRecords
    
End If

Set RstRep = Nothing
'Set rpt = Nothing   '  auto nothing in report view

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
'            Case ModWiseGWise
'                ListArray = Array("GroupWise", "VehicleType")
'                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VATDiffReg
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case Profitability
                ListArray = Array("Model", "Site")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)

        End Select
    Case List2
        Select Case GRepFormName
'           Case VehSalereg
'               ListArray = Array("Yes", "No")
'               Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)

        End Select
    Case List3
        Select Case GRepFormName
'           Case VehSalereg
'               ListArray = Array("PartyWise", "CityWise", "FinancierGrp", "FinancierName", "FormType", "All")
'               Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 6)
        End Select
    Case Cat2
        Select Case GRepFormName
'            Case ModFWiseMicro
'                FGrid.TextMatrix(Cat2, 1) = Val(FGrid.TextMatrix(Cat1, 1)) + 1
'                TxtGridLeave
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
        Select Case GRepFormName
'            Case ModFWiseMicro
'                NumPress TxtGrid(Index), KeyAscii, 4, 0
        End Select
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
        TxtGrid(0).TEXT = GetYear
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
        If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
    Case List3
        If TxtGrid(0).TEXT <> "" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            Select Case TxtGrid(0).TEXT
                Case "PartyWise"
                    Grid3Sql = "select '' as O,Name as Party_Name,SubCode  as code from Subgroup order by Name"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                Case "CityWise"
                    Grid3Sql = "select '' as O,CityName as City_Name,CityCode  as code from City order by CityName"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                Case "FinancierGrp"
                    Grid3Sql = "select '' as O,FinGrpName as FinGrp_Name,FinGrpCode  as code from FinGroup order by FinGrpName"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                Case "FinancierName"
                    Grid3Sql = "select '' as O,FinName as Financer_Name,FinCode  as code from ContractFinance order by FinName"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                Case "FormType"
                    Grid3Sql = "select '' as O,Form_Desc as City_Name,Form_Code  as code from TaxForms order by Form_Desc"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                Case "All"
                    GridSel(3).Visible = False: Check1(3).Visible = False
            End Select
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

Private Function FillString(GridArray As Variant, Gridindex As Integer, DataType As Byte, Optional DetailStr As Boolean) As String
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
      sitecond = "where  Site_code ='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If

Select Case GRepFormName
    Case FinCertificate
        With FGrid
            .TextMatrix(Date1, 0) = "For Month/Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date1: mHelpGridNo = 1
        Grid1Sql = "select distinct '' as O, FinName as Name, FinBankCode as Code from ContractFinance where FinCatg=0 order by FinName"
        GridInitialise 1, Grid1Sql
    
    Case InputTaxReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 1
        Grid1Sql = "select distinct '' as O, Site_Desc as Name, Site_Code as Code from Site " & sitecond & ""
        GridInitialise 1, Grid1Sql
    Case OutputTaxReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 1
        Grid1Sql = "select distinct '' as O, Site_Desc as Name, Site_Code as Code from Site " & sitecond & ""
        GridInitialise 1, Grid1Sql
        
    Case Profitability
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Group By": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Model"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 1
        Grid1Sql = "select distinct '' as O, Site_Desc as Name, Site_Code as Code from Site " & sitecond & ""
        GridInitialise 1, Grid1Sql
                
                
    Case VATDiffReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Margin": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Yes"
            
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 1
        Grid1Sql = "select distinct '' as O, Site_Desc as Name, Site_Code as Code from Site " & sitecond & ""
        GridInitialise 1, Grid1Sql
        
        Case DoPendingReport1
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 1
        Grid1Sql = "select distinct '' as O, Site_Desc as Name, Site_Code as Code from Site " & sitecond & ""
        GridInitialise 1, Grid1Sql
   Case DoRecive1
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 1
        Grid1Sql = "select distinct '' as O, Site_Desc as Name, Site_Code as Code from Site " & sitecond & ""
        GridInitialise 1, Grid1Sql
        
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
Dim RstCompDet As ADODB.Recordset


Select Case GRepFormName
    Case FinCertificate
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("Title")
'                    rpt.FormulaFields(i).Text = " 'FINANCER'S CERTIFICATE'"
            End Select
        Next
    Case InputTaxReg, OutputTaxReg, VATDiffReg, Profitability, DoPendingReport1, DoRecive1
        Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'")
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("Comp_Lst")
                    rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecLST & "  " & RstCompDet!V_SecLST_Date & "'"
                Case UCase("Comp_CST")
                    rpt.FormulaFields(I).TEXT = "'" & RstCompDet!V_SecCST & "  " & RstCompDet!V_SecCST_Date & "'"
                Case UCase("PubStartDate")
                    rpt.FormulaFields(I).TEXT = "'" & PubStartDate & "'"
                Case UCase("PubEndDate")
                    rpt.FormulaFields(I).TEXT = "'" & PubEndDate & "'"
            End Select
        Next



    'Case ModWiseGWise
    '    For i = 1 To rpt.FormulaFields.Count
    '        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
    '            Case UCase("Division")
    '                If GridString2 = "" Then
    '                    rpt.FormulaFields(i).Text = "'Division : All'"
    '                Else
    '                    rpt.FormulaFields(i).Text = "'Division :'+ '" & Replace(GridString2, "'", "") & "'"
    '                End If
    '        End Select
    '    Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub FinCertificateProc()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    Condstr = " where left(VO.Inv_DocID,1)='" & PubDivCode & "' and " & cCStr("Month(VO.Inv_Date)") & " + " & cCStr("Year(VO.Inv_Date)") & " = '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "'"
    
    mQry = " SELECT CF.FinName,VO.Model,VO.Chassis,SG.Name,VO.Inv_Prefix,VO.Inv_DocID,VO.Inv_Date,VO.Inv_Prefix+right(VO.Inv_DocID,8) as SalInvNo " & _
           " FROM (Veh_Order VO left Join SubGroup SG on VO.PartyCode=Sg.SubCode) " & _
        " left join ContractFinance CF on VO.FB_Code=CF.FinCode "

    mQry = mQry & Condstr & " Order By VO.Model,VO.Inv_Date"

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "FinCertificate"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

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

Private Function GetYear() As Integer
Dim mYear%
mYear = Val(TxtGrid(0).TEXT)
        If mYear = 0 Then mYear = Year(date)
        If mYear > 1999 Then mYear = Right(STR(mYear), 2)
        mYear = Val(mID(CStr(Year(date)), 1, 4 - Len(Trim(CStr(mYear)))) + Trim(CStr(mYear)))
GetYear = mYear
End Function
Private Sub InputTaxRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    Condstr = " where VP1.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VP1.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VP1.DocId", "3", "1") & " in (" & GridString1 & ")"
    
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VP1.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    
    mQry = "Select VP1.DocId,VP1.V_No,VP1.Tax_Per,VP1.Tax_Amt,SG.Name,SG.LSTNo, " & _
        " VP1.PBill_No as Party_Doc_No,VP1.PBill_Date as Party_Doc_Date, " & _
        " VP1.Tot_Amount as Net_Amt,VP1.Amount+VP1.Exsice+VP1.Addition-VP1.Deduction as Amount, " & _
        " VP1.V_Date, VS.ChassisNo, Sg.Add1, Sg.Add2, Sg.Add3, City.CityName, VP1.SatPer, VP1.SatAmt " & _
        " From ((Veh_Purch1 as VP1 " & _
        " Left Join SubGroup Sg on VP1.PartyCode=SG.SubCode) " & _
        " Left Join City  On Sg.CityCode = City.CityCode) " & _
        " Left Join Veh_Stock VS On VP1.DocId = VS.Pur_DocID "
        
    mQry = mQry & Condstr & " Order By VP1.V_Date,VP1.DocID"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    If StrCmp(left(PubComp_Name, 4), "Yash") Then
        RepName = "VehInputTaxReg_Yash"
    Else
        RepName = "VehInputTaxReg"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub OutPutTaxRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    Condstr = " where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Vo.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Vo.Inv_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    
    mQry = "Select VO.Inv_DocId as DocId,VO.Inv_No as V_No,VO.Tax_Per,VO.Tax_Amt,SG.Name,SG.LSTNo,VO.Net_Amount as Net_Amt, " & _
            "VO.SubTot as Amount,VO.Inv_Date as V_Date,  VStk.ChassisNo, VO.SatPer, VO.SatAmt " & _
            "From ((Veh_Order as VO Left Join Veh_Stock VStk On VO.Inv_DocId=VStk.Sal_DocId) Left Join SubGroup Sg on VO.PartyCode=SG.SubCode) "
        
    mQry = mQry & Condstr & " Order By VO.Inv_Date,VO.Inv_DocID"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "VehOutPutTaxReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VATDiffRegProc()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    Condstr = " where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VP1.DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VP1.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    mQry = "Select VO.Inv_DocId as DocId,VO.Inv_No as V_No,VO.Tax_Per," & cIIF("ST.L_C='Central'", "0", "VO.Tax_Amt") & " As SaleTax_Amt,SG.Name, " & _
            "SG.LSTNo,VO.Net_Amount as Net_Amt,VO.SubTot-VO.Rebate as Amount,VO.Inv_Date, " & _
            "VP1.DocId as DocId1,VP1.V_No as V_No1,VP1.Tax_Per," & cIIF("PT.L_C='Central'", "0", "VP1.Tax_Amt") & " As PurTax_Amt, " & _
            "VP1.PBill_No as Party_Doc_No,VP1.Tot_Amount as Net_Amt1, " & _
            "VP1.Amount+VP1.Addition-VP1.Deduction+VP1.Misc_Amt as Amount1,VStk.ChassisNo, " & _
            "" & cIIF("'" & PubSiebelActiveYn & "' = '1'", "MG.ModelGrp_Name", "VStk.Model") & " as Model, " & _
            "VStk.EngineNo,VP1.Tax_Amt as TaxAmt1,VP1.PBill_Date, VO.Rebate " & _
            "From ((((((Veh_Order as VO Left Join SubGroup Sg on VO.PartyCode=SG.SubCode) " & _
            "Left Join Veh_Stock VStk on VO.Chassis=Vstk.Chassisno) " & _
            "Left Join Model on Model.Model=Vstk.Model) " & _
            "Left Join Model_Grp MG On Model.Grp_Code=MG.ModelGrp_Code) " & _
            "Left Join VEH_Purch1 as VP1 on Vstk.Pur_Docid=VP1.Docid) " & _
            "Left Join TaxForms PT On PT.Form_Code=VP1.Form_Code)" & _
            "Left Join TaxForms ST On ST.Form_Code=VO.Form_Code"
    mQry = mQry & Condstr & " Order By VO.Inv_No"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    If FGrid.TextMatrix(List1, 1) = "Yes" Then
        RepName = "VATDiffRegMargin"
    Else
        RepName = "VATDiffReg"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ProcProfitability()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    Condstr = " where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VStk.sal_DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VP1.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    mQry = "Select VO.Inv_DocId as DocId,VO.Inv_No as V_No,VO.Tax_Per," & cIIF("ST.L_C='Central'", "0", "VO.Tax_Amt") & " As SaleTax_Amt,SG.Name, " & _
            "SG.LSTNo,VO.Net_Amount as Net_Amt, VO.Rebate, Vo.Subvention As Subvention_Given_to_Customer,VO.SubTot as Amount,VO.Inv_Date, " & _
            "VP1.DocId as DocId1,VP1.V_No as V_No1,VP1.Tax_Per," & cIIF("PT.L_C='Central'", "0", "VP1.Tax_Amt") & " As PurTax_Amt, " & _
            "VP1.PBill_No as Party_Doc_No,VP1.Tot_Amount as Net_Amt1, " & _
            "VP1.Amount+VP1.Addition-VP1.Deduction+VP1.Misc_Amt as Amount1,VStk.ChassisNo, " & _
            "" & cIIF("'" & PubSiebelActiveYn & "' = '1'", "MG.ModelGrp_Name", "VStk.Model") & " as Model, " & _
            "VStk.EngineNo,VP1.Tax_Amt as TaxAmt1,VP1.PBill_Date, (Select Sum(TataContribution) From Subvention Sv Where Sv.ModelGroup=Model.Grp_code And Vo.Inv_Date Between Sv.FromDate and Sv.ToDate) As SubventionTata, VO.SpecialDiscount, (Select Sum(rate) From Body_PurchDetail Where Body_PurchDetail.ChassisNo=VStk.ChassisNo) as BodyCharges, Site.Site_Desc " & _
            "From ((((((Veh_Order as VO Left Join SubGroup Sg on VO.PartyCode=SG.SubCode) " & _
            "Left Join Veh_Stock VStk on VO.Chassis=Vstk.Chassisno) " & _
            "Left Join Model on Model.Model=Vstk.Model) " & _
            "Left Join Model_Grp MG On Model.Grp_Code=MG.ModelGrp_Code) " & _
            "Left Join VEH_Purch1 as VP1 on Vstk.Pur_Docid=VP1.Docid) " & _
            "Left Join TaxForms PT On PT.Form_Code=VP1.Form_Code)" & _
            "Left Join TaxForms ST On ST.Form_Code=VO.Form_Code " & _
            "Left Join Site On Site.Site_Code=SubString(VO.Inv_DocID,3,1) "
    mQry = mQry & Condstr & " Order By VO.Inv_No"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
    If StrCmp(FGrid.TextMatrix(List1, 1), "MODEL") Then
    
        RepName = "Vehicle_Profitability_JMA"
    Else
        RepName = "Vehicle_ProfitabilitySiteWise_JMA"
    End If
Else
    RepName = "Vehicle_Profitability"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub DoPendingReport()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    Condstr = " where VO.DOIssueDate>''and VO.Inv_DocId ='' and VO.DOIssueDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.DOIssueDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Ord_SiteCode", "2", "2") & " in (" & GridString1 & ")"
    
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VP1.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    mQry = "SELECT VO.DoReciveDate ,VO.DOIssueDate ,Sg.Name,CF.FinName FROM Veh_Order VO LEFT JOIN SubGroup sg ON Vo.PartyCode =Sg.SubCode LEFT JOIN ContractFinance CF ON VO.FB_CODE  =CF.FinCode  "
    mQry = mQry & Condstr & " " ' and Order By VO.DOIssueDate"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
        RepName = "DoPending"

    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub DoRecive()
On Error GoTo ELoop
Dim mQry$, Condstr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    Condstr = " where VO.DOIssueDate>''and VO.Inv_DocId <>'' and VO.DOIssueDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.DOIssueDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Ord_SiteCode", "2", "2") & " in (" & GridString1 & ")"
    
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VP1.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    mQry = "SELECT Vo.Inv_No ,Vo.Inv_Date ,VO.DoReciveDate ,VO.DOIssueDate ,Sg.Name,CF.FinName FROM Veh_Order VO LEFT JOIN SubGroup sg ON Vo.PartyCode =Sg.SubCode LEFT JOIN ContractFinance CF ON VO.FB_CODE  =CF.FinCode  "
    mQry = mQry & Condstr & "  Order By VO.DOIssueDate"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "DoRecive"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


