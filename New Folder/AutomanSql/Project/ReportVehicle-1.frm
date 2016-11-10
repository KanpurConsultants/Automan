VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form ReportVehicle1 
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
      DownPicture     =   "ReportVehicle-1.frx":0000
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
      DownPicture     =   "ReportVehicle-1.frx":3132
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
Attribute VB_Name = "ReportVehicle1"
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




'///********VEHICLE***********///
Private Const ModBrWiseStk              As Byte = 1
Private Const ModBrWiseSD               As Byte = 2
Private Const MntTarSQty                As Byte = 3
Private Const ModFWiseSRep              As Byte = 4
Private Const ModWiseGWise              As Byte = 5
Private Const ArWiseModWYr              As Byte = 6
Private Const ModWiseFWise              As Byte = 7
Private Const ModFWiseMicro             As Byte = 8
Private Const ModWiseCustm              As Byte = 9
Private Const SalePurchDiff             As Byte = 10
Private Const DelayDelInt               As Byte = 11
Private Const AreaWiseFWise             As Byte = 12
Private Const eMailSale                 As Byte = 13
Private Const VehSaleFormW              As Byte = 14
Private Const VehPSAudit                As Byte = 15
Private Const MonthlySumm               As Byte = 16
Private Const MonthWiseModelWiseSales   As Byte = 17
'************************************




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
    Case ModBrWiseStk
         ModBrWiseStkProc
        If RepPrint = False Then Exit Sub
    Case ModBrWiseSD
        ModBrWiseSDProc
        If RepPrint = False Then Exit Sub
    Case MntTarSQty
        MntTarSQtyProc
        If RepPrint = False Then Exit Sub
    Case ModFWiseSRep
        ModFWiseSRepProc
        If RepPrint = False Then Exit Sub
    Case ModWiseGWise
        ModWiseGWiseProc
        If RepPrint = False Then Exit Sub
    Case ModWiseFWise    'Report for model/financier wise sales report(monthly)
        ModWiseFWiseProc
        If RepPrint = False Then Exit Sub
    Case ArWiseModWYr
        ArWiseModWYrProc
        If RepPrint = False Then Exit Sub
    Case ModFWiseMicro
        ModFWiseMicroProc
        If RepPrint = False Then Exit Sub
    Case ModWiseCustm
        ModWiseCustmProc
        If RepPrint = False Then Exit Sub
    Case SalePurchDiff
        SalePurchDiffProc
        If RepPrint = False Then Exit Sub
    Case DelayDelInt
        DelayDelIntProc
        If RepPrint = False Then Exit Sub
    Case AreaWiseFWise
        AreaWiseFWiseProc
        If RepPrint = False Then Exit Sub
    Case eMailSale
        eMailSaleProc
        If RepPrint = False Then Exit Sub
    Case VehSaleFormW
        VehSaleFormWProc
        If RepPrint = False Then Exit Sub
    Case VehPSAudit
        VehPSAuditProc
        If RepPrint = False Then Exit Sub
    Case MonthlySumm
        MonthlySummProc
        If RepPrint = False Then Exit Sub
    Case MonthWiseModelWiseSales
        MonthWiseModelWiseSalesProc
        If RepPrint = False Then Exit Sub
End Select
       
CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")

rpt.Database.SetDataSource RstRep
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
rpt.ReadRecords
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
        Case ModWiseGWise
                ListArray = Array("GroupWise", "VehicleType")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
        Case ModBrWiseSD
                ListArray = Array("Model", "Model Group")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                
          End Select
    Case List2
        Select Case GRepFormName
            Case MonthWiseModelWiseSales
                ListArray = Array("Model", "Model Group")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
        Case ModBrWiseSD
                ListArray = Array("Delivery Date", "Invoice Date")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                
        End Select
    Case List3
    Select Case GRepFormName
'     Case VehSalereg
'            ListArray = Array("PartyWise", "CityWise", "FinancierGrp", "FinancierName", "FormType", "All")
'            Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 6)
     End Select
    
    Case Cat2
    Select Case GRepFormName
    Case ModFWiseMicro
    
     FGrid.TextMatrix(Cat2, 1) = Val(FGrid.TextMatrix(Cat1, 1)) + 1
    TxtGridLeave
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
            Case ModFWiseMicro
                NumPress TxtGrid(Index), KeyAscii, 4, 0
        End Select
   
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
            TxtGrid(0).TEXT = GetYear
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
               If TxtGrid(0).TEXT <= (Format(PubStartDate, "YYYY") - 3) Or TxtGrid(0).TEXT >= (Format(PubEndDate, "YYYY")) Then
               MsgBox "Invalid Year Selection!"
                TxtGridLeave = False: Exit Function
               End If
               
        Case Cat2
             FGrid.TextMatrix(Cat2, 1) = Val(FGrid.TextMatrix(Cat1, 1)) + 1
'             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
        Case List1
            If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            If TxtGrid(0).TEXT <> "" And GRepFormName = ModWiseGWise Then
                If FGrid.TextMatrix(List1, 1) = "GroupWise" Then
                    Grid4Sql = "select '' as O,Model_Grp.ModelGrp_Name As ModelGroup,Model_Grp.ModelGrp_Code As Code from Model_Grp order by Model_Grp.ModelGrp_Name"
                Else
                    Grid4Sql = "select '' as O,Vehicle_Type.Vehicle_Type As VehicleType,Vehicle_Type.Vehicle_Type As Code from Vehicle_Type order by Vehicle_Type.Vehicle_Type"
                End If
                GridInitialise 4, Grid4Sql
            End If
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
      sitecond = "where site_code ='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
Select Case GRepFormName
    Case ModBrWiseSD   'vijay Vehicle
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Group By": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Filter On": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Model Group"
            .TextMatrix(List2, 1) = "Invoice Date"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql

    Case ModBrWiseStk, MntTarSQty   'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "As On Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date1: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
    Case ModFWiseSRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & "  order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,FinGroup.FinGrpName As FinancerGroup,FinGroup.FinGrpCode As Code from FinGroup order by FinGroup.FinGrpName"
        GridInitialise 4, Grid4Sql
        
    Case ModWiseCustm
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

       End With
        mFirstRow = Date1: mLastRow = Date2:
        mHelpGridNo = 3
        Grid1Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,ContractFinance.FinName As FinancerName,ContractFinance.FinCode As Code from ContractFinance order by ContractFinance.FinName"
        GridInitialise 2, Grid2Sql
        
         Grid3Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 3, Grid3Sql

    Case SalePurchDiff
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

       End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 3, Grid3Sql
        
    Case AreaWiseFWise
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

       End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Area.Areaname As Area,Area.Areacode As Code from Area order by Area.Areaname"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,ContractFinance.FinName As FinancerName,ContractFinance.FinCode As Code from ContractFinance order by ContractFinance.FinName"
        GridInitialise 4, Grid4Sql

    Case VehSaleFormW  'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

       End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,TaxForms.Form_desc As FormName,TaxForms.Form_code As Code from TaxForms order by TaxForms.Form_desc"
        GridInitialise 3, Grid3Sql
    Case ModWiseFWise  'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

       End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,ContractFinance.FinName As FinancerName,ContractFinance.FinCode As Code from ContractFinance order by ContractFinance.FinName"
        GridInitialise 4, Grid4Sql
    Case ArWiseModWYr
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

       End With
        mFirstRow = Date1: mLastRow = Date2:
        mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,area.areaname As Area,area.areacode As Code from area order by area.areaname"
        GridInitialise 4, Grid4Sql
          
    Case DelayDelInt, eMailSale 'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

       End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql

        
    Case ModFWiseMicro  'vijay Vehicle 16/11/02
        With FGrid
            
             .TextMatrix(Cat1, 0) = "For Year": .RowHeight(Cat1) = GridRowHeight
             .TextMatrix(Cat2, 0) = "For Year": .RowHeight(Cat2) = GridRowHeight
            .TextMatrix(Cat1, 1) = Format(PubStartDate, "YYYY")
                   
             .TextMatrix(Cat2, 1) = Format(PubEndDate, "YYYY")
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
          GridInitialise 2, Grid2Sql
    Case ModWiseGWise 'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "On Group/Type": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubEndDate
            .TextMatrix(List1, 1) = "GroupWise"
       End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
         
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,Model_Grp.ModelGrp_Name As ModelGroup,Model_Grp.ModelGrp_Code As Code from Model_Grp order by Model_Grp.ModelGrp_Name"
        GridInitialise 4, Grid4Sql
        
    Case VehPSAudit    'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
           
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,BMS.BMS_Name As Category,BMS.BMS_Code As Code from BMS order by BMS.BMS_Name"
          GridInitialise 2, Grid2Sql
    Case MonthlySumm
        With FGrid
            .TextMatrix(Date1, 0) = "For Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date1: mHelpGridNo = 3
        'Grid1Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        
        
    Case MonthWiseModelWiseSales
        With FGrid
            .TextMatrix(List1, 0) = "Taxable/TaxPaid/All": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Model/Group": .RowHeight(List2) = GridRowHeight

            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List2, 1) = "Model Group"
        End With
        
        mFirstRow = List1: mLastRow = List2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
            "left join " & FaTable("AcGroup") & " As AcGroup on SubGroup.GroupCode=AcGroup.GroupCode " & _
            "Where  " & _
            "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 4, Grid4Sql
        

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
Case ModBrWiseStk
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'As On Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + '' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next


Case ModBrWiseStk, MonthWiseModelWiseSales
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'For Financial Year :'+ '" & Year(PubStartDate) & "' + '-' + '" & Year(PubEndDate) & "'"
        End Select
    Next


Case MntTarSQty  'Vijay
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("Date")
                rpt.FormulaFields(I).TEXT = " '" & Format(FGrid.TextMatrix(Date1, 1), "MMM") & "'"
        End Select
    Next
Case SalePurchDiff, ModFWiseSRep, DelayDelInt, VehSaleFormW, VehPSAudit, ModWiseFWise, AreaWiseFWise 'Vijay
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
        If GRepFormName = VehSaleFormW And UCase(rpt.FormulaFields(I).FormulaFieldName) = "TOTCAPTION" Then
            rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
        End If
    Next
Case ModWiseGWise
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("Division")
                If GridString2 = "" Then
                    rpt.FormulaFields(I).TEXT = "'Division : All'"
                Else
                    rpt.FormulaFields(I).TEXT = "'Division :'+ '" & Replace(GridString2, "'", "") & "'"
                End If
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub MntTarSQtyProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim Target As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
'    CondStr = " where Veh_Order.Inv_Date <= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and Veh_Order.Inv_Date  >= #" & Format("01/" & Format(FGrid.TextMatrix(Date1, 1), "MMM") & "/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# "

'     mQRY = "select Veh_T.MODEL, Model.Vehicle_Type," & _
'           "Veh_T.tAPR, Veh_T.tMAY, Veh_T.tJUN, Veh_T.tJUL, Veh_T.tAUG, Veh_T.tSEP, Veh_T.tOCT," & _
'           "Veh_T.tNOV, Veh_T.tDEC, Veh_T.tJAN, Veh_T.tFEB, Veh_T.tMAR," & _
'           "iif( Veh_stock.Sal_VDate  >= #" & Format("01/" & Format(FGrid.TextMatrix(Date1, 1), "MMM"), "dd/MMM/yyyy") & "# AND Veh_stock.Sal_VDate <= #" & Format(FGrid.TextMatrix(Date1, 1), "DD/MMM/YYYY") & "#,1,0) AS MonthSale " & _
'           " FROM (Model LEFT JOIN Veh_Target AS Veh_T ON Model.MODEL = Veh_T.MODEL)" & _
'           "LEFT JOIN veh_Stock ON Model.MODEL = veh_Stock.MODEL"

       mQry = "select Veh_T.MODEL, Model.Vehicle_Type," & _
       " Veh_T.TargQTY_04, Veh_T.TargQTY_05, Veh_T.TargQTY_06, Veh_T.TargQTY_07, Veh_T.TargQTY_08, Veh_T.TargQTY_09, Veh_T.TargQTY_10," & _
       " Veh_T.TargQTY_11, Veh_T.TargQTY_12, Veh_T.TargQTY_01, Veh_T.TargQTY_02, Veh_T.TargQTY_03," & _
       " " & cIIF("(Veh_stock.Pur_VDate  >= " & ConvertDate(Format("01/" & Format(FGrid.TextMatrix(Date1, 1), "MMM"), "dd/MMM/yyyy")) & " AND Veh_stock.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "DD/MMM/YYYY")) & ")", "1", "0") & " as ChassisNo " & _
       " FROM (Model LEFT JOIN Veh_Forecast AS Veh_T ON Model.MODEL = Veh_T.MODEL)" & _
       " LEFT JOIN veh_Stock ON Model.MODEL = veh_Stock.MODEL "
    
    mQry = mQry

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "MntTarSQty"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ModWiseCustmProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim Target As String
'    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

 Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
'    Condstr = ""
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.Model in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = IIf(Condstr = "", "", Condstr + " And ") & " ContractFinance.FinCode in (" & GridString2 & ")"
    
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Veh_Order.OrdDocid", "3", "1") & " in (" & GridString3 & ")"
    
    If Check1(3).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_Order.OrdDocid", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    'Condstr = IIf(Condstr = "", "", " Where " + Condstr)

    mQry = "SELECT Veh_Order.MODEL, SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, " & _
           "City.CityName, SubGroup.Phone, SubGroup.Mobile, SubGroup.FAX, ContractFinance.FinName" & _
           " FROM ((Veh_Order LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) " & _
           " LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode"

    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ModWiseCustm"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SalePurchDiffProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim Target As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

    Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Veh_Order.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_order.Inv_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Inv_DocId,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.Model in (" & GridString3 & ")"
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        mQry = "SELECT S.Name,Veh_Order.Model, " & cTrim("RIGHT(Veh_Order.Inv_DocId,13)") & " AS V_No, Veh_Order.Inv_Date, Veh_Order.VRate+Veh_Order.Margine-Veh_Order.Rebate as Net_AMOUNT,Veh_Order.VRate+Veh_Order.Margine as MRP," & _
               "Veh_stock.PBILL_NO, veh_stock.vrate,Veh_stock.PBILL_DATE,(Veh_Order.Net_Amount-Veh_Stock.vrate) AS PriceDiff" & _
               " FROM (Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.inv_docid = Veh_Stock.sal_DocId) " & _
               " Left Join SubGroup S On Veh_Order.PartyCode=S.SubCode"
    Else
        mQry = "SELECT S.Name,Veh_Order.Model, " & cTrim("RIGHT(Veh_Order.Inv_DocId,13)") & " AS V_No, Veh_Order.Inv_Date, Veh_Order.Net_AMOUNT," & _
               "Veh_stock.PBILL_NO, veh_stock.vrate,Veh_stock.PBILL_DATE,(Veh_Order.Net_Amount-Veh_Stock.vrate) AS PriceDiff" & _
               " FROM (Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.inv_docid = Veh_Stock.sal_DocId) " & _
               " Left Join SubGroup S On Veh_Order.PartyCode=S.SubCode"
    End If
    mQry = mQry + Condstr + " Order By " & cTrim("RIGHT(Veh_Order.Inv_DocId,13)") & ""

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        RepName = "SalePurchDiffJMK"
    Else
        RepName = "SalePurchDiff"
    End If
    
    
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub DelayDelIntProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim Target As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

    Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Ord_SiteCode,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(Veh_Order.Ord_SiteCode,1) ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Inv_DocId,1) in (" & GridString2 & ")"
    
    mQry = "SELECT Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.RATE, Veh_Order.Inv_No, Veh_Order.Inv_Date," & _
          "Veh_Order.Net_AMOUNT, SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, City.CityName, Veh_Order.Interest, " & _
          "Veh_Order.DelCh_DT, Veh_Order.EXP_DATE,( Veh_Order.EXP_DATE-Veh_Order.DelCh_DT) AS Days ,Veh_Order.REBATE, Veh_Order.InterestPer" & _
          " FROM (Veh_Order LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) " & _
          " LEFT JOIN City ON SubGroup.CityCode = City.CityCode"


    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "DelayDelInt"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub AreaWiseFWiseProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim Target As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
     
     Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Veh_Order.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_Order.Inv_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Inv_DocId,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.area in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.fb_code in (" & GridString4 & ")"
    
    mQry = "SELECT Veh_Order.Inv_DocId, Area.AreaName, ContractFinance.FinName" & _
           " FROM (Veh_Order LEFT JOIN Area ON Veh_Order.AREA = Area.AreaCode) " & _
           "LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode"

    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "AreaWiseFWise"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VehSaleFormWProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim Target As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

     Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Veh_Order.Inv_DocID", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_Order.Inv_DocID", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Inv_DocID,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.Form_Code in (" & GridString3 & ")"

    mQry = "SELECT TaxForms.Form_Desc,Veh_Order.Form_Code, " & _
           "(Veh_Order.FIT_AMT+Veh_Order.MARGINE+Veh_Order.VRATE-Veh_Order.REBATE+Veh_Order.InciChrg+Veh_Order.Octroi+Veh_Order.RegTemp+Veh_Order.TransitInsu+Veh_Order.Transport+Veh_Order.MVT) AS Amount," & _
           "Veh_Order.Surcharge_Per, Veh_Order.TAX_Per," & _
           "Veh_Order.TAX_Amt,Veh_Order.Surcharge_Amt,Veh_Order.TOT_Amt, " & _
           "Veh_Order.Net_AMOUNT,Veh_Order.TOT_Per" & _
           " FROM Veh_Order LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code "
    mQry = mQry + Condstr
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "VehSaleFormW"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VehPSAuditProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
     
     Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Ord_SiteCode,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(Veh_Order.Ord_SiteCode,1) ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_Purch1.BMS_CATEGORY in (" & GridString2 & ")"
    
    mQry = "SELECT Veh_Purch1.V_Type AS PurType, Veh_Purch1.V_NO As PurNo, Veh_Purch1.V_Date AS PurDate," & _
           "Veh_Purch1.PBILL_NO, Veh_Purch1.PBILL_DATE," & _
           "Veh_Purch1.OBNO AS DlrNo, Veh_Purch1.OBDate as  DlrDate,Veh_Order.Chassis, Veh_Order.MODEL," & _
           "Veh_Purch1.Tot_AMOUNT, Veh_Purch1.P_Amount, Veh_Order.Inv_No,Veh_Order.Inv_Date, Veh_Order.TAX_Amt," & _
           "Veh_Order.Surcharge_Amt,  veh_Rate.margine as StdMargine, Veh_Order.MARGINE as ChgMargine," & _
           "Veh_Order.Transport, Veh_Order.FIT_AMT, Veh_Order.Net_AMOUNT AS NetSaleAmt" & _
           " FROM ((Veh_Stock LEFT JOIN Veh_Purch1 ON Veh_Stock.Pur_Docid = Veh_Purch1.DocID )" & _
           "Left Join Veh_order On Veh_Stock.Sal_Docid=Veh_Order.Inv_DocId)" & _
           "Left Join Veh_Rate ON  Veh_Rate.Model=Veh_Stock.Model"


    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "VehPSAudit"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub eMailSaleProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Dealer As String, constr1$, MyYesNo$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
     
     Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then constr1 = " and left(Veh_Order.Ord_SiteCode,1) in (" & GridString1 & ")"
     If constr1 <> "" Then
     MyYesNo = "Y"
     Else
     MyYesNo = "N"
    End If
    Dealer = GCn.Execute("select Dealer_Id from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    
    mQry = "SELECT Veh_Order.Chassis, Veh_Order.Inv_No, Veh_Order.Inv_Date," & _
           " '" & Dealer & " ' AS DealerI ," & _
           " '" & MyYesNo & " ' AS YNo " & _
           " FROM Veh_Order"


    mQry = mQry + Condstr + constr1

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "eMailSale"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ModWiseGWiseProc()
On Error GoTo ELoop
Dim mQry$, Condstr$, mQRY1$, CondStr1$
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
'    CondStr = " where VS.Pur_VDate  >= #" & Format(PubStartDate, "dd/MMM/yyyy") & "# and VS.Pur_VDate <= #" & Format(PubEndDate, "dd/MMM/yyyy") & "# "
'
'    If Check1(1).Value = Unchecked Then CondStr = CondStr & " And mid(VS.Pur_DocId,3,1) in (" & GridString1 & ")"
'    If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(Pur_DocId,1) in (" & GridString2 & ")"
'    If Check1(3).Value = Unchecked Then CondStr = CondStr & " and VS.Model in (" & GridString3 & ")"
'    If FGrid.TextMatrix(List1, 1) = "GroupWise" Then
'        If Check1(4).Value = Unchecked Then CondStr = CondStr & " and M.Vehicle_Type in (" & GridString4 & ")"
'    Else
'        If Check1(4).Value = Unchecked Then CondStr = CondStr & " and M.Grp_Code in (" & GridString4 & ")"
'    End If
'
' mQRY = "SELECT M.Vehicle_Type, M.Model,MG.ModelGrp_Name,  " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Apr/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Mar/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VType='V_OST',1,0) AS AprOpenPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Apr/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Apr/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VType='V_PB',1,0) AS AprPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "May/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "May/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS MayPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Jun/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Jun/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS JunPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Jul/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Jul/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS JulPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Aug/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Aug/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS AugPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Sep/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Sep/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS SepPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Oct/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Oct/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS octPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Nov/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Nov/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS NovPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Dec/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Dec/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS DecPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Jan/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Jan/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS JanPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format(fxLastDay(Format("27/" & "Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & "/Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS FebPurch, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Mar/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Mar/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND Pur_VType='V_PB',1,0) AS MarPurch, "
'
'mQRY1 = "iif( VS.Pur_VDate  >= #" & Format("01/" & "Apr/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Apr/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS AprOpenSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "May/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "May/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS MaySale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Jun/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Jun/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS JunSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Jul/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Jul/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS JulSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Aug/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Aug/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS AugSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Sep/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Sep/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS SepSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Oct/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Oct/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS octSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Nov/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("30/" & "Nov/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS NovSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Dec/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Dec/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS DecSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Jan/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Jan/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS JanSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format(fxLastDay(Format("27/" & "Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & "/Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS FebSale, " & _
'        "iif( VS.Pur_VDate  >= #" & Format("01/" & "Mar/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VS.Pur_VDate <= #" & Format("31/" & "Mar/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy") & "# AND VO.inv_VType='V_SB',1,0) AS MarSale " & _
'        " FROM ((Veh_Stock VS LEFT JOIN Model M ON VS.MODEL = M.MODEL) Left join Model_Grp MG on MG.ModelGrp_code =M.Grp_Code ) LEFT JOIN Veh_Order VO On VO.Inv_Docid=VS.Sal_Docid "

'mQRY = mQRY + mQRY1 + CondStr

    Condstr = " where (VS.Pur_VDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VS.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ") "
    CondStr1 = " where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then
        Condstr = Condstr & " And " & cMID("VS.Pur_DocId", "3", "1") & " in (" & GridString1 & ")"
        CondStr1 = CondStr1 & " And " & cMID("VO.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
        Else
         If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VS.Pur_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then CondStr1 = CondStr1 & " and " & cMID("VO.Inv_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
          
    End If
    If Check1(2).Value = Unchecked Then
        Condstr = Condstr & " and left(Pur_DocId,1) in (" & GridString2 & ")"
        CondStr1 = CondStr1 & " and left(Inv_DocId,1) in (" & GridString2 & ")"
    End If
    If Check1(3).Value = Unchecked Then
        Condstr = Condstr & " and VS.Model in (" & GridString3 & ")"
        CondStr1 = CondStr1 & " and VO.Model in (" & GridString3 & ")"
    End If
    If FGrid.TextMatrix(List1, 1) = "GroupWise" Then
        If Check1(4).Value = Unchecked Then
            Condstr = Condstr & " and M.Vehicle_Type in (" & GridString4 & ")"
            CondStr1 = CondStr1 & " and M.Vehicle_Type in (" & GridString4 & ")"
        End If
    Else
        If Check1(4).Value = Unchecked Then
            Condstr = Condstr & " and M.Grp_Code in (" & GridString4 & ")"
            CondStr1 = CondStr1 & " and M.Grp_Code in (" & GridString4 & ")"
        End If
    End If
    If PubBackEnd = "A" Then
        mQry = "SELECT M.Vehicle_Type, M.Model,MG.ModelGrp_Name,  " & _
            "" & cIIF("VS.Pur_VType='V_OST'", "1", "0") & " AS AprOpenPurch, " & _
            "" & cIIF("format(VS.Pur_VDate,'MM')='04' AND VS.Pur_VType='V_PB'", "1", "0") & " AS AprPurch, " & _
            "" & cIIF("format(VS.Pur_VDate,'MM')='05' AND Pur_VType='V_PB'", "1", "0") & " AS MayPurch, " & _
            "" & cIIF("format(VS.Pur_VDate,'MM')='06' AND Pur_VType='V_PB'", "1", "0") & " AS JunPurch, " & _
            "" & cIIF("format(VS.Pur_VDate,'MM')='07' AND Pur_VType='V_PB'", "1", "0") & " AS JulPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='08' AND Pur_VType='V_PB'", "1", "0") & " AS AugPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='09' AND Pur_VType='V_PB'", "1", "0") & " AS SepPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='10' AND Pur_VType='V_PB'", "1", "0") & " AS OctPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='11' AND Pur_VType='V_PB'", "1", "0") & " AS NovPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='12' AND Pur_VType='V_PB'", "1", "0") & " AS DecPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='01' AND Pur_VType='V_PB'", "1", "0") & " AS JanPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='02' AND Pur_VType='V_PB'", "1", "0") & " AS FebPurch, " & _
            "" & cIIF("format( VS.Pur_VDate,'MM')='03' AND Pur_VType='V_PB'", "1", "0") & " AS MarPurch, " & _
            " 0 AS AprOpenSale, 0 AS MaySale, 0 AS JunSale, " & _
            " 0 AS JulSale, 0 AS AugSale, 0 AS SepSale, 0 AS OctSale, " & _
            " 0 AS NovSale, 0 AS DecSale, 0 AS JanSale, 0 AS FebSale,0 AS MarSale " & _
            " FROM ((Veh_Stock VS LEFT JOIN Model M ON VS.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on MG.ModelGrp_code =M.Grp_Code ) "
        
        mQRY1 = "SELECT M.Vehicle_Type, M.Model,MG.ModelGrp_Name,  " & _
                "0 AS AprOpenPurch, 0 AS AprPurch,0 AS MayPurch,0 AS JunPurch, " & _
                "0 AS JulPurch, 0 AS AugPurch, 0 AS SepPurch, 0 AS OctPurch, " & _
                "0 AS NovPurch, 0 AS DecPurch, 0 AS JanPurch, 0 AS FebPurch, 0 AS MarPurch, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='04' AND VO.inv_VType='V_SB'", "1", "0") & " AS AprOpenSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='05' AND VO.inv_VType='V_SB'", "1", "0") & " AS MaySale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='06' AND VO.inv_VType='V_SB'", "1", "0") & " AS JunSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='07' AND VO.inv_VType='V_SB'", "1", "0") & " AS JulSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='08' AND VO.inv_VType='V_SB'", "1", "0") & " AS AugSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='09' AND VO.inv_VType='V_SB'", "1", "0") & " AS SepSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='10' AND VO.inv_VType='V_SB'", "1", "0") & " AS OctSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='11' AND VO.inv_VType='V_SB'", "1", "0") & " AS NovSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='12' AND VO.inv_VType='V_SB'", "1", "0") & " AS DecSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='01' AND VO.inv_VType='V_SB'", "1", "0") & " AS JanSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='02' AND VO.inv_VType='V_SB'", "1", "0") & " AS FebSale, " & _
                "" & cIIF(" format(VO.Inv_Date,'MM')='03' AND VO.inv_VType='V_SB'", "1", "0") & " AS MarSale " & _
                " FROM ((Veh_Order VO LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
                " Left join Model_Grp MG on MG.ModelGrp_code =M.Grp_Code ) "
    ElseIf PubBackEnd = "S" Then
        mQry = "SELECT M.Vehicle_Type, M.Model,MG.ModelGrp_Name,  " & _
            "" & cIIF("VS.Pur_VType='V_OST'", "1", "0") & " AS AprOpenPurch, " & _
            "" & cIIF("Month(Vs.Pur_VDate)='4' AND VS.Pur_VType='V_PB'", "1", "0") & " AS AprPurch, " & _
            "" & cIIF("Month(VS.Pur_VDate)='5' AND Pur_VType='V_PB'", "1", "0") & " AS MayPurch, " & _
            "" & cIIF("Month(VS.Pur_VDate)='6' AND Pur_VType='V_PB'", "1", "0") & " AS JunPurch, " & _
            "" & cIIF("Month(VS.Pur_VDate)='7' AND Pur_VType='V_PB'", "1", "0") & " AS JulPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='8' AND Pur_VType='V_PB'", "1", "0") & " AS AugPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='9' AND Pur_VType='V_PB'", "1", "0") & " AS SepPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='10' AND Pur_VType='V_PB'", "1", "0") & " AS OctPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='11' AND Pur_VType='V_PB'", "1", "0") & " AS NovPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='12' AND Pur_VType='V_PB'", "1", "0") & " AS DecPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='1' AND Pur_VType='V_PB'", "1", "0") & " AS JanPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='2' AND Pur_VType='V_PB'", "1", "0") & " AS FebPurch, " & _
            "" & cIIF("Month( VS.Pur_VDate)='3' AND Pur_VType='V_PB'", "1", "0") & " AS MarPurch, " & _
            " 0 AS AprOpenSale, 0 AS MaySale, 0 AS JunSale, " & _
            " 0 AS JulSale, 0 AS AugSale, 0 AS SepSale, 0 AS OctSale, " & _
            " 0 AS NovSale, 0 AS DecSale, 0 AS JanSale, 0 AS FebSale,0 AS MarSale " & _
            " FROM ((Veh_Stock VS LEFT JOIN Model M ON VS.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on MG.ModelGrp_code =M.Grp_Code ) "
        
        mQRY1 = "SELECT M.Vehicle_Type, M.Model,MG.ModelGrp_Name,  " & _
                "0 AS AprOpenPurch, 0 AS AprPurch,0 AS MayPurch,0 AS JunPurch, " & _
                "0 AS JulPurch, 0 AS AugPurch, 0 AS SepPurch, 0 AS OctPurch, " & _
                "0 AS NovPurch, 0 AS DecPurch, 0 AS JanPurch, 0 AS FebPurch, 0 AS MarPurch, " & _
                "" & cIIF(" Month(VO.Inv_Date)='4' AND VO.inv_VType='V_SB'", "1", "0") & " AS AprOpenSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='5' AND VO.inv_VType='V_SB'", "1", "0") & " AS MaySale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='6' AND VO.inv_VType='V_SB'", "1", "0") & " AS JunSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='7' AND VO.inv_VType='V_SB'", "1", "0") & " AS JulSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='8' AND VO.inv_VType='V_SB'", "1", "0") & " AS AugSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='9' AND VO.inv_VType='V_SB'", "1", "0") & " AS SepSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='10' AND VO.inv_VType='V_SB'", "1", "0") & " AS OctSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='11' AND VO.inv_VType='V_SB'", "1", "0") & " AS NovSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='12' AND VO.inv_VType='V_SB'", "1", "0") & " AS DecSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='1' AND VO.inv_VType='V_SB'", "1", "0") & " AS JanSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='2' AND VO.inv_VType='V_SB'", "1", "0") & " AS FebSale, " & _
                "" & cIIF(" Month(VO.Inv_Date)='3' AND VO.inv_VType='V_SB'", "1", "0") & " AS MarSale " & _
                " FROM ((Veh_Order VO LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
                " Left join Model_Grp MG on MG.ModelGrp_code =M.Grp_Code ) "
    End If

mQry = mQry & Condstr & " Union all " & mQRY1 & CondStr1

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ModWiseGWise"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub ModWiseFWiseProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, mQRY1 As String
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub

    Condstr = " where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Inv_Docid", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VO.Inv_Docid", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_Docid,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and  VO.Model in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VO.FB_CODE in (" & GridString4 & ")"

' mQRY = "SELECT M.Vehicle_Type, M.Model,ContractFinance.FinName , " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Apr/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("30/" & "Apr/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS AprSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "May/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("31/" & "May/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS MaySale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Jun/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("30/" & "Jun/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS JunSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Jul/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("31/" & "Jul/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS JulSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Aug/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("31/" & "Aug/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS AugSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Sep/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("30/" & "Sep/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS SepSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Oct/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("31/" & "Oct/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS octSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Nov/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("30/" & "Nov/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS NovSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Dec/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("31/" & "Dec/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS DecSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Jan/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("31/" & "Jan/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS JanSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Feb/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format(fxLastDay(Format("27/" & "Feb/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy")) & "/Feb/" & Format(Date2, "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS FebSale, " & _
'        "iif( VO.Inv_Date  >= #" & Format("01/" & "Mar/" & Format(FGrid.TextMatrix(Date1, 1), "yyyy"), "dd/MMM/yyyy") & "# AND VO.Inv_Date <= #" & Format("31/" & "Mar/" & Format(FGrid.TextMatrix(Date2, 1), "yyyy"), "dd/MMM/yyyy") & "# ,1,0) AS MarSale " & _
'        " FROM (Veh_Order VO LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
'        "LEFT JOIN ContractFinance ON ContractFinance.FinCode=VO.FB_Code"
    If PubBackEnd = "A" Then
        mQry = "SELECT M.Vehicle_Type, " & cIIF("'" & UCase(left(PubComp_Name, 4)) & "'='ENAR'", "M.Sales_Desc", "M.Model") & " as Model,ContractFinance.FinName , " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '04'", "1", "0") & " AS AprSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '05'", "1", "0") & " AS MaySale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '06'", "1", "0") & " AS JunSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '07'", "1", "0") & " AS JulSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '08'", "1", "0") & " AS AugSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '09'", "1", "0") & " AS SepSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '10'", "1", "0") & " AS OctSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '11'", "1", "0") & " AS NovSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '12'", "1", "0") & " AS DecSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '01'", "1", "0") & " AS JanSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '02'", "1", "0") & " AS FebSale, " & _
            "" & cIIF(" format(VO.Inv_Date,'MM')= '03'", "1", "0") & " AS MarSale " & _
            " FROM (Veh_Order VO LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
            "LEFT JOIN ContractFinance ON ContractFinance.FinCode=VO.FB_Code"
    ElseIf PubBackEnd = "S" Then
        mQry = "SELECT M.Vehicle_Type, " & cIIF("'" & UCase(left(PubComp_Name, 4)) & "'='ENAR'", "M.Sales_Desc", "M.Model") & " as Model,ContractFinance.FinName , " & _
            "" & cIIF(" Month(VO.Inv_Date)= '4'", "1", "0") & " AS AprSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '5'", "1", "0") & " AS MaySale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '6'", "1", "0") & " AS JunSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '7'", "1", "0") & " AS JulSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '8'", "1", "0") & " AS AugSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '9'", "1", "0") & " AS SepSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '10'", "1", "0") & " AS OctSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '11'", "1", "0") & " AS NovSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '12'", "1", "0") & " AS DecSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '1'", "1", "0") & " AS JanSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '2'", "1", "0") & " AS FebSale, " & _
            "" & cIIF(" Month(VO.Inv_Date)= '3'", "1", "0") & " AS MarSale " & _
            " ,M.Model_Desc FROM (Veh_Order VO LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
            "LEFT JOIN ContractFinance ON ContractFinance.FinCode=VO.FB_Code"
    End If
    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ModWiseFWise"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub ModFWiseMicroProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, mQRY1 As String
    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
'    CondStr = " where VO.Inv_Date  >= #" & Format(FGrid.TextMatrix(Cat1, 1), "dd/MMM/yyyy") & "# and VO.Inv_Date <= #" & Format(FGrid.TextMatrix(Cat2, 1), "dd/MMM/yyyy") & "#"
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " where left(VO.Ord_SiteCode,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(VO.Ord_SiteCode,1) ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and VO.Model in (" & GridString2 & ")"

 mQry = "SELECT M.Vehicle_Type, M.Model,ContractFinance.FinName , FGS.FIN_YR, FGS.QTY_T04, FGS.QTY_T05, FGS.QTY_T06, FGS.QTY_T07, " & _
        "FGS.QTY_T08, FGS.QTY_T09, FGS.QTY_T10, FGS.QTY_T11, FGS.QTY_T12, FGS.QTY_T01, FGS.QTY_T02, FGS.QTY_T03, FGS.QTY_S04, FGS.QTY_S05," & _
        "FGS.QTY_S06, FGS.QTY_S07, FGS.QTY_S08, FGS.QTY_S09, FGS.QTY_S10, FGS.QTY_S11, FGS.QTY_S12, FGS.QTY_S01, FGS.QTY_S02, FGS.QTY_S03," & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Apr/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Apr/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS AprSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "May/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "May/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS MaySale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Jun/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Jun/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS JunSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Jul/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Jul/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & "", "1", "0") & " AS JulSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Aug/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Aug/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS AugSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Sep/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Sep/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS SepSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Oct/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Oct/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS octSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Nov/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Nov/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS NovSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Dec/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Dec/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS DecSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Jan/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Jan/" & (FGrid.TextMatrix(Cat2, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS JanSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Feb/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format(fxLastDay(Format("27/" & "Feb/" & (FGrid.TextMatrix(Cat2, 1)), "dd/MMM/yyyy")) & "/Feb/" & (FGrid.TextMatrix(Cat2, 1)), "dd/MMM/yyyy")) & " ", "1", "0") & " AS FebSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Mar/" & (FGrid.TextMatrix(Cat1, 1)), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Mar/" & (FGrid.TextMatrix(Cat2, 1)), "dd/MMM/yyyy")) & "", "1", "0") & " AS MarSale " & _
        " FROM ((Veh_Order VO LEFT JOIN Model M ON VO.MODEL = M.MODEL)" & _
        " LEFT JOIN ContractFinance ON ContractFinance.FinCode=VO.FB_Code)" & _
        " LEFT JOIN FinGroupSummary AS FGS ON FGS.FinGrpCode=ContractFinance.UnderFinGrp"

mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ModFWiseMicro"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub ArWiseModWYrProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, mQRY1 As String
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where VO.Inv_Date  >= " & ConvertDate(Format(PubStartDate, "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(PubEndDate, "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Inv_docid", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VO.Inv_docid", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_docid,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.Model in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VO.AREA in (" & GridString4 & ")"

 mQry = "SELECT M.Vehicle_Type, M.Model,Area.AreaName , " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Apr/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Apr/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " ", "1", "0") & " AS AprSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "May/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "May/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS MaySale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Jun/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Jun/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS JunSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Jul/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Jul/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS JulSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Aug/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Aug/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS AugSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Sep/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Sep/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS SepSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Oct/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Oct/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS octSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Nov/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("30/" & "Nov/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS NovSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Dec/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Dec/" & Format((PubStartDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS DecSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Jan/" & Format((PubEndDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Jan/" & Format((PubEndDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS JanSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Feb/" & Format((PubEndDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format(fxLastDay(Format("27/" & "Feb/" & Format((PubEndDate), "yyyy"), "dd/MMM/yyyy")) & "/Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS FebSale, " & _
        "" & cIIF(" VO.Inv_Date  >= " & ConvertDate(Format("01/" & "Mar/" & Format((PubEndDate), "yyyy"), "dd/MMM/yyyy")) & " AND VO.Inv_Date <= " & ConvertDate(Format("31/" & "Mar/" & Format((PubEndDate), "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS MarSale " & _
        " FROM (Veh_Order VO LEFT JOIN Model M ON VO.MODEL = M.MODEL) " & _
        "LEFT JOIN Area ON Area.AreaCode=VO.Area "
    

mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ArWiseModWYr"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub ModBrWiseStkProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

    Condstr = " AND VS.Pur_Vdate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VS.Pur_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VS.Pur_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(VS.Pur_DocId,1) in (" & GridString2 & ")"

    mQry = "Select M.MODEL, M.Vehicle_Type, Site.Site_Desc, VS.Pur_DocId,VS.RATE" & _
           " FROM (Veh_Stock AS VS LEFT JOIN Model AS M ON VS.MODEL = M.MODEL) " & _
           " Left JOIN Site ON M.Site_Code = Site.Site_Code " & _
           " WHERE (VS.Sal_VDate Is Null OR VS.Sal_VDate  > " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
'           " AND (IsNull(VS.DelCh_Date) or VS.DelCh_Date  > #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# ) "

    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ModBrWiseStk"
    RepTitle = UCase("Model/BranchWise Stock Reports")
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ModBrWiseSDProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim mGrpField As String
Dim mDateField As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

    If StrCmp(FGrid.TextMatrix(List2, 1), "DELIVERY DATE") Then
        mDateField = "Veh_Order.DelCh_DT"
        Condstr = " Veh_Order.DelCh_DT  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.DelCh_DT <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    Else
        mDateField = "Veh_Order.Inv_Date"
        Condstr = " Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    End If
      
    
    If StrCmp(FGrid.TextMatrix(List1, 1), "Model") Then '
        mGrpField = "Model.Model"
    Else
        mGrpField = "MG.ModelGrp_Name"
    End If
    
    
    If StrCmp(FGrid.TextMatrix(List2, 1), "DELIVERY DATE") Then
        If Check1(1).Value = Unchecked Then
            Condstr = Condstr & " And " & cMID("Veh_Order.DelCh_DocId", "3", "1") & " in (" & GridString1 & ")"
        Else
             If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_Order.DelCh_DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
        
        If Check1(2).Value = Unchecked Then
            Condstr = Condstr & " and left(Veh_Order.DelCh_DocId,1) in (" & GridString2 & ")"
        End If
        
        mQry = "SELECT " & _
                    " Veh_Stock.Pur_DocId,Veh_Stock.Sal_DocId, Veh_Order.Inv_DocId, " & _
                    " Model.Model_Desc,Veh_Order.DelCh_DT, " & mGrpField & " as Model,Model.Vehicle_Type,Site.ShortName as Site_Desc " & _
                " FROM ((((Veh_Stock " & _
                    "INNER JOIN Model ON Veh_Stock.MODEL = Model.MODEL) " & _
                    "LEFT JOIN Model_Grp MG ON Model.Grp_Code =MG.ModelGrp_Code)  " & _
                    "INNER JOIN Veh_Order ON Veh_Stock.Sal_DocId = Veh_Order.Inv_DocId) " & _
                    "Inner Join Site on " & cMID("Veh_Order.DelCh_DocId", "3", "1") & "=Site.Site_Code) " & _
                    "where "
        mQry = mQry + Condstr
    Else
    
        If Check1(1).Value = Unchecked Then
            Condstr = Condstr & " And " & cMID("Veh_Order.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
        Else
            If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_Order.Inv_DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
        
        If Check1(2).Value = Unchecked Then
            Condstr = Condstr & " and left(Veh_Order.Inv_DocId,1) in (" & GridString2 & ")"
        End If
    
    
        mQry = "SELECT " & _
                    " Veh_Stock.Pur_DocId,Veh_Stock.Sal_DocId, Veh_Order.Inv_DocId, " & _
                    " Model.Model_Desc,Veh_Order.Inv_Date as DelCh_DT, " & mGrpField & " as Model, Model.Vehicle_Type, Site.ShortName as Site_Desc " & _
                " FROM ((((Veh_Stock " & _
                    "INNER JOIN Model ON Veh_Stock.MODEL = Model.MODEL) " & _
                    "LEFT JOIN Model_Grp MG ON Model.Grp_Code =MG.ModelGrp_Code)  " & _
                    "INNER JOIN Veh_Order ON Veh_Stock.Sal_DocId = Veh_Order.Inv_DocId) " & _
                    "Inner Join Site on " & cMID("Veh_Order.Inv_DocId", "3", "1") & "=Site.Site_Code) " & _
                    "where "
            mQry = mQry + Condstr
    End If
    

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ModBrWiseSD"
    RepTitle = UCase(Me.CAPTION)
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ModFWiseSRepProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub

    Condstr = " WHERE VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " AND " & cMID("VO.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VO.Inv_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and LEFT(VO.Inv_DocId,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.Model in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and ContractFinance.UnderFinGrp in (" & GridString4 & ")"

    mQry = "Select M.MODEL, M.Vehicle_Type, VO.OrdDocId, ContractFinance.FinName,ContractFinance.UnderFinGrp,M.Model_Desc  " & _
           " FROM (Veh_Order AS VO RIGHT JOIN Model AS M ON VO.MODEL = M.MODEL) " & _
           " Left JOIN ContractFinance ON VO.FB_Code = ContractFinance.FinCode  "
    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "ModFWiseSRep"
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
Private Sub MonthlySummProc()
On Error GoTo ELoop
Dim mOpQry$, mQry$, mQRY1$, mQRY2$, Condstr$, ChasDivCond$, PurDivCond$, SalDivCond$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If Check1(3).Value = Unchecked Then
        ChasDivCond = " Chassis_RctDivCode in (" & GridString3 & ") and "
        PurDivCond = " left(VStk.Pur_DocId,1) in (" & GridString3 & ") and "
        SalDivCond = " left(VO.Inv_DocId,1) in (" & GridString3 & ") and "
    End If
    If PubBackEnd = "A" Then
        mOpQry = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,1 as DayOpen, 0 AS MonthPur, 0 as PurDay, 0 AS MonthSal, 0 as SalDay, '' as AreaName,0 as TranQty " & _
            " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
            " where (VStk.Pur_DocId='' and " & ChasDivCond & " Chassis_RctDate<" & ConvertDate("01/" & Right(FGrid.TextMatrix(Date1, 1), 8)) & ") " & _
            " or (" & PurDivCond & " VStk.Pur_VDate < " & ConvertDate("01/" & Right(FGrid.TextMatrix(Date1, 1), 8)) & ") " & _
            " and (VStk.Sal_VDate Is Null or VStk.Sal_VDate< " & ConvertDate("01/" & Right(FGrid.TextMatrix(Date1, 1), 8)) & " )"
        
        mQry = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,0 as DayOpen, 1 AS MonthPur, " & _
            " " & cIIF("VStk.Pur_VDate = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS PurDay, " & _
            " 0 as MonthSale, 0 AS SalDay, '' as AreaName,0 as TranQty " & _
            " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
            " where (VStk.Pur_DocId='' and " & ChasDivCond & " Format(Chassis_RctDate,'YYYYMM')='" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & "' and Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") and  VStk.Indate is not Null " & _
            " or (" & PurDivCond & " format(VStk.Pur_VDate,'YYYYMM')= '" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & "' and VStk.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    
        mQRY1 = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VO.Model,0 as DayOpen, 0 as MonthPur, 0 as PurDay, 1 AS MonthSale, " & _
            "" & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS SalDay, Area.AreaName,0 as TranQty " & _
            " FROM (((Veh_Order VO Left Join Model M on VO.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Area on VO.Area=Area.AreaCode " & _
            " where " & SalDivCond & _
            " format(VO.Inv_Date,'YYYYMM')= '" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & _
            "' and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ""
            
            
        mQRY2 = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,0 as DayOpen, 0 AS MonthPur, " & _
            " 0 AS PurDay, " & _
            " 0 as MonthSale, 0 AS SalDay, '' as AreaName,1 as TranQty " & _
            " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
            " where (VStk.Pur_DocId='' and " & ChasDivCond & " Format(Chassis_RctDate,'YYYYMM')='" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & "' and Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") and VStk.Indate is Null "
    ElseIf PubBackEnd = "S" Then
        mOpQry = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,1 as DayOpen, 0 AS MonthPur, 0 as PurDay, 0 AS MonthSal, 0 as SalDay, '' as AreaName,0 as TranQty " & _
            " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
            " where (VStk.Pur_DocId='' and " & ChasDivCond & " Chassis_RctDate<" & ConvertDate("01/" & Right(FGrid.TextMatrix(Date1, 1), 8)) & ") " & _
            " or (" & PurDivCond & " VStk.Pur_VDate < " & ConvertDate("01/" & Right(FGrid.TextMatrix(Date1, 1), 8)) & ") " & _
            " and (VStk.Sal_VDate Is Null or VStk.Sal_VDate< " & ConvertDate("01/" & Right(FGrid.TextMatrix(Date1, 1), 8)) & " )"
        
        mQry = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,0 as DayOpen, 1 AS MonthPur, " & _
            " " & cIIF("VStk.Pur_VDate = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS PurDay, " & _
            " 0 as MonthSale, 0 AS SalDay, '' as AreaName,0 as TranQty " & _
            " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
            " Where (VStk.Pur_DocId='' and " & ChasDivCond & " " & cCStr("Month(Chassis_RctDate)") & " + " & cCStr("Year(Chassis_RctDate)") & " = '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "' and Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") and  VStk.Indate is not Null " & _
            " or (" & PurDivCond & " " & cCStr("Month(VStk.Pur_VDate)") & " + " & cCStr("Year(VStk.Pur_VDate)") & "= '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "' and VStk.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    
        mQRY1 = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VO.Model,0 as DayOpen, 0 as MonthPur, 0 as PurDay, 1 AS MonthSale, " & _
            "" & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS SalDay, Area.AreaName,0 as TranQty " & _
            " FROM (((Veh_Order VO Left Join Model M on VO.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Area on VO.Area=Area.AreaCode " & _
            " Where " & SalDivCond & _
            " " & cCStr("Month(VO.Inv_Date)") & " + " & cCStr("Year(VO.Inv_Date)") & " = '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & _
            "' And VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ""
            
            
        mQRY2 = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,0 as DayOpen, 0 AS MonthPur, " & _
            " 0 AS PurDay, " & _
            " 0 as MonthSale, 0 AS SalDay, '' as AreaName,1 as TranQty " & _
            " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
            " Where (VStk.Pur_DocId='' and " & ChasDivCond & " " & cCStr("Month(Chassis_RctDate)") & " + " & cCStr("Year(Chassis_RctDate)") & " ='" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "' and Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") and VStk.Indate is Null "
    End If
    
    mQry = mOpQry & " Union all " & mQry & " Union all " & mQRY1 & " Union all " & mQRY2

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

    
    RepName = "MonthlySumm"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub MonthWiseModelWiseSalesProc()
On Error GoTo ELoop
Dim mQry As String, Condstr$, CondStr1$
Dim I As Integer
Dim mGroupField$
Dim mDate As Date
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    
    If FGrid.TextMatrix(List1, 1) = "Taxable" Then Condstr = " And VStk.Tax_YN = 1 "
    If FGrid.TextMatrix(List1, 1) = "Taxpaid" Then Condstr = " And VStk.Tax_YN = 0 "
    
    
    
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VStk.Pur_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VStk.Pur_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(1).Value = Unchecked Then CondStr1 = CondStr1 & " and " & cMID("VStk.Sal_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then CondStr1 = CondStr1 & " and " & cMID("VStk.Sal_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    
'
'    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_DocId,1) in (" & GridString2 & ")"
'    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.PartyCode in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VStk.Model in (" & GridString4 & ")"
        
        If FGrid.TextMatrix(List2, 1) = "Model" Then
            mGroupField = "VStk.Model"
        Else
            mGroupField = "ModelGrp_Name"
        End If
                
        mDate = CDate("01/Apr/" & Year(PubStartDate))
        I = 1
        mQry = " Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 1 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar  " & _
            "From (((Veh_Stock VStk  " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where ((VStk.Pur_VDate < " & ConvertDate(mDate) & " " & _
            "And (Sal_VDate >= " & ConvertDate(PubStartDate) & " Or Sal_VDate Is NULL ))) " & CondStr1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 1 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(mDate) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 1 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(mDate) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr



        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 1 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 1 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr



        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 1 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 1 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr

        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 1 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 1 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr



        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 1 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 1 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr


        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 1 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 1 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr




        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 1 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 1 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr



        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 1 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 1 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr



        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 1 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 1 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr



        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 1 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 1 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr


        I = I + 1
            
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 1 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 1 As SaleFeb, 0 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr


        I = I + 1
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 1 As PurchaseMar, 0 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Pur_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Pur_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " " & CondStr1
            
        mQry = mQry & " Union All Select " & mGroupField & " As GroupField, MC.ModelCat_Name, MG.ModelGrp_Name, VStk.Model, 0 as Opening, 0 As PurchaseApr, 0 As SaleApr, 0 As PurchaseMay, 0 As SaleMay, 0 As PurchaseJun, 0 As SaleJun, 0 As PurchaseJul, 0 As SaleJul, 0 As PurchaseAug, 0 As SaleAug, 0 As PurchaseSep, 0 As SaleSep, 0 As PurchaseOct, 0 As SaleOct, 0 As PurchaseNov, 0 As SaleNov, 0 As PurchaseDec, 0 As SaleDec, 0 as PurchaseJan, 0 as NetSaleJan, 0 As PurchaseFeb, 0 As SaleFeb, 0 As PurchaseMar, 1 As SaleMar " & _
            "From (((Veh_Stock VStk " & _
            "Left Join Model M On M.Model=VStk.Model) " & _
            "Left Join Model_Cat MC On M.Cat_Code=MC.ModelCat_Code)" & _
            "Left Join Model_Grp MG On M.Grp_Code=MG.ModelGrp_Code)" & _
            "Where VStk.Sal_VDate >= " & ConvertDate(DateAdd("M", I - 1, mDate)) & " " & _
            "And VStk.Sal_VDate <= " & ConvertDate(DateAdd("D", -1, DateAdd("M", I, mDate))) & " And VStk.Sal_DocId<>'' " & Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    Set RstRep1 = New ADODB.Recordset
    'RstRep1.Open "Select Sum(VRATE+Margine) as Cancel_Amt,sum(Tax_Amt) as CancelTax_Amt,sum(TOT_Amt) as CancelTOT_Amt from Veh_Order1 as VO1 Where VO1.Ord_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and VO1.Ord_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# ", GCn, adOpenStatic, adLockReadOnly
    ' For Speed print of report
'    If SpeedPrnVehSale = True And FGrid.TextMatrix(List1, 1) = "Summary" Then
'        SpeedPrintSumm
'        Exit Sub
'    ElseIf SpeedPrnVehSale = True And FGrid.TextMatrix(List1, 1) = "Detailed" Then
'        SpeedPrintDet
'         Exit Sub
'    End If
    ' End Print
    
    RepName = "MonthWiseModelWiseSales"
    
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

