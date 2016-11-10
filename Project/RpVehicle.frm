VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form RpVehicle 
   BackColor       =   &H00C8E8DA&
   Caption         =   "RpVehicle"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   1800
      TabIndex        =   17
      Top             =   6540
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DGHelp1 
      Height          =   2745
      Left            =   -1890
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6120
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   4842
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   0   'False
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Help List"
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
            ColumnWidth     =   2940.095
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   5550
      TabIndex        =   7
      Top             =   5310
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   5535
      TabIndex        =   5
      Top             =   3435
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   3
      Top             =   1515
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C8E8DA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5850
      Left            =   135
      TabIndex        =   12
      Top             =   480
      Width           =   4680
      Begin VB.CommandButton BTNEXIT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E&xit"
         DownPicture     =   "RpVehicle.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exit Form"
         Top             =   5055
         Width           =   1290
      End
      Begin VB.CommandButton BTNPRINT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Print"
         DownPicture     =   "RpVehicle.frx":3132
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print Report"
         Top             =   5055
         Width           =   1290
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
         Height          =   4755
         Left            =   75
         TabIndex        =   0
         Top             =   75
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   8387
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16512
         Rows            =   8
         Cols            =   3
         FixedRows       =   0
         BackColorFixed  =   13166810
         ForeColorFixed  =   16384
         BackColorSel    =   16711680
         ForeColorSel    =   12648447
         BackColorBkg    =   13166810
         GridColor       =   13166810
         GridColorFixed  =   13166810
         GridColorUnpopulated=   13166810
         GridLinesFixed  =   1
         BorderStyle     =   0
         Appearance      =   0
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Left            =   3690
         TabIndex        =   15
         Top             =   5040
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Frame FrmList 
         BorderStyle     =   0  'None
         Height          =   1830
         Left            =   4140
         TabIndex        =   13
         Top             =   5325
         Visible         =   0   'False
         Width           =   2520
         Begin MSComctlLib.ListView ListView 
            Height          =   1830
            Left            =   135
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   -30
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   585
         Left            =   795
         Shape           =   4  'Rounded Rectangle
         Top             =   4980
         Width           =   2955
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   4890
         Left            =   30
         Top             =   15
         Width           =   4575
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   -30
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6765
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1410
      Index           =   1
      Left            =   5520
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2487
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   128
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
      Height          =   1890
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   1455
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   3334
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   128
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
      Height          =   1890
      Index           =   4
      Left            =   5520
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   3334
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   128
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
      Height          =   1890
      Index           =   3
      Left            =   5520
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   3334
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   128
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
End
Attribute VB_Name = "RpVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CellBackColLeave As String = &HFFFFFF
Private Const CellBackColEnter As String = &HFFFFC0
Private Const CellBackColLeave1 As String = &HEDF7FE
Private Const CellBackColEnter1 As String = &HC0E0FF
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim RsGrid3 As ADODB.Recordset
Dim RsGrid4 As ADODB.Recordset
Dim RsDataGrid1 As ADODB.Recordset
Dim RepTitle As String, RepName As String
Dim RepPrint As Boolean
Dim RstRep As ADODB.Recordset
Dim RstRep1 As ADODB.Recordset
Dim SubRep1 As Boolean
Private Const GridRowHeight As Integer = 270
Public GRepFormName As Byte
'Constant
Private Const SprQuot As Byte = 1
Private Const SprSaleOrd As Byte = 2
Private Const SprSaleReg As Byte = 3
Private Const SprSaleRet As Byte = 4
Private Const SprPurOrd As Byte = 5
Private Const SprMatReg As Byte = 6
Private Const SprPurReg As Byte = 7
Private Const SprPurRet As Byte = 8
Private Const SprStkTrf As Byte = 9
Private Const SprStkReg As Byte = 10
Private Const SprStkSumm As Byte = 11
Private Const SprStkInHand As Byte = 12
Private Const VehMoneyRect As Byte = 13
Private Const WksEstimate As Byte = 14
Private Const WksPerforma As Byte = 15
Private Const WksSaleReg As Byte = 16
Private Const WksReqReg As Byte = 17
Private Const WksVehDiary As Byte = 18
Private Const WksJobReg As Byte = 19
'24-09 onwards by lps
Private Const SprABCRep As Byte = 20
Private Const SprFSNRep As Byte = 21
Private Const SprIndent As Byte = 22
Private Const SprDailySale As Byte = 23
Private Const SprMonthSale As Byte = 24

Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const List1 As Byte = 2
Private Const List2 As Byte = 3
Private Const List3 As Byte = 4

Private Const Cat1 As Byte = 5
Private Const Cat2 As Byte = 6

Private Const G1Top As Integer = 36
Private Const G2Top As Integer = 1785
Private Const G3Top As Integer = 4140
Private Const G4Top As Integer = 6480
Dim mLastRow As Integer
Dim mFirstRow As Integer
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
Select Case GRepFormName
    Case SprMonthSale, SprDailySale
        SprMonthDateSale
        If RepPrint = False Then Exit Sub
    Case SprIndent
        SprIndentReg
        If RepPrint = False Then Exit Sub
    Case SprFSNRep
        SprFSNAnalysis
        If RepPrint = False Then Exit Sub
    Case SprABCRep
        If SprABCAnalysis = False Then Exit Sub
    Case WksVehDiary
        WksVehicleDiary
        If RepPrint = False Then Exit Sub
    Case SprQuot, WksEstimate, WksPerforma
        SprQuotReg
        If RepPrint = False Then Exit Sub
    Case WksJobReg
        WksJobRegister
        If RepPrint = False Then Exit Sub
    Case WksReqReg
        WksRequisition
        If RepPrint = False Then Exit Sub
    Case SprPurOrd, SprSaleOrd
        SprSalePurOrd
        If RepPrint = False Then Exit Sub
    Case SprMatReg, SprStkTrf
        SprPurChl
        If RepPrint = False Then Exit Sub
    Case SprSaleReg, SprSaleRet, SprPurReg, SprPurRet, WksSaleReg
        SprSalePurReg
        If RepPrint = False Then Exit Sub
    Case SprStkReg, SprStkSumm, SprStkInHand
        SprStkRep
        If RepPrint = False Then Exit Sub
    Case VehMoneyRect
        VehMoneyRectFunc
        If RepPrint = False Then Exit Sub
End Select

CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")

rpt.Database.SetDataSource RstRep
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1

rpt.ReadRecords



Call Formulas
Call REPORT_VIEW(rpt, RepTitle, , False)
Set RstRep = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub Check1_Click(Index As Integer)
    If Check1(Index).Value = Unchecked Then
        GridSel(Index).Enabled = True
        If FGrid.Rows > 1 Then
            GridSel(Index).Row = 1: GridSel(Index).Col = 1
        End If
    Else
        GridSel(Index).Enabled = False
        If FGrid.Rows > 1 Then
            GridSel(Index).Row = 0: GridSel(Index).Col = 0
            GridSel(Index).RowSel = GridSel(Index).Rows - 1
        End If
    End If
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim i As Byte
WinSetting Me
   Global_Grid
   TopCtrl1.TopText2 = "Add"
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If GridSel(4).Visible = True Then Set RsGrid1 = Nothing
If GridSel(1).Visible = True Then Set RsGrid2 = Nothing
If GridSel(2).Visible = True Then Set RsGrid3 = Nothing
If GridSel(3).Visible = True Then Set RsGrid4 = Nothing
Set RstRep = Nothing
Set mListItem = Nothing
End Sub

Private Sub GridSel_Click(Index As Integer)
'Dim i As Integer
'If GridSel(Index).Col <> 0 Or GridSel(Index).Rows < 1 Then Exit Sub
'    GridSel(Index).CellFontName = "WINGDINGS"
'    GridSel(Index).CellFontSize = 14
'    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü", " ", "ü")
'    Select Case Index
'        Case 1
'            i = UBound(GridRow1) + 1
'            ReDim Preserve GridRow1(i)
'            GridRow1(i) = GridSel(Index).Row
'        Case 2
'            i = UBound(GridRow2) + 1
'            ReDim Preserve GridRow2(i)
'            GridRow2(i) = GridSel(Index).Row
'        Case 3
'            i = UBound(GridRow3) + 1
'            ReDim Preserve GridRow3(i)
'            GridRow3(i) = GridSel(Index).Row
'        Case 4
'            i = UBound(GridRow4) + 1
'            ReDim Preserve GridRow4(i)
'            GridRow4(i) = GridSel(Index).Row
'    End Select
End Sub

Private Sub GridSel_EnterCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColEnter1
End Sub

Private Sub GridSel_GotFocus(Index As Integer)
GridSel(Index).CellBackColor = CellBackColEnter1
End Sub

Private Sub GridSel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
If KeyCode = 13 Then SendKeys vbTab
If GridSel(Index).Rows < 1 Then Exit Sub
If KeyCode = vbKeySpace And GridSel(Index).Col = 0 Then
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü", " ", "ü")
    Select Case Index
        Case 1
            i = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(i)
            GridRow1(i) = GridSel(Index).Row
        Case 2
            i = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(i)
            GridRow2(i) = GridSel(Index).Row
        Case 3
            i = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(i)
            GridRow3(i) = GridSel(Index).Row
        Case 4
            i = UBound(GridRow4) + 1
            ReDim Preserve GridRow4(i)
            GridRow4(i) = GridSel(Index).Row
    End Select
End If
End Sub

Private Sub GridSel_KeyPress(Index As Integer, KeyAscii As Integer)
If GridSel(Index).Col = 0 Or GridSel(Index).Row = 0 Then Exit Sub
Select Case Index
    Case 1
       SelGridKeyPress TxtSearch, GridSel, Index, RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPress TxtSearch, GridSel, Index, RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPress TxtSearch, GridSel, Index, RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPress TxtSearch, GridSel, Index, RsGrid4, KeyAscii, RsGrid4.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
End Select
TxtSearch.Tag = Index
End Sub
Private Sub TxtSearch_Click()
TxtSearch.Visible = False: TxtSearch.Text = "": GridSel(Val(TxtSearch.Tag)).SetFocus
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If NavigationKey(KeyCode) = True Then TxtSearch.Visible = False: GridSel(Val(TxtSearch.Tag)).SetFocus
If KeyCode = vbKeyDelete Then TxtSearch.Text = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then TxtSearch.Visible = False: GridSel(Val(TxtSearch.Tag)).SetFocus
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
Select Case TxtSearch.Tag
    Case 1
       SelGridKeyPress TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPress TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPress TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPress TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid4, KeyAscii, RsGrid4.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
End Select
End Sub

Private Sub TxtSearch_LostFocus()
TxtSearch.Visible = False: TxtSearch.Text = ""
End Sub

Private Sub GridSel_LeaveCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub

Private Sub GridSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If GridSel(Index).Col <> 0 Then Exit Sub
mGridStartRow = GridSel(Index).Row
End Sub

Private Sub GridSel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
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
            i = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(i)
            GridRow1(i) = GridSel(Index).Row
        Case 2
            i = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(i)
            GridRow2(i) = GridSel(Index).Row
        Case 3
            i = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(i)
            GridRow3(i) = GridSel(Index).Row
        Case 4
            i = UBound(GridRow4) + 1
            ReDim Preserve GridRow4(i)
            GridRow4(i) = GridSel(Index).Row
    End Select
Next
mGridStartRow = 0
End Sub

Private Sub GridSel_Validate(Index As Integer, Cancel As Boolean)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub

Private Sub ListView_Click()
    TxtGrid(0).Text = ListView.SelectedItem.Text
    FrmList.Visible = False
    TxtGrid(0).SetFocus
End Sub


Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    'SprQuot,SprSaleOrd,SprSaleReg,SprSaleRet,SprPurOrd,
    'SprMatReg,SprPurReg,SprPurRet,SprStkTrf
    'SprStkReg,SprStkSumm,SprStkInHand,VehMoneyRect
    Select Case FGrid.Row
    Case Cat1, Cat2
        TxtGrid(0).MaxLength = 5
    Case List1
            Select Case GRepFormName
            Case SprSaleOrd
              ListArray = Array("All", "Pending")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprSaleReg, SprPurReg, WksSaleReg
                ListArray = Array("All", "Cash", "Credit")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprSaleRet, SprPurRet
                ListArray = Array("Cash", "Credit", "Transfer")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprPurOrd
              ListArray = Array("Annual", "Quarterly", "Monthly", "General(Casual)", "VOR")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 5)
            Case SprQuot, WksEstimate, WksPerforma
              ListArray = Array("Stores", "Workshop")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VehMoneyRect
              ListArray = Array("All", "Form-60", "Form-61", "N/A")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
            Case WksReqReg, WksJobReg
              ListArray = Array("All", "Closed", "UnClosed")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
          End Select
    Case List2
        Select Case GRepFormName
        Case SprPurOrd
              ListArray = Array("All", "Pending", "Excess")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
        Case WksReqReg
            ListArray = Array("All", "PDI", "Free Service", "Chargable", "Warranty", "Company Vehicle", "Complementary")
            Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 7)
        End Select
    Case List3
'    Case DGrid1
'        If DGHelp1.Visible = False Then DGHelp1.left = TxtGrid(Index).left: DGHelp1.top = TxtGrid(Index).top + TxtGrid(Index).Height
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
If KeyCode = vbKeyEscape Then
    TxtGrid(0).Text = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    TxtGrid(0).Visible = False
    Grid_Hide
    FGrid.SetFocus
    Exit Sub
End If
Select Case FGrid.Row
Case List1, List2, List3
        ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).Height + 25), TxtGrid(0).width
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
'Case DGrid1
'        DGridTxtKeyDown DGHelp1, TxtGrid, 0, RsDataGrid1, KeyCode, True, 1
'        If KeyCode = vbKeyReturn Then
'            If TxtGridLeave = True Then TxtKeyDown
'        End If
Case Date1, Date2, Cat1, Cat2
    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
        If TxtGridLeave = True Then TxtKeyDown
    End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
 Call CheckQuote(KeyAscii)
 Select Case FGrid.Row
    Case Cat1, Cat2
        NumPress TxtGrid(Index), KeyAscii, 2, 2
'     Case DGrid1
'        If DGHelp1.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsDataGrid1, KeyAscii, "Name"
 End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
    Select Case FGrid.Row
'        Case Cat1, Cat2
'             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
        
'        Case Cat2
'            'If Val(FGrid.TextMatrix(Cat2, 1)) > Val(FGrid.TextMatrix(Cat2, 1)) Then
            
        Case List1, List2, List3
            If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
            ListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
'       Case DGrid1
'           If KeyCode <> 13 And DGHelp1.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, 0, RsDataGrid1, KeyCode, "Name", True
    End Select
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Select Case FGrid.Row
        Case Cat1, Cat2
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
        Case List1, List2, List3
            If TxtGrid(0).Text <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.Text
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
'        Case DGrid1
'            If RsDataGrid1.RecordCount = 0 Or (RsDataGrid1.EOF = True Or RsDataGrid1.BOF = True) Or TxtGrid(0).Text = "" Then
'                FGrid.TextMatrix(FGrid.Row, 1) = ""
'                FGrid.TextMatrix(FGrid.Row, 2) = ""
'            Else
'                FGrid.TextMatrix(FGrid.Row, 1) = RsDataGrid1!Name
'                FGrid.TextMatrix(FGrid.Row, 2) = RsDataGrid1!Code
'            End If
End Select
End Sub

Private Function TxtGridLeave() As Boolean
Select Case FGrid.Row
        Case Cat1, Cat2
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
        Case List1, List2, List3
            If TxtGrid(0).Text <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.Text
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
'        Case DGrid1
'            If RsDataGrid1.RecordCount = 0 Or (RsDataGrid1.EOF = True Or RsDataGrid1.BOF = True) Or TxtGrid(0).Text = "" Then
'                FGrid.TextMatrix(FGrid.Row, 1) = ""
'                FGrid.TextMatrix(FGrid.Row, 2) = ""
'            Else
'                FGrid.TextMatrix(FGrid.Row, 1) = RsDataGrid1!Name
'                FGrid.TextMatrix(FGrid.Row, 2) = RsDataGrid1!Code
'            End If
End Select
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function

'******* Fuctions **********

Private Sub Global_Grid()
Dim i As Integer
Frame1.top = 775: Frame1.left = 300: FGrid.top = 75: FGrid.left = 75
FGrid.Rows = 7  '5
FGrid.Cols = 3
FGrid.FixedCols = 1
FGrid.ColWidth(0) = 2200
FGrid.ColWidth(1) = 2000
FGrid.ColWidth(2) = 0
FGrid.ColAlignment(1) = flexAlignLeftCenter
For i = 0 To FGrid.Rows - 1
    FGrid.RowHeight(i) = 0
Next
Ini_Grid
End Sub
Private Sub Grid_Hide()
If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub FGrid_DblClick()
    Select Case FGrid.Row
        Case Date1, Date2, List1, List2, List3, Cat1, Cat2
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
Dim i As Integer
    Select Case FGrid.Row
        Case Cat1, Cat2
            If KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
            Else
                KeyAscii = 0
            End If
        Case Date1, Date2, List1, List2, List3, Cat1, Cat2
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell--> Enter Cell-->KeyDown
If KeyCode = vbKeyUp And Val(FGrid.Tag) = mFirstRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = mLastRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys vbTab
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
        Case Date1, Date2, List1, List2, List3
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
        Case Cat1, Cat2
            
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
Dim Ac_Str As String
Dim i As Integer
Dim GridRow As Integer
    Ac_Str = ""
    For i = 0 To UBound(GridArray)
        If GridArray(i) = 0 Then GoTo NXT:
        GridRow = GridArray(i)
        If GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü" Then
                If DataType = 0 Then
                   Ac_Str = Ac_Str + IIf(Ac_Str = "", GridSel(Gridindex).TextMatrix(GridRow, 2), "," + GridSel(Gridindex).TextMatrix(GridRow, 2))
                ElseIf DataType = 1 Then
                   Ac_Str = Ac_Str + IIf(Ac_Str = "", "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'", "," + "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'")
                End If
            GridSel(Gridindex).TextMatrix(GridRow, 0) = ""
        Else
            GridArray(i) = 0
        End If
NXT:
    Next
    For i = 0 To UBound(GridArray)
        GridRow = GridArray(i)
        If GridArray(i) <> 0 Then
            GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü"
        End If
    Next
'    Erase GridArray
'    ReDim Preserve GridArray(0)
'    GridArray(0) = 0
    If Ac_Str = "" Then
        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
        GridSel(Gridindex).SetFocus
        RepPrint = False
        Exit Function
    End If
    FillString = Ac_Str
    Exit Function
End Function

Private Sub TxtKeyDown()
Dim i As Integer
    If FGrid.Row = mLastRow Then SendKeys vbTab: Exit Sub
    For i = FGrid.Row To FGrid.Rows - 1
         If FGrid.RowHeight(i + 1) <> 0 Then FGrid.Row = i + 1: Exit For
    Next
End Sub
Private Sub GridInitialise(Gridindex As Integer, GridSql As String)
Dim Index As Integer
Index = Gridindex
If Index = 1 Then
    Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
    RsGrid1.Open GridSql, GCn, adOpenDynamic, adLockOptimistic: Set GridSel(Index).DataSource = RsGrid1
    GridSel(Index).top = G1Top
    ReDim Preserve GridRow1(0)
    GridRow1(0) = 0
End If
If Index = 2 Then
    Set RsGrid2 = New ADODB.Recordset: RsGrid2.CursorLocation = adUseClient
    RsGrid2.Open GridSql, GCn, adOpenDynamic, adLockOptimistic: Set GridSel(Index).DataSource = RsGrid2
    GridSel(Index).top = G2Top
    ReDim Preserve GridRow2(0)
    GridRow2(0) = 0
End If
If Index = 3 Then
    Set RsGrid3 = New ADODB.Recordset: RsGrid3.CursorLocation = adUseClient
    RsGrid3.Open GridSql, GCn, adOpenDynamic, adLockOptimistic: Set GridSel(Index).DataSource = RsGrid3
    GridSel(Index).top = G3Top
        ReDim Preserve GridRow3(0)
        GridRow3(0) = 0
End If
If Index = 4 Then
    Set RsGrid4 = New ADODB.Recordset: RsGrid4.CursorLocation = adUseClient
    RsGrid4.Open GridSql, GCn, adOpenDynamic, adLockOptimistic: Set GridSel(Index).DataSource = RsGrid4
    GridSel(Index).top = G4Top
    ReDim Preserve GridRow4(0)
    GridRow4(0) = 0
End If
GridSel(Index).Visible = True: GridSel(Index).Enabled = False: Check1(Index).Visible = True
GridSel(Index).width = 5200: GridSel(Index).ColWidth(0) = 600: GridSel(Index).ColWidth(2) = 0: GridSel(Index).ColWidth(1) = 4000
Check1(Index).top = GridSel(Index).top + 20: Check1(Index).left = GridSel(Index).left + 40: Check1(Index).width = 560
Check1(Index).Height = GridSel(Index).RowHeight(0) + 40: Check1(Index).Value = Checked
End Sub

Private Sub Ini_Grid()
'Date1,Date2,List1,List1,List2,List3
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Select Case GRepFormName
    Case SprMonthSale, SprDailySale
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2
    
    Case SprIndent
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2
    
      Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
      GridInitialise 2, Grid2Sql
      Grid3Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
      GridInitialise 3, Grid3Sql
      Grid4Sql = "select '' as O,Part_no as PartNo,Part_no as code,Part_name as PartName from Part order by Part_no,part_name"
      GridInitialise 4, Grid4Sql
      GridSel(4).width = GridSel(4).width + 1000: GridSel(4).ColWidth(1) = 1500: GridSel(4).ColWidth(3) = 3500
    
    Case SprFSNRep  '24-09 lps
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Cat1, 0) = "Fast %": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Cat2, 0) = "Slow %": .RowHeight(Cat2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(Cat1, 1) = ""  'All"
            .TextMatrix(Cat2, 1) = ""  'All"
        End With
        mFirstRow = Date1: mLastRow = Cat2
        
    Case SprABCRep  '24-09 lps
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Cat1, 0) = "A Category %": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Cat2, 0) = "B Category %": .RowHeight(Cat2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(Cat1, 1) = ""  'All"
            .TextMatrix(Cat2, 1) = ""  'All"
        End With
        mFirstRow = Date1: mLastRow = Cat2
        
    Case WksVehDiary
        With FGrid
            .TextMatrix(Date1, 0) = "Upto Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date1
    Case WksReqReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Job Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Purpose Of Part": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List2, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List2
          Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
          Grid3Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
          GridInitialise 3, Grid3Sql
    Case WksJobReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Job Type": .RowHeight(List1) = GridRowHeight
          
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List1
          Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
          Grid3Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
          GridInitialise 3, Grid3Sql
    Case SprMatReg, SprStkTrf
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2
    Case SprSaleReg, SprSaleRet, SprPurReg, SprPurRet, WksSaleReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Cash/Credit": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Cash"
        End With
            mFirstRow = Date1: mLastRow = List1
    Case SprQuot
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Stores/Workshop": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Stores"
        End With
          mFirstRow = Date1: mLastRow = List1
          Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
    Case WksEstimate, WksPerforma
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Stores/Workshop": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Workshop"
        End With
            mFirstRow = Date1: mLastRow = List1
          Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
    Case SprPurOrd
        With FGrid
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Order Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Report Option": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Monthly"
            .TextMatrix(List2, 1) = "All"
        End With
          mFirstRow = Date2: FGrid.Row = mFirstRow: mLastRow = List2
          Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
          Grid3Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
          GridInitialise 3, Grid3Sql
          Grid4Sql = "select '' as O,cityname as CityName,CityCode as code from city  order by cityname"
          GridInitialise 4, Grid4Sql
    Case SprStkReg, SprStkSumm, SprStkInHand
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
            
          mFirstRow = Date1: mLastRow = Date2
          Grid2Sql = "select '' as O ,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
          
          Grid3Sql = "select '' as O,Part_no as PartNo,Part_no as code,Part_name as PartName from Part order by Part_no,part_name"
          GridInitialise 3, Grid3Sql
          GridSel(3).width = GridSel(3).width + 1000: GridSel(3).ColWidth(1) = 1500: GridSel(3).ColWidth(3) = 3500
    Case SprSaleOrd
        With FGrid
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Pending/All": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
          mFirstRow = Date2: FGrid.Row = mFirstRow
          mLastRow = List1
          Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
          Grid3Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
          GridInitialise 3, Grid3Sql
    Case VehMoneyRect
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Under Declaration": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
          mFirstRow = Date1: FGrid.Row = mFirstRow
          mLastRow = List1
          Grid2Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 2, Grid2Sql
          Grid3Sql = "select '' as O,Description as VoucherType,v_type as code from Voucher_Type where category='GENFA' order by v_type "
          GridInitialise 3, Grid3Sql
          Grid4Sql = "select '' as O,CityName as CityName,CityCode as code from City order by CityName"
          GridInitialise 4, Grid4Sql
'        Set RsDataGrid1 = New ADODB.Recordset
'        RsDataGrid1.CursorLocation = adUseClient
'        RsDataGrid1.Open "select Voucher_Type.v_type as code,Description as name from Voucher_Type where category='GENFA' order by v_type ", GCnFa, adOpenDynamic, adLockOptimistic
'        Set DGHelp1.DataSource = RsDataGrid1

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

Private Sub SprQuotReg()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    CondStr = " where E.v_Date >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and E.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#  and E.Stores_Works = '" & FGrid.TextMatrix(List1, 1) & "'"
    Select Case GRepFormName
        Case WksEstimate
            CondStr = CondStr & " and E.V_type = '" & WksEst & "'"
        Case WksPerforma
            CondStr = CondStr & " and E.V_type = '" & WksPro & "'"
        Case SprQuot
            CondStr = CondStr & " and E.V_type = '" & SprQuotation & "'"
    End Select
    
    If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(E.site_code,1) in (" & GridString2 & ") and E.v_Date >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and E.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#  and E.Stores_Works = '" & FGrid.TextMatrix(List1, 1) & "'"
    
    mQRY = "SELECT E.V_DATE,E.V_NO,E.Party_Name,(E.SprAmt_TB +E.SprAmt_TP + E.SprAmt_MRP_TB + E.SprAmt_MRP_TP) as SprAmt, (E.OilAmt_MRP_TB + E.OilAmt_MRP_TP + " & _
    "E.OilAmt_TB + E.OilAmt_TP) as OilAmt, (E.D_Amt_TB +  E.D_Amt_TP + E.D_Amt_MRP_TB + E.D_Amt_MRP_TP) as DisAmt, E.Total_Amt, E.Addition, E.Gen_Sur_Amt," & _
    "E.Trans_Amt, (E.Tax_Amt + E.Tax_AmtMRP) as TaxAmt, (E.Tax_Sur_Amt + E.TaxSur_AmtMRP) as SurAmt," & _
    "E.Packing, E.Lab_Amt, E.Lab_D_Amt " & _
    "FROM Estimate E"
    
    mQRY = mQRY + CondStr + " order by E.v_date"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprEstQuot"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub SprSalePurOrd()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
    RepPrint = True
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If GRepFormName = SprPurOrd Then If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    If GRepFormName = SprPurOrd And (FGrid.TextMatrix(List2, 1) = "All" Or FGrid.TextMatrix(List2, 1) = "Pending") Then
        CondStr = " where P.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# and (isnull(P.OrdClosDate) or P.OrdClosDate='') "
        If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(P.site_code,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then CondStr = CondStr & " and P.Party_Code in (" & GridString3 & ")"
        If Check1(4).Value = Unchecked Then CondStr = CondStr & " and C.CityCode in (" & GridString4 & ")"
        CondStr = CondStr & " and P.Order_type <> 'S_SO' and Right(P.Order_type,1) = '" & left(FGrid.TextMatrix(List1, 1), 1) & "'"
        If FGrid.TextMatrix(List2, 1) = "Pending" Then
            CondStr = CondStr & " and P1.QTY-P1.Sup_Qty > 0"
        End If
        mQRY = "SELECT C.CityName,P.OrderId,site.site_Desc,P.Party_Code, P.Order_No, P1.Amount, P.Site_Code, P.V_Date, P.Order_Prefix, Part.Part_Name, SubGroup.Name, P1.PART_NO, P.Order_Reg_No, P.Order_Reg_Dt, P1.QTY, P1.Sup_Qty, (P1.QTY-P1.Sup_Qty) AS BalQty " & _
    "FROM ((((SP_Order AS P LEFT JOIN SP_Order1 AS P1 ON P.OrderId = P1.OrderId) LEFT JOIN Part ON P1.PART_NO = Part.PART_NO) LEFT JOIN SubGroup ON P.Party_Code = SubGroup.SubCode) LEFT JOIN Site ON left(P.Site_Code,1) = Site.Site_Code) LEFT JOIN City as C ON SubGroup.CityCode = C.CityCode"
    mQRY = mQRY + CondStr + " order by site.site_Desc,P.v_date, P.Order_Prefix,P.Order_No"
    ElseIf GRepFormName = SprPurOrd And FGrid.TextMatrix(List2, 1) = "Excess" Then
        CondStr = " where P.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# and (isnull(P.OrdClosDate) or P.OrdClosDate='') "
        If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(P.site_code,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then CondStr = CondStr & " and P.Party_Code in (" & GridString3 & ")"
        If Check1(4).Value = Unchecked Then CondStr = CondStr & " and C.CityCode in (" & GridString4 & ")"
        CondStr = CondStr & " and P.Order_type <> 'S_SO' and Right(P.Order_type,1) = '" & left(FGrid.TextMatrix(List1, 1), 1) & "'"
        If FGrid.TextMatrix(List2, 1) = "Pending" Then
            CondStr = CondStr & " and P1.QTY-P1.Sup_Qty < 0"
        End If
        mQRY = "SELECT C.CityName,P.OrderId,site.site_Desc,P.Party_Code, P.Order_No, P1.Amount, P.Site_Code, P.V_Date, P.Order_Prefix, Part.Part_Name, SubGroup.Name, P1.PART_NO, P.Order_Reg_No, P.Order_Reg_Dt, P1.QTY, P1.Sup_Qty, (P1.QTY-P1.Sup_Qty) AS BalQty " & _
        "FROM ((((SP_Order AS P LEFT JOIN SP_Order1 AS P1 ON P.OrderId = P1.OrderId) LEFT JOIN Part ON P1.PART_NO = Part.PART_NO) LEFT JOIN SubGroup ON P.Party_Code = SubGroup.SubCode) LEFT JOIN Site ON left(P.Site_Code,1) = Site.Site_Code) LEFT JOIN City as C ON SubGroup.CityCode = C.CityCode"
        mQRY = mQRY + CondStr + " order by site.site_Desc,P.v_date, P.Order_Prefix,P.Order_No"
    ElseIf GRepFormName = SprSaleOrd Then
        CondStr = " where S.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# and (isnull(S.OrdClosDate) or S.OrdClosDate='') "
        If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(S.site_code,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then CondStr = CondStr & " and S.Party_Code in (" & GridString3 & ")"
        CondStr = CondStr & " and S.Order_type = 'S_SO'"
        If FGrid.TextMatrix(List1, 1) = "Pending" Then
            CondStr = CondStr & " and S1.QTY-S1.Sup_Qty <> 0"
        End If
        mQRY = "SELECT S.OrderId,site.site_Desc,S.Party_Code, S.Order_No, S1.Amount, S.Site_Code, S.V_Date, S.Order_Prefix, Part.Part_Name, SubGroup.Name, S1.PART_NO, S.Order_Reg_No, S.Order_Reg_Dt, S1.QTY, S1.Sup_Qty, (S1.QTY-S1.Sup_Qty) AS BalQty " & _
        "FROM (((SP_Order AS S LEFT JOIN SP_Order1 AS S1 ON S.OrderId = S1.OrderId) LEFT JOIN Part ON S1.PART_NO = Part.PART_NO) LEFT JOIN SubGroup ON S.Party_Code = SubGroup.SubCode) LEFT JOIN Site ON left(S.Site_Code,1) = Site.Site_Code"
        mQRY = mQRY + CondStr + " order by site.site_Desc,S.v_date, S.Order_Prefix,S.Order_No"
    End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprPurOrd"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub SprIndentReg()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    CondStr = " where Indent.Doc_Date >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and Indent.Doc_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
    If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(Indent.site_code,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then CondStr = CondStr & " and Indent.PartyCode in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then CondStr = CondStr & " and Indent.PART_NO in (" & GridString4 & ")"
        
    mQRY = "SELECT SubGroup.Name, Part.Part_Name, Indent.DocID, Indent.IDNo, Indent.Doc_Date, Indent.PART_NO, Indent.QTY, Indent.RATE, Indent.Remark " & _
    "FROM (Indent LEFT JOIN Part ON Indent.PART_NO = Part.PART_NO) LEFT JOIN SubGroup ON Indent.PartyCode = SubGroup.SubCode"
    mQRY = mQRY + CondStr + " order by Indent.IDNo"

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprIndent"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub SprMonthDateSale()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String, mQRY1 As String
    RepPrint = True
    SubRep1 = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    CondStr = " where SP_Sale.V_Date >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and SP_Sale.V_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
    ' "SYSIC","SYSIR","W_SIC","W_SIR"
    If GRepFormName = SprMonthSale Then
     mQRY = "SELECT month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_MRP_TB - SP_Sale.D_Amt_TB) as TaxableAmt,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_MRP_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt, sum(SP_Sale.Gen_Sur_Amt + SP_Sale.Tax_AmtMRP + SP_Sale.Tax_Amt +  + SP_Sale.Tax_Sur_Amt + SP_Sale.TaxSur_AmtMRP) as Tax,'S' as RepType  " & _
    "FROM SP_Sale" & CondStr & " and sp_sale.v_type in ('SYSIC','SYSIR')  group by month(v_date) " & _
    "Union All " & _
    "SELECT month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_MRP_TB - SP_Sale.D_Amt_TB) as TaxableAmt,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_MRP_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt, sum(SP_Sale.Gen_Sur_Amt + SP_Sale.Tax_AmtMRP + SP_Sale.Tax_Amt +  + SP_Sale.Tax_Sur_Amt + SP_Sale.TaxSur_AmtMRP) as Tax,'W' as RepType " & _
    "FROM SP_Sale" & CondStr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by month(v_date)"
    
    mQRY1 = "SELECT iif(SubGroupType.Description<>'',SubGroupType.Description,'Others') as Descrip, Sum(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB+SP_Sale.SprAmt_TB+SP_Sale.OilAmt_TB-SP_Sale.D_Amt_MRP_TB-SP_Sale.D_Amt_TB) AS TaxableAmt, Sum(SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TP-SP_Sale.D_Amt_MRP_TP-SP_Sale.D_Amt_TP) AS TaxpaidAmt, Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Tax_AmtMRP+SP_Sale.Tax_Amt++SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP) AS Tax " & _
    "FROM SP_Sale LEFT JOIN (SubGroup LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type) ON SP_Sale.Party_Code = SubGroup.SubCode" & CondStr & " and sp_sale.v_type in ('SYSIC','SYSIR') group by SubGroupType.Description"
    
    RepName = "SprMonthSale"
    
    ElseIf GRepFormName = SprDailySale Then
     mQRY = "SELECT SP_Sale.V_Date, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_MRP_TB - SP_Sale.D_Amt_TB) as TaxableAmt,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_MRP_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt, sum(SP_Sale.Gen_Sur_Amt + SP_Sale.Tax_AmtMRP + SP_Sale.Tax_Amt +  + SP_Sale.Tax_Sur_Amt + SP_Sale.TaxSur_AmtMRP) as Tax,'S' as RepType  " & _
    "FROM SP_Sale" & CondStr & " and sp_sale.v_type in ('SYSIC','SYSIR')  group by SP_Sale.V_Date " & _
    "Union All " & _
    "SELECT SP_Sale.V_Date, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_MRP_TB - SP_Sale.D_Amt_TB) as TaxableAmt,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_MRP_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt, sum(SP_Sale.Gen_Sur_Amt + SP_Sale.Tax_AmtMRP + SP_Sale.Tax_Amt +  + SP_Sale.Tax_Sur_Amt + SP_Sale.TaxSur_AmtMRP) as Tax,'W' as RepType " & _
    "FROM SP_Sale" & CondStr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by SP_Sale.V_Date"
    
    mQRY1 = "SELECT iif(SubGroupType.Description<>'',SubGroupType.Description,'Others') as Descrip, Sum(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB+SP_Sale.SprAmt_TB+SP_Sale.OilAmt_TB-SP_Sale.D_Amt_MRP_TB-SP_Sale.D_Amt_TB) AS TaxableAmt, Sum(SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TP-SP_Sale.D_Amt_MRP_TP-SP_Sale.D_Amt_TP) AS TaxpaidAmt, Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Tax_AmtMRP+SP_Sale.Tax_Amt++SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP) AS Tax " & _
    "FROM SP_Sale LEFT JOIN (SubGroup LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type) ON SP_Sale.Party_Code = SubGroup.SubCode" & CondStr & " and sp_sale.v_type in ('SYSIC','SYSIR') group by SubGroupType.Description"
        RepName = "SprDailySale"
    End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    
    Set RstRep1 = New Recordset
    RstRep1.CursorLocation = adUseClient
    RstRep1.Open (mQRY1), GCn, adOpenDynamic, adLockOptimistic
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub

Private Sub Formulas()
On Error GoTo ELoop
Dim i As Integer
Select Case GRepFormName
Case SprMatReg, SprStkTrf, WksReqReg, SprMonthSale, SprDailySale
    For i = 1 To rpt.FormulaFields.count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(i).Text = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
    End Select
    Next
Case SprPurOrd, WksVehDiary, SprSaleOrd, WksJobReg
    For i = 1 To rpt.FormulaFields.count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(i).Text = "'Upto Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "'"
    End Select
    Next
Case SprABCRep
    For i = 1 To rpt.FormulaFields.count
        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(i).Text = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("CatA")
                rpt.FormulaFields(i).Text = Val(FGrid.TextMatrix(Cat1, 1))
            Case UCase("CatB")
                rpt.FormulaFields(i).Text = Val(FGrid.TextMatrix(Cat2, 1))
            Case UCase("RepBase")
                rpt.FormulaFields(i).Text = "'Formula of %   = (Consumption Value of each Item *100)/Total Consumption Value'"
            Case UCase("RepBase2")
                rpt.FormulaFields(i).Text = "'Category  A = Top " & Val(FGrid.TextMatrix(Cat1, 1)) & "%,  B = Next " & Val(FGrid.TextMatrix(Cat2, 1)) & "%,  C = Remaining consumption value'"
        End Select
    Next

Case SprQuot, WksEstimate, WksPerforma
    For i = 1 To rpt.FormulaFields.count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("SPRWORK")
            rpt.FormulaFields(i).Text = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "'"
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(i).Text = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
    End Select
    Next
Case SprSaleReg, SprSaleRet, WksSaleReg
    For i = 1 To rpt.FormulaFields.count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(i).Text = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        Case UCase("List1")
            rpt.FormulaFields(i).Text = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Sales'"
    End Select
    Next
Case SprPurReg, SprPurRet
    For i = 1 To rpt.FormulaFields.count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(i).Text = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        Case UCase("List1")
            rpt.FormulaFields(i).Text = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Purchase'"
    End Select
    Next
End Select
Exit Sub
ELoop:
    MsgBox err.Description
End Sub

Private Sub SprPurChl()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If GRepFormName = SprMatReg Then
        CondStr = "SP_Purch.V_Type='" & SprMrRct & " ' And SP_Purch.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and SP_Purch.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
    ElseIf GRepFormName = SprStkTrf Then
        CondStr = "SP_Purch.V_Type='" & SprMrTrf & "' And SP_Purch.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and SP_Purch.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
    End If
    mQRY = "SELECT SP_Purch.DocID,SP_Purch.Party_Name, SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date, SP_Purch.GR_RR_No, " & _
        "SP_Purch.GR_RR_Date, SP_Purch.Cash_Credit, SP_Purch.Tot_No_of_Items, SP_Purch.Tot_Doc_Qty, SP_Purch.Tot_Phy_Qty," & _
        "SP_Purch.Tot_Goods_Value, SP_Purch.NET_AMT, TaxForms.Form_Desc, SP_Stock.Part_No, SP_Stock.Qty_Doc, SP_Stock.Qty_Rec," & _
        "SP_Stock.Rate, SP_Purch.V_Type, SP_Purch.V_No, SP_Stock.Amount,SP_Purch.V_Date " & _
        "FROM (SP_Purch LEFT JOIN SP_Stock ON (SP_Purch.V_Type = SP_Stock.V_Type) AND (SP_Purch.V_No = SP_Stock.V_No)) " & _
        "LEFT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code " & _
        "Where " & CondStr & ""
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprMatReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub SprSalePurReg()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    Select Case GRepFormName
    Case SprSaleReg, SprSaleRet
        If GRepFormName = SprSaleReg Then
            If FGrid.TextMatrix(List1, 1) = "All" Then CondStr = "SP_Sale.V_Type In ('" & SprSlCsh & "','" & SprSlCre & "') And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then CondStr = "SP_Sale.V_Type = '" & SprSlCre & "' And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then CondStr = "SP_Sale.V_Type = '" & SprSlCsh & "' And "
            CondStr = CondStr + "SP_Sale.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and SP_Sale.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
            If FGrid.TextMatrix(List1, 1) = "All" Then
                RepName = "SprSalRegAll"
            Else
                RepName = "SprSalReg"
            End If
        ElseIf GRepFormName = SprSaleRet Then
            If FGrid.TextMatrix(List1, 1) = "Transfer" Then CondStr = "SP_Sale.V_Type = '" & SprSlTrfRet & "' And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then CondStr = "SP_Sale.V_Type = '" & SprSlRetCre & "' And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then CondStr = "SP_Sale.V_Type = '" & SprSlRetCsh & "' And "
            CondStr = CondStr + "SP_Sale.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and SP_Sale.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
            RepName = "SprSalRegAll"
        End If
        mQRY = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type, SP_Sale.V_No, " & _
                 "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, " & _
                 "SP_Sale.SprAmt_MRP_TP,SP_Sale.SprAmt_TB + SP_Sale.SprAmt_MRP_TB as SprAmtTB, SP_Sale.SprAmt_TP  + SP_Sale.SprAmt_MRP_TP as SprAmtTP, SP_Sale.OilAmt_TB  + SP_Sale.OilAmt_MRP_TB as OilAmtTB , " & _
                 "SP_Sale.OilAmt_TP + SP_Sale.OilAmt_MRP_TB as OilAmtTP, SP_Sale.D_Per_TB, SP_Sale.D_Amt_TB, SP_Sale.D_Per_TP,SP_Sale.D_Amt_TP,SP_Sale.Addition, " & _
                 "SP_Sale.Packing, SP_Sale.Gen_Sur_Per, SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt, SP_Sale.Tax_Per, " & _
                 "SP_Sale.Tax_Amt, SP_Sale.Tax_Sur_Per, SP_Sale.Tax_Sur_Amt,SP_Sale.TOT_Per, SP_Sale.TOT_Amt, " & _
                 "SP_Sale.Rounded, SP_Sale.Total_Amt FROM SP_Sale Where " & CondStr & ""
    Case WksSaleReg
            If FGrid.TextMatrix(List1, 1) = "All" Then CondStr = "SP_Sale.V_Type In ('" & WksSlCsh & "','" & WksSlCre & "') And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then CondStr = "SP_Sale.V_Type = '" & WksSlCre & "' And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then CondStr = "SP_Sale.V_Type = '" & WksSlCsh & "' And "
            CondStr = CondStr + "Job_Card.JobCloseDate >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and Job_Card.JobCloseDate<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
            RepName = "WksSalReg"
            mQRY = "SELECT Job_Card.DocId, Job_Card.Job_No,Job_Card.JobCloseDate,Job_Card.NetLab_Amt,SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type, SP_Sale.V_No, SP_Sale.Party_Name," & _
            "SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, SP_Sale.SprAmt_MRP_TP, SP_Sale.SprAmt_TB+SP_Sale.SprAmt_MRP_TB AS SprAmtTB, SP_Sale.SprAmt_TP+SP_Sale.SprAmt_MRP_TP AS SprAmtTP, SP_Sale.OilAmt_TB+SP_Sale.OilAmt_MRP_TB AS OilAmtTB," & _
            "SP_Sale.OilAmt_TP+SP_Sale.OilAmt_MRP_TB AS OilAmtTP, SP_Sale.D_Per_TB, SP_Sale.D_Amt_TB, SP_Sale.D_Per_TP, SP_Sale.Gen_Sur_Per, SP_Sale.Gen_Sur_Amt, SP_Sale.Tax_Per, SP_Sale.Tax_Amt, SP_Sale.Tax_Sur_Per, SP_Sale.Tax_Sur_Amt,  SP_Sale.Rounded, SP_Sale.Total_Amt,DocId_InvSpr,DocId_InvLab " & _
            "FROM Job_Card LEFT JOIN SP_Sale ON Job_Card.DocId = SP_Sale.Job_DocID Where " & CondStr & ""
    Case SprPurReg, SprPurRet
        If GRepFormName = SprPurReg Then
            If FGrid.TextMatrix(List1, 1) = "All" Then CondStr = "SP_Purch.V_Type In ('" & SprPurCre & "','" & SprPurCsh & "') And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then CondStr = "SP_Purch.V_Type = '" & SprPurCre & "' And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then CondStr = "SP_Purch.V_Type = '" & SprPurCsh & "' And "
            CondStr = CondStr + "SP_Purch.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and SP_Purch.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
            mQRY = "SELECT docid,v_date, v_no,  mid(DocID,9,5) as VPrefix, party_name, (spramt_mrp_tb+spramt_tb) AS sptamtTB, (spramt_mrp_tp+spramt_tp) AS spramtTP, (oilamt_mrp_tp+oilamt_tp) AS oilamtTP, (oilamt_mrp_tb+ oilamt_tb) AS oilamtTB, tot_amt, tot_disc_amt, tot_ord_discamt, tax_amt, addition, deduction, net_amt " & _
                    "FROM sp_purch  WHERE " & CondStr & ""
             RepName = "SprPurReg"
        ElseIf GRepFormName = SprPurRet Then
            If FGrid.TextMatrix(List1, 1) = "Transfer" Then CondStr = "SP_Purch.V_Type = '" & SprPrTrfRet & "' And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then CondStr = "SP_Purch.V_Type = '" & SprPrRetCre & "' And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then CondStr = "SP_Purch.V_Type = '" & SprPrRetCsh & "' And "
            CondStr = CondStr + "SP_Purch.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and SP_Purch.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
            mQRY = "SELECT SP_Stock.docid,Part.Part_Name, SubGroup.Name, SP_Stock.Part_No, SP_Stock.Qty_Iss, SP_Stock.Rate, SP_Stock.V_No, mid(SP_Stock.DocID,9,5) as VPrefix, SP_Stock.V_Date " & _
            "FROM ((SP_Purch LEFT JOIN SP_Stock ON SP_Purch.DocID = SP_Stock.DocID) LEFT JOIN SubGroup ON SP_Purch.Party_Code = SubGroup.SubCode) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO"
            mQRY = mQRY & " where " & CondStr
            RepName = "SprPurRet"
        End If
    End Select
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub SprStkRep()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    'CondStr = " where SP_Stock.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and SP_Stock.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "#"
    If Check1(2).Value = Unchecked Then CondStr = " and left(SP_Stock.site_code,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then CondStr = " and SP_Stock.Part_No in (" & GridString3 & ")"
    Select Case GRepFormName
    Case SprStkReg
        mQRY = "SELECT SP_Stock.Part_No, '' as DocID,'' as SrlNo, null as V_Date, IIf(SP_Stock.Tax_YN=0,'Opening Taxable','Opening TaxPaid') as v_Prefix, 0 as V_No,'' as Job_DocID, IIf(SP_Stock.Tax_YN=0,sum(SP_Stock.Qty_Rec) - sum(SP_Stock.Qty_Iss),0) AS TPQtyRec, " & _
        "IIf(SP_Stock.Tax_YN=1,sum(SP_Stock.Qty_Rec) - sum(SP_Stock.Qty_Iss),0) AS TBQtyRec,0 AS TPQtyIss,0 AS TBQtyIss, " & _
        "'' AS SprPurPose, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        "WHERE SP_Stock.V_Date>= #" & Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy") & "#  and SP_Stock.V_Date<= #" & Format(DateAdd("D", -1, FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy") & "# " & _
        CondStr & _
        " GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name " & _
        " Union All " & _
        "SELECT SP_Stock.Part_No, SP_Stock.DocID,'Z' as SrlNo, SP_Stock.V_Date, Mid(SP_Stock.DocID,9,5) as v_Prefix, SP_Stock.V_No, SP_Stock.Job_DocID, IIf(SP_Stock.Tax_YN=0,SP_Stock.Qty_Rec,0) AS TPQtyRec, " & _
        "IIf(SP_Stock.Tax_YN=1,SP_Stock.Qty_Rec,0) AS TBQtyRec,IIf(SP_Stock.Tax_YN=0,SP_Stock.Qty_Iss,0) AS TPQtyIss,IIf(SP_Stock.Tax_YN=1,SP_Stock.Qty_Iss,0) AS TBQtyIss, " & _
        "IIf(SP_Stock.Purpose='P','PDI',IIf(SP_Stock.Purpose='F','Free Service',IIf(SP_Stock.Purpose='C','Chargeable',IIf(SP_Stock.Purpose='W','Warranty',IIf(SP_Stock.Purpose='O','Company Vehicle',IIf(SP_Stock.Purpose=' L','Complementary','')))))) AS SprPurPose, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        "where SP_Stock.V_Type<>'SXAO' and SP_Stock.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and SP_Stock.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# " & _
        CondStr & _
        ""
        RepName = "SprStkReg"
    Case SprStkSumm
        mQRY = "SELECT SP_Stock.Part_No, 0 AS TPQtyRec, 0 AS TBQtyRec, 0 AS TPQtyIss, 0 AS TBQtyIss, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Rec),0) AS TPQtyOpen, IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Rec),0) AS TBQtyOpen, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        "WHERE SP_Stock.V_Date>= #" & Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy") & "#  and SP_Stock.V_Date<= #" & Format(DateAdd("D", -1, FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy") & "# " & _
        CondStr & _
        " GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name " & _
        " Union All " & _
        "SELECT SP_Stock.Part_No, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Rec),0) AS TPQtyRec, IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Rec),0) AS TBQtyRec, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Iss),0) AS TPQtyIss,IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Iss),0) AS TBQtyIss, 0 AS TPQtyOpen, 0 AS TBQtyOpen, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        "WHERE SP_Stock.V_Type<>'SXAO' and SP_Stock.V_Date>= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#  and SP_Stock.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# " & _
        CondStr & _
        "GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name"
        
        RepName = "SprStkSumm"
    Case SprStkInHand
        mQRY = "SELECT SP_Stock.Part_No, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Rec)-Sum(SP_Stock.Qty_Iss),0) AS TPQty, IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Rec)-Sum(SP_Stock.Qty_Iss),0) AS TBQty, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        "WHERE SP_Stock.V_Date>= #" & Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy") & "#  and SP_Stock.V_Date<= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# " & _
        CondStr & _
        "GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name"

       RepName = "SprStkInHand"
    End Select
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub VehMoneyRectFunc()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(DGrid1, FGrid.TextMatrix(DGrid1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    CondStr = " where R.v_Date >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and R.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# "
    If FGrid.TextMatrix(List1, 1) <> "All" Then CondStr = CondStr & " and R.IFORM = '" & FGrid.TextMatrix(List1, 1) & "'"
    If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(R.site_code,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then CondStr = CondStr & " and R.v_type in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then CondStr = CondStr & " and R.Prov_Location in (" & GridString4 & ")"
    
    mQRY = "SELECT SubGroup.Name, R.DocId, R.V_Date,R.Narration,R.v_type, R.V_No, R.Prov_No, R.Prov_Date, R.AMOUNT, R.DDNo, R.DDDate, R.IFORM, SubGroup.PANNo " & _
    "FROM Rect R LEFT JOIN SubGroup ON R.PartyCode = SubGroup.SubCode"
    
    mQRY = mQRY + CondStr + " order by R.DocId"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "VehMoneyRect"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub WksRequisition()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
Dim mPurpose As String
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
End Select
    'P->PDI,F->Free Service, C->Chargable,W->Warranty,O->Company Vehicle,L->Complementary

    RepPrint = True
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    CondStr = " where SP_Stock.v_Date >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and SP_Stock.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# "
    If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(SP_Stock.site_code,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then CondStr = CondStr & " and Job_Card.Serv_Type in (" & GridString3 & ")"
    If FGrid.TextMatrix(List1, 1) = "UnClosed" Then CondStr = CondStr & " and (isnull(Job_Card.JobCloseDate) or Job_Card.JobCloseDate = '')"
    If FGrid.TextMatrix(List1, 1) = "Closed" Then CondStr = CondStr & " and (isnotnull(Job_Card.JobCloseDate) or Job_Card.JobCloseDate <> '')"
    If FGrid.TextMatrix(List2, 1) <> "All" Then CondStr = CondStr & " and SP_Stock.Purpose = '" & mPurpose & "'"
    CondStr = CondStr & " and SP_Stock.V_type in ('" & WksGenReq & "','" & WksReqWrt & "')"
            
    mQRY = "SELECT Part.Part_Name,SP_Stock.docid, SP_Stock.V_No, SP_Stock.V_Date, Job_Card.Job_No, Job_Card.Job_Date, Job_Card.JobCloseDate," & _
    "Job_Card.DocId_InvSpr, HisCard.RegNo, HisCard.Chassis, SP_Stock.Part_No, SP_Stock.Purpose, SP_Stock.Qty_Doc, SP_Stock.Qty_Iss," & _
    "SP_Stock.Qty_Ret, SP_Stock.Rate, SP_Stock.Amount, (SP_Stock.Claim_Div + SP_Stock.Claim_Site + SP_Stock.Claim_YearPrefix + SP_Stock.Claim_Type +  SP_Stock.Claim_No) as ClaimNo, SP_Stock.Claim_Date " & _
    "FROM ((SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO) LEFT JOIN Job_Card ON SP_Stock.Job_DocID = Job_Card.DocId)" & _
    "LEFT JOIN HisCard ON (Job_Card.CardNo = HisCard.CardNo)"
    mQRY = mQRY + CondStr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksReqReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub

Private Function SprABCAnalysis() As Boolean
On Error GoTo ELoop
Dim mQRY As String, CondStr As String

    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Function
    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Function
    If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Function
    If Val(FGrid.TextMatrix(Cat2, 1)) > Val(FGrid.TextMatrix(Cat1, 1)) Then
        MsgBox FGrid.TextMatrix(Cat2, 0) & ">" & FGrid.TextMatrix(Cat1, 0), vbOKOnly, "Validation"
        FGrid.SetFocus:  FGrid.Row = Cat2: FGrid.Col = 1
        RepPrint = False: Exit Function
    End If

    mQRY = "Select SP.Part_No,P.Part_Name,sum(SP.Net_Amt2) as NetAmt2 from SP_Stock SP Left Join Part P on SP.Part_No=P.Part_No " & _
        "where SP.Net_Amt2<>0 " & _
        " and SP.V_Date2>= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & _
        " and SP.V_Date2<= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & _
        " and mid(SP.Invoice_DocId,4,5) in ('" & SprSlCsh & "', '" & SprSlCre & "', '" & WksSlCsh & "', '" & WksSlCre & "') " & _
        " and SP.Purpose<>'W' " & _
        " Group by SP.Part_No,P.Part_Name Order by sum(Net_Amt2) Desc"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Function
    RepName = "SprABCRep"
    RepTitle = UCase(Me.CAPTION)
    SprABCAnalysis = False
    Exit Function
ELoop:
    MsgBox err.Description
End Function

Private Sub SprFSNAnalysis()
'On Error GoTo ELoop
'Dim mQRY As String, CondStr As String, TrnSQL$, TrnSQL2$
'Dim rsTemp As ADODB.Recordset, rsTemp2 As ADODB.Recordset, rsTemp3 As ADODB.Recordset
'Dim mRecQty As Double, mIssQty As Double, mStkVal As Double
'
'    RepPrint = True
'    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
'    If Val(FGrid.TextMatrix(Cat2, 1)) > Val(FGrid.TextMatrix(Cat1, 1)) Then
'        MsgBox FGrid.TextMatrix(Cat2, 0) & ">" & FGrid.TextMatrix(Cat1, 0), vbOKOnly, "Validation"
'        FGrid.SetFocus:  FGrid.Row = Cat2: FGrid.Col = 1
'        RepPrint = False: Exit Sub
'    End If
'
'    GSQL = "Select P.Part_No,P.Part_Name,P.Unit,P.MRP, P.TB_SRate, P.TP_SRate " & _
'        " from Part P " & _
'        " Order By P.Part_No"
'
'    TrnSQL = "SELECT SP.Part_No, sum(sp.Qty_Rec- (sp.Qty_Iss-sp.Qty_Ret)) as OpQty, " & _
'        " FROM SP_Stock SP " & _
'        " WHERE SP.V_Date1< " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & _
'        " Group by SP.Part_No"
'
'    mQRY = "Select SP.Part_No, Part_No+trim(str(Tax_YN))+trim(str(MRP_YN)) as SearchCode," & _
'        " sum(SP.Qty_Rec) as QtyRec " & _
'        " from SP_Stock SP " & _
'        " where SP.V_Date>= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & _
'        " and SP.V_Date<= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & _
'        " Group by SP.Part_No, Part_No+trim(str(Tax_YN))+trim(str(MRP_YN)) "
'
'    TrnSQL2 = "Select Part_No, Part_No+trim(str(Tax_YN))+trim(str(MRP_YN)) as SearchCode," & _
'        " sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as ClosQty " & _
'        " From Sp_Stock " & _
'        " where V_Date<= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & _
'        " Group by Part_No, Part_No+trim(str(Tax_YN))+trim(str(MRP_YN))"
'
'    Set GRs = New ADODB.Recordset
'    With GRs
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'    End With
'    Set GRs = GCn.Execute(GSQL)
'
'    Set rsTemp = New ADODB.Recordset
'    With rsTemp
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'    End With
'    Set rsTemp = GCn.Execute(TrnSQL)
'
'    Set rsTemp2 = New ADODB.Recordset
'    With rsTemp2
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'    End With
'    Set rsTemp2 = GCn.Execute(mQRY)
'
'
'    Set rsTemp3 = New ADODB.Recordset
'    With rsTemp3
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'    End With
'    Set rsTemp3 = GCn.Execute(TrnSQL2)
'
'    Set RstRep = New ADODB.Recordset
'    With RstRep
'        .Fields.Append "Part_No", adVarChar, 21, adFldIsNullable
'        .Fields.Append "Part_Name", adVarChar, 40, adFldIsNullable
'        .Fields.Append "Unit", adVarChar, 6, adFldIsNullable
'        .Fields.Append "OpQty", adDouble, 12, adFldIsNullable
'        .Fields.Append "RecQty", adDouble, 12, adFldIsNullable
'        .Fields.Append "IssQty", adDouble, 12, adFldIsNullable
'        .Fields.Append "ClosValue", adDouble, 12, adFldIsNullable
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockOptimistic
'        .Open
'    End With
'    If rsTemp.EOF = True And rsTemp2.EOF = True Then GoTo ELoop
'
'    Set GRs = New ADODB.Recordset
'    With GRs
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockReadOnly
'    End With
'    Set GRs = GCn.Execute(GSQL)
'
'    Do Until GRs.EOF
'        rsTemp.MoveFirst
'        rsTemp.FIND ("Part_No = '" & GRs!Part_No & "'")
'
'        rsTemp2.MoveFirst
'        rsTemp2.FIND ("Part_No = '" & GRs!Part_No & "'")
'
'        If rsTemp.EOF = False Or rsTemp2.EOF = False Then
'            With RstRep
'                mRecQty = 0
'                mIssQty = 0
'                .AddNew
'                .Fields("Part_No") = GRs!Part_No
'                .Fields("Part_Name") = GRs!Part_Name
'                .Fields("Unit") = GRs!Unit
'
'                'Non-MRP + TaxPaid
''                If rsTemp2!MRP_YN = 0 and rsTemp2!Tax_YN = 0 Then
'                rsTemp2.MoveFirst
'                rsTemp2.FIND ("SearchCode = '" & GRs!Part_No & "00")
'                If rsTemp2.EOF = False Then
'                    mRecQty = mRecQty + rsTemp2!RecQty
'                    mIssQty = mIssQty + rsTemp2!IssQty
'                End If
'
'                rsTemp3.MoveFirst
'                rsTemp3.FIND ("SearchCode = '" & GRs!Part_No & "00")
'                If rsTemp3.EOF = False Then
'                    mStkVal = mStkVal + Round(rsTemp3!ClosQty * GRs!TP_SRate, 2)
'                End If
'
'                'Non-MRP + Taxable
''                If rsTemp2!MRP_YN = 0 and rsTemp2!Tax_YN = 1 Then
'                rsTemp2.MoveFirst
'                rsTemp2.FIND ("SearchCode = '" & GRs!Part_No & "01")
'                If rsTemp2.EOF = False Then
'                    mRecQty = mRecQty + rsTemp2!RecQty
'                    mIssQty = mIssQty + rsTemp2!IssQty
'                End If
'                rsTemp3.MoveFirst
'                rsTemp3.FIND ("SearchCode = '" & GRs!Part_No & "01")
'                If rsTemp3.EOF = False Then
'                    mStkVal = mStkVal + Round(rsTemp3!ClosQty * GRs!TB_SRate, 2)
'                End If
'
'                'MRP + TaxPaid
''                If rsTemp2!MRP_YN = 0 and rsTemp2!Tax_YN = 0 Then
'                rsTemp2.MoveFirst
'                rsTemp2.FIND ("SearchCode = '" & GRs!Part_No & "10")
'                If rsTemp2.EOF = False Then
'                    mRecQty = mRecQty + rsTemp2!RecQty
'                    mIssQty = mIssQty + rsTemp2!IssQty
'                End If
'                rsTemp3.MoveFirst
'                rsTemp3.FIND ("SearchCode = '" & GRs!Part_No & "10")
'                If rsTemp3.EOF = False Then
'                    mStkVal = mStkVal + Round(rsTemp3!ClosQty * GRs!MRP, 2)
'                End If
'
'                'MRP + Taxable
''                If rsTemp2!MRP_YN = 0 and rsTemp2!Tax_YN = 1 Then
'                rsTemp2.MoveFirst
'                rsTemp2.FIND ("SearchCode = '" & GRs!Part_No & "11")
'                If rsTemp2.EOF = False Then
'                    mRecQty = mRecQty + rsTemp2!RecQty
'                    mIssQty = mIssQty + rsTemp2!IssQty
'                End If
'                rsTemp3.MoveFirst
'                rsTemp3.FIND ("SearchCode = '" & GRs!Part_No & "11")
'                If rsTemp3.EOF = False Then
'                    mStkVal = mStkVal + Round(rsTemp3!ClosQty * GRs!MRP, 2)
'                End If
'
'                .Fields("RecQty") = mRecQty
'                .Fields("IssQty") = mIssQty
'                .Fields("OpQty") = GRs!OpQty
'                .Fields("OpValue") = GRs!OpValue
'
'                .Fields("ClosValue") = Round(IIf(rsTemp1.EOF, 0, rsTemp1!IssQty) * GRs!Rate, 2)
'
'                .Update
'                mTotOpVal = mTotOpVal + GRs!OpValue
'                mTotRecVal = mTotRecVal + Round(IIf(rsTemp1.EOF, 0, rsTemp1!RecQty) * GRs!Rate, 2)
'                mTotIssVal = mTotIssVal + Round(IIf(rsTemp1.EOF, 0, rsTemp1!IssQty) * GRs!Rate, 2)
'            End With
'        End If
'        GRs.MoveNext
'    Loop
'    Set GRs = Nothing
'
'    Set RstRep = New Recordset
'    RstRep.CursorLocation = adUseClient
'    RstRep.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
'    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'    RepName = "SprABCRep"
'    RepTitle = UCase(Me.CAPTION)
'
'ELoop: If err.NUMBER <> 0 Then CheckError
'    Set GRs = Nothing
'    Set rsTemp = Nothing
'    Set rsTemp2 = Nothing
    
End Sub

Private Sub WksVehicleDiary()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String

    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub

    CondStr = " where Job_Card.Job_Date <= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and (isnull(Job_Card.JobCloseDate))"
    
    mQRY = "SELECT Job_Card.Job_Date, HisCard.RegNo,HisCard.name, HisCard.Chassis, Job_Card.Job_No, Job_Card.AtKMsHrs, Emp_Mast.Emp_Name, Job_Card.ExpDelDate, Job_Card.DocId, Job_Card.REMARK, Job_Demand.S_No, Job_Demand.Details " & _
    "FROM ((Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) LEFT JOIN Emp_Mast ON Job_Card.RecBy_Mechanic = Emp_Mast.Emp_Code) LEFT JOIN Job_Demand ON Job_Card.DocId = Job_Demand.Job_DocID"
    
    mQRY = mQRY + CondStr + " order by Job_Card.Job_Date desc"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksVehDiary"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub

Private Sub WksJobRegister()
On Error GoTo ELoop
Dim mQRY As String, CondStr As String
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    CondStr = " where Job_Card.Job_Date >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and Job_Card.Job_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# "
    If Check1(2).Value = Unchecked Then CondStr = CondStr & " and left(Job_Card.site_code,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then CondStr = CondStr & " and Job_Card.Serv_Type in (" & GridString3 & ")"
    If FGrid.TextMatrix(List1, 1) = "UnClosed" Then CondStr = CondStr & " and isnull(Job_Card.JobCloseDate)"
    If FGrid.TextMatrix(List1, 1) = "Closed" Then CondStr = CondStr & " and isnotnull(Job_Card.JobCloseDate)"
                
    mQRY = "SELECT Job_Card.DocId_InvSpr, Job_Card.DocId_InvLab, Job_Card.Job_Date, Job_Card.Job_No, Job_Card.DocId, Job_Card.JobCloseDate, HisCard.Name, Service_Type.Serv_Desc, Job_Card.NetLab_Amt, HisCard.RegNo, HisCard.Chassis, Job_Lab.Lab_Rate, Job_Lab.War_Lab_Rate, Job_Lab.LabourAmt, Job_Lab.Major_YN,0 as Net_Amt,0 as Total_Amt,0 as  Purpose,0 as V_No " & _
    "FROM ((Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) LEFT JOIN Job_Lab ON Job_Card.DocId = Job_Lab.Job_DocID) LEFT JOIN Service_Type ON Job_Card.Serv_Type = Service_Type.Serv_Type " & _
    "UNION ALL SELECT Job_Card.DocId_InvSpr, Job_Card.DocId_InvLab, Job_Card.Job_Date, Job_Card.Job_No, Job_Card.DocId, Job_Card.JobCloseDate, HisCard.Name, Service_Type.Serv_Desc, Job_Card.NetLab_Amt, HisCard.RegNo, HisCard.Chassis, 0 AS Lab_Rate, 0 AS War_Lab_Rate, 0 AS LabourAmt, 0 AS Major_YN, SP_Stock.Net_Amt, SP_Sale.Total_Amt, SP_Stock.Purpose, SP_Sale.V_No " & _
    "FROM (((Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) LEFT JOIN Service_Type ON Job_Card.Serv_Type = Service_Type.Serv_Type) LEFT JOIN SP_Sale ON Job_Card.DocId_InvSpr = SP_Sale.DocID) LEFT JOIN SP_Stock ON SP_Sale.DocID = SP_Stock.Invoice_DocId "

    mQRY = mQRY + CondStr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "WksJobReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub

'Private Function GetString() As String
'Dim Rst As ADODB.Recordset
'Dim LocStr As String, Tss As String
'Set Rst = GCn.Execute("select distinct Supl_Loca from Part where part_no in (select distinct part_no from sp_order1)")
'If Rst.RecordCount > 0 Then
'Do Until Rst.EOF
'Tss = IIf(IsNull(Rst!supl_loca), "", Rst!supl_loca)
'    If Len(Tss) > 0 Then
'        Do While Not Tss = ""
'            If InStr(1, LocStr, left(Tss, 4), vbBinaryCompare) = False Then
'                LocStr = LocStr + IIf(LocStr = "", "'" + left(Tss, 4) + "'", "," + "'" + left(Tss, 4) + "'")
'            End If
'            Tss = Mid(Tss, 5, Len(Tss))
'        Loop
'    End If
'    Rst.MoveNext
'Loop
'Else
'    LocStr = ""
'End If
'GetString = LocStr
'Set Rst = Nothing
'End Function
