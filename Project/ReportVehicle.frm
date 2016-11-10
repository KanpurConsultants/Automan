VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form ReportVehicle 
   BackColor       =   &H00C8E8DA&
   Caption         =   "ReprtForm"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11820
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   11820
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Speed Print"
      DownPicture     =   "ReportVehicle.frx":0000
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
      Index           =   1
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Print Report"
      Top             =   6075
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7620
      TabIndex        =   16
      Top             =   -30
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   90
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   15
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
         BackColor       =   16379351
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3942
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Print"
      DownPicture     =   "ReportVehicle.frx":3132
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
      Index           =   0
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   6075
      Width           =   1290
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
      DownPicture     =   "ReportVehicle.frx":6264
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
      Left            =   6645
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
      TabIndex        =   14
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
         TabIndex        =   15
         Top             =   0
         Width           =   4470
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
Attribute VB_Name = "ReportVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CellBackColLeave As String = &HFFFFFF
Private Const CellBackColEnter As String = &HFFFFC0
Private Const CellBackColLeave1 As String = &HEDF7FE
Private Const CellBackColEnter1 As String = &HFFFFC0
'Modishekhar 17 mar
Dim FormulaStr1 As String, FormulaStr2 As String, FormulaStr3 As String, FormulaStr4 As String
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
Private Const ChsRecReg As Byte = 1
Private Const PurchReg As Byte = 2
Private Const VehInTrans As Byte = 3
Private Const VehBooking As Byte = 4
Private Const RetSaleRep As Byte = 5
Private Const DailyRetail As Byte = 6 'Own
Private Const VehSaleReg As Byte = 7 '6
Private Const DelChaReg As Byte = 8 '7
Private Const AddFitRep As Byte = 9 '8
Private Const VehStkReg As Byte = 10 '9
Private Const VehStkBank As Byte = 11
Private Const VehStkHold As Byte = 12
Private Const VehSalePurRep As Byte = 14
Private Const VehTarget As Byte = 15
Private Const VehQuot As Byte = 16
Private Const DailyRetailTelco As Byte = 17 'for Telco
Private Const VehSaleCancelReg As Byte = 18 'for Telco
Private Const SalesManPenAmt As Byte = 19
Private Const OutPayRep As Byte = 20
Private Const VehIssReg As Byte = 23
Private Const VehFollowUp As Byte = 25
Private Const IncomeTaxReg As Byte = 26
Private Const SubVentionClaimReg As Byte = 27
Private Const ChequePaymentRegister As Byte = 30

'***********************************************************
Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const List1 As Byte = 2
Private Const List2 As Byte = 3
Private Const List3 As Byte = 4
Private Const List4 As Byte = 5

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
Dim SpeedPrnVehSale As Boolean

Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub BTNPRINT_Click(Index As Integer)
On Error GoTo ERRORHANDLER
SubRep1 = False
RepPrint = True
Select Case GRepFormName
    Case VehQuot
        VehQuotProc
        If RepPrint = False Then Exit Sub
    Case ChsRecReg
        ChsRecRegProc
        If RepPrint = False Then Exit Sub
    Case PurchReg
        PurchRegProc
        If RepPrint = False Then Exit Sub
    Case VehInTrans
        VehInTransProc
        If RepPrint = False Then Exit Sub
    Case VehBooking
        VehBookingProc
        If RepPrint = False Then Exit Sub
    Case RetSaleRep
        RetSaleRepProc
        If RepPrint = False Then Exit Sub
    Case VehSaleReg
        If Index = 1 Then SpeedPrnVehSale = True Else SpeedPrnVehSale = False
        VehSaleRegProc
        If RepPrint = False Then Exit Sub
    Case VehSaleCancelReg
        VehSaleCancelRegProc
        If RepPrint = False Then Exit Sub
    Case DelChaReg
        DelCharegProc
        If RepPrint = False Then Exit Sub
    Case AddFitRep
        AddFitRepProc
        If RepPrint = False Then Exit Sub
    Case VehStkReg
        VehStkRegProc
        If RepPrint = False Then Exit Sub
    Case DailyRetail
        DailyRetailProc
    Case DailyRetailTelco
        DailyRetailProcTelco
'******************************************************************
    Case VehStkBank
        VehStkBankProc
    Case VehStkHold
        VehStkHoldProc
'**********************************************************************
    Case VehSalePurRep
        VehSalePurRepProc
    Case VehTarget
        VehTargetProc
    Case SalesManPenAmt
        SalesManPenAmtProc
    Case OutPayRep
        ProcVehicleBillWiseOutstandingFIFO
        'OutPayRepProc
    Case VehIssReg
        VehIssRegProc
        If RepPrint = False Then Exit Sub
    Case VehFollowUp
        VehFollowUpProc
        If RepPrint = False Then Exit Sub
    Case IncomeTaxReg
        IncomeTaxRegProc
        If RepPrint = False Then Exit Sub
    Case SubVentionClaimReg
        SubVentionClaimRegProc
    
    Case ChequePaymentRegister
        ChequePaymentRegisterProc
End Select
If RepPrint = False Then Exit Sub
If SpeedPrnVehSale = True Then Exit Sub

CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")

rpt.Database.SetDataSource RstRep
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
rpt.ReadRecords
Set RstRep = Nothing

Call Formulas
Call Report_View(rpt, RepTitle, , False)
'Set rpt = Nothing   auto by report_view function
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
WinSetting Me, 6885, 11500
   Global_Grid
   TopCtrl1.TopText2 = "Add"
'   If Mid(UserPermission(Me.Name), 4, 1) = "*" Then BTNPRINT.Enabled = False
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
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
    BTNPRINT(1).Visible = False
    If UCase(Trim(FGrid.TextMatrix(List2, 1))) = "SALESMANWISE" Then
        Grid3Sql = "Select '' as O,Emp_Name as Name,Emp_Code as Code from Emp_Mast where Emp_Type=0"
        GridInitialise 3, Grid3Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "PARTYWISE" Then
        Grid3Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
                "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
                "Where " & _
                "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
                " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        GridInitialise 3, Grid3Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "CITYWISE" Then
        Grid3Sql = "Select '' as O,CityName as CityName,CityCode as Code from City where Site_Code='" & PubSiteCode & "'"
        GridInitialise 3, Grid3Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "FINANCIERGRP" Then
        Grid3Sql = "Select '' as O,FinGrpName as FinGrpName,FinGrpCode as Code from FinGroup where Site_Code='" & PubSiteCode & "'"
        GridInitialise 3, Grid3Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "FINANCIERNAME" Then
        Grid3Sql = "Select '' as O,ContractFinance.FinName + '  ' + City.CityName as FinName,ContractFinance.FinCode as Code from ContractFinance Left Join City on ContractFinance.City=City.CityCode where ContractFinance.Site_Code='" & PubSiteCode & "'"
        GridInitialise 3, Grid3Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "FORMTYPE" Then
        Grid3Sql = "Select DISTINCT '' as O,Veh_Order.Form_Code  as FormCode,Veh_Order.Form_Code  as Code from Veh_Order Left Join TaxForms on Veh_Order.Form_Code=TaxForms.Form_Code where right(Veh_Order.Ord_SiteCode,1)='" & PubSiteCode & "'"
        GridInitialise 3, Grid3Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "INSU.AUTH." Then
        Grid3Sql = "Select DISTINCT '' as O,FinName  as InsurancerName,FinCode  as Code from ContractFinance where FinCatg=3"
        GridInitialise 3, Grid3Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "ALL" And UCase(Me.CAPTION) = "VEHICLE SALE REGISTER" Then
        BTNPRINT(1).Visible = True: BTNPRINT(1).Enabled = True
End If
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
Dim RsTemp As ADODB.Recordset
FGrid.CellBackColor = CellBackColLeave
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)

Select Case FGrid.Row
    Case List1
        Select Case GRepFormName
            Case OutPayRep
                ListArray = Array("Financer", "Salesman")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case PurchReg
                ListArray = Array("Detail", "Summary")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VehBooking
                ListArray = Array("All", "Pending", "Supplied")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case RetSaleRep
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VehSaleReg
                ListArray = Array("Summary", "Detailed")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case DelChaReg
                ListArray = Array("All", "Delivered", "Pending")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case VehStkReg
                ListArray = Array("With Sale", "W/O Sale")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VehStkHold
                ListArray = Array("With Sale", "W/O Sale")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case ChsRecReg  'vijay for work shop '16/11/02
                ListArray = Array("All", "Pending")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VehSalePurRep
                ListArray = Array("V-No", "Telco Inv-No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VehIssReg
                ListArray = Array("Issued", "Recieved")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
        End Select
    Case List2
        Select Case GRepFormName
            Case VehBooking
                ListArray = Array("Party", "OrderNo", "Model")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case PurchReg, VehStkHold
                Set RsTemp = GCn.Execute("select BMS.BMS_Name as Name,BMS.BMS_Code As Code from BMS order by BMS.BMS_Name")
                ListView_Items_RecordSet_Local ListView, TxtGrid, Index, RsTemp
            Case VehSaleReg
                ListArray = Array("SalesManWise", "PartyWise", "CityWise", "FinancierGrp", "FinancierName", "FormType", "Insu.Auth.", "All")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 8)
         End Select
    Case List3
        Select Case GRepFormName
            Case PurchReg
                ListArray = Array("Voucher No", "Telco Bill No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
        End Select
    Case List4
        Select Case GRepFormName
            Case PurchReg
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
        End Select
End Select
Set RsTemp = Nothing
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
    Case List1, List2, List3, List4
        If FGrid.Row = List2 And (GRepFormName = PurchReg Or GRepFormName = VehStkHold) Then
            ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), 4000, 4000
            ListView.ColumnHeaders(1).width = 3500
            ListView.ColumnHeaders(2).width = 0
        Else
            ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
        End If
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
' Case Cat1
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
'    Case Cat2
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
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
            
        Case List1, List2, List3, List4
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
Select Case FGrid.Row
        Case Cat1, Cat2, Cat3, Cat4, Cat5
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
        
        Case List2, List1
            If FGrid.Row = List2 And (GRepFormName = PurchReg Or GRepFormName = VehStkHold) Then
                If TxtGrid(0).TEXT <> "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
                    FGrid.TextMatrix(FGrid.Row, 2) = ListView.SelectedItem.SubItems(1)
                End If
            Else
                If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            End If
            
        Case List3
            If TxtGrid(0).TEXT <> "" Then
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
                Select Case TxtGrid(0).TEXT
                Case "SalesManWise"
                    Grid3Sql = "select '' as O,Emp_Name as RepresentativeName,Emp_Code  as code from Emp_Mast where Emp_Type=0 order by Emp_Name"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
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
                   Grid3Sql = "select '' as O,Form_Desc as Form_Desc,Form_Code  as code from TaxForms order by Form_Desc"
                      GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
               Case "All"
                    GridSel(3).Visible = False: Check1(3).Visible = False
    
                End Select
            End If
                                
        
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
End Select

Select Case FGrid.Row
    Case List4
        If (GRepFormName = PurchReg) Then
            If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            Ini_Grid
        End If
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
BTNPRINT(0).left = (Pic.width - (BTNPRINT(0).width + BTNEXIT.width)) / 2: BTNPRINT(0).top = Pic.top + 10
BTNPRINT(1).left = (BTNPRINT(0).left - BTNPRINT(1).width): BTNPRINT(1).top = Pic.top + 10
BTNEXIT.left = BTNPRINT(0).left + BTNPRINT(0).width: BTNEXIT.top = Pic.top + 10

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
        Case Date1, Date2, List1, List2, List3, List4, Cat1, Cat2, Cat3, Cat4, Cat5
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
Dim I As Integer
    Select Case FGrid.Row
        Case Cat1, Cat2, Cat3, Cat4, Cat5
                 Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
        Case Date1, Date2, List1, List2, List3, List4
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
        Case Date1, Date2, List1, List2, List3, List4, Cat1, Cat2, Cat3, Cat4, Cat5
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
            If (Len(formulastr) + Len(GridSel(Gridindex).TextMatrix(GridRow, 2))) < 250 Then
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
Case VehQuot
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight

            .TextMatrix(Date1, 1) = IIf(.TextMatrix(Date2, 1) = "", PubStartDate, .TextMatrix(Date1, 1))
            .TextMatrix(Date2, 1) = IIf(.TextMatrix(Date2, 1) = "", PubLoginDate, .TextMatrix(Date2, 1))
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,ProfessionName as Profession,ProfessionCode as code from Profession order by ProfessionName"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,PurposeName as Purpose,PurposeCode as code from Purpose order by PurposeName"
        GridInitialise 4, Grid4Sql

Case ChsRecReg    'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "All/Pending": .RowHeight(List1) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
            "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
            "Where " & _
            "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        GridInitialise 3, Grid3Sql
Case VehBooking    'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "All/Pending/Suppl.": .RowHeight(List1) = GridRowHeight
           .TextMatrix(List2, 0) = "sort On": .RowHeight(List2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
           .TextMatrix(List2, 1) = "Party"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        
        If PubSiebelActiveYn = 1 Then
            Grid3Sql = "select '' as O,Model.Model + ' ' + ModelGrp_Name + ' ' + Col_desc As Model_Description,Model.Model As Code from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code order by Model.Model"
              GridInitialise 3, Grid3Sql
        Else
            Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
            GridInitialise 3, Grid3Sql
        End If
        Grid4Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
            "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
            "Where  " & _
            "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        GridInitialise 4, Grid4Sql
    
Case VehSaleReg    'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Summary/Detailed": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Type": .RowHeight(List2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Summary"
            .TextMatrix(List2, 1) = "PartyWise"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        If UCase(left(PubComp_Name, 4)) = "ENAR" Then
            Grid2Sql = "SELECT Distinct '' AS O, Sales_Desc As Sales_Description,  Sales_Desc AS Code,Vehicle_Type as Veh_Type FROM Model ORDER BY Vehicle_Type"
            GridInitialise 2, Grid2Sql
            GridSel(2).ColWidth(1) = 3000 ': GridSel(2).ColWidth(2) = 1250: GridSel(2).ColWidth(3) = 0
        Else
            Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
            GridInitialise 2, Grid2Sql
        End If
        Grid3Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
            "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
            "Where  " & _
            "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        GridInitialise 3, Grid3Sql
        If PubSiebelActiveYn = 1 Then
            Grid4Sql = "select '' as O,Model.Model + ' ' + ModelGrp_Name + ' ' + Col_desc As Model_Description,Model.Model As Code from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code order by Model.Model"
            GridInitialise 4, Grid4Sql
        Else
            Grid4Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
            GridInitialise 4, Grid4Sql
        End If
    Case VehSaleCancelReg
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
        
Case VehInTrans, AddFitRep  'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date2: mLastRow = Date2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Model.Model As ModelName,Model.Model As Code from Model order by Model.Model"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,Prod_Name As ItemName,Prod_Code As Code from Veh_AMDModel order by Prod_NAme"
        GridInitialise 4, Grid4Sql
Case DailyRetail, DailyRetailTelco, VehTarget 'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "For Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate - 1
        End With
        mFirstRow = Date1: mLastRow = Date1: mHelpGridNo = 3
        'Grid1Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        
Case RetSaleRep    'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Payment Detail(Y/N)": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Yes"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Model.Model,Model.Model As Code from Model order by Model.Model"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql

Case DelChaReg    'vijay Vehicle
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "All/Deliver/Pending": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
       End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        If PubSiebelActiveYn = 1 Then
            Grid2Sql = "select '' as O,Model.Model + ' ' + ModelGrp_Name + ' ' + Col_desc As Model_Description,Model.Model As Code from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code order by Model.Model"
            GridInitialise 2, Grid2Sql
        Else
            Grid2Sql = "select '' as O,Model.Model As Model,Model.Model As Code from Model order by Model.Model"
            GridInitialise 2, Grid2Sql
        End If
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
    
    Case PurchReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Report Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Select Category": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Sort On": .RowHeight(List3) = GridRowHeight
            .TextMatrix(List4, 0) = "Model Group": .RowHeight(List4) = GridRowHeight
            
            .TextMatrix(Date1, 1) = IIf(.TextMatrix(Date2, 1) = "", PubStartDate, .TextMatrix(Date1, 1))
            .TextMatrix(Date2, 1) = IIf(.TextMatrix(Date2, 1) = "", PubLoginDate, .TextMatrix(Date2, 1))
            .TextMatrix(List1, 1) = "Detail"
            .TextMatrix(List2, 1) = "All"
            .TextMatrix(List3, 1) = "Voucher No"
            .TextMatrix(List4, 1) = IIf(left(.TextMatrix(List4, 1), 1) = "Y", "Yes", "No")
        End With
        mFirstRow = Date1: mLastRow = List4: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        If UCase(left(PubComp_Name, 4)) = "ENAR" Then
            Grid2Sql = "select Distinct '' as O,Sales_Desc As Sales_Description, Sales_Desc As Code, Vehicle_Type As Veh_Type from Model order by Vehicle_Type"
            GridInitialise 2, Grid2Sql
            GridSel(2).ColWidth(1) = 3000
        Else
            Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
            GridInitialise 2, Grid2Sql
        End If
        If UCase(left(FGrid.TextMatrix(List4, 1), 1)) = "N" Then
            If PubSiebelActiveYn = 1 Then
                Grid3Sql = "select '' as O,Model.Model + ' ' + ModelGrp_Name + ' ' + " & xIsNull("Col_desc", "") & " As Model_Description,Model.Model As Code from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code order by Model.Model"
                GridInitialise 3, Grid3Sql
            Else
                Grid3Sql = "select '' as O,Model.Model  As Model_Description,Model.Model As Code from Model order by Model.Model"
                GridInitialise 3, Grid3Sql
            End If
        Else
            Grid3Sql = "select '' as O,Mg.ModelGrp_Name  As Model_Group,Mg.ModelGrp_Code As Code from Model_Grp MG order by Mg.ModelGrp_Name"
            GridInitialise 3, Grid3Sql
        End If
        Grid4Sql = "select '' as O,Subgroup.Name As Party_Name,Subgroup.SubCode As Code from SubGroup order by SubGroup.Name"
        GridInitialise 4, Grid4Sql

Case VehSalePurRep    'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Short On": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "V-No"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,BMS.BMS_Name As Category,BMS.BMS_Code As Code from BMS order by BMS.BMS_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select distinct '' as O, " & cUCase("left(PBILL_No,1)") & " as Prefix, " & cUCase("left(PBILL_No,1)") & " as Code From Veh_Stock"
        GridInitialise 4, Grid4Sql

Case VehStkHold    'vijay WKS 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "As On Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(List1, 0) = "Stock Scope ": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Select BMS": .RowHeight(List2) = GridRowHeight
            .TextMatrix(Cat1, 0) = "Holding > ": .RowHeight(Cat1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "With Sale"
            .TextMatrix(List2, 1) = "All"
            .TextMatrix(Cat1, 1) = "0"
        End With
        mFirstRow = Date1: mLastRow = Cat1: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
            Grid3Sql = "select '' as O, God_Name As Godown, God_Code As Code From Godown Where Appli_For=1 Order by God_Name"
            GridInitialise 3, Grid3Sql
        Else
            Grid3Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
                "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
                "Where  " & _
                "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
                " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
            GridInitialise 3, Grid3Sql
        
        End If
        
        Grid4Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 4, Grid4Sql
        
Case VehStkBank  'vijay Vehicle 16/11/02
        With FGrid
            .TextMatrix(Date1, 0) = "As On Date": .RowHeight(Date1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date1: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        If PubSiebelActiveYn = 1 Then
            Grid3Sql = "select '' as O,Model.Model + ' ' + ModelGrp_Name + ' ' + Col_desc As Model_Description,Model.Model As Code from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code order by Model.Model"
            GridInitialise 3, Grid3Sql
        Else
            Grid3Sql = "select '' as O,Model.Model As Model,Model.Model As Code from Model order by Model.Model"
            GridInitialise 3, Grid3Sql
        End If
        
        Grid4Sql = "select '' as O, God_Name As Godown, God_Code As Code From Godown Where Appli_For=1 Order by God_Name"
        GridInitialise 4, Grid4Sql
        
'        Grid4Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
'            "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
'            "Where FirmCode = '" & PubFirmCode & "' and " & _
'            "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
'            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
'        GridInitialise 4, Grid4Sql
Case VehStkReg    'vijay Vehicle
        With FGrid
            .TextMatrix(Date1, 0) = "As On Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(List1, 0) = "Sale Scope": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "With Sale"
       End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        If PubSiebelActiveYn = 1 Then
            Grid3Sql = "select '' as O,Model.Model + ' ' + ModelGrp_Name + ' ' + Col_desc As Model_Description,Model.Model As Code from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code order by Model.Model"
              GridInitialise 3, Grid3Sql
        Else
            Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
            GridInitialise 3, Grid3Sql
        End If
        Grid4Sql = "select '' as O,Col_Desc as Colour,Col_Code as code from ColMast order by Col_Desc"
        GridInitialise 4, Grid4Sql
        
Case SalesManPenAmt
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
       End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,emp_name as name,Emp_code as code from emp_mast where emp_type = 0  order by Emp_name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        
Case OutPayRep
        With FGrid
            .TextMatrix(Date1, 0) = "As on Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Billing Date ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Group on ": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Financer"
            '.TextMatrix(Date2, 1) = PubLoginDate
       End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,emp_name as name,Emp_code as code from emp_mast where emp_type = 0  order by Emp_name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        Grid4Sql = "select '' as O,F.FinName + ', ' + ISNULL(C.CityName, '') AS Financer, F.FinCode AS Code  FROM ContractFinance F LEFT JOIN City C ON C.CityCode = F.City  ORDER BY F.FinName "
        GridInitialise 4, Grid4Sql
        
 Case VehIssReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Trxn.Type": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Issued"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
            "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
            "Where  SubGroup.AliasYN <>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select distinct '' as O,ChassisNo As Chassis,ChassisNo As Code from Veh_Stock "
        GridInitialise 4, Grid4Sql
        
        
 Case VehFollowUp
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(Cat1, 0) = "I Service": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Cat2, 0) = "II Service": .RowHeight(Cat2) = GridRowHeight
            .TextMatrix(Cat3, 0) = "III Service": .RowHeight(Cat3) = GridRowHeight
            .TextMatrix(Cat4, 0) = "IV Service": .RowHeight(Cat4) = GridRowHeight
            .TextMatrix(Cat5, 0) = "V Service": .RowHeight(Cat5) = GridRowHeight
            
        End With
        mFirstRow = Date1: mLastRow = Cat5: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql
        
 Case IncomeTaxReg
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
        
        Grid3Sql = "select '' as O,CityName as City,Citycode  as code from City order by CityName"
        GridInitialise 3, Grid3Sql
Case SubVentionClaimReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
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

        Grid4Sql = "select '' as O,ModelGrp_Name,ModelGrp_Code from Model_Grp order by ModelGrp_Name"
        GridInitialise 4, Grid4Sql

Case ChequePaymentRegister
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O, Sg.Name As Bank_Name, Sg.SubCode As Code from SubGroup Sg Where " & xIsNull("Sg.ChequeReportName", "") & " <> '' order by Sg.Name"
        GridInitialise 3, Grid3Sql
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
Case SubVentionClaimReg, ChequePaymentRegister
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next

Case VehBooking
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("SortOrder")
              If FGrid.TextMatrix(List2, 1) = "Party" Then
                   rpt.FormulaFields(I).TEXT = "'P'"
              ElseIf FGrid.TextMatrix(List2, 1) = "OrderNo" Then
                   rpt.FormulaFields(I).TEXT = "'O'"
              Else
                   rpt.FormulaFields(I).TEXT = "'M'"
              End If
        End Select
    Next

Case RetSaleRep
    
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("Ledger")
              If FGrid.TextMatrix(List1, 1) = "Yes" Then
                   rpt.FormulaFields(I).TEXT = "1"
              Else
                   rpt.FormulaFields(I).TEXT = "0"
              End If
        End Select
    Next

Case ChsRecReg, PurchReg, VehInTrans, VehQuot, AddFitRep, VehInTrans, VehSalePurRep, IncomeTaxReg
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next
Case DailyRetail, DailyRetailTelco, VehStkReg, VehTarget, VehStkBank
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("RepTitle")
                rpt.FormulaFields(I).TEXT = "'Daily Sales Report-'+ '" & Format(FGrid.TextMatrix(Date1, 1), "mmm-yyyy") & "' "
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'As On Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' "
            Case UCase("list1")
                rpt.FormulaFields(I).TEXT = " '" & FGrid.TextMatrix(List1, 1) & "'"
        End Select
    Next
Case VehSaleReg  'Vijay
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("List3")
                rpt.FormulaFields(I).TEXT = " '" & FGrid.TextMatrix(List3, 1) & "'"
            Case UCase("CAmt")
                rpt.FormulaFields(I).TEXT = " '" & RstRep1!Cancel_Amt & "'"
            Case UCase("CTax")
                rpt.FormulaFields(I).TEXT = " '" & RstRep1!CancelTax_Amt & "'"
            Case UCase("CTot")
                rpt.FormulaFields(I).TEXT = " '" & RstRep1!CancelTOT_Amt & "'"
            Case UCase("TOTCaption")
                rpt.FormulaFields(I).TEXT = " '" & pubTOTCaption & "'"
        End Select
    Next
Case VehSaleCancelReg
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next

Case DelChaReg  'Vijay
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("List1")
                rpt.FormulaFields(I).TEXT = "'' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Type'"
        End Select
    Next
Case VehStkBank 'lps
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'As on Date :'+  '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("DateTo")
                rpt.FormulaFields(I).TEXT = "'" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        End Select
    Next
Case VehStkHold   'Vijay
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'As On Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + '' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("List1")
                rpt.FormulaFields(I).TEXT = "'' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Stock '"
        End Select
    Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub VehQuotProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
     
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Veh_Quot.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Quot.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Quot.Site_Code,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  left(Veh_Quot.Site_Code,1)  ='" & PubSiteCode & "' "
    End If
    
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Quot.DocId,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Veh_Quot.Profession in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Veh_Quot.PURPOSE in (" & GridString4 & ")"
    
      
    mQry = "SELECT Veh_Quot.DEL_DATE, Veh_Quot.EXP_DATE, Veh_Quot1.MODEL, Veh_Quot1.QTY, Veh_Quot1.RATE, Veh_Quot1.RSO_WORK, Veh_Quot1.AMOUNT, Profession.ProfessionName, Purpose.PurposeName, " & _
    "ProspectiveCust.NPrefix, ProspectiveCust.Name, ProspectiveCust.NSuffix,ProspectiveCust.FPrefix, ProspectiveCust.FName, ProspectiveCust.Add1, ProspectiveCust.Add2,ProspectiveCust.Add3, ProspectiveCust.AREA, City.CityName, Veh_Quot.V_Date, Veh_Quot.DocId,  Veh_Quot.V_No, Reffered.RefName, Emp_Mast.Emp_Name, ContractFinance.FinName, Veh_Quot.Call_Status, Veh_Quot.AMOUNT " & _
    "FROM ((((((((Veh_Quot LEFT JOIN ProspectiveCust ON Veh_Quot.Party_Code = ProspectiveCust.Cust_Code)" & _
    "LEFT JOIN City ON ProspectiveCust.CityCode = City.CityCode)" & _
    "LEFT JOIN Profession ON Veh_Quot.Profession = Profession.ProfessionCode) " & _
    "LEFT JOIN Purpose ON Veh_Quot.PURPOSE = Purpose.PurposeCode)" & _
    "LEFT JOIN Reffered ON Veh_Quot.REF_CODE = Reffered.RefCode)" & _
    "LEFT JOIN Emp_Mast ON Veh_Quot.REP_CODE = Emp_Mast.Emp_Code)" & _
    "LEFT JOIN ContractFinance ON Veh_Quot.FB_CODE = ContractFinance.FinCode) " & _
    "LEFT JOIN Veh_Quot1 ON Veh_Quot.DocId = Veh_Quot1.DocId ) "

    mQry = mQry & Condstr & "  Order By Veh_Quot.Profession,Veh_Quot.PURPOSE"
      
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "VehQuotReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub ChsRecRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'   If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Veh_Stock.Chassis_RctDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Stock.Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Stock.Chassis_RctSiteCode,1) in (" & GridString1 & ")"
    
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(Veh_Stock.Chassis_RctSiteCode,1) ='" & PubSiteCode & "' "
    End If
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_Stock.Chassis_RctDivCode in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Veh_Stock.PartyCode in (" & GridString3 & ")"
    If FGrid.TextMatrix(List1, 1) = "Pending" Then
        Condstr = Condstr + " and Veh_Stock.Pur_DocId =''"
    End If
    
    mQry = "SELECT ( " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  - Veh_Stock.Chassis_RctDate) AS Age,Veh_Stock.Chassis_RctDate,Veh_Stock.Srv_BookNo,godown.god_name, Veh_Stock.Chassis_RctDocNo, SubGroup.Name, Veh_Stock.Model, Veh_Stock.ChassisNo, Veh_Stock.EngineNo, Veh_Stock.SDM_STM_NO, Veh_Stock.Chassis_RctSiteCode, ColMast.Col_Desc" & _
           " FROM ((Veh_Stock LEFT JOIN SubGroup ON Veh_Stock.PartyCode = SubGroup.SubCode) LEFT JOIN godown ON Veh_Stock.godown = godown.god_code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code "

'    mQRY = "SELECT ( # " & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & " # - Veh_Stock.Chassis_RctDate) AS Age,Veh_Stock.Chassis_RctDate,Veh_Stock.Srv_BookNo,godown.god_name, Veh_Stock.Chassis_RctDocNo, SubGroup.Name, Veh_Stock.Model, Veh_Stock.ChassisNo, Veh_Stock.EngineNo, Veh_Stock.SDM_STM_NO, Veh_Stock.Chassis_RctSiteCode, ColMast.Col_Desc" & _
'           " FROM ((Veh_Stock LEFT JOIN SubGroup ON Veh_Stock.PartyCode = SubGroup.SubCode) LEFT JOIN godown ON Veh_Stock.godown = godown.god_code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code" & _
'           " Where Veh_Stock.Pur_DocId =''"

    mQry = mQry + Condstr
      
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "ChsRecReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VehIssRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If FGrid.TextMatrix(List1, 1) = "Issued" Then
        Condstr = " where Veh_Stock.Sal_VDate >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Stock.Sal_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and Veh_Stock.Sal_VType='V_TRF'"
    ElseIf FGrid.TextMatrix(List1, 1) = "Recieved" Then
        Condstr = " where Veh_Stock.Chassis_RctDate >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Stock.Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  and Veh_Stock.RectType='T'"
    End If
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Veh_Stock.Sal_Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  left(Veh_Stock.Sal_Site_Code,1)  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_Stock.TrfParty in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Stock.Sal_DocId,1) in (" & GridString3 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "Issued" Then
        mQry = "SELECT Veh_Stock.Sal_DocID,Veh_Stock.Model, Veh_Stock.ChassisNo, Veh_Stock.EngineNo,ColMast.Col_Desc,SubGroup.Name as PartyName,Veh_Stock.Sal_VDate,Veh_Stock.Remarks,Veh_Stock.Sal_VNo" & _
           " FROM (Veh_Stock LEFT JOIN SubGroup ON Veh_Stock.TrfParty = SubGroup.SubCode) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code "
    ElseIf FGrid.TextMatrix(List1, 1) = "Recieved" Then
        mQry = "SELECT Veh_Stock.Sal_DocID,Veh_Stock.Model, Veh_Stock.ChassisNo, Veh_Stock.EngineNo,ColMast.Col_Desc,SubGroup.Name as PartyName,Veh_Stock.Chassis_RctDate as Sal_VDate,Veh_Stock.Remarks,Veh_Stock.Chassis_RctDocNo as Sal_VNo" & _
           " FROM (Veh_Stock LEFT JOIN SubGroup ON Veh_Stock.TrfParty = SubGroup.SubCode) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code "
    End If
    
    

    mQry = mQry + Condstr
      
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "VehIssReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub VehStkRegProc()
On Error GoTo ELoop
Dim mQry$, mQRY1$, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " and (VStk.InDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and (" & cMID("VStk.Pur_DocId", "3", "1") & " in (" & GridString1 & ") or VStk.Chassis_RctDivCode in (" & GridString1 & "))"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("VStk.Pur_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and (left(VStk.Pur_DocId ,1) in (" & GridString2 & ") or Chassis_RctSiteCode in (" & GridString2 & "))"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VStk.Model in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VStk.Colour_Code in (" & GridString4 & ")"
    
    mQry = "SELECT 1 as Qty,VStk.Chassis_RctDocNo,Model_Grp.ModelGrp_Name as Model,VStk.ChassisNo,VStk.EngineNo,ColMast.Col_Desc,VStk.Pur_DocId,VStk.Pur_VDate," & _
        "VStk.PBILL_NO,VStk.PBILL_DATE,S.Name AS SupplierName,VStk.TAX_YN,VStk.Rate, " & vIsNull("VP.Amount+VP.Addition-VP.Deduction+VP.Misc_Amt", "VStk.Rate") & " As VRate,VStk.Sal_DocID,VStk.Sal_VDate," & _
        "S1.Name As PartyName " & _
        "FROM (((((((Veh_Stock as VStk LEFT JOIN Veh_Order as VO ON VStk.Sal_DocId = VO.Inv_DocId) " & _
        "LEFT JOIN SubGroup S ON VStk.PartyCode = S.SubCode) " & _
        "LEFT JOIN Subgroup as S1 on VO.PartyCode = S1.SubCode) " & _
        "LEFT JOIN Model ON VStk.Model = Model.Model) " & _
        "LEFT JOIN Model_Grp on Model_Grp.ModelGrp_Code=Model.Grp_Code) " & _
        "LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code)" & _
        "Left Join Veh_Purch1 VP On VStk.Pur_DocId=VP.DocId)"
        
    mQRY1 = "SELECT 1 as Qty,VStk.Chassis_RctDocNo,Model_Grp.ModelGrp_Name as Model,VStk.ChassisNo,VStk.EngineNo,ColMast.Col_Desc,VStk.Pur_DocId,VStk.Pur_VDate," & _
        "VStk.PBILL_NO,VStk.PBILL_DATE,S.Name AS SupplierName,VStk.TAX_YN,VStk.Rate, " & vIsNull("VP.Amount+VP.Addition-VP.Deduction+VP.Misc_Amt", "VStk.Rate") & " as VRate,VStk.Sal_DocID,VStk.Sal_VDate," & _
        "'' as PartyName " & _
        "FROM (((((Veh_Stock as VStk LEFT JOIN SubGroup S ON VStk.PartyCode = S.SubCode) " & _
        "LEFT JOIN Model ON VStk.Model = Model.Model) " & _
        "LEFT JOIN Model_Grp on Model_Grp.ModelGrp_Code=Model.Grp_Code) " & _
        "LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code)" & _
        "Left Join Veh_Purch1 VP On VP.DocId=VStk.Pur_DocId)"
    
    If FGrid.TextMatrix(List1, 1) = "With Sale" Then
'        mQRY = mQRY & " WHERE (Not IsNull(VStk.Sal_VDate) and VStk.Sal_VDate< #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#)"
'        mQRY1 = mQRY1 & " WHERE (IsNull(VStk.Sal_VDate) )"
        mQry = mQry & " WHERE (VO.Inv_Date Is Not Null and VO.Inv_Date< " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
        mQRY1 = mQRY1 & " WHERE (VStk.Sal_DocId Is Null or VStk.Sal_DocID='' )"
    End If
    If FGrid.TextMatrix(List1, 1) = "W/O Sale" Then
'        mQRY = mQRY & " WHERE (VStk.Sal_VDate > #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#)"
'        mQRY1 = mQRY1 & " WHERE (IsNull(VStk.Sal_VDate) )"
        mQry = mQry & " WHERE (VO.Inv_Date > " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
        mQRY1 = mQRY1 & " WHERE (VStk.Sal_VDate is NULL or (VStk.Sal_VDate >=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VStk.Sal_Vtype='V_TRF'))"
    End If
    mQry = mQry & Condstr & " Union " & mQRY1 & Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List1, 1) = "With Sale" Then
        RepName = "VehStkReg"
    End If
    If FGrid.TextMatrix(List1, 1) = "W/O Sale" Then
        RepName = "VehStkReg1"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub VehStkBankProc()
On Error GoTo ELoop
Dim mQry$, mQRY1$, Condstr$
Dim sQry$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " and VStk.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and (" & cMID("VStk.Pur_DocId", "3", "1") & " in (" & GridString1 & ") or VStk.Chassis_RctDivCode in (" & GridString1 & "))"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("VStk.Pur_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and (left(VStk.Pur_DocId ,1) in (" & GridString2 & ") or VStk.Chassis_RctSiteCode in (" & GridString2 & "))"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VStk.Model in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VStk.Godown in (" & GridString4 & ")"
    
    
    If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
        sQry = "(Select Amount+Addition-Deduction From Veh_Purch1 VP Where VP.DocId=VStk.Pur_DocId)"
    Else
        sQry = "VStk.Rate"
    End If
    
    
    mQry = "SELECT 1 as Qty,VStk.Chassis_RctDocNo,Model_Grp.ModelGrp_Name as Model,VStk.ChassisNo,VStk.EngineNo,ColMast.Col_Desc,VStk.Pur_DocId,VStk.Pur_VDate," & _
        "VStk.PBILL_NO,VStk.PBILL_DATE,S.Name AS SupplierName,VStk.TAX_YN, " & sQry & " As Rate,VStk.VRate,VStk.Sal_DocID,VStk.Sal_VDate," & _
        "S1.Name As PartyName,Model.Model as Model_No, Model.Model_Desc, Model.Sales_Desc,Model_Cat.ModelCat_Name, G.God_Name, Site.Site_Desc " & _
        "FROM (((((((((Veh_Stock as VStk LEFT JOIN Veh_Order as VO ON VStk.Sal_DocId=VO.Inv_DocId) " & _
        "LEFT JOIN SubGroup S ON VStk.PartyCode = S.SubCode) " & _
        "LEFT JOIN Subgroup as S1 on VO.PartyCode=S1.SubCode) " & _
        "LEFT JOIN Model ON VStk.Model = Model.Model) " & _
        "Left Join Godown G On VStk.Godown=G.God_Code) " & _
        "LEFT JOIN Model_Grp on Model_Grp.ModelGrp_Code=Model.Grp_Code) " & _
        "LEFT JOIN Model_Cat on Model_Cat.ModelCat_Code=Model.Cat_Code) " & _
        "LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
        "Left Join Site On Site.Site_Code = " & cMID("VStk.Pur_DocId", "3", "1") & " )"
        
    mQRY1 = "SELECT 1 as Qty,VStk.Chassis_RctDocNo,Model_Grp.ModelGrp_name as Model,VStk.ChassisNo,VStk.EngineNo,ColMast.Col_Desc,VStk.Pur_DocId,VStk.Pur_VDate," & _
        "VStk.PBILL_NO,VStk.PBILL_DATE,S.Name AS SupplierName,VStk.TAX_YN, " & sQry & " As  Rate,VStk.VRate,VStk.Sal_DocID,VStk.Sal_VDate," & _
        "'' as PartyName,Model.Model as Model_No, Model.Model_Desc, Model.Sales_Desc,Model_Cat.ModelCat_Name, G.God_Name, Site.Site_Desc  " & _
        "FROM (((((((Veh_Stock as VStk LEFT JOIN SubGroup S ON VStk.PartyCode = S.SubCode) " & _
        "LEFT JOIN Model ON VStk.Model = Model.Model) " & _
        "Left Join Godown G On VStk.Godown=G.God_Code) " & _
        "LEFT JOIN Model_Grp on Model_Grp.ModelGrp_Code=Model.Grp_Code) " & _
        "LEFT JOIN Model_Cat on Model_Cat.ModelCat_Code=Model.Cat_Code) " & _
        "LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
        "Left Join Site On Site.Site_Code = " & cMID("VStk.Pur_DocId", "3", "1") & " ) "

'    If FGrid.TextMatrix(List1, 1) = "W/O Sale" Then
'        mQRY = mQRY & " WHERE (VStk.Sal_VDate > #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#)"
'        mQRY1 = mQRY1 & " WHERE (IsNull(VStk.Sal_VDate) )"
        mQry = mQry & " WHERE (VO.Inv_Date > " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
        mQRY1 = mQRY1 & " WHERE (VStk.Sal_DocID Is Null or VStk.Sal_DocID='' )"
'    End If
    mQry = mQry & Condstr & " Union " & mQRY1 & Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    If UCase(left(PubComp_Name, 4)) <> "ENAR" And UCase(left(PubComp_Name, 6)) <> "J.M.A." Then
        RepName = "VehStkBank"
    Else
        RepName = "VehStkBank_Enar"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub VehStkHoldProc()
On Error GoTo ELoop
Dim mQry As String, mQRY1 As String, Condstr As String, CondStr1 As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and (" & cMID("VStk.Pur_DocId", "3", "1") & " in (" & GridString1 & ") or VStk.Chassis_RctDivCode in (" & GridString1 & "))"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("VStk.Pur_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and (left(VStk.Pur_DocId ,1) in (" & GridString2 & ") or VStk.Chassis_RctSiteCode in (" & GridString2 & "))"
    If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VSTK.GODOWN in (" & GridString3 & ")"
    Else
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.PartyCode in (" & GridString3 & ")"
    End If
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VStk.MODEL  in (" & GridString4 & ")"
    
    If FGrid.TextMatrix(List2, 1) <> "All" Then
           CondStr1 = " and VP1.BMS_CATEGORY  = '" & FGrid.TextMatrix(List2, 2) & "' "
          
    End If
'      If Check1(1).Value = Unchecked Then CondStr1 = CondStr1 & " and (" & cMID("vp1.Pur_DocId", "3", "1") & " in (" & GridString1 & ") or VStk.Chassis_RctDivCode in (" & GridString1 & "))"
'           If Check1(1).Value = Checked Then
'          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then CondStr1 = CondStr1 & " and  " & cMID("VStk.Pur_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
'        End If
'
        
    mQry = "SELECT 1 as Qty,VP1.BMS_CATEGORY,VStk.Godown,Godown.God_Name,VStk.Model,VStk.ChassisNo,VStk.EngineNo,ColMast.Col_Desc,VStk.Chassis_RctSiteCode," & _
        "VStk.Pur_DocID,VP1.BMS_CATEGORY,VStk.PBILL_NO,VStk.PBILL_DATE,S.NamePrefix,S.Name,S.FPrefix,S.FName,S.Add1,S.Add2,S.Add3,S.AREA,City.CityName, " & _
        "VStk.InDate,VStk.TAX_YN,VStk.VRate,VStk.Sal_DocID,VStk.Sal_VDate,'' AS Model_Desc,0 AS Rate, DateDiff(Day, VStk.InDate, (Case When IsNull(VStk.Sal_DocID,'') ='' Then '" & FGrid.TextMatrix(Date1, 1) & "' Else VStk.Sal_VDate End)) as HoldDays  " & _
        "FROM ((((((Veh_Stock as VStk LEFT JOIN Veh_Order as VO ON VStk.Sal_DocId=VO.Inv_DocId) " & _
        "LEFT JOIN SubGroup S ON VO.PartyCode = S.SubCode) " & _
        "LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code) " & _
        " LEFT JOIN Godown ON VStk.Godown = Godown.God_Code) " & _
        " LEFT JOIN City ON S.CityCode = City.CityCode)" & _
        " Left Join Veh_Purch1 as VP1 on VStk.Pur_DocID=VP1.DocID)"
        
    mQRY1 = "SELECT 1 as Qty,VP1.BMS_CATEGORY,VStk.Godown,Godown.God_Name,VStk.Model,VStk.ChassisNo,VStk.EngineNo,ColMast.Col_Desc,VStk.Chassis_RctSiteCode," & _
        "VStk.Pur_DocID,VP1.BMS_CATEGORY,VStk.PBILL_NO,VStk.PBILL_DATE,'' as NamePrefix,'' as Name,'' as FPrefix,'' as FName,'' as Add1,'' as Add2,'' as Add3,'' as AREA,'' as CityName, " & _
        "VStk.InDate,VStk.TAX_YN,VStk.VRate,VStk.Sal_DocID,VStk.Sal_VDate,Model.Model_Desc,VStk.Rate, DateDiff(Day, VStk.InDate, (Case When IsNull(VStk.Sal_DocID,'') ='' Then  '" & FGrid.TextMatrix(Date1, 1) & "' Else  VStk.Sal_VDate End)) as HoldDays  " & _
        "FROM (((Veh_Stock as VStk LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code)" & _
        " LEFT JOIN Godown ON VStk.Godown = Godown.God_Code) " & _
        " Left Join Veh_Purch1 as VP1 on VStk.Pur_DocID=VP1.DocID LEFT JOIN Model ON VStk.Model=Model.MODEL )"
            
    If FGrid.TextMatrix(List1, 1) = "With Sale" Then
'        mQRY = mQRY & " WHERE (Not IsNull(VStk.Sal_VDate) and VStk.Sal_VDate< #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#)"
        mQry = mQry & " WHERE (VO.Inv_Date Is Not Null and VO.Inv_Date< " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    ElseIf FGrid.TextMatrix(List1, 1) = "W/O Sale" Then
'        mQRY = mQRY & " WHERE (VStk.Sal_VDate > #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "#)"
     '  mQry = mQry & " WHERE (VO.Inv_Date > " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
     If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
       mQry = mQry & " WHERE (VO.Inv_Date > " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VStk.Pur_VDate<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    Else
       mQry = mQry & " WHERE (VO.Inv_Date > " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    End If
    End If
'    mQRY1 = mQRY1 & " WHERE (IsNull(VStk.Sal_VDate) )"
    If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
           mQRY1 = mQRY1 & " WHERE (VStk.Sal_DocID Is Null or VStk.Sal_DocID='' and VStk.Pur_VDate< =" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    Else
          mQRY1 = mQRY1 & " WHERE (VStk.Sal_DocID Is Null or VStk.Sal_DocID='' "
    End If
    If FGrid.TextMatrix(Cat1, 1) > 0 Then
        mQry = mQry & " And DateDiff(Day, VStk.InDate, (Case When IsNull(VStk.Sal_DocID,'') ='' Then  '" & FGrid.TextMatrix(Date1, 1) & "' Else  VStk.Sal_VDate End)) > " & Val(FGrid.TextMatrix(Cat1, 1)) & " "
        mQRY1 = mQRY1 & " And DateDiff(Day, VStk.InDate, (Case When IsNull(VStk.Sal_DocID,'') ='' Then  '" & FGrid.TextMatrix(Date1, 1) & "' Else  VStk.Sal_VDate End)) > " & Val(FGrid.TextMatrix(Cat1, 1)) & " "
    End If
    
    
    If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
        mQry = mQry & Condstr & CondStr1 & " Union " & mQRY1 & Condstr & CondStr1
    Else
        If Check1(3).Value <> Unchecked Then
            mQry = mQry & CondStr1 & " Union " & mQRY1 & CondStr1
        Else
            mQry = mQry & Condstr & CondStr1
        End If
    End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
        If MsgBox("Do you want to see report Stock Summary (Godown Wise)", vbYesNo) = vbYes Then
            RepName = "StockReportGodownWiseCrossTab"
            RepTitle = " Stock Summary (Godown Wise) "
        Else
            RepName = "VehStkHold"
            RepTitle = UCase(Me.CAPTION)
        End If
    Else
        RepName = "VehStkHold"
        RepTitle = UCase(Me.CAPTION)
    End If
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description

End Sub

Private Sub RetSaleRepProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

    Condstr = " WHERE  V_O.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and V_O.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("V_O.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("V_O.Inv_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and V_S.Model in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(V_O.Inv_DocId,1) in (" & GridString3 & ")"

'    If FGrid.TextMatrix(List1, 1) = "Yes" Then 'CondStr = CondStr & " and isnull(Job_Card.JobCloseDate)"
'
'    mQRY = "SELECT SG.Name, SG.Add1, SG.Add2, SG.Add3, V_O.ord_Date, V_S.ChassisNo, V_S.MODEL," & _
'           " V_S.PBILL_NO AS TelcoInvNo, V_S.PBILL_DATE AS TelcoDate, V_O.MARGINE,(V_O.Net_Amount - V_O.VRate) As Difference, " & _
'           "V_S.VRATE,iif(ContractFinance.FinCatg=0,ContractFinance.FinName,'')AS finName,V_O.Interest,V_O.RTO,Veh_Rate.MARGINE AS StdMargine,Rect.V_Type,Rect.V_Date,Rect.V_No,Rect.AMOUNT,Rect.DrCr,SubGroupType.Description,City.cityName " & _
'           " FROM ((((Veh_Order V_O   LEFT JOIN ((SubGroup SG LEFT JOIN SubGroupType ON SG.Party_Type = SubGroupType.Party_Type) LEFT JOIN City ON SG.CityCode = City.CityCode) ON V_O.PartyCode = SG.SubCode) " & _
'           "LEFT JOIN Veh_Stock V_S ON V_O.Inv_DocId = V_S.Sal_DocId)" & _
'           " LEFT JOIN ContractFinance ON V_O.FB_CODE = ContractFinance.FinCode)" & _
'           " Left JOIN Veh_rate On V_O.Model=Veh_Rate.Model) " & _
'           " Left JOIN Rect on Rect.Ord_DocId=V_O.OrdDocId "
'    End If
'    If FGrid.TextMatrix(List1, 1) = "No" Then
    
'    mQRY = "SELECT SG.Name, SG.Add1, SG.Add2, SG.Add3, V_O.ord_Date, V_S.ChassisNo, V_S.MODEL," & _
           " V_S.PBILL_NO AS TelcoInvNo, V_S.PBILL_DATE AS TelcoDate, V_O.MARGINE,(V_O.Net_Amount - V_O.VRate) As Difference, " & _
           " V_S.VRATE,iif(ContractFinance.FinCatg=0,ContractFinance.FinName,'')AS finName,V_O.Interest,V_O.RTO,V_O.MARGINE AS StdMargine,SubGroupType.Description,City.cityName,V_O.Transport,V_O.Inv_No,V_O.Inv_Date,V_O.Net_Amount as SaleRate,V_O.Inv_DocId " & _
           " FROM ((((Veh_Order V_O LEFT JOIN SubGroup SG on V_O.PartyCode = SG.SubCode) " & _
           " LEFT JOIN SubGroupType ON SG.Party_Type = SubGroupType.Party_Type) " & _
           " LEFT JOIN City ON SG.CityCode = City.CityCode)  " & _
           " LEFT JOIN Veh_Stock V_S ON V_O.Inv_DocId = V_S.Sal_DocId)" & _
           " LEFT JOIN ContractFinance ON V_O.FB_CODE = ContractFinance.FinCode "
    mQry = "SELECT SG.Name, SG.Add1, SG.Add2, SG.Add3, V_O.ord_Date, V_S.ChassisNo, V_S.MODEL," & _
           " V_S.PBILL_NO AS TelcoInvNo, V_S.PBILL_DATE AS TelcoDate, ((V_O.Net_Amount-V_O.Tax_Amt-V_O.TOT_Amt) - V_O.VRate) as MARGINE,(V_O.Net_Amount - V_O.VRate) As Difference, " & _
           " V_S.VRATE, " & cIIF("ContractFinance.FinCatg=0", "ContractFinance.FinName", "''") & "AS finName,V_O.Interest,V_O.RTO,V_o.margine as stdmargine,Rect.V_Type,Rect.V_Date,Rect.V_No,Rect.AMOUNT,Rect.DrCr,SubGroupType.Description,City.CityName,V_O.Transport,V_O.Inv_No,V_O.Inv_Date,V_O.Net_Amount as SaleRate,V_O.Inv_DocId,SG.NamePrefix " & _
           " FROM ((((Veh_Order as V_O LEFT JOIN ((SubGroup SG LEFT JOIN SubGroupType ON SG.Party_Type = SubGroupType.Party_Type) " & _
           " LEFT JOIN City ON SG.CityCode = City.CityCode) ON V_O.PartyCode = SG.SubCode) " & _
           " LEFT JOIN Veh_Stock V_S ON V_O.Inv_DocId = V_S.Sal_DocId)" & _
           " LEFT JOIN ContractFinance ON V_O.FB_CODE = ContractFinance.FinCode)" & _
           " Left JOIN Rect on Rect.Ord_DocId=V_O.OrdDocId ) "
'      End If
    
    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "RetSaleRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub DelCharegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'   If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If FGrid.TextMatrix(List1, 1) = "Delivered" Then
        Condstr = " where VO.DelCh_DT  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.DelCh_DT <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    Else
        Condstr = " where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    End If
      
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("VO.Inv_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and VO.Model in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_DocId,1) in (" & GridString3 & ")"
    
    mQry = "SELECT VO.Inv_DocId,VO.DelCh_DocId,VO.DelCh_No,VO.VRATE, VO.DelCh_DT, Right(VO.Inv_DocId ,13) as InvNo, " & _
           "VO.Inv_Date,(VO.VRATE+VO.MARGINE) AS SaleRate, VStk.PBILL_NO, VStk.PBILL_DATE, VStk.ChassisNo," & _
           "VStk.EngineNo,SG.NamePrefix, SG.Name, SG.FPrefix, SG.FName, SG.Add1, SG.Add2, SG.Add3, SG.AREA, City.CityName, VStk.Srv_BookNo, VO.Interest, VO.TDS_Amt," & _
           "VO.REG_FEE, VO.S_CHARGE, VO.STAMP_DUTY, VO.INS_FEE," & _
           "Ins_NOTE, VP2.PROD_CODE, VP2.Trn_Type, Model.MODEL " & _
           "FROM (((((Veh_Order as VO LEFT JOIN Subgroup as SG ON VO.PartyCode = SG.SubCode) " & _
           "LEFT JOIN City ON SG.CityCode = City.CityCode)" & _
           "Left JOIN Model ON VO.MODEL = Model.MODEL)" & _
           "LEFT JOIN Veh_Stock as VStk ON VO.Inv_DocId = VStk.Sal_DocId)" & _
           "LEFT JOIN Veh_Purch2 as VP2 ON VO.OrdDocId = VP2.DocID) "
            
    If FGrid.TextMatrix(List1, 1) = "Delivered" Then
           Condstr = Condstr & " and VO.DelCh_DT Is Not Null "
    ElseIf FGrid.TextMatrix(List1, 1) = "Pending" Then
           Condstr = Condstr & " and VO.DelCh_DT Is Null "
    End If
    
    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List1, 1) = "Pending" Then
        RepName = "PendDelChareg"
    Else
        RepName = "DelChareg"
    End If
    
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub AddFitRepProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub

    'CondStr = " where Veh_Order.Inv_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and Veh_Order.Inv_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# "
    Condstr = " where Veh_Purch2.V_TYPE ='V_SB'  and  Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Veh_Order.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("Veh_Order.Inv_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Inv_DocId,1) in (" & GridString3 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.MODEL in (" & GridString2 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Veh_Purch2.PROD_CODE in (" & GridString4 & ")"
    
'    mQRY = "SELECT Veh_Order.Inv_DocId,right(Veh_Order.Inv_DocId,13) AS Inv_No, Veh_Order.Inv_Date, SubGroup.Name," & _
           "Veh_Order.MODEL, Veh_Stock.ChassisNo," & _
           "Veh_Purch2.Trn_Type, Veh_Purch2.PROD_CODE, Veh_Purch2.QTY, Veh_Purch2.RATE," & _
           "Veh_Purch2.TAX_AMT, Veh_Purch2.TaxSur_AMT," & _
           "Veh_Order.TAX_Amt, Veh_Order.Surcharge_Amt, Veh_Order.MISC_INFO," & _
           "Veh_Order.REBATE, Veh_Order.MARGINE," & _
           "Veh_Order.VRATE, Veh_Order.InciChrg, Veh_Order.Octroi," & _
           "Veh_Order.RegTemp, Veh_Order.TransitInsu," & _
           "Veh_Order.Transport, Veh_Order.MVT,Veh_Order.Net_Amount,Veh_Order.OtherChrg,Veh_AMDModel.Prod_Name " & _
           "FROM (((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_docId = Veh_Stock.Sal_DocId) LEFT JOIN " & _
           "SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode)" & _
           "LEFT JOIN Veh_Purch2 ON  Veh_Purch2.DocID = Veh_Order.Inv_DocId) LEFT JOIN Veh_AMDModel ON  Veh_AMDModel.Prod_Code = Veh_Purch2.PROD_CODE "
    
    mQry = "SELECT Veh_Order.Inv_DocId,right(Veh_Order.Inv_DocId,13) AS Inv_No, Veh_Order.Inv_Date, SubGroup.Name," & _
           "Veh_Order.MODEL, Veh_Stock.ChassisNo," & _
           "Veh_Purch2.Trn_Type, Veh_Purch2.PROD_CODE, Veh_Purch2.QTY, Veh_Purch2.RATE," & _
           "Veh_Purch2.TAX_AMT, Veh_Purch2.TaxSur_AMT," & _
           "Veh_Order.TAX_Amt, Veh_Order.Surcharge_Amt, Veh_Order.MISC_INFO," & _
           "Veh_Order.REBATE, Veh_Order.MARGINE," & _
           "Veh_Order.VRATE, Veh_Order.InciChrg, Veh_Order.Octroi," & _
           "Veh_Order.RegTemp, Veh_Order.TransitInsu," & _
           "Veh_Order.Transport, Veh_Order.MVT,Veh_Order.Net_Amount,Veh_Order.OtherChrg,Veh_AMDModel.Prod_Name " & _
           "FROM (((Veh_Purch2 LEFT JOIN Veh_Order ON Veh_Order.Inv_DocId = Veh_Purch2.DocID) " & _
           "LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId= Veh_Stock.Sal_DocId) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode)  " & _
           "LEFT JOIN Veh_AMDModel ON Veh_AMDModel.Prod_Code = Veh_Purch2.PROD_CODE"
    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "AddFitRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub DailyRetailProc()
On Error GoTo ELoop
Dim mOpQry$, mQry$, mQRY1$, Condstr$, ChasDivCond$, PurDivCond$, SalDivCond$
Dim Rst As ADODB.Recordset, rstArea As ADODB.Recordset, mQryArea$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If Check1(3).Value = Unchecked Then
        ChasDivCond = " Chassis_RctDivCode in (" & GridString3 & ") and "
        PurDivCond = " left(VStk.Pur_DocId,1) in (" & GridString3 & ") and "
        SalDivCond = " left(VO.Inv_DocId,1) in (" & GridString3 & ") and "
    End If
    'VStk.MODEL
    
'    If Check1(1).Value = Unchecked Then PurDivCond = PurDivCond & "  " & cMID("VStk.Pur_DocId", "3", "1") & "  in (" & GridString1 & ") and "
'    If Check1(1).Value = Checked Then
'      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then PurDivCond = PurDivCond & "  " & cMID("VStk.Pur_DocId", "3", "1") & "  ='" & PubSiteCode & "' and  "
'    End If
'
'    If Check1(1).Value = Unchecked Then SalDivCond = SalDivCond & "   " & cMID("VStk.sal_DocId", "3", "1") & "  in (" & GridString1 & ") and "
'    If Check1(1).Value = Checked Then
'      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then SalDivCond = SalDivCond & "  " & cMID("VStk.sal_DocId", "3", "1") & "  ='" & PubSiteCode & "' and  "
'    End If
    
    
    
    If Check1(2).Value = Unchecked Then SalDivCond = SalDivCond & "  VO.MODEL in (" & GridString2 & ") and "
    If Check1(2).Value = Unchecked Then PurDivCond = PurDivCond & "  VStk.MODEL in (" & GridString2 & ") and "
    
    
    If PubBackEnd = "A" Then
        mOpQry = "SELECT M.Div_Code,D.Div_SName,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,1 as MonthOpen, 0 AS MonthPurTL, 0 as MonthPurDL, 0 as PurDay, 0 AS MonthSal, 0 as SalDay, '' as AreaName " & _
            " FROM (((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Division D on M.Div_Code=D.Div_Code " & _
            " where ((VStk.Pur_DocId='' and " & ChasDivCond & " format(Chassis_RctDate,'YYYYMM')<'" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & "') " & _
            " or (" & PurDivCond & " format(VStk.Pur_VDate,'yyyymm') < '" & Format(FGrid.TextMatrix(Date1, 1), "yyyymm") & "')) " & _
            " and (VStk.Sal_VDate Is Null or format(VStk.Sal_VDate,'yyyymm')< '" & Format(FGrid.TextMatrix(Date1, 1), "yyyymm") & "')"
        
        mQry = "SELECT M.Div_Code,D.Div_SName,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,0 as MonthOpen, " & _
            " " & cIIF("VStk.Pur_VDate < " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VP.OBNO<>''", "1", "0") & " AS MonthPurTL, " & _
            " " & cIIF("VStk.Pur_VDate < " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VP.OBNO<>''", "0", "1") & " AS MonthPurDL, " & _
            " " & cIIF("VStk.Pur_VDate = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS PurDay, " & _
            " 0 as MonthSal, 0 AS SalDay, '' as AreaName " & _
            " FROM ((((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Veh_Purch1 VP on VStk.Pur_DocId=VP.DocID) " & _
            " Left join Division D on M.Div_Code=D.Div_Code " & _
            " where (VStk.Pur_DocId='' and " & ChasDivCond & " Format(Chassis_RctDate,'YYYYMM')='" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & "' and Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
            " or (" & PurDivCond & " format(VStk.Pur_VDate,'YYYYMM')= '" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & "' and VStk.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    
    'Grp_Code  ModelGrp_Code    ModelGrp_Name
    'Cat_Code  ModelCat_Code    ModelCat_Name
        mQRY1 = "SELECT M.Div_Code,D.Div_SName,MG.ModelGrp_Name,MC.ModelCat_Name,VO.Model,0 as MonthOpen, 0 as MonthPurTL, 0 as MonthPurDL, 0 as PurDay," & _
            " " & cIIF("VO.Inv_Date< " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS MonthSal, " & _
            " " & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS SalDay, Area.AreaName " & _
            " FROM ((((Veh_Order VO Left Join Model M on VO.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Area on VO.Area=Area.AreaCode) " & _
            " Left join Division D on M.Div_Code=D.Div_Code " & _
            " where " & SalDivCond & _
            " format(VO.Inv_Date,'YYYYMM')= '" & Format(FGrid.TextMatrix(Date1, 1), "YYYYMM") & _
            "' and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ""
    ElseIf PubBackEnd = "S" Then
        mOpQry = "SELECT M.Div_Code,D.Div_SName,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,1 as MonthOpen, 0 AS MonthPurTL, 0 as MonthPurDL, 0 as PurDay, 0 AS MonthSal, 0 as SalDay, '' as AreaName " & _
            " FROM (((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Division D on M.Div_Code=D.Div_Code " & _
            " where ((VStk.Pur_DocId='' and " & ChasDivCond & " " & cCStr("Month(Chassis_RctDate)") & "+" & CStr("Year(Chassis_RctDate)") & " <'" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "') " & _
            " or (" & PurDivCond & " " & cCStr("Month(VStk.Pur_VDate)") & " + " & cCStr("Year(VStk.Pur_VDate)") & " < '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "')) " & _
            " and (VStk.Sal_VDate Is Null or " & cCStr("Month(VStk.Sal_VDate)") & "+ " & cCStr("Year(VStk.Sal_VDate)") & " < '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "')"
        
        mQry = "SELECT M.Div_Code,D.Div_SName,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,0 as MonthOpen, " & _
            " " & cIIF("VStk.Pur_VDate < " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VP.OBNO<>''", "1", "0") & " AS MonthPurTL, " & _
            " " & cIIF("VStk.Pur_VDate < " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VP.OBNO<>''", "0", "1") & " AS MonthPurDL, " & _
            " " & cIIF("VStk.Pur_VDate = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS PurDay, " & _
            " 0 as MonthSal, 0 AS SalDay, '' as AreaName " & _
            " FROM ((((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Veh_Purch1 VP on VStk.Pur_DocId=VP.DocID) " & _
            " Left join Division D on M.Div_Code=D.Div_Code " & _
            " where (VStk.Pur_DocId='' and " & ChasDivCond & "  " & cCStr("Month(Chassis_RctDate)") & " + " & cCStr("Year(Chassis_RctDate)") & " = '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "' and Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
            " or (" & PurDivCond & " " & cCStr("Month(VStk.Pur_VDate)") & "+" & cCStr("Year(VStk.Pur_VDate)") & "= '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "' and VStk.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    
    'Grp_Code  ModelGrp_Code    ModelGrp_Name
    'Cat_Code  ModelCat_Code    ModelCat_Name
        mQRY1 = "SELECT M.Div_Code,D.Div_SName,MG.ModelGrp_Name,MC.ModelCat_Name,VO.Model,0 as MonthOpen, 0 as MonthPurTL, 0 as MonthPurDL, 0 as PurDay," & _
            " " & cIIF("VO.Inv_Date< " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS MonthSal, " & _
            " " & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS SalDay, Area.AreaName " & _
            " FROM ((((Veh_Order VO Left Join Model M on VO.MODEL = M.MODEL) " & _
            " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
            " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
            " Left join Area on VO.Area=Area.AreaCode) " & _
            " Left join Division D on M.Div_Code=D.Div_Code " & _
            " where " & SalDivCond & _
            " " & cCStr("Month(VO.Inv_Date)") & "+" & cCStr("Year(VO.Inv_Date)") & "= '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & _
            "' and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ""
    
    End If
    mQry = mOpQry & " Union all " & mQry & " Union all " & mQRY1
    'Area Qry
    mQryArea = "SELECT Distinct Area.AreaName " & _
        " FROM ((Veh_Order VO Left Join Model M on VO.MODEL = M.MODEL) " & _
        " Left join Area on VO.Area=Area.AreaCode) " & _
        " Left join Division D on M.Div_Code=D.Div_Code " & _
        " where " & SalDivCond & _
        " VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " Order By Area.AreaName"
    
    Set rstArea = New Recordset
    rstArea.CursorLocation = adUseClient
    
    
    rstArea.Open (mQryArea), GCn, adOpenStatic, adLockReadOnly

'    Set RstRep = New Recordset
'    RstRep.CursorLocation = adUseClient
'    RstRep.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
''    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: RepPrint = False: Exit Sub
'Temp Table
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
'        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Set Rst = Nothing: Exit Sub

        'Create temp table
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "Div_Code", adChar, 1, adFldIsNullable
            .Fields.Append "Div_SName", adChar, 5, adFldIsNullable
            .Fields.Append "ModelGrp_Name", adChar, 15, adFldIsNullable
            .Fields.Append "ModelCat_Name", adChar, 20, adFldIsNullable
            .Fields.Append "Model", adChar, 20, adFldIsNullable
            .Fields.Append "MonthOpen", adInteger, 4, adFldIsNullable
            .Fields.Append "MonthPurTL", adInteger, 4, adFldIsNullable
            .Fields.Append "MonthPurDL", adInteger, 4, adFldIsNullable
            .Fields.Append "PurDay", adInteger, 4, adFldIsNullable
            .Fields.Append "MonthSal", adInteger, 4, adFldIsNullable
            .Fields.Append "SalDay", adInteger, 4, adFldIsNullable
'            .Fields.Append "AreaName", adChar, 15, adFldIsNullable
            .Fields.Append "Head1", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head2", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head3", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head4", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head5", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head6", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head7", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head8", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head9", adVarChar, 5, adFldIsNullable
            .Fields.Append "Head10", adVarChar, 5, adFldIsNullable

            .Fields.Append "Val1", adInteger, 4, adFldIsNullable   '1
            .Fields.Append "Val2", adInteger, 4, adFldIsNullable     '2
            .Fields.Append "Val3", adInteger, 4, adFldIsNullable     '3
            .Fields.Append "Val4", adInteger, 4, adFldIsNullable    '4
            .Fields.Append "Val5", adInteger, 4, adFldIsNullable    '5
            .Fields.Append "Val6", adInteger, 4, adFldIsNullable    '6
            .Fields.Append "Val7", adInteger, 4, adFldIsNullable    '7
            .Fields.Append "Val8", adInteger, 4, adFldIsNullable    '8
            .Fields.Append "Val9", adInteger, 4, adFldIsNullable    '9
            .Fields.Append "Val10", adInteger, 4, adFldIsNullable    '10

            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
        End With
        'temp table created

        Do While Rst.EOF = False
            With RstRep
                .AddNew
                .Fields("Div_Code") = Rst!Div_Code
                .Fields("Div_SName") = Rst!Div_SName
                .Fields("ModelGrp_Name") = Rst!ModelGrp_Name
                .Fields("ModelCat_Name") = Rst!ModelCat_NAME
                .Fields("Model") = Rst!Model
                .Fields("MonthOpen") = Rst!MonthOpen
                .Fields("MonthPurDL") = Rst!MonthPurDL
                .Fields("PurDay") = Rst!PurDay
                .Fields("MonthSal") = Rst!MonthSal
                .Fields("SalDay") = Rst!SalDay
                'Area
                'If rstArea.RecordCount  > 0 Then
                    If IsNull(Rst!AreaName) Or Rst!AreaName = "" Then
                    Else
                        If rstArea.RecordCount > 0 Then rstArea.MoveFirst
                        rstArea.FIND ("AreaName='" & Rst!AreaName & "'")
                        If rstArea.AbsolutePosition <= 9 And rstArea.AbsolutePosition > 0 Then
                            .Fields("Val" & rstArea.AbsolutePosition) = Rst!SalDay
                        Else
                            .Fields("Val10") = Rst!SalDay
                        End If
                    End If
                    If rstArea.RecordCount > 0 Then rstArea.MoveFirst
                    Do While rstArea.EOF = False
                        If rstArea.AbsolutePosition <= 9 Then
                            .Fields("Head" & rstArea.AbsolutePosition) = left(rstArea!AreaName, 4)
                        Else
                            .Fields("Head10") = "OTH"
                        End If
                        rstArea.MoveNext
                    Loop
               ' End If
                .Update
            End With
            Rst.MoveNext
        Loop
    Set Rst = Nothing
    Set rstArea = Nothing
    
    RepName = "DailyRetail"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub DailyRetailProcTelco()

On Error GoTo ELoop
Dim mOpQry$, mQry$, mQRY1$, Condstr$, ChasDivCond$, PurDivCond$, SalDivCond$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
'    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If Check1(3).Value = Unchecked Then
        ChasDivCond = " Chassis_RctDivCode in (" & GridString3 & ") and "
        PurDivCond = " left(VStk.Pur_DocId,1) in (" & GridString3 & ") and "
        SalDivCond = " left(VO.Inv_DocId,1) in (" & GridString3 & ") and "
    End If
    
    mOpQry = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,1 as DayOpen, 0 AS MonthPur, 0 as PurDay, 0 AS MonthSal, 0 as SalDay, '' as AreaName " & _
        " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
        " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
        " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
        " where (VStk.Pur_DocId='' and " & ChasDivCond & " Chassis_RctDate<" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
        " or (" & PurDivCond & " VStk.Pur_VDate < " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
        " and (VStk.Sal_VDate Is Null or VStk.Sal_VDate<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"
    
    mQry = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VStk.Model,0 as DayOpen, 1 AS MonthPur, " & _
        " " & cIIF("VStk.Pur_VDate = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS PurDay, " & _
        " 0 as MonthSale, 0 AS SalDay, '' as AreaName " & _
        " FROM ((Veh_Stock VStk LEFT JOIN Model M ON VStk.MODEL = M.MODEL) " & _
        " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
        " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code " & _
        " where (VStk.Pur_DocId='' and " & ChasDivCond & " " & cCStr("Month(Chassis_RctDate)") & " + " & cCStr("Year(Chassis_RctDate)") & "='" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "' and Chassis_RctDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") " & _
        " or (" & PurDivCond & " " & cCStr("Month(VStk.Pur_VDate)") & " + " & cCStr("Year(VStk.Pur_VDate)") & " = '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & "' and VStk.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")"

    mQRY1 = "SELECT M.Div_Code,MG.ModelGrp_Name,MC.ModelCat_Name,VO.Model,0 as DayOpen, 0 as MonthPur, 0 as PurDay, 1 AS MonthSale, " & _
        "" & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "1", "0") & " AS SalDay, Area.AreaName " & _
        " FROM (((Veh_Order VO Left Join Model M on VO.MODEL = M.MODEL) " & _
        " Left join Model_Grp MG on M.Grp_Code=MG.ModelGrp_Code) " & _
        " Left join Model_Cat MC on M.Cat_Code=MC.ModelCat_Code) " & _
        " Left join Area on VO.Area=Area.AreaCode " & _
        " where " & SalDivCond & _
        " " & cCStr("Month(VO.Inv_Date)") & " + " & cCStr("Year(VO.Inv_Date)") & " = '" & Format(FGrid.TextMatrix(Date1, 1), "MYYYY") & _
        "' and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ""
    
    mQry = mOpQry & " Union all " & mQry & " Union all " & mQRY1

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
'    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: RepPrint = False: Exit Sub
    
    RepName = "DailyRetailTelco"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description

End Sub

Private Sub VehTargetProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " And veh_f.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  left(veh_f.Site_Code,1)   ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " And veh_F.Model in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " And Model.Div_Code in (" & GridString3 & ")"
    
    mQry = "SELECT Model.Model_Desc,Veh_F.MODEL," & _
           "Veh_F.QTY_04 AS OApr, Veh_F.QTY_05 AS OMay, Veh_F.QTY_06 AS OJun, Veh_F.QTY_07 AS OJul, Veh_F.QTY_08 AS OAug," & _
           "Veh_F.QTY_09 AS OSep, Veh_F.QTY_10 AS OOct, Veh_F.QTY_11 AS ONov, Veh_F.QTY_12 AS ODec, Veh_F.QTY_01 AS OJan," & _
           " Veh_F.QTY_02 AS Feb, Veh_F.QTY_03 AS OMar, Veh_F.TargQty_04 AS TApr, Veh_F.TargQty_05 AS TMay , Veh_F.TargQty_06 AS TJun," & _
           " Veh_F.TargQty_07 AS TJul, Veh_F.TargQty_08 AS TAug, Veh_F.TargQty_09 AS TSep, Veh_F.TargQty_10 AS TOct, Veh_F.TargQty_11 AS TNov," & _
           " Veh_F.TargQty_12 AS TDec, Veh_F.TargQty_01 AS TJan , Veh_F.TargQty_02 AS TFeb, Veh_F.TargQty_03 AS TMar, " & _
           "" & cIIF("Pur_VDate  >= " & ConvertDate(Format("01/" & "Apr/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("30/" & "Apr/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS AprPurch, " & _
           "" & cIIF("Pur_VDate  >= " & ConvertDate(Format("01/" & "May/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("31/" & "May/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS MayPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Jun/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("30/" & "Jun/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS JunPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Jul/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("31/" & "Jul/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS JulPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Aug/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("31/" & "Aug/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS AugPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Sep/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("30/" & "Sep/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS SepPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Oct/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("31/" & "Oct/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS octPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Nov/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("30/" & "Nov/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS NovPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Dec/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("31/" & "Dec/" & Format(PubStartDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS DecPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Jan/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("31/" & "Jan/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS JanPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format(fxLastDay(Format("27/" & "Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & "/Feb/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS FebPurch, " & _
           "" & cIIF(" Pur_VDate  >= " & ConvertDate(Format("01/" & "Mar/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & " AND Pur_VDate <= " & ConvertDate(Format("31/" & "Mar/" & Format(PubEndDate, "yyyy"), "dd/MMM/yyyy")) & "", "1", "0") & " AS MarPurch " & _
           " FROM ((Model left JOIN Veh_Forecast Veh_F ON veh_f.MODEL = model.MODEL) " & _
           " LEFT JOIN veh_Stock ON model.Model = veh_Stock.model) "
    mQry = mQry + Condstr

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "VehTarget"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VehSaleRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub

    Condstr = " Where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  " & cMID("VO.Inv_DocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Model.Sales_Desc in (" & GridString2 & ")"
    Else
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_DocId,1) in (" & GridString2 & ")"
    End If
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VO.Model in (" & GridString4 & ")"

    Select Case FGrid.TextMatrix(List2, 1)
        Case "PartyWise"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.PartyCode in (" & GridString3 & ")"
        Case "CityWise"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.PartyCode In (Select SG1.SubCode From SubGroup as SG1 where SG1.CityCode In (" & GridString3 & "))"
             
        Case "FinancierGrp"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.FB_CODE in (Select ContractFinance.FinCode From ContractFinance where ContractFinance.UnderFinGrp in (" & GridString3 & "))"
            
        Case "FinancierName"
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.FB_CODE in (" & GridString3 & ")"
        Case "FormType"
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.Form_Code in (" & GridString3 & ")"
        Case "Insu.Auth."
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.RegBy in (" & GridString3 & ")"
        Case "SalesManWise"
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.Rep_Code in (" & GridString3 & ")"
             
            mQry = "SELECT VO.Inv_DocId, VO.Inv_Date, VO.OrdDocId,VO.Ord_Date,SG.NamePrefix,SG.NAME, SG.FPrefix, SG.FName, SG.Add1, SG.Add2, SG.Add3,'' as CityName, " & _
               "" & cIIF("'" & UCase(left(PubComp_Name, 4)) & "'='ENAR'", "Model.Sales_Desc", "Model.Model") & " as Model,VO.Chassis as ChassisNo,VStk.EngineNo,VO.VRATE, VStk.PBILL_NO, VStk.PBILL_DATE," & _
               "VO.MARGINE, VO.InciChrg, VO.Octroi, VO.Transport, VO.RegTemp," & _
               "VO.TAX_Amt,VO.Surcharge_Amt, VO.OtherChrg, VO.Net_AMOUNT," & _
               "VO.Form_Code,TF.Form_Desc,FinGroup.FinGrpName,ContractFinance.FinName,VO.Fin_Amt,VO.MISC_INFO,VO.TOT_Amt,Emp_Mast.Emp_Name,VO.SubTot, VO.RtoFee, Vo.Insurance " & _
               " FROM ((((((((Veh_Order as VO LEFT JOIN Veh_Stock as VStk ON VO.chassis = VStk.ChassisNo) " & _
               "LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
               "LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
               "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
               "LEFT JOIN ContractFinance ON VO.FB_CODE = ContractFinance.FinCode) " & _
               "LEFT JOIN FinGroup ON ContractFinance.UnderFinGrp = FinGroup.FinGrpCode) " & _
               "LEFT JOIN TaxForms as TF ON VO.Form_Code = TF.Form_Code) " & _
               "LEFT JOIN Emp_Mast ON VO.Rep_Code = Emp_Mast.Emp_Code )"
            GoTo NXT
        
    End Select
      
    mQry = "SELECT VO.Inv_DocId, VO.Inv_Date, VO.OrdDocId,VO.Ord_Date,SG.NamePrefix, SG.Name, SG.FPrefix, SG.FName, SG.Add1, SG.Add2, SG.Add3, City.CityName," & _
        "" & cIIF("'" & UCase(left(PubComp_Name, 4)) & "'='ENAR'", "Model.Sales_Desc", "Model.Model") & " as Model,VO.Chassis as ChassisNo,VStk.EngineNo,VO.VRATE, VStk.PBILL_NO, VStk.PBILL_DATE," & _
        "VO.MARGINE, VO.InciChrg, VO.Octroi, VO.Transport, VO.RegTemp," & _
        "VO.TAX_Amt,VO.Surcharge_Amt, VO.OtherChrg, VO.Net_AMOUNT," & _
        "VO.Form_Code,TF.Form_Desc,FinGroup.FinGrpName,ContractFinance.FinName,VO.Fin_Amt,VO.MISC_INFO,VO.TOT_Amt,'" & FGrid.TextMatrix(List2, 1) & "' as ReportType,VO.SubTot,sg.phone, VO.RtoFee, VO.Insurance " & _
        " FROM (((((((Veh_Order as VO LEFT JOIN Veh_Stock as VStk ON VO.chassis = VStk.ChassisNo) " & _
        "LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
        "LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
        "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
        "LEFT JOIN ContractFinance ON VO.FB_CODE = ContractFinance.FinCode) " & _
        "LEFT JOIN FinGroup ON ContractFinance.UnderFinGrp = FinGroup.FinGrpCode) " & _
        "LEFT JOIN TaxForms as TF ON VO.Form_Code = TF.Form_Code) "


NXT:
    mQry = mQry + Condstr & " and (Vstk.Sal_Vtype<>'V_TRF' Or VStk.Sal_VType Is Null) order by VO.Inv_Date,right(VO.Inv_DocId,8) "

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    Set RstRep1 = New ADODB.Recordset
    RstRep1.Open "Select Sum(VRATE+Margine) as Cancel_Amt,sum(Tax_Amt) as CancelTax_Amt,sum(TOT_Amt) as CancelTOT_Amt from Veh_Order1 as VO1 Where VO1.Ord_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO1.Ord_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " ", GCn, adOpenStatic, adLockReadOnly
    ' For Speed print of report
    If SpeedPrnVehSale = True And FGrid.TextMatrix(List1, 1) = "Summary" Then
        SpeedPrintSumm
        Exit Sub
    ElseIf SpeedPrnVehSale = True And FGrid.TextMatrix(List1, 1) = "Detailed" Then
        SpeedPrintDet
         Exit Sub
    End If
    ' End Print
    
    If FGrid.TextMatrix(List1, 1) = "Summary" Then
        If FGrid.TextMatrix(List3, 1) = "All" Then
            RepName = "VehSaleRegSumAll"
        Else
            RepName = "VehSaleRegSum"
        End If
    ElseIf FGrid.TextMatrix(List1, 1) = "Detailed" Then
        If FGrid.TextMatrix(List3, 1) = "All" Then
            RepName = "VehSaleRegDetAll"
        Else
            RepName = "VehSaleRegDet"
        End If
    End If
    If FGrid.TextMatrix(List2, 1) = "SalesManWise" Then
        If FGrid.TextMatrix(List1, 1) = "Summary" Then
            RepName = "SalesManWiseSaleRepSum"
        Else
            RepName = "SalesManWiseSaleRep"
        End If
        Me.CAPTION = "Sales Man Wise Sale Report"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SpeedPrintSumm()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double, mCounter As Double
    Dim isLast As Boolean, mRec As Integer, PageNo As Double
    Dim RstCompDet As ADODB.Recordset, TotalNetAmt As Double
    Dim fob As New FileSystemObject
    
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    PageLength = PubPageLength
    PageWidth = 132
    mRec = 9
    'Header printing
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!S_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
    If PubComp_Add2 <> "" Then
        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    
    Print #1, PRN_TIT("Vehicle Sale Register", "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, "From : " & FGrid.TextMatrix(Date1, 1) & "  To : " & FGrid.TextMatrix(Date2, 1)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & Space(5) & Space(10) & PSTR("Sales", 20) & PSTR("Order", 20) & PSTR("Name of Customer", 35) & PSTR("Model", 15) & PSTR("Telco", 15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
    mHeader = mHeader + 1
    Print #1, PSTR("#", 5) & PSTR("Invoice", 10) & PSTR("Date", 20) & PSTR("Date", 20) & PSTR("Address", 35) & PSTR("Chassis-No", 15) & PSTR("Bill No", 15) & PSTR("Sale-Amt", 15) & PSTR("Financer Group", 35) & Space(15) & PSTR("Form", 5) & PSTR("Spl.Info", 20)
    mHeader = mHeader + 1
    Print #1, PSTR("#", 5) & PSTR("Prefix", 10) & PSTR("Inv-No", 20) & PSTR("No.", 20) & PSTR("Name Of City", 35) & PSTR("Engine No", 15) & PSTR("Date", 15) & Space(15) & PSTR("Financer Name", 35) & PSTR("Fin-Amt", 15) & PSTR("Type", 5) & Space(20) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    RstRep.MoveFirst
    mHeader = 1
    While Not RstRep.EOF = True
        If Counter <= mRec Then
            Counter = Counter + 1
            mCounter = mCounter + 1
            Print #1, mChr17 & PSTR(STR(mCounter), 5) & PSTR(mID(RstRep!Inv_DocId, 9, 5), 10) & PSTR("'" & RstRep!Inv_Date & "'", 20) & PSTR("'" & RstRep!Ord_Date & "'", 20) & PSTR(RstRep!Name, 33) & Space(2) & PSTR(RstRep!Model, 20) & PSTR("'" & RstRep!PBILL_NO & "'", 15) & PSTR(IIf(RstRep!Net_Amount = 0, "", STR(RstRep!Net_Amount)), 15) & PSTR(RstRep!FinGrpName, 35) & PSTR(IIf(RstRep!Fin_Amt = 0, "", STR(RstRep!Fin_Amt)), 15) & PSTR(RstRep!Form_Code, 5) & PSTR(RstRep!MISC_INFO, 20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & PSTR(PrinID(RstRep!Inv_DocId), 20) & PSTR(PrinID(RstRep!OrdDocId), 20) & PSTR(RstRep!Add1, 33) & Space(2) & PSTR(RstRep!ChassisNo, 20) & PSTR("'" & RstRep!PBILL_DATE & "'", 15) & Space(15) & PSTR(RstRep!FinName, 35) & Space(15) & Space(5) & Space(20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & Space(20) & Space(20) & PSTR(RstRep!Add2, 33) & Space(2) & PSTR(RstRep!EngineNo, 20) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & Space(20) & Space(20) & PSTR(RstRep!Add3, 33) & Space(2) & Space(15) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & Space(20) & Space(20) & PSTR(RstRep!CityName, 33) & Space(2) & Space(15) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20) & mChr18
            mHeader = mHeader + 1
            TotalNetAmt = TotalNetAmt + Val(RstRep!Net_Amount)
            If Counter = mRec Then isLast = True
            RstRep.MoveNext
        Else
            If isLast Then
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = 0: Counter = 0
                isLast = False
                Print #1, Space(PageWidth / 2) & "Page :" & PageNo + 1
                PageNo = PageNo + 1
                Print #1, mEject
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT("Vehicle Sale Register", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & Space(5) & Space(10) & PSTR("Sales", 20) & PSTR("Order", 20) & PSTR("Name of Customer", 35) & PSTR("Model", 15) & PSTR("Telco", 15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
                mHeader = mHeader + 1
                Print #1, PSTR("#", 5) & PSTR("Invoice", 10) & PSTR("Date", 20) & PSTR("Date", 20) & PSTR("Address", 35) & PSTR("Chassis-No", 15) & PSTR("Bill No", 15) & PSTR("Sale-Amt", 15) & PSTR("Financer Group", 35) & Space(15) & PSTR("Form", 5) & PSTR("Spl.Info", 20)
                mHeader = mHeader + 1
                Print #1, PSTR("#", 5) & PSTR("Prefix", 10) & PSTR("Inv-No", 20) & PSTR("No.", 20) & PSTR("Name Of City", 35) & PSTR("Engine No", 15) & PSTR("Date", 15) & Space(15) & PSTR("Financer Name", 35) & PSTR("Fin-Amt", 15) & PSTR("Type", 5) & Space(20) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
            End If
        End If
    
    Wend
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, Space(5) & Space(10) & Space(20) & Space(20) & Space(33) & Space(2) & Space(20) & Space(15) & Space(15) & Space(35) & PSTR("Total --- >", 20) & PSTR(IIf(TotalNetAmt = 0, "", STR(TotalNetAmt)), 20, , AlignRight)
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt > Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
    End If
End Sub
Private Sub SpeedPrintDet()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double, mCounter As Double
    Dim isLast As Boolean, mRec As Integer, PageNo As Double
    Dim TotalInciChrg As Double, TotalRegTemp As Double, TotalTaxAmt As Double
    Dim TotalTOTAmt As Double, TotalNetAmount As Double, TotalOctroi As Double
    Dim TotalSurchargeAmt As Double, TotalTransport As Double, TotalOtherChrg As Double
    Dim RstCompDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    PageLength = PubPageLength
    PageWidth = 132
    mRec = 7
    'Header printing
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!S_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
    If PubComp_Add2 <> "" Then
        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    
    Print #1, PRN_TIT("Vehicle Sale Register", "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, "From : " & FGrid.TextMatrix(Date1, 1) & "  To : " & FGrid.TextMatrix(Date2, 1)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & Space(5) & Space(6) & PSTR("Sales", 18) & PSTR("Order", 18) & PSTR("Name of Customer", 35) & PSTR("Model", 15) & PSTR("Telco", 15) & Space(15) & Space(30) & Space(15) & Space(5) & Space(10) & Space(10) & Space(10) & Space(10) & Space(10)
    mHeader = mHeader + 1
    Print #1, PSTR("#", 5) & PSTR("Inv.", 6) & PSTR("Date", 18) & PSTR("Date", 18) & PSTR("Address", 30) & PSTR("Chassis-No", 15) & PSTR("Bill No", 15) & PSTR("Sale-Amt", 15) & PSTR("Financer Group", 35) & Space(15) & PSTR("Form", 5) & PSTR("Inci-Chg", 10) & PSTR("RegTemp", 10) & PSTR("Tax Amt", 10) & PSTR(pubTOTCaption & " Amt", 10) & PSTR("NET Amt", 10)
    mHeader = mHeader + 1
    Print #1, PSTR("#", 5) & PSTR("Pref.", 6) & PSTR("Inv-No", 18) & PSTR("No.", 18) & PSTR("Name Of City", 30) & PSTR("Engine No", 15) & PSTR("Date", 15) & Space(15) & PSTR("Financer Name", 35) & PSTR("Fin-Amt", 15) & PSTR("Type", 5) & PSTR("Octroi", 10) & PSTR("Transport", 10) & PSTR("Sur Amt", 10) & PSTR("MisChrg", 10) & Space(10) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    RstRep.MoveFirst
    mHeader = 1
    While Not RstRep.EOF = True
        If Counter <= mRec Then
            Counter = Counter + 1
            mCounter = mCounter + 1
            Print #1, mChr17 & PSTR(STR(mCounter), 5) & PSTR(mID(RstRep!Inv_DocId, 9, 5), 6) & PSTR("'" & RstRep!Inv_Date & "'", 18) & PSTR("'" & RstRep!Ord_Date & "'", 18) & PSTR(RstRep!Name, 30) & Space(2) & PSTR(RstRep!Model, 20) & PSTR("'" & RstRep!PBILL_NO & "'", 15) & PSTR(IIf(RstRep!vrate + RstRep!Margine = 0, "", STR(RstRep!vrate + RstRep!Margine)), 15) & PSTR(RstRep!FinGrpName, 35) & PSTR(IIf(RstRep!Fin_Amt = 0, "", STR(RstRep!Fin_Amt)), 15) & PSTR(RstRep!Form_Code, 5) & PSTR(IIf(RstRep!InciChrg = 0, "", STR(RstRep!InciChrg)), 10) & PSTR(IIf(RstRep!RegTemp = 0, "", STR(RstRep!RegTemp)), 10) & PSTR(IIf(RstRep!Tax_Amt = 0, "", STR(RstRep!Tax_Amt)), 10) & PSTR(IIf(RstRep!Tot_Amt = 0, "", STR(RstRep!Tot_Amt)), 10) & PSTR(IIf(RstRep!Net_Amount = 0, "", STR(RstRep!Net_Amount)), 10)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(6) & PSTR(PrinID(RstRep!Inv_DocId), 18) & PSTR(PrinID(RstRep!OrdDocId), 18) & PSTR(RstRep!Add1, 30) & Space(2) & PSTR(RstRep!ChassisNo, 20) & PSTR("'" & RstRep!PBILL_DATE & "'", 15) & Space(15) & PSTR(RstRep!FinName, 35) & Space(15) & Space(5) & PSTR(IIf(RstRep!Octroi = 0, "", STR(RstRep!Octroi)), 10) & PSTR(IIf(RstRep!Transport = 0, "", STR(RstRep!Transport)), 10) & PSTR(IIf(RstRep!Surcharge_Amt = 0, "", STR(RstRep!Surcharge_Amt)), 10) & PSTR(IIf(RstRep!OtherChrg = 0, "", STR(RstRep!OtherChrg)), 10) & Space(10)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(6) & Space(18) & Space(18) & PSTR(RstRep!Add2, 30) & Space(2) & PSTR(RstRep!EngineNo, 20) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(10) & Space(10) & Space(10) & Space(10) & Space(10)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(6) & Space(18) & Space(18) & PSTR(RstRep!Add3, 30) & Space(2) & Space(15) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(10) & Space(10) & Space(10) & Space(10) & Space(10)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(6) & Space(18) & Space(18) & PSTR(RstRep!CityName, 30) & Space(2) & Space(15) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(10) & Space(10) & Space(10) & Space(10) & Space(10) & mChr18
            mHeader = mHeader + 1
            
            TotalInciChrg = TotalInciChrg + Val(RstRep!InciChrg)
            TotalRegTemp = TotalRegTemp + Val(RstRep!RegTemp)
            TotalTaxAmt = TotalTaxAmt + Val(RstRep!Tax_Amt)
            TotalTOTAmt = TotalTOTAmt + Val(RstRep!Tot_Amt)
            TotalNetAmount = TotalNetAmount + Val(RstRep!Net_Amount)
            TotalOctroi = TotalOctroi + Val(RstRep!Octroi)
            TotalTransport = TotalTransport + Val(RstRep!Transport)
            TotalSurchargeAmt = TotalSurchargeAmt + Val(RstRep!Surcharge_Amt)
            TotalOtherChrg = TotalOtherChrg + Val(RstRep!OtherChrg)
            
            If Counter = mRec Then isLast = True
            RstRep.MoveNext
        Else
            If isLast Then
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = 0: Counter = 0
                isLast = False
                Print #1, Space(PageWidth / 2) & "Page :" & PageNo + 1
                PageNo = PageNo + 1
                Print #1, mEject
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT("Vehicle Sale Register", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & Space(5) & Space(6) & PSTR("Sales", 18) & PSTR("Order", 18) & PSTR("Name of Customer", 30) & PSTR("Model", 15) & PSTR("Telco", 15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(10) & Space(10) & Space(10) & Space(10) & Space(10)
                mHeader = mHeader + 1
                Print #1, PSTR("#", 5) & PSTR("Invoice", 6) & PSTR("Date", 18) & PSTR("Date", 18) & PSTR("Address", 30) & PSTR("Chassis-No", 15) & PSTR("Bill No", 15) & PSTR("Sale-Amt", 15) & PSTR("Financer Group", 35) & Space(15) & PSTR("Form", 5) & PSTR("Inci-Chg", 10) & PSTR("RegTemp", 10) & PSTR("Tax Amt", 10) & PSTR(pubTOTCaption & " Amt", 10) & PSTR("NET Amt", 10)
                mHeader = mHeader + 1
                Print #1, PSTR("#", 5) & PSTR("Prefix", 6) & PSTR("Inv-No", 18) & PSTR("No.", 18) & PSTR("Name Of City", 30) & PSTR("Engine No", 15) & PSTR("Date", 15) & Space(15) & PSTR("Financer Name", 35) & PSTR("Fin-Amt", 15) & PSTR("Type", 5) & PSTR("Octroi", 10) & PSTR("Transport", 10) & PSTR("Sur Amt", 10) & PSTR("MisChrg", 10) & Space(10) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
            End If
        End If
   
    Wend
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & Space(5) & Space(6) & Space(18) & Space(18) & Space(30) & Space(2) & Space(20) & Space(15) & Space(15) & Space(35) & Space(15) & PSTR(IIf(TotalInciChrg = 0, "", STR(TotalInciChrg)), 10) & Space(1) & PSTR(IIf(TotalRegTemp = 0, "", STR(TotalRegTemp)), 10) & Space(1) & PSTR(IIf(TotalTaxAmt = 0, "", STR(TotalTaxAmt)), 10) & Space(1) & PSTR(IIf(TotalTOTAmt = 0, "", STR(TotalTOTAmt)), 10) & Space(1) & PSTR(IIf(TotalNetAmount = 0, "", STR(TotalNetAmount)), 10)
    Print #1, mChr17 & Space(5) & Space(6) & Space(18) & Space(18) & Space(30) & Space(2) & Space(20) & Space(15) & Space(15) & Space(35) & Space(15) & PSTR(IIf(TotalOctroi = 0, "", STR(TotalOctroi)), 10) & Space(1) & PSTR(IIf(TotalTransport = 0, "", STR(TotalTransport)), 10) & Space(1) & PSTR(IIf(TotalSurchargeAmt = 0, "", STR(TotalSurchargeAmt)), 10) & Space(1) & PSTR(IIf(TotalOtherChrg = 0, "", STR(TotalOtherChrg)), 10) & Space(10)
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt > Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
    End If
End Sub
Private Sub VehSaleCancelRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

    Condstr = " Where VO.Inv_UEntDt  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_UEntDt <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VO.Inv_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_DocId,1) in (" & GridString2 & ")"
    
    
      
    mQry = "SELECT VO.Inv_DocId, VO.Inv_Date, VO.OrdDocId,VO.Ord_Date,SG.NamePrefix, SG.Name, SG.FPrefix, SG.FName, SG.Add1, SG.Add2, SG.Add3, City.CityName," & _
        "" & cIIF("'" & UCase(left(PubComp_Name, 4)) & "'='ENAR'", "Model.Sales_Desc", "Model.Model") & " as Model,VO.Chassis as ChassisNo,VStk.EngineNo,VO.VRATE, VStk.PBILL_NO, VStk.PBILL_DATE," & _
        "VO.MARGINE, VO.InciChrg, VO.Octroi, VO.Transport, VO.RegTemp," & _
        "VO.TAX_Amt,VO.Surcharge_Amt, VO.OtherChrg, VO.Net_AMOUNT,VO.Inv_UEntDt as CancelDate," & _
        "VO.Form_Code,TF.Form_Desc,FinGroup.FinGrpName,ContractFinance.FinName,VO.Fin_Amt,VO.MISC_INFO,VO.TOT_Amt" & _
        " FROM (((((((Veh_Order1 as VO LEFT JOIN Veh_Stock as VStk ON VO.chassis = VStk.ChassisNo) " & _
        "LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
        "LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
        "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
        "LEFT JOIN ContractFinance ON VO.FB_CODE = ContractFinance.FinCode) " & _
        "LEFT JOIN FinGroup ON ContractFinance.UnderFinGrp = FinGroup.FinGrpCode) " & _
        "LEFT JOIN TaxForms as TF ON VO.Form_Code = TF.Form_Code) "

    mQry = mQry + Condstr & "order by VO.Inv_Date,VO.Inv_DocId"

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "VehSaleCancelReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub VehInTransProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
'    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

    Condstr = " where VP1.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(VP1.site_code,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(VP1.site_code,1) ='" & PubSiteCode & "' "
    End If
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VP1.Docid,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VStk.Model in (" & GridString3 & ")"
   
   mQry = "SELECT VP1.V_NO, VP1.V_Date, VP1.PBILL_NO, VP1.PBILL_DATE," & _
          "SubGroup.Name, ColMast.Col_Desc,VP1.DocId," & _
          "VStk.EngineNo, VStk.SDM_STM_NO, VStk.INDATE," & _
          "(" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " - VP1.V_Date) AS TDay, " & _
          " VStk.VRATE, Model.Model,VStk.ChassisNo" & _
          " FROM (((Veh_Purch1 as VP1 LEFT JOIN Veh_Stock as VStk ON VP1.DocID = VStk.Pur_DocId) " & _
          " LEFT JOIN SubGroup ON VP1.PARTYCODE = SubGroup.SubCode)" & _
          " LEFT JOIN ColMast ON VStk.Colour_Code = ColMast.Col_Code)" & _
          " LEFT JOIN Model ON VStk.MODEL = Model.MODEL"
  
    mQry = mQry + Condstr + "AND Subgroup.nature='Supplier' And VStk.InDate Is Null"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub

    RepName = "VehInTrans"

    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub PurchRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
FormulaStr1 = ""
FormulaStr2 = ""
FormulaStr3 = ""
FormulaStr4 = ""

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List3, FGrid.TextMatrix(List3, 1)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
   
     Condstr = "  VP1.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VP1.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
'     If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = Condstr & " and (VStk.TAX_YN = 1 or VStk.TAX_YN = 0) and VP1.V_Type='V_PB'"
'     If FGrid.TextMatrix(List1, 1) = "Taxable" Then Condstr = Condstr & " and VStk.TAX_YN  = 1 and VP1.V_Type='V_PB'"
'     If FGrid.TextMatrix(List1, 1) = "TaxPaid" Then Condstr = Condstr & " and Vstk.TAX_YN=0 and VP1.V_Type='V_PB'"

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(VP1.site_code,1) in (" & GridString1 & ")"
    
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(VP1.site_code,1) ='" & PubSiteCode & "' "
    End If
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Model.Sales_Desc in (" & GridString2 & ")"
    Else
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VP1.docid,1) in (" & GridString2 & ")"
    End If
    
    If FGrid.TextMatrix(List4, 1) = "Yes" Then
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and MODEL.Grp_Code  in (" & GridString3 & ")"
    Else
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VStk.MODEL  in (" & GridString3 & ")"
    End If
    
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VStk.PartyCode  in (" & GridString4 & ")"

'    If FGrid.TextMatrix(List1, 1) = "All" Then 'CondStr = CondStr & " and isnull(Job_Card.JobCloseDate)"
    
      mQry = "SELECT SG.Name,BMS.BMS_Name,ColMast.Col_Desc,VP1.DocID,VP1.V_No,VP1.PBILL_NO AS BillNo,VP1.PBILL_DATE AS BillDate," & _
            "" & cIIF("'" & UCase(left(PubComp_Name, 4)) & "'='ENAR'", "Model.Sales_Desc", IIf(FGrid.TextMatrix(List4, 1) = "Yes", "MG.ModelGrp_Name", "VStk.Model")) & " as MODEL,VStk.ChassisNo,VStk.EngineNo,VStk.INDATE,VP1.V_Date As PurchDate,VP1.V_NO AS MemoNo,VP1.Exsice," & _
            "VP1.AMOUNT,VP1.Addition,VP1.Deduction,((" & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ") - VP1.PBILL_DATE) AS AgeDays," & _
            "VP1.TaxSur_Amt, VP1.MISC_AMT,VP1.Tot_Amount,VStk.VRate,VStk.Rate,"
            
        mQry = mQry & "VP1.TAX_Amt AS TaxAmt "
    mQry = mQry & ",VP1.AMOUNT+VP1.Addition-VP1.Deduction As PurVRate, VP1.SubventionCredit, VP1.Exsice, " & cIIF("Tf.L_C='Local'", "VP1.TAX_Amt", 0) & " As TaxAmt_Local, " & cIIF("Tf.L_C='Central'", "VP1.TAX_Amt", 0) & " As TaxAmt_Central,Model.Model_Desc " & _
           " FROM (((((((Veh_Purch1 as VP1 LEFT JOIN Veh_Stock as VStk ON VP1.DocID = VStk.Pur_DocId)) " & _
           " LEFT JOIN SubGroup as SG ON VStk.PartyCode = SG.SubCode) " & _
           " Left Join TaxForms Tf On VP1.Form_Code = Tf.Form_Code) " & _
           " LEFT JOIN Model ON VStk.Model = Model.Model) " & _
           " Left Join Model_Grp MG On Model.Grp_Code=MG.ModelGrp_Code)" & _
           " LEFT JOIN ColMast On ColMast.Col_Code=VStk.Colour_Code)" & _
           " LEFT JOIN BMS On BMS.BMS_Code=VP1.BMS_Category " & _
           " WHERE VStk.Pur_VType='V_PB' and " & Condstr
    
    If FGrid.TextMatrix(List3, 1) = "Voucher No" Then
        mQry = mQry & "Order By VP1.V_Date,VP1.Docid"
    Else
        mQry = mQry & "Order By VP1.PBILL_NO"
    End If

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    If UCase(left(PubComp_Name, 4)) <> "ENAR" Then
        If StrCmp(FGrid.TextMatrix(List1, 1), "Detail") Then
            RepName = "PurchReg"
        Else
            RepName = "PurchRegSumm"
        End If
    Else
        RepName = "PurchReg_eNAR"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub VehSalePurRepProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " WHERE VStk.Pur_vDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VStk.Pur_vDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and ( " & cMID("VStk.Pur_DocId", "3", "1") & " in (" & GridString1 & ") or VStk.Chassis_RctDivCode in (" & GridString1 & "))"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("VStk.Pur_DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and VP1.BMS_CATEGORY  in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and (left(VStk.Pur_DocId ,1) in (" & GridString3 & ") or Chassis_RctSiteCode in (" & GridString3 & "))"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and left(VStk.PBILL_NO,1) in (" & GridString4 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "V-No" Then
         mQry = "SELECT " & _
                "VStk.Pur_DocID, VStk.Pur_VDate," & _
                "VStk.PBILL_NO, VStk.PBILL_DATE, VP1.OBNO, VP1.OBDate, S1.NamePrefix, S1.Name,S1.FPrefix, S1.FName, S1.Add1, S1.Add2, S1.Add3, S1.AREA, City.CityName, " & _
                "" & cIIF("SubGroup.Nature='Supplier'", "SubGroup.name", "''") & " AS P_PartyName, VStk.MODEL," & _
                "VStk.ChassisNo, VStk.EngineNo, VO.Inv_DocId," & _
                "VO.Inv_Date, VO.TAX_Amt AS SaleTax, VO.Surcharge_Amt AS SaleSurAmt," & _
                "VO.Net_AMOUNT AS Net_InvVal,VStk.Vrate, " & cIIF("S1.Nature='Customer'", "S1.name", "''") & " AS S_PartyName,VO.Rebate,VO.SpecialDiscount,VP1.Deduction,M.Model_Desc" & _
            " FROM (((((Veh_Stock as Vstk LEFT JOIN Veh_Order as VO ON VStk.Sal_DocId = VO.inv_DocId)" & _
                "LEFT JOIN SubGroup ON VStk.PartyCode = SubGroup.SubCode)" & _
                "LEFT JOIN SubGroup AS S1 ON VO.PartyCode=S1.SubCode) " & _
                "LEFT JOIN City ON SubGroup.CityCode = City.CityCode)" & _
                "LEFT JOIN Veh_Purch1 as VP1 ON VStk.Pur_DocId = VP1.DocID) LEFT JOIN Model M ON Vstk.Model=M.Model"
                mQry = mQry + Condstr & "  order by VStk.Pur_DocID "
    ElseIf FGrid.TextMatrix(List1, 1) = "Telco Inv-No" Then
        mQry = "SELECT " & _
                "RIGHT(VStk.Pur_DocID,13) AS PVNo, VStk.Pur_VDate," & _
                "VStk.PBILL_NO, VStk.PBILL_DATE, VP1.OBNO, VP1.OBDate, S1.NamePrefix, S1.Name,S1.FPrefix, S1.FName, S1.Add1, S1.Add2, S1.Add3, S1.AREA, City.CityName, " & _
                "" & cIIF("SubGroup.Nature='Supplier'", "SubGroup.name", "''") & " AS P_PartyName, VStk.MODEL," & _
                "VStk.ChassisNo, VStk.EngineNo,VO.Inv_docId," & _
                "VO.Inv_Date, VO.TAX_Amt AS SaleTax, VO.Surcharge_Amt AS SaleSurAmt," & _
                "VO.Net_AMOUNT AS Net_InvVal,VStk.Vrate, " & cIIF("S1.Nature='Customer'", "S1.name", "''") & " AS S_PartyName,VO.Rebate,VO.SpecialDiscount,VP1.Deduction,M.Model_Desc" & _
            " FROM (((((Veh_Stock as Vstk LEFT JOIN Veh_Order as VO ON VStk.Sal_DocId = VO.inv_DocId) " & _
                "LEFT JOIN SubGroup ON VStk.PartyCode = SubGroup.SubCode)" & _
                "LEFT JOIN SubGroup AS S1 ON VO.PartyCode=S1.SubCode) " & _
                "LEFT JOIN City ON SubGroup.CityCode = City.CityCode)" & _
                "LEFT JOIN Veh_Purch1 as VP1 ON VStk.Pur_DocId = VP1.DocID)  LEFT JOIN Model M ON Vstk.Model=M.Model "
                mQry = mQry + Condstr & "  order by VStk.PBILL_NO ASC"
    End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        
    RepName = "VehSalePurRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub VehBookingProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub

    Condstr = " where Veh_Order.Ord_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Ord_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    'CondStr1 = " And Veh_Order.Ord_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and Veh_Order.Ord_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & "# "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Veh_Order.OrdDocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_Order.OrdDocId", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.OrdDocId,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.MODEL  in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.PartyCode in (" & GridString4 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "All" Then 'CondStr = CondStr & " and isnull(Job_Card.JobCloseDate)"
         mQry = "SELECT Veh_Order.OrdDocId,Veh_Order.Ord_Date AS VehOrdDate, right(Veh_Order.OrdDocId,13) AS VehOrdNo, veh_order.Model," & _
            "SubGroup.Name AS PartyName,SubGroup.NamePrefix,City.CityName,SubGroup.Add1,SubGroup.Add2, Rect.V_Date AS RectDate, Rect.V_No AS RectNo,FinName," & _
            "Rect.AMOUNT AS RectAmt, Rect.Narration AS RectNarr, Rect.Narration1 AS RectNarr1," & _
            "ColMast.Col_Desc AS Colour, Veh_Order.Inv_No,Veh_Stock.Pbill_No,Veh_Stock.Pbill_Date,Veh_Order.vrate," & _
            "Veh_Order.Inv_Date AS InvDate, Veh_Order.Net_Amount  AS InvoiceAmt,Veh_Order.Inv_DocId" & _
            " FROM (((((Veh_Order LEFT JOIN" & _
            "(Rect LEFT JOIN  " & FaTable("Voucher_Type") & " ON Rect.V_Type = Voucher_Type.V_Type) ON Veh_Order.OrdDocId = Rect.ord_DocId)" & _
            " LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode)" & _
            " LEFT JOIN ColMast ON Veh_Order.Colour_Code = ColMast.Col_Code)" & _
            " LEFT JOIN ContractFinance  on ContractFinance.FinCode=Veh_order.FB_Code)" & _
            " Left Join City On City.CityCode = SubGroup.CityCode)" & _
            " Left Join Veh_Stock On Veh_Stock.Ord_DocId=Veh_Order.OrdDocId"
          
    End If
    If FGrid.TextMatrix(List1, 1) = "Pending" Then
        Condstr = Condstr & " and Veh_Order.Inv_Date Is Null Or Veh_Order.Inv_Date  > " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
        mQry = "SELECT Veh_Order.OrdDocId,Veh_Order.Ord_Date AS VehOrdDate, Veh_Order.Ord_No AS VehOrdNo, veh_order.Model," & _
            "SubGroup.Name AS PartyName,SubGroup.NamePrefix,City.CityName,SubGroup.Add1,SubGroup.Add2, Rect.V_Date AS RectDate, Rect.V_No AS RectNo,FinName," & _
            "Rect.AMOUNT AS RectAmt, Rect.Narration AS RectNarr, Rect.Narration1 AS RectNarr1," & _
            "ColMast.Col_Desc AS Colour, Veh_Order.Inv_No,Veh_Stock.Pbill_No,Veh_Stock.Pbill_Date,Veh_Order.vrate," & _
            "'#01/04/04#' AS InvDate, Veh_Order.Net_Amount  AS InvoiceAmt,'' as Inv_DocId" & _
            " FROM (((((Veh_Order LEFT JOIN" & _
            "(Rect LEFT JOIN  " & FaTable("Voucher_Type") & " ON Rect.V_Type = Voucher_Type.V_Type) ON Veh_Order.OrdDocId = Rect.ord_DocId)" & _
            " LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode)" & _
            " LEFT JOIN ColMast ON Veh_Order.Colour_Code = ColMast.Col_Code)" & _
            " LEFT JOIN ContractFinance  on ContractFinance.FinCode=Veh_order.FB_Code)" & _
            " Left Join City On City.CityCode = SubGroup.CityCode)" & _
            " Left Join Veh_Stock On Veh_Stock.Ord_DocId=Veh_Order.OrdDocId"
    End If
        
    If FGrid.TextMatrix(List1, 1) = "Supplied" Then
        Condstr = Condstr & " and (Veh_Order.Inv_Date Is Not Null Or Veh_Order.Inv_Date  > " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ")"
        mQry = "SELECT Veh_Order.OrdDocId,Veh_Order.Ord_Date AS VehOrdDate, Veh_Order.Ord_No AS VehOrdNo, veh_order.Model," & _
            "SubGroup.Name AS PartyName,SubGroup.NamePrefix,City.CityName,SubGroup.Add1,SubGroup.Add2, Rect.V_Date AS RectDate, Rect.V_No AS RectNo,FinName," & _
            "Rect.AMOUNT AS RectAmt, Rect.Narration AS RectNarr, Rect.Narration1 AS RectNarr1," & _
            "ColMast.Col_Desc AS Colour, Veh_Order.Inv_No,Veh_Stock.Pbill_No,Veh_Stock.Pbill_Date,Veh_Order.vrate," & _
            "Veh_Order.Inv_Date AS InvDate, Veh_Order.Net_Amount  AS InvoiceAmt,Veh_Order.Inv_DocId" & _
            " FROM (((((Veh_Order LEFT JOIN" & _
            "(Rect LEFT JOIN  " & FaTable("Voucher_Type") & " ON Rect.V_Type = Voucher_Type.V_Type) ON Veh_Order.OrdDocId = Rect.ord_DocId)" & _
            " LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode)" & _
            " LEFT JOIN ColMast ON Veh_Order.Colour_Code = ColMast.Col_Code)" & _
            " LEFT JOIN ContractFinance  on ContractFinance.FinCode=Veh_order.FB_Code)" & _
            " Left Join City On City.CityCode = SubGroup.CityCode)" & _
            " Left Join Veh_Stock On Veh_Stock.Ord_DocId=Veh_Order.OrdDocId"
    End If
    
    mQry = mQry + Condstr
  
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        
    RepName = "VehBookingReg"
       
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
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then Exit Sub
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

Public Function ListView_Items_RecordSet_Local(LV As Object, Txt As Object, Index As Integer, Rst As ADODB.Recordset) As ListItem
    Dim xName As ListItem
    Dim I As Long
    LV.ListItems.Clear
        
    If Rst.RecordCount <= 0 Then Exit Function
    Set xName = LV.ListItems.Add(, , "All")
    xName.SubItems(1) = ""
    Do Until Rst.EOF
        Set xName = LV.ListItems.Add(, , Rst.Fields("Name").Value)
        If Not IsNull(Rst.Fields("Code").Value) Then
            xName.SubItems(1) = CStr(Rst.Fields("code").Value)
        End If
    Rst.MoveNext
    Loop
    Set xName = LV.FindItem(Txt(Index), 0, , 1)
    If xName Is Nothing Then
        Exit Function
    Else
        xName.EnsureVisible
        xName.SELECTED = True
    End If
    Set ListView_Items_RecordSet_Local = xName
End Function

Private Sub SalesManPenAmtProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
Dim TmpRst As ADODB.Recordset
FormulaStr1 = ""
FormulaStr2 = ""
FormulaStr3 = ""
FormulaStr4 = ""

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
    
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " and  Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Ord_sitecode,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(Veh_Order.Ord_SiteCode,1)  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.Rep_Code  in   (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and  left(Veh_Order.Orddocid,1) in (" & GridString3 & ")"
    
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "SalesMan", adChar, 50, adFldIsNullable
        .Fields.Append "Party", adChar, 50, adFldIsNullable
        .Fields.Append "FinName", adChar, 50, adFldIsNullable
        .Fields.Append "Model", adChar, 50, adFldIsNullable
        .Fields.Append "ChlDate", adDate, 20, adFldIsNullable
        .Fields.Append "OstdAmt", adDouble, 20, adFldIsNullable
        .Fields.Append "Model_Desc", adChar, 50, adFldIsNullable
        .Fields.Append "Inv_No", adChar, 20, adFldIsNullable
        .Fields.Append "InvDate", adDate, 20, adFldIsNullable
        
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    
    
    Set TmpRst = GCn.Execute("Select * from Veh_Order LEFT JOIN Model M ON Veh_Order.MODEL =M.MODEL where Fin_Amt>0 and DelCh_DocId<>''  " & Condstr)
    
    If TmpRst.RecordCount > 0 Then
        Do While TmpRst.EOF = False
            If GCn.Execute("Select * from Rect where Rect.PARTYCODE='" & TmpRst!PartyCode & "' and Rect.RectCatg='BAL'").RecordCount > 0 Then
            Else
                With RstRep
                    .AddNew
                    .Fields("SalesMan") = GCn.Execute("Select Emp_Name from Emp_Mast where Emp_Code='" & TmpRst!REP_CODE & "'").Fields(0).Value
                    .Fields("Party") = G_FaCn.Execute("Select Name from SubGroup where SubCode='" & TmpRst!PartyCode & "'").Fields(0).Value
                     If GCn.Execute("Select FinName from ContractFinance where  FinCode ='" & TmpRst!FB_Code & "'").RecordCount > 0 Then
                        .Fields("FinName") = GCn.Execute("Select FinName from ContractFinance where  FinCode ='" & TmpRst!FB_Code & "'").Fields(0).Value
                    End If
                    .Fields("Model") = XNull(TmpRst!Model)
                    .Fields("ChlDate") = VNull(TmpRst!DelCh_DT)
                    .Fields("OstdAmt") = VNull(TmpRst!Fin_Amt)
                    .Fields("Model_Desc") = XNull(TmpRst!Model_Desc)
                    .Fields("Inv_No") = XNull(TmpRst!Inv_No)
                    .Fields("InvDate") = VNull(TmpRst!Inv_Date)
                    .Update
                   
                End With
            End If
        TmpRst.MoveNext
        Loop
    End If
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "SalesPendAmt"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ProcVehicleBillWiseOutstandingFIFO()
On Error GoTo ELoop
Dim mQry$, Condstr$, CondDivCode$, CondMarkYN$, CondPartNos$, CondPartNos1$, CondDivCode1$
Dim CondStrMRP$, CondStropbal$, CondPartNosOpStk$, CondPartNosTrn$, Part_Name$, DivStr1$
Dim mRecQty As Double, mIssQty As Double, mStkVal As Double, noofDiv As Integer, d As Integer
Dim XRecNo As Double, DivStr$
Dim mQty As Double, mReqQty As Double, mFifoCost As Double
Dim mNo As Byte, NoUpto As Byte
Dim RstDiv As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim RsPart As ADODB.Recordset
Dim RsParty As ADODB.Recordset
    RepPrint = True
    GridString1 = Empty: GridString2 = Empty: GridString3 = Empty
    
    
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
    
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("O.Inv_DocID", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("O.Inv_DocID", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and O.Rep_Code  in   (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and  Left(O.Inv_DocId,1) in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and  O.FB_Code in (" & GridString4 & ")"
    
    
    
            
           
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Party_Name", adVarChar, 100, adFldIsNullable
        .Fields.Append "Bill_No", adVarChar, 21, adFldIsNullable
        .Fields.Append "Bill_Amt", adDouble, 12, adFldIsNullable
        .Fields.Append "OutStd_Amt", adDouble, 12, adFldIsNullable
        .Fields.Append "Finanser", adVarChar, 100, adFldIsNullable
        .Fields.Append "SalesMan", adVarChar, 100, adFldIsNullable
        .Fields.Append "Bill_Date", adDate, , adFldIsNullable
        .Fields.Append "Site_Desc", adVarChar, 100, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
           
           
    Dim bNetOsd As Double
    Dim bBillOsd As Double
           
    mQry = " Select L.SubCode, Sum(L.AmtDr-L.AmtCr) As Balance From Ledger L Where L.SubCode In (SELECT DISTINCT O.PartyCode  FROM Veh_Order O) And L.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "  Group By L.SubCode Having Sum(L.AmtDr-L.AmtCr)>0 "
    Set RsParty = GCn.Execute(mQry)
    
    With RsParty
        If RsParty.RecordCount > 0 Then
            Do While Not RsParty.EOF
                bNetOsd = VNull(RsParty!Balance)
                Set RsTemp = GCn.Execute("SELECT O.Inv_DocId, O.Inv_Date, O.Net_Amount, O.PartyCode, S.NAME + ', ' + ISNULL(CS.CityName,'') AS Party, F.FinName + ', ' + ISNULL(CF.CityName,'' ) AS Financer, E.Emp_Name,s1.Site_Desc " & _
                                        "FROM Veh_Order O " & _
                                        "LEFT JOIN SubGroup S ON O.PartyCode = S.SubCode " & _
                                        "LEFT JOIN City CS ON S.CityCode = CS.CityCode " & _
                                        "LEFT JOIN ContractFinance F ON O.FB_CODE = F.FinCode " & _
                                        "LEFT JOIN City CF ON F.City = CF.CityCode " & _
                                        "LEFT JOIN Emp_Mast E ON O.REP_CODE = E.Emp_Code " & _
                                        "LEFT JOIN Site s1 ON RIGHT(O.Inv_SiteCode,1 )=s1.Site_Code  " & _
                                        "WHERE ISNULL(O.Inv_DocId,'')<>'' " & _
                                        "AND O.PartyCode = '" & XNull(RsParty!SubCode) & "' And O.Inv_Date <=  " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "  " & Condstr & _
                                        "Order By O.Inv_Date Desc")
                If RsTemp.RecordCount > 0 Then
                    Do Until RsTemp.EOF
                        bBillOsd = 0
                        If bNetOsd > 0 Then
                            bBillOsd = IIf(VNull(RsTemp!Net_Amount) > bNetOsd, bNetOsd, VNull(RsTemp!Net_Amount))
                            bNetOsd = bNetOsd - bBillOsd
                            
                            
                            RstRep.AddNew
                            RstRep!Party_Name = RsTemp!Party
                            RstRep!Bill_No = RsTemp!Inv_DocId
                            RstRep!Bill_Amt = RsTemp!Net_Amount
                            RstRep!OutStd_Amt = bBillOsd
                            RstRep!Finanser = RsTemp!Financer
                            RstRep!SalesMan = RsTemp!Emp_Name
                            RstRep!Bill_Date = RsTemp!Inv_Date
                            RstRep!Site_Desc = RsTemp!Site_Desc
                            RstRep.Update
                            
                        End If
                        RsTemp.MoveNext
                    Loop
                End If
                RsParty.MoveNext
            Loop
        End If
    End With
    
   
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        
        
        ' For Speed Printing of report
        RepTitle = UCase(Me.CAPTION)
        
    If StrCmp(FGrid.TextMatrix(List1, 1), "Financer") Then
        RepName = "VehicleBillWiseOutstanding"
    Else
        RepName = "VehicleBillWiseOutstandingSalesman"
    End If

ELoop:
    Set RsTemp = Nothing
    If err.NUMBER <> 0 Then CheckError

End Sub



'Private Sub OutPayRepProc()
'On Error GoTo ELoop
'Dim mQRY As String, Condstr As String, CondStr1 As String
'Dim TmpRst, RstCrAmt As ADODB.Recordset
'Dim RsTemp As ADODB.Recordset
'Dim InvAmt As Double
'FormulaStr1 = ""
'FormulaStr2 = ""
'FormulaStr3 = ""
'FormulaStr4 = ""
'
'    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
'
'
'    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
'    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
'    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
'    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
'
'    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("O.Inv_DocID", "3", "1") & " in (" & GridString1 & ")"
'    If Check1(1).Value = Checked Then
'      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("O.Inv_DocID", "3", "1") & "  ='" & PubSiteCode & "' "
'    End If
'
'    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and O.Rep_Code  in   (" & GridString2 & ")"
'    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and  Left(O.Inv_DocId,1) in (" & GridString3 & ")"
'    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and  O.FB_Code in (" & GridString4 & ")"
'
'
'
'
'
'
'    mQRY = "SELECT F.FinName + ', ' + ISNULL(C.CityName,'') AS Fin_City, S.Name + ', ' + ISNULL(CityParty.CityName,'' ) AS Party, R.Emp_Name  AS Salesman, PartyCode,REP_CODE,O.FB_CODE, " & _
'            "DelCh_DocId , DelCh_DT, Inv_DocId, Inv_Date, Net_Amount, l.balance,s1.Site_Desc " & _
'            "FROM dbo.Veh_Order O " & _
'            "LEFT JOIN SubGroup S ON O.PartyCode = S.SubCode " & _
'            "LEFT JOIN Emp_Mast  R ON O.REP_CODE = R.Emp_Code " & _
'            "LEFT JOIN ContractFinance F ON O.FB_CODE = F.FinCode " & _
'            "LEFT JOIN City C ON F.City = C.CityCode " & _
'            "LEFT JOIN City CityParty ON S.CityCode = CityParty.CityCode " & _
'            "LEFT JOIN (SELECT L.SubCode, SUM(AmtDr-AmtCr) AS Balance FROM Ledger L Where L.V_Date<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " GROUP BY L.SubCode ) L ON L.SubCode = O.PartyCode  LEFT JOIN Site s1 ON RIGHT(O.Inv_SiteCode,1 )=s1.Site_Code " & _
'            "WHERE L.balance>0  And O.Inv_Date <= '" & FGrid.TextMatrix(Date2, 1) & "' "
'    mQRY = mQRY & Condstr
'    mQRY = mQRY & " Order By O.Inv_Date, O.Inv_DocID "
'    Set RstRep = GCn.Execute(mQRY)
'
'
'
'
'
'
'
'    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'    RepName = "OstdPayRep"
'    RepTitle = UCase(Me.CAPTION)
'    Exit Sub
'ELoop:
'    RepPrint = False
'    MsgBox err.Description
'End Sub
Private Sub VehFollowUpProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2$, mDlrName$, mModel$
Dim TmpRst As ADODB.Recordset
Dim LastDt, CurDt As Date
Dim LstJobNo, CurJobNo, CurKms, LstKms  As Double
Dim CurMech, LstMech As String

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub


    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
    Condstr = " where Veh_Order.Inv_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Veh_Order.ordDocid", "3", "1") & " in (" & GridString1 & ")"
    
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Veh_Order.ordDocid", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    Set RstRep = New ADODB.Recordset
    With RstRep
    
        .Fields.Append "Chassis", adVarChar, 25, adFldIsNullable
        .Fields.Append "Model", adVarChar, 25, adFldIsNullable
        .Fields.Append "PartyName", adVarChar, 50, adFldIsNullable
        .Fields.Append "PartyPhone", adVarChar, 30, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 25, adFldIsNullable
        
        .Fields.Append "FirstSerDueOn", adDate, 15, adFldIsNullable
        .Fields.Append "FirstSerDoneDt", adDate, 15, adFldIsNullable
        
        .Fields.Append "SecdSerDueOn", adDate, 15, adFldIsNullable
        .Fields.Append "SecdSerDoneDt", adDate, 15, adFldIsNullable
        
        .Fields.Append "ThirdSerDueOn", adDate, 15, adFldIsNullable
        .Fields.Append "ThirdSerDoneDt", adDate, 15, adFldIsNullable
        
        .Fields.Append "FourSerDueOn", adDate, 15, adFldIsNullable
        .Fields.Append "FourSerDoneDt", adDate, 15, adFldIsNullable
        
        .Fields.Append "FifthSerDueOn", adDate, 15, adFldIsNullable
        .Fields.Append "FifthSerDoneDt", adDate, 15, adFldIsNullable
        
        
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Set RstRep1 = GCn.Execute("Select Veh_Order.*,HisCard.CardNo,HisCard.RegNo from Veh_Order Left join HisCard on Veh_Order.Chassis=HisCard.Chassis " & Condstr)
    If RstRep1.RecordCount > 0 Then
        Do While RstRep1.EOF = False
            With RstRep
                Set TmpRst = GCn.Execute("Select * from Job_Card left join Service_Type ST on Job_Card.Serv_type=ST.Serv_type where CardNo='" & RstRep1!CardNo & "' and ST.Serv_Catg <> 'P' Order By Job_date")
                    If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
                        If TmpRst.RecordCount = 0 Then
                            If DateDiff("D", RstRep1!Inv_Date, PubLoginDate) >= 1 Then
                            'If DateDiff("D", RstRep1!Inv_Date, PubLoginDate) >= 1 Then
                                .AddNew
                                .Fields("Chassis") = RstRep1!Chassis
                                .Fields("Model") = RstRep1!Model
                                .Fields("PartyName") = GCn.Execute("Select Name from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("PartyPhone") = GCn.Execute("Select Phone from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("RegNo") = RstRep1!RegNo
                                .Fields("FirstSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat1, 1))
                                '.Fields("FirstSerDoneDt") = TmpRst!Job_Date
                                
                                .Fields("SecdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat2, 1))
                                '.Fields("SecdSerDoneDt") = ""
                                
                                .Fields("ThirdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat3, 1))
                                '.Fields("ThirdSerDoneDt") = ""
                                
                                .Fields("FourSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat4, 1))
                                '.Fields("FourthSerDoneDt") = ""
                                
                                .Fields("FifthSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat5, 1))
                                '.Fields("FifthSerDoneDt") = ""
                                .Update
                            End If
                        ElseIf TmpRst.RecordCount = 1 Then
                            If DateDiff("D", RstRep1!Inv_Date, PubLoginDate) >= 1 Then
                                .AddNew
                                .Fields("Chassis") = RstRep1!Chassis
                                .Fields("Model") = RstRep1!Model
                                .Fields("PartyName") = GCn.Execute("Select Name from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("PartyPhone") = GCn.Execute("Select Phone from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("RegNo") = RstRep1!RegNo
                                .Fields("FirstSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat1, 1))
                                .Fields("FirstSerDoneDt") = TmpRst!Job_Date
                                
                                'TmpRst.MoveNext
                                .Fields("SecdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat2, 1))
                                '.Fields("SecdSerDoneDt") = TmpRst!Job_Date
                                
                                .Fields("ThirdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat3, 1))
                                '.Fields("ThirdSerDoneDt") = ""
                                
                                .Fields("FourSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat4, 1))
                                '.Fields("FourthSerDoneDt") = ""
                                
                                .Fields("FifthSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat5, 1))
                                '.Fields("FifthSerDoneDt") = ""
                                .Update
                            End If
                            
                        ElseIf TmpRst.RecordCount = 2 Then
                        
                            If DateDiff("D", RstRep1!Inv_Date, PubLoginDate) >= 1 Then
                                .AddNew
                                .Fields("Chassis") = GCn.Execute("Select Chassis from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                                .Fields("Model") = GCn.Execute("Select Model from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                                .Fields("PartyName") = GCn.Execute("Select Name from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("PartyPhone") = GCn.Execute("Select Phone from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("RegNo") = RstRep1!RegNo
                                .Fields("FirstSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat1, 1))
                                .Fields("FirstSerDoneDt") = TmpRst!Job_Date
                                
                                TmpRst.MoveNext
                                .Fields("SecdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat2, 1))
                                .Fields("SecdSerDoneDt") = TmpRst!Job_Date
                                
                                'TmpRst.MoveNext
                                .Fields("ThirdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat3, 1))
                                '.Fields("ThirdSerDoneDt") = TmpRst!Job_Date
                                
                                .Fields("FourSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat4, 1))
                                '.Fields("FourthSerDoneDt") = ""
                                
                                .Fields("FifthSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat5, 1))
                                '.Fields("FifthSerDoneDt") = ""
                                .Update
                            End If
                            
                        ElseIf TmpRst.RecordCount = 3 Then
                            If DateDiff("D", RstRep1!Inv_Date, PubLoginDate) >= 1 Then
                                .AddNew
                                .Fields("Chassis") = GCn.Execute("Select Chassis from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                                .Fields("Model") = GCn.Execute("Select Model from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                                .Fields("PartyName") = GCn.Execute("Select Name from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("PartyPhone") = GCn.Execute("Select Phone from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("RegNo") = RstRep1!RegNo
                                .Fields("FirstSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat1, 1))
                                .Fields("FirstSerDoneDt") = TmpRst!Job_Date
                                
                                TmpRst.MoveNext
                                .Fields("SecdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat2, 1))
                                .Fields("SecdSerDoneDt") = TmpRst!Job_Date
                                
                                TmpRst.MoveNext
                                .Fields("ThirdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat3, 1))
                                .Fields("ThirdSerDoneDt") = TmpRst!Job_Date
                                
                                'TmpRst.MoveNext
                                .Fields("FourSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat4, 1))
                                '.Fields("FourSerDoneDt") = TmpRst!Job_Date
                                
                                .Fields("FifthSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat5, 1))
                                '.Fields("FifthSerDoneDt") = ""
                                .Update
                            End If
                        ElseIf TmpRst.RecordCount = 4 Then
                            If DateDiff("D", RstRep1!Inv_Date, PubLoginDate) >= 1 Then
                                .AddNew
                                .Fields("Chassis") = GCn.Execute("Select Chassis from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                                .Fields("Model") = GCn.Execute("Select Model from HisCard where CardNo='" & RstRep1!CardNo & "'").Fields(0).Value
                                .Fields("PartyName") = GCn.Execute("Select Name from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("PartyPhone") = GCn.Execute("Select Phone from Subgroup where SubCode='" & RstRep1!PartyCode & "'").Fields(0).Value
                                .Fields("RegNo") = RstRep1!RegNo
                                .Fields("FirstSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat1, 1))
                                .Fields("FirstSerDoneDt") = TmpRst!Job_Date
                                
                                TmpRst.MoveNext
                                .Fields("SecdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat2, 1))
                                .Fields("SecdSerDoneDt") = TmpRst!Job_Date
                                
                                TmpRst.MoveNext
                                .Fields("ThirdSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat3, 1))
                                .Fields("ThirdSerDoneDt") = TmpRst!Job_Date
                                
                                TmpRst.MoveNext
                                .Fields("FourSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat4, 1))
                                .Fields("FourSerDoneDt") = TmpRst!Job_Date
                                
                                .Fields("FifthSerDueOn") = RstRep1!Inv_Date + Val(FGrid.TextMatrix(Cat5, 1))
                                '.Fields("FifthSerDoneDt") = ""
                                .Update
                            End If
                        End If
            End With
            RstRep1.MoveNext
        Loop
    End If
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "VehFollowUp"
    RepTitle = UCase(Me.CAPTION)
    SubRep1 = False
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub IncomeTaxRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " Where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("vO.Inv_DocID", "3", "1") & "  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_DocId,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and SG.CityCode in (" & GridString3 & ")"
    
      
    mQry = "SELECT VO.Inv_DocId, VO.Inv_Date, VO.OrdDocId,VO.Ord_Date,SG.NamePrefix, SG.Name, SG.FPrefix, SG.FName, SG.Add1, SG.Add2, SG.Add3, City.CityName," & _
        "Model.Model,VO.Chassis as ChassisNo,VStk.EngineNo,VO.VRATE, VStk.PBILL_NO, VStk.PBILL_DATE," & _
        "VO.MARGINE, VO.InciChrg, VO.Octroi, VO.Transport, VO.RegTemp," & _
        "VO.TAX_Amt,VO.Surcharge_Amt, VO.OtherChrg, VO.Net_AMOUNT," & _
        "VO.Form_Code,TF.Form_Desc,FinGroup.FinGrpName,ContractFinance.FinName,VO.Fin_Amt,VO.MISC_INFO,VO.TOT_Amt,'" & FGrid.TextMatrix(List2, 1) & "' as ReportType,VO.SubTot,sg.phone" & _
        " FROM (((((((Veh_Order as VO LEFT JOIN Veh_Stock as VStk ON VO.chassis = VStk.ChassisNo) " & _
        "LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
        "LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
        "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
        "LEFT JOIN ContractFinance ON VO.FB_CODE = ContractFinance.FinCode) " & _
        "LEFT JOIN FinGroup ON ContractFinance.UnderFinGrp = FinGroup.FinGrpCode) " & _
        "LEFT JOIN TaxForms as TF ON VO.Form_Code = TF.Form_Code) "


NXT:
    mQry = mQry + Condstr & " and Vstk.Sal_Vtype<>'V_TRF' order by SG.Name  "

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "IncomeTaxReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub SubVentionClaimRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1 As String
Dim TmpRst, RstCrAmt As ADODB.Recordset
Dim InvAmt As Double
FormulaStr1 = ""
FormulaStr2 = ""
FormulaStr3 = ""
FormulaStr4 = ""

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
    
    Condstr = " Where Vo.Inv_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And Vo.Inv_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " "
    CondStr1 = " Where Vp.V_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And Vp.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " And VP.V_Type<>'V_OST' "
    If StrCmp(left(PubComp_Name, 3), "LMP") Then
        Condstr = Condstr & " And " & vIsNull("Vo.Subvention", 0) & " > 0 "
    Else
        Condstr = Condstr & " And Vo.Inv_Date Between Sv.FromDate and Sv.ToDate "
    End If

    GridString1 = "": GridString2 = "": GridString3 = "": GridString4 = ""
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Vo.Inv_SiteCode,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  left(Vo.Inv_SiteCode,1)  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Vo.Inv_DocId,1)  in   (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and  Vs.Model in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and  Mg.ModelGrp_Code in (" & GridString4 & ")"
    If StrCmp(left(PubComp_Name, 3), "LMP") Then
        mQry = "Select 'Subvention' As mGroup, Vo.Inv_Date As Inv_Date, Vo.Inv_DocId As Inv_DocId, Vo.Inv_No As Inv_No, Vo.Inv_VType As Inv_VType,Left(Vo.Inv_SiteCode,1) As Site,Vo.SubventionScheme As SubventionScheme, " & _
                " Vo.Subvention as Subvention,VO.DealerContribution As DealerContribution,VO.TataContribution As TataContribution,Vo.Net_Amount As Net_Amount,VS.ChassisNo as ChassisNo,VS.EngineNo As EngineNo,Vs.Model As Model,Mg.ModelGrp_Name As ModelGrp_Name, '' As Category " & _
                " From ((Veh_Order As Vo Left Join Veh_Stock As Vs On Vo.Inv_DocId = Vs.Sal_DocId) " & _
                "                        Left Join Model On Vs.Model = Model.Model) " & _
                "                        Left Join Model_Grp As Mg On Model.Grp_Code=Mg.ModelGrp_Code "
        mQry = mQry & Space(1) & Condstr
    
        mQry = mQry & "Union All Select  'Offtake' As mGroup, Null As Inv_Date, '' As Inv_DocId, 0 As Inv_No, '' As Inv_Vtype, " & PubSiteCode & " As Site, Max(O.SchemeNo) As SubventionScheme, " & _
                " Max(O.Qty) As Subvention, Sum(1) As DealerContribution, Sum(O.Amount) As TataContribution, 0 As NetAmount, '' As ChassisNo, '' As EngineNo, '' As Model, '' As ModelGrp_Name, Max(O.SchemeNo) + ' From ' + " & cCStr("O.FromDate") & " + ' To ' +  " & cCStr("O.ToDate") & "  As Category " & _
                " From ((((Veh_Purch1 As Vp Left Join Veh_Stock As Vs On Vp.DocId = Vs.Pur_DocId) " & _
                "                        Left Join Model On Vs.Model = Model.Model) " & _
                "                        Left Join OffTake1 O1 On Model.Grp_Code = O1.ModelGrp) " & _
                "                        Left Join OffTake O On O.Code=O1.Code And  (V_Date Between O.FromDate And O.ToDate)) " & _
                "                        Left Join Model_Grp As Mg On Model.Grp_Code=Mg.ModelGrp_Code "
        mQry = mQry & Space(1) & CondStr1 & " Group By O.Code, O.FromDate, O.ToDate Having Sum(1) >= Max(O.Qty) "
    Else
        mQry = "Select 'Subvention' As mGroup, Vo.Inv_Date As Inv_Date, Vo.Inv_DocId As Inv_DocId, Vo.Inv_No As Inv_No, Vo.Inv_VType As Inv_VType,Left(Vo.Inv_SiteCode,1) As Site,Sv.SchemeNo As SubventionScheme, " & _
                " Vo.Subvention as Subvention,SV.DealerContribution As DealerContribution,SV.TataContribution As TataContribution,Vo.Net_Amount As Net_Amount,VS.ChassisNo as ChassisNo,VS.EngineNo As EngineNo,Vs.Model As Model,Mg.ModelGrp_Name As ModelGrp_Name, '' As Category " & _
                " From ((((Veh_Order As Vo " & _
                " Left Join Veh_Stock As Vs On Vo.Inv_DocId = Vs.Sal_DocId) " & _
                " Left Join Model On Vs.Model = Model.Model) " & _
                " Left Join Model_Grp As Mg On Model.Grp_Code=Mg.ModelGrp_Code) " & _
                " Left Join Subvention Sv On Model.Grp_Code = Sv.ModelGroup)"
        mQry = mQry & Space(1) & Condstr
    
        mQry = mQry & "Union All Select  'Offtake' As mGroup, Null As Inv_Date, '' As Inv_DocId, 0 As Inv_No, '' As Inv_Vtype, '" & PubSiteCode & "' As Site, Max(O.SchemeNo) As SubventionScheme, " & _
                " Max(O.Qty) As Subvention, Sum(1) As DealerContribution, Sum(O.Amount) As TataContribution, 0 As NetAmount, '' As ChassisNo, '' As EngineNo, '' As Model, '' As ModelGrp_Name, Max(O.SchemeNo) + '      From ' + " & cDt("O.FromDate") & " + ' To ' +  " & cDt("O.ToDate") & "  As Category " & _
                " From ((((Veh_Purch1 As Vp " & _
                " Left Join Veh_Stock As Vs On Vp.DocId = Vs.Pur_DocId) " & _
                " Left Join Model On Vs.Model = Model.Model) " & _
                " Left Join OffTake1 O1 On Model.Grp_Code = O1.ModelGrp) " & _
                " Left Join OffTake O On O.Code=O1.Code) " & _
                " Left Join Model_Grp As Mg On Model.Grp_Code=Mg.ModelGrp_Code "
        mQry = mQry & Space(1) & CondStr1 & "  And  (VP.V_Date Between O.FromDate And O.ToDate) Group By O.Code, O.FromDate, O.ToDate Having Sum(1) >= Max(O.Qty) "
    End If
    mQry = mQry & " Order By Inv_Date"
            
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
     
    RepName = "SubVentionClaimReg"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ChequePaymentRegisterProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim TmpRst, RstCrAmt As ADODB.Recordset
Dim InvAmt As Double
Dim mSubTable As String
    FormulaStr1 = "": FormulaStr2 = "": FormulaStr3 = "": FormulaStr4 = ""

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
    
    Condstr = " Where P.V_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And P.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " "
    'Condstr = Condstr & " And Vt.NCat In ('" & Voucher_NCat_BankPayment & "') "
    
    GridString1 = "": GridString2 = "": GridString3 = "": GridString4 = ""
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(P.Site_Code,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  left(p.Site_Code,1)  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(P.DocId,1)  in   (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and  Sg.SubCode  in (" & GridString3 & ")"
    
    
    mSubTable = "Select L.DocId, Name, L.SubCode From Ledger L Left Join  SubGroup On L.SubCode=SubGroup.SubCode Where SubGroup.Nature='Bank'"
    
    mQry = "Select P.DocId, P.Site_Code, P.V_Date, P.V_Type, P.V_No, P.PartyCode, P.Amount, P.AcCode, P.Chq_No, " & _
            " P.Chq_Date, P.Clg_Date, P.Narration, P.PayTo1, P.PayTo2, P.Printed, P.AcPayeeCheque, P.AcPostByU_Name, " & _
            " P.AcPostByU_EntDt, Vt.Description As VType_Description, Vt.NCat, Sg.Name As AcName, SgP.Name As PartyName, " & _
            " SgP.Phone, SgP.Mobile, SgP.Add1, SgP.Add2, SgP.Add3, C.CityName " & _
            " From (((((Payment As P  " & _
            " Left Join Voucher_Type As Vt On P.V_Type = Vt.V_Type " & _
            " Left Join (" & mSubTable & ") As Sg On Sg.DocId = P.DocId " & _
            " Left Join SubGroup As SgP On P.PartyCode = SgP.SubCode " & _
            " Left Join City C On SgP.CityCode = C.CityCode ))))) " & Condstr
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
     
    RepName = "ChequePaymentRegister"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


