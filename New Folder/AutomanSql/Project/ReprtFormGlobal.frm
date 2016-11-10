VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form ReprtFormGlobal 
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
   Begin VB.CheckBox ChkOpeningStockOnly 
      BackColor       =   &H00C8E8DA&
      Caption         =   "Opening Stock Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6195
      TabIndex        =   20
      Top             =   1365
      Visible         =   0   'False
      Width           =   2670
   End
   Begin MSDataGridLib.DataGrid DgHelp 
      Height          =   2295
      Left            =   6810
      Negotiate       =   -1  'True
      TabIndex        =   19
      Top             =   2130
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "code"
         Caption         =   "Code"
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
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Name"
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
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2534.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton BtnSpeed 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Speed Print"
      DownPicture     =   "ReprtFormGlobal.frx":0000
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
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   6090
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
      DownPicture     =   "ReprtFormGlobal.frx":3132
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
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exit Form"
      Top             =   6090
      Width           =   1620
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Windows &Print"
      DownPicture     =   "ReprtFormGlobal.frx":6264
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
      Left            =   4635
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Print Report"
      Top             =   6090
      Width           =   1620
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
      TabIndex        =   17
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
         TabIndex        =   18
         Top             =   0
         Width           =   4470
      End
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7290
      TabIndex        =   15
      Top             =   -675
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   405
         TabIndex        =   16
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      AllowUserResizing=   1
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
Attribute VB_Name = "ReprtFormGlobal"
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
Dim RsMark As ADODB.Recordset
Dim myRst1 As ADODB.Recordset

Dim RepTitle As String, RepName As String, RsHelp As ADODB.Recordset
Dim RepPrint As Boolean
Dim RstRep As ADODB.Recordset
Dim RstRep1 As ADODB.Recordset
Dim SubRep1 As Boolean
Dim SpeedPrn As Boolean
'Modishekhar 17 mar
Dim FormulaStr1 As String, FormulaStr2 As String, FormulaStr3 As String, FormulaStr4 As String
Private Const GridRowHeight As Integer = 270




Private Const SprSOrdReg        As Byte = 2
Private Const SprSaleReg        As Byte = 3
Private Const SprSaleRet        As Byte = 4
Private Const SprPOrdReg        As Byte = 5
Private Const SprMatReg         As Byte = 6
Private Const SprPurReg         As Byte = 7
Private Const SprPurRet         As Byte = 8
Private Const SprStkTrf         As Byte = 9
Private Const SprStkReg         As Byte = 10
Private Const SprStkSumm        As Byte = 11
Private Const SprStkInHand      As Byte = 12
Private Const VehMoneyRect      As Byte = 13
Private Const WksSaleReg        As Byte = 14
Private Const SprIndent         As Byte = 15
Private Const SprDailySale      As Byte = 16
Private Const SprMonthSale      As Byte = 17
Private Const SprPartPur        As Byte = 18
Private Const SprPurSum         As Byte = 19
Private Const SprPartSale       As Byte = 20
Private Const SprSaleSum        As Byte = 21
Private Const SprStkReOrd       As Byte = 22
Private Const SprStkBin         As Byte = 23
Private Const SprPartMovement   As Byte = 24
Private Const SprStkAgeing      As Byte = 25
Private Const SprCtrRateVari    As Byte = 26
Private Const SprPurRateVari    As Byte = 27
Private Const SprDailySaleReg   As Byte = 28
Private Const SprMRPTaxClaimReg As Byte = 29
Private Const SprOthPurReg      As Byte = 30
Private Const SprSaleTaxCtrlStmt As Byte = 31
Private Const SprSaleTrfReg     As Byte = 32
Private Const WarTaxReimbReg    As Byte = 33
Private Const InputTaxReg       As Byte = 34
Private Const OutputTaxReg      As Byte = 35
Private Const DailyLubCon       As Byte = 36
Private Const SaleSumm          As Byte = 37
Private Const SalesmanWPending  As Byte = 38
Private Const CashBankBook      As Byte = 39
Private Const PurTaxSumm        As Byte = 40
Private Const SaleTaxSumm       As Byte = 41
Private Const BudgetExpVariRep  As Byte = 42
Private Const SalesManCostRep   As Byte = 43
Private Const BillWiseOutstanding   As Byte = 44
Private Const StockValue        As Byte = 45
Private Const SpareSaleAccount  As Byte = 46
Private Const SparePurchaseAccount  As Byte = 47





Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const List1 As Byte = 2
Private Const List2 As Byte = 3
Private Const List3 As Byte = 4
Private Const List4 As Byte = 5

Private Const Cat1 As Byte = 6
Private Const Cat2 As Byte = 7
Private Const Cat3 As Byte = 8
Private Const Cat4 As Byte = 9
Private Const Cat5 As Byte = 10

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
Private Const WksSlCsh$ = "W_SIC"          'Workshop Cash Sale
Private Const WksSlCre$ = "W_SIR"          'Workshop Credit Sale
Private Const SprSlRetCsh$ = "SXSRC"       'Cash Sale Return
Private Const SprSlRetCre$ = "SXSRR"       'Credit Sale Return
Private Const SprSlTrfRet$ = "SXSRT"       'Transfer Issue Return
Private Const SprPurCsh$ = "SXPIC"         'Cash Purchase
Private Const SprPurCre$ = "SXPIR"         'Credit Purchase
Private Const OthPurCsh$ = "SXPTC"         'Cash Purchase
Private Const OthPurCre$ = "SXPTR"         'Credit Purchase
Private Const SprPrRetCsh$ = "SYPRC"       'Purchase Return Cash
Private Const SprPrRetCre$ = "SYPRR"       'Purchase Return Credit
Private Const SprPrTrfRet$ = "SYPRT"       'Transfer Receipt Return
Private Const SprQuotation$ = "S_QU"       'Spare Quotation
Private Const WksEst$ = "W_EST"       'Workshop Estimation
Private Const WksPro$ = "W_PL"       'Workshop Proforma Labour
Private Const WksGenReq$ = "W_RG"       'Workshop General Reqisition
Private Const WksReqWrt$ = "W_RW"       'Workshop Warranti Reqisition
Dim mListItem As ListItem
Dim mDays As String
Dim CashOpening As Double, BankOpening As Double
Dim CashClosing As Double, BankClosing As Double
Dim TotalInterest As Double
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub BTNPRINT_Click()
On Error GoTo ERRORHANDLER
SubRep1 = False
RepPrint = True
Select Case GRepFormName
    Case SprCtrRateVari, SprPurRateVari
        SprRateVariation
        
    Case SprStkAgeing
        SprStkAge
        
    Case SprPartMovement
        SprPartMove
        
    Case SprStkBin
        SprStkBinLoc
    Case SprStkReOrd
        SprStkReOrder
    Case SprPurSum, SprSaleSum, SprMRPTaxClaimReg
        SprPurSalSum
    Case SprPartPur, SprPartSale
        SprPartPurSal
    Case SprMonthSale, SprDailySale
        SprMonthDateSale
    Case SprIndent
        SprIndentReg
    Case SprPOrdReg, SprSOrdReg
        SprSalePurOrd
    Case SprMatReg, SprStkTrf
        SprPurChl
    Case SprSaleTrfReg
        SprSaleTransfer
    Case SprSaleReg, SprSaleRet, SprPurReg, SprPurRet, WksSaleReg, WarTaxReimbReg
        SprSalePurReg
    Case SprStkReg, SprStkSumm, SprStkInHand
        SprStkRep
    Case StockValue
        ProcStockValue
    Case VehMoneyRect
        VehMoneyRectFunc
    Case SprDailySaleReg
        SprDailySaleRegFunc
    Case SprOthPurReg
        SprOthPurRegs
    Case SprSaleTaxCtrlStmt
        SprSaleTaxCtrlStmtRegs
    Case InputTaxReg
        InputTaxRegProc
    Case OutputTaxReg
        OutPutTaxRegProc
    Case DailyLubCon
        DailyLubConProc
    Case SaleSumm
        SaleSummary
    Case SalesmanWPending
        OutPayRepProc
    Case CashBankBook
        CashBankBookProc
    Case PurTaxSumm
        PurTaxSummProc
    Case SaleTaxSumm
        SaleTaxSummProc
    Case SpareSaleAccount
        ProcSpareSaleAccount
    Case SparePurchaseAccount
        ProcSparePurchaseAccount
    Case BudgetExpVariRep
        ProcBudgetExpVariRep
    Case SalesManCostRep
        ProcSalesManCostRep
    Case BillWiseOutstanding
        ProcBillWiseOutstanding
        
End Select
If RepPrint = False Then Exit Sub
If SpeedPrn = True Then SpeedPrn = False: Exit Sub
CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")
rpt.Database.SetDataSource RstRep
If GRepFormName = CashBankBook Then
    SubRep1 = True
    If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep
    If SubRep1 = True Then rpt.OpenSubreport("SUBREP2").Database.SetDataSource RstRep
Else
    If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
End If
rpt.ReadRecords
Set RstRep = Nothing
Call Formulas
Set myRst1 = Nothing
mRepName = RepName
Call Report_View(rpt, RepTitle, IIf(SpeedPrn, 1, 0), False)
SpeedPrn = False
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub BtnSpeed_Click()
SpeedPrn = True
BTNPRINT_Click
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
WinSetting Me
   Global_Grid
   TopCtrl1.TopText2 = "Add"
   'If InStr(UserPermission(Me.Name), "P") = 0 Then BTNPRINT.Enabled = False
   Select Case GRepFormName
        Case SprStkInHand
            BtnSpeed.Visible = True
        Case SprStkReg
            BtnSpeed.Visible = True
   End Select
   BtnSpeed.Visible = False
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
Dim Grid2Sql As String
If GRepFormName = SprSaleReg Then
    If UCase(Trim(FGrid.TextMatrix(List2, 1))) = "SALESMANWISE" Then
            Grid2Sql = "Select '' as O,Emp_Name as Name,Emp_Code as Code from Emp_Mast where Emp_Type=0"
            GridInitialise 2, Grid2Sql
    ElseIf UCase(Trim(FGrid.TextMatrix(List2, 1))) = "PARTYWISE" Then
             Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
             GridInitialise 2, Grid2Sql
'    Else
'           Call Ini_Grid
    End If
End If
If GRepFormName <> SprStkBin Then
    If TxtGrid.Item(Index) = "Yes" Then Check1(2).Enabled = False Else Check1(2).Enabled = True
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
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    'SprQuot,SprSaleOrd,SprSaleReg,SprSaleRet,SprPurOrd,
    'SprMatReg,SprPurReg,SprPurRet,SprStkTrf
    'SprStkReg,SprStkSumm,SprStkInHand,VehMoneyRect
    Select Case FGrid.Row
    Case List1
        Select Case GRepFormName
            'modi lps
            Case SprDailySale
                ListArray = Array("Detail", "Summary")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprMonthSale
                ListArray = Array("With Sale Ret.", "Without Sale Ret.")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            'EOF
            Case SprCtrRateVari, SprPurRateVari
                ListArray = Array("High", "Low", "Both")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprStkAgeing
                ListArray = Array("QtyWise", "ValueWise", "Both")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprPartMovement
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprStkBin
                ListArray = Array("Bin + PartNo", "PartNo")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprStkReOrd
                ListArray = Array("Above Maximum", "Below Minimum", "Below ReOrder")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprMatReg
                ListArray = Array("General", "PartyWise")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprSOrdReg
                ListArray = Array("All", "Pending")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprSaleReg, SprPurReg, WksSaleReg
                ListArray = Array("All", "Cash", "Credit")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprSaleRet, SprPurRet
                ListArray = Array("Cash", "Credit", "Transfer", "All")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
            Case SprPOrdReg
              ListArray = Array("Annual", "Quarterly", "Monthly", "General(Casual)", "VOR")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 5)
            Case VehMoneyRect
              ListArray = Array("All", "Form-60", "Form-61", "N/A")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprDailySaleReg
              ListArray = Array("Both", "Counter", "Workshop")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprStkSumm
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprStkInHand
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprStkReOrd
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprOthPurReg
                ListArray = Array("Cash", "Credit", "All")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
             Case SprStkReg
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprSaleTrfReg
                ListArray = Array("All", "Sale Challan", "Transfer", "Pending")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
            Case InputTaxReg, PurTaxSumm, SparePurchaseAccount
                ListArray = Array("Purchase", "Return")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case OutputTaxReg, SaleTaxSumm, SpareSaleAccount
                ListArray = Array("Sale", "Return")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                'O->Oil Filter, F->Fuel Filter,E->Engine Oil,G->Gear Oil,R->Rear Axle Oil,A->Front Axle Oil,S->Steering Oil
            Case SprPartSale, SprPartPur, SalesManCostRep
                ListArray = Array("Summary", "Detail")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case BillWiseOutstanding
                ListArray = Array("Debtors", "Creditors")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)

            'Case
            '    ListArray = Array("Oil Filter", "Fuel Filter")
            '    Set mListItem = ListView_Items(ListView, txtGrid, Index, ListArray, 2)
          End Select
    Case List2
        Select Case GRepFormName
            Case PurTaxSumm, SaleTaxSumm, SpareSaleAccount, SparePurchaseAccount
                  ListArray = Array("All", "Local", "Central")
                  Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprPOrdReg
                  ListArray = Array("All", "Pending", "Excess")
                  Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprDailySaleReg, SprMatReg
                  ListArray = Array("All", "Cash", "Credit")
                  Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprStkBin, SalesManCostRep
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprPartMovement
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprStkAgeing
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case VehMoneyRect
                ListArray = Array("All", "Taxable", "TaxPaid")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprSaleReg
                  ListArray = Array("General", "PartyWise", "SalesManWise")
                  Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprSaleRet, SprPurRet, WksSaleReg
                  ListArray = Array("General", "BillWise")
                  Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprPurReg
                  ListArray = Array("With Detail", "Without Detail")
                  Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprStkInHand
                ListArray = Array("MRP", "NDP")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprSaleTrfReg
                ListArray = Array("Summary", "Detail")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprPartSale
                ListArray = Array("Counter", "Workshop", "Both")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)

            End Select
    Case List3
     Select Case GRepFormName
        Case SprMatReg
              ListArray = Array("Pending", "Billed", "All")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
        Case VehMoneyRect
              ListArray = Array("Summery", "Detail")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
         Case SprSaleReg
                ListArray = Array("Both", "Counter", "Workshop")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
         Case SprStkBin
                ListArray = Array("Summery", "Detail")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
         Case SprPurReg
                ListArray = Array("Inv.Date", "Rec.Date")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
        Case SprStkInHand
            DgHelp.left = TxtGrid(0).left: DgHelp.top = TxtGrid(0).top + TxtGrid(0).height + 30
            If RsHelp.RecordCount = 0 Or (RsHelp.EOF = True Or RsHelp.BOF = True) Or FGrid.TextMatrix(FGrid.Row, 1) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, 2) <> RsHelp!Code Then
                RsHelp.MoveFirst
                RsHelp.FIND "Code ='" & FGrid.TextMatrix(FGrid.Row, 2) & "'"
            End If
    End Select
    Case List4
        Select Case GRepFormName
            Case VehMoneyRect
                  ListArray = Array("All", "M.M.", "BAL", "FULL", "Staff")
                  Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 5)
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
       
        Case List3
           Select Case GRepFormName
             Case SprStkInHand
                DGridTxtKeyDown DgHelp, TxtGrid, 0, RsHelp, KeyCode, True, 1
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols - 1
                    End If
                End If
            Case Else
                      ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
                  If KeyCode = vbKeyReturn Then
                      If TxtGridLeave = True Then TxtKeyDown
                  End If
           End Select
               
    Case List1, List2, List4
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
            Case SprStkAgeing
                NumPress TxtGrid(Index), KeyAscii, 3, 0
        End Select
    Case Cat2
        Select Case GRepFormName
            Case SprStkAgeing
                NumPress TxtGrid(Index), KeyAscii, 3, 0
        End Select
    Case Cat3
        Select Case GRepFormName
            Case SprStkAgeing
                NumPress TxtGrid(Index), KeyAscii, 3, 0
        End Select
    Case Cat4
        Select Case GRepFormName
            Case SprStkAgeing
                NumPress TxtGrid(Index), KeyAscii, 3, 0
        End Select
    Case Cat5
        Select Case GRepFormName
            Case SprStkAgeing
                NumPress TxtGrid(Index), KeyAscii, 3, 0
        End Select
    Case List3
         Select Case GRepFormName
            Case SprStkInHand
                    If DgHelp.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsHelp, KeyAscii, "Name"
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
            
        Case List1, List2, List4
            If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
            ListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
        Case List3
          Select Case GRepFormName
            
            Case "SprStkInHand"
                If KeyCode <> 13 Then TxtGrid_KeyDown Index, GridKey, 0
                 If DgHelp.Visible = True Then DGridTxtKeyUp1 TxtGrid, 0, RsHelp, KeyCode, "Name"
             Case Else
                 If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                 ListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
        End Select
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
        Case List1, List2
            If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
           
        Case List3
          Select Case GRepFormName
             Case SprStkInHand
                    If RsHelp.RecordCount = 0 Or (RsHelp.EOF = True Or RsHelp.BOF = True) Or TxtGrid(0).TEXT = "" Then
                        FGrid.TextMatrix(FGrid.Row, 1) = ""
                        FGrid.TextMatrix(FGrid.Row, 2) = ""
                    Else
                        FGrid.TextMatrix(FGrid.Row, 1) = RsHelp!Name
                        FGrid.TextMatrix(FGrid.Row, 2) = RsHelp!Code
                    End If
             Case Else
                If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
        End Select
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
BtnSpeed.left = (Pic.width - (BtnSpeed.width + BTNPRINT.width + BTNEXIT.width)) / 2: BtnSpeed.top = Pic.top + 10
BTNPRINT.left = BtnSpeed.left + BtnSpeed.width: BTNPRINT.top = Pic.top + 10
BTNEXIT.left = BTNPRINT.left + BTNPRINT.width: BTNEXIT.top = Pic.top + 10

FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
If GRepFormName = SprDailySaleReg Then
    FGrid.Rows = 4
Else
    FGrid.Rows = 11  '5
End If
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
    If DgHelp.Visible = True Then DgHelp.Visible = False
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
'If GRepFormName = SprPurReg Then
'   If FGrid.TextMatrix(List2, 1) = "Part Grade" Then
'        GridInitialise 4, "select '' as O,PartGrade_Name as Grade,PartGrade_code  as code from Part_Grade order by PartGrade_Name"
'   ElseIf FGrid.TextMatrix(List2, 1) = "TaxForms" Then
'      GridInitialise 4, "select '' as O,Form_Desc as Form,Form_Code  as code from TaxForms order by Form_Desc"
'   End If
'End If
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
Private Sub GridInitialise(Gridindex As Integer, GridSql As String, Optional UseFaCn As Boolean)
Dim Index As Integer
Index = Gridindex
If Index = 1 Then
    Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
    RsGrid1.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid1
    ReDim Preserve GridRow1(0)
    GridRow1(0) = 0
End If
If Index = 2 Then
    Set RsGrid2 = New ADODB.Recordset: RsGrid2.CursorLocation = adUseClient
    RsGrid2.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
    ReDim Preserve GridRow2(0)
    GridRow2(0) = 0
End If
If Index = 3 Then
    Set RsGrid3 = New ADODB.Recordset: RsGrid3.CursorLocation = adUseClient
    RsGrid3.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid3
    ReDim Preserve GridRow3(0)
    GridRow3(0) = 0
End If
If Index = 4 Then
    Set RsGrid4 = New ADODB.Recordset: RsGrid4.CursorLocation = adUseClient
    If UseFaCn Then
        RsGrid4.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid4
    Else
        RsGrid4.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid4
    End If
    ReDim Preserve GridRow4(0)
    GridRow4(0) = 0
End If
GridSel(Index).height = 1600
GridSel(Index).Visible = True: GridSel(Index).Enabled = False: Check1(Index).Visible = True
GridSel(Index).width = 5200: GridSel(Index).ColWidth(0) = 600: GridSel(Index).ColWidth(2) = 0: GridSel(Index).ColWidth(1) = 4000
Check1(Index).width = 580: Check1(Index).height = GridSel(Index).RowHeight(0) + 20: Check1(Index).Value = Checked
End Sub

Private Sub Ini_Grid()
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
 Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where site_code='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
Select Case GRepFormName
    Case SprSaleTaxCtrlStmt
       With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
                       
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
             
       End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName, site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName, Div_code  as code from Division  order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case SprOthPurReg
       With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Cash/Credit/All": .RowHeight(List1) = GridRowHeight
                  
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 3
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select distinct '' as O,Name as PartyName, SubCode as code from SubGroup order by Name"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case BudgetExpVariRep
       With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Branch Wise": .RowHeight(List1) = GridRowHeight
                  
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Yes"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O, site_desc as SiteName, site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select distinct '' as O,Name as Expence_Account_Name, SubCode as code from SubGroup Where Nature In ('Expenses') Order by Name"
        GridInitialise 2, Grid2Sql
        
        
    Case SalesManCostRep
       With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Only SalesMan": .RowHeight(List2) = GridRowHeight
                  
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Summary"
            .TextMatrix(List2, 1) = "No"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O, site_desc as SiteName, site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select distinct '' as O,Emp_Name as Employee_Name, Emp_Code as code From Emp_Mast Order by Emp_Name"
        GridInitialise 2, Grid2Sql
        
        
    Case BillWiseOutstanding
       With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Type": .RowHeight(List1) = GridRowHeight
                  
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Debtors"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O, site_desc as SiteName, site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select distinct '' as O,Name as Ledger_Ac_Name, SubCode as code From SubGroup Order by Name"
        GridInitialise 2, Grid2Sql
        
        
    Case SprStkReOrd
        With FGrid
            .TextMatrix(List1, 0) = "Select Option": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List1, 1) = "Below ReOrder"
            
        End With
        mFirstRow = List1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 3
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case SprStkAgeing
        With FGrid
            .TextMatrix(Date1, 0) = "As On Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(List1, 0) = "Select Option": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Marked Part": .RowHeight(List2) = GridRowHeight
            .TextMatrix(Cat1, 0) = "No Of Days[1]": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Cat2, 0) = "No Of Days[2]": .RowHeight(Cat2) = GridRowHeight
            .TextMatrix(Cat3, 0) = "No Of Days[3]": .RowHeight(Cat3) = GridRowHeight
            .TextMatrix(Cat4, 0) = "No Of Days[4]": .RowHeight(Cat4) = GridRowHeight
            .TextMatrix(Cat5, 0) = "No Of Days[5]": .RowHeight(Cat5) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "QtyWise"
            .TextMatrix(List2, 1) = "No"
            .TextMatrix(Cat1, 1) = 10
            .TextMatrix(Cat2, 1) = 20
            .TextMatrix(Cat3, 1) = 30
            .TextMatrix(Cat4, 1) = 40
            .TextMatrix(Cat5, 1) = 50
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Cat5: mHelpGridNo = 3
    
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        GridSel(2).ColWidth(1) = 1500: GridSel(2).ColWidth(3) = 2500

    Case SprCtrRateVari, SprPurRateVari
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
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Cat1: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
        
    Case SprPartMovement
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Move Parts": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Marked Part": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Yes"
            .TextMatrix(List2, 1) = "No"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 3
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 2, Grid2Sql
         
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql
        GridSel(2).ColWidth(1) = 1500: GridSel(2).ColWidth(3) = 3000
        
    Case SprStkBin
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Index On": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Marked Part": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Summery/Detail": .RowHeight(List3) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Bin + PartNo"
            .TextMatrix(List2, 1) = "No"
            .TextMatrix(List3, 1) = "Detail"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List3: mHelpGridNo = 3
                
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & "  order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select distinct '' as O, P.Bin_Loca as BinLoca, P.Bin_Loca from SP_Stock SPStk Left Join Part P on SpStk.Part_No=P.Part_NO  ORDER BY P.BIN_LOCA"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case SprMRPTaxClaimReg, SprDailySale, SprSaleSum, SprPurSum, SprStkTrf

        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1:: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
        
    Case SprMonthSale

        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Report Type": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "With Sale Ret."
        End With
        mFirstRow = Date1:: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
        
    Case SprPartPur
    
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "UpTo Date": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Summary"
        End With
        mFirstRow = Date1:: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select distinct '' as O,Part_No as Part_No,Part_No as code from SP_Stock Order by Part_No"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        
        
        
    Case SprPartSale
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Rep Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Counter/WorkShop": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Summary"
            .TextMatrix(List2, 1) = "Both"
        End With
        mFirstRow = Date1:: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select distinct '' as O,Part_No as Part_No,Part_No as code from SP_Stock Order by Part_No"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select '' as O,PartGrade_Name as Grade,PartGrade_code  as code from Part_Grade  order by PartGrade_Name"
        GridInitialise 4, Grid4Sql

    Case SprIndent
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 4
    
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
        GridInitialise 4, Grid4Sql
        
        GridSel(3).ColWidth(1) = 1500: GridSel(3).ColWidth(3) = 2500
      
    Case SprMatReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "General/PartyWise": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Type": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Pending/All": .RowHeight(List3) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "General"
            .TextMatrix(List2, 1) = "All"
            .TextMatrix(List3, 1) = "All"
            
        End With
        mFirstRow = Date1: mLastRow = List3: mHelpGridNo = 2
        Grid1Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 1, Grid1Sql

  Grid2Sql = "select '' as O,site_desc as SiteName, site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 2, Grid2Sql
        
    Case SprSaleTrfReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Challan Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Report Type": .RowHeight(List2) = GridRowHeight

            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List2, 1) = "Summary"

        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 3
            
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        
    
    Case SprSaleRet, SprPurRet, WksSaleReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Cash/Credit": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "General/Billwise": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Cash"
            .TextMatrix(List2, 1) = "General"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 4
            
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
         
    Case SprPurReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Cash/Credit": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Report Type": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Based On": .RowHeight(List3) = GridRowHeight
                        
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Cash"
            .TextMatrix(List2, 1) = "With Detail"
            .TextMatrix(List3, 1) = "Rec.Date"
            
        End With
        mFirstRow = Date1: mLastRow = List3: mHelpGridNo = 4
            
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql

        If StrCmp(left(PubComp_Name, 4), "Yash") Then
            Grid3Sql = "Select '' as O, Form_Desc as [Form Description], Form_Code as Code from TaxForms WHERE Trn_Type ='Purchase' AND Vehicle_Yn = 0 Order by Form_Desc"
            GridInitialise 3, Grid3Sql
        Else
            Grid3Sql = "select '' as O, Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
            GridInitialise 3, Grid3Sql
        End If
        
        Grid4Sql = "select '' as O,PartGrade_Name as Grade,PartGrade_code  as code from Part_Grade  order by PartGrade_Name"
        GridInitialise 4, Grid4Sql
    
    Case SprSaleReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Cash/Credit": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Report Type": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Counter/Workshop": .RowHeight(List3) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Cash"
            .TextMatrix(List2, 1) = "General"
            .TextMatrix(List3, 1) = "Both"
        End With
        mFirstRow = Date1: mLastRow = List3: mHelpGridNo = 3
            
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case WarTaxReimbReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
            
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case PurTaxSumm
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Purchase/Return": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Local/Central": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Purchase"
            .TextMatrix(List2, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 1
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        
        
    Case SaleTaxSumm
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Sale/Return": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Local/Central": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Sale"
            .TextMatrix(List2, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from  site " & sitecond & "  order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
        
    Case SpareSaleAccount
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Sale/Return": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Local/Central": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Sale"
            .TextMatrix(List2, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from  site " & sitecond & "  order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql

    Case SparePurchaseAccount
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Sale/Return": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Local/Central": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Purchase"
            .TextMatrix(List2, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from  site " & sitecond & "  order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql

        
    Case InputTaxReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Purchase/Return": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Purchase"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 1
        
         Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        
    Case SalesmanWPending
         With FGrid
            .TextMatrix(Date1, 0) = "As on Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Billing Date ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
       End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,emp_name as name,Emp_code as code from emp_mast where emp_type = 0  order by Emp_name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 3, Grid3Sql
        
    Case OutputTaxReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Sale/Return": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Sale"
        End With
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 2
        
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
        
        
    Case SprPOrdReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Order Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Report Option": .RowHeight(List2) = GridRowHeight
                        
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Monthly"
            .TextMatrix(List2, 1) = "All"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,cityname as CityName,CityCode as code from city  order by cityname"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 4, Grid4Sql

    Case SprStkReg, SprStkSumm
        ChkOpeningStockOnly.Visible = True
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Marked Part": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "No"
        End With
            
        mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select '' as O,PartGrade_Name as PartGrade,PartGrade_code  as code from Part_Grade "
        GridInitialise 4, Grid4Sql
        
        'For Part Only
        GridSel(2).ColWidth(1) = 1200: GridSel(2).ColWidth(3) = 2800
    Case SprStkInHand
        ChkOpeningStockOnly.Visible = True
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Marked Part": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Rate Type": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Site": .RowHeight(List3) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "No"
            .TextMatrix(List2, 1) = "NDP"
            '.TextMatrix(List3, 1) = PubSiteName: .TextMatrix(List3, 2) = PubSiteCode
        End With
            
        mFirstRow = Date1: mLastRow = List3: mHelpGridNo = 4
        
  
        Set RsHelp = GCn.Execute("select '' as O,site_desc as Name,site_code  as code from site " & sitecond & " order by site_desc")
        Set DgHelp.DataSource = RsHelp
        
        Grid1Sql = "select distinct '' as O,Bin_Loca as Bin,Bin_Loca  as code from Part order by Bin_Loca"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 2, Grid2Sql
        
        If StrCmp(left(PubComp_Name, 4), "Yash") Then
            Grid3Sql = "select '' as O, God_Name, God_Code  as code FROM Godown WHERE Appli_For ='0'  order BY God_Name "
            GridInitialise 3, Grid3Sql
        Else
            Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
            GridInitialise 3, Grid3Sql
        End If
        
        Grid4Sql = "select '' as O,PartGrade_Name as PartGrade,PartGrade_code  as code from Part_Grade "
        GridInitialise 4, Grid4Sql
        
        'For Part Only
        GridSel(2).ColWidth(1) = 1200: GridSel(2).ColWidth(3) = 2800
        
        
    Case StockValue
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
            
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 4
        
  
        Set RsHelp = GCn.Execute("select '' as O,site_desc as Name,site_code  as code from site " & sitecond & " order by site_desc")
        Set DgHelp.DataSource = RsHelp
        
        Grid1Sql = "select distinct '' as O,Bin_Loca as Bin,Bin_Loca  as code from Part order by Bin_Loca"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select '' as O,PartGrade_Name as PartGrade,PartGrade_code  as code from Part_Grade "
        GridInitialise 4, Grid4Sql
        
        'For Part Only
        GridSel(2).ColWidth(1) = 1200: GridSel(2).ColWidth(3) = 2800
        
    Case SprSOrdReg
        With FGrid
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Pending/All": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
        mFirstRow = Date2: FGrid.Row = mFirstRow: mHelpGridNo = 3
        mLastRow = List1
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql

        
    Case VehMoneyRect
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Under Declaration": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Taxable/TaxPaid": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Summery/Detail": .RowHeight(List3) = GridRowHeight
            .TextMatrix(List4, 0) = "Vr.Cat.Type": .RowHeight(List4) = GridRowHeight

            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List2, 1) = "All"
            .TextMatrix(List3, 1) = "Summery"
            .TextMatrix(List4, 1) = "All"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List4: mHelpGridNo = 4
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,CityName as Location ,CityCode as code from City order by CityName"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select '' as O,Description as VoucherType,v_type as code from Voucher_Type where category='GENFA' order by v_type "
        GridInitialise 4, Grid4Sql, True
    Case SprDailySaleReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Workshop/Counter": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Sale Type": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Both"
            .TextMatrix(List2, 1) = "All"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
                
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case DailyLubCon
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            '.TextMatrix(List1, 0) = "Type": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            '.TextMatrix(List1, 1) = "N/A"
            
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case CashBankBook
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division order by Div_Name"
        GridInitialise 2, Grid2Sql
        
    Case SaleSumm
        With FGrid
            .TextMatrix(Date1, 0) = "As on Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date1: mHelpGridNo = 2
            
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql

        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division  order by Div_Name"
        GridInitialise 2, Grid2Sql
   
        
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
Private Sub SprSalePurOrd()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If GRepFormName = SprPOrdReg Then If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    'If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    If GRepFormName = SprPOrdReg Then If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If GRepFormName = SprPOrdReg And (FGrid.TextMatrix(List2, 1) = "All" Or FGrid.TextMatrix(List2, 1) = "Pending") Then
        Condstr = " where left(P.OrderId,1)='" & PubDivCode & "' And  P.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and P.OrdClosDate Is Null "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("P.OrderId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("P.OrderId", "3", "1") & " ='" & PubSiteCode & "' "
         End If


        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and P.Party_Code in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and SubGroup.CityCode in (" & GridString3 & ")"
        'If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Left(P.OrderId,1) in (" & GridString4 & ")"
        Condstr = Condstr & " and P.Order_type <> 'S_SO' and Right(P.Order_type,1) = '" & left(FGrid.TextMatrix(List1, 1), 1) & "'"
        If FGrid.TextMatrix(List2, 1) = "Pending" Then
            Condstr = Condstr & " and P1.QTY-P1.Sup_Qty  > 0"
        End If
        mQry = "SELECT C.CityName,P.OrderId,site.site_Desc,P.Party_Code, P.Order_No, P1.Amount, P.Site_Code, P.V_Date, P.Order_Prefix, Part.Part_Name, SubGroup.Name, P1.PART_NO, P.Order_Reg_No, P.Order_Reg_Dt, P1.QTY, P1.Sup_Qty, (P1.QTY-P1.Sup_Qty) AS BalQty " & _
            "FROM ((((SP_Order AS P LEFT JOIN SP_Order1 AS P1 ON P.OrderId = P1.OrderId) " & _
            "LEFT JOIN Part ON P1.PART_NO = Part.PART_NO and Part.Div_Code = left(P1.orderid,1)) " & _
            "LEFT JOIN SubGroup ON P.Party_Code = SubGroup.SubCode) " & _
            "LEFT JOIN Site ON left(P.Site_Code,1) = Site.Site_Code) " & _
            "LEFT JOIN City as C ON SubGroup.CityCode = C.CityCode"
        mQry = mQry + Condstr + " order by site.site_Desc,P.v_date, P.Order_Prefix,P.Order_No"
        RepName = "SprPOrdReg"
    ElseIf GRepFormName = SprPOrdReg And FGrid.TextMatrix(List2, 1) = "Excess" Then
        Condstr = " where left(P.OrderId,1)='" & PubDivCode & "' And P.v_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " And P.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and P.OrdClosDate Is Null "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("P.OrderId", "3", "1") & " in (" & GridString1 & ")"
       
             If Check1(1).Value = Checked Then
            If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("P.OrderId", "3", "1") & " ='" & PubSiteCode & "' "
            End If

        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and P.Party_Code in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and C.CityCode in (" & GridString3 & ")"
        Condstr = Condstr & " and P.Order_type <> 'S_SO' and Right(P.Order_type,1) = '" & left(FGrid.TextMatrix(List1, 1), 1) & "'"
        If FGrid.TextMatrix(List2, 1) = "Pending" Then
            Condstr = Condstr & " and P1.QTY-P1.Sup_Qty < 0"
        End If
        mQry = "SELECT C.CityName,P.OrderId,site.site_Desc,P.Party_Code, P.Order_No, P1.Amount, P.Site_Code, P.V_Date, P.Order_Prefix, Part.Part_Name, SubGroup.Name, P1.PART_NO, P.Order_Reg_No, P.Order_Reg_Dt, P1.QTY, P1.Sup_Qty, (P1.QTY-P1.Sup_Qty) AS BalQty " & _
            "FROM ((((SP_Order AS P LEFT JOIN SP_Order1 AS P1 ON P.OrderId = P1.OrderId) " & _
            "LEFT JOIN Part ON P1.PART_NO = Part.PART_NO and Part.Div_Code = left(p1.orderid,1)) " & _
            "LEFT JOIN SubGroup ON P.Party_Code = SubGroup.SubCode) " & _
            "LEFT JOIN Site ON left(P.Site_Code,1) = Site.Site_Code) " & _
            "LEFT JOIN City as C ON SubGroup.CityCode = C.CityCode"
        mQry = mQry + Condstr + " order by site.site_Desc,P.v_date, P.Order_Prefix,P.Order_No"
        RepName = "SprPOrdReg"
    ElseIf GRepFormName = SprSOrdReg Then
        Condstr = " where left(S.OrderId,1)='" & PubDivCode & "' And  S.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and S.OrdClosDate Is Null "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("s.OrderId", "3", "1") & " in (" & GridString1 & ")"
             If Check1(1).Value = Checked Then
             If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("s.OrderId", "3", "1") & " ='" & PubSiteCode & "' "
            End If

        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and S.Party_Code in (" & GridString2 & ")"
        Condstr = Condstr & " and S.Order_type = 'S_SO'"
            If FGrid.TextMatrix(List1, 1) = "Pending" Then
                Condstr = Condstr & " and S1.QTY-S1.Sup_Qty <> 0"
            End If
        mQry = "SELECT C.CityName,S.OrderId,site.site_Desc,S.Party_Code, S.Order_No, S1.Amount, S.Site_Code, S.V_Date, S.Order_Prefix, Part.Part_Name, SubGroup.Name, S1.PART_NO, S.Order_Reg_No, S.Order_Reg_Dt, S1.QTY, S1.Sup_Qty, (S1.QTY-S1.Sup_Qty) AS BalQty " & _
            "FROM ((((SP_Order AS S LEFT JOIN SP_Order1 AS S1 ON S.OrderId = S1.OrderId) " & _
            "LEFT JOIN Part ON S1.PART_NO = Part.PART_NO and Part.Div_Code = left(s1.orderid,1)) " & _
            "LEFT JOIN SubGroup ON S.Party_Code = SubGroup.SubCode) " & _
            "LEFT JOIN Site ON left(S.Site_Code,1) = Site.Site_Code) " & _
            "Left Join City as C on SubGroup.CityCode=C.CityCode"
        mQry = mQry + Condstr + " order by site.site_Desc,S.v_date, S.Order_Prefix,S.Order_No"
        RepName = "SprSOrdReg"
    End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepTitle = UCase(Me.CAPTION) + "[" + FGrid.TextMatrix(List1, 1) + "]"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprIndentReg()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    Condstr = " where Indent.Doc_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Indent.Doc_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Indent.site_code,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(indent.site_code,1) ='" & PubSiteCode & "' "
        End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Indent.PartyCode in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Indent.PART_NO in (" & GridString3 & ")"
        
    mQry = "SELECT SubGroup.Name, Part.Part_Name, Indent.DocID, Indent.IDNo, Indent.Doc_Date, Indent.PART_NO, Indent.QTY, Indent.RATE, Indent.Remark " & _
    "FROM (Indent LEFT JOIN Part ON Indent.PART_NO = Part.PART_NO and Part.Div_Code = left(indent.Docid,1)) LEFT JOIN SubGroup ON Indent.PartyCode = SubGroup.SubCode"
    mQry = mQry + Condstr + " order by Indent.IDNo"

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprIndent"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SprMonthDateSale()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, mQRY1 As String
    'SubRep1 = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    Condstr = " Where SP_Sale.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Sale.DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("sp_sale.docid", "3", "1") & " ='" & PubSiteCode & "' "
        End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Sale.DocId,1) in (" & GridString2 & ")"
    
    If GRepFormName = SprMonthSale Then
        If Not FGrid.TextMatrix(List1, 1) = "With Sale Ret." Then
'             mQRY = "SELECT " & _
'                    "month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt,0 as TaxableAmtRet,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt,0 as TaxpaidAmtRet,sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP ) as Tax,0 as TaxRet,'S' as RepType,0 as warrTB,0 as WarrTP,sp_sale.v_type " & _
'                "FROM SP_Sale " & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR') " & _
'                    "group by month(v_date),sp_sale.v_type " & _
'                " Union All " & _
'                "SELECT " & _
'                   "month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt,0 as TaxableAmtRet,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP -  SP_Sale.D_Amt_TP) as TaxpaidAmt,0 as TaxpaidAmtRet, sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP) as Tax,0 as TaxRet,'W' as RepType,0 as warrTB,0 as WarrTP,sp_sale.v_type " & _
'                "FROM SP_Sale " & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by month(sp_sale.v_date),sp_sale.v_type" & _
'                " Union All " & _
'                "SELECT " & _
'                   "month(SP_Sale.V_Date) as SaleMonth, 0 as TaxableAmt,0 as TaxableAmtRet,0 as TaxpaidAmt,0 as TaxpaidAmtRet, 0 as Tax,0 as TaxRet,'W' as RepType " & _
'                   ",iif(SP_Stock.Purpose='W' and SP_Stock.tax_yn=1,sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate),0) as warrTB,iif(SP_Stock.Purpose='W' and SP_Stock.tax_yn=0,sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate),0) as warrTP,sp_sale.v_type " & _
'                "FROM SP_Sale Left Join Sp_Stock on SP_Sale.Docid=SP_Stock.Invoice_DocId " & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by month(sp_sale.v_date),SP_Stock.Purpose, SP_Stock.tax_yn,sp_stock.Amount,sp_stock.QTY_Iss,sp_stock.QTY_Ret,sp_stock.v_Rate,sp_sale.v_type"
                
             mQry = "SELECT " & _
                    "month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB -SP_Sale.TaxSur_AmtMRP) as TaxableAmt,0 as TaxableAmtRet,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt,0 as TaxpaidAmtRet,sum(SP_Sale.Tax_Amt +Sp_Sale.SatAmt+ SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt +SP_Sale.TaxSur_AmtMRP ) as Tax,0 as TaxRet,'S' as RepType,0 as warrTB,0 as WarrTP,sp_sale.v_type " & _
                "FROM SP_Sale " & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR') " & _
                    "group by month(v_date),sp_sale.v_type " & _
                " Union All " & _
                "SELECT " & _
                   "month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt,0 as TaxableAmtRet,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP -  SP_Sale.D_Amt_TP) as TaxpaidAmt,0 as TaxpaidAmtRet, sum(SP_Sale.Tax_Amt +Sp_Sale.SatAmt+ SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP) as Tax,0 as TaxRet,'W' as RepType,0 as warrTB,0 as WarrTP,sp_sale.v_type " & _
                "FROM SP_Sale " & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by month(sp_sale.v_date),sp_sale.v_type" & _
                " Union All " & _
                "SELECT " & _
                   "month(SP_Sale.V_Date) as SaleMonth, 0 as TaxableAmt,0 as TaxableAmtRet,0 as TaxpaidAmt,0 as TaxpaidAmtRet, 0 as Tax,0 as TaxRet,'W' as RepType " & _
                   "," & cIIF("SP_Stock.Purpose='W' and SP_Stock.tax_yn=1", "sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate)", "0") & " as warrTB, " & cIIF("SP_Stock.Purpose='W' and SP_Stock.tax_yn=0", "sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate)", "0") & " as warrTP,sp_sale.v_type " & _
                "FROM SP_Sale Left Join Sp_Stock on SP_Sale.Docid=SP_Stock.Invoice_DocId " & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by month(sp_sale.v_date),SP_Stock.Purpose, SP_Stock.tax_yn,sp_stock.Amount,sp_stock.QTY_Iss,sp_stock.QTY_Ret,sp_stock.v_Rate,sp_sale.v_type"
                
                
        mQRY1 = "SELECT " & _
                    "" & cIIF("SubGroupType.Description<>''", "SubGroupType.Description", "'Others'") & " as Descrip, Sum(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB+SP_Sale.SprAmt_TB+SP_Sale.OilAmt_TB-SP_Sale.D_Amt_TB) AS TaxableAmt, Sum(SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TP-SP_Sale.D_Amt_TP) AS TaxpaidAmt, Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Tax_Amt+SP_Sale.Tax_Sur_Amt+sp_sale.Rounded) AS Tax " & _
                "FROM SP_Sale " & _
                    "LEFT JOIN (SubGroup LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type) ON SP_Sale.Party_Code = SubGroup.SubCode " & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR') group by SubGroupType.Description"
        Else
         mQry = "SELECT " & _
                    "month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt,0 as TaxableAmtRet,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt,0 as TaxpaidAmtRet,sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP +Sp_Sale.SatAmt) as Tax,0 as TaxRet,'S' as RepType,0 as warrTB,0 as WarrTP,sp_sale.v_type " & _
                "FROM SP_Sale " & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR') " & _
                    "group by month(v_date),sp_sale.v_type " & _
                " Union All " & _
                "SELECT " & _
                    "month(SP_Sale.V_Date) as SaleMonth,0 as TaxableAmt,sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmtRet,0 as TaxpaidAmt,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmtRet,0 as Tax,sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP +Sp_Sale.SatAmt ) as TaxRet,'S' as RepType,0 as warrTB,0 as WarrTP,sp_sale.v_type " & _
                "FROM SP_Sale " & Condstr & " and sp_sale.v_type in ('SXSRC','SXSRR') " & _
                    "group by month(v_date),sp_sale.v_type " & _
                " Union All " & _
                "SELECT " & _
                   "month(SP_Sale.V_Date) as SaleMonth, sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt,0 as TaxableAmtRet,sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP -  SP_Sale.D_Amt_TP) as TaxpaidAmt,0 as TaxpaidAmtRet, sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP +Sp_Sale.SatAmt) as Tax,0 as TaxRet,'W' as RepType,0 as warrTB,0 as WarrTP,sp_sale.v_type " & _
                "FROM SP_Sale " & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by month(sp_sale.v_date),sp_sale.v_type" & _
                " Union All " & _
                "SELECT " & _
                   "month(SP_Sale.V_Date) as SaleMonth, 0 as TaxableAmt,0 as TaxableAmtRet,0 as TaxpaidAmt,0 as TaxpaidAmtRet, 0 as Tax,0 as TaxRet,'W' as RepType " & _
                   "," & cIIF("SP_Stock.Purpose='W' and SP_Stock.tax_yn=1", "sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate)", "0") & " as warrTB, " & cIIF("SP_Stock.Purpose='W' and SP_Stock.tax_yn=0", "sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate)", "0") & " as warrTP,sp_sale.v_type " & _
                "FROM SP_Sale Left Join Sp_Stock on SP_Sale.Docid=SP_Stock.Invoice_DocId " & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by month(sp_sale.v_date),SP_Stock.Purpose, SP_Stock.tax_yn,sp_stock.Amount,sp_stock.QTY_Iss,sp_stock.QTY_Ret,sp_stock.v_Rate,sp_sale.v_type"
                
        mQRY1 = "SELECT " & _
                    "" & cIIF("SubGroupType.Description<>''", "SubGroupType.Description", "'Others'") & " as Descrip, Sum(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB+SP_Sale.SprAmt_TB+SP_Sale.OilAmt_TB-SP_Sale.D_Amt_TB) AS TaxableAmt, Sum(SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TP-SP_Sale.D_Amt_TP) AS TaxpaidAmt, Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Tax_Amt+SP_Sale.Tax_Sur_Amt+sp_sale.Rounded +Sp_Sale.SatAmt) AS Tax " & _
                "FROM SP_Sale " & _
                    "LEFT JOIN (SubGroup LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type) ON SP_Sale.Party_Code = SubGroup.SubCode " & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR') group by SubGroupType.Description"
        End If
        RepName = "SprMonthSale"

    ElseIf GRepFormName = SprDailySale Then
'        mQRY = "SELECT 'S' as RepType,SP_Sale.V_Date, iif(sp_sale.v_type='SYSIC','Cash  ','Credit') as CashCr, " & _
'            " 0 as TaxableAmt,0 as TaxpaidAmt,0 as GenSurTrns,0 as Tax, 0 as TaxOnMRP, 0 as Pack,0 as TotalAmt," & _
'            "sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_MRP_TB - SP_Sale.D_Amt_TB) as TaxableAmt2," & _
'            "sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_MRP_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt2," & _
'            "Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) as GenSurTrns2, sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt ) as Tax2," & _
'            "sum(SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_AmtMRP) as TaxOnMRP2," & _
'            "Sum(SP_Sale.Packing) as Pack2,sum(Sp_Sale.Total_Amt) as TotalAmt2 " & _
'            "FROM SP_Sale" & CondStr & " and sp_sale.v_type in ('SYSIC','SYSIR')  group by SP_Sale.V_Date,sp_sale.v_type " & _
'            "Union All " & _
'            "SELECT 'W' as RepType,SP_Sale.V_Date,  iif(sp_sale.v_type='W_SIC','Cash  ','Credit') as CashCr, " & _
'            "sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_MRP_TB - SP_Sale.D_Amt_TB) as TaxableAmt," & _
'            "sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_MRP_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt, " & _
'            "Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) as GenSurTrns,sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt ) as Tax," & _
'            "sum(SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_AmtMRP) as TaxOnMRP," & _
'            "Sum(SP_Sale.Packing) as Pack,sum(Sp_Sale.Total_Amt) as TotalAmt, " & _
'            "0 as TaxableAmt2,0 as TaxpaidAmt2,0 as GenSurTrns2,0 as Tax2, 0 as TaxOnMRP2, 0 as Pack2,0 as TotalAmt2 " & _
'            "FROM SP_Sale" & CondStr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by SP_Sale.V_Date,SP_Sale.V_Type"
        
   '************VIKAS 10 OCT 03
'made comm by arpit 4/9/06 at JMK
''        mQRY = "SELECT 'S' as RepType,SP_Sale.V_Date, iif(sp_sale.v_type='SYSIC','Cash  ','Credit') as CashCr, " & _
''            " 0 as TaxableAmt,0 as TaxpaidAmt,0 as GenSurTrns,0 as Tax, 0 as TaxOnMRP, 0 as Pack,0 as TotalAmt," & _
''            "sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt2," & _
''            "sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt2," & _
''            "Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) as GenSurTrns2, sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP) as Tax2," & _
''            "0 as TaxOnMRP2," & _
''            "Sum(SP_Sale.Packing) as Pack2,sum(Sp_Sale.Total_Amt) as TotalAmt2,0 as warrTB,0 as WarrTP " & _
''            "FROM SP_Sale" & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR')  group by SP_Sale.V_Date,sp_sale.v_type " & _
''            "Union All " & _
''            "SELECT 'W' as RepType,SP_Sale.V_Date,  iif(sp_sale.v_type='W_SIC','Cash  ','Credit') as CashCr, " & _
''            "sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt," & _
''            "sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt, " & _
''            "Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) as GenSurTrns,sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP) as Tax," & _
''            "0  as TaxOnMRP,Sum(SP_Sale.Packing) as Pack,sum(Sp_Sale.Total_Amt) as TotalAmt,0 as TaxableAmt2,0 as TaxpaidAmt2,0 as GenSurTrns2,0 as Tax2, 0 as TaxOnMRP2, 0 as Pack2,0 as TotalAmt2,0 as warrTB,0 as WarrTP " & _
''            "FROM SP_Sale" & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by SP_Sale.V_Date,SP_Sale.V_Type" & _
''            " Union All " & _
''            "SELECT 'W' as RepType,SP_Sale.V_Date,  iif(sp_sale.v_type='W_SIC','Cash  ','Credit') as CashCr, " & _
''            "0 as TaxableAmt," & _
''            "0 as TaxpaidAmt,0 as GenSurTrns, " & _
''            "0 as Tax," & _
''            "0  as TaxOnMRP,0 as pack," & _
''            "0 as TotalAmt, " & _
''            "0 as TaxableAmt2,0 as TaxpaidAmt2,0 as GenSurTrns2,0 as Tax2, 0 as TaxOnMRP2, 0 as Pack2,0 as TotalAmt2,iif(SP_Stock.Purpose='W' and SP_Stock.tax_yn=1,sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate),0) as warrTB,iif(SP_Stock.Purpose='W' and SP_Stock.tax_yn=0,sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.v_Rate),0) as warrTP  " & _
''            "FROM SP_Sale Left Join Sp_Stock on SP_Sale.Docid=SP_Stock.Invoice_DocId" & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by SP_Sale.V_Date,SP_Sale.V_Type,SP_Stock.Purpose, SP_Stock.tax_yn,sp_stock.Amount"


        mQry = "SELECT 'S' as RepType,SP_Sale.V_Date, " & cIIF("sp_sale.v_type='SYSIC'", "'Cash  '", "'Credit'") & " as CashCr, " & _
            " 0 as TaxableAmt,0 as TaxpaidAmt,0 as GenSurTrns,0 as Tax, 0 as TaxOnMRP, 0 as Pack,0 as TotalAmt," & _
            "sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt2," & _
            "sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt2," & _
            "Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) as GenSurTrns2, sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP) as Tax2," & _
            "0 as TaxOnMRP2," & _
            "Sum(SP_Sale.Packing) as Pack2,sum(Sp_Sale.Total_Amt) as TotalAmt2,0 as warrTB,0 as WarrTP, 0 As SatAmt, Sum(Sp_Sale.SatAmt) as SatAmt2 " & _
            "FROM SP_Sale" & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR')  group by SP_Sale.V_Date,sp_sale.v_type " & _
            "Union All " & _
            "SELECT 'W' as RepType,SP_Sale.V_Date,  " & cIIF("sp_sale.v_type='W_SIC'", "'Cash  '", "'Credit'") & " as CashCr, " & _
            "sum(SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_TB-SP_Sale.Tax_AmtMRP -SP_Sale.TaxSur_AmtMRP) as TaxableAmt," & _
            "sum(SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_TP) as TaxpaidAmt, " & _
            "Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) as GenSurTrns,sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP) as Tax," & _
            "0  as TaxOnMRP,Sum(SP_Sale.Packing) as Pack,sum(Sp_Sale.Total_Amt) as TotalAmt,0 as TaxableAmt2,0 as TaxpaidAmt2,0 as GenSurTrns2,0 as Tax2, 0 as TaxOnMRP2, 0 as Pack2,0 as TotalAmt2,0 as warrTB,0 as WarrTP, Sum(Sp_Sale.SatAmt) as SatAmt, 0 as SatAmt2 " & _
            "FROM SP_Sale" & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by SP_Sale.V_Date,SP_Sale.V_Type" & _
            " Union All " & _
            "SELECT 'W' as RepType,SP_Sale.V_Date,  " & cIIF("sp_sale.v_type='W_SIC'", "'Cash  '", "'Credit'") & " as CashCr, " & _
            "0 as TaxableAmt," & _
            "0 as TaxpaidAmt,0 as GenSurTrns, " & _
            "0 as Tax," & _
            "0  as TaxOnMRP,0 as pack," & _
            "0 as TotalAmt, " & _
            "0 as TaxableAmt2,0 as TaxpaidAmt2,0 as GenSurTrns2,0 as Tax2, 0 as TaxOnMRP2, 0 as Pack2,0 as TotalAmt2, " & cIIF("SP_Stock.Purpose='W' and SP_Stock.tax_yn=1", "sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.Rate)", "0") & " as warrTB, " & cIIF("SP_Stock.Purpose='W' and SP_Stock.tax_yn=0", "sum((sp_stock.QTY_Iss-sp_stock.QTY_Ret) *sp_stock.Rate)", "0") & " as warrTP, 0 as SatAmt, 0 as SatAmt2  " & _
            "FROM SP_Sale Left Join Sp_Stock on SP_Sale.Docid=SP_Stock.Invoice_DocId" & Condstr & " and sp_sale.v_type in ('W_SIC','W_SIR')  group by SP_Sale.V_Date,SP_Sale.V_Type,SP_Stock.Purpose, SP_Stock.tax_yn,sp_stock.Amount"

            
        mQRY1 = "SELECT Taxforms.Form_desc, " & cIIF("SubGroupType.Description<>''", "SubGroupType.Description", "'Others'") & " as Descrip, " & _
            "Sum(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB+SP_Sale.SprAmt_TB+SP_Sale.OilAmt_TB-SP_Sale.D_Amt_TB) AS TaxableAmt, " & _
            "Sum(SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TP-SP_Sale.D_Amt_TP) AS TaxpaidAmt, " & _
            "Sum(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) as GenSurTrns,sum(SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt+SP_Sale.TOT_Amt+SP_Sale.ReSalTax_Amt+SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_AmtMRP ) as Tax," & _
            "0 as TaxOnMRP,sum(Sp_Sale.Total_Amt) as TotalAmt " & _
            "FROM (SP_Sale LEFT JOIN (SubGroup LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type) ON SP_Sale.Party_Code = SubGroup.SubCode) Left join taxforms on taxforms.Form_Code = sp_sale.Form_Code " & Condstr & " and sp_sale.v_type in ('SYSIC','SYSIR') group by SubGroupType.Description,Taxforms.Form_desc"
        RepName = "SprDailySale"
    End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    Set RstRep1 = New Recordset
    RstRep1.CursorLocation = adUseClient
    RstRep1.Open (mQRY1), GCn, adOpenDynamic, adLockOptimistic
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub Formulas()
On Error GoTo ELoop
Dim RstCompDet As ADODB.Recordset
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
    'SprStkReg
Select Case GRepFormName
    Case SprMRPTaxClaimReg, SprMatReg, SprStkTrf, SprMonthSale, SprDailySale, SprPartPur, SprPartSale, SprSaleSum, SprPurSum, SprStkBin, SprPartMovement, VehMoneyRect, DailyLubCon
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "   :  Report Type :  " & FGrid.TextMatrix(List1, 1) & " '"
            End Select
        If (GRepFormName = SprMRPTaxClaimReg Or GRepFormName = SprSaleSum) And UCase(rpt.FormulaFields(I).FormulaFieldName) = "TOTCAPTION" Then
             rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
        End If
        Next
    Case SprOthPurReg
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
        
    Case SprStkAgeing
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("List1")
                    rpt.FormulaFields(I).TEXT = "'" & FGrid.TextMatrix(List1, 1) & "'"
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'As On Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "'"
                Case UCase("TDay1")
                    rpt.FormulaFields(I).TEXT = "'0-' + '" & FGrid.TextMatrix(Cat1, 1) & "'"
                Case UCase("TDay2")
                    If FGrid.TextMatrix(Cat1, 1) <> "" Then
                        rpt.FormulaFields(I).TEXT = "'" & Val(FGrid.TextMatrix(Cat1, 1) + 1) & "' + '-' + '" & FGrid.TextMatrix(Cat2, 1) & "'"
                    End If
                Case UCase("TDay3")
                    If FGrid.TextMatrix(Cat2, 1) <> "" Then
                        rpt.FormulaFields(I).TEXT = "'" & Val(FGrid.TextMatrix(Cat2, 1) + 1) & "' + '-' + '" & FGrid.TextMatrix(Cat3, 1) & "'"
                    End If
                Case UCase("TDay4")
                    If FGrid.TextMatrix(Cat3, 1) <> "" Then
                        rpt.FormulaFields(I).TEXT = "'" & Val(FGrid.TextMatrix(Cat3, 1) + 1) & "' + '-' + '" & FGrid.TextMatrix(Cat4, 1) & "'"
                    End If
                Case UCase("TDay5")
                    If FGrid.TextMatrix(Cat4, 1) <> "" Then
                        rpt.FormulaFields(I).TEXT = "'" & Val(FGrid.TextMatrix(Cat4, 1) + 1) & "' + '-' + '" & FGrid.TextMatrix(Cat5, 1) & "'"
                    End If
                Case UCase("TDayLast")
                    rpt.FormulaFields(I).TEXT = "'Above ' + '" & mDays & "'"
            End Select
        Next
    Case SprPOrdReg, SprSOrdReg, SalesmanWPending
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'Upto Date :'+ '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next

    Case SprCtrRateVari, SprPurRateVari
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
    
    Case SprStkReg, SprSaleRet, SprSaleTrfReg, WarTaxReimbReg, InputTaxReg, OutputTaxReg
        Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax, LstNoS, LstDateS from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("List1")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Sales'"
                Case UCase("Comp_LST")
                    rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecLST & "  " & RstCompDet!S_SecLST_Date & "'"
                Case UCase("Comp_CST")
                    rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecCST & "  " & RstCompDet!S_SecCST_Date & "'"
                Case UCase("PubStartDate")
                    rpt.FormulaFields(I).TEXT = "'" & PubStartDate & "'"
                Case UCase("PubEndDate")
                    rpt.FormulaFields(I).TEXT = "'" & PubEndDate & "'"
                Case UCase("PubStartDate")
                    rpt.FormulaFields(I).TEXT = "'" & PubStartDate & "'"
            End Select
            If (GRepFormName = SprSaleTrfReg Or GRepFormName = SprSaleRet) And UCase(rpt.FormulaFields(I).FormulaFieldName) = "TOTCAPTION" Then
                    rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
            End If
        Next
    
    Case PurTaxSumm, SaleTaxSumm, SpareSaleAccount, SparePurchaseAccount
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("List1")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' "
                Case UCase("List2")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List2, 1) & "' + ' Purchase'"
            End Select
            If (GRepFormName = SprSaleTrfReg Or GRepFormName = SprSaleRet) And UCase(rpt.FormulaFields(I).FormulaFieldName) = "TOTCAPTION" Then
                    rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
            End If
        Next
    Case WksSaleReg
            For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("List1")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Sales'"
                Case "CashCredit"
                    rpt.FormulaFields(I).TEXT = "'" & FGrid.TextMatrix(List1, 1) & "'"
            End Select
            If (GRepFormName = SprSaleTrfReg Or GRepFormName = SprSaleRet) And UCase(rpt.FormulaFields(I).FormulaFieldName) = "TOTCAPTION" Then
                    rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
            End If
        Next
    
    Case SprSaleReg
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("List1")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Sales'"
                Case UCase("RepGrp")
                    If UCase(FGrid.TextMatrix(List2, 1)) = "SALESMANWISE" Then
                        rpt.FormulaFields(I).TEXT = "Emp_Name"
                    Else
                        rpt.FormulaFields(I).TEXT = "Party_Name"
                    End If
            End Select
            If GRepFormName = SprSaleReg And UCase(rpt.FormulaFields(I).FormulaFieldName) = "TOTCAPTION" Then
                    rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
            End If
        Next
    Case SprSaleTaxCtrlStmt
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("GTO")
                    rpt.FormulaFields(I).TEXT = myRst1!GTO '- myRst1!TotRet
                Case UCase("DISCOUNT")
                    rpt.FormulaFields(I).TEXT = myRst1!DisAmt
                Case UCase("SALEAFTDISC")
                    rpt.FormulaFields(I).TEXT = myRst1!SprAmtMrpTB + myRst1!SprAmtMrpTP + myRst1!OilAmtMrpTB + myRst1!OilAmtMrpTP + myRst1!SprAmtTB + myRst1!SprAmtTP + myRst1!OilAmtTB + myRst1!OilAmtTP
                Case UCase("TP_SPRSAL")
                    If PubDiscOnLube = 1 Then
                        rpt.FormulaFields(I).TEXT = (myRst1!SprAmtMrpTP + myRst1!SprAmtTP)
                    Else
                        rpt.FormulaFields(I).TEXT = (myRst1!SprAmtMrpTP + myRst1!SprAmtTP) - (myRst1!SprAmtMrpTPRet + myRst1!SprAmtTPRet)
                    End If
                Case UCase("TOTLUBSAL")
                    If PubDiscOnLube = 1 Then
                        rpt.FormulaFields(I).TEXT = (myRst1!OilAmtMrpTB + myRst1!OilAmtMrpTP + myRst1!OilAmtTB + myRst1!OilAmtTP)
                    Else
                        rpt.FormulaFields(I).TEXT = (myRst1!OilAmtMrpTB + myRst1!OilAmtMrpTP + myRst1!OilAmtTB + myRst1!OilAmtTP) - (myRst1!OilAmtMrpTBRet + myRst1!OilAmtMrpTPRet + myRst1!OilAmtTBRet + myRst1!OilAmtTPRet)
                    End If
            End Select
        Next
    Case SprPurReg, SprPurRet
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("List1")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Purchase'"
            End Select
        Next
    Case SprStkReOrd
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("FldTitle1")
                    Select Case FGrid.TextMatrix(List1, 1)
                        Case "Above Maximum"
                            rpt.FormulaFields(I).TEXT = "'Max Lvl'"
                        Case "Below Minimum"
                            rpt.FormulaFields(I).TEXT = "'Min Lvl'"
                        Case "Below ReOrder"
                            rpt.FormulaFields(I).TEXT = "'ReOrd Lvl'"
                    End Select
                Case UCase("FldTitle2")
                    Select Case FGrid.TextMatrix(List1, 1)
                        Case "Above Maximum"
                            rpt.FormulaFields(I).TEXT = "'Excess'"
                        Case "Below Minimum", "Below ReOrder"
                            rpt.FormulaFields(I).TEXT = "'Short'"
                    End Select
            End Select
        Next
     Case CashBankBook
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("CASHOP")
                    rpt.FormulaFields(I).TEXT = "'" & Format(CashOpening, "0.00") & "'"
                Case UCase("VIJBOP")
                    rpt.FormulaFields(I).TEXT = "'" & Format(BankOpening, "0.00") & "'"
                Case UCase("CASHCL")
                    rpt.FormulaFields(I).TEXT = "'" & Format(CashClosing, "0.00") & "'"
                Case UCase("VIJBCL")
                    rpt.FormulaFields(I).TEXT = "'" & Format(BankClosing, "0.00") & "'"
                Case UCase("INTEREST")
                    rpt.FormulaFields(I).TEXT = "'" & Format(TotalInterest, "0.00") & "'"
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub SprPurChl()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If GRepFormName = SprMatReg Then
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        Condstr = " Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And  SP_Purch.V_Type='" & SprMrRct & " ' And SP_Purch.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Purch.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and SP_Purch.Party_Code in (" & GridString1 & ")"
        If FGrid.TextMatrix(List1, 1) = "General" Then
            RepName = "SprMatReg"
        Else
            RepName = "SprMatRegParty"
        End If
        
         If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_purch.DocId", "3", "1") & " in (" & GridString2 & ")"
    If Check1(2).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_purch.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If


        If FGrid.TextMatrix(List2, 1) = "Cash" Then
            Condstr = Condstr & " and Sp_purch.Cash_Credit='Cash' "
        ElseIf FGrid.TextMatrix(List2, 1) = "Credit" Then
            Condstr = Condstr & " and Sp_purch.Cash_Credit='Credit' "
        End If
        
        
        If StrCmp(FGrid.TextMatrix(List3, 1), "Pending") Then
            Condstr = Condstr & " And (Sp_Purch.Invoice_DocId is null Or Sp_Purch.Invoice_DocId='') "
        ElseIf StrCmp(FGrid.TextMatrix(List3, 1), "Billed") Then
            Condstr = Condstr & " And (Sp_Purch.Invoice_DocId is Not null And Sp_Purch.Invoice_DocId<>'') "
        End If
    ElseIf GRepFormName = SprStkTrf Then
        Condstr = "SP_Purch.V_Type='" & SprMrTrf & "' And SP_Purch.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Purch.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP_Purch.DocID", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP_Purch.DocID", "3", "1") & " ='" & PubSiteCode & "' "
        End If
    
        RepName = "SprMatReg"
    End If
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        mQry = "SELECT SP_Purch.DocID,SP_Purch.Party_Name, SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date, SP_Purch.GR_RR_No, " & _
            "SP_Purch.GR_RR_Date, SP_Purch.Cash_Credit, SP_Purch.Tot_No_of_Items, SP_Purch.Tot_Doc_Qty, SP_Purch.Tot_Phy_Qty," & _
            "SP_Purch.Tot_Amt as Tot_Goods_Value, SP_Purch.NET_AMT, TaxForms.Form_Desc, SP_Stock.Part_No, SP_Stock.Qty_Doc, SP_Stock.Qty_Rec," & _
            "SP_Stock.Rate, SP_Purch.V_Type, SP_Purch.V_No, SP_Stock.Amount,SP_Purch.V_Date,(SP_Purch.Tot_Disc_Amt+SP_Purch.Tot_Ord_DiscAmt) as Tot_Disc_Amt,SP_Purch.Addition,SP_Purch.Deduction,SP_Purch.Tax_Amt " & _
            "FROM (SP_Purch INNER JOIN SP_Stock ON (SP_Purch.V_Type = SP_Stock.V_Type) AND (SP_Purch.V_No = SP_Stock.V_No)) " & _
            "LEFT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code " & _
            "Where " & Condstr & ""

    Else
  
    mQry = "SELECT SP_Purch.DocID, " & cIIF("Sp_Purch.Cash_Credit='Credit'", "(Select Name From SubGroup Where SubCode=Sp_Purch.Party_Code)", "SP_Purch.Party_Name") & " As Party_Name, SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date, SP_Purch.GR_RR_No, " & _
        "SP_Purch.GR_RR_Date, SP_Purch.Cash_Credit, SP_Purch.Tot_No_of_Items, SP_Purch.Tot_Doc_Qty, SP_Purch.Tot_Phy_Qty," & _
        "SP_Purch.Tot_Goods_Value, SP_Purch.NET_AMT, TaxForms.Form_Desc, SP_Stock.Part_No, SP_Stock.Qty_Doc, SP_Stock.Qty_Rec," & _
        "SP_Stock.Rate, SP_Purch.V_Type, SP_Purch.V_No, SP_Stock.Amount,SP_Purch.V_Date,SP_Purch.Tot_Disc_Amt,SP_Purch.Addition,SP_Purch.Deduction,SP_Purch.Tax_Amt,SP_Purch.Tot_Amt " & _
        "FROM (SP_Purch INNER JOIN SP_Stock ON (SP_Purch.V_Type = SP_Stock.V_Type) AND (SP_Purch.V_No = SP_Stock.V_No)) " & _
        "LEFT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code " & _
        "Where " & Condstr & ""
    End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
                   
    RepTitle = UCase(Me.CAPTION) + "[" + FGrid.TextMatrix(List1, 1) + "]"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SprSalePurReg()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset
Dim TmpRst1 As ADODB.Recordset
Dim TotAmtTb, TotAmtTp As Variant
Dim JDocId, InvdocId As String
Dim Clodate, VDt As Date
Dim NetLab, Misc, SDTAMT, NetVal As Double


Dim BillAmts As Double, I As Integer

'Date1,Date2,List1,List1,List1,List2,List1,List1
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If GRepFormName <> WarTaxReimbReg Then: If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub

If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub


Select Case GRepFormName
    Case SprSaleReg, SprSaleRet
        If GRepFormName = SprSaleReg Then
            If FGrid.TextMatrix(List3, 1) = "Both" Then
                If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Sale.V_Type In ('" & SprSlCsh & "','" & SprSlCre & "','" & WksSlCsh & "','" & WksSlCre & "','W_WWC','W_WWR') And "
                If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = "SP_Sale.V_Type in('" & SprSlCre & "','" & WksSlCre & "','W_WWR') And "
                If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = "SP_Sale.V_Type in('" & SprSlCsh & "','" & WksSlCsh & "','W_WWC') And "
            ElseIf FGrid.TextMatrix(List3, 1) = "Counter" Then
                If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Sale.V_Type In ('" & SprSlCsh & "','" & SprSlCre & "') And "
                If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = "SP_Sale.V_Type = '" & SprSlCre & "' And "
                If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = "SP_Sale.V_Type = '" & SprSlCsh & "' And "
            ElseIf FGrid.TextMatrix(List3, 1) = "Workshop" Then
                If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Sale.V_Type In ('" & WksSlCsh & "','" & WksSlCre & "','W_WWC','W_WWR') And "
                If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = "SP_Sale.V_Type IN ('" & WksSlCre & "','W_WWR') And "
                If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = "SP_Sale.V_Type IN ('" & WksSlCsh & "','W_WWC') And "
            End If
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " " & cMID("SP_Sale.DocId", "3", "1") & " in (" & GridString1 & ") AND "
            If Check1(1).Value = Checked Then
            If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " " & cMID("Sp_sale.docid", "3", "1") & " ='" & PubSiteCode & "' and "
           End If

            If FGrid.TextMatrix(List2, 1) = "SalesManWise" Then
                If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Sale.Rep_Code in (" & GridString2 & ") AND "
            Else
                If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Sale.Party_Code in (" & GridString2 & ") AND "
            End If
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " Left(SP_Sale.DocId,1) in (" & GridString3 & ") AND "
            
            Condstr = Condstr + "SP_Sale.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Sale.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

            If FGrid.TextMatrix(List1, 1) = "All" Then
                If FGrid.TextMatrix(List2, 1) = "General" Then
                    RepName = "SprSalRegAll"
                ElseIf FGrid.TextMatrix(List2, 1) = "SalesManWise" Then
                    RepName = "SprSalRegSManWise"
                Else
                    RepName = "SprSalRegAllParty"
                End If
            Else
                If FGrid.TextMatrix(List2, 1) = "General" Then
                    RepName = "SprSalReg"
                ElseIf FGrid.TextMatrix(List2, 1) = "SalesManWise" Then
                    RepName = "SprSalRegSManWise"
                Else
                    RepName = "SprSalRegParty"
                End If
            End If
        ElseIf GRepFormName = SprSaleRet Then
            If FGrid.TextMatrix(List1, 1) = "Transfer" Then Condstr = "SP_Sale.V_Type = '" & SprSlTrfRet & "' And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = "SP_Sale.V_Type = '" & SprSlRetCre & "' And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = "SP_Sale.V_Type = '" & SprSlRetCsh & "' And "
            If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Sale.V_Type in('" & SprSlRetCsh & "','" & SprSlRetCre & "') And "
            If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & "  " & cMID("SP_Sale.DocId", "3", "1") & " in (" & GridString1 & ") and "
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & "  " & cMID("SP_Sale.DocId", "3", "1") & " ='" & PubSiteCode & "' and "
    End If


            If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Sale.Party_Code in (" & GridString2 & ") AND "
            Condstr = Condstr + "SP_Sale.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Sale.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
            If FGrid.TextMatrix(List2, 1) = "General" Then
                RepName = "SprSalRegAll"
            Else
                RepName = "SprSalRegAllParty"
            End If
        End If
        'Disc on Lube Applicablity
        If PubDiscOnLube = 1 Then
            CondStr3 = "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB ) - ((SP_Sale.SprAmt_TB)* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB)) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " AS SprAmtTB, " & _
                "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.SprAmt_TP) - ((SP_Sale.SprAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP)) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP )))") & " AS SprAmtTP, " & _
                "" & cIIF("((SP_Sale.SprAmt_MRP_TB)+(OilAmt_MRP_TB))=0", "0", "((SP_Sale.SprAmt_MRP_TB) - ((SP_Sale.SprAmt_MRP_TB) * (SP_Sale.D_Amt_MRP_TB)) / ((SP_Sale.SprAmt_MRP_TB) + (SP_Sale.OilAmt_MRP_TB)))") & " AS SprAmtMRPTB, " & _
                "" & cIIF("((SP_Sale.SprAmt_MRP_TP)+(OilAmt_MRP_TP))=0", "0", "((SP_Sale.SprAmt_MRP_TP) - ((SP_Sale.SprAmt_MRP_TP) * (SP_Sale.D_Amt_MRP_TP)) / ((SP_Sale.SprAmt_MRP_TP) + (SP_Sale.OilAmt_MRP_TP)))") & " AS SprAmtMRPTP, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB) - ((SP_Sale.OilAmt_TB )* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB)) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " as OilAmtTB , " & _
                "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.OilAmt_TP) - ((SP_Sale.OilAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TB)) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP)))") & " as OilAmtTP , " & _
                "" & cIIF("((SP_Sale.SprAmt_MRP_TB)+(OilAmt_MRP_TB))=0", "0", "((SP_Sale.OilAmt_MRP_TB) - ((SP_Sale.OilAmt_MRP_TB) * (SP_Sale.D_Amt_MRP_TB)) / ((SP_Sale.SprAmt_MRP_TB) + (SP_Sale.OilAmt_MRP_TB)))") & " AS OilAmtMRPTB, " & _
                "" & cIIF("((SP_Sale.SprAmt_MRP_TP)+(OilAmt_MRP_TP))=0", "0", "((SP_Sale.OilAmt_MRP_TP) - ((SP_Sale.OilAmt_MRP_TP) * (SP_Sale.D_Amt_MRP_TP)) / ((SP_Sale.SprAmt_MRP_TP) + (SP_Sale.OilAmt_MRP_TP)))") & " AS OilAmtMRPTP, "
        Else
            CondStr3 = "" & cIIF("(SP_Sale.SprAmt_TB)=0", "0", "((SP_Sale.SprAmt_TB ) - (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB))") & " AS SprAmtTB, " & _
                "" & cIIF("(SP_Sale.SprAmt_TP)=0", "0", "((SP_Sale.SprAmt_TP) - (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP))") & " AS SprAmtTP, " & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB) - (SP_Sale.D_Amt_MRP_TB))") & " AS SprAmtMRPTB, " & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP)=0", "0", "((SP_Sale.SprAmt_MRP_TP) - (SP_Sale.D_Amt_MRP_TP))") & " AS SprAmtMRPTP, " & _
                "SP_Sale.OilAmt_TB AS OilAmtTB , " & _
                "SP_Sale.OilAmt_TP AS OilAmtTP , " & _
                "OilAmt_MRP_TB AS OilAmtMRPTB, " & _
                "OilAmt_MRP_TP AS OilAmtMRPTP, "
        End If
        If PubVATYN <> 1 Then
            If UCase(left(PubComp_Name, 3)) = "JMK" Then
                mQry = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type,(left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & "+ " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, " & _
                    "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB,SP_Sale.SprAmt_MRP_TP, " & _
                    CondStr3 & _
                    "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " AS SprTransTB, " & _
                    "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " as OilTransTB , " & _
                    "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                    "SP_Sale.Tax_Amt as TaxAmt,SP_Sale.Tax_Sur_Amt AS Tax_Sur_Amt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                    "SP_Sale.Tax_AmtMRP AS Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP AS TaxSur_AmtMRP,SP_Sale.TOT_Amt as TOT_AmtMRP," & _
                    "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.NetLab_Amt,Emp_Mast.Emp_Name,Job_Card.Lab_TaxAmt,Job_Card.LabAmt_TB,Job_Card.Lab_D_Amt " & _
                    " FROM ((SP_Sale LEFT JOIN Job_Card ON SP_Sale.Job_DocId=Job_Card.DocId) LEFT JOIN Emp_Mast on SP_Sale.Rep_Code=Emp_Mast.Emp_Code) Where " & Condstr & " order by SP_Sale.V_Date,Sp_Sale.V_No"
            Else
                mQry = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type,(left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, " & _
                    "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB,SP_Sale.SprAmt_MRP_TP, " & _
                    CondStr3 & _
                    "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " AS SprTransTB, " & _
                    "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " as OilTransTB , " & _
                    "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                    "SP_Sale.Tax_Amt as TaxAmt,SP_Sale.Tax_Sur_Amt AS Tax_Sur_Amt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                    "SP_Sale.Tax_AmtMRP AS Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP AS TaxSur_AmtMRP,SP_Sale.TOT_AmtMRP as TOT_AmtMRP," & _
                    "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.NetLab_Amt,Emp_Mast.Emp_Name,Job_Card.Lab_TaxAmt,Job_Card.LabAmt_TB,Job_Card.Lab_D_Amt " & _
                    " FROM ((SP_Sale LEFT JOIN Job_Card ON SP_Sale.Job_DocId=Job_Card.DocId) LEFT JOIN Emp_Mast on SP_Sale.Rep_Code=Emp_Mast.Emp_Code) Where " & Condstr & " "
                mQry = mQry & IIf(PubBackEnd = "A", "  order by SP_Sale.V_Date,Sp_Sale.V_No ", "")
            End If
         Else
            
              mQry = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type,(left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, " & _
                    "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB,SP_Sale.SprAmt_MRP_TP, " & _
                    CondStr3 & _
                    "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " AS SprTransTB, " & _
                    "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " as OilTransTB , " & _
                    "(SP_Sale.D_Amt_TB) as D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                    "0 AS TaxAmt,SP_Sale.Tax_Sur_Amt AS Tax_Sur_Amt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                    "0 AS Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP AS TaxSur_AmtMRP,SP_Sale.TOT_AmtMRP as TOT_AmtMRP," & _
                    "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.NetLab_Amt,Emp_Mast.Emp_Name,Job_Card.Lab_TaxAmt,Job_Card.LabAmt_TB,Job_Card.Lab_D_Amt,0 as DiscAmt, Sp_Sale.SatAmt " & _
                    " FROM ((SP_Sale LEFT JOIN Job_Card ON SP_Sale.Job_DocId=Job_Card.DocId) LEFT JOIN Emp_Mast on SP_Sale.Rep_Code=Emp_Mast.Emp_Code) Where " & Condstr & " Group By SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type,SP_Sale.V_No, SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB,SP_Sale.SprAmt_MRP_TP, SP_Sale.SprAmt_TB,SP_Sale.D_Amt_TB,SP_Sale.D_Amt_MRP_TB,SP_Sale.SprAmt_TP,SP_Sale.D_Amt_TP,SP_Sale.D_Amt_MRP_TP,SP_Sale.SprAmt_MRP_TB," & _
                    "SP_Sale.D_Amt_MRP_TB,SP_Sale.SprAmt_MRP_TP,SP_Sale.D_Amt_MRP_TP, SP_Sale.OilAmt_TB , SP_Sale.OilAmt_TP,SP_Sale.OilAmt_MRP_TB ," & _
                    "SP_Sale.OilAmt_MRP_TP ,SP_Sale.Trans_Amt,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt,SP_Sale.Tax_Amt ,SP_Sale.Tax_Sur_Amt , SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt,SP_Sale.Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP ,SP_Sale.TOT_AmtMRP ,SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.NetLab_Amt,Emp_Mast.Emp_Name,Job_Card.Lab_TaxAmt,Job_Card.LabAmt_TB,Job_Card.Lab_D_Amt, Sp_Sale.SatAmt  "
            mQry = mQry & IIf(PubBackEnd = "A", "  order by SP_Sale.V_Date,Sp_Sale.V_No ", "")
                    
                    
            mQry = mQry + " Union All SELECT SP_Sale.DocID, SP_Sale.V_Date, '' as V_Type,'' as V_No, " & _
                "SP_Sale.Party_Name, Max(Sp_Sale.Cash_Credit) As Cash_Credit, 0 as SprAmt_MRP_TB,0 as SprAmt_MRP_TP, " & _
                "0 AS SprAmtTB, " & _
                "0 AS SprAmtTP, " & _
                "0 AS SprAmtMRPTB, " & _
                "0 AS SprAmtMRPTP, " & _
                "0 AS OilAmtTB , " & _
                "0 AS OilAmtTP , " & _
                "0 AS OilAmtMRPTB, " & _
                "0 AS OilAmtMRPTP, " & _
                "0 AS SprTransTB, " & _
                "0 as OilTransTB , " & _
                "0 as D_Amt_TB,0 as D_Amt_TP,0 as Gen_Sur_Amt, 0 as Trans_Amt," & _
                "" & cIIF(cCStr(xIsNull("SP_Stock.MRP_YN", "")) & "+" & cCStr(xIsNull("SP_Stock.Tax_YN", "")) & "='01' and SP_Stock.Purpose in ('C','')", "sum(SP_Stock.TaxAmt)", "0") & " as TaxAmt,0 as Tax_Sur_Amt, 0 as TOT_Amt,0 as ReSalTax_Amt," & _
                "" & cIIF(cCStr(xIsNull("SP_Stock.MRP_YN", "")) & "+" & cCStr(xIsNull("SP_Stock.Tax_YN", "")) & "='11' and SP_Stock.Purpose in ('C','')", "sum(SP_Stock.TaxAmt)", "0") & " as Tax_AmtMRP,0 AS TaxSur_AmtMRP,0 as TOT_AmtMRP," & _
                "0 as Packing,0 as Rounded, 0 as Total_Amt,0 as NetLab_Amt,'' as Emp_Name,0 as Lab_TaxAmt,0 as LabAmt_TB,0 as Lab_D_Amt,Sum(Disc_amt2) as DiscAmt, 0 As SatAmt " & _
                " FROM ((SP_Sale LEFT JOIN Job_Card ON SP_Sale.Job_DocId=Job_Card.DocId) LEFT JOIN Emp_Mast on SP_Sale.Rep_Code=Emp_Mast.Emp_Code) Left Join Sp_Stock on SP_Sale.DocId=SP_Stock.Invoice_DocId Where " & Condstr & " Group By SP_Sale.DocId,SP_Sale.V_Date,SP_Sale.Party_Name,SP_Stock.MRP_YN,SP_Stock.Tax_YN,SP_Stock.TaxAmt,SP_Stock.Purpose  "
            mQry = mQry & IIf(PubBackEnd = "A", "  order by SP_Sale.V_Date ", "")
          End If
            
    Case WarTaxReimbReg
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " " & cMID("SP_Stock.DocId", "3", "1") & " in (" & GridString1 & ") AND "
            If Check1(1).Value = Checked Then
            If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & "  " & cMID("Sp_stock.docid", "3", "1") & " ='" & PubSiteCode & "' and "
            End If

            If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Sale.Party_Code in (" & GridString2 & ") AND "
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " Left(SP_Stock.DocId,1) in (" & GridString3 & ") AND "
                        
            Condstr = Condstr + "SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
        
        mQry = "SELECT SP_Stock.DocID, SP_Stock.V_Date, SP_Stock.V_Type,(left(SP_Stock.Docid,1)+ " & cMID("SP_Stock.Docid", "3", "2") & " + " & cMID("SP_Stock.Docid", "8", "1") & " + " & cCStr("SP_Stock.V_No") & ") as V_No, " & _
            "SP_Stock.Tax_YN,SP_Stock.MRP_YN,SP_Stock.Amount, " & _
            "" & cIIF("(SP_Stock.MRP_YN=0)and (SP_Stock.Tax_YN=1)", "SP_Stock.Amount", "0") & " AS SprAmtTB, " & _
            "" & cIIF("((SP_Stock.MRP_YN=0) or (SP_Stock.MRP_YN=1)) and(SP_Stock.Tax_YN=0)", "SP_Stock.Amount", "0") & " AS SprAmtTP, " & _
            "" & cIIF("(SP_Stock.MRP_YN=1) and (SP_Stock.Tax_YN=1)", "SP_Stock.Amount", "0") & " AS SprAmtMRPTB, " & _
            "SP_Sale.Tax_Per as TaxPer,SP_Sale.Tax_Sur_Per AS Tax_Sur_Per" & _
            " FROM (SP_Stock LEFT JOIN SP_Sale ON SP_Stock.Job_DocId=SP_Sale.Job_DocId) Where SP_Stock.Purpose='W' and SP_Stock.Invoice_DocID <> '' and " & Condstr & " order by SP_Stock.V_Date,Sp_Stock.V_No"
            
            RepName = "WarTaxReimbReg"
            
    Case WksSaleReg
            If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Sale.V_Type In ('" & WksSlCsh & "','" & WksSlCre & "','W_WWC','W_WWR') And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = "SP_Sale.V_Type IN ('" & WksSlCre & "','W_WWR') And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = "SP_Sale.V_Type IN ('" & WksSlCsh & "','W_WWC') And "
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ") AND "
            If Check1(1).Value = Checked Then
            If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & "  " & cMID("Job_card.docid", "3", "1") & " ='" & PubSiteCode & "' and "
            End If

            If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Sale.Party_Code in (" & GridString2 & ") AND "
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " Left(Job_Card.DocId,1) in (" & GridString3 & ")  and "
            Condstr = Condstr + "Job_Card.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and Job_Card.JobCloseDate<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
            
            If FGrid.TextMatrix(List2, 1) = "General" Then
            
                If UCase(left(PubComp_Name, 3)) = "JMK" Then
                
                    If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Sale.V_Type In ('" & WksSlCsh & "','" & WksSlCre & "','SYSIC','SYSIR','W_WWC','W_WWR') And "
                    If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = "SP_Sale.V_Type in('" & WksSlCre & "','SYSIR','W_WWR') And "
                    If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = "SP_Sale.V_Type in('" & WksSlCsh & "','SYSIC','W_WWC') And "
                    If Check1(1).Value = Unchecked Then Condstr = Condstr & " " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ") AND "
                    If Check1(1).Value = Checked Then
                    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & "  " & cMID("Job_card.docid", "3", "1") & " ='" & PubSiteCode & "' and "
                    End If

                    If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Sale.Party_Code in (" & GridString2 & ") AND "
                    If Check1(3).Value = Unchecked Then Condstr = Condstr & " Left(Job_Card.DocId,1) in (" & GridString3 & ")  and ( "
                    Condstr = Condstr + " ( Job_Card.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and Job_Card.JobCloseDate<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
                    Condstr = Condstr + " or SP_Sale.V_date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Sale.V_date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " )"
                
'                    mQRY = "SELECT Job_Card.DocId as JobDocID,Job_Card.JobCloseDate," & _
'                        "SP_Sale.DocID as InvDocID, SP_Sale.V_Date, SP_Sale.V_Type, (left(SP_Sale.Docid,1)& mid(SP_Sale.Docid,3,2) & mid(SP_Sale.Docid,8,1)& SP_Sale.V_No) as V_No, SP_Sale.Party_Name," & _
'                        "SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, SP_Sale.SprAmt_MRP_TP, " & _
'                        "(SP_Sale.SprAmt_TB+SP_Sale.SprAmt_MRP_TB - SP_Sale.D_Amt_TB) AS SprAmtTB, " & _
'                        "(SP_Sale.SprAmt_TP+SP_Sale.SprAmt_MRP_TP - SP_Sale.D_Amt_TP) AS SprAmtTP, " & _
'                        "(SP_Sale.OilAmt_TB+SP_Sale.OilAmt_MRP_TB) AS OilAmtTB, " & _
'                        "(SP_Sale.OilAmt_TP+SP_Sale.OilAmt_MRP_TP) AS OilAmtTP, " & _
'                        "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
'                        "(SP_Sale.Tax_Amt+SP_Sale.Tax_Sur_Amt) as TaxAmt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
'                        "(SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_AmtMRP) as TaxOnMRP," & _
'                        "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.DocId_InvSpr," & _
'                        "Job_Card.NetLab_Amt,Job_Card.DocId_InvLab,Job_Card.Lab_D_Amt,iif(Job_Card.Serv_type='W','Warranty','WorkShop Sale') as servType,Job_Card.Job_no " & _
'                        "FROM Job_Card LEFT JOIN SP_Sale ON Job_Card.DocId = SP_Sale.Job_DocID " & _
'                        "Where " & Condstr & " Order By Job_Card.JobCloseDate,Job_Card.DocId_InvLab"

                    Dim sQrySaleVat12$, sQrySaleVat4$, sQryVat12$, sQryVat4$
                     Dim SQrySaleSat3$, SQrySaleSat1_5$, SQrySaleSat1$, SQrySaleSat0_5$

                    If MsgBox("Do You Want Categorised VAT Report?", vbYesNo) = vbYes Then
                        sQrySaleVat12 = "Select Sum(Net_Amt) From Sp_Stock Where Sp_Stock.TaxPer>=12.5 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId"
                        sQrySaleVat4 = "Select Sum(Net_Amt) From Sp_Stock Where Sp_Stock.TaxPer<12 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId"
                        sQryVat12 = "Select Sum(TaxAmt) From Sp_Stock Where Sp_Stock.TaxPer>=12.5 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId"
                        sQryVat4 = "Select Sum(TaxAmt) From Sp_Stock Where Sp_Stock.TaxPer<12 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId"
                    Else
                        sQrySaleVat12 = "0"
                        sQrySaleVat4 = "0"
                        sQryVat12 = "0"
                        sQryVat4 = "0"
                    End If
                    'kunal start
                        SQrySaleSat3 = "Select Sum(SatAmt) From Sp_Stock Where Sp_Stock.SatPer=3 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId"
                        SQrySaleSat1_5 = "Select Sum(SatAmt) From Sp_Stock Where Sp_Stock.SatPer=1.5 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId"
                        SQrySaleSat1 = "Select Sum(SatAmt) From Sp_Stock Where Sp_Stock.SatPer=1 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId "
                        SQrySaleSat0_5 = "Select Sum(SatAmt) From Sp_Stock Where Sp_Stock.SatPer=0.5 and Sp_Stock.Invoice_DocId=Sp_Sale.DocId "
                     'kunal end

'                    mQry = "SELECT Job_Card.DocId as JobDocID,Job_Card.JobCloseDate," & _
'                        "SP_Sale.DocID as InvDocID, SP_Sale.V_Date, SP_Sale.V_Type, (left(SP_Sale.Docid,1)+" & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, SP_Sale.Party_Name," & _
'                        "SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, SP_Sale.SprAmt_MRP_TP, " & _
'                        "(SP_Sale.SprAmt_TB+SP_Sale.SprAmt_MRP_TB - SP_Sale.D_Amt_TB) AS SprAmtTB, " & _
'                        "(SP_Sale.SprAmt_TP+SP_Sale.SprAmt_MRP_TP - SP_Sale.D_Amt_TP) AS SprAmtTP, " & _
'                        "(SP_Sale.OilAmt_TB+SP_Sale.OilAmt_MRP_TB) AS OilAmtTB, " & _
'                        "(SP_Sale.OilAmt_TP+SP_Sale.OilAmt_MRP_TP) AS OilAmtTP, " & _
'                        "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
'                        "(SP_Sale.Tax_Amt+SP_Sale.Tax_Sur_Amt) as TaxAmt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
'                        "(SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP) as TaxOnMRP," & _
'                        "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.DocId_InvSpr," & _
'                        "Job_Card.NetLab_Amt,Job_Card.DocId_InvLab,Job_Card.Lab_D_Amt, " & _
'                        "" & cIIF("Job_Card.Serv_type='W'", "'Warranty'", cIIF("SP_Sale.Job_DocId=''", "''", "'WorkShop Sale'")) & " as servType,Job_Card.Job_no, " & cIIF("SP_Sale.Job_DocId=''", "'A.Counter'", "'B.WorkShop'") & " as SaleType,Sp_Sale.Packing as Misc,Sp_Sale.Party_name as PartyName,Sp_Sale.Cash_Credit,Job_Card.Lab_TaxAmt,Job_Card.Lab_RoundOff,Sp_Sale.Rounded, (" & sQrySaleVat12 & ") As SaleAmtVat12, (" & sQrySaleVat4 & ") As SaleAmtVat4, (" & sQryVat12 & ") As Vat12, (" & sQryVat4 & ") As Vat4, Sp_Sale.SatAmt " & _
'                        "FROM SP_Sale LEFT JOIN Job_Card ON Job_Card.DocId = SP_Sale.Job_DocID " & _
'                        "Where " & Condstr & " Order By Job_Card.JobCloseDate,Job_Card.DocId_InvLab"
                 mQry = "SELECT Job_Card.DocId as JobDocID,Job_Card.JobCloseDate," & _
                        "SP_Sale.DocID as InvDocID, SP_Sale.V_Date, SP_Sale.V_Type, (left(SP_Sale.Docid,1)+" & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, SP_Sale.Party_Name," & _
                        "SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, SP_Sale.SprAmt_MRP_TP, " & _
                        "(SP_Sale.SprAmt_TB+SP_Sale.SprAmt_MRP_TB - SP_Sale.D_Amt_TB) AS SprAmtTB, " & _
                        "(SP_Sale.SprAmt_TP+SP_Sale.SprAmt_MRP_TP - SP_Sale.D_Amt_TP) AS SprAmtTP, " & _
                        "(SP_Sale.OilAmt_TB+SP_Sale.OilAmt_MRP_TB) AS OilAmtTB, " & _
                        "(SP_Sale.OilAmt_TP+SP_Sale.OilAmt_MRP_TP) AS OilAmtTP, " & _
                        "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                        "(SP_Sale.Tax_Amt+SP_Sale.Tax_Sur_Amt) as TaxAmt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                        "(SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP) as TaxOnMRP," & _
                        "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.DocId_InvSpr," & _
                        "Job_Card.NetLab_Amt,Job_Card.DocId_InvLab,Job_Card.Lab_D_Amt, " & _
                        "" & cIIF("Job_Card.Serv_type='W'", "'Warranty'", cIIF("SP_Sale.Job_DocId=''", "''", "'WorkShop Sale'")) & " as servType,Job_Card.Job_no, " & cIIF("SP_Sale.Job_DocId=''", "'A.Counter'", "'B.WorkShop'") & " as SaleType,Sp_Sale.Packing as Misc,Sp_Sale.Party_name as PartyName,Sp_Sale.Cash_Credit,Job_Card.Lab_TaxAmt,Job_Card.Lab_RoundOff,Sp_Sale.Rounded, (" & sQrySaleVat12 & ") As SaleAmtVat12, " & _
                        "(" & SQrySaleSat3 & ") As SaleSat3,(" & SQrySaleSat1_5 & ") As SaleSat1_5,(" & SQrySaleSat1 & ") As SaleSat1 ,(" & SQrySaleSat0_5 & ") As SaleSat0_5, " & _
                        "(" & sQrySaleVat4 & ") As SaleAmtVat4, (" & sQryVat12 & ") As Vat12, (" & sQryVat4 & ") As Vat4, Sp_Sale.SatAmt " & _
                        "FROM SP_Sale LEFT JOIN Job_Card ON Job_Card.DocId = SP_Sale.Job_DocID " & _
                        "Where " & Condstr & " Order By Job_Card.JobCloseDate,Job_Card.DocId_InvLab"
                        
                    
                        RepName = "WksSalRegJMK"
                Else
                    mQry = "SELECT Job_Card.DocId as JobDocID,Job_Card.JobCloseDate," & _
                        "SP_Sale.DocID as InvDocID, SP_Sale.V_Date, SP_Sale.V_Type, (left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, SP_Sale.Party_Name," & _
                        "SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, SP_Sale.SprAmt_MRP_TP, " & _
                        "(SP_Sale.SprAmt_TB+SP_Sale.SprAmt_MRP_TB - SP_Sale.D_Amt_TB) AS SprAmtTB, " & _
                        "(SP_Sale.SprAmt_TP+SP_Sale.SprAmt_MRP_TP - SP_Sale.D_Amt_TP) AS SprAmtTP, " & _
                        "(SP_Sale.OilAmt_TB+SP_Sale.OilAmt_MRP_TB) AS OilAmtTB, " & _
                        "(SP_Sale.OilAmt_TP+SP_Sale.OilAmt_MRP_TP) AS OilAmtTP, " & _
                        "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                        "(SP_Sale.Tax_Amt+SP_Sale.Tax_Sur_Amt) as TaxAmt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                        "(SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_AmtMRP) as TaxOnMRP," & _
                        "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.DocId_InvSpr," & _
                        "Job_Card.NetLab_Amt,Job_Card.DocId_InvLab,Job_Card.Lab_D_Amt,Job_Card.Serv_type,Job_Card.Lab_TaxPer,Job_Card.Lab_TaxAmt  " & _
                        "FROM Job_Card LEFT JOIN SP_Sale ON Job_Card.DocId = SP_Sale.Job_DocID " & _
                        "Where " & Condstr
                    mQry = mQry & IIf(PubBackEnd = "A", "  Order By Job_Card.JobCloseDate,Job_Card.DocId_InvLab ", " ")
                    
                    RepName = "WksSalReg"
                End If
            Else
                mQry = "SELECT Job_Card.DocId as JobDocID,Job_Card.JobCloseDate," & _
                "SP_Sale.DocID as InvDocID, SP_Sale.V_Date, SP_Sale.V_Type, (left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, SP_Sale.Party_Name," & _
                "SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, SP_Sale.SprAmt_MRP_TP, " & _
                "(SP_Sale.SprAmt_TB+SP_Sale.SprAmt_MRP_TB - SP_Sale.D_Amt_TB) AS SprAmtTB, " & _
                "(SP_Sale.SprAmt_TP+SP_Sale.SprAmt_MRP_TP - SP_Sale.D_Amt_TP) AS SprAmtTP, " & _
                "(SP_Sale.OilAmt_TB+SP_Sale.OilAmt_MRP_TB) AS OilAmtTB, " & _
                "(SP_Sale.OilAmt_TP+SP_Sale.OilAmt_MRP_TP) AS OilAmtTP, " & _
                "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                "(SP_Sale.Tax_Amt+SP_Sale.Tax_Sur_Amt) as TaxAmt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                "(SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_AmtMRP) as TaxOnMRP," & _
                "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,Job_Card.DocId_InvSpr," & _
                "Job_Card.NetLab_Amt,Job_Card.DocId_InvLab,Job_Card.Lab_D_Amt,Job_Card.Serv_type " & _
                "FROM Job_Card LEFT JOIN SP_Sale ON Job_Card.DocId = SP_Sale.Job_DocID " & _
                "Where " & Condstr
                
                mQry = mQry & IIf(PubBackEnd = "A", "   Order By Job_Card.JobCloseDate,Job_Card.DocId_InvLab ", "")
                
                RepName = "WksSalRegParty"
            End If
            
        
            
    Case SprPurReg, SprPurRet
        If GRepFormName = SprPurReg Then
            Dim RptType$
            If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
            If FGrid.TextMatrix(List3, 1) = "Inv.Date" Then
                Condstr = "SP.Party_Doc_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP.Party_Doc_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
            Else
                Condstr = "SP.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
            End If
            Condstr2 = "and SP_Stock.V_Date2 >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date2<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
            
            If FGrid.TextMatrix(List1, 1) = "All" Then
                Condstr = Condstr & " And SP.V_Type In ('" & SprPurCre & "','" & SprPurCsh & "') "
                Condstr2 = Condstr2 & " And " & cMID("SP_Stock.Invoice_DocId", "4", "5") & " In ('" & SprPurCre & "','" & SprPurCsh & "') "
            ElseIf FGrid.TextMatrix(List1, 1) = "Credit" Then
                Condstr = Condstr & " And SP.V_Type in ('" & SprPurCre & "')"
                Condstr2 = Condstr2 & " And " & cMID("SP_Stock.Invoice_DocId", "4", "5") & " ='" & SprPurCre & "' "
            ElseIf FGrid.TextMatrix(List1, 1) = "Cash" Then
                Condstr = Condstr & " And SP.V_Type in ('" & SprPurCsh & "')"
                Condstr2 = Condstr2 & " And " & cMID("SP_Stock.Invoice_DocId", "4", "5") & " ='" & SprPurCsh & "' "
            End If
            'Checking of Report Type
            If FGrid.TextMatrix(List2, 1) = "With Detail" Then
                RptType = "W"
            Else
                RptType = "O"
            End If
            
            If Check1(2).Value = Unchecked Then
                Condstr = Condstr & " AND SP.Party_Code in (" & GridString2 & ") "
                Condstr2 = Condstr2 & " AND SP_Stock.Party_Code in (" & GridString2 & ") "
            End If
            
            If Check1(3).Value = Unchecked Then
                If StrCmp(left(PubComp_Name, 4), "Yash") Then
                    Condstr = Condstr & " and SP.Form_Code in (" & GridString3 & ")"
                    'Condstr2 = Condstr2 & " and left(SP_Stock.Invoice_DocId,1) in (" & GridString3 & ")"
                Else
                    Condstr = Condstr & " and left(SP_Stock.DocId,1) in (" & GridString3 & ")"
                    Condstr = Condstr & " and Part.Div_Code in (" & GridString3 & ")"
                    'Condstr2 = Condstr2 & " and left(SP_Stock.Invoice_DocId,1) in (" & GridString3 & ")"
                End If
            End If
            If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Part.Part_Grade in (" & GridString4 & ")"
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " in (" & GridString1 & ")"
            If Check1(1).Value = Checked Then
            If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("sp.DocId", "3", "1") & " ='" & PubSiteCode & "' "
            End If

            
            'for lube purchase
            If GridString4 = "'L'" Then
                mQry = "SELECT '" & RptType & "' as RptType, SP.Docid,SP.V_Date,(left(SP.Docid,1)+ " & cMID("SP.Docid", "3", "2") & " + " & cMID("SP.Docid", "8", "1") & "+ " & cCStr("SP.V_No") & ") as V_No, " & cMID("SP.DocID", "9", "5") & " as VPrefix,SP.Party_Doc_No,SP.Party_Doc_Date,SP.Party_Name,SP.RoadPermit_No," & _
                    " (SP.SprAmt_MRP_TB+SP.SprAmt_TB) AS SprAmtTB,(SP.SprAmt_MRP_TP+SP.SprAmt_TP) AS SprAmtTP, " & _
                    " (SP.OilAmt_MRP_TP+SP.OilAmt_TP) AS OilAmtTP,(SP.OilAmt_MRP_TB+SP.OilAmt_TB) AS OilAmtTB, " & _
                    " " & cIIF("SP_Stock.Net_Amt2=0", "SP_Stock.Net_Amt", "SP_Stock.Net_Amt2") & " AS Tot_Goods_Value,SP.Tot_Amt,SP.Tot_Disc_Amt,SP.Tot_Ord_DiscAmt,SP.Tax_Amt,SP.Addition,SP.Deduction, " & _
                    " SP.Net_Amt,SP.EntryTaxPer,SP.EntryTaxAmt,(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation) as TotPurAmt, " & _
                    " " & cIIF("SP.V_Type='" & SprPurCsh & "'", "(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation)", "0") & " as TotCash, " & _
                    " " & cIIF("SP.V_Type='" & SprPurCre & "'", "(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation)", "0") & " as TotCr,SP_Stock.Order_DocId,0 as ChkVal,SP.Transportation, Sp.L_C " & _
                    " FROM (((SP_Purch as SP left join SP_Stock on SP.Docid=SP_Stock.Invoice_Docid) " & _
                    " Left Join Part on SP_Stock.Part_No=Part.Part_No and Part.Div_Code=Left(SP_Stock.DocId,1)) " & _
                    " Left Join TaxForms on SP.Form_Code=TaxForms.Form_Code) " & _
                    " WHERE " & Condstr & " and SP_Stock.V_Type='SXGR' "
                mQry = mQry & IIf(PubBackEnd = "A", " order by SP.V_Date,SP.Docid ", "")
            
            'For All Purchases
            ElseIf GridString4 = "" Then
                mQry = "SELECT '" & RptType & "' as RptType, SP.Docid,SP.V_Date,(left(SP.Docid,1)+ " & cMID("SP.Docid", "3", "2") & "  + " & cMID("SP.Docid", "8", "1") & " + " & cCStr("SP.V_No") & ") as V_No, " & cMID("SP.DocID", "9", "5") & " as VPrefix,SP.Party_Doc_No,SP.Party_Doc_Date,SP.Party_Name,SP.RoadPermit_No," & _
                    " (SP.SprAmt_MRP_TB+SP.SprAmt_TB) AS SprAmtTB,(SP.SprAmt_MRP_TP+SP.SprAmt_TP) AS SprAmtTP, " & _
                    " (SP.OilAmt_MRP_TP+SP.OilAmt_TP) AS OilAmtTP,(SP.OilAmt_MRP_TB+SP.OilAmt_TB) AS OilAmtTB, " & _
                    " " & cIIF("SP_Stock.Net_Amt2=0", "SP_Stock.Net_Amt", "SP_Stock.Net_Amt2") & " as Tot_Goods_Value,SP.Tot_Amt,SP.Tot_Disc_Amt,SP.Tot_Ord_DiscAmt,SP.Tax_Amt,SP.Addition,SP.Deduction, " & _
                    " SP.Net_Amt,SP.EntryTaxPer,SP.EntryTaxAmt,(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation) as TotPurAmt, " & _
                    " " & cIIF("SP.V_Type='" & SprPurCsh & "'", "(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation)", "0") & " as TotCash, " & _
                    " " & cIIF("SP.V_Type='" & SprPurCre & "'", "(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation)", "0") & " as TotCr,SP_Stock.Order_DocId,0 as ChkVal,SP.Transportation, SP.L_C,SP.SatAmt  " & _
                    " FROM (((SP_Purch as SP left join SP_Stock on SP.Docid=SP_Stock.Invoice_Docid) " & _
                    " Left Join Part on SP_Stock.Part_No=Part.Part_No and Part.Div_Code=Left(SP_Stock.DocId,1)) " & _
                    " Left Join TaxForms on SP.Form_Code=TaxForms.Form_Code) " & _
                    " WHERE " & Condstr & " and SP_Stock.V_Type='SXGR' "
                mQry = mQry & " order by SP.V_Date,SP.Docid "
            Else
                mQry = "SELECT '" & RptType & "' as RptType, SP.Docid,SP.V_Date,(left(SP.Docid,1)+ " & cMID("SP.Docid", "3", "2") & " + " & cMID("SP.Docid", "8", "1") & " + " & cCStr("SP.V_No") & " ) as V_No, " & cMID("SP.DocID", "9", "5") & " as VPrefix,SP.Party_Doc_No,SP.Party_Doc_Date,SP.Party_Name,SP.RoadPermit_No," & _
                    " (SP.SprAmt_MRP_TB+SP.SprAmt_TB) AS SprAmtTB,(SP.SprAmt_MRP_TP+SP.SprAmt_TP) AS SprAmtTP, " & _
                    " (SP.OilAmt_MRP_TP+SP.OilAmt_TP) AS OilAmtTP,(SP.OilAmt_MRP_TB+SP.OilAmt_TB) AS OilAmtTB, " & _
                    " " & cIIF("SP_Stock.Net_Amt2=0", "SP_Stock.Net_Amt", "SP_Stock.Net_Amt2") & " as Tot_Goods_Value,SP.Tot_Amt,SP.Tot_Disc_Amt,SP.Tot_Ord_DiscAmt,SP.Tax_Amt,SP.Addition,SP.Deduction, " & _
                    " SP.Net_Amt,SP.EntryTaxPer,SP.EntryTaxAmt,(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation) as TotPurAmt, " & _
                    " " & cIIF("SP.V_Type='" & SprPurCsh & "'", "(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation)", "0") & " as TotCash, " & _
                    " " & cIIF("SP.V_Type='" & SprPurCre & "'", "(SP.Net_Amt+SP.EntryTaxAmt+SP.Transportation)", "0") & " as TotCr,SP_Stock.Order_DocId,1 as ChkVal,SP.Transportation, SP.L_C " & _
                    " FROM (((SP_Purch as SP left join SP_Stock on SP.Docid=SP_Stock.Invoice_Docid) " & _
                    " Left Join Part on SP_Stock.Part_No=Part.Part_No and Part.Div_Code=Left(SP_Stock.DocId,1)) " & _
                    " Left Join TaxForms on SP.Form_Code=TaxForms.Form_Code) " & _
                    " WHERE " & Condstr & " "
                mQry = mQry & " order by SP.V_Date,SP.Docid "
            End If
            
            RepName = "SprPurReg"
        ElseIf GRepFormName = SprPurRet Then
            Set RstRep1 = New ADODB.Recordset
            With RstRep1
                .Fields.Append "DocId", adChar, 21, adFldIsNullable
                .Fields.Append "Part_Name", adChar, 40, adFldIsNullable
                .Fields.Append "Name", adChar, 40, adFldIsNullable
                .Fields.Append "Part_No", adChar, 40, adFldIsNullable
                .Fields.Append "Qty_Iss", adDouble, 10, adFldIsNullable
                .Fields.Append "Rate", adDouble, 10, adFldIsNullable
                .Fields.Append "V_No", adChar, 20, adFldIsNullable
                .Fields.Append "Prefix", adChar, 5, adFldIsNullable
                .Fields.Append "V_Date", adDate, 20, adFldIsNullable
                .Fields.Append "PartyName", adChar, 40, adFldIsNullable
                .Fields.Append "NetAmt", adDouble, 10, adFldIsNullable
                .Fields.Append "TaxAmt", adDouble, 10, adFldIsNullable
                .Fields.Append "Party_Doc_No", adChar, 40, adFldIsNullable
                .Fields.Append "BillAmt", adDouble, 10, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            
            If FGrid.TextMatrix(List1, 1) = "Transfer" Then Condstr = "SP_Purch.V_Type = '" & SprPrTrfRet & "' And "
            If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = "SP_Purch.V_Type = '" & SprPrRetCre & "' And "
            If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = "SP_Purch.V_Type = '" & SprPrRetCsh & "' And "
            If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Purch.V_Type In ('" & SprPrRetCsh & "','" & SprPrRetCre & "','" & SprPrTrfRet & "') And "
            
            If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Purch.Party_Code in (" & GridString2 & ") AND "
            Condstr = Condstr + "SP_Purch.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Purch.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP_Stock.docid", "3", "1") & " in (" & GridString1 & ")  "
               If Check1(1).Value = Checked Then
                    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP_Stock.docid", "3", "1") & " ='" & PubSiteCode & "'  "
               End If
                
                
            mQry = "SELECT SP_Stock.docid,Part.Part_Name, SubGroup.Name, SP_Stock.Part_No, SP_Stock.Qty_Iss, " & _
                "SP_Stock.Rate, (left(SP_Stock.Docid,1)+ " & cMID("SP_Stock.Docid", "3", "2") & " + " & cMID("SP_Stock.Docid", "8", "1") & " + " & cCStr("SP_Stock.V_NO") & ") as V_No, " & _
                "" & cMID("SP_Stock.DocID", "9", "5") & " as VPrefix, SP_Stock.V_Date,SP_Purch.Party_Name, Sp_Stock.TaxAmt, SP_Purch.Net_Amt, SP_Purch.Party_Doc_No, 0 as BIllAmt " & _
                "FROM ((SP_Purch LEFT JOIN SP_Stock ON SP_Purch.DocID = SP_Stock.DocID) LEFT JOIN SubGroup ON SP_Purch.Party_Code = SubGroup.SubCode) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) "
            mQry = mQry & " Where  Left(Sp_Purch.DocId,1)='" & PubDivCode & "' And " & Condstr
            If FGrid.TextMatrix(List2, 1) = "General" Then
                RepName = "SprPurRet"
            Else
                RepName = "SprPurRetParty"
            End If
        Set RstRep = New ADODB.Recordset
        RstRep.CursorLocation = adUseClient
        RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        If RstRep.RecordCount > 0 Then
        RstRep.MoveFirst
        For I = 1 To RstRep.RecordCount
            With RstRep1
                .AddNew
                .Fields("DocId") = RstRep!DocID
                .Fields("Part_Name") = RstRep!Part_Name
                .Fields("Name") = RstRep!Name
                .Fields("Part_No") = RstRep!Part_No
                .Fields("Qty_Iss") = RstRep!Qty_Iss
                .Fields("Rate") = RstRep!Rate
                .Fields("V_No") = RstRep!V_NO
                .Fields("Prefix") = RstRep!vPrefix
                .Fields("V_Date") = RstRep!V_DATE
                .Fields("PartyName") = RstRep!Party_Name
                .Fields("TaxAmt") = RstRep!TaxAmt
                .Fields("NetAmt") = RstRep!Net_Amt
                .Fields("Party_Doc_No") = RstRep!Party_Doc_No
                
                Set TmpRst = GCn.Execute("Select Net_Amt from SP_Purch where Left(Sp_Purch.DocId,1)='" & PubDivCode & "' and V_Type in ('SXPIC','SXPIR') and SP_Purch.Party_Doc_No='" & Trim(RstRep!Party_Doc_No) & "'")
                If TmpRst.RecordCount > 0 Then
                    BillAmts = TmpRst.Fields(0).Value
                End If
                If BillAmts > 0 Then
                    .Fields("BillAmt") = BillAmts
                End If
                .Update
            End With
        RstRep.MoveNext
        Next
        Set RstRep = RstRep1.Clone
        GoTo NXT
    End If
    End If
    End Select
    Set RstRep = New ADODB.Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
    
        If GRepFormName = WksSaleReg Then
            Set TmpRst = GCn.Execute(mQry)
                Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "JobDocId", adVarChar, 30, adFldIsNullable
                .Fields.Append "JobCloseDate", adDate, 15, adFldIsNullable
                .Fields.Append "InvDocID", adVarChar, 30, adFldIsNullable
                .Fields.Append "V_Date", adChar, 30, adFldIsNullable
                .Fields.Append "SprAmtTB", adDouble, 15, adFldIsNullable
                .Fields.Append "SprAmtTP", adDouble, 35, adFldIsNullable
                .Fields.Append "OilAmtTB", adDouble, 10, adFldIsNullable
                .Fields.Append "OilAmtTP", adDouble, 10, adFldIsNullable
                .Fields.Append "Total_Amt", adDouble, 10, adFldIsNullable
                .Fields.Append "NetLab_Amt", adDouble, 10, adFldIsNullable
                .Fields.Append "Misc", adDouble, 10, adFldIsNullable
                .Fields.Append "Tot_Amt", adDouble, 10, adFldIsNullable
                .Fields.Append "SaleType", adVarChar, 10, adFldIsNullable
                .Fields.Append "Job_No", adVarChar, 10, adFldIsNullable
                .Fields.Append "SprAmt_MRP_TB", adDouble, 15, adFldIsNullable
                .Fields.Append "TaxAmt", adDouble, 15, adFldIsNullable
                .Fields.Append "CName", adVarChar, 50, adFldIsNullable
                .Fields.Append "CRDr", adVarChar, 15, adFldIsNullable
                .Fields.Append "LabTax", adDouble, 15, adFldIsNullable
                .Fields.Append "LabROff", adDouble, 15, adFldIsNullable
                .Fields.Append "SprROff", adDouble, 15, adFldIsNullable
                .Fields.Append "SaleAmtVat12", adDouble, 10, adFldIsNullable
                .Fields.Append "SaleAmtVat4", adDouble, 10, adFldIsNullable
                .Fields.Append "Vat12", adDouble, 10, adFldIsNullable
                .Fields.Append "Vat4", adDouble, 10, adFldIsNullable
                .Fields.Append "SatAmt", adDouble, 10, adFldIsNullable
                'kunal
                .Fields.Append "SaleSat3", adDouble, 10, adFldIsNullable
                .Fields.Append "SaleSat1_5", adDouble, 10, adFldIsNullable
                .Fields.Append "SaleSat1", adDouble, 10, adFldIsNullable
                .Fields.Append "SaleSat0_5", adDouble, 10, adFldIsNullable
    
                
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            
            Set TmpRst = GCn.Execute(mQry)
            Do While TmpRst.EOF = False
                 With RstRep
                    .AddNew
                        !JobDocID = XNull(TmpRst!JobDocID)
                        !JobCloseDate = IIf(IsNull(TmpRst!JobCloseDate), TmpRst!V_DATE, TmpRst!JobCloseDate)
                        !InvdocId = XNull(TmpRst!InvdocId)
                        !V_DATE = XNull(TmpRst!V_DATE)
                        !SprAmtTB = VNull(TmpRst!SprAmtTB)
                        !SprAmtTP = XNull(TmpRst!SprAmtTP)
                        !OilAmtTB = XNull(TmpRst!OilAmtTB)
                        !OilAmtTP = XNull(TmpRst!OilAmtTP)
                        !NetLab_Amt = VNull(TmpRst!NetLab_Amt)
                        !Misc = VNull(TmpRst!Packing)
                        !Tot_Amt = VNull(TmpRst!Tot_Amt)
                        !Total_Amt = VNull(TmpRst!Total_Amt)
                        !SaleType = IIf(IsNull(TmpRst!JobDocID), "A-Counter", "B-WorkShop")
                        !Job_No = VNull(TmpRst!Job_No)
                        !SprAmt_MRP_TB = VNull(TmpRst!SprAmt_MRP_TB)
                        !TaxAmt = VNull(Val(TmpRst!TaxAmt)) '+ Val(TmpRst!TaxonMrp)
                        !CName = XNull(TmpRst!PartyName)
                        !CRDR = XNull(TmpRst!Cash_Credit)
                        !LabTax = VNull(TmpRst!Lab_TaxAmt)
                        !LabROff = VNull(TmpRst!Lab_RoundOff)
                        !SprROff = VNull(TmpRst!Rounded)
                        !SaleAmtVat12 = VNull(TmpRst!SaleAmtVat12)
                        !SaleAmtVat4 = VNull(TmpRst!SaleAmtVat4)
                        !Vat12 = VNull(TmpRst!Vat12)
                        !Vat4 = VNull(TmpRst!Vat4)
                        !SatAmt = VNull(TmpRst!SatAmt)
                        'kunal
                        !SaleSat3 = VNull(TmpRst!SaleSat3)
                        !SaleSat1_5 = VNull(TmpRst!SaleSat1_5)
                        !SaleSat1 = VNull(TmpRst!SaleSat1)
                        !SaleSat0_5 = VNull(TmpRst!SaleSat0_5)
                    .Update
                End With
                TmpRst.MoveNext
            Loop
            If TmpRst.RecordCount > 0 Then: TmpRst.MoveFirst
            
            Do While TmpRst.EOF = False
                Set TmpRst1 = GCn.Execute("Select Qty_Iss-Qty_Ret as NetQty,Tax_Yn,Part.TB_SRate,Sp_Stock.V_Rate from Sp_Stock left Join Part on Part.Part_No=Sp_Stock.Part_No where Sp_Stock.Purpose='W' and Sp_Stock.Job_DocId='" & TmpRst!JobDocID & "'")
                
                TotAmtTb = 0: TotAmtTp = 0
                
                If TmpRst1.RecordCount > 0 And TmpRst!JobDocID <> "" Then
                    Do While TmpRst1.EOF = False
                        If TmpRst1!Tax_YN = 1 Then
                            'TotAmtTb = TotAmtTb + Round((Val(TmpRst1!NetQty) * (TmpRst1!TB_SRate)), 0)
                            TotAmtTb = TotAmtTb + Round((VNull(TmpRst1!NetQty) * VNull(TmpRst1!V_Rate)), 0)
                        Else
                            'TotAmtTp = TotAmtTp + Round((Val(TmpRst1!NetQty) * (TmpRst1!TB_SRate)), 0)
                            TotAmtTp = TotAmtTp + Round((VNull(TmpRst1!NetQty) * VNull(TmpRst1!V_Rate)), 0)
                        End If
                        TmpRst1.MoveNext
                    Loop
                        JDocId = TmpRst!JobDocID
                        InvdocId = TmpRst!InvdocId
                        VDt = TmpRst!V_DATE
                        NetVal = TotAmtTb + TotAmtTp
                        Clodate = TmpRst!JobCloseDate
                    
                    With RstRep
                    .AddNew
                        !JobDocID = XNull(JDocId)
                        !JobCloseDate = XNull(Clodate)
                        !InvdocId = XNull(TmpRst!InvdocId)
                        !V_DATE = XNull(TmpRst!V_DATE)
                        !SprAmtTB = TotAmtTb
                        !SprAmtTP = TotAmtTp
                        !NetLab_Amt = 0
                        !Misc = 0
                        !Tot_Amt = 0
                        !Total_Amt = TotAmtTb + TotAmtTp
                        !SaleType = "Warranty"
                        !Job_No = VNull(TmpRst!Job_No)
                        !SprAmt_MRP_TB = TotAmtTb
                        !CName = XNull(TmpRst!PartyName)
                        
                    .Update
                End With
                        
                End If
                TmpRst.MoveNext
            Loop
            
        End If
    End If
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If GRepFormName = WarTaxReimbReg Then
        RepTitle = UCase(Me.CAPTION)
    ElseIf (GRepFormName = SprPurRet Or GRepFormName = SprPurReg) Then
        RepTitle = UCase(Me.CAPTION)
    Else
        RepTitle = UCase(Me.CAPTION) + "[" + FGrid.TextMatrix(List1, 1) + "]"
    End If
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprStkRep()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim RateCond$
'Date1,Date2,List1,List1,List1,List2,List1,List1
Condstr = ""
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub

If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If FGrid.TextMatrix(List1, 1) = "Yes" Then GridString2 = MarkRecCalculate 'Else MsgBox "No"
If GridString2 = Empty And FGrid.TextMatrix(List1, 1) = "Yes" Then: MsgBox "** No Records Found to Print **": RepPrint = False: Exit Sub: Else If FGrid.TextMatrix(List1, 1) = "Yes" Then Condstr = Condstr & " and SP_Stock.Part_No in (" & GridString2 & ")"  'Else MsgBox "No"
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub

If GRepFormName = SprStkInHand Then
    If Check1(1).Value = Unchecked Then Condstr = " and Part.Bin_Loca in (" & GridString1 & ")"
    If FGrid.TextMatrix(List3, 1) <> "" Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " = '" & FGrid.TextMatrix(List3, 2) & "'"
Else
    If Check1(1).Value = Unchecked Then Condstr = " and left(SP_Stock.site_code,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(sp_stock.site_code,1) ='" & PubSiteCode & "' "
End If

End If
'If FGrid.TextMatrix(List3, 1) <> "" Then Condstr = Condstr & " and SP_Stock.Part_No in (" & GridString2 & ")"

If Check1(2).Value = Unchecked Then Condstr = Condstr & " and SP_Stock.Part_No in (" & GridString2 & ")"
If StrCmp(left(PubComp_Name, 4), "Yash") Then
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and SP_Stock.Godown in (" & GridString3 & ")"
Else
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(SP_Stock.DocId,1) in (" & GridString3 & ")"
End If
If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Part.Part_Grade in (" & GridString4 & ")"

Select Case GRepFormName
    Case SprStkReg
        If ChkOpeningStockOnly.Value = 1 Then Condstr = Condstr & " And Sp_stock.V_Type = 'SXAO' "
        mQry = "SELECT  SP_Stock.Part_No,Part.Part_Name , 'Opening' As DocID,'0' as SrlNo, '" & DateAdd("d", -1, CDate(FGrid.TextMatrix(Date1, 1))) & "' as V_Date, " & cIIF("SP_Stock.Tax_YN=0", "'Op.TP'", "'Op.TB'") & " as v_Prefix, 0 As V_No, ' '  As Job_DocID, " & _
        " Sum(" & cIIF("SP_Stock.Tax_YN=0", "SP_Stock.Qty_Rec-Sp_Stock.Qty_Iss+Sp_Stock.Qty_Ret", "0") & ") AS TPQtyRec, " & _
        " Sum(" & cIIF("SP_Stock.Tax_YN=1", "SP_Stock.Qty_Rec-Sp_Stock.Qty_Iss+Sp_Stock.Qty_Ret", "0") & ") AS TBQtyRec, " & _
        " 0 AS TPQtyIss,0 AS TBQtyIss,'' AS SprPurPose, '' As Party_Doc_No, Null As Party_Doc_Date,' ' As Bin_Loca,0 as v_Rate, 'SXAO' As V_type, SP_Stock.MRP_YN " & _
        " FROM ((SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and  left(SP_Stock.DocId,1) = Part.Div_Code) LEFT JOIN SP_Purch on SP_Purch.DocID=SP_Stock.DocId) " & _
        " WHERE SP_Stock.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(DateAdd("D", -1, FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  And (SP_Stock.V_Type = " & cIIF("SP_Stock.V_Date = " & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or  SP_Stock.V_Type <> " & cIIF("SP_Stock.V_Date > " & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & ") " & _
        Condstr & " Group By Sp_Stock.Part_No, Part.Part_Name, SP_Stock.Mrp_Yn, Sp_Stock.Tax_Yn "
        
        If PubBackEnd = "A" Then
            mQry = mQry & " Union All " & _
            "SELECT  SP_Stock.Part_No,Part.Part_Name, SP_Stock.DocID,'Z' as SrlNo, SP_Stock.V_Date, " & cMID("SP_Stock.DocID", "9", "5") & " as v_Prefix, SP_Stock.V_No, SP_Stock.Job_DocID, " & _
            " " & cIIF("SP_Stock.Tax_YN=0", "SP_Stock.Qty_Rec", "0") & " AS TPQtyRec, " & _
            " " & cIIF("SP_Stock.Tax_YN=1", "SP_Stock.Qty_Rec", "0") & " AS TBQtyRec, " & _
            " " & cIIF("SP_Stock.Tax_YN=0", "(SP_Stock.Qty_Iss-Sp_Stock.Qty_Ret)", "0") & " AS TPQtyIss," & _
            " " & cIIF("SP_Stock.Tax_YN=1", "(SP_Stock.Qty_Iss-Sp_Stock.Qty_Ret)", "0") & " AS TBQtyIss, " & _
            " " & cIIF("SP_Stock.Purpose='P'", "'PDI'", cIIF("SP_Stock.Purpose='F'", "'FreeSer'", cIIF("SP_Stock.Purpose='C'", "'Chargeble'", cIIF("SP_Stock.Purpose='W'", "'Warrant'", cIIF("SP_Stock.Purpose='O'", "'CompVeh'", cIIF("SP_Stock.Purpose='L'", "'Complementory'", cIIF("SP_Stock.Purpose='A'", "'AMC'", "''"))))))) & " AS SprPurPose,SP_Purch.Party_Doc_No,SP_Purch.Party_Doc_Date,Part.Bin_Loca, " & cIIF("Sp_stock.qty_Rec > 0", "Sp_stock.Rate", "0") & " as v_rate,SP_Stock.V_type,SP_Stock.MRP_YN " & _
            " FROM ((SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and left(SP_Stock.DocId,1) = Part.Div_Code) LEFT JOIN SP_Purch on SP_Purch.DocID=SP_Stock.DocId) " & _
            " where SP_Stock.V_Type<>'SXAO' and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & _
            Condstr & _
            " Order By V_Date, V_No, V_Type asc"
        Else
            mQry = mQry & " Union All " & _
            "SELECT  SP_Stock.Part_No,Part.Part_Name, SP_Stock.DocID,'Z' as SrlNo, SP_Stock.V_Date, " & cMID("SP_Stock.DocID", "9", "5") & " as v_Prefix, SP_Stock.V_No, SP_Stock.Job_DocID, " & _
            " " & cIIF("SP_Stock.Tax_YN=0", "SP_Stock.Qty_Rec", "0") & " AS TPQtyRec, " & _
            " " & cIIF("SP_Stock.Tax_YN=1", "SP_Stock.Qty_Rec", "0") & " AS TBQtyRec, " & _
            " " & cIIF("SP_Stock.Tax_YN=0", "(SP_Stock.Qty_Iss-Sp_Stock.Qty_Ret)", "0") & " AS TPQtyIss," & _
            " " & cIIF("SP_Stock.Tax_YN=1", "(SP_Stock.Qty_Iss-Sp_Stock.Qty_Ret)", "0") & " AS TBQtyIss, " & _
            " " & cIIF("SP_Stock.Purpose='P'", "'PDI'", cIIF("SP_Stock.Purpose='F'", "'FreeSer'", cIIF("SP_Stock.Purpose='C'", "'Chargeble'", cIIF("SP_Stock.Purpose='W'", "'Warrant'", cIIF("SP_Stock.Purpose='O'", "'CompVeh'", cIIF("SP_Stock.Purpose=' L'", "'Complem'", cIIF("SP_Stock.Purpose='A'", "'AMC'", "''"))))))) & " AS SprPurPose,SP_Purch.Party_Doc_No,SP_Purch.Party_Doc_Date,Part.Bin_Loca," & cIIF("Sp_stock.qty_Rec > 0", "Sp_stock.Rate", "0") & " as v_rate,SP_Stock.V_type,SP_Stock.MRP_YN " & _
            " FROM ((SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and left(SP_Stock.DocId,1) = Part.Div_Code) LEFT JOIN SP_Purch on SP_Purch.DocID=SP_Stock.DocId) " & _
            " where SP_Stock.V_Type<>'SXAO' and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & _
            Condstr & _
            " order by V_Date"
        End If
        
        RepName = "SprStkReg"
    Case SprStkSumm
        If ChkOpeningStockOnly.Value = 1 Then Condstr = Condstr & " And Sp_stock.V_Type = 'SXAO' "
        mQry = "SELECT SP_Stock.Part_No, 0 AS TPQtyRec, 0 AS TBQtyRec, 0 AS TPQtyIss, 0 AS TBQtyIss, " & cIIF("SP_Stock.Tax_YN=0", "Sum(SP_Stock.Qty_Rec-Sp_Stock.Qty_Iss+Sp_Stock.Qty_Ret)", "0") & " AS TPQtyOpen, " & cIIF("SP_Stock.Tax_YN=1", "Sum(SP_Stock.Qty_Rec-Sp_Stock.Qty_Iss+Sp_Stock.Qty_Ret)", "0") & " AS TBQtyOpen, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and left(SP_Stock.DocId,1) = Part.Div_Code " & _
        "Where SP_Stock.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(DateAdd("D", -1, FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  And SP_Stock.V_Type = " & cIIF("SP_Stock.V_Date = " & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & "  " & _
        Condstr & _
        " GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name " & _
        " Union All " & _
        "SELECT SP_Stock.Part_No, " & cIIF("SP_Stock.Tax_YN=0", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TPQtyRec, " & cIIF("SP_Stock.Tax_YN=1", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TBQtyRec, " & cIIF("SP_Stock.Tax_YN=0", "Sum(SP_Stock.Qty_Iss-SP_Stock.Qty_Ret)", "0") & " AS TPQtyIss, " & cIIF("SP_Stock.Tax_YN=1", "Sum(SP_Stock.Qty_Iss-SP_Stock.Qty_Ret)", "0") & " AS TBQtyIss, 0 AS TPQtyOpen, 0 AS TBQtyOpen, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and left(SP_Stock.DocId,1) = Part.Div_Code " & _
        "WHERE SP_Stock.V_Type<>'SXAO' and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & _
        Condstr & _
        "GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name"
        
        RepName = "SprStkSumm"
        
    Case SprStkInHand
        If ChkOpeningStockOnly.Value = 1 Then Condstr = Condstr & " And Sp_stock.V_Type = 'SXAO' "
        If FGrid.TextMatrix(List2, 1) = "MRP" Then
            RateCond = "" & cIIF("SP_Stock.Tax_YN=0", "Max(Part.MRP)", "0") & " as TP_SRate, " & cIIF("SP_Stock.Tax_YN=1", "Max(Part.MRP)", "0") & " as TB_SRate"
            
        ElseIf FGrid.TextMatrix(List2, 1) = "NDP" Then
            RateCond = "" & cIIF("SP_Stock.Tax_YN=0", "Max(Part.NDP)", "0") & " as TP_SRate, " & cIIF("SP_Stock.Tax_YN=1", "Max(Part.NDP)", "0") & " as TB_SRate"
            
        ElseIf FGrid.TextMatrix(List2, 1) = "Unit Price" Then
            RateCond = "" & cIIF("SP_Stock.Tax_YN=0", "Max(Part.TB_SRate)", "0") & " as TP_SRate, " & cIIF("SP_Stock.Tax_YN=1", "Max(Part.TB_SRate)", "0") & " as TB_SRate"
            
        ElseIf FGrid.TextMatrix(List2, 1) = "Warr.Price" Then
            RateCond = "" & IIf("SP_Stock.Tax_YN=0", "Max(Part.WarrRate)", "0") & " as TP_SRate," & cIIF("SP_Stock.Tax_YN=1", "Max(Part.WarrRate)", "0") & " as TB_SRate"
            
        ElseIf FGrid.TextMatrix(List2, 1) = "Sale Rate" Then
            RateCond = "" & cIIF("SP_Stock.Tax_YN=0", "Max(SP_Stock.Rate)", "0") & " as TP_SRate," & cIIF("SP_Stock.Tax_YN=1", "Max(SP_Stock.Rate)", "0") & " as TB_SRate"
        Else
            RateCond = "Max(Part.PurRate) as TP_SRate, Max(Part.PurRate) As TB_SRate "
        End If
    
        
                
        mQry = "SELECT " & cTrim("Sp_Stock.Part_No") & " , Max(Part.Part_Name) as Part_Name, " & cIIF("Sp_Stock.Tax_YN=0", "Sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret)", "0") & " As Qty_TP, " & _
                 " " & cIIF("Sp_Stock.Tax_YN=1", "Sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret)", "0") & " As Qty_TB, " & RateCond & " " & _
                 "From ((SP_Stock  " & _
            "Left Join Part  On Part.Part_No = Sp_Stock.Part_No) " & _
            "Left Join Sp_Purch On SP_Stock.DocId = SP_Purch.DocId) " & _
            "Where SP_Stock.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & " and SP_Stock.V_Date<= " & ConvertDate(Format(DateAdd("D", -1, FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & " And SP_Stock.V_Type = " & cIIF("SP_Stock.V_Date = " & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & "  " & _
            Condstr & _
            " Group By Sp_Stock.Part_No, Sp_Stock.Tax_YN,SP_Stock.Rate " & _
            " Union All " & _
            "SELECT " & cTrim("Sp_Stock.Part_No") & " , Max(Part.Part_Name) as Part_Name, " & cIIF("Sp_Stock.Tax_YN=0", "Sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret)", "0") & " As Qty_TP, " & _
                 "" & cIIF("Sp_Stock.Tax_YN=1", "Sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret)", "0") & " As Qty_TB, " & RateCond & " " & _
            "From ((SP_Stock  " & _
            "Left Join Part  On Part.Part_No = Sp_Stock.Part_No) " & _
            "Left Join Sp_Purch On SP_Stock.DocId = SP_Purch.DocId) " & _
            " where SP_Stock.V_Type<>'SXAO' and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & _
            Condstr & _
            "Group By Sp_Stock.Part_No, Sp_Stock.Tax_YN "


        RepName = "SprStkInHand"
    End Select
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    ' Speed printing of stock reports
    If SpeedPrn = True Then
        Select Case GRepFormName
            Case SprStkInHand
                SpeedPrnStkInHnd
                Exit Sub
            Case SprStkReg
                SpeedPrnStkLedger
                Exit Sub
        End Select
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ProcStockValue()
        On Error GoTo ELoop
        Dim mQry As String, Condstr As String
        Dim RateCond$
        'Date1,Date2,List1,List1,List1,List2,List1,List1
        Condstr = ""
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        
        
        
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
        If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
        If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
        
        
        If Check1(1).Value = Unchecked Then Condstr = " and " & cMID("S.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("s.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and S.Part_No in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(S.DocId,1) in (" & GridString3 & ")"
        If Check1(4).Value = Unchecked Then Condstr = Condstr & " and P.Part_Grade in (" & GridString4 & ")"



        mQry = "Select X.Part_No, MAX(X.NDP) AS NDP, SUM(X.Opening) AS Opening, SUM(X.Opening)*MAX(X.NDP) AS OpeningValue , SUM(X.Qty_Rec) AS Qty_Rec, SUM(X.Qty_Rec)*MAX(X.NDP) AS Qty_RecValue,  SUM(X.Qty_Iss) AS Qty_Iss,  SUM(X.Qty_Iss)*MAX(X.NDP) AS Qty_IssValue, SUM(X.Opening) + SUM(X.Qty_Rec) - SUM(X.Qty_Iss) AS Balance, (SUM(X.Opening) + SUM(X.Qty_Rec) - SUM(X.Qty_Iss))*MAX(X.NDP) AS BalanceValue  " & _
               " From " & _
               " ( " & _
               " SELECT S.Part_No, MAX(P.NDP) AS NDP, SUM(S.Qty_Rec-S.Qty_Iss+S.Qty_Ret) AS Opening, 0 AS Qty_Rec, 0 AS Qty_Iss  FROM SP_Stock S LEFT JOIN Part P ON S.Part_No = P.PART_NO  WHERE S.V_Date >= '" & DateAdd("D", -1, CDate(PubStartDate)) & "' And S.V_Date < " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " AND (V_Type = (CASE WHEN V_Date = '" & DateAdd("D", -1, CDate(PubStartDate)) & "' THEN 'SXAO' End) Or V_Type <> (CASE WHEN V_Date <> '" & DateAdd("D", -1, CDate(PubStartDate)) & "' THEN 'SXAO' End)) " & Condstr & " GROUP BY S.Part_No " & _
               " Union All " & _
               " SELECT S.Part_No, MAX(P.NDP) AS NDP, 0 AS Opening, SUM(S.Qty_Rec) AS Qty_Rec, SUM(S.Qty_Iss-S.Qty_Ret)AS Qty_Iss FROM SP_Stock S LEFT JOIN Part P ON S.Part_No = P.PART_NO  WHERE S.V_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And S.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & Condstr & "   GROUP BY s.Part_No " & _
               " ) AS X " & _
               " GROUP BY X.Part_No "


        
        RepName = "StockValue"

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    ' Speed printing of stock reports
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub VehMoneyRectFunc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(DGrid1, FGrid.TextMatrix(DGrid1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where R.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and R.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    If FGrid.TextMatrix(List1, 1) <> "All" Then Condstr = Condstr & " and R.IFORM = '" & FGrid.TextMatrix(List1, 1) & "'"
    If FGrid.TextMatrix(List2, 1) = "Taxable" Then Condstr = Condstr & " and R.Tax_Amt <> 0 and Surcharge_Amt <> 0 and TOT_Amt <> 0"
    If FGrid.TextMatrix(List2, 1) = "TaxPaid" Then Condstr = Condstr & " and R.Tax_Amt = 0 and Surcharge_Amt = 0 and TOT_Amt = 0"
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(R.site_code,1) in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(r.site_code,1) ='" & PubSiteCode & "' "
    End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and R.Prov_Location in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and left(R.Docid,1) in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and R.v_type in (" & GridString4 & ")"
    
    mQry = "SELECT SubGroup.Name, R.DocId, R.V_Date,R.Narration,R.v_type, R.V_No, R.Prov_No, R.Prov_Date, R.AMOUNT, R.DDNo, R.DDDate, R.IFORM,R.Veh_Amt,R.Tax_Amt,R.Surcharge_Amt,R.TOT_Amt,SubGroup.PANNo,cITY.CITYNAME  " & _
    " FROM (Rect AS R " & _
    " LEFT JOIN SubGroup ON R.PartyCode = SubGroup.SubCode) " & _
    " LEFT JOIN City ON R.Prov_Location = City.CityCode "
    
    If FGrid.TextMatrix(List4, 1) <> "All" Then
        Condstr = Condstr & " and R.RectCatG='" & FGrid.TextMatrix(List4, 1) & "'"
    End If

    mQry = mQry + Condstr + " order by R.V_DATE,R.DOCID"
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List3, 1) = "Summery" Then
        RepName = "VehMoneyRectSumm"
    ElseIf FGrid.TextMatrix(List3, 1) = "Detail" Then
        RepName = "VehMoneyRect"
    End If
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SprPartPurchase()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim PartyType As Byte

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    Condstr = "Where Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Stock.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    Condstr = Condstr & " and SP_Stock.V_type = '" & SprMrRct & "'"
    
    PartyType = GCn.Execute("select PartyType from SYCTRL").Fields(0).Value

    mQry = "SELECT Part.Part_Name, SP_Stock.Part_No, " & cIIF("SubGroup.Party_Type=" & PartyType & "", "Sum(sp_Stock.Qty_Rec)", "0") & " AS telcoqty, " & cIIF("SubGroup.Party_Type<>" & PartyType & "", "Sum(sp_Stock.Qty_Rec)", "0") & " AS localqty, " & cIIF("SubGroup.Party_Type=" & PartyType & "", "Sum(SP_Stock.Amount)", "0") & " AS telcoamt, " & cIIF("SubGroup.Party_Type<>" & PartyType & "", "Sum(SP_Stock.Amount)", "0") & " AS localamt " & _
    "FROM (SP_Stock LEFT JOIN SubGroup ON SP_Stock.Party_Code = SubGroup.SubCode) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) "
        
    mQry = mQry + Condstr
    
    mQry = mQry + " GROUP BY Part.Part_Name, SP_Stock.Part_No, SubGroup.Party_Type"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprPartPur"
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


Private Sub SprPurSalSum()
On Error GoTo ELoop
Dim mQry As String, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    If GRepFormName = SprMRPTaxClaimReg Then
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Sale.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("sp_sale.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Sale.DocId,1) in (" & GridString2 & ")"
        If PubVATYN = 1 Then
            If PubBackEnd = "A" Then
                mQry = "SELECT " & _
                    "switch(" & cMID("max(SP_Sale.DocID)", "4", "5") & "='" & SprSlCsh & "',' Sale-Coun'," & _
                    "" & cMID("max(SP_Sale.DocID)", "4", "5") & "='" & SprSlCre & "',' Sale-Coun'," & _
                    "" & cMID("max(SP_Sale.DocID)", "4", "5") & "='" & WksSlCsh & "',' Sale-Work'," & _
                    "" & cMID("max(SP_Sale.DocID)", "4", "5") & "='" & WksSlCre & "',' Sale-Work'," & _
                    "" & cMID("max(SP_Sale.DocID)", "4", "5") & "='" & SprSlRetCsh & "','Return    '," & _
                    "" & cMID("max(SP_Sale.DocID)", "4", "5") & "='" & SprSlRetCre & "','Return    ') as TrnType," & _
                    "Max(SP_Sale.DocID) as DocId,Max(SP_Sale.V_Date) as V_Date,Max(SP_Sale.Party_Name) as Party_Name," & _
                    "(Max(SP_Sale.SprAmt_MRP_TB)+Max(SP_Sale.OilAmt_MRP_TB)-Max(SP_Sale.D_Amt_MRP_TB)) AS TBMRPAmt," & _
                    "sum(SP_Stock.TaxAmt) as Tax_AmtMRP,0.00 as TaxSur_AmtMrp, " & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "Max(SP_Sale.TOT_Amt)", "Max(SP_Sale.TOT_AmtMrp)") & " as TOT_AmtMRP,(Max(SP_Sale.Tax_AmtMRP)+Max(SP_Sale.TaxSur_AmtMRP)+max(SP_Sale.TOT_Amt)) as TaxSurAmt," & _
                    "Max(SP_Sale.Total_Amt) as Total_Amt,Max(TaxForms.Form_Code) as Form_Code,Max(TaxForms.Form_Desc) as Form_Desc,Max(TaxForms.L_C) as L_C,Max(TaxForms.Tax_Per) as Tax_Per,Max(TaxForms.Tax_Sur_Per) as Tax_Sur_Per" & _
                    " FROM (SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code)" & _
                    " Left Join SP_Stock On SP_Stock.Invoice_DocId=SP_Sale.DocId " & _
                    " Where  SP_Sale.Total_Amt<>0 and SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "')  And SP_Stock.Tax_YN=1 and SP_Stock.MRP_YN=1 " & Condstr & " Group By SP_Stock.INVoice_DocId "
            ElseIf PubBackEnd = "S" Then
                mQry = "SELECT " & _
                    " Case " & cMID("max(SP_Sale.DocID)", "4", "5") & " When '" & SprSlCsh & "' Then ' Sale-Coun' " & _
                    " When '" & SprSlCre & "' Then ' Sale-Coun' " & _
                    " When '" & WksSlCsh & "' Then ' Sale-Work' " & _
                    " When '" & WksSlCre & "' Then ' Sale-Work' " & _
                    " When '" & SprSlRetCsh & "' Then 'Return    ' " & _
                    " When '" & SprSlRetCre & "' Then 'Return    ' End As TrnType," & _
                    " Max(SP_Sale.DocID) as DocId,Max(SP_Sale.V_Date) as V_Date,Max(SP_Sale.Party_Name) as Party_Name," & _
                    " (Max(SP_Sale.SprAmt_MRP_TB)+Max(SP_Sale.OilAmt_MRP_TB)-Max(SP_Sale.D_Amt_MRP_TB)) AS TBMRPAmt," & _
                    " sum(SP_Stock.TaxAmt) as Tax_AmtMRP,0.00 as TaxSur_AmtMrp, " & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "Max(SP_Sale.TOT_Amt)", "Max(SP_Sale.TOT_AmtMrp)") & " as TOT_AmtMRP,(Max(SP_Sale.Tax_AmtMRP)+Max(SP_Sale.TaxSur_AmtMRP)+max(SP_Sale.TOT_Amt)) as TaxSurAmt," & _
                    " Max(SP_Sale.Total_Amt) as Total_Amt,Max(TaxForms.Form_Code) as Form_Code,Max(TaxForms.Form_Desc) as Form_Desc,Max(TaxForms.L_C) as L_C,Max(TaxForms.Tax_Per) as Tax_Per,Max(TaxForms.Tax_Sur_Per) as Tax_Sur_Per" & _
                    " FROM (SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code)" & _
                    " Left Join SP_Stock On SP_Stock.Invoice_DocId=SP_Sale.DocId " & _
                    " Where  SP_Sale.Total_Amt<>0 and SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "')  And SP_Stock.Tax_YN=1 and SP_Stock.MRP_YN=1 " & Condstr & " Group By SP_Stock.INVoice_DocId "
            End If
        Else
            If PubBackEnd = "A" Then
                mQry = "SELECT " & _
                    "switch(" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCsh & "',' Sale-Coun'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCre & "',' Sale-Coun'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCsh & "',' Sale-Work'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCre & "',' Sale-Work'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCsh & "','Return    '," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCre & "','Return    ') as TrnType," & _
                    "SP_Sale.DocID,SP_Sale.V_Date,SP_Sale.Party_Name," & _
                    "(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB-SP_Sale.D_Amt_MRP_TB) AS TBMRPAmt," & _
                    "SP_Sale.Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP," & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "SP_Sale.TOT_Amt", "SP_Sale.TOT_AmtMrp") & " as TOT_AmtMRP,(SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_Amt) as TaxSurAmt," & _
                    "SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per" & _
                    " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                    " Where  SP_Sale.Total_Amt<>0 and SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "') " & Condstr
            ElseIf PubBackEnd = "S" Then
                mQry = "SELECT " & _
                    " Case " & cMID("SP_Sale.DocID", "4", "5") & " When '" & SprSlCsh & "' Then ' Sale-Coun' " & _
                    " When '" & SprSlCre & "' Then ' Sale-Coun' " & _
                    " When '" & WksSlCsh & "' Then ' Sale-Work' " & _
                    " When '" & WksSlCre & "' Then ' Sale-Work' " & _
                    " When '" & SprSlRetCsh & "' Then 'Return    ' " & _
                    " When '" & SprSlRetCre & "' Then 'Return    ' End as TrnType," & _
                    "SP_Sale.DocID,SP_Sale.V_Date,SP_Sale.Party_Name," & _
                    "(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB-SP_Sale.D_Amt_MRP_TB) AS TBMRPAmt," & _
                    "SP_Sale.Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP," & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "SP_Sale.TOT_Amt", "SP_Sale.TOT_AmtMrp") & " as TOT_AmtMRP,(SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+SP_Sale.TOT_Amt) as TaxSurAmt," & _
                    "SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per" & _
                    " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                    " Where  SP_Sale.Total_Amt<>0 and SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "') " & Condstr
            End If
        End If
        RepName = "SprMRPTax"
        
    ElseIf GRepFormName = SprSaleSum Then
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Sale.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_Sale.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Sale.DocId,1) in (" & GridString2 & ")"
        If PubBackEnd = "A" Then
            mQry = "SELECT " & _
                "switch(" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCsh & "',' Sale-Coun'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCre & "',' Sale-Coun'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCsh & "',' Sale-Work'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCre & "',' Sale-Work'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCsh & "','Return    '," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCre & "','Return    ') as TrnType," & _
                "SP_Sale.DocID,SP_Sale.V_Date," & _
                "(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB+SP_Sale.SprAmt_TB+ SP_Sale.OilAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt-(SP_Sale.D_Amt_TB+SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP)) AS TBAmt," & _
                "(SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP-(SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TP)) AS TPamt," & _
                "SP_Sale.Packing,SP_Sale.TOT_Amt,SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per,SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP as TaxAmt,SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP as TaxSurAmt" & _
                " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                " Where  SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "') " & Condstr
        ElseIf PubBackEnd = "S" Then
            mQry = "SELECT " & _
                " Case " & cMID("SP_Sale.DocID", "4", "5") & " When '" & SprSlCsh & "' Then ' Sale-Coun' " & _
                " When '" & SprSlCre & "' Then ' Sale-Coun' " & _
                " When '" & WksSlCsh & "' Then ' Sale-Work' " & _
                " When '" & WksSlCre & "' Then ' Sale-Work' " & _
                " When '" & SprSlRetCsh & "' Then 'Return    ' " & _
                " When '" & SprSlRetCre & "' Then 'Return    ' End As TrnType," & _
                "SP_Sale.DocID,SP_Sale.V_Date," & _
                "(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB+SP_Sale.SprAmt_TB+ SP_Sale.OilAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt-(SP_Sale.D_Amt_TB+SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP)) AS TBAmt," & _
                "(SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP-(SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TP)) AS TPamt," & _
                "SP_Sale.Packing,SP_Sale.TOT_Amt,SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per,SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP as TaxAmt,SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP as TaxSurAmt" & _
                " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                " Where  SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "') " & Condstr
        End If
    If PubDiscOnLube = 1 Then
        If PubBackEnd = "A" Then
            mQry = "SELECT switch(" & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCsh & "',' Sale-Coun'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCre & "',' Sale-Coun'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCsh & "',' Sale-Work'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCre & "',' Sale-Work'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCsh & "','Return    '," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCre & "','Return    ') as TrnType," & _
                "SP_Sale.DocID,SP_Sale.V_Date," & _
                "" & cIIF("(SP_Sale.SprAmt_TB+ SP_Sale.OilAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt)=0", "0", "((SP_Sale.SprAmt_TB+ SP_Sale.OilAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt)-(SP_Sale.D_Amt_TB))") & " AS TBAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB)- (SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP + SP_Sale.D_Amt_MRP_TB))") & " AS TBMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)=0", "0", "(SP_Sale.SprAmt_MRP_TP - ((SP_Sale.SprAmt_MRP_TP * SP_Sale.D_Amt_MRP_TP) / (SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)))") & " AS TPSprMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)=0", "0", "(SP_Sale.OilAmt_MRP_TP - ((SP_Sale.OilAmt_MRP_TP * SP_Sale.D_Amt_MRP_TP) / (SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)))") & " AS TPOilMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)=0", "0", "(SP_Sale.SprAmt_TP - ((SP_Sale.SprAmt_TP * (SP_Sale.D_Amt_TP - D_Amt_MRP_TP)) / (SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)))") & " AS TPSprAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)=0", "0", "(SP_Sale.OilAmt_TP - ((SP_Sale.OilAmt_TP * (SP_Sale.D_Amt_TP - D_Amt_MRP_TP)) / (SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)))") & " AS TPOilAmt," & _
                "SP_Sale.Packing,SP_Sale.TOT_Amt,SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per,SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP as TaxAmt,SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP as TaxSurAmt,(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP) as DiscAmt" & _
                " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                " Where  SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "')  " & Condstr & " order by SP_Sale.V_Date"
        ElseIf PubBackEnd = "S" Then
            mQry = "SELECT  Case " & cMID("SP_Sale.DocID", "4", "5") & " " & _
                " When '" & SprSlCsh & "' Then ' Sale-Coun' " & _
                " When '" & SprSlCre & "' Then ' Sale-Coun' " & _
                " When '" & WksSlCsh & "' Then ' Sale-Work' " & _
                " When '" & WksSlCre & "' Then ' Sale-Work' " & _
                " When '" & SprSlRetCsh & "' Then 'Return    ' " & _
                " When '" & SprSlRetCre & "' Then 'Return    ' End As TrnType," & _
                "SP_Sale.DocID,SP_Sale.V_Date," & _
                "" & cIIF("(SP_Sale.SprAmt_TB+ SP_Sale.OilAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt)=0", "0", "((SP_Sale.SprAmt_TB+ SP_Sale.OilAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt)-(SP_Sale.D_Amt_TB))") & " AS TBAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB+SP_Sale.OilAmt_MRP_TB)- (SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP + SP_Sale.D_Amt_MRP_TB))") & " AS TBMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)=0", "0", "(SP_Sale.SprAmt_MRP_TP - ((SP_Sale.SprAmt_MRP_TP * SP_Sale.D_Amt_MRP_TP) / (SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)))") & " AS TPSprMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)=0", "0", "(SP_Sale.OilAmt_MRP_TP - ((SP_Sale.OilAmt_MRP_TP * SP_Sale.D_Amt_MRP_TP) / (SP_Sale.SprAmt_MRP_TP+ SP_Sale.OilAmt_MRP_TP)))") & " AS TPOilMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)=0", "0", "(SP_Sale.SprAmt_TP - ((SP_Sale.SprAmt_TP * (SP_Sale.D_Amt_TP - D_Amt_MRP_TP)) / (SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)))") & " AS TPSprAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)=0", "0", "(SP_Sale.OilAmt_TP - ((SP_Sale.OilAmt_TP * (SP_Sale.D_Amt_TP - D_Amt_MRP_TP)) / (SP_Sale.SprAmt_TP+ SP_Sale.OilAmt_TP)))") & " AS TPOilAmt," & _
                "SP_Sale.Packing,SP_Sale.TOT_Amt,SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per,SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP as TaxAmt,SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP as TaxSurAmt,(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP) as DiscAmt" & _
                " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                " Where  SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "')  " & Condstr & " order by SP_Sale.V_Date"
        End If
    Else
        If UCase(left(PubComp_Name, 3)) = "JMK" Then
            mQry = "SELECT switch(" & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCsh & "',' Sale-Coun'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCre & "',' Sale-Coun'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCsh & "',' Sale-Work'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCre & "',' Sale-Work'," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCsh & "','Return    '," & _
                "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCre & "','Return    ') as TrnType," & _
                "SP_Sale.DocID,SP_Sale.V_Date," & _
                "" & cIIF("(SP_Sale.SprAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt)=0", "0", "((SP_Sale.SprAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) - (TOT_Amt+SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB))") & " AS TBAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB) - (SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+SP_Sale.D_Amt_MRP_TB))") & " AS TBMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP)=0", "0", "((SP_Sale.SprAmt_MRP_TP) - (SP_Sale.D_Amt_MRP_TP))") & " AS TPSprMRPAmt," & _
                "SP_Sale.OilAmt_MRP_TP AS TPOilMRPAmt," & _
                "" & cIIF("(SP_Sale.SprAmt_TP)=0", "0", "((SP_Sale.SprAmt_TP) - (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP))") & " AS TPSprAmt," & _
                "SP_Sale.OilAmt_TP AS TPOilAmt," & _
                "SP_Sale.Packing,SP_Sale.TOT_Amt,SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per,SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP as TaxAmt,SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP as TaxSurAmt,(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP) as DiscAmt" & _
                " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                " Where  SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "') " & Condstr & " order by SP_Sale.V_Date"
        Else
            If PubBackEnd = "A" Then
                mQry = "SELECT switch(" & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCsh & "',' Sale-Coun'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlCre & "',' Sale-Coun'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCsh & "',' Sale-Work'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & WksSlCre & "',' Sale-Work'," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCsh & "','Return    '," & _
                    "" & cMID("SP_Sale.DocID", "4", "5") & "='" & SprSlRetCre & "','Return    ') as TrnType," & _
                    "SP_Sale.DocID,SP_Sale.V_Date," & _
                    "" & cIIF("(SP_Sale.SprAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt)=0", "0", "((SP_Sale.SprAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) - (TOT_Amt+SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB))") & " AS TBAmt," & _
                    "" & cIIF("(SP_Sale.SprAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB) - (SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+TOT_AmtMRP+SP_Sale.D_Amt_MRP_TB))") & " AS TBMRPAmt," & _
                    "" & cIIF("(SP_Sale.SprAmt_MRP_TP)=0", "0", "((SP_Sale.SprAmt_MRP_TP) - (SP_Sale.D_Amt_MRP_TP))") & " AS TPSprMRPAmt," & _
                    "SP_Sale.OilAmt_MRP_TP AS TPOilMRPAmt," & _
                    "" & cIIF("(SP_Sale.SprAmt_TP)=0", "0", "((SP_Sale.SprAmt_TP) - (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP))") & " AS TPSprAmt," & _
                    "SP_Sale.OilAmt_TP AS TPOilAmt," & _
                    "SP_Sale.Packing,SP_Sale.TOT_Amt,SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per,SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP as TaxAmt,SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP as TaxSurAmt,(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP) as DiscAmt" & _
                    " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                    " Where  SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "') " & Condstr & " order by SP_Sale.V_Date"
            ElseIf PubBackEnd = "S" Then
                mQry = "SELECT Case " & cMID("SP_Sale.DocID", "4", "5") & " " & _
                    " When '" & SprSlCsh & "' Then ' Sale-Coun' " & _
                    " When '" & SprSlCre & "' Then ' Sale-Coun' " & _
                    " When '" & WksSlCsh & "' Then ' Sale-Work' " & _
                    " When '" & WksSlCre & "' Then ' Sale-Work' " & _
                    " When '" & SprSlRetCsh & "' Then 'Return    ' " & _
                    " When '" & SprSlRetCre & "' Then 'Return    ' End as TrnType," & _
                    "SP_Sale.DocID,SP_Sale.V_Date," & _
                    "" & cIIF("(SP_Sale.SprAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt)=0", "0", "((SP_Sale.SprAmt_TB+SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt) - (TOT_Amt+SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB))") & " AS TBAmt," & _
                    "" & cIIF("(SP_Sale.SprAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB) - (SP_Sale.Tax_AmtMRP+SP_Sale.TaxSur_AmtMRP+TOT_AmtMRP+SP_Sale.D_Amt_MRP_TB))") & " AS TBMRPAmt," & _
                    "" & cIIF("(SP_Sale.SprAmt_MRP_TP)=0", "0", "((SP_Sale.SprAmt_MRP_TP) - (SP_Sale.D_Amt_MRP_TP))") & " AS TPSprMRPAmt," & _
                    "SP_Sale.OilAmt_MRP_TP AS TPOilMRPAmt," & _
                    "" & cIIF("(SP_Sale.SprAmt_TP)=0", "0", "((SP_Sale.SprAmt_TP) - (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP))") & " AS TPSprAmt," & _
                    "SP_Sale.OilAmt_TP AS TPOilAmt," & _
                    "SP_Sale.Packing,SP_Sale.TOT_Amt,SP_Sale.Total_Amt,TaxForms.Form_Code,TaxForms.Form_Desc,TaxForms.L_C,TaxForms.Tax_Per,TaxForms.Tax_Sur_Per,SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP as TaxAmt,SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP as TaxSurAmt,(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP) as DiscAmt" & _
                    " FROM SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code" & _
                    " Where  SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID("SP_Sale.DocID", "4", "5") & " in ('" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "') " & Condstr & " order by SP_Sale.V_Date"
            End If
        End If
    End If
        RepName = "SprSaleSum"
    ElseIf GRepFormName = SprPurSum Then
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Purch.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_Purch.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Purch.DocId,1) in (" & GridString2 & ")"
        
'        mQRY = "SELECT TaxForms.Form_Code, TaxForms.Form_Desc, TaxForms.L_C, TaxForms.Tax_Per," & _
               "TaxForms.Tax_Sur_Per, SP" & _
               " FROM SP_Purch LEFT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code" & _
               " where SP_Purch.v_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and SP_Purch.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & _
               "# and mid(SP_Purch.DocID,4,5) in ('" & SprPurCsh & "','" & SprPurCre & "','" & SprPrRetCsh & "','" & SprPrRetCre & "') " & CondStr
        mQry = " SELECT " & _
             " SP_Purch.Form_Code, TaxForms.Form_Desc, TaxForms.L_C, TaxForms.Tax_Per,TaxForms.Tax_Sur_Per," & _
             " SP_Purch.Docid," & cMID("SP_Purch.Docid", "4", "5") & " as trn_type,SP_Purch.V_Date,SP_Purch.Tot_Goods_Value,SP_Purch.Tax_Amt,SP_Purch.Addition,SP_Purch.Deduction,SP_Purch.NET_AMT" & _
             " FROM SP_Purch RIGHT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code" & _
             " where SP_Purch.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Purch.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
            " and " & cMID("SP_Purch.DocID", "4", "5") & " in ('" & SprPurCsh & "','" & SprPurCre & "','" & SprPrRetCsh & "','" & SprPrRetCre & "') " & Condstr
        RepName = "SprPurSum"
    End If
      
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprPartPurSal()
On Error GoTo ELoop
Dim mQry As String, Condstr$, CondStr1$
Dim PartyType As Byte
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    PartyType = GCn.Execute("select PartyType from SYCTRL").Fields(0).Value
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_stock.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Sp_Stock.Part_No in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Stock.DocId,1) in (" & GridString3 & ")"
    
    
    If FGrid.TextMatrix(List2, 1) = "Counter" Then
        CondStr1 = "'" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "'"
    ElseIf FGrid.TextMatrix(List2, 1) = "Workshop" Then
        CondStr1 = "'" & WksSlCsh & "','" & WksSlCre & "'"
    ElseIf FGrid.TextMatrix(List2, 1) = "Both" Then
        CondStr1 = "'" & WksSlCsh & "','" & WksSlCre & "','" & SprSlCsh & "','" & SprSlCre & "','" & SprSlRetCsh & "','" & SprSlRetCre & "'"
    End If
    
    If GRepFormName = SprPartSale Then
            If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
            If Check1(4).Value = Unchecked Then Condstr = Condstr & " and Part.Part_Grade in (" & GridString4 & ")"
            
            If FGrid.TextMatrix(List1, 1) = "Summary" Then
                mQry = "SELECT Part.Part_Name, SP_Stock.Part_No, Sum(sp_Stock.Qty_Iss-Qty_Ret) AS Qty, Sum(SP_Stock.Amount) AS Amt " & _
                    " FROM Sp_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) " & _
                    " Where SP_Stock.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Stock.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID$("SP_Stock.Invoice_DocID", "4", "5") & " in (" & CondStr1 & ") " & Condstr & _
                    " GROUP BY Part.Part_Name, SP_Stock.Part_No"
                RepName = "SprPartSale"
            Else
                mQry = "SELECT SP_Stock.DocID,SP_Stock.V_Type,SG.Name as PartyName,SP_Stock.Job_DocId,H.RegNo,H.Chassis,Part.Part_Name, SP_Stock.Part_No, Sum(sp_Stock.Qty_Iss-Qty_Ret) AS Qty, Sum(SP_Stock.Amount) AS Amt,SP_Stock.V_Date " & _
                    " FROM (((Sp_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1)) " & _
                    " LEFT JOIN SubGroup SG on SP_Stock.Party_Code = SG.SubCode) " & _
                    " LEFT JOIN Job_Card ON SP_Stock.Job_Docid = Job_Card.DocId) " & _
                    " LEFT JOIN HisCard H ON Job_Card.CardNo = H.CardNo " & _
                    " Where SP_Stock.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Stock.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " and " & cMID$("SP_Stock.Invoice_DocID", "4", "5") & " in (" & CondStr1 & ") " & Condstr & _
                    " GROUP BY Part.Part_Name, SP_Stock.Part_No,SP_Stock.DocID,SP_Stock.V_Type,SG.Name ,SP_Stock.Job_DocId,H.RegNo,H.Chassis,SP_Stock.V_Date"
                RepName = "SprPartSaleDet"
    
            End If
        
    ElseIf GRepFormName = SprPartPur Then
    
        If FGrid.TextMatrix(List1, 1) = "Summary" Then
            mQry = "SELECT Part.Part_Name, SP_Stock.Part_No, " & cIIF("SP_Stock.L_C='C'", "Sum(sp_Stock.Qty_Rec)", "0") & " AS telcoqty, " & cIIF("SP_Stock.L_C='L'", "Sum(sp_Stock.Qty_Rec)", "0") & " AS localqty, " & cIIF("SP_Stock.L_C='C'", "Sum(SP_Stock.Amount)", "0") & " AS telcoamt, " & cIIF("SP_Stock.L_C='L'", "Sum(SP_Stock.Amount)", "0") & " AS localamt " & _
               "FROM (SP_Stock LEFT JOIN SubGroup ON SP_Stock.Party_Code = SubGroup.SubCode) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) " & _
                " Where SP_Stock.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Stock.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                " and " & cMID$("SP_Stock.Invoice_DocID", "4", "5") & " in ('" & SprPurCsh & "','" & SprPurCre & "','" & SprPrRetCsh & "','" & SprPrRetCre & "') " & Condstr & _
                " GROUP BY Part.Part_Name, SP_Stock.Part_No, SubGroup.Party_Type,SP_Stock.L_C"
            RepName = "SprPartPur"
        Else
'            mQRY = "SELECT Part.Part_Name, SP_Stock.Part_No, IIf(SP_Stock.L_C='C',Sum(sp_Stock.Qty_Rec),0) AS telcoqty, IIf(SP_Stock.L_C='L',Sum(sp_Stock.Qty_Rec),0) AS localqty, IIf(SP_Stock.L_C='C',Sum(SP_Stock.Amount),0) AS telcoamt, IIf(SP_Stock.L_C='L',Sum(SP_Stock.Amount),0) AS localamt " & _
'               " FROM (SP_Stock LEFT JOIN SubGroup ON SP_Stock.Party_Code = SubGroup.SubCode) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) " & _
'                " Where SP_Stock.v_Date  >= #" & Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy") & "# and SP_Stock.v_Date <= #" & Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy") & _
'                "# and mid$(SP_Stock.Invoice_DocID,4,5) in ('" & SprPurCsh & "','" & SprPurCre & "','" & SprPrRetCsh & "','" & SprPrRetCre & "') " & Condstr & _
'                " GROUP BY Part.Part_Name, SP_Stock.Part_No, SubGroup.Party_Type,SP_Stock.L_C"

            mQry = "SELECT Part.Part_Name, SP_Stock.Part_No, Sp_Stock.Qty_Rec AS TotRec,  SP_Stock.Amount2 as  TotAmt,Sp_Purch.Party_Name,Sp_Purch.Party_Doc_No,Sp_Purch.Party_Doc_Date,Sp_Stock.Disc_Per,Sp_Stock.V_No,Sp_Stock.V_Date,Sp_Stock.Disc_Amt,Part.Mrp ,Sp_Stock.Rate2 " & _
                    " FROM ((SP_Stock LEFT JOIN SubGroup ON SP_Stock.Party_Code = SubGroup.SubCode) " & _
                    " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1)) " & _
                    " Left Join Sp_Purch On Sp_Stock.DocId=Sp_Purch.DocId " & _
                    " Where SP_Stock.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Stock.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & _
                    " And " & cMID$("SP_Stock.Invoice_DocID", "4", "5") & " in ('" & SprPurCsh & "','" & SprPurCre & "','" & SprPrRetCsh & "','" & SprPrRetCre & "') " & Condstr
                    
            RepName = "SprPartPurDet"
        End If
        
        
        
    End If
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION:  RepPrint = False: Exit Sub
    
'    RstRep.MoveFirst
'    Do While RstRep.EOF = False
'        Set TmpRst = G_FaCn.Execute("select Sum(AmtDr-AmtCr) as Exp   from ledger  where VehNo='" & RstRep!TruckNo & "' and V_Date=" & ConvertDate(RstRep!V_DATE) & "")
'            If TmpRst.EOF = False Then
'                RstRep!truckExp = VNull(TmpRst!Exp)
'                RstRep.Update
'            End If
'            RstRep.MoveNext
'    Loop
    
    
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description




End Sub

Private Sub SprStkReOrder()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    '"Above Maximum", "Below Minimum", "Below ReOrder"
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    'If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Part.Part_No in (" & GridString2 & ")"
    
     If Check1(1).Value = Unchecked Then Condstr = Condstr & "  and part.Site_Code in (" & GridString1 & ")  "
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & "  and part.site_code ='" & PubSiteCode & "'  "
       
    End If
    
    Select Case FGrid.TextMatrix(List1, 1)
        Case "Above Maximum"
        'SELECT SP_Stock.Part_No, 0 AS TPQtyRec, 0 AS TBQtyRec, 0 AS TPQtyIss, 0 AS TBQtyIss, 0 as TPQtyBalance,0 as TBQtyBalance,Part.Part_Name FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) WHERE   GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name  UNION ALL SELECT SP_Stock.Part_No, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Rec),0) AS TPQtyRec, IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Rec),0) AS TBQtyRec, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Iss),0) AS TPQtyIss,IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Iss),0) AS TBQtyIss,(TPQtyRec)-(TPQtyIss) as TPQtyBalance,(TBQtyRec)-(TBQtyIss) as TBQtyBalance, Part.Part_Name FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) WHERE Where Div_Code='" & PubDivCode & "' And (Cur_MRP_TBStk+ Cur_MRP_TPStk + Cur_TB_Stk + Cur_TP_Stk)  > Max_Lvl"

            mQry = "SELECT PART_NO, Part_Name, Min_Lvl as Lvl, (Cur_MRP_TBStk + Cur_TB_Stk ) as TBStk ,(Cur_MRP_TPStk + Cur_TP_Stk) as TPStk,High_Pur_Rate, (Cur_MRP_TBStk+ Cur_MRP_TPStk + Cur_TB_Stk + Cur_TP_Stk) AS CurrStk  " & _
            "FROM PART Where Div_Code='" & PubDivCode & "' And (Cur_MRP_TBStk+ Cur_MRP_TPStk + Cur_TB_Stk + Cur_TP_Stk) <  Max_Lvl"
        Case "Below Minimum"
            mQry = "SELECT PART_NO, Part_Name, Min_Lvl as Lvl, (Cur_MRP_TBStk + Cur_TB_Stk ) as TBStk ,(Cur_MRP_TPStk + Cur_TP_Stk) as TPStk,High_Pur_Rate, (Cur_MRP_TBStk+ Cur_MRP_TPStk + Cur_TB_Stk + Cur_TP_Stk) AS CurrStk  " & _
            "FROM PART Where Div_Code='" & PubDivCode & "' And (Cur_MRP_TBStk+ Cur_MRP_TPStk + Cur_TB_Stk + Cur_TP_Stk) <  Min_Lvl"
             ' mQRY = "SELECT SP_Stock.Part_No, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Rec),0) AS TPQtyRec, IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Rec),0) AS TBQtyRec, IIf(SP_Stock.Tax_YN=0,Sum(SP_Stock.Qty_Iss),0) AS TPQtyIss, IIf(SP_Stock.Tax_YN=1,Sum(SP_Stock.Qty_Iss),0) AS TBQtyIss, (TBQtyRec)-(TBQtyIss) AS TBStk, (TPQtyRec)-(TPQtyIss) AS TPStk, Part.Part_Name,Part.High_Pur_Rate FROM SP_Stock LEFT JOIN Part ON (Part.Div_Code = left(SP_Stock.Docid,1)) AND (SP_Stock.Part_No = Part.PART_NO) WHERE Part.Min_Lvl<=2 and Div_Code='" & PubDivCode & "' GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name,High_Pur_Rate"

        Case "Below ReOrder"
            mQry = "SELECT PART_NO, Part_Name, ReOrd_Lvl as Lvl, (Cur_MRP_TBStk + Cur_TB_Stk ) as TBStk ,(Cur_MRP_TPStk + Cur_TP_Stk) as TPStk,High_Pur_Rate, (Cur_MRP_TBStk+ Cur_MRP_TPStk + Cur_TB_Stk + Cur_TP_Stk) AS CurrStk  " & _
            "FROM PART Where Div_Code='" & PubDivCode & "' And (Cur_MRP_TBStk+ Cur_MRP_TPStk + Cur_TB_Stk + Cur_TP_Stk) <  ReOrd_Lvl"
    End Select
   
    mQry = mQry + Condstr
        
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprStkReOrd"
    RepTitle = UCase(Me.CAPTION) + "[" + FGrid.TextMatrix(List1, 1) + "]"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprStkBinLoc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    'Condstr = "WHERE SP_Stock.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & " and SP_Stock.V_Date<= " & ConvertDate(Format(DateAdd("D", -1, FGrid.TextMatrix(Date2, 1)), "dd/MMM/yyyy")) & " "
    If FGrid.TextMatrix(List2, 1) = "Yes" Then GridString2 = MarkRecCalculate 'Else MsgBox "No"
    If GridString2 = Empty And FGrid.TextMatrix(List2, 1) = "Yes" Then: MsgBox "** No Records Found to Print **": RepPrint = False: Exit Sub: Else If FGrid.TextMatrix(List2, 1) = "Yes" Then Condstr = Condstr & " and SP_Stock.Part_No in (" & left(GridString2, (Len(GridString2) - 1)) & ")" 'Else MsgBox "No"
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
      
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " in (" & GridString1 & ")"
   If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

    'If GridString2 = "''" Then
    '    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and (Part.Bin_Loca is null or Part.Bin_Loca='')"
    'Else
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Part.Bin_Loca in (" & GridString2 & ")"
    'End If
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Stock.DocId,1) in (" & GridString3 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "PartNo" Then
    mQry = "SELECT Part.Bin_Loca,SP_Stock.Part_No, 0 AS TPQtyRec, 0 AS TBQtyRec, 0 AS TPQtyIss, 0 AS TBQtyIss, " & cIIF("SP_Stock.Tax_YN=0 and SP_Stock.V_Type='SXAO'", "Sum(SP_Stock.Qty_Rec)-Sum(SP_Stock.Qty_Iss)", "0") & " AS TPQtyOpen, " & cIIF("SP_Stock.Tax_YN=1 and SP_Stock.V_Type='SXAO'", "Sum(SP_Stock.Qty_Rec)-Sum(SP_Stock.Qty_Iss)", "0") & " AS TBQtyOpen, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1)  " & _
        "Where Sp_Stock.V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " " & _
        "" & Condstr & _
        " GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name,Part.Bin_Loca,SP_Stock.V_Type " & _
        " Union All " & _
        "SELECT Part.Bin_Loca,SP_Stock.Part_No, " & cIIF("SP_Stock.Tax_YN=0 and SP_Stock.V_Type <> 'SXAO'", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TPQtyRec, " & cIIF("SP_Stock.Tax_YN=1 and SP_Stock.V_Type <> 'SXAO'", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TBQtyRec, " & cIIF("SP_Stock.Tax_YN=0", "Sum(SP_Stock.Qty_Iss)-Sum(SP_Stock.Qty_Ret)", "0") & " AS TPQtyIss," & cIIF("SP_Stock.Tax_YN=1", "Sum(SP_Stock.Qty_Iss)-Sum(SP_Stock.Qty_Ret)", "0") & " AS TBQtyIss, 0 AS TPQtyOpen, 0 AS TBQtyOpen, Part.Part_Name " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1)" & _
        "Where Sp_Stock.V_Date>=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And Sp_Stock.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & "" & _
        "" & Condstr & _
        "GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name, Part.Bin_Loca,SP_Stock.V_Type"
        
        If FGrid.TextMatrix(List3, 1) = "Detail" Then
            RepName = "SprstkBin"
        Else
            RepName = "SprstkBinSum"
        End If
    ElseIf FGrid.TextMatrix(List1, 1) = "Bin + PartNo" Then
    mQry = "SELECT Part.Bin_Loca,SP_Stock.Part_No,  0 AS TPQtyRec, 0 AS TBQtyRec, 0 AS TPQtyIss,0 AS TBQtyIss, " & cIIF("SP_Stock.Tax_YN=0 and SP_Stock.V_Type='SXAO'", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TPQtyOpen, " & cIIF("SP_Stock.Tax_YN=1 and SP_Stock.V_Type='SXAO'", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TBQtyOpen, Part.Part_Name,Part.PhyStk " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        "and Part.Div_Code = left(SP_Stock.Docid,1)  " & _
        "Where Sp_Stock.V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " " & _
        "" & Condstr & _
        " GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name,Part.PhyStk,Part.Bin_Loca,SP_Stock.V_Type " & _
        " Union All " & _
        "SELECT Part.Bin_Loca,SP_Stock.Part_No, " & cIIF("SP_Stock.Tax_YN=0 and SP_Stock.V_Type <> 'SXAO'", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TPQtyRec, " & cIIF("SP_Stock.Tax_YN=1 and SP_Stock.V_Type <> 'SXAO'", "Sum(SP_Stock.Qty_Rec)", "0") & " AS TBQtyRec, " & cIIF("SP_Stock.Tax_YN=0", "Sum(SP_Stock.Qty_Iss)-Sum(SP_Stock.Qty_Ret)", "0") & " AS TPQtyIss, " & cIIF("SP_Stock.Tax_YN=1", "Sum(SP_Stock.Qty_Iss)-Sum(SP_Stock.Qty_Ret)", "0") & " AS TBQtyIss, 0 AS TPQtyOpen, 0 AS TBQtyOpen, Part.Part_Name,Part.PhyStk " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        "and Part.Div_Code = left(SP_Stock.Docid,1)" & _
        "Where Sp_Stock.V_Date>=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And Sp_Stock.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & "" & _
        "" & Condstr & _
        "GROUP BY SP_Stock.Part_No, SP_Stock.Tax_YN, Part.Part_Name,Part.PhyStk,Part.Bin_Loca,SP_Stock.V_Type "
        If FGrid.TextMatrix(List3, 1) = "Detail" Then
            RepName = "SprstkBinIndex"
        Else
            RepName = "SprstkBinIndexSum"
        End If
        
    End If
'    mQRY = mQRY + Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprOthPurRegs()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
   
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    Condstr = " Where SP_Purch.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Purch.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

     If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = Condstr & " and SP_Purch.V_Type In ('" & OthPurCre & "','" & OthPurCsh & "') "
     If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = Condstr & " and SP_Purch.V_Type = '" & OthPurCre & "' "
     If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = Condstr & " and SP_Purch.V_Type = '" & OthPurCsh & "' "
     If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Purch.DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_Purch.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

     If Check1(2).Value = Unchecked Then Condstr = Condstr & " and SP_Purch.Party_Code in (" & GridString2 & ")"
     If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Purch.DocId,1) in (" & GridString3 & ")"
    
     mQry = "SELECT docid,v_date, v_no,  " & cMID("DocID", "9", "5") & " as VPrefix,Cash_Credit,party_name,L_C,Form_Code,Party_Doc_No,Party_Doc_Date,Remarks,Net_Amt " & _
             "FROM sp_purch " & Condstr
         
            
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "OthPurReg"
    RepTitle = UCase(Me.CAPTION) + "[" + FGrid.TextMatrix(List1, 1) + "]"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ProcBudgetExpVariRep()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
   
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

    Condstr = " Where SP_Purch.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Purch.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

     If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = Condstr & " and SP_Purch.V_Type In ('" & OthPurCre & "','" & OthPurCsh & "') "
     If FGrid.TextMatrix(List1, 1) = "Credit" Then Condstr = Condstr & " and SP_Purch.V_Type = '" & OthPurCre & "' "
     If FGrid.TextMatrix(List1, 1) = "Cash" Then Condstr = Condstr & " and SP_Purch.V_Type = '" & OthPurCsh & "' "
     
     If Check1(1).Value = Unchecked Then Condstr = Condstr & " and BE.Site_Code in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and be.site_code ='" & PubSiteCode & "' "
      End If

     If Check1(2).Value = Unchecked Then Condstr = Condstr & " and BE.ExpAc in (" & GridString2 & ")"
     
    
     
    mQry = "Select BE.ExpAc, S.Name As ExpenceAcName, BE.Site_Code, Site.Site_Desc, BE.Amount As Budgeted_Amount, M.Name As Month_Name, " & _
           "(Select Sum(AmtDr-AmtCr) From Ledger L Where SubCode=BE.ExpAc And Right(L.Site_Code,1)=BE.Site_Code And " & cMth("L.V_Date") & "=BE.Month) As Actual_Amount " & _
           "From (((Budget_Exp BE " & _
           "Left Join Chas_Mth M On BE.Month=M.Code) " & _
           "Left Join SubGroup S On BE.ExpAc=S.SubCode) " & _
           "Left Join Site On Site.Site_Code=BE.Site_Code) " & _
           "Order By S.Name, Site.Site_Desc, BE.Month"
    
         
            
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "BudgetExpVariation"
    If UCase(FGrid.TextMatrix(List1, 1)) = "Yes" Then
        RepTitle = UCase(Me.CAPTION) + "[Branch Wise]"
    End If
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ProcSalesManCostRep()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1$
'Date1,Date2,List1,List1,List1,List2,List1,List1
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
   
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " Where EE.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and EE.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    CondStr1 = " Where Veh_Order.Inv_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and Veh_Order.Inv_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
     
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        Condstr = Condstr & " And E.Designation In ('Sales Manager', 'Sales Representative') "
    End If
     
     
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and E.Site_Code in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and e.site_code ='" & PubSiteCode & "' "
        End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and EE1.Emp_Code in (" & GridString2 & ")"
                 
    mQry = "Select EE1.Emp_Code, Max(E.Emp_Name) As Emp_Name, Max(M.Name) As MonthName, Sum(EE1.Amount) As ExpAmount, " & _
           "Max(VO.ChasCount) As ChasCount, Max(VO.SaleValue) As SaleValue " & _
           "From ((((Exp_Emp1 EE1 " & _
           "Left Join Ledger EE On EE.DocId = EE1.DocId and EE.SubCode = EE1.SubCode )  " & _
           "Left Join (Select Rep_Code, " & cMth("Inv_Date") & " As Inv_Date,Count(*) As ChasCount, Sum(Net_Amount) As SaleValue From Veh_Order  " & CondStr1 & " Group By Rep_Code, " & cMth("Inv_Date") & ") As VO On EE1.Emp_Code=VO.Rep_Code And " & cMth("EE.V_Date") & " =  VO.Inv_Date ) " & _
           "Left Join Emp_Mast E On EE1.Emp_Code=E.Emp_Code) " & _
           "Left Join Chas_Mth M On M.Code =  " & cMth("EE.V_Date") & ") " & Condstr & _
           "Group By EE1.Emp_Code, " & cMth("EE.V_Date") & ""
            
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If UCase(FGrid.TextMatrix(List1, 1)) = "SUMMARY" Then
        RepName = "SalesManCostRepSum"
    Else
        RepName = "SalesManCostRep"
    End If
    
    RepTitle = UCase(Me.CAPTION) & " [" & UCase(FGrid.TextMatrix(List1, 1)) & "]"

    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub



Private Sub ProcBillWiseOutstanding()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, CondStr1$
'Date1,Date2,List1,List1,List1,List2,List1,List1
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
   
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " Where S.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and S.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " And L.DocId Is Not Null "
    CondStr1 = " Where S.Inv_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and S.Inv_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " And L.DocId Is Not Null "
     
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and S.Site_Code in (" & GridString1 & ") And L.DocId Is Not Null"
    If Check1(1).Value = Checked Then
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and  s.site_code ='" & PubSiteCode & "' and l.docid is not null "
        End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and SG.SubCode in (" & GridString2 & ")"
                 
    If FGrid.TextMatrix(List1, 1) = "Debtors" Then
        mQry = "Select 'Spare' As Category, S.DocId, Max(S.V_Date) As V_Date, S.Party_Code, Max(SG.Name) As Party_Name, Sum(S.Total_Amt) As Bill_Amount, Sum(LA.Adj_Amt) As Adj_Amount " & _
               "From (((Sp_Sale S " & _
               "Left Join Ledger L on S.DocId = L.DocId ) " & _
               "Left Join (Select DocId2 As DocId_Bill, Sum(Cr) As Adj_Amt From LedgerAdj Group By DocId2) As LA On S.DocId=LA.DocId_Bill) " & _
               "Left Join SubGroup SG On SG.SubCode=S.Party_Code) " & Condstr & " And S.Cash_Credit='Credit'  " & _
               " Group By S.Party_Code, S.DocId"
        mQry = mQry & " Union All " & _
               "Select 'Vehicle' As Category, S.Inv_DocId As DocId, Max(Inv_Date) As V_Date, S.PartyCode As Party_Code, Max(SG.Name) As Party_Name, Sum(S.Net_Amount) As Bill_Amount, Sum(LA.Adj_Amt) As Adj_Amount " & _
               "From (((Veh_Order S " & _
               "Left Join Ledger L On L.DocId=S.Inv_DocId) " & _
               "Left Join (Select DocId2 As DocId_Bill, Sum(Cr) As Adj_Amt From LedgerAdj Group By DocId2) As LA On S.Inv_DocId=LA.DocId_Bill) " & _
               "Left Join SubGroup SG On SG.SubCode=S.PartyCode) " & CondStr1 & _
               " Group By S.PartyCode, S.Inv_DocId"
    Else
        mQry = "Select 'Spare' As Category, S.DocId, Max(S.V_Date) As V_Date, S.Party_Code, Max(SG.Name) As Party_Name, Sum(S.Net_Amt) As Bill_Amount, Sum(LA.Adj_Amt) As Adj_Amount " & _
               "From (((Sp_Purch S " & _
               "Left Join Ledger L On L.DocId=S.DocID) " & _
               "Left Join (Select DocId1 As DocId_Bill, Sum(Cr) As Adj_Amt From LedgerAdj Group By DocId1) As LA On S.DocId=LA.DocId_Bill) " & _
               "Left Join SubGroup SG On SG.SubCode=S.Party_Code) " & Condstr & " And S.Cash_Credit='Credit' " & _
               " Group By S.Party_Code, S.DocId"
        mQry = mQry & " Union All " & _
               "Select 'Vehicle' As Category, S.DocId, Max(S.V_Date) As V_Date, S.PartyCode As Party_Code, Max(SG.Name) As Party_Name, Sum(S.Tot_Amount) As Bill_Amount, Sum(LA.Adj_Amt) As Adj_Amount " & _
               "From (((Veh_Purch1 S " & _
               "Left Join Ledger L On L.DocId = S.DocId) " & _
               "Left Join (Select DocId1 As DocId_Bill, Sum(Cr) As Adj_Amt From LedgerAdj Group By DocId1) As LA On S.DocId=LA.DocId_Bill) " & _
               "Left Join SubGroup SG On SG.SubCode=S.PartyCode) " & Condstr & _
               " Group By S.PartyCode, S.DocId"
    End If
    
            
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "BillWiseOutstanding"
    
    RepTitle = UCase(Me.CAPTION) & " [" & UCase(FGrid.TextMatrix(List1, 1)) & "]"

    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprSaleTaxCtrlStmtRegs()
On Error GoTo ELoop
Dim mQry As String, mQRY1 As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim sprAmtTBAftDisc As Long, OilAmtTBAftDisc As Long
Dim MyRst As ADODB.Recordset, MyRstRet As ADODB.Recordset
Set MyRst = New ADODB.Recordset
Set MyRstRet = New ADODB.Recordset
Set myRst1 = New ADODB.Recordset
'Date1,Date2,List1,List1,List1,List2,List1,List1
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
   
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    Condstr = " and SP.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

     
     If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " in (" & GridString1 & ")"
     If Check1(1).Value = Checked Then
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If

     If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(SP.DocId,1) in (" & GridString2 & ")"
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "TBSal", adDouble, 12, adFldIsNullable
        .Fields.Append "TaxPer", adDouble, 12, adFldIsNullable
        .Fields.Append "SurchargePer", adDouble, 12, adFldIsNullable
        .Fields.Append "TaxAmt", adDouble, 12, adFldIsNullable
        .Fields.Append "SurchargeAmt", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    'Discount on Lube Applicability
    If PubDiscOnLube = 1 Then
        Condstr2 = "sum(" & cIIF("((SP.SprAmt_MRP_TP)+(SP.OilAmt_MRP_TP))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_MRP_TP) - ((SP.SprAmt_MRP_TP) * (SP.D_Amt_MRP_TP)) / ((SP.SprAmt_MRP_TP) + (SP.OilAmt_MRP_TP)))") & ")  AS SprAmtMRPTP, " & _
                   "sum(" & cIIF("((SP.SprAmt_MRP_TB)+(SP.OilAmt_MRP_TB))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_MRP_TB) - ((SP.SprAmt_MRP_TB) * (SP.D_Amt_MRP_TB)) / ((SP.SprAmt_MRP_TB) + (SP.OilAmt_MRP_TB)))") & ")  AS SprAmtMRPTB, " & _
                   "sum(" & cIIF("((SP.SprAmt_TB)+(SP.OilAmt_TB ))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_TB ) - ((SP.SprAmt_TB)* (SP.D_Amt_TB-SP.D_Amt_MRP_TB)) / ((SP.SprAmt_TB)+(SP.OilAmt_TB))) ") & ")  AS SprAmtTB, " & _
                   "sum(" & cIIF("((SP.SprAmt_TP)+(SP.OilAmt_TP ))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_TP) - ((SP.SprAmt_TP)* (SP.D_Amt_TP-SP.D_Amt_MRP_TP)) / ((SP.SprAmt_TP)+(SP.OilAmt_TP)))") & ")  AS SprAmtTP, " & _
                   "sum(" & cIIF("((SP.SprAmt_MRP_TB)+(SP.OilAmt_MRP_TB))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.OilAmt_MRP_TB) - ((SP.OilAmt_MRP_TB) * (SP.D_Amt_MRP_TB)) / ((SP.SprAmt_MRP_TB) + (SP.OilAmt_MRP_TB)))") & ")  AS OilAmtMRPTB, " & _
                   "sum(" & cIIF("((SP.SprAmt_MRP_TP)+(SP.OilAmt_MRP_TP))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.OilAmt_MRP_TP) - ((SP.OilAmt_MRP_TP) * (SP.D_Amt_MRP_TP)) / ((SP.SprAmt_MRP_TP) + (SP.OilAmt_MRP_TP)))") & ")  AS OilAmtMRPTP, " & _
                   "sum(" & cIIF("((SP.SprAmt_TB)+(SP.OilAmt_TB ))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.OilAmt_TB) - ((SP.OilAmt_TB )* (SP.D_Amt_TB-SP.D_Amt_MRP_TB)) / ((SP.SprAmt_TB)+(SP.OilAmt_TB)))") & ")  as OilAmtTB , " & _
                   "sum(" & cIIF("((SP.SprAmt_TP)+(SP.OilAmt_TP ))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.OilAmt_TP) - ((SP.OilAmt_TP)* (SP.D_Amt_TP-SP.D_Amt_MRP_TP)) / ((SP.SprAmt_TP)+(SP.OilAmt_TP)))") & ")  as OilAmtTP  "
        ' Taxable sale for each tax percentage
        CondStr3 = "sum(" & cIIF("((SP.SprAmt_MRP_TP)+(SP.OilAmt_MRP_TP))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.OilAmt_MRP_TP) - ((SP.OilAmt_MRP_TP) * (SP.D_Amt_MRP_TP)) / ((SP.SprAmt_MRP_TP) + (SP.OilAmt_MRP_TP)))") & ")  AS OilAmtMRPTP, " & _
                   "sum(" & cIIF("((SP.SprAmt_MRP_TB)+(SP.OilAmt_MRP_TB))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.OilAmt_MRP_TB) - ((SP.OilAmt_MRP_TB) * (SP.D_Amt_MRP_TB)) / ((SP.SprAmt_MRP_TB) + (SP.OilAmt_MRP_TB)))") & ") AS OilAmtMRPTB, " & _
                   "sum(" & cIIF("((SP.SprAmt_TB)+(SP.OilAmt_TB ))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.OilAmt_TB) - ((SP.OilAmt_TB )* (SP.D_Amt_TB-SP.D_Amt_MRP_TB)) / ((SP.SprAmt_TB)+(SP.OilAmt_TB)))") & ")  as OilAmtTB , " & _
                   "sum(" & cIIF("((SP.SprAmt_MRP_TB)+(SP.OilAmt_MRP_TB))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_MRP_TB) - ((SP.SprAmt_MRP_TB) * (SP.D_Amt_MRP_TB)) / ((SP.SprAmt_MRP_TB) + (SP.OilAmt_MRP_TB)))") & ") AS SprAmtMRPTB, " & _
                   "sum(" & cIIF("((SP.SprAmt_TB)+(SP.OilAmt_TB ))=0 and SP.V_Type Not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_TB ) - ((SP.SprAmt_TB)* (SP.D_Amt_TB)) / ((SP.SprAmt_TB)+(SP.OilAmt_TB)))") & ") AS SprAmtTB, "
        
    Else
        Condstr2 = "sum(" & cIIF("(SP.SprAmt_MRP_TB)=0 and SP.V_Type not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_MRP_TB) - (SP.D_Amt_MRP_TB)))") & " AS SprAmtMRPTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "((SP.SprAmt_MRP_TB) - (SP.D_Amt_MRP_TB))", "0") & ") AS SprAmtMRPTBRet, " & _
                   "sum(" & cIIF("(SP.SprAmt_MRP_TP)=0 and SP.V_Type not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_MRP_TP) - (SP.D_Amt_MRP_TP))") & ") AS SprAmtMRPTP, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "((SP.SprAmt_MRP_TP) - (SP.D_Amt_MRP_TP))", "0") & ") AS SprAmtMRPTPRet, " & _
                   "sum(" & cIIF("(SP.SprAmt_TB)=0 and SP.V_Type not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_TB ) - (SP.D_Amt_TB-SP.D_Amt_MRP_TB))") & ") AS SprAmtTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "((SP.SprAmt_TB) - (SP.D_Amt_TB-SP.D_Amt_MRP_TB))", "0") & ") AS SprAmtTBRet, " & _
                   "sum(" & cIIF("(SP.SprAmt_TP)=0 and SP.V_Type not in ('SXSRC','SXSRR')", "0", "((SP.SprAmt_TP) - (SP.D_Amt_TP-SP.D_Amt_MRP_TP))") & ") AS SprAmtTP, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "((SP.SprAmt_TP) - (SP.D_Amt_TP-SP.D_Amt_MRP_TP))", "0") & ") AS SprAmtTPRet, " & _
                   "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.OilAmt_MRP_TB)", "0") & ") AS OilAmtMRPTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.OilAmt_MRP_TB)", "0") & ") AS OilAmtMRPTBRet, " & _
                   "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.OilAmt_MRP_TP)", "0") & ") AS OilAmtMRPTP, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.OilAmt_MRP_TP)", "0") & ") AS OilAmtMRPTPRet, " & _
                   "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.OilAmt_TB)", "0") & ") AS OilAmtTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.OilAmt_TB)", "0") & ") AS OilAmtTBRet, " & _
                   "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.OilAmt_TP)", "0") & ") AS OilAmtTP, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.OilAmt_TP)", "0") & ") AS OilAmtTPRet "
' Taxable sale for each tax percentage
        CondStr3 = "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.OilAmt_MRP_TB)", "0") & ") AS OilAmtMRPTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.OilAmt_MRP_TB)", "0") & ") AS OilAmtMRPTBRet, " & _
                   "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.OilAmt_TB)", "0") & ") AS OilAmtTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.OilAmt_TB)", "0") & ") AS OilAmtTBRet, " & _
                   "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.SprAmt_MRP_TB) - (SP.D_Amt_MRP_TB)", "0") & ") AS SprAmtMRPTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.SprAmt_MRP_TB) - (SP.D_Amt_MRP_TB)", "0") & ") AS SprAmtMRPTBRet, " & _
                   "sum(" & cIIF("SP.V_Type not in ('SXSRC','SXSRR')", "(SP.SprAmt_TB-(SP.D_Amt_TB-SP.D_Amt_MRP_TB))", "0") & ") AS SprAmtTB, " & _
                   "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.SprAmt_TB-(SP.D_Amt_TB-SP.D_Amt_MRP_TB))", "0") & ") AS SprAmtTBRet, "
    End If
    'GTO , Disc , AfterDiscSale , TP_SpareSal,LubSal Calculation
     mQry = "SELECT sum(SP.Total_Amt) AS GTO,sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "SP.Total_Amt", "0") & ") as TotRet, " & _
            "Sum(SP.D_Amt_TB+D_Amt_MRP_TB+SP.D_Amt_TP+D_Amt_MRP_TP) AS DisAmt," & _
            Condstr2 & _
            "FROM SP_Sale SP  " & _
            "Where SP.V_Type In ('SYSIC','SYSIR','W_SIC','W_SIR','SXSRC','SXSRR')" & Condstr
     myRst1.Open mQry, GCn, adOpenDynamic, adLockOptimistic
     
    
    ' TB Sale on All Tax Percentages
     Dim I As Integer
     mQry = "SELECT " & _
            CondStr3 & _
            "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.Tax_AmtMRP)", "0") & ") AS Tax_AmtMRP, " & _
            "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.Tax_AmtMRP)", "0") & ") AS Tax_AmtMRPRet, " & _
            "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.TaxSur_AmtMRP)", "0") & ") AS TaxSur_AmtMRP, " & _
            "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.TaxSur_AmtMRP)", "0") & ") AS TaxSur_AmtMRPRet, " & _
            "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.TOT_Amt)", "0") & ") AS TOT_Amt, " & _
            "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.TOT_Amt)", "0") & ") AS TOT_AmtRet, " & _
            "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.Tax_Amt+SP.Tax_AmtMRP)", "0") & ") AS Tax_Amt, " & _
            "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.Tax_Amt+SP.Tax_AmtMRP)", "0") & ") AS Tax_AmtRet, " & _
            "sum(" & cIIF("SP.V_Type Not in ('SXSRC','SXSRR')", "(SP.Tax_Sur_Amt+SP.TaxSur_AmtMRP)", "0") & ") AS Tax_Sur_Amt, " & _
            "sum(" & cIIF("SP.V_Type in ('SXSRC','SXSRR')", "(SP.Tax_Sur_Amt+SP.TaxSur_AmtMRP)", "0") & ") AS Tax_Sur_AmtRet, " & _
            "SP.Tax_Per,SP.Tax_Sur_Per " & _
            "FROM SP_Sale SP " & _
            "Where SP.V_Type In ('SYSIC','SYSIR','W_SIC','W_SIR','SXSRC','SXSRR') " & Condstr & _
            "group by SP.Form_Code,SP.Tax_Per,SP.Tax_Sur_Per"
            
    MyRst.Open mQry, GCn, adOpenDynamic, adLockOptimistic
    If PubDiscOnLube = 1 Then
        For I = 1 To MyRst.RecordCount
            RstRep.AddNew
            RstRep!TBSal = (MyRst!SprAmtTB + MyRst!SprAmtMrpTB) - (MyRst!Tax_AmtMRP + MyRst!TaxSur_AmtMRP)
            RstRep!TaxPer = MyRst!Tax_Per
            RstRep!Surchargeper = MyRst!Tax_Sur_Per
            RstRep!TaxAmt = MyRst!Tax_Amt
            RstRep!SurchargeAmt = MyRst!Tax_Sur_Amt
        MyRst.MoveNext
        Next
    Else
        For I = 1 To MyRst.RecordCount
            RstRep.AddNew
            RstRep!TBSal = (((MyRst!SprAmtTB + MyRst!SprAmtMrpTB) - (MyRst!SprAmtTBRet + MyRst!SprAmtMrpTBRet)) - ((MyRst!Tax_AmtMRP + MyRst!TaxSur_AmtMRP) - (MyRst!Tax_AmtMRPRet + MyRst!TaxSur_AmtMRPRet)))
            RstRep!TaxPer = MyRst!Tax_Per
            RstRep!Surchargeper = MyRst!Tax_Sur_Per
            RstRep!TaxAmt = MyRst!Tax_Amt - MyRst!Tax_AmtRet
            RstRep!SurchargeAmt = MyRst!Tax_Sur_Amt - MyRst!Tax_Sur_AmtRet
        MyRst.MoveNext
        Next
    End If
    MyRst.Close
    Set MyRst = Nothing
    Set MyRstRet = Nothing
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprSaleTaxCtrlStmt"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SprRateVariation()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
Dim CompOperater As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    Condstr = " and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If


    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Stock.DocId,1) in (" & GridString2 & ")"

    If FGrid.TextMatrix(List1, 1) = "High" Then
        CompOperater = " >"
    ElseIf FGrid.TextMatrix(List1, 1) = "Low" Then
        CompOperater = "<"
    Else
        CompOperater = "<>"
    End If
    
    If GRepFormName = SprCtrRateVari Then
    mQry = "SELECT SP_Stock.V_Type, SP_Stock.DocID, SP_Stock.V_Date," & _
        "SP_Stock.Qty_Iss,SP_Stock.Rate,SP_Stock.Part_No, Part.Part_Name," & _
        "" & cIIF("SP_Stock.MRP_YN = 1", "Part.MRP", cIIF("SP_Stock.Tax_YN = 1", "Part.TB_SRate", "Part.TP_SRate")) & " as PartRate, " & cIIF("SP_Stock.Tax_YN = 1", "Part.TB_Effect_Dt", "MRP_Effect_Dt") & " as EffDate " & _
        "FROM SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) " & _
        "where SP_Stock.Rate " & CompOperater & " " & cIIF("SP_Stock.MRP_YN = 1", "Part.MRP", cIIF("SP_Stock.Tax_YN = 1", "Part.TB_SRate", "Part.TP_SRate")) & " and sp_stock.v_type in ('SYSIC','SYSIR')"
    Else
     mQry = "SELECT SP_Stock.V_Type, SP_Stock.DocID, SP_Stock.V_Date, SP_Stock.DocID," & _
        "SP_Stock.Qty_Rec as Qty_Iss,SP_Stock.Rate,SP_Stock.Part_No, Part.Part_Name," & _
        "" & cIIF("SP_Stock.MRP_YN = 1", "Part.MRP - Part_DiscFactor.PurcDisc_Per", cIIF("SP_Stock.Tax_YN = 1", "Part.TB_SRate -Part_DiscFactor.PurcDisc_Per", "Part.TP_SRate - Part_DiscFactor.PurcDisc_Per")) & " as PartRate, " & cIIF("SP_Stock.Tax_YN = 1", "Part.TB_Effect_Dt", "MRP_Effect_Dt") & " as EffDate " & _
        "FROM (SP_Stock LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1)) " & _
        "LEFT JOIN Part_DiscFactor ON Part.Disc_Factor = Part_DiscFactor.DiscFac_Catg " & _
        "where SP_Stock.Rate " & CompOperater & " " & cIIF("SP_Stock.MRP_YN = 1", "Part.MRP - Part_DiscFactor.PurcDisc_Per", cIIF("SP_Stock.Tax_YN = 1", "Part.TB_SRate - Part_DiscFactor.PurcDisc_Per", "Part.TP_SRate - Part_DiscFactor.PurcDisc_Per")) & " and sp_stock.v_type in ('SXPIC','SXPIR')"
    End If
    mQry = mQry + Condstr
        
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprCtrRateVari"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprPartMove()
On Error GoTo ELoop
Dim mQry As String, Condstr As String
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_Stock.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    
    If Check1(2).Value = Unchecked Then Condstr = " and SP_Stock.Part_No in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Stock.DocId,1) in (" & GridString3 & ") and Part.Div_Code in (" & GridString3 & ") "
    
    If FGrid.TextMatrix(List1, 1) = "Yes" Then
        
        mQry = "SELECT  " & _
            " SP_Stock.Part_No, Sum(SP_Stock.Qty_Rec)-Sum(SP_Stock.Qty_Iss) AS QtyOpen,0 AS QtyRec, 0 AS QtyRet,0 as counter1, 0 AS Workshop , " & _
            " 0 as PDI,0 as Warranty,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name FROM SP_Stock " & _
            " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        " WHERE SP_Stock.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(DateAdd("D", -1, FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "" & Condstr & _
            " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All "
        mQry = mQry + " SELECT SP_Stock.Part_No, 0 as QtyOpen,Sum(SP_Stock.Qty_Rec) AS QtyRec, 0 AS QtyRet,0 as counter1, 0 AS Workshop , " & _
            " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
            " FROM SP_Stock " & _
            " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        " WHERE SP_Stock.V_Type in ('SXGR','SXGRT','SXRAD','SXPIC','SXPIR') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
            " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec, Sum(SP_Stock.Qty_Rec) AS QtyRet,0 as counter1,0 AS Workshop, " & _
        " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        " WHERE SP_Stock.V_Type in ('SXSRC','SXSRR') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All "
        mQry = mQry + " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,Sum(SP_Stock.Qty_iss) as counter1, 0 AS Workshop, " & _
        " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE SP_Stock.V_Type in ('SYSC','SYPRC','SYPRR') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, Sum(SP_Stock.Qty_iss-SP_Stock.Qty_Ret) AS Workshop, " & _
        " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " " & cIIF("SP_Stock.Purpose='P'", "Sum(SP_Stock.Qty_Iss-SP_Stock.Qty_Ret)", "0") & " AS  PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose "

         mQry = mQry + " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI, " & cIIF("SP_Stock.Purpose='W'", "Sum(SP_Stock.Qty_Iss-SP_Stock.Qty_Ret)", "0") & " as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO " & _
        " WHERE SP_Stock.V_Type in ('W_RG','W_RW','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No,        0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI,0 as Warranti, " & cIIF("SP_Stock.Purpose='F'", "Sum(SP_Stock.Qty_Iss-SP_Stock.Qty_Ret)", "0") & " as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose " & _
        " Union All "
        mQry = mQry + " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI,0 as Warranti,0 as Free," & cIIF("SP_Stock.Purpose='O'", "Sum(SP_Stock.Qty_Iss-SP_Stock.Qty_Ret)", "0") & " as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI,0 as Warranti,0 as Free,0 as CoVeh,Sum(SP_Stock.Qty_Iss-SP_Stock.Qty_Ret) as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE SP_Stock.V_Type in ('SXSRT','SYPRT') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name "

    ElseIf FGrid.TextMatrix(List1, 1) = "No" Then
        mQry = "SELECT  " & _
        " SP_Stock.Part_No, Sum(SP_Stock.Qty_Rec)-Sum(SP_Stock.Qty_Iss) AS QtyOpen,0 AS QtyRec, 0 AS QtyRet,0 as counter1, 0 AS Workshop , " & _
        "  0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name FROM SP_Stock " & _
        "  " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(DateAdd("D", -1, FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All "
        mQry = mQry + " SELECT SP_Stock.Part_No, 0 as QtyOpen,Sum(SP_Stock.Qty_Rec) AS QtyRec, 0 AS QtyRet,0 as counter1, 0 AS Workshop , " & _
        " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('SXGR','SXGRT') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec, Sum(SP_Stock.Qty_Iss) AS QtyRet,0 as counter1,0 AS Workshop, " & _
        " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('SXSRC','SXSRR') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All "
        mQry = mQry + " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,Sum(SP_Stock.Qty_iss) as counter1, 0 AS Workshop, " & _
        " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('SYSC') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, Sum(SP_Stock.Qty_iss) AS Workshop, " & _
        " 0 as PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        "  GROUP BY SP_Stock.Part_No,Part.Part_Name " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " " & cIIF("SP_Stock.Purpose='P'", "Sum(SP_Stock.Qty_Iss)", "0") & " AS  PDI,0 as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose "

         mQry = mQry + " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI," & cIIF("SP_Stock.Purpose='W'", "Sum(SP_Stock.Qty_Iss)", "0") & " as Warranti,0 as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No,        0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI,0 as Warranti," & cIIF("SP_Stock.Purpose='F'", "Sum(SP_Stock.Qty_Iss)", "0") & " as Free,0 as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose " & _
        " Union All "
        mQry = mQry + " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI,0 as Warranti,0 as Free," & cIIF("SP_Stock.Purpose='O'", "Sum(SP_Stock.Qty_Iss)", "0") & " as CoVeh,0 as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('W_RG','W_RW') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name,SP_Stock.Purpose " & _
        " Union All " & _
        " SELECT SP_Stock.Part_No, 0 as QtyOpen,0 AS QtyRec,0 AS QtyRet,0 as counter1, 0 AS Workshop, " & _
        " 0 AS  PDI,0 as Warranti,0 as Free,0 as CoVeh,Sum(SP_Stock.Qty_Iss) as Transfer,Part.Part_Name " & _
        " FROM SP_Stock " & _
        " LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO  " & _
        " WHERE Left(Sp_Stock.DocId,1)='" & PubDivCode & "' And SP_Stock.V_Type in ('SXSRT','SYPRT') and SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " " & Condstr & _
        " GROUP BY SP_Stock.Part_No,Part.Part_Name "
    End If
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "SprPartMovement"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprStkAge()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset
Dim mQry As String, Condstr As String
Dim TotQty As Double
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(Cat3, FGrid.TextMatrix(Cat3, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(Cat4, FGrid.TextMatrix(Cat4, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(Cat5, FGrid.TextMatrix(Cat5, 0)) = False Then RepPrint = False: Exit Sub
    
    If FGrid.TextMatrix(List2, 1) = "Yes" Then GridString2 = MarkRecCalculate 'Else MsgBox "No"
    If GridString2 = Empty And FGrid.TextMatrix(List2, 1) = "Yes" Then: MsgBox "** No Records Found to Print **": RepPrint = False: Exit Sub: Else If FGrid.TextMatrix(List2, 1) = "Yes" Then Condstr = Condstr & " and Stk.Part_No in (" & left(GridString2, (Len(GridString2) - 1)) & ")" 'Else MsgBox "No"
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
      
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Stk.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Stk.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Stk.Part_No in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(Stk.DocId,1) in (" & GridString3 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "QtyWise" Then
        If FGrid.TextMatrix(Cat1, 1) <> "" Then
            mQry = "SELECT Stk.Part_No, Sum(Stk.Qty_Rec)-Sum(Stk.Qty_Iss-Stk.Qty_Ret) AS OPQty,0 as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name " & _
                "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, Sum(Stk.Qty_Rec) as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name ,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        End If
        If FGrid.TextMatrix(Cat2, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,Sum(Stk.Qty_Rec)  as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat1, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat3, 1) <> "" Then
            mQry = mQry & "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,Sum(Stk.Qty_Rec)  as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name ,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat2, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat4, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,Sum(Stk.Qty_Rec)  as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat3, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat5, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,Sum(Stk.Qty_Rec) as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name ,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        End If
        
    ElseIf FGrid.TextMatrix(List1, 1) = "ValueWise" Then
        If FGrid.TextMatrix(Cat1, 1) <> "" Then
            mQry = "SELECT Stk.Part_No, Sum(Stk.Qty_Rec*PurRate)-Sum(Stk.Qty_Iss-Stk.Qty_Ret) AS OPQty,0 as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name " & _
                "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, Sum(Stk.Qty_Rec*PurRate) as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        End If
        If FGrid.TextMatrix(Cat2, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,Sum(Stk.Qty_Rec*PurRate)  as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat1, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat3, 1) <> "" Then
            mQry = mQry & "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,Sum(Stk.Qty_Rec*PurRate)  as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat2, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat4, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,Sum(Stk.Qty_Rec*PurRate)  as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat3, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat5, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,Sum(Stk.Qty_Rec*PurRate) as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        End If
    ElseIf FGrid.TextMatrix(List1, 1) = "Both" Then
        If FGrid.TextMatrix(Cat1, 1) <> "" Then
            mQry = "SELECT Stk.Part_No, Sum(Stk.Qty_Rec)-Sum(Stk.Qty_Iss-Stk.Qty_Ret) AS OPQty,0 as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name " & _
                "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, Sum(Stk.Qty_Rec) as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name ,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        End If
        If FGrid.TextMatrix(Cat2, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,Sum(Stk.Qty_Rec)  as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat1, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat3, 1) <> "" Then
            mQry = mQry & "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,Sum(Stk.Qty_Rec)  as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name ,'Q' as RepType" & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat2, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat4, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,Sum(Stk.Qty_Rec)  as Day4,0 as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat3, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat5, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,Sum(Stk.Qty_Rec) as day5,0 as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name ,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec) as LastDay,Part.Part_Name,'Q' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        End If
        
        If FGrid.TextMatrix(Cat1, 1) <> "" Then
            mQry = mQry & " UNION ALL SELECT Stk.Part_No, Sum(Stk.Qty_Rec*PurRate)-Sum(Stk.Qty_Iss-Stk.Qty_Ret) AS OPQty,0 as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date >= " & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name " & _
                "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, Sum(Stk.Qty_Rec*PurRate) as Day1,0 as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        End If
        If FGrid.TextMatrix(Cat2, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,Sum(Stk.Qty_Rec*PurRate)  as Day2,0 as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat1, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat1, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat3, 1) <> "" Then
            mQry = mQry & "Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,Sum(Stk.Qty_Rec*PurRate)  as Day3,0 as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat2, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat2, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat4, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,Sum(Stk.Qty_Rec*PurRate)  as Day4,0 as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat3, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat3, 1)
                GoTo NXT
        End If
        If FGrid.TextMatrix(Cat5, 1) <> "" Then
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,Sum(Stk.Qty_Rec*PurRate) as day5,0 as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date > " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  and Stk.V_Date<= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat4, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "  " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        Else
            mQry = mQry & " Union All " & _
                "SELECT Stk.Part_No,0 AS OPQty, 0 as Day1,0 as Day2,0 as Day3,0  as Day4,0 as day5,Sum(Stk.Qty_Rec*PurRate) as LastDay,Part.Part_Name,'V' as RepType " & _
                "FROM SP_Stock as Stk LEFT JOIN Part ON Stk.Part_No = Part.PART_NO and Part.Div_Code = left(Stk.Docid,1) " & _
                "WHERE Left(Stk.DocId,1)='" & PubDivCode & "' And Stk.V_Date <= " & ConvertDate(Format(DateAdd("D", -1 * Val((FGrid.TextMatrix(Cat5, 1))), FGrid.TextMatrix(Date1, 1)), "dd/MMM/yyyy")) & "   " & Condstr & _
                "GROUP BY Stk.Part_No,Part.Part_Name "
                mDays = FGrid.TextMatrix(Cat4, 1)
        End If
    End If
NXT:
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "SprStkAge"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub SprDailySaleRegFunc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, mQRY1 As String

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " Where SP_Sale.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Sale.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Sp_Sale.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Sale.DocId,1) in (" & GridString2 & ")"
    
    If FGrid.TextMatrix(List1, 1) = "Workshop" Then
        If FGrid.TextMatrix(List2, 1) = "Cash" Then
            Condstr = Condstr & " and SP_Sale.V_Type in ('W_SIC','W_WWC')"
        ElseIf FGrid.TextMatrix(List2, 1) = "Credit" Then
            Condstr = Condstr & " and SP_Sale.V_Type in ('W_SIR','W_WWR')"
        Else
            Condstr = Condstr & " and SP_Sale.V_Type in ('W_SIC','W_SIR','W_WWC','W_WWR')"
        End If
    ElseIf FGrid.TextMatrix(List1, 1) = "Counter" Then
        If FGrid.TextMatrix(List2, 1) = "Cash" Then
            Condstr = Condstr & " and SP_Sale.V_Type in ('SYSIC')"
        ElseIf FGrid.TextMatrix(List2, 1) = "Credit" Then
            Condstr = Condstr & " and SP_Sale.V_Type in ('SYSIR')"
        Else
            Condstr = Condstr & " and SP_Sale.V_Type in ('SYSIR','SYSIC')"
        End If
    Else
        If FGrid.TextMatrix(List2, 1) = "Cash" Then
            Condstr = Condstr & " and SP_Sale.V_Type in ('W_SIC','SYSIC','W_WWC')"
        ElseIf FGrid.TextMatrix(List2, 1) = "Credit" Then
            Condstr = Condstr & " and SP_Sale.V_Type in ('W_SIR','SYSIR','W_WWR')"
        Else
            Condstr = Condstr & " and SP_Sale.V_Type in ('W_SIC','SYSIC','W_SIR','SYSIR','W_WWC','W_WWR')"
        End If
    End If
    ' "SYSIC","SYSIR","W_SIC","W_SIR"
    mQry = "SELECT " & cMID("SP_Sale.DocID", "4", "5") & " as V_Type,SP_Sale.DocID,SP_Sale.V_Date,SP_sale.Job_DocID,SP_Sale.Cash_Credit,SubGroup.Name," & _
        "SP_Sale.SprAmt_MRP_TB + SP_Sale.OilAmt_MRP_TB + SP_Sale.SprAmt_TB + SP_Sale.OilAmt_TB - SP_Sale.D_Amt_MRP_TB - SP_Sale.D_Amt_TB as TaxableAmt," & _
        "SP_Sale.SprAmt_MRP_TP + SP_Sale.OilAmt_MRP_TP + SP_Sale.SprAmt_TP + SP_Sale.OilAmt_TP - SP_Sale.D_Amt_MRP_TP - SP_Sale.D_Amt_TP as TaxpaidAmt, " & _
        "SP_Sale.Gen_Sur_Amt +  SP_Sale.Tax_Amt + SP_Sale.Tax_Sur_Amt  as Tax, SP_Sale.Tax_AmtMRP +SP_Sale.TaxSur_AmtMRP as TaxOnMRP, " & _
        "SP_Sale.Total_Amt,J.NetLab_Amt,J.DocId_InvLab,  " & _
        "SP_Sale.SprAmt_MRP_TB + Sp_Sale.SprAmt_Mrp_TP +  SP_Sale.SprAmt_TB + Sp_Sale.SprAmt_TP as SpareAmt," & _
        "SP_Sale.OilAmt_MRP_TB + Sp_Sale.OilAmt_Mrp_TP + SP_Sale.OilAmt_TB + Sp_Sale.OilAmt_TP as OilAmt," & _
        "SP_Sale.D_Amt_TB + SP_Sale.D_Amt_TP as Discount, Sp_Sale.Packing As Misc_Chg " & _
        "FROM ((SP_Sale left join SubGroup on SubGroup.SubCode=SP_Sale.Party_Code) " & _
        "left join Job_Card as J on SP_Sale.Job_DocId=J.DocID) " & Condstr
    mQry = mQry & " Order By SP_Sale.DocId "
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "SprSalRegDaily"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Function MarkRecCalculate()
     Dim formulastr As String
     Set RsMark = New Recordset
     RsMark.Open "Select PART_NO from Part where MARK_YN ='Y'", GCn, adOpenDynamic, adLockOptimistic, adCmdText
     While Not RsMark.EOF = True
        formulastr = formulastr & "'" & RsMark!Part_No & "',": RsMark.MoveNext
     Wend
     Set RsMark = Nothing
     MarkRecCalculate = formulastr
End Function

Private Sub SprSaleTransfer()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String
'Date1,Date2,List1,List1,List1,List2,List1,List1
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

    If FGrid.TextMatrix(List1, 1) = "All" Then Condstr = "SP_Sale.V_Type In ('" & SprSlChal & "','" & SprTrfChal & "') And "
    If FGrid.TextMatrix(List1, 1) = "Transfer" Then Condstr = "SP_Sale.V_Type = '" & SprTrfChal & "' And "
    If FGrid.TextMatrix(List1, 1) = "Sale Challan" Then Condstr = "SP_Sale.V_Type = '" & SprSlChal & "' And "
    If FGrid.TextMatrix(List1, 1) = "Pending" Then Condstr = "SP_Sale.V_Type In ('" & SprSlChal & "','" & SprTrfChal & "') And (SP_Sale.Invoice_Docid='' or len(SP_Sale.Invoice_Docid)<=0) And "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " " & cMID("SP_Sale.DocId", "3", "1") & " in (" & GridString1 & ") AND "
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & "  " & cMID("SP_Sale.DocId", "3", "1") & " ='" & PubSiteCode & "' and "
    End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " SP_Sale.Party_Code in (" & GridString2 & ") AND "
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " Left(SP_Sale.DocId,1) in (" & GridString3 & ") AND "
            
    Condstr = Condstr + "SP_Sale.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Sale.V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
        
If FGrid.TextMatrix(List2, 1) = "Summary" Then
        RepName = "SprSalTrfReg"
        If StrCmp(left(PubComp_Name, 4), "Yash") Then
            mQry = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type,(left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, " & _
                "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB,SP_Sale.SprAmt_MRP_TP, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB ) - ((SP_Sale.SprAmt_TB)* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB))))") & " + " & cIIF("((SP_Sale.SprAmt_Mrp_TB)+(SP_Sale.OilAmt_Mrp_TB ))=0", "0", "((SP_Sale.SprAmt_Mrp_TB ) - ((SP_Sale.SprAmt_Mrp_TB)* SP_Sale.D_Amt_MRP_TB / ((SP_Sale.SprAmt_Mrp_TB)+(SP_Sale.OilAmt_Mrp_TB))))") & " AS SprAmtTB, " & _
                "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.SprAmt_TP) - ((SP_Sale.SprAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))))") & " + " & cIIF("((SP_Sale.SprAmt_Mrp_TP)+(SP_Sale.OilAmt_Mrp_TP ))=0", "0", "((SP_Sale.SprAmt_Mrp_TP ) - ((SP_Sale.SprAmt_Mrp_TP)* SP_Sale.D_Amt_MRP_TP / ((SP_Sale.SprAmt_Mrp_TP)+(SP_Sale.OilAmt_Mrp_TP))))") & " AS SprAmtTP, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB) - ((SP_Sale.OilAmt_TB )* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB))))") & " + " & cIIF("((SP_Sale.SprAmt_Mrp_TB)+(SP_Sale.OilAmt_Mrp_TB ))=0", "0", "((SP_Sale.OilAmt_Mrp_TB) - ((SP_Sale.OilAmt_Mrp_TB )* SP_Sale.D_Amt_MRP_TB / ((SP_Sale.SprAmt_Mrp_TB)+(SP_Sale.OilAmt_Mrp_TB))))") & " as OilAmtTB  , " & _
                "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.OilAmt_TP) - ((SP_Sale.OilAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP))))") & " + " & cIIF("((SP_Sale.SprAmt_Mrp_TP)+(SP_Sale.OilAmt_Mrp_TP ))=0", "0", "((SP_Sale.OilAmt_Mrp_TP) - ((SP_Sale.OilAmt_Mrp_TP)* SP_Sale.D_Amt_MRP_TP / ((SP_Sale.SprAmt_Mrp_TP)+(SP_Sale.OilAmt_Mrp_TP))))") & " as OilAmtTP , " & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TB+OilAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB+OilAmt_MRP_TB) - (SP_Sale.D_Amt_MRP_TB))") & " AS SprAmtMRPTB, " & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP+OilAmt_MRP_TP)=0", "0", "((SP_Sale.SprAmt_MRP_TP+OilAmt_MRP_TP) - (SP_Sale.D_Amt_MRP_TP))") & " AS SprAmtMRPTP, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " AS SprTransTB, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " as OilTransTB , " & _
                "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                "SP_Sale.Tax_Amt as TaxAmt,SP_Sale.Tax_Sur_Amt AS Tax_Sur_Amt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                "SP_Sale.Tax_AmtMRP AS Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP AS TaxSur_AmtMRP,SP_Sale.TOT_AmtMRP as TOT_AmtMRP," & _
                "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt " & _
                " FROM (SP_Sale LEFT JOIN SubGroup ON SubGroup.SubCode=SP_Sale.Party_Code) Where Left(Sp_Sale.DocId,1)='" & PubDivCode & "' And " & Condstr & ""
        Else
            mQry = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type,(left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, " & _
                "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB,SP_Sale.SprAmt_MRP_TP, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB ) - ((SP_Sale.SprAmt_TB)* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB))))") & " AS SprAmtTB, " & _
                "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.SprAmt_TP) - ((SP_Sale.SprAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))))") & " AS SprAmtTP, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB) - ((SP_Sale.OilAmt_TB )* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB))))") & " as OilAmtTB  , " & _
                "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.OilAmt_TP) - ((SP_Sale.OilAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP))))") & " as OilAmtTP , " & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TB+OilAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB+OilAmt_MRP_TB) - (SP_Sale.D_Amt_MRP_TB))") & " AS SprAmtMRPTB, " & _
                "" & cIIF("(SP_Sale.SprAmt_MRP_TP+OilAmt_MRP_TP)=0", "0", "((SP_Sale.SprAmt_MRP_TP+OilAmt_MRP_TP) - (SP_Sale.D_Amt_MRP_TP))") & " AS SprAmtMRPTP, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " AS SprTransTB, " & _
                "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " as OilTransTB , " & _
                "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
                "SP_Sale.Tax_Amt as TaxAmt,SP_Sale.Tax_Sur_Amt AS Tax_Sur_Amt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
                "SP_Sale.Tax_AmtMRP AS Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP AS TaxSur_AmtMRP,SP_Sale.TOT_AmtMRP as TOT_AmtMRP," & _
                "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt " & _
                " FROM (SP_Sale LEFT JOIN SubGroup ON SubGroup.SubCode=SP_Sale.Party_Code) Where Left(Sp_Sale.DocId,1)='" & PubDivCode & "' And " & Condstr & ""
        
        End If
Else
    RepName = "SprSalTrfRegDet"

        mQry = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type,(left(SP_Sale.Docid,1)+ " & cMID("SP_Sale.Docid", "3", "2") & " + " & cMID("SP_Sale.Docid", "8", "1") & " + " & cCStr("SP_Sale.V_No") & ") as V_No, " & _
            "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB,SP_Sale.SprAmt_MRP_TP, " & _
            "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB ) - ((SP_Sale.SprAmt_TB)* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB))))") & " AS SprAmtTB, " & _
            "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.SprAmt_TP) - ((SP_Sale.SprAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))))") & " AS SprAmtTP, " & _
            "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB) - ((SP_Sale.OilAmt_TB )* (SP_Sale.D_Amt_TB-SP_Sale.D_Amt_MRP_TB) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB))))") & " as OilAmtTB , " & _
            "" & cIIF("((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP ))=0", "0", "((SP_Sale.OilAmt_TP) - ((SP_Sale.OilAmt_TP)* (SP_Sale.D_Amt_TP-SP_Sale.D_Amt_MRP_TP) / ((SP_Sale.SprAmt_TP)+(SP_Sale.OilAmt_TP))))") & " as OilAmtTP , " & _
            "" & cIIF("(SP_Sale.SprAmt_MRP_TB+OilAmt_MRP_TB)=0", "0", "((SP_Sale.SprAmt_MRP_TB+OilAmt_MRP_TB) - (SP_Sale.D_Amt_MRP_TB))") & " AS SprAmtMRPTB, " & _
            "" & cIIF("(SP_Sale.SprAmt_MRP_TP+OilAmt_MRP_TP)=0", "0", "((SP_Sale.SprAmt_MRP_TP+OilAmt_MRP_TP) - (SP_Sale.D_Amt_MRP_TP))") & " AS SprAmtMRPTP, " & _
            "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.SprAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " AS SprTransTB, " & _
            "" & cIIF("((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB ))=0", "0", "((SP_Sale.OilAmt_TB)* (SP_Sale.Trans_Amt) / ((SP_Sale.SprAmt_TB)+(SP_Sale.OilAmt_TB)))") & " as OilTransTB , " & _
            "SP_Sale.D_Amt_TB,SP_Sale.D_Amt_TP,SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt," & _
            "SP_Sale.Tax_Amt as TaxAmt,SP_Sale.Tax_Sur_Amt AS Tax_Sur_Amt, SP_Sale.TOT_Amt,SP_Sale.ReSalTax_Amt," & _
            "SP_Sale.Tax_AmtMRP AS Tax_AmtMRP,SP_Sale.TaxSur_AmtMRP AS TaxSur_AmtMRP,SP_Sale.TOT_AmtMRP as TOT_AmtMRP," & _
            "SP_Sale.Packing,SP_Sale.Rounded, SP_Sale.Total_Amt,SP_Stock.Part_No,SP_Stock.Qty_Iss,SP_Stock.Rate,SP_Stock.Amount,Part.Part_Name  " & _
            " FROM ((SP_Sale LEFT JOIN SubGroup ON SubGroup.SubCode=SP_Sale.Party_Code) Left join SP_Stock on SP_Sale.DocId=SP_Stock.DocId) Left Join Part on SP_Stock.Part_No=Part.Part_No Where Left(Sp_Sale.DocId,1)='" & PubDivCode & "' And " & Condstr & ""
End If
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepTitle = UCase(Me.CAPTION) + "[" + FGrid.TextMatrix(List1, 1) + "]"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SpeedPrnStkInHnd()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double
    Dim isLast As Boolean, mRec As Integer, PageNo As Double
    Dim TotalTBVal As Double, TotalTPVal As Double, TotalVal As Double
    Dim TotalTBQty As Double, TotalTPQty As Double, TotalQty As Double
    Dim GTotalTBVal As Double, GTotalTPVal As Double, GTotalVal As Double
    Dim GTotalTBQty As Double, GTotalTPQty As Double, GTotalQty As Double
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
    PageWidth = 80
    mRec = 45
    'Header printing
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
'    If XNull(RstCompDet!S_SecSpeciality) <> "" Then
'        Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
'        mHeader = mHeader + 1
'    End If
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
    
    Print #1, PRN_TIT("Stock In Hand", "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, "From : " & FGrid.TextMatrix(Date1, 1) & "  To : " & FGrid.TextMatrix(Date2, 1)
    mHeader = mHeader + 1
    Print #1, "For MRP Parts : " & FGrid.TextMatrix(List2, 1)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("#", 10) & PSTR("Part No.", 22) & PSTR("Part Name", 28) & PSTR("TB Qty", 10, , AlignRight) & PSTR("TP Qty", 10, , AlignRight) & PSTR("TotalQty", 10, , AlignRight) & PSTR("Rate", 10, , AlignRight) & PSTR("TB Val", 12, , AlignRight) & PSTR("TP Val", 12, , AlignRight) & PSTR("Total Val", 16, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    RstRep.Sort = "Part_No"
    RstRep.MoveFirst
    mHeader = 1
    Dim StartCounter As Double, I As Double, LastPartNo$, LastPartName$, mRate As Double
    While Not RstRep.EOF = True
        mRate = VNull(RstRep!LPRate)
        If mHeader <= mRec Then
            If RstRep.EOF = True Then GoTo aa
            For I = StartCounter To RstRep.RecordCount
                If LastPartNo <> "" And LastPartNo <> RstRep!Part_No Then Exit For
                TotalTBQty = TotalTBQty + (Val(RstRep!TBQtyRec) - Val(RstRep!TBQtyIss))
                TotalTPQty = TotalTPQty + (Val(RstRep!TPQtyRec) - Val(RstRep!TPQtyIss))
                TotalQty = TotalTBQty + TotalTPQty
                
                TotalTBVal = TotalTBVal + ((Val(RstRep!TBQtyRec) - Val(RstRep!TBQtyIss)) * Val(RstRep!LPRate))
                TotalTPVal = TotalTPVal + ((Val(RstRep!TPQtyRec) - Val(RstRep!TPQtyIss)) * Val(RstRep!LPRate))
                TotalVal = (TotalTBVal + TotalTPVal)
                StartCounter = StartCounter + 1
                LastPartNo = RstRep!Part_No: LastPartName = RstRep!Part_Name
                RstRep.MoveNext
                If RstRep.EOF = True Then Exit For
            Next
            Counter = Counter + 1
            Print #1, mChr17 & PSTR(STR(Counter), 10) & PSTR(LastPartNo, 22) & PSTR(LastPartName, 28) & PSTR(IIf(TotalTBQty = 0, "", STR(Format(TotalTBQty, ".00"))), 10, , AlignRight) & PSTR(IIf(TotalTPQty = 0, "", STR(Format(TotalTPQty, ".00"))), 10, , AlignRight) & PSTR(IIf(TotalQty = 0, "", STR(Format(TotalQty, ".00"))), 10, , AlignRight) & PSTR(IIf(TotalQty = 0, "", STR(Format(mRate, ".00"))), 10, , AlignRight) & PSTR(IIf(TotalTBVal = 0, "", STR(Format(TotalTBVal, ".00"))), 12, , AlignRight) & PSTR(IIf(TotalTPVal = 0, "", STR(Format(TotalTPVal, ".00"))), 12, , AlignRight) & PSTR(IIf(TotalVal = 0, "", STR(Format(TotalVal, ".00"))), 16, , AlignRight) & mChr18
            mHeader = mHeader + 1
            GTotalTBQty = GTotalTBQty + TotalTBQty: GTotalTPQty = GTotalTPQty + TotalTPQty: GTotalQty = GTotalQty + TotalQty
            GTotalTBVal = GTotalTBVal + TotalTBVal: GTotalTPVal = GTotalTPVal + TotalTPVal: GTotalVal = GTotalVal + TotalVal
            
            TotalTBQty = 0: TotalTBVal = 0: TotalTPQty = 0: TotalTPVal = 0: TotalQty = 0: TotalVal = 0
            LastPartNo = "": LastPartName = ""
            If mHeader = mRec Then isLast = True
        Else
            If isLast Then
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = 0
                isLast = False
                Print #1, Space(PageWidth / 2) & "Page :" & PageNo + 1
                PageNo = PageNo + 1
                Print #1, mEject
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT("Stock In Hand", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & PSTR("#", 10) & PSTR("Part No.", 22) & PSTR("Part Name", 28) & PSTR("TB Qty", 10, , AlignRight) & PSTR("TP Qty", 10, , AlignRight) & PSTR("TotalQty", 10, , AlignRight) & PSTR("TB Val", 12, , AlignRight) & PSTR("TP Val", 12, , AlignRight) & PSTR("Total Val", 16, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
            End If
        End If
    Wend
aa:
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("Total-- >", 10) & Space(22) & Space(28) & PSTR(IIf(GTotalTBQty = 0, "", STR(Format(GTotalTBQty, ".00"))), 10, , AlignRight) & PSTR(IIf(GTotalTPQty = 0, "", STR(Format(GTotalTPQty, ".00"))), 10, , AlignRight) & PSTR(IIf(GTotalQty = 0, "", STR(Format(GTotalQty, ".00"))), 10, , AlignRight) & Space(10) & PSTR(IIf(GTotalTBVal = 0, "", STR(Format(GTotalTBVal, ".00"))), 12, , AlignRight) & PSTR(IIf(GTotalTPVal = 0, "", STR(Format(GTotalTPVal, ".00"))), 12, , AlignRight) & PSTR(IIf(GTotalVal = 0, "", STR(Format(GTotalVal, ".00"))), 16, , AlignRight) & mChr18
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
Private Sub SpeedPrnStkLedger()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double
    Dim isLast As Boolean, mRec As Integer, PageNo As Double
    Dim TotalTBRec As Double, TotalTBIss As Double, TotalTBBal As Double
    Dim TotalTPRec As Double, TotalTPIss As Double, TotalTPBal As Double
    Dim RstCompDet As ADODB.Recordset
    Dim VDate As String, mDocId As String, Vdate1 As String
    Dim fob As New FileSystemObject
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    PageLength = PubPageLength
    PageWidth = 80
    mRec = 45
    'Header printing
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
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
    
    Print #1, PRN_TIT("Stock Ledger", "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, "From : " & FGrid.TextMatrix(Date1, 1) & "  To : " & FGrid.TextMatrix(Date2, 1)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("Part No & Part Name", 33) & Chr(27) & Chr(45) & Chr(1) & Space(4) & Chr(27) & Chr(45) & Chr(0) & Chr(27) & Chr(45) & Chr(1) & PSTR("Recieved", 19) & Chr(27) & Chr(45) & Chr(0) & Space(1) & Chr(27) & Chr(45) & Chr(1) & PSTR("Issued", 19) & Chr(27) & Chr(45) & Chr(0) & Space(1) & Chr(27) & Chr(45) & Chr(1) & PSTR("Balance", 15) & Chr(27) & Chr(45) & Chr(0) & PSTR("Job Inv. No", 20, , AlignLeft) & PSTR("Purpose", 7, , AlignRight) & PSTR("Supplier", 10, , AlignRight) & PSTR("Supplier", 10, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("V_Date", 12) & PSTR("V_No", 21) & PSTR("TB Qty", 10) & PSTR("TP Qty", 10) & PSTR("TB Qty", 10) & PSTR("TP Qty", 10) & PSTR("TB Qty", 10) & PSTR("TP Qty", 10) & Space(20) & Space(7) & PSTR("Inv.No", 10, , AlignRight) & PSTR("Inv.Date", 10, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    RstRep.Sort = "Part_No,V_No"
    RstRep.MoveFirst
    mHeader = 1
    Dim StartCounter As Double, I As Double, LastPartNo$, LastPartName$
    For I = 1 To RstRep.RecordCount
        If mHeader <= mRec Then
            If RstRep.EOF = True Then GoTo aa
                If LastPartNo <> RstRep!Part_No Then
                    TotalTBRec = 0: TotalTBIss = 0: TotalTPRec = 0
                    TotalTPIss = 0: TotalTBBal = 0: TotalTPBal = 0
                    Print #1, mChr17 & mDoub & PSTR(RstRep!Part_No & "   " & RstRep!Part_Name, 50) & Space(107) & mDoub1 & mChr18
                    mHeader = mHeader + 1
                End If
                    TotalTBRec = TotalTBRec + Val(VNull(RstRep!TBQtyRec))
                    TotalTBIss = TotalTBIss + Val(VNull(RstRep!TBQtyIss))
                    TotalTPRec = TotalTPRec + Val(VNull(RstRep!TPQtyRec))
                    TotalTPIss = TotalTPIss + Val(VNull(RstRep!TPQtyIss))
                    TotalTBBal = TotalTBRec - TotalTBIss
                    TotalTPBal = TotalTPRec - TotalTPIss
                    VDate = RstRep!V_DATE
                    Vdate1 = XNull(RstRep!Party_Doc_Date)
                    If XNull(RstRep!job_docid) = "" Then
                        mDocId = ""
                    Else
                        mDocId = PrinID(RstRep!job_docid)
                    End If
                    Print #1, mChr17 & PSTR(VDate, 12) & Space(2) & PSTR(PrinID(RstRep!DocID), 22) & PSTR(IIf(VNull(RstRep!TBQtyRec) = 0, "", STR(RstRep!TBQtyRec)), 10) & PSTR(IIf(VNull(RstRep!TPQtyRec) = 0, "", STR(RstRep!TPQtyRec)), 11) & PSTR(IIf(VNull(RstRep!TBQtyIss) = 0, "", STR(RstRep!TBQtyIss)), 11) & PSTR(IIf(VNull(RstRep!TPQtyIss) = 0, "", STR(RstRep!TPQtyIss)), 10) & PSTR(IIf(TotalTBBal = 0, "", STR(TotalTBBal)), 9) & PSTR(IIf(TotalTPBal = 0, "", STR(TotalTPBal)), 9) & PSTR(mDocId, 15, , AlignRight) & PSTR(left(RstRep!SprPurPose, 1), 7, , AlignRight) & PSTR(VNull(RstRep!Party_Doc_No), 10, , AlignRight) & PSTR(VDate, 12, , AlignRight) & mChr18
                    mHeader = mHeader + 1
                    
                    If RstRep.EOF = True Then Exit For
                    LastPartNo = RstRep!Part_No: LastPartName = RstRep!Part_Name
                    RstRep.MoveNext
                    If mHeader = mRec Then isLast = True
        Else
            If isLast Then
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = 0
                isLast = False
                Print #1, Space(PageWidth / 2) & "Page :" & PageNo + 1
                PageNo = PageNo + 1
                Print #1, mEject
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT("Stock Ledger", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & PSTR("Part No & Part Name", 33) & Chr(27) & Chr(45) & Chr(1) & PSTR("Recieved", 20) & Chr(27) & Chr(45) & Chr(0) & Chr(27) & Chr(45) & Chr(1) & PSTR("Issued", 20) & Chr(27) & Chr(45) & Chr(0) & Chr(27) & Chr(45) & Chr(1) & PSTR("Balance", 20) & Chr(27) & Chr(45) & Chr(0) & PSTR("Job Inv. No", 20, , AlignRight) & PSTR("Purpose", 7, , AlignRight) & PSTR("Supplier", 10, , AlignRight) & PSTR("Supplier", 10, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, mChr17 & PSTR("V_Date", 12) & PSTR("V_Date", 21) & PSTR("TP Qty", 10) & PSTR("TP Qty", 10) & PSTR("TP Qty", 10) & PSTR("TP Qty", 10) & PSTR("TP Qty", 10) & PSTR("TP Qty", 10) & Space(20) & Space(7) & PSTR("Inv.No", 10, , AlignRight) & PSTR("Inv.Date", 10, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
            End If
        End If
    Next
aa:
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
'    Print #1, mChr17 & PSTR("Total-- >", 10) & Space(22) & Space(28) & PSTR(str(Format(GTotalTBQty, ".00")), 10, , AlignRight) & PSTR(str(Format(GTotalTPQty, ".00")), 10, , AlignRight) & PSTR(str(Format(GTotalQty, ".00")), 10, , AlignRight) & PSTR(str(Format(GTotalTBVal, ".00")), 12, , AlignRight) & PSTR(str(Format(GTotalTPVal, ".00")), 12, , AlignRight) & PSTR(str(Format(GTotalVal, ".00")), 16, , AlignRight) & mChr18
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
Private Sub InputTaxRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset
'Date1,Date2,List1,List1,List1,List2,List1,List1
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

Condstr = "where SP.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
 If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
 
If Check1(1).Value = Unchecked Then Condstr = Condstr & "  and " & cMID("SP.DocId", "3", "1") & " in (" & GridString1 & ")  "
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & "  and " & cMID("SP.DocId", "3", "1") & " ='" & PubSiteCode & "'  "
    End If


If FGrid.TextMatrix(List1, 1) = "Purchase" Then
    If UCase(left(PubComp_Name, 5)) = "SAMAL" Then
        mQry = "Select SP.DocId,SP.V_No,TF.Tax_Per as TaxPer,Stk.TaxAmt,SG.Name,SG.LSTNo,SP.Party_Doc_No,Stk.Net_Amt2 as Net_Amt,Stk.Net_Amt2 as Amount,SP.V_Date,SP.Party_Doc_Date, Stk.SatPer, Stk.SatAmt " & _
            " From ((SP_Purch as SP Left Join SP_Stock as Stk on SP.DocID=Stk.Invoice_DocID)" & _
            " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
            " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code "
            
        mQry = mQry & Condstr
        mQry = mQry & " and SP.V_Type IN ('SXPIR','SXPIC')"
    Else
        mQry = "Select SP.DocId,SP.V_No,TF.Tax_Per as TaxPer,SP.Tax_Amt,SG.Name,SG.LSTNo,SP.Party_Doc_No,Sp.Net_Amt as Net_Amt,SP.Tot_Goods_Value as Amount,SP.V_Date,SP.Party_Doc_Date, Sg.Add1, Sg.Add2, Sg.Add3, City.CityName,SP.SatAmt " & _
            " From ((SP_Purch as SP " & _
            " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
            " Left Join City on City.CityCode=SG.CityCode) " & _
            " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code "
        
'        mQry = "Select SP.DocId,SP.V_No,Stk.TaxPer as TaxPer,Stk.TaxAmt,SG.Name,SG.LSTNo,SP.Party_Doc_No,Sp.Net_Amt as Net_Amt,Stk.Net_Amt2 as Amount,SP.V_Date,SP.Party_Doc_Date, Sg.Add1, Sg.Add2, Sg.Add3, City.CityName, Stk.SatPer, Stk.SatAmt " & _
'            " From (((SP_Purch as SP Left Join SP_Stock as Stk on SP.DocID=Stk.Invoice_DocID)" & _
'            " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
'            " Left Join City on City.CityCode=SG.CityCode) " & _
'            " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code "
        
        
        mQry = mQry & Condstr
        mQry = mQry & " and SP.V_Type IN ('SXPIR','SXPIC') and SP.L_C='L'"
    End If
Else
    mQry = "Select distinct SP.DocId,SP.V_No.V_No,Stk.TaxPer as TaxPer,Stk.TaxAmt as TaxAmt,SG.Name,SG.LSTNo,SP.Party_Doc_No,SP.Net_Amt,SP.Net_Amt as Amount, Sg.Add1, Sg.Add2, Sg.Add3, City.CityName,SP.SatAmt " & _
        " From (((SP_Purch as SP Left Join SP_Stock as Stk on SP.DocID=Stk.Invoice_DocID)" & _
        " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
        " Left Join City on City.CityCode=SG.CityCode) " & _
        " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code "
    mQry = mQry & Condstr
    mQry = mQry & " and SP.V_Type IN ('SYPRC','SYPRR')"
End If

Set RstRep = New ADODB.Recordset
RstRep.CursorLocation = adUseClient
RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    If StrCmp(left(PubComp_Name, 4), "Yash") Then
        RepName = "InputTaxReg_Yash"
    Else
        RepName = "InputTaxReg"
    End If
Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub PurTaxSummProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset




If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub



Condstr = "where SP.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""



If FGrid.TextMatrix(List2, 1) = "Local" Then
    Condstr = Condstr & " And Sp.L_C = 'L'"
ElseIf FGrid.TextMatrix(List2, 1) = "Central" Then
    Condstr = Condstr & " And Sp.L_C = 'C'"
End If
   If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
  
If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("sp.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("sp.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    


mQry = "Select SP.DocId,SP.V_No,TF.Tax_Per,SP.Tax_Amt, SP.Tot_Goods_Value, Sp.Addition, " & _
    "Sp.Deduction, SG.Name + ', ' + " & xIsNull("C.CityName", "") & " As Name,SG.LSTNo, " & _
    " SP.Party_Doc_No,Sp.Net_Amt,SP.V_Date,SP.Party_Doc_Date,SP.Sat_YN,SP.SatAmt " & _
    " From (((SP_Purch as SP " & _
    " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
    " Left Join City C On C.CityCode = SG.CityCode) " & _
    " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code) "
mQry = mQry & Condstr



If FGrid.TextMatrix(List1, 1) = "Purchase" Then
    mQry = mQry & " and SP.V_Type IN ('SXPIR','SXPIC')"
Else
    mQry = mQry & " and SP.V_Type IN ('SYPRC','SYPRR')"
End If



Set RstRep = New ADODB.Recordset
RstRep.CursorLocation = adUseClient
RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    
    
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "PurTaxSumm"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub SaleTaxSummProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset




If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = "where SP.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
     
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(SP.DocId,1) in (" & GridString2 & ")"






If FGrid.TextMatrix(List2, 1) = "Local" Then
    Condstr = Condstr & " And Sp.L_C = 'L'"
ElseIf FGrid.TextMatrix(List2, 1) = "Local" Then
    Condstr = Condstr & " And Sp.L_C = 'C'"
End If



mQry = "Select SP.DocId, Sp.V_Type, SP.V_No, TF.Tax_Per, SP.Tax_Amt, SP.Total_Amt, Sp.Addition, " & _
    "Sp.D_Amt_TB + Sp.D_Amt_TP  As Discount, " & _
    "Sp.SprAmt_TB + Sp.SprAmt_TP + Sp.SprAmt_Mrp_TB + Sp.SprAmt_Mrp_TP As SpareAmt, " & _
    "Sp.OilAmt_TB + Sp.OilAmt_TP+ Sp.OilAmt_Mrp_TB + Sp.OilAmt_Mrp_TP As OilAmt, " & _
    "SG.Name + ', ' + " & xIsNull("C.CityName", "") & " As Name, SG.LSTNo, SP.V_Date, V.Description As VoucherDesc, " & _
    "(Sp.SprAmt_TB + Sp.SprAmt_TP + Sp.SprAmt_Mrp_TB + Sp.SprAmt_Mrp_TP + Sp.OilAmt_TB + Sp.OilAmt_TP+ Sp.OilAmt_Mrp_TB + Sp.OilAmt_Mrp_TP) As GrossTotal, Sp.SatAmt " & _
    "From (((SP_Sale as SP " & _
    "Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
    "Left Join City C On C.CityCode = SG.CityCode) " & _
    "Left Join TaxForms TF on SP.Form_Code=TF.Form_Code) " & _
    "Left Join " & FaTable("Voucher_Type") & "  V  On V.V_Type = Sp.V_Type "
mQry = mQry & Condstr & " And (Sp.SprAmt_TB + Sp.SprAmt_TP + Sp.SprAmt_Mrp_TB + Sp.SprAmt_Mrp_TP+Sp.OilAmt_TB + Sp.OilAmt_TP+ Sp.OilAmt_Mrp_TB + Sp.OilAmt_Mrp_TP)>0 "



If FGrid.TextMatrix(List1, 1) = "Sale" Then
    mQry = mQry & " and SP.V_Type IN ('" & SprSlCre & "', '" & SprSlCsh & "','" & WksSlCre & "','" & WksSlCsh & "')"
Else
    mQry = mQry & " and SP.V_Type IN ('" & SprSlRetCre & "','" & SprSlRetCsh & "')"
End If

mQry = mQry & " Order By SP.DocId"

Set RstRep = New ADODB.Recordset
RstRep.CursorLocation = adUseClient
RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    
    
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "SaleTaxSumm"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub ProcSpareSaleAccount()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset




If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " And S.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and S.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
     
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("S.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("S.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(S.DocId,1) in (" & GridString2 & ")"






    If FGrid.TextMatrix(List2, 1) = "Local" Then
        Condstr = Condstr & " And S.L_C = 'L'"
    ElseIf FGrid.TextMatrix(List2, 1) = "Central" Then
        Condstr = Condstr & " And S.L_C = 'C'"
    End If




    If FGrid.TextMatrix(List1, 1) = "Sale" Then
        Condstr = Condstr & " and S.V_Type IN ('" & SprSlCre & "', '" & SprSlCsh & "','" & WksSlCre & "','" & WksSlCsh & "')"
    Else
        Condstr = Condstr & " and S.V_Type IN ('" & SprSlRetCre & "','" & SprSlRetCsh & "')"
    End If

    mQry = "SELECT  Max(S.V_Date) AS Date, Max(S.V_Type) as SaleType, Max(Sg.Name) As PartyName, S.DocID, Max(S.V_No) AS BillNo,  " & _
         "sum(CASE WHEN Stk.TaxPer=14 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_14,  " & _
         "sum (CASE WHEN Stk.TaxPer=10 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_10, " & _
         "sum(CASE WHEN Stk.TaxPer=5 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_5, " & _
         "sum(CASE WHEN Stk.TaxPer=0 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_0, " & _
         "Sum(Stk.Net_Amt2) AS VAT_ASSESABLE, " & _
         "sum (CASE WHEN Stk.TaxPer=14 Then stk.TaxAmt ELSE 0 End) AS VAT_14, " & _
         "sum (CASE WHEN Stk.TaxPer=10 Then stk.TaxAmt ELSE 0 End) AS VAT_10, " & _
         "sum (CASE WHEN Stk.TaxPer=5 Then stk.TaxAmt ELSE 0 End) AS VAT_5, " & _
         "sum (CASE WHEN Stk.TaxPer=0 Then stk.TaxAmt ELSE 0 End) AS VAT_0, " & _
         "Sum(Stk.TaxAmt) AS Tax_Amt, Sum(Stk.Amount2) AS TotalAmount, " & _
         "IsNull(Sum(J.LabAmt_TB + J.LabAmt_TP - J.Lab_D_Amt),0) AS LabourCharges, " & _
         "IsNull(Sum(J.Lab_TaxAmt),0) AS ServiceTax, " & _
         "IsNull(Sum(J.LabAmt_TB + J.LabAmt_TP - J.Lab_D_Amt+J.Lab_TaxAmt),0) AS GrossLabourAmt " & _
         "FROM sp_sale S " & _
         "LEFT JOIN Job_Card J ON J.DocId = s.Job_DocID " & _
         "LEFT JOIN SP_Stock Stk ON S.DocID = Stk.Invoice_DocId " & _
         "LEFT JOIN Subgroup Sg ON S.Party_Code = Sg.Subcode " & _
         "WHERE Stk.Amount2 <>0 " & Condstr & _
         "GROUP BY S.DocID " & _
         "ORDER BY  Max(S.V_Date),Max(S.V_No ) "



    Set RstRep = New ADODB.Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    
    
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "SpareSaleAccount"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub ProcSparePurchaseAccount()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset




If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " And S.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and S.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
     
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("S.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("S.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(S.DocId,1) in (" & GridString2 & ")"



    If FGrid.TextMatrix(List2, 1) = "Local" Then
        Condstr = Condstr & " And S.L_C = 'L'"
    ElseIf FGrid.TextMatrix(List2, 1) = "Central" Then
        Condstr = Condstr & " And S.L_C = 'C'"
    End If








If FGrid.TextMatrix(List1, 1) = "Purchase" Then
    Condstr = Condstr & " and S.V_Type IN ('SXPIR','SXPIC')"
Else
    Condstr = Condstr & " and S.V_Type IN ('SYPRC','SYPRR')"
End If



mQry = "SELECT  Max(S.V_Date) AS Date, Max(S.V_Type) as SaleType, Max(Sg.Name) as PartyName, S.DocID, Max(S.V_No) AS BillNo, " & _
    "sum(CASE WHEN Stk.TaxPer=14 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_14, " & _
    "sum (CASE WHEN Stk.TaxPer=10 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_10, " & _
    "sum(CASE WHEN Stk.TaxPer=5 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_5, " & _
    "sum(CASE WHEN Stk.TaxPer=0 Then stk.Net_Amt2 ELSE 0 End) AS VAT_ASSESABLE_0, " & _
    "Sum(Stk.Net_Amt2) AS VAT_ASSESABLE, " & _
    "sum (CASE WHEN Stk.TaxPer=14 Then stk.TaxAmt ELSE 0 End) AS VAT_14, " & _
    "sum (CASE WHEN Stk.TaxPer=10 Then stk.TaxAmt ELSE 0 End) AS VAT_10, " & _
    "sum (CASE WHEN Stk.TaxPer=5 Then stk.TaxAmt ELSE 0 End) AS VAT_5, " & _
    "sum (CASE WHEN Stk.TaxPer=0 Then stk.TaxAmt ELSE 0 End) AS VAT_0, " & _
    "Sum(Stk.TaxAmt) AS Tax_Amt, Sum(Stk.Amount2) AS TotalAmount " & _
     "FROM SP_Purch  S " & _
     "LEFT JOIN SP_Stock Stk ON S.DocID = Stk.Invoice_DocId " & _
     "LEFT JOIN Subgroup Sg ON S.Party_Code = Sg.Subcode " & _
     "WHERE Stk.Amount2 <>0 " & Condstr & _
     "AND Stk.Amount2 <>0 " & _
     "GROUP BY S.DocID " & _
     "ORDER BY  Max(S.V_Date),Max(S.V_No ) "




Set RstRep = New ADODB.Recordset
RstRep.CursorLocation = adUseClient
RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    
    
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "SparePurchaseAccount"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub




Private Sub OutPutTaxRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset
'Date1,Date2,List1,List1,List1,List2,List1,List1
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub


    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = "where SP.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " And SP.Total_Amt>0"
     
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " in (" & GridString1 & ")"
    
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(SP.DocId,1) in (" & GridString2 & ")"




If FGrid.TextMatrix(List1, 1) = "Sale" Then
'    mQRY = "Select  SP.DocId,SP.V_No,SPStk.TaxAmt, " & _
'        " " & cIIF("SP.V_type in ('SYSIC','W_SIC')", "SP.Party_Name", "SG.Name") & " as Name,SG.LSTNo, " & _
'        " SP.total_Amt as Net_Amt,SPStk.Net_Amt2 as Amount, SP.V_Date, Stk.TaxPer, Max(SP.L_C) as L_C, " & _
'        " Max(SpStk.Tax_YN) as Tax_YN " & _
'        " From (((SP_Sale as SP Left Join SP_Stock as Stk on SP.DocID=Stk.Invoice_DocID)" & _
'        " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
'        " Left Join Sp_Stock SPStk on SP.Docid=SPStk.Invoice_Docid) " & _
'        " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code "
    mQry = "Select  SP.DocId,SP.V_No,stk.TaxAmt, " & _
        " " & cIIF("SP.V_type in ('SYSIC','W_SIC')", "SP.Party_Name", "SG.Name") & " as Name,SG.LSTNo, " & _
        " SP.total_Amt as Net_Amt,stk.Net_Amt2 as Amount, SP.V_Date, Stk.TaxPer, Max(SP.L_C) as L_C, " & _
        " Max(Stk.Tax_YN) as Tax_YN, Max(D_Amt_TB)+Max(D_Amt_TP) As Discount, Stk.SatPer, Stk.SatAmt " & _
        " From (((SP_Sale as SP Left Join SP_Stock as Stk on SP.DocID=Stk.Invoice_DocID)" & _
        " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
        " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code) "
        
    mQry = mQry & Condstr
    mQry = mQry & " and SP.V_Type IN ('SYSIC','SYSIR','W_SIC','W_SIR') Group by SP.DocId,SP.V_No,stk.TaxAmt,SP.V_type,SP.Party_Name,SG.Name,stk.Net_Amt2,SP.V_Date,Stk.TaxPer,SG.LSTNo,SP.total_Amt, stk.Part_No, stk.Part_SrlNo, Stk.SatPer, Stk.SatAmt "
Else
    mQry = "Select distinct SP.DocId,SP.V_No,TF.Tax_Per as TaxPer,SP.Tax_Amt as TaxAmt," & cIIF("SP.V_type ='SXSRC'", "SP.Party_Name", "SG.Name") & " as Name,SG.LSTNo,SP.total_Amt as Net_Amt,(SP.SprAmt_MRP_TB+SP.OilAmt_MRP_TB+SP.SprAmt_TB+SP.OilAmt_TB)-SP.D_Amt_TB as Amount,SP.V_Date, Stk.SatPer, Stk.SatAmt" & _
        " From ((SP_Sale as SP Left Join SP_Stock as Stk on SP.DocID=Stk.DocID)" & _
        " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
        " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code "
    mQry = mQry & Condstr
    mQry = mQry & " and SP.V_Type IN ('SXSRC','SXSRR')"
End If

Set RstRep = New ADODB.Recordset
RstRep.CursorLocation = adUseClient
RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "OutputTaxReg"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub DailyLubConProc()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, Condstr2 As String, CondStr3 As String
Dim TmpRst As ADODB.Recordset
'Date1,Date2,List1,List1,List1,List2,List1,List1
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    Condstr = "where SP_Stock.V_Date >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "  and SP_Stock.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("SP_Stock.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP_stock.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(SP_Stock.DocId,1) in (" & GridString2 & ")"

    mQry = "Select V_Date,Sum(SP_Stock.Qty_Iss - SP_Stock.Qty_Ret) as QtyIss,Max(SP_Stock.Rate) as Rate,SP_Stock.Lub_Category from SP_Stock  "

        
    mQry = mQry & Condstr
    mQry = mQry & " and SP_Stock.V_Type='W_RG' and SP_Stock.Lub_Category Is Not Null and SP_Stock.Lub_Category<>'' Group by SP_Stock.V_Date,SP_Stock.Lub_Category Order by SP_Stock.V_Date  "


    Set RstRep = New ADODB.Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    RepName = "DailyLubCon"
    
NXT:
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    RepName = "DailyLubCon"
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SaleSummary()
On Error GoTo ELoop
Dim mQry As String, Condstr As String, mQRY1 As String
Dim TmpRst As ADODB.Recordset
Dim TmpRst1 As ADODB.Recordset
Dim TotToday, TotTilDt, TotLstMon As Integer

Dim TotPdTd, TotPdTil, TotCouTd, TotCopTil, TotPdiTd, TotPdiTil, TotAccTd, TotAccTil, TotTd, TotTil, TotLstMonTil As Integer

Dim TotSprToday, TotSprTilDt, TotSprLstMon As Double
Dim TotLbrToday, TotLbrTilDt, TotLbrLstMon As Double

Dim PageWidth As Byte, PageLength As Integer, mHeader As Double
Dim fob As New FileSystemObject
Dim FirstDate$

FirstDate = "01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)

    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " Where SP_Sale.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Sp_Sale.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("SP_sale.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Sp_Sale.DocId,1) in (" & GridString2 & ")"
    
    
    Set RstRep = New ADODB.Recordset
        With RstRep
        
            '*****************Spare Party Fields**********************
            .Fields.Append "CSaleToday", adDouble, 20, adFldIsNullable
            .Fields.Append "CSaleTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "CSaleLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "WorkConToday", adDouble, 20, adFldIsNullable
            .Fields.Append "WorkConTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "WorkConLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "WarrRepToday", adDouble, 20, adFldIsNullable
            .Fields.Append "WarrRepTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "WarrRepLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "AccRepToday", adDouble, 20, adFldIsNullable
            .Fields.Append "AccRepTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "AccRepLstMon", adDouble, 20, adFldIsNullable
            
            '.Fields.Append "DenAndPenToday", adDouble, 20, adFldIsNullable
            '.Fields.Append "DenAndPenTilDt", adDouble, 20, adFldIsNullable
            '.Fields.Append "DenAndPenLstMon", adDouble, 20, adFldIsNullable
            
            .Fields.Append "PDIConPaidToday", adDouble, 20, adFldIsNullable
            .Fields.Append "PDIConPaidTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "PDIConPaidLstMon", adDouble, 20, adFldIsNullable
            
            .Fields.Append "PDIConUWToday", adDouble, 20, adFldIsNullable
            .Fields.Append "PDIConUWTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "PDIConUWLstMon", adDouble, 20, adFldIsNullable
            
            
            '****************Labour & Oil Fields********************
            
            .Fields.Append "PaidLbrToday", adDouble, 20, adFldIsNullable
            .Fields.Append "PaidLbrTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "PaidLbrLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "OutLbrToday", adDouble, 20, adFldIsNullable
            .Fields.Append "OutLbrTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "OutLbrLstMon", adDouble, 20, adFldIsNullable
            
            
            
            .Fields.Append "PaintToday", adDouble, 20, adFldIsNullable
            .Fields.Append "PaintTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "PaintLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "AccidentToday", adDouble, 20, adFldIsNullable
            .Fields.Append "AccidentTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "AccidentLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "WarrToday", adDouble, 20, adFldIsNullable
            .Fields.Append "WarrTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "WarrLstMon", adDouble, 20, adFldIsNullable
            
            .Fields.Append "CouponToday", adDouble, 20, adFldIsNullable
            .Fields.Append "CouponTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "CouponLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "PDIToday", adDouble, 20, adFldIsNullable
            .Fields.Append "PDITilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "PDILstMon", adDouble, 20, adFldIsNullable
            
            
            
            .Fields.Append "OilToday", adDouble, 20, adFldIsNullable
            .Fields.Append "OilTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "OilLstMon", adDouble, 20, adFldIsNullable
            
            
            
            'Vehicle Attended Fields
            
            .Fields.Append "PaidToday", adInteger, 20, adFldIsNullable
            .Fields.Append "PaidTilDt", adInteger, 20, adFldIsNullable
            .Fields.Append "PaidLstMon", adInteger, 20, adFldIsNullable
            
            
            .Fields.Append "CoupToday", adInteger, 20, adFldIsNullable
            .Fields.Append "CoupTilDt", adInteger, 20, adFldIsNullable
            .Fields.Append "CoupLstMon", adInteger, 20, adFldIsNullable
            
            
            .Fields.Append "AccToday", adInteger, 20, adFldIsNullable
            .Fields.Append "AccTilDt", adInteger, 20, adFldIsNullable
            .Fields.Append "AccLstMon", adInteger, 20, adFldIsNullable
            
            
            .Fields.Append "WarrServToday", adInteger, 20, adFldIsNullable
            .Fields.Append "WarrServTilDt", adInteger, 20, adFldIsNullable
            .Fields.Append "WarrServLstMon", adInteger, 20, adFldIsNullable
            
            
            .Fields.Append "PDIServToday", adInteger, 20, adFldIsNullable
            .Fields.Append "PDIServTilDt", adInteger, 20, adFldIsNullable
            .Fields.Append "PDIServLstMon", adInteger, 20, adFldIsNullable
            
            
            
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
    End With
    
    With RstRep
        .AddNew
            '************ Adding Spare Part Amt Details************************
            .Fields("CSaleToday") = GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleToday from Sp_Sale where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1))).Fields(0).Value
            .Fields("CSaleTilDt") = GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1))).Fields(0).Value
            .Fields("CSaleLstMon") = GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1)))).Fields(0).Value
            
            
            .Fields("WarrRepToday") = GCn.Execute("Select Sum(Sp_Stock.Amount) from ((Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "  and Sp_Stock.Purpose='W'").Fields(0).Value
            .Fields("WarrRepTilDt") = GCn.Execute("Select Sum(Sp_Stock.Amount)  from ((Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Sp_Stock.Purpose='W'").Fields(0).Value
            .Fields("WarrRepLstMon") = GCn.Execute("Select Sum(Sp_Stock.Amount) from ((Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Sp_Stock.Purpose='W' ").Fields(0).Value
            
            If UCase(PubSiteName) = "GWALTOLI" Then
                .Fields("AccRepToday") = 0
                .Fields("AccRepTilDt") = 0
                .Fields("AccRepLstMon") = 0
            Else
                .Fields("AccRepToday") = GCn.Execute("Select Sum(SprAmt_TB+SprAmt_TP) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where Sp_Sale.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Type='A'").Fields(0).Value
                .Fields("AccRepTilDt") = GCn.Execute("Select Sum(SprAmt_TB+SprAmt_TP) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where Sp_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='A'").Fields(0).Value
                .Fields("AccRepLstMon") = GCn.Execute("Select Sum(SprAmt_TB+SprAmt_TP) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where Sp_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='A'").Fields(0).Value
            End If
            
            .Fields("PDIConPaidToday") = GCn.Execute("Select Sum(Sp_Stock.Amount) from Sp_Sale Left Join  Sp_stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Sp_Stock.Purpose='P'").Fields(0).Value
            .Fields("PDIConPaidTilDt") = GCn.Execute("Select Sum(Sp_Stock.Amount) from Sp_Sale Left Join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Sp_Stock.Purpose='P'").Fields(0).Value
            .Fields("PDIConPaidLstMon") = GCn.Execute("Select Sum(Sp_Stock.Amount) from Sp_Sale Left Join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Sp_Stock.Purpose='P'").Fields(0).Value
            
            .Fields("PDIConUWToday") = GCn.Execute("Select Sum(Sp_Stock.Amount)  from ((Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='P' and Sp_Stock.Purpose='W' ").Fields(0).Value
            .Fields("PDIConUWTilDt") = GCn.Execute("Select Sum(Sp_Stock.Amount)  from ((Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='P' and Sp_Stock.Purpose='W'").Fields(0).Value
            .Fields("PDIConUWLstMon") = GCn.Execute("Select Sum(Sp_Stock.Amount) from ((Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left join Sp_Stock on Sp_Sale.DocId=Sp_Stock.Invoice_DocId where Sp_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='P' and Sp_Stock.Purpose='W' ").Fields(0).Value
            
            
            If UCase(PubSiteName) = "GWALTOLI" Then
                .Fields("WorkConToday") = GCn.Execute("Select Sum(SprAmt_TB+SprAmt_TP) from Sp_Sale where Sp_Sale.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and V_Type in ('W_SIC','W_SIR')").Fields(0).Value
                .Fields("WorkConTilDt") = GCn.Execute("Select Sum(SprAmt_TB+SprAmt_TP) from Sp_Sale where Sp_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and V_Type in ('W_SIC','W_SIR')").Fields(0).Value
                .Fields("WorkConLstMon") = GCn.Execute("Select Sum(SprAmt_TB+SprAmt_TP) from Sp_Sale where Sp_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and V_Type in ('W_SIC','W_SIR')").Fields(0).Value
                
                '.Fields("WorkConToday") = GCn.Execute("Select Sum(Sp_Stock.Net_Amt) from ((Sp_Stock Left Join Job_Card on Sp_Stock.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Part on SP_Stock.Part_No=Part.Part_No where Sp_Stock.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "  and Sp_Stock.V_type in('W_RG') and SP_Stock.Purpose <> 'F' and Part.Part_Grade <> 'P' and len(SP_Stock.Invoice_Docid) > 0 ").Fields(0).Value
                '.Fields("WorkConTilDt") = GCn.Execute("Select Sum(Sp_Stock.Net_Amt) from ((Sp_Stock Left Join Job_Card on Sp_Stock.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Part on SP_Stock.Part_No=Part.Part_No where Sp_Stock.V_Date Between #01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8) & "#" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Sp_Stock.V_type in('W_RG') and SP_Stock.Purpose <> 'F' and Part.Part_Grade <> 'P' and len(SP_Stock.Invoice_Docid) > 0 ").Fields(0).Value
                '.Fields("WorkConLstMon") = GCn.Execute("Select Sum(Sp_Stock.Net_Amt) from ((Sp_Stock Left Join Job_Card on Sp_Stock.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Part on SP_Stock.Part_No=Part.Part_No where Sp_Stock.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & "  and Sp_Stock.V_type in('W_RG') and SP_Stock.Purpose <> 'F' and Part.Part_Grade <> 'P' and len(SP_Stock.Invoice_Docid) > 0").Fields(0).Value
            Else
                .Fields("WorkConToday") = GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where Sp_Sale.V_Date=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "  and Sp_Sale.V_type in('W_SIC','W_SIR')").Fields(0).Value
                .Fields("WorkConTilDt") = GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where Sp_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value
                .Fields("WorkConLstMon") = GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where Sp_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & "  and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value
            End If
            
            '*****************Adding Labour & Oil Details***********************
            
            If UCase(PubSiteName) = "GWALTOLI" Then
                .Fields("PaidLbrToday") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code  where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Labour.Lab_Group not in ('X','Z')").Fields(0).Value
                .Fields("PaidLbrTilDt") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Labour.Lab_Group not in ('X','Z')").Fields(0).Value
                .Fields("PaidLbrLstMon") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Labour.Lab_Group not in ('X','Z')").Fields(0).Value
            Else
                .Fields("PaidLbrToday") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "").Fields(0).Value
                .Fields("PaidLbrTilDt") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "").Fields(0).Value
                .Fields("PaidLbrLstMon") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & "").Fields(0).Value
            End If
            If UCase(PubSiteName) = "GWALTOLI" Then
                'Used for Denting Labour **************
                .Fields("OutLbrToday") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code  where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Labour.Lab_Group in ('X')").Fields(0).Value
                .Fields("OutLbrTilDt") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Labour.Lab_Group  in ('X')").Fields(0).Value
                .Fields("OutLbrLstMon") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Labour.Lab_Group in ('X')").Fields(0).Value
            Else
                .Fields("OutLbrToday") = GCn.Execute("Select Sum(LabAmt_Out) from Job_Card   where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1))).Fields(0).Value
                .Fields("OutLbrTilDt") = GCn.Execute("Select Sum(LabAmt_Out) from Job_Card   where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1))).Fields(0).Value
                .Fields("OutLbrLstMon") = GCn.Execute("Select Sum(LabAmt_Out) from Job_Card   where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1)))).Fields(0).Value
            End If
            If UCase(PubSiteName) = "GWALTOLI" Then
                .Fields("PaintToday") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code  where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Labour.Lab_Group in ('Z')").Fields(0).Value
                .Fields("PaintTilDt") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Labour.Lab_Group  in ('Z')").Fields(0).Value
                .Fields("PaintLstMon") = GCn.Execute("Select Sum(LabourAmt) from (Job_Lab Left Join  Job_Card on Job_Lab.Job_DocId=Job_Card.DocID) Left Join Labour on Labour.Lab_Code=Job_Lab.Lab_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Labour.Lab_Group in ('Z')").Fields(0).Value
            Else
                .Fields("PaintToday") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId  where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='D' and Job_Lab.Chrg_type='C' ").Fields(0).Value
                .Fields("PaintTilDt") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='D' and Job_Lab.Chrg_type='C' ").Fields(0).Value
                .Fields("PaintLstMon") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='D' and Job_Lab.Chrg_type='C' ").Fields(0).Value
            End If
            If UCase(PubSiteName) = "GWALTOLI" Then
                .Fields("AccidentToday") = 0
                .Fields("AccidentTilDt") = 0
                .Fields("AccidentLstMon") = 0
            Else
                .Fields("AccidentToday") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='A' and Job_Lab.Chrg_type='C'").Fields(0).Value
                .Fields("AccidentTilDt") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & " " & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='A' and Job_Lab.Chrg_type='C' ").Fields(0).Value
                .Fields("AccidentLstMon") = GCn.Execute("Select Sum(LabAmt_TB+LabAmt_TP-Lab_D_Amt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='A' and Job_Lab.Chrg_type='C' ").Fields(0).Value
            End If
            
            .Fields("WarrToday") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and  Job_Lab.Chrg_type='W' ").Fields(0).Value
            .Fields("WarrTilDt") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and  Job_Lab.Chrg_type='W'").Fields(0).Value
            .Fields("WarrLstMon") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and  Job_Lab.Chrg_type='W'").Fields(0).Value
            
            .Fields("CouponToday") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and  Job_Lab.Chrg_type='F' and Job_Lab.Chrg_from='M'").Fields(0).Value
            .Fields("CouponTilDt") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and  Job_Lab.Chrg_type='F' and Job_Lab.Chrg_from='M' ").Fields(0).Value
            .Fields("CouponLstMon") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and  Job_Lab.Chrg_type='F' and Job_Lab.Chrg_from='M' ").Fields(0).Value
            
            .Fields("PDIToday") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and  Job_Lab.Chrg_type='P'").Fields(0).Value
            .Fields("PDITilDt") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and  Job_Lab.Chrg_type='P' ").Fields(0).Value
            .Fields("PDILstMon") = GCn.Execute("Select Sum(Job_Lab.LabourAmt) from (Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join Job_Lab on Job_Card.Docid=Job_Lab.Job_DocId where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and  Job_Lab.Chrg_type='P' ").Fields(0).Value
            
            
            .Fields("OilToday") = GCn.Execute("Select Sum(OilAmt_TB+OilAmt_TP+OilAmt_MRP_TB+OilAmt_MRP_TP) from Sp_Sale where Sp_Sale.V_Date = " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and SP_Sale.V_type in ('W_SIC','W_SIR')").Fields(0).Value
            .Fields("OilTilDt") = GCn.Execute("Select Sum(OilAmt_TB+OilAmt_TP+OilAmt_MRP_TB+OilAmt_MRP_TP) from Sp_Sale where Sp_Sale.V_Date Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and SP_Sale.V_type in ('W_SIC','W_SIR')").Fields(0).Value
            .Fields("OilLstMon") = GCn.Execute("Select Sum(OilAmt_TB+OilAmt_TP+OilAmt_MRP_TB+OilAmt_MRP_TP) from Sp_Sale where Sp_Sale.V_Date Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and SP_Sale.V_type in ('W_SIC','W_SIR')").Fields(0).Value
           
        .Update
    End With
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    PageLength = PubPageLength
    PageWidth = 80
    'mRec = 45
    'Header printing
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
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
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, PRN_TIT("Daily WorkShop Return Report For " & FGrid.TextMatrix(Date1, 1), "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, "                               Todays           Til Date       Last Month Growth"
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, mEmph & "Spare Party Details" & mEmph1
    Print #1, ""
    Print #1, ""
    mHeader = mHeader + 3
    
    TotSprToday = Val(VNull(RstRep!CSaleToday)) + Val(VNull(RstRep!WorkConToday)) + Val(VNull(RstRep!WarrRepToday)) + Val(VNull(RstRep!AccRepToday)) + Val(VNull(RstRep!PDIConPaidToday)) + Val(VNull(RstRep!PDIConUWToday))
    TotSprTilDt = Val(VNull(RstRep!CSaleTilDt)) + Val(VNull(RstRep!WorkConTilDt)) + Val(VNull(RstRep!WarrRepTilDt)) + Val(VNull(RstRep!AccRepTilDt)) + Val(VNull(RstRep!PDIConPaidTilDt)) + Val(VNull(RstRep!PDIConUWTilDt))
    TotSprLstMon = Val(VNull(RstRep!CSaleLstMon)) + Val(VNull(RstRep!WorkConLstMon)) + Val(VNull(RstRep!WarrRepLstMon)) + Val(VNull(RstRep!AccRepLstMon)) + Val(VNull(RstRep!PDIConPaidLstMon)) + Val(VNull(RstRep!PDIConUWLstMon))
    
    
    
    
    Print #1, "1." & "Counter Sale............" & Space(3) & PSTR(Format(RstRep!CSaleToday, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!CSaleTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!CSaleLstMon, "0.00"), 8, , AlignRight)
    Print #1, "2." & "Work Shop Consumption..." & Space(3) & PSTR(Format(RstRep!WorkConToday, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!WorkConTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!WorkConLstMon, "0.00"), 8, , AlignRight)
    Print #1, "3." & "Warranty Replacement...." & Space(3) & PSTR(Format(VNull(RstRep!WarrRepToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!WarrRepTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!WarrRepLstMon, "0.00"), 8, , AlignRight)
    Print #1, "4." & "Accident Repair........." & Space(3) & PSTR(Format(VNull(RstRep!AccRepToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!AccRepTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!AccRepLstMon, "0.00"), 8, , AlignRight)
    Print #1, "5." & "P.D.I.Consumption Paid.." & Space(3) & PSTR(Format(VNull(RstRep!PDIConPaidToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!PDIConPaidTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!PDIConPaidLstMon, "0.00"), 8, , AlignRight)
    Print #1, "6." & "P.D.I.Consumption U/W..." & Space(3) & PSTR(Format(VNull(RstRep!PDIConUWToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!PDIConUWTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!PDIConUWLstMon, "0.00"), 8, , AlignRight)
    
    
    mHeader = mHeader + 5
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, "Total===>>>" & Space(18) & PSTR(Format(TotSprToday, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(TotSprTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(TotSprLstMon, "0.00"), 8, , AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 2
    
    Print #1, mEmph & "Labour & Oil Details" & mEmph1
    mHeader = mHeader + 1
    
    TotLbrToday = Val(VNull(RstRep!PaidLbrToday)) + Val(VNull(RstRep!OutLbrToday)) + Val(VNull(RstRep!PaintToday)) + Val(VNull(RstRep!AccidentToday)) + Val((VNull(RstRep!WarrRepToday) * 10) / 100) + Val(VNull(RstRep!CouponToday)) + Val(VNull(RstRep!PdIToday) + ((VNull(RstRep!OilToday) * 40) / 100))
    TotLbrTilDt = Val(VNull(RstRep!PaidLbrTilDt)) + Val(VNull(RstRep!OutLbrTilDt)) + Val(VNull(RstRep!PaintTilDt)) + Val(VNull(RstRep!AccidentTilDt)) + Val((VNull(RstRep!WarrRepTilDt) * 10) / 100) + Val(VNull(RstRep!CouponTilDt)) + Val(VNull(RstRep!PdItildt) + ((VNull(RstRep!OilTilDt) * 40) / 100))
    TotLbrLstMon = Val(VNull(RstRep!PaidLbrLstMon)) + Val(VNull(RstRep!OutLbrLstMon)) + Val(VNull(RstRep!PaintLstMon)) + Val(VNull(RstRep!AccidentLstMon)) + Val((VNull(RstRep!WarrRepLstMon) * 10) / 100) + Val(VNull(RstRep!CouponLstMon)) + Val(VNull(RstRep!PdILstMon) + ((VNull(RstRep!OilLstMon) * 40) / 100))
    
    Print #1, "1." & "Paid Labour............. " & Space(2) & PSTR(Format(VNull(RstRep!PaidLbrToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!PaidLbrTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!PaidLbrLstMon, "0.00"), 8, , AlignRight)
    If UCase(PubSiteName) = "GWALTOLI" Then
        Print #1, "2." & "Denting Charges......... " & Space(2) & PSTR(Format(VNull(RstRep!OutLbrToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!OutLbrTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!OutLbrLstMon, "0.00"), 8, , AlignRight)
    Else
        Print #1, "2." & "Out Side Labour......... " & Space(2) & PSTR(Format(VNull(RstRep!OutLbrToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!OutLbrTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!OutLbrLstMon, "0.00"), 8, , AlignRight)
    End If
    Print #1, "3." & "Painting Charges........ " & Space(2) & PSTR(Format(VNull(RstRep!PaintToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!PaintTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!PaintLstMon, "0.00"), 8, , AlignRight)
    Print #1, "4." & "Accident Labour......... " & Space(2) & PSTR(Format(VNull(RstRep!AccidentToday), "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!AccidentTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!AccidentLstMon, "0.00"), 8, , AlignRight)
    Print #1, "5." & "Warranty Labour......... " & Space(2) & PSTR(Format((VNull(RstRep!WarrRepToday) * 10) / 100, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format((VNull(RstRep!WarrRepTilDt) * 10) / 100, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format((VNull(RstRep!WarrRepLstMon) * 10) / 100, "0.00"), 8, , AlignRight)
    Print #1, "6." & "Coupon Labour........... " & Space(2) & PSTR(Format(RstRep!CouponToday, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!CouponTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!CouponLstMon, "0.00"), 8, , AlignRight)
    Print #1, "7." & "PDI Labour.............. " & Space(2) & PSTR(Format(RstRep!PdIToday, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!PdItildt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!PdILstMon, "0.00"), 8, , AlignRight)
    Print #1, "8." & "Total Oil Sale.......... " & Space(2) & PSTR(Format(RstRep!OilToday, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(RstRep!OilTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(RstRep!OilLstMon, "0.00"), 8, , AlignRight)
    Print #1, "9." & "Oil Profit 40 % Of Sale. " & Space(2) & PSTR(Format((RstRep!OilToday * 40) / 100, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format((RstRep!OilTilDt * 40) / 100, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format((RstRep!OilLstMon * 40) / 100, "0.00"), 8, , AlignRight)
    
    
    mHeader = mHeader + 8
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    Print #1, "Total===>>>" & Space(18) & PSTR(Format(TotLbrToday, "0.00"), 8, , AlignRight) & Space(11) & PSTR(Format(TotLbrTilDt, "0.00"), 8, , AlignRight) & Space(16) & PSTR(Format(TotLbrLstMon, "0.00"), 8, , AlignRight)
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 2
    
    Print #1, mEmph & "Vehicle Attended Details" & mEmph1
    mHeader = mHeader + 1
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    
    '********Recordset for the veh Attended details********
    
    Set TmpRst = New ADODB.Recordset
    
        With TmpRst
            .Fields.Append "VehGroup", adChar, 20, adFldIsNullable
            .Fields.Append "VehCode", adChar, 20, adFldIsNullable
            
            .Fields.Append "PdToday", adDouble, 20, adFldIsNullable
            .Fields.Append "PdTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "PdLstMon", adDouble, 20, adFldIsNullable
            
            .Fields.Append "CouponToday", adDouble, 20, adFldIsNullable
            .Fields.Append "CouponTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "CouponLstMon", adDouble, 20, adFldIsNullable
            
            .Fields.Append "AccToday", adDouble, 20, adFldIsNullable
            .Fields.Append "AccTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "AccLstMon", adDouble, 20, adFldIsNullable
            
            
            .Fields.Append "WarrToday", adDouble, 20, adFldIsNullable
            .Fields.Append "WarrTilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "WarrLstMon", adDouble, 20, adFldIsNullable
            
            .Fields.Append "PDIToday", adDouble, 20, adFldIsNullable
            .Fields.Append "PDITilDt", adDouble, 20, adFldIsNullable
            .Fields.Append "PDILstMon", adDouble, 20, adFldIsNullable
            
            
            
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
        End With
        
        Set TmpRst1 = GCn.Execute("Select ModelCat_Code,Modelcat_Name from Model_Cat Order by Modelcat_Name")
        
        Do While TmpRst1.EOF = False
                With TmpRst
                    .AddNew
                    !VehGroup = XNull(TmpRst1!ModelCat_NAME)
                    !VehCode = XNull(TmpRst1!ModelCat_Code)
                    .Update
                End With
            TmpRst1.MoveNext
        Loop
        If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
        
        Do While TmpRst.EOF = False
            With TmpRst
                If GCn.Execute("Select Model_Cat.ModelCat_Name from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='C' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !PdToday = GCn.Execute("Select Job_Card.*  from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='C' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !PdToday = 0
                End If
                
                If GCn.Execute("Select Model_Cat.ModelCat_Name from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='C' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !PdTilDt = GCn.Execute("Select Job_Card.CardNo from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='C' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !PdTilDt = 0
                End If
                
                If GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='C' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !PdLstMon = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='C' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !PdLstMon = 0
                End If
                
                If GCn.Execute("Select Model_Cat.ModelCat_Name from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg IN('F','W') and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !CouponToday = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg IN('F','W') and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !CouponToday = 0
                End If
                
                If GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg IN('F','W') and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !CouponTilDt = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg IN('F','W') and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !CouponTilDt = 0
                End If
                
                If GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg IN('F','W') and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !CouponLstMon = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg IN('F','W') and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !CouponLstMon = 0
                End If
                
                
                If GCn.Execute("Select Model_Cat.ModelCat_Name from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='A' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !AccToday = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='A' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !AccToday = 0
                End If
                
                If GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='A' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !AccTilDt = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='A' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !AccTilDt = 0
                End If
                
                If GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='A' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !AccLstMon = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='A' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !AccLstMon = 0
                End If
                
                If GCn.Execute("Select Job_Card.*,Model_Cat.ModelCat_Name from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='P' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !PdIToday = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='P' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !PdIToday = 0
                End If
                
                If GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='P' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !PdItildt = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate("01/" & mID(FGrid.TextMatrix(Date1, 1), 4, 8)) & "" & " and " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Service_type.Serv_Catg='P' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !PdItildt = 0
                End If
                
                If GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='P' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount > 0 Then
                    !PdILstMon = GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate Between " & ConvertDate(DateAdd("m", -1, FirstDate)) & " and " & ConvertDate(DateAdd("m", -1, FGrid.TextMatrix(Date1, 1))) & " and Service_type.Serv_Catg='P' and model_Cat.ModelCat_Code='" & Trim(TmpRst!VehCode) & "'").RecordCount
                Else
                    !PdILstMon = 0
                End If
                
                
                .Update
            End With
            
            TmpRst.MoveNext
        Loop
        
        
    Print #1, "                         -Paid-      Coupon    Accident     -PDI-    Total  Total"
    Print #1, "                        FD    TD     FD   TD   FD    TD    FD  TD   FD   TD L/M TD"
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 3
    
    If TmpRst.RecordCount > 0 Then
        TmpRst.MoveFirst
        Do While TmpRst.EOF = False
            TotToday = Val(VNull(TmpRst!PdToday)) + Val(VNull(TmpRst!CouponToday)) + Val(VNull(TmpRst!AccToday)) + Val(VNull(TmpRst!PdIToday))
            TotTilDt = Val(VNull(TmpRst!PdTilDt)) + Val(VNull(TmpRst!CouponTilDt)) + Val(VNull(TmpRst!AccTilDt)) + Val(VNull(TmpRst!PdItildt))
            TotLstMon = Val(VNull(TmpRst!PdLstMon)) + Val(VNull(TmpRst!CouponLstMon)) + Val(VNull(TmpRst!AccLstMon)) + Val(VNull(TmpRst!PdILstMon))
            
            Print #1, SETW(TmpRst!VehGroup, 20) & Space(3) & PSTR(Format(VNull(TmpRst!PdToday), "0"), 3, , AlignRight) & Space(3) & PSTR(Format(VNull(TmpRst!PdTilDt), "0"), 3, , AlignRight) & Space(4) & PSTR(Format(VNull(TmpRst!CouponToday), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TmpRst!CouponTilDt), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TmpRst!AccToday), "0"), 3, , AlignRight) & Space(3) & PSTR(Format(VNull(TmpRst!AccTilDt), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TmpRst!PdIToday), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TmpRst!PdItildt), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TotToday), "0"), 3, , AlignRight) & PSTR(Format(VNull(TotTilDt), "0"), 5, , AlignRight) & Space(2) & PSTR(Format(VNull(TotLstMon), "0"), 3, , AlignRight)
            
            TotPdTd = TotPdTd + VNull(TmpRst!PdToday)
            TotPdTil = TotPdTil + VNull(TmpRst!PdTilDt)
            
            TotCouTd = TotCouTd + VNull(TmpRst!CouponToday)
            TotCopTil = TotCopTil + VNull(TmpRst!CouponTilDt)
            
            TotAccTd = TotAccTd + VNull(TmpRst!AccToday)
            TotAccTil = TotAccTil + VNull(TmpRst!AccTilDt)
            
            TotPdiTd = TotPdiTd + VNull(TmpRst!PdIToday)
            TotPdiTil = TotPdiTil + VNull(TmpRst!PdItildt)
            
            TotTd = TotTd + Val(VNull(TotToday))
            TotTil = TotTil + Val(VNull(TotTilDt))
                        
            TotLstMonTil = TotLstMonTil + Val(TotLstMon)
            TmpRst.MoveNext
        
        Loop
    End If
    
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 3
    Print #1, SETW("Total===>>>", 20) & Space(3) & PSTR(Format(VNull(TotPdTd), "0"), 3, , AlignRight) & Space(3) & PSTR(Format(VNull(TotPdTil), "0"), 3, , AlignRight) & Space(4) & PSTR(Format(VNull(TotCouTd), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TotCopTil), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TotAccTd), "0"), 3, , AlignRight) & Space(3) & PSTR(Format(VNull(TotAccTil), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TotPdiTd), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TotPdiTil), "0"), 3, , AlignRight) & Space(2) & PSTR(Format(VNull(TotTd), "0"), 3, , AlignRight) & PSTR(Format(VNull(TotTil), "0"), 5, , AlignRight) & Space(2) & PSTR(Format(VNull(TotLstMonTil), "0"), 3, , AlignRight)
    'Print #1,        3) & Format(SETN(, 3), "0") & Space(3) & Format(SETN(VNull(TotPdTil), 3), "0") & Space(4) & Format(SETN(VNull(TotCouTd), 3), "0") & Space(2) & Format(SETN(VNull(TotCopTil), 3), "0") & Space(2) & Format(SETN(VNull(TotAccTd), 3), "0") & Space(3) & Format(SETN(VNull(TotAccTil), 3), "0") & Space(2) & Format(SETN(VNull(TotPdiTd), 3), "0") & Space(2) & Format(SETN(VNull(TotPdiTil), 3), "0") & Space(2) & Format(SETN(VNull(TotTd), 3), "0") & Format(SETN(VNull(TotTil), 3), "0") & Space(2) & Format(SETN(Val(TotLstMonTil), 3), "0")
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 2
    
    
    
    
    'Print #1, mEject
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

    'If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    'RepTitle = UCase(Me.CAPTION)
    RepPrint = False
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub OutPayRepProc()

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
    
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(SP_Sale.site_code,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(sp_sale.site_code,1) ='" & PubSiteCode & "' "
    End If

    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and SP_Sale.Rep_Code  in   (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and  left(SP_Sale.DocId,1) in (" & GridString3 & ")"
    
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "SalesMan", adChar, 50, adFldIsNullable
        .Fields.Append "Party", adChar, 50, adFldIsNullable
        .Fields.Append "BillNo", adDouble, 8, adFldIsNullable
        .Fields.Append "BillDate", adDate, 20, adFldIsNullable
        .Fields.Append "BillAmt", adDouble, 20, adFldIsNullable
        .Fields.Append "OstdAmt", adDouble, 20, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
    Set TmpRst = GCn.Execute("Select * from SP_Sale where V_Date<= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and V_type in ('SYSIR') and total_AMT > 0")
    
    If TmpRst.RecordCount > 0 Then
        Do While TmpRst.EOF = False
            Set RstCrAmt = G_FaCn.Execute("Select Sum(AmtDr-AmtCr) as TotDr from Ledger where SubCode='" & TmpRst!Party_code & "' and Ledger.V_Date<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "")
                If RstCrAmt.RecordCount > 0 Then
                    If VNull(RstCrAmt!TotDr) > 0 Then
                        With RstRep
                             .AddNew
                                 If TmpRst!REP_CODE = "" Then
                                    .Fields("SalesMan") = "Direct"
                                 Else
                                    .Fields("SalesMan") = GCn.Execute("Select Emp_Name from Emp_Mast where Emp_Code='" & TmpRst!REP_CODE & "'").Fields(0).Value
                                 End If
                                 .Fields("Party") = G_FaCn.Execute("Select Name from SubGroup where SubCode='" & TmpRst!Party_code & "'").Fields(0).Value
                                 .Fields("BillNo") = VNull(TmpRst!V_NO)
                                 .Fields("BillDate") = XNull(TmpRst!V_DATE)
                                 .Fields("BillAmt") = VNull(TmpRst!Total_Amt)
                                 .Fields("OstdAmt") = RstCrAmt!TotDr
                            .Update
                        End With
                    End If
                End If
            TmpRst.MoveNext
        Loop
    End If
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "OstdPayRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub CashBankBookProc()
On Error GoTo ELoop
Dim RST1 As ADODB.Recordset, RstTmp As ADODB.Recordset, mGROUP_rs As ADODB.Recordset, SUBGROUP_rs As ADODB.Recordset, TmpGrs As ADODB.Recordset, TmpGrs1 As ADODB.Recordset
Dim DrAc As String, CrAc As String, oBAL As Double, mAcCode As String, mDocNo As String, mDocNo1 As String, mAcCode1 As String
Dim mNARR1 As String, mNARR2 As String, TmpDate As Date, mDate1 As Date, mDate2 As Date, TinTin As Integer
Dim mFLAG1 As Boolean, mFLAG2 As Boolean, mFLAG11 As Boolean, mFLAG22 As Boolean, mFLAG111 As Boolean, mFLAG222 As Boolean, mSiteCode As String
Dim TXTS_DATE As Date
Dim TXTE_DATE As Date
Dim PubDatamanFa As New DMFa.ClsFa
Dim I As Integer, mBType As String
Dim IntDays As Integer

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

Set RstTmp = New ADODB.Recordset
Set RstTmp = PubDatamanFa.FaCASHTMP1(RstTmp)

TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
'If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
mSiteCode = ""
'TXTACC_CODE = Trim(FGrid.TextMatrix(List3, 1))
mAcCode = "11000001" 'Trim(FGrid.TextMatrix(List3, 2))
DrAc = ""
CrAc = ""
oBAL = 0
    
Set RST1 = G_FaCn.Execute("SELECT GROUPNATURE FROM PARTY_LIST WHERE SUBCODE='" & mAcCode & "'")

Set RstTmp = New ADODB.Recordset
With RstTmp
    .Fields.Append "V_DATE", adDate, , adFldIsNullable
    .Fields.Append "V_No", adDouble, 8, adFldIsNullable
    .Fields.Append "V_Type", adVarChar, 5, adFldIsNullable
    .Fields.Append "V_SNO", adDouble, 8, adFldIsNullable
    .Fields.Append "V_ADD", adVarChar, 80, adFldIsNullable
    .Fields.Append "CR", adDouble, 16, adFldIsNullable
    .Fields.Append "ADJAMT", adDouble, 16, adFldIsNullable
    .Fields.Append "SUBCODE", adVarChar, 8, adFldIsNullable
    .Fields.Append "NAME", adVarChar, 600, adFldIsNullable
    .Fields.Append "ADJQTY", adDouble, 16, adFldIsNullable
    .Fields.Append "VType", adVarChar, 5, adFldIsNullable
    .Fields.Append "VNo", adDouble, 8, adFldIsNullable
    .Fields.Append "VADD", adVarChar, 80, adFldIsNullable
    .Fields.Append "VSNO", adDouble, 8, adFldIsNullable
    .Fields.Append "VAL", adDouble, 16, adFldIsNullable
    .Fields.Append "NAME1", adVarChar, 600, adFldIsNullable
    .Fields.Append "ADDRESS1", adVarChar, 80, adFldIsNullable
    .Fields.Append "DOCNO", adVarChar, 25, adFldIsNullable
    .Fields.Append "DOCNO1", adVarChar, 25, adFldIsNullable
    
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With
If RST1.RecordCount <= 0 Then Exit Sub
For I = 1 To 2
    If I = 1 Then mAcCode1 = "11000001"
    If I = 2 Then mAcCode1 = "11200338"
    If RST1!GroupNature = "A" Or RST1!GroupNature = "L" Then
        If PubBackEnd = "S" Then
            oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
        ElseIf PubBackEnd = "A" Then
            oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "- " & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
        End If
    Else
        If PubBackEnd = "S" Then
            oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
        ElseIf PubBackEnd = "A" Then
            oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "-" & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
        End If
    End If
    If oBAL <> 0 Then
        If oBAL < 0 Then
            If I = 1 Then
                CashOpening = Abs(oBAL)
            ElseIf I = 2 Then
                BankOpening = Abs(oBAL)
            End If
            
        Else
            If I = 1 Then
                CashOpening = Abs(oBAL)
            ElseIf I = 2 Then
                BankOpening = Abs(oBAL)
            End If
        End If
    End If
Next

If PubBackEnd = "S" Then
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,subgroup.SUBCODE,CONVERT(VARCHAR,VIEWLEDGER.CHQ_DATE,103)AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND CREDIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
    Set mGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,subgroup.SUBCODE,CONVERT(VARCHAR,VIEWLEDGER.CHQ_DATE,103)AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND DEBIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
ElseIf PubBackEnd = "A" Then
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,subgroup.SUBCODE,FORMAT(VIEWLEDGER.CHQ_DATE,'DD/MM/YY') AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND CREDIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
    Set mGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,subgroup.SUBCODE,FORMAT(VIEWLEDGER.CHQ_DATE,'DD/MM/YY') AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND DEBIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
End If

If Not (mGROUP_rs.EOF) Then mDate2 = mGROUP_rs!V_DATE
If Not (SUBGROUP_rs.EOF) Then mDate1 = SUBGROUP_rs!V_DATE
mFLAG1 = False
mFLAG2 = False
mFLAG11 = False
mFLAG22 = False
mFLAG111 = False
mFLAG222 = False
mNARR1 = ""
mNARR2 = ""

mDocNo1 = ""
TmpDate = TXTS_DATE
Do Until mGROUP_rs.EOF And SUBGROUP_rs.EOF
    mDocNo = ""
    If mDate1 = TmpDate Or mDate2 = TmpDate Then
        RstTmp.AddNew
        If mDate1 = TmpDate Then
            RstTmp!V_Type = SUBGROUP_rs!V_Type
            RstTmp!V_DATE = mDate1
            RstTmp!V_NO = SUBGROUP_rs!V_NO
            RstTmp!V_SNo = SUBGROUP_rs!V_SNo
            If PubFaSiteType = 1 Then
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(SUBGROUP_rs!V_Type), 1) + Trim(mID(Trim(SUBGROUP_rs!V_Type), 3, 3))
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(SUBGROUP_rs!V_NO))
            Else
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(SUBGROUP_rs!V_Type)
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(SUBGROUP_rs!V_NO))
            End If
            Select Case XNull(SUBGROUP_rs!V_Type)
                Case "SYSIC", "W_SIC", "SYPRC"
                    mDocNo = "C " & mDocNo
                Case Else
                    If XNull(SUBGROUP_rs!Party1) = "11000001" Then
                        mDocNo = "C " & mDocNo
                    ElseIf XNull(SUBGROUP_rs!Party) = "11200338" Then
                        mDocNo = "B " & mDocNo
                    Else
                        mDocNo = "N " & mDocNo
                    End If
            End Select
            
            RstTmp!DOCNO = mDocNo
            If mFLAG1 = False Or mFLAG11 = False Or mFLAG111 = False Then
                mFLAG1 = True
                mNARR1 = ""
                If FaXNull(Trim(SUBGROUP_rs!Chq_No)) <> "" Then mNARR1 = mNARR1 + "Ch.No:" + Trim(FaXNull(SUBGROUP_rs!Chq_No)) + " Ch.Dt: " + CStr(FaXNull(SUBGROUP_rs!ChqDate))
                mNARR1 = mNARR1 + Trim(FaXNull(SUBGROUP_rs!mNarr)) + Trim(FaXNull(SUBGROUP_rs!Narr))
                If Len(FaXNull(SUBGROUP_rs!Name)) <> 0 Then
                    RstTmp!Name = FaXNull(SUBGROUP_rs!Name)
                    RstTmp!cr = Format(SUBGROUP_rs!Credit, "0.00")
'                    If mAcCode = SUBGROUP_rs!Party Then oBAL = oBAL - Format(SUBGROUP_rs!CREDIT, "0.00")
                    mFLAG111 = True
                    mFLAG11 = True
                Else
                    If mFLAG11 = False And Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then
                        If mFLAG111 = False Then
                            RstTmp!cr = Format(SUBGROUP_rs!Credit, "0.00")
'                            If mAcCode = SUBGROUP_rs!Party Then oBAL = oBAL - Format(SUBGROUP_rs!CREDIT, "0.00")
                            RstTmp!Name = "As Per Detail"
                            mFLAG111 = True
                            If Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then Set TmpGrs = G_FaCn.Execute("SELECT subgroup.NAME,VIEWLEDGER.DEBIT,VIEWLEDGER.Credit FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE VIEWLEDGER.DOCID='" & SUBGROUP_rs!DocID & "' AND V_SNO<>" & SUBGROUP_rs!V_SNo)
                        Else
                            If Not TmpGrs.EOF Then
                                If TmpGrs!Debit > 0 Then
                                    RstTmp!Name = Space(2) + FaSetW(TmpGrs!Name, 22) + " " + FaSetN(FaSNull(TmpGrs!Debit), 12) + " Dr"
                                Else
                                    RstTmp!Name = Space(2) + FaSetW(TmpGrs!Name, 22) + " " + FaSetN(FaSNull(TmpGrs!Credit), 12) + " Cr"
                                End If
                                TmpGrs.MoveNext
                            End If
                            If TmpGrs.EOF = True Then mFLAG11 = True
                        End If
                    Else
                        RstTmp!Name = mNARR1 'Space(2) + Trim(mID(mNARR1, 1, 36))
                        RstTmp!cr = Format(SUBGROUP_rs!Credit, "0.00")
'                        If mAcCode = SUBGROUP_rs!Party Then oBAL = oBAL - Format(SUBGROUP_rs!CREDIT, "0.00")
                        'mNARR1 = Trim(mID(mNARR1, 37, 510))
                        mFLAG111 = True
                        mFLAG11 = True
                    End If
                End If
                If Len(mNARR1) <= 0 And mFLAG11 = True And mFLAG111 = True Then
                    mFLAG1 = False
                    SUBGROUP_rs.MoveNext
                    If Not SUBGROUP_rs.EOF Then
                        mDate1 = SUBGROUP_rs!V_DATE
                    Else
                        mDate1 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            Else
                mNARR1 = Trim(mNARR1)
                RstTmp!Name = mNARR1 'Space(2) + Trim(mID(mNARR1, 1, 36))
                'mNARR1 = Trim(mID(mNARR1, 37, 510))
                'If Len(mNARR1) <= 0 Then
                    mFLAG1 = False
                    mFLAG11 = False
                    mFLAG111 = False
                    SUBGROUP_rs.MoveNext
                    If Not SUBGROUP_rs.EOF Then
                        mDate1 = SUBGROUP_rs!V_DATE
                    Else
                        mDate1 = DateAdd("D", 1, TXTE_DATE)
                    End If
                'End If
            End If
        End If
        mDocNo1 = ""
        If mDate2 = TmpDate Then
            RstTmp!VType = mGROUP_rs!V_Type
            RstTmp!VNo = mGROUP_rs!V_NO
            RstTmp!V_DATE = mDate2
            RstTmp!VSNo = mGROUP_rs!V_SNo
            If PubFaSiteType = 1 Then
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + left(Trim(mGROUP_rs!V_Type), 1) + Trim(mID(Trim(mGROUP_rs!V_Type), 3, 3))
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(STR(mGROUP_rs!V_NO))
            Else
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(mGROUP_rs!V_Type)
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(STR(mGROUP_rs!V_NO))
            End If
            Select Case XNull(mGROUP_rs!V_Type)
                Case "SXPIC", "SXSRC"
                    mDocNo1 = "C " & mDocNo1
                Case Else
                    If XNull(mGROUP_rs!Party1) = "11000001" Then
                        mDocNo1 = "C " & mDocNo1
                    ElseIf XNull(mGROUP_rs!Party1) = "11200338" Then
                        mDocNo1 = "B " & mDocNo1
                    Else
                        mDocNo1 = "N " & mDocNo1
                    End If
            End Select
            RstTmp!DocNo1 = mDocNo1
            If mFLAG2 = False Or mFLAG22 = False Or mFLAG222 = False Then
                mFLAG2 = True
                mNARR2 = ""
                If FaXNull(Trim(mGROUP_rs!Chq_No)) <> "" Then mNARR2 = mNARR2 + "Ch.No:" + Trim(FaXNull(mGROUP_rs!Chq_No)) + " Ch.Dt: " + CStr(FaXNull(mGROUP_rs!ChqDate))
                mNARR2 = mNARR2 + Trim(FaXNull(mGROUP_rs!mNarr)) + Trim(FaXNull(mGROUP_rs!Narr))
                If Len(FaXNull(mGROUP_rs!Name)) <> 0 Then
                    RstTmp!Name1 = mGROUP_rs!Name
                    RstTmp!ADJAMT = Format(mGROUP_rs!Debit, "0.00")
                    mFLAG222 = True
                    mFLAG22 = True
                Else
                    If mFLAG22 = False And Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then
                        If mFLAG222 = False Then
                            RstTmp!ADJAMT = Format(mGROUP_rs!Debit, "0.00")
                            RstTmp!Name1 = "As Per Detail"
                            mFLAG222 = True
                            If Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then Set TmpGrs1 = G_FaCn.Execute("SELECT subgroup.NAME,VIEWLEDGER.DEBIT,VIEWLEDGER.Credit FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE VIEWLEDGER.DOCID='" & mGROUP_rs!DocID & "' AND V_SNO<>" & mGROUP_rs!V_SNo)
                        Else
                            If Not TmpGrs1.EOF Then
                                If TmpGrs1!Debit > 0 Then
                                    RstTmp!Name1 = Space(2) + FaSetW(TmpGrs1!Name, 22) + " " + FaSetN(FaSNull(TmpGrs1!Debit), 12) + " Dr"
                                Else
                                    RstTmp!Name1 = Space(2) + FaSetW(TmpGrs1!Name, 22) + " " + FaSetN(FaSNull(TmpGrs1!Credit), 12) + " Cr"
                                End If
                                TmpGrs1.MoveNext
                            End If
                            If TmpGrs1.EOF = True Then mFLAG22 = True
                        End If
                    Else
                        RstTmp!Name1 = mNARR2 'Space(2) + Trim(mID(mNARR2, 1, 36))
                        'mNARR2 = Trim(mID(mNARR2, 37, 510))
                        RstTmp!ADJAMT = Format(mGROUP_rs!Debit, "0.00")
'                        If mAcCode = mGROUP_rs!Party Then oBAL = oBAL + Format(mGROUP_rs!DEBIT, "0.00")
                        mFLAG222 = True
                        mFLAG22 = True
                    End If
                End If
                If Len(mNARR2) <= 0 And mFLAG22 = True And mFLAG222 = True Then
                    mFLAG2 = False
                    mGROUP_rs.MoveNext
                    If Not mGROUP_rs.EOF Then
                        mDate2 = mGROUP_rs!V_DATE
                    Else
                        mDate2 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            Else
                mNARR2 = Trim(mNARR2)
                RstTmp!Name1 = mNARR2 'Space(2) + Trim(mID(mNARR2, 1, 36))
                'mNARR2 = Trim(mID(mNARR2, 37, 510))
                'If Len(mNARR2) <= 0 Then
                    mFLAG2 = False
                    mFLAG22 = False
                    mFLAG222 = False
                    mGROUP_rs.MoveNext
                    If Not mGROUP_rs.EOF Then
                        mDate2 = mGROUP_rs!V_DATE
                    Else
                        mDate2 = DateAdd("D", 1, TXTE_DATE)
                    End If
                'End If
            End If
        End If
        RstTmp.Update
    Else
    For I = 1 To 2
        If I = 1 Then mAcCode1 = "11000001"
        If I = 2 Then mAcCode1 = "11200338"
        If RST1!GroupNature = "A" Or RST1!GroupNature = "L" Then
            If PubBackEnd = "S" Then
                oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            ElseIf PubBackEnd = "A" Then
                oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "-" & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            End If
        Else
            If PubBackEnd = "S" Then
                oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            ElseIf PubBackEnd = "A" Then
                oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "- " & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            End If
        End If
        If oBAL <> 0 Then
            If oBAL < 0 Then
                If I = 1 Then
                    CashClosing = Abs(oBAL)
                ElseIf I = 2 Then
                    BankClosing = Abs(oBAL)
                End If
                
            Else
                If I = 1 Then
                    CashClosing = Abs(oBAL)
                ElseIf I = 2 Then
                    BankClosing = Abs(oBAL)
                End If
            End If
        End If
    Next
        If mDate1 <= mDate2 Then
            If mDate1 = CDate("12:00:00 AM") Then
                TmpDate = mDate2
            Else
                TmpDate = mDate1
            End If
        Else
            If mDate2 = CDate("12:00:00 AM") Then
                TmpDate = mDate1
            Else
                TmpDate = mDate2
            End If
        End If
        
        If oBAL <> 0 Then
            RstTmp.AddNew
            RstTmp!V_DATE = TmpDate
            If oBAL < 0 Then
                RstTmp!Name = "OPENING BALANCE"
                RstTmp!cr = Abs(oBAL)
            Else
                RstTmp!Name1 = "OPENING BALANCE"
                RstTmp!ADJAMT = Abs(oBAL)
            End If
            RstTmp.Update
        End If
    End If
Loop
 For I = 1 To 2
    If I = 1 Then mAcCode1 = "11000001"
    If I = 2 Then mAcCode1 = "11200338"
    If RST1!GroupNature = "A" Or RST1!GroupNature = "L" Then
        If PubBackEnd = "S" Then
            oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
        ElseIf PubBackEnd = "A" Then
            oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "-" & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
        End If
    Else
        If PubBackEnd = "S" Then
            oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
        ElseIf PubBackEnd = "A" Then
            oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "-" & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
        End If
    End If
    If oBAL <> 0 Then
        If oBAL < 0 Then
            If I = 1 Then
                CashClosing = Abs(oBAL)
            ElseIf I = 2 Then
                BankClosing = Abs(oBAL)
            End If
            
        Else
            If I = 1 Then
                CashClosing = Abs(oBAL)
            ElseIf I = 2 Then
                BankClosing = Abs(oBAL)
            End If
        End If
    End If
Next
' Interest Calculation
TotalInterest = 0
IntDays = DateDiff("d", CDate(TXTS_DATE), CDate(TXTE_DATE))
IntDays = IntDays + 1
For I = 1 To IntDays
    mAcCode1 = "11200338"
    If RST1!GroupNature = "A" Or RST1!GroupNature = "L" Then
        If PubBackEnd = "S" Then
            oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<=" & FaConvertDate(DateAdd("d", I, CDate(TXTS_DATE))) & " ").Fields(0)
        ElseIf PubBackEnd = "A" Then
            oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "-" & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND V_DATE<=" & FaConvertDate(DateAdd("d", I, CDate(TXTS_DATE))) & " " & mSiteCode & "").Fields(0)
        End If
    Else
        If PubBackEnd = "S" Then
            oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND  V_DATE<=" & FaConvertDate(DateAdd("d", I, CDate(TXTS_DATE))) & " " & mSiteCode & "").Fields(0)
        ElseIf PubBackEnd = "A" Then
            oBAL = G_FaCn.Execute("SELECT " & vIsNull("SUM(CREDIT)", "0") & "-" & vIsNull("SUM(DEBIT)", "0") & " FROM VIEWLEDGER WHERE PARTY='" & mAcCode1 & "' AND v_DATE<=" & FaConvertDate(DateAdd("d", I, CDate(TXTS_DATE))) & " " & mSiteCode & "").Fields(0)
        End If
    End If
    TotalInterest = TotalInterest + (((oBAL * 10.5) / 100) / 365)
Next
RepName = "CashBankBook"
If RstTmp.RecordCount > 0 Then RstTmp.MoveFirst
Set RstRep = RstTmp.Clone
If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
EXIT_SUB:
    Set RST1 = Nothing
    Set mGROUP_rs = Nothing
    Set SUBGROUP_rs = Nothing
    Set TmpGrs = Nothing
    Set TmpGrs1 = Nothing
    Set RstTmp = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub

