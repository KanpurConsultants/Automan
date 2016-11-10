VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form RepSprMIS 
   BackColor       =   &H00C8E8DA&
   Caption         =   "Spare MIS Reports"
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
   Begin VB.CheckBox ChkOpStkOnly 
      BackColor       =   &H00C8E8DA&
      Caption         =   "Only Opening Stock"
      Height          =   240
      Left            =   2940
      TabIndex        =   19
      Top             =   195
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C8E8DA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5850
      Left            =   495
      TabIndex        =   6
      Top             =   480
      Width           =   4680
      Begin VB.CommandButton BTNPRINT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&SpeedPrint"
         DownPicture     =   "RepSprMIS.frx":0000
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
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print Report"
         Top             =   5055
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Frame FrmList 
         BorderStyle     =   0  'None
         Height          =   1830
         Left            =   4140
         TabIndex        =   11
         Top             =   5325
         Visible         =   0   'False
         Width           =   2520
         Begin MSComctlLib.ListView ListView 
            Height          =   1830
            Left            =   135
            TabIndex        =   12
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
         TabIndex        =   10
         Top             =   5055
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton BTNPRINT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Print"
         DownPicture     =   "RepSprMIS.frx":3132
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
         Index           =   0
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print Report"
         Top             =   5055
         Width           =   1290
      End
      Begin VB.CommandButton BTNEXIT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E&xit"
         DownPicture     =   "RepSprMIS.frx":6264
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit Form"
         Top             =   5055
         Width           =   1290
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
         Height          =   4755
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   8387
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16512
         Rows            =   10
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
         FocusRect       =   0
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
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   4890
         Left            =   30
         Top             =   15
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   585
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   4980
         Width           =   4515
      End
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
      Left            =   5880
      TabIndex        =   5
      Top             =   345
      Visible         =   0   'False
      Width           =   915
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
      Left            =   5895
      TabIndex        =   4
      Top             =   1740
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
      Left            =   5895
      TabIndex        =   3
      Top             =   3240
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
      Index           =   4
      Left            =   5910
      TabIndex        =   2
      Top             =   4665
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   2160
      TabIndex        =   0
      Top             =   6540
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DGHelp1 
      Height          =   2745
      Left            =   -1530
      Negotiate       =   -1  'True
      TabIndex        =   1
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
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   330
      TabIndex        =   13
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
      Left            =   5880
      TabIndex        =   14
      Top             =   240
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
      Height          =   1410
      Index           =   2
      Left            =   5880
      TabIndex        =   15
      Top             =   1695
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
      Height          =   1410
      Index           =   4
      Left            =   5880
      TabIndex        =   16
      Top             =   4635
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
      Height          =   1410
      Index           =   3
      Left            =   5880
      TabIndex        =   17
      Top             =   3165
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
End
Attribute VB_Name = "RepSprMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CellBackColLeave$ = &HFFFFFF
Private Const CellBackColEnter$ = &HFFFFC0
Private Const CellBackColLeave1$ = &HEDF7FE
Private Const CellBackColEnter1$ = &HC0E0FF
Dim TRec1Qty As Single, TRec2Qty As Single
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim RsGrid3 As ADODB.Recordset
Dim RsGrid4 As ADODB.Recordset

Dim FormulaStr1$, FormulaStr2$, FormulaStr3$, FormulaStr4$

Dim RsDataGrid1 As ADODB.Recordset
Dim RepTitle$, RepName$
Dim RepPrint As Boolean
Dim RstRep As ADODB.Recordset           '' For Report SQL
Dim RstRep1 As ADODB.Recordset          '' For SubReport
Dim SubRep1 As Boolean                  ''
Private Const GridRowHeight As Integer = 270
Private Const SprABCRep As Byte = 1
''VJ
Private Const SprFSNRep As Byte = 2
Private Const SprXYZRep As Byte = 3
Private Const SprStkLedVal As Byte = 4      'Spr Stock Ledger FIFO
Private Const SprStkLedValSum As Byte = 5   'Spr Stock Valuation FIFO
Private Const SprPartProfit As Byte = 6
Private Const SprSaleInventory As Byte = 7
Private Const SprProjection As Byte = 8
Private Const DeleteLog As Byte = 9

Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const List1 As Byte = 2
Private Const List2 As Byte = 3
Private Const List3 As Byte = 4
Private Const List4 As Byte = 5

Private Const Cat1 As Byte = 6
Private Const Cat2 As Byte = 7
'' Vishal   23-10-02
Private Const Cat3 As Byte = 8
Private Const Cat4 As Byte = 9

Private Const G1Top As Integer = 240
Private Const G2Top As Integer = 1695
Private Const G3Top As Integer = 3165
Private Const G4Top As Integer = 4635

Public GRepFormName$

Dim mLastRow As Integer
Dim mFirstRow As Integer
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

''Following Variables are for Stock Valuation Purpose
Dim TRec1 As ADODB.Recordset, TRec2 As ADODB.Recordset, Temp06 As ADODB.Recordset
Dim RstPart As ADODB.Recordset
Dim RstStock As ADODB.Recordset, RstStock2 As ADODB.Recordset, RstStock3 As ADODB.Recordset
Dim mOP_TB_QTY As Double, mOP_TP_QTY As Double, mOP_TB_VAL As Double, mOP_TP_VAL As Double
Dim mTrf As Boolean, mRate As Double, mPART_ADD As Boolean, TQty As Double
Dim mname$, mInv_No$, mInv_Date$, mNarr$
Dim mRec_TB_Qty As Double, mRec_TB_Val As Double, mRec_TP_Qty As Double, mRec_TP_Val As Double
Dim mIss_TB_Qty As Double, mIss_TB_Val As Double, mIss_TP_Qty As Double, mIss_TP_Val As Double
Dim xMOP_TBQty As Double, xMOP_TBVal As Double, xMOP_TPQty As Double, xMOP_TPVal As Double

Dim tempRst As ADODB.Recordset, TempRst1 As ADODB.Recordset
Dim mVRate As Double, mOPVRate As Double, mDisPer As Double, TempVal As Double, I As Integer
Dim SpeedPrnStkRep As Boolean

Private Sub BTNPRINT_Click(Index As Integer)
On Error GoTo ERRORHANDLER
SubRep1 = False
Select Case GRepFormName
    Case SprABCRep
        SprABCAnalysis
    ''VJ
    Case SprFSNRep
        SprFSNAnalysis
    Case SprXYZRep
        SprXYZAnalysis
    Case SprStkLedVal
        SprStkLedValCalc
    Case SprStkLedValSum
        SprStkLedValSumCalc1
        
        ''Code For Open Report With Old Code
'        If Index = 1 Then SpeedPrnStkRep = True Else SpeedPrnStkRep = False
'        If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
'            SprStkLedValSumCalc1UJWAL
'        Else
'            If FGrid.TextMatrix(Cat4, 1) = "Separate Division" Then
'                SprStkLedValSumCalc1
'            Else
'                SprStkLedValSumCalc
'            End If
'        End If
    Case SprPartProfit
        'If UCase(left(PubComp_Name, 3)) = "JMK" Then
            SprPartProfitCalcJMK
        'Else
        '    SprPartProfitCalc
        'End If
    Case SprSaleInventory
        SprSaleInventoryCalc
    Case SprProjection
        SprProjectionCalc
    Case DeleteLog
        DeleteLogProc
End Select
If RepPrint = False Then Exit Sub
If SpeedPrnStkRep = True Then Exit Sub
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
If MDIForm1.Picture1.Visible = True Then MDIForm1.Picture1.Visible = False
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub Form_Activate()
    If GRepFormName = SprStkLedValSum Then
        BTNPRINT(1).Visible = True
    End If
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
FGrid.CellBackColor = CellBackColLeave
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Select Case FGrid.Row
    Case Cat1, Cat2, Cat3
        TxtGrid(0).MaxLength = 10
    Case List1
        Select Case GRepFormName
            Case SprStkLedVal, SprStkLedValSum, SprPartProfit, SprFSNRep, SprXYZRep, SprProjection
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case DeleteLog
                ListArray = Array("Edited", "Deleted", "All")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
        End Select
    Case List2
        Select Case GRepFormName
            Case SprProjection
                ListArray = Array("All", "Short", "Excess")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            Case SprStkLedVal
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprStkLedValSum
                ListArray = Array("Yes", "No")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case SprFSNRep
                ListArray = Array("All", "Fast", "Slow", "Dead")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
            Case SprXYZRep
                ListArray = Array("All", "X Cat", "Y Cat", "Z Cat")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
            Case SprPartProfit
                ListArray = Array("Summary", "DateWise")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            
        End Select
    Case List3
        ListArray = Array("Summary", "Detail")
        Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
    Case List4
        ListArray = Array("Yes", "No")
        Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
    Case Cat4
        ListArray = Array("Separate Division", "Merge Division")
        Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
End Select
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    TxtGrid(0).Visible = False
    Grid_Hide
    FGrid.SetFocus
    Exit Sub
End If
Select Case FGrid.Row
    Case List1, List2, List3, List4, Cat4
        ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
    Case Date1, Date2, Cat1, Cat2, Cat3
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case FGrid.Row
    Case Cat1, Cat2, Cat3
        NumPress TxtGrid(Index), KeyAscii, 7, 2
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case FGrid.Row
        Case List1, List2, List3, List4, Cat4
            If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
            ListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
End Select
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Select Case FGrid.Row
        Case Cat1, Cat2, Cat3
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
        Case List1, List2, List3, List4, Cat4
           If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            If (GRepFormName = SprStkLedVal Or GRepFormName = SprStkLedValSum Or GRepFormName = SprPartProfit Or GRepFormName = SprXYZRep Or GRepFormName = SprProjection) And FGrid.Row = List1 Then     '' Marked Part Only
                If UCase(TxtGrid(0).TEXT) = UCase("Yes") Then
                    Check1(3).Enabled = False
                Else
                    Check1(3).Enabled = True
                End If
            End If
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
End Select
End Sub
Private Function TxtGridLeave() As Boolean
Select Case FGrid.Row
        Case Cat1, Cat2, Cat3
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
        Case List1, List2, List3, List4, Cat4
            If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            If (GRepFormName = SprStkLedVal Or GRepFormName = SprStkLedValSum Or GRepFormName = SprPartProfit Or GRepFormName = SprXYZRep) And FGrid.Row = List1 Then  '' Marked Part Only
                If UCase(TxtGrid(0).TEXT) = UCase("Yes") Then
                    Check1(3).Enabled = False
                Else
                    Check1(3).Enabled = True
                End If
            End If
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
End Select
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function
Private Sub FGrid_DblClick()
    Select Case FGrid.Row
        Case Date1, Date2, List1, List2, List3, List4, Cat4, Cat1, Cat2, Cat3
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
Dim I As Integer
    Select Case FGrid.Row
        Case Cat1, Cat2, Cat3
            If KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
            Else
                KeyAscii = 0
            End If
        Case Date1, Date2, List1, List2, List3, List4, Cat4
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
        Case Date1, Date2, List1, List2, List3, List4, Cat4
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
        Case Cat1, Cat2, Cat3
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
Private Sub Ini_Grid()
Dim Grid1Sql$, Grid2Sql$, Grid3Sql$, Grid4Sql$
 Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where site_code='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
Select Case GRepFormName
    Case SprStkLedVal
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Marked Part": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Include MRP": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List3, 0) = "Summary/Detail": .RowHeight(List3) = GridRowHeight
            
                       
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "No"
            .TextMatrix(List2, 1) = "Yes"
            .TextMatrix(List3, 1) = "Detail"
        End With
        
        mFirstRow = Date1: mLastRow = List3
        
        Grid1Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 1, Grid1Sql
        GridSel(1).width = GridSel(1).width + 1000: GridSel(1).ColWidth(1) = 1500: GridSel(1).ColWidth(1) = 3500
        
        Grid3Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 3, Grid3Sql
        GridSel(3).width = GridSel(3).width + 1000: GridSel(3).ColWidth(1) = 1500: GridSel(3).ColWidth(3) = 3500
        
    Case SprStkLedValSum
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            '.TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
'            .TextMatrix(List1, 0) = "Marked Part": .RowHeight(List1) = GridRowHeight
'            .TextMatrix(List2, 0) = "Include MRP": .RowHeight(List2) = GridRowHeight
'            .TextMatrix(List3, 0) = "Summary/Detail": .RowHeight(List3) = GridRowHeight
'            .TextMatrix(List4, 0) = "UnBilled(Y/N)": .RowHeight(List4) = GridRowHeight
'            .TextMatrix(Cat4, 0) = "Stock Val. Method": .RowHeight(Cat4) = GridRowHeight
            
            
            .TextMatrix(Date1, 1) = PubLoginDate   'PubStartDate
'            .TextMatrix(Date2, 1) = PubLoginDate
'            .TextMatrix(List1, 1) = "No"
'            .TextMatrix(List2, 1) = "Yes"
'            .TextMatrix(List3, 1) = "Detail"
'            .TextMatrix(List4, 1) = "No"
'            .TextMatrix(Cat4, 1) = "Separate Division"
            
        End With
        mFirstRow = Date1: mLastRow = Date1
                
        Grid1Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 1, Grid1Sql
        GridSel(1).width = GridSel(1).width + 1000: GridSel(1).ColWidth(1) = 1500: GridSel(1).ColWidth(1) = 3500
        
        Grid2Sql = "select '' as O,PartGrade_Name as Grade,PartGrade_code  as code from Part_Grade Order by PartGrade_Name"
        GridInitialise 2, Grid2Sql
        GridSel(2).width = GridSel(2).width + 1000: GridSel(2).ColWidth(2) = 1500: GridSel(2).ColWidth(2) = 3500
        
        Grid3Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 3, Grid3Sql
        GridSel(3).width = GridSel(3).width + 1000: GridSel(3).ColWidth(1) = 1500: GridSel(3).ColWidth(3) = 3500
        
        ChkOpStkOnly.Visible = True
    Case SprSaleInventory
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 2, Grid2Sql
        
        
    Case SprPartProfit
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Marked Part": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Type": .RowHeight(List2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "No"
            .TextMatrix(List2, 1) = "Summary"
        End With
        mFirstRow = Date2: mLastRow = List2
        
        Grid1Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 1, Grid1Sql
        
        GridSel(1).width = GridSel(1).width + 1000: GridSel(1).ColWidth(1) = 1500: GridSel(1).ColWidth(1) = 3500
        
        Grid3Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 3, Grid3Sql
        
        GridSel(3).width = GridSel(3).width + 1000: GridSel(3).ColWidth(1) = 1500: GridSel(3).ColWidth(3) = 3500
        
    Case SprFSNRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Include Receipts": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Fast/Slow/Dead/All": .RowHeight(List2) = GridRowHeight
            .TextMatrix(Cat1, 0) = "Fast %": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Cat2, 0) = "Slow %": .RowHeight(Cat2) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Yes"
            .TextMatrix(List2, 1) = "All"
            .TextMatrix(Cat1, 1) = ""  'All"
            .TextMatrix(Cat2, 1) = ""  'All"
        End With
        mFirstRow = Date1: mLastRow = Cat2
        
    Case SprProjection
        With FGrid
            .TextMatrix(Date2, 0) = "As on Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Marked Parts": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "For Stk Status": .RowHeight(List2) = GridRowHeight
            .TextMatrix(Cat1, 0) = "Last Days in View": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Cat2, 0) = "Projection for Day": .RowHeight(Cat2) = GridRowHeight
            .TextMatrix(Cat3, 0) = "For Amount >=": .RowHeight(Cat3) = GridRowHeight

            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "No"
            .TextMatrix(List2, 1) = "All"
            .TextMatrix(Cat1, 1) = ""  'All"
            .TextMatrix(Cat2, 1) = ""  'All"
            .TextMatrix(Cat3, 1) = "0.00"  'All"
        End With
        mFirstRow = Date1: mLastRow = Cat3
        
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select distinct '' as O,Part_no as PartNo,Part_no as code,Part_name as PartName from Part order by Part_no,part_name"
'        Grid3Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 3, Grid3Sql
        GridSel(3).width = GridSel(3).width + 1000: GridSel(3).ColWidth(1) = 1500: GridSel(3).ColWidth(3) = 3500
        
    Case SprXYZRep  '24-09 lps
        With FGrid
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Cat1, 0) = "X Cat %": .RowHeight(Cat1) = GridRowHeight
            .TextMatrix(Cat2, 0) = "Y Cat %": .RowHeight(Cat2) = GridRowHeight

            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(Cat1, 1) = ""  'All"
            .TextMatrix(Cat2, 1) = ""  'All"
        End With

        mFirstRow = Date1: mLastRow = Cat2
    
        Grid1Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 1, Grid1Sql
        GridSel(1).width = GridSel(1).width + 1000: GridSel(1).ColWidth(1) = 1500: GridSel(1).ColWidth(1) = 3500
        
        Grid3Sql = "Select Distinct '' as O,Part.Part_No,Part.Part_No as Code,Part.Part_Name as PartName From Part where Part.Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk) Order by Part_No"
        GridInitialise 3, Grid3Sql
        GridSel(3).width = GridSel(3).width + 1000: GridSel(3).ColWidth(1) = 1500: GridSel(3).ColWidth(3) = 3500
    
    Case SprABCRep
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
        
    Case DeleteLog
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Edit/Delete": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
        mFirstRow = Date1: mLastRow = List1
                
                
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Order by Div_Name"
        GridInitialise 2, Grid2Sql
                
        Grid3Sql = "Select Distinct '' as O,User_Name, User_Name as Code From DeleteLog Order by User_Name"
        GridInitialise 3, Grid3Sql
        
        Grid4Sql = "select Distinct '' as O,Type as Voucher_Type,Type as code from DeleteLog Order by Type"
        GridInitialise 4, Grid4Sql
        
        
End Select
End Sub
Private Sub Formulas()
On Error GoTo ELoop
Dim I As Integer


    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("Formulastr1")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr1 & "'"
            Case UCase("Formulastr2")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr2 & "'"
            Case UCase("Formulastr3")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr3 & "'"
            Case UCase("Formulastr4")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr4 & "'"
            Case UCase("RepTitle")
                rpt.FormulaFields(I).TEXT = "'Daily Sales Report-'+ '" & Format(FGrid.TextMatrix(Date1, 1), "mmm-yyyy") & "' "
            Case UCase("list1")
                rpt.FormulaFields(I).TEXT = " '" & FGrid.TextMatrix(List1, 1) & "'"
            Case UCase("List3")
                rpt.FormulaFields(I).TEXT = " '" & FGrid.TextMatrix(List3, 1) & "'"
                
        End Select
    Next
    FormulaStr1 = "": FormulaStr2 = "": FormulaStr3 = "": FormulaStr4 = ""



Select Case GRepFormName
'Case SprPurOrd
'    For i = 1 To rpt.FormulaFields.count
'    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
'        Case UCase("DATEBETWEEN")
'            rpt.FormulaFields(i).Text = "'Upto Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "'"
'    End Select
'    Next
Case SprABCRep
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("CatA")
                rpt.FormulaFields(I).TEXT = Val(FGrid.TextMatrix(Cat1, 1))
            Case UCase("CatB")
                rpt.FormulaFields(I).TEXT = Val(FGrid.TextMatrix(Cat2, 1))
            Case UCase("RepBase")
                rpt.FormulaFields(I).TEXT = "'Formula of %   = (Consumption Value of each Item *100)/Total Consumption Value'"
            Case UCase("RepBase2")
                rpt.FormulaFields(I).TEXT = "'Category  A = Top " & Val(FGrid.TextMatrix(Cat1, 1)) & "%,  B = Next " & Val(FGrid.TextMatrix(Cat2, 1)) & "%,  C = Remaining consumption value'"
        End Select
    Next
Case SprFSNRep
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("Fast")
                rpt.FormulaFields(I).TEXT = "'Fast :  >=" & FGrid.TextMatrix(Cat1, 1) & "'"
            Case UCase("Slow")
                rpt.FormulaFields(I).TEXT = "'Slow :  >=" & FGrid.TextMatrix(Cat2, 1) & " and <" & FGrid.TextMatrix(Cat1, 1) & "'"
            Case UCase("Dead")
                rpt.FormulaFields(I).TEXT = "'Dead : <" & FGrid.TextMatrix(Cat2, 1) & "'"
            Case UCase("RepBase")
                rpt.FormulaFields(I).TEXT = IIf(FGrid.TextMatrix(List1, 1) = "Yes", "'Movement % includes Receipts'", "'Movement % includes without Receipts'")
            Case UCase("RepBase2")
                rpt.FormulaFields(I).TEXT = "'For " & FGrid.TextMatrix(List2, 1) & " Type Movement, Valuation Based on FIFO Method'"
        End Select
    Next
Case SprXYZRep
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'As on Date " & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("Fast")
                rpt.FormulaFields(I).TEXT = "'X Cat :  >=" & FGrid.TextMatrix(Cat1, 1) & "'"
            Case UCase("Slow")
                rpt.FormulaFields(I).TEXT = "'Y Cat :  >=" & FGrid.TextMatrix(Cat2, 1) & " and <" & FGrid.TextMatrix(Cat1, 1) & "'"
            Case UCase("Dead")
                rpt.FormulaFields(I).TEXT = "'Z Cat : <" & FGrid.TextMatrix(Cat2, 1) & "'"
            Case UCase("RepBase")
                rpt.FormulaFields(I).TEXT = IIf(FGrid.TextMatrix(List1, 1) = "Yes", "'Marked Parts Only'", IIf(Check1(3).Value = Unchecked, "'Only For Selected Parts'", "'For All Parts'"))
            Case UCase("RepBase2")
                rpt.FormulaFields(I).TEXT = "'Valuation Based on FIFO Method'"
        End Select
    Next

Case SprStkLedVal, SprStkLedValSum
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("PartSelect")
                rpt.FormulaFields(I).TEXT = "'" & IIf(FGrid.TextMatrix(List1, 1) = "Yes", "For Marked Parts", IIf(Check1(3).Value = Unchecked, "For Selective Parts", "For All Parts")) & "'"
            Case UCase("MRP_YN")
                rpt.FormulaFields(I).TEXT = "'" & IIf(FGrid.TextMatrix(List2, 1) = "Yes", "For MRP Stock", IIf(FGrid.TextMatrix(List2, 1) = "No", "For Non-MRP Stock", "For MRP and Non-MRP Stock")) & "'"
        End Select
    Next
Case SprProjection
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'As on Date " & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("PartSelect")
                rpt.FormulaFields(I).TEXT = "'" & IIf(FGrid.TextMatrix(List1, 1) = "Yes", "For Marked Parts", IIf(Check1(3).Value = Unchecked, "For Selective Parts", "For All Parts")) & "'"
            Case UCase("ForAmount")
                If Val(FGrid.TextMatrix(Cat3, 1)) <> 0 Then
                    rpt.FormulaFields(I).TEXT = "'For Short/Excess Value is  >=" & FGrid.TextMatrix(Cat3, 1) & "'"
                End If
            Case UCase("StkStatus")
                rpt.FormulaFields(I).TEXT = "'" & IIf(FGrid.TextMatrix(List2, 1) = "All", "For Short And Excess in Stock", IIf(FGrid.TextMatrix(List2, 1) = "Short", "For Shortage in Stock", "For Excess Stock")) & "'"
        End Select
    Next

Case SprPartProfit
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("PartSelect")
                rpt.FormulaFields(I).TEXT = "'" & IIf(FGrid.TextMatrix(List1, 1) = "Yes", "For Marked Parts", IIf(Check1(3).Value = Unchecked, "For Selective Parts", "For All Parts")) & "'"
        End Select
    Next
Case SprSaleInventory
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
Private Sub SprABCAnalysis()
On Error GoTo ELoop
Dim mQry$, Condstr$
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    If IsInLimit(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsInLimit(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    If Val(FGrid.TextMatrix(Cat2, 1)) > Val(FGrid.TextMatrix(Cat1, 1)) Then
        MsgBox FGrid.TextMatrix(Cat2, 0) & " >" & FGrid.TextMatrix(Cat1, 0), vbOKOnly, "Validation"
        FGrid.SetFocus:  FGrid.Row = Cat2: FGrid.Col = 1
        RepPrint = False: Exit Sub
    
    End If
'   rEMOVED PART FROM TRHE BELOW QUERY
'   and mid(SP.Invoice_DocId,4,5) in ('" & SprSlCsh & "', '" & SprSlCre & "', '" & WksSlCsh & "', '" & WksSlCre & "') "
'    and SP.Purpose<>'W' "
    mQry = "Select SP.Part_No,P.Part_Name,sum(SP.Amount) as NetAmt2 from (SP_Stock SP Left Join Part P on SP.Part_No=P.Part_No) " & _
        "where  Qty_Iss > 0 " & _
        " and SP.V_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & _
        " and SP.V_Date<= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & _
        " and P.Div_Code='" & PubDivCode & _
        "' Group by SP.Part_No,P.Part_Name Order by sum(Amount) Desc"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprABCRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    MsgBox err.Description
End Sub
Private Sub SprStkLedValCalc()
On Error GoTo ELoop
Dim mQry$, Condstr$, CondDivCode$, CondDivCode1$, Condstr2$, CondStrMRP$, CondMarkYN$
Dim CondPartNos$, CondPartNos1$, CondPartNosOpStk$
Dim mRecQty As Double, mIssQty As Double, mStkVal As Double
Dim XRecNo As Double
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    Condstr = "where SPStk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    If Check1(1).Value = Unchecked Then
        CondDivCode = " and left(SPStk.DocID,1) in (" & GridString1 & ")"
        CondDivCode1 = " and left(Stk.DocID,1) in (" & GridString1 & ")"
    End If
    
    If FGrid.TextMatrix(List1, 1) = "Yes" Then          '' Only for Marked Parts
        Condstr2 = " and Mark_yn='Y'"
    Else
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and SPStk.Part_No in (" & GridString3 & ")"
    End If
        Condstr = Condstr & CondDivCode
    If FGrid.TextMatrix(List1, 1) = "Yes" Then          '' Only for Marked Parts
        CondMarkYN = " and Mark_YN='Y'"
    Else
        If Check1(3).Value = Unchecked Then
            CondPartNos = " and SPStk.Part_No in (" & GridString3 & ")"
            CondPartNos1 = " Part.Part_No in " & "(" & GridString3 & ")"
            CondPartNosOpStk = CondPartNos
        Else
            CondPartNos = " and Part_No in (select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
            CondPartNos1 = " Part.Part_No in ( select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
            CondPartNosOpStk = ""
        End If
    End If
    If FGrid.TextMatrix(List2, 1) = "No" Then       '' For MRP_YN
        CondStrMRP = " and SPStk.MRP_YN=0"
    End If
    Condstr = Condstr & CondStrMRP


    If FGrid.TextMatrix(List2, 1) = "No" Then          '' Only for MRP Parts
        CondDivCode = CondDivCode & " and MRP_yn=0"    '' Only for Non-MRP Parts
    End If

    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
        
    Set TRec1 = New ADODB.Recordset
    Set TRec1 = TmpTRec1(TRec1)
    
    Set TRec2 = New ADODB.Recordset
    Set TRec2 = TmpTRec1(TRec2)
    
'    Set RstPart = GCn.Execute("Select Distinct SPStk.Part_No,Part.Part_Name From SP_Stock as SPStk Left Join Part On SPStk.Part_No=Part.Part_No and Part.Div_Code = left(SPStk.Docid,1) " & CondStr)
    GSQL = "Select Part.Part_No,Part.Part_Name From Part where "
    If Check1(1).Value = Unchecked Then
        GSQL = "Select Part.Part_No,Part.Part_Name From Part where Part.Div_Code in (" & GridString1 & ") " & Condstr2 & _
                " and Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk " & Condstr & ") Order By Part.Part_No"
    Else
        GSQL = "Select Distinct Part.Part_No,Part.Part_Name From Part where " & _
                "Part_No in (select Distinct SPStk.Part_No from SP_Stock as SPStk " & Condstr & ") " & Condstr2 & " Order By Part.Part_No"
    End If
    Set RstPart = GCn.Execute(GSQL)
    Do While Not RstPart.EOF = True
        Do While TRec1.RecordCount > 0
           If TRec1.RecordCount > 0 Then TRec1.MoveFirst
           TRec1.Delete
           TRec1.Update
        Loop
        Do While TRec2.RecordCount > 0
           If TRec2.RecordCount > 0 Then TRec2.MoveFirst
           TRec2.Delete
           TRec2.Update
        Loop
        
        mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
'        Set RstStock = GCn.Execute("select SPStk.*,Vt.Stktrn " & _
                                "From Sp_Stock left Join [" & PubSFADataPath & "].Voucher_Type vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn")
        GSQL = "select SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,SPStk.Qty_Iss,SPStk.Qty_Ret,SPStk.v_rate,Vt.Stktrn From SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' "
        If Check1(1).Value = Unchecked Then
            GSQL = GSQL & "and left(SPStk.Docid,1) in (" & GridString1 & ") and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
        Else
            GSQL = GSQL & "and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
        End If
        Set RstStock = GCn.Execute(GSQL)
        Do While Not RstStock.EOF = True
            '' Add Record for Received Side
            If RstStock!StkTrn = "+" Then
                If RstStock!Tax_YN = 1 Then     '' Taxable
                    With TRec1
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("Part_No") = RstStock!Part_No
                        .Fields("Qty") = RstStock!Qty_Rec
                        .Fields("Rate") = RstStock!V_Rate
                        .Update
                    End With
                Else
                    With TRec2
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("Part_No") = RstStock!Part_No
                        .Fields("Qty") = RstStock!Qty_Rec
                        .Fields("Rate") = RstStock!V_Rate
                        .Update
                    End With
                End If
            End If
            RstStock.MoveNext
        Loop
        mTrf = False
        
        TRec1.Sort = "Date"
        TRec2.Sort = "Date"
        
        If TRec1.RecordCount > 0 Then TRec1.MoveFirst
        If TRec2.RecordCount > 0 Then TRec2.MoveFirst
        
'        Set RstStock = GCn.Execute("select SPStk.*,Vt.Stktrn From Sp_Stock left Join [" & PubSFADataPath & "].Voucher_Type vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' and v_date<" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn")
        GSQL = "select SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,SPStk.Qty_Iss,SPStk.Qty_Ret,SPStk.v_rate,Vt.Stktrn From SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & "   vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' "
        If Check1(1).Value = Unchecked Then
            GSQL = GSQL & " and left(SPStk.Docid,1) in (" & GridString1 & ") and v_date<" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
        Else
            GSQL = GSQL & " and v_date<" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
        End If
        Set RstStock = GCn.Execute(GSQL)
        Do While Not RstStock.EOF
            mTrf = True
            If RstStock!StkTrn = "-" Then
                If RstStock!Tax_YN = 1 Then     '' Taxable
                    mRate = 0
                    Call X_Val1(Temp06, TRec1, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate)
                Else
                    mRate = 0
                    Call X_Val2(Temp06, TRec2, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate)
                End If
            ElseIf RstStock!StkTrn = "+" Then
                If RstStock!Tax_YN = 1 Then     '' Taxable
                    mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                    mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                Else
                    mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                    mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                End If
            End If
            RstStock.MoveNext
        Loop
        
        If mOP_TB_QTY <> 0 Then
            If TRec1.RecordCount > 0 Then
                With TRec1
                    XRecNo = .AbsolutePosition
                    .MoveFirst
                    .Fields("Date") = Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")
                    .Fields("Part_No") = RstPart!Part_No
                    .Update
                    .Bookmark = XRecNo
                End With
            End If
        End If
        If mOP_TP_QTY <> 0 Then
            If TRec2.RecordCount > 0 Then
                With TRec2
                    XRecNo = .AbsolutePosition
                    .MoveFirst
                    .Fields("Date") = Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")
                    .Fields("Part_No") = RstPart!Part_No
                    .Update
                    .Bookmark = XRecNo
                End With
            End If
        End If
    
    
        mTrf = False: mPART_ADD = False
        If mOP_TB_QTY <> 0 Or mOP_TP_QTY <> 0 Then
            mPART_ADD = True
            With Temp06
                .AddNew
                .Fields("Part_Name") = RstPart!Part_Name
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Job_Age") = "Y"
                
                .AddNew
                .Fields("Part_Name") = "Opening Balance"
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Re_TB") = 0
                .Fields("Is_TB") = 0
                .Fields("Tb_Val") = 0
                .Fields("tb_Bqty") = mOP_TB_QTY
                .Fields("tb_BVal") = mOP_TB_VAL
                
                .Fields("Re_TP") = 0
                .Fields("Is_TP") = 0
                .Fields("TP_Val") = 0
                .Fields("TP_Bqty") = mOP_TP_QTY
                .Fields("TP_BVal") = mOP_TP_VAL
                .Update
            End With
        End If
        
'        Set RstStock = GCn.Execute("select SPStk.*,Vt.Stktrn,Vt.Description From Sp_Stock left Join [" & PubSFADataPath & "].Voucher_Type vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' and v_date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn")
        GSQL = "select (H.REGNO & ' ' & H.chassis) as RegChassis,SPStk.V_DATE,SPStk.V_Type,SPStk.DocId,SPStk.V_NO,SPStk.job_docid,SPStk.Party_code,SG.Name,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,SPStk.Qty_Iss,SPStk.Qty_Ret,SPStk.Invoice_DocID,SPStk.v_date2,SPStk.v_rate,Vt.Stktrn,Vt.Description " & _
                "From ((((SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=SPStk.V_type) " & _
                "Left Join SubGroup as SG on SPStk.Party_Code=SG.SubCode) " & _
                "Left Join Job_Card as J on SPStk.job_docid=J.DocID)" & _
                "Left Join Hiscard as H on J.CardNo=H.CardNo) " & _
                "where Part_No='" & RstPart!Part_No & "' "
        
        'GCn.Execute("select (H.REGNO & ' ' & H.chassis) as RegChassis From Job_Card Left Join Hiscard H on Job_card.CardNo=H.CardNo where Job_Card.DocId='" & RstStock!job_docid & "'").Fields(0)
        
        If Check1(1).Value = Unchecked Then
            GSQL = GSQL & "and left(SPStk.Docid,1) in (" & GridString1 & ") and v_date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
        Else
            GSQL = GSQL & "and v_date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
        End If
        Set RstStock = GCn.Execute(GSQL)
        Do While Not RstStock.EOF = True
            mname = IIf(IsNull(RstStock!Name), "", RstStock!Name)
'            If Not IsNull(RstStock!Party_code) Then
'                mname = GCn.Execute("select iif(isnull(Max(Name)),'',Max(Name)) From SubGroup where Subcode='" & RstStock!Party_code & "'").Fields(0).Value
'            End If
            mInv_No = "N.A."
            mInv_Date = "N.A."
            mNarr = RstStock!Description
            If RstStock!V_Type = WksGenReq Or RstStock!V_Type = WksReqWrt Then
'                mInv_No = DeCodeDocID(RstStock!Invoice_DocID, Document_No)
                mInv_No = PrinID(RstStock!Invoice_DocID)
                mInv_Date = IIf(IsNull(RstStock!V_DATE2), "", Format(RstStock!V_DATE2, "dd/mm/yyyy"))
                mname = IIf(IsNull(RstStock!RegChassis), "", RstStock!RegChassis)
'                mname = GCn.Execute("select H.REGNO+' '+H.chassis From Job_Card Left Join Hiscard H on Job_card.CardNo=H.CardNo where Job_Card.DocId='" & RstStock!job_docid & "'").Fields(0)
            End If
            If RstStock!StkTrn = "-" Then
                If RstStock!Tax_YN = 1 Then     '' Taxable
                    mRate = 0
                    Call X_Val1(Temp06, TRec1, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate, mNarr)
                Else
                    mRate = 0
                    Call X_Val2(Temp06, TRec2, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate, mNarr)
                End If
            ElseIf RstStock!StkTrn = "+" Then
                If RstStock!Tax_YN = 1 Then     '' Taxable
                    mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                    mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = PrinID(RstStock!DocID) 'RstStock!V_Type + "-" + str(RstStock!v_no)
                        .Fields("Narr") = left(mNarr, 25)
                        .Fields("Part_Name") = mname
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Rate") = RstStock!V_Rate
                        
                        .Fields("re_Tb") = RstStock!Qty_Rec
                        .Fields("Tb_Val") = RstStock!Qty_Rec * RstStock!V_Rate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("re_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                                            
                        .Update
                    End With
                Else
                    mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                    mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = PrinID(RstStock!DocID) 'RstStock!V_Type + "-" + str(RstStock!v_no)
                        .Fields("Narr") = mNarr
                        .Fields("Part_Name") = mname
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Rate") = RstStock!V_Rate
                        
                        .Fields("re_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("re_Tp") = RstStock!Qty_Rec
                        .Fields("Tp_Val") = RstStock!Qty_Rec * RstStock!V_Rate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        
                        .Update
                    End With
                End If
            End If
            RstStock.MoveNext
        Loop
        RstPart.MoveNext
    Loop
    
    Set RstRep = Temp06.Clone
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "StockLedValue"
    RepTitle = UCase(Me.CAPTION)
ELoop:
    If err.NUMBER <> 0 Then CheckError
    Set GRs = Nothing
End Sub

Private Sub X_Val1(ByRef Temp06 As ADODB.Recordset, ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        xRate = 0
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        Exit Sub
    End If
    If xQty = TRec1!Qty Then
        TRec1.Fields("QTY") = 0
        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        TRec1.MoveNext
    ElseIf xQty < TRec1!Qty Then
        TRec1.Fields("QTY") = TRec1!Qty - xQty
        TRec1.Update
        
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
    ElseIf xQty > TRec1!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
            If TRec1!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec1!Qty <= TQty Then
                TQty = TQty - TRec1!Qty
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = left(xNARR, 25)
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TRec1!Qty
                        .Fields("Tb_Val") = TRec1!Qty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TRec1.Fields("QTY") = 0
                    TRec1.Update
                End If
            Else
                TRec1.Fields("QTY") = TRec1!Qty - TQty
                TRec1.Update
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TQty
                        .Fields("Tb_Val") = TQty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = TQty
                    .Fields("Tb_Val") = TQty * xRate
                    .Fields("Tb_BQty") = mOP_TB_QTY
                    .Fields("Tb_BVal") = mOP_TB_VAL
                    
                    .Fields("Is_Tp") = 0
                    .Fields("Tp_Val") = 0
                    .Fields("Tp_BQty") = 0
                    .Fields("Tp_BVal") = 0
                    .Update
                End With
            
            End If
        Loop
    End If
End Sub

Private Sub X_Val2(ByRef Temp06 As ADODB.Recordset, ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        xRate = 0
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = PrinID(RstStock!DocID)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        Exit Sub
    End If
    
    If xQty = TRec2!Qty Then
        TRec2.Fields("QTY") = 0
        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        TRec2.MoveNext
    ElseIf xQty < TRec2!Qty Then
        TRec2.Fields("QTY") = TRec2!Qty - xQty
        TRec2.Update
        
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                .Update
            End With
        End If
    ElseIf xQty > TRec2!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
            If TRec2!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec2!Qty <= TQty Then
                TQty = TQty - TRec2!Qty
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = xNARR
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TRec2!Qty
                        .Fields("Tp_Val") = TRec2!Qty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TRec2.Fields("QTY") = 0
                    TRec2.Update
                End If
            Else
                TRec2.Fields("QTY") = TRec2!Qty - TQty
                TRec2.Update
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TQty
                        .Fields("Tp_Val") = TQty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = 0
                    .Fields("Tb_Val") = 0
                    .Fields("Tb_BQty") = 0
                    .Fields("Tb_BVal") = 0
                    
                    .Fields("Is_Tp") = TQty
                    .Fields("Tp_Val") = TQty * xRate
                    .Fields("Tp_BQty") = mOP_TP_QTY
                    .Fields("Tp_BVal") = mOP_TP_VAL
                    .Update
                End With
            End If
        Loop
    End If
End Sub
Private Sub SprStkLedValSumCalc()
On Error GoTo ELoop
Dim mQry$, Condstr$, CondDivCode$, CondMarkYN$, CondPartNos$, CondPartNos1$, CondDivCode1$
Dim CondStrMRP$, CondPartNosOpStk$, CondPartNosTrn$, Part_Name$
Dim mRecQty As Double, mIssQty As Double, mStkVal As Double
Dim XRecNo As Double, DivStr$
Dim mNo As Byte, NoUpto As Byte
    RepPrint = True
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

    Condstr = " where SPStk.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    
    If Check1(1).Value = Unchecked Then
        CondDivCode = " and left(SPStk.DocID,1) in (" & GridString1 & ")"
        CondDivCode1 = " and left(Stk.DocID,1) in (" & GridString1 & ")"
    End If
    
    Condstr = Condstr & CondDivCode
    If FGrid.TextMatrix(List1, 1) = "Yes" Then          '' Only for Marked Parts
        CondMarkYN = " and Mark_YN='Y'"
    Else
       If Check1(3).Value = Unchecked Then
            CondPartNos = " and SPStk.Part_No in (" & GridString3 & ")"
            CondPartNos1 = " Part.Part_No in " & "(" & GridString3 & ")"
            CondPartNosOpStk = CondPartNos
        Else
            CondPartNos = " and SPStk.Part_No in (select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
            CondPartNos1 = " Part.Part_No in ( select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
            CondPartNosOpStk = ""
        End If
    
    End If
    ' Properitary grade check *******
    If Check1(2).Value = Unchecked Then
             CondPartNos1 = CondPartNos1 & " and Part.Part_Grade in (" & GridString2 & ")"
    End If
    '***************
    
    ' Bill Status Check *******
    If FGrid.TextMatrix(List4, 1) = "Yes" Then
        If FGrid.TextMatrix(List3, 1) = "Detail" Or FGrid.TextMatrix(List3, 1) = "Summary" Then
            CondPartNos = CondPartNos & " and SPStk.Invoice_DocID = '' and SPStk.V_Type in('W_RG','SYSC')"
        End If
    End If
    ' ************
    If FGrid.TextMatrix(List2, 1) = "No" Then       '' For MRP_YN
         CondStrMRP = " and SPStk.MRP_YN=0"
    End If
    
    Condstr = Condstr & CondStrMRP
    
    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
    
    'For RstPart, SQL
    GSQL = "Select Distinct Part.Part_No From Part " & _
        "where " & CondPartNos1
    If Check1(1).Value = Unchecked Then
        GSQL = GSQL & IIf(CondPartNos1 = "", "", " and") & " Part.Div_Code in (" & GridString1 & ") "
        DivStr = "and Part.Div_Code in(" & GridString1 & ")"
    Else
        DivStr = ""
    End If
    GSQL = GSQL & CondMarkYN & " Order By Part.Part_No"
    
    Set RstPart = GCn.Execute(GSQL)
    'Process Stock Ledger FIFO
    If GRepFormName = SprStkLedVal Then
        RepName = "StockLedValue"
        Set TRec1 = New ADODB.Recordset
        Set TRec1 = TmpTRec1(TRec1)
        
        Set TRec2 = New ADODB.Recordset
        Set TRec2 = TmpTRec1(TRec2)
        Do While Not RstPart.EOF = True
            Do While TRec1.RecordCount > 0
               If TRec1.RecordCount > 0 Then TRec1.MoveFirst
               TRec1.Delete
               TRec1.Update
            Loop
            Do While TRec2.RecordCount > 0
               If TRec2.RecordCount > 0 Then TRec2.MoveFirst
               TRec2.Delete
               TRec2.Update
            Loop
            
            mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
            GSQL = "select SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,SPStk.Qty_Iss,SPStk.Qty_Ret,SPStk.v_rate,Vt.Stktrn From SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' "
            If Check1(1).Value = Unchecked Then
                GSQL = GSQL & "and left(SPStk.Docid,1) in (" & GridString1 & ") and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
            Else
                GSQL = GSQL & "and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
            End If
            Set RstStock = GCn.Execute(GSQL)
            Do While Not RstStock.EOF = True
                '' Add Record for Received Side
                If RstStock!StkTrn = "+" Then
                    If RstStock!Tax_YN = 1 Then     '' Taxable
                        With TRec1
                            .AddNew
                            .Fields("Date") = RstStock!V_DATE
                            .Fields("Part_No") = RstStock!Part_No
                            .Fields("Qty") = RstStock!Qty_Rec
                            .Fields("Rate") = RstStock!V_Rate
                            .Update
                        End With
                    Else
                        With TRec2
                            .AddNew
                            .Fields("Date") = RstStock!V_DATE
                            .Fields("Part_No") = RstStock!Part_No
                            .Fields("Qty") = RstStock!Qty_Rec
                            .Fields("Rate") = RstStock!V_Rate
                            .Update
                        End With
                    End If
                End If
                RstStock.MoveNext
            Loop
            mTrf = False
        
            TRec1.Sort = "Date"
            TRec2.Sort = "Date"
        
            If TRec1.RecordCount > 0 Then TRec1.MoveFirst
            If TRec2.RecordCount > 0 Then TRec2.MoveFirst
        
            GSQL = "select SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,SPStk.Qty_Iss,SPStk.Qty_Ret,SPStk.v_rate,Vt.Stktrn From SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' "
            If Check1(1).Value = Unchecked Then
                GSQL = GSQL & " and left(SPStk.Docid,1) in (" & GridString1 & ") and v_date<" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
            Else
                GSQL = GSQL & " and v_date<" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
            End If
            Set RstStock = GCn.Execute(GSQL)
            Do While Not RstStock.EOF
                If RstStock!Part_No = "1540" Then
                    MsgBox "sss"
                End If
                mTrf = True
                If RstStock!StkTrn = "-" Then
                    If RstStock!Tax_YN = 1 Then     '' Taxable
                        mRate = 0
                        Call X_Val1(Temp06, TRec1, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate)
                    Else
                        mRate = 0
                        Call X_Val2(Temp06, TRec2, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate)
                    End If
                ElseIf RstStock!StkTrn = "+" Then
                    If RstStock!Tax_YN = 1 Then     '' Taxable
                        mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                        mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                    Else
                        mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                        mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                    End If
                End If
                RstStock.MoveNext
            Loop
            If mOP_TB_QTY <> 0 Then
                If TRec1.RecordCount > 0 Then
                    With TRec1
                        XRecNo = .AbsolutePosition
                        .MoveFirst
                        .Fields("Date") = Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")
                        .Fields("Part_No") = RstPart!Part_No
                        .Update
                        .Bookmark = XRecNo
                    End With
                End If
            End If
            If mOP_TP_QTY <> 0 Then
                If TRec2.RecordCount > 0 Then
                    With TRec2
                        XRecNo = .AbsolutePosition
                        .MoveFirst
                        .Fields("Date") = Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")
                        .Fields("Part_No") = RstPart!Part_No
                        .Update
                        .Bookmark = XRecNo
                    End With
                End If
            End If
            mTrf = False: mPART_ADD = False
            If mOP_TB_QTY <> 0 Or mOP_TP_QTY <> 0 Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    
                    .AddNew
                    .Fields("Part_Name") = "Opening Balance"
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Re_TB") = 0
                    .Fields("Is_TB") = 0
                    .Fields("Tb_Val") = 0
                    .Fields("tb_Bqty") = mOP_TB_QTY
                    .Fields("tb_BVal") = mOP_TB_VAL
                    
                    .Fields("Re_TP") = 0
                    .Fields("Is_TP") = 0
                    .Fields("TP_Val") = 0
                    .Fields("TP_Bqty") = mOP_TP_QTY
                    .Fields("TP_BVal") = mOP_TP_VAL
                    .Update
                End With
            End If
            GSQL = "select (H.REGNO & ' ' & H.chassis) as RegChassis,SPStk.V_DATE,SPStk.V_Type,SPStk.DocId,SPStk.V_NO,SPStk.job_docid,SPStk.Party_code,SG.Name,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,SPStk.Qty_Iss,SPStk.Qty_Ret,SPStk.Invoice_DocID,SPStk.v_date2,SPStk.v_rate,Vt.Stktrn,Vt.Description " & _
                "From ((((SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=SPStk.V_type) " & _
                "Left Join SubGroup as SG on SPStk.Party_Code=SG.SubCode) " & _
                "Left Join Job_Card as J on SPStk.job_docid=J.DocID)" & _
                "Left Join Hiscard as H on J.CardNo=H.CardNo) " & _
                "where Part_No='" & RstPart!Part_No & "' "
            If Check1(1).Value = Unchecked Then
                GSQL = GSQL & "and left(SPStk.Docid,1) in (" & GridString1 & ") and v_date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
            Else
                GSQL = GSQL & "and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode & " order By SPStk.V_Date,Vt.StkTrn"
            End If
            Set RstStock = GCn.Execute(GSQL)
            Do While Not RstStock.EOF = True
                mname = IIf(IsNull(RstStock!Name), "", RstStock!Name)
                mInv_No = "N.A."
                mInv_Date = "N.A."
                mNarr = RstStock!Description
                If RstStock!V_Type = WksGenReq Or RstStock!V_Type = WksReqWrt Then
                    mInv_No = PrinID(RstStock!Invoice_DocID)
                    mInv_Date = IIf(IsNull(RstStock!V_DATE2), "", Format(RstStock!V_DATE2, "dd/mm/yyyy"))
                    mname = IIf(IsNull(RstStock!RegChassis), "", RstStock!RegChassis)
                End If
                If RstStock!StkTrn = "-" Then
                    If RstStock!Tax_YN = 1 Then     '' Taxable
                        mRate = 0
                        Call X_Val1(Temp06, TRec1, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate, mNarr)
                    Else
                        mRate = 0
                        Call X_Val2(Temp06, TRec2, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate, mNarr)
                    End If
                ElseIf RstStock!StkTrn = "+" Then
                    If RstStock!Tax_YN = 1 Then     '' Taxable
                        mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                        mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                        If mPART_ADD = False Then
                            mPART_ADD = True
                            With Temp06
                                .AddNew
                                .Fields("Part_Name") = RstPart!Part_Name
                                .Fields("Part_No") = RstPart!Part_No
                                .Fields("Job_Age") = "Y"
                                .Update
                            End With
                        End If
                        With Temp06
                            .AddNew
                            .Fields("Date") = RstStock!V_DATE
                            .Fields("V_No") = PrinID(RstStock!DocID) 'RstStock!V_Type + "-" + str(RstStock!v_no)
                            .Fields("Narr") = left(mNarr, 25)
                            .Fields("Part_Name") = mname
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Inv_No") = mInv_No
                            .Fields("Inv_Date") = mInv_Date
                            .Fields("Rate") = RstStock!V_Rate
                            
                            .Fields("re_Tb") = RstStock!Qty_Rec
                            .Fields("Tb_Val") = RstStock!Qty_Rec * RstStock!V_Rate
                            .Fields("Tb_BQty") = mOP_TB_QTY
                            .Fields("Tb_BVal") = mOP_TB_VAL
                            
                            .Fields("re_Tp") = 0
                            .Fields("Tp_Val") = 0
                            .Fields("Tp_BQty") = 0
                            .Fields("Tp_BVal") = 0
                                                
                            .Update
                        End With
                    Else
                        mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                        mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                        If mPART_ADD = False Then
                            mPART_ADD = True
                            With Temp06
                                .AddNew
                                .Fields("Part_Name") = RstPart!Part_Name
                                .Fields("Part_No") = RstPart!Part_No
                                .Fields("Job_Age") = "Y"
                                .Update
                            End With
                        End If
                        
                        With Temp06
                            .AddNew
                            .Fields("Date") = RstStock!V_DATE
                            .Fields("V_No") = PrinID(RstStock!DocID) 'RstStock!V_Type + "-" + str(RstStock!v_no)
                            .Fields("Narr") = mNarr
                            .Fields("Part_Name") = mname
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Inv_No") = mInv_No
                            .Fields("Inv_Date") = mInv_Date
                            .Fields("Rate") = RstStock!V_Rate
                            
                            .Fields("re_Tb") = 0
                            .Fields("Tb_Val") = 0
                            .Fields("Tb_BQty") = 0
                            .Fields("Tb_BVal") = 0
                            
                            .Fields("re_Tp") = RstStock!Qty_Rec
                            .Fields("Tp_Val") = RstStock!Qty_Rec * RstStock!V_Rate
                            .Fields("Tp_BQty") = mOP_TP_QTY
                            .Fields("Tp_BVal") = mOP_TP_VAL
                            
                            .Update
                        End With
                    End If
                End If
                RstStock.MoveNext
            Loop
            RstPart.MoveNext
        Loop
    Else    'Stock Valuation FIFO Summary
        If FGrid.TextMatrix(List3, 1) = "Summary" Then
            RepName = "StockLedValueSum"
        ElseIf FGrid.TextMatrix(List3, 1) = "Detail" Then
            RepName = "StockLedValueDet"
        End If
        '********** Taxable Qty
        mQry = "select SPStk.Part_No,SPStk.V_DATE,SPStk.Qty_Rec as Qty,SPStk.MRP_YN,SPStk.V_Rate as Rate " & _
            "From " & _
            "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
            Condstr & CondPartNos
        GSQL = mQry & " and SpStk.Tax_YN=1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
        Set TRec1 = New Recordset
        With TRec1
            .CursorLocation = adUseClient
            .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
        End With
        '******* Taxpaid Qty
        GSQL = mQry & " and SpStk.Tax_YN<>1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
        Set TRec2 = New Recordset
        With TRec2
            .CursorLocation = adUseClient
            .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
        End With
        '******* Taxable + Taxpaid Qty for Opening Loop
        
        mQry = "select SPStk.V_Type,SPStk.Part_No,SPStk.V_DATE,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate,Vt.StkTrn " & _
            "From " & _
            "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & "  as VT on Vt.V_type=SPStk.V_type " & _
            Condstr & CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
        GSQL = mQry & " and SPStk.V_Type='SXAO' Order By SPStk.Part_No,SPStk.V_Date," & cMID("SPStk.DocID", "4", "5") & ""
        Set RstStock = GCn.Execute(GSQL)
        '******* Taxable + Taxpaid Qty for With in Date Period Loop
        GSQL = "select SPStk.V_Type,SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate,Vt.StkTrn,Vt.Description " & _
            "From " & _
            "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
            "where (SPStk.V_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And SPStk.v_date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " ) " & _
            CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
        GSQL = GSQL & " Order By SPStk.Part_No,SPStk.V_Date,SPStk.Tax_YN, " & cMID("SPStk.DocID", "4", "5") & ""
        Set RstStock2 = GCn.Execute(GSQL)
        '***********
        Set tempRst = RstStock.Clone
        Set TempRst1 = RstStock2.Clone
        Dim I As Integer
         MDIForm1.Picture1.Visible = True
         Do While Not RstPart.EOF
            'NRA Update
            MDIForm1.Label1.CAPTION = "Process Status : " & RstPart.AbsolutePosition & "/" & RstPart.RecordCount
            MDIForm1.Label1.Refresh
            mVRate = 0: mOPVRate = 0: mDisPer = 0: TempVal = 0
'                'For Opening Calculate
'                tempRst.Filter = ("Part_No='" & RstPart!Part_No & "'")
'                If tempRst.RecordCount  > 0 Then
'                    tempRst.Sort = "V_Date Asc"
'                    tempRst.MoveFirst
'                    If tempRst!V_Rate <> 0 Then
'                        mOPVRate = tempRst!V_Rate
'                    Else
'                        If tempRst!MRP_YN = 1 Then
'                            mOPVRate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
'                        Else
'                            If tempRst!Tax_YN = 1 Then
'                                mOPVRate = GCn.Execute("Select TB_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
'                            Else
'                                mOPVRate = GCn.Execute("Select TP_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
'                            End If
'                            mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
'                            mOPVRate = mOPVRate - ((mOPVRate * mDisPer) / 100)
'                        End If
'                    End If
'                End If
'
'                'For NON Opening Calculate
'                TempRst1.Filter = ("Part_No='" & RstPart!Part_No & "'")
'                If TempRst1.RecordCount  > 0 Then
'                    TempRst1.Sort = "V_Date Asc"
'                    TempRst1.MoveFirst
'                    If TempRst1!V_Rate <> 0 Then
'                        mVrate = TempRst1!V_Rate
'                    Else
'                        If TempRst1!MRP_YN = 1 Then
'                            mVrate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
'                        Else
'                            If TempRst1!Tax_YN = 1 Then
'                                mVrate = GCn.Execute("Select TB_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
'                            Else
'                                mVrate = GCn.Execute("Select TP_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
'                            End If
'                            mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
'                            mVrate = mVrate - ((mVrate * mDisPer) / 100)
'                        End If
'                    End If
'                End If
        
            If FGrid.TextMatrix(List2, 1) = "Yes" Then          '' Only for MRP Parts
                NoUpto = 1
                mNo = 1
            ElseIf FGrid.TextMatrix(List2, 1) = "No" Then          '' Only for Non-MRP Parts
                NoUpto = 0
                mNo = 0
            ElseIf FGrid.TextMatrix(List2, 1) = "All" Then          '' For All Parts
                NoUpto = 1
                mNo = 0
            End If
            TRec1Qty = 0
            TRec2Qty = 0
            
            mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
            mIss_TB_Qty = 0: mIss_TB_Val = 0: mIss_TP_Qty = 0: mIss_TP_Val = 0
            mRec_TB_Qty = 0: mRec_TB_Val = 0: mRec_TP_Qty = 0: mRec_TP_Val = 0
            
            TRec1.Filter = ""
            mOPVRate = 0
            If TRec1.RecordCount > 0 Then    'Taxable Rect
                TRec1.MoveFirst
                TRec1.Filter = ("Part_No='" & RstPart!Part_No & "'")
                'Nra Update
                If TRec1.RecordCount > 0 Then
                    TRec1.MoveFirst
                    If TRec1!Rate <> 0 Then
                        mOPVRate = TRec1!Rate
                    Else
                        If TRec1!MRP_YN = 1 Then
                            mOPVRate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                        Else
                            mOPVRate = GCn.Execute("Select TB_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                            mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
                            mOPVRate = mOPVRate - ((mOPVRate * mDisPer) / 100)
                        End If
                    End If
                TRec1.MoveFirst
                End If

                'End Update
                
                If TRec1.EOF = False Then
                    TRec1Qty = TRec1!Qty
                End If
            End If
            mVRate = 0
            TRec2.Filter = ""
            If TRec2.RecordCount > 0 Then    'Taxpaid Rect
                TRec2.MoveFirst
                TRec2.Filter = ("Part_No='" & RstPart!Part_No & "'")
                'Nra Update
                If TRec2.RecordCount > 0 Then
                    TRec2.MoveLast
                    If TRec2!Rate <> 0 Then
                        mVRate = TRec2!Rate
                    Else
                        If TRec2!MRP_YN = 1 Then
                            mVRate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                        Else
                            mVRate = GCn.Execute("Select TP_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                            mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
                            mVRate = mVRate - ((mVRate * mDisPer) / 100)
                        End If
                    End If
                TRec2.MoveFirst
                End If
                
                'End Update
                If TRec2.EOF = False Then
                    TRec2Qty = TRec2!Qty
                End If
            End If
            If RstStock.RecordCount > 0 Then
                RstStock.MoveFirst
                RstStock.FIND ("Part_No='" & RstPart!Part_No & "'")
                If RstStock.EOF = False Then
                    Do While RstStock!Part_No = RstPart!Part_No    'Opening Calculation
                        If RstStock!StkTrn = "-" Then
                            If RstStock!Tax_YN = 1 Then     '' Taxable
                                mRate = 0
                                Call X_VAL11(TRec1, RstStock!Qty_Iss, mRate)
                            Else
                                mRate = 0
                                Call X_VAL22(TRec2, RstStock!Qty_Iss, mRate)
                            End If
                        ElseIf RstStock!StkTrn = "+" Then
                            If RstStock!Tax_YN = 1 Then     '' Taxable
                                mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                                mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * mOPVRate)
                            Else
                                mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                                mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * mVRate)
                            End If
                        End If
                        RstStock.MoveNext
                        If RstStock.EOF Then
                            Exit Do
                        ElseIf RstStock!Part_No <> RstPart!Part_No Then
                            Exit Do
                        End If
                    Loop
                End If
            End If
            xMOP_TBQty = mOP_TB_QTY:        xMOP_TPQty = mOP_TP_QTY
            xMOP_TBVal = mOP_TB_VAL:        xMOP_TPVal = mOP_TP_VAL
            '**
            mIss_TB_Qty = 0:                mIss_TB_Val = 0
            mIss_TP_Qty = 0:                mIss_TP_Val = 0
            '**
            mTrf = False
            
            If RstStock2.RecordCount > 0 Then
                RstStock2.MoveFirst
                RstStock2.FIND ("Part_No='" & RstPart!Part_No & "'")
                If RstStock2.EOF = False Then
                    Do While RstStock2!Part_No = RstPart!Part_No
                        mNarr = ""
                        If RstStock2!StkTrn = "-" Then
                            If RstStock2!Tax_YN = 1 Then     '' Taxable
                                mRate = 0
                                Call X_VAL11(TRec1, RstStock2!Qty_Iss, mRate, mNarr)
                            Else
                                mRate = 0
                                Call X_VAL22(TRec2, RstStock2!Qty_Iss, mRate, mNarr)
                            End If
                        ElseIf RstStock2!StkTrn = "+" Then
                            If RstStock2!Tax_YN = 1 Then     '' Taxable
                                mOP_TB_QTY = mOP_TB_QTY + RstStock2!Qty_Rec
                                mOP_TB_VAL = mOP_TB_VAL + (RstStock2!Qty_Rec * mOPVRate)
                            
                                mRec_TB_Qty = mRec_TB_Qty + RstStock2!Qty_Rec
                                mRec_TB_Val = mRec_TB_Val + (RstStock2!Qty_Rec * mOPVRate)
                            Else
                                mOP_TP_QTY = mOP_TP_QTY + RstStock2!Qty_Rec
                                mOP_TP_VAL = mOP_TP_VAL + (RstStock2!Qty_Rec * mVRate)
                                mRec_TP_Qty = mRec_TP_Qty + RstStock2!Qty_Rec
                                mRec_TP_Val = mRec_TP_Val + (RstStock2!Qty_Rec * mVRate)
                            End If
                        End If
                        RstStock2.MoveNext
                        If RstStock2.EOF Then
                            Exit Do
                        ElseIf RstStock2!Part_No <> RstPart!Part_No Then
                            Exit Do
                        End If
                    Loop
                End If
            End If
            If (xMOP_TBQty + mOP_TB_QTY) <> 0 Or (xMOP_TPQty + mOP_TP_QTY) <> 0 Then
                If mOP_TB_QTY = 0 Then
                    mOP_TB_VAL = 0
                ElseIf mOP_TB_QTY < 0 Then
'                    If mOP_TB_VAL  > 0 Then
'                        mOP_TB_VAL = -1 * mOP_TB_VAL
'                    Else
'                        mOP_TB_VAL = 0
'                    End If
                End If
                If mOP_TP_QTY = 0 Then
                    mOP_TP_VAL = 0
                ElseIf mOP_TP_QTY < 0 Then
'                    If mOP_TP_VAL  > 0 Then
'                        mOP_TP_VAL = -1 * mOP_TP_VAL
'                    Else
'                        mOP_TP_VAL = 0
'                    End If
                End If
                RsPart1.Filter = ("Code='" & RstPart!Part_No & "'")
                RsPart1.MoveFirst
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Part_Name") = RsPart1!Name
                    
                    .Fields("TB_OQty") = xMOP_TBQty
                    .Fields("TB_OVal") = xMOP_TBVal
                    .Fields("TP_OQty") = xMOP_TPQty
                    .Fields("TP_OVal") = xMOP_TPVal
                    
                    .Fields("RE_TB") = mRec_TB_Qty
                    .Fields("RE_TBV") = mRec_TB_Val
                    .Fields("RE_TP") = mRec_TP_Qty
                    .Fields("RE_TPV") = mRec_TP_Val
                    
                    .Fields("IS_TB") = mIss_TB_Qty
                    .Fields("IS_TBV") = mIss_TB_Val
                    .Fields("IS_TP") = mIss_TP_Qty
                    .Fields("IS_TPV") = mIss_TP_Val
                    
                    .Fields("TB_BQty") = mOP_TB_QTY
                    .Fields("TB_BVal") = mOP_TB_VAL
                    .Fields("TP_BQty") = mOP_TP_QTY
                    .Fields("TP_BVal") = mOP_TP_VAL
                    
                    .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
                    .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
                    
                    .Fields("Narr") = mNo
                    
                    .Update
                End With
            End If
            RstPart.MoveNext
        Loop
    End If
    Set RstRep = Temp06.Clone
    Set TRec1 = Nothing
    Set TRec2 = Nothing
    Set RstStock = Nothing
    Set RstStock2 = Nothing
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If MDIForm1.Picture1.Visible = True Then MDIForm1.Picture1.Visible = False
    ' For Speed Printing of report
    If SpeedPrnStkRep = True And FGrid.TextMatrix(List3, 1) = "Summary" Then
        SpeedPrintStkValFIFOSumm
        Exit Sub
    ElseIf SpeedPrnStkRep = True And FGrid.TextMatrix(List3, 1) = "Detail" Then
        SpeedPrintStkValFIFODet
        Exit Sub
    End If
    RepTitle = UCase(Me.CAPTION)
ELoop:
    Set TRec1 = Nothing
    Set TRec2 = Nothing
    Set RstStock = Nothing
    Set RstStock2 = Nothing
    Set GRs = Nothing
    If err.NUMBER <> 0 Then CheckError
End Sub
Private Sub SpeedPrintStkValFIFOSumm()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double
    Dim isLast As Boolean, mRec As Integer, PageNo As Double
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
    
    Print #1, PRN_TIT("Stock Valuation FIFO [" & FGrid.TextMatrix(List3, 1) & "]", "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, "From : " & FGrid.TextMatrix(Date1, 1) & "  To : " & FGrid.TextMatrix(Date2, 1)
    mHeader = mHeader + 1
    Print #1, "For MRP Parts : " & FGrid.TextMatrix(List2, 1)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & Space(10) & Space(22) & Space(28) & "<------------TAXABLE------------ >" & "<------------TAXPAID------------ >" & Space(10)
    mHeader = mHeader + 1
    Print #1, PSTR("#", 10) & PSTR("Part No.", 22) & PSTR("Part Name", 28) & PSTR("OP Qty", 8, , AlignRight) & PSTR("Rec Qty", 8, , AlignRight) & PSTR("Iss Qty", 8, , AlignRight) & PSTR("Bal Qty", 8, , AlignRight) & PSTR("OP Qty", 8, , AlignRight) & PSTR("Rec Qty", 8, , AlignRight) & PSTR("Iss Qty", 8, , AlignRight) & PSTR("Bal Qty", 8, , AlignRight) & PSTR("Clos.Qty", 10, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    RstRep.MoveFirst
    mHeader = 1
    While Not RstRep.EOF = True
        If mHeader <= mRec Then
            Counter = Counter + 1
            Print #1, mChr17 & PSTR(STR(Counter), 10) & PSTR(RstRep!Part_No, 22) & PSTR(RstRep!Part_Name, 28) & PSTR(IIf(RstRep!TB_OQty = 0, "", STR(Format(RstRep!TB_OQty, ".00"))), 8) & PSTR(IIf(RstRep!RE_TB = 0, "", STR(Format(RstRep!RE_TB, ".00"))), 8) & PSTR(IIf(RstRep!IS_TB = 0, "", STR(Format(RstRep!IS_TB, ".00"))), 8) & PSTR(IIf(RstRep!TB_BQty = 0, "", STR(Format(RstRep!TB_BQty, ".00"))), 8) & PSTR(IIf(RstRep!TP_OQty = 0, "", STR(Format(RstRep!TP_OQty, ".00"))), 8) & PSTR(IIf(RstRep!RE_TP = 0, "", STR(Format(RstRep!RE_TP, ".00"))), 8) & PSTR(IIf(RstRep!IS_TP = 0, "", STR(Format(RstRep!IS_TP, ".00"))), 8) & PSTR(IIf(RstRep!TP_BQty = 0, "", STR(Format(RstRep!TP_BQty, ".00"))), 8) & PSTR(IIf(RstRep!Net_Qty = 0, "", STR(Format(RstRep!Net_Qty, ".00"))), 10) & mChr18
            mHeader = mHeader + 1
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
                Print #1, PRN_TIT("Stock Valuation FIFO [" & FGrid.TextMatrix(List3, 1) & "]", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & (10) & Space(22) & Space(28) & "<------------TAXABLE------------ >" & "<------------TAXPAID------------ >" & Space(10)
                mHeader = mHeader + 1
                Print #1, PSTR("#", 10) & PSTR("Part No.", 22) & PSTR("Part Name", 28) & PSTR("OP Qty", 8, , AlignRight) & PSTR("Rec Qty", 8, , AlignRight) & PSTR("Iss Qty", 8, , AlignRight) & PSTR("Bal Qty", 8, , AlignRight) & PSTR("OP Qty", 8, , AlignRight) & PSTR("Rec Qty", 8, , AlignRight) & PSTR("Iss Qty", 8, , AlignRight) & PSTR("Bal Qty", 8, , AlignRight) & PSTR("Clos.Qty", 10, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
            End If
        End If
    RstRep.MoveNext
    Wend
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
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
Private Sub SpeedPrintStkValFIFODet()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double
    Dim isLast As Boolean, mRec As Integer, PageNo As Double
    Dim TotalTBVal As Double, TotalTPVal As Double, TotalVal As Double
    Dim TotalTBQty As Double, TotalTPQty As Double, TotalQty As Double
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
    
    Print #1, PRN_TIT("Stock Valuation FIFO [" & FGrid.TextMatrix(List3, 1) & "]", "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, "From : " & FGrid.TextMatrix(Date1, 1) & "  To : " & FGrid.TextMatrix(Date2, 1)
    mHeader = mHeader + 1
    Print #1, "For MRP Parts : " & FGrid.TextMatrix(List2, 1)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("#", 10) & PSTR("Part No.", 22) & PSTR("Part Name", 28) & PSTR("TB Qty", 10, , AlignRight) & PSTR("TP Qty", 10, , AlignRight) & PSTR("TotalQty", 10, , AlignRight) & PSTR("TB Val", 12, , AlignRight) & PSTR("TP Val", 12, , AlignRight) & PSTR("Total Val", 16, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    RstRep.MoveFirst
    mHeader = 1
    While Not RstRep.EOF = True
        If mHeader <= mRec Then
            Counter = Counter + 1
            Print #1, mChr17 & PSTR(STR(Counter), 10) & PSTR(RstRep!Part_No, 22) & PSTR(RstRep!Part_Name, 28) & PSTR(IIf(RstRep!TB_BQty = 0, "", STR(Format(RstRep!TB_BQty, ".00"))), 10, , AlignRight) & PSTR(IIf(RstRep!TP_BQty = 0, "", STR(Format(RstRep!TP_BQty, ".00"))), 10, , AlignRight) & PSTR(IIf(RstRep!TB_BQty + RstRep!TP_BQty = 0, "", STR(Format(RstRep!TB_BQty + RstRep!TP_BQty, ".00"))), 10, , AlignRight) & PSTR(IIf(RstRep!TB_BVal = 0, "", STR(Format(RstRep!TB_BVal, ".00"))), 12, , AlignRight) & PSTR(IIf(RstRep!TP_BVal = 0, "", STR(Format(RstRep!TP_BVal, ".00"))), 12, , AlignRight) & PSTR(IIf(RstRep!TB_BVal + RstRep!TP_BVal = 0, "", STR(Format(RstRep!TB_BVal + RstRep!TP_BVal, ".00"))), 16, , AlignRight) & mChr18
            
            TotalTBQty = TotalTBQty + Val(RstRep!TB_BQty): TotalTPQty = TotalTPQty + Val(RstRep!TP_BQty)
            TotalQty = TotalQty + (Val(RstRep!TB_BQty) + Val(RstRep!TP_BQty))
            
            TotalTBVal = TotalTBVal + Val(RstRep!TB_BVal): TotalTPVal = TotalTPVal + Val(RstRep!TP_BVal)
            TotalVal = TotalVal + (Val(RstRep!TB_BVal) + Val(RstRep!TP_BVal))
            
            mHeader = mHeader + 1
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
                Print #1, PRN_TIT("Stock Valuation FIFO [" & FGrid.TextMatrix(List3, 1) & "]", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & PSTR("#", 10) & PSTR("Part No.", 22) & PSTR("Part Name", 28) & PSTR("TB Qty", 10, , AlignRight) & PSTR("TP Qty", 10, , AlignRight) & PSTR("TotalQty", 10, , AlignRight) & PSTR("TB Val", 12, , AlignRight) & PSTR("TP Val", 12, , AlignRight) & PSTR("Total Val", 16, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
            End If
        End If
    RstRep.MoveNext
    Wend
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("Total-- >", 10) & Space(22) & Space(28) & PSTR(IIf(TotalTBQty = 0, "", STR(Format(TotalTBQty, ".00"))), 10, , AlignRight) & PSTR(IIf(TotalTPQty = 0, "", STR(Format(TotalTPQty, ".00"))), 10, , AlignRight) & PSTR(IIf(TotalQty = 0, "", STR(Format(TotalQty, ".00"))), 10, , AlignRight) & PSTR(IIf(TotalTBVal = 0, "", STR(Format(TotalTBVal, ".00"))), 12, , AlignRight) & PSTR(IIf(TotalTPVal = 0, "", STR(Format(TotalTPVal, ".00"))), 12, , AlignRight) & PSTR(IIf(TotalVal = 0, "", STR(Format(TotalVal, ".00"))), 16, , AlignRight) & mChr18
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

Private Sub X_VAL11(ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
On Error GoTo ErrLoop
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        If mOP_TB_VAL <> 0 And mOP_TB_QTY <> 0 Then
            xRate = Round(mOP_TB_VAL / mOP_TB_QTY, 3)
        Else
            xRate = 0
        End If
            mOP_TB_QTY = mOP_TB_QTY - xQty
            mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
            mIss_TB_Qty = mIss_TB_Qty + xQty
            mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
          Exit Sub
    End If
    If xQty = TRec1Qty Then
        TRec1Qty = 0
        xRate = VNull(TRec1!Rate)
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
        TRec1.MoveNext
        If TRec1.EOF = False Then
            TRec1Qty = TRec1!Qty
        End If
'    ElseIf xQty < TRec1!Qty Then
    ElseIf xQty < TRec1Qty Then
        TRec1Qty = TRec1Qty - xQty
        xRate = VNull(TRec1!Rate)
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
'    ElseIf xQty  > TRec1!Qty Then
    ElseIf xQty > TRec1Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
'            If TRec1!Qty <= TQty Then
            If TRec1Qty <= TQty Then
'                TQty = TQty - TRec1!Qty
                TQty = TQty - TRec1Qty
                xRate = VNull(TRec1!Rate)
                mOP_TB_QTY = mOP_TB_QTY - TRec1Qty 'TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + (TRec1Qty) '(TRec1!Qty)
                mIss_TB_Val = mIss_TB_Val + (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                TRec1Qty = 0
'                TRec1!Qty = 0
'                TRec1.Update
            Else
                TRec1Qty = TRec1Qty - TQty
'                TRec1!Qty = TRec1!Qty - TQty
'                TRec1.Update
                xRate = VNull(TRec1!Rate)
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
            End If
            If TRec1.EOF = False Then
                TRec1Qty = TRec1!Qty
            End If
        Loop
    End If
Exit Sub
ErrLoop:
     If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub X_VAL22(ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        'xRate = 0
        
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        Exit Sub
    End If
'    If xQty = TRec2!Qty Then
    If xQty = TRec2Qty Then
        TRec2Qty = 0
'        TRec2!Qty = 0
'        TRec2.Update
        xRate = TRec2!Rate
      
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        TRec2.MoveNext
        If TRec2.EOF = False Then
            TRec2Qty = TRec2!Qty
        End If
'    ElseIf xQty < TRec2!Qty Then
    ElseIf xQty < TRec2Qty Then
        TRec2Qty = TRec2Qty - xQty
'        TRec2!Qty = TRec2!Qty - xQty
'        TRec2.Update
        xRate = TRec2!Rate
       
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
'    ElseIf xQty  > TRec2!Qty Then
    ElseIf xQty > TRec2Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
'            If TRec2!Qty <= TQty Then
            If TRec2Qty <= TQty Then
                TQty = TQty - TRec2Qty 'TRec2!Qty
                xRate = TRec2!Rate
              
                mOP_TP_QTY = mOP_TP_QTY - TRec2Qty 'TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2Qty * xRate)   '(TRec2!Qty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + (TRec2Qty)     '(TRec2!Qty)
                mIss_TP_Val = mIss_TP_Val + (TRec2Qty * xRate) '(TRec2!Qty * xRate)
                TRec2Qty = 0
'                TRec2!Qty = 0
'                TRec2.Update
            Else
                TRec2Qty = TRec2Qty - TQty
'                TRec2!Qty = TRec2!Qty - TQty
'                TRec2.Update
                xRate = TRec2!Rate
             
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
            End If
            If TRec2.EOF = False Then
                TRec2Qty = TRec2!Qty
            End If
        Loop
    End If
End Sub

Private Sub SprPartProfitCalc()
On Error GoTo ELoop
Dim mQry$, Condstr$, CondStr1$, Condstr2$
Dim mSale As Double, mCost As Double, mAmount As Double, mQty As Double, mProf As Double
Dim XRecNo As Double, xRate As Double, xRate1 As Double
Dim D_Per_MRP_TB1 As Double, D_Per_MRP_TP1 As Double, Gen_Sur_Per1 As Double, D_Per_TP1 As Double, D_Per_TB1 As Double
Dim mNo As Byte, NoUpto As Byte, TRec1Qty As Double

RepPrint = True
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    ''**
    Condstr = "where SPStk.V_Date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and SPStk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    If Check1(1).Value = Unchecked Then CondStr1 = " and left(SPStk.DocID,1) in (" & GridString1 & ")"
    
    Condstr = Condstr & CondStr1
    If FGrid.TextMatrix(List1, 1) = "Yes" Then          '' Only for Marked Parts
        Condstr2 = " and Mark_yn='Y'"
    Else
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and SPStk.Part_No in (" & GridString3 & ")"
    End If
    
    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
    
    'For RstPart, SQL
    GSQL = "Select Distinct Part.Part_No,Part.Part_Name From Part " & _
        "where Part.Part_No in " & _
        "(select Distinct SPStk.Part_No from Sp_Stock as SPStk " & Condstr & ") "
    If Check1(1).Value = Unchecked Then
        GSQL = GSQL & " and Part.Div_Code in (" & GridString1 & ") " & Condstr2
    Else
        GSQL = GSQL & Condstr2
    End If
    GSQL = GSQL & " Order By Part.Part_No"
    
    Set RstPart = GCn.Execute(GSQL)
    
    '********** Taxable+Taxpaid Qty
    mQry = "select SPStk.Part_No,SPStk.V_DATE,SPStk.Qty_Rec as Qty, " & cIIF("SPStk.V_Rate=0", "SPStk.Rate", "SPStk.V_Rate") & " as Rate " & _
        "From " & _
        "Sp_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
        Condstr & " and Part_No in (select Distinct Stk.Part_No from Sp_Stock as Stk where Stk.V_Date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and  Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    If Check1(1).Value = Unchecked Then
        mQry = mQry & " and left(Stk.Docid,1) in (" & GridString1 & ")) "
    Else
        mQry = mQry & ")"
    End If
    GSQL = mQry & " and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date"
    Set TRec1 = New Recordset
    With TRec1
        .CursorLocation = adUseClient
        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
    End With
    mQry = "select SPStk.DocID,SPStk.V_Type,SPStk.Part_No,SPStk.V_DATE,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate,Vt.StkTrn,SPStk.Invoice_DocID,SPStk.Amount as Net_Amt " & _
        "From " & _
        "Sp_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
        "where SPStk.V_Date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and SPStk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & _
        CondStr1 & " AND Vt.StkTrn='-' and SPStk.Part_No in " & _
        "(select Distinct Stk.Part_No from Sp_Stock as Stk where Stk.V_Date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    If Check1(1).Value = Unchecked Then
        mQry = mQry & " and left(Stk.Docid,1) in (" & GridString1 & ")) "
    Else
        mQry = mQry & ")"
    End If
    GSQL = mQry & " and SPSTK.Qty_Iss > 0 Order By SPStk.Part_No,SPStk.V_Date, " & cMID("SPStk.DocID", "4", "5") & ""
    Set RstStock = GCn.Execute(GSQL)
    Do While Not RstPart.EOF
        MDIForm1.Picture1.Visible = True
        MDIForm1.Label1.CAPTION = "Process Status : " & RstPart.AbsolutePosition & "/" & RstPart.RecordCount
        MDIForm1.Label1.Refresh
     
        mQty = 0: mCost = 0: mSale = 0
        TRec1.Sort = "V_Date"
        RstStock.Sort = "V_Date"
        TRec1.Filter = ""
        RstStock.Filter = ""
        If TRec1.RecordCount > 0 Then
            TRec1.MoveFirst
            TRec1.Filter = ("Part_No='" & RstPart!Part_No & "'")
            If TRec1.EOF = False Then
                TRec1Qty = TRec1!Qty
            End If
        End If
'        If UCase(left(PubComp_Name, 3)) <> "JMK" Then
'            If TRec1.EOF = False Then
'
'                RstStock.Filter = ("Part_No='" & RstPart!Part_No & "'")
'                If RstStock.EOF = False Then
'                RstStock.MoveFirst
'                    Do While RstStock!Part_No = RstPart!Part_No
'                        Set GRs = GCn.Execute("select Gen_Sur_Per,D_Per_TB,D_Per_Tp,D_Per_MRP_TB,D_Per_MRP_Tp From Sp_Sale where DocID='" & IIf(IsNull(RstStock!Invoice_DocID) Or RstStock!Invoice_DocID = "", RstStock!DocId, RstStock!Invoice_DocID) & "'")
'                        If GRs.RecordCount > 0 Then
'                            D_Per_MRP_TB1 = GRs!D_PER_MRP_TB
'                            D_Per_MRP_TP1 = GRs!D_PER_MRP_TP
'                            Gen_Sur_Per1 = GRs!Gen_Sur_Per
'                            D_Per_TP1 = GRs!D_Per_TP
'                            D_Per_TB1 = GRs!D_Per_TB
'                        Else
'                            D_Per_MRP_TB1 = 0
'                            D_Per_MRP_TP1 = 0
'                            Gen_Sur_Per1 = 0
'                            D_Per_TP1 = 0
'                            D_Per_TB1 = 0
'                        End If
'                        If RstStock!Qty_Iss = TRec1Qty Then 'TRec1!Qty Then
'    '                       TRec1.Fields("QTY") = 0
'    '                       TRec1.Update
'                            TRec1Qty = 0
'                            mQty = mQty + RstStock!Qty_Iss
'                            mAmount = RstStock!Net_Amt
'                            If RstStock!Tax_YN = 1 And RstStock!MRP_YN = 1 Then  '' Taxable & MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_MRP_TB1 / 100, 2)
'                                'mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                            ElseIf RstStock!Tax_YN = 1 And RstStock!MRP_YN = 0 Then  '' Taxable & Non-MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_TB1 / 100, 2)
'                                mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                            ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 1 Then  '' Taxpaid & MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_MRP_TP1 / 100, 2)
'                            ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 0 Then  '' TaxPaid & Non-MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_TP1 / 100, 2)
'                            End If
'                            mSale = mSale + mAmount
'                            mCost = mCost + Round(RstStock!Qty_Iss * TRec1!Rate, 2)
'                            TRec1.MoveNext
'                        ElseIf RstStock!Qty_Iss < TRec1Qty Then    'TRec1!Qty Then
'    '                        TRec1.Fields("QTY") = TRec1!Qty - RstStock!Qty_Iss
'    '                        TRec1.Update
'                            TRec1Qty = TRec1Qty - RstStock!Qty_Iss
'                            mQty = mQty + RstStock!Qty_Iss
'                            mAmount = RstStock!Net_Amt
'                            If RstStock!Tax_YN = 1 And RstStock!MRP_YN = 1 Then  '' Taxable & MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_MRP_TB1 / 100, 2)
'                               'mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                            ElseIf RstStock!Tax_YN = 1 And RstStock!MRP_YN = 0 Then  '' Taxable & Non-MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_TB1 / 100, 2)
'                                mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                            ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 1 Then  '' Taxpaid & MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_MRP_TP1 / 100, 2)
'                            ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 0 Then  '' TaxPaid & Non-MRP
'                                mAmount = mAmount - Round(mAmount * D_Per_TP1 / 100, 2)
'                            End If
'                            mSale = mSale + mAmount
'                            mCost = mCost + Round(RstStock!Qty_Iss * TRec1!Rate, 2)
'                        ElseIf RstStock!Qty_Iss > TRec1Qty Then     'TRec1!Qty Then
'                            TQty = RstStock!Qty_Iss
'    '                        Do While TQty <> 0 And TRec1.EOF()
'    '                            If TRec1!Qty <= TQty Then
'                            If TRec1.EOF = False Then
'                                Do While TRec1!Part_No = RstStock!Part_No
'                                    If TRec1Qty >= TQty Then
'                                        TQty = TQty - TRec1Qty 'TRec1!Qty
'                                        mQty = mQty + TRec1Qty 'TRec1!Qty
'
'                                        '' Cost Rate
'                                        xRate = TRec1!Rate
'                                        '' Calculate the Rate of Item for issued qty
'                                        xRate1 = Round(RstStock!Net_Amt / RstStock!Qty_Iss, 2)
'        '                                mAmount = Round(TRec1!Qty * xRate1, 2)
'                                        mAmount = Round(TRec1Qty * xRate1, 2)
'                                        If RstStock!Tax_YN = 1 And RstStock!MRP_YN = 1 Then  '' Taxable & MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_MRP_TB1 / 100, 2)
'                                            'mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                                        ElseIf RstStock!Tax_YN = 1 And RstStock!MRP_YN = 0 Then  '' Taxable & Non-MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_TB1 / 100, 2)
'                                            mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                                        ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 1 Then  '' Taxpaid & MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_MRP_TP1 / 100, 2)
'                                        ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 0 Then  '' TaxPaid & Non-MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_TP1 / 100, 2)
'                                        End If
'                                        mSale = mSale + mAmount
'        '                                mCost = mCost + Round(TRec1!Qty * xRate, 2)
'                                        mCost = mCost + Round(TRec1Qty * xRate, 2)
'        '                                TRec1.Fields("QTY") = 0
'        '                                TRec1.Update
'                                    Else
'        '                                TRec1.Fields("QTY") = TRec1!Qty - TQty
'        '                                TRec1.Update
'                                        If TRec1Qty <= 0 Then
'                                            TRec1Qty = TRec1!Qty
'                                        End If
'                                        TRec1Qty = TRec1Qty - TQty
'                                        mQty = mQty + TQty
'
'                                        '' Calculate the Rate of Item for issued qty
'                                        If RstStock!Qty_Iss > 0 Then
'                                            xRate1 = Round(RstStock!Net_Amt / RstStock!Qty_Iss, 2)
'                                        End If
'                                        mAmount = TQty * xRate1
'                                        If RstStock!Tax_YN = 1 And RstStock!MRP_YN = 1 Then  '' Taxable & MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_MRP_TB1 / 100, 2)
'                                            'mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                                        ElseIf RstStock!Tax_YN = 1 And RstStock!MRP_YN = 0 Then  '' Taxable & Non-MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_TB1 / 100, 2)
'                                            mAmount = mAmount + Round(mAmount * Gen_Sur_Per1 / 100, 2)
'                                        ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 1 Then  '' Taxpaid & MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_MRP_TP1 / 100, 2)
'                                        ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 0 Then  '' TaxPaid & Non-MRP
'                                            mAmount = mAmount - Round(mAmount * D_Per_TP1 / 100, 2)
'                                        End If
'                                        mSale = mSale + mAmount
'                                        mCost = mCost + Round(TQty * TRec1!Rate, 2)
'                                        TQty = 0
'                                        Exit Do
'                                    End If
'                                    TRec1.MoveNext
'                                    If TQty = 0 Or TRec1.EOF Then
'                                        Exit Do
'                                    End If
'                                Loop
'                            End If
'                        End If
'                        RstStock.MoveNext
'                        If RstStock.EOF Then
'                            Exit Do
'                        ElseIf RstStock!Part_No <> RstPart!Part_No Then
'                            Exit Do
'                        End If
'                    Loop
'                End If
'            Else
'            End If
'        Else
'           Set RstStock = GCn.Execute("select Sum(SPStk.Qty_Iss-SPStk.Qty_ret) as Qty1,Sum(Net_Amt) as Amt From Sp_Stock as SPStk left Join [" & PubSFADataPath & "].Voucher_Type vt on Vt.V_type=SPStk.V_type where Part_No='" & RstPart!Part_No & "' and vt.StkTrn='-' and v_date<" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondStr1)
            If UCase(left(PubComp_Name, 3)) <> "JMK" Then
                Set GRs = GCn.Execute("select Sum(SPStk.Qty_Iss-SPStk.Qty_ret) as Qty1,Sum(Net_Amt) as Amt,V_Date " & _
                       "From Sp_Stock as SPStk left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=SPStk.V_type " & _
                       "where Part_No='" & RstPart!Part_No & _
                       "' and vt.StkTrn='-' " & _
                       " and V_Date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondStr1 & " group by V_Date")
            Else
                Set GRs = GCn.Execute("select Sum(SPStk.Qty_Iss-SPStk.Qty_ret) as Qty1,Sum(Amount) as Amt,V_Date " & _
                       "From Sp_Stock as SPStk left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=SPStk.V_type " & _
                       "where Part_No='" & RstPart!Part_No & _
                       "' and vt.StkTrn='-' " & _
                       " and V_Date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondStr1 & " group by V_Date")
            End If
           ' Set GRs = GCn.Execute("select Sum(SPStk.Qty_Iss-SPStk.Qty_ret) as Qty1,Sum(Net_Amt) as Amt " & _
                    "From Sp_Stock as SPStk " & _
                    "where Part_No='" & RstPart!Part_No & _
                    "' and v_date<" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondStr1)
                    
                'mQty = IIf(IsNull(GRs!Qty1), 0, GRs!Qty1)
                
                'mSale = IIf(IsNull(GRs!Amt), 0, GRs!Amt)
                'Set GRs = Nothing
        'End If
        Do While GRs.EOF = False
            mQty = IIf(IsNull(GRs!Qty1), 0, GRs!Qty1)
            mSale = IIf(IsNull(GRs!Amt), 0, GRs!Amt)
            mCost = GCn.Execute("Select " & vIsNull("V_Rate", "0") & " from SP_Stock where Part_No='" & RstPart!Part_No & "'").Fields(0).Value
            mCost = mCost * mQty
            If mQty + mCost + mSale > 0 Then
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("TB_OQty") = mQty
                    .Fields("TB_OVal") = mCost
                    .Fields("TP_OQty") = mSale
                    .Fields("TP_OVal") = mSale - mCost
                    .Fields("Inv_date") = GRs!V_DATE
                    .Update
                End With
            End If
            GRs.MoveNext
        Loop
        RstPart.MoveNext
    Loop
    Set RstRep = Temp06.Clone
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List2, 1) = "Detailed" Then
        RepName = "SparePartProfit"
    Else
        RepName = "SparePartProfitSum"
    End If
    RepTitle = UCase(Me.CAPTION)
ELoop:
    If err.NUMBER <> 0 Then CheckError
    Set GRs = Nothing
End Sub

Private Sub SprProjectionCalc()
On Error GoTo ELoop
Dim mQry$, CondStr1$, Condstr$, TrnSQL$, TrnSQL2$
Dim mClQty As Double, mClAmt As Double, mBAL As Double, mDate1 As Date, mReqQty As Double
Dim mSaleQty As Double, mOtherQty As Double

    RepPrint = True
    
    'If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat3, FGrid.TextMatrix(Cat3, 0)) = False Then RepPrint = False: Exit Sub


    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

    Condstr = "where V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1))

    If FGrid.TextMatrix(List1, 1) = "Yes" Then          '' Only for Marked Parts
        Condstr = Condstr & " and Mark_yn='Y'"
    Else
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Part.Part_No in (" & GridString3 & ")"
    End If

    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
    
    Set RstPart = GCn.Execute("Select Distinct Sp_Stock.Part_No,Part.Part_Name,Part.MRP,Part.TB_SRate From Sp_Stock Left Join Part On SP_Stock.Part_No=Part.Part_No and Part.Div_Code = left(SP_Stock.Docid,1) " & Condstr)
    Do While Not RstPart.EOF
        mClQty = 0: mClAmt = 0: mBAL = 0: mReqQty = 0: mSaleQty = 0: mOtherQty = 0
        
        Set RstStock = GCn.Execute("select " & vIsNull("Sum(Qty_rec)", "0") & " as Rect, " & vIsNull("Sum(Qty_Iss)", "0") & " as Issue, " & vIsNull("Sum(Qty_Ret)", "0") & " as Ret,MRP_YN,Tax_YN From Sp_Stock where Part_No='" & RstPart!Part_No & "' and v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " Group By Mrp_YN,Tax_YN")
        If RstStock.RecordCount > 0 Then
            Do While Not RstStock.EOF
                mBAL = (RstStock!Rect + RstStock!Ret - RstStock!Issue)
                If mBAL <> 0 Then
                    mClQty = mClQty + mBAL
                    mClAmt = mClAmt + GetFIFOAmt(RstPart!Part_No, FGrid.TextMatrix(Date2, 1), mBAL, RstStock!MRP_YN, RstStock!Tax_YN)
                End If
                RstStock.MoveNext
            Loop
        End If
        mDate1 = CDate(FGrid.TextMatrix(Date2, 1)) - Val(FGrid.TextMatrix(Cat1, 1))
        
        Set RstStock = GCn.Execute("select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and v_DATE >= " & ConvertDate(mDate1) & " AND  v_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " And StkTrn='-' order By Sp_Stock.V_Date,Vt.StkTrn")
        Do While Not RstStock.EOF
            Select Case RstStock!V_Type
                Case SprSlChal, SprSlCsh, SprSlCre, WksGenReq, WksReqWrt    'WksSlCsh,WksSlCre
                    mSaleQty = mSaleQty + (RstStock!Qty_Iss - RstStock!Qty_Ret)
                Case Else       ''SprTrfChal
                    mOtherQty = mOtherQty + (RstStock!Qty_Iss - RstStock!Qty_Ret)
                    '' SprSlRetCsh,SprSlRetCre,SprSlTrfRet
            End Select
            RstStock.MoveNext
        Loop
        
        mReqQty = Round(Val(FGrid.TextMatrix(Cat2, 1)) * ((mSaleQty + mOtherQty) / Val(FGrid.TextMatrix(Cat1, 1))), 2)
        
        If Round(mReqQty, 2) <> Round(mClQty, 2) Then
            If Val(FGrid.TextMatrix(Cat2, 1)) > 0 Then
                If Abs(mClQty - mReqQty) * IIf(RstPart!MRP > RstPart!TB_SRate, RstPart!MRP, RstPart!TB_SRate) < Val(FGrid.TextMatrix(Cat2, 1)) Then
                    GoTo MyNextPart
                End If
            End If
            
            If FGrid.TextMatrix(List2, 1) <> "All" Then
                If FGrid.TextMatrix(List2, 1) = "Short" And mReqQty < mClQty Then
                    GoTo MyNextPart
                ElseIf FGrid.TextMatrix(List2, 1) = "Excess" And mReqQty > mClQty Then
                    GoTo MyNextPart
                End If
            End If
            
            With Temp06
                .AddNew
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Part_Name") = RstPart!Part_Name
                .Fields("Inv_No") = Val(FGrid.TextMatrix(Cat1, 1))   '' Last Days
                .Fields("V_no") = Val(FGrid.TextMatrix(Cat2, 1))     '' Next days
                .Fields("TB_OQty") = mClQty
                .Fields("TB_OVal") = mClAmt
                .Fields("TP_OQty") = mSaleQty
                .Fields("TP_OVal") = mOtherQty
                .Fields("Re_Tb") = mReqQty
                If mReqQty > mClQty Then
                    .Fields("Re_TBv") = mReqQty - mClQty
                    .Fields("Narr") = "Short"
                    .Fields("Is_Tb") = (-1) * Abs(mClQty - mReqQty) * IIf(RstPart!MRP > RstPart!TB_SRate, RstPart!MRP, RstPart!TB_SRate)
                Else
                    .Fields("Re_TPv") = mClQty - mReqQty
                    .Fields("Narr") = "Excess"
                    .Fields("Is_Tb") = Abs(mClQty - mReqQty) * IIf(RstPart!MRP > RstPart!TB_SRate, RstPart!MRP, RstPart!TB_SRate)
                End If
                
                .Update
            End With
        End If
MyNextPart:
        RstPart.MoveNext
    Loop
    Set RstRep = Temp06.Clone
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprProjection"
    RepTitle = UCase(Me.CAPTION)
ELoop:
    If err.NUMBER <> 0 Then CheckError
    Set GRs = Nothing
End Sub


Private Sub SprSaleInventoryCalc()
On Error GoTo ELoop
Dim mQry$, CondStr1$, Condstr$, TrnSQL$, TrnSQL2$
Dim xDtTO As Date, xDtFR As Date
Dim mCMonth$, mMth_No As Byte
Dim mJOB_SAL As Double, mJOB_FREE As Double, mCOU_SAL As Double, mCOU_TRF As Double, mCL_VAL As Double
Dim D_Per_MRP_TB1 As Double, D_Per_MRP_TP1 As Double, Gen_Sur_Per1 As Double, D_Per_TP1 As Double, D_Per_TB1 As Double
Dim mAmount As Double, mInvValue As Double
Dim mNo As Byte, NoUpto As Byte

    RepPrint = True
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    Condstr = "where V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1))

    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
        
    Set TRec1 = New ADODB.Recordset
    Set TRec1 = TmpTRec1(TRec1)
    
    
    Set TRec2 = New ADODB.Recordset
    Set TRec2 = TmpTRec1(TRec2)
    
    xDtTO = CDate(FGrid.TextMatrix(Date1, 1))
    
    Do While xDtTO <= CDate(FGrid.TextMatrix(Date2, 1))
         MDIForm1.Picture1.Visible = True
         MDIForm1.Label1.CAPTION = "Process For Date : " & xDtTO
         MDIForm1.Label1.Refresh
         
        xDtFR = FLDAY(xDtTO, "F")
        xDtTO = IIf(CDate(FGrid.TextMatrix(Date2, 1)) >= FLDAY(xDtTO, "L"), FLDAY(xDtTO, "L"), CDate(FGrid.TextMatrix(Date2, 1)))
        
        mCMonth = Format(xDtTO, "MMMM")
        mMth_No = IIf(Format(xDtTO, "MM") >= 4, Format(xDtTO, "MM") - 3, Format(xDtTO, "MM") + 9)
        mJOB_SAL = 0: mJOB_FREE = 0: mCOU_SAL = 0: mCOU_TRF = 0: mCL_VAL = 0
        
        '' Sales Calcluation
        Set RstStock = GCn.Execute("select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where  (v_date >=" & ConvertDate(xDtFR) & " and v_date<=" & ConvertDate(xDtTO) & ") and Vt.StkTrn ='-' and len(SP_Stock.Invoice_Docid) > 0  order By Sp_Stock.V_Date,Vt.StkTrn")
        Do While Not RstStock.EOF
            Select Case RstStock!V_Type
                Case SprSlChal, SprTrfChal, SprSlRetCsh, SprSlRetCre, SprSlTrfRet '' Counter Sales Type
                    Set GRs = GCn.Execute("select Gen_Sur_Per,D_Per_TB,D_Per_Tp,D_Per_MRP_TB,D_Per_MRP_Tp From Sp_Sale where DocID='" & IIf(IsNull(RstStock!Invoice_DocID) Or RstStock!Invoice_DocID = "", RstStock!DocID, RstStock!Invoice_DocID) & "'")
                    If GRs.RecordCount > 0 Then
                        D_Per_MRP_TB1 = GRs!D_PER_MRP_TB
                        D_Per_MRP_TP1 = GRs!D_PER_MRP_TP
                        Gen_Sur_Per1 = GRs!Gen_Sur_Per
                        D_Per_TP1 = GRs!D_Per_TP
                        D_Per_TB1 = GRs!D_Per_TB
                    Else
                        D_Per_MRP_TB1 = 0
                        D_Per_MRP_TP1 = 0
                        Gen_Sur_Per1 = 0
                        D_Per_TP1 = 0
                        D_Per_TB1 = 0
                    End If
                    mAmount = RstStock!Net_Amt
                    If RstStock!Tax_YN = 1 And RstStock!MRP_YN = 1 Then  '' Taxable & MRP
                        mAmount = mAmount - Round(mAmount * D_Per_MRP_TB1 / 100, 2)
                    ElseIf RstStock!Tax_YN = 1 And RstStock!MRP_YN = 0 Then  '' Taxable & Non-MRP
                        mAmount = mAmount - Round(mAmount * D_Per_TB1 / 100, 2)
                    ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 1 Then  '' Taxpaid & MRP
                        mAmount = mAmount - Round(mAmount * D_Per_MRP_TP1 / 100, 2)
                    ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 0 Then  '' TaxPaid & Non-MRP
                        mAmount = mAmount - Round(mAmount * D_Per_TP1 / 100, 2)
                    End If
                    If RstStock!StkTrn = "-" Then       '' Sales
                        If RstStock!V_Type = SprSlChal Then
                            mCOU_SAL = mCOU_SAL + mAmount
                        Else ''If RstStock!v_TYPE = SprTrfChal Then  '' Transfer
                            mCOU_TRF = mCOU_TRF + mAmount
                        End If
                    ElseIf RstStock!StkTrn = "+" Then       '' Sales return
                        If RstStock!V_Type = SprSlRetCsh Or RstStock!V_Type = SprSlRetCre Then
                            mCOU_SAL = mCOU_SAL - mAmount
                        Else    '' if RstStock!v_TYPE = SprSlTrfRet '' Transfer Return
                            mCOU_TRF = mCOU_TRF - mAmount
                        End If
                    End If
                
                Case WksGenReq, WksReqWrt  '' WorkShop Sales Type
                    Set GRs = GCn.Execute("select Gen_Sur_Per,D_Per_TB,D_Per_Tp,D_Per_MRP_TB,D_Per_MRP_Tp From Sp_Sale where DocID='" & IIf(IsNull(RstStock!Invoice_DocID) Or RstStock!Invoice_DocID = "", RstStock!DocID, RstStock!Invoice_DocID) & "'")
                    If GRs.RecordCount > 0 Then
                        D_Per_MRP_TB1 = GRs!D_PER_MRP_TB
                        D_Per_MRP_TP1 = GRs!D_PER_MRP_TP
                        Gen_Sur_Per1 = GRs!Gen_Sur_Per
                        D_Per_TP1 = GRs!D_Per_TP
                        D_Per_TB1 = GRs!D_Per_TB
                    Else
                        D_Per_MRP_TB1 = 0
                        D_Per_MRP_TP1 = 0
                        Gen_Sur_Per1 = 0
                        D_Per_TP1 = 0
                        D_Per_TB1 = 0
                    End If
                    mAmount = RstStock!Net_Amt
                    If RstStock!Tax_YN = 1 And RstStock!MRP_YN = 1 Then  '' Taxable & MRP
                        mAmount = mAmount - Round(mAmount * D_Per_MRP_TB1 / 100, 2)
                    ElseIf RstStock!Tax_YN = 1 And RstStock!MRP_YN = 0 Then  '' Taxable & Non-MRP
                        mAmount = mAmount - Round(mAmount * D_Per_TB1 / 100, 2)
                    ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 1 Then  '' Taxpaid & MRP
                        mAmount = mAmount - Round(mAmount * D_Per_MRP_TP1 / 100, 2)
                    ElseIf RstStock!Tax_YN = 0 And RstStock!MRP_YN = 0 Then  '' TaxPaid & Non-MRP
                        mAmount = mAmount - Round(mAmount * D_Per_TP1 / 100, 2)
                    End If
                    If RstStock!Purpose = "C" Then
                        mJOB_SAL = mJOB_SAL + mAmount
                    Else
                        mJOB_FREE = mJOB_FREE + mAmount
                    End If
            End Select
            RstStock.MoveNext
        Loop
        mInvValue = 0
        
        ''' INVENTORY Calculation
        Set RstPart = GCn.Execute("Select Distinct Sp_Stock.Part_No,Part.Part_Name From Sp_Stock Left Join Part On SP_Stock.Part_No=Part.Part_No and Part.Div_Code = left(SP_Stock.Docid,1) " & Condstr)
        Do While Not RstPart.EOF
            NoUpto = 1
            mNo = 0
            Do While mNo <= NoUpto
                CondStr1 = " and MRP_yn=" & mNo
                Do While TRec1.RecordCount > 0
                   If TRec1.RecordCount > 0 Then TRec1.MoveFirst
                   TRec1.Delete
                   TRec1.Update
                Loop
                Do While TRec2.RecordCount > 0
                   If TRec2.RecordCount > 0 Then TRec2.MoveFirst
                   TRec2.Delete
                   TRec2.Update
                Loop
                
                mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
                mIss_TB_Qty = 0: mIss_TB_Val = 0: mIss_TP_Qty = 0: mIss_TP_Val = 0
                
                Set RstStock = GCn.Execute("select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and v_date<=" & ConvertDate(xDtTO) & CondStr1 & " order By Sp_Stock.V_Date,Vt.StkTrn")
                Do While Not RstStock.EOF
                    '' Add Record for Received Side
                    If RstStock!StkTrn = "+" Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            With TRec1
                                .AddNew
                                .Fields("Date") = RstStock!V_DATE
                                .Fields("Part_No") = RstStock!Part_No
                                .Fields("Qty") = RstStock!Qty_Rec
                                .Fields("Rate") = RstStock!V_Rate
                                .Update
                            End With
                        Else
                            With TRec2
                                .AddNew
                                .Fields("Date") = RstStock!V_DATE
                                .Fields("Part_No") = RstStock!Part_No
                                .Fields("Qty") = RstStock!Qty_Rec
                                .Fields("Rate") = RstStock!V_Rate
                                .Update
                            End With
                        End If
                    End If
                    RstStock.MoveNext
                Loop
                
                TRec1.Sort = "Date"
                TRec2.Sort = "Date"
                
                If TRec1.RecordCount > 0 Then TRec1.MoveFirst
                If TRec2.RecordCount > 0 Then TRec2.MoveFirst
                    
                Set RstStock = GCn.Execute("select Sp_stock.*,Vt.Stktrn,Vt.Description From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and  v_date<=" & ConvertDate(xDtTO) & CondStr1 & " order By Sp_Stock.V_Date,Vt.StkTrn")
                Do While Not RstStock.EOF
                    If RstStock!StkTrn = "-" Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            mRate = 0
                            Call X_VAL11(TRec1, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate, mNarr)
                        Else
                            mRate = 0
                            Call X_VAL22(TRec2, (RstStock!Qty_Iss - RstStock!Qty_Ret), mRate, mNarr)
                        End If
                    ElseIf RstStock!StkTrn = "+" Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                            mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * VNull(RstStock!V_Rate))
                        Else
                            mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                            mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * VNull(RstStock!V_Rate))
                        End If
                    End If
                    RstStock.MoveNext
                Loop
                mInvValue = mInvValue + mOP_TP_VAL + mOP_TB_VAL
                mNo = mNo + 1
            Loop
            RstPart.MoveNext
        Loop
        
        If Temp06.RecordCount > 0 Then
            Temp06.MoveFirst
            Temp06.FIND ("Net_Qty=" & mMth_No)
        End If
        If Temp06.EOF = True Or Temp06.BOF = True Then
            Temp06.AddNew
            Temp06.Fields("Net_Qty") = mMth_No
            Temp06.Fields("Narr") = mCMonth
'            Temp06.Update
 '           Temp06.FIND ("Net_Qty=" & mMth_No)
        End If
        With Temp06
            .Fields("TB_OQTY") = IIf(IsNull(Temp06!TB_OQty), 0, Temp06!TB_OQty) + mJOB_SAL
            .Fields("TB_OVAL") = IIf(IsNull(Temp06!TB_OVal), 0, Temp06!TB_OVal) + mJOB_FREE
            .Fields("TP_OQTY") = IIf(IsNull(Temp06!TP_OQty), 0, Temp06!TP_OQty) + mCOU_SAL
            .Fields("TP_OVAL") = IIf(IsNull(Temp06!TP_OVal), 0, Temp06!TP_OVal) + mCOU_TRF
            .Fields("Net_Val") = IIf(IsNull(Temp06!Net_Val), 0, Temp06!Net_Val) + mInvValue
            .Update
        End With
        xDtTO = xDtTO + 1
    Loop
    
    Set RstRep = Temp06.Clone
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprSaleInventory"
    RepTitle = UCase(Me.CAPTION)
ELoop:
    If err.NUMBER <> 0 Then CheckError
    Set GRs = Nothing
End Sub


Private Sub SprXYZAnalysis()
On Error GoTo ELoop
Dim mQry$, CondStr1$, Condstr$, CondStr3$
Dim mRecQty As Double, mIssQty As Double, mStkVal As Double
Dim XRecNo As Double
Dim mTB_StkVal As Double, mTP_StkVal As Double, mTB_StkQty As Double, mTP_StkQty As Double
Dim mNo As Byte, NoUpto As Byte, mVRate As Double, mDisPer As Double
Dim mCat$, mAdjVal As Double, mTotalPer As Double
    RepPrint = True
    
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    If IsInLimit(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsInLimit(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = "where SPStk.V_Date >=" & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SPStk.V_Date <=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    If Check1(1).Value = Unchecked Then CondStr1 = " and left(SPStk.DocID,1) in (" & GridString1 & ")"
    If Check1(3).Value = Unchecked Then CondStr3 = " and SPStk.Part_No in (" & GridString3 & ")"
    
    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
    
    mStkVal = 0
    'For RstPart, SQL
    GSQL = "Select Distinct SP_Stock.Part_No,Part.Part_Name From SP_Stock Left Join Part on Sp_Stock.Part_No=Part.Part_No "

    
    
    Set RstPart = GCn.Execute(GSQL)
    '********** Taxable Qty
    mQry = "Select SPStk.Part_No,SPStk.V_DATE,SPStk.Qty_Rec as Qty, " & cIIF("SPStk.V_Type='SXAO'", "SPStk.Rate", "SPStk.V_Rate") & " as Rate " & _
        "From " & _
        "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
        "where SPStk.V_Date>=" & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SPStk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & _
        " and Part_No in (select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    mQry = mQry & CondStr1 & CondStr3 & ") "
    '****
    GSQL = mQry & " and SpStk.Tax_YN=1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date"
    Set TRec1 = New Recordset
    With TRec1
        .CursorLocation = adUseClient
        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
    End With
    '********* Taxpaid Qty
    GSQL = mQry & " and SpStk.Tax_YN<>1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date"
    Set TRec2 = New Recordset
    With TRec2
        .CursorLocation = adUseClient
        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
    End With
    '********* Taxable + Taxpaid Qty for Processing Loop
    mQry = "select SPStk.V_Type,SPStk.Part_No,SPStk.V_DATE,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss, " & cIIF("SPStk.V_Type='SXAO'", "SPStk.Rate", "SPStk.V_Rate") & " as V_Rate,Vt.StkTrn " & _
        "From " & _
        "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
        "where SPStk.V_Date>=" & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SPStk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & _
        " and SPStk.Part_No in " & _
        "(select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
   If Check1(1).Value = Unchecked Then
        mQry = mQry & " and left(Stk.Docid,1) in (" & GridString1 & ")" & CondStr1 & ") "
    Else
        mQry = mQry & CondStr1 & ")"
    End If
    GSQL = mQry & " Order By SPStk.Part_No,SPStk.V_Date, " & cMID("SPStk.DocID", "4", "5") & ""
    Set RstStock = GCn.Execute(GSQL)
    
    Do While Not RstPart.EOF
        MDIForm1.Picture1.Visible = True
        MDIForm1.Label1.CAPTION = "Process Status : " & RstPart.AbsolutePosition & "/" & RstPart.RecordCount
        MDIForm1.Label1.Refresh
        NoUpto = 1
        mNo = 0
     
        
        mTB_StkVal = 0: mTP_StkVal = 0
        mTB_StkQty = 0: mTP_StkQty = 0
'        Do While mNo <= NoUpto
            
            mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
            mIss_TB_Qty = 0: mIss_TB_Val = 0: mIss_TP_Qty = 0: mIss_TP_Val = 0
            'mRec_TB_Qty = 0: mRec_TB_Val = 0: mRec_TP_Qty = 0: mRec_TP_Val = 0
           
            If TRec1.RecordCount > 0 Then
                TRec1.MoveFirst
                TRec1.Filter = ("Part_No='" & RstPart!Part_No & "'")
                If TRec1.EOF = False Then
                    TRec1Qty = TRec1!Qty
                End If
            End If
            
            If TRec2.RecordCount > 0 Then
                TRec2.MoveFirst
                TRec2.Filter = ("Part_No='" & RstPart!Part_No & "'")
                If TRec2.EOF = False Then
                    TRec2Qty = TRec2!Qty
                End If
            End If
            
            RstStock.MoveFirst
            RstStock.FIND ("Part_No='" & RstPart!Part_No & "'")
            If RstStock.EOF = False Then
                Do While RstStock!Part_No = RstPart!Part_No    'Opening Calculation
                    mNarr = ""
                    If RstStock!StkTrn = "-" Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            mRate = 0
                            Call X_VAL11(TRec1, RstStock!Qty_Iss, mRate, mNarr)
                        Else
                            mRate = 0
                            Call X_VAL22(TRec2, RstStock!Qty_Iss, mRate, mNarr)
                        End If
                    ElseIf RstStock!StkTrn = "+" Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                            mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * VNull(RstStock!V_Rate))
                        Else
                            mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                            mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * VNull(RstStock!V_Rate))
                        End If
                    End If
                    RstStock.MoveNext
                    If RstStock.EOF Then
                        Exit Do
                    ElseIf RstStock!Part_No <> RstPart!Part_No Then
                        Exit Do
                    End If
                Loop
            End If
            mTB_StkQty = mTB_StkQty + mOP_TB_QTY
            mTP_StkQty = mTP_StkQty + mOP_TP_QTY
            mTB_StkVal = mTB_StkVal + mOP_TB_VAL
            mTP_StkVal = mTP_StkVal + mOP_TP_VAL
            mNo = mNo + 1
 '       Loop
        
        If mTB_StkQty + mTP_StkQty + mTB_StkVal + mTP_StkVal > 0 Then
            With Temp06
                .AddNew
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Part_Name") = RstPart!Part_Name
                
                .Fields("TB_BQty") = mTB_StkQty
                .Fields("TB_BVal") = mTB_StkVal
                .Fields("TP_BQty") = mTP_StkQty
                .Fields("TP_BVal") = mTP_StkVal
                
                .Fields("Net_Qty") = mTB_StkQty + mTP_StkQty
                .Fields("Net_Val") = mTB_StkVal + mTP_StkVal
                .Update
            End With
            mStkVal = mStkVal + mTB_StkVal + mTP_StkVal
        End If
        RstPart.MoveNext
    Loop
    
    mCat = "X"
    mAdjVal = 0
    If Temp06.RecordCount > 0 Then
        Temp06.Sort = "Net_Val"
        Temp06.MoveLast
        Do While Not Temp06.BOF
            mAdjVal = mAdjVal + Temp06!Net_Val
            
            With Temp06
            .Fields("Rate") = Round(Temp06!Net_Val * 100 / mStkVal, 3)
            mTotalPer = mTotalPer + .Fields("Rate")
            If mTotalPer < Val(FGrid.TextMatrix(Cat1, 1)) Then
                mCat = "X"
            ElseIf mTotalPer >= Val(FGrid.TextMatrix(Cat1, 1)) And mTotalPer <= (Val(FGrid.TextMatrix(Cat1, 1)) + Val(FGrid.TextMatrix(Cat2, 1))) Then
                mCat = "Y"
            ElseIf mTotalPer > (Val(FGrid.TextMatrix(Cat1, 1)) + Val(FGrid.TextMatrix(Cat2, 1))) Then
                mCat = "Z"
            End If
                .Fields("Narr") = mCat
                .Update
            End With
            Temp06.MovePrevious
        Loop
    End If
'    If Temp06.RecordCount > 0 Then
'        Temp06.Sort = "Net_Val"
'        Temp06.MoveLast
'        Do While Not Temp06.BOF
'            mAdjVal = mAdjVal + Temp06!Net_Val
'
'            With Temp06
'            .Fields("Rate") = Round(Temp06!Net_Val * 100 / mStkVal, 3)
'            mTotalPer = mTotalPer + .Fields("Rate")
'            If mTotalPer < Val(FGrid.TextMatrix(Cat1, 1)) Then
'                mCat = "X"
'            ElseIf mTotalPer >= Val(FGrid.TextMatrix(Cat1, 1)) And mTotalPer <= (Val(FGrid.TextMatrix(Cat1, 1)) + Val(FGrid.TextMatrix(Cat2, 1))) Then
'                mCat = "Y"
'            ElseIf mTotalPer > (Val(FGrid.TextMatrix(Cat1, 1)) + Val(FGrid.TextMatrix(Cat2, 1))) Then
'                mCat = "Z"
'            End If
'                .Fields("Narr") = mCat
'                .Update
'            End With
'            Temp06.MovePrevious
'        Loop
'    End If
'
    Set RstRep = Temp06.Clone
    Set TRec1 = Nothing
    Set TRec2 = Nothing
    Set RstStock = Nothing
    Set RstStock2 = Nothing
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprXYZ"
    RepTitle = UCase(Me.CAPTION)
    MDIForm1.Picture1.Visible = False
ELoop:
    Set TRec1 = Nothing
    Set TRec2 = Nothing
    Set RstStock = Nothing
    Set RstStock2 = Nothing
    Set GRs = Nothing
    If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub SprFSNAnalysis()
On Error Resume Next
Dim Condstr$
Dim mOP_TB_Stk As Double, mOP_TP_Stk As Double, MRPOP_TB_Stk As Double, MRPOP_TP_Stk As Double
Dim mTB_Issue As Double, mTP_Issue As Double, mTB_Rect As Double, mTP_Rect As Double
Dim MRPTB_Issue As Double, MRPTP_Issue As Double, MRPTB_Rect As Double, MRPTP_Rect As Double
Dim mOP_Stock As Double, mRect As Double, mIssue As Double, mTB_ClStk As Double, mTP_ClStk As Double, mTB_CLVal As Double, mTP_CLVal As Double
Dim MRPOP_Stock As Double, MRPRect As Double, MRPIssue As Double, MRPTB_ClStk As Double, MRPTP_ClStk As Double, MRPTB_ClVal As Double, MRPTP_ClVal As Double
Dim mCLos_Stk As Double, MRPClos_Stk As Double, mPer As Double, Tot_Issues As Double, mCLos_Val As Double
Dim mType$, mCat As Byte
Dim RstStock1 As ADODB.Recordset
    RepPrint = True
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    If IsInLimit(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then RepPrint = False: Exit Sub
    If IsInLimit(Cat2, FGrid.TextMatrix(Cat2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Val(FGrid.TextMatrix(Cat2, 1)) > Val(FGrid.TextMatrix(Cat1, 1)) Then
        MsgBox FGrid.TextMatrix(Cat2, 0) & " >" & FGrid.TextMatrix(Cat1, 0), vbOKOnly, "Validation"
        FGrid.SetFocus:  FGrid.Row = Cat2: FGrid.Col = 1
        RepPrint = False: Exit Sub
    End If
    Condstr = "where V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
    Set RstPart = GCn.Execute("Select Distinct Sp_Stock.Part_No,Part.Part_Name,PurcDisc_Per From (Sp_Stock Left Join Part On SP_Stock.Part_No=Part.Part_No) Left Join Part_DiscFactor on Part.Disc_Factor=Part_DiscFactor.DiscFac_Catg " & Condstr)
    Do While Not RstPart.EOF
        MDIForm1.Picture1.Visible = True
        MDIForm1.Label1.CAPTION = "Process Status : " & RstPart.AbsolutePosition & "/" & RstPart.RecordCount
        MDIForm1.Label1.Refresh
        mOP_TB_Stk = 0: mOP_TP_Stk = 0
        MRPOP_TB_Stk = 0: MRPOP_TP_Stk = 0
        mTB_Issue = 0: mTP_Issue = 0: mTB_Rect = 0: mTP_Rect = 0
        MRPTB_Issue = 0: MRPTP_Issue = 0: MRPTB_Rect = 0: MRPTP_Rect = 0
        If PubBackEnd = "A" Then
            GSQL = "select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and (v_date >= " & ConvertDate(PubStartDate) & " AND v_date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ")  order By Sp_Stock.V_Date,Vt.StkTrn"
            GSQL = GSQL + " Union All "
            GSQL = GSQL + "select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and (v_date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SP_Stock.V_Type='SXAO')  order By Sp_Stock.V_Date,Vt.StkTrn"
        ElseIf PubBackEnd = "S" Then
            GSQL = "select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and (v_date >= " & ConvertDate(PubStartDate) & " AND v_date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ")"
            GSQL = GSQL + " Union All "
            GSQL = GSQL + "select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and (v_date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SP_Stock.V_Type='SXAO')  order By Sp_Stock.V_Date,Vt.StkTrn"
        End If
        
        Set RstStock = GCn.Execute(GSQL)
        
        'Set RstStock = GCn.Execute("select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join [" & PubSFADataPath & "].Voucher_Type vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and (v_date >= #" & PubStartDate & "# AND v_date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ")  order By Sp_Stock.V_Date,Vt.StkTrn")
        
        
'        Set RstStock = GCn.Execute("select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join [" & PubSFADataPath & "].Voucher_Type vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & RstPart!Part_No & "' and  v_date >= #" & PubStartDate & "# and v_date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " order By Sp_Stock.V_Date,Vt.StkTrn")
        Do While Not RstStock.EOF
            If RstStock!V_DATE >= CDate(FGrid.TextMatrix(Date1, 1)) And RstStock!V_DATE <= CDate(FGrid.TextMatrix(Date2, 1)) Then
                If RstStock!Tax_YN = 1 Then     '' Taxable
                    mTB_Issue = mTB_Issue + (RstStock!Qty_Iss - RstStock!Qty_Ret)
                Else
                    mTP_Issue = mTP_Issue + (RstStock!Qty_Iss - RstStock!Qty_Ret)
                End If
            End If
            If RstStock!MRP_YN = 1 Then
                If RstStock!StkTrn = "-" Then
                    If RstStock!V_DATE < CDate(FGrid.TextMatrix(Date1, 1)) Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            MRPOP_TB_Stk = MRPOP_TB_Stk - (RstStock1!Qty_Iss - RstStock1!Qty_Ret)
                        Else
                            MRPOP_TP_Stk = MRPOP_TP_Stk - (RstStock1!Qty_Iss - RstStock1!Qty_Ret)
                        End If
                    Else
                        If RstStock!V_DATE >= CDate(FGrid.TextMatrix(Date1, 1)) And RstStock!V_DATE <= CDate(FGrid.TextMatrix(Date2, 1)) Then
                            If RstStock!Tax_YN = 1 Then     '' Taxable
                                MRPTB_Issue = MRPTB_Issue + (RstStock!Qty_Iss - RstStock!Qty_Ret)
                            Else
                                MRPTP_Issue = MRPTP_Issue + (RstStock!Qty_Iss - RstStock!Qty_Ret)
                            End If
                        End If
                    End If
                Else
                    If RstStock!V_DATE < CDate(FGrid.TextMatrix(Date1, 1)) Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            MRPOP_TB_Stk = MRPOP_TB_Stk + RstStock!Qty_Rec
                        Else
                            MRPOP_TP_Stk = MRPOP_TP_Stk + RstStock!Qty_Rec
                        End If
                    Else
                        If RstStock!V_DATE >= CDate(FGrid.TextMatrix(Date1, 1)) And RstStock!V_DATE <= CDate(FGrid.TextMatrix(Date2, 1)) Then
                            If RstStock!Tax_YN = 1 Then     '' Taxable
                                MRPTB_Rect = MRPTB_Rect + RstStock!Qty_Rec
                            Else
                                MRPTP_Rect = MRPTP_Rect + RstStock!Qty_Rec
                            End If
                        End If
                    End If
                End If
            Else
                If RstStock!StkTrn = "-" Then
                    If RstStock!V_DATE < CDate(FGrid.TextMatrix(Date1, 1)) Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            mOP_TB_Stk = mOP_TB_Stk - (RstStock!Qty_Iss - RstStock!Qty_Ret)
                        Else
                            mOP_TP_Stk = mOP_TP_Stk - (RstStock!Qty_Iss - RstStock!Qty_Ret)
                        End If
                    
                    End If
                Else
                    If RstStock!V_DATE < CDate(FGrid.TextMatrix(Date1, 1)) Then
                        If RstStock!Tax_YN = 1 Then     '' Taxable
                            mOP_TB_Stk = mOP_TB_Stk + RstStock!Qty_Rec
                        Else
                            mOP_TP_Stk = mOP_TP_Stk + RstStock!Qty_Rec
                        End If
                    Else
                        If RstStock!V_DATE >= CDate(FGrid.TextMatrix(Date1, 1)) And RstStock!V_DATE <= CDate(FGrid.TextMatrix(Date2, 1)) Then
                            If RstStock!Tax_YN = 1 Then     '' Taxable
                                mTB_Rect = mTB_Rect + RstStock!Qty_Rec
                            Else
                                mTP_Rect = mTP_Rect + RstStock!Qty_Rec
                            End If
                        End If
                    End If
                End If
            End If
            RstStock.MoveNext
        Loop
        
        mOP_Stock = mOP_TB_Stk + mOP_TP_Stk
        mRect = mTB_Rect + mTP_Rect
        mIssue = mTB_Issue + mTP_Issue
        mTB_ClStk = (mOP_TB_Stk + mTB_Rect) - mTB_Issue
        mTP_ClStk = (mOP_TP_Stk + mTP_Rect) - mTP_Issue
        'mTB_CLVal = GetFIFOAmt(RstPart!Part_No, CDate(FGrid.TextMatrix(Date2, 1)), mTB_ClStk, 0, 1)
        'mTP_CLVal = GetFIFOAmt(RstPart!Part_No, CDate(FGrid.TextMatrix(Date2, 1)), mTP_ClStk, 0, 0)
        
        MRPOP_Stock = MRPOP_TB_Stk + MRPOP_TP_Stk
        MRPRect = MRPTB_Rect + MRPTP_Rect
        MRPIssue = MRPTB_Issue + MRPTP_Issue
        MRPTB_ClStk = (MRPOP_TB_Stk + MRPTB_Rect) - MRPTB_Issue
        MRPTP_ClStk = (MRPOP_TP_Stk + MRPTP_Rect) - MRPTP_Issue
        'MRPTB_ClVal = GetFIFOAmt(RstPart!Part_No, CDate(FGrid.TextMatrix(Date2, 1)), MRPTB_ClStk, 1, 1)
        'MRPTP_ClVal = GetFIFOAmt(RstPart!Part_No, CDate(FGrid.TextMatrix(Date2, 1)), MRPTP_ClStk, 1, 0)
        
        mCLos_Stk = mTB_ClStk + mTP_ClStk
        MRPClos_Stk = MRPTB_ClStk + MRPTP_ClStk
        mCLos_Val = 0
        If left(UCase(PubComp_Name), 3) = "JMK" Then
            mCLos_Val = GCn.Execute("select TB_SRate from Part where Part_No='" & RstPart!Part_No & "'").Fields(0).Value
        Else
            mCLos_Val = GCn.Execute("select MRP from Part where Part_No='" & RstPart!Part_No & "'").Fields(0).Value
        End If
        
      
           If (mCLos_Stk + MRPClos_Stk) <> 0 Then
            With Temp06
                .AddNew
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Part_Name") = RstPart!Part_Name
                .Fields("Inv_No") = mCat        '' 1/2/3
                .Fields("TB_OQty") = mOP_Stock + MRPOP_Stock
                
                If left(UCase(PubComp_Name), 3) = "JMK" Then
                    If FGrid.TextMatrix(List1, 1) = "No" Then
                        If ((mOP_Stock + MRPOP_Stock) - (mIssue)) > 0 Then
                            .Fields("Re_TB") = 0
                            .Fields("Is_TB") = mIssue ' + MRPIssue
                            .Fields("TB_BQty") = (mOP_Stock + MRPOP_Stock) - (mIssue) 'mCLos_Stk '+ MRPClos_Stk
                        End If
                    Else
                        .Fields("Re_TB") = mRect + MRPRect
                        .Fields("Is_TB") = mIssue ' + MRPIssue
                        .Fields("TB_BQty") = (mOP_Stock + MRPOP_Stock) + MRPRect - (mIssue - mRect) 'mCLos_Stk '+ MRPClos_Stk
                    End If
                Else
                    .Fields("Re_TB") = mRect + MRPRect
                    .Fields("Is_TB") = mIssue ' + MRPIssue
                    .Fields("TB_BQty") = (mOP_Stock + MRPOP_Stock) + MRPRect - (mIssue - mRect) 'mCLos_Stk '+ MRPClos_Stk
                End If
                .Fields("Net_Val") = .Fields("TB_BQty") * mCLos_Val
                If .Fields("Net_Val") >= 0 Then
                   If left(UCase(PubComp_Name), 3) = "JMK" Then
                        If (.Fields("TB_OQty") + .Fields("Re_TB")) > 0 Then
                            mPer = Round((.Fields("Is_TB") * 100) / (.Fields("TB_OQty") + .Fields("Re_TB")), 2)
                        End If
                        .Fields("MovePer") = mPer
                        If mPer >= Val(FGrid.TextMatrix(Cat1, 1)) Then
                            .Fields("mType") = "FAST"
                            
                        ElseIf mPer >= Val(FGrid.TextMatrix(Cat2, 1)) And mPer < Val(FGrid.TextMatrix(Cat1, 1)) Then
                            .Fields("mType") = "SLOW"
                        ElseIf mPer < Val(FGrid.TextMatrix(Cat2, 1)) Then
                            .Fields("mType") = "DEAD"
                        End If
                   Else
                        If FGrid.TextMatrix(List1, 1) = "No" Then
                            mPer = Round((.Fields("Is_TB") / .Fields("TB_OQty")) * 100, 2)
                        Else
                            If (.Fields("TB_OQty") + .Fields("Re_TB")) > 0 Then
                                mPer = Round((.Fields("Is_TB") * 100) / (.Fields("TB_OQty") + .Fields("Re_TB")), 2)
                            End If
                        End If
                        .Fields("MovePer") = mPer
                        If mPer >= Val(FGrid.TextMatrix(Cat1, 1)) Then
                            .Fields("mType") = "FAST"
                            
                        ElseIf mPer >= Val(FGrid.TextMatrix(Cat2, 1)) And mPer < Val(FGrid.TextMatrix(Cat1, 1)) Then
                            .Fields("mType") = "SLOW"
                        ElseIf mPer < Val(FGrid.TextMatrix(Cat2, 1)) Then
                            .Fields("mType") = "DEAD"
                        End If
                    End If
                    .Update
                End If
            End With
        End If
MyNext:
        RstPart.MoveNext
    Loop
                  
    Set RstRep = Temp06.Clone
    If StrCmp(FGrid.TextMatrix(List2, 1), "Fast") Then
        RstRep.Filter = "mType='Fast'"
    ElseIf StrCmp(FGrid.TextMatrix(List2, 1), "Slow") Then
        RstRep.Filter = "mType='Slow'"
    ElseIf StrCmp(FGrid.TextMatrix(List2, 1), "Dead") Then
        RstRep.Filter = "mType='Dead'"
    End If
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    'RstRep.MoveFirst
    'While Not RstRep.EOF = True
    '    Tot_Issues = Tot_Issues + RstRep.Fields("Is_TB")
    '    RstRep.MoveNext
    'Wend
'    RstRep.MoveFirst
'    While Not RstRep.EOF = True
'        mPer = Round((RstRep.Fields("Is_TB") / RstRep.Fields("TB_OQty")) * 100, 2)
'                If mPer >= Val(FGrid.TextMatrix(Cat1, 1)) Then
'                    mType = "FAST"
'                    mCat = 1
'                ElseIf mPer >= Val(FGrid.TextMatrix(Cat2, 1)) And mPer < Val(FGrid.TextMatrix(Cat1, 1)) Then
'                    mType = "SLOW"
'                    mCat = 2
'                ElseIf mPer < Val(FGrid.TextMatrix(Cat2, 1)) Then
'                    mType = "DEAD"
'                    mCat = 2
'                End If
'                If FGrid.TextMatrix(List2, 1) = "All" Then
'                    RstRep.Fields("Inv_Date") = mType: GoTo MyNext1
'                End If
'                If UCase(FGrid.TextMatrix(List2, 1)) = mType Then
'                    RstRep.Fields("Inv_Date") = mType: GoTo MyNext1
'                Else
'                    RstRep.Delete
''                ElseIf UCase(FGrid.TextMatrix(List2, 1)) = mType Then
''                    RstRep.Fields("Inv_Date") = mType: GoTo MyNext1
''                ElseIf UCase(FGrid.TextMatrix(List2, 1)) = mType Then
''                    RstRep.Fields("Inv_Date") = mType: GoTo MyNext1
'                End If
'
'                'RstRep.Fields("Inv_Date") = mType       '' Fast/Slow/Dead
'                'End If
'MyNext1:
'        RstRep.MoveNext
'    Wend
    RepName = "SprFSN"
    RepTitle = UCase(Me.CAPTION)
    MDIForm1.Picture1.Visible = False
Exit Sub
ELoop:
    If err.NUMBER <> 0 Then CheckError
    Set GRs = Nothing
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

''
''Private Function FillString(GridArray As Variant, Gridindex As Integer, DataType As Byte) As String
''Dim ac_str$
''Dim i As Integer
''Dim GridRow As Integer
''    ac_str = ""
''    For i = 0 To UBound(GridArray)
''        If GridArray(i) = 0 Then GoTo NXT:
''        GridRow = GridArray(i)
''        If GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü" Then
''                If DataType = 0 Then
''                   ac_str = ac_str + IIf(ac_str = "", GridSel(Gridindex).TextMatrix(GridRow, 2), "," + GridSel(Gridindex).TextMatrix(GridRow, 2))
''                ElseIf DataType = 1 Then
''                   ac_str = ac_str + IIf(ac_str = "", "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'", "," + "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'")
''                End If
''            GridSel(Gridindex).TextMatrix(GridRow, 0) = ""
''        Else
''            GridArray(i) = 0
''        End If
''NXT:
''    Next
''    For i = 0 To UBound(GridArray)
''        GridRow = GridArray(i)
''        If GridArray(i) <> 0 Then
''            GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü"
''        End If
''    Next
'''    Erase GridArray
'''    ReDim Preserve GridArray(0)
'''    GridArray(0) = 0
''    If ac_str = "" Then
''        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
''        GridSel(Gridindex).SetFocus
''        RepPrint = False
''        Exit Function
''    End If
''    FillString = ac_str
''    Exit Function
''End Function

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
    If Index = 2 Then: GridSel(Index).ColWidth(1) = 2000: GridSel(Index).ColWidth(2) = 2000
    Check1(Index).top = GridSel(Index).top + 20: Check1(Index).left = GridSel(Index).left + 40: Check1(Index).width = 560
    Check1(Index).height = GridSel(Index).RowHeight(0) + 40: Check1(Index).Value = Checked
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

Private Function IsInLimit(FieldRow As Integer, FieldCaption As String) As Boolean
    If Val(FGrid.TextMatrix(FieldRow, 1)) > 100 Then
        MsgBox FieldCaption & " Should not be Greater then 100 %.", vbInformation, "Validation Check"
        FGrid.SetFocus
        FGrid.Row = FieldRow
        FGrid.Col = 1
        IsInLimit = False
    Else
        IsInLimit = True
    End If
End Function
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
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
    Dim I As Byte
    
        If GRepFormName = SprStkLedValSum Then
            Set RsPart1 = New ADODB.Recordset
            RsPart1.CursorLocation = adUseClient
            If PubBackEnd = "A" Then
                RsPart1.Open "Select P.Part_No AS Code, " & xIsNull("P.Part_Name", "") & " as Name, " & xIsNull("P.Local_Name", "") & " AS LName, " & xIsNull("P.Unit", "") & " as Unit," _
                & "format(P.MRP,'0.00') as MRP,format(P.TB_SRate,'0.00') as TB_SRate,format(P.TP_SRate,'0.00') as TP_SRate,P.Bin_Loca " _
                & "From Part P " _
                & "Order By P.Part_No,P.Part_Name,P.Local_Name", GCn, adOpenDynamic, adLockOptimistic
            ElseIf PubBackEnd = "S" Then
                RsPart1.Open "Select P.Part_No AS Code, " & xIsNull("P.Part_Name", "") & " as Name, " & xIsNull("P.Local_Name", "") & " AS LName, " & xIsNull("P.Unit", "") & " as Unit," _
                & "P.MRP as MRP, P.TB_SRate as TB_SRate,P.TP_SRate as TP_SRate,P.Bin_Loca " _
                & "From Part P " _
                & "Order By P.Part_No,P.Part_Name,P.Local_Name", GCn, adOpenDynamic, adLockOptimistic
            End If
    
            RsPart1.Sort = "Code"
        End If
    
    WinSetting Me
    Global_Grid
    TopCtrl1.TopText2 = "Add"
'       If Mid(UserPermission(Me.Name), 4, 1) = "*" Then BTNPRINT.Enabled = False
       Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MDIForm1.Picture1.Visible = True Then MDIForm1.Picture1.Visible = False
    If GridSel(4).Visible = True Then Set RsGrid1 = Nothing
    If GridSel(1).Visible = True Then Set RsGrid2 = Nothing
    If GridSel(2).Visible = True Then Set RsGrid3 = Nothing
    If GridSel(3).Visible = True Then Set RsGrid4 = Nothing
    Set RstRep = Nothing
    Set mListItem = Nothing
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
           SelGridKeyPress TxtSearch, GridSel(Index), RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
        Case 2
           SelGridKeyPress TxtSearch, GridSel(Index), RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
        Case 3
           SelGridKeyPress TxtSearch, GridSel(Index), RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
        Case 4
           SelGridKeyPress TxtSearch, GridSel(Index), RsGrid4, KeyAscii, RsGrid4.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    End Select
    TxtSearch.Tag = Index
End Sub

Private Sub TxtSearch_Click()
    TxtSearch.Visible = False: TxtSearch.TEXT = "": GridSel(Val(TxtSearch.Tag)).SetFocus
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If NavigationKey(KeyCode) = True Then TxtSearch.Visible = False: GridSel(Val(TxtSearch.Tag)).SetFocus
    If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then TxtSearch.Visible = False: GridSel(Val(TxtSearch.Tag)).SetFocus
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
    TxtSearch.Visible = False: TxtSearch.TEXT = ""
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

Private Sub Grid_Hide()
If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub btnexit_Click()
    Unload Me
End Sub
'******* Fuctions **********

Private Sub Global_Grid()
Dim I As Integer
    Frame1.top = 775: Frame1.left = 300: FGrid.top = 75: FGrid.left = 75
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
End Sub

Private Function FLDAY(ByVal xxCURDATE As Date, xxOPT As String) As Date
Dim xxFdate As Date, xxLDate As Date
Dim xxMONTH  As Byte, xxYEAR As Integer
    xxFdate = Format("01/" & Format(xxCURDATE, "MMM") & "/" & Format(xxCURDATE, "yyyy"), "dd/MMM/yyyy")
    xxMONTH = Format(xxCURDATE, "MM")
    xxYEAR = Format(xxCURDATE, "yyyy")
    Do While Format(xxCURDATE, "MM") = xxMONTH And Format(xxCURDATE, "yyyy") = xxYEAR
        xxCURDATE = xxCURDATE + 1
    Loop
    xxLDate = xxCURDATE - 1
    FLDAY = (IIf(UCase(xxOPT) = "F", xxFdate, xxLDate))
End Function

Private Function GetFIFOAmt(ByVal mPartNo As String, ByVal Uptodate As Date, ByVal BalStk As Double, ByVal mMRP_YN As Byte, ByVal mTax_YN As Byte) As Double
Dim mAmt As Double, tAmt As Double, mRate As Double
    Set GRs = GCn.Execute("select Sp_stock.*,Vt.Stktrn From Sp_Stock left Join " & FaTable("Voucher_Type") & " vt on Vt.V_type=Sp_Stock.V_type where Part_No='" & mPartNo & "' and v_date<=" & ConvertDate(Uptodate) & " and MRP_YN=" & mMRP_YN & " and Tax_YN=" & mTax_YN & " and vt.stktrn='+' order By Sp_Stock.V_Date,Vt.StkTrn")
    mAmt = 0
    If GRs.RecordCount > 0 Then
        Do While Not GRs.EOF And BalStk > 0
            tAmt = GRs!Qty_Rec * IIf(IsNull(GRs!V_Rate), 0, GRs!V_Rate)
            If GRs!Qty_Rec < BalStk Then
                BalStk = BalStk - GRs!Qty_Rec
                mAmt = mAmt + tAmt
            Else
                mRate = Round(tAmt / GRs!Qty_Rec, 8)
                mAmt = mAmt + Round(BalStk * mRate, 2)
                BalStk = 0
                Exit Do
            End If
            GRs.MoveNext
            If GRs.EOF = True Then Exit Do
        Loop
    End If
    GetFIFOAmt = mAmt
    Set GRs = Nothing
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

Private Sub SprStkLedValSumCalc1_Old()
On Error GoTo ELoop
Dim mQry$, Condstr$, CondDivCode$, CondMarkYN$, CondPartNos$, CondPartNos1$, CondDivCode1$
Dim CondStrMRP$, CondPartNosOpStk$, CondPartNosTrn$, Part_Name$, DivStr1$
Dim mRecQty As Double, mIssQty As Double, mStkVal As Double, noofDiv As Integer, d As Integer
Dim XRecNo As Double, DivStr$
Dim mNo As Byte, NoUpto As Byte
Dim RstDiv As ADODB.Recordset
    RepPrint = True
    GridString1 = Empty: GridString2 = Empty: GridString3 = Empty
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

'   Condstr = " where SPStk.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    Condstr = "where SPStk.V_Date >=" & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SPStk.V_Date <=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
    
   
    DivStr1 = GridString1
    Set RstDiv = GCn.Execute("Select Div_Code from Division")
    If Len(GridString1) = 3 Then
        noofDiv = 1
    ElseIf GridString1 = "" Then
        noofDiv = RstDiv.RecordCount
    Else
        noofDiv = 2
    End If
    RstDiv.MoveFirst
    For d = 1 To noofDiv
                If GridString1 = "" Or Len(DivStr1) = 7 Or noofDiv = 2 Then
                    GridString1 = "'" & RstDiv!Div_Code & "'"
                End If
                CondDivCode = " and left(SPStk.DocID,1) in (" & GridString1 & ")"
                CondDivCode1 = " and left(Stk.DocID,1) in (" & GridString1 & ")"
                
                Condstr = Condstr & CondDivCode
                If FGrid.TextMatrix(List1, 1) = "Yes" Then          '' Only for Marked Parts
                    CondMarkYN = " and Mark_YN='Y'"
                Else
                   If Check1(3).Value = Unchecked Then
                        CondPartNos = " and SPStk.Part_No in (" & GridString3 & ")"
                        CondPartNos1 = " Part.Part_No in " & "(" & GridString3 & ")"
                        CondPartNosOpStk = CondPartNos
                    Else
                        CondPartNos = " and SPStk.Part_No in (select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
                        CondPartNos1 = " Part.Part_No in ( select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
                        CondPartNosOpStk = ""
                    End If
                
                End If
                ' Properitary grade check *******
                If Check1(2).Value = Unchecked Then
                         CondPartNos1 = CondPartNos1 & " and Part.Part_Grade in (" & GridString2 & ")"
                End If
                '***************
                
                ' Bill Status Check *******
                If FGrid.TextMatrix(List4, 1) = "Yes" Then
                    If FGrid.TextMatrix(List3, 1) = "Detail" Or FGrid.TextMatrix(List3, 1) = "Summary" Then
                        CondPartNos = CondPartNos & " and SPStk.Invoice_DocID = '' and SPStk.V_Type in('W_RG','SYSC')"
                    End If
                End If
                ' ************
                If FGrid.TextMatrix(List2, 1) = "No" Then       '' For MRP_YN
                     CondStrMRP = " and SPStk.MRP_YN=0"
                End If
                
                Condstr = Condstr & CondStrMRP
                
               
                
                'For RstPart, SQL
                GSQL = "Select Distinct Part.Part_No From Part " & _
                    "where " & CondPartNos1
                If Check1(1).Value = Unchecked Then
                    GSQL = GSQL & IIf(CondPartNos1 = "", "", " and") & " Part.Div_Code in (" & GridString1 & ") "
                    DivStr = "and Part.Div_Code in(" & GridString1 & ")"
                Else
                    DivStr = ""
                End If
                GSQL = GSQL & CondMarkYN & " Order By Part.Part_No"
                
                Set RstPart = GCn.Execute(GSQL)
                    
                    'Stock Valuation FIFO Summary
                    If FGrid.TextMatrix(List3, 1) = "Summary" Then
                        RepName = "StockLedValueSum"
                    ElseIf FGrid.TextMatrix(List3, 1) = "Detail" Then
                        RepName = "StockLedValueDet"
                    End If
                    '********** Taxable Qty
                    mQry = "select SPStk.Part_No,SPStk.V_DATE,SPStk.Qty_Rec as Qty,SPStk.MRP_YN, " & cIIF("SPStk.V_Type='SXAO'", "SPStk.Rate", "SPStk.V_Rate") & " as Rate " & _
                        "From " & _
                        "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
                        Condstr & CondPartNos
                    GSQL = mQry & " and SpStk.Tax_YN=1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
                    Set TRec1 = New Recordset
                    With TRec1
                        .CursorLocation = adUseClient
                        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
                    End With
                    '******* Taxpaid Qty
                    GSQL = mQry & " and SpStk.Tax_YN<>1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
                    Set TRec2 = New Recordset
                    With TRec2
                        .CursorLocation = adUseClient
                        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
                    End With
                    '******* Taxable + Taxpaid Qty for Opening Loop
                    
                    mQry = "select SPStk.V_Type,SPStk.Part_No,SPStk.V_DATE,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss, " & cIIF("SPStk.V_Type='SXAO'", "SPStk.Rate", "SPStk.V_Rate") & " as V_Rate,Vt.StkTrn " & _
                        "From " & _
                        "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
                        Condstr & CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
                    GSQL = mQry & " and SPStk.V_Type='SXAO' Order By SPStk.Part_No,SPStk.V_Date, " & cMID("SPStk.DocID", "4", "5") & ""
                    Set RstStock = GCn.Execute(GSQL)
                    '******* Taxable + Taxpaid Qty for With in Date Period Loop
                    GSQL = "select SPStk.V_Type,SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss, " & cIIF("SPStk.V_Type='SXAO'", "SPStk.Rate", "SPStk.V_Rate") & " as V_Rate,Vt.StkTrn,Vt.Description " & _
                        "From " & _
                        "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & "  as VT on Vt.V_type=SPStk.V_type " & _
                        "where (SPStk.V_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And SPStk.v_date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " ) " & _
                        CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
                    GSQL = GSQL & " Order By SPStk.Part_No,SPStk.V_Date,SPStk.Tax_YN, " & cMID("SPStk.DocID", "4", "5") & ""
                    Set RstStock2 = GCn.Execute(GSQL)
                    '***********
                    Set tempRst = RstStock.Clone
                    Set TempRst1 = RstStock2.Clone
                    Dim I As Integer
                    MDIForm1.Picture1.Visible = True
                    Do While Not RstPart.EOF
                        'NRA Update
                        MDIForm1.Label1.CAPTION = "Process Status : " & RstPart.AbsolutePosition & "/" & RstPart.RecordCount
                        MDIForm1.Label1.Refresh
                        mVRate = 0: mOPVRate = 0: mDisPer = 0: TempVal = 0
                
                        If FGrid.TextMatrix(List2, 1) = "Yes" Then          '' Only for MRP Parts
                            NoUpto = 1
                            mNo = 1
                        ElseIf FGrid.TextMatrix(List2, 1) = "No" Then          '' Only for Non-MRP Parts
                            NoUpto = 0
                            mNo = 0
                        ElseIf FGrid.TextMatrix(List2, 1) = "All" Then          '' For All Parts
                            NoUpto = 1
                            mNo = 0
                        End If
                        TRec1Qty = 0
                        TRec2Qty = 0
                        
                        mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
                        mIss_TB_Qty = 0: mIss_TB_Val = 0: mIss_TP_Qty = 0: mIss_TP_Val = 0
                        mRec_TB_Qty = 0: mRec_TB_Val = 0: mRec_TP_Qty = 0: mRec_TP_Val = 0
                        
                        TRec1.Filter = ""
                        mOPVRate = 0
                        If TRec1.RecordCount > 0 Then    'Taxable Rect
                            TRec1.MoveFirst
                            TRec1.Filter = ("Part_No='" & RstPart!Part_No & "'")
                            'Nra Update
                            If TRec1.RecordCount > 0 Then
                                TRec1.MoveFirst
                                If TRec1!Rate <> 0 Then
                                    mOPVRate = TRec1!Rate
                                Else
                                    If TRec1!MRP_YN = 1 Then
                                        mOPVRate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                                    Else
                                        mOPVRate = GCn.Execute("Select TB_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                                        mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
                                        mOPVRate = mOPVRate - ((mOPVRate * mDisPer) / 100)
                                    End If
                                End If
                            TRec1.MoveFirst
                            End If
            
                            'End Update
                            
                            If TRec1.EOF = False Then
                                TRec1Qty = TRec1!Qty
                            End If
                        End If
                        mVRate = 0
                        TRec2.Filter = ""
                        If TRec2.RecordCount > 0 Then    'Taxpaid Rect
                            TRec2.MoveFirst
                            TRec2.Filter = ("Part_No='" & RstPart!Part_No & "'")
                            'Nra Update
                            If TRec2.RecordCount > 0 Then
                                TRec2.MoveLast
                                If TRec2!Rate <> 0 Then
                                    mVRate = TRec2!Rate
                                Else
'                                    If TRec2!MRP_YN = 1 Then
'                                        mVRate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
'                                    Else
                                        mVRate = GCn.Execute("Select TP_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                                        mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
                                        mVRate = mVRate - ((mVRate * mDisPer) / 100)
                                    'End If
                                End If
                            TRec2.MoveFirst
                            End If
                            
                            'End Update
                            If TRec2.EOF = False Then
                                TRec2Qty = TRec2!Qty
                            End If
                        End If
                        If RstStock.RecordCount > 0 Then
                            RstStock.MoveFirst
                            RstStock.FIND ("Part_No='" & RstPart!Part_No & "'")
                            If RstStock.EOF = False Then
                                Do While RstStock!Part_No = RstPart!Part_No    'Opening Calculation
                                    If RstStock!StkTrn = "-" Then
                                        If RstStock!Tax_YN = 1 Then     '' Taxable
                                            mRate = 0
                                            Call X_VAL11(TRec1, RstStock!Qty_Iss, mRate)
                                        Else
                                            mRate = 0
                                            Call X_VAL22(TRec2, RstStock!Qty_Iss, mRate)
                                        End If
                                    ElseIf RstStock!StkTrn = "+" Then
                                        If RstStock!Tax_YN = 1 Then     '' Taxable
                                            mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                                            mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * mOPVRate)
                                        Else
                                            mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                                            mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * mVRate)
                                        End If
                                    End If
                                    RstStock.MoveNext
                                    If RstStock.EOF Then
                                        Exit Do
                                    ElseIf RstStock!Part_No <> RstPart!Part_No Then
                                        Exit Do
                                    End If
                                Loop
                            End If
                        End If
                        xMOP_TBQty = mOP_TB_QTY:        xMOP_TPQty = mOP_TP_QTY
                        xMOP_TBVal = mOP_TB_VAL:        xMOP_TPVal = mOP_TP_VAL
                        '**
                        mIss_TB_Qty = 0:                mIss_TB_Val = 0
                        mIss_TP_Qty = 0:                mIss_TP_Val = 0
                        '**
                        mTrf = False
                        
                        If RstStock2.RecordCount > 0 Then
                            RstStock2.MoveFirst
                            RstStock2.FIND ("Part_No='" & RstPart!Part_No & "'")
                            If RstStock2.EOF = False Then
                                Do While RstStock2!Part_No = RstPart!Part_No
                                    mNarr = ""
                                    If RstStock2!StkTrn = "-" Then
                                        If RstStock2!Tax_YN = 1 Then     '' Taxable
                                            mRate = 0
                                            Call X_VAL11(TRec1, RstStock2!Qty_Iss, mRate, mNarr)
                                        Else
                                            mRate = 0
                                            Call X_VAL22(TRec2, RstStock2!Qty_Iss, mRate, mNarr)
                                        End If
                                    ElseIf RstStock2!StkTrn = "+" Then
                                        If RstStock2!Tax_YN = 1 Then     '' Taxable
                                            mOP_TB_QTY = mOP_TB_QTY + RstStock2!Qty_Rec
                                            mOP_TB_VAL = mOP_TB_VAL + (RstStock2!Qty_Rec * VNull(RstStock2!V_Rate))
                                        
                                            mRec_TB_Qty = mRec_TB_Qty + RstStock2!Qty_Rec
                                            mRec_TB_Val = mRec_TB_Val + (RstStock2!Qty_Rec * VNull(RstStock2!V_Rate))
                                        Else
                                            mOP_TP_QTY = mOP_TP_QTY + RstStock2!Qty_Rec
                                            mOP_TP_VAL = mOP_TP_VAL + (RstStock2!Qty_Rec * RstStock2!V_Rate)
                                            mRec_TP_Qty = mRec_TP_Qty + RstStock2!Qty_Rec
                                            mRec_TP_Val = mRec_TP_Val + (RstStock2!Qty_Rec * RstStock2!V_Rate)
                                        End If
                                    End If
                                    RstStock2.MoveNext
                                    If RstStock2.EOF Then
                                        Exit Do
                                    ElseIf RstStock2!Part_No <> RstPart!Part_No Then
                                        Exit Do
                                    End If
                                Loop
                            End If
                        End If
                        If (xMOP_TBQty + mOP_TB_QTY) > 0 Or (xMOP_TPQty + mOP_TP_QTY) > 0 Then
                            If mOP_TB_QTY = 0 Then
                                mOP_TB_VAL = 0
                            ElseIf mOP_TB_QTY < 0 Then
            '                    If mOP_TB_VAL  > 0 Then
            '                        mOP_TB_VAL = -1 * mOP_TB_VAL
            '                    Else
            '                        mOP_TB_VAL = 0
            '                    End If
                            End If
                            If mOP_TP_QTY = 0 Then
                                mOP_TP_VAL = 0
                            ElseIf mOP_TP_QTY < 0 Then
            '                    If mOP_TP_VAL  > 0 Then
            '                        mOP_TP_VAL = -1 * mOP_TP_VAL
            '                    Else
            '                        mOP_TP_VAL = 0
            '                    End If
                            End If
                            RsPart1.Filter = ("Code='" & RstPart!Part_No & "'")
                            RsPart1.MoveFirst
                            Temp06.FIND ("Part_No='" & RstPart!Part_No & "'")
                                If Not Temp06.EOF = True Then
                                    With Temp06
                                        .Fields("Part_No") = RstPart!Part_No
                                        .Fields("Part_Name") = RsPart1!Name
                                        
                                        .Fields("TB_OQty") = .Fields("TB_OQty") + xMOP_TBQty
                                        .Fields("TB_OVal") = .Fields("TB_OVal") + xMOP_TBVal
                                        .Fields("TP_OQty") = .Fields("TP_OQty") + xMOP_TPQty
                                        .Fields("TP_OVal") = .Fields("TP_OVal") + xMOP_TPVal
                                        
                                        .Fields("RE_TB") = .Fields("RE_TB") + mRec_TB_Qty
                                        .Fields("RE_TBV") = .Fields("RE_TBV") + mRec_TB_Val
                                        .Fields("RE_TP") = .Fields("RE_TP") + mRec_TP_Qty
                                        .Fields("RE_TPV") = .Fields("RE_TPV") + mRec_TP_Val
                                        
                                        .Fields("IS_TB") = .Fields("IS_TB") + mIss_TB_Qty
                                        .Fields("IS_TBV") = .Fields("IS_TBV") + mIss_TB_Val
                                        .Fields("IS_TP") = .Fields("IS_TP") + mIss_TP_Qty
                                        .Fields("IS_TPV") = .Fields("IS_TPV") + mIss_TP_Val
                                        
                                        .Fields("TB_BQty") = .Fields("TB_BQty") + mOP_TB_QTY
                                        .Fields("TB_BVal") = .Fields("TB_BVal") + mOP_TB_VAL
                                        .Fields("TP_BQty") = .Fields("TP_BQty") + mOP_TP_QTY
                                        .Fields("TP_BVal") = .Fields("TP_BVal") + mOP_TP_VAL
                                        
                                        .Fields("Net_Qty") = .Fields("Net_Qty") + mOP_TB_QTY + mOP_TP_QTY
                                        .Fields("Net_Val") = .Fields("Net_Val") + mOP_TB_VAL + mOP_TP_VAL
                                        
                                        .Fields("Narr") = mNo
                                        
                                        .Update
                                    End With
                                Else
                                    With Temp06
                                        .AddNew
                                        .Fields("Part_No") = RstPart!Part_No
                                        .Fields("Part_Name") = RsPart1!Name
                                        
                                        .Fields("TB_OQty") = xMOP_TBQty
                                        .Fields("TB_OVal") = xMOP_TBVal
                                        .Fields("TP_OQty") = xMOP_TPQty
                                        .Fields("TP_OVal") = xMOP_TPVal
                                        
                                        .Fields("RE_TB") = mRec_TB_Qty
                                        .Fields("RE_TBV") = mRec_TB_Val
                                        .Fields("RE_TP") = mRec_TP_Qty
                                        .Fields("RE_TPV") = mRec_TP_Val
                                        
                                        .Fields("IS_TB") = mIss_TB_Qty
                                        .Fields("IS_TBV") = mIss_TB_Val
                                        .Fields("IS_TP") = mIss_TP_Qty
                                        .Fields("IS_TPV") = mIss_TP_Val
                                        
                                        .Fields("TB_BQty") = mOP_TB_QTY
                                        .Fields("TB_BVal") = mOP_TB_VAL
                                        .Fields("TP_BQty") = mOP_TP_QTY
                                        .Fields("TP_BVal") = mOP_TP_VAL
                                        
                                        .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
                                        .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
                                        
                                        .Fields("Narr") = mNo
                                        
                                        .Update
                                    End With
                                End If
                        End If
                        RstPart.MoveNext
                    Loop
        If Not RstDiv.EOF = True Then RstDiv.MoveNext
        Next
               
                Set RstRep = Temp06.Clone
                Set TRec1 = Nothing
                Set TRec2 = Nothing
                Set RstStock = Nothing
                Set RstStock2 = Nothing
                If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
                If MDIForm1.Picture1.Visible = True Then MDIForm1.Picture1.Visible = False
                ' For Speed Printing of report
                If SpeedPrnStkRep = True And FGrid.TextMatrix(List3, 1) = "Summary" Then
                    SpeedPrintStkValFIFOSumm
                    Exit Sub
                ElseIf SpeedPrnStkRep = True And FGrid.TextMatrix(List3, 1) = "Detail" Then
                    SpeedPrintStkValFIFODet
                    Exit Sub
                End If
                RepTitle = UCase(Me.CAPTION)
Exit Sub
ELoop:
                Set TRec1 = Nothing
                Set TRec2 = Nothing
                Set RstStock = Nothing
                Set RstStock2 = Nothing
                Set GRs = Nothing
                If err.NUMBER <> 0 Then CheckError
            End Sub


Private Sub SprStkLedValSumCalc1()
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
    RepPrint = True
    GridString1 = Empty: GridString2 = Empty: GridString3 = Empty
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    'If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub


    'Condstr = "where SPStk.V_Date >=" & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SPStk.V_Date <=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " And Left(S.DocId,1) In (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " And " & cMID("S.DocId", "2", "1") & " In (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " And S.Part_No In (" & GridString3 & ")"
            
    If ChkOpStkOnly.Value = 1 Then
        Condstr = Condstr & "And S.V_Type = 'SXAO' "
    End If
                
    'Condstr = Condstr & " And S.Part_No = '269126010122' "

           
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Div_Code", adChar, 1, adFldIsNullable
        .Fields.Append "Part_No", adVarChar, 25, adFldIsNullable
        .Fields.Append "Part_Name", adVarChar, 50, adFldIsNullable
        .Fields.Append "Qty", adDouble, 12, adFldIsNullable
        .Fields.Append "Tax_Yn", adChar, 1, adFldIsNullable
        .Fields.Append "Mrp_Yn", adChar, 1, adFldIsNullable
        .Fields.Append "Rate", adDouble, 12, adFldIsNullable
        .Fields.Append "Amount", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
           
           
           
           
           
    mQry = "SELECT Left(S.DocId,1) As Div_Code, " & cTrim("s.Part_No") & "As Part_No, (Select Part_Name From Part Where Part_No=s.Part_No And Div_Code ='" & PubDivCode & "') As Part_Name, " & _
           "sum(S.Qty_Rec)-Sum(S.Qty_Iss)+Sum(S.Qty_Ret) As mQty, Tax_Yn, " & _
           "Max(V_Rate) As mRate, Max(Amount) As Amount, Max(S.V_Rate) As V_Rate " & _
           "From Sp_Stock S " & _
           "WHERE (S.V_Type= " & cIIF("S.v_Date=" & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "", "'SXAO'") & " Or " & _
           "S.V_Type<> " & cIIF("S.V_Date>= " & ConvertDate(Format(PubStartDate, "dd/MMM/yyyy")) & " And s.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "'SXAO'") & ")" & Condstr & _
           "Group By Left(S.DocId,1), Part_No, Tax_Yn " & _
           "Having (sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret))<>0 "
               
           

           
           
        Set RsPart = New ADODB.Recordset
        If RsPart.State <> 0 Then RsPart.Close
        RsPart.CursorLocation = adUseClient
        RsPart.Open mQry, GCn, adOpenDynamic, adLockBatchOptimistic
                
                
        With RsPart
            If .RecordCount > 0 Then
                Do While Not .EOF
                    Set RsTemp = GCn.Execute("Select Part_No, V_Date, Qty_Rec, V_Rate as Rate " & _
                                          "From Sp_Stock S " & _
                                          "Where S.Part_No='" & !Part_No & "' And S.Tax_Yn = " & !Tax_YN & " " & _
                                          "And Left(S.DocId,1)='" & !Div_Code & "' And Qty_Rec>0  And S.V_Date<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & IIf(ChkOpStkOnly.Value = 1, " And S.V_Type='SXAO' ", "") & "  " & _
                                          "Order by V_Date Desc")
                                                                
                    mQty = 0
                    mReqQty = 0
                    mFifoCost = 0
                    Debug.Print RsTemp.RecordCount
                    Do Until RsTemp.EOF
                        If mQty < VNull(!mQty) Then
                            mReqQty = IIf((mQty + VNull(RsTemp!Qty_Rec)) > VNull(!mQty), VNull(!mQty) - mQty, RsTemp!Qty_Rec)
                            mQty = mQty + VNull(RsTemp!Qty_Rec)
                            
                            mFifoCost = mFifoCost + (mReqQty * VNull(RsTemp!Rate))
                            RsTemp.MoveNext
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    RstRep.AddNew
                    RstRep!Div_Code = RsPart!Div_Code
                    RstRep!Part_No = RsPart!Part_No
                    RstRep!Part_Name = RsPart!Part_Name
                    RstRep!Qty = !mQty
                    RstRep!Tax_YN = RsPart!Tax_YN
                    RstRep!MRP_YN = 1
                    RstRep!Rate = mFifoCost / VNull(!mQty)
                    RstRep!Amount = mFifoCost
                    RstRep.Update
                                        
                    .MoveNext
                Loop
            End If
        End With
    
    
   
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        If MDIForm1.Picture1.Visible = True Then MDIForm1.Picture1.Visible = False
        
        ' For Speed Printing of report
        RepTitle = UCase(Me.CAPTION)
        
        RepName = "StockValuationFifo"

ELoop:
    Set RsTemp = Nothing
    If err.NUMBER <> 0 Then CheckError

End Sub




Private Sub SprStockValuationFIFO()
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
    RepPrint = True
    GridString1 = Empty: GridString2 = Empty: GridString3 = Empty
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub


    'Condstr = "where SPStk.V_Date >=" & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SPStk.V_Date <=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " And Left(S.DocId,1) In (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " And " & cMID("S.DocId", "2", "1") & " In (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " And S.Part_No In (" & GridString3 & ")"
            
           
    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Div_Code", adChar, 1, adFldIsNullable
        .Fields.Append "Part_No", adVarChar, 25, adFldIsNullable
        .Fields.Append "Part_Name", adVarChar, 50, adFldIsNullable
        .Fields.Append "Qty", adDouble, 12, adFldIsNullable
        .Fields.Append "Tax_Yn", adChar, 1, adFldIsNullable
        .Fields.Append "Mrp_Yn", adChar, 1, adFldIsNullable
        .Fields.Append "Rate", adDouble, 12, adFldIsNullable
        .Fields.Append "Amount", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
           
           
    mQry = "SELECT Left(S.DocId,1) As Div_Code, " & cTrim("s.Part_No") & "As Part_No, (Select Part_Name From Part Where Part_No=s.Part_No And Div_Code ='" & PubDivCode & "') As Part_Name, " & _
           "sum(S.Qty_Rec)-Sum(S.Qty_Iss)+Sum(S.Qty_Ret) As mQty, Tax_Yn, Mrp_Yn, " & _
           "Max(V_Rate) As mRate, Max(Amount) As Amount, Max(S.V_Rate) As V_Rate " & _
           "From Sp_Stock S " & _
           "WHERE (S.V_Type= " & cIIF("S.v_Date=" & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & "", "'SXAO'") & " Or " & _
           "S.V_Type<> " & cIIF("S.V_Date>= " & ConvertDate(Format(PubStartDate, "dd/MMM/yyyy")) & " And s.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "", "'SXAO'") & ")" & Condstr & _
           "Group By Left(S.DocId,1), Part_No, Tax_Yn, Mrp_Yn " & _
           "Having (sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret))<>0 "
           
           
           
           
        Set RsPart = New ADODB.Recordset
        If RsPart.State <> 0 Then RsPart.Close
        RsPart.CursorLocation = adUseClient
        RsPart.Open mQry, GCn, adOpenDynamic, adLockBatchOptimistic
                
                
        With RsPart
            If .RecordCount > 0 Then
                Do While Not .EOF
                    Set RsTemp = GCn.Execute("Select Part_No, V_Date, Qty_Rec, V_Rate as Rate " & _
                                          "From Sp_Stock S " & _
                                          "Where S.Part_No='" & !Part_No & "' And S.Tax_Yn = " & !Tax_YN & " And S.Mrp_Yn = " & !MRP_YN & " " & _
                                          "And Left(S.DocId,1)='" & !Div_Code & "' And Qty_Rec>0  And S.V_Date<=" & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & "" & _
                                          "Order by V_Date Desc")
                                                                
                    mQty = 0
                    mReqQty = 0
                    mFifoCost = 0
                    Debug.Print RsTemp.RecordCount
                    Do Until RsTemp.EOF
                        If mQty < VNull(!mQty) Then
                            mReqQty = IIf((mQty + VNull(RsTemp!Qty_Rec)) > VNull(!mQty), VNull(!mQty) - mQty, RsTemp!Qty_Rec)
                            mQty = mQty + VNull(RsTemp!Qty_Rec)
                            
                            mFifoCost = mFifoCost + (mReqQty * VNull(RsTemp!Rate))
                            RsTemp.MoveNext
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    RstRep.AddNew
                    RstRep!Div_Code = RsPart!Div_Code
                    RstRep!Part_No = RsPart!Part_No
                    RstRep!Part_Name = RsPart!Part_Name
                    RstRep!Qty = !mQty
                    RstRep!Tax_YN = RsPart!Tax_YN
                    RstRep!MRP_YN = RsPart!MRP_YN
                    RstRep!Rate = mFifoCost / VNull(!mQty)
                    RstRep!Amount = mFifoCost
                    RstRep.Update
                                        
                    .MoveNext
                Loop
            Else
                !Amount = 0
            End If
        End With
    
    
   
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        If MDIForm1.Picture1.Visible = True Then MDIForm1.Picture1.Visible = False
        
        ' For Speed Printing of report
        RepTitle = UCase(Me.CAPTION)
        
        RepName = "StockLedValueDet"

ELoop:
    Set RsTemp = Nothing
    If err.NUMBER <> 0 Then CheckError

End Sub



Private Sub SprPartProfitCalcJMK()
On Error GoTo ELoop
Dim mQry$, Condstr$, CondStr1$, Condstr2$
Dim mSale As Double, mCost As Double, mAmount As Double, mQty As Double, mProf As Double
Dim XRecNo As Double, xRate As Double, xRate1 As Double
Dim D_Per_MRP_TB1 As Double, D_Per_MRP_TP1 As Double, Gen_Sur_Per1 As Double, D_Per_TP1 As Double, D_Per_TB1 As Double
Dim mNo As Byte, NoUpto As Byte, TRec1Qty As Double

RepPrint = True
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    'If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = "where Sale.V_Date >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Sale.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " and Sale.V_Type in ('SYSIC','SYSIR','W_SIC','W_SIR') And Purpose Not In ('W') "
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and  " & cMID("Sale.DocId", "2", "1") & " in (" & GridString3 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Stk.Part_No in (" & GridString3 & ")"
    
    If PubBackEnd = "A" Then
        mQry = "Select Part.Part_No,Part.Part_Name,Stk.Tax_YN,Sale.V_Date,Sale.V_Type,Stk.Qty_Iss,Stk.Qty_Rec,Stk.Qty_Ret,(Stk.Qty_Iss-Stk.Qty_Ret) as Net_Qty,Stk.Rate, Net_Amt + IIF(Mrp_Yn<>1,Net_Amt*(Tax_Per+Tot_Per)/100,0) as Sale,(Net_Qty*Part.PurRate) as Cost " & _
              "From (Sp_Sale Sale Left Join Sp_Stock Stk On Sale.DocId=Stk.Invoice_DocId) " & _
              "Left Join Part On Part.Part_No=Stk.Part_No " & Condstr & " Order By Part.Part_No"
    Else
        If StrCmp(left(PubComp_Name, 3), "JMK") Then
            mQry = "Select Part.Part_No,Part.Part_Name,Stk.Tax_YN,Sale.V_Date,Sale.V_Type,Stk.Qty_Iss,Stk.Qty_Rec,Stk.Qty_Ret,(Stk.Qty_Iss-Stk.Qty_Ret) as Net_Qty,Stk.Rate, Net_Amt + " & cIIF("Mrp_Yn<>1", "Net_Amt*(Tax_Per+Tot_Per)/100", "0") & " as Sale,((Stk.Qty_Iss-Stk.Qty_Ret)*Part.PurRate) as Cost " & _
                  "From (Sp_Sale Sale Left Join Sp_Stock Stk On Sale.DocId=Stk.Invoice_DocId) " & _
                  "Left Join Part On Part.Part_No=Stk.Part_No " & Condstr & " Order By Part.Part_No"
        Else
            mQry = "Select Part.Part_No,Part.Part_Name,Stk.Tax_YN,Sale.V_Date,Sale.V_Type,Stk.Qty_Iss,Stk.Qty_Rec,Stk.Qty_Ret,(Stk.Qty_Iss-Stk.Qty_Ret) as Net_Qty,Stk.Rate, Net_Amt as Sale,((Stk.Qty_Iss-Stk.Qty_Ret)*Part.NDP) as Cost " & _
                  "From (Sp_Sale Sale Left Join Sp_Stock Stk On Sale.DocId=Stk.Invoice_DocId) " & _
                  "Left Join Part On Part.Part_No=Stk.Part_No " & Condstr & " Order By Part.Part_No"
        End If
    End If

    Set RstRep = GCn.Execute(mQry)

    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List2, 1) = "Detailed" Then
        RepName = "SparePartProfitJMK"
    ElseIf FGrid.TextMatrix(List2, 1) = "Summary" Then
        RepName = "SparePartProfitSummJMK"
    Else
        RepName = "SparePartProfitDateJMK"
    End If
    RepTitle = UCase(Me.CAPTION)
ELoop:
    If err.NUMBER <> 0 Then CheckError
    Set GRs = Nothing
End Sub


Private Sub SprStkLedValSumCalc1UJWAL()
On Error GoTo ELoop
Dim mQry$, Condstr$, CondDivCode$, CondMarkYN$, CondPartNos$, CondPartNos1$, CondDivCode1$
Dim CondStrMRP$, CondPartNosOpStk$, CondPartNosTrn$, Part_Name$, DivStr1$
Dim mRecQty As Double, mIssQty As Double, mStkVal As Double, noofDiv As Integer, d As Integer
Dim XRecNo As Double, DivStr$
Dim mNo As Byte, NoUpto As Byte
Dim RstDiv, TBIss, TPIss As ADODB.Recordset
    RepPrint = True
    GridString1 = Empty: GridString2 = Empty: GridString3 = Empty
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

'   Condstr = " where SPStk.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    Condstr = "where SPStk.V_Date >=" & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and SPStk.V_Date <=" & ConvertDate(FGrid.TextMatrix(Date2, 1))
    Set Temp06 = New ADODB.Recordset
    Set Temp06 = TmpTemp06(Temp06)
    
   
    DivStr1 = GridString1
    Set RstDiv = GCn.Execute("Select Div_Code from Division")
    If Len(GridString1) = 3 Then
        noofDiv = 1
    ElseIf GridString1 = "" Then
        noofDiv = RstDiv.RecordCount
    Else
        noofDiv = 2
    End If
    RstDiv.MoveFirst
    For d = 1 To noofDiv
                If GridString1 = "" Or Len(DivStr1) = 7 Or noofDiv = 2 Then
                    GridString1 = "'" & RstDiv!Div_Code & "'"
                End If
                CondDivCode = " and left(SPStk.DocID,1) in (" & GridString1 & ")"
                CondDivCode1 = " and left(Stk.DocID,1) in (" & GridString1 & ")"
                
                Condstr = Condstr & CondDivCode
                If FGrid.TextMatrix(List1, 1) = "Yes" Then          '' Only for Marked Parts
                    CondMarkYN = " and Mark_YN='Y'"
                Else
                   If Check1(3).Value = Unchecked Then
                        CondPartNos = " and SPStk.Part_No in (" & GridString3 & ")"
                        CondPartNos1 = " Part.Part_No in " & "(" & GridString3 & ")"
                        CondPartNosOpStk = CondPartNos
                    Else
                        CondPartNos = " and SPStk.Part_No in (select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
                        CondPartNos1 = " Part.Part_No in ( select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & CondDivCode1 & ")"
                        CondPartNosOpStk = ""
                    End If
                
                End If
                
                ' Properitary grade check *******
                If Check1(2).Value = Unchecked Then
                    CondPartNos1 = CondPartNos1 & " and Part.Part_Grade in (" & GridString2 & ")"
                End If
                '***************
                
                
                If FGrid.TextMatrix(List2, 1) = "No" Then       '' For MRP_YN
                     CondStrMRP = " and SPStk.MRP_YN=0"
                End If
                
                Condstr = Condstr & CondStrMRP
               
                
                'For RstPart, SQL
                GSQL = "Select Distinct Part.Part_No From Part " & _
                    "where " & CondPartNos1
                If Check1(1).Value = Unchecked Then
                    GSQL = GSQL & IIf(CondPartNos1 = "", "", " and") & " Part.Div_Code in (" & GridString1 & ") "
                    DivStr = "and Part.Div_Code in(" & GridString1 & ")"
                Else
                    DivStr = ""
                End If
                GSQL = GSQL & CondMarkYN & " Order By Part.Part_No"
                
                Set RstPart = GCn.Execute(GSQL)
                    
                    'Stock Valuation FIFO Summary
                    If FGrid.TextMatrix(List3, 1) = "Summary" Then
                        RepName = "StockLedValueSum"
                    ElseIf FGrid.TextMatrix(List3, 1) = "Detail" Then
                        RepName = "StockLedValueDet"
                    End If
                    
                    '********** Taxable Qty Recd.
                    mQry = "select SPStk.Part_No,SPStk.V_DATE,SPStk.Qty_Rec as Qty,SPStk.MRP_YN,SPStk.V_Rate as Rate,SpStk.Qty_Rec*SpStk.V_Rate as RecAmt " & _
                        "From " & _
                        "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
                        Condstr & CondPartNos
                    GSQL = mQry & " and SpStk.Tax_YN=1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
                    'GSQL = mQRY & " and SpStk.Tax_YN=1 and Vt.StkTrn='+' Order By SPStk.V_Date"
                    
                    Set TRec1 = New Recordset
                    With TRec1
                        .CursorLocation = adUseClient
                        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
                    End With
                    
                    '******* Taxpaid Qty Recd.
                    GSQL = mQry & " and SpStk.Tax_YN<>1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
                    'GSQL = mQRY & " and SpStk.Tax_YN<>1 and Vt.StkTrn='+' Order By SPStk.V_Date"
                    Set TRec2 = New Recordset
                    With TRec2
                        .CursorLocation = adUseClient
                        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
                    End With
                    
                    
                    '********** Taxable Qty Issued
                    
                    mQry = "select SPStk.Part_No,SPStk.V_DATE,SPStk.Qty_Iss as Qty,SPStk.MRP_YN,SPStk.V_Rate as Rate,SpStk.Qty_iss*SpStk.V_Rate as IssAmt " & _
                        "From " & _
                        "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
                        Condstr & CondPartNos
                    GSQL = mQry & " and SpStk.Tax_YN=1 and Vt.StkTrn='-' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
                    
                    Set TBIss = New Recordset
                    With TBIss
                        .CursorLocation = adUseClient
                        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
                    End With
                    
                    '******* Taxpaid Qty Issued
                    GSQL = mQry & " and SpStk.Tax_YN<>1 and Vt.StkTrn='-' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
                    Set TPIss = New Recordset
                    With TPIss
                        .CursorLocation = adUseClient
                        .Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
                    End With
                    
                    '******* Taxable + Taxpaid Qty for Opening Loop*************
                    
                    mQry = "select SPStk.V_Type,SPStk.Part_No,SPStk.V_DATE,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate as V_Rate,Vt.StkTrn " & _
                        "From " & _
                        " SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
                         Condstr & CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
                    GSQL = mQry & " and SPStk.V_Type='SXAO' Order By SPStk.Part_No,SPStk.V_Date, " & cMID("SPStk.DocID", "4", "5") & ""
                    
                    Set RstStock = GCn.Execute(GSQL)
                    
'                    ******* Taxable + Taxpaid Qty for With in Date Period Loop

                    GSQL = "select SpStk.V_No,SPStk.V_Type,SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss, " & cIIF("SPStk.V_Type='SXAO'", "SPStk.Rate", "SPStk.V_Rate") & " as V_Rate,Vt.StkTrn,Vt.Description,SubGroup.Name as Party " & _
                        " From " & _
                        " ((SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type) " & _
                        " Left Join Sp_Purch Sp on Sp.DocId=SpStk.DocId) " & _
                        " left Join Subgroup on SubGroup.SubCode=Sp.Party_Code " & _
                        " where (SPStk.V_Date >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " And SPStk.v_date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & " ) " & _
                        CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
                    GSQL = GSQL & " Order By SPStk.Part_No,SPStk.V_Date,SPStk.Tax_YN, " & cMID("SPStk.DocID", "4", "5") & ""

                    Set RstStock2 = GCn.Execute(GSQL)
                    
                    '***********
                    Set tempRst = RstStock.Clone
                    Set TempRst1 = RstStock2.Clone
                    Dim I As Integer
                    MDIForm1.Picture1.Visible = True
                    Do While Not RstPart.EOF
                        'NRA Update
                        MDIForm1.Label1.CAPTION = "Process Status : " & RstPart.AbsolutePosition & "/" & RstPart.RecordCount
                        MDIForm1.Label1.Refresh
                        mVRate = 0: mOPVRate = 0: mDisPer = 0: TempVal = 0
                
                        If FGrid.TextMatrix(List2, 1) = "Yes" Then          '' Only for MRP Parts
                            NoUpto = 1
                            mNo = 1
                        ElseIf FGrid.TextMatrix(List2, 1) = "No" Then          '' Only for Non-MRP Parts
                            NoUpto = 0
                            mNo = 0
                        ElseIf FGrid.TextMatrix(List2, 1) = "All" Then          '' For All Parts
                            NoUpto = 1
                            mNo = 0
                        End If
                        TRec1Qty = 0
                        TRec2Qty = 0
                        
                        mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
                        mIss_TB_Qty = 0: mIss_TB_Val = 0: mIss_TP_Qty = 0: mIss_TP_Val = 0
                        mRec_TB_Qty = 0: mRec_TB_Val = 0: mRec_TP_Qty = 0: mRec_TP_Val = 0
                        
                        TRec1.Filter = ""
                        mOPVRate = 0
                        If TRec1.RecordCount > 0 Then    'Taxable Rect
                            TRec1.MoveFirst
                            TRec1.Filter = ("Part_No='" & RstPart!Part_No & "'")
                            'Nra Update
                            If TRec1.RecordCount > 0 Then
                                TRec1.MoveFirst
                                'If TRec1!Rate <> 0 Then
                                    mOPVRate = TRec1!Rate
                                'Else
                                 '   If TRec1!MRP_YN = 1 Then
                                 '       mOPVRate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                                 '   Else
                                 '       mOPVRate = GCn.Execute("Select TB_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                                 '       mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
                                 '       mOPVRate = mOPVRate - ((mOPVRate * mDisPer) / 100)
                                 '   End If
                                'End If
                            TRec1.MoveFirst
                            End If
            
                            'End Update
                            
                            If TRec1.EOF = False Then
                                TRec1Qty = TRec1!Qty
                            End If
                        End If
                        mVRate = 0
                        TRec2.Filter = ""
                        If TRec2.RecordCount > 0 Then    'Taxpaid Rect
                            TRec2.MoveFirst
                            TRec2.Filter = ("Part_No='" & RstPart!Part_No & "'")
                            'Nra Update
                            If TRec2.RecordCount > 0 Then
                                TRec2.MoveLast
                                'If TRec2!Rate <> 0 Then
                                    mVRate = TRec2!Rate
                                'Else
                                '    If TRec2!MRP_YN = 1 Then
                                '        mVRate = GCn.Execute("Select MRP from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                                '    Else
                                '        mVRate = GCn.Execute("Select TP_SRate from Part where Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(0).Value
                                '        mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstPart!Part_No & "'" & DivStr & "").Fields(1).Value)
                                '        mVRate = mVRate - ((mVRate * mDisPer) / 100)
                                '    End If
                                'End If
                            TRec2.MoveFirst
                            End If
                            
                            'End Update
                            If TRec2.EOF = False Then
                                TRec2Qty = TRec2!Qty
                            End If
                        End If
                        If RstStock.RecordCount > 0 Then
                            RstStock.MoveFirst
                            RstStock.FIND ("Part_No='" & RstPart!Part_No & "'")
                            If RstStock.EOF = False Then
                                Do While RstStock!Part_No = RstPart!Part_No    'Opening Calculation
                                    If RstStock!StkTrn = "-" Then
                                        If RstStock!Tax_YN = 1 Then     '' Taxable
                                            mRate = 0
                                            Call X_VAL11(TRec1, RstStock!Qty_Iss, mRate)
                                        Else
                                            mRate = 0
                                            Call X_VAL22(TRec2, RstStock!Qty_Iss, mRate)
                                        End If
                                        If FGrid.TextMatrix(List3, 1) = "Detail" Then
                                            With Temp06
                                                .AddNew
                                                .Fields("Part_No") = RstPart!Part_No
                                                .Fields("Part_Name") = RsPart1!Name
                                                
                                                
                                                .Fields("TB_OQty") = xMOP_TBQty
                                                .Fields("TB_OVal") = xMOP_TBVal
                                                .Fields("TP_OQty") = xMOP_TPQty
                                                .Fields("TP_OVal") = xMOP_TPVal
                                                
                                                .Fields("RE_TB") = mRec_TB_Qty
                                                .Fields("RE_TBV") = mRec_TB_Val
                                                .Fields("RE_TP") = mRec_TP_Qty
                                                .Fields("RE_TPV") = mRec_TP_Val
                                                
                                                .Fields("IS_TB") = mIss_TB_Qty
                                                .Fields("IS_TBV") = mIss_TB_Val
                                                .Fields("IS_TP") = mIss_TP_Qty
                                                .Fields("IS_TPV") = mIss_TP_Val
                                                
                                                .Fields("TB_BQty") = mOP_TB_QTY
                                                .Fields("TB_BVal") = mOP_TB_VAL
                                                .Fields("TP_BQty") = mOP_TP_QTY
                                                .Fields("TP_BVal") = mOP_TP_VAL
                                                
                                                .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
                                                .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
                                                
                                                .Fields("Narr") = mNo
                                                .Update
                                            End With
                                        End If
                                    ElseIf RstStock!StkTrn = "+" Then
                                        If RstStock!Tax_YN = 1 Then     '' Taxable
                                            mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
                                            mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                                        Else
                                            mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
                                            mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
                                        End If
                                        If FGrid.TextMatrix(List3, 1) = "Detail" Then
                                            With Temp06
                                                .AddNew
                                                .Fields("Part_No") = RstPart!Part_No
                                                .Fields("Part_Name") = RsPart1!Name
                                                .Fields("Party_Name") = RstStock2!Party
                                                
                                                .Fields("TB_OQty") = xMOP_TBQty
                                                .Fields("TB_OVal") = xMOP_TBVal
                                                .Fields("TP_OQty") = xMOP_TPQty
                                                .Fields("TP_OVal") = xMOP_TPVal
                                                
                                                .Fields("RE_TB") = mRec_TB_Qty
                                                .Fields("RE_TBV") = mRec_TB_Val
                                                .Fields("RE_TP") = mRec_TP_Qty
                                                .Fields("RE_TPV") = mRec_TP_Val
                                                
                                                .Fields("IS_TB") = mIss_TB_Qty
                                                .Fields("IS_TBV") = mIss_TB_Val
                                                .Fields("IS_TP") = mIss_TP_Qty
                                                .Fields("IS_TPV") = mIss_TP_Val
                                                
                                                .Fields("TB_BQty") = mOP_TB_QTY
                                                .Fields("TB_BVal") = mOP_TB_VAL
                                                .Fields("TP_BQty") = mOP_TP_QTY
                                                .Fields("TP_BVal") = mOP_TP_VAL
                                                
                                                .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
                                                .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
                                                
                                                .Fields("Narr") = mNo
                                                .Update
                                            End With
                                        End If
                                    End If
                                    RstStock.MoveNext
                                    If RstStock.EOF Then
                                        Exit Do
                                    ElseIf RstStock!Part_No <> RstPart!Part_No Then
                                        Exit Do
                                    End If
                                Loop
                            End If
                        End If
                        xMOP_TBQty = mOP_TB_QTY:        xMOP_TPQty = mOP_TP_QTY
                        xMOP_TBVal = mOP_TB_VAL:        xMOP_TPVal = mOP_TP_VAL
                        '**
                        mIss_TB_Qty = 0:                mIss_TB_Val = 0
                        mIss_TP_Qty = 0:                mIss_TP_Val = 0
                        '**
                        mTrf = False
                        
                        If RstStock2.RecordCount > 0 Then
                            RstStock2.MoveFirst
                            RstStock2.FIND ("Part_No='" & RstPart!Part_No & "'")
                            If RstStock2.EOF = False Then
                                Do While RstStock2!Part_No = RstPart!Part_No
                                    mNarr = ""
                                    If RstStock2!StkTrn = "-" Then
                                        If RstStock2!Tax_YN = 1 Then     '' Taxable
                                            mRate = 0
                                            Call X_VAL11(TRec1, RstStock2!Qty_Iss, mRate, mNarr)
                                        Else
                                            mRate = 0
                                            Call X_VAL22(TRec2, RstStock2!Qty_Iss, mRate, mNarr)
                                        End If
                                        If FGrid.TextMatrix(List3, 1) = "Detail" Then
                                            With Temp06
                                                .AddNew
                                                .Fields("Part_No") = RstPart!Part_No
                                                .Fields("Part_Name") = RsPart1!Name
                                                .Fields("V_type") = RstStock2!V_Type
                                                .Fields("V_No") = RstStock2!V_NO
                                                .Fields("Date") = RstStock2!V_DATE
                                                .Fields("Party_Name") = RstStock2!Party
                                                
                                                .Fields("TB_OQty") = xMOP_TBQty
                                                .Fields("TB_OVal") = xMOP_TBVal
                                                .Fields("TP_OQty") = xMOP_TPQty
                                                .Fields("TP_OVal") = xMOP_TPVal
                                                
                                                .Fields("RE_TB") = mRec_TB_Qty
                                                .Fields("RE_TBV") = mRec_TB_Val
                                                .Fields("RE_TP") = mRec_TP_Qty
                                                .Fields("RE_TPV") = mRec_TP_Val
                                                
                                                .Fields("IS_TB") = mIss_TB_Qty
                                                .Fields("IS_TBV") = mIss_TB_Val
                                                .Fields("IS_TP") = mIss_TP_Qty
                                                .Fields("IS_TPV") = mIss_TP_Val
                                                
                                                .Fields("TB_BQty") = mOP_TB_QTY
                                                .Fields("TB_BVal") = mOP_TB_VAL
                                                .Fields("TP_BQty") = mOP_TP_QTY
                                                .Fields("TP_BVal") = mOP_TP_VAL
                                                
                                                .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
                                                .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
                                                
                                                .Fields("Narr") = mNo
                                                .Update
                                            End With
                                        End If
                                        
                                    ElseIf RstStock2!StkTrn = "+" Then
                                        If RstStock2!Tax_YN = 1 Then     '' Taxable
                                            mOP_TB_QTY = mOP_TB_QTY + RstStock2!Qty_Rec
                                            mOP_TB_VAL = mOP_TB_VAL + (RstStock2!Qty_Rec * RstStock2!V_Rate)

                                            mRec_TB_Qty = mRec_TB_Qty + RstStock2!Qty_Rec
                                            mRec_TB_Val = mRec_TB_Val + (RstStock2!Qty_Rec * RstStock2!V_Rate)
                                        Else
                                            mOP_TP_QTY = mOP_TP_QTY + RstStock2!Qty_Rec
                                            mOP_TP_VAL = mOP_TP_VAL + (RstStock2!Qty_Rec * RstStock2!V_Rate)
                                            mRec_TP_Qty = mRec_TP_Qty + RstStock2!Qty_Rec
                                            mRec_TP_Val = mRec_TP_Val + (RstStock2!Qty_Rec * RstStock2!V_Rate)
                                        End If
                                        If FGrid.TextMatrix(List3, 1) = "Detail" Then
                                            With Temp06
                                                .AddNew
                                                .Fields("Part_No") = RstPart!Part_No
                                                .Fields("Part_Name") = RsPart1!Name
                                                .Fields("V_type") = RstStock2!V_Type
                                                .Fields("V_No") = RstStock2!V_NO
                                                .Fields("Date") = RstStock2!V_DATE
                                                .Fields("Party_Name") = RstStock2!Party
                                                
                                                
                                                .Fields("TB_OQty") = xMOP_TBQty
                                                .Fields("TB_OVal") = xMOP_TBVal
                                                .Fields("TP_OQty") = xMOP_TPQty
                                                .Fields("TP_OVal") = xMOP_TPVal
                                                
                                                .Fields("RE_TB") = mRec_TB_Qty
                                                .Fields("RE_TBV") = mRec_TB_Val
                                                .Fields("RE_TP") = mRec_TP_Qty
                                                .Fields("RE_TPV") = mRec_TP_Val
                                                
                                                .Fields("IS_TB") = mIss_TB_Qty
                                                .Fields("IS_TBV") = mIss_TB_Val
                                                .Fields("IS_TP") = mIss_TP_Qty
                                                .Fields("IS_TPV") = mIss_TP_Val
                                                
                                                .Fields("TB_BQty") = mOP_TB_QTY
                                                .Fields("TB_BVal") = mOP_TB_VAL
                                                .Fields("TP_BQty") = mOP_TP_QTY
                                                .Fields("TP_BVal") = mOP_TP_VAL
                                                
                                                .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
                                                .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
                                                
                                                .Fields("Narr") = mNo
                                                .Update
                                            End With
                                        End If
                                    End If
                                    RstStock2.MoveNext
                                    If RstStock2.EOF Then
                                        Exit Do
                                    ElseIf RstStock2!Part_No <> RstPart!Part_No Then
                                        Exit Do
                                    End If
                                Loop
                            End If
                        End If
                        If (xMOP_TBQty + mOP_TB_QTY) > 0 Or (xMOP_TPQty + mOP_TP_QTY) > 0 Then
                            If mOP_TB_QTY = 0 Then
                                mOP_TB_VAL = 0
                            ElseIf mOP_TB_QTY < 0 Then
            '                    If mOP_TB_VAL  > 0 Then
            '                        mOP_TB_VAL = -1 * mOP_TB_VAL
            '                    Else
            '                        mOP_TB_VAL = 0
            '                    End If
                            End If
                            If mOP_TP_QTY = 0 Then
                                mOP_TP_VAL = 0
                            ElseIf mOP_TP_QTY < 0 Then
            '                    If mOP_TP_VAL  > 0 Then
            '                        mOP_TP_VAL = -1 * mOP_TP_VAL
            '                    Else
            '                        mOP_TP_VAL = 0
            '                    End If
                            End If
                            If FGrid.TextMatrix(List3, 1) = "Summary" Then
                                RsPart1.Filter = ("Code='" & RstPart!Part_No & "'")
                                RsPart1.MoveFirst
                                Temp06.FIND ("Part_No='" & RstPart!Part_No & "'")
                                    
                                    If Not Temp06.EOF = True Then
                                        With Temp06
                                            .Fields("Part_No") = RstPart!Part_No
                                            .Fields("Part_Name") = RsPart1!Name
                                            
                                            .Fields("TB_OQty") = .Fields("TB_OQty") + xMOP_TBQty
                                            .Fields("TB_OVal") = .Fields("TB_OVal") + xMOP_TBVal
                                            .Fields("TP_OQty") = .Fields("TP_OQty") + xMOP_TPQty
                                            .Fields("TP_OVal") = .Fields("TP_OVal") + xMOP_TPVal
                                            
                                            .Fields("RE_TB") = .Fields("RE_TB") + mRec_TB_Qty
                                            .Fields("RE_TBV") = .Fields("RE_TBV") + mRec_TB_Val
                                            .Fields("RE_TP") = .Fields("RE_TP") + mRec_TP_Qty
                                            .Fields("RE_TPV") = .Fields("RE_TPV") + mRec_TP_Val
                                            
                                            .Fields("IS_TB") = .Fields("IS_TB") + mIss_TB_Qty
                                            .Fields("IS_TBV") = .Fields("IS_TBV") + mIss_TB_Val
                                            .Fields("IS_TP") = .Fields("IS_TP") + mIss_TP_Qty
                                            .Fields("IS_TPV") = .Fields("IS_TPV") + mIss_TP_Val
                                            
                                            '.Fields("TB_BQty") = .Fields("TB_BQty") + mOP_TB_QTY
                                            .Fields("TB_BVal") = .Fields("TB_BVal") + mOP_TB_VAL
                                            .Fields("TP_BQty") = .Fields("TP_BQty") + mOP_TP_QTY
                                            .Fields("TP_BVal") = .Fields("TP_BVal") + mOP_TP_VAL
                                            
                                            .Fields("Net_Qty") = .Fields("Net_Qty") + mOP_TB_QTY + mOP_TP_QTY
                                            .Fields("Net_Val") = .Fields("Net_Val") + mOP_TB_VAL + mOP_TP_VAL
                                            
                                            .Fields("Narr") = mNo
                                            
                                            .Update
                                        End With
                                    Else
                                            If TRec1.RecordCount > 0 Then
                                                TRec1.MoveFirst: mRec_TB_Val = 0
                                                Do While TRec1.EOF = False
                                                    mRec_TB_Val = mRec_TB_Val + TRec1!RecAmt
                                                    TRec1.MoveNext
                                                Loop
                                            End If
                                            
                                            If TRec2.RecordCount > 0 Then
                                                TRec2.MoveFirst: mRec_TP_Val = 0
                                                Do While TRec2.EOF = False
                                                    mRec_TP_Val = mRec_TP_Val + TRec2!RecAmt
                                                    TRec2.MoveNext
                                                Loop
                                            End If
                                            
                                        With Temp06
                                            .AddNew
                                            .Fields("Part_No") = RstPart!Part_No
                                            .Fields("Part_Name") = RsPart1!Name
                                            
                                            .Fields("TB_OQty") = xMOP_TBQty
                                            .Fields("TB_OVal") = xMOP_TBVal
                                            .Fields("TP_OQty") = xMOP_TPQty
                                            .Fields("TP_OVal") = xMOP_TPVal
                                            
                                            .Fields("RE_TB") = mRec_TB_Qty
                                            .Fields("RE_TBV") = mRec_TB_Val
                                            .Fields("RE_TP") = mRec_TP_Qty
                                            .Fields("RE_TPV") = mRec_TP_Val
                                            
                                            .Fields("IS_TB") = mIss_TB_Qty
                                            .Fields("IS_TBV") = mIss_TB_Val
                                            .Fields("IS_TP") = mIss_TP_Qty
                                            .Fields("IS_TPV") = mIss_TP_Val
                                            
                                            .Fields("TB_BQty") = mOP_TB_QTY
                                            .Fields("TB_BVal") = mOP_TB_VAL
                                            .Fields("TP_BQty") = mOP_TP_QTY
                                            .Fields("TP_BVal") = mOP_TP_VAL
                                            
                                            .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
                                            .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
                                            
                                            .Fields("Narr") = mNo
                                            
                                            .Update
                                        End With
                                    End If
                            End If
                        End If
                        RstPart.MoveNext
                    Loop
        If Not RstDiv.EOF = True Then RstDiv.MoveNext
        Next
               
                Set RstRep = Temp06.Clone
                Set TRec1 = Nothing
                Set TRec2 = Nothing
                Set RstStock = Nothing
                Set RstStock2 = Nothing
                If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
                If MDIForm1.Picture1.Visible = True Then MDIForm1.Picture1.Visible = False
                ' For Speed Printing of report
                If SpeedPrnStkRep = True And FGrid.TextMatrix(List3, 1) = "Summary" Then
                    SpeedPrintStkValFIFOSumm
                    Exit Sub
                ElseIf SpeedPrnStkRep = True And FGrid.TextMatrix(List3, 1) = "Detail" Then
                    SpeedPrintStkValFIFODet
                    Exit Sub
                End If
                RepTitle = UCase(Me.CAPTION)
ELoop:
                Set TRec1 = Nothing
                Set TRec2 = Nothing
                Set RstStock = Nothing
                Set RstStock2 = Nothing
                Set GRs = Nothing
                If err.NUMBER <> 0 Then CheckError
End Sub
Private Sub DeleteLogProc()
On Error GoTo ELoop
Dim mQry$, Condstr$
RepPrint = True
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub

    Condstr = "where Convert(SmallDateTime,D.VDate) >=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and Convert(SmallDateTime, D.VDate)<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Mid(D.DocID,2,1) in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(D.DocID,1) in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and D.User_Name in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and D.Type in (" & GridString4 & ")"
    
    
    If FGrid.TextMatrix(List1, 1) = "Edited" Then
        Condstr = Condstr & " And EditDate is not Null"
    ElseIf FGrid.TextMatrix(List1, 1) = "Deleted" Then
        Condstr = Condstr & " And Del_Date is not Null"
    End If
    
    mQry = "Select D.DocId, D.Bill_Amt, D.User_Name, D.Del_Date, D.Del_Time, D.Type, D.VDate, " & _
           "D.Total_Item, D.Total_Qty, D.GoodsValue, D.Discount, D.Addition, D.Deduction, " & _
           "D.LabDiscount, D.LabAmount, D.AutoYn, D.EditDate, D.EditTime " & _
           "from DeleteLog D " & Condstr
    Set RstRep = GCn.Execute(mQry)

    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "DeleteLog"
    RepTitle = UCase(Me.CAPTION)
ELoop:
    If err.NUMBER <> 0 Then CheckError
    Set GRs = Nothing
End Sub


