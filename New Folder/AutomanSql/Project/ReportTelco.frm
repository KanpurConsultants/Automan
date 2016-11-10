VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form ReportTelco 
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
   Begin VB.Frame BackFrm 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   435
      TabIndex        =   18
      Top             =   1065
      Visible         =   0   'False
      Width           =   7830
      Begin MSChart20Lib.MSChart Chart1 
         Height          =   1065
         Left            =   6510
         OleObjectBlob   =   "ReportTelco.frx":0000
         TabIndex        =   29
         Top             =   195
         Width           =   660
      End
      Begin VB.TextBox txtGridView 
         BackColor       =   &H008080FF&
         Height          =   285
         Index           =   0
         Left            =   0
         MaxLength       =   50
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   3435
         Top             =   3810
      End
      Begin VB.CommandButton BtnViewExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1605
         TabIndex        =   27
         Top             =   3690
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton BTNPRINT 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   420
         TabIndex        =   26
         Top             =   3675
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.OptionButton ChartType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "3D Step Bar"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   4
         Left            =   1260
         TabIndex        =   25
         Top             =   2865
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.OptionButton ChartType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "3D Bar"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   3
         Left            =   195
         TabIndex        =   24
         Top             =   2910
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.OptionButton ChartType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "3D Area"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   2
         Left            =   2310
         TabIndex        =   23
         Top             =   2595
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.OptionButton ChartType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "2D Pie"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   1
         Left            =   1260
         TabIndex        =   22
         Top             =   2565
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.OptionButton ChartType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "2D Bar"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   2610
         Visible         =   0   'False
         Width           =   810
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Height          =   2310
         Left            =   3240
         TabIndex        =   20
         Top             =   1245
         Visible         =   0   'False
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   4075
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483634
         GridColor       =   -2147483634
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   2370
         Left            =   30
         TabIndex        =   19
         Top             =   45
         Visible         =   0   'False
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   4180
         _Version        =   393216
         BackColor       =   14875388
         ForeColor       =   64
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   16711680
         BackColorSel    =   16761024
         BackColorBkg    =   14875388
         GridColor       =   -2147483640
         GridColorFixed  =   16777215
         GridColorUnpopulated=   16777215
         ScrollTrack     =   -1  'True
         HighLight       =   0
         GridLines       =   0
         GridLinesUnpopulated=   1
         Appearance      =   0
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Print"
      DownPicture     =   "ReportTelco.frx":2356
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
      DownPicture     =   "ReportTelco.frx":5488
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
   Begin VB.TextBox txtGrid 
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
Attribute VB_Name = "ReportTelco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********Reports Created by Nra ********
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
Private Const WksPerRep As Byte = 1
Private Const TimeDevRep As Byte = 2
Private Const CostDevRep As Byte = 3
Private Const ModWCompRep As Byte = 4
Private Const AggreCompRep As Byte = 5
Private Const RepairOrdAnaRep As Byte = 6
Private Const ModWReptComp As Byte = 7
Private Const RepeatJobSumRep As Byte = 10
Private Const RepeatJobRep As Byte = 11
Private Const SrvWiseJob As Byte = 12
Private Const ServAnaDet As Byte = 13
Private Const QuaCompRep As Byte = 14
Private Const DailyCallsCRO As Byte = 16
Private Const DissatisfiedCust As Byte = 17
Private Const CustSatRep As Byte = 18
Private Const WorkProfitRep As Byte = 19

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
'Chart array Declaire
Dim WeekDeviArr(2, 3) As String
Dim ModCompArr() As String
Dim AggreCompArr(2, 6) As String
Dim RepeatJobAnalysisArr(2, 4) As String
Dim WeekDeviSumArr(4, 4) As Double
Dim ComplaintArr(5, 2) As String
Dim QuanComplaintArr(2, 4) As String
Dim Troubletype(2, 7) As String

Dim chrttype As Integer
Dim TotalModel As Double
Dim mListItem As ListItem
Dim ShowChrt As Boolean
Dim VehAtt As Double, NoofRes As Double
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub BTNPRINT_Click(Index As Integer)
'On Error GoTo ERRORHANDLER
SubRep1 = False
RepPrint = True
Select Case GRepFormName
    Case WksPerRep, SrvWiseJob, ServAnaDet, RepairOrdAnaRep, RepeatJobSumRep, DailyCallsCRO, DissatisfiedCust, CustSatRep
        TelcoReportsProc
    Case TimeDevRep, CostDevRep, ModWCompRep, AggreCompRep, ModWReptComp, RepeatJobRep, QuaCompRep
        If Index = 0 Then
            TelcoReportsView ("V")
            Exit Sub
        Else
            TelcoReportsView ("P")
        End If
    Case WorkProfitRep
        If left(FGrid.TextMatrix(List1, 1), 3) <> "All" Then
            WorkProfitRepProc
        Else
            WorkProfitRepProcAll
        End If
End Select
If RepPrint = False Then Exit Sub
    
CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True

If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")
rpt.Database.SetDataSource RstRep
'End Updation
If SubRep1 = True Then
    If GRepFormName = RepeatJobSumRep Then
        rpt.OpenSubreport("SubRep1").Database.SetDataSource RstRep1
    Else
        rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
    End If
End If
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
Private Sub BtnViewExit_Click()
    BackFrm.Visible = False
    
End Sub
Private Sub Chart1_Click()
    If ShowChrt = True Then
        Timer1.Enabled = True
        ShowChrt = False
    Else
        Timer1.Enabled = False
        ShowChrt = True
    End If
End Sub
Private Sub ChartType_Click(Index As Integer)
Select Case GRepFormName
    Case TimeDevRep, CostDevRep
        DispChart Chart1, Index + 1, WeekDeviArr, Chart1.left, (Me.height / 2) - 2200, 4300, 5000
    Case ModWCompRep, ModWReptComp
        DispChart Chart1, Index + 1, ModCompArr, 7260, (Me.height / 2) - 2200, 4300, 5000
    Case AggreCompRep
        DispChart Chart1, Index + 1, AggreCompArr, 7260, (Me.height / 2) - 2200, 4300, 5000
    Case RepeatJobRep
        DispChart Chart1, Index + 1, RepeatJobAnalysisArr, 2000, 3500, 8000, 3000
    Case QuaCompRep
        DispChart Chart1, Index + 1, QuanComplaintArr, 2000, 3000, 8000, 3000
End Select
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

Private Sub Form_Activate()
    If GRepFormName = TimeDevRep Or GRepFormName = CostDevRep Or GRepFormName = ModWCompRep Or GRepFormName = AggreCompRep _
    Or GRepFormName = ModWReptComp Or GRepFormName = RepeatJobRep Then
        BTNPRINT(0).CAPTION = "View"
    End If
        Troubletype(0, 0) = "001": Troubletype(1, 0) = "Engine Related"
        Troubletype(0, 1) = "002": Troubletype(1, 1) = "Gear-Box Related"
        Troubletype(0, 2) = "003": Troubletype(1, 2) = "Suspension/Steering Related"
        Troubletype(0, 3) = "004": Troubletype(1, 3) = "AC Related"
        Troubletype(0, 4) = "005": Troubletype(1, 4) = "Electrical Related"
        Troubletype(0, 5) = "006": Troubletype(1, 5) = "Body Shell Related"
        Troubletype(0, 6) = "007": Troubletype(1, 6) = "Misclleneous"
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


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
Select Case GRepFormName
    Case RepeatJobRep
        If Grid1.Col = 4 Then
            Get_Text Me, Grid1, txtGridView, 0, False, KeyAscii
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Grid1_LeaveCell()
    Grid1.CellBackColor = &HE2FAFC
    Grid1.CellForeColor = &H80000012
End Sub
Private Sub Grid1_RowColChange()
    Grid1.FocusRect = flexFocusNone
    Grid1.CellBackColor = vbBlue
    Grid1.CellForeColor = vbWhite
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
Private Sub Option2_Click()
End Sub

Private Sub Timer1_Timer()
   chrttype = chrttype + 1
   If chrttype = 5 Then
        chrttype = 0
   End If
   ChartType_Click (chrttype)
   ChartType(chrttype).Value = True
End Sub

Private Sub txtGridView_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then txtGridView(0).TEXT = txtGridView(0).Tag: Exit Sub
       If KeyCode = vbKeyReturn Then
            If txtGridView(0).Visible = False Then
                GridTxtDown Grid1, txtGridView, Index, KeyCode, TAddMode, 4, 1, 4
            Else
                Grid1.TextMatrix(Grid1.Row, Grid1.Col) = txtGridView(0).TEXT
                txtGridView(0).Visible = False
                Grid1.CellBackColor = vbBlue
            End If
       End If
Exit Sub
ELoop:
    CheckError
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
            Case ServAnaDet
              ListArray = Array("All", "PDI", "Free Service", "Chargable", "Warranty", "Company Vehicle", "Complementary", "AMC")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 8)
            Case WorkProfitRep
              ListArray = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "All")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 13)
        End Select
    Case List2
        Select Case GRepFormName
            Case ServAnaDet
              ListArray = Array("Yes", "No")
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
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        Else
            ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
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
        Case List1, List2
            If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(TxtGrid(Index), "hh:mm")
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
End Select
Select Case GRepFormName
        Case TimeDevRep, CostDevRep, AggreCompRep, RepairOrdAnaRep, ModWReptComp, RepeatJobRep, RepeatJobSumRep, QuaCompRep, DailyCallsCRO, DissatisfiedCust, CustSatRep
            If FGrid.Row = Date1 Then
                If FGrid.TextMatrix(Date2, 1) <> "" Then
                    If Month(CDate(FGrid.TextMatrix(Date1, 1))) <> Month(CDate(FGrid.TextMatrix(Date2, 1))) Then
                        FGrid.TextMatrix(Date1, 1) = FLDate(FGrid.TextMatrix(Date1, 1), "F")
                    End If
                End If
            Else
                If FGrid.TextMatrix(Date1, 1) <> "" Then
                    If Month(CDate(FGrid.TextMatrix(Date2, 1))) <> Month(CDate(FGrid.TextMatrix(Date1, 1))) Then
                        FGrid.TextMatrix(Date2, 1) = FLDate(CDate(FGrid.TextMatrix(Date1, 1)), "L")
                        
                    End If
                End If
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
Dim I As Integer, Cnt As Integer, GridHeight As Integer

Pic.top = Me.top - Pic.width - 10
BTNPRINT(0).left = (Pic.width - (BTNPRINT(0).width + BTNEXIT.width)) / 2: BTNPRINT(0).top = Pic.top + 10
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
      sitecond = "where site_code='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    
Select Case GRepFormName
    Case WksPerRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Work Time Start": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Work time End": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case TimeDevRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Work Time Start": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Work Time End": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case CostDevRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Job Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Job Date": .RowHeight(Date2) = GridRowHeight
            '.TextMatrix(List1, 0) = "Print Only Chart": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            '.TextMatrix(List1, 1) = "No"
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case ModWCompRep, AggreCompRep, RepairOrdAnaRep, ModWReptComp, RepeatJobRep, RepeatJobSumRep, QuaCompRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Job Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Job Date": .RowHeight(Date2) = GridRowHeight
            '.TextMatrix(List1, 0) = "Print Only Chart": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            '.TextMatrix(List1, 1) = "No"
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
        GridInitialise 2, Grid2Sql
        Grid3Sql = "select '' as O,Model.Model As Model_Description,Model.Model As Code from Model order by Model.Model"
        GridInitialise 3, Grid3Sql
    
    Case SrvWiseJob, DailyCallsCRO, DissatisfiedCust, CustSatRep
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
          
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate

        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
        GridInitialise 2, Grid2Sql
    Case ServAnaDet
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Part Purpose": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Only PartCons.Jobs": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List2, 1) = "Yes"
            
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 3
'       By Rahul At U.N.Automobile Udaipur 11-04-2003
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        
        Grid2Sql = "select '' as O,Serv_Desc as ServiceType,serv_Type  as code from Service_Type order by Serv_desc"
        GridInitialise 2, Grid2Sql
        
        Grid3Sql = "select '' as O,Div_Name as DivName,Div_Code  as code from Division Order by Div_Name"
        GridInitialise 3, Grid3Sql
    Case WorkProfitRep
        With FGrid
            .TextMatrix(List1, 0) = "For Month": .RowHeight(List1) = GridRowHeight
          
            .TextMatrix(List1, 1) = "All"

        End With
        mFirstRow = List1: mLastRow = List1: mHelpGridNo = 2
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Div_Name as DivName,Div_code  as code from Division Where Div_Code='" & PubDivCode & "' order by Div_Name"
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
    Case WksPerRep, RepairOrdAnaRep
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
    Case TimeDevRep, CostDevRep, QuaCompRep
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
         For I = 1 To rpt.OpenSubreport("SUBREP1").FormulaFields.Count
            Select Case UCase(rpt.OpenSubreport("SUBREP1").FormulaFields(I).FormulaFieldName)
                Case UCase("Month")
                    rpt.OpenSubreport("SUBREP1").FormulaFields(I).TEXT = "'" & cMonth(Month(CDate(FGrid.TextMatrix(Date1, 1)))) & "'"
            End Select
        Next
    Case ModWCompRep, AggreCompRep, ModWReptComp, DailyCallsCRO, DissatisfiedCust
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Next
    Case CustSatRep
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("NoofVeh")
                    rpt.FormulaFields(I).TEXT = "" & VehAtt & ""
                Case UCase("NoofRes")
                    rpt.FormulaFields(I).TEXT = "" & NoofRes & ""
            End Select
        Next
    Case RepeatJobRep
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("VehAttended")
                    rpt.FormulaFields(I).TEXT = "'" & Grid2.TextMatrix(0, 1) & "'"
                Case UCase("TotalComp")
                    rpt.FormulaFields(I).TEXT = "'" & Grid2.TextMatrix(1, 1) & "'"
                Case UCase("Month")
                    rpt.FormulaFields(I).TEXT = "'" & Grid2.TextMatrix(0, 3) & "'"
                Case UCase("RepDate")
                    rpt.FormulaFields(I).TEXT = "'" & Grid2.TextMatrix(1, 3) & "'"
            End Select
        Next
    Case RepeatJobSumRep
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("Month1")
                    rpt.FormulaFields(I).TEXT = "'" & cMonth(Month(date)) & "-" & Year(date) & "'"
            End Select
        Next
    Case ServAnaDet
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("DATEBETWEEN")
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
                Case UCase("List1")
                    rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Service'"
            End Select
        Next
       
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub TelcoReportsProc()
'On Error GoTo ELoop
Dim TmpRst As ADODB.Recordset, TmpRst1 As ADODB.Recordset, I As Double, j As Double, K As Double
Dim mQry$, Condstr$, WrkHrs As Integer, TotalComp As Double
Select Case GRepFormName
    Case WksPerRep
            If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
            If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
            If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
            If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
                
            If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
            If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
            
            If Val(FGrid.TextMatrix(List1, 1)) > Val(FGrid.TextMatrix(List2, 1)) Then
                MsgBox "End Time is Less than Start Time", vbInformation
                RepPrint = False
                Exit Sub
            Else
                WrkHrs = Val(FGrid.TextMatrix(List2, 1)) - Val(FGrid.TextMatrix(List1, 1))
            End If
            Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
            
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
            If Check1(1).Value = Checked Then
              If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
            End If
    
            If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
            
            mQry = " SELECT JC.DocId_InvSpr, JC.DocId_InvLab, JC.Job_Date, JC.Job_No, JC.DocId, JC.JobCloseDate,JC.ExpDelDate,JC.GP_No,JC.DelBy,JC.Recp_Time,Left(JC.JobComp_Dt_Time,11) as Comp_Date,right(JC.JobComp_Dt_Time,11) as Comp_Time, " & _
                     " H.Name,JC.Serv_Type, ST.Serv_Desc, 0 as NetLab_Amt, H.RegNo, H.Chassis,H.Model, " & _
                     " JL.War_Lab_Rate,JL.LabourAmt, JL.Major_YN,JL.External_YN,JL.Hrs_Taken as WorkingHrs, 0 as Amount,0 as Net_Amt,0 as Total_Amt,'' as  Purpose,'' as Lub_Category,0 as V_No,0 as MisCharged, " & WrkHrs & " as WorkHrs,EM.Emp_Name,Div.Div_Name " & _
                " FROM ((((Job_Card as JC LEFT JOIN Hiscard as H ON JC.CardNo = H.CardNo)" & _
                     " LEFT JOIN Job_Lab as JL ON JC.DocId = JL.Job_DocID)" & _
                     " LEFT JOIN Emp_Mast as EM ON JC.DelBy = EM.Emp_Code)" & _
                     " LEFT JOIN Division as Div ON Left(JC.DocId,1) = Div.Div_Code)" & _
                     " LEFT JOIN Service_Type as ST ON JC.Serv_Type = ST.Serv_Type " & Condstr & _
                " UNION ALL " & _
                " SELECT JC.DocId_InvSpr, JC.DocId_InvLab, JC.Job_Date, JC.Job_No," & _
                     " JC.DocId, JC.JobCloseDate,JC.ExpDelDate,JC.GP_No,JC.DelBy,JC.Recp_Time,Left(JC.JobComp_Dt_Time,11) as Comp_Date,right(JC.JobComp_Dt_Time,11) as Comp_Time, H.Name,JC.Serv_Type, ST.Serv_Desc, 0 as NetLab_Amt," & _
                     " H.RegNo, H.Chassis,H.Model," & _
                     " 0 AS War_Lab_Rate, 0 AS LabourAmt, 0 AS Major_YN,0 as External_YN,0 as WorkingHrs,((SP_Stock.qty_iss - SP_Stock.qty_ret)*SP_Stock.Rate) as amount, SP_Stock.Net_Amt, 0 as Total_Amt, SP_Stock.Purpose,SP_Stock.Lub_Category, " & cVal("right(JC.Docid_InvSpr,8)") & " as V_No,0 as MisCharged ," & WrkHrs & " as WorkHrs,EM.Emp_Name,Div.Div_Name" & _
                " FROM ((((Job_Card as JC LEFT JOIN Hiscard as H ON JC.CardNo = H.CardNo)" & _
                     " LEFT JOIN Service_Type as ST ON JC.Serv_Type = ST.Serv_Type) LEFT JOIN Emp_Mast as EM ON JC.DelBy = EM.Emp_Code) LEFT JOIN Division as Div ON Left(JC.DocId,1) = Div.Div_Code)" & _
                     " LEFT JOIN SP_Stock ON JC.DocID = SP_Stock.Job_DocId " & Condstr & _
                " Union All " & _
                " SELECT JC.DocId_InvSpr, JC.DocId_InvLab, JC.Job_Date, JC.Job_No," & _
                     " JC.DocId, JC.JobCloseDate,JC.ExpDelDate,JC.GP_No,JC.DelBy,JC.Recp_Time,Left(JC.JobComp_Dt_Time,11) as Comp_Date,right(JC.JobComp_Dt_Time,11) as Comp_Time,H.Name,JC.Serv_Type, ST.Serv_Desc, JC.NetLab_Amt," & _
                     " H.RegNo, H.Chassis,H.Model," & _
                     " 0 AS War_Lab_Rate, 0 AS LabourAmt, 0 AS Major_YN,0 as External_YN,0 as WorkingHrs,0 as amount, 0 as Net_Amt, SP_Sale.Total_Amt, '' as Purpose,'' as Lub_Category,SP_Sale.V_No, " & _
                     "(SP_Sale.Gen_Sur_Amt+SP_Sale.Trans_Amt+SP_Sale.Tax_Amt+SP_Sale.Tax_AmtMRP+SP_Sale.Tax_Sur_Amt+SP_Sale.TaxSur_AmtMRP+SP_Sale.Packing+SP_Sale.TOT_AmtMRP+SP_Sale.ReSalTax_Amt+SP_Sale.Rounded) as MisCharged," & WrkHrs & " as WorkHrs ,EM.Emp_Name,Div.Div_Name" & _
                " FROM ((((Job_Card as JC LEFT JOIN Hiscard as H ON JC.CardNo = H.CardNo) " & _
                     " LEFT JOIN Service_Type as ST ON JC.Serv_Type = ST.Serv_Type) LEFT JOIN Emp_Mast as EM ON JC.DelBy = EM.Emp_Code) LEFT JOIN Division as Div ON Left(JC.DocId,1) = Div.Div_Code)" & _
                     " LEFT JOIN SP_Sale ON JC.DocId_InvSpr = SP_Sale.DocID " & Condstr

            RepName = "WksPerRep"
            
            Set RstRep = New ADODB.Recordset
                With RstRep
                    .Fields.Append "DocId_InvSpr", adChar, 21, adFldIsNullable
                    .Fields.Append "DocId_InvLab", adChar, 21, adFldIsNullable
                    .Fields.Append "Job_Date", adChar, 15, adFldIsNullable
                    .Fields.Append "Job_No", adChar, 10, adFldIsNullable
                    .Fields.Append "DocId", adChar, 21, adFldIsNullable
                    .Fields.Append "JobCloseDate", adDate, 10, adFldIsNullable
                    .Fields.Append "ExpDalDate", adDate, 10, adFldIsNullable
                    .Fields.Append "GP_No", adChar, 20, adFldIsNullable
                    .Fields.Append "Del_By", adChar, 40, adFldIsNullable
                    .Fields.Append "Recp_Time", adChar, 20, adFldIsNullable
                    
                    .Fields.Append "Comp_Date", adDate, 12, adFldIsNullable
                    .Fields.Append "Comp_Time", adChar, 12, adFldIsNullable
                    .Fields.Append "Name", adChar, 40, adFldIsNullable
                    .Fields.Append "Serv_Type", adChar, 7, adFldIsNullable
                    .Fields.Append "Serv_Desc", adChar, 35, adFldIsNullable
                    .Fields.Append "NetLab_Amt", adDouble, 12, adFldIsNullable
                    .Fields.Append "RegNo", adVarChar, 20, adFldIsNullable
                    .Fields.Append "Chassis", adVarChar, 20, adFldIsNullable
                    .Fields.Append "Model", adVarChar, 20, adFldIsNullable
                    .Fields.Append "War_Lab_Rate", adDouble, 12, adFldIsNullable
                    .Fields.Append "LabourAmt", adDouble, 12, adFldIsNullable
                    .Fields.Append "Major_YN", adInteger, 1, adFldIsNullable
                    .Fields.Append "External_YN", adInteger, 1, adFldIsNullable
                    .Fields.Append "Amount", adDouble, 12, adFldIsNullable
                    .Fields.Append "Net_Amt", adDouble, 12, adFldIsNullable
                    .Fields.Append "Total_Amt", adDouble, 12, adFldIsNullable
                    .Fields.Append "Purpose", adVarChar, 1, adFldIsNullable
                    .Fields.Append "Lub_Category", adVarChar, 20, adFldIsNullable
                    .Fields.Append "V_No", adChar, 10, adFldIsNullable
                    .Fields.Append "MisCharged", adDouble, 12, adFldIsNullable
                    .Fields.Append "WorkHrs", adDouble, 2, adFldIsNullable
                    .Fields.Append "TotalHrs", adChar, 20, adFldIsNullable
                    .Fields.Append "WorkingHrs", adDouble, 10, adFldIsNullable
                    .Fields.Append "Div_Name", adChar, 40, adFldIsNullable
                    .CursorLocation = adUseClient
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open
                End With
            Set TmpRst = New ADODB.Recordset
            TmpRst.CursorLocation = adUseClient
            TmpRst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
            For I = 1 To TmpRst.RecordCount
                With RstRep
                    .AddNew
                    !DocId_InvSpr = TmpRst!DocId_InvSpr
                    !DocID_InvLab = TmpRst!DocID_InvLab
                    !Job_Date = TmpRst!Job_Date
                    !Job_No = TmpRst!Job_No
                    !DocID = TmpRst!DocID
                    !JobCloseDate = TmpRst!JobCloseDate
                    !ExpDalDate = TmpRst!ExpDelDate
                    !gp_no = TmpRst!gp_no
                    !Del_By = TmpRst!Emp_Name
                    !Recp_Time = Trim(Format(TmpRst!Recp_Time, "HH:MM"))
                    !Comp_Date = TmpRst!Comp_Date
                    !Comp_Time = TmpRst!Comp_Time
                    !Name = TmpRst!Name
                    !Serv_Type = TmpRst!Serv_Type
                    !Serv_Desc = TmpRst!Serv_Desc
                    !NetLab_Amt = TmpRst!NetLab_Amt
                    !RegNo = TmpRst!RegNo
                    !Chassis = TmpRst!Chassis
                    !Model = TmpRst!Model
                    !War_Lab_Rate = TmpRst!War_Lab_Rate
                    !LabourAmt = TmpRst!LabourAmt
                    !Major_YN = TmpRst!Major_YN
                    !External_yn = TmpRst!External_yn
                    !Amount = TmpRst!Amount
                    !Net_Amt = TmpRst!Net_Amt
                    !Total_Amt = TmpRst!Total_Amt
                    !Purpose = TmpRst!Purpose
                    !Lub_Category = TmpRst!Lub_Category
                    !V_NO = TmpRst!V_NO
                    !MisCharged = TmpRst!MisCharged
                    !WorkHrs = TmpRst!WorkHrs
                    If XNull(TmpRst!Job_Date) <> "" And XNull(TmpRst!Comp_Date) <> "" And XNull(TmpRst!Recp_Time) <> "" And XNull(TmpRst!Comp_Time) <> "" And VNull(TmpRst!WorkHrs) <> 0 Then
                        !TotalHrs = CalcHrs(TmpRst!Job_Date, TmpRst!Comp_Date, TmpRst!Recp_Time, TmpRst!Comp_Time, TmpRst!WorkHrs)
                    End If
                    !WorkingHrs = ConvertHr((GetHr(VNull(TmpRst!WorkingHrs)) * 60) + GetMinuts(VNull(TmpRst!WorkingHrs)))
                    !Div_Name = TmpRst!Div_Name
                    .Update
                End With
                TmpRst.MoveNext
            Next
            If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
            RepTitle = UCase(Me.CAPTION)
            Exit Sub
    Case RepairOrdAnaRep
            If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
            If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
            
            If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
            If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
            Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
            If Check1(1).Value = Checked Then
              If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
            End If
            
            
            If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString3 & ")"
            
            RepName = "RepOrdAnalysis"
            
            Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Aggregate", adChar, 40, adFldIsNullable
                .Fields.Append "Complaint", adChar, 40, adFldIsNullable
                .Fields.Append "NoofComp", adDouble, 10, adFldIsNullable
                .Fields.Append "KM1000", adDouble, 6, adFldIsNullable
                .Fields.Append "KM2000", adDouble, 6, adFldIsNullable
                .Fields.Append "KM3000", adDouble, 6, adFldIsNullable
                .Fields.Append "KM4000", adDouble, 6, adFldIsNullable
                .Fields.Append "KM5000", adDouble, 6, adFldIsNullable
                .Fields.Append "KM6000", adDouble, 6, adFldIsNullable
                .Fields.Append "KM6000Above", adDouble, 6, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            
            For I = 0 To 6
                Set TmpRst = GCn.Execute("SELECT T.Trouble_Name,T.Trouble_Code FROM Trouble T where T.Trelated = '" & Troubletype(0, I) & "' and TType='Complaint'")
                If TmpRst.RecordCount > 0 Then
                    TmpRst.MoveFirst
                    For j = 1 To TmpRst.RecordCount
                        TotalComp = GCn.Execute("SELECT Jd.Details,JC.AtKMsHrs FROM ((Job_Card as JC Left Join HisCard H On JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocId) Left Join Trouble T on JD.Code=T.Trouble_Code  " & Condstr & " and JD.Code = '" & TmpRst!trouble_code & "'").RecordCount
                        If TotalComp > 0 Then
                            With RstRep
                                .AddNew
                                !Aggregate = Trim(Troubletype(1, I))
                                !Complaint = TmpRst!trouble_name
                                !NoofComp = GCn.Execute("SELECT Jd.Details,JC.AtKMsHrs FROM ((Job_Card as JC Left Join HisCard H On JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocId) Left Join Trouble T on JD.Code=T.Trouble_Code  " & Condstr & " and JD.Code = '" & TmpRst!trouble_code & "'").RecordCount
                                
                                Set TmpRst1 = GCn.Execute("SELECT JC.AtKMsHrs FROM ((Job_Card as JC Left Join HisCard H On JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocId) Left Join Trouble T on JD.Code=T.Trouble_Code  " & Condstr & " and JD.Code = '" & TmpRst!trouble_code & "'")
                                    If TmpRst1.RecordCount > 0 Then
                                        TmpRst1.MoveFirst
                                        For K = 1 To TmpRst1.RecordCount
                                            Select Case TmpRst1!AtKMsHrs
                                                Case Is <= 1000
                                                    RstRep!KM1000 = VNull(RstRep!KM1000) + 1
                                                Case Is <= 2000
                                                    RstRep!KM2000 = VNull(RstRep!KM2000) + 1
                                                Case Is <= 3000
                                                    RstRep!KM3000 = VNull(RstRep!KM3000) + 1
                                                Case Is <= 4000
                                                    RstRep!KM4000 = VNull(RstRep!KM4000) + 1
                                                Case Is <= 5000
                                                    RstRep!KM5000 = VNull(RstRep!KM5000) + 1
                                                Case Is <= 6000
                                                    RstRep!KM6000 = VNull(RstRep!KM6000) + 1
                                                Case Is > 6000
                                                    RstRep!KM6000Above = VNull(RstRep!KM6000Above) + 1
                                            End Select
                                        TmpRst1.MoveNext
                                        Next
                                    End If
                                .Update
                            End With
                        End If
                        TmpRst.MoveNext
                    Next
                End If
            Next
            If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
            RepTitle = UCase(Me.CAPTION)
            Exit Sub
   Case SrvWiseJob
            If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
            If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
                
            If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
            If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
            
            
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
            If Check1(1).Value = Checked Then
              If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
            End If
    
            
            If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Job_Card.DocId,1) in (" & GridString2 & ")"
            RepName = "SrvWiseJob"
            
            Set RstRep = New ADODB.Recordset
                With RstRep
                    .Fields.Append "Serv_Type", adChar, 2, adFldIsNullable
                    .Fields.Append "Serv_Desc", adChar, 20, adFldIsNullable
                    .Fields.Append "PrvJob", adDouble, 10, adFldIsNullable
                    .Fields.Append "NewJob", adDouble, 10, adFldIsNullable
                    .Fields.Append "CloseJob", adDouble, 10, adFldIsNullable
                    .Fields.Append "Target", adDouble, 10, adFldIsNullable
                                     
                    .Fields.Append "Lab_Chrg", adDouble, 10, adFldIsNullable
                    .Fields.Append "Lab_Paid", adDouble, 10, adFldIsNullable
                    .Fields.Append "Net_Lab", adDouble, 12, adFldIsNullable
                    .Fields.Append "Spr_Chrg", adDouble, 10, adFldIsNullable
                    .Fields.Append "Spr_Free", adDouble, 10, adFldIsNullable
                    .CursorLocation = adUseClient
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open
                End With
                
            Set TmpRst = GCn.Execute("Select Serv_Type,Serv_Desc,Serv_Target from Service_Type")
            If TmpRst.RecordCount > 0 Then
                For I = 1 To TmpRst.RecordCount
                    Condstr = Condstr & " and Job_Card.Serv_Type ='" & TmpRst!Serv_Type & "'"
                    With RstRep
                        .AddNew
                            !Serv_Type = TmpRst!Serv_Type
                            !Serv_Desc = TmpRst!Serv_Desc
                            !Target = VNull(TmpRst!Serv_Target)
                            
                            !PrvJob = GCn.Execute("Select DocId from Job_Card where ((Job_Date < " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") and ((JobCloseDate  > " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ")  or JobCloseDate is null))" & Condstr).RecordCount
                            !NewJob = GCn.Execute("Select DocId from Job_Card where Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Date <=  " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "" & Condstr).RecordCount
                            !CloseJob = GCn.Execute("Select DocId from Job_Card where  JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "" & Condstr).RecordCount
                            
                            !Lab_Chrg = VNull(GCn.Execute("Select sum(NetLab_Amt)from Job_Card where JobCLOSEDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JobCLOSEDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "" & Condstr).Fields(0).Value)
                            !Lab_Paid = VNull(GCn.Execute("Select sum(Lab_Paid)from Job_Card where JobCLOSEDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JobCLOSEDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "" & Condstr).Fields(0).Value)
                            !Net_Lab = Round(VNull(!Lab_Chrg) + VNull(!Lab_Paid), 2)
                            !Spr_Chrg = VNull(GCn.Execute("Select sum(SP_Stock.Net_Amt)from Job_Card Left Join SP_Stock On Job_Card.DocID=SP_Stock.Job_DocId where JobCLOSEDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JobCLOSEDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and SP_Stock.Purpose = 'C'" & Condstr).Fields(0).Value)
                            !Spr_Free = VNull(GCn.Execute("Select sum(SP_Stock.Net_Amt)from Job_Card Left Join SP_Stock On Job_Card.DocID=SP_Stock.Job_DocId where JobCLOSEDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JobCLOSEDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " and SP_Stock.Purpose = 'F'" & Condstr).Fields(0).Value)
                        .Update
                    End With
                TmpRst.MoveNext
                Next
            End If
            If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
            RepTitle = UCase(Me.CAPTION)
            Exit Sub
Case ServAnaDet
    Dim mQRY1$, mQRY2$, CondStr1$
    Dim mPurpose As String, TotRec As Integer
    Dim MyRst As ADODB.Recordset
    Dim myRst1 As ADODB.Recordset
    Dim myRst2 As ADODB.Recordset
    On Error Resume Next
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
    If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Job_Card.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
    End If
    
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Job_Card.Serv_Type in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " and Left(Job_Card.DocId,1) in (" & GridString3 & ")"
    If FGrid.TextMatrix(List1, 1) <> "All" Then CondStr1 = " and SP_Stock.Purpose = '" & mPurpose & "'"
            
    Set RstRep = New Recordset
    With RstRep
        .Fields.Append "DocId", adVarChar, 21, adFldIsNullable
        .Fields.Append "Job_Date", adVarChar, 16, adFldIsNullable
        .Fields.Append "Job_No", adVarChar, 21, adFldIsNullable
        .Fields.Append "JobCloseDate", adVarChar, 16, adFldIsNullable
        .Fields.Append "CustName", adVarChar, 40, adFldIsNullable
        .Fields.Append "RegNo", adVarChar, 20, adFldIsNullable
        .Fields.Append "Chassis", adVarChar, 20, adFldIsNullable
        .Fields.Append "Engine", adVarChar, 20, adFldIsNullable
        .Fields.Append "AtKMsHrs", adVarChar, 20, adFldIsNullable
        .Fields.Append "Serv_Type", adVarChar, 2, adFldIsNullable
        .Fields.Append "Details", adVarChar, 35, adFldIsNullable
        .Fields.Append "Part_No", adVarChar, 35, adFldIsNullable
        .Fields.Append "Purpose", adVarChar, 2, adFldIsNullable
        .Fields.Append "Qty_Doc", adInteger, 10, adFldIsNullable
        .Fields.Append "Qty_Iss", adInteger, 10, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
            
    Set MyRst = New Recordset
    Set myRst1 = New Recordset
    Set myRst2 = New ADODB.Recordset
            
'Total Job Counting
    mQry = "SELECT Job_Card.DocId,Job_Card.Job_No, Job_Card.JobCloseDate, Job_Card.Job_Date, Job_Card.AtKMsHrs, Job_Card.Serv_Type,HisCard.RegNo, HisCard.Chassis, HisCard.Engine,HisCard.Name" & _
           " from (Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo) "
    mQry = mQry + Condstr
    MyRst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    For I = 1 To MyRst.RecordCount
'Spare Issued on Job Counting
        mQRY1 = "SELECT Job_Card.DocId,SP_Stock.Part_No,SP_Stock.Qty_Doc, SP_Stock.Purpose, SP_Stock.Qty_Rec, SP_Stock.Qty_Iss, SP_Stock.Qty_Ret " & _
           "FROM ((Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo)LEFT JOIN " & _
           "SP_Stock ON sp_stock.job_docid=job_card.docid)"
          
        mQRY1 = mQRY1 & Condstr & CondStr1 & " and Job_Card.DocId= '" & MyRst!DocID & "'"
        myRst1.Open (mQRY1), GCn, adOpenDynamic, adLockOptimistic
'Job Details Counting
        mQRY2 = "SELECT Job_Demand.Details from  Job_Demand WHERE Job_Demand.Job_DocId= '" & MyRst!DocID & "'"
        myRst2.Open (mQRY2), GCn, adOpenDynamic, adLockOptimistic
        
        If myRst1.RecordCount > myRst2.RecordCount Then
            TotRec = myRst1.RecordCount
        Else
            TotRec = myRst2.RecordCount
        End If
        If FGrid.TextMatrix(List2, 1) = "Yes" And myRst1.RecordCount = 1 Then: GoTo xxx
        For j = 1 To TotRec
                With RstRep
                    .AddNew
                    .Fields("DocId") = MyRst!DocID
                    .Fields("Job_No") = MyRst!Job_No
                    .Fields("Job_Date") = Format(MyRst!Job_Date, "DD/MM/YY")
                    .Fields("JobCloseDate") = Format(MyRst!JobCloseDate, "DD/MM/YY")
                    .Fields("CustName") = MyRst!Name
                    .Fields("RegNo") = MyRst!RegNo
                    .Fields("Chassis") = MyRst!Chassis
                    .Fields("Engine") = MyRst!Engine
                    .Fields("AtKMsHrs") = MyRst!AtKMsHrs
                    .Fields("Serv_Type") = MyRst!Serv_Type
                    .Fields("Details") = IIf(myRst2.EOF <> True, myRst2!Details, "")
                    .Fields("Part_No") = myRst1!Part_No
                    .Fields("Purpose") = myRst1!Purpose
                    .Fields("Qty_Doc") = myRst1!Qty_Doc
                    .Fields("Qty_Iss") = myRst1!Qty_Iss - myRst1!Qty_Ret
                    .Update
                    myRst1.MoveNext
                    myRst2.MoveNext
                End With
        Next
xxx:
        MyRst.MoveNext
        myRst1.Close
        myRst2.Close
    Next
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "ServAnaDet"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
Case RepeatJobSumRep
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
        If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString3 & ")"
        
        VehAtt = GCn.Execute("SELECT DocId FROM Job_Card JC Left Join HisCard H On JC.CardNo=H.CardNo " & Condstr & "").RecordCount
        TotalComp = GCn.Execute("SELECT ISNULL(sum(JD.Repeat_YN),0) FROM (Job_Card as JC Left Join Job_Demand JD on JC.DocId=JD.Job_DocID) Left Join HisCard H On JC.CardNo=H.CardNo  " & Condstr & "").Fields(0).Value
        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "VehAtt", adDouble, 6, adFldIsNullable
                .Fields.Append "RepeatComp", adDouble, 6, adFldIsNullable
                .Fields.Append "NoOfFailures", adDouble, 6, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            With RstRep
                .AddNew
                .Fields("VehAtt") = VehAtt
                .Fields("RepeatComp") = TotalComp
                .Fields("NoOfFailures") = GCn.Execute("Select RJ.Job_DocId From (Job_Card JC Left Join RepeatJob RJ On  JC.DocID=RJ.Job_DocID) Left Join HisCard H On JC.CardNo=H.CardNo " & Condstr & " and (Imp_Date1 is not Null or Len(Imp_Date1) >1)").RecordCount
                .Update
            End With
        Set RstRep1 = GCn.Execute("Select Comp_Name,Count(Comp_Name) as FailureNo From (Job_Card JC Left Join RepeatJob RJ On  JC.DocID=RJ.Job_DocID) Left Join HisCard H On JC.CardNo=H.CardNo " & Condstr & " and (Imp_Date1 is not Null or Len(Imp_Date1) >1) Group by Comp_Name")
        SubRep1 = True
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "RepeatJobSum"
        RepTitle = UCase(Me.CAPTION)
        Exit Sub
 Case DailyCallsCRO
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

        Condstr = " where Job_Card.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.JobCLoseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Job_Card.DocId,1) in (" & GridString2 & ")"
        
        GSQL = "SELECT Job_Card.JobCloseDate, Count(Job_Card.DocID) AS NoofJobs, CustFeedback.FeedbackStat, count(CustFeedback.FeedbackStat) AS NoofCalls " & _
             "FROM Job_Card LEFT JOIN CustFeedback ON Job_Card.DocID=CustFeedback.Job_DocId " & Condstr & "" & _
             " GROUP BY Job_Card.JobCloseDate, CustFeedback.FeedbackStat"

        Set RstRep = New ADODB.Recordset
        RstRep.CursorLocation = adUseClient
        RstRep.Open GSQL, GCn, adOpenDynamic, adLockOptimistic

        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "DailyCallsCRO"
        RepTitle = UCase(Me.CAPTION)
        Exit Sub
    Case DissatisfiedCust
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

        Condstr = " where Job_Card.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Job_Card.JobCLoseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("Job_Card.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(Job_Card.DocId,1) in (" & GridString2 & ")"
        
        GSQL = "SELECT Job_Card.JobCloseDate,HisCard.Name,Emp_Mast.Emp_Name,CustFeedback.CompNature,CustFeedback.NxtVisit,CustFeedback.FeedbackStat " & _
             "FROM ((Job_Card LEFT JOIN HisCard on Job_Card.CardNo=HisCard.CardNo)" & _
             " Left Join Emp_Mast ON Job_Card.RecBy_Supervisor=Emp_Mast.Emp_Code)" & _
             " Left Join CustFeedback ON Job_Card.DocID=CustFeedback.Job_DocId " & _
             "" & Condstr & "order by Job_Card.JobCloseDate "
        Set RstRep = New ADODB.Recordset
        RstRep.CursorLocation = adUseClient
        RstRep.Open GSQL, GCn, adOpenDynamic, adLockOptimistic

        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "DissatisfiedCust"
        RepTitle = UCase(Me.CAPTION)
        Exit Sub
    Case CustSatRep
        
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

        Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCLoseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
        
        VehAtt = GCn.Execute("SELECT DocId FROM Job_Card JC Left Join HisCard H On JC.CardNo=H.CardNo" & Condstr & "").RecordCount
        NoofRes = GCn.Execute("SELECT CF.Job_DocId FROM (Job_Card JC Left Join HisCard H On JC.CardNo=H.CardNo) Left Join CustFeedback CF on JC.DocID=CF.Job_DocID " & Condstr & "  and  CF.Job_DocId is not null").RecordCount
        
        GSQL = "Select Sum(CF.Point1) as Parameter1,Sum(CF.Point2) as Parameter2,Sum(CF.Point3) as Parameter3," & _
               "Sum(CF.Point4) as Parameter4,Sum(CF.Point5) as Parameter5,Sum(CF.Point6) as Parameter6,Sum(CF.Point7) as Parameter7,Sum(CF.Point8) as Parameter8 " & _
               " From Job_Card JC Left Join CustFeedback CF on  JC.DocId=CF.Job_DocId " & _
               "" & Condstr & ""
             
        Set RstRep = New ADODB.Recordset
        RstRep.CursorLocation = adUseClient
        RstRep.Open GSQL, GCn, adOpenDynamic, adLockOptimistic

        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "CustSatRep"
        RepTitle = UCase(Me.CAPTION)
        Exit Sub
    End Select
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "RepeatJobSum"
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
Private Sub TelcoReportsView(prnType As String)
Dim TmpRst As ADODB.Recordset, I As Double, j As Double
Dim mQry$, Condstr$, WrkHrs As Integer, PrintRep As Boolean
Dim totalMin As Double, totalmin1 As Double, TotalComp As Double

Select Case GRepFormName
    Case TimeDevRep
        If prnType = "P" Then
            PrintRep = True
        End If
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

        If Val(FGrid.TextMatrix(List1, 1)) > Val(FGrid.TextMatrix(List2, 1)) Then
            MsgBox "End Time is Less than Start Time", vbInformation
            RepPrint = False
            Exit Sub
        Else
            WrkHrs = Val(FGrid.TextMatrix(List2, 1)) - Val(FGrid.TextMatrix(List1, 1))
        End If
        Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"

        mQry = " SELECT Jc.DocId,JC.Job_Date, JC.Job_No,JC.JobCloseDate,Left(JC.ExpDelDate,11) as ExpDelDate,Right(JC.ExpDelDate,11) as ExpDelTime,JC.GP_No,JC.DelBy,JC.Recp_Time,Left(JC.JobComp_Dt_Time,11) as Comp_Date,right(JC.JobComp_Dt_Time,11) as Comp_Time,JC.Serv_Type," & WrkHrs & " as WorkHrs" & _
               " FROM Job_Card as JC " & Condstr

        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Job_Date", adChar, 15, adFldIsNullable
                .Fields.Append "Job_No", adChar, 10, adFldIsNullable
                .Fields.Append "SrvAdv", adChar, 20, adFldIsNullable
                .Fields.Append "DocId", adChar, 21, adFldIsNullable
                .Fields.Append "JobCloseDate", adDate, 10, adFldIsNullable
                .Fields.Append "ExpDalDate", adDate, 10, adFldIsNullable
                .Fields.Append "ExpDalTime", adChar, 12, adFldIsNullable
                .Fields.Append "GP_No", adChar, 20, adFldIsNullable
                .Fields.Append "Del_By", adChar, 40, adFldIsNullable
                .Fields.Append "Recp_Time", adChar, 20, adFldIsNullable

                .Fields.Append "Comp_Date", adDate, 12, adFldIsNullable
                .Fields.Append "Comp_Time", adChar, 12, adFldIsNullable
                .Fields.Append "Serv_Type", adChar, 7, adFldIsNullable
                .Fields.Append "CommitHrs", adChar, 10, adFldIsNullable
                .Fields.Append "TotalHrs", adChar, 10, adFldIsNullable
                .Fields.Append "DeviHrs", adDouble, 10, adFldIsNullable
                .Fields.Append "DeviPer", adDouble, 10, adFldIsNullable
                .Fields.Append "WorkHrs", adDouble, 2, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
        Set TmpRst = New ADODB.Recordset
        TmpRst.CursorLocation = adUseClient
        TmpRst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        For I = 1 To TmpRst.RecordCount
            With RstRep
                .AddNew
                !Job_Date = TmpRst!Job_Date
                !Job_No = TmpRst!Job_No
                !SrvAdv = "          "
                !DocID = TmpRst!DocID
                !JobCloseDate = TmpRst!JobCloseDate
                !ExpDalDate = TmpRst!ExpDelDate
                !ExpDalTime = Format(TmpRst!ExpDelTime, "hh:mm")
                !gp_no = TmpRst!gp_no
                !Recp_Time = Format(TmpRst!Recp_Time, "hh:mm")
                !Comp_Date = TmpRst!Comp_Date
                !Comp_Time = Format(TmpRst!Comp_Time, "hh:mm")
                !Serv_Type = TmpRst!Serv_Type
                !WorkHrs = TmpRst!WorkHrs
                If XNull(TmpRst!Job_Date) <> "" And XNull(TmpRst!ExpDelDate) <> "" And XNull(TmpRst!Recp_Time) <> "" And XNull(TmpRst!ExpDelTime) <> "" And VNull(TmpRst!WorkHrs) <> 0 Then
                    !CommitHrs = CalcHrs(TmpRst!Job_Date, TmpRst!ExpDelDate, TmpRst!Recp_Time, TmpRst!ExpDelTime, TmpRst!WorkHrs)
                End If
                If XNull(TmpRst!Job_Date) <> "" And XNull(TmpRst!Comp_Date) <> "" And XNull(TmpRst!Recp_Time) <> "" And XNull(TmpRst!Comp_Time) <> "" And VNull(TmpRst!WorkHrs) <> 0 Then
                    !TotalHrs = CalcHrs(TmpRst!Job_Date, TmpRst!Comp_Date, TmpRst!Recp_Time, TmpRst!Comp_Time, TmpRst!WorkHrs)
                End If
                totalMin = (Int(Val(XNull(!TotalHrs))) * 60) + GetMinuts(Trim(XNull(!TotalHrs)))
                totalmin1 = (Int(Val(XNull(!CommitHrs))) * 60) + GetMinuts(Trim(XNull(!CommitHrs)))
                !DeviHrs = ConvertHr(totalMin - totalmin1)
                If Val(!CommitHrs) = 0 Then
                    !DeviPer = 100
                Else
                    !DeviPer = Round(Val((!DeviHrs) * 100) / Val(!CommitHrs), 2)
                End If
                .Update
            End With
            TmpRst.MoveNext
        Next
        
        ' Week wise deviation
        If RstRep.RecordCount = 0 Then MsgBox "****** No Data to View ******": RepPrint = False: Exit Sub
        RstRep.Sort = "JobCloseDate"
        RstRep.MoveFirst
        For I = 1 To RstRep.RecordCount
                If Day(RstRep!JobCloseDate) <= 7 Then
                    WeekDeviSumArr(0, 0) = WeekDeviSumArr(0, 0) + Val(Trim(RstRep!CommitHrs))
                    WeekDeviSumArr(0, 1) = WeekDeviSumArr(0, 1) + Val(Trim(RstRep!DeviHrs))
                    WeekDeviArr(1, 0) = Round(IIf(WeekDeviSumArr(0, 0) = 0, 0, (WeekDeviSumArr(0, 1) / WeekDeviSumArr(0, 0))), 2)
                ElseIf Day(RstRep!JobCloseDate) > 7 And Day(RstRep!JobCloseDate) <= 14 Then
                    WeekDeviSumArr(1, 0) = WeekDeviSumArr(1, 0) + Val(Trim(RstRep!CommitHrs))
                    WeekDeviSumArr(1, 1) = WeekDeviSumArr(1, 1) + Val(Trim(RstRep!DeviHrs))
                    WeekDeviArr(1, 1) = Round(IIf(WeekDeviSumArr(1, 0) = 0, 0, (WeekDeviSumArr(1, 1) / WeekDeviSumArr(1, 0))), 2)
                ElseIf Day(RstRep!JobCloseDate) > 14 And Day(RstRep!JobCloseDate) <= 21 Then
                    WeekDeviSumArr(2, 0) = WeekDeviSumArr(2, 0) + Val(Trim(RstRep!CommitHrs))
                    WeekDeviSumArr(2, 1) = WeekDeviSumArr(2, 1) + Val(Trim(RstRep!DeviHrs))
                    WeekDeviArr(1, 2) = Round(IIf(WeekDeviSumArr(2, 0) = 0, 0, (WeekDeviSumArr(2, 1) / WeekDeviSumArr(2, 0))), 2)
                Else
                    WeekDeviSumArr(3, 0) = WeekDeviSumArr(3, 0) + Val(Trim(RstRep!CommitHrs))
                    WeekDeviSumArr(3, 1) = WeekDeviSumArr(3, 1) + Val(Trim(RstRep!DeviHrs))
                    WeekDeviArr(1, 3) = Round(IIf(WeekDeviSumArr(3, 0) = 0, 0, (WeekDeviSumArr(3, 1) / WeekDeviSumArr(3, 0))), 2)
                End If
        RstRep!CommitHrs = Replace(RstRep!CommitHrs, ".", ":")
        RstRep.MoveNext
        Next
        With Grid2
            .Visible = True: .left = 7260: .top = 100: .height = 950
            .Rows = 4: .Cols = 5: .FontFixed.Bold = True: .FixedRows = 3
            .FixedCols = 1
            .ColWidth(0) = 1000: .ColWidth(1) = 1000: .ColWidth(2) = 800: .ColWidth(3) = 800
            .TextMatrix(0, 0) = "Frequency": .TextMatrix(0, 1) = ": Weekly"
            .TextMatrix(1, 0) = "Month": .TextMatrix(1, 1) = " : " & cMonth(Month(CDate(FGrid.TextMatrix(Date1, 1))))
            .TextMatrix(2, 1) = "Week 1"
            .TextMatrix(3, 1) = WeekDeviArr(1, 0)
            .TextMatrix(2, 2) = "Week 2"
            .TextMatrix(3, 2) = WeekDeviArr(1, 1)
            .TextMatrix(2, 3) = "Week 3"
            .TextMatrix(3, 3) = WeekDeviArr(1, 2)
            .TextMatrix(2, 4) = "Week 4"
            .TextMatrix(3, 4) = WeekDeviArr(1, 3)
            .TextMatrix(3, 0) = "% Devi."
        
            WeekDeviArr(0, 0) = .TextMatrix(2, 1)
            WeekDeviArr(0, 1) = .TextMatrix(2, 2)
            WeekDeviArr(0, 2) = .TextMatrix(2, 3)
            WeekDeviArr(0, 3) = .TextMatrix(2, 4)
        End With
        
        DispChart Chart1, 2, WeekDeviArr, 7260, (Me.height / 2) - 2200, 4300, 5000
        RstRep.MoveFirst
        ' For report Printing
        If PrintRep Then
            RepTitle = UCase(Me.CAPTION)
            RepName = "TimeDevRep"
            'For Graph and Weekly Deviation Printing
            Set RstRep1 = New ADODB.Recordset
                With RstRep1
                    .Fields.Append "Week1", adDouble, 10, adFldIsNullable
                    .Fields.Append "Week2", adDouble, 10, adFldIsNullable
                    .Fields.Append "Week3", adDouble, 10, adFldIsNullable
                    .Fields.Append "Week4", adDouble, 10, adFldIsNullable
                    .CursorLocation = adUseClient
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open
                End With
            With RstRep1
                 .AddNew
                 !Week1 = Val(WeekDeviArr(1, 0))
                 !Week2 = Val(WeekDeviArr(1, 1))
                 !Week3 = Val(WeekDeviArr(1, 2))
                 !Week4 = Val(WeekDeviArr(1, 3))
                 .Update
            End With
            SubRep1 = True
            Exit Sub
        End If
    Case CostDevRep
        If prnType = "P" Then
            PrintRep = True
        End If
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

        Condstr = " where JC.JobCloseDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.JobCloseDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
        
        Condstr = Condstr & "  Group by Jc.DocId,JC.Job_Date, JC.Job_No,JC.JobCloseDate,JC.Est_SpCost,JC.Est_LabCost,JC.NetLab_Amt Order by JC.DocID"

        mQry = " SELECT Jc.DocId,JC.Job_Date, JC.Job_No,JC.JobCloseDate,JC.Est_SpCost,JC.Est_LabCost,JC.NetLab_Amt,Sum(SP.Total_Amt) as Total_Amt" & _
               " FROM Job_Card as JC Left Join SP_Sale SP on JC.DocId=SP.Job_DocId " & Condstr
            

        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Job_Date", adChar, 15, adFldIsNullable
                .Fields.Append "Job_No", adChar, 10, adFldIsNullable
                .Fields.Append "DocId", adChar, 21, adFldIsNullable
                .Fields.Append "SrvAdv", adChar, 20, adFldIsNullable
                .Fields.Append "JobCloseDate", adDate, 10, adFldIsNullable
                .Fields.Append "EstCost", adDouble, 10, adFldIsNullable
                .Fields.Append "Net_Amt", adDouble, 10, adFldIsNullable
                .Fields.Append "DeviCost", adDouble, 10, adFldIsNullable
                .Fields.Append "DeviPer", adDouble, 10, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
        Set TmpRst = New ADODB.Recordset
        TmpRst.CursorLocation = adUseClient
        TmpRst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        For I = 1 To TmpRst.RecordCount
            With RstRep
                .AddNew
                !Job_Date = TmpRst!Job_Date
                !Job_No = TmpRst!Job_No
                !SrvAdv = "          "
                !DocID = TmpRst!DocID
                !JobCloseDate = TmpRst!JobCloseDate
                !EstCost = VNull(TmpRst!Est_SpCost) + VNull(TmpRst!Est_LabCost)
                !Net_Amt = VNull(TmpRst!Total_Amt) + VNull(TmpRst!NetLab_Amt)
                !DeviCost = VNull(!Net_Amt) - VNull(!EstCost)
                If Val(!EstCost) = 0 Then
                    !DeviPer = 100
                Else
                    !DeviPer = Round((VNull(!Net_Amt) - VNull(!EstCost)) / VNull(!EstCost) * 100, 2)
                End If
                .Update
            End With
            TmpRst.MoveNext
        Next
        ' Week wise deviation
        If RstRep.RecordCount = 0 Then MsgBox "****** No Data to View ******": RepPrint = False: Exit Sub
        RstRep.Sort = "JobCloseDate"
        RstRep.MoveFirst
        For I = 1 To RstRep.RecordCount
                If Day(RstRep!JobCloseDate) <= 7 Then
                    WeekDeviSumArr(0, 0) = WeekDeviSumArr(0, 0) + Val(Trim(RstRep!EstCost))
                    WeekDeviSumArr(0, 1) = WeekDeviSumArr(0, 1) + Val(Trim(RstRep!DeviCost))
                    If WeekDeviSumArr(0, 0) <> 0 Then
                        WeekDeviArr(1, 0) = Round(WeekDeviSumArr(0, 1) / WeekDeviSumArr(0, 0), 2)
                    End If
                ElseIf Day(RstRep!JobCloseDate) > 7 And Day(RstRep!JobCloseDate) <= 14 Then
                    WeekDeviSumArr(1, 0) = WeekDeviSumArr(1, 0) + Val(Trim(RstRep!EstCost))
                    WeekDeviSumArr(1, 1) = WeekDeviSumArr(1, 1) + Val(Trim(RstRep!DeviCost))
                    If WeekDeviSumArr(1, 0) <> 0 Then
                        WeekDeviArr(1, 1) = Round(WeekDeviSumArr(1, 1) / WeekDeviSumArr(1, 0), 2)
                    End If
                ElseIf Day(RstRep!JobCloseDate) > 14 And Day(RstRep!JobCloseDate) <= 21 Then
                    WeekDeviSumArr(2, 0) = WeekDeviSumArr(2, 0) + Val(Trim(RstRep!EstCost))
                    WeekDeviSumArr(2, 1) = WeekDeviSumArr(2, 1) + Val(Trim(RstRep!DeviCost))
                    If WeekDeviSumArr(2, 0) <> 0 Then
                        WeekDeviArr(1, 2) = Round(WeekDeviSumArr(2, 1) / WeekDeviSumArr(2, 0), 2)
                    End If
                Else
                    WeekDeviSumArr(3, 0) = WeekDeviSumArr(3, 0) + Val(Trim(RstRep!EstCost))
                    WeekDeviSumArr(3, 1) = WeekDeviSumArr(3, 1) + Val(Trim(RstRep!DeviCost))
                    If WeekDeviSumArr(1, 0) <> 0 Then
                        WeekDeviArr(1, 3) = Round(WeekDeviSumArr(3, 1) / WeekDeviSumArr(3, 0), 2)
                    End If
                End If
        RstRep.MoveNext
        Next
        With Grid2
            .Visible = True: .left = 7260: .top = 100
            .height = 950: .Rows = 4: .Cols = 5
            .FontFixed.Bold = True: .FixedRows = 3
            .FixedCols = 1
            .ColWidth(0) = 1000: .ColWidth(1) = 1000: .ColWidth(2) = 800: .ColWidth(3) = 800
            .TextMatrix(0, 0) = "Frequency": .TextMatrix(0, 1) = ": Weekly"
            .TextMatrix(1, 0) = "Month": .TextMatrix(1, 1) = " : " & cMonth(Month(CDate(FGrid.TextMatrix(Date1, 1))))
            .TextMatrix(2, 1) = "Week 1"
            .TextMatrix(3, 1) = WeekDeviArr(1, 0)
            .TextMatrix(2, 2) = "Week 2"
            .TextMatrix(3, 2) = WeekDeviArr(1, 1)
            .TextMatrix(2, 3) = "Week 3"
            .TextMatrix(3, 3) = WeekDeviArr(1, 2)
            .TextMatrix(2, 4) = "Week 4"
            .TextMatrix(3, 4) = WeekDeviArr(1, 3)
            .TextMatrix(3, 0) = "% Devi."
        
            WeekDeviArr(0, 0) = .TextMatrix(2, 1)
            WeekDeviArr(0, 1) = .TextMatrix(2, 2)
            WeekDeviArr(0, 2) = .TextMatrix(2, 3)
            WeekDeviArr(0, 3) = .TextMatrix(2, 4)
        End With
        
        DispChart Chart1, 2, WeekDeviArr, 7260, (Me.height / 2) - 2200, 4300, 5000
        RstRep.MoveFirst
        ' For report Printing
        If PrintRep Then
            RepTitle = UCase(Me.CAPTION)
            RepName = "CostDevRep"
            'For Graph and Weekly Deviation Printing
            Set RstRep1 = New ADODB.Recordset
                With RstRep1
                    .Fields.Append "Week1", adDouble, 10, adFldIsNullable
                    .Fields.Append "Week2", adDouble, 10, adFldIsNullable
                    .Fields.Append "Week3", adDouble, 10, adFldIsNullable
                    .Fields.Append "Week4", adDouble, 10, adFldIsNullable
                    .CursorLocation = adUseClient
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open
                End With
            With RstRep1
                 .AddNew
                 !Week1 = Val(WeekDeviArr(1, 0))
                 !Week2 = Val(WeekDeviArr(1, 1))
                 !Week3 = Val(WeekDeviArr(1, 2))
                 !Week4 = Val(WeekDeviArr(1, 3))
                 .Update
            End With
            SubRep1 = True
            Exit Sub
        End If
   Case ModWCompRep
        If prnType = "P" Then
            PrintRep = True
        End If
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
        If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString3 & ")"
        mQry = " SELECT H.Model,sum(JD.Complaint_YN) as NoofComp" & _
               " FROM (Job_Card as JC Left Join HisCard H  on JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocID  " & Condstr & " Group by H.Model"
            

        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Model", adChar, 21, adFldIsNullable
                .Fields.Append "NoofComp", adDouble, 10, adFldIsNullable
                .Fields.Append "PerofComp", adDouble, 10, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
        
        Set TmpRst = New ADODB.Recordset
        TmpRst.CursorLocation = adUseClient
        TmpRst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
        
        For I = 1 To TmpRst.RecordCount
            TotalComp = TotalComp + VNull(TmpRst!NoofComp)
            TmpRst.MoveNext
        Next
        
        If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
        ReDim ModCompArr(2, TmpRst.RecordCount)
        TotalModel = TmpRst.RecordCount
        For I = 1 To TmpRst.RecordCount
            With RstRep
                .AddNew
                !Model = Trim(TmpRst!Model)
                !NoofComp = VNull(TmpRst!NoofComp)
                If TotalComp > 0 Then
                    !PerOfComp = Round((VNull(TmpRst!NoofComp) * 100) / TotalComp, 2)
                Else
                    !PerOfComp = 0
                End If
                .Update
                'Chart Array Filling
                ModCompArr(0, I - 1) = Trim(!Model)
                ModCompArr(1, I - 1) = !PerOfComp
            End With
            TmpRst.MoveNext
        Next
        ' Week wise deviation
        If RstRep.RecordCount = 0 Then MsgBox "****** No Data to View ******": RepPrint = False: Exit Sub
        
        DispChart Chart1, 2, ModCompArr, 7260, (Me.height / 2) - 2200, 4300, 5000
        RstRep.MoveFirst
        ' For report Printing
        If PrintRep = True Then
            RepTitle = UCase(Me.CAPTION)
            RepName = "ModWCompRep"
            
            Set RstRep1 = New ADODB.Recordset
            With RstRep1
            For I = 0 To TotalModel - 1
                .Fields.Append "" & ModCompArr(0, I) & "", adDouble, 10, adFldIsNullable
            Next
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            RstRep1.AddNew
            For I = 0 To TotalModel - 1
                With RstRep1
                     .Fields("" & ModCompArr(0, I) & "") = Val(Trim(ModCompArr(1, I)))
                     .Update
                End With
            Next
            SubRep1 = True
            Exit Sub
        End If
    Case AggreCompRep
        If prnType = "P" Then
            PrintRep = True
        End If
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
        If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString3 & ")"
            
            Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Aggregate", adChar, 40, adFldIsNullable
                .Fields.Append "NoofComp", adDouble, 10, adFldIsNullable
                .Fields.Append "PerofComp", adDouble, 10, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            TotalComp = GCn.Execute("SELECT T.TRelated FROM ((Job_Card as JC Left Join HisCard H On JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocId) Left Join Trouble T on JD.Code=T.Trouble_Code  " & Condstr & " and T.Trelated in ('001','002','003','004','005','006','007')").RecordCount
            For I = 0 To 6
                With RstRep
                    .AddNew
                    !Aggregate = Trim(Troubletype(1, I))
                    !NoofComp = GCn.Execute("SELECT T.TRelated FROM ((Job_Card as JC Left Join HisCard H On JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocId) Left Join Trouble T on JD.Code=T.Trouble_Code  " & Condstr & " and T.Trelated = '" & Troubletype(0, I) & "'").RecordCount
                    If TotalComp > 0 Then
                        !PerOfComp = Round((VNull(!NoofComp) * 100) / TotalComp, 2)
                    Else
                        !PerOfComp = 0
                    End If
                    .Update
                    'Chart Array Filling
                    AggreCompArr(0, I) = Troubletype(1, I)
                    AggreCompArr(1, I) = !PerOfComp
                End With
            Next
        If RstRep.RecordCount = 0 Then MsgBox "****** No Data to View ******": RepPrint = False: Exit Sub
        
        DispChart Chart1, 2, AggreCompArr, 7260, (Me.height / 2) - 2200, 4300, 5000
        RstRep.MoveFirst
        ' For report Printing
        If PrintRep = True Then
            RepTitle = UCase(Me.CAPTION)
            RepName = "AggreCompRep"
            
            Set RstRep1 = New ADODB.Recordset
            With RstRep1
            For I = 0 To 6
                .Fields.Append "" & AggreCompArr(0, I) & "", adDouble, 10, adFldIsNullable
            Next
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            RstRep1.AddNew
            For I = 0 To 6
                With RstRep1
                     .Fields("" & AggreCompArr(0, I) & "") = Val(Trim(AggreCompArr(1, I)))
                     .Update
                End With
            Next
            SubRep1 = True
            Exit Sub
        End If
Case ModWReptComp
        If prnType = "P" Then
            PrintRep = True
        End If
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
        If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & "  ='" & PubSiteCode & "' "
        End If
    
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString3 & ")"
        mQry = " SELECT H.Model,sum(JD.Repeat_YN) as NoofComp" & _
               " FROM (Job_Card as JC Left Join HisCard H  on JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocID  " & Condstr & " Group by H.Model"
            

        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Model", adChar, 21, adFldIsNullable
                .Fields.Append "NoofComp", adDouble, 10, adFldIsNullable
                .Fields.Append "PerofComp", adDouble, 10, adFldIsNullable
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
        
        Set TmpRst = New ADODB.Recordset
        TmpRst.CursorLocation = adUseClient
        TmpRst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
        
        For I = 1 To TmpRst.RecordCount
            TotalComp = TotalComp + VNull(TmpRst!NoofComp)
            TmpRst.MoveNext
        Next
        
        If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
        ReDim ModCompArr(2, TmpRst.RecordCount)
        TotalModel = TmpRst.RecordCount
        For I = 1 To TmpRst.RecordCount
            With RstRep
                .AddNew
                !Model = Trim(TmpRst!Model)
                !NoofComp = VNull(TmpRst!NoofComp)
                If TotalComp > 0 Then
                    !PerOfComp = Round((VNull(TmpRst!NoofComp) * 100) / TotalComp, 2)
                Else
                    !PerOfComp = 0
                End If
                .Update
                'Chart Array Filling
                ModCompArr(0, I - 1) = Trim(!Model)
                ModCompArr(1, I - 1) = !PerOfComp
            End With
            TmpRst.MoveNext
        Next
        ' Week wise deviation
        If RstRep.RecordCount = 0 Then MsgBox "****** No Data to View ******": RepPrint = False: Exit Sub
        
        DispChart Chart1, 2, ModCompArr, 7260, (Me.height / 2) - 2200, 4300, 5000
        RstRep.MoveFirst
        ' For report Printing
        If PrintRep Then
            RepTitle = UCase(Me.CAPTION)
            RepName = "ModWReptComp"
            
            Set RstRep1 = New ADODB.Recordset
            With RstRep1
            For I = 0 To TotalModel - 1
                .Fields.Append "" & ModCompArr(0, I) & "", adDouble, 10, adFldIsNullable
            Next
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
            RstRep1.AddNew
            For I = 0 To TotalModel - 1
                With RstRep1
                     .Fields("" & ModCompArr(0, I) & "") = Val(Trim(ModCompArr(1, I)))
                     .Update
                End With
            Next
            SubRep1 = True
            Exit Sub
        End If
  Case RepeatJobRep
        Dim VehAtt As Double
        If prnType = "P" Then
            PrintRep = True
        End If
        If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
        If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
        
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
        If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub

        Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""

        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
        End If
        
        
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString3 & ")"
        
'        mQRY = " SELECT count(H.Model) as NoofVehicle,sum(JD.Repeat_YN) as NoofComp" & _
'               " FROM (Job_Card as JC Left Join HisCard H  on JC.CardNo=H.CardNo) Left Join Job_Demand JD on JC.DocId=JD.Job_DocID  " & Condstr & " Group by H.Model"
        
        VehAtt = GCn.Execute("SELECT DocId FROM Job_Card JC Left Join HisCard H On JC.CardNo=H.CardNo" & Condstr & "").RecordCount
        TotalComp = GCn.Execute("SELECT isnull(sum(JD.Repeat_YN),0) FROM (Job_Card as JC Left Join Job_Demand JD on JC.DocId=JD.Job_DocID) Left Join HisCard H On JC.CardNo=H.CardNo  " & Condstr & "").Fields(0).Value
        
        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "SrNo", adDouble, 2, adFldIsNullable
                .Fields.Append "CompReason", adChar, 200, adFldIsNullable
                .Fields.Append "NoOfComplaints", adDouble, 6, adFldIsNullable
                .Fields.Append "PerOfVehAtt", adDouble, 6, adFldIsNullable
                .Fields.Append "ActionTaken", adChar, 50, adFldIsNullable
                
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
        'Fill Complaint array.
        ComplaintArr(0, 0) = "During the previous visit the defect was not diagnosed properly by floor superwiser  / service advisor. "
        ComplaintArr(1, 0) = "Defect was not diagnosed properly as road test was not done."
        ComplaintArr(2, 0) = "Technician does not know the correct procedure for doing the job."
        ComplaintArr(3, 0) = "Technician did not do the proper job due to negligence / casualness."
        ComplaintArr(4, 0) = "Defect could not be attended as the part was not available"
        ComplaintArr(0, 1) = "RJ.Imp_Date2"
        ComplaintArr(1, 1) = "RJ.Imp_Date3"
        ComplaintArr(2, 1) = "RJ.Imp_Date4"
        ComplaintArr(3, 1) = "RJ.Imp_Date5"
        ComplaintArr(4, 1) = "RJ.Imp_Date6"
        For I = 0 To 4
            With RstRep
                .AddNew
                .Fields("SrNo") = I + 1
                .Fields("CompReason") = ComplaintArr(I, 0)
                .Fields("NoOfComplaints") = GCn.Execute("Select RJ.Job_DocId From (Job_Card JC Left Join RepeatJob RJ On  JC.DocID=RJ.Job_DocID) Left Join HisCard H On JC.CardNo=H.CardNo " & Condstr & " and (" & ComplaintArr(I, 1) & " is not Null or Len(" & ComplaintArr(I, 1) & ") >1) ").RecordCount
                .Fields("PerOfVehAtt") = Round((Val(.Fields("NoOfComplaints")) * 100) / VehAtt, 0)
                .Fields("ActionTaken") = ""
                .Update
                'Fill chart array
                RepeatJobAnalysisArr(0, I) = "Reason" & I + 1
                RepeatJobAnalysisArr(1, I) = .Fields("PerOfVehAtt")
            End With
        Next
        With Grid2
            .Visible = True: .width = 7350: .height = 550
            .top = 100: .left = (Me.width / 2) - (.width / 2): .Rows = 2
            .GridLines = flexGridRaised: .GridColor = vbBlack: .Font.Bold = True
            .ForeColor = vbBlue: .Cols = 4: .ColAlignment(3) = vbRightJustify
            .FixedRows = 0
            .FixedCols = 0
            .ColWidth(0) = 3000: .ColWidth(1) = 1000: .ColWidth(2) = 2000: .ColWidth(3) = 1300
            .TextMatrix(0, 0) = "No of Veh Attended ": .TextMatrix(0, 1) = VehAtt: .TextMatrix(0, 2) = "Month": .TextMatrix(0, 3) = cMonth(Month(CDate(FGrid.TextMatrix(Date1, 1))))
            .TextMatrix(1, 0) = "No of Repeat Complaints": .TextMatrix(1, 1) = TotalComp: .TextMatrix(1, 2) = "Date of Reports": .TextMatrix(1, 3) = CStr(PubLoginDate)
        End With
        DispChart Chart1, 2, RepeatJobAnalysisArr, 2000, 3500, 8000, 3000
        RstRep.MoveFirst
        ' For report Printing
        If PrintRep Then
            RepTitle = UCase(Me.CAPTION)
            RepName = "RepeatJobAnalysis"
            'For Graph and Weekly Deviation Printing
            Set RstRep1 = New ADODB.Recordset
                With RstRep1
                    For I = 0 To 4
                        .Fields.Append "" & RepeatJobAnalysisArr(0, I) & "", adDouble, 10, adFldIsNullable
                    Next
                    .CursorLocation = adUseClient
                    .CursorType = adOpenDynamic
                    .LockType = adLockOptimistic
                    .Open
                End With
            With RstRep1
                 .AddNew
                 For I = 0 To 4
                    .Fields("" & RepeatJobAnalysisArr(0, I) & "") = Val(Trim(RepeatJobAnalysisArr(1, I)))
                    .Update
                 Next
            End With
            With RstRep
                 .MoveFirst
                 For I = 0 To 4
                    .Fields("ActionTaken") = Trim(Grid1.TextMatrix(I + 2, 4))
                    .Update
                    .MoveNext
                 Next
            End With
            SubRep1 = True
            Exit Sub
        End If
  Case QuaCompRep
            Dim Complaint1 As Double, Complaint2 As Double, Complaint3 As Double, Complaint4 As Double, TotalJobs As Double
            If prnType = "P" Then
                PrintRep = True
            End If
            If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
            If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
            
            If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
            If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
            Condstr = " where JC.Job_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and JC.Job_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
            If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " in (" & GridString1 & ")"
            If Check1(1).Value = Checked Then
              If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and " & cMID("JC.DocId", "3", "1") & " ='" & PubSiteCode & "' "
            End If
    
            
            If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Left(JC.DocId,1) in (" & GridString2 & ")"
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and H.Model in (" & GridString3 & ")"
            
            mQry = "SELECT JC.DocId, Count(JD.Job_DocID) As TotalJob FROM (Job_Card AS JC LEFT JOIN Job_Demand AS JD ON JC.DocId=JD.Job_DocID) Left Join HisCard H On JC.CardNo=H.CardNo " & Condstr & " GROUP BY JD.Job_DocID, JC.DocID "
            Set TmpRst = New ADODB.Recordset
            Set TmpRst = New ADODB.Recordset
            TmpRst.CursorLocation = adUseClient
            TmpRst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
            TotalJobs = GCn.Execute(mQry).RecordCount
            If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
            
            Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "SrNo", adChar, 2, adFldIsNullable
                .Fields.Append "Complain", adChar, 30, adFldIsNullable
                .Fields.Append "NoofJobs", adDouble, 6, adFldIsNullable
                .Fields.Append "AnalysisinPer", adDouble, 6, adFldIsNullable
                
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
        'Fill Complaint array.
            QuanComplaintArr(0, 0) = "Job With One Complain"
            QuanComplaintArr(0, 1) = "Job With Two Complains"
            QuanComplaintArr(0, 2) = "Job With Three Complains"
            QuanComplaintArr(0, 3) = "Job With Complains  >4"
            
            For I = 0 To TmpRst.RecordCount - 1
                Select Case TmpRst!TotalJob
                    Case 1
                        Complaint1 = Complaint1 + 1
                    Case 2
                        Complaint2 = Complaint2 + 1
                    Case 3
                        Complaint3 = Complaint3 + 1
                    Case Is >= 4
                        Complaint4 = Complaint4 + 1
                End Select
                TmpRst.MoveNext
            Next
            For I = 0 To 3
                With RstRep
                    .AddNew
                    .Fields("SrNo") = I + 1
                    .Fields("Complain") = Trim(QuanComplaintArr(0, I))
                    Select Case I
                    Case 0
                        .Fields("NoofJobs") = Complaint1
                        .Fields("AnalysisinPer") = Round(((Complaint1 * 100) / TotalJobs), 0)
                        QuanComplaintArr(1, I) = .Fields("AnalysisinPer")
                    Case 1
                        .Fields("NoofJobs") = Complaint2
                        .Fields("AnalysisinPer") = Round(((Complaint2 * 100) / TotalJobs), 0)
                        QuanComplaintArr(1, I) = .Fields("AnalysisinPer")
                    Case 2
                        .Fields("NoofJobs") = Complaint3
                        .Fields("AnalysisinPer") = Round(((Complaint3 * 100) / TotalJobs), 0)
                        QuanComplaintArr(1, I) = .Fields("AnalysisinPer")
                    Case 3
                        .Fields("NoofJobs") = Complaint4
                        .Fields("AnalysisinPer") = Round(((Complaint4 * 100) / TotalJobs), 0)
                        QuanComplaintArr(1, I) = .Fields("AnalysisinPer")
                    End Select
                    .Update
           
                End With
            Next
            DispChart Chart1, 2, QuanComplaintArr, 2000, 3000, 8000, 3500
            RstRep.MoveFirst
            ' For report Printing
             If PrintRep Then
                RepTitle = UCase(Me.CAPTION)
                RepName = "QuaCompAnalysis"
                Set RstRep1 = New ADODB.Recordset
                    With RstRep1
                        For I = 0 To 3
                            .Fields.Append "" & QuanComplaintArr(0, I) & "", adDouble, 10, adFldIsNullable
                        Next
                        .CursorLocation = adUseClient
                        .CursorType = adOpenDynamic
                        .LockType = adLockOptimistic
                        .Open
                    End With
                With RstRep1
                    .AddNew
                     For I = 0 To 3
                        .Fields("" & QuanComplaintArr(0, I) & "") = Val(Trim(QuanComplaintArr(1, I)))
                        .Update
                     Next
                End With
                SubRep1 = True
                Exit Sub
            End If
    End Select
        For I = 0 To 4
            ChartType(I).Visible = True
            If I = 0 Then
                ChartType(0).left = Chart1.left - 80
                ChartType(0).top = Chart1.top - 500
            Else
                ChartType(I).left = ChartType(I - 1).left + ChartType(I - 1).width
                ChartType(I).top = ChartType(I - 1).top
            End If
        Next
        ChartType(1).Value = True
        BackFrm.Visible = True
        BackFrm.top = Me.top
        BackFrm.left = Me.left
        BackFrm.width = Me.width
        BackFrm.height = Me.height
        
        ini_ViewGrid
        Timer1.Enabled = True

ELoop:
    RepPrint = False
'    MsgBox err.Description
End Sub
Private Sub ini_ViewGrid()
  Dim HeadArr(2, 9) As String
  Dim WidthArr(9) As Double
  Dim Fldarr(9) As String
Select Case GRepFormName
    Case TimeDevRep
        
        HeadArr(0, 0) = "JobCard": HeadArr(1, 0) = "No."
        HeadArr(0, 1) = "Open Date" ': HeadArr(1, 1) = "Advisor"
        HeadArr(0, 2) = "JobCard": HeadArr(1, 2) = "Open Time"
        HeadArr(0, 3) = "Comtd.Time": HeadArr(1, 3) = "of Delivery"
        HeadArr(0, 4) = "Commited": HeadArr(1, 4) = "ManHrs"
        HeadArr(0, 5) = "Vehicle": HeadArr(1, 5) = "Deliv. Date"
        HeadArr(0, 6) = "Vehicle": HeadArr(1, 6) = "Deliv.Time"
        HeadArr(0, 7) = "Deviation from": HeadArr(1, 7) = "Comtd Time ManHrs"
        HeadArr(0, 8) = "Deviation as %": HeadArr(1, 8) = "of Commited"
        
        WidthArr(0) = 1100: WidthArr(1) = 1500: WidthArr(2) = 1100: WidthArr(3) = 1100: WidthArr(4) = 1100
        WidthArr(5) = 1500: WidthArr(6) = 1200: WidthArr(7) = 1700: WidthArr(8) = 1500
        
        Fldarr(0) = "Job_No": Fldarr(1) = "Job_Date": Fldarr(2) = "Recp_Time": Fldarr(3) = "ExpDalTime": Fldarr(4) = "CommitHrs": Fldarr(5) = "JobCloseDate"
        Fldarr(6) = "Comp_Time": Fldarr(7) = "DeviHrs": Fldarr(8) = "DeviPer"
        
        ViewGrid Grid1, Me.left + 100, Me.top + 100, (Me.width / 2) + 1100, Me.height - 600, RstRep, HeadArr, WidthArr, Fldarr, 2, 9
    Case CostDevRep
        HeadArr(0, 0) = "JobCard": HeadArr(1, 0) = "No."
        HeadArr(0, 1) = "Initial of Srv": HeadArr(1, 1) = "Advisor"
        HeadArr(0, 2) = "Est. Cost": HeadArr(1, 2) = "Sprs+Lab"
        HeadArr(0, 3) = "Actual Cost": HeadArr(1, 3) = "Sprs+Lab"
        HeadArr(0, 4) = "Deviation %": HeadArr(1, 4) = "of Cost"
        
        WidthArr(0) = 1100: WidthArr(1) = 1200: WidthArr(2) = 1100: WidthArr(3) = 1100: WidthArr(4) = 1100
        Fldarr(0) = "Job_No": Fldarr(1) = "SrvAdv": Fldarr(2) = "EstCost": Fldarr(3) = "Net_Amt": Fldarr(4) = "DeviPer"
        
        ViewGrid Grid1, Me.left + 100, Me.top + 100, (Me.width / 2) + 1100, Me.height - 600, RstRep, HeadArr, WidthArr, Fldarr, 2, 5
    Case ModWCompRep, ModWReptComp
        HeadArr(0, 0) = "Model"
        HeadArr(0, 1) = "Number of": HeadArr(1, 1) = "Complaints"
        HeadArr(0, 2) = "Analysis": HeadArr(1, 2) = "(in %)"
        
        WidthArr(0) = 3000: WidthArr(1) = 1000: WidthArr(2) = 1100
        Fldarr(0) = "Model": Fldarr(1) = "NoofComp": Fldarr(2) = "PerofComp"
        
        ViewGrid Grid1, Me.left + 100, Me.top + 100, (Me.width / 2) + 1100, Me.height - 600, RstRep, HeadArr, WidthArr, Fldarr, 2, 3
    Case AggreCompRep
        HeadArr(0, 0) = "Aggregate"
        HeadArr(0, 1) = "Number of": HeadArr(1, 1) = "Complaints"
        HeadArr(0, 2) = "Analysis": HeadArr(1, 2) = "(in %)"
        
        WidthArr(0) = 3000: WidthArr(1) = 1000: WidthArr(2) = 1100
        Fldarr(0) = "Aggregate": Fldarr(1) = "NoofComp": Fldarr(2) = "PerofComp"
        
        ViewGrid Grid1, Me.left + 100, Me.top + 100, (Me.width / 2) + 1100, Me.height - 600, RstRep, HeadArr, WidthArr, Fldarr, 2, 3
        
    Case RepeatJobRep, QuaCompRep
        If GRepFormName = RepeatJobRep Then
            HeadArr(0, 0) = "Sr.": HeadArr(1, 0) = "No."
            HeadArr(0, 1) = "Reason"
            HeadArr(0, 2) = "No. of": HeadArr(1, 2) = "Complaints"
            HeadArr(0, 3) = "Per of Total": HeadArr(1, 3) = "Veh. Attended"
            HeadArr(0, 4) = "Action Taken"
            
            WidthArr(0) = 400: WidthArr(1) = 7000: WidthArr(2) = 1000: WidthArr(3) = 1100: WidthArr(4) = 2000
            Fldarr(0) = "SrNo": Fldarr(1) = "CompReason": Fldarr(2) = "NoOfComplaints": Fldarr(3) = "PerOfVehAtt": Fldarr(4) = "ActionTaken"
            
            ViewGrid Grid1, Me.left + 100, (Grid2.top + Grid2.height) + 200, (Me.width) - 350, (Me.height / 4) + 200, RstRep, HeadArr, WidthArr, Fldarr, 2, 5
        ElseIf GRepFormName = QuaCompRep Then
            HeadArr(0, 0) = "Sr.": HeadArr(1, 0) = "No."
            HeadArr(0, 1) = "Description"
            HeadArr(0, 2) = "No. of": HeadArr(1, 2) = "Jobs"
            HeadArr(0, 3) = "Analysis": HeadArr(1, 3) = "(in %)"
            
            WidthArr(0) = 400: WidthArr(1) = 5400: WidthArr(2) = 1000: WidthArr(3) = 1100
            Fldarr(0) = "SrNo": Fldarr(1) = "Complain": Fldarr(2) = "NoofJobs": Fldarr(3) = "AnalysisinPer"
            
            ViewGrid Grid1, 2000, Me.top + 200, 8000, (Me.height / 4) + 200, RstRep, HeadArr, WidthArr, Fldarr, 2, 4
        End If
        
        With BTNPRINT(1)
            .Visible = True
            .top = Me.top + Me.height - 1000
            .left = Me.left + Me.width - 3000
        End With
        With BtnViewExit
            .Visible = True
            .top = BTNPRINT(1).top
            .left = BTNPRINT(1).left + BTNPRINT(1).width + 100
        End With
        Exit Sub
End Select
        With BTNPRINT(1)
            .Visible = True
            .top = Chart1.top + Chart1.height + 100
            .left = Chart1.left + 1000
        End With
        With BtnViewExit
            .Visible = True
            .top = BTNPRINT(1).top
            .left = BTNPRINT(1).left + BTNPRINT(1).width + 100
        End With
End Sub
Private Sub WorkProfitRepProc()
On Error GoTo ELoop
Dim TmpRst As ADODB.Recordset, TmpRst1 As ADODB.Recordset, I As Double, j As Double, K As Double
Dim mQry$, Condstr$, WrkHrs As Integer, TotalComp As Double
Dim FirstDate, LastDate As Date
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Month", adVarChar, 15, adFldIsNullable
                .Fields.Append "BayNos", adDouble, 2, adFldIsNullable
                .Fields.Append "WorkHrsPerDay", adDouble, 2, adFldIsNullable
                .Fields.Append "DaysPerWeek", adDouble, 2, adFldIsNullable
                .Fields.Append "TotAtms", adDouble, 2, adFldIsNullable
                .Fields.Append "Labour", adDouble, 12, adFldIsNullable
                .Fields.Append "SparesWork", adDouble, 12, adFldIsNullable
                .Fields.Append "SparesCou", adDouble, 12, adFldIsNullable
                .Fields.Append "FreeService", adDouble, 12, adFldIsNullable
                .Fields.Append "Warranty", adDouble, 12, adFldIsNullable
                .Fields.Append "PaidService", adDouble, 12, adFldIsNullable
                .Fields.Append "PaidRepairs", adDouble, 12, adFldIsNullable
                .Fields.Append "WorkRunCost", adDouble, 12, adFldIsNullable
                
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
                   
             Select Case left(FGrid.TextMatrix(List1, 1), 3)
                Case "Jan"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=1")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Feb"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=2")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Mar"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=3")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Apr"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=4")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= #31/Apr/" & Year(PubStartDate) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "May"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=5")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/May/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Jun"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=6")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= #01/Jun/" & Year(PubStartDate) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Jul"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=7")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Aug"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=8")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Sep"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=9")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Oct"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=10")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Nov"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=11")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
                Case "Dec"
                    Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=12")
                    If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of " & FGrid.TextMatrix(List1, 1) & " ": RepPrint = False: Exit Sub
                    
                    With RstRep
                        .AddNew
                        .Fields("Month") = FGrid.TextMatrix(List1, 1): .Fields("BayNos") = VNull(TmpRst!TotalBays)
                        .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                        .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                        .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & "").Fields(0).Value)
                        .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                        .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                        .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                        .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                        .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                        .Update
                    End With
            End Select
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "WorkProfitReg"
        RepTitle = UCase(Me.CAPTION)
        Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub WorkProfitRepProcAll()
On Error GoTo ELoop
Dim TmpRst As ADODB.Recordset, TmpRst1 As ADODB.Recordset, I As Double, j As Double, K As Double
Dim mQry$, Condstr$, WrkHrs As Integer, TotalComp As Double
Dim FirstDate, LastDate As Date
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
        Set RstRep = New ADODB.Recordset
            With RstRep
                .Fields.Append "Month", adVarChar, 15, adFldIsNullable
                .Fields.Append "BayNos", adDouble, 2, adFldIsNullable
                .Fields.Append "WorkHrsPerDay", adDouble, 2, adFldIsNullable
                .Fields.Append "DaysPerWeek", adDouble, 2, adFldIsNullable
                .Fields.Append "TotAtms", adDouble, 2, adFldIsNullable
                .Fields.Append "Labour", adDouble, 12, adFldIsNullable
                .Fields.Append "SparesWork", adDouble, 12, adFldIsNullable
                .Fields.Append "SparesCou", adDouble, 12, adFldIsNullable
                .Fields.Append "FreeService", adDouble, 12, adFldIsNullable
                .Fields.Append "Warranty", adDouble, 12, adFldIsNullable
                .Fields.Append "PaidService", adDouble, 12, adFldIsNullable
                .Fields.Append "PaidRepairs", adDouble, 12, adFldIsNullable
                .Fields.Append "WorkRunCost", adDouble, 12, adFldIsNullable
                
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
            End With
                   
             Select Case left(FGrid.TextMatrix(List1, 1), 3)
                Case "All"
                    'Jan
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 1 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=1")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "January"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of January": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of January": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of January": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of January": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of January": RepPrint = False: Exit Sub
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jan/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jan/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                      End If
                    End If
                    'Feb
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 2 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=2")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "February"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of February": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of February": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of February": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of February": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of February": RepPrint = False: Exit Sub
                            'Feb
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Feb/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Feb/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Mar
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 3 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=3")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "March"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of March": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of March": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of March": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of March": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of March": RepPrint = False: Exit Sub
                            'Mar
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Mar/" & Year(PubEndDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Mar/" & Year(PubEndDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Apr
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 4 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=4")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "April"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of April": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of April": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of April": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of April": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of April": RepPrint = False: Exit Sub
                            'Apr
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Apr/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Apr/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'May
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 5 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=5")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "May"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of May": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of May": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of May": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of May": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of May": RepPrint = False: Exit Sub
                            'May
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/May/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/May/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/May/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Jun
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 6 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=6")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "June"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of June": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of June": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of June": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of June": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of June": RepPrint = False: Exit Sub
                            'Jun
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("01/Jun/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Jul
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 7 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=7")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "July"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of July": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of July": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of July": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of July": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of July": RepPrint = False: Exit Sub
                            'Jul
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Jul/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Jul/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Aug
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 8 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=8")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "August"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of Auguest": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of Auguest": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of Auguest": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of Auguest": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of Auguest": RepPrint = False: Exit Sub
                            'Aug
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Aug/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Aug/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Sep
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 9 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=9")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "September"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of September": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of September": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of September": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of September": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of September": RepPrint = False: Exit Sub
                            'Sep
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID)  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Sep/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Sep/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Oct
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 10 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=10")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "October"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of October ": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of October": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of October": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of October": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of October": RepPrint = False: Exit Sub
                            'oct
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Oct/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Oct/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Nov
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) > 11 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=11")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "November"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of November": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of November": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of November": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of November": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of November": RepPrint = False: Exit Sub
                            'Nov
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >=" & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Nov/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Nov/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
                    'Dec
                    If Month(GCn.Execute("Select Max(V_Date) from SP_Sale").Fields(0).Value) = 12 Then
                        Set TmpRst = GCn.Execute("Select Bays_Wash+Bays_Serv+Bays_Repair+Bays_AcRepair as TotalBays,HrsPerDay,TotAtms,WDays,WorkRunCost from Wrk_Details where MonthNo=12")
                        RstRep.AddNew
                        RstRep.Fields("Month") = "December"
                        If TmpRst.RecordCount > 0 Then
                            If VNull(TmpRst!TotalBays) = 0 Then MsgBox "Fill the Bays Details in Workshop Details Master for the month of December": RepPrint = False: Exit Sub
                            If VNull(TmpRst!HrsPerDay) = 0 Then MsgBox "Fill the Working Hours Per Day Detail in Workshop Details Master for the month of December": RepPrint = False: Exit Sub
                            If VNull(TmpRst!TotAtms) = 0 Then MsgBox "Fill the Total ATM Detail in Workshop Details Master for the month of December": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WDays) = 0 Then MsgBox "Fill the Working Days Per Week Detail in Workshop Details Master for the month of December": RepPrint = False: Exit Sub
                            If VNull(TmpRst!WorkRunCost) = 0 Then MsgBox "Fill the Workshop Running Cost Detail in Workshop Details Master for the month of December": RepPrint = False: Exit Sub
                            'Dec
                            With RstRep
                                .Fields("BayNos") = VNull(TmpRst!TotalBays)
                                .Fields("WorkHrsPerDay") = VNull(TmpRst!HrsPerDay): .Fields("DaysPerWeek") = VNull(TmpRst!WDays)
                                .Fields("TotAtms") = VNull(TmpRst!TotAtms): .Fields("Labour") = VNull(GCn.Execute("Select Sum(Job_Lab.LabourAmt) from ((Job_Card Left join Job_Lab on Job_Card.DocId=Job_Lab.Job_DocID) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type)  where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("SparesWork") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) from (Sp_Sale Left Join Job_Card on Sp_Sale.Job_DocId=Job_Card.DocId) Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type where SP_Sale.V_Date >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Sp_Sale.V_type in('W_SIC','W_SIR') ").Fields(0).Value)
                                .Fields("SparesCou") = VNull(GCn.Execute("Select (Sum(Sp_Sale.SprAmt_MRP_TB+SP_Sale.SprAmt_MRP_TP+SP_Sale.OilAmt_MRP_TB+SP_Sale.OilAmt_MRP_TP+SP_Sale.SprAmt_TB+SP_Sale.SprAmt_TP+SP_Sale.OilAmt_TB+SP_Sale.OilAmt_TP)-sum(SP_Sale.D_Amt_TB+SP_Sale.D_Amt_TP+SP_Sale.D_Amt_MRP_TB+SP_Sale.D_Amt_MRP_TP)) as CntSaleDate from Sp_Sale  where Sp_Sale.V_type in('SYSIC','SYSIR') and SP_Sale.V_Date >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and SP_Sale.V_Date <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & "").Fields(0).Value)
                                .Fields("FreeService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('F')").RecordCount)
                                .Fields("Warranty") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('W')").RecordCount)
                                .Fields("PaidService") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Catg in ('C')").RecordCount)
                                .Fields("PaidRepairs") = VNull(GCn.Execute("Select Job_Card.* from (((Job_Card  Left Join Service_Type on Job_Card.Serv_type=Service_Type.Serv_Type) Left Join HisCard on Job_Card.CardNo=Hiscard.CardNo) Left Join Model on HisCard.Model=Model.Model) Left Join Model_Cat on Model.Cat_Code=Model_Cat.ModelCat_Code where Job_Card.JobCloseDate >= " & ConvertDate("01/Dec/" & Year(PubStartDate)) & " and Job_Card.JobCloseDate <= " & ConvertDate("31/Dec/" & Year(PubStartDate)) & " and Service_type.Serv_Catg IN('C') and Service_Type.Serv_Type in ('PR')").RecordCount)
                                .Fields("WorkRunCost") = VNull(TmpRst!WorkRunCost)
                                .Update
                            End With
                        End If
                    End If
            End Select
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "WorkProfitReg"
        RepTitle = UCase(Me.CAPTION)
        Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


