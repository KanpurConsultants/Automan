VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmVehRate 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Rate Declaration"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      HideSelection   =   0   'False
      Left            =   10110
      TabIndex        =   28
      Top             =   6585
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Frame FrmPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00CAECF0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   2880
      TabIndex        =   21
      Top             =   2460
      Visible         =   0   'False
      Width           =   5040
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         DisabledPicture =   "frmVehRate.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3420
         MaskColor       =   &H00FFC0FF&
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Printer "
         Top             =   1020
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00CAECF0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4710
         MousePointer    =   99  'Custom
         Picture         =   "frmVehRate.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Delete Current Record"
         Top             =   0
         Width           =   315
      End
      Begin VB.OptionButton OptPrn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Both"
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   3495
         TabIndex        =   24
         Top             =   420
         Width           =   1260
      End
      Begin VB.OptionButton OptPrn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Govt."
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   23
         Top             =   420
         Width           =   1260
      End
      Begin VB.OptionButton OptPrn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "General"
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   375
         TabIndex        =   22
         Top             =   420
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Printer Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Index           =   3
         Left            =   15
         TabIndex        =   27
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.CommandButton CmdAppyi 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Show List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1110
      Width           =   1275
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   780
      MaxLength       =   20
      TabIndex        =   1
      Top             =   660
      Width           =   2505
   End
   Begin MSDataGridLib.DataGrid DGDate 
      Height          =   2910
      Left            =   780
      Negotiate       =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5133
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Select Date"
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
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   1665
      Negotiate       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Site Name"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7650
      MaxLength       =   8
      TabIndex        =   3
      Top             =   660
      Width           =   675
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   5760
      TabIndex        =   6
      Top             =   4530
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   10590
      MaxLength       =   4
      TabIndex        =   4
      Top             =   660
      Width           =   675
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4440
      MaxLength       =   12
      TabIndex        =   2
      Top             =   660
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   5595
      Left            =   60
      TabIndex        =   7
      Top             =   1560
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   9869
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   3
      Cols            =   26
      FixedRows       =   2
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   13298928
      BackColorBkg    =   12243913
      GridColor       =   0
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   26
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Chassis Type :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   7605
      TabIndex        =   20
      Top             =   1155
      Width           =   2340
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   2475
      TabIndex        =   19
      Top             =   1155
      Width           =   5130
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Model :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   30
      TabIndex        =   18
      Top             =   1155
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   17
      Top             =   675
      Width           =   330
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   615
      TabIndex        =   16
      Top             =   660
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Y/N"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   6405
      TabIndex        =   13
      Top             =   675
      Width           =   1035
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   92
      Left            =   7455
      TabIndex        =   12
      Top             =   660
      Width           =   75
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   8
      Left            =   10350
      TabIndex        =   11
      Top             =   660
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RSO Purchase Y/N"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   27
      Left            =   8700
      TabIndex        =   10
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   91
      Left            =   4275
      TabIndex        =   9
      Top             =   660
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   3780
      TabIndex        =   8
      Top             =   660
      Width           =   405
   End
End
Attribute VB_Name = "frmVehRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsSite As ADODB.Recordset
Dim RstMain As ADODB.Recordset
Dim RsDate  As ADODB.Recordset
Dim GridKey As Integer
'grid color scheme
Private Const CellBackColLeave As String = &HCFE0E0
'Private Const CellForeColLeave As String = &H0&
'Private Const CellBackColEnter As String = &HC0E0FF
'Private Const GridBackColorBkg As String = &HBAD3C9

Private Const VDate As Byte = 0
Private Const TaxYN As Byte = 1
Private Const RSO_WORK As Byte = 2
Private Const SiteCode As Byte = 3


Private Const GenSalePrice  As Byte = 1
Private Const GovtSalePrice As Byte = 2
Private Const GenXGodownPrice As Byte = 3
Private Const GovtXGodownPrice As Byte = 4
Private Const GenDMargin    As Byte = 5
Private Const GovtDMargin   As Byte = 6
Private Const Model         As Byte = 7
Private Const NDP           As Byte = 8
Private Const GenSalePrice2  As Byte = 9
Private Const GovtSalePrice2 As Byte = 10
Private Const InsCharge  As Byte = 11
Private Const Octoroi  As Byte = 12
Private Const TempReg  As Byte = 13
Private Const TransIns  As Byte = 14
Private Const Transport As Byte = 15
Private Const HandlingCharges As Byte = 16
Private Const MVT As Byte = 17
Private Const STotal As Byte = 18
Private Const TaxPer As Byte = 19
Private Const TaxAmt As Byte = 20
Private Const TaxSurPer As Byte = 21
Private Const TaxSurAmt As Byte = 22
Private Const RegFee As Byte = 23
Private Const RegFeeCom As Byte = 24
Private Const InsFee As Byte = 25
Private Const ModName As Byte = 26
Private Const ChasType As Byte = 27


Dim TAddMode As Boolean

Private Sub CmdAppyi_Click()
Dim i As Integer, mEffectDate As Date, mRSO As Byte, mTaxable As Byte
If IsValid(txt(SiteCode), "Site Code") = False Then Exit Sub
If IsValid(txt(VDate), "Date") = False Then Exit Sub
If IsValid(txt(TaxYN), "Tax YN") = False Then Exit Sub
If IsValid(txt(RSO_WORK), "RSO YN") = False Then Exit Sub

GSQL = "SELECT MODEL " & _
    "FROM Veh_Rate " & _
    "where (Veh_Rate.Effective_Date=" & ConvertDate(txt(VDate)) & " and Veh_Rate.Site_Code='" & txt(SiteCode).Tag & "' and Veh_Rate.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & " and Veh_Rate.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & ") "
Set RstMain = New Recordset
RstMain.CursorLocation = adUseClient
RstMain.Open GSQL, GCn, adOpenStatic, adLockReadOnly

If RstMain.RecordCount > 0 Then
    GSQL = "SELECT Model.MODEL as ModelCode, Veh_Rate.*, Model.Chas_Type, Model.Model_Desc " & _
        "FROM Model LEFT JOIN Veh_Rate ON Model.MODEL=Veh_Rate.MODEL " & _
        "where Model.Model not in (Select Model from Veh_Rate VH where VH.Effective_Date=" & ConvertDate(txt(VDate)) & " and Vh.Site_Code='" & txt(SiteCode).Tag & "' and VH.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & " and Vh.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & " ) " & _
        " or (Veh_Rate.Effective_Date=" & ConvertDate(txt(VDate)) & " and Veh_Rate.Site_Code='" & txt(SiteCode).Tag & "' and Veh_Rate.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & " and Veh_Rate.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & ") " & _
        "order by Model.model"
Else
    If MsgBox("No Record Found Of Given Criteria.Do You Want To Copy & Create New Rate List ? ", vbYesNo + vbCritical + vbDefaultButton2, "No Matching Record!") = vbYes Then
        GSQL = "Select top 1 Effective_Date From Veh_Rate " & _
            " Where Effective_Date<" & ConvertDate(txt(VDate)) & _
            " and Veh_Rate.Site_Code='" & txt(SiteCode).Tag & _
            "' and Veh_Rate.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & _
            " and Veh_Rate.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & ""
        
        Set RstMain = New Recordset
        RstMain.CursorLocation = adUseClient
        RstMain.Open GSQL, GCn, adOpenStatic, adLockReadOnly
        If RstMain.RecordCount > 0 Then
            mEffectDate = RstMain!Effective_Date
            GSQL = "SELECT Model.MODEL as ModelCode, Veh_Rate.*, Model.Chas_Type, Model.Model_Desc " & _
                "FROM Model LEFT JOIN Veh_Rate ON Model.MODEL=Veh_Rate.MODEL " & _
                "where Model.Model not in (Select Model from Veh_Rate VH where VH.Effective_Date=" & ConvertDate(mEffectDate) & " and Vh.Site_Code='" & txt(SiteCode).Tag & "' and VH.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & " and Vh.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & " ) " & _
                " or (Veh_Rate.Effective_Date=" & ConvertDate(mEffectDate) & " and Veh_Rate.Site_Code='" & txt(SiteCode).Tag & "' and Veh_Rate.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & " and Veh_Rate.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & ") " & _
                "order by Model.model"
        Else
            mRSO = IIf(txt(RSO_WORK) = "Yes", 1, 0)
            mTaxable = IIf(txt(TaxYN) = "Yes", 1, 0)
            GSQL = "select Model as ModelCode,MODEL," & ConvertDate(txt(VDate)) & " as Effective_Date," & _
                mRSO & " as RSO_WORK," & mTaxable & " as TAXABLE_YN,'" & PubSiteCode & "' as Site_Code," & _
                "0 as P_RATE,0 as S_RATE,0 as G_RATE, 0 as INCI_CHRG,0 as OCTROI,0 as Reg_Temp,0 as Ins_Trn," & _
                "0 as Transport, 0 As HandlingChanges,0 as Tax,0 as Tax_Surcharge,0 as MVT,0 as Reg_Fee,0 as Ins_Fee," & _
                "0 as Margine,Chas_Type, model_desc " & _
                "from Model Order by model"
        End If
    Else
        FGrid.Rows = 2
        FGrid.AddItem FGrid.Rows - 1
        FGrid.FixedRows = 2
   End If
End If
    
    Set RstMain = New Recordset
    RstMain.CursorLocation = adUseClient
    RstMain.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    FGrid.Rows = 2
    FGrid.Redraw = False
    If RstMain.RecordCount > 0 Then
        TopCtrl1.tPrn = True
        i = 2
        Do Until RstMain.EOF
            With FGrid
                .AddItem ""
                .TextMatrix(i, 0) = i - 1
                .TextMatrix(i, Model) = RstMain!ModelCode
                .TextMatrix(i, NDP) = Format(IIf(IsNull(RstMain!p_rate) Or RstMain!p_rate = 0, "", RstMain!p_rate), "0.00")
                .TextMatrix(i, GenSalePrice2) = Format(IIf(IsNull(RstMain!S_Rate) Or RstMain!S_Rate = 0, "", RstMain!S_Rate), "0.00")
                .TextMatrix(i, GovtSalePrice2) = Format(IIf(IsNull(RstMain!G_Rate) Or RstMain!G_Rate = 0, "", RstMain!G_Rate), "0.00")
                .TextMatrix(i, GenXGodownPrice) = Format(IIf(IsNull(RstMain!GenExGodRate) Or RstMain!GenExGodRate = 0, "", RstMain!GenExGodRate), "0.00")
                .TextMatrix(i, GovtXGodownPrice) = Format(IIf(IsNull(RstMain!GovtExGodRate) Or RstMain!GovtExGodRate = 0, "", RstMain!GovtExGodRate), "0.00")
                .TextMatrix(i, InsCharge) = Format(IIf(IsNull(RstMain!INCI_CHRG) Or RstMain!INCI_CHRG = 0, "", RstMain!INCI_CHRG), "0.00")
                .TextMatrix(i, Octoroi) = Format(IIf(IsNull(RstMain!Octroi) Or RstMain!Octroi = 0, "", RstMain!Octroi), "0.00")
                .TextMatrix(i, TempReg) = Format(IIf(IsNull(RstMain!REG_TEMP) Or RstMain!REG_TEMP = 0, "", RstMain!REG_TEMP), "0.00")
                .TextMatrix(i, TransIns) = Format(IIf(IsNull(RstMain!INS_TRN) Or RstMain!INS_TRN = 0, "", RstMain!INS_TRN), "0.00")
                .TextMatrix(i, Transport) = Format(IIf(IsNull(RstMain!Transport) Or RstMain!Transport = 0, "", RstMain!Transport), "0.00")
                .TextMatrix(i, HandlingCharges) = Format(IIf(IsNull(RstMain!HandlingCharges) Or RstMain!HandlingCharges = 0, "", RstMain!HandlingCharges), "0.00")
                .TextMatrix(i, MVT) = Format(IIf(IsNull(RstMain!MVT) Or RstMain!MVT = 0, "", RstMain!MVT), "0.00")
                .TextMatrix(i, TaxPer) = Format(IIf(IsNull(RstMain!TAX) Or RstMain!TAX = 0, "", RstMain!TAX), "0.00")
                .TextMatrix(i, TaxSurPer) = Format(IIf(IsNull(RstMain!tax_SURCHARGE) Or RstMain!tax_SURCHARGE = 0, "", RstMain!tax_SURCHARGE), "0.00")
                .TextMatrix(i, RegFee) = Format(IIf(IsNull(RstMain!REG_FEE) Or RstMain!REG_FEE = 0, "", RstMain!REG_FEE), "0.00")
                .TextMatrix(i, RegFeeCom) = Format(IIf(IsNull(RstMain!REG_FEECom) Or RstMain!REG_FEECom = 0, "", RstMain!REG_FEECom), "0.00")
                .TextMatrix(i, InsFee) = Format(IIf(IsNull(RstMain!INS_FEE) Or RstMain!INS_FEE = 0, "", RstMain!INS_FEE), "0.00")
                .TextMatrix(i, ModName) = IIf(IsNull(RstMain!Model_Desc), "", RstMain!Model_Desc)
                .TextMatrix(i, ChasType) = IIf(IsNull(RstMain!Chas_Type), "", RstMain!Chas_Type)
                .TextMatrix(i, GenDMargin) = Format(Val(.TextMatrix(i, GenSalePrice2)) - Val(.TextMatrix(i, NDP)), "0.00")
                .TextMatrix(i, GovtDMargin) = Format(Val(.TextMatrix(i, GovtSalePrice2)) - Val(.TextMatrix(i, NDP)), "0.00")
                .TextMatrix(i, STotal) = Format(Val(.TextMatrix(i, NDP)) + Val(.TextMatrix(i, GenDMargin)) _
                + Val(.TextMatrix(i, InsCharge)) + Val(.TextMatrix(i, Octoroi)) + Val(.TextMatrix(i, TempReg)) + Val(.TextMatrix(i, TransIns)) _
                + Val(.TextMatrix(i, Transport)) + Val(.TextMatrix(i, HandlingCharges)) + Val(.TextMatrix(i, MVT)), "0.00")
                 
                If txt(TaxYN) = "Yes" Then
                    .TextMatrix(i, TaxAmt) = Val(.TextMatrix(i, TaxPer)) * Val(.TextMatrix(i, STotal)) / 100
                    .TextMatrix(i, TaxSurAmt) = Val(.TextMatrix(i, TaxSurPer)) * Val(.TextMatrix(i, TaxAmt)) / 100
                Else
                    .TextMatrix(i, TaxAmt) = Val(.TextMatrix(i, TaxPer)) * (Val(.TextMatrix(i, NDP)) + Val(.TextMatrix(i, GenDMargin))) / 100
                    .TextMatrix(i, TaxSurAmt) = Val(.TextMatrix(i, TaxSurPer)) * Val(.TextMatrix(i, TaxAmt)) / 100
                End If
            
                
'                .TextMatrix(I, GovtSalePrice) = Format(Val(.TextMatrix(I, NDP)) + Val(.TextMatrix(I, GovtDMargin)) _
'                + Val(.TextMatrix(I, InsCharge)) + Val(.TextMatrix(I, Octoroi)) + Val(.TextMatrix(I, TempReg)) + Val(.TextMatrix(I, TransIns)) _
'                + Val(.TextMatrix(I, Transport)) + Val(.TextMatrix(I, HandlingCharges)) + Val(.TextMatrix(I, MVT)) + Val(.TextMatrix(I, RegFee)) + Val(.TextMatrix(I, InsFee)), "0.00")
                
                .TextMatrix(i, GovtSalePrice) = Format(Val(.TextMatrix(i, NDP)) + Val(.TextMatrix(i, GovtDMargin)) _
                + Val(.TextMatrix(i, InsCharge)) + Val(.TextMatrix(i, Octoroi)) + Val(.TextMatrix(i, TempReg)) + Val(.TextMatrix(i, TransIns)) _
                + Val(.TextMatrix(i, Transport)) + Val(.TextMatrix(i, HandlingCharges)) + Val(.TextMatrix(i, MVT)), "0.00")
                
                
'                .TextMatrix(I, GenSalePrice) = Format(Val(.TextMatrix(I, NDP)) + Val(.TextMatrix(I, GenDMargin)) _
'                + Val(.TextMatrix(I, InsCharge)) + Val(.TextMatrix(I, Octoroi)) + Val(.TextMatrix(I, TempReg)) + Val(.TextMatrix(I, TransIns)) _
'                + Val(.TextMatrix(I, Transport)) + Val(.TextMatrix(I, HandlingCharges)) + Val(.TextMatrix(I, MVT)) + Val(.TextMatrix(I, TaxAmt)) + Val(.TextMatrix(I, TaxSurAmt)) + Val(.TextMatrix(I, RegFee)) + Val(.TextMatrix(I, InsFee)), "0.00")
                
                .TextMatrix(i, GenSalePrice) = Format(Val(.TextMatrix(i, NDP)) + Val(.TextMatrix(i, GenDMargin)) _
                + Val(.TextMatrix(i, InsCharge)) + Val(.TextMatrix(i, Octoroi)) + Val(.TextMatrix(i, TempReg)) + Val(.TextMatrix(i, TransIns)) _
                + Val(.TextMatrix(i, Transport)) + Val(.TextMatrix(i, HandlingCharges)) + Val(.TextMatrix(i, MVT)) + Val(.TextMatrix(i, TaxAmt)) + Val(.TextMatrix(i, TaxSurAmt)), "0.00")
                
            End With
            RstMain.MoveNext
           i = i + 1
        Loop
    Else
        MsgBox "Open Model From Master", vbInformation, "No Model"
    End If
    If FGrid.Rows = 2 Then FGrid.AddItem FGrid.Rows - 1
    FGrid.FixedRows = 2
    FGrid.Redraw = True
    FGrid.SetFocus
End Sub

Private Sub CmdPrint_Click(Index As Integer)
Dim mQRY As String, mQRY1 As String, Rst As ADODB.Recordset, i As Integer
Select Case Index
    Case 1
        FrmPrint.Visible = False
    Case 0
        If OptPrn(0).Value = True Then
            mQRY1 = "Veh_rate.S_Rate  > 0"
        ElseIf OptPrn(1).Value = True Then
            mQRY1 = "Veh_rate.G_Rate  > 0"
        Else
           mQRY1 = "(Veh_rate.S_Rate + Veh_rate.G_Rate  > 0)"
        End If
            
    mQRY = "SELECT Veh_Rate.*, Model.Chas_Type, Model.Model_Desc " & _
    " FROM Veh_Rate LEFT JOIN Model ON Veh_Rate.MODEL = Model.MODEL where " & _
    "Veh_Rate.Effective_Date=" & ConvertDate(txt(VDate)) & " and Veh_Rate.Site_Code='" & txt(SiteCode).Tag & "' and " & _
    "Veh_Rate.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & " and Veh_Rate.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & " and " & mQRY1 & "  " & _
    "order by veh_rate.model"
        
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

        CreateFieldDefFile Rst, PubRepoPath + "\VehRateList.TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath & "\VehRateList.RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        
        For i = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
            Case UCase("GenPrice")
                rpt.FormulaFields(i).TEXT = "'" & IIf(OptPrn(0).Value = True Or OptPrn(2).Value = True, "General", "") & "'"
            Case UCase("GovtPrice")
                rpt.FormulaFields(i).TEXT = "'" & IIf(OptPrn(1).Value = True Or OptPrn(2).Value = True, "Govt.", "") & "'"
        End Select
        Next

        Call Report_View(rpt, "Vehicle Rate List")
Exit Sub
    
End Select
End Sub

Private Sub FGrid_RowColChange()
If FGrid.Row > 2 Then
Label1(0).CAPTION = "Model : " & FGrid.TextMatrix(FGrid.Row, Model)
Label1(1).CAPTION = "Name : " & FGrid.TextMatrix(FGrid.Row, ModName)
Label1(2).CAPTION = "Chassis Type : " & FGrid.TextMatrix(FGrid.Row, ChasType)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
Exit Sub

ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub


Private Sub Form_Load()
On Error GoTo ELoop
Dim i As Byte
TopCtrl1.Tag = UserPermission(Me.Name)
    
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DgSite.DataSource = RsSite
    
    Set RsDate = New ADODB.Recordset
    RsDate.CursorLocation = adUseClient
    RsDate.Open "select distinct " & cCStr("Effective_Date") & " as code from Veh_Rate", GCn, adOpenDynamic, adLockOptimistic
    Set DGDate.DataSource = RsDate
    WinSetting Me
    Ini_Grid
 
    Disp_Text False
    TopCtrl1.tPrn = False
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Resize()
'Ini_Grid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsSite = Nothing
Set RsDate = Nothing
End Sub




Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
 Dim i As Integer
    Disp_Text True
    txt(SiteCode).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub


Private Sub TopCtrl1_eCancel()
Dim i As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        FGrid.Rows = 2
        FGrid.AddItem FGrid.Rows - 1
        FGrid.FixedRows = 2
    Disp_Text False
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
FrmPrint.left = 1335
FrmPrint.top = 350
FrmPrint.Visible = True
CmdPrint(0).SetFocus
End Sub

Private Sub TopCtrl1_eSave()
    Dim i As Integer
    Dim mTrans As Boolean
'    On Error GoTo errlbl

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(txt(SiteCode), "Site Code") = False Then Exit Sub
    If IsValid(txt(VDate), "Date") = False Then Exit Sub
    If IsValid(txt(TaxYN), "Tax YN") = False Then Exit Sub
    If IsValid(txt(RSO_WORK), "RSO YN") = False Then Exit Sub

    GCn.BeginTrans
    mTrans = True

    GCn.Execute ("delete from Veh_Rate where " & _
    "Veh_Rate.Effective_Date=" & ConvertDate(txt(VDate)) & " and Veh_Rate.Site_Code='" & txt(SiteCode).Tag & "' and " & _
    "Veh_Rate.RSO_WORK=" & IIf(txt(RSO_WORK) = "Yes", 1, 0) & " and Veh_Rate.TAXABLE_YN=" & IIf(txt(TaxYN) = "Yes", 1, 0) & "")
    
    For i = 2 To FGrid.Rows - 1
        'If FGrid.TextMatrix(i, NDP) <> "" Then
            GCn.Execute ("insert into Veh_Rate( " & _
            "MODEL,Effective_Date,RSO_WORK, " & _
            "TAXABLE_YN,Site_Code,P_RATE, " & _
            "S_RATE,G_RATE,INCI_CHRG, " & _
            "OCTROI,REG_TEMP,INS_TRN, " & _
            "TRANSPORT, HandlingCharges,TAX,Tax_SURCHARGE, " & _
            "MVT,REG_FEE, Reg_FeeCom,INS_FEE, " & _
            "MARGINE, GenExGodRate, GovtExGodRate,U_Name,U_EntDt,U_AE) " & _
            "values('" & FGrid.TextMatrix(i, Model) & "'," & ConvertDate(txt(VDate)) & "," & IIf(txt(RSO_WORK).TEXT = "Yes", 1, 0) & ", " & _
            "" & IIf(txt(TaxYN).TEXT = "Yes", 1, 0) & ",'" & txt(SiteCode).Tag & "'," & Val(FGrid.TextMatrix(i, NDP)) & ", " & _
            "" & Val(FGrid.TextMatrix(i, GenSalePrice2)) & "," & Val(FGrid.TextMatrix(i, GovtSalePrice2)) & "," & Val(FGrid.TextMatrix(i, InsCharge)) & ", " & _
            "" & Val(FGrid.TextMatrix(i, Octoroi)) & "," & Val(FGrid.TextMatrix(i, TempReg)) & "," & Val(FGrid.TextMatrix(i, TransIns)) & ", " & _
            "" & Val(FGrid.TextMatrix(i, Transport)) & "," & Val(FGrid.TextMatrix(i, HandlingCharges)) & "," & Val(FGrid.TextMatrix(i, TaxPer)) & "," & Val(FGrid.TextMatrix(i, TaxSurPer)) & ", " & _
            "" & Val(FGrid.TextMatrix(i, MVT)) & ", " & _
            "" & Val(FGrid.TextMatrix(i, RegFee)) & ", " & Val(FGrid.TextMatrix(i, RegFeeCom)) & "," & Val(FGrid.TextMatrix(i, InsFee)) & "," & Val(FGrid.TextMatrix(i, GenDMargin)) & ", " & Val(FGrid.TextMatrix(i, GenXGodownPrice)) & ", " & Val(FGrid.TextMatrix(i, GovtXGodownPrice)) & ", " & _
            "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
        'End If
    Next

GCn.CommitTrans
mTrans = False
RsDate.Requery
    Disp_Text False
    Exit Sub
errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub


Private Sub Txt_GotFocus(Index As Integer)
    TxtGrid(0).Visible = False
    Ctrl_GetFocus txt(Index)
    Grid_Hide
Select Case Index
    Case SiteCode
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txt(Index).Tag = RsSite!Code
            txt(Index).TEXT = RsSite!Name
        Else
            If txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txt(Index).TEXT & "'"
            End If
        End If
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case SiteCode
        DGridTxtKeyDown DgSite, txt, Index, RsSite, KeyCode, False, 1
    Case VDate
        DGridTxtKeyDown_Mast DGDate, txt, Index, RsDate, KeyCode, False, 0
End Select

If DgSite.Visible = False And DGDate.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If Index <> SiteCode And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case Index
    Case SiteCode
        If DgSite.Visible = True Then DGridTxtKeyPress txt, Index, RsSite, KeyAscii, "Name"
    Case TaxYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txt(Index) = ""
        End If
        KeyAscii = 0
    Case RSO_WORK
        If UCase(Chr(KeyAscii)) = "Y" Then
            txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txt(Index) = ""
        End If
        KeyAscii = 0
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case VDate
        If DGDate.Visible = True Then DGridTxtKeyUp_Mast txt, Index, RsDate, KeyCode, "code"
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Select Case Index
     Case SiteCode
        If IsValid(txt(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsSite!Name
            txt(Index).Tag = RsSite!Code
        End If
    Case VDate
        If IsValid(txt(Index), "Date") = False Then Cancel = True: Exit Sub
        If RsDate.EOF = False And RsDate.BOF = False And txt(Index) <> "" Then
            txt(Index) = Format(RsDate(0), "DD/MMM/YYYY")
        End If
        txt(Index).TEXT = RetDate(txt(Index))
     
End Select
End Sub

Private Sub DgDate_Click()
DGDate.Visible = False
If RsDate.RecordCount > 0 Then
            txt(VDate).TEXT = RsDate!Code
End If
    txt(VDate).SetFocus
End Sub
Private Sub DGSite_Click()
    If RsSite.RecordCount > 0 Then
        txt(SiteCode).TEXT = RsSite!Name
        FGrid.TextMatrix(FGrid.Row, SiteCode) = RsSite!Code
    End If
    DgSite.Visible = False
    txt(SiteCode).SetFocus
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, TaxPer, TaxSurPer, RegFee, InsFee, RegFeeCom
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            Amt_Cal
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, TaxPer, TaxSurPer, RegFee, InsFee, RegFeeCom
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, TaxPer, TaxSurPer, RegFee, InsFee, RegFeeCom
        Call GridDblClick(Me, FGrid, TxtGrid, 0)
End Select
TAddMode = False
End Sub

Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
Select Case FGrid.Col
    Case NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, TaxPer, TaxSurPer, RegFee, InsFee, RegFeeCom
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case Model
        If TopCtrl1.TopText2 <> "Browse" Then
            SelGridKeyPress TxtSearch, FGrid, RstMain, KeyAscii, "ModelCode", CellBackColEnter, vbWhite, 2
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = FGrid.BackColor  'CellBackColLeave
'    FGrid.CellForeColor = CellForeColLeave
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
For i = 0 To txt.Count - 1
    txt(i).TEXT = ""
Next i
End Sub

Private Sub Ini_Grid()
    With FGrid
        .top = 1560
        .left = Me.left '+ 45
        .width = Me.width - 90
        .Rows = 24
        .RowHeightMin = PubGridRowHeight
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .MergeRow(0) = True
        .ColAlignmentFixed = flexAlignRightCenter
        

        .TextMatrix(0, 0) = "S.No."
        .TextMatrix(1, 0) = "S.No."
        .ColAlignmentFixed(0) = flexAlignLeftCenter
        .ColWidth(0) = 550

        .MergeCol(1) = True
        .TextMatrix(0, Model) = "Model"
        '.TextMatrix(1, Model) = "Model"
        .ColAlignmentFixed(Model) = flexAlignLeftCenter
        .ColWidth(Model) = 2500

        '.MergeCol(GenSalePrice) = True
        .TextMatrix(0, GenSalePrice) = "Net Sale Price"
        .TextMatrix(1, GenSalePrice) = "Gen"
        .ColWidth(GenSalePrice) = 1200
        .ColAlignment(GenSalePrice) = flexAlignRightCenter
        
        '.TextMatrix(0, GovtSalePrice) = .TextMatrix(0, GenSalePrice)
        .TextMatrix(0, GovtSalePrice) = "Net Sale Price"
        .TextMatrix(1, GovtSalePrice) = "Govt"
        .ColWidth(GovtSalePrice) = 1200
        .ColAlignment(GovtSalePrice) = flexAlignRightCenter


        .TextMatrix(0, GenXGodownPrice) = "Ex Godown Price"
        .TextMatrix(1, GenXGodownPrice) = "Gen"
        .ColWidth(GenXGodownPrice) = 1200
        .ColAlignment(GenXGodownPrice) = flexAlignRightCenter
        
        '.TextMatrix(0, GovtSalePrice) = .TextMatrix(0, GenSalePrice)
        .TextMatrix(0, GovtXGodownPrice) = "Ex Godown Price"
        .TextMatrix(1, GovtXGodownPrice) = "Govt"
        .ColWidth(GenXGodownPrice) = 1200
        .ColAlignment(GenXGodownPrice) = flexAlignRightCenter


        .TextMatrix(0, GenDMargin) = "Dealer Margin          "
        .TextMatrix(1, GenDMargin) = "Gen"
        .ColWidth(GenDMargin) = 1200
        .ColAlignment(GenDMargin) = flexAlignRightCenter

        .TextMatrix(0, GovtDMargin) = .TextMatrix(0, GenDMargin)
        .TextMatrix(1, GovtDMargin) = "Govt"
        .ColWidth(GovtDMargin) = 1200
        .ColAlignment(GovtDMargin) = flexAlignRightCenter
        
        .TextMatrix(0, NDP) = .TextMatrix(0, GenSalePrice)
        .TextMatrix(0, NDP) = "NDP"
        .ColWidth(NDP) = 1200
        
        .TextMatrix(0, GenSalePrice2) = "Sale Price"
        .TextMatrix(1, GenSalePrice2) = "Gen"
        .ColWidth(GenSalePrice2) = 1200

        .TextMatrix(0, GovtSalePrice2) = .TextMatrix(0, GenSalePrice2)
        .TextMatrix(1, GovtSalePrice2) = "Govt"
        .ColWidth(GovtSalePrice2) = 1200
       
        .TextMatrix(0, InsCharge) = "Incidental"
        .TextMatrix(1, InsCharge) = "Charges   "
        .ColWidth(InsCharge) = 1000

        .TextMatrix(0, Octoroi) = "Octroi"
        .ColWidth(Octoroi) = 1000
        
        .TextMatrix(0, TempReg) = "Temp."
        .TextMatrix(1, TempReg) = "Reg."
        .ColWidth(TempReg) = 1000
        
        .TextMatrix(0, TransIns) = "Transit"
        .TextMatrix(1, TransIns) = "Insu."
        .ColWidth(TransIns) = 1000
        
        .TextMatrix(0, Transport) = "Transport   "
        .ColAlignmentFixed(Transport) = flexAlignRightCenter
        .ColWidth(Transport) = 1100
        
        .TextMatrix(0, HandlingCharges) = "Hand. Chg.   "
        .ColAlignmentFixed(HandlingCharges) = flexAlignRightCenter
        .ColWidth(HandlingCharges) = 1100
        
        .TextMatrix(0, MVT) = "MVT     "
        .ColAlignmentFixed(MVT) = flexAlignRightCenter
        .ColWidth(MVT) = 1000
        
        .TextMatrix(0, STotal) = "Sub Total   "
        .ColAlignmentFixed(STotal) = flexAlignRightCenter
        .ColWidth(STotal) = 1200
        
        .TextMatrix(0, TaxPer) = "Tax Details"
        .TextMatrix(1, TaxPer) = "%"
        .ColWidth(TaxPer) = 1000
       
        .TextMatrix(0, TaxAmt) = .TextMatrix(0, TaxPer)
        .TextMatrix(1, TaxAmt) = "Amount"
        .ColWidth(TaxAmt) = 1000
        
        .TextMatrix(0, TaxSurPer) = "Surcharge Detail"
        .TextMatrix(1, TaxSurPer) = "%"
        .ColWidth(TaxSurPer) = 1150
        
        .TextMatrix(0, TaxSurAmt) = .TextMatrix(0, TaxSurPer)
        .TextMatrix(1, TaxSurAmt) = "Amount"
        .ColWidth(TaxSurAmt) = 1000
       
        .TextMatrix(0, RegFee) = "RegFee"
        .TextMatrix(1, RegFee) = "Personal"
        .ColAlignmentFixed(RegFee) = flexAlignRightCenter
        .ColWidth(RegFee) = 1100
        
        .TextMatrix(0, RegFeeCom) = "Reg. Fee"
        .TextMatrix(1, RegFeeCom) = "Commercial"
        .ColAlignmentFixed(RegFeeCom) = flexAlignRightCenter
        .ColWidth(RegFeeCom) = 1100
        
        
        .TextMatrix(0, InsFee) = "Insurance"
        .TextMatrix(1, InsFee) = "Charges"
        .ColAlignmentFixed(InsFee) = flexAlignRightCenter
        .ColWidth(InsFee) = 1100
        
        .ColWidth(ModName) = 0
        .ColWidth(ChasType) = 0
        
        .Rows = 2
        .AddItem FGrid.Rows - 1
        .FixedRows = 2
        .FixedCols = 7
End With
DgSite.left = txt(SiteCode).left:   DgSite.top = txt(SiteCode).top + txt(SiteCode).height
DGDate.left = txt(VDate).left: DGDate.top = txt(VDate).top + txt(VDate).height

End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim i As Integer
For i = 0 To txt.Count - 1
    txt(i).Enabled = Enb
    txt(i).ForeColor = CtrlFColOrg
Next
If Enb = True Then
    TopCtrl1.tEdit = False
    TopCtrl1.tCancel = True
    TopCtrl1.tSave = True
    CmdAppyi.Enabled = True
    TopCtrl1.TopText2.CAPTION = "Edit"
Else
    TopCtrl1.tEdit = True
    TopCtrl1.tCancel = False
    TopCtrl1.tSave = False
    CmdAppyi.Enabled = False
    TopCtrl1.TopText2.CAPTION = "Browse"
End If
TopCtrl1.tAdd = False
TopCtrl1.tRef = False
TopCtrl1.tDel = False
TopCtrl1.tFirst = False
TopCtrl1.tPrev = False
TopCtrl1.tNext = False
TopCtrl1.tLast = False
TopCtrl1.tFind = False
TopCtrl1.tExit = True
TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If DgSite.Visible = True Then DgSite.Visible = False
    If DGDate.Visible = True Then DGDate.Visible = False
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyEscape Then
                TxtGrid(0).TEXT = TxtGrid(0).Tag
                TxtGrid_KeyUp Index, KeyCode, Shift
                TxtGrid(0).Visible = False
                FGrid.SetFocus
                Grid_Hide
                Exit Sub
            End If
        Select Case FGrid.Col
            Case NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, TaxPer, TaxSurPer, RegFee, InsFee, RegFeeCom
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, InsFee
                        End If
                End If
        End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
     Case NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, TaxPer, TaxSurPer, RegFee, InsFee, RegFeeCom
        Call NumPress(TxtGrid(0), KeyAscii, 8, 2)
End Select
End Sub


Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
            Case TaxPer, TaxSurPer, NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, RegFee, InsFee, RegFeeCom
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
                Amt_Cal
End Select
End Sub


Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Dim j As Integer
Select Case FGrid.Col
     Case NDP, GenSalePrice2, GovtSalePrice2, InsCharge, Octoroi, TempReg, TransIns, Transport, HandlingCharges, MVT, TaxPer, TaxSurPer, RegFee, InsFee
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(TxtGrid(0).TEXT, "0.00")
                Amt_Cal
End Select
End Sub
Private Function TxtGridLeave() As Boolean
Dim j As Integer
Select Case FGrid.Col
     Case Model
            FGrid.TextMatrix(FGrid.Row, Model) = Format(TxtGrid(0).TEXT, "0.00")
            Amt_Cal
    End Select
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function

Private Sub Amt_Cal()

Dim mExGodownSubTotal As Double
With FGrid
     .TextMatrix(FGrid.Row, GenDMargin) = Format(Val(.TextMatrix(FGrid.Row, GenSalePrice2)) - Val(.TextMatrix(FGrid.Row, NDP)), "0.00")
    .TextMatrix(FGrid.Row, GovtDMargin) = Format(Val(.TextMatrix(FGrid.Row, GovtSalePrice2)) - Val(.TextMatrix(FGrid.Row, NDP)), "0.00")
    
    .TextMatrix(FGrid.Row, STotal) = Format(Val(.TextMatrix(FGrid.Row, NDP)) + Val(.TextMatrix(FGrid.Row, GenDMargin)) _
    + Val(.TextMatrix(FGrid.Row, InsCharge)) + Val(.TextMatrix(FGrid.Row, Octoroi)) + Val(.TextMatrix(FGrid.Row, TempReg)) + Val(.TextMatrix(FGrid.Row, TransIns)) _
    + Val(.TextMatrix(FGrid.Row, Transport)) + Val(.TextMatrix(FGrid.Row, HandlingCharges)) + Val(.TextMatrix(FGrid.Row, MVT)), "0.00")
     
    If txt(TaxYN) = "Yes" Then
        .TextMatrix(FGrid.Row, TaxAmt) = Val(.TextMatrix(FGrid.Row, TaxPer)) * Val(.TextMatrix(FGrid.Row, STotal)) / 100
        .TextMatrix(FGrid.Row, TaxSurAmt) = Val(.TextMatrix(FGrid.Row, TaxSurPer)) * Val(.TextMatrix(FGrid.Row, TaxAmt)) / 100
    Else
        .TextMatrix(FGrid.Row, TaxAmt) = Val(.TextMatrix(FGrid.Row, TaxPer)) * (Val(.TextMatrix(FGrid.Row, NDP)) + Val(.TextMatrix(FGrid.Row, GenDMargin))) / 100
        .TextMatrix(FGrid.Row, TaxSurAmt) = Val(.TextMatrix(FGrid.Row, TaxSurPer)) * Val(.TextMatrix(FGrid.Row, TaxAmt)) / 100
    End If

    .TextMatrix(FGrid.Row, GovtSalePrice) = Format(Val(.TextMatrix(FGrid.Row, NDP)) + Val(.TextMatrix(FGrid.Row, GovtDMargin)) _
    + Val(.TextMatrix(FGrid.Row, InsCharge)) + Val(.TextMatrix(FGrid.Row, Octoroi)) + Val(.TextMatrix(FGrid.Row, TempReg)) + Val(.TextMatrix(FGrid.Row, TransIns)) _
    + Val(.TextMatrix(FGrid.Row, Transport)) + Val(.TextMatrix(FGrid.Row, HandlingCharges)) + Val(.TextMatrix(FGrid.Row, MVT)), "0.00")
    
    .TextMatrix(FGrid.Row, GenSalePrice) = Format(Val(.TextMatrix(FGrid.Row, NDP)) + Val(.TextMatrix(FGrid.Row, GenDMargin)) _
    + Val(.TextMatrix(FGrid.Row, InsCharge)) + Val(.TextMatrix(FGrid.Row, Octoroi)) + Val(.TextMatrix(FGrid.Row, TempReg)) + Val(.TextMatrix(FGrid.Row, TransIns)) _
    + Val(.TextMatrix(FGrid.Row, Transport)) + Val(.TextMatrix(FGrid.Row, HandlingCharges)) + Val(.TextMatrix(FGrid.Row, MVT)) + Val(.TextMatrix(FGrid.Row, TaxAmt)) + Val(.TextMatrix(FGrid.Row, TaxSurAmt)), "0.00")
      
    .TextMatrix(FGrid.Row, GovtXGodownPrice) = Format(Val(.TextMatrix(FGrid.Row, NDP)) + Val(.TextMatrix(FGrid.Row, GovtDMargin)) _
    + Val(.TextMatrix(FGrid.Row, InsCharge)) + Val(.TextMatrix(FGrid.Row, TempReg)) + Val(.TextMatrix(FGrid.Row, TransIns)) _
    + Val(.TextMatrix(FGrid.Row, Transport)) + Val(.TextMatrix(FGrid.Row, HandlingCharges)) + Val(.TextMatrix(FGrid.Row, MVT)), "0.00")
      
    mExGodownSubTotal = Format(Val(.TextMatrix(FGrid.Row, NDP)) + Val(.TextMatrix(FGrid.Row, GenDMargin)) _
    + Val(.TextMatrix(FGrid.Row, InsCharge)) + Val(.TextMatrix(FGrid.Row, TempReg)) + Val(.TextMatrix(FGrid.Row, TransIns)) _
    + Val(.TextMatrix(FGrid.Row, Transport)) + Val(.TextMatrix(FGrid.Row, HandlingCharges)) + Val(.TextMatrix(FGrid.Row, MVT)), "0.00")
            
    .TextMatrix(FGrid.Row, GenXGodownPrice) = mExGodownSubTotal + (mExGodownSubTotal * Val(.TextMatrix(FGrid.Row, TaxPer)) / 100) + ((mExGodownSubTotal * Val(.TextMatrix(FGrid.Row, TaxPer)) / 100) * Val(.TextMatrix(FGrid.Row, TaxSurPer)) / 100)
      
End With
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If NavigationKey(KeyCode) = True Then FGrid.SetFocus: TxtSearch.Visible = False
    If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
        FGrid.Col = Model: FGrid.SetFocus: TxtSearch.Visible = False
        TxtSearch = ""
    End If
    
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
    Select Case FGrid.Col
        Case Model
            If TopCtrl1.TopText2 <> "Browse" Then
                SelGridKeyPress TxtSearch, FGrid, RstMain, KeyAscii, "ModelCode", CellBackColEnter, vbWhite, 2: KeyAscii = 0
            End If
    End Select
End Sub

Private Sub TxtSearch_LostFocus()
    TxtSearch = ""
End Sub
