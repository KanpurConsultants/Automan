VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form RepFormTax 
   BackColor       =   &H00C8E8DA&
   Caption         =   "Tax Report's"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9855
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Print"
      DownPicture     =   "RepFormTax.frx":0000
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
      Left            =   1965
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   6075
      Width           =   1290
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
      DownPicture     =   "RepFormTax.frx":3132
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
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exit Form"
      Top             =   6075
      Width           =   1290
   End
   Begin VB.PictureBox Pic 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808000&
      Enabled         =   0   'False
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9795
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6045
      Width           =   9855
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
Attribute VB_Name = "RepFormTax"
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

' vijay's work
Private Const SprPurForm As Byte = 29       '**** Spare
Private Const SprSaleForm As Byte = 30      '**** Spare
Private Const VehPurForm As Byte = 31       '**** Vehicle
Private Const VehSaleForm As Byte = 32      '**** Vehicle
Private Const SprFormUtz As Byte = 33       '**** Spare
Private Const VehFormUtz As Byte = 34       '**** Vehicle
Private Const SprFormRemind As Byte = 35    '**** Spare
Private Const VehFormRemind As Byte = 36    '**** Vehicle

Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const Date3 As Byte = 2            '*****constant defined by vijay for form reminder

Private Const List1 As Byte = 2
Private Const List2 As Byte = 3
Private Const List3 As Byte = 4

Private Const Cat1 As Byte = 5
Private Const Cat2 As Byte = 6

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
'**************work done by vijay****************************************************
    Case SprFormRemind                   '***** Printing for Spare Reminder form
        SprFormReminder
    Case VehFormRemind                   '***** Printing for Vehicle Reminder form
        VehFormReminder
    Case VehFormUtz                   '***** Printing for vehicle road  utilization form
        VehFormUtilize
    Case SprFormUtz                  '***** printing for spare road utilization form
        SprFormUtilize
    Case VehPurForm, VehSaleForm     '****printing for vehicle purchase/sales Tax forms
        VehPurSaleFormTax
    Case SprPurForm, SprSaleForm     '****** printing for spare purchase/sale Tax Form
        SprPurFormTax
End Select
If RepPrint = False Then Exit Sub
CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")

rpt.Database.SetDataSource RstRep
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
rpt.ReadRecords

Call Formulas
Call Report_View(rpt, RepTitle, , False)
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

Private Sub GridSel_KeyPress(Index As Integer, keyascii As Integer)
If GridSel(Index).Col = 0 Or GridSel(Index).Row = 0 Then Exit Sub
Select Case Index
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid1, keyascii, RsGrid1.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid2, keyascii, RsGrid2.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid3, keyascii, RsGrid3.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid4, keyascii, RsGrid4.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
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

Private Sub TxtSearch_KeyPress(keyascii As Integer)
Select Case TxtSearch.Tag
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid1, keyascii, RsGrid1.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid2, keyascii, RsGrid2.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid3, keyascii, RsGrid3.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid4, keyascii, RsGrid4.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
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

Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Row
        Case Cat1, Cat2
            TxtGrid(0).MaxLength = 5
        Case List1
            Select Case GRepFormName
                Case SprPurForm, SprSaleForm, VehPurForm, VehSaleForm
                    ListArray = Array("All", "Pending")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            End Select
        Case List2
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
Case List1, List2, List3
        ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
Case Date1, Date2, Date3, Cat1, Cat2
    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
        If TxtGridLeave = True Then TxtKeyDown
    End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(keyascii)
 Select Case FGrid.Row
    Case Cat1, Cat2
        NumPress TxtGrid(Index), keyascii, 2, 2
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
        Case Cat1, Cat2
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
        Case List1, List2, List3
            If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
        Case Date1, Date2
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        TxtGrid(0).Visible = False
        FGrid.SetFocus
    End If
End Function

'******* Fuctions **********

Private Sub Global_Grid()
Dim I As Integer, Cnt As Integer

Pic.top = Me.top - Pic.width - 10
BTNPRINT.left = (Pic.width - (BTNPRINT.width + BTNEXIT.width)) / 2: BTNPRINT.top = Pic.top + 10

BTNEXIT.left = BTNPRINT.left + BTNPRINT.width: BTNEXIT.top = Pic.top + 10

FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75

FGrid.Rows = 7  '5
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
FGrid.height = ((mLastRow + 1) * PubGridRowHeight) + 500

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
        Case Date1, Date2, Date3, List1, List2, List3, Cat1, Cat2
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
Dim I As Integer
    Select Case FGrid.Row
        Case Cat1, Cat2
            If keyascii = 46 Or (keyascii >= 48 And keyascii <= 57) Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, True, keyascii)
            Else
                keyascii = 0
            End If
        Case Date1, Date2, Date3, List1, List2, List3, Cat1, Cat2
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, keyascii)
    End Select
If keyascii <> vbKeyReturn Then TAddMode = True
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
        Case Date1, Date2, Date3, List1, List2, List3
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
'Date1,Date2,List1,List1,List2,List3
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Select Case GRepFormName      'vijay (16/10/02) work for part wise purchase/sale and summary for purchase/sale
    
    Case SprFormUtz ' vijay (date 18/10/02) for road permit utilization for spare
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 1
        
          Grid1Sql = "select '' as O,Form_Desc as FormDescription,form_Code as FormCode from TaxForms where taxforms.trn_type='form 31' AND taxforms.FormTrnType <> 0 order by Form_Desc"
          GridInitialise 1, Grid1Sql
          
    Case VehFormUtz ' vijay (date 18/10/02) for road permit utilization for Vehicle
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 1
        
          Grid1Sql = "select '' as O,Form_Desc as FormDescription,form_Code as FormCode from TaxForms where taxforms.trn_type='form 31' AND taxforms.FormTrnType <> 0 order by Form_Desc"
          GridInitialise 1, Grid1Sql
          
    Case SprFormRemind, VehFormRemind ' vijay (date 18/10/02) for Reminder Date
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date3, 0) = "Reminder Date": .RowHeight(Date3) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(Date3, 1) = PubLoginDate
        End With
        mFirstRow = Date1: mLastRow = Date3: mHelpGridNo = 1
        
          Grid1Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
          GridInitialise 1, Grid1Sql
    
    Case SprPurForm, SprSaleForm     'vijay (date 16/10/02) for tax forms
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Pending/All": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
          mFirstRow = Date1: FGrid.Row = mFirstRow: mHelpGridNo = 3
          mLastRow = List1
          Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 1, Grid1Sql
          Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
          GridInitialise 2, Grid2Sql
        If GRepFormName = SprPurForm Then
          Grid3Sql = "select '' as O,Form_Desc as FormDescription,form_Code as FormCode from TaxForms  where taxforms.trn_type='Purchase' AND Taxforms.Vehicle_YN=0 AND taxforms.FormTrnType <> 0 order by Form_Desc"
          GridInitialise 3, Grid3Sql
        ElseIf GRepFormName = SprSaleForm Then
          Grid3Sql = "select '' as O,Form_Desc as FormDescription,form_Code as FormCode from TaxForms  where taxforms.trn_type='Sale' AND Taxforms.Vehicle_YN=0 And taxforms.FormTrnType <> 0 order by Form_Desc"
          GridInitialise 3, Grid3Sql
        End If
    Case VehPurForm, VehSaleForm     'vijay  (date 17/10/02) for vehicle tax forms
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Pending/All": .RowHeight(List1) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
        End With
          mFirstRow = Date1: FGrid.Row = mFirstRow: mHelpGridNo = 3
          mLastRow = List1
          Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site order by site_desc"
          GridInitialise 1, Grid1Sql
          Grid2Sql = "select '' as O,SubGroup.NAME as PartyName,SubGroup.Subcode as code from SubGroup order by SubGroup.name"
          GridInitialise 2, Grid2Sql
        If GRepFormName = VehPurForm Then
          Grid3Sql = "select '' as O,Form_Desc as FormDescription,form_Code as FormCode from TaxForms  where taxforms.trn_type='Purchase' AND Taxforms.Vehicle_YN=1 and taxforms.FormTrnType <> 0 order by Form_Desc"
          GridInitialise 3, Grid3Sql
        ElseIf GRepFormName = VehSaleForm Then
          Grid3Sql = "select '' as O,Form_Desc as FormDescription,form_Code as FormCode from TaxForms  where taxforms.trn_type='Sale' AND Taxforms.Vehicle_YN=1 and taxforms.FormTrnType <> 0 order by Form_Desc"
          GridInitialise 3, Grid3Sql
        End If
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

Private Sub SprFormUtilize()
On Error GoTo ELoop
Dim mQRY As String, Condstr As String, mQRY1 As String
    RepPrint = True
    'SubRep1 = False
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where SP_Purch.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Purch.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and TaxForms.Form_Code in (" & GridString1 & ")"
     mQRY = "SELECT SP_Purch.RoadPermit_FormCode, SP_Purch.RoadPermit_No, TaxFormStk.RecDate, SP_Purch.V_No,SP_Purch.V_Date, SP_Purch.NET_AMT, SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date, SubGroup.Name,taxforms.form_desc FROM ((SP_Purch LEFT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON SP_Purch.Party_Code = SubGroup.SubCode) LEFT JOIN TaxFormStk ON SP_Purch.Form_Code = TaxFormStk.Form_Code" & _
            "" & Condstr & " and SP_Purch.v_type in ('SXPIC','SXPIR') "

    
  
    RepName = "SprFormUtz"
    
   
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    
    
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VehFormUtilize()
On Error GoTo ELoop
Dim mQRY As String, Condstr As String, mQRY1 As String
    RepPrint = True
    'SubRep1 = False
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Veh_Purch1.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Purch1.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and TaxForms.Form_Code in (" & GridString1 & ")"
     mQRY = "SELECT Veh_Purch1.RoadPermit_FormCode, Veh_Purch1.RoadPermit_No, TaxFormStk.RecDate, Veh_Purch1.V_No,Veh_Purch1.V_Date, Veh_Purch1.Tot_Amount, Veh_Purch1.PBILL_NO, Veh_Purch1.PBILL_Date, SubGroup.Name,taxforms.form_desc FROM ((Veh_Purch1 LEFT JOIN TaxForms ON Veh_Purch1.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON Veh_Purch1.PartyCode = SubGroup.SubCode) LEFT JOIN TaxFormStk ON Veh_Purch1.Form_Code = TaxFormStk.Form_Code" & _
            "" & Condstr & " "
  
    RepName = "VehFormUtz"
  
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
      
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub SprFormReminder()
On Error GoTo ELoop
Dim mQRY As String, Condstr As String, mQRY1 As String
    RepPrint = True
    'SubRep1 = False
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where SP_Sale.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and SP_Sale.Party_Code in (" & GridString1 & ")"
    
    mQRY = "SELECT SP_Sale.Form_Code, SP_Sale.V_No, SP_Sale.V_Date, SP_Sale.TOT_Amt, SubGroup.Name, SubGroup.Add1,SubGroup.Add2, SubGroup.Add3, TaxForms.Form_Desc, City.CityName, SP_Sale.Invoice_DocId" & _
           " FROM ((SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON" & _
           " SP_Sale.Party_Code = SubGroup.SubCode) LEFT JOIN City ON SP_Sale.Site_Code = City.Site_Code" & _
            "" & Condstr & " and SP_Sale.v_type in ('SXSIC','SXSIR') "

    RepName = "SprFormRemind"
  
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    
 '   If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Private Sub VehFormReminder()
On Error GoTo ELoop
Dim mQRY As String, Condstr As String, mQRY1 As String
    RepPrint = True
    'SubRep1 = False
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date3, FGrid.TextMatrix(Date3, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    
    Condstr = " where Veh_Order.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ""
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.PartyCode in (" & GridString1 & ")"
    
    mQRY = "SELECT Veh_Order.Form_Code, Veh_Order.Inv_No, Veh_Order.Inv_Date, Veh_Order.Net_Amount, SubGroup.Name, SubGroup.Add1,SubGroup.Add2, SubGroup.Add3, TaxForms.Form_Desc, City.CityName" & _
           " FROM ((Veh_Order LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON" & _
           " Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN City ON Veh_Order.Ord_SiteCode = City.Site_Code" & _
            "" & Condstr & " "

    RepName = "VehFormRemind"
  
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    
    
    
  If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub



Private Sub Formulas()
On Error GoTo ELoop
Dim I As Integer

Select Case GRepFormName
Case SprFormUtz, VehFormUtz
    For I = 1 To rpt.FormulaFields.Count                    'work done by vijay from SprPartPur Onwards
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
    End Select
    Next
    
Case SprFormRemind, VehFormRemind
    For I = 1 To rpt.FormulaFields.Count                    'work done by vijay on 18/10/02 reminder form
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        Case UCase("REMINDERDATE")
            rpt.FormulaFields(I).TEXT = "'Reminder Date :'+ '" & Format(FGrid.TextMatrix(Date3, 1), "dd/mmm/yyyy") & "' "
    End Select
    Next
Case SprPurForm, SprSaleForm, VehPurForm, VehSaleForm   '********Vijay (date 17/10/02)
    For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
        Case UCase("List1")
            rpt.FormulaFields(I).TEXT = "'For ' + '" & FGrid.TextMatrix(List1, 1) & "' + ' Forms'"
    End Select
    Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub SprPartPurchase()
On Error GoTo ELoop
Dim mQRY As String, Condstr As String
Dim PartyType As Byte

    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    Condstr = " where SP_Stock.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Stock.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    Condstr = Condstr & " and SP_Stock.V_type = '" & SprMrRct & "'"
            
    PartyType = GCn.Execute("select PartyType from SYCTRL").Fields(0).Value



    mQRY = "SELECT Part.Part_Name, SP_Stock.Part_No, " & cIIF("SubGroup.Party_Type=" & PartyType & "", "Sum(sp_Stock.Qty_Rec)", "0") & " AS telcoqty, " & cIIF("SubGroup.Party_Type<>" & PartyType & "", "Sum(sp_Stock.Qty_Rec)", "0") & " AS localqty, " & cIIF("SubGroup.Party_Type=" & PartyType & "", "Sum(SP_Stock.Amount)", "0") & " AS telcoamt, " & cIIF("SubGroup.Party_Type<>" & PartyType & "", "Sum(SP_Stock.Amount)", "0") & " AS localamt " & _
    "FROM (SP_Stock LEFT JOIN SubGroup ON SP_Stock.Party_Code = SubGroup.SubCode) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO and Part.Div_Code = left(SP_Stock.Docid,1) "
        
    mQRY = mQRY + Condstr
    
    mQRY = mQRY + " GROUP BY Part.Part_Name, SP_Stock.Part_No, SubGroup.Party_Type"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SprPartPur"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Public Sub SelGridKeyPressLocal(Txt As Object, FGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef keyascii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If FGrid(Index).Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then Txt.TEXT = "": Exit Sub
    If keyascii = vbKeyEscape Or keyascii = vbKeyReturn Or keyascii = vbKeyDelete Then Exit Sub
        If keyascii = vbKeyBack And Len(Txt.SelText) <> 1 Then
            Txt.SelLength = Len(Txt.SelText) - 1
            FindStr = Txt.SelText
        Else
            FindStr = Txt.SelText + Chr(keyascii)
        End If
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
        Else    'character serach
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
        End If
        keyascii = 0
       If Rst.AbsolutePosition <> adPosEOF And Rst.AbsolutePosition <> adPosBOF Then
            FGrid(Index).CellBackColor = CellBackColLeave
            FGrid(Index).Row = Rst.AbsolutePosition
            FGrid(Index).CellBackColor = CellBackColEnter
            Txt.TEXT = Rst.Fields(FindFldName).Value
            Txt.SelLength = Len(FindStr)
            Txt.left = FGrid(Index).CellLeft + FGrid(Index).left
            Txt.top = FGrid(Index).CellTop + FGrid(Index).top
            If Txt.Visible = False Then
                Txt.Visible = True: Txt.ZOrder 0: Txt.SetFocus: Txt.BackColor = FGrid(Index).CellBackColor
                 Txt.ForeColor = FGrid(Index).CellForeColor: Txt.width = FGrid(Index).CellWidth: Txt.height = FGrid(Index).CellHeight
            End If
       End If
End Sub

Private Sub SprPurFormTax()
On Error GoTo ELoop
Dim mQRY As String, Condstr As String
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If GRepFormName = SprPurForm Then
    
        Condstr = " where SP_Purch.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Purch.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(SP_Purch.site_code,1) in (" & GridString1 & ")"
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and SP_Purch.Party_Code in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and TaxForms.Form_Code in (" & GridString3 & ")"
        If FGrid.TextMatrix(List1, 1) = "Pending" Then
         Condstr = Condstr & " and (SP_Purch.FormNo='' or isnull(SP_Purch.FormNo))"
        End If
        
   mQRY = "SELECT SP_Purch.DocID,SP_Purch.V_No, SP_Purch.V_Date, SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date," & _
               "SubGroup.Name As Party_Name, SP_Purch.NET_AMT, SP_Purch.FormNo, SP_Purch.FormIssRecDate, TaxForms.Form_Desc," & _
               "TaxForms.Form_Code" & _
               " FROM (SP_Purch LEFT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON SubGroup.SubCode=SP_Purch.Party_Code"
   mQRY = mQRY + Condstr
   
   RepName = "SprPurForm"
   
   ElseIf GRepFormName = SprSaleForm Then
   
        Condstr = " where SP_Sale.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and SP_Sale.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(SP_Sale.site_code,1) in (" & GridString1 & ")"
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and SP_Sale.Party_Code in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and TaxForms.Form_Code in (" & GridString3 & ")"
        If FGrid.TextMatrix(List1, 1) = "Pending" Then
         Condstr = Condstr & " and (SP_Sale.FormNo='' or isnull(SP_Sale.FormNo))"
        End If
        
   mQRY = "SELECT SP_Sale.DocID,SP_Sale.V_No,SP_Sale.V_Type,SP_Sale.V_Date, " & _
               "SubGroup.Name As Party_Name, SP_Sale.Total_Amt, SP_Sale.FormNo, SP_Sale.FormIssRecDate, TaxForms.Form_Desc," & _
               "TaxForms.Form_Code" & _
               " FROM (SP_Sale LEFT JOIN TaxForms ON SP_Sale.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON Subgroup.SubCode=SP_Sale.Party_Code"
   mQRY = mQRY + Condstr
   
   RepName = "SprSaleForm"
   
   End If
        
       
    Set RstRep = New ADODB.Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'    RepName = "SprPurForm"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub
Private Sub VehPurSaleFormTax()
On Error GoTo ELoop
Dim mQRY As String, Condstr As String
    RepPrint = True
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Sub
    
    If GRepFormName = VehPurForm Then
        Condstr = " where Veh_Purch1.v_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Purch1.v_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Purch1.site_code,1) in (" & GridString1 & ")"
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_Purch1.PartyCode in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and TaxForms.Form_Code in (" & GridString3 & ")"
        If FGrid.TextMatrix(List1, 1) = "Pending" Then
            Condstr = Condstr & " and (Veh_Purch1.Form_No='' or isnull(Veh_Purch1.Form_No))"
        End If
        
        mQRY = "SELECT Veh_Purch1.V_No, Veh_Purch1.V_Date, Veh_Purch1.PBill_No, Veh_Purch1.PBill_Date," & _
               "Subgroup.Name, Veh_Purch1.Tot_Amount, Veh_Purch1.Form_No, Veh_Purch1.Form_Date, TaxForms.Form_Desc," & _
               "TaxForms.Form_Code" & _
               " FROM (Veh_Purch1 LEFT JOIN TaxForms ON Veh_Purch1.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON SubGroup.SubCode=Veh_Purch1.Partycode"
        mQRY = mQRY + Condstr
        RepName = "VehPurForm"
        
   ElseIf GRepFormName = VehSaleForm Then
        Condstr = " where Veh_Order.Ord_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Veh_Order.Ord_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(Veh_Order.Ord_SiteCode,1) in (" & GridString1 & ")"
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and Veh_Order.PartyCode in (" & GridString2 & ")"
        If Check1(3).Value = Unchecked Then Condstr = Condstr & " and TaxForms.Form_Code in (" & GridString3 & ")"
        If FGrid.TextMatrix(List1, 1) = "Pending" Then
            Condstr = Condstr & " and (Veh_Order.Form_No='' or isnull(Veh_Order.Form_No))"
        End If
        mQRY = "SELECT Veh_Order.OrdDocId,Veh_Order.Ord_VType,Veh_Order.Ord_Date, " & _
               "Veh_Order.Net_Amount,Veh_Order.Form_No,Veh_Order.Form_Date,TaxForms.Form_Desc," & _
               "TaxForms.Form_Code,SubGroup.Name As PartyName" & _
               " FROM (Veh_Order LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN SubGroup ON SubGroup.SubCode=Veh_Order.PartyCode"
        mQRY = mQRY + Condstr
        RepName = "VehSaleForm"
    End If
    Set RstRep = New ADODB.Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'    RepName = "SprPurForm"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

