VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form subtrial 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subtrial Balance"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "SubTrial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10455
   Begin VB.Frame Frame6 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Display Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   645
      Left            =   270
      TabIndex        =   31
      Top             =   3450
      Visible         =   0   'False
      Width           =   3750
      Begin VB.OptionButton OptGrp 
         BackColor       =   &H00CFE0E0&
         Caption         =   "A/c Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2100
         TabIndex        =   33
         Top             =   240
         Width           =   1290
      End
      Begin VB.OptionButton OptLdg 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Ledger A/c"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   360
         TabIndex        =   32
         Top             =   240
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Options"
      ForeColor       =   &H00800080&
      Height          =   510
      Left            =   5640
      TabIndex        =   28
      Top             =   825
      Visible         =   0   'False
      Width           =   3450
      Begin VB.OptionButton OptChoice 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Without Opening"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1665
         TabIndex        =   30
         Top             =   180
         Width           =   1755
      End
      Begin VB.OptionButton OptChoice 
         BackColor       =   &H00CFE0E0&
         Caption         =   "With Opening"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   180
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4215
      TabIndex        =   25
      Top             =   2100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Amount Slab Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   540
      Left            =   5640
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   3975
      Begin VB.OptionButton Option4 
         BackColor       =   &H00CFE0E0&
         Caption         =   "No"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Yes"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   22
         Top             =   225
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CFE0E0&
      Height          =   585
      Left            =   5640
      TabIndex        =   18
      Top             =   1215
      Visible         =   0   'False
      Width           =   2475
      Begin VB.OptionButton Option3 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1230
         TabIndex        =   20
         Top             =   255
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.TextBox TXTE_DATE 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1395
   End
   Begin VB.TextBox TXTS_DATE 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1395
      Width           =   1395
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H8000000A&
      Caption         =   "&Dos Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   6090
      TabIndex        =   5
      ToolTipText     =   "Print Reports"
      Top             =   4965
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox Txt_CL_BAL 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4170
      MaxLength       =   12
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Enter the date from which the data has to be printed."
      Top             =   2085
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox Txt_PR_CL_BAL 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4170
      MaxLength       =   12
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Enter the date from which the data has to be printed."
      Top             =   2400
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   270
      TabIndex        =   9
      Top             =   4185
      Visible         =   0   'False
      Width           =   2910
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Adjustment Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   390
         TabIndex        =   10
         Top             =   300
         Width           =   2085
      End
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H8000000A&
      Caption         =   "E&xit"
      DisabledPicture =   "SubTrial.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8310
      Picture         =   "SubTrial.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   4965
      Width           =   1110
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H8000000A&
      Caption         =   "&Print"
      DisabledPicture =   "SubTrial.frx":0686
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   7200
      Picture         =   "SubTrial.frx":07C8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Reports"
      Top             =   4965
      Width           =   1110
   End
   Begin MSDataListLib.DataCombo TXTACC_CODE 
      Height          =   315
      Left            =   4170
      TabIndex        =   2
      Top             =   2070
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Include Opening Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   285
      TabIndex        =   13
      Top             =   2775
      Visible         =   0   'False
      Width           =   3750
      Begin VB.OptionButton Option1 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00CFE0E0&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1725
         TabIndex        =   14
         Top             =   240
         Width           =   1725
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   1
      Left            =   4185
      TabIndex        =   27
      Top             =   2070
      Visible         =   0   'False
      Width           =   6285
      _ExtentX        =   11086
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
      ScrollBars      =   2
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   810
      TabIndex        =   24
      Top             =   375
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cl.Stock Values (Prev.Year)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   1650
      TabIndex        =   17
      Top             =   2430
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label Label51 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cl.Stock Value (This Yr.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2130
      TabIndex        =   16
      Top             =   2130
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3150
      TabIndex        =   11
      Top             =   2115
      Width           =   885
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   495
      TabIndex        =   8
      Top             =   1470
      Width           =   885
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   1830
      Width           =   765
   End
End
Attribute VB_Name = "subtrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oBAL As Double
Dim Rst3 As ADODB.Recordset, GString As String
Private Const CellBackColLeave As String = &HFFFFFF
Private Const CellBackColEnter As String = &HFFFFC0
Private Const CellBackColLeave1 As String = &HEDF7FE
Private Const CellBackColEnter1 As String = &HFFFFC0
Dim RsGrid1 As ADODB.Recordset

Dim RsGrid2 As ADODB.Recordset
Dim RsGrid3 As ADODB.Recordset
Dim RsGrid4 As ADODB.Recordset

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
Private Sub Check2_Click()
    If Check2.Value = Unchecked Then
        GridSel(1).Enabled = True
        If GridSel(1).Rows > 1 Then
           GridSel(1).Row = 1: GridSel(1).Col = 1
        End If
    Else
        GridSel(1).Enabled = False
        If GridSel(1).Rows > 1 Then
            GridSel(1).Row = 0: GridSel(1).Col = 0
            GridSel(1).RowSel = GridSel(1).Rows - 1
        End If
    End If
End Sub

Private Sub Check2_GotFocus()
Check2.BackColor = &HFF&
End Sub

Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Check2_Validate(Cancel As Boolean)
Check2.BackColor = &H800000
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
Private Sub GridSel_LeaveCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub

Private Sub GridSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If GridSel(Index).Col <> 0 Then Exit Sub
mGridStartRow = GridSel(Index).Row
End Sub

Private Sub GridSel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Private Sub TEXT_GotFocus(Index As Integer)
    SendKeys "{Home}+{End}"
End Sub
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
Private Sub Form_Activate()
Select Case Me.Tag
    Case "7", "8"
'        Call INI_COMBO("select V_TYPE,Description+' Book' AS NAME from Voucher_type WHERE Category='FA' AND V_TYPE<>'F_AO' order by Description", TXTACC_CODE, "NAME", "V_tYPE")
    'modishekhar
        Call INI_COMBO("select V_TYPE,Description+' Book' AS NAME,Category from Voucher_type " & _
        "WHERE (Category='FA' AND V_TYPE<>'F_AO') or " & _
        "(V_TYPE in ('SXPIC','SXPIR','SYSIC','SYSIR','V_PB','V_SB','W_LIC','W_LIR','W_SIC','W_SIR')) " & _
        "order by Description", TXTACC_CODE, "NAME", "V_tYPE")
    Case "27"
        Call INI_COMBO("select SUBCODE,NAME from SubGroup order by name", TXTACC_CODE, "NAME", "SUBCODE")
    Case "5", "30"
        Call INI_COMBO("select SUBCODE,NAME from SubGroup where nature='Cash' order by name", TXTACC_CODE, "NAME", "SUBCODE")
    Case "22", "23", "6"
        Call INI_COMBO("select SUBCODE,NAME from SubGroup where nature='Bank' order by name", TXTACC_CODE, "NAME", "SUBCODE")
    Case "29", "10", "13", "14", "15"
        GridSel(1).Visible = True
        Check2.Visible = True
        GridSel(1).left = 4185
        GridSel(1).top = 2070
        GridSel(1).Height = 2580
        GridSel(1).width = 5775
        GridSel(1).ColWidth(1) = 4500: GridSel(1).ColWidth(2) = 0
        Check2.top = GridSel(1).top + 20
        Check2.left = GridSel(1).left + 30
        Check2.Height = GridSel(1).RowHeight(0) + 20
        Check2.width = 950
        Check2.Value = Checked
        TxtSearch.width = GridSel(1).ColWidth(1) - 50
        Select Case Me.Tag
            Case "29", "10"
                Ini_Grid 1
            Case "13"
                Ini_Grid 2
            Case "14", "15"
                Ini_Grid 1
        End Select
    Case Else
        Call INI_COMBO("select GROUPCODE AS CODE,GROUPNAME AS NAME from ACGROUP order by GROUPname", TXTACC_CODE, "NAME", "CODE")
End Select
If Me.Tag = "11" Or Me.Tag = "12" Then
    Frame5.Visible = True
    Frame5.left = Frame3.left
    Frame5.top = Frame3.top
    Frame3.Visible = False
ElseIf Me.Tag = "17" Or "19" Then
    Frame6.Visible = True
End If
End Sub
Private Sub Form_Load()
    Call WinSetting(Me)
    TXTE_DATE = PubLoginDate
    TXTS_DATE = PubStartDate
End Sub
Private Sub Txt_CL_BAL_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And (KeyAscii < 46 Or KeyAscii > 58) Then KeyAscii = 0
End Sub
Private Sub Txt_PR_CL_BAL_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And (KeyAscii < 46 Or KeyAscii > 58) Then KeyAscii = 0
End Sub
Private Sub BTNPRINT_Click(Index As Integer)
On Error GoTo ErrLoop
Dim mGroup_Rs As ADODB.Recordset, SubGroup_Rs As ADODB.Recordset, Rst1 As ADODB.Recordset
Dim rstTmp As ADODB.Recordset, Age_Rs As ADODB.Recordset
Dim DrAc$, CrAc$, mQRY1$, mQRY$, X11
Dim OP_CR As Double, DR As Double, CR As Double, x As Double, Cr1 As Double, m_Pre_Bal As Double
Dim ARR(7) As Double, Dr1 As Double, Dr2 As Double, Dr3 As Double, Dr4 As Double, Dr5 As Double, Dr6 As Double, Dr7 As Double
Dim Date1 As Date, Date2 As Date, TmpDate As Date, i As Integer, Days As Integer
Dim Tot_Amt As Double, TOT_AMTDR As Double, TOT_AMTCR As Double, MyRs As New ADODB.Recordset, ac_str$
Dim MyOpBal As Double, MyCloBal As Double, MyDrStr$, MyCrStr$
Dim MyRst As New ADODB.Recordset
Dim mNarr1$, mNarr2$, mFlag1 As Boolean, mFlag2 As Boolean, TmpNarr$, Cnt As Integer
Dim db As DAO.Database, QryDef As QueryDef, TrnStartDt$

If Me.Tag = "14" Or Me.Tag = "15" Then
    If VALID_DATE_CHK(TXTE_DATE, "Date Upto") = False Then Exit Sub
Else
    If VALID_DATE(Me) = 0 Then Exit Sub
End If
If TXTACC_CODE.Visible = True Then If IsValid(TXTACC_CODE, Label(2)) = False Then Exit Sub
If Me.Tag = "28" Then
    If Val(Txt_CL_BAL) = 0 Then MsgBox "Vr.No.Required": Txt_CL_BAL.SetFocus: Exit Sub
    If Val(Txt_PR_CL_BAL) = 0 Then MsgBox "Vr.No.Required": Txt_PR_CL_BAL.SetFocus: Exit Sub
    If Val(Txt_CL_BAL) > Val(Txt_PR_CL_BAL) Then MsgBox "Invalid Vr.No.": Txt_CL_BAL.SetFocus: Exit Sub
End If
Select Case Me.Tag
    Case "12"
        mQRY = "SELECT MAX(GROUP_TRANS.TYPE) AS GRTYPE,MAX(GROUP_TRANS.GR_NAME) AS GrName,MAX(GROUP_TRANS.CODE) AS MG_CODE,MAX(GROUP_TRANS.NAME) AS MG_NAME,MAX(PARTY_LIST.SUBCODE) AS SUB_SUBCODE,MAX(PARTY_LIST.NAME) AS SUB_NAME,SUM(DEBIT) AS OP_DR,SUM(CREDIT) AS OP_CR,0 AS DR,0 AS CR FROM (PARTY_LIST LEFT JOIN GROUP_TRANS ON GROUP_TRANS.CODE=PARTY_LIST.CODE) LEFT JOIN ViewLedger ON ViewLedger.PARTY=PARTY_LIST.SUBCODE WHERE V_DATE< #" & CDate(TXTS_DATE) & "# AND GROUP_TRANS.TYPE NOT IN ('E','R') GROUP BY GROUP_TRANS.CODE,PARTY_LIST.SUBCODE " & _
        " Union " & _
        "SELECT MAX(GROUP_TRANS.TYPE) AS GRTYPE,MAX(GROUP_TRANS.GR_NAME) AS GrName,MAX(GROUP_TRANS.CODE) AS MG_CODE,MAX(GROUP_TRANS.NAME) AS MG_NAME,MAX(PARTY_LIST.SUBCODE) AS SUB_SUBCODE,MAX(PARTY_LIST.NAME) AS SUB_NAME,SUM(DEBIT) AS OP_DR,SUM(CREDIT) AS OP_CR,0 AS DR,0 AS CR FROM (PARTY_LIST LEFT JOIN GROUP_TRANS ON GROUP_TRANS.CODE=PARTY_LIST.CODE) LEFT JOIN ViewLedger ON ViewLedger.PARTY=PARTY_LIST.SUBCODE  WHERE V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO' GROUP BY GROUP_TRANS.CODE,PARTY_LIST.SUBCODE " & _
        " Union " & _
        "SELECT MAX(GROUP_TRANS.TYPE) AS GRTYPE,MAX(GROUP_TRANS.GR_NAME) AS GrName,MAX(GROUP_TRANS.CODE) AS MG_CODE,MAX(GROUP_TRANS.NAME) AS MG_NAME,MAX(PARTY_LIST.SUBCODE) AS SUB_SUBCODE,MAX(PARTY_LIST.NAME) AS SUB_NAME,0 AS OP_DR,0 AS OP_CR,SUM(DEBIT) AS DR,SUM(CREDIT) AS CR FROM (PARTY_LIST LEFT JOIN GROUP_TRANS ON GROUP_TRANS.CODE=PARTY_LIST.CODE) LEFT JOIN ViewLedger ON ViewLedger.PARTY=PARTY_LIST.SUBCODE  WHERE V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO' GROUP BY GROUP_TRANS.CODE,PARTY_LIST.SUBCODE"
    
    Case "11"        ' FOR TRIAL BALANCE
        mQRY = "SELECT MAX(GROUP_TRANS.TYPE) AS GRTYP,Max(GROUP_TRANS.GR_CODE) AS GRCODE,Max(GROUP_TRANS.GR_NAME) AS GrName,Max(GROUP_TRANS.CODE) AS ACCODE,Max(GROUP_TRANS.NAME) AS ACNAME,SUM(DEBIT) AS OPDR,SUM(CREDIT) AS OPCR,0 AS DR,0 AS CR FROM ViewLedger LEFT JOIN (PARTY_LIST LEFT JOIN GROUP_TRANS ON PARTY_LIST.CODE = GROUP_TRANS.CODE) ON ViewLedger.party = PARTY_LIST.SUBCODE WHERE V_DATE<#" & CDate(TXTS_DATE) & "# AND GROUP_TRANS.TYPE NOT IN ('E','R') GROUP BY GROUP_TRANS.CODE " & _
        "Union " & _
        "SELECT MAX(GROUP_TRANS.TYPE) AS GRTYP,Max(GROUP_TRANS.GR_CODE) AS GRCODE,Max(GROUP_TRANS.GR_NAME) AS GrName,Max(GROUP_TRANS.CODE) AS ACCODE,Max(GROUP_TRANS.NAME) AS ACNAME,SUM(DEBIT) AS OPDR,SUM(CREDIT) AS OPCR,0 AS DR,0 AS CR FROM ViewLedger LEFT JOIN (PARTY_LIST LEFT JOIN GROUP_TRANS ON PARTY_LIST.CODE = GROUP_TRANS.CODE) ON ViewLedger.party = PARTY_LIST.SUBCODE WHERE  V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO'  GROUP BY GROUP_TRANS.CODE " & _
        "Union " & _
        "SELECT MAX(GROUP_TRANS.TYPE) AS GRTYP,Max(GROUP_TRANS.GR_CODE) AS GRCODE,Max(GROUP_TRANS.GR_NAME) AS GrName,Max(GROUP_TRANS.CODE) AS ACCODE,Max(GROUP_TRANS.NAME) AS ACNAME,0 AS OPDR,0 AS OPCR,SUM(DEBIT) AS DR,SUM(CREDIT) AS CR FROM ViewLedger LEFT JOIN (PARTY_LIST LEFT JOIN GROUP_TRANS ON PARTY_LIST.CODE = GROUP_TRANS.CODE) ON ViewLedger.party = PARTY_LIST.SUBCODE WHERE V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO' GROUP BY GROUP_TRANS.CODE"
            
    Case "20", "21"
        If Me.Tag = "20" Then
            Set mGroup_Rs = GCnFa.Execute("SELECT SUM(ViewLedger.CREDIT) AS TOP_CR,SUM(ViewLedger.DEBIT) AS TOP_DR,MAX(GROUP_TRANS.NAME) AS TNAME,MAX(GROUP_TRANS.gr_name) As TS_NAME From ((ViewLedger LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=ViewLedger.PARTY) LEFT JOIN PARTY_LIST AS PARTY_LIST1 ON PARTY_LIST1.SUBCODE=ViewLedger.PARTY1) LEFT JOIN GROUP_TRANS ON GROUP_TRANS.CODE=PARTY_LIST.CODE WHERE ViewLedger.V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND PARTY_LIST.NATURE IN ('Cash','Bank') OR PARTY_LIST1.NATURE IN ('Cash','Bank') GROUP BY GROUP_TRANS.NAME")
        ElseIf Me.Tag = "21" Then
            Set mGroup_Rs = GCnFa.Execute("SELECT SUM(ViewLedger.CREDIT) AS TOP_CR,SUM(ViewLedger.DEBIT) AS TOP_DR,MAX(GROUP_TRANS.NAME) AS TNAME,MAX(GROUP_TRANS.gr_name) As TS_NAME From ((ViewLedger LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=ViewLedger.PARTY) LEFT JOIN PARTY_LIST AS PARTY_LIST1 ON PARTY_LIST1.SUBCODE=ViewLedger.PARTY1) LEFT JOIN GROUP_TRANS ON GROUP_TRANS.CODE=PARTY_LIST.CODE WHERE ViewLedger.V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# GROUP BY GROUP_TRANS.NAME")
        End If
        Set rstTmp = New ADODB.Recordset
        With rstTmp
            .Fields.Append "OP_CR", adDouble, 19, 5
            .Fields.Append "OP_DR", adDouble, 19, 5
            .Fields.Append "G_NAME", adChar, 35
            .Fields.Append "NAME", adChar, 35
            .Fields.Append "S_NAME", adChar, 35
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        Do Until mGroup_Rs.EOF
            rstTmp.AddNew
            rstTmp!Name = XNull(mGroup_Rs!TNAME)
            rstTmp!S_NAME = XNull(mGroup_Rs!TS_NAME)
            If mGroup_Rs!TOP_DR > mGroup_Rs!TOP_CR Then
                rstTmp!OP_DR = mGroup_Rs!TOP_DR - mGroup_Rs!TOP_CR
                rstTmp!OP_CR = 0
                rstTmp!G_NAME = "Application Of Funds"
            Else
                rstTmp!OP_CR = mGroup_Rs!TOP_CR - mGroup_Rs!TOP_DR
                rstTmp!OP_DR = 0
                rstTmp!G_NAME = "Source Of Funds"
            End If
            rstTmp.Update
            mGroup_Rs.MoveNext
        Loop
    Case "10"       'Annexure
        ac_str = ""
        If Check2.Value = Unchecked Then
            ac_str = FillString(GridRow1, 1, 1)
            If ac_str = "" Then Exit Sub

            mQRY = "SELECT MAX(ACGROUP.GROUPNAME)As GroupName,MAX(PARTY) AS PARTYCODE,MAX(PARTY_NAME) AS PARTYNAME,MAX(ADD1) AS ADDR1,MAX(ADD2) AS ADDR2,MAX(CITY_NAME) AS CITYNAME,SUM(CREDIT) as op_cr,sum(DEBIT) AS OP_dr,0 AS BALANCEdr,0 as balancecr,0 As Bal " & _
                   "FROM (ViewLedger LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=ViewLedger.PARTY)" & _
                   "LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=PARTY_LIST.GROUPCODE " & _
                   "WHERE V_DATE<#" & CDate(TXTS_DATE) & "# AND CODE IN ( " & ac_str & ") " & _
                   "GROUP BY ACGROUP.GROUPNAME,PARTY " & _
                   "Union " & _
                   "SELECT MAX(ACGROUP.GROUPNAME)As GroupName ,MAX(PARTY) AS PARTYCODE,MAX(PARTY_NAME) AS PARTYNAME,MAX(ADD1) AS ADDR1,MAX(ADD2) AS ADDR2,MAX(CITY_NAME) AS CITYNAME,0 as op_cr,0 AS OP_dr,SUM(DEBIT) AS BALANCEdr,sum(credit) as balancecr,sum(credit)-SUM(DEBIT) As Bal " & _
                   "FROM (ViewLedger LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=ViewLedger.PARTY) " & _
                   "LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=PARTY_LIST.GROUPCODE " & _
                   "WHERE V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND CODE IN ( " & ac_str & ") " & _
                   "GROUP BY ACGROUP.GROUPNAME,PARTY"
        Else
            mQRY = "SELECT MAX(ACGROUP.GROUPNAME)As GroupName,MAX(PARTY) AS PARTYCODE,MAX(PARTY_NAME) AS PARTYNAME,MAX(ADD1) AS ADDR1,MAX(ADD2) AS ADDR2,MAX(CITY_NAME) AS CITYNAME,SUM(CREDIT) as op_cr,sum(DEBIT) AS OP_dr,0 AS BALANCEdr,0 as balancecr,0 As Bal " & _
                   "FROM (ViewLedger LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=ViewLedger.PARTY)" & _
                   "LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=PARTY_LIST.GROUPCODE " & _
                   "WHERE V_DATE<#" & CDate(TXTS_DATE) & "# GROUP BY ACGROUP.GROUPNAME,PARTY " & _
                   "Union " & _
                   "SELECT MAX(ACGROUP.GROUPNAME)As GroupName,MAX(PARTY) AS PARTYCODE,MAX(PARTY_NAME) AS PARTYNAME,MAX(ADD1) AS ADDR1,MAX(ADD2) AS ADDR2,MAX(CITY_NAME) AS CITYNAME,0 as op_cr,0 AS OP_dr,SUM(DEBIT) AS BALANCEdr,sum(credit) as balancecr,sum(credit)-SUM(DEBIT) As Bal " & _
                   "FROM (ViewLedger LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=ViewLedger.PARTY) " & _
                   "LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=PARTY_LIST.GROUPCODE " & _
                   "WHERE V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# GROUP BY ACGROUP.GROUPNAME,PARTY"
        End If
    Case "5", "6"       'Cashbook English Format
        DrAc = ""
        CrAc = ""
        oBAL = 0
'        oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM ViewLedger WHERE (V_DATE<#" & CDate(TXTS_DATE) & "# OR (V_DATE=#" & CDate(TXTS_DATE) & "# AND V_TYPE='F_AO')) AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText)).Fields(0)
        TrnStartDt = IIf(IsNull(GCnFa.Execute("select Min(V_Date) from ViewLedger where V_DATE between #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText)).Fields(0)), "", GCnFa.Execute("select Min(V_Date) from ViewLedger where V_DATE between #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText)).Fields(0))
        If TrnStartDt = "" Then
            TrnStartDt = TXTS_DATE
        End If
        oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM ViewLedger WHERE (V_DATE<#" & CDate(TrnStartDt) & "# OR (V_DATE BETWEEN  #" & CDate(TrnStartDt) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO')) AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText)).Fields(0)

        Set rstTmp = New ADODB.Recordset
        Set rstTmp = CASHTMP1(rstTmp)
        If oBAL <> 0 Then
            rstTmp.AddNew
            rstTmp!V_DATE = TrnStartDt
            rstTmp!Name = "OPENING BALANCE"
            If oBAL < 0 Then
                rstTmp!CR = Abs(oBAL)
            Else
                rstTmp!AdjAmt = Abs(oBAL)
            End If
            rstTmp.Update
        End If
        Set SubGroup_Rs = GCnFa.Execute("SELECT ViewLedger.V_NO,SubGroup.NAME, ViewLedger.V_DATE,ViewLedger.DEBIT AS DRAMOUNT, ViewLedger.CREDIT AS CrAmount, ViewLedger.V_TYPE, ViewLedger.NARRATION, ViewLedger.V_SNO,ViewLedger.CHQ_NO,FORMAT(ViewLedger.CHQ_DATE,'DD/MM/YY') AS CHQDATE " & _
            "FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.PARTY1=SubGroup.SUBCODE " & _
            "WHERE ViewLedger.V_DATE Between #" & CDate(TrnStartDt) & "# And #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO' AND (ViewLedger.PARTY=" & Chk_Text(TXTACC_CODE.BoundText) & ") AND CREDIT + DEBIT>0 " & _
            "ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
        If Not (SubGroup_Rs.EOF) Then Date1 = SubGroup_Rs!V_DATE
        mFlag1 = False
        mFlag2 = False
        mNarr1 = ""
        mNarr2 = ""
        TmpDate = TrnStartDt
        Do Until SubGroup_Rs.EOF
            If Date1 = TmpDate Then
                rstTmp.AddNew
                    rstTmp!V_Type = SubGroup_Rs!V_Type
                    rstTmp!V_DATE = Date1
                    rstTmp!v_no = SubGroup_Rs!v_no
                    rstTmp!v_sno = SubGroup_Rs!v_sno
                    mNarr1 = ""
                    If XNull(Trim(SubGroup_Rs!Chq_No)) <> "" Then mNarr1 = mNarr1 + "Ch.No:" + Trim(XNull(SubGroup_Rs!Chq_No)) + " Ch.Dt: " + CStr(SubGroup_Rs!CHQDATE)
                    mNarr1 = mNarr1 + Trim(XNull(SubGroup_Rs!Narration))
                    rstTmp!Name = IIf(IsNull(SubGroup_Rs!Name), "", SubGroup_Rs!Name)
                    rstTmp!NARRATION1 = Trim(mNarr1)
                    If SubGroup_Rs!DrAmount > 0 Then
                        rstTmp!CR = Format(SubGroup_Rs!DrAmount, "0.00")
                        oBAL = oBAL - SubGroup_Rs!DrAmount
                    ElseIf SubGroup_Rs!CrAmount > 0 Then
                        rstTmp!AdjAmt = Format(SubGroup_Rs!CrAmount, "0.00")
                        oBAL = oBAL + SubGroup_Rs!CrAmount
                    End If
'                    If SubGroup_Rs!DrAmount > 0 Then
'                        oBAL = oBAL + Format(SubGroup_Rs!DrAmount, "0.00")
'                    ElseIf SubGroup_Rs!CrAmount > 0 Then
'                        oBAL = oBAL - Format(SubGroup_Rs!CrAmount, "0.00")
'                    End If
                    SubGroup_Rs.MoveNext
                    If Not SubGroup_Rs.EOF Then
                        Date1 = SubGroup_Rs!V_DATE
                    Else
                        Date1 = DateAdd("D", 1, TXTE_DATE)
                    End If
                rstTmp.Update
            Else
                If Date1 <= Date2 Then
                    If Date1 = CDate("12:00:00 AM") Then
                        TmpDate = Date2
                    Else
                        TmpDate = Date1
                    End If
                Else
                    If Date2 = CDate("12:00:00 AM") Then
                        TmpDate = Date1
                    Else
                        TmpDate = Date2
                    End If
                End If
                If oBAL <> 0 Then
                    rstTmp.AddNew
                    rstTmp!V_DATE = TmpDate
                    rstTmp!Name = "OPENING BALANCE"
                    If oBAL < 0 Then
                        rstTmp!CR = Abs(oBAL)
                    Else
                        rstTmp!AdjAmt = Abs(oBAL)
                    End If
                    rstTmp.Update
                End If
            End If
        Loop
        If rstTmp.RecordCount > 0 Then rstTmp.MoveFirst

    Case "6", "30"  'FOR CASH BooK (Long format), Bank Book
    '*******************************************************
        DrAc = ""
        CrAc = ""
        oBAL = 0
'        oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM ViewLedger WHERE (V_DATE<#" & CDate(TXTS_DATE) & "# OR (V_DATE=#" & CDate(TXTS_DATE) & "# AND V_TYPE='F_AO')) AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText)).Fields(0)
        TrnStartDt = GCnFa.Execute("select Min(V_Date) from ViewLedger where V_DATE between #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText)).Fields(0)
        If TrnStartDt = "" Then
            TrnStartDt = TXTS_DATE
        End If
        oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM ViewLedger WHERE (V_DATE<#" & CDate(TrnStartDt) & "# OR (V_DATE BETWEEN  #" & CDate(TrnStartDt) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO')) AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText)).Fields(0)
        Set rstTmp = New ADODB.Recordset
        Set rstTmp = CASHTMP1(rstTmp)
        If oBAL <> 0 Then
            rstTmp.AddNew
            rstTmp!V_DATE = TrnStartDt
            If oBAL < 0 Then
                rstTmp!Name = "OPENING BALANCE"
                rstTmp!CR = Abs(oBAL)
            Else
                rstTmp!NAME1 = "OPENING BALANCE"
                rstTmp!AdjAmt = Abs(oBAL)
            End If
            rstTmp.Update
        End If
        Set mGroup_Rs = GCnFa.Execute("SELECT ViewLedger.V_NO,SubGroup.NAME, ViewLedger.V_DATE, ViewLedger.CREDIT AS AMOUNT, ViewLedger.V_TYPE, ViewLedger.NARRATION, ViewLedger.V_SNO,ViewLedger.CHQ_NO,FORMAT(ViewLedger.CHQ_DATE,'DD/MM/YY') AS CHQDATE " & _
            "FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.PARTY1=SubGroup.SUBCODE " & _
            "WHERE ViewLedger.V_DATE Between #" & CDate(TrnStartDt) & "# And #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO' AND (ViewLedger.PARTY=" & Chk_Text(TXTACC_CODE.BoundText) & ") AND CREDIT>0 " & _
            "ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
        Set SubGroup_Rs = GCnFa.Execute("SELECT ViewLedger.V_NO,SubGroup.NAME, ViewLedger.V_DATE, ViewLedger.DEBIT AS AMOUNT, ViewLedger.V_TYPE, ViewLedger.NARRATION, ViewLedger.V_SNO,ViewLedger.CHQ_NO,FORMAT(ViewLedger.CHQ_DATE,'DD/MM/YY') AS CHQDATE " & _
            "FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.PARTY1=SubGroup.SUBCODE " & _
            "WHERE ViewLedger.V_DATE Between #" & CDate(TrnStartDt) & "# And #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO' AND (ViewLedger.PARTY=" & Chk_Text(TXTACC_CODE.BoundText) & ") AND DEBIT>0 " & _
            "ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
        If Not (mGroup_Rs.EOF) Then Date2 = mGroup_Rs!V_DATE
        If Not (SubGroup_Rs.EOF) Then Date1 = SubGroup_Rs!V_DATE
        mFlag1 = False
        mFlag2 = False
        mNarr1 = ""
        mNarr2 = ""
        TmpDate = TrnStartDt
        Do Until mGroup_Rs.EOF And SubGroup_Rs.EOF
            If Date1 = TmpDate Or Date2 = TmpDate Then
                rstTmp.AddNew
                If Date1 = TmpDate Then
                    rstTmp!V_Type = SubGroup_Rs!V_Type
                    rstTmp!V_DATE = Date1
                    rstTmp!v_no = SubGroup_Rs!v_no
                    rstTmp!v_sno = SubGroup_Rs!v_sno
                    'changed by Avinash
                    Set GRs = GCnFa.Execute("SELECT ledger.*,SubGroup.Name " & _
                        " FROM Ledger INNER JOIN SubGroup ON Ledger.SubCode = SubGroup.SubCode " & _
                        "WHERE LEDGER.V_DATE = #" & SubGroup_Rs!V_DATE & _
                        "# and V_SNo <> " & SubGroup_Rs!v_sno & " AND V_TYPE='" & SubGroup_Rs!V_Type & "' and left(docid,3)&'-'&trim(right(docid,8))='" & SubGroup_Rs!v_no & "'")
                    If GRs.RecordCount > 0 Then
                        Cnt = 1
                        TmpNarr = ""
                        GRs.MoveFirst
                        While Not GRs.EOF
                            TmpNarr = TmpNarr + UCase(IIf(Len(Mid(GRs!Name, 1, 35)) < 35, Mid(GRs!Name, 1, 35) + Space(35 - Len(Mid(GRs!Name, 1, 35))), Mid(GRs!Name, 1, 35))) + Space(1) + Space(11 - Len(CStr(Format(Abs(GRs!AmtCr - GRs!AmtDr), "0.00")))) + CStr(Format(Abs(GRs!AmtCr - GRs!AmtDr), "0.00")) + IIf(GRs!AmtCr - GRs!AmtDr > 0, "Cr", "Dr") + vbCrLf
                            GRs.MoveNext
                            Cnt = Cnt + 1
                        Wend
                        rstTmp!Narration5 = TmpNarr
                    End If
                    'end change
                    If mFlag1 = False Then
                        mFlag1 = True
                        mNarr1 = ""
                        If XNull(Trim(SubGroup_Rs!Chq_No)) <> "" Then mNarr1 = mNarr1 + "Ch.No:" + Trim(XNull(SubGroup_Rs!Chq_No)) + " Ch.Dt: " + CStr(SubGroup_Rs!CHQDATE)
                        'If Not IsNull(SubGroup_Rs!CHQDATE) And SubGroup_Rs!CHQDATE <> "" Then mNARR1 = mNARR1 + " Ch.Dt: " + CStr(SubGroup_Rs!CHQDATE)
                        mNarr1 = mNarr1 + Trim(XNull(SubGroup_Rs!Narration))
                        If CrAc <> SubGroup_Rs!Name Then
                            rstTmp!Name = SubGroup_Rs!Name
                            CrAc = SubGroup_Rs!Name
                        Else
                            rstTmp!Name = Trim(Mid(mNarr1, 1, 29))
                            mNarr1 = Trim(Mid(mNarr1, 30, 100))
                        End If
                        rstTmp!CR = Format(SubGroup_Rs!AMOUNT, "0.00")
                        oBAL = oBAL - Format(SubGroup_Rs!AMOUNT, "0.00")
                        If Len(mNarr1) <= 0 Then
                            mFlag1 = False
                            SubGroup_Rs.MoveNext
                            If Not SubGroup_Rs.EOF Then
                                Date1 = SubGroup_Rs!V_DATE
                            Else
                                Date1 = DateAdd("D", 1, TXTE_DATE)
                            End If
                        End If
                    Else
                        mNarr1 = Trim(mNarr1)
                        rstTmp!Name = Trim(Mid(mNarr1, 1, 29))
                        mNarr1 = Trim(Mid(mNarr1, 30, 100))
                        If Len(mNarr1) <= 0 Then
                            mFlag1 = False
                            SubGroup_Rs.MoveNext
                            If Not SubGroup_Rs.EOF Then
                                Date1 = SubGroup_Rs!V_DATE
                            Else
                                Date1 = DateAdd("D", 1, TXTE_DATE)
                            End If
                        End If
                    End If
                End If
                If Date2 = TmpDate Then
                    rstTmp!Vtype = mGroup_Rs!V_Type
                    rstTmp!VNo = mGroup_Rs!v_no
                    rstTmp!V_DATE = Date2
                    rstTmp!VSNo = mGroup_Rs!v_sno
                    'changed by Avinash
                        Set GRs = GCnFa.Execute("SELECT ledger.*,SubGroup.Name " & _
                            " FROM Ledger INNER JOIN SubGroup ON Ledger.SubCode = SubGroup.SubCode " & _
                            "WHERE LEDGER.V_DATE = #" & mGroup_Rs!V_DATE & _
                            "# and V_SNo <> " & mGroup_Rs!v_sno & " AND V_TYPE='" & mGroup_Rs!V_Type & "' and left(docid,3)&'-'&trim(right(docid,8))='" & mGroup_Rs!v_no & "'")
                        If GRs.RecordCount > 0 Then
                            Cnt = 1
                            TmpNarr = ""
                            GRs.MoveFirst
                            While Not GRs.EOF
                                TmpNarr = TmpNarr + UCase(IIf(Len(Mid(GRs!Name, 1, 35)) < 35, Mid(GRs!Name, 1, 35) + Space(35 - Len(Mid(GRs!Name, 1, 35))), Mid(GRs!Name, 1, 35))) + Space(1) + Space(11 - Len(CStr(Format(Abs(GRs!AmtCr - GRs!AmtDr), "0.00")))) + CStr(Format(Abs(GRs!AmtCr - GRs!AmtDr), "0.00")) + IIf(GRs!AmtCr - GRs!AmtDr > 0, "Cr", "Dr") + vbCrLf
                                GRs.MoveNext
                                Cnt = Cnt + 1
                            Wend
                            rstTmp!Narration4 = TmpNarr
'                            MsgBox Len(!Narration3)
                        End If

                    'end change
                    If mFlag2 = False Then
                        mFlag2 = True
                        mNarr2 = ""
                        If XNull(Trim(mGroup_Rs!Chq_No)) <> "" Then mNarr2 = mNarr2 + "Ch.No:" + Trim(XNull(mGroup_Rs!Chq_No)) + " Ch.Dt: " + CStr(mGroup_Rs!CHQDATE)
                        'If Not IsNull(mGROUP_rs!CHQDATE) And mGROUP_rs!CHQDATE <> "" Then mNARR2 = mNARR2 + " Ch.Dt: " + CStr(mGROUP_rs!CHQDATE)
                        mNarr2 = mNarr2 + Trim(XNull(mGroup_Rs!Narration))
                        If DrAc <> mGroup_Rs!Name Then
                            rstTmp!NAME1 = mGroup_Rs!Name
                            DrAc = mGroup_Rs!Name
                        Else
                            rstTmp!NAME1 = Trim(Mid(mNarr2, 1, 29))
                            mNarr2 = Trim(Mid(mNarr2, 30, 100))
                        End If
                        rstTmp!AdjAmt = Format(mGroup_Rs!AMOUNT, "0.00")
                        oBAL = oBAL + Format(mGroup_Rs!AMOUNT, "0.00")
                        If Len(mNarr2) <= 0 Then
                            mFlag2 = False
                            mGroup_Rs.MoveNext
                            If Not mGroup_Rs.EOF Then
                                Date2 = mGroup_Rs!V_DATE
                            Else
                                Date2 = DateAdd("D", 1, TXTE_DATE)
                            End If
                        End If
                    Else
                        mNarr2 = Trim(mNarr2)
                        rstTmp!NAME1 = Trim(Mid(mNarr2, 1, 29))
                        mNarr2 = Trim(Mid(mNarr2, 30, 100))
                        If Len(mNarr2) <= 0 Then
                            mFlag2 = False
                            mGroup_Rs.MoveNext
                            If Not mGroup_Rs.EOF Then
                                Date2 = mGroup_Rs!V_DATE
                            Else
                                Date2 = DateAdd("D", 1, TXTE_DATE)
                            End If
                        End If
                    End If
                End If
                rstTmp.Update
            Else
                If Date1 <= Date2 Then
                    If Date1 = CDate("12:00:00 AM") Then
                        TmpDate = Date2
                    Else
                        TmpDate = Date1
                    End If
                Else
                    If Date2 = CDate("12:00:00 AM") Then
                        TmpDate = Date1
                    Else
                        TmpDate = Date2
                    End If
                End If
                If oBAL <> 0 Then
                    rstTmp.AddNew
                    rstTmp!V_DATE = TmpDate
                    If oBAL < 0 Then
                        rstTmp!Name = "OPENING BALANCE"
                        rstTmp!CR = Abs(oBAL)
                    Else
                        rstTmp!NAME1 = "OPENING BALANCE"
                        rstTmp!AdjAmt = Abs(oBAL)
                    End If
                    rstTmp.Update
                End If
            End If
        Loop
        If rstTmp.RecordCount > 0 Then rstTmp.MoveFirst
        
    '********************************************************
        
    Case "13"   'FOR BANK REGISTER
        If Check2.Value = Unchecked Then
            ac_str = FillString(GridRow1, 1, 1)
            If ac_str = "" Then Exit Sub

            mQRY = " SELECT " & _
                    " MAX(ViewLedger.PARTY) AS PARTY_cODE,MAX(SubGroup.NAME) AS PARTY_NAME, " & ConvertDate(DateAdd("D", -1, TXTS_DATE)) & " AS V_DATE,Sum(DEBIT) AS DR,Sum(CREDIT) AS CR,'' AS v_type, 0 AS v_no,'' AS v_add,'' AS CHQ_NO,'' AS CHQ_DATE,'' AS CLG_DATE,'' AS NARRATION,'Op.Balance' AS Name1 " & _
                    " FROM ViewLedger " & _
                    " LEFT JOIN SubGroup ON SubGroup.SUBCODE = ViewLedger.PARTY " & _
                    " WHERE (V_DATE<#" & CDate(TXTS_DATE) & "# OR (V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_tYPE='F_AO')) AND ViewLedger.PARTY In ( " & ac_str & ") GROUP BY PARTY " & _
                    " Union " & _
                    " SELECT " & _
                    " ViewLedger.PARTY AS PARTY_CODE,SubGroup.NAME AS PARTY_NAME,V_DATE,DEBIT AS DR,CREDIT AS CR,v_type,v_no,v_add,CHQ_NO,FORMAT(CHQ_DATE,'DD/MM/YY'),FORMAT(CLG_DATE,'DD/MM/YY'),NARRATION,SubGroup1.NAME AS NAME1 " & _
                    " FROM (ViewLedger " & _
                    " LEFT JOIN SubGroup ON SubGroup.SUBCODE = ViewLedger.PARTY)" & _
                    " LEFT JOIN SubGroup SubGroup1 ON ViewLedger.party1=SubGroup1.SUBCODE " & _
                    " WHERE ViewLedger.PARTY In (" & ac_str & " ) AND V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_tYPE<>'F_AO'"
        Else
            mQRY = " SELECT " & _
                    " MAX(ViewLedger.PARTY) AS PARTY_cODE,MAX(SubGroup.NAME) AS PARTY_NAME, " & ConvertDate(DateAdd("D", -1, TXTS_DATE)) & " AS V_DATE,Sum(DEBIT) AS DR,Sum(CREDIT) AS CR,'' AS v_type, 0 AS v_no,'' AS v_add,'' AS CHQ_NO,'' AS CHQ_DATE,'' AS CLG_DATE,'' AS NARRATION,'Op.Balance' AS Name1 " & _
                    " FROM ViewLedger " & _
                    " LEFT JOIN SubGroup ON SubGroup.SUBCODE = ViewLedger.PARTY " & _
                    " WHERE (V_DATE<#" & CDate(TXTS_DATE) & "# OR (V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_tYPE='F_AO')) AND SubGroup.Nature='Bank' GROUP BY PARTY " & _
                    " Union " & _
                    " SELECT " & _
                    " ViewLedger.PARTY AS PARTY_CODE,SubGroup.NAME AS PARTY_NAME,V_DATE,DEBIT AS DR,CREDIT AS CR,v_type,v_no,v_add,CHQ_NO,FORMAT(CHQ_DATE,'DD/MM/YY'),FORMAT(CLG_DATE,'DD/MM/YY'),NARRATION,SubGroup1.NAME AS NAME1 " & _
                    " FROM (ViewLedger " & _
                    " LEFT JOIN SubGroup ON SubGroup.SUBCODE = ViewLedger.PARTY)" & _
                    " LEFT JOIN SubGroup SubGroup1 ON ViewLedger.party1=SubGroup1.SUBCODE " & _
                    " WHERE SubGroup.Nature='Bank' AND V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_tYPE<>'F_AO'"
        
        End If
        
    Case "14", "15"  'AGEING ANALYSIS
        Set Age_Rs = GCnFa.Execute("SELECT * FROM AGEPARAMFA")
        If Age_Rs.RecordCount = 0 Then MsgBox " ** Please Set Ageing Parameters ** ", vbInformation, Me.Caption:      Exit Sub
        If Me.Tag = "14" Then mQRY1 = " AcGroup.Nature='Customer' "
        If Me.Tag = "15" Then mQRY1 = " AcGroup.Nature='Supplier' "
        
        ac_str = ""
        If Check2.Value = Unchecked Then
            ac_str = FillString(GridRow1, 1, 1)
            If ac_str = "" Then Exit Sub

            Set SubGroup_Rs = GCnFa.Execute("SELECT SubGroup.SUBCODE,SubGroup.GROUPCODE AS CODE,SubGroup.NAME,ACGROUP.GroupName FROM SubGroup LEFT JOIN ACGROUP ON SubGroup.GROUPCODE=ACGROUP.GROUPCODE WHERE ACGROUP.GROUPCODE In (" & ac_str & ") And " & mQRY1 & " ORDER BY SubGroup.NAME")
        Else
            Set SubGroup_Rs = GCnFa.Execute("SELECT SubGroup.SUBCODE,SubGroup.GROUPCODE AS CODE,SubGroup.NAME,ACGROUP.GroupName FROM SubGroup LEFT JOIN ACGROUP ON SubGroup.GROUPCODE=ACGROUP.GROUPCODE Where " & mQRY1 & " ORDER BY SubGroup.NAME")
        End If
        
        Set rstTmp = New ADODB.Recordset
        Set rstTmp = AGETMP(rstTmp)
        Label1.Visible = True
        Label1.Caption = ""

        Do Until SubGroup_Rs.EOF
            Dr1 = 0
            Dr2 = 0
            Dr3 = 0
            Dr4 = 0
            Dr5 = 0
            Dr6 = 0
            Dr7 = 0
            DR = 0
            CR = 0
            Set mGroup_Rs = GCnFa.Execute("SELECT V_DATE,DEBIT,CREDIT,PARTY AS SUBCODE FROM ViewLedger WHERE  " & Chk_Text(SubGroup_Rs!SubCode) & "=ViewLedger.PARTY AND (V_DATE<=#" & CDate(TXTE_DATE) & "#)")
            Label1.Caption = SubGroup_Rs!Name
            Do Until mGroup_Rs.EOF
                
                If Me.Tag = "14" Then
                    If mGroup_Rs!credit > 0 Then
                        CR = mGroup_Rs!credit + CR
                    Else
                        Days = DateDiff("D", mGroup_Rs!V_DATE, TXTE_DATE)
                        If Days <= Age_Rs!p1 Then
                            Dr1 = Dr1 + mGroup_Rs!Debit
                        ElseIf Days > Age_Rs!p1 And Days <= Age_Rs!p2 Then
                            Dr2 = Dr2 + mGroup_Rs!Debit
                        ElseIf Days > Age_Rs!p2 And Days <= Age_Rs!p3 Then
                            Dr3 = Dr3 + mGroup_Rs!Debit
                        ElseIf Days > Age_Rs!p3 And Days <= Age_Rs!p4 Then
                            Dr4 = Dr4 + mGroup_Rs!Debit
                        ElseIf Days > Age_Rs!p4 And Days <= Age_Rs!p5 Then
                            Dr5 = Dr5 + mGroup_Rs!Debit
                        ElseIf Days > Age_Rs!p5 And Days <= Age_Rs!p6 Then
                            Dr6 = Dr6 + mGroup_Rs!Debit
                        Else
                            Dr7 = Dr7 + mGroup_Rs!Debit
                        End If
                    End If
                ElseIf Me.Tag = "15" Then
                    If mGroup_Rs!Debit > 0 Then
                        CR = mGroup_Rs!Debit + CR
                    Else
                        Days = DateDiff("D", mGroup_Rs!V_DATE, TXTE_DATE)
                        If Days <= Age_Rs!p1 Then
                            Dr1 = Dr1 + mGroup_Rs!credit
                        ElseIf Days > Age_Rs!p1 And Days <= Age_Rs!p2 Then
                            Dr2 = Dr2 + mGroup_Rs!credit
                        ElseIf Days > Age_Rs!p2 And Days <= Age_Rs!p3 Then
                            Dr3 = Dr3 + mGroup_Rs!credit
                        ElseIf Days > Age_Rs!p3 And Days <= Age_Rs!p4 Then
                            Dr4 = Dr4 + mGroup_Rs!credit
                        ElseIf Days > Age_Rs!p4 And Days <= Age_Rs!p5 Then
                            Dr5 = Dr5 + mGroup_Rs!credit
                        ElseIf Days > Age_Rs!p5 And Days <= Age_Rs!p6 Then
                            Dr6 = Dr6 + mGroup_Rs!credit
                        Else
                            Dr7 = Dr7 + mGroup_Rs!credit
                        End If
                    End If
                End If
                mGroup_Rs.MoveNext
            Loop
            ARR(0) = Dr1
            ARR(1) = Dr2
            ARR(2) = Dr3
            ARR(3) = Dr4
            ARR(4) = Dr5
            ARR(5) = Dr6
            ARR(6) = Dr7
            x = 7
            Do While x <> 0
                If ARR(x - 1) > 0 Then
                    If CR >= ARR(x - 1) Then
                        CR = CR - ARR(x - 1)
                        ARR(x - 1) = 0
                    ElseIf ARR(x - 1) > CR Then
                        ARR(x - 1) = ARR(x - 1) - CR
                        CR = 0
                    End If
                End If
                x = x - 1
            Loop
            x = ARR(0) + ARR(1) + ARR(2) + ARR(3) + ARR(4) + ARR(5) + ARR(6)
            If Not ((x = 0) And (CR = 0)) Then
                With rstTmp
                    .AddNew
                    !DEBIT1 = ARR(0)
                    !DEBIT2 = ARR(1)
                    !DEBIT3 = ARR(2)
                    !DEBIT4 = ARR(3)
                    !DEBIT5 = ARR(4)
                    !DEBIT6 = ARR(5)
                    !Debit = ARR(6)
                    !TOTALDR = ARR(0) + ARR(1) + ARR(2) + ARR(3) + ARR(4) + ARR(5) + ARR(6)
                    !credit = CR
                    !ACC_NAME = SubGroup_Rs!Name
                    !ANAME = SubGroup_Rs!GroupName
                    .Update
                End With
            End If
            SubGroup_Rs.MoveNext
        Loop
        If rstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\AGEING.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\AGEING.RPT")
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "title"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "ACNAME"
                    rpt.FormulaFields(i).Text = "'For A/C  : " & TXTACC_CODE.Text & "'"
                Case "DATE"
                    rpt.FormulaFields(i).Text = "'Upto Date : " & TXTE_DATE & "'"
                Case "P1"
                    rpt.FormulaFields(i).Text = "'" & "0 - " + str(Age_Rs!p1) & "'"
                Case "P2"
                    rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p1 + 1) & " - " & str(Age_Rs!p2) & "'"
                Case "P3"
                    rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p2 + 1) + " - " + str(Age_Rs!p3) & "'"
                Case "P4"
                    rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p3 + 1) + " - " + str(Age_Rs!p4) & "'"
                Case "P5"
                    rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p4 + 1) + " - " + str(Age_Rs!p5) & "'"
                Case "P6"
                    rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p5 + 1) + " - " + str(Age_Rs!p6) & "'"
                Case "P7"
                    rpt.FormulaFields(i).Text = "'Above " & str(Age_Rs!p6) & "'"
                Case "P8"
                    If Me.Tag = "14" Then
                        rpt.FormulaFields(i).Text = "'Total Debit'"
                    ElseIf Me.Tag = "15" Then
                        rpt.FormulaFields(i).Text = "'Total Credit'"
                    End If
                Case "P9"
                    If Me.Tag = "14" Then
                        rpt.FormulaFields(i).Text = "'Total Credit'"
                    ElseIf Me.Tag = "15" Then
                        rpt.FormulaFields(i).Text = "'Total Debit'"
                    End If
                Case "headi"
                    If Me.Tag = "14" Then
                        rpt.FormulaFields(i).Text = "'<----------------------------- AMOUNT DEBITED FROM DAYS ------------------------------>'"
                    ElseIf Me.Tag = "15" Then
                        rpt.FormulaFields(i).Text = "'<----------------------------- AMOUNT CREDITED FROM DAYS ------------------------------>'"
                    End If
            End Select
        Next
        rpt.Database.SetDataSource rstTmp
    Case "17"   'Profit & Loss A/c
        Set rstTmp = New ADODB.Recordset
        Set rstTmp = TEMPSTAT(rstTmp)
        mQRY1 = ""
        If Option1.Value = True Then
            mQRY = " L.V_DATE<=" & ConvertDate(CDate(TXTE_DATE))
            mQRY1 = " AND V_DATE<=" & ConvertDate(CDate(TXTE_DATE))
        Else
            mQRY = " L.V_DATE BETWEEN " & ConvertDate(CDate(TXTS_DATE)) & " AND " & ConvertDate(CDate(TXTE_DATE))
            mQRY1 = " AND V_DATE BETWEEN " & ConvertDate(CDate(TXTS_DATE)) & " AND " & ConvertDate(CDate(TXTE_DATE))
        End If
        CR = 0: DR = 0: Dr1 = 0: Cr1 = 0
        If OptLdg.Value Then
            GSQL = "select IIF(ISNULL(SUM(L.AmtDr)),0,SUM(L.AmtDr))-IIF(ISNULL(SUM(L.AmtCr)),0,SUM(L.AmtCr)) AS Tot_Amt," & _
                    "MAX(L.SubCode) AS SubCode,MAX(SG.NAME) AS NAME,max(SG.GroupNature) as GroupNature " & _
                    "FROM Ledger as L LEFT JOIN SubGroup as SG ON L.SubCode=SG.SubCode " & _
                    "WHERE " & mQRY & " and SG.GroupNature IN ('E','R')  GROUP BY L.SubCode,SG.Name"
            Set mGroup_Rs = GCnFa.Execute(GSQL)
            Do Until mGroup_Rs.EOF
                If mGroup_Rs!Tot_Amt <> 0 Then
                    If mGroup_Rs!GroupNature = "E" Then
                        With rstTmp
                            .AddNew
                            !Code = mGroup_Rs!SubCode
                            !Name = XNull(mGroup_Rs!Name)
                            !S_NAME = ""
                            !Grp_Code = 1
                            !OP_DR = VNull(mGroup_Rs!Tot_Amt)
                            !G_NAME = "Expenditure"
                            .Update
                        End With
                        DR = DR + mGroup_Rs!Tot_Amt
                    ElseIf mGroup_Rs!GroupNature = "R" Then
                        With rstTmp
                            .AddNew
                            !Code = mGroup_Rs!SubCode
                            !Name = XNull(mGroup_Rs!Name)
                            !S_NAME = ""
                            !Grp_Code = 1
                            !OP_DR = VNull(-mGroup_Rs!Tot_Amt)
                            !G_NAME = "Revenue"
                            .Update
                        End With
                        CR = CR + mGroup_Rs!Tot_Amt
                    End If
                End If
                mGroup_Rs.MoveNext
            Loop
        Else
            GSQL = "select IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT))-IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) AS Tot_Amt,MAX(GROUP_TRANS.CODE) AS CODE,MAX(GROUP_TRANS.NAME) AS NAME,MAX(gr_NAME) AS GrName,MAX(GROUP_TRANS.TYPE) AS TYPE,MAX(GROUP_TRANS.LASTYEARBALANCE) As LAST_BAL " & _
                    "FROM ViewLedger_GROUP LEFT JOIN GROUP_TRANS ON GROUP_TRANS.CODE=ViewLedger_GROUP.CODE " & _
                    "WHERE ViewLedger_GROUP.GroupNature IN ('E','R') " & mQRY1 & " GROUP BY GROUP_TRANS.name"
            Set mGroup_Rs = GCnFa.Execute(GSQL)
            Do Until mGroup_Rs.EOF
                If mGroup_Rs!LAST_BAL <> 0 Or mGroup_Rs!Tot_Amt <> 0 Then
                    If mGroup_Rs!Type = "E" Then
                        With rstTmp
                            .AddNew
                            !Code = mGroup_Rs!Code
                            !Name = XNull(mGroup_Rs!Name)
                            !S_NAME = XNull(mGroup_Rs!GrName)
                            !Grp_Code = 1
                            !OP_DR = VNull(mGroup_Rs!Tot_Amt)
                            !G_NAME = "Expenditure"
                            .Update
                        End With
                        DR = DR + mGroup_Rs!Tot_Amt
                    ElseIf mGroup_Rs!Type = "R" Then
                        With rstTmp
                            .AddNew
                            !Code = mGroup_Rs!Code
                            !Name = XNull(mGroup_Rs!Name)
                            !S_NAME = XNull(mGroup_Rs!GrName)
                            !Grp_Code = 1
                            !OP_DR = VNull(-mGroup_Rs!Tot_Amt)
                            !G_NAME = " Revenue"
                            .Update
                        End With
                        CR = CR + mGroup_Rs!Tot_Amt
                    End If
                End If
                mGroup_Rs.MoveNext
            Loop
        End If
        If Val(Txt_CL_BAL.Text) > 0 Or Val(Txt_PR_CL_BAL.Text) > 0 Then
            With rstTmp
                .AddNew
                !OP_DR = Val(Txt_CL_BAL.Text)
                !LAST_DR = Val(Txt_PR_CL_BAL.Text)
                !G_NAME = " Revenue"
                !Name = "Closing Stock"
                !Grp_Code = 1
                .Update
            End With
        End If
        DR = Abs(DR)
        CR = Abs(CR) + Val(Txt_CL_BAL.Text)
        Dr1 = Abs(Dr1)
        Cr1 = Abs(Cr1) + Val(Txt_PR_CL_BAL.Text)
    Case "19"   'Balance Sheet
        Set rstTmp = New ADODB.Recordset
        Set rstTmp = TEMPSTAT(rstTmp)
        mQRY1 = ""
        If Option1.Value = True Then
            mQRY1 = " WHERE V_DATE<=" & ConvertDate(CDate(TXTE_DATE))
        Else
            mQRY1 = " WHERE V_DATE BETWEEN " & ConvertDate(CDate(TXTS_DATE)) & " AND " & ConvertDate(CDate(TXTE_DATE))
        End If
        CR = 0: DR = 0: Dr1 = 0: Cr1 = 0
        If OptLdg.Value Then
            GSQL = "select IIF(ISNULL(SUM(L.AmtDr)),0,SUM(L.AmtDr))-IIF(ISNULL(SUM(L.AmtCr)),0,SUM(L.AmtCr)) AS Tot_Amt," & _
                    "MAX(L.SubCode) AS SubCode,MAX(SG.NAME) AS NAME,max(SG.GroupNature) as GroupNature " & _
                    "FROM Ledger as L LEFT JOIN SubGroup as SG ON L.SubCode=SG.SubCode " & mQRY1 & " GROUP BY L.SubCode,SG.Name"
            Set mGroup_Rs = GCnFa.Execute(GSQL)
            Do Until mGroup_Rs.EOF
                TOT_AMTDR = 0
                TOT_AMTCR = 0
                Tot_Amt = 0
                If mGroup_Rs!Tot_Amt > 0 Then
                    TOT_AMTDR = TOT_AMTDR + Abs(mGroup_Rs!Tot_Amt)
                End If
                If mGroup_Rs!Tot_Amt < 0 Then
                    TOT_AMTCR = TOT_AMTCR + Abs(mGroup_Rs!Tot_Amt)
                End If
                If mGroup_Rs!GroupNature = "A" Then
                    If TOT_AMTDR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = XNull(mGroup_Rs!Name)
                            !G_NAME = IIf(XNull(mGroup_Rs!GroupNature) = "A", "Assets", "Liabilities")
                            !S_NAME = ""
                            !Grp_Code = 1
                            !OP_DR = Abs(TOT_AMTDR)
                            .Update
                        End With
                    End If
                    If TOT_AMTCR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = Trim(XNull(mGroup_Rs!Name)) + " **"
                            !G_NAME = IIf(XNull(mGroup_Rs!GroupNature) = "A", "Assets", "Liabilities")
                            !S_NAME = ""
                            !Grp_Code = 1
                            !OP_DR = -Abs(TOT_AMTCR)
                            .Update
                        End With
                    End If
                ElseIf mGroup_Rs!GroupNature = "L" Then
                    If TOT_AMTDR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = Trim(XNull(mGroup_Rs!Name)) + " **"
                            !G_NAME = IIf(XNull(mGroup_Rs!GroupNature) = "A", "Assets", "Liabilities")
                            !S_NAME = ""
                            !Grp_Code = 1
                            !OP_DR = -Abs(TOT_AMTDR)
                            .Update
                        End With
                    End If
                    If TOT_AMTCR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = XNull(mGroup_Rs!Name)
                            !G_NAME = IIf(XNull(mGroup_Rs!GroupNature) = "A", "Assets", "Liabilities")
                            !S_NAME = ""
                            !Grp_Code = 1
                            !OP_DR = Abs(TOT_AMTCR)
                            .Update
                        End With
                    End If
                ElseIf mGroup_Rs!GroupNature = "E" Then
                    DR = DR + TOT_AMTDR - TOT_AMTCR
                ElseIf mGroup_Rs!GroupNature = "R" Then
                    CR = CR + TOT_AMTCR - TOT_AMTDR
                End If
                mGroup_Rs.MoveNext
            Loop
        Else
            Set mGroup_Rs = GCnFa.Execute("SELECT GROUP_TRANS.CODE AS CODE,GROUP_TRANS.NAME AS NAME,gr_NAME AS GrName,GROUP_TRANS.TYPE AS TYPE FROM GROUP_TRANS ORDER BY GROUP_TRANS.NAME")
            Do Until mGroup_Rs.EOF
                TOT_AMTDR = 0
                TOT_AMTCR = 0
                Tot_Amt = 0
                Set SubGroup_Rs = GCnFa.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT))-IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) AS Tot_Amt FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.PARTY=SubGroup.SUBCODE " & mQRY1 & " AND SubGroup.GROUPCODE=" & Chk_Text(mGroup_Rs!Code) & " GROUP BY ViewLedger.PARTY")
                Do Until SubGroup_Rs.EOF
                    If SubGroup_Rs!Tot_Amt > 0 Then
                        TOT_AMTDR = TOT_AMTDR + Abs(SubGroup_Rs!Tot_Amt)
                    End If
                    If SubGroup_Rs!Tot_Amt < 0 Then
                        TOT_AMTCR = TOT_AMTCR + Abs(SubGroup_Rs!Tot_Amt)
                    End If
                    SubGroup_Rs.MoveNext
                Loop
                If mGroup_Rs!Type = "A" Then
                    If TOT_AMTDR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = XNull(mGroup_Rs!Name)
                            !G_NAME = IIf(XNull(mGroup_Rs!Type) = "A", "Assets", "Liabilities")
                            !S_NAME = XNull(mGroup_Rs!GrName)
                            !Grp_Code = 1
                            !OP_DR = Abs(TOT_AMTDR)
                            .Update
                        End With
                    End If
                    If TOT_AMTCR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = Trim(XNull(mGroup_Rs!Name)) + " **"
                            !G_NAME = IIf(XNull(mGroup_Rs!Type) = "A", "Assets", "Liabilities")
                            !S_NAME = XNull(mGroup_Rs!GrName)
                            !Grp_Code = 1
                            !OP_DR = -Abs(TOT_AMTCR)
                            .Update
                        End With
                    End If
                ElseIf mGroup_Rs!Type = "L" Then
                    If TOT_AMTDR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = Trim(XNull(mGroup_Rs!Name)) + " **"
                            !G_NAME = IIf(XNull(mGroup_Rs!Type) = "A", "Assets", "Liabilities")
                            !S_NAME = XNull(mGroup_Rs!GrName)
                            !Grp_Code = 1
                            !OP_DR = -Abs(TOT_AMTDR)
                            .Update
                        End With
                    End If
                    If TOT_AMTCR <> 0 Then
                        With rstTmp
                            .AddNew
                            !Name = XNull(mGroup_Rs!Name)
                            !G_NAME = IIf(XNull(mGroup_Rs!Type) = "A", "Assets", "Liabilities")
                            !S_NAME = XNull(mGroup_Rs!GrName)
                            !Grp_Code = 1
                            !OP_DR = Abs(TOT_AMTCR)
                            .Update
                        End With
                    End If
                ElseIf mGroup_Rs!Type = "E" Then
                    DR = DR + TOT_AMTDR - TOT_AMTCR
                ElseIf mGroup_Rs!Type = "R" Then
                    CR = CR + TOT_AMTCR - TOT_AMTDR
                End If
                mGroup_Rs.MoveNext
            Loop
        End If
        If Val(Txt_CL_BAL.Text) > 0 Or Val(Txt_PR_CL_BAL.Text) > 0 Then
            With rstTmp
                .AddNew
                !OP_DR = Val(Txt_CL_BAL.Text)
                !LAST_DR = Val(Txt_PR_CL_BAL.Text)
                !G_NAME = "Assets"
                !S_NAME = "Closing Stock"
                !Grp_Code = 1
                .Update
            End With
        End If
        DR = Abs(DR)
        CR = Abs(CR) + Val(Txt_CL_BAL.Text)
        Dr1 = Abs(Dr1)
        Cr1 = Abs(Cr1) + Val(Txt_PR_CL_BAL.Text)
    Case "27"
        mQRY = "SELECT 'OP' AS V_TYPE,0 AS V_NO,' ' AS V_ADD,0 AS V_SNO,'OP.Balance' AS PARTY, SUM(CREDIT)-SUM(DEBIT) AS OPBAL,0 AS DEB,0 AS CRED,'' AS V_DATE from ViewLedger WHERE PARTY=" & Chk_Text(TXTACC_CODE.BoundText) & " AND (V_DATE < #" & CDate(TXTS_DATE) & "# OR (V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO')) UNION " & _
        "SELECT V_TYPE,V_NO,V_ADD,V_SNO,PARTY,0 AS OPBAL,DEBIT AS DEB,CREDIT AS CRED,V_DATE from ViewLedger WHERE PARTY=" & Chk_Text(TXTACC_CODE.BoundText) & " AND V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO'"
End Select

Select Case Me.Tag
    Case "12"
        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\SUBTRIAL.ttx", True)
        If OptChoice(0).Value = True Then
            Set rpt = rdApp.OpenReport(PubRepoPath + "\New_SUBTRIALOp.RPT")
        Else
            Set rpt = rdApp.OpenReport(PubRepoPath + "\New_SUBTRIAL.RPT")
        End If
        rpt.Database.SetDataSource Rst1
        rpt.FormulaFields(1).Text = "'From   : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
    Case "11"
        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\TRIAL.ttx", True)
        If OptChoice(0).Value = True Then
            Set rpt = rdApp.OpenReport(PubRepoPath + "\New_TrialOp.RPT")
        Else
            Set rpt = rdApp.OpenReport(PubRepoPath + "\New_Trial.RPT") ' "\TrialDet.RPT")
        End If
        rpt.Database.SetDataSource Rst1
        rpt.FormulaFields(4).Text = "'" & Me.Caption + Space(1) + IIf(Option3(1).Value = True, "Summary", "") & "'"
        rpt.FormulaFields(6).Text = "'From   : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
    Case "20", "21"
        If rstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\CASHFL.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\CASHFL.RPT")
        rpt.Database.SetDataSource rstTmp
        rpt.FormulaFields(2).Text = "'" & Me.Caption & "'"
        rpt.FormulaFields(1).Text = "'From Date : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
    Case "7", "8"       ', "9"
        'Journal Book
        'by lps
        Dim CatStr$
        CatStr = GCnFa.Execute("select Category from Voucher_type WHERE V_TYPE='" & TXTACC_CODE.BoundText & "'").Fields(0).Value
        '***
        mQRY = "SELECT V_DATE,credit,debit,v_type,v_no,v_add,CHQ_NO,CHQ_DATE,NARRATION,NAME FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.party=SubGroup.SUBCODE where V_date BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND trim(V_TYPE)=" & Chk_Text(TXTACC_CODE.BoundText) & " ORDER BY V_DATE,V_TYPE,V_NO,V_ADD,V_SNO"
        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\JRNL.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + IIf(Me.Tag = "7", "\JRNL.RPT", "\JRNL1.RPT"))
        rpt.Database.SetDataSource Rst1
        For i = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
                Case "TITLE"
                    rpt.FormulaFields(i).Text = "'" & TXTACC_CODE & "'"
                Case "DATE"
                    rpt.FormulaFields(i).Text = "'From Date : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
                Case "PRINTVOUCHERSUM"   'lps to display voucher total
                    If CatStr = "FA" Then
                        rpt.FormulaFields(i).Text = 1
                    End If
            End Select
        Next
    Case "28"
        mQRY = "SELECT V_DATE,credit,debit,v_type,v_no,v_add,CHQ_NO,CHQ_DATE,NARRATION,NAME FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.party=SubGroup.SUBCODE where V_NO BETWEEN " & Val(Txt_CL_BAL) & " AND " & Val(Txt_PR_CL_BAL) & " AND V_TYPE='J' ORDER BY V_NO,V_ADD,CREDIT,DEBIT"
        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\JRNL28.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\JRNL28.RPT")
        rpt.Database.SetDataSource Rst1
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "TITLE"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "DATE"
                    rpt.FormulaFields(i).Text = "'From Vr.No. : " & Val(Txt_CL_BAL) & " To : " & Val(Txt_PR_CL_BAL) & "'"
            End Select
        Next
    Case "10"   'ANNEXURE
            Set db = OpenDatabase(PubSFADataPath)
MyLbl:
        'For Annexure Without Slab
        For i = 0 To db.QueryDefs.Count - 1
            If db.QueryDefs(i).Name = "ACQRY1" Then
                db.QueryDefs.Delete "ACQRY1"
                Exit For
            End If
        Next
        Set QryDef = db.CreateQueryDef("ACQRY1", mQRY)
        Set Rst1 = GCnFa.Execute("SELECT CITYNAME,GroupName,PARTYCODE,PARTYNAME,Sum(op_cr) AS SumOfop_cr,Sum(OP_dr) AS SumOfOP_dr,Sum(BALANCEdr) AS SumOfBALANCEdr,Sum(balancecr)AS SumOfbalancecr,(Sum(OP_dr)+Sum(BALANCEdr))-(Sum(op_cr)+Sum(balancecr)) AS MyBal From ACQRY1 GROUP BY CITYNAME,GroupName,PARTYCODE,PARTYNAME")
        
        Set Age_Rs = GCnFa.Execute("SELECT * FROM AgeParamFA")
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        If Option4(0).Value = True Then
            X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\ANNEXURE.ttx", True)
            Set rpt = rdApp.OpenReport(PubRepoPath + "\ANNEXURE.RPT")
            rpt.Database.SetDataSource Rst1
            rpt.FormulaFields(4).Text = "'" & Me.Caption & "'"
            rpt.FormulaFields(6).Text = "'From   : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
            rpt.FormulaFields(7).Text = "'For A/C: " & TXTACC_CODE & "'"
        Else
            '********************************************************************************************
        'For Annexure With Slab
MyLbl1:
            For i = 0 To db.QueryDefs.Count - 1
                If db.QueryDefs(i).Name = "ACQRY" Then
                    db.QueryDefs.Delete "ACQRY"
                    Exit For
                End If
            Next
            Set QryDef = db.CreateQueryDef("ACQRY", mQRY)
            Set rstTmp = New ADODB.Recordset
            Set rstTmp = AGETMP(rstTmp)
            Label1.Visible = True
            Label1.Caption = ""
            Set Rst1 = GCnFa.Execute("SELECT GroupName,PARTYNAME, Sum(op_cr) AS SumOfop_cr, Sum(OP_dr) AS SumOfOP_dr, Sum(BALANCEdr) AS SumOfBALANCEdr, Sum(balancecr) AS SumOfbalancecr, (Sum(OP_dr)+Sum(BALANCEdr))-(Sum(op_cr)+Sum(balancecr)) AS Bal From ACQRY GROUP BY GroupName,PARTYNAME")
            Do Until Rst1.EOF
                If Rst1!BAL <> 0 Then
                    rstTmp.AddNew
                    Label1.Caption = Rst1!PartyName
                    rstTmp!ACC_NAME = Rst1!PartyName
                    rstTmp!ANAME = Rst1!GroupName
                    If Abs(Rst1!BAL) <= Age_Rs!p1 Then
                        rstTmp!DEBIT1 = Rst1!BAL
                    ElseIf Abs(Rst1!BAL) > Age_Rs!p1 And Abs(Rst1!BAL) <= Age_Rs!p2 Then
                        rstTmp!DEBIT2 = Rst1!BAL
                    ElseIf Abs(Rst1!BAL) > Age_Rs!p2 And Abs(Rst1!BAL) <= Age_Rs!p3 Then
                        rstTmp!DEBIT3 = Rst1!BAL
                    ElseIf Abs(Rst1!BAL) > Age_Rs!p3 And Abs(Rst1!BAL) <= Age_Rs!p4 Then
                        rstTmp!DEBIT4 = Rst1!BAL
                    ElseIf Abs(Rst1!BAL) > Age_Rs!p4 And Abs(Rst1!BAL) <= Age_Rs!p5 Then
                        rstTmp!DEBIT5 = Rst1!BAL
                    ElseIf Abs(Rst1!BAL) > Age_Rs!p5 And Abs(Rst1!BAL) <= Age_Rs!p6 Then
                        rstTmp!DEBIT6 = Rst1!BAL
                    Else
                        rstTmp!Debit = Rst1!BAL
                    End If
                    rstTmp.Update
                End If
                Rst1.MoveNext
            Loop
            
            
            X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\AGEING.ttx", True)
            Set rpt = rdApp.OpenReport(PubRepoPath + "\ANNEXURE2.RPT")
            
            For i = 1 To rpt.FormulaFields.Count
                Select Case rpt.FormulaFields(i).FormulaFieldName
                    Case "title"
                        rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                    Case "ACNAME"
                        rpt.FormulaFields(i).Text = "'For A/C  : " & TXTACC_CODE.Text & "'"
                    Case "DATE"
                        rpt.FormulaFields(i).Text = "'Upto Date : " & TXTE_DATE & "'"
                    Case "P1"
                        rpt.FormulaFields(i).Text = "'" & "0 - " + str(Age_Rs!p1) & "'"
                    Case "P2"
                        rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p1 + 1) & " - " & str(Age_Rs!p2) & "'"
                    Case "P3"
                        rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p2 + 1) + " - " + str(Age_Rs!p3) & "'"
                    Case "P4"
                        rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p3 + 1) + " - " + str(Age_Rs!p4) & "'"
                    Case "P5"
                        rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p4 + 1) + " - " + str(Age_Rs!p5) & "'"
                    Case "P6"
                        rpt.FormulaFields(i).Text = "'" & str(Age_Rs!p5 + 1) + " - " + str(Age_Rs!p6) & "'"
                    Case "P7"
                        rpt.FormulaFields(i).Text = "'Above " & str(Age_Rs!p6) & "'"
                End Select
            Next
            rpt.Database.SetDataSource rstTmp
        End If

    Case "5", "6"   'Cashbook/Bank format English,
        If rstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        If Index = 0 Then
            X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\CashBookLed.ttx", True)
            Set rpt = rdApp.OpenReport(PubRepoPath + "\CashBookLed.RPT")
        Else
            X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\CashBookLedDOS.ttx", True)
            Set rpt = rdApp.OpenReport(PubRepoPath + "\CashBookLedDOS.RPT")
        End If
        rpt.Database.SetDataSource rstTmp
        rpt.FormulaFields(17).Text = "'" & Me.Caption & "'"
        rpt.FormulaFields(13).Text = "'" & TXTACC_CODE & "'"
        
    Case "30"   'Cash Book Format Long
        If rstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        If Index = 0 Then
            X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\cashbook.ttx", True)
            Set rpt = rdApp.OpenReport(PubRepoPath + "\cashbook.RPT")
        Else
            X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\cashbookDOS.ttx", True)
            Set rpt = rdApp.OpenReport(PubRepoPath + "\cashbookDOS.RPT")
        End If
        rpt.Database.SetDataSource rstTmp
        rpt.FormulaFields(17).Text = "'" & Me.Caption & "'"
        rpt.FormulaFields(13).Text = "'" & TXTACC_CODE & "'"
    Case "6"
        If rstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        If Index = 0 Then
            X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\cashbook.ttx", True)
            Set rpt = rdApp.OpenReport(PubRepoPath + "\Bankbook.RPT")
        Else
'            X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\cashbookDOS.ttx", True)
'            Set rpt = rdApp.OpenReport(PubRepoPath + "\cashbookDOS.RPT")
        End If
        rpt.Database.SetDataSource rstTmp
        rpt.FormulaFields(17).Text = "'" & Me.Caption & "'"
        rpt.FormulaFields(13).Text = "'" & TXTACC_CODE & "'"
    Case "13"
        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\BANKREG.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\BANKREG.RPT")
        rpt.Database.SetDataSource Rst1
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "title"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "DATE"
                    rpt.FormulaFields(i).Text = "'From Date : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
                Case "ACNAME"
                    rpt.FormulaFields(i).Text = "'From A/C : " & TXTACC_CODE.Text & "'"
            End Select
        Next
    Case "22", "23"
        If Me.Tag = "22" Then
            mQRY1 = " WHERE V_date BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText) & " AND CHQ_NO<> ''AND CHQ_NO IS NOT NULL AND CLG_DATE IS NOT NULL"
        ElseIf Me.Tag = "23" Then
            mQRY1 = " WHERE V_date BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND PARTY=" & Chk_Text(TXTACC_CODE.BoundText) & " AND CHQ_NO<> '' AND CHQ_NO IS NOT NULL AND CLG_DATE IS NULL"
        End If
        mQRY = "SELECT V_DATE,credit,debit,v_type,v_no,v_add,CHQ_NO,CHQ_DATE,CLG_DATE,NARRATION,PARTY_NAME,PARTY_LIST.ADD1,PARTY_LIST.ADD2,CITY_NAME,SubGroup.NAME AS CONTRA_NAME FROM (ViewLedger LEFT JOIN SubGroup ON ViewLedger.party1=SubGroup.SUBCODE) LEFT JOIN PARTY_LIST ON ViewLedger.party=PARTY_LIST.SUBCODE " & mQRY1
        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\CLG.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\CLG.RPT")
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "TITLE"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "DATE"
                    rpt.FormulaFields(i).Text = "'From Date : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
            End Select
        Next
        rpt.Database.SetDataSource Rst1
    Case "16", "17", "18"   '17=Profit & Loss A/c
        If rstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\PRLOSS.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\PRLOSS.RPT")
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "title"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "START_DATE"
                    rpt.FormulaFields(i).Text = "'" & TXTS_DATE & "'"
                Case "END_DATE"
                    rpt.FormulaFields(i).Text = "'" & TXTE_DATE & "'"
                Case "BALANCE"
                    rpt.FormulaFields(i).Text = "'" & CStr(Format(IIf(CR > DR, CR - DR, DR - CR), "0.00")) & "'"
                Case "BALANCE_TYPE"
                    rpt.FormulaFields(i).Text = "'" & IIf(CR > DR, "Gross Profit ", "Gross Loss ") & "'"
                Case "bal1"
                    rpt.FormulaFields(i).Text = "'" & CStr(Format(IIf(Cr1 > Dr1, Cr1 - Dr1, Dr1 - Cr1), "0.00")) & "'"
                Case "BAL_TYPE1"
                    rpt.FormulaFields(i).Text = "'" & IIf(Cr1 > Dr1, "Gross Profit (Prv.Yr.)", "Gross Loss (Prv.Yr.)") & "'"
            End Select
        Next
        rpt.Database.SetDataSource rstTmp
    Case "19"
        If rstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\BLNSHEET.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\BLNSHEET.RPT")
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "title"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "START_DATE"
                    rpt.FormulaFields(i).Text = "'" & TXTS_DATE & "'"
                Case "END_DATE"
                    rpt.FormulaFields(i).Text = "'" & TXTE_DATE & "'"
                Case "TOT_INT"
                    rpt.FormulaFields(i).Text = IIf(CR > DR, CR - DR, DR - CR)
                Case "ADD1"
                    rpt.FormulaFields(i).Text = "'" & IIf(CR > DR, "Net Profit From P/L A/C", "Net Loss From P/L A/C") & "'"
                Case "BALANCE"
                    rpt.FormulaFields(i).Text = IIf(Cr1 > Dr1, Cr1 - Dr1, Dr1 - Cr1)
                Case "ADD2"
                    rpt.FormulaFields(i).Text = "'" & IIf(Cr1 > Dr1, "Net Profit From P/L A/C", "Net Loss From P/L A/C") & "'"
            End Select
        Next
        rpt.Database.SetDataSource rstTmp
    Case "27"
        If CDate(TXTS_DATE.Text) = PubStartDate Then
            Set GRs = GCnFa.Execute("SELECT SubCode,iif(IsNull(Sum(Ledger.AmtCr)),0,Sum(Ledger.AmtCr)) AS SumOfAmtCr, iif(Isnull(Sum(Ledger.AmtDr)),0,Sum(Ledger.AmtDr)) AS SumOfAmtDr From Ledger Where Ledger.V_Type='F_AO' And  Ledger.SubCode='" & TXTACC_CODE.BoundText & "' And Ledger.V_Date <= #" & CDate(TXTS_DATE.Text) & "#  GROUP BY Ledger.SubCode")
        Else
            Set GRs = GCnFa.Execute("SELECT SubCode,iif(IsNull(Sum(Ledger.AmtCr)),0,Sum(Ledger.AmtCr)) AS SumOfAmtCr, iif(Isnull(Sum(Ledger.AmtDr)),0,Sum(Ledger.AmtDr)) AS SumOfAmtDr From Ledger Where Ledger.V_Type<>'F_AO' And Ledger.SubCode='" & TXTACC_CODE.BoundText & "' And Ledger.V_Date < #" & CDate(TXTS_DATE.Text) & "#  GROUP BY Ledger.SubCode")
            Set MyRs = GCnFa.Execute("SELECT SubCode,iif(IsNull(Sum(Ledger.AmtCr)),0,Sum(Ledger.AmtCr))-iif(Isnull(Sum(Ledger.AmtDr)),0,Sum(Ledger.AmtDr)) AS OpDr From Ledger Where Ledger.V_Type='F_AO' And  Ledger.SubCode='" & TXTACC_CODE.BoundText & "' And Ledger.V_Date <= #" & CDate(PubStartDate) & "#  GROUP BY Ledger.SubCode")
            If MyRs.RecordCount > 0 Then
                TOT_AMTDR = MyRs!OpDr
            Else
                TOT_AMTDR = 0
            End If
        End If
        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        X11 = CreateFieldDefFile(Rst1, PubRepoPath + "\DAILYSUM.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\DAILYSUM.RPT")
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "title"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "DATE"
                    rpt.FormulaFields(i).Text = "'From Date : " & TXTS_DATE & " To " & TXTE_DATE & "'"
                Case "PARTYNAME"
                    rpt.FormulaFields(i).Text = "'For Party : " & TXTACC_CODE & "'"
                Case "MyOpBal"
                    If GRs.RecordCount > 0 Then
                        rpt.FormulaFields(i).Text = (GRs!SumOfAmtDr - GRs!SumOfAmtCr + Abs(TOT_AMTDR))
                    Else
                        rpt.FormulaFields(i).Text = 0
                    End If
            End Select
        Next
        rpt.Database.SetDataSource Rst1
    Case "29"
        ac_str = ""
        If Check2.Value = Unchecked Then
            ac_str = FillString(GridRow1, 1, 1)
            If ac_str = "" Then Exit Sub
            Set Rst1 = GCnFa.Execute("Select  SubGroup.GroupCode,SubGroup.Name, SubGroup.SubCode  From SubGroup Where SubGroup.SubCode Not In (Select SubCode From Ledger Where V_Date Between #" & CDate(TXTS_DATE) & "# And #" & CDate(TXTE_DATE) & "#) AND GroupCode In( " & ac_str & ") GROUP BY  SubGroup.GroupCode,SubGroup.Name, SubGroup.SubCode ")
        Else
            Set Rst1 = GCnFa.Execute("Select  SubGroup.GroupCode,SubGroup.Name, SubGroup.SubCode From SubGroup Where SubGroup.SubCode Not In (Select SubCode From Ledger Where V_Date Between #" & CDate(TXTS_DATE) & "# And #" & CDate(TXTE_DATE) & "#)  GROUP BY  SubGroup.GroupCode,SubGroup.Name, SubGroup.SubCode ")
        End If
        
'        Set Rst1 = GCnFa.Execute(mQRY)
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
        Set rstTmp = New ADODB.Recordset
        Set rstTmp = AGETMP(rstTmp)
        Label1.Visible = True
        While Not Rst1.EOF
            Label1.Caption = Rst1!Name
            'Opening Calculating
            Set MyRst = GCn.Execute("Select (sum(IIF(ISNULL(AmtDr),0,AmtDr))-sum(IIF(ISNULL(AmtCr),0,AmtCr))) As OpBal  From Ledger Where SubCode='" & Rst1!SubCode & "' And V_Type='F_AO'")
            If MyRst.RecordCount > 0 Then MyOpBal = IIf(IsNull(MyRst!OpBal), 0, MyRst!OpBal)
            'Closing Calculating
            Set MyRst = GCn.Execute("SELECT iif(isnull(Sum(Ledger.AmtDr)),0,Sum(Ledger.AmtDr))-iif(isnull(Sum(Ledger.AmtCr)),0,Sum(Ledger.AmtCr)) AS ClBal FROM Ledger Where V_Date <=#" & CDate(TXTE_DATE) & "# And SubCode='" & Rst1!SubCode & "' And V_Type<>'F_AO'")
            
            
            If MyRst.RecordCount > 0 Then MyCloBal = IIf(IsNull(MyRst!ClBal), 0, MyRst!ClBal)
            'Last Dr Details
            Set MyRst = GCn.Execute("SELECT SubGroup.Name, ViewLedger.V_DATE, ViewLedger.credit, ViewLedger.debit, ViewLedger.v_type, ViewLedger.v_no FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.party1 = SubGroup.SubCode WHERE  ViewLedger.V_DATE <=#" & CDate(TXTE_DATE) & "# And ViewLedger.party='" & Rst1!SubCode & "' AND  ViewLedger.v_type<>'F_AO' And   ViewLedger.Debit>0 ORDER BY ViewLedger.V_DATE DESC")
            If MyRst.RecordCount > 0 Then
                MyDrStr = CStr(MyRst!V_DATE) + Space(1) + CStr(MyRst!V_Type) + Space(1) + CStr(MyRst!v_no) + Space(1) + CStr(MyRst!Debit)
            End If
            'Last Cr Details
            Set MyRst = GCn.Execute("SELECT SubGroup.Name, ViewLedger.V_DATE, ViewLedger.credit, ViewLedger.debit, ViewLedger.v_type, ViewLedger.v_no FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.party1 = SubGroup.SubCode WHERE  ViewLedger.V_DATE <=#" & CDate(TXTE_DATE) & "# And ViewLedger.party='" & Rst1!SubCode & "' AND  ViewLedger.v_type<>'F_AO' And   ViewLedger.Credit>0 ORDER BY ViewLedger.V_DATE DESC")
            If MyRst.RecordCount > 0 Then
                MyCrStr = CStr(MyRst!V_DATE) + Space(1) + CStr(MyRst!V_Type) + Space(1) + CStr(MyRst!v_no) + Space(1) + CStr(MyRst!credit)
            End If
            'Insert Record
            With rstTmp
                .AddNew
                !ACC_NAME = Rst1!Name
                !OpBal = Format(MyOpBal, "0.00")
                !ClBal = Format((MyOpBal + MyCloBal), "0.00")
                !LastDrVNo = MyDrStr
                !LastCrVNo = MyCrStr
                !ANAME = GCn.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!GroupCode & "'").Fields(0).Value
                .Update
            End With
            Rst1.MoveNext
        Wend
        X11 = CreateFieldDefFile(rstTmp, PubRepoPath + "\NonTran.ttx", True)
        Set rpt = rdApp.OpenReport(PubRepoPath + "\NonTran.RPT")
        For i = 1 To rpt.FormulaFields.Count
            Select Case rpt.FormulaFields(i).FormulaFieldName
                Case "title"
                    rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
                Case "DATE"
                    rpt.FormulaFields(i).Text = "'From Date : " & TXTS_DATE & " To " & TXTE_DATE & "'"
                Case "PARTYNAME"
                    rpt.FormulaFields(i).Text = "'For Party : " & TXTACC_CODE & "'"
            End Select
        Next
        rpt.Database.SetDataSource rstTmp
End Select
Select Case Me.Tag
    Case "5"
        rpt.ReadRecords
        Report_View rpt, "Cash Book", , False
        
    Case "30", "6", "28"
        rpt.ReadRecords
        Report_View rpt, Me.Caption, , False
    Case "10", "11", "12", "13", "20", "21", "14", "15", "16", "17", "18", "19", "9", "27", "22", "23", "29"
        rpt.ReadRecords
        Report_View rpt, Me.Caption, , False
    Case "7", "8"
        rpt.ReadRecords
        Report_View rpt, TXTACC_CODE, , False
End Select
Set mGroup_Rs = Nothing
Set SubGroup_Rs = Nothing
Set rstTmp = Nothing
Set Age_Rs = Nothing
Set MyRst = Nothing

TXTACC_CODE.BoundText = ""
    Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical, Me.Caption: Exit Sub
End Sub
Private Sub TXTE_DATE_Validate(Cancel As Boolean)
    TXTE_DATE = RetDate(TXTE_DATE)
End Sub
Private Sub TXTS_DATE_Validate(Cancel As Boolean)
    TXTS_DATE = RetDate(TXTS_DATE)
End Sub
Private Function FillString(GridArray As Variant, Gridindex As Integer, DataType As Byte) As String
Dim ac_str As String
Dim i As Integer
Dim GridRow As Integer
    ac_str = ""
    For i = 0 To UBound(GridArray)
        If GridArray(i) = 0 Then GoTo NXT:
        GridRow = GridArray(i)
        If GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü" Then
                If DataType = 0 Then
                   ac_str = ac_str + IIf(ac_str = "", GridSel(Gridindex).TextMatrix(GridRow, 2), "," + GridSel(Gridindex).TextMatrix(GridRow, 2))
                ElseIf DataType = 1 Then
                   ac_str = ac_str + IIf(ac_str = "", "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'", "," + "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'")
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
    If ac_str = "" Then
        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
        GridSel(Gridindex).SetFocus
        
        Exit Function
    End If
    FillString = ac_str
    Exit Function
End Function
Private Sub GridInitialise(Gridindex As Integer, GridSql As String, Optional GridSetting As Boolean)
Dim Index As Integer
Index = Gridindex
If Index = 1 Then
    Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
    RsGrid1.Open GridSql, GCnFa, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid1
    ReDim Preserve GridRow1(0)
    GridRow1(0) = 0
End If
End Sub
Private Sub Ini_Grid(Index As Integer)
'Date1,Date2,List1,List1,List2,List3
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Select Case Index
    Case 0
        Grid1Sql = "SELECT '' as O,NAME as Account,SUBCODE as AccId from SubGroup order by name "
        GridInitialise 1, Grid1Sql, True
    Case 1
        If Me.Tag = "14" Then
            Grid1Sql = "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup Where Nature='Customer' order by GroupName"
        ElseIf Me.Tag = "15" Then
            Grid1Sql = "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup Where Nature='Supplier' order by GroupName"
        Else
            Grid1Sql = "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup order by GroupName"
        End If
        GridInitialise 1, Grid1Sql, True
    Case 2
        Grid1Sql = "SELECT '' as O,NAME as Account,SUBCODE as AccId from SubGroup where nature='Bank' order by name "
        GridInitialise 1, Grid1Sql, True
    Case 3
        Grid1Sql = "SELECT '' as O,NAME as Account,SUBCODE as AccId from SubGroup where nature='Cash' order by name "
        GridInitialise 1, Grid1Sql, True
End Select
End Sub

Public Sub SelGridKeyPressLocal(txt As Object, FGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If FGrid(Index).Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then txt.Text = "": Exit Sub
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack And Len(txt.SelText) <> 1 Then
            txt.SelLength = Len(txt.SelText) - 1
            FindStr = txt.SelText
        Else
            FindStr = txt.SelText + Chr(KeyAscii)
        End If
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
            Rst.FIND "" & FindFldName & " >=" & Val(FindStr) & ""
        Else    'character serach
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
        End If
        KeyAscii = 0
       If Rst.AbsolutePosition <> adPosEOF And Rst.AbsolutePosition <> adPosBOF Then
            FGrid(Index).CellBackColor = CellBackColLeave
            FGrid(Index).Row = Rst.AbsolutePosition
            FGrid(Index).CellBackColor = CellBackColEnter
            txt.Text = Rst.Fields(FindFldName).Value
            txt.SelLength = Len(FindStr)
            txt.left = FGrid(Index).CellLeft + FGrid(Index).left
            txt.top = FGrid(Index).CellTop + FGrid(Index).top
            If txt.Visible = False Then
                txt.Visible = True: txt.ZOrder 0: txt.SetFocus: txt.BackColor = FGrid(Index).CellBackColor
                 txt.ForeColor = FGrid(Index).CellForeColor: txt.width = FGrid(Index).CellWidth: txt.Height = FGrid(Index).CellHeight
            End If
       End If
End Sub
Private Function NavigationKey(KeyCode As Integer) As Boolean
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyUp _
    Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        NavigationKey = True
    End If
End Function
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
TxtSearch.Visible = False: TxtSearch.Text = ""
End Sub

