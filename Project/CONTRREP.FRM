VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form CONTRAREP 
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contra Ledger"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   Icon            =   "CONTRREP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7830
   ScaleWidth      =   11250
   Begin VB.CheckBox Check1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "As Per Detail"
      Height          =   315
      Left            =   2625
      TabIndex        =   35
      Top             =   4290
      Width           =   2220
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
      Left            =   4920
      TabIndex        =   34
      Top             =   4350
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TXTE_DATE 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1935
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1395
   End
   Begin VB.TextBox TXTS_DATE 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1935
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1395
   End
   Begin VB.TextBox TXTINT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1935
      MaxLength       =   5
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "10"
      ToolTipText     =   "Enter the Interest Rate."
      Top             =   2970
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1935
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1980
      Width           =   1395
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1935
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2310
      Width           =   1395
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1935
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2640
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
      Left            =   6075
      TabIndex        =   25
      ToolTipText     =   "Print Reports"
      Top             =   5025
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   2625
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3600
      Width           =   4335
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   2625
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3945
      Width           =   4335
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00CFE0E0&
      Caption         =   "TxN.Amt.   >="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   330
      TabIndex        =   2
      Top             =   2692
      Width           =   1590
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Cr.Balance >="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   330
      TabIndex        =   1
      Top             =   2362
      Width           =   1590
   End
   Begin VB.CheckBox Check 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Dr.Balance >="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   1987
      Width           =   1590
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H8000000A&
      Caption         =   "&Print"
      DisabledPicture =   "CONTRREP.frx":0442
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
      Left            =   7185
      Picture         =   "CONTRREP.frx":0584
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Reports"
      Top             =   5025
      Width           =   1110
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H8000000A&
      Caption         =   "E&xit"
      DisabledPicture =   "CONTRREP.frx":06C6
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
      Left            =   8295
      Picture         =   "CONTRREP.frx":07C8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   5025
      Width           =   1110
   End
   Begin VB.Frame Frame_GROUP 
      BackColor       =   &H00CFE0E0&
      Height          =   705
      Left            =   4920
      TabIndex        =   18
      Top             =   960
      Width           =   4470
      Begin MSDataListLib.DataCombo TXTGRP_CODE 
         Height          =   315
         Left            =   45
         TabIndex        =   19
         Top             =   240
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
   End
   Begin VB.Frame Frame_SELECTED 
      BackColor       =   &H00CFE0E0&
      Height          =   1695
      Left            =   4920
      TabIndex        =   17
      Top             =   960
      Width           =   4470
      Begin MSDataListLib.DataCombo TXTACC_CODE1 
         Height          =   315
         Left            =   90
         TabIndex        =   21
         Top             =   750
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo TXTACC_CODE 
         Height          =   315
         Left            =   105
         TabIndex        =   20
         Top             =   315
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Merge"
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
      Index           =   4
      Left            =   4035
      TabIndex        =   14
      Top             =   45
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Group"
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
      Index           =   3
      Left            =   3435
      TabIndex        =   13
      Top             =   1965
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Optional"
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
      Index           =   2
      Left            =   3435
      TabIndex        =   11
      Top             =   1605
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Selected"
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
      Index           =   1
      Left            =   3435
      TabIndex        =   10
      Top             =   1245
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00CFE0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   3045
      TabIndex        =   4
      Top             =   135
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame_OPTIONAL 
      BackColor       =   &H00CFE0E0&
      Height          =   2580
      Left            =   4920
      TabIndex        =   15
      Top             =   960
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid FGrid1 
         Height          =   2370
         Left            =   0
         TabIndex        =   16
         Top             =   180
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   3
         FormatString    =   "       | A/C Name                                                           |"
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   1
      Left            =   4920
      TabIndex        =   33
      Top             =   4320
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration Not Having"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   300
      TabIndex        =   23
      Top             =   3960
      Width           =   2190
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration Having"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   720
      TabIndex        =   22
      Top             =   3630
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   3435
      TabIndex        =   12
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   1005
      TabIndex        =   9
      Top             =   3007
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   735
      TabIndex        =   8
      Top             =   1357
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   240
      Left            =   810
      TabIndex        =   7
      Top             =   1687
      Width           =   975
   End
End
Attribute VB_Name = "CONTRAREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 3 And Index <> 4 Then NumPress Text(Index), KeyAscii, 7, 2
End Sub
Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 3 And Index <> 4 Then Text(Index) = Validate_Numeric(Text(Index))
End Sub
Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 3 And Index <> 4 Then NumDown Text(Index), KeyCode, 7, 2
End Sub
Private Sub Check_Click(Index As Integer)
    Text(Index).Enabled = IIf(Check(Index).Value = 1, True, False)
End Sub
Private Sub Form_Activate()
Dim mQRY$
If Me.Tag = "24" Or Me.Tag = "25" Then
'        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.Name"

    GSQL = "SELECT SubCode,TRIM(IIF(ISNULL(SubGroup.Name),'',SubGroup.Name))&TRIM(IIF(ISNULL(SubGroup.Add1),'',', '&SubGroup.Add1))&TRIM(IIf(IsNull(City.CityName),'',', '&City.CityName)) As TName " & _
        "FROM ((SubGroup LEFT JOIN City on City.CityCode=SubGroup.CityCode) " & _
        "left join [" & PubFADataPath & "].AcGroup on SubGroup.GroupCode=AcGroup.GroupCode) " & _
        "Where FirmCode = '" & PubFirmCode & "' and "
    '***
    mQRY = "select CODE,NAME from GROUP_TRANS Where GR_CODE<>'' AND GR_CODE IS NOT NULL and "
    'Nature='" & IIf(Me.Tag = "24", "Customer", "Supplier") & "' order by name", TXTGRP_CODE, "NAME", "CODE")

    If Me.Tag = "24" Then
        GSQL = GSQL & " left(AcGroup.MainGrCode,6)='" & pubSundryDrSysMainGrCode & "'"
        mQRY = mQRY & "left(Group_Trans.MainGrCode,6)='" & pubSundryDrSysMainGrCode & "'"
    Else
        GSQL = GSQL & " left(AcGroup.MainGrCode,6)='" & pubSundryCrSysMainGrCode & "'"
    End If
    GSQL = GSQL & " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' Order by SubGroup.Name"
    Call INI_COMBO(GSQL, TXTACC_CODE, "TName", "SubCode")

'    Call INI_COMBO("SELECT SubCode,TRIM(IIF(ISNULL(SubGroup.Name),'',SubGroup.Name))+', '+TRIM(IIF(ISNULL(SubGroup.Add1),'',SubGroup.Add1))+', '+TRIM(IIf(IsNull(City.CityName),'', City.CityName)) As TName FROM SubGroup LEFT JOIN City on City.CityCode=SubGroup.CityCode Where Nature='" & IIf(Me.Tag = "24", "Customer", "Supplier") & "' ORDER BY SubGroup.Name", TXTACC_CODE1, "TName", "SubCode")
    Call INI_COMBO("select CODE,NAME from GROUP_TRANS Where GR_CODE<>'' AND GR_CODE IS NOT NULL and Nature='" & IIf(Me.Tag = "24", "Customer", "Supplier") & "' order by name", TXTGRP_CODE, "NAME", "CODE")
Else
    Call INI_COMBO("SELECT SubCode,TRIM(IIF(ISNULL(SubGroup.Name),'',SubGroup.Name))+', '+TRIM(IIF(ISNULL(SubGroup.Add1),'',SubGroup.Add1))+', '+TRIM(IIf(IsNull(City.CityName),'', City.CityName)) As TName FROM SubGroup LEFT JOIN City on City.CityCode=SubGroup.CityCode ORDER BY SubGroup.Name", TXTACC_CODE, "TName", "SubCode")
    Call INI_COMBO("SELECT SubCode,TRIM(IIF(ISNULL(SubGroup.Name),'',SubGroup.Name))+', '+TRIM(IIF(ISNULL(SubGroup.Add1),'',SubGroup.Add1))+', '+TRIM(IIf(IsNull(City.CityName),'', City.CityName)) As TName FROM SubGroup LEFT JOIN City on City.CityCode=SubGroup.CityCode ORDER BY SubGroup.Name", TXTACC_CODE1, "TName", "SubCode")
    Call INI_COMBO("select CODE,NAME from GROUP_TRANS WHERE GR_CODE<>'' AND GR_CODE IS NOT NULL order by name", TXTGRP_CODE, "NAME", "CODE")
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
Private Sub Form_Load()
    Call WinSetting(Me)
    TXTE_DATE = PubLoginDate
    TXTS_DATE = PubStartDate
    Frame_GROUP.Visible = False
    Frame_OPTIONAL.Visible = False
    Frame_SELECTED.Visible = False
    Option1(2).Visible = False
    Option1(4).left = Option1(3).left
    Option1(4).top = Option1(3).top
    Option1(3).left = Option1(2).left
    Option1(3).top = Option1(2).top
    
    Frame_GROUP.Visible = False
    Frame_OPTIONAL.Visible = False
    GridSel(1).Visible = False
    Frame_SELECTED.Visible = False
    Ini_Grid 0

    
End Sub
Private Sub Option1_Click(Index As Integer)
Dim Rst1 As ADODB.Recordset
Select Case Index
    Case 0
        Frame_GROUP.Visible = False
        Frame_OPTIONAL.Visible = False
        Frame_SELECTED.Visible = False
        GridSel(1).Visible = False
        Check2.Visible = False
    Case 1
        Ini_Grid 0
        Frame_GROUP.Visible = False
        Frame_OPTIONAL.Visible = False
        Frame_SELECTED.Visible = False
        TXTACC_CODE1.Visible = False
        GridSel(1).Visible = True
        Check2.Visible = True
        GridSel(1).left = Frame_OPTIONAL.left
        GridSel(1).top = Frame_OPTIONAL.top
        GridSel(1).Height = Frame_OPTIONAL.Height
        GridSel(1).width = Frame_OPTIONAL.width
        GridSel(1).ColWidth(1) = 4500: GridSel(1).ColWidth(2) = 0
        Check2.top = GridSel(1).top + 20
        Check2.left = GridSel(1).left + 30
        Check2.Height = GridSel(1).RowHeight(0) + 20
        Check2.width = 950
        Check2.Value = Checked
        TxtSearch.width = GridSel(1).ColWidth(1) - 50
    Case 2
        Frame_GROUP.Visible = False
        'Frame_OPTIONAL.Visible = True
        GridSel(1).Visible = True
        Check2.Visible = True
        GridSel(1).left = Frame_OPTIONAL.left
        GridSel(1).top = Frame_OPTIONAL.top
        GridSel(1).Height = Frame_OPTIONAL.Height
        GridSel(1).width = Frame_OPTIONAL.width
        GridSel(1).ColWidth(1) = 4500: GridSel(1).ColWidth(2) = 0
        Check2.top = GridSel(1).top + 20
        Check2.left = GridSel(1).left + 40
        Check2.Height = GridSel(1).RowHeight(0) + 20
        Check2.width = 970
        Check2.Value = Checked
        Frame_SELECTED.Visible = False
        TxtSearch.width = GridSel(1).ColWidth(1) - 50
'       FGrid1.Rows = 0
'       Set Rst1 = GCnFa.Execute("select SubCode,NAME from SubGroup order by name")
'       Do Until Rst1.EOF
'           FGrid1.AddItem "" & Chr(9) & Rst1!Name & Chr(9) & Rst1!SubCode
'           Rst1.MoveNext
'       Loop
    Case 3
    
''        Frame_GROUP.Visible = True
''        Frame_OPTIONAL.Visible = False
''        Frame_SELECTED.Visible = False
''        GridSel(1).Visible = False
''        Check2.Visible = False
        Ini_Grid 1
        Frame_GROUP.Visible = False
        'Frame_OPTIONAL.Visible = True
        GridSel(1).Visible = True
        Check2.Visible = True
        GridSel(1).left = Frame_OPTIONAL.left
        GridSel(1).top = Frame_OPTIONAL.top
        GridSel(1).Height = Frame_OPTIONAL.Height
        GridSel(1).width = Frame_OPTIONAL.width
        GridSel(1).ColWidth(1) = 4500: GridSel(1).ColWidth(2) = 0
        Check2.top = GridSel(1).top + 20
        Check2.left = GridSel(1).left + 40
        Check2.Height = GridSel(1).RowHeight(0) + 20
        Check2.width = 970
        Check2.Value = Checked
        Frame_SELECTED.Visible = False
        TxtSearch.width = GridSel(1).ColWidth(1) - 50

        
        
    Case 4
        Frame_GROUP.Visible = False
        Frame_OPTIONAL.Visible = False
        Frame_SELECTED.Visible = True
        TXTACC_CODE1.Visible = True
        GridSel(1).Visible = False
        Check2.Visible = False
End Select
Set Rst1 = Nothing
End Sub
Private Sub FGrid1_Click()
    FGrid1.Col = 0
    FGrid1.CellFontName = "WINGDINGS"
    FGrid1.CellFontSize = 14
    FGrid1.TextMatrix(FGrid1.Row, 0) = IIf(FGrid1.TextMatrix(FGrid1.Row, 0) = "ü", " ", "ü")
End Sub
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub TXTINT_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub TXTINT_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And (KeyAscii < 46 Or KeyAscii > 58) Then KeyAscii = 0
End Sub
Private Sub BTNPRINT_Click(Index As Integer)
Dim Rst1 As ADODB.Recordset, rst2 As ADODB.Recordset, ac_str$
Dim i As Integer, oBAL As Double, AC_Name$, Ac_Code$, Qry$, Cnt As Integer, TmpNarr$, TmpAmt As Double
Dim TmpRst As ADODB.Recordset, NarrStr$, X11, rsTName As ADODB.Recordset
'On Error GoTo ErrLoop
If DateDiff("d", TXTS_DATE, TXTE_DATE) < 0 Then
    MsgBox " Ending Date Less than Starting Date ", vbCritical
    Exit Sub
End If
If Me.Tag = "4" Then If IsValid(TXTINT, "Enter Interest Rate") = False Then Exit Sub
'If Option1(3).Value = True Then If IsValid(TXTGRP_CODE, "Select A/C") = False Then Exit Sub
If Option1(4).Value = True And (TXTACC_CODE = "" Or TXTACC_CODE1 = "") Then MsgBox " ** Select A/C ** ", vbCritical, Me.Caption: Exit Sub
If Option1(4).Value = True And TXTACC_CODE = TXTACC_CODE1 Then MsgBox " ** Both A/C are Same** ", vbCritical, Me.Caption: Exit Sub
If Option1(0).Value = True Then

    If Me.Tag = "24" Then
        Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST Where Nature='Customer' ORDER BY NAME")
    ElseIf Me.Tag = "25" Then
        Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST Where Nature='Supplier' ORDER BY NAME")
    Else
        Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST ORDER BY NAME")
    End If
ElseIf Option1(1).Value = True Then
    ac_str = ""
    If Check2.Value = Unchecked Then
       ac_str = FillString(GridRow1, 1, 1)
       Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST WHERE SubCode IN (" & ac_str & ") ORDER BY NAME")
    Else
          Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST ORDER BY NAME")
    End If
ElseIf Option1(2).Value = True Then
    ac_str = ""
    For i = 0 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(i, 0) = "ü" Then ac_str = ac_str + IIf(ac_str = "", "'" & FGrid1.TextMatrix(i, 2) & "'", "," + "'" & FGrid1.TextMatrix(i, 2) & "'")
    Next
    If ac_str = "" Then MsgBox " ** Select A/C ** ", vbCritical, Me.Caption: Exit Sub
    Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST WHERE SubCode IN (" & ac_str & ") ORDER BY NAME")
ElseIf Option1(3).Value = True Then
    ac_str = ""
    If Check2.Value = Unchecked Then
        ac_str = FillString(GridRow1, 1, 1)
        Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST WHERE CODE IN (" & ac_str & ") ORDER BY NAME")
    Else
        Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST ORDER BY NAME")
    End If
'    Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,groupCode as CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST WHERE CODE='" & TXTGRP_CODE.BoundText & "' ORDER BY NAME")
ElseIf Option1(4).Value = True Then
    Set Rst1 = GCnFa.Execute("SELECT GROUPNATURE,CODE,SubCode,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST WHERE SubCode IN ('" & TXTACC_CODE.BoundText & "','" & TXTACC_CODE1.BoundText & "') ORDER BY NAME")
End If
Qry = ""
If Check(2).Value = 1 Then Qry = Qry + "AND VIEWLEDGER.DEBIT+VIEWLEDGER.CREDIT>=" & Val(Text(2))
If Trim(Text(3)) <> "" Then Qry = Qry + " AND INSTR(1,UCASE(VIEWLEDGER.NARRATION),UCASE(TRIM('" & Text(3) & "')))"
If Trim(Text(4)) <> "" Then Qry = Qry + " AND NOT INSTR(1,UCASE(VIEWLEDGER.NARRATION),UCASE(TRIM('" & Text(4) & "')))"
Set TmpRst = New ADODB.Recordset
Set TmpRst = ADTMP1(TmpRst)
AC_Name = ""
If Option1(4).Value = True Then
    Do Until Rst1.EOF
        AC_Name = AC_Name + Trim(Rst1!Name) + ","
        Rst1.MoveNext
    Loop
End If
If Rst1.RecordCount > 0 Then Rst1.MoveFirst
Do Until Rst1.EOF
    If Check(0).Value = 1 Then If GCnFa.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT))-IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE<=#" & CDate(TXTE_DATE) & "#").Fields(0) < Val(Text(0)) Then GoTo EXIT_LOOP
    If Check(1).Value = 1 Then If GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE<=#" & CDate(TXTE_DATE) & "#").Fields(0) < Val(Text(1)) Then GoTo EXIT_LOOP
    oBAL = 0
    Ac_Code = IIf(Option1(4).Value = True, 0, Rst1!SubCode)
    If Option1(4).Value = False Then AC_Name = Trim(Rst1!Name)
    If Me.Tag = "24" Then
        If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
            oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE CREDIT >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE <=#" & CDate(TXTE_DATE) & "# ").Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE DEBIT>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND (V_DATE<#" & CDate(TXTS_DATE) & "# OR (V_DATE<=#" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO'))").Fields(0)
        Else
            oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE CREDIT >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN #" & CDate(PubStartDate) & "# AND #" & CDate(TXTE_DATE) & "#").Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE DEBIT>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN #" & CDate(PubStartDate) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO'").Fields(0)
        End If
        If oBAL < 0 Then
            With TmpRst
                .AddNew
                !V_Type = "": !v_no = 0: !v_add = "": !v_sno = 0: !Name = "Opening Balance"
                !V_DATE = TXTS_DATE: !CR = Abs(oBAL): !ADJQTY = Abs(oBAL): !Val = "1"
                !NAME1 = AC_Name: !SubCode = Ac_Code: !CITY_NAME = XNull(Rst1!CITY_NAME)
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                !GrCode = Rst1!Code
                !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value
                .Update
            End With
            oBAL = 0
        End If
        Set rst2 = GCnFa.Execute("SELECT VIEWLEDGER.*,SubGroup.Name FROM VIEWLEDGER LEFT JOIN SubGroup ON SubGroup.SubCode=VIEWLEDGER.PARTY1 WHERE VIEWLEDGER.DEBIT>0 AND VIEWLEDGER.PARTY=" & Chk_Text(Rst1!SubCode) & " AND VIEWLEDGER.V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE<>'F_AO' ORDER BY V_DATE,V_TYPE,V_NO")
        Do Until rst2.EOF
            If oBAL >= rst2!Debit Then
                oBAL = oBAL - rst2!Debit
            Else
                With TmpRst
                    .AddNew
                    !v_add = rst2!v_add: !v_no = rst2!v_no: !Name = XNull(rst2!Name)
                    !V_DATE = rst2!V_DATE: !CR = rst2!Debit: !ADJQTY = rst2!Debit - oBAL
                    !V_Type = rst2!V_Type: !NARRATION1 = Mid(rst2!Narration, 1, 150)
                    !v_sno = rst2!v_sno: !NAME1 = AC_Name: !SubCode = Ac_Code: !CITY_NAME = XNull(Rst1!CITY_NAME)
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    !GrCode = Rst1!Code
                    !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value

                    .Update
                End With
                oBAL = 0
            End If
            rst2.MoveNext
        Loop
        If oBAL > 0 Then
            With TmpRst
                .AddNew
                !V_Type = "": !v_no = 0: !v_add = "": !v_sno = 0: !Name = "Excess Credit"
                !V_DATE = TXTE_DATE: !AdjAmt = Abs(oBAL): !NAME1 = AC_Name: !SubCode = Ac_Code
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                !CITY_NAME = XNull(Rst1!CITY_NAME)
                !GrCode = Rst1!Code
                !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value

                .Update
            End With
        End If
    ElseIf Me.Tag = "25" Then
        If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
            oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE DEBIT >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE <=#" & CDate(TXTE_DATE) & "# ").Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE CREDIT>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND (V_DATE<#" & CDate(TXTS_DATE) & "# OR (V_DATE<=#" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO'))").Fields(0)
        Else
            oBAL = GCnFa.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE DEBIT >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN #" & CDate(PubStartDate) & "# AND #" & CDate(TXTE_DATE) & "#").Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE CREDIT>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN #" & CDate(PubStartDate) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO'").Fields(0)
        End If
        If oBAL < 0 Then
            With TmpRst
                .AddNew
                !V_Type = "": !v_no = 0: !v_add = "": !v_sno = 0: !NARRATION1 = "OPENING BALANCE"
                !NARRATION2 = "": !Name = "OPENING BALANCE": !V_DATE = TXTS_DATE: !CR = Abs(oBAL)
                !ADJQTY = Abs(oBAL): !Val = "1": !NAME1 = AC_Name: !SubCode = Ac_Code
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                !CITY_NAME = XNull(Rst1!CITY_NAME)
                !GrCode = Rst1!Code
                !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value

                .Update
            End With
            oBAL = 0
        End If
        Set rst2 = GCnFa.Execute("SELECT VIEWLEDGER.*,SubGroup.Name FROM VIEWLEDGER LEFT JOIN SubGroup ON SubGroup.SubCode=VIEWLEDGER.PARTY1 WHERE VIEWLEDGER.CREDIT>0 AND VIEWLEDGER.PARTY=" & Chk_Text(Rst1!SubCode) & " AND VIEWLEDGER.V_DATE BETWEEN #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_tYPE<>'F_AO' ORDER BY V_DATE,V_TYPE,V_NO")
        Do Until rst2.EOF
            If oBAL >= rst2!credit Then
                oBAL = oBAL - rst2!credit
            Else
                With TmpRst
                    .AddNew
                    !v_add = rst2!v_add: !v_no = rst2!v_no: !Name = IIf(IsNull(rst2!Name), "", rst2!Name): !V_DATE = rst2!V_DATE
                    !CR = rst2!credit: !ADJQTY = rst2!credit - oBAL: !V_Type = rst2!V_Type
                    !NARRATION1 = IIf(IsNull(rst2!Narration), "", rst2!Narration): !v_sno = rst2!v_sno: !NAME1 = AC_Name
                    !SubCode = Ac_Code: !CITY_NAME = XNull(Rst1!CITY_NAME)
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    !GrCode = Rst1!Code
                    !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value

                    .Update
                End With
                oBAL = 0
            End If
            rst2.MoveNext
        Loop
        If oBAL > 0 Then
            With TmpRst
                .AddNew
                !V_Type = "": !v_no = 0: !v_add = "": !v_sno = 0: !Name = "Excess Debit"
                !V_DATE = TXTE_DATE: !AdjAmt = Abs(oBAL): !NAME1 = AC_Name: !SubCode = Ac_Code
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), " ", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), " ", ", " + Rst1!Add2)))
                !CITY_NAME = XNull(Rst1!CITY_NAME)
                !GrCode = Rst1!Code
                !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value

                .Update
            End With
        End If
    Else
        If Qry = "" Then
            If Me.Tag = "2" Or Me.Tag = "3" Or Me.Tag = "4" Then
                If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
                    oBAL = oBAL + GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM  VIEWLEDGER WHERE (V_DATE<#" & CDate(TXTS_DATE) & "# OR (V_DATE BETWEEN  #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO')) AND PARTY=" & Chk_Text(Rst1!SubCode)).Fields(0)
                Else
                    oBAL = oBAL + GCnFa.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM  VIEWLEDGER WHERE (V_DATE BETWEEN  #" & CDate(TXTS_DATE) & "# AND #" & CDate(TXTE_DATE) & "# AND V_TYPE='F_AO') AND PARTY=" & Chk_Text(Rst1!SubCode)).Fields(0)
                End If
            End If
            If oBAL <> 0 Then
                With TmpRst
                    .AddNew
                    !V_Type = "": !v_no = 0: !v_add = "": !v_sno = 0: !Name = "OPENING BALANCE"
                    !V_DATE = TXTS_DATE: !CR = IIf(oBAL > 0, Abs(oBAL), 0)
                    !AdjAmt = IIf(oBAL < 0, Abs(oBAL), 0): !Val = "1": !NAME1 = AC_Name
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), " ", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), " ", ", " + Rst1!Add2)))
                    !CITY_NAME = XNull(Rst1!CITY_NAME): !SubCode = Ac_Code
                    !GrCode = Rst1!Code
                    !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value
                    .Update
                End With
            End If
        End If
        If Me.Tag = "2" Or Me.Tag = "3" Or Me.Tag = "4" Then
            Set rst2 = GCnFa.Execute("SELECT VIEWLEDGER.*,SubGroup.Name FROM VIEWLEDGER LEFT JOIN SubGroup ON VIEWLEDGER.PARTY1 = SubGroup.SubCode WHERE VIEWLEDGER.V_DATE Between #" & CDate(TXTS_DATE) & "# And #" & CDate(TXTE_DATE) & "# AND (VIEWLEDGER.PARTY=" & Chk_Text(Rst1!SubCode) & ") AND V_TYPE<>'F_AO'" & Qry)
            Do Until rst2.EOF
                 With TmpRst
                    .AddNew
                    !v_add = rst2!v_add: !v_no = rst2!v_no
                    If Me.Tag = "3" Then
                        !NARRATION1 = Mid(XNull(rst2!Narration), 1, 150)
                        Set GRs = GCnFa.Execute("SELECT ledger.*,SubGroup.Name " & _
                            " FROM Ledger INNER JOIN SubGroup ON Ledger.SubCode = SubGroup.SubCode " & _
                            "WHERE LEDGER.V_DATE = #" & rst2!V_DATE & _
                            "# and V_SNo <> " & rst2!v_sno & " AND V_TYPE='" & rst2!V_Type & _
                            "' and left(docid,1) & mid(docid,3,1) &'/'& left(trim(mid(docid,4,5)),1) &   mid(trim(mid(docid,4,5)),3, len(trim(mid(docid,4,5)))-2) & '/' & trim(right(docid,8)) ='" & rst2!v_no & "'")
                        If GRs.RecordCount >= 1 Then
                            If GRs.RecordCount = 1 Then
                                !Name = Mid(IIf(IsNull(GRs!Name), "", GRs!Name), 1, 35)
                            Else
                                Cnt = 1
                                TmpNarr = ""
                                TmpAmt = 0
                                If Check1.Value = vbChecked Then
                                    !Name = "As per detail:"
                                Else
                                    TmpNarr = "As per detail:" + vbCrLf
                                End If
                                GRs.MoveFirst
                                While Not GRs.EOF
                                    TmpNarr = TmpNarr + UCase(IIf(Len(Mid(GRs!Name, 1, 35)) < 35, Mid(GRs!Name, 1, 35) + Space(35 - Len(Mid(GRs!Name, 1, 35))), Mid(GRs!Name, 1, 35))) + Space(1) + Space(11 - Len(CStr(Format(Abs(GRs!AmtCr - GRs!AmtDr), "0.00")))) + CStr(Format(Abs(GRs!AmtCr - GRs!AmtDr), "0.00")) + IIf(GRs!AmtCr - GRs!AmtDr > 0, "Cr", "Dr") + vbCrLf
                                    If Check1.Value = vbUnchecked Then
                                        If rst2!credit = 0 Then
                                            If (GRs!AmtCr - GRs!AmtDr > 0 And Abs(GRs!AmtCr - GRs!AmtDr) > TmpAmt) Then
                                                !Name = Mid(GRs!Name, 1, 35)
                                                TmpAmt = Abs(GRs!AmtCr - GRs!AmtDr)
                                            End If
                                        Else
                                            If (GRs!AmtCr - GRs!AmtDr < 0 And Abs(GRs!AmtCr - GRs!AmtDr) > TmpAmt) Then
                                                !Name = Mid(GRs!Name, 1, 35)
                                                TmpAmt = Abs(GRs!AmtCr - GRs!AmtDr)
                                            End If
                                        End If
                                    End If
                                    GRs.MoveNext
                                    Cnt = Cnt + 1
                                Wend
                                !Narration3 = TmpNarr
                            End If
                        End If
                    Else
                        If Trim(rst2!Name) = "" Or IsNull(rst2!Name) Then
                            !Name = Mid(XNull(rst2!Narration), 1, 35): !NARRATION1 = Mid(XNull(rst2!Narration), 36, 255)
                        Else
                            !Name = XNull(rst2!Name): !NARRATION1 = XNull(rst2!Narration)
                        End If
                    End If
                    !V_DATE = rst2!V_DATE: !CR = rst2!credit: !AdjAmt = rst2!Debit
                    !V_Type = rst2!V_Type: !v_sno = rst2!v_sno: !Val = IIf(rst2!credit > 0, "2", "3")
                    !NAME1 = Mid(AC_Name, 1, 35): !SubCode = Ac_Code
                    NarrStr = ""
                    If XNull(rst2!Chq_No) <> "" Then NarrStr = NarrStr & "Chq.No:" & Trim(XNull(rst2!Chq_No))
                    If Not IsNull(rst2!Chq_Date) And rst2!Chq_Date <> "" Then NarrStr = NarrStr & " Chq.Dt: " & Format(rst2!Chq_Date, "dd/MMM/YYYY")
                    !NARRATION2 = NarrStr
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    !CITY_NAME = XNull(Rst1!CITY_NAME)
                    !GrCode = Rst1!Code
                    !GrName = GCnFa.Execute("Select GroupName From AcGroup Where GroupCode='" & Rst1!Code & "'").Fields(0).Value
                    .Update
                End With
                rst2.MoveNext
            Loop
        End If
    End If
EXIT_LOOP:
    Rst1.MoveNext
Loop
If TmpRst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
X11 = CreateFieldDefFile(TmpRst, PubRepoPath + "\LEDGER.ttx", True)   ' dos printing
If Me.Tag = "24" And Index = 1 Then
        Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGADALET.RPT")
ElseIf MsgBox("Do You Want Each A/C on Separate Page", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
    Select Case Me.Tag
        Case "2"
            If Index = 0 Then
                Set rpt = rdApp.OpenReport(PubRepoPath + "\NEW_CLEDG.RPT")
            Else
                Set rpt = rdApp.OpenReport(PubRepoPath + "\NEWCLEDGDOS.RPT")
            End If
        Case "3"
            If Index = 0 Then
                Set rpt = rdApp.OpenReport(PubRepoPath + "\NEW_LEDG.RPT")
            Else
                Set rpt = rdApp.OpenReport(PubRepoPath + "\NEWLEDGDOS.RPT")   ' dos printing
            End If
        Case "4"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\INTEREST.RPT")
        Case "24"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGADADR.RPT")
        Case "25"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGADACR.RPT")
    End Select
Else
    Select Case Me.Tag
        Case "2"
            If Index = 0 Then
                Set rpt = rdApp.OpenReport(PubRepoPath + "\CLEDG.RPT")
            Else
                Set rpt = rdApp.OpenReport(PubRepoPath + "\CLEDGDOS.RPT")
            End If
        Case "3"
            If Index = 0 Then
                Set rpt = rdApp.OpenReport(PubRepoPath + "\LEDG.RPT")
            Else
                Set rpt = rdApp.OpenReport(PubRepoPath + "\LEDGDOS.RPT")
            End If
        Case "4"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\INTER.RPT")
        Case "24"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGDR.RPT")
        Case "25"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGCR.RPT")
    End Select
End If
For i = 1 To rpt.FormulaFields.Count
    Select Case rpt.FormulaFields(i).FormulaFieldName
        Case "TITLE"
            rpt.FormulaFields(i).Text = "'" & Me.Caption & "'"
        Case "DT"
            rpt.FormulaFields(i).Text = "'From : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
        Case "INT_RATE"
            If Me.Tag = "4" Then rpt.FormulaFields(i).Text = "" & TXTINT & ""
        Case "END_DATE"
            If Me.Tag = "4" Then rpt.FormulaFields(i).Text = "DATE(#" & Format(CDate(TXTE_DATE), "YYYY,MM,DD") & "#)"
        Case "DetYn"
           If Me.Tag = "3" Then
                If Check1.Value = vbChecked Then
                    rpt.FormulaFields(i).Text = "'" & "Y" & "'"
                End If
           End If
    End Select
Next
rpt.Database.SetDataSource TmpRst
rpt.ReadRecords
Set Rst1 = Nothing
Set rst2 = Nothing
Set TmpRst = Nothing

Select Case Me.Tag
    Case "2", "3", "24"
        Report_View rpt, Me.Caption, , False
    Case Else
        Report_View rpt, Me.Caption, , False
End Select
TXTACC_CODE = ""
TXTACC_CODE1 = ""
TXTGRP_CODE = ""
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical, Me.Caption
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
'    Erase GridArray
'    ReDim Preserve GridArray(0)
'    GridArray(0) = 0
    If ac_str = "" Then
        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
        GridSel(Gridindex).SetFocus
        
        Exit Function
    End If
    FillString = ac_str
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
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Select Case Index
    Case 0
        Grid1Sql = "SELECT '' as O,NAME as Account,SubCode as AccId from SubGroup order by name "
        GridInitialise 1, Grid1Sql, True
    Case 1
        Grid1Sql = "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup order by GroupName"
        GridInitialise 1, Grid1Sql, True
End Select
End Sub
Public Sub SelGridKeyPressLocal(txt As Object, FGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
Dim LPlace As Byte
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


