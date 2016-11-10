VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ContraRep 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contra Ledger"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11850
   Icon            =   "ContraRep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7260
   ScaleWidth      =   11850
   Begin VB.TextBox TxtE_Date 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1935
      TabIndex        =   1
      Top             =   1650
      Width           =   1395
   End
   Begin VB.TextBox TxtS_Date 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1935
      TabIndex        =   0
      Top             =   1320
      Width           =   1395
   End
   Begin VB.TextBox TxtInt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1935
      MaxLength       =   5
      TabIndex        =   8
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
      TabIndex        =   3
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
      TabIndex        =   5
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
      TabIndex        =   7
      Text            =   "0"
      Top             =   2640
      Width           =   1395
   End
   Begin VB.CommandButton BtnPrint 
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
      TabIndex        =   22
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
      TabIndex        =   20
      Top             =   3600
      Width           =   4335
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   2625
      TabIndex        =   21
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
      TabIndex        =   6
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
      TabIndex        =   4
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
      TabIndex        =   2
      Top             =   1987
      Width           =   1590
   End
   Begin VB.CommandButton BtnPrint 
      BackColor       =   &H8000000A&
      Caption         =   "&Print"
      DisabledPicture =   "ContraRep.frx":0442
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
      Picture         =   "ContraRep.frx":0584
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Print Reports"
      Top             =   5025
      Width           =   1110
   End
   Begin VB.CommandButton BtnExit 
      BackColor       =   &H8000000A&
      Caption         =   "E&xit"
      DisabledPicture =   "ContraRep.frx":06C6
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
      Picture         =   "ContraRep.frx":07C8
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Exit"
      Top             =   5025
      Width           =   1110
   End
   Begin VB.Frame Frame_GROUP 
      BackColor       =   &H00CFE0E0&
      Height          =   705
      Left            =   4920
      TabIndex        =   14
      Top             =   960
      Width           =   4470
      Begin MSDataListLib.DataCombo TxtGrp_Code 
         Height          =   315
         Left            =   45
         TabIndex        =   15
         Top             =   255
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
      TabIndex        =   16
      Top             =   960
      Width           =   4470
      Begin MSDataListLib.DataCombo TxtAcc_Code1 
         Height          =   315
         Left            =   45
         TabIndex        =   17
         Top             =   750
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo TxtAcc_Code 
         Height          =   315
         Left            =   105
         TabIndex        =   29
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
      Left            =   3435
      TabIndex        =   13
      Top             =   2535
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
      TabIndex        =   12
      Top             =   2220
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
      Top             =   1905
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
      Top             =   1590
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
      Left            =   3435
      TabIndex        =   9
      Top             =   1275
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.Frame Frame_OPTIONAL 
      BackColor       =   &H00CFE0E0&
      Height          =   2580
      Left            =   4920
      TabIndex        =   18
      Top             =   945
      Width           =   4470
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
         Height          =   2385
         Left            =   45
         TabIndex        =   19
         Top             =   150
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   4207
         _Version        =   393216
         ForeColor       =   8388608
         BackColorFixed  =   15595518
         ForeColorFixed  =   192
         BackColorBkg    =   13623520
         GridColor       =   8438015
         GridColorFixed  =   8438015
         GridLinesFixed  =   1
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
      TabIndex        =   31
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
      TabIndex        =   30
      Top             =   3630
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Selection"
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
      TabIndex        =   28
      Top             =   960
      Width           =   1395
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   1687
      Width           =   975
   End
End
Attribute VB_Name = "ContraRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const OptAll As Byte = 0
Private Const OptSelected As Byte = 1
Private Const OptOptional As Byte = 2
Private Const OptGroup As Byte = 3
Private Const OptMerge As Byte = 4

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
    If Me.Tag = "24" Or Me.Tag = "25" Then
        Call INI_COMBO("SELECT SubCode,TRIM(IIF(IsNull(SubGroup.NAME),'',SubGroup.NAME)) +TRIM(IIF(IsNull(SubGroup.ADD1) or SubGroup.ADD1='','',', '+SubGroup.ADD1)) +TRIM(IIf(IsNull(CITY.CITYName) or CITY.CITYName='','',', '+ CITY.CITYName)) As TNAME FROM SubGroup LEFT JOIN CITY ON CITY.CITYCode=SubGroup.CITYCode where Nature='" & IIf(Me.Tag = "24", "Customer", "Supplier") & "' ORDER BY SubGroup.NAME", TxtAcc_Code, "TNAME", "SubCode")
        Call INI_COMBO("SELECT SubCode,TRIM(IIF(IsNull(SubGroup.NAME),'',SubGroup.NAME)) +TRIM(IIF(IsNull(SubGroup.ADD1) or SubGroup.ADD1='','',', '+SubGroup.ADD1)) +TRIM(IIf(IsNull(CITY.CITYName) or CITY.CITYName='','',', '+CITY.CITYName)) As TNAME FROM SubGroup LEFT JOIN CITY ON CITY.CITYCode=SubGroup.CITYCode where Nature='" & IIf(Me.Tag = "24", "Customer", "Supplier") & "' ORDER BY SubGroup.NAME", TxtAcc_Code1, "TNAME", "SubCode")
        Call INI_COMBO("select Code,NAME from GROUP_TRANS where GR_Code<>'' AND GR_Code IS NOT NULL and Nature='" & IIf(Me.Tag = "24", "Customer", "Supplier") & "' order by name", TxtGrp_Code, "NAME", "Code")
    Else
        Call INI_COMBO("SELECT SubCode,TRIM(IIF(IsNull(SubGroup.NAME),'',SubGroup.NAME)) +TRIM(IIF(IsNull(SubGroup.ADD1) or SubGroup.ADD1='','',', '+SubGroup.ADD1)) +TRIM(IIf(IsNull(CITY.CITYName) or CITY.CITYName='','',', '+CITY.CITYName)) As TNAME FROM SubGroup LEFT JOIN CITY ON CITY.CITYCode=SubGroup.CITYCode ORDER BY SubGroup.NAME", TxtAcc_Code, "TNAME", "SubCode")
        Call INI_COMBO("SELECT SubCode,TRIM(IIF(IsNull(SubGroup.NAME),'',SubGroup.NAME)) +TRIM(IIF(IsNull(SubGroup.ADD1) or SubGroup.ADD1='','',', '+SubGroup.ADD1)) +TRIM(IIf(IsNull(CITY.CITYName) or CITY.CITYName='','',', '+CITY.CITYName)) As TNAME FROM SubGroup LEFT JOIN CITY ON CITY.CITYCode=SubGroup.CITYCode ORDER BY SubGroup.NAME", TxtAcc_Code1, "TNAME", "SubCode")
        Call INI_COMBO("select Code,NAME from GROUP_TRANS where GR_Code<>'' AND GR_Code IS NOT NULL order by name", TxtGrp_Code, "NAME", "Code")
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
Private Sub Form_Load()
    Call WinSetting(Me)
    BtnPrint(1).Visible = False
    TxtE_Date = PubLoginDate
    TxtS_Date = PubStartDate
    Frame_GROUP.Visible = False
    Frame_OPTIONAL.Visible = False
    Frame_SELECTED.Visible = False
End Sub
Private Sub Option1_Click(Index As Integer)
Dim Rst1 As ADODB.Recordset
Select Case Index
    Case 0
        Frame_GROUP.Visible = False
        Frame_OPTIONAL.Visible = False
        Frame_SELECTED.Visible = False
    Case 1
        Frame_GROUP.Visible = False
        Frame_OPTIONAL.Visible = False
        Frame_SELECTED.Visible = True
        TxtAcc_Code1.Visible = False
    Case 2
        Frame_GROUP.Visible = False
        Frame_OPTIONAL.Visible = True
        Frame_SELECTED.Visible = False
        FGrid1.Rows = 0
        Set Rst1 = GCnFa.Execute("select NAME,SubCode from SubGroup order by name")
        Set FGrid1.DataSource = Rst1
        FGrid1.ColWidth(0) = 345
        FGrid1.ColWidth(1) = 3745
        FGrid1.ColWidth(2) = 0
    Case 3
        Frame_GROUP.Visible = True
        Frame_OPTIONAL.Visible = False
        Frame_SELECTED.Visible = False
    Case 4
        Frame_GROUP.Visible = False
        Frame_OPTIONAL.Visible = False
        Frame_SELECTED.Visible = True
        TxtAcc_Code1.Visible = True
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

Private Sub TxtInt_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub TxtInt_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And (KeyAscii < 46 Or KeyAscii > 58) Then KeyAscii = 0
End Sub

Private Sub BTNPRINT_Click(Index As Integer)
Dim Rst1 As ADODB.Recordset, rst2 As ADODB.Recordset, Ac_Str$
Dim i As Integer, oBAL As Double, Ac_Name$, Ac_Code$, Qry$, CondStr$
Dim tmprst As ADODB.Recordset, NarrStr$, X11, rstName As ADODB.Recordset
On Error GoTo Errloop
If DateDiff("d", TxtS_Date, TxtE_Date) < 0 Then
    MsgBox " Ending Date Less than Starting Date ", vbCritical
    Exit Sub
End If
If Me.Tag = "4" Then If IsValid(TxtInt, "Enter Interest Rate") = False Then Exit Sub
If Option1(OptSelected).Value = True Then If IsValid(TxtAcc_Code, "Select A/C") = False Then Exit Sub
If Option1(OptGroup).Value = True Then If IsValid(TxtGrp_Code, "Select A/C") = False Then Exit Sub
If Option1(OptMerge).Value = True And (TxtAcc_Code = "" Or TxtAcc_Code1 = "") Then MsgBox " ** Select A/c ** ", vbCritical, Me.CAPTION: Exit Sub
If Option1(OptMerge).Value = True And TxtAcc_Code = TxtAcc_Code1 Then MsgBox " ** Both A/C are Same** ", vbCritical, Me.CAPTION: Exit Sub
GSQL = "SELECT GroupNature,GroupCode as Code,SubCode,Name,Add1,Add2,CITY_Name FROM Party_List "
If Option1(OptAll).Value = True Then
    If Me.Tag = "24" Then
        CondStr = "where Nature='Customer' "
    ElseIf Me.Tag = "25" Then
        CondStr = "where Nature='Supplier' "
    Else
        CondStr = ""
    End If
ElseIf Option1(OptSelected).Value = True Then
    CondStr = "where SubCode=" & Chk_Text(TxtAcc_Code.BoundText) & " "
ElseIf Option1(OptOptional).Value = True Then
    Ac_Str = ""
    For i = 0 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(i, 0) = "ü" Then Ac_Str = Ac_Str + IIf(Ac_Str = "", FGrid1.TextMatrix(i, 2), "," + Trim(FGrid1.TextMatrix(i, 2)))
    Next
    If Ac_Str = "" Then MsgBox " ** Select A/C ** ", vbCritical, Me.CAPTION: Exit Sub
    CondStr = " where SubCode IN (" & Ac_Str & ")"
ElseIf Option1(OptGroup).Value = True Then
    CondStr = " where Code='" & TxtGrp_Code.BoundText & "' "
ElseIf Option1(OptMerge).Value = True Then
    GSQL = "SELECT GroupNature,Code,SubCode,NAME,ADD1,ADD2,CITY_NAME " & _
        "FROM Party_List where SubCode IN ('" & TxtAcc_Code.BoundText & "','" & TxtAcc_Code1.BoundText & "')"
End If
GSQL = GSQL & CondStr & " Order By Name"
Set Rst1 = GCnFa.Execute(GSQL)

Qry = ""
If Check(2).Value = 1 Then Qry = Qry + "AND ViewLedger.Debit+ViewLedger.Credit>=" & Val(Text(2))
If Trim(Text(3)) <> "" Then Qry = Qry + " AND INSTR(1,UCASE(ViewLedger.Narration),UCASE(TRIM('" & Text(3) & "')))"
If Trim(Text(4)) <> "" Then Qry = Qry + " AND NOT INSTR(1,UCASE(ViewLedger.Narration),UCASE(TRIM('" & Text(4) & "')))"
Set tmprst = New ADODB.Recordset
Set tmprst = ADTMP1(tmprst)
Ac_Name = ""
If Option1(OptMerge).Value = True Then
    Do Until Rst1.EOF
        Ac_Name = Ac_Name + Trim(Rst1!Name) + ","
        Rst1.MoveNext
    Loop
End If
Rst1.MoveFirst
Do Until Rst1.EOF
    If Check(0).Value = 1 Then If GCnFa.Execute("SELECT IIF(IsNull(SUM(Debit)),0,SUM(Debit))-IIF(IsNull(SUM(Credit)),0,SUM(Credit)) FROM ViewLedger where PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE<=" & ConvertDate(TxtE_Date)).Fields(0) < Val(Text(0)) Then GoTo Exit_Loop
    If Check(1).Value = 1 Then If GCnFa.Execute("SELECT IIF(IsNull(SUM(Credit)),0,SUM(Credit))-IIF(IsNull(SUM(Debit)),0,SUM(Debit)) FROM ViewLedger where PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE<=" & ConvertDate(TxtE_Date)).Fields(0) < Val(Text(1)) Then GoTo Exit_Loop
    oBAL = 0
    Ac_Code = IIf(Option1(OptMerge).Value = True, 0, Rst1!SubCode)
    If Option1(OptMerge).Value = False Then Ac_Name = Trim(Rst1!Name)
    If Me.Tag = "24" Then
        If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
            oBAL = GCnFa.Execute("SELECT IIF(IsNull(SUM(Credit)),0,SUM(Credit)) FROM ViewLedger where Credit >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE <=" & ConvertDate(TxtE_Date)).Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(IsNull(SUM(Debit)),0,SUM(Debit)) FROM ViewLedger where Debit>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND (V_DATE<" & ConvertDate(TxtS_Date) & " OR (V_DATE<=" & ConvertDate(TxtE_Date) & " AND V_TYPE='F_AO'))").Fields(0)
        Else
            oBAL = GCnFa.Execute("SELECT IIF(IsNull(SUM(Credit)),0,SUM(Credit)) FROM ViewLedger where Credit >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN " & ConvertDate(PubStartDate) & " AND " & ConvertDate(TxtE_Date)).Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(IsNull(SUM(Debit)),0,SUM(Debit)) FROM ViewLedger where Debit>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN " & ConvertDate(PubStartDate) & " AND " & ConvertDate(TxtE_Date) & " AND V_TYPE='F_AO'").Fields(0)
        End If
        If oBAL < 0 Then
            With tmprst
                .AddNew
                !V_Type = "": !V_NO = 0: !V_ADD = "": !v_sno = 0: !Name = "Opening Balance"
                !V_Date = TxtS_Date: !cr = Abs(oBAL): !ADJQTY = Abs(oBAL): !Val = "1"
                !Name1 = Ac_Name: !SubCode = Ac_Code: !CITY_NAME = XNull(Rst1!CITY_NAME)
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                .Update
            End With
            oBAL = 0
        End If
        Set rst2 = GCnFa.Execute("SELECT ViewLedger.*,SubGroup.NAME FROM ViewLedger LEFT JOIN SubGroup ON SubGroup.SubCode=ViewLedger.PARTY1 where ViewLedger.Debit>0 AND ViewLedger.PARTY=" & Chk_Text(Rst1!SubCode) & " AND ViewLedger.V_DATE BETWEEN " & ConvertDate(TxtS_Date) & " AND " & ConvertDate(TxtE_Date) & " AND V_TYPE<>'F_AO' ORDER BY V_DATE,V_TYPE,V_NO")
        Do Until rst2.EOF
            If oBAL >= rst2!Debit Then
                oBAL = oBAL - rst2!Debit
            Else
                With tmprst
                    .AddNew
                    !V_ADD = rst2!V_ADD: !V_NO = rst2!V_NO: !Name = XNull(rst2!Name)
                    !V_Date = rst2!V_Date: !cr = rst2!Debit: !ADJQTY = rst2!Debit - oBAL
                    !V_Type = rst2!V_Type: !Narration1 = XNull(rst2!Narration)
                    !v_sno = rst2!v_sno: !Name1 = Ac_Name: !SubCode = Ac_Code: !CITY_NAME = XNull(Rst1!CITY_NAME)
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    .Update
                End With
                oBAL = 0
            End If
            rst2.MoveNext
        Loop
        If oBAL > 0 Then
            With tmprst
                .AddNew
                !V_Type = "": !V_NO = 0: !V_ADD = "": !v_sno = 0: !Name = "Excess Credit"
                !V_Date = TxtE_Date: !ADJAMT = Abs(oBAL): !Name1 = Ac_Name: !SubCode = Ac_Code
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                !CITY_NAME = XNull(Rst1!CITY_NAME)
                .Update
            End With
        End If
    ElseIf Me.Tag = "25" Then
        If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
            oBAL = GCnFa.Execute("SELECT IIF(IsNull(SUM(Debit)),0,SUM(Debit)) FROM ViewLedger where Debit >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE <=" & ConvertDate(TxtE_Date)).Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(IsNull(SUM(Credit)),0,SUM(Credit)) FROM ViewLedger where Credit>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND (V_DATE<" & ConvertDate(TxtS_Date) & " OR (V_DATE<=" & ConvertDate(TxtE_Date) & " AND V_TYPE='F_AO'))").Fields(0)
        Else
            oBAL = GCnFa.Execute("SELECT IIF(IsNull(SUM(Debit)),0,SUM(Debit)) FROM ViewLedger where Debit >0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN " & ConvertDate(PubStartDate) & " AND " & ConvertDate(TxtE_Date)).Fields(0)
            oBAL = oBAL - GCnFa.Execute("SELECT IIF(IsNull(SUM(Credit)),0,SUM(Credit)) FROM ViewLedger where Credit>0 AND PARTY=" & Chk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN " & ConvertDate(PubStartDate) & " AND " & ConvertDate(TxtE_Date) & " AND V_TYPE='F_AO'").Fields(0)
        End If
        If oBAL < 0 Then
            With tmprst
                .AddNew
                !V_Type = "": !V_NO = 0: !V_ADD = "": !v_sno = 0: !Narration1 = "OPENING BALANCE"
                !Narration2 = "": !Name = "OPENING BALANCE": !V_Date = TxtS_Date: !cr = Abs(oBAL)
                !ADJQTY = Abs(oBAL): !Val = "1": !Name1 = Ac_Name: !SubCode = Ac_Code
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                !CITY_NAME = XNull(Rst1!CITY_NAME)
                .Update
            End With
            oBAL = 0
        End If
        Set rst2 = GCnFa.Execute("SELECT ViewLedger.*,SubGroup.NAME FROM ViewLedger LEFT JOIN SubGroup ON SubGroup.SubCode=ViewLedger.PARTY1 where ViewLedger.Credit>0 AND ViewLedger.PARTY=" & Chk_Text(Rst1!SubCode) & " AND ViewLedger.V_DATE BETWEEN " & ConvertDate(TxtS_Date) & " AND " & ConvertDate(TxtE_Date) & " AND V_tYPE<>'F_AO' ORDER BY V_DATE,V_TYPE,V_NO")
        Do Until rst2.EOF
            If oBAL >= rst2!Credit Then
                oBAL = oBAL - rst2!Credit
            Else
                With tmprst
                    .AddNew
                    !V_ADD = rst2!V_ADD: !V_NO = rst2!V_NO: !Name = rst2!Name: !V_Date = rst2!V_Date
                    !cr = rst2!Credit: !ADJQTY = rst2!Credit - oBAL: !V_Type = rst2!V_Type
                    !Narration1 = XNull(rst2!Narration): !v_sno = rst2!v_sno: !Name1 = Ac_Name
                    !SubCode = Ac_Code: !CITY_NAME = XNull(Rst1!CITY_NAME)
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    .Update
                End With
                oBAL = 0
            End If
            rst2.MoveNext
        Loop
        If oBAL > 0 Then
            With tmprst
                .AddNew
                !V_Type = "": !V_NO = 0: !V_ADD = "": !v_sno = 0: !Name = "Excess Debit"
                !V_Date = TxtE_Date: !ADJAMT = Abs(oBAL): !Name1 = Ac_Name: !SubCode = Ac_Code
                !Address1 = Trim(IIf(IsNull(Rst1!Add1), " ", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), " ", ", " + Rst1!Add2)))
                !CITY_NAME = XNull(Rst1!CITY_NAME)
                .Update
            End With
        End If
    Else
        If Qry = "" Then
            If Me.Tag = "2" Or Me.Tag = "3" Or Me.Tag = "4" Then
                If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
                    oBAL = oBAL + GCnFa.Execute("SELECT IIF(IsNull(SUM(Credit)),0,SUM(Credit))-IIF(IsNull(SUM(Debit)),0,SUM(Debit)) FROM ViewLedger where (V_DATE<" & ConvertDate(TxtS_Date) & " OR (V_DATE BETWEEN  " & ConvertDate(TxtS_Date) & " AND " & ConvertDate(TxtE_Date) & " and V_TYPE='F_AO'))  AND PARTY=" & Chk_Text(Rst1!SubCode)).Fields(0)
                Else
                    oBAL = oBAL + GCnFa.Execute("SELECT IIF(IsNull(SUM(Credit)),0,SUM(Credit))-IIF(IsNull(SUM(Debit)),0,SUM(Debit)) FROM ViewLedger where (V_DATE BETWEEN  " & ConvertDate(TxtS_Date) & " AND " & ConvertDate(TxtE_Date) & " AND V_TYPE='F_AO') AND PARTY=" & Chk_Text(Rst1!SubCode)).Fields(0)
                End If
            End If
            If oBAL <> 0 Then
                With tmprst
                    .AddNew
                    !V_Type = "": !V_NO = 0: !V_ADD = "": !v_sno = 0: !Name = "OPENING BALANCE"
                    !V_Date = TxtS_Date: !cr = IIf(oBAL > 0, Abs(oBAL), 0)
                    !ADJAMT = IIf(oBAL < 0, Abs(oBAL), 0): !Val = "1": !Name1 = Ac_Name
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), " ", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), " ", ", " + Rst1!Add2)))
                    !CITY_NAME = XNull(Rst1!CITY_NAME): !SubCode = Ac_Code
                    .Update
                End With
            End If
        End If
        If Me.Tag = "2" Or Me.Tag = "3" Or Me.Tag = "4" Then
            Set rst2 = GCnFa.Execute("SELECT ViewLedger.*,SubGroup.NAME FROM ViewLedger LEFT JOIN SubGroup ON ViewLedger.PARTY1 = SubGroup.SubCode where ViewLedger.V_DATE Between " & ConvertDate(TxtS_Date) & " And " & ConvertDate(TxtE_Date) & " AND (ViewLedger.PARTY=" & Chk_Text(Rst1!SubCode) & ") AND V_TYPE<>'F_AO'" & Qry)
            Do Until rst2.EOF
                 With tmprst
                    .AddNew
                    !V_ADD = rst2!V_ADD: !V_NO = rst2!V_NO
                    If Me.Tag = "3" Then
                        !Name = Mid(XNull(rst2!Narration), 1, 35): !Narration1 = Mid(XNull(rst2!Narration), 36, 255)
                    Else
                        If Trim(rst2!Name) = "" Or IsNull(rst2!Name) Then
                            !Name = Mid(XNull(rst2!Narration), 1, 35): !Narration1 = Mid(XNull(rst2!Narration), 36, 255)
                        Else
                            !Name = XNull(rst2!Name): !Narration1 = XNull(rst2!Narration)
                        End If
                    End If
                    !V_Date = rst2!V_Date: !cr = rst2!Credit: !ADJAMT = rst2!Debit
                    !V_Type = rst2!V_Type: !v_sno = rst2!v_sno: !Val = IIf(rst2!Credit > 0, "2", "3")
                    !Name1 = Mid(Ac_Name, 1, 35): !SubCode = Ac_Code
                    NarrStr = ""
                    If Trim(XNull(rst2!CHQ_NO)) <> "" Then
                        NarrStr = NarrStr + "Chq.No:" + Trim(XNull(rst2!CHQ_NO))
                    End If
                    If Not IsNull(rst2!Chq_Date) Then NarrStr = NarrStr + " Chq.Dt: " + str(rst2!Chq_Date)
                    !Narration2 = NarrStr
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    !CITY_NAME = XNull(Rst1!CITY_NAME)
                    .Update
                End With
                rst2.MoveNext
            Loop
        End If
    End If
Exit_Loop:
    Rst1.MoveNext
Loop
If tmprst.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: Exit Sub
X11 = CreateFieldDefFile(tmprst, PubRepoPath + "\LEDGER.ttx", True)
If Me.Tag = "24" And Index = 1 Then
        Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGADALET.RPT")
''ElseIf MsgBox("Do You Want Each A/C on Separate Page", vbYesNo + vbQuestion + vbDefaultButton2, Me.CAPTION) = vbYes Then
Else
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
                Set rpt = rdApp.OpenReport(PubRepoPath + "\NEWLEDGDOS.RPT")
            End If
        Case "4"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\INTEREST.RPT")
        Case "24"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGADADR.RPT")
        Case "25"
            Set rpt = rdApp.OpenReport(PubRepoPath + "\TAGADACR.RPT")
    End Select
End If
For i = 1 To rpt.FormulaFields.Count
    Select Case rpt.FormulaFields(i).FormulaFieldName
        Case "TITLE"
            rpt.FormulaFields(i).Text = "'" & Me.CAPTION & "'"
        Case "DT"
            rpt.FormulaFields(i).Text = "'From : " & TxtS_Date & " To : " & TxtE_Date & "'"
        Case "INT_RATE"
            If Me.Tag = "4" Then rpt.FormulaFields(i).Text = "" & TxtInt & ""
        Case "END_DATE"
            If Me.Tag = "4" Then rpt.FormulaFields(i).Text = "DATE(" & Format(CDate(TxtE_Date), "YYYY,MM,DD") & ")"
    End Select
Next
rpt.Database.SetDataSource tmprst
rpt.ReadRecords
TxtAcc_Code = ""
TxtAcc_Code1 = ""
TxtGrp_Code = ""
Set Rst1 = Nothing
Set rst2 = Nothing
Set tmprst = Nothing
Select Case Me.Tag
    Case "2", "3", "24"
        Report_View rpt, Me.CAPTION, Index
    Case Else
        Report_View rpt, Me.CAPTION, 0
End Select
Exit Sub
Errloop:    MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TXTE_DATE_Validate(Cancel As Boolean)
    TxtE_Date = RetDate(TxtE_Date)
End Sub

Private Sub TXTS_DATE_Validate(Cancel As Boolean)
    TxtS_Date = RetDate(TxtS_Date)
End Sub
