VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FaAdjust 
   BackColor       =   &H00E6AC86&
   Caption         =   "Adjustment Entry"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   Icon            =   "FaAdjust.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11550
   Begin VB.TextBox TXT_DATE 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F0DF&
      Height          =   300
      Left            =   1995
      MaxLength       =   12
      TabIndex        =   1
      Top             =   120
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   7935
      Width           =   9405
      Begin VB.Label COMP_LABEL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   23
         Top             =   165
         Width           =   5070
      End
   End
   Begin VB.CommandButton BtnFill 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fill Detail"
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
      Left            =   8025
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   195
      Width           =   1365
   End
   Begin VB.CommandButton BtnAdd 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Add"
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
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add"
      Top             =   195
      Width           =   1365
   End
   Begin VB.CommandButton BtnExit 
      BackColor       =   &H00C0FFFF&
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
      Height          =   390
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit"
      Top             =   195
      Width           =   1365
   End
   Begin MSDataListLib.DataCombo TXT_ACCOUNT 
      Height          =   315
      Left            =   1995
      TabIndex        =   3
      Top             =   765
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16249055
      Text            =   ""
   End
   Begin VB.Frame FRAMEADJUST 
      BackColor       =   &H00B7A4F0&
      Caption         =   "Adjustment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5430
      Left            =   -645
      TabIndex        =   8
      Top             =   2880
      Width           =   11835
      Begin VB.CommandButton BTS_AUTO_ADJ 
         Caption         =   "Auto"
         DisabledPicture =   "FaAdjust.frx":000C
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11055
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ok"
         Top             =   450
         Width           =   585
      End
      Begin VB.TextBox TXTADJ_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F0DF&
         Height          =   285
         Left            =   4905
         TabIndex        =   24
         Top             =   0
         Width           =   1065
      End
      Begin VB.TextBox TXTNARRATION 
         BackColor       =   &H00D5E2E3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   795
         Left            =   90
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   4590
         Visible         =   0   'False
         Width           =   9210
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   11055
         Picture         =   "FaAdjust.frx":010E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Full Adjuatments"
         Top             =   945
         Width           =   585
      End
      Begin VB.CommandButton ADJ_OK 
         DisabledPicture =   "FaAdjust.frx":0550
         Height          =   495
         Left            =   11055
         Picture         =   "FaAdjust.frx":0652
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ok"
         Top             =   1440
         Width           =   585
      End
      Begin VB.CommandButton ADJ_CANCLE 
         DisabledPicture =   "FaAdjust.frx":0794
         Height          =   495
         Left            =   11055
         Picture         =   "FaAdjust.frx":08D6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel Changes"
         Top             =   1935
         Width           =   585
      End
      Begin MSFlexGridLib.MSFlexGrid FgridAdjust 
         Height          =   4125
         Left            =   90
         TabIndex        =   13
         Top             =   465
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   7276
         _Version        =   393216
         Cols            =   13
         BackColor       =   14873589
         ForeColor       =   64
         BackColorSel    =   8388608
         ForeColorSel    =   65535
         BackColorBkg    =   12035312
         GridColor       =   0
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"FaAdjust.frx":0A18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label ADJ_LAB7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "KK"
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
         Height          =   240
         Left            =   5625
         TabIndex        =   19
         Top             =   195
         Width           =   1050
      End
      Begin VB.Label ADJ_LAB6 
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "Pend."
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
         Height          =   240
         Left            =   4770
         TabIndex        =   18
         Top             =   195
         Width           =   870
      End
      Begin VB.Label ADJ_LAB3 
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "Tr.Amt."
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
         Height          =   240
         Left            =   2280
         TabIndex        =   17
         Top             =   195
         Width           =   675
      End
      Begin VB.Label ADJ_LAB4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MMMM"
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
         Height          =   240
         Left            =   2955
         TabIndex        =   16
         Top             =   195
         Width           =   1290
      End
      Begin VB.Label ADJ_LAB5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   4245
         TabIndex        =   15
         Top             =   195
         Width           =   510
      End
      Begin VB.Label ADJ_LAB8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   6720
         TabIndex        =   14
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.Frame FRAMELEDGER 
      BackColor       =   &H00A9A765&
      Caption         =   "Credit List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5080
      Left            =   0
      TabIndex        =   7
      Top             =   1290
      Width           =   11190
      Begin VB.TextBox TxtNARATION1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   660
         Left            =   30
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   4380
         Width           =   10335
      End
      Begin VB.CommandButton BTSADJUST 
         Caption         =   "A&djust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10440
         TabIndex        =   6
         Top             =   1155
         Width           =   705
      End
      Begin MSFlexGridLib.MSFlexGrid FGridLedger 
         Height          =   4125
         Left            =   45
         TabIndex        =   5
         Top             =   255
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   7276
         _Version        =   393216
         Cols            =   13
         BackColor       =   14807499
         BackColorBkg    =   11118437
         GridColor       =   0
         FocusRect       =   0
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"FaAdjust.frx":0AC5
      End
   End
   Begin MSDataListLib.DataCombo TXT_Group 
      Height          =   315
      Left            =   1995
      TabIndex        =   2
      Top             =   435
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16249055
      Text            =   ""
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "For A/c Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   28
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Upto Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   870
      TabIndex        =   27
      Top             =   150
      Width           =   1065
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "For Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   690
      TabIndex        =   26
      Top             =   810
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   525
      Left            =   6570
      Shape           =   4  'Rounded Rectangle
      Top             =   135
      Width           =   4275
   End
End
Attribute VB_Name = "FaAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB_LE As ADODB.Recordset
Private CURR_ROW As Integer, AMT_ADJ As Double, AC__AMT As Double, OLD_AMT1 As Double
Private PubDatamanFa As New DMFa.ClsFa

Private Sub ADJ_CANCLE_Click()
    BTSADJUST.Enabled = False
    BtnAdd.Enabled = True
    BtnExit.Enabled = True
    BTNFILL.Enabled = True
    TXT_DATE.Enabled = False
    TXT_Group.Enabled = False
    TXT_ACCOUNT.Enabled = False
    FRAMELEDGER.Visible = True
    FRAMELEDGER.ZOrder 0
    BTSADJUST.Enabled = True
    TXTNARRATION.Visible = False
    TXTADJ_AMT.Visible = False
    FRAMEADJUST.Visible = False
End Sub
Private Sub ADJ_OK_Click()
Dim K As Integer, BeginTrans As Byte
On Error GoTo ELoop
    BeginTrans = 0
    BTSADJUST.Enabled = False
    BtnAdd.Enabled = True
    BtnExit.Enabled = True
    BTNFILL.Enabled = True
    TXT_DATE.Enabled = False
    TXT_Group.Enabled = True
    TXT_ACCOUNT.Enabled = False
    FRAMELEDGER.Visible = True
    FRAMELEDGER.ZOrder 0
    BTSADJUST.Enabled = True
    TXTNARRATION.Visible = False
    TXTADJ_AMT.Visible = False
    FRAMEADJUST.Visible = False
    G_FaCn.BeginTrans
    BeginTrans = 1
    For K = 1 To FgridAdjust.Rows - 1
        If Val(FgridAdjust.TextMatrix(K, 9)) > 0 Then
            G_FaCn.Execute ("Insert Into LEDGERAdj (DocId1,V_SNo1,DocId2,V_SNo2,CR,SubCode,U_Name,U_EntDt,U_AE) Values ('" & FGridLedger.TextMatrix(FGridLedger.Row, 12) & "'," & Val(FGridLedger.TextMatrix(FGridLedger.Row, 5)) & ",'" & FgridAdjust.TextMatrix(K, 12) & "'," & Val(FgridAdjust.TextMatrix(K, 4)) & "," & Val(FgridAdjust.TextMatrix(K, 9)) & ",'" & TXT_ACCOUNT.BoundText & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A')")
            FGridLedger.TextMatrix(FGridLedger.Row, 10) = Format(Val(FGridLedger.TextMatrix(FGridLedger.Row, 10)) + Val(FgridAdjust.TextMatrix(K, 9)), "0.00")
            FGridLedger.TextMatrix(FGridLedger.Row, 11) = Format(Val(FGridLedger.TextMatrix(FGridLedger.Row, 7)) - Val(FGridLedger.TextMatrix(FGridLedger.Row, 10)), "0.00")
        End If
    Next
    FgridAdjust.Row = 0
    TXTNARRATION.TEXT = FgridAdjust.TextMatrix(FgridAdjust.Row, 10)
    G_FaCn.CommitTrans
    BeginTrans = 0
Exit Sub
ELoop:  G_FaCn.RollbackTrans
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FgridAdjust_RowColChange()
    TXTNARRATION.TEXT = FgridAdjust.TextMatrix(FgridAdjust.Row, 10)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
Private Sub Form_Load()
    BTSADJUST.Enabled = False
    TXT_DATE.Enabled = False
    TXT_ACCOUNT.Enabled = False
    TXT_Group.Enabled = False
    CURR_ROW = 1
    Me.left = 0
    Me.top = 0
    Me.height = 6975
    Me.width = 11300
    '''''''''''''
    PubDatamanFa.FaBackEnd = PubBackEnd
    PubDatamanFa.FaPubLoginDate = PubLoginDate
    PubDatamanFa.FaPubDivCode = PubDivCode
    PubDatamanFa.FaPubSiteCode = PubSiteCode
    PubDatamanFa.FaPubSiteCodeDisplay = PubSiteCodeDisplay
    PubDatamanFa.FaPubSiteName = PubSiteName
    PubDatamanFa.FapubUName = pubUName
    PubDatamanFa.FaDosPort = PubFaDosPort
    PubDatamanFa.FaRunPIF = PubRunPIF
    PubDatamanFa.FaPubSiteType = PubFaSiteType
    Set PubDatamanFa.SetG_FaCn = G_FaCn
    Set PubDatamanFa.SetG_CompCn = G_CompCn
    Set PubDatamanFa.SetrsUserPerm = rsUserPerm.Clone
    Set PubDatamanFa.SetMasterRst = FaMasterRst.Clone
    '''''''''''''
    FRAMELEDGER.left = 0
    FRAMELEDGER.top = 1290
    FRAMELEDGER.width = 11190
    FRAMELEDGER.height = 5080
    FRAMEADJUST.left = FRAMELEDGER.left
    FRAMEADJUST.top = FRAMELEDGER.top
    FRAMEADJUST.width = FRAMELEDGER.width
    FRAMEADJUST.height = FRAMELEDGER.height
    FgridAdjust.left = FGridLedger.left
    FgridAdjust.top = FGridLedger.top + 200
    FgridAdjust.width = FGridLedger.width
    FgridAdjust.height = FGridLedger.height - 200
    TXTNARRATION.top = TxtNARATION1.top
    TXTNARRATION.left = TxtNARATION1.left
    TXTNARRATION.width = TxtNARATION1.width
    TXTNARRATION.height = TxtNARATION1.height
    BTS_AUTO_ADJ.left = 10450
    Command2.left = 10450
    ADJ_OK.left = 10450
    ADJ_CANCLE.left = 10450
    FGridLedger.ColAlignment(1) = flexAlignLeftCenter
    FGridLedger.ColAlignment(3) = flexAlignLeftCenter
    FGridLedger.ColAlignment(7) = flexAlignRightCenter
    FGridLedger.ColAlignment(10) = flexAlignRightCenter
    FGridLedger.ColAlignment(11) = flexAlignRightCenter
    FGridLedger.ColWidth(8) = 0
    FGridLedger.ColWidth(9) = 0
    FGridLedger.ColWidth(4) = 0
    FGridLedger.ColWidth(12) = 0
    FgridAdjust.ColAlignment(2) = flexAlignLeftCenter
    FgridAdjust.ColAlignment(5) = flexAlignLeftCenter
    FgridAdjust.ColAlignment(7) = flexAlignRightCenter
    FgridAdjust.ColAlignment(8) = flexAlignRightCenter
    FgridAdjust.ColAlignment(9) = flexAlignRightCenter
    FgridAdjust.ColWidth(3) = 0
    FgridAdjust.ColWidth(10) = 0
    FgridAdjust.ColWidth(11) = 0
    FgridAdjust.ColWidth(12) = 0
    TXTADJ_AMT.Visible = False
    FRAMEADJUST.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PubDatamanFa = Nothing
End Sub

Private Sub TXT_ACCOUNT_Click(Area As Integer)
    BTNFILL.Enabled = True
End Sub
Private Sub FGridLEDGER_RowColChange()
    TxtNARATION1.TEXT = FGridLedger.TextMatrix(FGridLedger.Row, 8)
End Sub
Private Sub BtnAdd_Click()
    FgridAdjust.Rows = 1
    FgridAdjust.AddItem ""
    FGridLedger.Rows = 1
    FGridLedger.AddItem ""
    TXT_DATE.Enabled = True
    TXT_DATE = Format(Now, "Short Date")
    TXT_ACCOUNT.Enabled = True
    TXT_Group.Enabled = True
    TXT_DATE.SetFocus
    FaIniCombo "SELECT GroupName,GroupCode FROM ACGROUP ORDER BY GroupName", TXT_Group, "GroupName", "GroupCode"
    If PubBackEnd = "A" Then
        FaIniCombo "SELECT NAME,SUBCODE FROM SUBGROUP WHERE SUBCODE IN (SELECT SUBGROUP.SubCode " & _
        " FROM (Ledger LEFT JOIN LedgerAdj ON (Ledger.DocId = LedgerAdj.DocId1) AND (Ledger.V_SNo = LedgerAdj.V_SNo1)) LEFT JOIN SUBGROUP ON Ledger.SubCode = SUBGROUP.SubCode Where (((Ledger.AmtCr) > 0)) GROUP BY SUBGROUP.SubCode, SUBGROUP.Name, Ledger.DocId, Ledger.V_SNo Having MAX(Ledger.AmtCr) > IIf(IsNull(Sum(LEDGERADJ.cr)), 0, Sum(LEDGERADJ.cr))) ORDER BY NAME", TXT_ACCOUNT, "NAME", "SUBCODE"
    ElseIf PubBackEnd = "S" Then
        FaIniCombo "SELECT NAME,SUBCODE FROM SUBGROUP WHERE SUBCODE IN (SELECT SUBGROUP.SubCode " & _
        " FROM (Ledger LEFT JOIN LedgerAdj ON (Ledger.DocId = LedgerAdj.DocId1) AND (Ledger.V_SNo = LedgerAdj.V_SNo1)) LEFT JOIN SUBGROUP ON Ledger.SubCode = SUBGROUP.SubCode Where (((Ledger.AmtCr) > 0)) GROUP BY SUBGROUP.SubCode, SUBGROUP.Name, Ledger.DocId, Ledger.V_SNo Having MAX(Ledger.AmtCr) > IsNull(Sum(LEDGERADJ.cr), 0)) ORDER BY NAME", TXT_ACCOUNT, "NAME", "SUBCODE"
    End If
End Sub
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub BTNFILL_Click()
Dim J As Integer, Already_Adjusted As Double, G_Rs As ADODB.Recordset
FGridLedger.Rows = 1
If TXT_ACCOUNT.TEXT <> "" Then
    Set G_Rs = G_FaCn.Execute("Select L.DocId,V_SNo,L.V_Type,L.V_Prefix,L.V_No,L.V_Date,AmtCr As Credit,M.NARRATION+' '+L.Narration AS NARR,ContraSub AS Party1,SG.Name as PartyName From (Ledger L Left Join SubGroup SG on SG.SubCode=L.CONTRASUB) LEFT JOIN LEDGERM M ON M.DOCID=L.DOCID Where L.SubCode='" & TXT_ACCOUNT.BoundText & "' AND L.V_Date<=" & FaConvertDate(TXT_DATE) & " AND AMTCR>0 Order By L.V_Date,L.V_Type,L.V_No")
    Do Until G_Rs.EOF
        If PubBackEnd = "S" Then
            Already_Adjusted = G_FaCn.Execute("Select IsNull(Sum(CR),0) From LEDGERAdj Where DocId1='" & G_Rs!DocID & "' And V_SNo1=" & G_Rs!V_SNo & " And SubCode='" & TXT_ACCOUNT.BoundText & "'").Fields(0)
        ElseIf PubBackEnd = "A" Then
            Already_Adjusted = G_FaCn.Execute("Select IIF(IsNull(Sum(CR)),0,Sum(CR)) From LEDGERAdj Where DocId1='" & G_Rs!DocID & "' And V_SNo1=" & G_Rs!V_SNo & " And SubCode='" & TXT_ACCOUNT.BoundText & "'").Fields(0)
        End If
        If Already_Adjusted < G_Rs!CREDIT Then
            FGridLedger.AddItem "" & Chr(9) & Format(G_Rs!V_DATE, "Short Date") & Chr(9) & G_Rs!V_Type & Chr(9) & G_Rs!V_NO & Chr(9) & G_Rs!v_Prefix & Chr(9) & G_Rs!V_SNo & Chr(9) & G_Rs!PartyName & Chr(9) & Format(G_Rs!CREDIT, "0.00") & Chr(9) & G_Rs!nARR & Chr(9) & G_Rs!Party1 & Chr(9) & Format(Already_Adjusted, "0.00") & Chr(9) & Format(G_Rs!CREDIT - Already_Adjusted, "0.00") & Chr(9) & G_Rs!DocID
        End If
        G_Rs.MoveNext
    Loop
    If FGridLedger.Rows = 1 Then
        FGridLedger.AddItem ""
    Else
        TxtNARATION1.TEXT = FGridLedger.TextMatrix(FGridLedger.Row, 8)
    End If
    If FGridLedger.TextMatrix(1, 1) = "" Then MsgBox "* No Entries to Adjust *", vbInformation, Me.CAPTION: Exit Sub
    BTSADJUST.Enabled = True
    BtnAdd.Enabled = False
    BTNFILL.Enabled = False
End If
End Sub
Private Sub BTSADJUST_Click()
Dim TEMP_AMT_ADJ, ADJ_ADD, THIS_VR, THIS_OT_VR, J As Integer
    If Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) <= 0 Then MsgBox "* Amount Already Adjusted *", vbInformation, Me.CAPTION: Exit Sub
    FRAMELEDGER.Visible = False
    FRAMEADJUST.Visible = True
    FgridAdjust.Redraw = False
    AC__AMT = 0
    AMT_ADJ = 0
    THIS_OT_VR = 0
    THIS_VR = 0
    ADJ_LAB4.CAPTION = FGridLedger.TextMatrix(FGridLedger.Row, 11)
    ADJ_LAB5.CAPTION = "Cr."
    FgridAdjust.Rows = 1
    Set DB_LE = G_FaCn.Execute("Select DocId,V_SNo,V_Type,V_Prefix,V_No,V_Date,AmtDr As Amount,Narration,SG.Name As PartyName FROM Ledger L Left Join SubGroup SG on SG.SubCode=L.SubCode Where L.SubCode='" & TXT_ACCOUNT.BoundText & "' ORDER BY V_Date,V_Type,V_No")
    Do Until DB_LE.EOF
        If PubBackEnd = "S" Then
            AC__AMT = G_FaCn.Execute("SELECT IsNull(Sum(CR),0) AS TSum From LEDGERAdj Where SubCode='" & TXT_ACCOUNT.BoundText & "' And DocId2='" & DB_LE!DocID & "' And V_SNo2=" & DB_LE!V_SNo).Fields(0)
        ElseIf PubBackEnd = "A" Then
            AC__AMT = G_FaCn.Execute("SELECT IIF(IsNull(Sum(CR)),0,Sum(CR)) AS TSum From LEDGERAdj Where SubCode='" & TXT_ACCOUNT.BoundText & "' And DocId2='" & DB_LE!DocID & "' And V_SNo2=" & DB_LE!V_SNo).Fields(0)
        End If
        If DB_LE!AMOUNT - AC__AMT > 0 Then
            FgridAdjust.AddItem "" & Chr(9) & DB_LE!V_Type & Chr(9) & DB_LE!V_NO & Chr(9) & DB_LE!v_Prefix & Chr(9) & DB_LE!V_SNo & Chr(9) & Format(DB_LE!V_DATE, "Short Date") & Chr(9) & DB_LE!PartyName & Chr(9) & Format(DB_LE!AMOUNT, "0.00") & Chr(9) & Format(DB_LE!AMOUNT - AC__AMT + THIS_VR + THIS_OT_VR, "0.00") & Chr(9) & Format(THIS_VR, "0.00") & Chr(9) & DB_LE!Narration & Chr(9) & "" & Chr(9) & DB_LE!DocID
        End If
        DB_LE.MoveNext
    Loop
    If FgridAdjust.Rows <= 1 Then
        FgridAdjust.Redraw = True
        MsgBox "   Entries Not Found   ", vbInformation, Me.CAPTION
        BTSADJUST.Enabled = False
        BtnAdd.Enabled = True
        BtnExit.Enabled = True
        BTNFILL.Enabled = True
        TXT_DATE.Enabled = False
        TXT_ACCOUNT.Enabled = False
        FRAMELEDGER.Visible = True
        FRAMELEDGER.ZOrder 0
        BTSADJUST.Enabled = True
        TXTNARRATION.Visible = False
        TXTADJ_AMT.Visible = False
        FRAMEADJUST.Visible = False
    Exit Sub
    End If
    AMT_ADJ = 0
    FRAMEADJUST.Visible = True
    FRAMEADJUST.ZOrder 0
    TXTNARRATION.Visible = True
    BTSADJUST.Enabled = False
    ADJ_CANCLE.Enabled = True
    ADJ_OK.Enabled = True
    BTS_AUTO_ADJ.Enabled = True
    ADJ_LAB7.CAPTION = Format((Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - AMT_ADJ), "0.00")
    ADJ_LAB8.CAPTION = "Dr"
    TXTNARRATION.TEXT = FgridAdjust.TextMatrix(FgridAdjust.Row, 10)
    FgridAdjust.Redraw = True
    BtnExit.Enabled = False
    FgridAdjust.SetFocus
End Sub
Private Sub FGRIDADJUST_Click()
    UPD_ADJ
End Sub
Private Sub FGRIDADJUST_KeyUp(KeyCode As Integer, Shift As Integer)
    UPD_ADJ
End Sub
Private Sub FGRIDADJUST_Scroll()
    UPD_ADJ
End Sub
Private Sub BTS_AUTO_ADJ_Click()
Dim K As Integer, OLD_AMT_ADJ As Long
For K = FgridAdjust.Row To FgridAdjust.Rows - 1
    If Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) > AMT_ADJ Then
        If Val(FgridAdjust.TextMatrix(K, 9)) < Val(FgridAdjust.TextMatrix(K, 8)) And AMT_ADJ <= FGridLedger.TextMatrix(FGridLedger.Row, 11) Then
            OLD_AMT_ADJ = Val(FgridAdjust.TextMatrix(K, 9))
            If (Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - (AMT_ADJ - OLD_AMT_ADJ)) >= Val(FgridAdjust.TextMatrix(K, 8)) Then
                FgridAdjust.TextMatrix(K, 9) = Val(FgridAdjust.TextMatrix(K, 8))
            Else
                FgridAdjust.TextMatrix(K, 9) = (Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - (AMT_ADJ - OLD_AMT_ADJ))
            End If
        End If
    End If
    UPD_ADJ
Next
End Sub
Private Sub Command2_Click()
Dim OLD_AMT_ADJ As Long
If Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 9)) < Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 8)) And AMT_ADJ < FGridLedger.TextMatrix(FGridLedger.Row, 11) Then
    OLD_AMT_ADJ = Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 9))
    If (Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - (AMT_ADJ - OLD_AMT_ADJ)) >= Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 8)) Then
        FgridAdjust.TextMatrix(FgridAdjust.Row, 9) = Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 8))
    Else
        FgridAdjust.TextMatrix(FgridAdjust.Row, 9) = (Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - (AMT_ADJ - OLD_AMT_ADJ))
    End If
    UPD_ADJ
End If
End Sub
Private Sub UPD_ADJ()
CURR_ROW = FgridAdjust.Row
If FgridAdjust.Col = 9 Then
    TXTADJ_AMT.Visible = True
    TXTADJ_AMT.ZOrder 0
    TXTADJ_AMT.width = FgridAdjust.ColWidth(9)
    TXTADJ_AMT.top = FgridAdjust.top + FgridAdjust.CellTop
    TXTADJ_AMT.left = FgridAdjust.left + FgridAdjust.CellLeft
    TXTADJ_AMT.TEXT = Val(FgridAdjust.TextMatrix(CURR_ROW, 9))
    TXTADJ_AMT.SetFocus
End If
CAL_ADJ_TOT
End Sub
Private Sub TXT_DATE_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub TXT_DATE_Validate(Cancel As Boolean)
    TXT_DATE = PubDatamanFa.FaRetDateFunc(TXT_DATE)
End Sub

Private Sub TXT_Group_Change()
    TXT_ACCOUNT.BoundText = ""
End Sub
Private Sub TXT_Group_Validate(Cancel As Boolean)
Dim mQRY1 As String
mQRY1 = ""
If TXT_Group.BoundText <> "" Then
    mQRY1 = " And GROUPCODE='" & TXT_Group.BoundText & "'"
    If PubBackEnd = "A" Then
        FaIniCombo "SELECT NAME,SUBCODE FROM SUBGROUP WHERE SUBCODE IN (SELECT SUBGROUP.SubCode  FROM (Ledger LEFT JOIN LedgerAdj ON (Ledger.DocId = LedgerAdj.DocId1) AND (Ledger.V_SNo=LedgerAdj.V_SNo1)) LEFT JOIN SUBGROUP ON Ledger.SubCode=SUBGROUP.SubCode Where (((Ledger.AmtCr) > 0)) GROUP BY SUBGROUP.SubCode, SUBGROUP.Name, Ledger.DocId, Ledger.V_SNo Having MAX(Ledger.AmtCr) > IIf(IsNull(Sum(LEDGERADJ.cr)), 0, Sum(LEDGERADJ.cr))) " & mQRY1 & " ORDER BY NAME", TXT_ACCOUNT, "NAME", "SUBCODE"
    ElseIf PubBackEnd = "S" Then
        FaIniCombo "SELECT NAME,SUBCODE FROM SUBGROUP WHERE SUBCODE IN (SELECT SUBGROUP.SubCode FROM (Ledger LEFT JOIN LedgerAdj ON (Ledger.DocId = LedgerAdj.DocId1) AND (Ledger.V_SNo = LedgerAdj.V_SNo1)) LEFT JOIN SUBGROUP ON Ledger.SubCode = SUBGROUP.SubCode Where (((Ledger.AmtCr) > 0)) GROUP BY SUBGROUP.SubCode, SUBGROUP.Name, Ledger.DocId, Ledger.V_SNo Having MAX(Ledger.AmtCr) > IsNull(Sum(LEDGERADJ.cr), 0)) " & mQRY1 & " ORDER BY NAME", TXT_ACCOUNT, "NAME", "SUBCODE"
    End If
End If
End Sub
Private Sub TXTADJ_AMT_GotFocus()
    OLD_AMT1 = Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 9))
    SendKeys "{Home}+{End}"
End Sub
Private Sub TXTADJ_AMT_KeyPress(KeyAscii As Integer)
    FaNumPress TXTADJ_AMT, KeyAscii, 10, 2
End Sub
Private Sub TXTADJ_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
    FaNumDown TXTADJ_AMT, KeyCode, 10, 2
End Sub
Private Sub TXTADJ_AMT_Validate(Cancel As Boolean)
    TXTADJ_AMT = FaValidate_Numeric(TXTADJ_AMT)
    If Val(TXTADJ_AMT) > Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - AMT_ADJ + OLD_AMT1 Or Val(FgridAdjust.TextMatrix(CURR_ROW, 8)) < Val(TXTADJ_AMT.TEXT) Then
        MsgBox " Amount is Greater Then Pendng Adj.Amt. You Can Adjust Only " + LTrim(RTrim(IIf(Val(FgridAdjust.TextMatrix(CURR_ROW, 8)) > Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - AMT_ADJ + OLD_AMT1, Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - AMT_ADJ + OLD_AMT1, FgridAdjust.TextMatrix(CURR_ROW, 8)))) + " Here", vbCritical, Me.CAPTION
        FgridAdjust.Row = CURR_ROW
        Cancel = True
    Else
        FgridAdjust.TextMatrix(CURR_ROW, 9) = Format(TXTADJ_AMT, "0.00")
        CAL_ADJ_TOT
        TXTADJ_AMT.Visible = False
    End If
End Sub
Private Sub CAL_ADJ_TOT()
Dim I As Integer
    AMT_ADJ = 0
    For I = 1 To FgridAdjust.Rows - 1
        AMT_ADJ = AMT_ADJ + Val(FgridAdjust.TextMatrix(I, 9))
    Next
    ADJ_LAB7.CAPTION = Format((Val(FGridLedger.TextMatrix(FGridLedger.Row, 11)) - AMT_ADJ), "0.00")
End Sub



'''''
'FgridAdjust
'0 |1 V_Type |2 V_NO |3 v_Prefix |4 V_SNo |5 V_Date |6 PartyName | 7 AMOUNT  |8 | 9 |10 |11 Narration
