VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FaAdjustDel 
   BackColor       =   &H00E6AC86&
   Caption         =   "Adjustment Delete"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FaAdjustDel.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton BTNFILL 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fill Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   195
      Width           =   1365
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9870
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   195
      Width           =   1365
   End
   Begin VB.CommandButton btndelete 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Delete"
      Top             =   195
      Width           =   1365
   End
   Begin MSDataListLib.DataCombo TXT_ACCOUNT 
      DataField       =   "PARTY"
      DataSource      =   "master"
      Height          =   315
      Left            =   2370
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16249055
      Text            =   "DataCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid FG1 
      Height          =   6840
      Left            =   15
      TabIndex        =   4
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12065
      _Version        =   393216
      Cols            =   12
      BackColor       =   14873589
      ForeColor       =   64
      BackColorSel    =   8388608
      ForeColorSel    =   65535
      BackColorBkg    =   15117446
      GridColor       =   0
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"FaAdjustDel.frx":030A
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   525
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   135
      Width           =   4215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   270
      Width           =   1545
   End
End
Attribute VB_Name = "FaAdjustDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BTNDELETE_Click()
On Error GoTo ELoop
If FG1.Rows > 1 Then
    If FG1.TextMatrix(1, 2) <> "" Then
        If MsgBox("Do You Want to Delete It ", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Confirmation") = 6 Then
            G_FaCn.BeginTrans
            G_FaCn.Execute "Delete From LEDGERADJ Where DocId1='" & FG1.TextMatrix(FG1.Row, 10) & "' AND V_SNo1=" & FG1.TextMatrix(FG1.Row, 4) & " AND DocId2='" & FG1.TextMatrix(FG1.Row, 11) & "' AND V_SNo2=" & FG1.TextMatrix(FG1.Row, 8) & " AND SubCode='" & TXT_ACCOUNT.BoundText & "'"
            G_FaCn.CommitTrans
        End If
    End If
End If
BTNFILL_Click
Exit Sub
ELoop:  G_FaCn.RollbackTrans
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_Load()
    FaIniCombo "SELECT NAME,SUBCODE FROM SUBGROUP ORDER BY NAME", TXT_ACCOUNT, "NAME", "SUBCODE"
    FaWinSetting Me
    btndelete.Enabled = False
    TXT_ACCOUNT.Enabled = True
    FG1.ColAlignment(2) = flexAlignLeftCenter
    FG1.ColAlignment(4) = flexAlignLeftCenter
    FG1.ColAlignment(6) = flexAlignLeftCenter
    FG1.ColAlignment(8) = flexAlignLeftCenter
    FG1.ColAlignment(9) = flexAlignRightCenter
    FG1.ColWidth(3) = 0
    FG1.ColWidth(7) = 0
    FG1.ColWidth(10) = 0
    FG1.ColWidth(11) = 0
    FG1.left = 15: FG1.top = 720
End Sub
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub BTNFILL_Click()
Dim G_Rs As ADODB.Recordset
If TXT_ACCOUNT <> "" Then
    FG1.Rows = 1
    Set G_Rs = G_FaCn.Execute("SELECT * FROM LEDGERADJ WHERE SUBCODE='" & TXT_ACCOUNT.BoundText & "'")
    Do Until G_Rs.EOF
        FG1.AddItem "" & Chr(9) & Mid(G_Rs!DocID1, 4, 5) & Chr(9) & Trim(Right(G_Rs!DocID1, 8)) & Chr(9) & "" & Chr(9) & G_Rs!V_SNo1 & Chr(9) & Mid(G_Rs!DocID2, 4, 5) & Chr(9) & Trim(Right(G_Rs!DocID2, 8)) & Chr(9) & "" & Chr(9) & G_Rs!V_SNo2 & Chr(9) & Format(G_Rs!cr, "0.00") & Chr(9) & G_Rs!DocID1 & Chr(9) & G_Rs!DocID2
        G_Rs.MoveNext
    Loop
    If FG1.Rows = 1 Then
        MsgBox "No Record Found to Fill", vbInformation, Me.CAPTION
        FG1.AddItem ""
    Else
        btndelete.Enabled = True
    End If
End If
Set G_Rs = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
