VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FaCurrBalUpdate 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Current Balance Updation"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7335
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   720
      Left            =   615
      TabIndex        =   3
      Top             =   1395
      Width           =   6105
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   -15
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Processing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1222"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1650
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1233"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4650
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808080&
      Caption         =   "Accounts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   345
      Index           =   1
      Left            =   615
      TabIndex        =   2
      Top             =   525
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   315
      Left            =   5370
      TabIndex        =   1
      Top             =   2970
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   315
      Left            =   4065
      TabIndex        =   0
      Top             =   2970
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo TXT_ACCOUNT 
      DataField       =   "PARTY"
      DataSource      =   "master"
      Height          =   315
      Left            =   2145
      TabIndex        =   8
      Top             =   540
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   12640511
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   390
      Left            =   615
      TabIndex        =   9
      Top             =   2385
      Visible         =   0   'False
      Width           =   6270
   End
End
Attribute VB_Name = "FaCurrBalUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click(Index As Integer)
    If Check1(1).Value = vbChecked Then
           TXT_ACCOUNT.Visible = True
           TXT_ACCOUNT.SetFocus
    ElseIf Check1(1).Value = vbUnchecked Then
           TXT_ACCOUNT.Visible = False
    End If
End Sub
Private Sub Command1_Click()
Dim I As Integer
Dim GRs As New ADODB.Recordset
On Error GoTo ErrorLoop
    Label4(0).Visible = True
    Label4(1).Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label1.Visible = True
    Label5.CAPTION = ""
    Label6.CAPTION = ""
    'Account Balance Updation
    If Check1(1).Value = vbChecked Then
        Set GRs = G_FaCn.Execute("SELECT SubGroup.SubCode, SubGroup.Name,Sum(Ledger.AmtCr)-Sum(Ledger.AmtDr) AS Balance FROM SubGroup LEFT JOIN Ledger ON SubGroup.SubCode = Ledger.SubCode where SubGroup.SubCode ='" & TXT_ACCOUNT.BoundText & "' GROUP BY SubGroup.SubCode, SubGroup.Name")
    Else
        Set GRs = G_FaCn.Execute("SELECT SubGroup.SubCode, SubGroup.Name,Sum(Ledger.AmtCr)-Sum(Ledger.AmtDr) AS Balance FROM SubGroup LEFT JOIN Ledger ON SubGroup.SubCode = Ledger.SubCode GROUP BY SubGroup.SubCode, SubGroup.Name")
    End If
    
        Label6.CAPTION = ""
        Label5.CAPTION = ""
        Label5.CAPTION = CStr(GRs.RecordCount)
        Label5.Refresh
        Do Until GRs.EOF
            Label6.CAPTION = GRs.AbsolutePosition
            Label6.Refresh
            Label1.CAPTION = GRs!Name: Label1.Refresh
            G_FaCn.BeginTrans
                If Not IsNull(GRs!Balance) Then
                    G_FaCn.Execute ("Update SubGroup Set Curr_Bal=" & GRs!Balance & " Where SubCode='" & GRs!SubCode & "'")
                Else
                    G_FaCn.Execute ("Update SubGroup Set Curr_Bal=0 Where SubCode='" & GRs!SubCode & "'")
                End If
            G_FaCn.CommitTrans
            GRs.MoveNext
        Loop
        MsgBox "Opening Balance Updation Has been Completed", vbInformation, "Information"
    Label4(0).Visible = False
    Label4(1).Visible = False
    Label1.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Set GRs = Nothing
    Exit Sub
ErrorLoop:      G_FaCn.RollbackTrans
                MsgBox err.Description, vbCritical, "Posting Error": Exit Sub
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 Then Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ErrorLoop
    Me.Icon = MDIForm1.Icon
    Label4(0).Visible = False
    Label4(1).Visible = False
    Label5.CAPTION = ""
    Label6.CAPTION = ""
    Label5.Visible = False
    Label6.Visible = False
    FaIniCombo "SELECT NAME,SUBCODE FROM SUBGROUP ORDER BY NAME", TXT_ACCOUNT, "NAME", "SUBCODE"
    Exit Sub
ErrorLoop:      MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
