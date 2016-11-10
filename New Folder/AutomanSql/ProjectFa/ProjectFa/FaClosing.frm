VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FaClosing 
   BackColor       =   &H00CDCCFB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Environment Setting"
   ClientHeight    =   6510
   ClientLeft      =   180
   ClientTop       =   795
   ClientWidth     =   10455
   Icon            =   "FaClosing.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10455
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Status 
      BackColor       =   &H00CDCCFB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   7
      Top             =   1500
      Width           =   7965
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete"
      Top             =   2670
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
      Left            =   5685
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   2670
      Width           =   1365
   End
   Begin VB.CommandButton BtnOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ok"
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
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   2670
      Width           =   1365
   End
   Begin VB.TextBox TXT_DATE 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8EBFC&
      Height          =   300
      Left            =   2985
      MaxLength       =   12
      TabIndex        =   2
      Top             =   165
      Width           =   1410
   End
   Begin MSDataListLib.DataCombo DBPlac 
      Height          =   315
      Left            =   2985
      TabIndex        =   0
      Top             =   480
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16313340
      Text            =   ""
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Date"
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
      Left            =   1545
      TabIndex        =   3
      Top             =   195
      Width           =   1350
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "P/L Account"
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
      Left            =   1545
      TabIndex        =   1
      Top             =   517
      Width           =   1245
   End
End
Attribute VB_Name = "FaClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BTNDELETE_Click()
Dim RST1 As ADODB.Recordset, mSanSite As String
On Error GoTo ELoop
Select Case PubFaSiteType
    Case 0, 1
        mSanSite = PubSiteCode + PubSiteCode
    Case 2
        mSanSite = PubSiteCode
End Select
Set RST1 = G_FaCn.Execute("SELECT SUBCODE,AMTDR,AMTCR FROM LEDGER WHERE V_tYPE='PLCLS' AND V_DATE=" & FaConvertDate(PubEndDate) & " AND SITE_cODE='" & mSanSite & "'")
Do Until RST1.EOF
    If RST1!AmtCr > 0 Then
        FaCalCurrBal G_FaCn, RST1!SubCode, RST1!AmtCr, 0
    ElseIf RST1!AmtDr > 0 Then
        FaCalCurrBal G_FaCn, RST1!SubCode, 0, RST1!AmtDr
    End If
    RST1.MoveNext
Loop
G_FaCn.Execute "DELETE FROM LEDGER WHERE V_tYPE='PLCLS' AND V_DATE=" & FaConvertDate(PubEndDate) & " AND SITE_CODE='" & mSanSite & "'"
Set RST1 = Nothing
Exit Sub
ELoop:
        MsgBox err.Description
End Sub
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
Private Sub btncancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Me.left = 0
    Me.top = 0
    TXT_DATE = PubEndDate
    FaIniCombo "SELECT SUBCODE,NAME FROM VIEWSUBGROUP WHERE MAINGRCODES='999' AND SITE_CODE='" & PubSiteCode & "'", DBPlac, "Name", "Subcode"
End Sub
Private Sub TXT_DATE_Validate(Cancel As Boolean)
    TXT_DATE = RetDate(TXT_DATE)
End Sub
Private Sub btnok_Click()
Dim RST1 As ADODB.Recordset, Rst2 As ADODB.Recordset, I As Integer, mSanSite As String, mDocId As String
On Error GoTo ELoop
If FaIsValid(TXT_DATE, "Closing Date") = False Then Exit Sub
If FaIsValid(DBPlac, "P/L Account") = False Then Exit Sub
Select Case PubFaSiteType
    Case 0, 1
        mSanSite = PubSiteCode + PubSiteCode
    Case 2
        mSanSite = PubSiteCode
End Select
Set RST1 = G_FaCn.Execute("SELECT SUBCODE,AMTDR,AMTCR FROM LEDGER WHERE V_tYPE='PLCLS' AND V_DATE=" & FaConvertDate(PubEndDate) & " AND SITE_CODE='" & mSanSite & "'")
Do Until RST1.EOF
    If RST1!AmtCr > 0 Then
        FaCalCurrBal G_FaCn, RST1!SubCode, RST1!AmtCr, 0
    ElseIf RST1!AmtDr > 0 Then
        FaCalCurrBal G_FaCn, RST1!SubCode, 0, RST1!AmtDr
    End If
    RST1.MoveNext
Loop
G_FaCn.Execute "DELETE FROM LEDGER WHERE V_tYPE='PLCLS' AND V_DATE=" & FaConvertDate(PubEndDate) & " AND SITE_cODE='" & mSanSite & "'"

Set RST1 = G_FaCn.Execute("SELECT SUBGROUP.*,MAINGRCODE FROM SUBGROUP LEFT JOIN ACGROUP ON SUBGROUP.GROUPCODE=ACGROUP.GROUPCODE WHERE ACGROUP.MAINGRCODE<>'999' AND ((SUBGROUP.GROUPNATURE IN ('E','R')  AND ACGROUP.GROUPNATURE IN ('E','R')) OR ACGROUP.MAINGRCODE='060001') AND SUBGROUP.ALIASYN<>'Y' AND ACGROUP.ALIASYN<>'Y' AND SUBGROUP.SITE_CODE='" & PubSiteCode & "'  ORDER BY SUBGROUP.NAME")
I = 0
Do Until RST1.EOF
    Status = FaXNull(RST1!Name)
    Status.Refresh
    If RST1!MainGrCode = "060001" Then
        Set Rst2 = G_FaCn.Execute("SELECT SUM(AMTCR-AMTDR) AS BAL FROM LEDGER WHERE V_DATE<=" & FaConvertDate(PubStartDate) & " AND SUBCODE='" & RST1!SubCode & "' AND SITE_cODE='" & mSanSite & "'")
    Else
        Set Rst2 = G_FaCn.Execute("SELECT SUM(AMTCR-AMTDR) AS BAL FROM LEDGER WHERE V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubEndDate) & " AND SUBCODE='" & RST1!SubCode & "' AND SITE_cODE='" & mSanSite & "'")
    End If
    If Rst2.RecordCount > 0 Then
        If FaVNull(Rst2!BAL) > 0 Then
            I = I + 1
            Select Case PubFaSiteType
                Case 0, 1
                    mSanSite = PubSiteCode + PubSiteCode
                Case 2
                    mSanSite = PubSiteCode
            End Select
            mDocId = PubDivCode + mSanSite + "PLCLS" + FaSetW(CStr(Year(TXT_DATE)), 5) + FaSetN(CStr(I), 8)
            G_FaCn.Execute "INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "',1,'PLCLS'," & I & ",'" & CStr(Year(TXT_DATE)) & "','" & mSanSite & "'," & FaConvertDate(TXT_DATE) & ",'" & RST1!SubCode & "',0," & FaVNull(Rst2!BAL) & ",'" & DBPlac.BoundText & "','" & "Trans. To Balance Sheet" & "','" & pubUName & "'," & FaConvertDate(Now) & ",'A')"
            G_FaCn.Execute "INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "',2,'PLCLS'," & I & ",'" & CStr(Year(TXT_DATE)) & "','" & mSanSite & "'," & FaConvertDate(TXT_DATE) & ",'" & DBPlac.BoundText & "'," & FaVNull(Rst2!BAL) & ",0,'" & RST1!SubCode & "','" & "Trans. To Balance Sheet" & "','" & pubUName & "'," & FaConvertDate(Now) & ",'A')"
            FaCalCurrBal G_FaCn, RST1!SubCode, 0, FaVNull(Rst2!BAL)
            FaCalCurrBal G_FaCn, DBPlac.BoundText, FaVNull(Rst2!BAL), 0
        ElseIf FaVNull(Rst2!BAL) < 0 Then
            I = I + 1
            Select Case PubFaSiteType
                Case 0, 1
                    mSanSite = PubSiteCode + PubSiteCode
                Case 2
                    mSanSite = PubSiteCode
            End Select
            mDocId = PubDivCode + mSanSite + "PLCLS" + FaSetW(CStr(Year(TXT_DATE)), 5) + FaSetN(CStr(I), 8)
            G_FaCn.Execute "INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "',1,'PLCLS'," & I & ",'" & CStr(Year(TXT_DATE)) & "','" & mSanSite & "'," & FaConvertDate(TXT_DATE) & ",'" & RST1!SubCode & "'," & Abs(FaVNull(Rst2!BAL)) & ",0,'" & DBPlac.BoundText & "','" & "Trans. To Profit & Loss A/C" & "','" & pubUName & "'," & FaConvertDate(Now) & ",'A')"
            G_FaCn.Execute "INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "',2,'PLCLS'," & I & ",'" & CStr(Year(TXT_DATE)) & "','" & mSanSite & "'," & FaConvertDate(TXT_DATE) & ",'" & DBPlac.BoundText & "',0," & Abs(FaVNull(Rst2!BAL)) & ",'" & RST1!SubCode & "','" & "Trans. To Profit & Loss A/C" & "','" & pubUName & "'," & FaConvertDate(Now) & ",'A')"
            FaCalCurrBal G_FaCn, RST1!SubCode, Abs(FaVNull(Rst2!BAL)), 0
            FaCalCurrBal G_FaCn, DBPlac.BoundText, 0, Abs(FaVNull(Rst2!BAL))
        End If
    End If
    RST1.MoveNext
Loop
Status = "Compeleted"
Status.Refresh
Set RST1 = Nothing
Set Rst2 = Nothing
Exit Sub
ELoop:
        MsgBox err.Description
End Sub
