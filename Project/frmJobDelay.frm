VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmJobDelay 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Reason For Job Delay Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.Frame FrJob 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   2820
      TabIndex        =   5
      Top             =   1410
      Visible         =   0   'False
      Width           =   5220
      Begin MSDataGridLib.DataGrid DGJob 
         Height          =   3225
         Left            =   30
         TabIndex        =   6
         Top             =   345
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483648
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Code"
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
         BeginProperty Column01 
            DataField       =   "name"
            Caption         =   "Reason"
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
         BeginProperty Column02 
            DataField       =   "Code"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3089.764
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Reasons"
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
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   5175
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   225
      Index           =   1
      Left            =   2805
      MaxLength       =   30
      TabIndex        =   2
      Top             =   945
      Width           =   3765
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   225
      Index           =   0
      Left            =   2805
      MaxLength       =   2
      TabIndex        =   1
      Top             =   690
      Width           =   900
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason*"
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
      Left            =   1185
      TabIndex        =   4
      Top             =   975
      Width           =   735
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code*"
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
      Index           =   4
      Left            =   1185
      TabIndex        =   3
      Top             =   720
      Width           =   555
   End
End
Attribute VB_Name = "frmJobDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Dim mFlag As Byte
Private Const Code = 0, Desc = 1


Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_Load()
Me.top = 0: Me.left = 0
TopCtrl1.Tag = PubUParam ': TopCtrl1.TopText1 = "Reason For Job Delay Master"   ': TopCtrl1.TopText1.Width = 1000
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select * From JOB_DELAY  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by Code", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select Code as SearchCode, Job_Delay.* From JOB_DELAY Order by Code", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 Code as SearchCode, Job_Delay.* From JOB_DELAY Order by Code", GCn, adOpenDynamic, adLockOptimistic
End If

Set RstHelp = New ADODB.Recordset
'RstHelp.Open "Select CODE,R_DESC as name FROM JOB_DELAY  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by code", GCn, adOpenDynamic, adLockOptimistic
RstHelp.Open "Select CODE,R_DESC as name FROM JOB_DELAY Order by code", GCn, adOpenDynamic, adLockOptimistic
Set DGJob.DataSource = RstHelp
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
CtrlClckCol
MoveRec
ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
txt(Code).Tag = txt(Code)
Txt_GotFocus Code
ADDFLAG = 1
txt(Code).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(Code).Enabled = False
    txt(Desc).Tag = txt(Desc)
    Txt_GotFocus Desc
    ADDFLAG = 2
    txt(Desc).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo Errloop
Dim transFalg As Byte
transFalg = 0
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            XBM = RstMain.Bookmark
            transFalg = 1
            GCn.Execute ("delete * from JOB_DELAY Where Code= " & Chk_Text(Trim(txt(Code))))
            GCn.CommitTrans
            transFalg = 0
            RstMain.Requery
            RstHelp.Requery
            If RstMain.RecordCount >= XBM Then
                RstMain.Bookmark = XBM
            Else
                If RstMain.EOF = False Then RstMain.MoveLast
            End If
            Call MoveRec
        End If
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If

Exit Sub
Errloop:    If transFalg = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbExclamation, " Deletion Error "
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, RstMain, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, RstMain, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, RstMain, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, RstMain, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select Code as SearchCode,Code, R_Desc as Reason From JOB_DELAY Order by Code"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select Code as SearchCode, Job_Delay.* From JOB_DELAY Where Code  = '" & MyValue & "' Order by Code")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "DelayReason"
    mQRY = "SELECT * from Job_Delay"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
'        For i = 1 To rpt.FormulaFields.Count
'            Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
'                Case UCase("SubTitle")
'                    rpt.FormulaFields(i).Text = "'" & rst1!S_SecSpeciality & "'"
'                Case UCase("LST")
'                    rpt.FormulaFields(i).Text = "'" & rst1!S_SecLST & "'"
'                Case UCase("LSTDate")
'                    rpt.FormulaFields(i).Text = "'" & rst1!S_SecLST_Date & "'"
'                Case UCase("CST")
'                    rpt.FormulaFields(i).Text = "'" & rst1!S_SecCST & "'"
'                Case UCase("CSTDate")
'                    rpt.FormulaFields(i).Text = "'" & rst1!S_SecCST_Date & "'"
'                Case UCase("Phone")
'                    rpt.FormulaFields(i).Text = "'" & rst1!S_SecPhone & "'"
'                Case UCase("Fax")
'                    rpt.FormulaFields(i).Text = "'" & rst1!S_SecFax & "'"
'            End Select
'        Next
     rpt.Database.SetDataSource Rst
     rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , True)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo Errloop
    transFlag = 0
    If IsValid(txt(Code), "Code") = False Then Txt_GotFocus Code: Exit Sub
    If IsValid(txt(Desc), "Description") = False Then Txt_GotFocus Desc: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From JOB_DELAY Where Code= " & Chk_Text(Trim(txt(Code))) & " AND SITE_CODE='" & PubSiteCode & "'").Fields(0) > 0 Then MsgBox "Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus Code: txt(Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        GCn.Execute ("DELETE From JOB_DELAY Where Code= " & Chk_Text(Trim(txt(Code))) & " AND SITE_CODE='" & PubSiteCode & "'")
        GCn.Execute ("Insert Into JOB_DELAY(Code,Site_Code,R_DESC,U_Name,U_EntDt,U_AE) Values('" & txt(Code) & "','" & PubSiteCode & "','" & txt(Desc) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    ElseIf ADDFLAG = 2 Then
        GCn.Execute ("UPDATE JOB_DELAY SET Code=" & Chk_Text(Trim(txt(Code))) & ",Site_Code='" & PubSiteCode & "',R_DESC=" & Chk_Text(txt(Desc)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "' Where Code= " & Chk_Text(Trim(txt(Code))) & " AND SITE_CODE='" & PubSiteCode & "'")
    End If

    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select Code as SearchCode, Job_Delay.* From JOB_DELAY Where Code  = " & Chk_Text(Trim(txt(Code))) & " Order by Code")
    End If
    RstHelp.Requery
    RstMain.FIND ("Code=" & Chk_Text(Trim(txt(Code))))
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus Code
        txt(Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrJob.Visible = False
    End If
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
        CtrlClckCol
        FrJob.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(Code).BackColor = CtrlBColOrg:      txt(Code).ForeColor = CtrlFColOrg
    txt(Desc).BackColor = CtrlBColOrg:      txt(Desc).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
On Error GoTo Errloop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    txt(Code) = XNull(RstMain!Code)
    txt(Desc) = XNull(RstMain!R_DESC)
End If
TopCtrl1.tDel = False
Exit Sub
Errloop:        MsgBox err.Description
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub

Private Sub ColCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Code >=" & Chk_Text(Trim(txt(Code)))
End Sub
Private Sub ColNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "name >=" & Chk_Text(XNull(txt(Desc)))
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Code, Desc
            If RstHelp.RecordCount = 0 Then Exit Sub
            If FrJob.Visible = True Then FrJob.Visible = False
            FrJob.Visible = True
            FrJob.top = txt(Index).top + txt(Index).height + 10
            FrJob.left = txt(Index).left
            FrJob.ZOrder 0
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
DGJob.Columns(0).width = 1000.1: DGJob.Columns(1).width = 3535.024: DGJob.Columns(2).width = 1000.1
Dim mBookMark
    Ctrl_GetFocus txt(Index)
mFlag = 0
    If FrJob.Visible = True Then FrJob.Visible = False
    RST_BOF_EOF RstHelp
    txt(Index).Tag = txt(Index)
    Select Case Index
        Case Code, Desc
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case Code
            DGJob.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Code ASC"
            RstHelp.Bookmark = mBookMark
            ColCodeSearch
        Case Desc
           DGJob.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "NAME ASC"
            RstHelp.Bookmark = mBookMark
            ColNameSearch
    End Select
    If txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
Select Case Index
    Case Desc
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            FrJob.Visible = False
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt_GotFocus Index: txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
'                Txt_Click Index
                Txt_GotFocus Index
                txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If

End Select
Select Case Index
    Case Code
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case Code
        ColCodeSearch
    Case Desc
        ColNameSearch
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Code
            Set Rst = GCn.Execute("SELECT * FROM JOB_DELAY WHERE Code=" & Chk_Text(Trim(txt(Code))))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox " Code Already Exists", vbInformation, "Validation": txt(Code) = txt(Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Code <> RstMain!Code Then MsgBox "Code Already Exists", vbInformation, "Validation": txt(Code) = txt(Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Desc
            Set Rst = GCn.Execute("SELECT * FROM JOB_DELAY WHERE R_DESC=" & Chk_Text(txt(Desc)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Description Already Exists", vbInformation, "Validation": txt(Desc) = txt(Desc).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!R_DESC <> RstMain!R_DESC Then MsgBox "Description Already Exists", vbInformation, "Validation": txt(Desc) = txt(Desc).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
'CmbOrder.Enabled = IIf(AddFlag = 1, True, False)
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
Next
End Sub

'Private Sub Ini_Grid()
'    FGrid.RowHeightMin = 250
'    FGrid.ColWidth(25) = 0
'End Sub

