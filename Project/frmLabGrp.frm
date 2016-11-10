VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmLabGrp 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Labour Group Master"
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
   Begin VB.Frame FrLabG 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   1170
      TabIndex        =   5
      Top             =   2700
      Visible         =   0   'False
      Width           =   5220
      Begin MSDataGridLib.DataGrid DGLabG 
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
         BackColor       =   16777215
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Lab_Group"
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
            DataField       =   "LabGrp_Desc"
            Caption         =   "Labour Description"
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
            DataField       =   "Lab_Group"
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
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3089.764
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Labour Groups"
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
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1020
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
      MaxLength       =   1
      TabIndex        =   1
      Top             =   765
      Width           =   900
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Description*"
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
      Left            =   735
      TabIndex        =   4
      Top             =   1050
      Width           =   1710
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Group Code*"
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
      Left            =   735
      TabIndex        =   3
      Top             =   795
      Width           =   1785
   End
End
Attribute VB_Name = "frmLabGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Dim mFlag As Byte
Private Const Lab_Group = 0, LabGrp_Desc = 1

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
Me.top = 0: Me.left = 0
TopCtrl1.Tag = PubUParam ': TopCtrl1.TopText1 = "Labour Group Master"   ': TopCtrl1.TopText1.Width = 1000
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select Lab_Group as SearchCode, Labour_Group.* From Labour_Group  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by Lab_Group", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select Lab_Group as SearchCode, Labour_Group.* From Labour_Group Order by Lab_Group", GCn, adOpenDynamic, adLockOptimistic
Else
    Set RstMain = GCn.Execute("Select Top 1 Lab_Group as SearchCode, Labour_Group.* From Labour_Group Order by Lab_Group")
End If
Set RstHelp = New ADODB.Recordset
'RstHelp.Open "Select Lab_Group,LabGrp_Desc FROM Labour_Group  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by Lab_Group", GCn, adOpenDynamic, adLockOptimistic
RstHelp.Open "Select Lab_Group,LabGrp_Desc FROM Labour_Group Order by Lab_Group", GCn, adOpenDynamic, adLockOptimistic
CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0:    mFlag = 0
Set DGLabG.DataSource = RstHelp
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
txt(Lab_Group).Tag = txt(Lab_Group)
Txt_GotFocus Lab_Group
ADDFLAG = 1
txt(Lab_Group).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(Lab_Group).Enabled = False
    txt(LabGrp_Desc).Tag = txt(LabGrp_Desc)
    Txt_GotFocus LabGrp_Desc
    ADDFLAG = 2
    txt(LabGrp_Desc).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo Errloop
Dim mTrans As Byte
mTrans = 0
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        If GCn.Execute("Select Lab_Group from Labour where Lab_Group='" & txt(Lab_Group) & "'").RecordCount > 0 Then
            MsgBox "Transaction in Labour Description exists" & vbCrLf & "Delete Denied !", vbCritical, "Delete Denied!"
            Exit Sub
        End If
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            XBM = RstMain.Bookmark
            mTrans = 1
            GCn.Execute ("delete  from Labour_Group where Lab_Group= " & Chk_Text(Trim(txt(Lab_Group))))
            GCn.CommitTrans
            mTrans = 0
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
Errloop:    If mTrans = 1 Then GCn.RollbackTrans
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
    GSQL = "SELECT Lab_Group as SearchCode,Lab_Group,LabGrp_Desc FROM labour_group Order By Lab_Group"
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
        Set RstMain = GCn.Execute("Select Lab_Group as SearchCode, Labour_Group.* From Labour_Group Where Lab_Group ='" & MyValue & "' Order by Lab_Group")
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

    mRepName = "LabourGrp"
    mQRY = "SELECT * from labour_group Order By Lab_Group"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
     rpt.Database.SetDataSource Rst
     rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , False)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION

End Sub

Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo Errloop
    transFlag = 0
    If IsValid(txt(Lab_Group), "Labour Group Code") = False Then Txt_GotFocus Lab_Group: Exit Sub
    If IsValid(txt(LabGrp_Desc), "Labour Description") = False Then Txt_GotFocus LabGrp_Desc: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From Labour_Group Where Lab_Group= " & Chk_Text(Trim(txt(Lab_Group)))).Fields(0) > 0 Then MsgBox "Labour Group Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus Lab_Group: txt(Lab_Group).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        GCn.Execute ("Insert Into Labour_Group(Lab_Group,Site_Code,LabGrp_Desc,U_Name,U_EntDt,U_AE) Values(" & Chk_Text(txt(Lab_Group)) & ",'" & PubSiteCode & "'," & Chk_Text(txt(LabGrp_Desc)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    ElseIf ADDFLAG = 2 Then
        GCn.Execute ("UPDATE Labour_Group SET Site_Code='" & PubSiteCode & "',LabGrp_Desc='" & txt(LabGrp_Desc) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "' Where Lab_Group= '" & txt(Lab_Group) & "'")
    End If

    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select Lab_Group as SearchCode, Labour_Group.* From Labour_Group Where Lab_Group =" & Chk_Text(txt(Lab_Group)) & " Order by Lab_Group")
    End If
    RstHelp.Requery
    RstMain.FIND ("Lab_Group=" & Chk_Text(txt(Lab_Group)))
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus Lab_Group
        txt(Lab_Group).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrLabG.Visible = False
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
        FrLabG.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(Lab_Group).BackColor = CtrlBColOrg:      txt(Lab_Group).ForeColor = CtrlFColOrg
    txt(LabGrp_Desc).BackColor = CtrlBColOrg:      txt(LabGrp_Desc).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
On Error GoTo Errloop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    txt(Lab_Group) = XNull(RstMain!Lab_Group)
    txt(LabGrp_Desc) = XNull(RstMain!LabGrp_Desc)
End If
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

Private Sub LabGCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Lab_Group >=" & Chk_Text(Trim(txt(Lab_Group)))
End Sub
Private Sub LabGNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "LabGrp_Desc >=" & Chk_Text(XNull(txt(LabGrp_Desc)))
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Lab_Group, LabGrp_Desc
            If RstHelp.RecordCount = 0 Then Exit Sub
            If FrLabG.Visible = True Then FrLabG.Visible = False
            FrLabG.Visible = True
            FrLabG.top = txt(Index).top + txt(Index).height + 10
            FrLabG.left = txt(Index).left
            FrLabG.ZOrder 0
    End Select
End If
End Sub

Private Sub Txt_GotFocus(Index As Integer)
DGLabG.Columns(0).width = 1000.1: DGLabG.Columns(1).width = 3535.024: DGLabG.Columns(2).width = 1000.1
Dim mBookMark
    Ctrl_GetFocus txt(Index)
mFlag = 0
    If FrLabG.Visible = True Then FrLabG.Visible = False
    RST_BOF_EOF RstHelp
    txt(Index).Tag = txt(Index)
    Select Case Index
        Case Lab_Group, LabGrp_Desc
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case Lab_Group
            DGLabG.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Lab_Group ASC"
            RstHelp.Bookmark = mBookMark
            LabGCodeSearch
        Case LabGrp_Desc
            DGLabG.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "LabGrp_Desc ASC"
            RstHelp.Bookmark = mBookMark
            LabGNameSearch
    End Select
    If txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
Select Case Index
    Case LabGrp_Desc
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            FrLabG.Visible = False
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
    Case Lab_Group
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)

End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case Lab_Group
        LabGCodeSearch
    Case LabGrp_Desc
        LabGNameSearch
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Lab_Group
            Set Rst = GCn.Execute("SELECT * FROM Labour_Group WHERE Lab_Group=" & Chk_Text(txt(Lab_Group)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Labour Group Code Already Exists", vbInformation, "Validation": txt(Lab_Group) = txt(Lab_Group).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Lab_Group <> RstMain!Lab_Group Then MsgBox "Labour Group Code Already Exists", vbInformation, "Validation": txt(Lab_Group) = txt(Lab_Group).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case LabGrp_Desc
            Set Rst = GCn.Execute("SELECT * FROM Labour_Group WHERE LabGrp_Desc=" & Chk_Text(txt(LabGrp_Desc)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Labour Description Already Exists", vbInformation, "Validation": txt(LabGrp_Desc) = txt(LabGrp_Desc).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!LabGrp_Desc <> RstMain!LabGrp_Desc Then MsgBox "Labour Description Already Exists", vbInformation, "Validation": txt(LabGrp_Desc) = txt(LabGrp_Desc).Tag: Cancel = True: Exit Sub
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

