VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmArea 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Area Master"
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
   Begin VB.Frame FrArea 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   2865
      TabIndex        =   6
      Top             =   1845
      Visible         =   0   'False
      Width           =   4950
      Begin MSDataGridLib.DataGrid DGArea 
         Height          =   3225
         Left            =   30
         TabIndex        =   3
         Top             =   345
         Width           =   4890
         _ExtentX        =   8625
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
            Caption         =   "Area"
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
         Caption         =   "List of Area"
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
         Width           =   4890
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2865
      MaxLength       =   15
      TabIndex        =   2
      Top             =   930
      Width           =   3765
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2865
      MaxLength       =   2
      TabIndex        =   1
      Top             =   660
      Width           =   900
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   1605
      TabIndex        =   5
      Top             =   945
      Width           =   405
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   4
      Left            =   1605
      TabIndex        =   4
      Top             =   675
      Width           =   435
   End
End
Attribute VB_Name = "frmArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim mSearchCode As String
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset

Private Const Code = 0
Private Const Desc = 1

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
    CheckError
End Sub

Private Sub Form_Load()
    WinSetting Me, 5325, 9630
    TopCtrl1.Tag = PubUParam
    Set RstMain = New ADODB.Recordset
    If PubMoveRecYn Then
        RstMain.Open "Select AreaCode As SearchCode,Area.* From Area Where Site_Code='" & PubSiteCode & "' Order by AreaName", GCn, adOpenDynamic, adLockOptimistic
    Else
        RstMain.Open "Select Top 1 AreaCode As SearchCode,Area.* From Area Where Site_Code='" & PubSiteCode & "' Order by AreaName", GCn, adOpenDynamic, adLockOptimistic
    End If
    Set RstHelp = New ADODB.Recordset
    RstHelp.Open "Select " & cMID("AreaCode", "2", "2") & " As Code,AreaName as Name From Area Where Site_Code='" & PubSiteCode & "' Order by AreaCode", GCn, adOpenDynamic, adLockOptimistic
    Set DGArea.DataSource = RstHelp
    CtrlClckCol
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, RstMain)
    MoveRec
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, RstMain)
    Txt(Code).Tag = Txt(Code)
    Txt_GotFocus Code
    Txt(Code).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If RstMain.RecordCount > 0 Then
        Disp_Text SETS("EDIT", Me, RstMain)
        Txt(Code).Enabled = False
        Txt(Desc).Tag = Txt(Desc)
        Txt_GotFocus Desc
        Txt(Desc).SetFocus
    Else
        MsgBox "There Is No Record To Edit.", vbInformation, "Information"
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim mTrans As Byte
Dim XBM
Dim Res As Integer
mTrans = 0
    If RstMain.RecordCount > 0 Then
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
                XBM = RstMain.Bookmark
                mTrans = 1
                GCn.Execute ("Delete * From Area Where AreaCode='" & PubSiteCode + Trim(Txt(Code)) & "'")
            GCn.CommitTrans
            mTrans = 0
            RstMain.Requery
            RstHelp.Requery
            If RstMain.RecordCount >= XBM Then
                RstMain.Bookmark = XBM
            Else
                If RstMain.EOF = False Then RstMain.MoveLast
            End If
            MoveRec
        End If
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If
Exit Sub
ELoop:
    If mTrans = 1 Then GCn.RollbackTrans
    CheckError
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
On Error GoTo ELoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select AreaCode As SearchCode, " & cMID("AreaCode", "2", "2") & " As Code,AreaName as Name From Area Where Site_Code='" & PubSiteCode & "' Order by AreaName"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "Area"
    mQRY = "SELECT * from Area Order By AreaName"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , True)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION

End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Byte
On Error GoTo ELoop
    mTrans = 0
    If IsValid(Txt(Code), "Code") = False Then Txt_GotFocus Code: Exit Sub
    If IsValid(Txt(Desc), "Area Name") = False Then Txt_GotFocus Desc: Exit Sub
    If TopCtrl1.TopText2 = "Add" Then If GCn.Execute("Select Count(*) From Area Where AreaCode='" & PubSiteCode + Trim(Txt(Code)) & "' And Site_Code='" & PubSiteCode & "'").Fields(0) > 0 Then MsgBox "Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus Code: Txt(Code).SetFocus: Exit Sub
    GCn.BeginTrans
    mTrans = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("Insert Into Area(AreaCode,Site_Code,AreaName,U_Name,U_EntDt,U_AE) Values('" & PubSiteCode + Trim(Txt(Code)) & "','" & PubSiteCode & "','" & Txt(Desc) & "','" & pubUName & "'," & ConvertDate(Format(Now, "dd/MMM/yyyy HH:NN:SS")) & ",'A')")
    ElseIf TopCtrl1.TopText2 = "Edit" Then
        GCn.Execute ("Update Area Set AreaName='" & Txt(Desc) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(Format(Now, "dd/MMM/yyyy HH:NN:SS")) & ",U_AE='E' Where AreaCode='" & PubSiteCode + Trim(Txt(Code)) & "'")
    End If

    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    mTrans = 0
    mSearchCode = PubSiteCode + Trim(Txt(Code))
    
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select AreaCode As SearchCode,Area.* From Area Where Site_Code='" & PubSiteCode & "' And AreaCode = '" & mSearchCode & "' Order by AreaName")
    End If
    
    RstHelp.Requery
    If TopCtrl1.TopText2 = "Add" Then
        TopCtrl1_eAdd
        Exit Sub
    End If
    
    RstMain.FIND "AreaCode='" & mSearchCode & "'"
    Disp_Text SETS("INI", Me, RstMain)
    MoveRec
    CtrlClckCol
    FrArea.Visible = False
Exit Sub
ELoop:
    If mTrans = 1 Then GCn.RollbackTrans
    CheckError
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
    If MasterFormExit Then Unload Me: Exit Sub
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        FrArea.Visible = False
    Else
        Me.ActiveControl.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

'**********Functions***********
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select AreaCode As SearchCode,Area.* From Area Where Site_Code='" & PubSiteCode & "' And AreaCode = '" & MyValue & "' Order by AreaName")
    End If
    MoveRec
    BUTTONS True, Me, RstMain, 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub CtrlClckCol()
    Txt(Code).BackColor = CtrlBColOrg: Txt(Code).ForeColor = CtrlFColOrg
    Txt(Desc).BackColor = CtrlBColOrg: Txt(Desc).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
On Error GoTo ELoop
RST_BOF_EOF RstMain
TopCtrl1.tDel = False
    If RstMain.RecordCount <= 0 Then
        BlankText
    Else
        'Txt(Code) = RstMain!AreaCode
        Txt(Code) = mID(RstMain!AreaCode, 2, 2)
        mSearchCode = RstMain!AreaCode
        Txt(Desc) = RstMain!AreaName
    End If
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eRef()
    RstMain.Requery
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub

Private Sub ColCodeSearch()
    If RstHelp.RecordCount <= 0 Then Exit Sub
    RstHelp.MoveFirst
    RstHelp.FIND "Code >='" & Trim(Txt(Code)) & "'"
End Sub
Private Sub ColNameSearch()
    If RstHelp.RecordCount <= 0 Then Exit Sub
    RstHelp.MoveFirst
    RstHelp.FIND "Name >='" & Txt(Desc) & "'"
End Sub

Private Sub Txt_Change(Index As Integer)
If TopCtrl1.TopText2 <> "Browse" Then
    Select Case Index
    Case Code, Desc
        If RstHelp.RecordCount = 0 Then Exit Sub
        If FrArea.Visible = True Then FrArea.Visible = False
        FrArea.Visible = True
        FrArea.top = Txt(Index).top + Txt(Index).height + 10
        FrArea.left = Txt(Index).left
        FrArea.ZOrder 0
    End Select
End If
End Sub

Private Sub Txt_GotFocus(Index As Integer)
DGArea.Columns(0).width = 1000.1: DGArea.Columns(1).width = 3535.024: DGArea.Columns(2).width = 1000.1
Dim mBookMark
    Ctrl_GetFocus Txt(Index)
    If FrArea.Visible = True Then FrArea.Visible = False
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Select Case Index
        Case Code, Desc
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case Code
            DGArea.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Code"
            RstHelp.Bookmark = mBookMark
            ColCodeSearch
        Case Desc
           DGArea.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Name"
            RstHelp.Bookmark = mBookMark
            ColNameSearch
    End Select
    If Txt(Index) = "" Then Txt_Change Index
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
    Select Case Index
        Case Desc
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                FrArea.Visible = False
                If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                    Txt_Validate Index, result
                    If result = True Then Txt_GotFocus Index: Txt(Index).SetFocus: Exit Sub
                    TopCtrl1_eSave
                Else
                    Txt_GotFocus Index
                    Txt(Index).SetFocus
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
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Code
            Set Rst = GCn.Execute("Select * From Area Where AreaCode='" & PubSiteCode + Trim(Txt(Code)) & "'")
            If TopCtrl1.TopText2 = "Add" Then
                If Not Rst.EOF Then MsgBox " Code Already Exists", vbInformation, "Validation": Txt(Code) = Txt(Code).Tag: Cancel = True: Exit Sub
            ElseIf TopCtrl1.TopText2 = "Edit" Then
                If Not Rst.EOF Then
                    If Rst!AreaCode <> RstMain!AreaCode Then MsgBox "Code Already Exists", vbInformation, "Validation": Txt(Code) = Txt(Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Desc
            Set Rst = GCn.Execute("Select * From Area Where Site_Code='" & PubSiteCode & "' And AreaName='" & Txt(Desc) & "'")
            If TopCtrl1.TopText2 = "Add" Then
                If Not Rst.EOF Then MsgBox "Area Name Already Exists", vbInformation, "Validation": Txt(Desc) = Txt(Desc).Tag: Cancel = True: Exit Sub
            ElseIf TopCtrl1.TopText2 = "Edit" Then
                If Not Rst.EOF Then
                    If Rst!AreaName <> RstMain!AreaName Then MsgBox "Area Name Already Exists", vbInformation, "Validation": Txt(Desc) = Txt(Desc).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub

Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
    Next I
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
End Sub
