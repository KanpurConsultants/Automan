VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmInspCat 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Inspection Category  Master"
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   2805
      TabIndex        =   4
      Top             =   1455
      Width           =   1665
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   2805
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1200
      Width           =   900
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   2805
      MaxLength       =   2
      TabIndex        =   1
      Top             =   690
      Width           =   900
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   2805
      MaxLength       =   25
      TabIndex        =   2
      Top             =   945
      Width           =   4680
   End
   Begin VB.Frame FrJob 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   2805
      TabIndex        =   5
      Top             =   2595
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
         BackColor       =   16777215
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
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
            Caption         =   "Categories"
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
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
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
         Caption         =   "List of Categories"
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
      TabIndex        =   8
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   2805
      TabIndex        =   13
      Top             =   1860
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   0
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   16379351
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print On*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   915
      TabIndex        =   12
      Top             =   1485
      Width           =   795
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print / Display Index"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   915
      TabIndex        =   11
      Top             =   1230
      Width           =   1770
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   915
      TabIndex        =   10
      Top             =   720
      Width           =   555
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   915
      TabIndex        =   9
      Top             =   975
      Width           =   600
   End
End
Attribute VB_Name = "frmInspCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Dim mFlag As Byte
Private Const Code = 0, Desc = 1, PIndex = 2, POn = 3
Dim ListArray As Variant
Dim mListItem As ListItem

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
TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = "Inspection Category Master"   ': TopCtrl1.TopText1.Width = 1000
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select * From Inspection_catg  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by Insp_Code", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select Inspection_catg.Insp_Code as SearchCode, Inspection_catg.* From Inspection_catg Order by Insp_Code", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 Inspection_catg.Insp_Code as SearchCode, Inspection_catg.* From Inspection_catg Order by Insp_Code", GCn, adOpenDynamic, adLockOptimistic
End If

Set RstHelp = New ADODB.Recordset
'RstHelp.Open "Select Insp_Code as code,Insp_description as name FROM Inspection_catg  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by Insp_Code", GCn, adOpenDynamic, adLockOptimistic
RstHelp.Open "Select Insp_Code as code,Insp_description as name FROM Inspection_catg  Order by Insp_Code", GCn, adOpenDynamic, adLockOptimistic
Set DGJob.DataSource = RstHelp

Disp_Text SETS("INI", Me, RstMain)
ListArray = Array("Inspection Sheet", "Job Card", "None")
Set mListItem = ListView_Items(ListView, txt, POn, ListArray, 3)
CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
MoveRec
ADDFLAG = 0:    mFlag = 0
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
    If GCn.Execute("Select Inspection_Catg from Inspection_Element Where Inspection_Catg= " & Chk_Text(Trim(txt(Code))) & "").RecordCount > 0 Then
        MsgBox "Transaction in Inspection Element !", vbCritical, "Delete Denied"
        Exit Sub
    End If
    If RstMain.RecordCount > 0 Then
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            XBM = RstMain.Bookmark
            transFalg = 1
            GCn.Execute ("delete * from Inspection_catg Where insp_Code= " & Chk_Text(Trim(txt(Code))) & "")
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
    GSQL = "Select Insp_Code as SearchCode, Insp_Code as Code,Insp_description as Name FROM Inspection_catg  Order by Insp_Code"
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
        RstMain.FIND ("searchcode='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select Inspection_catg.Insp_Code as SearchCode, Inspection_catg.* From Inspection_catg Where Inspection_catg.Insp_Code = '" & MyValue & "'  Order by Insp_Code")
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

    mRepName = "InspCat"
    mQRY = "SELECT * from Inspection_Catg"

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
FrJob.Visible = False: FrmList.Visible = False
    transFlag = 0
    If IsValid(txt(Code), "Code") = False Then Exit Sub
    If IsValid(txt(Desc), "Category") = False Then Exit Sub
    If txt(POn) <> "None" Then
        If IsValid(txt(PIndex), "Print Index") = False Then Exit Sub
    End If
    
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From Inspection_catg Where insp_Code= " & Chk_Text(Trim(txt(Code))) & " AND SITE_CODE='" & PubSiteCode & "'").Fields(0) > 0 Then MsgBox "Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus Code: txt(Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        GCn.Execute ("DELETE From Inspection_catg Where insp_Code= " & Chk_Text(Trim(txt(Code))) & " AND SITE_CODE='" & PubSiteCode & "'")
        GCn.Execute ("Insert Into Inspection_catg(Insp_Code,Div_Code,Site_Code,Insp_description,report_index,print_on,U_Name,U_EntDt,U_AE) Values('" & txt(Code) & "','" & PubDivCode & "','" & PubSiteCode & "','" & txt(Desc) & "'," & Val(txt(PIndex)) & ",'" & left(txt(POn), 1) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    ElseIf ADDFLAG = 2 Then
        GCn.Execute ("UPDATE Inspection_catg SET insp_description=" & Chk_Text(txt(Desc)) & ",report_index=" & Val(txt(PIndex)) & ",print_on='" & left(txt(POn), 1) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "' Where insp_Code= " & Chk_Text(Trim(txt(Code))) & "")
    End If
    
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select Inspection_catg.Insp_Code as SearchCode, Inspection_catg.* From Inspection_catg Where Inspection_catg.Insp_Code = " & Chk_Text(Trim(txt(Code))) & "  Order by Insp_Code")
    End If
    RstHelp.Requery
    RstMain.FIND ("insp_Code=" & Chk_Text(Trim(txt(Code))))
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
        TopCtrl1.SetFocus
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(Code).BackColor = CtrlBColOrg:      txt(Code).ForeColor = CtrlFColOrg
    txt(Desc).BackColor = CtrlBColOrg:      txt(Desc).ForeColor = CtrlFColOrg
    txt(PIndex).BackColor = CtrlBColOrg:      txt(Desc).ForeColor = CtrlFColOrg
    txt(POn).BackColor = CtrlBColOrg:      txt(Desc).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
On Error GoTo Errloop
FrJob.Visible = False: FrmList.Visible = False
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    txt(Code) = XNull(RstMain!insp_Code)
    txt(Desc) = XNull(RstMain!Insp_description)
    txt(PIndex) = XNull(RstMain!Report_Index)
    Select Case XNull(RstMain!Print_On)
        Case "I"
                txt(POn) = "Inspection Sheet"
        Case "J"
                txt(POn) = "Job Card"
        Case "N", ""
                txt(POn) = "None"
    End Select
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
    Case POn
        If FrmList.Visible = True Then FrmList.Visible = False
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
    Case POn
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width + 100, 900
        If KeyCode = 13 Or KeyCode = vbKeyTab Then
            FrJob.Visible = False: FrmList.Visible = False
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt_GotFocus Index: txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_GotFocus Index
                txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp Then
            If FrmList.Visible = False Then SendKeys "+{Tab}": KeyCode = 0
        End If
End Select

Select Case Index
    Case Code
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case Desc
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case PIndex
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
CheckQuote (keyascii)
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
    Case POn
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Code
            Set Rst = GCn.Execute("SELECT * FROM Inspection_catg WHERE Insp_Code=" & Chk_Text(Trim(txt(Code))))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox " Code Already Exists", vbInformation, "Validation": txt(Code) = txt(Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!insp_Code <> RstMain!Code Then MsgBox "Code Already Exists", vbInformation, "Validation": txt(Code) = txt(Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Desc
            Set Rst = GCn.Execute("SELECT * FROM Inspection_catg WHERE insp_description=" & Chk_Text(txt(Desc)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Category Already Exists", vbInformation, "Validation": txt(Desc) = txt(Desc).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Insp_description <> RstMain!Insp_description Then MsgBox "Category Already Exists", vbInformation, "Validation": txt(Desc) = txt(Desc).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case POn
            If txt(POn).TEXT <> "" Then txt(POn).TEXT = ListView.SelectedItem.TEXT
            If IsValid(txt(POn), "Print On") = False Then Cancel = True:   Exit Sub
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

