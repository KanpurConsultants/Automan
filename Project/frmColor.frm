VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmColor 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Color Master"
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
   Begin VB.Frame FrCol 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   1425
      TabIndex        =   5
      Top             =   2715
      Visible         =   0   'False
      Width           =   5220
      Begin MSDataGridLib.DataGrid DGCol 
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
            DataField       =   "Col_Code"
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
            DataField       =   "Col_Desc"
            Caption         =   "Description"
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
            DataField       =   "Col_Code"
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
         Caption         =   "List of Color"
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2805
      MaxLength       =   20
      TabIndex        =   2
      Top             =   975
      Width           =   3765
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2805
      MaxLength       =   4
      TabIndex        =   1
      Top             =   690
      Width           =   900
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   1185
      TabIndex        =   4
      Top             =   1005
      Width           =   1200
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   4
      Left            =   1185
      TabIndex        =   3
      Top             =   720
      Width           =   570
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Dim mFlag As Byte
Private Const Col_Code = 0, Col_Desc = 1

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
TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = "Color Master"
Set RstMain = New ADODB.Recordset
   Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
     sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
    sitecond = ""
    End If
    
If PubMoveRecYn Then
    RstMain.Open "Select col_code as searchcode, ColMast.* From ColMast " & sitecond & "Order by Col_Code", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 col_code as searchcode, ColMast.* From ColMast " & sitecond & " Order by Col_Code", GCn, adOpenDynamic, adLockOptimistic
End If
Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select col_code as searchcode,col_code,Col_Desc FROM ColMast Order by Col_Code", GCn, adOpenDynamic, adLockOptimistic
CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0:    mFlag = 0
Set DGCol.DataSource = RstHelp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form_Unload (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
BlankText
If ADDFLAG <> 1 Then Disp_Text SETS("ADD", Me, RstMain)
'txt(Col_Code).Tag = txt(Col_Code)
Txt(Col_Code) = PubSiteCode
Txt_GotFocus Col_Code
Txt(Col_Code).SelStart = Len(Txt(Col_Code))
Txt(Col_Code).SetFocus
ADDFLAG = 1
Exit Sub

ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(Col_Code).Enabled = False
    Txt(Col_Desc).Tag = Txt(Col_Desc)
    Txt_GotFocus Col_Desc
    ADDFLAG = 2
    Txt(Col_Desc).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ErrLoop
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
            GCn.Execute ("delete * from ColMast where col_code= '" & Txt(Col_Code) & "'")
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
ErrLoop:    If transFalg = 1 Then GCn.RollbackTrans
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
Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
     sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
    sitecond = ""
    End If
    
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT Col_Code as SearchCode,Col_Code,Col_Desc FROM COLMAST " & sitecond & " order by COL_code"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim rep As CrystalReport, Form1 As frmMastList
Set Form1 = New frmMastList
With Form1
    .g_FormID = 14
    .LblName.CAPTION = Me.CAPTION
    .CAPTION = Me.CAPTION
    .Show
End With
Set Form1 = Nothing
Set rep = Nothing

End Sub

Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo ErrLoop
    transFlag = 0
    If Len(Trim(Txt(Col_Code))) = 1 Then MsgBox "Code should be filled ", vbOKOnly, "Validation": Txt(Col_Code).SetFocus: Exit Sub ' Txt_GotFocus Col_Code: Exit Sub
    
    If IsValid(Txt(Col_Desc), "Color Description") = False Then Txt_GotFocus Col_Desc: Exit Sub
    If TopCtrl1.TopText2 = "Add" Then
        If GCn.Execute("Select COUNT(*) From ColMast Where Col_Code= '" & Txt(Col_Code) & "'").Fields(0) > 0 Then MsgBox "Color Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus Col_Code: Txt(Col_Code).SetFocus: Exit Sub
    End If
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("Insert Into ColMast(Col_Code,Site_Code,Col_Desc,U_Name,U_EntDt,U_AE) Values('" & Txt(Col_Code) & "','" & PubSiteCode & "'," & Chk_Text(Txt(Col_Desc)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
    Else
        GCn.Execute ("UPDATE ColMast SET Col_Desc=" & Chk_Text(Txt(Col_Desc)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "' Where Col_Code= '" & Txt(Col_Code) & "'")
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select col_code as searchcode, ColMast.* From ColMast Where col_code  = " & Txt(Col_Code) & " Order by Col_Code")
    End If
    RstHelp.Requery
    RstMain.FIND ("searchcode='" & Txt(Col_Code) & "'")
    If TopCtrl1.TopText2 = "Add" Then
        TopCtrl1_eAdd
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrCol.Visible = False
    End If
Exit Sub
ErrLoop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
        CtrlClckCol
        FrCol.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    Txt(Col_Code).BackColor = CtrlBColOrg:      Txt(Col_Code).ForeColor = CtrlFColOrg
    Txt(Col_Desc).BackColor = CtrlBColOrg:      Txt(Col_Desc).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
On Error GoTo ErrLoop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    Txt(Col_Code) = XNull(RstMain!Col_Code)
    Txt(Col_Desc) = XNull(RstMain!Col_Desc)
End If
'TopCtrl1.tPrn = False
TopCtrl1.tDel = False
Exit Sub
ErrLoop:        MsgBox err.Description
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
RstHelp.FIND "searchcode  >=" & Chk_Text(Trim(Txt(Col_Code)))
End Sub

Private Sub ColNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Col_Desc >= " & Chk_Text(XNull(Txt(Col_Desc)))
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Col_Code, Col_Desc
            If RstHelp.RecordCount = 0 Then Exit Sub
            If FrCol.Visible = True Then FrCol.Visible = False
            FrCol.Visible = True
            FrCol.top = Txt(Index).top + Txt(Index).height + 10
            FrCol.left = Txt(Index).left
            FrCol.ZOrder 0
    End Select
End If
End Sub

Private Sub Txt_GotFocus(Index As Integer)
DGCol.Columns(0).width = 1000.1: DGCol.Columns(1).width = 3535.024: DGCol.Columns(2).width = 1000.1
Dim mBookMark
    Ctrl_GetFocus Txt(Index)
    mFlag = 0
    If FrCol.Visible = True Then FrCol.Visible = False
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Select Case Index
        Case Col_Code, Col_Desc
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case Col_Code
            DGCol.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "searchcode ASC"
            RstHelp.Bookmark = mBookMark
            ColCodeSearch
        Case Col_Desc
            DGCol.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Col_Desc ASC"
            RstHelp.Bookmark = mBookMark
            ColNameSearch
    End Select
    If Txt(Index) = "" Then Txt_Change Index
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean   ', i As Integer
'If KeyCode = vbKeyEscape Then Txt(Index).Text = ""
Select Case Index
    Case Col_Desc
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            FrCol.Visible = False
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
    Case Col_Code
        'SiteCode Edit restricted
        KeyCode = RestrictKey(1, KeyCode, Txt(Index), Shift)
        
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
'        ElseIf KeyCode = vbKeyUp Then
'            If Len(Txt(Index)) = 1 Then
'                KeyCode = 0
'            End If
        End If
        
End Select
If FrCol.Visible = False Then
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> Col_Code Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Col_Desc Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case Col_Code
        KeyAscii = RestrictKey(1, KeyAscii, Txt(Index), 0)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case Col_Code
        ColCodeSearch
    Case Col_Desc
        ColNameSearch
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Col_Code
            Set Rst = GCn.Execute("SELECT * FROM ColMast WHERE Col_Code=" & Chk_Text(Trim(Txt(Col_Code))))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Color Code Already Exists", vbInformation, "Validation": Txt(Col_Code) = Txt(Col_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Col_Code <> RstMain!Col_Code Then MsgBox "Color Code Already Exists", vbInformation, "Validation": Txt(Col_Code) = Txt(Col_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Col_Desc
            Set Rst = GCn.Execute("SELECT * FROM ColMast WHERE Col_Desc=" & Chk_Text(Txt(Col_Desc)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Color Name Already Exists", vbInformation, "Validation": Txt(Col_Desc) = Txt(Col_Desc).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Col_Desc <> RstMain!Col_Desc Then MsgBox "Color Name Already Exists", vbInformation, "Validation": Txt(Col_Desc) = Txt(Col_Desc).Tag: Cancel = True: Exit Sub
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
'CmbOrder.Enabled = IIf(AddFlag = 1, True, False)
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
Next
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select col_code as searchcode, ColMast.* From ColMast Where col_code  = " & MyValue & " Order by Col_Code")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

    
