VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmProduceMaster 
   Caption         =   "Produce Master"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11205
   Begin VB.Frame FrCity 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   660
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   4980
      Begin MSDataGridLib.DataGrid DGCity 
         Height          =   3225
         Left            =   30
         TabIndex        =   4
         Top             =   345
         Width           =   4920
         _ExtentX        =   8678
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
            DataField       =   "ModelCat_Code"
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
            DataField       =   "ModelCat_Name"
            Caption         =   "Name"
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
            DataField       =   "ModelCat_Code"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
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
         Caption         =   "List of Produce"
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
         TabIndex        =   5
         Top             =   30
         Width           =   4935
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
      Left            =   2355
      MaxLength       =   15
      TabIndex        =   2
      Top             =   840
      Width           =   4245
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
      Left            =   2355
      MaxLength       =   3
      TabIndex        =   1
      Top             =   555
      Width           =   1260
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   661
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produce Name"
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
      Left            =   630
      TabIndex        =   7
      Top             =   840
      Width           =   1560
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produce Code"
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
      Left            =   630
      TabIndex        =   6
      Top             =   555
      Width           =   1500
   End
End
Attribute VB_Name = "FrmProduceMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Don't Change Tag Property of (Txt) Control as it is used in other activities
'FORM COLOR &H00C0FFFF&
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Private Const Produce_Code = 0, Produce_Name = 1
'd


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
TopCtrl1.Tag = PubUParam    ': TopCtrl1.TopText1 = Me.Caption '"Vehicle Model Category Master"
Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select Produce_Code,Produce_Name FROM Produce where left(Produce_Code,1)='" & PubDivCode & "' Order by Produce_Name", GCn, adOpenDynamic, adLockOptimistic
Set DGCity.DataSource = RstHelp

FrCity.Visible = False
Set RstMain = New ADODB.Recordset


Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
If PubMoveRecYn Then
    RstMain.Open "Select Produce_Code as searchcode,Produce.* From Produce  where left(Produce_Code,1)='" & PubDivCode & "' " & sitecond & " Order by Produce_Name", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 Produce_Code as searchcode,Produce.* From Produce  where left(Produce_Code,1)='" & PubDivCode & "' " & sitecond & " Order by Produce_Name", GCn, adOpenDynamic, adLockOptimistic
End If
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
CtrlClckCol
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form_Unload (-1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing
    Set RstHelp = Nothing
End Sub
Private Sub CtrlClckCol()
    Txt(Produce_Code).BackColor = CtrlBColOrg:      Txt(Produce_Code).ForeColor = CtrlFColOrg
    Txt(Produce_Name).BackColor = CtrlBColOrg:      Txt(Produce_Name).ForeColor = CtrlFColOrg
End Sub
Private Sub Disp_Text(Enb As Boolean)
    Txt(Produce_Code).Enabled = Enb
    Txt(Produce_Name).Enabled = Enb
End Sub
Private Sub MakeBlank()
    Txt(Produce_Code) = ""
    Txt(Produce_Name) = ""
End Sub
Private Sub MoveRec()
On Error GoTo ErrLoop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    Txt(Produce_Code) = XNull(RstMain!Produce_Code)
    Txt(Produce_Name) = XNull(RstMain!Produce_Name)
End If
TopCtrl1.tDel = False
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
MakeBlank
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
Txt(Produce_Code).Tag = Txt(Produce_Code)
Txt_GotFocus Produce_Code
Txt(Produce_Code) = PubDivCode
Txt(Produce_Code).SelStart = Len(Txt(Produce_Code))
Txt(Produce_Code).SetFocus
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(Produce_Code).Enabled = False
    Txt(Produce_Name).Tag = Txt(Produce_Name)
    Txt_GotFocus Produce_Name
    Txt(Produce_Name).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ErrLoop
Dim transFalg As Byte, XBM As Variant, Res As Byte
transFalg = 0
If RstMain.RecordCount > 0 Then
    If MsgBox("Are You Sure to Delete This Record ?", vbYesNo, "Confirmation") = vbYes Then
        GCn.BeginTrans
        XBM = RstMain.Bookmark
        transFalg = 1
        GCn.Execute ("Delete From Produce Where Produce_Code=" & Chk_Text(Trim(Txt(Produce_Code))))
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
    MsgBox "No Records To Delete", vbInformation, "Information"
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
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "Select Produce_Code AS SEARCHCODE,Produce_Name ,Produce_Code FROM Produce where left(Produce_Code,1)='" & PubDivCode & "' " & sitecond & " Order by Produce_Name"
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
        .g_FormID = 3
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
    If Len(Trim(Txt(Produce_Code))) = 1 Then MsgBox "Category Code should be filled ", vbOKOnly, "Validation": Txt(Produce_Code).SetFocus: Exit Sub
    If IsValid(Txt(Produce_Name), "Category Name") = False Then Txt_GotFocus Produce_Name: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From Produce Where Produce_Code=" & Chk_Text(Trim(Txt(Produce_Code)))).Fields(0) > 0 Then MsgBox "Category Code Already Exists", vbInformation, "Godown Code Validation": Txt_GotFocus Produce_Code: Txt(Produce_Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("DELETE From Produce Where Produce_Code=" & Chk_Text(Trim(Txt(Produce_Code))))
        GCn.Execute ("Insert Into Produce (Produce_Code,Site_Code,Produce_Name,U_Name,U_EntDt,U_AE) Values('" & Trim(Txt(Produce_Code)) & "','" & PubSiteCode & "'," & Chk_Text(Txt(Produce_Name)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Else
        GCn.Execute ("update  Produce set Site_Code='" & PubSiteCode & "',Produce_Name=" & Chk_Text(Txt(Produce_Name)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "'" & " Where Produce_Code=" & Chk_Text(Trim(Txt(Produce_Code))))
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    RstMain.Requery
    RstHelp.Requery
    RstMain.FIND ("Produce_Code=" & Chk_Text(Trim(Txt(Produce_Code))))
    If ADDFLAG = 1 Then
        MakeBlank
        Txt_GotFocus Produce_Code
        Txt(Produce_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrCity.Visible = False
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
        MoveRec
        CtrlClckCol
        FrCity.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub Produce_CodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Produce_Code >=" & Chk_Text(XNull(Trim(Txt(Produce_Code))))
End Sub
Private Sub Produce_NAMESearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Produce_Name >=" & Chk_Text(XNull(Txt(Produce_Name)))
End Sub
Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Produce_Code, Produce_Name
            FrCity.Visible = True
            FrCity.top = Txt(Index).top + Txt(Index).height + 10
            FrCity.left = Txt(Index).left
            FrCity.ZOrder 0
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
    If FrCity.Visible = True Then FrCity.Visible = False
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    DGCity.Columns(0).width = 1000.1: DGCity.Columns(1).width = 3000: DGCity.Columns(2).width = 800
    Select Case Index
        Case Produce_Code
            DGCity.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Produce_Code ASC"
            RstHelp.Bookmark = mBookMark
            Produce_CodeSearch
        Case Produce_Name
            DGCity.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Produce_Name ASC"
            RstHelp.Bookmark = mBookMark
            Produce_NAMESearch
    End Select
    If Txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Select Case Index
    Case Produce_Code
        'Div Code Edit restricted
        KeyCode = RestrictCode(KeyCode, Txt(Index), Shift)
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case Produce_Name
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_Click Index
                Txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case Produce_Code
        KeyAscii = RestrictCode(KeyAscii, Txt(Index), 0)
End Select

End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case Produce_Code
        Produce_CodeSearch
    Case Produce_Name
        Produce_NAMESearch
End Select
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Produce_Code
            Set Rst = GCn.Execute("SELECT * FROM Produce WHERE Produce_Code=" & Chk_Text(Txt(Produce_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Category Code Already Exists", vbInformation, "Validation": Txt(Produce_Code) = Txt(Produce_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Produce_Code <> RstMain!Produce_Code Then MsgBox "Category Code Already Exists", vbInformation, "Validation": Txt(Produce_Code) = Txt(Produce_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Produce_Name
            Set Rst = GCn.Execute("SELECT * FROM Produce WHERE Produce_Name=" & Chk_Text(Txt(Produce_Name)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Category Name Already Exists", vbInformation, "Validation": Txt(Produce_Name) = Txt(Produce_Name).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Produce_Name <> RstMain!Produce_Name Then MsgBox "Category Name Already Exists", vbInformation, "Validation": Txt(Produce_Name) = Txt(Produce_Name).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select Produce_Code as searchcode,Produce.* From Produce  where left(Produce_Code,1)='" & PubDivCode & "' And Produce_Code  = '" & MyValue & "' Order by Produce_Name")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub


