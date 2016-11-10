VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmGodown 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Godown Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10320
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
      Left            =   2340
      TabIndex        =   3
      Text            =   "3"
      Top             =   1080
      Width           =   1260
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   2340
      MaxLength       =   3
      TabIndex        =   1
      Top             =   570
      Width           =   630
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
      Left            =   2340
      MaxLength       =   30
      TabIndex        =   2
      Top             =   825
      Width           =   4245
   End
   Begin VB.Frame FrCity 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   660
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   4980
      Begin MSDataGridLib.DataGrid DGCity 
         Height          =   3225
         Left            =   30
         TabIndex        =   8
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
            DataField       =   "God_Code"
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
            DataField       =   "God_Name"
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
            DataField       =   "God_Code"
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
               ColumnWidth     =   3225.26
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
         Caption         =   "List of Godown"
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
         TabIndex        =   6
         Top             =   30
         Width           =   4935
      End
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(S)pare / (V)ehicle"
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
      Left            =   3750
      TabIndex        =   10
      Top             =   1110
      Width           =   1620
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicable For"
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
      Left            =   615
      TabIndex        =   9
      Top             =   1080
      Width           =   1200
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
      Left            =   630
      TabIndex        =   7
      Top             =   570
      Width           =   555
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Godown Name*"
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
      Left            =   630
      TabIndex        =   4
      Top             =   825
      Width           =   1350
   End
End
Attribute VB_Name = "frmGodown"
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
Private Const God_Code As Byte = 0, God_Name As Byte = 1, ApplicableFor As Byte = 2

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
On Error GoTo ELoop

'Me.left = 0: Me.top = 0
WinSetting Me, 5500, 8500
TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = Me.CAPTION '"Godown Master"
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select Godown.*,Godown.God_Code as SearchCode From Godown WHERE SITE_CODE=" & Chk_Text(PubSiteCode) & " Order by God_Name", GCn, adOpenDynamic, adLockOptimistic


    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
If PubMoveRecYn Then
    RstMain.Open "Select Godown.*, Godown.God_Code as SearchCode From Godown " & sitecond & " Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 Godown.*, Godown.God_Code as SearchCode From Godown " & sitecond & " Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
End If

Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select God_Code,God_Name FROM GODOWN Order by God_Name", GCn, adOpenDynamic, adLockOptimistic

Lbl(1).CAPTION = IIf(PubVCompCode = "", "(S)pare", Lbl(1).CAPTION)

CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
Set DGCity.DataSource = RstHelp
FrCity.Visible = False
TopCtrl1.tDel = False

Exit Sub

ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub
Private Sub CtrlClckCol()
    Txt(God_Code).BackColor = CtrlBColOrg:      Txt(God_Code).ForeColor = CtrlFColOrg
    Txt(God_Name).BackColor = CtrlBColOrg:      Txt(God_Name).ForeColor = CtrlFColOrg
End Sub

Private Sub Disp_Text(Enb As Boolean)
    Txt(God_Code).Enabled = Enb
    Txt(God_Name).Enabled = Enb
    Txt(ApplicableFor).Enabled = Enb
End Sub

Private Sub MakeBlank()
    Txt(God_Code) = ""
    Txt(God_Name) = ""
    Txt(ApplicableFor) = ""
End Sub

Private Sub MoveRec()
On Error GoTo ErrLoop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    Txt(God_Code) = XNull(RstMain!God_Code)
    Txt(God_Name) = XNull(RstMain!God_Name)
    Txt(ApplicableFor) = IIf(RstMain!Appli_For = 0, "Spare", "Vehicle")
End If
'TopCtrl1.tPrn = False
TopCtrl1.tDel = False
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
MakeBlank
If ADDFLAG <> 1 Then Disp_Text SETS("ADD", Me, RstMain)
ADDFLAG = 1
Txt(God_Code) = PubSiteCode
Txt_GotFocus God_Code
Txt(God_Code).SelStart = Len(Txt(God_Code))
If PubVCompCode = "" Then
    Txt(ApplicableFor) = "Spare"
End If
Txt(God_Code).SetFocus
Exit Sub

ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(God_Code).Enabled = False
    Txt(God_Code).Tag = Txt(God_Code)
    Txt_GotFocus God_Name
    Txt(God_Name).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
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
    GSQL = "SELECT God_code as searchcode,God_Code,God_Name FROM Godown " & sitecond & " order by God_code"
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
        .g_FormID = 9
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
    If Len(Trim(Txt(God_Code))) = 1 Then MsgBox "Code should be filled ", vbOKOnly, "Validation": Txt(God_Code).SelStart = Len(Txt(God_Code)): Txt(God_Code).SetFocus: Exit Sub ' Txt_GotFocus God_Code: Exit Sub
    
    If IsValid(Txt(God_Name), "Godown Name") = False Then Txt_GotFocus God_Name: Exit Sub
    If IsValid(Txt(ApplicableFor), "Applicable For") = False Then Txt_GotFocus ApplicableFor: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From GODOWN Where GOD_Code=" & Chk_Text(PubSiteCode + Trim(Txt(God_Code)))).Fields(0) > 0 Then MsgBox "Godown Code Already Exists", vbInformation, "Godown Code Validation": Txt_GotFocus God_Code: Txt(God_Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("Insert Into GODOWN (God_Code,Site_Code,God_Name,Appli_For,U_Name,U_EntDt,U_AE) Values('" & Trim(Txt(God_Code)) & "','" & PubSiteCode & "'," & Chk_Text(Txt(God_Name)) & ", " & IIf(Txt(ApplicableFor) = "Spare", 0, 1) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
    Else
        GCn.Execute ("update GODOWN set Site_Code='" & PubSiteCode & "',God_Name=" & Chk_Text(Txt(God_Name)) & ",Appli_For=" & IIf(Txt(ApplicableFor) = "Spare", 0, 1) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & " Where GOD_Code=" & Chk_Text(Trim(Txt(God_Code))))
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select Godown.*, Godown.God_Code as SearchCode From Godown Where Godown.God_Code = " & Chk_Text(Trim(Txt(God_Code))) & " Order by God_Name")
    End If
    RstHelp.Requery
    RstMain.FIND ("GOD_CODE=" & Chk_Text(Trim(Txt(God_Code))))
    If ADDFLAG = 1 Then
        TopCtrl1_eAdd
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrCity.Visible = False
        TopCtrl1.tDel = False
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
        TopCtrl1.tDel = False
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

Private Sub godCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "GOD_CODE  >=" & Chk_Text(XNull(Trim(Txt(God_Code))))
End Sub

Private Sub godNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "GOD_Name  >=" & Chk_Text(XNull(Txt(God_Name)))
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case God_Code, God_Name
            FrCity.Visible = True
            FrCity.top = Txt(Index).top + Txt(Index).height + 10
            FrCity.left = Txt(Index).left
            FrCity.ZOrder 0
    End Select
End If
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    DGCity.Columns(0).width = 1000.1: DGCity.Columns(1).width = 3200: DGCity.Columns(2).width = 800
    Select Case Index
        Case God_Code
            DGCity.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "GOD_CODE ASC"
            RstHelp.Bookmark = mBookMark
            godCodeSearch
        Case God_Name
            DGCity.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "GOD_NAME ASC"
            RstHelp.Bookmark = mBookMark
            godNameSearch
    End Select
    If FrCity.Visible = True Then FrCity.Visible = False
    If Txt(Index) = "" Then Txt_Change Index
End Sub

Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Select Case Index
    Case God_Code, God_Name
        If Index = God_Code Then KeyCode = RestrictKey(1, KeyCode, Txt(Index), Shift)
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case ApplicableFor
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_Click Index
                Txt(ApplicableFor).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case God_Code
        KeyAscii = RestrictKey(1, KeyAscii, Txt(Index), 0)
End Select

End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case God_Code
        godCodeSearch
    Case God_Name
        godNameSearch
    Case ApplicableFor
        If KeyCode = vbKeyReturn Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Exit Sub
        If KeyCode = vbKeyS Then
            Txt(Index) = "Spare"
        ElseIf KeyCode = vbKeyV Then
            Txt(Index) = "Vehicle"
        End If
End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo lblExit
Dim Rst As ADODB.Recordset
    Select Case Index
        Case God_Code
            Set Rst = GCn.Execute("SELECT god_code FROM GODOWN WHERE GOD_CODE=" & Chk_Text(Txt(God_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Godown Code Already Exists", vbInformation, "Validation": Txt(God_Code) = Txt(God_Code).Tag: Txt(God_Code).SelStart = Len(Trim(Txt(God_Code))): Cancel = True: GoTo lblExit
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!God_Code <> RstMain!God_Code Then MsgBox "Godown Code Already Exists", vbInformation, "Validation": Txt(God_Code) = Txt(God_Code).Tag: Cancel = True:   GoTo lblExit
                End If
            End If
        Case God_Name
            Set Rst = GCn.Execute("SELECT God_Name FROM GODOWN WHERE GOD_NAME=" & Chk_Text(Txt(God_Name)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Godown Name Already Exists", vbInformation, "Validation": Txt(God_Name) = Txt(God_Name).Tag: Cancel = True:  GoTo lblExit
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!God_Name <> RstMain!God_Name Then MsgBox "Godown Name Already Exists", vbInformation, "Validation": Txt(God_Name) = Txt(God_Name).Tag: Cancel = True:   GoTo lblExit
                End If
            End If
        Case ApplicableFor
            If Txt(ApplicableFor) = "" Then MsgBox "Applicable for Spare / Vehicle", vbInformation, "Validation": Cancel = True:  GoTo lblExit
    End Select
lblExit:
    Set Rst = Nothing
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select Godown.*, Godown.God_Code as SearchCode From Godown Where Godown.God_Code = '" & MyValue & "' Order by God_Name")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub


