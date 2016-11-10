VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "topctl.ocx"
Begin VB.Form FaTDSCat 
   BackColor       =   &H00FCB5A0&
   Caption         =   "T.D.S.Category Entry"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   ForeColor       =   &H00000000&
   Icon            =   "FaTDSCat.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   8880
   Begin MSDataGridLib.DataGrid Dg 
      Height          =   2460
      Left            =   60
      TabIndex        =   3
      Top             =   1665
      Visible         =   0   'False
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   4339
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   13234931
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Category"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
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
         DataField       =   "TDSLimit"
         Caption         =   "Limit"
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
      BeginProperty Column03 
         DataField       =   "TDS_Percentage"
         Caption         =   "%"
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
         MarqueeStyle    =   5
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            DividerStyle    =   1
            ColumnWidth     =   134.929
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   1
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            DividerStyle    =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            DividerStyle    =   1
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F2F3&
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
      Index           =   2
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1290
      Width           =   1440
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F2F3&
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
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1020
      Width           =   1440
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F2F3&
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
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   0
      Top             =   750
      Width           =   4215
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.D.S.%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   870
      TabIndex        =   7
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.D.S.Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   870
      TabIndex        =   6
      Top             =   1020
      Width           =   1140
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   870
      TabIndex        =   4
      Top             =   750
      Width           =   1200
   End
End
Attribute VB_Name = "FaTDSCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CtrlBColOrg = &HC2FAF5, CtrlFColOrg = &H80000008
Dim ADDFLAG As Byte, Master As ADODB.Recordset, RstHelp As ADODB.Recordset, mFlag As Byte
Private Const Name1 As Byte = 0, TDSLimit As Byte = 1, TDSPercentage As Byte = 2
Private PubDatamanFa As New DMFa.ClsFa

Private Sub Form_Load()
Dim I As Byte
On Error GoTo Errloop
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    If PubSec = "SANJEEV" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_NAME='" & Me.CAPTION & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    ElseIf PubSec = "RAHUL" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_CODE='" & Me.Name & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    End If
    '''''''''''''
    PubDatamanFa.FaBackEnd = PubBackEnd
    PubDatamanFa.FaPubLoginDate = PubLoginDate
    PubDatamanFa.FaPubDivCode = PubDivCode
    PubDatamanFa.FaPubSiteCode = PubSiteCode
    PubDatamanFa.FaPubSiteCodeDisplay = PubSiteCodeDisplay
    PubDatamanFa.FaPubSiteName = PubSiteName
    PubDatamanFa.FapubUName = pubUName
    PubDatamanFa.FaDosPort = PubFaDosPort
    PubDatamanFa.FaRunPIF = PubRunPIF
    PubDatamanFa.FaPubSiteType = PubFaSiteType
    Set PubDatamanFa.SetG_FaCn = G_FaCn
    Set PubDatamanFa.SetG_CompCn = G_CompCn
    Set PubDatamanFa.SetrsUserPerm = rsUserPerm.Clone
    Set PubDatamanFa.SetMasterRst = FaMasterRst.Clone
    '''''''''''''
    Me.top = 0
    Me.left = 0
'    Me.BackColor = FrmBackCol
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    If PubSiteCodeWiseMasterRst = True Then
        Set Master = G_FaCn.Execute("SELECT * FROM TDSCAT WHERE LEFT(CODE,1)='" & PubSiteCode & "' ORDER BY Name")
    Else
        Set Master = G_FaCn.Execute("SELECT * FROM TDSCAT ORDER BY Name")
    End If
    Set RstHelp = G_FaCn.Execute("SELECT Code,name FROM TDSCAT ORDER BY Name")
    Dg.left = 0
    Dg.top = 1665
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    ADDFLAG = 0
    mFlag = 0
    Dg.Columns(0).Visible = False
    Set Dg.DataSource = RstHelp
    Me.TopCtrl1.TopText1.left = 5800
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RstHelp = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Errloop
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown, vbKeyUp
            Select Case KeyCode
                Case vbKeyDown, vbKeyUp
                    If Dg.Visible = True Then Exit Sub
            End Select
            If TypeOf Me.ActiveControl Is TextBox Then Txt_Validate Me.ActiveControl.Index, False
            If PubDatamanFa.FaManageKeysControl(Me, KeyCode, Shift) = True Then SaveMsg 0
            KeyCode = 0
        Case Else
            FaFormKeyDown Me, KeyCode, Shift
    End Select
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
Next
End Sub
Private Sub MakeBlank()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
Next
End Sub
Private Sub MoveRec()
On Error GoTo Errloop
FaRstBofEof Master
If Master.RecordCount <= 0 Then
    MakeBlank
Else
    Txt(Name1).TEXT = Master!Name
    Txt(TDSLimit) = FaVNull(Master!TDSLimit)
    Txt(TDSPercentage) = FaVNull(Master!TDSPercentage)
End If
Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo Errloop
    MakeBlank
    ADDFLAG = 1
    Disp_Text SETS("ADD", Me, Master)
    Txt(Name1).SetFocus
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
    If Master.RecordCount > 0 Then
        ADDFLAG = 2
        Disp_Text SETS("EDIT", Me, Master)
        Txt(Name1).SetFocus
    Else
        MsgBox "There Is No Record To Edit.", vbInformation, "Information"
    End If
    Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
Dim transFalg As Byte
transFalg = 0
On Error GoTo Errloop
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure to Delete This Record", vbYesNo, "Confirmation") = vbYes Then
            G_FaCn.BeginTrans
            transFalg = 1
            G_FaCn.Execute ("Delete From TDSCAT Where code='" & Master!Code & "'")
            G_FaCn.CommitTrans
            transFalg = 0
            Master.Requery
            RstHelp.Requery
            Disp_Text SETS("INI", Me, Master)
            MoveRec
        End If
    Else
        MsgBox "There Is No Record To Delete.", vbInformation, "Information"
    End If
    Exit Sub
Errloop:
    If transFalg = 1 Then
        G_FaCn.RollbackTrans
        MsgBox err.Description, vbExclamation, " Deletion Error "
    End If
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select code As SearchCode,Name FROM TDSCAT Order by Name"
    Set SearchForm = Me
    FAFind.Show vbModal
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.Find ("Code='" & MyValue & "'")
    MoveRec
    Exit Sub
ErrorLoop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
Dim RST1 As ADODB.Recordset, X11, I As Integer
On Error GoTo ERRORHANDLER
Set RST1 = G_FaCn.Execute("select * FROM TDSCAT order by Name")
If RST1.RecordCount = 0 Then MsgBox "No record Found to Print": Exit Sub
'X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaTDSCat.ttx", True)
Set rpt = PubDatamanFa.FaTDSCatRpt
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("Title")
            rpt.FormulaFields(I).TEXT = "'T.D.S. Category List'"
    End Select
Next
rpt.Database.SetDataSource RST1
rpt.ReadRecords
FaReport_View rpt, 0, Me.CAPTION, True
Set RST1 = Nothing
Exit Sub
ERRORHANDLER:  MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte, MySql$, mCode As String, Rst As ADODB.Recordset, MaxCode As Long
On Error GoTo Errloop
    Ctrl_BckColor
    transFlag = 0
    If FaIsValid(Txt(Name1), "Description") = False Then Txt_GotFocus Name1: Exit Sub
    If FaIsValid(Txt(TDSLimit), "T.D.S.Limit") = False Then Txt_GotFocus TDSLimit: Exit Sub
    If FaIsValid(Txt(TDSPercentage), "T.D.S.Percentage") = False Then Txt_GotFocus TDSPercentage: Exit Sub
    If ADDFLAG = 1 Then
        If PubBackEnd = "A" Then
            Set Rst = G_FaCn.Execute("Select Max(MID(Code,3,Len(Code)-2)) As tCode From TDSCAT Where Left(Code,2)='" & PubSiteCode & left(Txt(Name1), 1) & "'")
        ElseIf PubBackEnd = "S" Then
            Set Rst = G_FaCn.Execute("Select Max(SubString(Code,3,Len(Code)-2)) As tCode From TDSCAT Where Left(Code,2)='" & PubSiteCode & left(Txt(Name1), 1) & "'")
        End If
        If Rst.RecordCount > 0 Then
            If Not IsNull(Rst!tCode) Then
                MaxCode = Rst!tCode + 1
            Else
                MaxCode = 1
            End If
        Else
            MaxCode = 1
        End If
        mCode = PubSiteCode & UCase(left(Txt(Name1), 1)) & Format(MaxCode, "00")
    Else
        mCode = Master!Code
    End If
    G_FaCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        G_FaCn.Execute "Insert Into TDSCAT (code,name,TDSLimit,TDSPercentage,U_Name,U_EntDt,U_AE) Values('" & mCode & "','" & Txt(Name1).TEXT & "'," & Val(Txt(TDSLimit)) & "," & Val(Txt(TDSPercentage)) & ",'" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A' )"
    Else
        G_FaCn.Execute "UPDATE TDSCAT SET name='" & Txt(Name1) & "',TDSLimit=" & Val(Txt(TDSLimit)) & ",TDSPercentage=" & Val(Txt(TDSPercentage)) & ",U_name='" & pubUName & "',U_EntDt=" & FaConvertDate(PubLoginDate) & " ,U_AE='E' WHERE code='" & mCode & "'"
    End If
    G_FaCn.CommitTrans
    transFlag = 0
    Master.Requery
    RstHelp.Requery
    Dg.Refresh
    Master.Find ("code='" & mCode & "'")
    If ADDFLAG = 1 Then
        MakeBlank
        Txt_GotFocus Name1
        Txt(Name1).SetFocus
    Else
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        ADDFLAG = 0
    End If
    Set Rst = Nothing
    Exit Sub
Errloop:        If transFlag = 1 Then G_FaCn.RollbackTrans
                MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
    Ctrl_BckColor
    ADDFLAG = 0
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
Errloop:    MsgBox err.Description, vbCritical, "Information"
End Sub
Private Sub nameSearch()
    If RstHelp.RecordCount <= 0 Then Exit Sub
    RstHelp.MoveFirst
    RstHelp.Find "name>='" & Txt(Name1) & "'"
End Sub
Private Sub Txt_Change(Index As Integer)
    If ADDFLAG <> 0 Then
        Select Case Index
            Case Name1
                Dg.Visible = True
                Dg.top = Txt(Index).top + Txt(Index).height + 10
                Dg.left = Txt(Index).left
                Dg.ZOrder 0
                nameSearch
        End Select
    End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
On Error GoTo Errloop
    mFlag = 0
    Call Ctrl_GetFocus(Index)
    If Dg.Visible = True Then Dg.Visible = False
    FaRstBofEof RstHelp
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    Select Case Index
        Case Name1
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case Name1
            mBookMark = RstHelp.Bookmark
            RstHelp.Bookmark = mBookMark
    End Select
    If Txt(Index) = "" Then Txt_Change Index
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_Click(Index As Integer)
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim MyResult As Boolean, I As Integer
On Error GoTo Errloop
    If KeyCode = vbKeyEscape Then
        Dg.Visible = False
        Exit Sub
    End If
    Select Case Index
        Case Name1
            FaDGridTxtKeyDown_Mast Dg, Txt, Name1, RstHelp, KeyCode, False, 1
    End Select
    Select Case Index
        Case TDSPercentage
            Select Case KeyCode
                Case 13
                    If MsgBox("Save Record?", vbYesNo, "Save Entry") = vbYes Then
                        TopCtrl1_eSave
                        Exit Sub
                    Else
                        Me.ActiveControl.SetFocus
                    End If
            End Select
    End Select
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case TDSLimit
        FaNumPress Txt(Index), KeyAscii, 8, 2
    Case TDSPercentage
        FaNumPress Txt(Index), KeyAscii, 2, 4
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case Name1
            If Dg.Visible = True Then FaDGridTxtKeyUp_Mast Txt, Name1, RstHelp, KeyCode, "Name"
    End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    Call Ctrl_validate(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo Errloop
Select Case Index
    Case Name1
        If ADDFLAG = 1 Then
            If G_FaCn.Execute("Select COUNT(*) From TDSCAT Where name='" & Txt(Name1).TEXT & "'").Fields(0) > 0 Then MsgBox "Name Already Exists", vbInformation, "Validation": Cancel = True: Exit Sub
        ElseIf ADDFLAG = 2 Then
            If G_FaCn.Execute("Select COUNT(*) From TDSCAT Where name='" & Txt(Name1).TEXT & "' AND CODE<>'" & Master!Code & "'").Fields(0) > 0 Then MsgBox "Name Already Exists", vbInformation, "Validation": Cancel = True: Exit Sub
        End If
End Select
Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Ctrl_GetFocus(Index As Integer)
    Txt(Index).BackColor = CtrlBCol
    Txt(Index).ForeColor = CtrlFCol
    Txt(Index).BorderStyle = 1
End Sub
Private Sub Ctrl_validate(Index As Integer)
    Txt(Index).BackColor = CtrlBColOrg
    Txt(Index).ForeColor = CtrlFColOrg
    Txt(Index).BorderStyle = 0
End Sub
Private Sub Ctrl_BckColor()
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).BorderStyle = 0
    Next
End Sub
Private Sub SaveMsg(Index As Integer)
Dg.Visible = False
If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
    TopCtrl1_eSave
Else
    Txt(Index).SetFocus
End If
End Sub
