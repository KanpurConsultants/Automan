VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form FaNarrMast 
   BackColor       =   &H00CAF1FD&
   Caption         =   "Narration Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10515
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   661
   End
   Begin VB.Frame FrameList 
      BackColor       =   &H00FF0000&
      Caption         =   "Narration List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3705
      Left            =   1440
      TabIndex        =   3
      Top             =   2100
      Width           =   4740
      Begin MSDataGridLib.DataGrid DGMaster 
         Bindings        =   "FaNarrMast.frx":0000
         Height          =   3375
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Name"
            Caption         =   "Narration"
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
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   3899.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   1950
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   930
      Width           =   4980
   End
   Begin MSDataGridLib.DataGrid DGHelp 
      Height          =   3330
      Left            =   6480
      Negotiate       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
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
         DataField       =   "Name"
         Caption         =   "Narration"
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
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   975
      TabIndex        =   2
      Top             =   945
      Width           =   765
   End
End
Attribute VB_Name = "FaNarrMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mSname As String, RsHelp As ADODB.Recordset, Master As ADODB.Recordset, mOldName As String
Private Const Name1 As Byte = 0, ShortName As Byte = 1
Private PubDatamanFa As New DMFa.ClsFa

Private Sub DGHelp_Click()
    DGHelp.Visible = False
    If RsHelp.RecordCount > 0 Then
        Txt(Name1).Tag = RsHelp!Code
        Txt(Name1).Text = RsHelp!Name
    End If
    Txt(Name1).SetFocus
End Sub
Private Sub DGMaster_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then Exit Sub
    MoveRec
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FaFormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
TopCtrl1.TopText1 = Me.CAPTION
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
DGHelp.left = Txt(Name1).left
DGHelp.top = Txt(Name1).top + Txt(Name1).height + 30
Set RsHelp = New ADODB.Recordset
RsHelp.CursorLocation = adUseClient
RsHelp.Open "Select Code,Name From NarrMast Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
Set DGHelp.DataSource = RsHelp
Set Master = New ADODB.Recordset
Master.CursorLocation = adUseClient
If PubSiteCodeWiseMasterRst = True Then
    Master.Open "Select N.*,N.Code as SearchCode From NarrMast N WHERE LEFT(CODE,1)='" & PubSiteCode & "' Order by N.Name", G_FaCn, adOpenDynamic, adLockOptimistic
Else
    Master.Open "Select N.*,N.Code as SearchCode From NarrMast N Order by N.Name", G_FaCn, adOpenDynamic, adLockOptimistic
End If
Set DGMaster.DataSource = Master
Disp_Text SETS("INI", Me, Master)
MoveRec
Me.height = 7200
Me.width = 7755
Me.top = 0
Me.left = 0
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_Unload (-1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsHelp = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    Disp_Text SETS("ADD", Me, Master)
    BlankText
    Txt(Name1).SetFocus
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM, J As Byte, mBeginTrans As Byte
On Error GoTo ELoop
If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    mBeginTrans = 1
    G_FaCn.BeginTrans
    XBM = Master.Bookmark
    G_FaCn.Execute ("Delete From NarrMast Where Code='" & Master!SearchCode & "'")
    G_FaCn.CommitTrans
    mBeginTrans = 0
    Master.Requery
    RsHelp.Requery
    If Master.RecordCount >= XBM Then
        Master.Bookmark = XBM
    Else
        If Master.EOF = False Then Master.MoveLast
    End If
    MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
ELoop:  If mBeginTrans = 1 Then G_FaCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    mOldName = Txt(Name1).Text
    Txt(Name1).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select N.Code as SearchCode,N.Name FROM NarrMast N Order by N.Name"
    Set SearchForm = Me
    FAFind.Show vbModal
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.Find ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eCancel()
Dim I As Byte
On Error GoTo ELoop
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
    End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_ePrn()
Dim X1, Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
If Master.RecordCount <= 0 Then Exit Sub
Set Rst = G_FaCn.Execute("Select * FROM NarrMast Order by Name")
'X1 = CreateFieldDefFile(Rst, PubFaReportPath + "\FaNarrMast.ttx", True)
Set rpt = PubDatamanFa.FaNarrMastRpt
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("Title")
            rpt.FormulaFields(I).Text = "'Narration List'"
    End Select
Next
rpt.Database.SetDataSource Rst
rpt.ReadRecords
FaReport_View rpt, 0, Me.CAPTION, True
Set Rst = Nothing
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eRef()
    RsHelp.Requery
'    Master.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim Rst As ADODB.Recordset, SearchCode As String, MaxCode As Integer, mBeginTrans As Byte
On Error GoTo ELoop
Grid_Hide
If FaIsValid(Txt(Name1), "Godown Name") = False Then Exit Sub
If RsHelp.RecordCount <> 0 Then
    If TopCtrl1.TopText2 = "Add" Then
        If G_FaCn.Execute("Select Count(*) From NarrMast Where Name='" & Txt(Name1) & "'").Fields(0).Value > 0 Then MsgBox "Duplicate Godown Name", vbInformation, "Information": Txt(Name1).SetFocus: Exit Sub
    Else
        If G_FaCn.Execute("Select Count(*) From NarrMast Where Name='" & Txt(Name1) & "' And Name<>'" & mOldName & "'").Fields(0).Value > 0 Then MsgBox "Duplicate Godown Name", vbInformation, "Information": Txt(Name1).SetFocus:  Exit Sub
    End If
End If
mBeginTrans = 1
G_FaCn.BeginTrans
If TopCtrl1.TopText2.CAPTION = "Add" Then
    If PubBackEnd = "A" Then
        Set Rst = G_FaCn.Execute("Select Max(MID(Code,3,Len(Code)-2)) As tCode From NarrMast Where Left(Code,2)='" & PubSiteCode & left(Txt(Name1), 1) & "'")
    ElseIf PubBackEnd = "S" Then
        Set Rst = G_FaCn.Execute("Select Max(SubString(Code,3,Len(Code)-2)) As tCode From NarrMast Where Left(Code,2)='" & PubSiteCode & left(Txt(Name1), 1) & "'")
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
    Txt(Name1).Tag = PubSiteCode & UCase(left(Txt(Name1), 1)) & Format(MaxCode, "000")
    Replace Txt(Name1), Chr(13), ""
    G_FaCn.Execute ("Insert Into NarrMast(Code,Name,U_Name,U_EntDt,U_AE) Values('" & Txt(Name1).Tag & "','" & Replace(Replace(Txt(Name1), Chr(10), ""), Chr(13), "") & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A')")
Else
    G_FaCn.Execute ("Update NarrMast Set Name='" & Replace(Replace(Txt(Name1), Chr(10), ""), Chr(13), "") & "',U_Name='" & pubUName & "',U_EntDt=" & FaConvertDate(PubLoginDate) & ",U_AE='E' Where Code='" & Txt(Name1).Tag & "'")
End If
G_FaCn.CommitTrans
mBeginTrans = 0
SearchCode = Txt(Name1).Tag
Master.Requery
RsHelp.Requery
Master.Find "SearchCode ='" & SearchCode & "'"
If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, Master)
Set Rst = Nothing
Exit Sub
ELoop:  If mBeginTrans = 1 Then G_FaCn.RollbackTrans
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Txt_GotFocus(Index As Integer)
    FaCtrl_GetFocus Txt(Index)
    Grid_Hide
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case Name1
        FaDGridTxtKeyDown_Mast DGHelp, Txt, Name1, RsHelp, KeyCode, False, 1
End Select
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    FaCheckQuote KeyAscii
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case Name1
        If DGHelp.Visible = True Then FaDGridTxtKeyUp_Mast Txt, Name1, RsHelp, KeyCode, "Name"
End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    FaCtrl_validate Txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case Name1
        If RsHelp.RecordCount = 0 Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then
            If G_FaCn.Execute("Select Count(*) From NarrMast Where Name='" & Txt(Name1) & "'").Fields(0).Value > 0 Then MsgBox "Duplicate Godown Name", vbInformation, "Information": Txt(Name1).SetFocus: Exit Sub
        Else
            If G_FaCn.Execute("Select Count(*) From NarrMast Where Name='" & Txt(Name1) & "' And Name<>'" & mOldName & "'").Fields(0).Value > 0 Then MsgBox "Duplicate Godown Name", vbInformation, "Information": Txt(Name1).SetFocus:  Exit Sub
        End If
End Select
End Sub
Private Sub BlankText()
Dim I As Byte
mOldName = ""
For I = 0 To Txt.Count - 1
    Txt(I).Text = ""
    Txt(I).Tag = ""
Next I
End Sub
Private Sub MoveRec()
On Error GoTo ELoop
If Master.RecordCount > 0 Then
    Txt(Name1).Tag = Master!Code
    Txt(Name1) = FaXNull(Master!Name)
Else
    BlankText
End If
Grid_Hide
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
Next
DGMaster.Enabled = Not Enb
End Sub
Private Sub Grid_Hide()
    If DGHelp.Visible = True Then DGHelp.Visible = False
End Sub
Private Sub SaveMsg(Index As Integer)
Grid_Hide
If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
    TopCtrl1_eSave
Else
    Txt(Index).SetFocus
End If
End Sub

