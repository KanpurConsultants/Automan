VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmAggregate 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Aggregate Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10320
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
      Height          =   210
      Index           =   0
      Left            =   2355
      MaxLength       =   2
      TabIndex        =   1
      Top             =   825
      Width           =   1260
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
      Height          =   210
      Index           =   2
      Left            =   2355
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "No"
      Top             =   1305
      Width           =   1260
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
      Height          =   210
      Index           =   1
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1065
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
      Width           =   4590
      Begin MSDataGridLib.DataGrid DGCity 
         Height          =   3225
         Left            =   30
         TabIndex        =   9
         Top             =   345
         Width           =   4530
         _ExtentX        =   7990
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
            DataField       =   "Aggre_Code"
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
            DataField       =   "Aggre_Name"
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
            DataField       =   "Aggre_Code"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   2805.166
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
         Caption         =   "List of Aggregate Groups"
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
         Width           =   4530
      End
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es/(N)o"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   750
      TabIndex        =   8
      Top             =   825
      Width           =   450
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine"
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
      Index           =   3
      Left            =   750
      TabIndex        =   7
      Top             =   1305
      Width           =   570
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aggregate Name"
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
      Left            =   750
      TabIndex        =   4
      Top             =   1065
      Width           =   1440
   End
End
Attribute VB_Name = "frmAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Don't Change Tag Property of (Txt) Control as it is used in other activities
'FORM COLOR &H00C0FFFF&
Option Explicit
Public MasterFormExit As Boolean
'Private Const CtrlBColOrg = &HC2D5B9           'Orginal BackColour
'Private Const CtrlFColOrg = &H80000012      'Orginal ForeColour
'Private Const CtrlBCol = &H80000008         'Changed BackColour
'Private Const CtrlFCol = &H8000000E         'Changed ForeColour
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Private Const Aggre_Code = 0, Aggre_Name = 1, AggreEngine = 2

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
TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = "Aggregate Group Master"             ': TopCtrl1.TopText1.Width = 1000
Set RstMain = New ADODB.Recordset
RstMain.Open "Select * From AGGREGATE Order by Aggre_Name", GCn, adOpenDynamic, adLockOptimistic
Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select * FROM AGGREGATE Order by Aggre_Name", GCn, adOpenDynamic, adLockOptimistic
CtrlClckCol
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
Set DgCity.DataSource = RstHelp
FrCity.Visible = False
End Sub
Private Sub Form_Resize()
'    TopCtrl1.Width = Me.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub
Private Sub CtrlClckCol()
    txt(Aggre_Code).BackColor = CtrlBColOrg:      txt(Aggre_Code).ForeColor = CtrlFColOrg
    txt(Aggre_Name).BackColor = CtrlBColOrg:      txt(Aggre_Name).ForeColor = CtrlFColOrg
    txt(AggreEngine).BackColor = CtrlBColOrg:     txt(AggreEngine).ForeColor = CtrlFColOrg
End Sub
Private Sub Disp_Text(Enb As Boolean)
    txt(Aggre_Code).Enabled = Enb
    txt(Aggre_Name).Enabled = Enb
    txt(AggreEngine).Enabled = Enb
End Sub
Private Sub MakeBlank()
    txt(Aggre_Code) = ""
    txt(Aggre_Name) = ""
    txt(AggreEngine) = "No"
End Sub
Private Sub MoveRec()
On Error GoTo Errloop
RST_BOF_EOF RstMain
TopCtrl1.tDel = False
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    txt(Aggre_Code) = XNull(RstMain!Aggre_Code)
    txt(Aggre_Name) = XNull(RstMain!Aggre_Name)
    txt(AggreEngine) = IIf(RstMain!AggreEngine = "Y", "Yes", "No")
End If
Exit Sub
Errloop:        MsgBox err.Description
End Sub
Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
MakeBlank
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
txt(Aggre_Code).Tag = txt(Aggre_Code)
Txt_GotFocus Aggre_Code
txt(Aggre_Code).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(Aggre_Code).Enabled = False
    txt(Aggre_Name).Tag = txt(Aggre_Name)
    Txt_GotFocus Aggre_Name
    txt(Aggre_Name).SetFocus
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
If RstMain.RecordCount > 0 Then
'    If gCN.Execute("Select COUNT(*) From SubGroup Where CityCode='" & PubSiteCode + Trim(Txt(CityCode)) & "'").Fields(0)  > 0 Then
'        MsgBox "Relative Record(s) Exists Under This City," & vbCrLf & "Can not Delete this Record", vbInformation, "Delete Check"
'        Exit Sub
'    End If
    If MsgBox("Are You Sure to Delete This Record", vbYesNo, "Confirmation") = vbYes Then
        GCn.BeginTrans
        transFalg = 1
        GCn.Execute ("Delete From AGGREGATE Where Aggre_Code=" & Chk_Text(Trim(txt(Aggre_Code))))
        GCn.CommitTrans
        transFalg = 0
        RstMain.Requery
        RstHelp.Requery
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
    End If
Else
    MsgBox "There Is No Record To Delete.", vbInformation, "Information"
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
'On Error GoTo ErrorLoop
'    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
'    GSQL = "Select * From AGGREGATE Order by Aggre_Name"
'    Set SearchForm = Me
'    FAFind.IsNonFaFind = True
'    FAFind.Show vbModal
'    Exit Sub
'ErrorLoop:
'    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
'prn
    MsgBox "print"
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo Errloop
    transFlag = 0
    If IsValid(txt(Aggre_Code), "Aggregate Code") = False Then Txt_GotFocus Aggre_Code: Exit Sub
    If IsValid(txt(Aggre_Name), "Aggregate Name") = False Then Txt_GotFocus Aggre_Name: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From AGGREGATE Where Aggre_Code=" & Chk_Text(Trim(txt(Aggre_Code)))).Fields(0) > 0 Then MsgBox "Aggregate Group Code Already Exists", vbInformation, "Aggregate Group Code Validation": Txt_GotFocus Aggre_Code: txt(Aggre_Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("DELETE From AGGREGATE Where Aggre_Code=" & Chk_Text(Trim(txt(Aggre_Code))))
        GCn.Execute ("Insert Into AGGREGATE (Aggre_Code,Site_Code,Aggre_Name,AggreHelp,AggreEngine,U_Name,U_EntDt,U_AE) Values('" & Trim(txt(Aggre_Code)) & "','" & PubSiteCode & "'," & Chk_Text(txt(Aggre_Name)) & "," & Chk_Text(FilterString(txt(Aggre_Name))) & ",'" & left(txt(AggreEngine), 1) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Else
        GCn.Execute ("update AGGREGATE set Aggre_Code='" & Trim(txt(Aggre_Code)) & "',Site_Code='" & PubSiteCode & "',Aggre_Name=" & Chk_Text(txt(Aggre_Name)) & ",AggreHelp=" & Chk_Text(FilterString(txt(Aggre_Name))) & ",AggreEngine='" & left(txt(AggreEngine), 1) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "'" & " Where Aggre_Code=" & Chk_Text(Trim(txt(Aggre_Code))))
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    RstMain.Requery
    RstHelp.Requery
    RstMain.FIND ("Aggre_Code=" & Chk_Text(Trim(txt(Aggre_Code))))
    If ADDFLAG = 1 Then
        MakeBlank
        Txt_GotFocus Aggre_Code
        txt(Aggre_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrCity.Visible = False
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
        MoveRec
        CtrlClckCol
        FrCity.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub aggCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Aggre_Code >=" & Chk_Text(XNull(Trim(txt(Aggre_Code))))
End Sub
Private Sub aggNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Aggre_Name >=" & Chk_Text(XNull(txt(Aggre_Name)))
End Sub
Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Aggre_Code, Aggre_Name
            FrCity.Visible = True
            FrCity.top = txt(Index).top + txt(Index).height + 10
            FrCity.left = txt(Index).left
            FrCity.ZOrder 0
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
    RST_BOF_EOF RstHelp
    txt(Index).Tag = txt(Index)
    Txt_Click Index
    Select Case Index
        Case Aggre_Code, Aggre_Name
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    DgCity.Columns(0).width = 1000.1: DgCity.Columns(1).width = 2850.024: DgCity.Columns(2).width = 1000.1
    Select Case Index
        Case Aggre_Code
            DgCity.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Aggre_Code ASC"
            RstHelp.Bookmark = mBookMark
            aggCodeSearch
        Case Aggre_Name
            DgCity.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Aggre_Name ASC"
            RstHelp.Bookmark = mBookMark
            aggNameSearch
    End Select
    If FrCity.Visible = True Then FrCity.Visible = False
    If txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    txt(Index).ForeColor = CtrlFCol: txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
Select Case Index
    Case AggreEngine
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt_GotFocus Index: txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_Click Index
                Txt_GotFocus Index
                txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
End Select
Select Case Index
    Case Aggre_Code
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case Aggre_Name
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
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
    Case Aggre_Code
        aggCodeSearch
    Case Aggre_Name
        aggNameSearch
    Case AggreEngine
        If Len(txt(AggreEngine)) = 0 Or UCase(mID(txt(AggreEngine), 1, 1)) = "Y" Then
            txt(AggreEngine) = "Yes"
        Else
            txt(AggreEngine) = "No"
        End If
End Select
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Aggre_Code
            Set Rst = GCn.Execute("SELECT * FROM AGGREGATE WHERE Aggre_Code=" & Chk_Text(txt(Aggre_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Code Already Exists", vbInformation, "Validation": txt(Aggre_Code) = txt(Aggre_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Aggre_Code <> RstMain!Aggre_Code Then MsgBox "Code Already Exists", vbInformation, "Validation": txt(Aggre_Code) = txt(Aggre_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Aggre_Name
            Set Rst = GCn.Execute("SELECT * FROM AGGREGATE WHERE Aggre_Name=" & Chk_Text(txt(Aggre_Name)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Aggregate Group Name Already Exists", vbInformation, "Validation": txt(Aggre_Name) = txt(Aggre_Name).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Aggre_Name <> RstMain!Aggre_Name Then MsgBox "Aggregate Group Name Already Exists", vbInformation, "Validation": txt(Aggre_Name) = txt(Aggre_Name).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub
