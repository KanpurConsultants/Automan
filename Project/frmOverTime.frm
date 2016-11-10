VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmOverTime 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Over Time Entry"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11820
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
   LinkTopic       =   "form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGEmp 
      Height          =   4515
      Left            =   8070
      Negotiate       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Employee Name"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   2970
      MaxLength       =   12
      TabIndex        =   1
      Top             =   750
      Width           =   1380
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   2970
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1005
      Width           =   4785
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   2970
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1515
      Width           =   4785
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Left            =   2970
      TabIndex        =   3
      Top             =   1260
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Over Time Date*....."
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
      Index           =   26
      Left            =   1260
      TabIndex        =   8
      Top             =   750
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name*......"
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
      Left            =   1260
      TabIndex        =   7
      Top             =   1020
      Width           =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remark's"
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
      Index           =   38
      Left            =   1245
      TabIndex        =   6
      Top             =   1530
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Over Time*.............."
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
      Index           =   27
      Left            =   1260
      TabIndex        =   5
      Top             =   1275
      Width           =   1845
   End
End
Attribute VB_Name = "frmOverTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim ExitCtrl As Boolean

Dim RsEmp As ADODB.Recordset
Dim Master As ADODB.Recordset
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Dim ListArray As Variant
Dim mListItem As ListItem

Private Const OTDate As Byte = 0
Private Const Emp_Code As Byte = 1
Private Const OverTime As Byte = 2
Private Const Remarks As Byte = 3

Private Sub DGEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 If RsEmp.RecordCount > 0 Then
        txt(Emp_Code).TEXT = RsEmp!Name
        txt(Emp_Code).Tag = RsEmp!Code
    End If
    DGEmp.Visible = False
    txt(Emp_Code).SetFocus
End If
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
    TopCtrl1.Tag = PubUParam: WinSetting Me
    DGEmp.left = Me.width - (DGEmp.width + mRtScale): DGEmp.top = mTopScale

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        If PubBackEnd = "A" Then
            Master.Open "select (OT.Div_Code+ " & cCStr("OT.OT_Date") & " + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Order by OT.OT_Date,E.Emp_Name", GCn, adOpenDynamic, adLockOptimistic
        ElseIf PubBackEnd = "S" Then
            Master.Open "select (OT.Div_Code+ Convert(nVarChar,OT.OT_Date,3) + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Order by OT.OT_Date,E.Emp_Name", GCn, adOpenDynamic, adLockOptimistic
        End If
    Else
        If PubBackEnd = "A" Then
            Master.Open "select Top 1 (OT.Div_Code+ " & cCStr("OT.OT_Date") & " + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Order by OT.OT_Date,E.Emp_Name", GCn, adOpenDynamic, adLockOptimistic
        ElseIf PubBackEnd = "S" Then
            Master.Open "select Top 1 (OT.Div_Code+ Convert(nVarChar,OT.OT_Date,3) + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Order by OT.OT_Date,E.Emp_Name", GCn, adOpenDynamic, adLockOptimistic
        End If
    End If
   
    Set RsEmp = New ADODB.Recordset
    RsEmp.CursorLocation = adUseClient
    RsEmp.Open "select Emp_code as code,emp_name as name from emp_mast where Div_Code='" & PubDivCode & "' order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGEmp.DataSource = RsEmp
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    txt(OTDate).Tag = PubLoginDate
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
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
Set RsEmp = Nothing
Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
'Dim VNo As Long
'Dim i As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(OTDate) = Format(txt(OTDate).Tag, "dd/mm/yyyy")
    txt(OTDate).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim mTrans As Boolean, vBook As Variant
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        mTrans = True
        GCn.BeginTrans
        GCn.Execute ("Delete from OverTime where OT_Date=" & ConvertDate(txt(OTDate)) & " and Emp_Code = '" & txt(Emp_Code).Tag & "'")
        GCn.CommitTrans
        mTrans = False
        Master.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        Call MoveRec
    End If
eloop1:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(OverTime).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    If PubBackEnd = "A" Then
        GSQL = "select (OT.Div_Code + OT.OT_Date + OT.Emp_Code) as searchcode," & _
            " OT.OT_Date as OverTime_Date,Emp.Emp_Name " & _
            " from OverTime as OT left Join Emp_Mast Emp on OT.Emp_Code=Emp.Emp_Code" & _
            " Where Ot.Div_Code='" & PubDivCode & "' Order by OT.OT_Date,Emp.Emp_Name"
    ElseIf PubBackEnd = "S" Then
        GSQL = "select (OT.Div_Code + Convert(nVarChar,OT.OT_Date,3) + OT.Emp_Code) as searchcode," & _
            " Convert(nVarChar,OT.OT_Date,3) as OverTime_Date,Emp.Emp_Name " & _
            " from OverTime as OT left Join Emp_Mast Emp on OT.Emp_Code=Emp.Emp_Code" & _
            " Where Ot.Div_Code='" & PubDivCode & "' Order by OT.OT_Date,Emp.Emp_Name"
    End If
    
    Set SearchForm = Me
    If PubBackEnd = "A" Then
        FIND2.Show vbModal
    Else
        FAFind.Show vbModal
    End If
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        If PubBackEnd = "A" Then
            Set Master = GCn.Execute("select (OT.Div_Code+ " & cCStr("OT.OT_Date") & " + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Where (OT.Div_Code+ " & cCStr("OT.OT_Date") & " + OT.Emp_Code) = '" & MyValue & "' Order by OT.OT_Date,E.Emp_Name")
        ElseIf PubBackEnd = "S" Then
            Set Master = GCn.Execute("select (OT.Div_Code+ Convert(nVarChar,OT.OT_Date,3) + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Where (OT.Div_Code+ Convert(nVarChar,OT.OT_Date,3) + OT.Emp_Code) = '" & MyValue & "' Order by OT.OT_Date,E.Emp_Name")
        End If
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, Master, 1
  Call MoveRec
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, Master, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, Master, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, Master, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    RsEmp.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, mStr$
    Dim mTrans As Boolean
'    On Error GoTo errlbl
    
    Grid_Hide
    If IsValid(txt(OTDate), "Over Time  Date") = False Then Exit Sub
    If IsValid(txt(Emp_Code), "Employee name") = False Then Exit Sub
    If IsValid(txt(OverTime), "Over Time") = False Then Exit Sub
    If Val(txt(OverTime)) = 0 Then
        MsgBox "Over Time hour must be greater than zero"
        txt(OverTime).SetFocus
        Exit Sub
    End If
 RemoveTxtNull
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        Set GRs = New ADODB.Recordset
        GRs.CursorLocation = adUseClient
        GRs.Open "select Emp_Code from OverTime where Div_Code + " & cCStr("OT_Date") & " + Emp_Code= '" & PubDivCode & txt(OTDate) & txt(Emp_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        If GRs.RecordCount > 0 Then
            Set GRs = Nothing
            MsgBox "Over Time for Selected Employee Already Feeded!" & vbCrLf & "Try Edit !", vbCritical, "Validation"
            txt(Emp_Code).SetFocus
            Exit Sub
        End If
        Set GRs = Nothing
        GCn.Execute ("Insert into OverTime(Div_Code,OT_Date,Emp_Code,HrMinute,REMARKS,Site_Code,U_Name,U_EntDt,U_AE)" & _
            " values('" & PubDivCode & "'," & ConvertDate(txt(OTDate)) & ",'" & txt(Emp_Code).Tag & "','" & txt(OverTime) & "', '" & txt(Remarks) & "','" & PubSiteCode & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
    Else
        GCn.Execute ("update OverTime set Site_Code='" & PubSiteCode & PubSiteCode & "',Emp_Code='" & txt(Emp_Code).Tag & "',OT_Date=" & ConvertDate(txt(OTDate)) & ",HrMinute='" & txt(OverTime) & "',Remarks='" & txt(Remarks) & "',U_Name = '" & pubUName & "', U_EntDt = " & ConvertDate(PubServerDate) & ", U_AE ='E' " & _
            " where Div_Code + " & cCStr("OT_Date") & " + Emp_Code= '" & PubDivCode & txt(OTDate) & txt(Emp_Code).Tag & "'")
    End If
GCn.CommitTrans
mTrans = False

    If PubBackEnd = "A" Then
        mStr = PubDivCode & txt(OTDate) & txt(Emp_Code).Tag
    ElseIf PubBackEnd = "S" Then
        mStr = PubDivCode & Format(txt(OTDate), "DD/MM/YY") & txt(Emp_Code).Tag
    End If
    
    If PubMoveRecYn Then
        Master.Requery
    Else
        If PubBackEnd = "A" Then
            Set Master = GCn.Execute("select (OT.Div_Code+ " & cCStr("OT.OT_Date") & " + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Where (OT.Div_Code+ " & cCStr("OT.OT_Date") & " + OT.Emp_Code) = '" & mStr & "' Order by OT.OT_Date,E.Emp_Name")
        ElseIf PubBackEnd = "S" Then
            Set Master = GCn.Execute("select (OT.Div_Code+ Convert(nVarChar,OT.OT_Date,3) + OT.Emp_Code) as searchcode,E.Emp_Name from OverTime as OT Left Join Emp_Mast as E on OT.Emp_Code=E.Emp_Code Where (OT.Div_Code+ Convert(nVarChar,OT.OT_Date,3) + OT.Emp_Code) = '" & mStr & "' Order by OT.OT_Date,E.Emp_Name")
        End If
    End If
    
    Master.FIND "SearchCode = '" & mStr & "'"
    
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        txt(OTDate).Tag = Format(txt(OTDate), "dd/mm/yyyy")
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
 Grid_Hide
Ctrl_GetFocus txt(Index)
Select Case Index
    Case Emp_Code
        DGEmp.Tag = 1
        If RsEmp.RecordCount = 0 Or (RsEmp.EOF = True Or RsEmp.BOF = True) Or txt(Emp_Code).TEXT = "" Then Exit Sub
        If txt(Emp_Code).TEXT <> RsEmp!Name Then
            RsEmp.MoveFirst
            RsEmp.FIND "name ='" & txt(Emp_Code).TEXT & "'"
        End If
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    txt(Index).TEXT = ""
    Grid_Hide
    Exit Sub
End If

'Select Case Index
'    Case Emp_Code
'        DGridTxtKeyDown DGEmp, txt, Index, RsEmp, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
'End Select
If DGEmp.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Remarks Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    Else
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    End If
    If TopCtrl1.TopText2 = "Add" And Index <> OTDate Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    If TopCtrl1.TopText2 = "Edit" And Index <> Remarks Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
If Not keyascii = 13 And Index = Emp_Code Then
    DGEmp.Visible = True
    DGEmp.SetFocus
End If
Select Case Index
    Case Emp_Code
        If DGEmp.Visible = True Then DGridTxtKeyPress txt, Index, RsEmp, keyascii, "Name"
    Case OverTime
        Call NumPress(txt(Index), keyascii, 2, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
'    Case Visit_Call
'        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
'    Case Call_Status
'        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case Emp_Code
        If IsValid(txt(Emp_Code), "Representative name") = False Then: txt(Emp_Code).SetFocus: Exit Sub
        If RsEmp.RecordCount = 0 Or (RsEmp.EOF = True Or RsEmp.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsEmp!Name
            txt(Index).Tag = RsEmp!Code
        End If
    Case OTDate
        txt(Index).TEXT = RetDate(txt(Index))
    Case OverTime
        txt(Index) = Format(txt(Index), "hh:mm")
End Select
End Sub

Private Sub DGEmp_Click()
    If RsEmp.RecordCount > 0 Then
        txt(Emp_Code).TEXT = RsEmp!Name
        txt(Emp_Code).Tag = RsEmp!Code
    End If
    DGEmp.Visible = False
    txt(Emp_Code).SetFocus
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    If I <> OTDate Then txt(I).TEXT = "": txt(I).Tag = ""
Next I
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset, I As Integer
On Error GoTo error1
TopCtrl1.tPrn = False
If Master.RecordCount > 0 Then
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    If PubBackEnd = "A" Then
        Master1.Open "select * from OverTime where (Div_Code+ " & cCStr("OT_Date") & " + Emp_Code)='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
    ElseIf PubBackEnd = "S" Then
        Master1.Open "select * from OverTime where (Div_Code+ Convert(nVarChar,OT_Date,3) + Emp_Code)='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
    End If
    

    txt(OTDate) = Master1!OT_Date
    txt(OverTime) = Format(XNull(Master1!HrMinute), "hh:mm")
    txt(Remarks) = IIf(IsNull(Master1!Remarks), "", Master1!Remarks)
    txt(Emp_Code).Tag = Master1!Emp_Code
    If txt(Emp_Code).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & txt(Emp_Code).Tag & "'").RecordCount > 0 Then
        txt(Emp_Code).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & txt(Emp_Code).Tag & "'").Fields(0).Value
    Else
        txt(Emp_Code).TEXT = ""
    End If
Else
    Call BlankText
End If
Grid_Hide
Set Master1 = Nothing
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
Next
If TopCtrl1.TopText2 = "Edit" Then
    txt(OTDate).Enabled = False
    txt(Emp_Code).Enabled = False
End If
'txtDisabled_Color Me
End Sub
Private Sub Grid_Hide()
    If DGEmp.Visible = True Then DGEmp.Visible = False
End Sub

Private Sub RemoveTxtNull()
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).TEXT = IIf(IsNull(txt(I).TEXT), "", txt(I).TEXT)
Next I
End Sub
