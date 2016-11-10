VERSION 5.00
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmDeprecation_itemMaster 
   Caption         =   "Deprecation Item Master"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   2190
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1125
      Width           =   795
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   2190
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1365
      Width           =   810
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   2205
      MaxLength       =   1
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   6600
      TabIndex        =   4
      Top             =   2475
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   2190
      MaxLength       =   40
      TabIndex        =   1
      Top             =   885
      Width           =   3885
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name....."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1005
      TabIndex        =   9
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deprecation %......."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1005
      TabIndex        =   8
      Top             =   1395
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code.................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1005
      TabIndex        =   7
      Top             =   660
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type......................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1005
      TabIndex        =   6
      Top             =   915
      Width           =   1725
   End
End
Attribute VB_Name = "FrmDeprecation_itemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCategory As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim RsCity As ADODB.Recordset

Private Const tCode         As Byte = 0
Private Const tName         As Byte = 1
Private Const Dep_per          As Byte = 2
Private Const ShortNAme          As Byte = 3


Dim EditName        As String
Dim EditDesc        As String
Dim ListArray       As Variant
Dim mListItem       As ListItem



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub


Private Sub Form_Load()
On Error GoTo ELoop
    
    TopCtrl1.Tag = PubUParam
    WinSetting Me, 4500, 8715
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "Select code as SearchCode, I.Description As Name,I.* from Deprecation_itemMaster I Order by Description", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 Code as SearchCode, I.Description As Name,I.* from Deprecation_itemMaster I Order by Description", GCn, adOpenDynamic, adLockOptimistic
    End If
   
   Set RsCategory = GCn.Execute("Select Code, Description, Dep_per As Name From Deprecation_itemMaster Order By Description")
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
  
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsCategory = Nothing
Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(tName).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                    GCn.Execute ("delete from Deprecation_itemMaster where code= '" & Master!SearchCode & "'")
                GCn.CommitTrans
                
                Master.Requery
                Call MoveRec
                RsCategory.Requery
                
                BUTTONS True, Me, Master, 0
            End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    EditName = txt(tCode).TEXT
    EditDesc = txt(tName).TEXT
    txt(tName).SetFocus
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
    GSQL = "select I.code as SearchCode, I.Description From Deprecation_itemMaster I Order By I.Description"
    Set SearchForm = Me
    'FIND.Show vbModal
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal

    
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
        Set Master = GCn.Execute("Select code as SearchCode, I.Description As Name,I.*, C.CityName from Deprecation_itemMaster I Left Join City C On C.CityCode=I.CityCode Where code = '" & MyValue & "'  Order by Description")
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
    RsCategory.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim ItemCode As Integer
    Dim Rst As ADODB.Recordset
    Dim mMaxID As Long
    Dim mCondStr$
'   On Error GoTo errlbl
   
     
     If IsValid(txt(tName), "Objective Desc") = False Then Exit Sub
     
    If TopCtrl1.TopText2 = "Edit" Then mCondStr = " And Description <> '" & Master!Name & "'"
    Set Rst = GCn.Execute("select Description from Deprecation_itemMaster where Description = '" & txt(tName) & "' " & mCondStr & " ")
    If Rst.RecordCount > 0 Then
        MsgBox "Duplicate Name", vbInformation, "Validation Check": txt(tName).SetFocus: Exit Sub
    End If
    Set Rst = Nothing

    
    
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        mMaxID = GCn.Execute("Select " & vIsNull("Max(" & cVal("code") & ")", "0") & "+1 From Deprecation_itemMaster").Fields(0).Value
                
        GCn.Execute ("insert into Deprecation_itemMaster(code, Description,ShortName, Dep_per, Site_Code, U_Name, U_EntDt, U_AE) " & _
            " values('" & mMaxID & "' ,'" & txt(tName) & "','" & txt(ShortNAme) & "', '" & txt(Dep_per) & "', '" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
    Else
        GCn.Execute "update Deprecation_itemMaster  set Description='" & txt(tName) & "',ShortName='" & txt(ShortNAme) & "', Dep_per='" & txt(Dep_per) & "', U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "' Where Code = '" & Master!SearchCode & "'"
        mMaxID = Master!SearchCode
    End If
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
'        Set Master = GCn.Execute("Select code as SearchCode, I.Description As Name,I.*, C.CityName from Deprecation_itemMaster I Left Join City C On C.CityCode=I.CityCode Where code = '" & mMaxID & "'  Order by Description")
    End If
    RsCategory.Requery
    Master.FIND "SearchCode = '" & Master!SearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Grid_Hide
    Ctrl_GetFocus txt(Index)
    Select Case Index
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
        
        
End Select
'If DgCategory.Visible = False And FrmList.Visible = False And DgCity.Visible = False Then
'        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Contact Then Ctrl_DownKeyDown KeyCode, Shift
'        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Contact Then
'            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
'        End If
'        If Index <> tCode Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
'End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case Dep_per
        NumPress txt(Index), KeyAscii, 3, 2, True
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
    txt(I).Tag = ""
Next I
End Sub

Private Sub MoveRec()
On Error GoTo error1

With Master
    If .RecordCount > 0 Then
        txt(tCode) = !SearchCode
        txt(tName) = XNull(!Description)
           txt(ShortNAme) = XNull(!ShortNAme)
        txt(Dep_per) = XNull(!Dep_per)
    End If
End With

TopCtrl1.tPrn = False
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next
    txtDisabled_Color Me
End Sub
Private Sub Grid_Hide()
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
    Select Case Index
    End Select
End Sub












