VERSION 5.00
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmRateType 
   Caption         =   "Rate Type Master"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   7995
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   2550
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1290
      Width           =   3885
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
      Index           =   0
      Left            =   2565
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1035
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   2550
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1545
      Width           =   825
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   661
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name...................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1365
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code.................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1365
      TabIndex        =   6
      Top             =   1065
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variation %......."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1365
      TabIndex        =   5
      Top             =   1575
      Width           =   1425
   End
End
Attribute VB_Name = "FrmRateType"
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
Private Const VariationPer          As Byte = 2



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
        Master.Open "Select code as SearchCode, I.Description As Name,I.* from RateType I Order by Description", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 Code as SearchCode, I.Description As Name,I.* from RateType I Order by Description", GCn, adOpenDynamic, adLockOptimistic
    End If
   
   Set RsCategory = GCn.Execute("Select Code, Description, VariationPer As Name From RateType Order By Description")
    
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
    Txt(tName).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                    GCn.Execute ("delete from RateType where code= '" & Master!SearchCode & "'")
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
    EditName = Txt(tCode).TEXT
    EditDesc = Txt(tName).TEXT
    Txt(tName).SetFocus
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
    GSQL = "select I.code as SearchCode, I.Description From RateType I Order By I.Description"
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
        Set Master = GCn.Execute("Select code as SearchCode, I.Description As Name,I.*, C.CityName from RateType I Left Join City C On C.CityCode=I.CityCode Where code = '" & MyValue & "'  Order by Description")
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
    Dim mMaxId As Long
    Dim mCondStr$
'   On Error GoTo errlbl
   
     
     If IsValid(Txt(tName), "Objective Desc") = False Then Exit Sub
     
    If TopCtrl1.TopText2 = "Edit" Then mCondStr = " And Description <> '" & Master!Name & "'"
    Set Rst = GCn.Execute("select Description from RateType where Description = '" & Txt(tName) & "' " & mCondStr & " ")
    If Rst.RecordCount > 0 Then
        MsgBox "Duplicate Name", vbInformation, "Validation Check": Txt(tName).SetFocus: Exit Sub
    End If
    Set Rst = Nothing

    
    
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        mMaxId = GCn.Execute("Select " & vIsNull("Max(" & cVal("code") & ")", "0") & "+1 From RateType").Fields(0).Value
                
        GCn.Execute ("insert into RateType(code, Description, VariationPer, Site_Code, U_Name, U_EntDt, U_AE) " & _
            " values('" & mMaxId & "' ,'" & Txt(tName) & "', '" & Txt(VariationPer) & "', '" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
    Else
        GCn.Execute "update RateType  set Description='" & Txt(tName) & "', VariationPer='" & Txt(VariationPer) & "', U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "' Where Code = '" & Master!SearchCode & "'"
        mMaxId = Master!SearchCode
    End If
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select code as SearchCode, I.Description As Name,I.*, C.CityName from RateType I Left Join City C On C.CityCode=I.CityCode Where code = '" & mMaxId & "'  Order by Description")
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
    Ctrl_GetFocus Txt(Index)
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

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
Select Case Index
    Case VariationPer
        NumPress Txt(Index), keyascii, 3, 2, True
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
    Txt(I).Tag = ""
Next I
End Sub

Private Sub MoveRec()
On Error GoTo error1

With Master
    If .RecordCount > 0 Then
        Txt(tCode) = !SearchCode
        Txt(tName) = XNull(!Description)
        Txt(VariationPer) = XNull(!VariationPer)
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
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    Txt(I).ForeColor = CtrlFColOrg
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










