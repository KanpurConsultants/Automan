VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmBodyType 
   Caption         =   "Body Type Master"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   7080
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   1875
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1290
      Width           =   3840
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   6600
      TabIndex        =   1
      Top             =   2475
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   3201
         View            =   3
         Arrange         =   1
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
            Name            =   "MS Sans Serif"
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
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   1875
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1035
      Visible         =   0   'False
      Width           =   345
   End
   Begin MSDataGridLib.DataGrid DgCategory 
      Height          =   2040
      Left            =   2340
      Negotiate       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2910
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3598
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Passport Category"
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
            ColumnWidth     =   3479.811
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   661
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name...................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   180
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
      Left            =   180
      TabIndex        =   6
      Top             =   1065
      Visible         =   0   'False
      Width           =   1530
   End
End
Attribute VB_Name = "FrmBodyType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCategory As ADODB.Recordset

Dim Master As ADODB.Recordset

Private Const tCode             As Byte = 0
Private Const tName             As Byte = 1



Dim EditName        As String
Dim EditDesc        As String
Dim ListArray       As Variant
Dim mListItem       As ListItem

Private Sub ListView_Click()
On Error GoTo ELoop
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txt(Val(ListView.Tag)).SetFocus
Exit Sub
ELoop:
    CheckError
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
    
    TopCtrl1.Tag = PubUParam
    WinSetting Me, 4500, 8715
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "Select BodyTypeCode as SearchCode, I.BodyTypeDesc As Name,I.* from BodyType I  Order by BodyTypeDesc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 BodyTypeCode as SearchCode, I.BodyTypeDesc As Name,I.* from BodyType I  Order by BodyTypeDesc", GCn, adOpenDynamic, adLockOptimistic
    End If
   
    Set RsCategory = GCn.Execute("select BodyTypeCode as Code, BodyTypeDesc Name  from BodyType Order by BodyTypeDesc")
    Set DgCategory.DataSource = RsCategory
    
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
                    GCn.Execute ("delete from BodyType where BodyTypeCode= '" & Master!SearchCode & "'")
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
    GSQL = "select I.BodyTypeCode as SearchCode, I.BodyTypeDesc From BodyType I Order By I.BodyTypeDesc"
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
        Set Master = GCn.Execute("Select BodyTypeCode as SearchCode, I.BodyTypeDesc As Name,I.* from BodyType I  Where BodyTypeCode ='" & MyValue & "' Order by BodyTypeDesc")
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
    Dim mCondStr$, mSearchCode$
'   On Error GoTo errlbl
   
     
     If IsValid(txt(tName), "Objective Desc") = False Then Exit Sub
     
    If TopCtrl1.TopText2 = "Edit" Then mCondStr = " And BodyTypeDesc <> '" & Master!Name & "'"
    Set Rst = GCn.Execute("select BodyTypeDesc from BodyType where BodyTypeDesc = '" & txt(tName) & "' " & mCondStr & " ")
    If Rst.RecordCount > 0 Then
        MsgBox "Duplicate Objective Company Name", vbInformation, "Validation Check": txt(tName).SetFocus: Exit Sub
    End If
    Set Rst = Nothing

    
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        mMaxId = GCn.Execute("Select " & vIsNull("Max(" & cVal("BodyTypeCode") & ")", "0") & "+1 From BodyType").Fields(0).Value
                
        GCn.Execute ("insert into BodyType(BodyTypeCode, BodyTypeDesc, Site_Code,  U_Name, U_EntDt, U_AE) " & _
            " values('" & mMaxId & "' ,'" & txt(tName) & "', '" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
        mSearchCode = mMaxId
    Else
        GCn.Execute "update BodyType  set BodyTypeDesc='" & txt(tName) & "', U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "' Where BodyTypeCode = '" & Master!SearchCode & "'"
        mSearchCode = Master!SearchCode
    End If
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select BodyTypeCode as SearchCode, I.BodyTypeDesc As Name,I.* from BodyType I  Where BodyTypeCode ='" & mSearchCode & "' Order by BodyTypeDesc")
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
'        Case tCity
'            'DGridColSwap DgCity, 1
'            RsCity.Sort = "Name"
'            If RsCity.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
'            If RsCity.EOF = True Or RsCity.BOF = True Then Exit Sub
'            If txt(Index).TEXT <> RsCity!Name Then
'                RsCity.MoveFirst
'                RsCity.FIND "Name ='" & txt(Index).TEXT & "'"
'            End If
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
    Case tName
        DGridTxtKeyDown_Mast DgCategory, txt, Index, RsCategory, KeyCode, False, 1
End Select
If DgCategory.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> tName Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = tName Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If Index <> tCode Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
Select Case Index
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case tName
        DGridTxtKeyUp_Mast txt, Index, RsCategory, KeyCode, "Name"
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
        txt(tName) = XNull(!BodyTypeDesc)
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
    If DgCategory.Visible = True Then DgCategory.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
    Select Case Index
        
    End Select
End Sub








