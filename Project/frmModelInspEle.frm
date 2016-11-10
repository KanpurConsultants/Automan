VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmModelInspEle 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Model Check Sheet Item Master"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11460
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3945
      Left            =   825
      TabIndex        =   2
      Top             =   735
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   6959
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   8421504
      FocusRect       =   0
      Appearance      =   0
      FormatString    =   "    |Code          |Description   |Default value           |Print index| "
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   661
      tAdd            =   0   'False
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   3735
      MaxLength       =   25
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmModelInspEle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstLab As ADODB.Recordset
Dim GridKey As Integer
Dim ExitCtrl As Boolean
'Item_Code
'Item_Description
'Default_Value
'Report_Index
'AddEdit
Private Const ItemCode As Byte = 1
Private Const Description As Byte = 2
Private Const DefVal As Byte = 3
Private Const PIndex As Byte = 4
Private Const AddEdit As Byte = 5
Dim TAddMode As Boolean
'Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Sub FGrid_Click()
txt(0).Visible = False
SetMaxLength
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
FGrid.BackColorSel = BackColorSelEnter
FGrid.ForeColorSel = ForeColorSelEnter
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
Dim result As Boolean
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case Description, DefVal
            Call Get_Text(Me, FGrid, txt, 0, False, 48)
        Case PIndex
            Call Get_Text(Me, FGrid, txt, 0, True, 48)
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Description
            Call GridDblClick(Me, FGrid, txt, 0)
            TAddMode = False
        Case DefVal, PIndex
            If FGrid.TextMatrix(FGrid.Row, Description) <> "" Then
                Call GridDblClick(Me, FGrid, txt, 0)
                TAddMode = False
            End If
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
SetMaxLength
Select Case FGrid.Col
    Case Description
        Call Get_Text(Me, FGrid, txt, 0, False, KeyAscii)
    Case DefVal
        If FGrid.TextMatrix(FGrid.Row, Description) <> "" Then
            Call Get_Text(Me, FGrid, txt, 0, False, KeyAscii)
        End If
    Case PIndex
        If FGrid.TextMatrix(FGrid.Row, Description) <> "" Then
           Call Get_Text(Me, FGrid, txt, 0, True, KeyAscii)
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = 46 And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        MsgBox "Check in Table"
'            If FGrid.Rows  > 2 Then
'                FGrid.RemoveItem (FGrid.Row)
'            Else
'                FGrid.Rows = 1
'                FGrid.AddItem ""
'                FGrid.FixedRows = 1
'            End If
         End If
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
    txt(0).Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
WinSetting Me: Ini_Grid
TopCtrl1.Tag = PubUParam
MoveRec
ADDFLAG = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
 Set RstHelp = Nothing
End Sub

Private Sub TopCtrl1_eDel()
'On Error GoTo Errloop
'Dim Res As Integer
'Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
'        If Res = 6 Then
'            GCn.BeginTrans
'            GCn.Execute ("DELETE  inspection_element.*  from (inspection_element left join inspection_catg on inspection_element.inspection_catg=inspection_catg.insp_code) where inspection_catg.print_on ='" & left(Txt(POn), 1) & "'" & " AND inspection_catg.SITE_CODE='" & PubSiteCode & "'")
'            GCn.CommitTrans
'            FGrid.Rows = 1
'            FGrid.AddItem ""
'            FGrid.FixedRows = 1
'        End If
'Exit Sub
'ErrLoop:    GCn.RollbackTrans
'            MsgBox err.Description, vbExclamation, " Deletion Error "
End Sub

Private Sub TopCtrl1_eEdit()
Dim rs As Recordset
On Error GoTo Errloop
        TopCtrl1.TopText2 = "Edit"
        TopCtrl1.TopText2.ForeColor = RGB(255, 0, 0)
        TopCtrl1.tDel = False
        TopCtrl1.tEdit = False
        TopCtrl1.tSave = True
        TopCtrl1.tCancel = True
        TopCtrl1.tFirst = False
        TopCtrl1.tNext = False
        TopCtrl1.tLast = False
        TopCtrl1.tPrev = False

    ADDFLAG = 2
    FGrid.AddItem "" & FGrid.Rows
    FGrid.SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_ePrn()
Dim i As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "ModelCheckSheet"
    mQRY = "select * from ModelCheckListMast Order by Item_Description,Report_Index"
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , False)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, mNewCode As Integer
Dim i As Byte
On Error GoTo Errloop
    
    If GCn.Execute("select max(Item_Code) from ModelCheckListMast").RecordCount > 0 Then
        mNewCode = GCn.Execute("select " & cVal("max(Item_Code)") & " from ModelCheckListMast").Fields(0).Value
    End If
    
    GCn.BeginTrans
    mTrans = True
    For i = 1 To FGrid.Rows - 1
        'Auto Code generation for new records
        If Len(FGrid.TextMatrix(i, Description)) <> 0 And FGrid.TextMatrix(i, AddEdit) = "" Then
            mNewCode = mNewCode + 1
            FGrid.TextMatrix(i, ItemCode) = Right("0000" & mNewCode, 4)
        End If
    Next
    
    For i = 1 To FGrid.Rows - 1
        If Len(FGrid.TextMatrix(i, ItemCode)) <> 0 Then
            If FGrid.TextMatrix(i, AddEdit) = "D" Then 'Delete
                GCn.Execute ("Delete from ModelCheckListMast where Item_Code='" & FGrid.TextMatrix(i, ItemCode) & "'")
            ElseIf FGrid.TextMatrix(i, AddEdit) = "" Then  'Add
                GCn.Execute ("Insert Into ModelCheckListMast (Site_Code,Item_Code,Item_Description,Default_Value,Report_Index,U_Name,U_EntDt,U_AE) Values('" & PubSiteCode & "','" & FGrid.TextMatrix(i, ItemCode) & "','" & FGrid.TextMatrix(i, Description) & "','" & FGrid.TextMatrix(i, DefVal) & "'," & Val(FGrid.TextMatrix(i, PIndex)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
            ElseIf FGrid.TextMatrix(i, AddEdit) = "E" Then 'Edit
                GCn.Execute ("Update ModelCheckListMast set Item_Description='" & FGrid.TextMatrix(i, Description) & "',Default_Value='" & FGrid.TextMatrix(i, DefVal) & "',Report_Index=" & Val(FGrid.TextMatrix(i, PIndex)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where Item_Code='" & FGrid.TextMatrix(i, ItemCode) & "'")
            End If
        End If
    Next
    GCn.CommitTrans
    mTrans = False
    TopCtrl1.TopText2 = "Browse"
    TopCtrl1.TopText2.ForeColor = RGB(0, 0, 0)
    MoveRec
    txt(0).Visible = False
'    FGrid.CellBackColor = CellBackColLeave
Exit Sub
Errloop:    If mTrans Then GCn.RollbackTrans: CheckError
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        MoveRec
        TopCtrl1.TopText2 = "Browse"
        TopCtrl1.TopText2.ForeColor = RGB(0, 0, 0)
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

Private Sub MoveRec()
Dim rs As Recordset, i As Integer
On Error GoTo Errloop
'Item_Code,Item_Description,Default_Value,Report_Index
'U_Name
'U_EntDt
'U_AE
'Trf_Date
'Site_Code
GSQL = "select * from ModelCheckListMast Order by Report_Index,Item_Description"
Set rs = New Recordset
Set rs = GCn.Execute(GSQL)
    If rs.RecordCount > 0 Then
        FGrid.Rows = 1
        i = 1
        Do Until rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(i, 0) = i
                .TextMatrix(i, ItemCode) = IIf(IsNull(rs!Item_Code), "", rs!Item_Code)
                .TextMatrix(i, Description) = IIf(IsNull(rs!Item_Description), "", rs!Item_Description)
                .TextMatrix(i, DefVal) = IIf(IsNull(rs!Default_Value), "", rs!Default_Value)
                .TextMatrix(i, PIndex) = IIf(IsNull(rs!Report_Index), "", rs!Report_Index)
                .TextMatrix(i, AddEdit) = "N"
            End With
            rs.MoveNext
            i = i + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.Rows = 1
        FGrid.AddItem ""
        FGrid.FixedRows = 1
    End If
TopCtrl1.tNext = False: TopCtrl1.tLast = False
TopCtrl1.tPrev = False: TopCtrl1.tFirst = False
TopCtrl1.tFind = False
TopCtrl1.tDel = False
TopCtrl1.tRef = False

TopCtrl1.tEdit = True
TopCtrl1.tSave = False
TopCtrl1.tCancel = False
Exit Sub
Errloop:        MsgBox err.Description
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus txt(Index)
txt(Index).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Dim result As Boolean
'Dim i As Byte
If KeyCode = vbKeyEscape Then txt(0) = txt(0).Tag: Exit Sub
Select Case FGrid.Col
    Case Description, DefVal
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
             If TxtGridLeave = True Then
                  GridTxtDown FGrid, txt, Index, KeyCode, TAddMode, PIndex      ', 3
             End If
         End If
    Case PIndex
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
             If TxtGridLeave = True Then
                  GridTxtDown FGrid, txt, Index, KeyCode, TAddMode, PIndex, , Description
             End If
         End If
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case Index
    Case 0
        Select Case FGrid.Col
            Case PIndex
                NumPress txt(0), KeyAscii, 2, 0
        End Select
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case Description
'        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(Index).Text
    Case DefVal
'        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(Index).Text
    Case PIndex
'        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Val(Txt(Index).Text)
End Select

If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    txt(0).Visible = False
End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim i As Integer, tst As String, coun As Integer
Select Case FGrid.Col
    Case Description
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        If FGrid.TextMatrix(FGrid.Row, FGrid.Col) <> txt(0) Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txt(0): FGridAddEditDel
        End If
        If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
    Case DefVal, PIndex
        If FGrid.TextMatrix(FGrid.Row, FGrid.Col) <> txt(0) Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txt(0): FGridAddEditDel
        End If
        If FGrid.TextMatrix(FGrid.Rows - 1, Description) <> "" Then FGrid.AddItem FGrid.Rows
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    txt(0).Visible = False
End If
End Function

Private Sub BlankText()
Dim i As Byte
    txt(0).TEXT = ""
    FGrid.Rows = 1
    FGrid.AddItem ""
    FGrid.FixedRows = 1
End Sub

Private Sub Ini_Grid()
Dim i As Byte
    With FGrid
        .width = 8100 'Me.width - 120
        .left = (Me.width - FGrid.width) / 2
        .RowHeightMin = PubGridRowHeight '220
        .height = .RowHeight(0) * 22
'        .top = 1575
        .Cols = 6
        
        .ColAlignmentFixed = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignRightCenter
        .ColWidth(0) = 510
        
        .TextMatrix(0, ItemCode) = "Code"
        .ColAlignment(ItemCode) = flexAlignLeftCenter
        .ColWidth(ItemCode) = 600

        .TextMatrix(0, Description) = "Item Description"
        .ColAlignment(Description) = flexAlignLeftCenter
        .ColWidth(Description) = 3400
        
        .TextMatrix(0, DefVal) = "Default Value"
        .ColAlignment(DefVal) = flexAlignLeftCenter
        .ColWidth(DefVal) = 1785
        
        .TextMatrix(0, PIndex) = "Print Index"
        .ColAlignment(PIndex) = flexAlignRightCenter
        .ColWidth(PIndex) = 975
        .TextMatrix(0, AddEdit) = "Add/Edit"
        .ColWidth(AddEdit) = 300
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
End Sub

Private Function ChkDuplicate() As Boolean
Dim i As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte
Select Case FGrid.Col
    Case Description
        Col1 = ItemCode
        Col2 = Description
    End Select
    X = UCase(Trim(txt(0).TEXT))
    For i = 1 To FGrid.Rows - 1
        If i = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(i, FGrid.Col))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Description Not Allowed", vbInformation, "Validation"
            txt(0).SetFocus
            Ctrl_GetFocus txt(0)
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Sub SetMaxLength()
Select Case FGrid.Col   'Index
    Case Description
        txt(0).MaxLength = 25
    Case DefVal
        txt(0).MaxLength = 10
    Case PIndex
        txt(0).MaxLength = 2
    Case Else
        txt(0).MaxLength = 0
End Select
End Sub

Private Sub FGridAddEditDel()
If FGrid.TextMatrix(FGrid.Row, AddEdit) <> "" Then FGrid.TextMatrix(FGrid.Row, AddEdit) = "E"
End Sub

