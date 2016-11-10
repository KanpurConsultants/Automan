VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmWarrantyDispatch 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Warranty Claim Dispatch Entry"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6045
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
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   6045
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtgrid1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   6960
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   705
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   7515
      Left            =   180
      TabIndex        =   2
      Top             =   855
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   13256
      _Version        =   393216
      BackColor       =   12243913
      BackColorFixed  =   4210816
      ForeColorFixed  =   65535
      BackColorSel    =   16711680
      BackColorBkg    =   14667998
      GridColor       =   128
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmWarrantyDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TAddMode As Boolean
Dim ExitCtrl As Boolean
Dim GridKey As Integer

Dim ForSiteCode As String

Dim MyIndex As Byte
Dim Rst As ADODB.Recordset
Dim Master As ADODB.Recordset

'Text Box (Form)
Private Const PendingOnly As Byte = 1

'Fgrid1 Columns
Private Const C_DispNo As Byte = 1
Private Const C_DispDate As Byte = 2
Private Const C_Year As Byte = 3
Private Const C_ClaimType As Byte = 4
Private Const C_ClaimNo As Byte = 5
Private Const C_ClaimDt As Byte = 6
Private Const C_JobNo As Byte = 7
Private Const C_RegNo As Byte = 8
Private Const C_Chassis As Byte = 9
Private Const C_ZSRPass As Byte = 10
Private Const C_ZSRDate As Byte = 11
Private Const C_Div As Byte = 12
Private Const C_Site As Byte = 13

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
    If KeyCode <> vbKeyF10 Then
        If TopCtrl1.PrvKeyCode = vbKeyEscape Then
            TopCtrl1.PrvKeyCode = 0
        Else
            TopCtrl1.PrvKeyCode = KeyCode
        End If
    End If
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
Dim SrNo As Integer
    TopCtrl1.Tag = "*E*P"
    ForSiteCode = PubSiteCode
    Call BlankText
    Ini_Grid
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select JW1.Year_Prefix+RIGHT(SPACE(2)+JW1.Claim_Type,2)+RIGHT(SPACE(10)+JW1.Claim_No,10) AS CODE,JW1.Year_Prefix+RIGHT(SPACE(2)+JW1.Claim_Type,2)+RIGHT(SPACE(10)+JW1.Claim_No,10) AS SearchCode, JW1.*,Jc.Job_no,HC.RegNo,HC.Chassis " _
                & "FROM (JOB_WARR1 AS JW1 left join Job_Card as Jc on Jw1.Job_DocId=Jc.DocId) left join Hiscard as HC on JC.Cardno=Hc.Cardno  where left(JW1.site_code,1)='" & PubSiteCode & "' AND left(JW1.site_code,1)='" & PubSiteCode & "' order by JW1.CLAIM_TYPE,JW1.CLAIM_NO", GCn, adOpenDynamic, adLockOptimistic
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
    TopCtrl1.tFirst = True
    TopCtrl1.tLast = True
    TopCtrl1.tPrev = True
    TopCtrl1.tNext = True
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMaximized Then
        Me.left = MDIForm1.left
    End If
    Ini_Grid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
        Disp_Text SETS("ADD", Me, Master)
        FGrid1.SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
    Exit Sub
End Sub

Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    FGrid1.SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    If TopCtrl1.TopText2 = "Browse" Then Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = Master.Source
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
Dim I As Integer
    For I = 1 To FGrid1.Rows - 1
        If MyValue = FGrid1.TextMatrix(I, C_Year) + Right(Space(2) + FGrid1.TextMatrix(I, C_ClaimType), 2) + Right(Space(10) + FGrid1.TextMatrix(I, C_ClaimNo), 10) Then
            FGrid1.Row = I
            FGrid1.SetFocus
            Exit For
        End If
    Next
    Exit Sub
End Sub

Private Sub TopCtrl1_eFirst()
    Call moveFGrid(1)
End Sub

Private Sub TopCtrl1_eLast()
    Call moveFGrid(4)
End Sub

Private Sub TopCtrl1_eNext()
    Call moveFGrid(3)
End Sub

Private Sub TopCtrl1_ePrev()
    Call moveFGrid(2)
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
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
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim AddFlg As String
    On Error GoTo errlbl
    
    If TxtGrid1(0).Visible = True Then
        If TxtGrid1Leave = False Then
            TxtGrid1_LostFocus 0
            TxtGrid1(0).SetFocus
            Exit Sub
        End If
    End If
    
    Grid_Hide
    GCn.BeginTrans
    mTrans = True

    For I = 1 To FGrid1.Rows - 1
        GSQL = "Update Job_WARR1 set DispatchNo='" & FGrid1.TextMatrix(I, C_DispNo) & "',DispatchDate=" & ConvertDate(FGrid1.TextMatrix(I, C_DispDate)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='A' where Div_code='" & FGrid1.TextMatrix(I, C_Div) & "' and Site_code='" & FGrid1.TextMatrix(I, C_Site) & "' and Claim_Type='" & FGrid1.TextMatrix(I, C_ClaimType) & "' and claim_no='" & FGrid1.TextMatrix(I, C_ClaimNo) & "' and Year_Prefix='" & FGrid1.TextMatrix(I, C_Year) & "'"
        GCn.Execute GSQL
    Next I

    GCn.CommitTrans
    mTrans = False

    Master.Requery
    Call UpdRequery
    
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


'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
End Sub

Private Sub MoveRec()
Dim rs As Recordset
Dim mVor As String
Dim I As Integer
On Error GoTo error1
    If Master.RecordCount > 0 Then
        Call Fill_Grid
    Else
        Call BlankText
    End If
    Grid_Hide
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    
    With FGrid1
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 14
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        
        .TextMatrix(0, C_DispNo) = "Dispatch No."
        .ColAlignment(C_DispNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_DispNo) = flexAlignLeftCenter
        .ColWidth(C_DispNo) = 1000
        
        .TextMatrix(0, C_DispDate) = "Dispatch Date"
        .ColAlignment(C_DispDate) = flexAlignLeftCenter
        .ColAlignmentFixed(C_DispDate) = flexAlignLeftCenter
        .ColWidth(C_DispDate) = 1200
        
        .TextMatrix(0, C_Year) = "Year"
        .ColAlignment(C_Year) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Year) = flexAlignLeftCenter
        .ColWidth(C_Year) = 500
        
        .TextMatrix(0, C_ClaimType) = "Type"
        .ColAlignment(C_ClaimType) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ClaimType) = flexAlignLeftCenter
        .ColWidth(C_ClaimType) = 500
        
        .TextMatrix(0, C_ClaimNo) = "Claim No"
        .ColAlignment(C_ClaimNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ClaimNo) = flexAlignLeftCenter
        .ColWidth(C_ClaimNo) = 1000
        
        
        .TextMatrix(0, C_ClaimDt) = "Claim Date"
        .ColAlignment(C_ClaimDt) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ClaimDt) = flexAlignLeftCenter
        .ColWidth(C_ClaimDt) = 1200
        
        .TextMatrix(0, C_JobNo) = "JobCard No"
        .ColAlignment(C_JobNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_JobNo) = flexAlignLeftCenter
        .ColWidth(C_JobNo) = 1000
        
        .TextMatrix(0, C_RegNo) = "Reg No"
        .ColAlignment(C_RegNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_RegNo) = flexAlignLeftCenter
        .ColWidth(C_RegNo) = 1200
        
        .TextMatrix(0, C_Chassis) = "Chassis No"
        .ColAlignment(C_Chassis) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Chassis) = flexAlignLeftCenter
        .ColWidth(C_Chassis) = 1350
        
        .TextMatrix(0, C_ZSRPass) = "ZSR Pass"
        .ColAlignment(C_ZSRPass) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ZSRPass) = flexAlignLeftCenter
        .ColWidth(C_ZSRPass) = 850
        
        .TextMatrix(0, C_ZSRDate) = "ZSR Date"
        .ColAlignment(C_ZSRDate) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ZSRDate) = flexAlignLeftCenter
        .ColWidth(C_ZSRDate) = 1200
        
        .TextMatrix(0, C_Div) = "Div"
        .ColAlignment(C_Div) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Div) = flexAlignLeftCenter
        .ColWidth(C_Div) = 0
        
        .TextMatrix(0, C_Site) = "Site"
        .ColAlignment(C_Site) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Site) = flexAlignLeftCenter
        .ColWidth(C_Site) = 0
    End With
    
    FGrid1.width = Me.width - 40: FGrid1.left = 20
    FGrid1.height = 6965: FGrid1.top = 575
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer

End Sub

Private Sub Grid_Hide()
    ''
End Sub

Private Sub UpdRequery()
    ''
End Sub

Private Sub Fill_Grid()
Dim MyRst As ADODB.Recordset
Dim I As Integer
    FGrid1.Rows = 1
    I = 1
    If Master.RecordCount > 0 Then
        Master.MoveFirst
        Do Until Master.EOF
            With FGrid1
                .AddItem ""
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_DispNo) = XNull(Master!DispatchNo)
                If IsNull(Master!DispatchDate) Then
                    .TextMatrix(I, C_DispDate) = ""
                Else
                    .TextMatrix(I, C_DispDate) = Format(Master!DispatchDate, "dd/MMM/yyyy")
                End If
                .TextMatrix(I, C_Year) = Master!Year_Prefix
                .TextMatrix(I, C_ClaimType) = Master!claim_type
                .TextMatrix(I, C_ClaimNo) = Master!claim_no
                .TextMatrix(I, C_ClaimDt) = Format(Master!Claim_Date, "dd/MMM/yyyy")
                .TextMatrix(I, C_JobNo) = Master!Job_No
                .TextMatrix(I, C_RegNo) = XNull(Master!RegNo)
                .TextMatrix(I, C_Chassis) = Master!Chassis
                .TextMatrix(I, C_ZSRPass) = IIf(Master!zsr_Pass = 0, "No", "Yes")
                .TextMatrix(I, C_ZSRDate) = Format(Master!Zsr_Date, "dd/MMM/yyyy")
                .TextMatrix(I, C_Div) = Master!Div_Code
                .TextMatrix(I, C_Site) = Master!Site_Code
            End With
            I = I + 1
            Master.MoveNext
        Loop
        
        FGrid1.FixedRows = 1
    Else
        FGrid1.Rows = FGrid1.Rows
        FGrid1.AddItem FGrid1.Rows
        FGrid1.FixedRows = 1
    End If
End Sub

Private Sub FGrid1_Click()
    TxtGrid1(0).Visible = False
End Sub

Private Sub FGrid1_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid1.Col
        Case C_DispNo, C_DispDate
            GridDblClick Me, FGrid1, TxtGrid1, 0
    End Select
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_EnterCell()
    FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.CellBackColor = CellBackColEnter
    TxtGrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
        FGrid1.CellBackColor = CellBackColLeave
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
        If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
        FGrid1.SetFocus
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    Select Case FGrid1.Col
        Case C_DispNo, C_DispDate
            If KeyCode = vbKeyDelete And Shift = 0 Then
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            End If
    End Select
    If KeyCode = vbKeyReturn Then
        Select Case FGrid1.Col
            Case C_DispNo, C_DispDate
                GridDblClick Me, FGrid1, TxtGrid1, 0
        End Select
        TAddMode = False
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    Select Case FGrid1.Col
        Case C_DispNo, C_DispDate
            Get_Text Me, FGrid1, TxtGrid1, 0, False, KeyAscii
        Case Else
            FGrid1_LeaveCell
            If FGrid1.Col = C_ZSRDate Then
                If FGrid1.Rows > FGrid1.Row + 1 Then
                    FGrid1.Row = FGrid1.Row + 1
                End If
                FGrid1.Col = 1
            Else
                FGrid1.Col = FGrid1.Col + 1
            End If
            FGrid1_EnterCell
            FGrid1.SetFocus
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid1.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid1.Row >= 1 Then
            If MsgBox("Are You Sure To Delete Entry ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid1.Rows > 2 Then
                    FGrid1.RemoveItem (FGrid1.Row)
                Else
                    FGrid1.Rows = 1
                    FGrid1.AddItem FGrid1.Rows
                    FGrid1.FixedRows = 1
                End If
            End If
            For I = 1 To FGrid1.Rows - 1
                FGrid1.TextMatrix(I, 0) = I
            Next
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid1.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_LeaveCell()
    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid1_Scroll()
    TxtGrid1(0).Visible = False
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
On Error GoTo ELoop
If ExitCtrl = False Then Exit Sub
    Ctrl_GetFocus TxtGrid1(0)
    Grid_Hide
    FGrid1.CellBackColor = CellBackColLeave
    TxtGrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    Select Case FGrid1.Col
        Case C_DispNo
            TxtGrid1(0).MaxLength = 10
        Case C_DispDate
            TxtGrid1(0).MaxLength = 15
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid1(0).TEXT = TxtGrid1(0).Tag
        TxtGrid1_KeyUp Index, KeyCode, Shift
        TxtGrid1(0).Visible = False
        FGrid1.SetFocus
        Exit Sub
    End If
    Select Case FGrid1.Col
        Case C_DispNo, C_DispDate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                    FGrid1_LeaveCell
                    If FGrid1.Col = C_DispDate And KeyCode = vbKeyReturn Then
                        If FGrid1.Rows > FGrid1.Row + 1 Then
                            FGrid1.Row = FGrid1.Row + 1
                        End If
                        FGrid1.Col = C_DispNo
                    Else
                        GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, C_DispDate
                        'FGrid1.Col = FGrid1.Col + 1
                    End If
                    FGrid1_EnterCell
                Else
                    TxtGrid1_LostFocus 0
                    TxtGrid1(0).SetFocus
                End If
            End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    CheckQuote KeyAscii
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_validate TxtGrid1(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer
On Error GoTo ELoop
    Select Case FGrid1.Col
        Case C_DispNo
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid1(Index).TEXT
        Case C_DispDate
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = RetDate(TxtGrid1(Index))
    End Select
    TxtGrid1(0).MaxLength = 15
    Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave() As Boolean
Dim I As Integer
    Select Case FGrid1.Col
        Case C_DispNo
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid1(0).TEXT
        Case C_DispDate
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = RetDate(TxtGrid1(0))
    End Select
    TxtGrid1(0).Visible = False
    ExitCtrl = True
    TxtGrid1Leave = True
    FGrid1.SetFocus
End Function


Private Sub moveFGrid(ByVal MyAction As Byte)
    Select Case MyAction
        Case 1
            FGrid1.Row = 1
        Case 2
            If FGrid1.Row <> 1 Then
                FGrid1.Row = FGrid1.Row - 1
            End If
        Case 3
            If FGrid1.Row <> FGrid1.Rows - 1 Then
                FGrid1.Row = FGrid1.Row + 1
            End If
        Case 4
            FGrid1.Row = FGrid1.Rows - 1
    End Select
    FGrid1_EnterCell
    FGrid1.SetFocus
End Sub
