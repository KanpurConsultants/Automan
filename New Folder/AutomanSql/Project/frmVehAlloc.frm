VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehAlloc 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "Vehicle Allocation Entry"
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   Begin VB.CommandButton CmdAppyi 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Show Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4755
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   705
      Width           =   1200
   End
   Begin VB.OptionButton OPtSel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Selected"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   2970
      TabIndex        =   2
      Top             =   697
      Width           =   1140
   End
   Begin VB.OptionButton OptAll 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   712
      Width           =   1005
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
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   6
      Top             =   3585
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   6285
      Left            =   90
      TabIndex        =   5
      Top             =   1530
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   11086
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   8
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   12701168
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   8421504
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
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
      _Band(0).Cols   =   8
   End
   Begin MSFlexGridLib.MSFlexGrid GridHlp 
      Height          =   7350
      Left            =   8805
      TabIndex        =   4
      Top             =   450
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   12965
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   16761087
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   8421504
      FocusRect       =   2
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "        |Model                            "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmVehAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Master As adodb.Recordset
Dim TAddMode As Boolean
Dim GridKey As Integer


Private Const SrNo As Byte = 0
Private Const SiteCode As Byte = 1
Private Const Model As Byte = 2
Private Const ChassisNo As Byte = 3
Private Const Allot As Byte = 4
Private Const Colours As Byte = 5
Private Const ChassisDocid  As Byte = 6
Private Const PDocid  As Byte = 7

Private Sub CmdAppyi_Click()
Dim ac_str As String
Dim I As Integer
Dim Rs As adodb.Recordset
    If OptAll.Value = False Then
       ac_str = ""
        For I = 0 To GridHlp.Rows - 1
            If GridHlp.TextMatrix(I, 0) = "ü" Then
                ac_str = ac_str + IIf(ac_str = "", "'" + GridHlp.TextMatrix(I, 1) + "'", "," + "'" + GridHlp.TextMatrix(I, 1) + "'")
'     Numeric           Ac_Str = Ac_Str + IIf(Ac_Str = "", GridHlp.TextMatrix(i, 1), "," + GridHlp.TextMatrix(i, 1))
            End If
        Next
        If ac_str = "" Then
            MsgBox "Select Model From List", vbInformation, "Massage"
            GridHlp.SetFocus
            Exit Sub
        End If
        
        Set Rs = New Recordset
        Set Rs = GCn.Execute("SELECT Veh_Stock.Pur_DocId,site.site_desc,ColMast.Col_Desc,Veh_Stock.al_name,Veh_Stock.Chassis_RctDocNo,Veh_Stock.MODEL,Veh_Stock.ChassisNo,Veh_Stock.Colour_Code " & _
        "FROM (Veh_Stock LEFT JOIN Site ON right(Veh_Stock.Pur_SiteCode,1) = Site.Site_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
                " where Veh_Stock.model in (" & ac_str & ") and (Veh_Stock.Sal_DocId Is Null or Veh_Stock.Sal_DocId ='')")
                
    Else
        Set Rs = New Recordset
        Set Rs = GCn.Execute("SELECT Veh_Stock.Pur_DocId,site.site_desc,ColMast.Col_Desc,Veh_Stock.al_name,Veh_Stock.Chassis_RctDocNo,Veh_Stock.MODEL,Veh_Stock.ChassisNo,Veh_Stock.Colour_Code " & _
        "FROM (Veh_Stock LEFT JOIN Site ON right(Veh_Stock.Pur_SiteCode,1) = Site.Site_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
                " where  (Veh_Stock.Sal_DocId ='' Or Veh_Stock.Sal_DocId Is Null) and left(Veh_Stock.Pur_DocID,1)='" & PubDivCode & "'")
    End If
    FGrid.Rows = 1
    I = 1
    If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            With FGrid
                .AddItem ""
                .TextMatrix(I, 0) = I
                .TextMatrix(I, SiteCode) = IIf(IsNull(Rs!Site_Desc), "", Rs!Site_Desc)
                .TextMatrix(I, Model) = Rs!Model
                .TextMatrix(I, Allot) = IIf(IsNull(Rs!al_name), "", Rs!al_name)
                .TextMatrix(I, ChassisNo) = Rs!ChassisNo
                .TextMatrix(I, Colours) = IIf(IsNull(Rs!Col_Desc), "", Rs!Col_Desc)
                .TextMatrix(I, ChassisDocid) = IIf(IsNull(Rs!Chassis_RctDocNo), "", Rs!Chassis_RctDocNo)
                .TextMatrix(I, PDocid) = IIf(IsNull(Rs!Pur_DocId), "", Rs!Pur_DocId)
            End With
            Rs.MoveNext
           I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else

        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
Set Rs = Nothing
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Eloop
FormKeyDown Me, KeyCode, Shift
Exit Sub
Eloop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()

On Error GoTo Eloop
Dim I As Byte
Dim Rst As adodb.Recordset: WinSetting Me: TopCtrl1.Tag = PubULabel: Ini_Grid
    GridHlp.ColAlignment(0) = flexAlignLeftCenter
    GridHlp.ColAlignment(1) = flexAlignLeftCenter
    Set Rst = New Recordset
      Dim sitecond As String
  
    Rst.CursorLocation = adUseClient
    Rst.Open "Select model from Veh_Stock where Left(Pur_DocId,1)='" & PubDivCode & "' " & sitecond & "order by model ", GCn, adOpenDynamic, adLockOptimistic
    GridHlp.Rows = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            GridHlp.AddItem "" & Chr(9) & Rst.Fields(0).Value
            Rst.MoveNext
        Loop
    End If
    Disp_Text False
    Set Rst = Nothing
    Exit Sub
Eloop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub GridHlp_Click()
    GridHlp.Col = 0
    GridHlp.CellFontName = "WINGDINGS"
    GridHlp.CellFontSize = 14
    GridHlp.TextMatrix(GridHlp.Row, 0) = IIf(GridHlp.TextMatrix(GridHlp.Row, 0) = "ü", " ", "ü")
End Sub


Private Sub GridHlp_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
    GridHlp.Col = 0
    GridHlp.CellFontName = "WINGDINGS"
    GridHlp.CellFontSize = 14
    GridHlp.TextMatrix(GridHlp.Row, 0) = IIf(GridHlp.TextMatrix(GridHlp.Row, 0) = "ü", " ", "ü")
End If

End Sub

Private Sub OptAll_Click()
Dim I As Integer
GridHlp.Enabled = False
For I = 0 To GridHlp.Rows - 1
GridHlp.TextMatrix(I, 0) = ""
Next
End Sub

Private Sub OPtSel_Click()
GridHlp.Enabled = True
End Sub

Private Sub TopCtrl1_eEdit()
' On Error GoTo eloop1
    Disp_Text True
'    MoveRec
    OptAll.SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub



Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text False
'    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
WindowsPrint
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean

'    On Error GoTo errlbl

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
mTrans = True
GCn.BeginTrans
    For I = 1 To FGrid.Rows - 1
      If FGrid.TextMatrix(I, Model) <> "" Then
            GCn.Execute "update veh_stock set al_name = '" & FGrid.TextMatrix(I, Allot) & "' where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "'"
      End If
    Next
GCn.CommitTrans

mTrans = False
    Disp_Text False
    Exit Sub
errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub


Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
     Select Case FGrid.Col
        Case Allot
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Allot
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case Allot
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub

Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub


Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case Allot
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
'    FGrid.CellForeColor = CellForeColLeave
End Sub

'******* Fuctions **********
'Private Sub BlankText()
'Dim i As Byte
'For i = 0 To txt.Count - 1
'    txt(i).Text = ""
'Next i
'End Sub

Private Sub Ini_Grid()
Dim I As Byte
With FGrid
    .left = Me.left ' + 45
    .Cols = 8
    .RowHeightMin = PubGridRowHeight

    .TextMatrix(0, 0) = "S.No."
    .ColAlignment(0) = flexAlignLeftCenter
    .ColWidth(0) = 450
    
    .TextMatrix(0, SiteCode) = " Site"
    .ColAlignment(SiteCode) = flexAlignLeftCenter
    .ColWidth(SiteCode) = 1200

    .TextMatrix(0, Model) = " Model Code"
    .ColAlignment(Model) = flexAlignLeftCenter
    .ColWidth(Model) = 1590
    
    .TextMatrix(0, ChassisNo) = "Chassis No"
    .ColAlignment(ChassisNo) = flexAlignLeftCenter
    .ColWidth(ChassisNo) = 1500
    
    .TextMatrix(0, Allot) = "Alloted To"
    .ColAlignment(Allot) = flexAlignLeftCenter
    .ColWidth(Allot) = 2205
    
    .TextMatrix(0, Colours) = "Colours"
    .ColAlignment(Colours) = flexAlignLeftCenter
    .ColWidth(Colours) = 1755
    
    .ColWidth(ChassisDocid) = 0
    .ColWidth(PDocid) = 0
End With
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
If Enb = True Then
    GridHlp.Enabled = True
    FGrid.Enabled = True
    TopCtrl1.tEdit = False
    TopCtrl1.tCancel = True
    TopCtrl1.tSave = True
    OPtSel.Enabled = True
    OptAll.Enabled = True
    CmdAppyi.Enabled = True
    TopCtrl1.TopText2.CAPTION = "Edit"
Else
    GridHlp.Enabled = False
    FGrid.Enabled = False
    TopCtrl1.tEdit = True
    TopCtrl1.tCancel = False
    TopCtrl1.tSave = False
    OPtSel.Enabled = False
    OptAll.Enabled = False
    CmdAppyi.Enabled = False
    TopCtrl1.TopText2.CAPTION = "Browse"
End If
TopCtrl1.tAdd = False
TopCtrl1.tRef = False
TopCtrl1.tDel = False
TopCtrl1.tFirst = False
TopCtrl1.tPrev = False
TopCtrl1.tNext = False
TopCtrl1.tLast = False
TopCtrl1.tFind = False
TopCtrl1.tExit = True
TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid(0).Visible = False
    FGrid.SetFocus
    Exit Sub
End If
If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
    If TxtGridLeave = True Then
         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Allot
         FGrid.Col = Allot
    End If
End If
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
End Sub

Private Function TxtGridLeave() As Boolean
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function
Private Sub WindowsPrint()
Dim Rs As adodb.Recordset ', mQry As String
Dim I As Integer, ac_str As String, mRepName As String
On Error GoTo ERRORHANDLER
    If OptAll.Value = False Then
       ac_str = ""
        For I = 0 To GridHlp.Rows - 1
            If GridHlp.TextMatrix(I, 0) = "ü" Then
                ac_str = ac_str + IIf(ac_str = "", "'" + GridHlp.TextMatrix(I, 1) + "'", "," + "'" + GridHlp.TextMatrix(I, 1) + "'")
'     Numeric           Ac_Str = Ac_Str + IIf(Ac_Str = "", GridHlp.TextMatrix(i, 1), "," + GridHlp.TextMatrix(i, 1))
            End If
        Next
        If ac_str = "" Then
            MsgBox "Select Model From List", vbInformation, "Massage"
            GridHlp.SetFocus
            Exit Sub
        End If
        
        Set Rs = New Recordset
        Set Rs = GCn.Execute("SELECT Veh_Stock.Pur_DocId,site.site_desc,ColMast.Col_Desc,Veh_Stock.al_name,Veh_Stock.Chassis_RctDocNo,Veh_Stock.MODEL,Veh_Stock.ChassisNo,Veh_Stock.Colour_Code " & _
        "FROM (Veh_Stock LEFT JOIN Site ON right(Veh_Stock.Pur_SiteCode,1) = Site.Site_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
                " where Veh_Stock.model in (" & ac_str & ") and (isnull(Veh_Stock.Sal_DocId) or Veh_Stock.Sal_DocId ='')")
                
    Else
        Set Rs = New Recordset
        Set Rs = GCn.Execute("SELECT Veh_Stock.Pur_DocId,site.site_desc,ColMast.Col_Desc,Veh_Stock.al_name,Veh_Stock.Chassis_RctDocNo,Veh_Stock.MODEL,Veh_Stock.ChassisNo,Veh_Stock.Colour_Code " & _
        "FROM (Veh_Stock LEFT JOIN Site ON right(Veh_Stock.Pur_SiteCode,1) = Site.Site_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
                " where  Veh_Stock.Sal_DocId ='' and left(Veh_Stock.Pur_DocID,1)='" & PubDivCode & "'")
                
    End If
    If Rs.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    mRepName = "VehAllocation"
    CreateFieldDefFile Rs, PubRepoPath + "\" & mRepName & ".ttx", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("compname")
                rpt.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
            Case UCase("compadd")
                rpt.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
            Case UCase("compcity")
                rpt.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
            Case UCase("RepTitle")
                rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
        End Select
    Next
    rpt.Database.SetDataSource Rs
    rpt.ReadRecords
    
    Call Report_View(rpt, Me.CAPTION, , False)
Exit Sub
ERRORHANDLER:
        CheckError
End Sub

