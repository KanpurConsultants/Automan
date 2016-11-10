VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmBudgetExp 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Expences Budgeting"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11325
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   4
      Top             =   930
      Width           =   5205
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   8490
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   75
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         Sorted          =   -1  'True
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
            Name            =   "Verdana"
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
   Begin VB.TextBox txtgrid1 
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
      Left            =   5010
      MaxLength       =   40
      TabIndex        =   0
      Top             =   1950
      Visible         =   0   'False
      Width           =   1170
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1695
      Left            =   975
      TabIndex        =   5
      Top             =   1290
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   0
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "ddd"
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSDataGridLib.DataGrid DGHelp 
      Height          =   2730
      Left            =   1215
      Negotiate       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5535
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      RowDividerStyle =   1
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
         DataField       =   "name"
         Caption         =   "Expence A/c Name"
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
            ColumnWidth     =   3420.284
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMonth 
      Height          =   2730
      Left            =   6285
      Negotiate       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5730
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      RowDividerStyle =   1
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
         DataField       =   "name"
         Caption         =   "Month Name"
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
            ColumnWidth     =   3420.284
         EndProperty
      EndProperty
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expencence A/c "
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
      Left            =   945
      TabIndex        =   6
      Top             =   960
      Width           =   1425
   End
End
Attribute VB_Name = "FrmBudgetExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset
Dim RsHelp As ADODB.Recordset
Dim RsMonth As ADODB.Recordset

Dim mFlag As Byte
Dim GridKey As Integer
Dim RsTrb As ADODB.Recordset


Private Const Col_Month_Desc    As Byte = 1
Private Const Col_Amount        As Byte = 2
Private Const Col_Date          As Byte = 3
Private Const Col_Month         As Byte = 4


Private Const ExpenceAc As Byte = 0


Private Sub FGrid1_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    txtgrid1(0).Visible = False
End Sub

Private Sub FGrid1_DblClick()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid1.Col
        Case Col_Month_Desc, Col_Amount
            Call GridDblClick(Me, FGrid1, txtgrid1, 0)
    End Select
End Sub

Private Sub FGrid1_EnterCell()
'FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = FaBackColorSelEnter

    'FGrid1.Col = Col_Month_Desc
    txtgrid1(0).Visible = False
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
    If MsgBox("Do You Want to Save?", vbYesNo) = vbYes Then TopCtrl1_eSave
'    SendKeysA vbKeyTab, True
'    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid1.Col
        Case Col_Amount
            FGrid1 = ""
        Case Col_Month_Desc
            FGrid1 = ""
            FGrid1.TextMatrix(Col_Month, FGrid1.Row) = ""
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid1.Col
        Case Col_Month, Col_Amount
            Call GridDblClick(Me, FGrid1, txtgrid1, 0)
            
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid1_KeyPress(keyascii As Integer)
Select Case FGrid1.Col
    Case Col_Month_Desc
       Call Get_Text(Me, FGrid1, txtgrid1, 0, False, keyascii)
    Case Col_Amount
        Call Get_Text(Me, FGrid1, txtgrid1, 0, True, keyascii)
    Case Col_Date
       Call Get_Text(Me, FGrid1, txtgrid1, 0, False, keyascii)
End Select

End Sub

Private Sub FGrid1_LostFocus()
FGrid1.BackColorSel = FaCellBackColLeave1

FGrid1_Validate (True)
End Sub

Private Sub FGrid1_Scroll()
    txtgrid1(0).Visible = False

End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid1.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid1.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
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
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
'    FGrid1.CellBackColor = CellBackColLeave
End Sub

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
Me.width = 9345: Me.height = 6870
TopCtrl1.Tag = PubUParam
Ini_Grid

Set RstMain = New ADODB.Recordset
RstMain.Open "Select Distinct BE.ExpAc as SearchCode, S.Name From Budget_Exp BE Left Join SubGroup S On S.SubCode = BE.ExpAc Where BE.Site_Code='" & PubSiteCode & "' Order by S.Name", GCn, adOpenDynamic, adLockOptimistic


Set RsHelp = GCn.Execute("Select SubCode As Code, Name FROM SubGroup Where Nature In ('Expenses') Order by Name")
Set DGHelp.DataSource = RsHelp


Set RsMonth = GCn.Execute("Select Code As Code, Name From Chas_Mth Order By Name")
Set DGMonth.DataSource = RsMonth

Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0:    mFlag = 0




Grid_Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set RstMain = Nothing: Set RsHelp = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)

ADDFLAG = 1
Txt(ExpenceAc).SetFocus



FGrid1.Rows = 1
FGrid1.AddItem ""
FGrid1.FixedRows = 1

Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(ExpenceAc).SetFocus
    ADDFLAG = 2
    
    FGrid1.AddItem ""
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
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        If MsgBox("Sure To Delete Record", vbYesNo) = vbYes Then
            GCn.Execute "Delete From Budget_Exp Where ExpAc = '" & RstMain!SearchCode & "' And Site_Code = '" & PubSiteCode & "'"
        End If
        RstMain.Requery
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If

Exit Sub
Errloop:
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
On Error GoTo ErrorLoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select ExpAc as SearchCode, S.Name As Expence_Account_Name from Budget_Exp B Left Join Subgroup S On B.ExpAc = S.SubCode Where B.Site_Code = '" & PubSiteCode & "' Order By S.Name"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    RstMain.MoveFirst
    RstMain.FIND ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "Budget_Exp"
    mQRY = "Select BE.ExpAc, BE.Month, Sum(BE.Amount) As Amount, Max(S.Name) As ExpAcName, Max(M.Name) As Month_Desc " & _
         "from Budget_Exp BE Left Join SubGroup S On S.SubCode = BE.ExpAc " & _
         "Left Join Chas_Mth M On M.Code = BE.Month Where BE.Site_Code = '" & PubSiteCode & "' " & _
         "Group By BE.ExpAc, BE.Month Order By ExpAcName, BE.Month "
    
    
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
Dim mTrans As Byte
Dim mSearchCode$
Dim I As Integer
On Error GoTo Errloop

    If IsValid(Txt(ExpenceAc), "Expences Ac") = False Then Txt(ExpenceAc).SetFocus: Exit Sub
    
    
    mSearchCode = IIf(TopCtrl1.TopText2 = "Add", Txt(ExpenceAc).Tag, RstMain!SearchCode)
    
    If TopCtrl1.TopText2 = "Add" Then
        If GCn.Execute("Select Count(*) From Budget_Exp Where ExpAc='" & mSearchCode & "' And Site_Code = '" & PubSiteCode & "'").Fields(0).Value > 0 Then
            MsgBox "Expence Ac Already Exist"
            Txt(ExpenceAc).SetFocus
            Exit Sub
        End If
    Else
        If GCn.Execute("Select Count(*) From Budget_Exp Where ExpAc='" & Txt(ExpenceAc).Tag & "' And ExpAc <> '" & mSearchCode & "' And Site_Code = '" & PubSiteCode & "'").Fields(0).Value > 0 Then
            MsgBox "Expence Ac Already Exist"
            Txt(ExpenceAc).SetFocus
            Exit Sub
        End If
    End If
    
    mTrans = 1
    GCn.BeginTrans
        If TopCtrl1.TopText2 = "Add" Then
            GCn.Execute "Delete From Budget_Exp Where ExpAc = '" & mSearchCode & "' And Site_Code = '" & PubSiteCode & "'"
        Else
            If RstMain.RecordCount > 0 Then GCn.Execute "Delete From Budget_Exp Where ExpAc = '" & mSearchCode & "'"
        End If
        For I = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(I, Col_Month_Desc) <> "" Then
                GCn.Execute "Insert Into Budget_Exp (ExpAc, VDate, Month, Amount, Site_Code, U_Name, U_EntDt, U_AE) " & _
                            "Values ('" & mSearchCode & "', " & ConvertDate(FGrid1.TextMatrix(I, Col_Date)) & ", '" & FGrid1.TextMatrix(I, Col_Month) & "', " & Val(FGrid1.TextMatrix(I, Col_Amount)) & ", '" & PubSiteCode & "', '" & pubUName & "', " & ConvertDate(PubLoginDate) & ", '" & left(TopCtrl1.TopText2, 1) & "') "
            End If
        Next I
    GCn.CommitTrans
    mTrans = 0
                    
    RstMain.Requery
    RsHelp.Requery
    RstMain.FIND ("SearchCode = '" & mSearchCode & "'")

    
    Disp_Text SETS("INI", Me, RstMain)
    MoveRec
    Grid_Hide

Exit Sub
Errloop:
    If mTrans = 1 Then GCn.RollbackTrans
    MsgBox err.Description, vbCritical
End Sub


Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub MoveRec()
Dim Rs As Recordset
Dim I As Integer
On Error GoTo Errloop


Grid_Hide
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    
    Set Rs = GCn.Execute("SELECT BE.*, S.Name As ExpenceAcName, M.Name As Month_Desc " & _
        " FROM (Budget_Exp BE Left Join SubGroup S On S.SubCode =  BE.ExpAc) " & _
        " Left Join Chas_Mth M On M.Code = BE.Month " & _
        " WHERE BE.ExpAc = '" & RstMain!SearchCode & "' And BE.Site_Code = '" & PubSiteCode & "' Order By BE.Month")
    
    I = 1
    FGrid1.Rows = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            FGrid1.AddItem ""
            
            FGrid1.TextMatrix(I, 0) = I
            Txt(ExpenceAc).Tag = XNull(Rs!ExpAc)
            Txt(ExpenceAc) = XNull(Rs!ExpenceAcName)
            
            FGrid1.TextMatrix(I, Col_Month) = XNull(Rs!Month)
            FGrid1.TextMatrix(I, Col_Month_Desc) = XNull(Rs!Month_Desc)
            FGrid1.TextMatrix(I, Col_Amount) = XNull(Rs!Amount)
            FGrid1.TextMatrix(I, Col_Date) = XNull(Rs!VDate)
            
            
            
            I = I + 1
            Rs.MoveNext
        Loop
        
        FGrid1.FixedRows = 1
    Else
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
    End If
    
    
End If
Exit Sub
Errloop:        MsgBox err.Description
End Sub
Private Sub TopCtrl1_eRef()
    RsHelp.Requery
    RsMonth.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub


Private Sub Txt_GotFocus(Index As Integer)

txtgrid1(0).Visible = False
Ctrl_GetFocus Txt(Index)
Grid_Hide

Select Case Index
    Case ExpenceAc
        DGHelp.Move Txt(Index).left, Txt(Index).top + Txt(Index).height + 20
        If RsHelp.RecordCount = 0 Or (RsHelp.EOF = True Or RsHelp.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsHelp!Name Then
            RsHelp.MoveFirst
            RsHelp.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If

Select Case Index
    Case ExpenceAc
        DGridTxtKeyDown DGHelp, Txt, ExpenceAc, RsHelp, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
End Select

If DGHelp.Visible = False And DGMonth.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = ExpenceAc Then Txt_Validate Index, True
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> ExpenceAc Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> ExpenceAc Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If

End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case Index
    Case ExpenceAc
        If DGHelp.Visible = True Then DGridTxtKeyPress Txt, ExpenceAc, RsHelp, keyascii, "Name"
End Select

'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Integer
Dim mDays As Integer
Select Case Index
    Case ExpenceAc
        If RsHelp.RecordCount = 0 Or (RsHelp.EOF = True Or RsHelp.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsHelp!Name
            Txt(Index).Tag = RsHelp!Code
        End If
End Select
Set Rst = Nothing

End Sub



Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
    Txt(I).Tag = ""
Next I

FGrid1.Rows = 1
FGrid1.AddItem ""
FGrid1.FixedRows = 1
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
Next


End Sub

'Private Sub Ini_Grid()
'    FGrid.RowHeightMin = 250
'    FGrid.ColWidth(25) = 0
'End Sub

Sub Grid_Hide()
    DGHelp.Visible = False
    DGMonth.Visible = False
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid1(Index)
    Grid_Hide
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    
    Select Case FGrid1.Col
        Case Col_Month_Desc
            DGMonth.Move txtgrid1(0).left, txtgrid1(0).top + txtgrid1(0).height + 20
            If RsMonth.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, Col_Month) = "" Then Exit Sub
            RsMonth.MoveFirst
            RsMonth.FIND "Code ='" & FGrid1.TextMatrix(FGrid1.Row, Col_Month) & "'"
    End Select
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtgrid1(0).TEXT = txtgrid1(0).Tag
        TxtGrid1_KeyUp Index, KeyCode, Shift
        FGrid1.SetFocus
        txtgrid1(0).Visible = False
        Exit Sub
    End If
    Select Case FGrid1.Col
        Case Col_Month_Desc
            If DGMonth.Visible = False Then DGridColSwap DGMonth, 0
            DGridTxtKeyDown DGMonth, txtgrid1, Index, RsMonth, KeyCode, False, 1
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGMonth.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, Col_Amount, , Col_Amount
                End If
            End If
        Case Col_Amount, Col_Date
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGMonth.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, Col_Amount, 1
                End If
            End If
                        
                        
    End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, keyascii As Integer)
    Call CheckQuote(keyascii)
    Select Case FGrid1.Col
        Case Col_Month_Desc
            DGridTxtKeyPress txtgrid1, Index, RsMonth, keyascii, "Name"
        Case Col_Amount
            NumPress txtgrid1(0), keyascii, 8, 2
    End Select
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid1.Col
            Case Col_Month_Desc
                If KeyCode <> 13 And DGHelp.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid1, Index, RsMonth, KeyCode, "Name"

        End Select
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
    'If ExitCtrl = False Then Exit Sub
    txtgrid1(Index).Visible = False
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim Repeat$
Select Case FGrid1.Col
    Case Col_Month_Desc
        If RsMonth.RecordCount = 0 Or txtgrid1(0).TEXT = "" Or RsMonth.EOF = True Or RsMonth.BOF = True Then
            FGrid1.TextMatrix(FGrid1.Row, Col_Month) = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_Month_Desc) = ""
        Else
            FGrid1.TextMatrix(FGrid1.Row, Col_Month) = RsMonth!Code
            FGrid1.TextMatrix(FGrid1.Row, Col_Month_Desc) = RsMonth!Name
        End If
            
    Case Col_Amount
        FGrid1 = Format(txtgrid1(0), "0.00")
        FGrid1.TextMatrix(FGrid1.Row, Col_Date) = PubLoginDate
    Case Col_Date
        FGrid1 = RetDate(txtgrid1(0))
End Select
TxtGrid1Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid1.SetFocus
    txtgrid1(0).Visible = False
End If
End Function


Sub Ini_Grid()
    With FGrid1
        .Cols = 5
        .TextMatrix(0, 0) = "Srl."
                                
        .TextMatrix(0, Col_Month_Desc) = "Month"
        .ColAlignment(Col_Month_Desc) = flexAlignLeftCenter
        .ColWidth(Col_Month_Desc) = 1200
                
        .ColWidth(Col_Month) = 0
                
        .TextMatrix(0, Col_Amount) = "Budget Amt"
        .ColAlignment(Col_Amount) = flexAlignRightCenter
        .ColWidth(Col_Amount) = 1200
        
        .TextMatrix(0, Col_Date) = "Date"
        .ColAlignment(Col_Date) = flexAlignLeftCenter
        .ColWidth(Col_Date) = 1200
        
    End With
End Sub


