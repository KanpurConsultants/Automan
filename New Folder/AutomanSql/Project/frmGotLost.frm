VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmGotLost 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Order Got / Lost Entry"
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
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGCont 
      Height          =   2910
      Left            =   180
      Negotiate       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4665
      Visible         =   0   'False
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   5133
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
         Weight          =   400
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Inv_No"
         Caption         =   "Invoice No"
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
         DataField       =   "Cust_Name"
         Caption         =   "Customer Name"
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
            ColumnWidth     =   2445.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2594.835
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGLostCat 
      Height          =   4515
      Left            =   1680
      Negotiate       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4410
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
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
         Caption         =   "Lost Category"
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
            ColumnWidth     =   3225.26
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDF4B5&
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
      Left            =   510
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3930
      Left            =   90
      TabIndex        =   3
      Top             =   1500
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   6932
      _Version        =   393216
      BackColorFixed  =   13623520
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   12640511
      GridColorUnpopulated=   13623520
      FocusRect       =   0
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
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
      _Band(0).Cols   =   2
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2970
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1035
      Width           =   1245
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   0
      Left            =   9945
      MaxLength       =   12
      TabIndex        =   7
      Text            =   "VFalse"
      Top             =   390
      Visible         =   0   'False
      Width           =   1530
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
      Height          =   225
      Index           =   1
      Left            =   2970
      MaxLength       =   40
      TabIndex        =   1
      Top             =   780
      Width           =   4785
   End
   Begin MSDataGridLib.DataGrid DGRep 
      Height          =   4515
      Left            =   7965
      Negotiate       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   570
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
         Caption         =   "Representative Name"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   5
      Left            =   1275
      TabIndex        =   8
      Top             =   1050
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Executive*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   1275
      TabIndex        =   4
      Top             =   795
      Width           =   1455
   End
End
Attribute VB_Name = "frmGotLost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim ExitCtrl As Boolean

Dim rsLostCat As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim rsCont As ADODB.Recordset
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Dim ListArray As Variant
Dim mListItem As ListItem

Private Const REP_CODE As Byte = 1
Private Const RepPWD As Byte = 2
'Grid Columns
Private Const PartyCode As Byte = 1
Private Const PartyName As Byte = 2
Private Const Site_Code As Byte = 3
Private Const Rep_Code2 As Byte = 4
Private Const StartDate As Byte = 5
Private Const Model As Byte = 6
Private Const Call_Status2 As Byte = 7
Private Const Got_Lost As Byte = 8
Private Const GotLost_Date As Byte = 9
Private Const Lost_Cat As Byte = 10
Private Const Lost_CatName As Byte = 11
Private Const QuotDocId As Byte = 12
Private Const PurchModel As Byte = 13
Private Const QuotSrl_No As Byte = 14

Private Sub Ini_Grid()
Dim I As Byte
    
    With FGrid
        .left = Me.left
        .width = Me.width - 120
        .Cols = 15
        .top = 1500
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 525
        .ColWidth(PartyCode) = 0
        .ColWidth(Site_Code) = 0
        .ColWidth(Rep_Code2) = 0
        
        .TextMatrix(0, PartyName) = "Party"
        .ColAlignmentFixed(PartyName) = flexAlignCenterCenter
        .ColAlignment(PartyName) = flexAlignLeftCenter
        .ColWidth(PartyName) = 3500
        
        .TextMatrix(0, StartDate) = "Start Date"
        .ColAlignmentFixed(StartDate) = flexAlignCenterCenter
        .ColAlignment(StartDate) = flexAlignLeftCenter
        .ColWidth(StartDate) = 1245
       
        .TextMatrix(0, Model) = "Model"
        .ColAlignmentFixed(Model) = flexAlignCenterCenter
        .ColAlignment(Model) = flexAlignLeftCenter
        .ColWidth(Model) = 1680
        
        .TextMatrix(0, Call_Status2) = "Call Stat."
        .ColAlignmentFixed(Call_Status2) = flexAlignCenterCenter
        .ColAlignmentFixed(Call_Status2) = flexAlignLeftCenter
        .ColWidth(Call_Status2) = 1050
        
        .TextMatrix(0, Got_Lost) = "Got/Lost"
        .ColAlignmentFixed(Got_Lost) = flexAlignCenterCenter
        .ColAlignmentFixed(Got_Lost) = flexAlignLeftCenter
        .ColWidth(Got_Lost) = 1000
        
        .TextMatrix(0, GotLost_Date) = "Got/Lost Dt."
        .ColAlignmentFixed(GotLost_Date) = flexAlignCenterCenter
        .ColAlignment(GotLost_Date) = flexAlignLeftCenter
        .ColWidth(GotLost_Date) = 1380
        
        .ColWidth(Lost_Cat) = 0
        .TextMatrix(0, Lost_CatName) = "Lost Cat."
        .ColAlignmentFixed(Lost_CatName) = flexAlignCenterCenter
        .ColAlignment(Lost_CatName) = flexAlignLeftCenter
        .ColWidth(Lost_CatName) = 1860
  
        .TextMatrix(0, QuotDocId) = "OrderDocId"
        .ColAlignmentFixed(QuotDocId) = flexAlignCenterCenter
        .ColAlignment(QuotDocId) = flexAlignRightCenter
        .ColWidth(QuotDocId) = 2000
        
        .TextMatrix(0, PurchModel) = "PurchModel"
        .ColAlignmentFixed(PurchModel) = flexAlignCenterCenter
        .ColAlignment(PurchModel) = flexAlignRightCenter
        .ColWidth(PurchModel) = 1650
        
        .TextMatrix(0, QuotSrl_No) = "SNo"
        .ColAlignmentFixed(QuotSrl_No) = flexAlignRightCenter
        .ColWidth(QuotSrl_No) = 555
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    DGLostCat.left = mLtScale: DGLostCat.top = mTopScale
    DGCont.width = 5000: DGCont.left = Me.width - (DGCont.width + mRtScale): DGCont.top = mTopScale: DGCont.height = 5000
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid_KeyPress (vbKeyReturn)
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
    If FGrid.BackColorSel = BackColorSelLeave Then FGrid.Col = Got_Lost
    FGrid.BackColorSel = BackColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
    End If
    Exit Sub
'        SendKeysA vbKeyTab, True
'        KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case Got_Lost
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Lost_Cat) = ""
            FGrid.TextMatrix(FGrid.Row, Lost_CatName) = ""
            FGrid.TextMatrix(FGrid.Row, GotLost_Date) = ""
        Case GotLost_Date
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case Lost_CatName
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, Lost_Cat) = ""
        Case PurchModel
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If
KeyCode = 0
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Or FGrid.TextMatrix(FGrid.Row, Model) = "" Then Exit Sub
'SetMaxLength
Select Case FGrid.Col
    Case Got_Lost
        If UCase(Chr(KeyAscii)) = "G" Then
            FGrid.TextMatrix(FGrid.Row, Got_Lost) = "Got"
        ElseIf UCase(Chr(KeyAscii)) = "L" Then
            FGrid.TextMatrix(FGrid.Row, Got_Lost) = "Lost"
        End If
    Case GotLost_Date
        If FGrid.TextMatrix(FGrid.Row, Got_Lost) <> "" Then
            Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        End If
    Case Lost_CatName
        If FGrid.TextMatrix(FGrid.Row, Got_Lost) = "Lost" And _
            FGrid.TextMatrix(FGrid.Row, GotLost_Date) <> "" Then
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
        End If
    Case QuotDocId
        If FGrid.TextMatrix(FGrid.Row, Got_Lost) = "Got" Then
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
        End If
    Case PurchModel
        If FGrid.TextMatrix(FGrid.Row, Got_Lost) = "Lost" Then
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_LostFocus()
    If TxtGrid(0).Visible = False Then FGrid.BackColorSel = BackColorSelLeave
End Sub
Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
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
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
       Dim SiteCond As String
       
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = "where  " & cMID("Veh_SubGroupQuot.QuotDocId", "3", "1") & "='" & PubSiteCode & "'"
    Else
      SiteCond = ""
    End If
    
    If PubMoveRecYn Then
        Master.Open "select distinct Rep_Code as SearchCode, E.Emp_Name from Veh_SubGroupQuot left join Emp_Mast E on Veh_SubGroupQuot.Rep_Code=E.Emp_Code " & SiteCond & " order by E.Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select distinct Top 1 Rep_Code as SearchCode, E.Emp_Name from Veh_SubGroupQuot left join Emp_Mast E on Veh_SubGroupQuot.Rep_Code=E.Emp_Code " & SiteCond & "  order by E.Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    Set rsLostCat = New ADODB.Recordset
    rsLostCat.CursorLocation = adUseClient
    rsLostCat.Open "select Code,Name from Veh_OrdLostCatg order by Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGLostCat.DataSource = rsLostCat
    
    Set rsCont = New ADODB.Recordset
    rsCont.CursorLocation = adUseClient
    rsCont.Open "Select VO.OrdDocId as Inv_No,SG.Name as Cust_Name,VS.Sal_VDate,VS.MODEL,VS.ChassisNo,C.Col_Desc from ((Veh_Stock VS Left Join ColMast C on VS.Colour_Code=C.Col_Code) Left Join Veh_Order VO on VS.Sal_DocId=VO.Inv_DocID) Left Join SubGroup SG On VO.PartyCode=SG.SubCode where Len(VS.Sal_DocID)  > 1", GCn, adOpenDynamic, adLockOptimistic
    Set DGCont.DataSource = rsCont
    rsCont.Sort = "Inv_No"
    rsCont.Sort = "Cust_Name"
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
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
Set rsLostCat = Nothing
Set Master = Nothing
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(RepPWD).Enabled = True
    txt(RepPWD).SetFocus
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
          Dim SiteCond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = "where  " & cMID("Veh_SubGroupQuot.QuotDocId", "3", "1") & "='" & PubSiteCode & "'"
    Else
      SiteCond = ""
    End If
    
    GSQL = "select distinct Rep_Code as SearchCode,E.Emp_Name from Veh_SubGroupQuot left join Emp_Mast E on Veh_SubGroupQuot.Rep_Code=E.Emp_Code " & SiteCond & " order by E.Emp_Name"
    Set SearchForm = Me
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
        Set Master = GCn.Execute("select distinct Rep_Code as SearchCode, E.Emp_Name from Veh_SubGroupQuot left join Emp_Mast E on Veh_SubGroupQuot.Rep_Code=E.Emp_Code Where Rep_Code = '" & MyValue & "' order by E.Emp_Name")
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
    rsLostCat.Requery
    Master.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, mRepPWD$
    Dim mTrans As Boolean
    On Error GoTo errlbl
    
    Grid_Hide
    mRepPWD = GCn.Execute("Select Access_PWD from Emp_Mast where Emp_Code='" & txt(REP_CODE).Tag & "'").Fields(0).Value
    If txt(RepPWD) <> mRepPWD Then
        MsgBox "Please enter valid password!", vbOKOnly, "Authorisation Checking"
        txt(RepPWD).SetFocus
        Exit Sub
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Got_Lost) <> "" Then
            If FGrid.TextMatrix(I, GotLost_Date) = "" Then MsgBox "Fill Got/Lost Date in Row No " & I, vbInformation, "Required Data": FGrid.Row = I: FGrid.Col = GotLost_Date: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Got_Lost) = "Lost" Then
                If FGrid.TextMatrix(I, Lost_CatName) = "" Then MsgBox "Fill Lost Category in Row No " & I, vbInformation, "Required Data": FGrid.Row = I: FGrid.Col = Lost_CatName: FGrid.SetFocus: Exit Sub
            End If
        End If
    Next
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Got_Lost) = "Got" Then
            If FGrid.TextMatrix(I, QuotDocId) = "" Then MsgBox "Fill Order DocID in Row No. " & I, vbInformation, "Required Data": FGrid.Row = I: FGrid.Col = QuotDocId: FGrid.SetFocus: Exit Sub
        End If
    Next
 GCn.BeginTrans
    mTrans = True
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, StartDate) <> "" Then
            GSQL = "Update Veh_SubGroupQuot Set Got_Lost='" & FGrid.TextMatrix(I, Got_Lost) & _
                "',GotLost_Date=" & ConvertDate(FGrid.TextMatrix(I, GotLost_Date)) & " ,Lost_Cat='" & FGrid.TextMatrix(I, Lost_Cat) & "',OrdDocID='" & FGrid.TextMatrix(I, QuotDocId) & "',PurchModel='" & FGrid.TextMatrix(I, PurchModel) & _
                "',U_Name='" & pubUName & "' , U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='" & left(TopCtrl1.TopText2.CAPTION, 1) & _
                "' where REP_CODE='" & txt(REP_CODE).Tag & _
                "' and PartyCode='" & FGrid.TextMatrix(I, PartyCode) & _
                "' and StartDate=" & ConvertDate(FGrid.TextMatrix(I, StartDate)) & _
                "  and MODEL='" & FGrid.TextMatrix(I, Model) & "'"
            GCn.Execute (GSQL)
        End If
    Next
GCn.CommitTrans
mTrans = False
'    Master.FIND "SearchCode = '" & mStr & "'"
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
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    txt(Index).TEXT = ""
    Grid_Hide
    Exit Sub
End If
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
'    If TopCtrl1.TopText2 = "Add" And Index <> VisitDate Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    If TopCtrl1.TopText2 = "Edit" And Index <> RepPWD Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'Select Case Index
'    Case Visit_Call
'        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
'    Case Call_Status
'        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
'End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub DGLostCat_Click()
    If rsLostCat.RecordCount > 0 Then
        TxtGrid(0).TEXT = rsLostCat!Name
        TxtGrid(0).Tag = rsLostCat!Code
    End If
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGLostCat.Visible = False
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = "": txt(I).Tag = ""
Next I
End Sub

Private Sub MoveRec()
Dim Rs As Recordset, I As Integer, mPartyName$
On Error GoTo error1
TopCtrl1.tPrn = False
TopCtrl1.tAdd = False
TopCtrl1.tDel = False
If Master.RecordCount > 0 Then
    txt(REP_CODE).Tag = Master!SearchCode
    txt(REP_CODE).TEXT = Master!Emp_Name
    txt(RepPWD) = ""
    'Fill VehSubGroupQuot Records
    Set Rs = New Recordset
    GSQL = "Select * from Veh_SubGroupQuot where Rep_Code='" & Master!SearchCode & "'"
    Set Rs = GCn.Execute(GSQL)
    FGrid.Rows = 1: FGrid.Redraw = False
    I = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            If Rs!ProspectiveCust_SubGroup = 1 Then 'Yes
                GSQL = "select name from subgroup where SubCode = '" & Rs!PartyCode & "'"
            Else
                GSQL = "select name from ProspectiveCust where CUST_CODE = '" & Rs!PartyCode & "'"
            End If
            If GCn.Execute(GSQL).RecordCount > 0 Then
                mPartyName = GCn.Execute(GSQL).Fields(0).Value
            End If
            With FGrid
                .AddItem I
                .TextMatrix(I, PartyCode) = Rs!PartyCode
                .TextMatrix(I, PartyName) = mPartyName
                .TextMatrix(I, Site_Code) = Rs!Site_Code
                .TextMatrix(I, Rep_Code2) = Rs!REP_CODE
                .TextMatrix(I, StartDate) = Rs!StartDate
                .TextMatrix(I, Model) = Rs!Model
                .TextMatrix(I, Call_Status2) = FxCallStatus(Rs!Call_Status, True)
                .TextMatrix(I, Got_Lost) = XNull(Rs!Got_Lost)
                .TextMatrix(I, GotLost_Date) = XNull(Format(Rs!GotLost_Date, "dd/mmm/yyyy"))
                .TextMatrix(I, Lost_Cat) = IIf(IsNull(Rs!Lost_Cat) Or Rs!Lost_Cat = "", "", Rs!Lost_Cat)
                If IsNull(Rs!Lost_Cat) Or Rs!Lost_Cat = "" Then
                    .TextMatrix(I, Lost_CatName) = ""
                Else
                    .TextMatrix(I, Lost_CatName) = GCn.Execute("Select NAME from Veh_OrdLostCatg where code='" & Rs!Lost_Cat & "'").Fields(0).Value
                End If
                .TextMatrix(I, QuotDocId) = IIf(IsNull(Rs!OrdDocId), "", Rs!OrdDocId)
                .TextMatrix(I, PurchModel) = IIf(IsNull(Rs!PurchModel), "", Rs!PurchModel)
                .TextMatrix(I, QuotSrl_No) = IIf(IsNull(Rs!QuotSrl_No), "", Rs!QuotSrl_No)
            End With
            Rs.MoveNext
           I = I + 1
        Loop
        FGrid.FixedRows = 1
    End If
Else
    Call BlankText
    Ini_Grid
    FGrid.Rows = 1
End If
If FGrid.Rows = 1 Then FGrid.AddItem FGrid.Rows:  FGrid.FixedRows = 1
FGrid.Redraw = True
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = False
'    txt(i).ForeColor = CtrlFColOrg
Next
'txtDisabled_Color Me
End Sub

Private Sub Grid_Hide()
    If DGLostCat.Visible = True Then DGLostCat.Visible = False
    If DGCont.Visible = True Then DGCont.Visible = False
End Sub

Private Sub RemoveTxtNull()
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).TEXT = IIf(IsNull(txt(I).TEXT), "", txt(I).TEXT)
Next I
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
         Case Lost_CatName
            If rsLostCat.RecordCount = 0 Or (rsLostCat.EOF = True Or rsLostCat.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Lost_Cat) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Lost_Cat) <> rsLostCat!Code Then
                rsLostCat.MoveFirst
                rsLostCat.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, Lost_Cat) & "'"
            End If
    End Select
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            TxtGrid(0).TEXT = TxtGrid(0).Tag
            TxtGrid_KeyUp Index, KeyCode, Shift
            TxtGrid(0).Visible = False
            Grid_Hide
            FGrid.SetFocus
            Exit Sub
        End If
        Select Case FGrid.Col
            Case GotLost_Date
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, QuotSrl_No
                    End If
                End If
            Case Lost_CatName    '1
                DGridTxtKeyDown DGLostCat, TxtGrid, Index, rsLostCat, KeyCode, True, 0
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, QuotSrl_No
                    End If
                End If
            Case QuotDocId
                DGridTxtKeyDown DGCont, TxtGrid, Index, rsCont, KeyCode, False, 0
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, QuotSrl_No
                    End If
                End If
            Case PurchModel
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, QuotSrl_No
                    End If
                End If
        End Select
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case Lost_CatName
        If DGLostCat.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsLostCat, KeyAscii, "Name"
    Case QuotDocId
        If DGCont.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsCont, KeyAscii, "Inv_No"
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case Lost_CatName
        If KeyCode <> 13 And DGLostCat.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, rsLostCat, KeyCode, "Name", True
    Case QuotDocId
        If KeyCode <> 13 And DGCont.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, rsCont, KeyCode, "Inv_No", True
End Select
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim j As Integer
Dim Rst As ADODB.Recordset
Dim GridCol As Byte
GridCol = FGrid.Col
Select Case GridCol
    Case GotLost_Date
        TxtGrid(0).TEXT = RetDate(TxtGrid(0))
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    Case Lost_CatName
        If rsLostCat.RecordCount = 0 Or (rsLostCat.EOF = True Or rsLostCat.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, Lost_Cat) = ""
            FGrid.TextMatrix(FGrid.Row, Lost_CatName) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, Lost_Cat) = rsLostCat!Code
            FGrid.TextMatrix(FGrid.Row, Lost_CatName) = rsLostCat!Name
        End If
    Case QuotDocId
            FGrid.TextMatrix(FGrid.Row, QuotDocId) = rsCont!Inv_No
    Case PurchModel
            FGrid.TextMatrix(FGrid.Row, PurchModel) = TxtGrid(0).TEXT
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End If
End Function
Private Function FxCallStatus(CallStatus As Variant, VarTypeNumber As Boolean) As Variant
If VarTypeNumber Then
    If Not IsNull(CallStatus) Then
        Select Case CallStatus
            Case 0
                FxCallStatus = "Cold"
            Case 1
                FxCallStatus = "Warm"
            Case 2
                FxCallStatus = "Hot"
            Case 3
                FxCallStatus = "Nill"
        End Select
    End If
Else
    If Not IsNull(CallStatus) Then
        Select Case CallStatus
            Case "Cold"
                FxCallStatus = 0
            Case "Warm"
                FxCallStatus = 1
            Case "Hot"
                FxCallStatus = 2
            Case "Nill"
                FxCallStatus = 3
        End Select
    End If
End If
End Function

Private Sub SetMaxLength()
    Select Case FGrid.Col
        Case StartDate
            TxtGrid(0).MaxLength = 12
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Else
            TxtGrid(0).MaxLength = 0
    End Select
End Sub

