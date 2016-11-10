VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form FrmVouPre 
   BackColor       =   &H00E8DDB7&
   Caption         =   "Voucher Prefix Creation"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11610
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
   ScaleHeight     =   6930
   ScaleWidth      =   11610
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8DDB7&
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6165
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8DDB7&
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6225
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6165
      Width           =   1290
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DGGroup 
      Height          =   2580
      Left            =   -2205
      Negotiate       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6165
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   4551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   0   'False
      ForeColor       =   8388608
      HeadLines       =   1.5
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   ""
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
   Begin MSDataGridLib.DataGrid DgSpot 
      Height          =   2580
      Left            =   -1485
      Negotiate       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   4551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   0   'False
      ForeColor       =   8388608
      HeadLines       =   1.5
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   ""
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8DDB7&
      Caption         =   "Cancel  && E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6165
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8DDB7&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6165
      Width           =   1290
   End
   Begin MSDataGridLib.DataGrid DGItem 
      Height          =   2580
      Left            =   -1785
      Negotiate       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6060
      Visible         =   0   'False
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   4551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   0   'False
      ForeColor       =   8388608
      HeadLines       =   1.5
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   ""
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
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1620
      TabIndex        =   3
      Top             =   3075
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   5460
      Left            =   15
      TabIndex        =   0
      Top             =   495
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   9631
      _Version        =   393216
      BackColor       =   15525079
      Cols            =   8
      BackColorFixed  =   14940925
      ForeColorFixed  =   8388608
      BackColorSel    =   16711680
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14865856
      GridColor       =   14940925
      GridColorFixed  =   12632319
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      FormatString    =   "SrNo."
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
End
Attribute VB_Name = "FrmVouPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Master As ADODB.Recordset
Public VouType As String
Dim mAdd As Boolean, mEdit As Boolean, mDel As Boolean, mPrn As Boolean
Private Const CellBackColLeave As String = &HECE4D7    '&HECE4D7   '&HEDF7FE
Private Const CellForeColLeave As String = &HFF00FF
Private Const CellBackColEnter As String = &HF0D5BF
Private Const GridBackColorBkg As String = &HE2D5C0
Private Const Description As Byte = 1
Private Const ManAuto As Byte = 2
Private Const DIV As Byte = 3
Private Const DateFrom As Byte = 4
Private Const Prefix As Byte = 5
Private Const StartNo As Byte = 6
Private Const Vtype As Byte = 7
Dim EditName As String
Dim GridKey As Integer
Dim DataAddMode As Boolean
Dim TAddMode As Boolean
Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 2
        Select Case FGrid.Col
            Case Description
                Master.Sort = "Description"
            Case Vtype
                Master.Sort = "Vtype"
            Case StartNo
                Master.Sort = "StartNo"
            Case ManAuto
                Master.Sort = "AutomaticManual"
            Case Prefix
                Master.Sort = "Prefix"
            Case DateFrom
                Master.Sort = "DateFrom"
        End Select
'        Set FGrid.DataSource = Master
        Ini_Grid
    Case 3
        Set Master = GCnFa.Execute("SELECT distinct Voucher_Type.Description,Voucher_Type.Number_Method as AutomaticManual," & PubDivCode & " as Div, '' as DateFrom, Voucher_Prefix.Prefix, 0 as StartNo, Voucher_Prefix.V_Type as VType " & _
        "FROM Voucher_Prefix LEFT JOIN Voucher_Type ON Voucher_Prefix.V_Type = Voucher_Type.V_Type where Voucher_Prefix.div_Code ='" & PubDivCode & "'")
        DataAddMode = True
'        FGrid.Redraw = False
        Set FGrid.DataSource = Master
        With FGrid
            .RowHeightMin = PubGridRowHeight
            .TextMatrix(0, 0) = ""
            .ColAlignment(0) = flexAlignLeftCenter
            .ColWidth(0) = 450
    
            .TextMatrix(0, Description) = "Description"
            .ColAlignment(Description) = flexAlignLeftCenter
            .ColWidth(Description) = 2500
            
            .TextMatrix(0, ManAuto) = "Auto/Manual"
            .ColAlignment(ManAuto) = flexAlignLeftCenter
            .ColWidth(ManAuto) = 1500
            
            .TextMatrix(0, DIV) = "Div"
            .ColAlignment(DIV) = flexAlignLeftCenter
            .ColWidth(DIV) = 700
            
            .TextMatrix(0, DateFrom) = "DateFrom"
            .ColAlignment(DateFrom) = flexAlignLeftCenter
            .ColWidth(DateFrom) = 1100
    
            .TextMatrix(0, Prefix) = "Prefix"
            .ColAlignment(Prefix) = flexAlignLeftCenter
            .ColWidth(Prefix) = 700
    
            .TextMatrix(0, StartNo) = "StartNo"
            .ColAlignmentFixed(StartNo) = flexAlignRightCenter
            .ColWidth(StartNo) = 700
            
            .TextMatrix(0, Vtype) = "VType"
            .ColAlignment(Vtype) = flexAlignLeftCenter
            .ColWidth(Vtype) = 700
        End With '            FGrid.Redraw = True
        If mAdd = True Then FGrid.AddItem "**"
    Case 1
        Unload Me
    Case 0
        SaveData
        Unload Me
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
'FormKeyDown Me, KeyCode, Shift
If KeyCode = vbKeyS And Shift = 2 Then Call SaveData: FGrid.SetFocus
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
    TopCtrl1.TopText2 = "Add"
    If InStr(PubUParam, "A") <> 0 Then mAdd = True Else mAdd = False
    If InStr(PubUParam, "E") <> 0 Then mEdit = True Else mEdit = False
    If InStr(PubUParam, "D") <> 0 Then mDel = True Else mDel = False
    If InStr(PubUParam, "P") <> 0 Then mPrn = True Else mPrn = False

    TopCtrl1.Tag = PubUParam
    WinSetting Me
       
    Set Master = GCnFa.Execute("SELECT Voucher_Type.Description,Voucher_Type.Number_Method as AutomaticManual,Voucher_Prefix.Div_Code as Div,Voucher_Prefix.Date_From as DateFrom, Voucher_Prefix.Prefix, Voucher_Prefix.Start_Srl_No as StartNo, Voucher_Prefix.V_Type as VType " & _
    "FROM Voucher_Prefix LEFT JOIN Voucher_Type ON Voucher_Prefix.V_Type = Voucher_Type.V_Type where Voucher_Prefix.div_Code ='" & PubDivCode & "' order by Voucher_Prefix.Date_From desc")
    
'    MoveRec
    Ini_Grid
    DataAddMode = False
  Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set RsItem = Nothing
'Set RsSpot = Nothing
'Set RsGroup = Nothing
Set Master = Nothing
End Sub
Private Sub DGItem_Click()
'DGItem.Visible = False
'If RsItem.RecordCount > 0 Then
'        TxtGrid(0).Text = RsItem!Name
'        FGrid.TextMatrix(FGrid.Row, ItemName) = RsItem!Name
'        FGrid.TextMatrix(FGrid.Row, ItemCode) = RsItem!Code
'End If
End Sub
Private Sub DGGroup_Click()
'DGGroup.Visible = False
'If RsGroup.RecordCount > 0 Then
'        TxtGrid(0).Text = RsGroup!Name
'        FGrid.TextMatrix(FGrid.Row, ItemGroup) = RsGroup!Name
'        FGrid.TextMatrix(FGrid.Row, GroupCode) = RsGroup!Code
'End If
End Sub

Private Sub DGSpot_Click()
'DgSpot.Visible = False
'If RsSpot.RecordCount > 0 Then
'        TxtGrid(0).Text = RsSpot!Name
'        FGrid.TextMatrix(FGrid.Row, SpotName) = RsSpot!Name
'End If
End Sub

Private Sub SaveData()
    Dim i As Integer
    Dim mTrans As Boolean
    On Error GoTo errlbl
    GCnFa.BeginTrans
    mTrans = True
    For i = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(i, Vtype) <> "" And (FGrid.TextMatrix(i, DateFrom) = "" Or FGrid.TextMatrix(i, Prefix) = "" Or FGrid.TextMatrix(i, StartNo) = "" Or FGrid.TextMatrix(i, ManAuto) = "") Then
            MsgBox "Fill Data In Row No " & i, vbInformation, "Validation Check": FGrid.Row = i: TxtGrid(0).left = FGrid.left: Exit Sub
        End If
    Next
    If DataAddMode = False Then
        For i = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(i, 0) = "*" Then
                If FGrid.TextMatrix(i, Vtype) <> "" Then
                      GCnFa.Execute "update Voucher_Type set Number_Method='" & IIf(left(FGrid.TextMatrix(i, ManAuto), 1) = "M", "Manual", "Automatic") & "' where V_Type = '" & FGrid.TextMatrix(i, Vtype) & "'"
                End If
            End If
        Next
    End If
    For i = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(i, Vtype) <> "" Then
            If DataAddMode = True Then
                If FGrid.TextMatrix(i, 0) = "*" Then
                    GCnFa.Execute ("insert into Voucher_Prefix(V_Type,Date_From,Div_Code,Prefix,Start_Srl_No)  " & _
                    "values('" & FGrid.TextMatrix(i, Vtype) & "'," & ConvertDate(FGrid.TextMatrix(i, DateFrom)) & ",'" & PubDivCode & "','" & FGrid.TextMatrix(i, Prefix) & "'," & _
                    "" & Val(FGrid.TextMatrix(i, StartNo)) & ")")
                End If
            Else
                If FGrid.TextMatrix(i, 0) = "*" Then
                    GCnFa.Execute "update Voucher_Prefix set Prefix='" & FGrid.TextMatrix(i, Prefix) & "',Start_Srl_No=" & Val(FGrid.TextMatrix(i, StartNo)) & " where V_Type ='" & FGrid.TextMatrix(i, Vtype) & "' and Date_From=" & ConvertDate(FGrid.TextMatrix(i, DateFrom)) & " and Div_Code='" & PubDivCode & "'"
                End If
            End If
        End If
    Next
GCnFa.CommitTrans
mTrans = False
Master.Requery
MsgBox "Data Updation Complete"
'    RsSpot.Requery
    'Call MoveRec
 DataAddMode = False
    Exit Sub
errlbl:
    If mTrans = True Then
        GCnFa.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell--> Enter Cell-->KeyDown
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys vbTab
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyReturn Then
       Select Case FGrid.Col
            Case StartNo
                    FGrid.TextMatrix(FGrid.Row, 0) = "*"
                    Call GridDblClick(Me, FGrid, TxtGrid, 0)
                    TAddMode = False
            Case Prefix
                If DataAddMode = True Then
                    FGrid.TextMatrix(FGrid.Row, 0) = "*"
                    Call GridDblClick(Me, FGrid, TxtGrid, 0)
                    TAddMode = False
                End If
            Case ManAuto
                If DataAddMode = False And FGrid.TextMatrix(FGrid.Row, Vtype) <> "F_AO" And FGrid.TextMatrix(FGrid.Row, Vtype) <> "SXAO" Then
                    FGrid.TextMatrix(FGrid.Row, 0) = "*"
                    Call GridDblClick(Me, FGrid, TxtGrid, 0)
                    TAddMode = False
                End If
            Case DateFrom
                If DataAddMode = True Then
                    FGrid.TextMatrix(FGrid.Row, 0) = "*"
                    Call GridDblClick(Me, FGrid, TxtGrid, 0)
                    TAddMode = False
                End If
        End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_DblClick()
Select Case FGrid.Col
    Case StartNo
            FGrid.TextMatrix(FGrid.Row, 0) = "*"
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    Case Prefix
        If DataAddMode = True Then
            FGrid.TextMatrix(FGrid.Row, 0) = "*"
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
        End If
    Case ManAuto
        If DataAddMode = False And FGrid.TextMatrix(FGrid.Row, Vtype) <> "F_AO" And FGrid.TextMatrix(FGrid.Row, Vtype) <> "SXAO" Then
            FGrid.TextMatrix(FGrid.Row, 0) = "*"
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
        End If
    Case DateFrom
        If DataAddMode = True Then
            FGrid.TextMatrix(FGrid.Row, 0) = "*"
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
        End If
End Select
TAddMode = False
End Sub
Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim i As Integer
'If FGrid.ColSel = False Then Exit Sub
'If KeyCode = vbKeyD And Shift = 2 Then
'If mDel = False Then MsgBox "Permission denied", vbInformation, "Permission": FGrid.SetFocus: Exit Sub
'    If FGrid.Row >= 1 Then
'        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
'            If FGrid.TextMatrix(FGrid.Row, SpotName) <> "" Then GcnFA.Execute ("delete from spot where spotname = '" & FGrid.TextMatrix(FGrid.Row, SpotName) & "'")
'            Master.Requery
'            Ini_Grid
'        Else
'            Exit Sub
'        End If
'    Else
'        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
'    End If
'    RsSpot.Requery
'    FGrid.SetFocus
'End If
'Exit Sub
End Sub


Private Sub FGrid_KeyPress(KeyAscii As Integer)
If FGrid.TextMatrix(FGrid.Row, 0) <> "**" Then If mEdit = True Then FGrid.TextMatrix(FGrid.Row, 0) = "*"
    Select Case FGrid.Col
        Case Prefix
            If DataAddMode = True Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
            End If
        Case ManAuto
            If DataAddMode = False And FGrid.TextMatrix(FGrid.Row, Vtype) <> "F_AO" And FGrid.TextMatrix(FGrid.Row, Vtype) <> "SXAO" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
            End If
        Case DateFrom
            If DataAddMode = True Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
            End If
        Case StartNo
            Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub


Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
'    Select Case FGrid.Col
'        Case SpotName
'            If DgSpot.Visible = False Then DgSpot.left = TxtGrid(0).left: DgSpot.top = TxtGrid(0).top + TxtGrid(0).Height + 20
'        Case ItemName
'            If DGItem.Visible = False Then DGItem.left = TxtGrid(0).left: DGItem.top = TxtGrid(0).top + TxtGrid(0).Height + 20
'            If RsItem.RecordCount = 0 Or (RsItem.EOF = True Or RsItem.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ItemName) = "" Then Exit Sub
'            If FGrid.TextMatrix(FGrid.Row, ItemName) <> RsItem!Name Then
'                RsItem.MoveFirst
'                RsItem.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, ItemName) & "'"
'            End If
'        Case ItemGroup
'            If DGGroup.Visible = False Then DGGroup.left = TxtGrid(0).left: DGGroup.top = TxtGrid(0).top + TxtGrid(0).Height + 20
'            If RsGroup.RecordCount = 0 Or (RsGroup.EOF = True Or RsGroup.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ItemGroup) = "" Then Exit Sub
'            If FGrid.TextMatrix(FGrid.Row, ItemGroup) <> RsGroup!Name Then
'                RsItem.MoveFirst
'                RsItem.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, ItemGroup) & "'"
'            End If
'    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).Text = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        TxtGrid(0).Visible = False
        Grid_Hide
        FGrid.SetFocus
        Exit Sub
    End If
    Select Case FGrid.Col
'        Case SpotName
'            DGridTxtKeyDown_Mast DgSpot, TxtGrid, 0, RsSpot, KeyCode, True, 0
'            If KeyCode = vbKeyReturn Then
'                If TxtGridLeave = True Then
'                    GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, 9, , , False
'                End If
'            End If
'        Case ItemName
'            DGridTxtKeyDown DGItem, TxtGrid, 0, RsItem, KeyCode, True, 1
'            If KeyCode = vbKeyReturn Then
'                If TxtGridLeave = True Then
'                    GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, 9, , , False
'                End If
'            End If
'        Case ItemGroup
'            DGridTxtKeyDown DGGroup, TxtGrid, 0, RsGroup, KeyCode, True, 1
'            If KeyCode = vbKeyReturn Then
'                If TxtGridLeave = True Then
'                    GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, 9, , , False
'                End If
'            End If
        Case ManAuto, DateFrom, Prefix, StartNo
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, StartNo, , , False
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
 Call CheckQuote(KeyAscii)
 
Select Case FGrid.Col
'    Case ItemName
'        If DGItem.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsItem, KeyAscii, "Name"
'    Case ItemGroup
'        If DGGroup.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsGroup, KeyAscii, "Name"
    Case Prefix
        If VouType = "FA" Then
            If Len(TxtGrid(Index)) = 4 Then KeyAscii = 0
        Else
            If Len(TxtGrid(Index)) = 5 Then KeyAscii = 0
        End If
    Case StartNo
        Call NumPress(TxtGrid(Index), KeyAscii, 5, 0)
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
Select Case FGrid.Col
    Case ManAuto
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            TxtGrid(Index) = ""
        ElseIf UCase(left$(TxtGrid(Index), 1)) = "A" Then
            TxtGrid(Index) = "Automatic"
        Else
            TxtGrid(Index) = "Manual"
        End If
End Select
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Select Case FGrid.Col
    Case Prefix, ManAuto
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
    Case StartNo
        If DataAddMode = False And Val(TxtGrid(0).Text) < Val(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) Then
            MsgBox "Tou Can't Fill Less Value than the existing Value"
            Cancel = True
            Exit Sub
        End If
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).Text), "0")
    Case DateFrom
        TxtGrid(0) = RetDate(TxtGrid(0))
        If TxtGrid(0) <> "" Then
            If DataAddMode = True Then
                Set GRs = GCnFa.Execute("select * from Voucher_Prefix where date_From=#" & TxtGrid(0) & "# and V_Type ='" & FGrid.TextMatrix(FGrid.Row, Vtype) & "' and Div_Code='" & PubDivCode & "'")
                If GRs.RecordCount = 0 Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
                Else
                    MsgBox "Voucher Prefix Already Exist For This Date", vbInformation, "Validation Check": Cancel = True
                End If
            End If
        End If
End Select
End Sub

Private Function TxtGridLeave() As Boolean
Dim i As Integer
Select Case FGrid.Col
    Case Prefix, ManAuto
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
    Case StartNo
        If DataAddMode = False And Val(TxtGrid(0).Text) < Val(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) Then
            MsgBox "Tou Can't Fill Less Value than the existing Value"
            TxtGridLeave = False
            Exit Function
        End If
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).Text), "0")
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).Text), "0")
    Case DateFrom
        TxtGrid(0) = RetDate(TxtGrid(0))
        If TxtGrid(0) <> "" Then
            If DataAddMode = True Then
                Set GRs = GCnFa.Execute("select * from Voucher_Prefix where date_From=#" & TxtGrid(0) & "# and V_Type ='" & FGrid.TextMatrix(FGrid.Row, Vtype) & "' and Div_Code='" & PubDivCode & "'")
                If GRs.RecordCount = 0 Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
                Else
                    MsgBox "Voucher Prefix Already Exist For This Date", vbInformation, "Validation Check": TxtGridLeave = False: Exit Function
                End If
            End If
        End If
End Select
TxtGridLeave = True
TxtGrid(0).Visible = False
FGrid.SetFocus
End Function

'******* Fuctions **********
Private Sub BlankText()
End Sub
Private Sub MoveRec()
'On Error GoTo error1
'If Master.RecordCount > 0 Then
'        FGrid.Rows = 1
'        Do Until Master.EOF
'            FGrid.AddItem "" & Chr(9) & Master!SpotName & Chr(9) & Master!ItemName & Chr(9) & Master!GroupName & Chr(9) & Format(Master!Qty, "0.000") & Chr(9) & Master!Pcs & Chr(9) & Master!IssDate & Chr(9) & Master!RecDate & Chr(9) & Master!ItemCode & Chr(9) & Master!ItemGroup
'            Master.MoveNext
'        Loop
'        FGrid.AddItem ""
'        FGrid.FixedRows = 1
'Else
'        FGrid.Rows = 1
'        FGrid.AddItem ""
'        FGrid.FixedRows = 1
'End If
'Grid_Hide
'Exit Sub
'error1:
'        CheckError
End Sub

Private Sub Ini_Grid()
Dim i As Byte
If Master.RecordCount > 0 Then
'Description,DateFrom,Prefix,StartNo,VType
'    FGrid.Redraw = False
    Set FGrid.DataSource = Master
    With FGrid
        .RowHeightMin = PubGridRowHeight
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, Description) = "Description"
        .ColAlignment(Description) = flexAlignLeftCenter
        .ColWidth(Description) = 2500
        
        .TextMatrix(0, ManAuto) = "Auto/Manual"
        .ColAlignment(ManAuto) = flexAlignLeftCenter
        .ColWidth(ManAuto) = 1500
        
        .TextMatrix(0, DIV) = "Div"
        .ColAlignment(DIV) = flexAlignLeftCenter
        .ColWidth(DIV) = 700

        .TextMatrix(0, DateFrom) = "DateFrom"
        .ColAlignment(DateFrom) = flexAlignLeftCenter
        .ColWidth(DateFrom) = 1100

        .TextMatrix(0, Prefix) = "Prefix"
        .ColAlignment(Prefix) = flexAlignLeftCenter
        .ColWidth(Prefix) = 700

        .TextMatrix(0, StartNo) = "StartNo"
        .ColAlignmentFixed(StartNo) = flexAlignRightCenter
        .ColWidth(StartNo) = 700
        
        .TextMatrix(0, Vtype) = "VType"
        .ColAlignment(Vtype) = flexAlignLeftCenter
        .ColWidth(Vtype) = 700
    End With
        FGrid.Redraw = True
        If mAdd = True Then FGrid.AddItem "**"
End If
Exit Sub

End Sub

Private Sub Grid_Hide()
'    If DGItem.Visible = True Then DGItem.Visible = False
'    If DgSpot.Visible = True Then DgSpot.Visible = False
'    If DGGroup.Visible = True Then DGGroup.Visible = False
End Sub
    

