VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmPhysicalStk 
   BackColor       =   &H80000004&
   Caption         =   "Physical Stock Updation"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   10215
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame FrameUpdate 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   2205
      TabIndex        =   4
      Top             =   2265
      Visible         =   0   'False
      Width           =   5685
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   330
         TabIndex        =   10
         Top             =   2475
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblupdate 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   2340
         TabIndex        =   9
         Top             =   1470
         Width           =   1980
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   2340
         TabIndex        =   8
         Top             =   1110
         Width           =   2010
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Updatable Records :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   315
         TabIndex        =   7
         Top             =   1440
         Width           =   2010
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Records :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   810
         TabIndex        =   6
         Top             =   1095
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   3570
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Process Status"
         Height          =   270
         Left            =   2025
         TabIndex        =   5
         Top             =   330
         Width           =   1665
      End
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
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
      Height          =   240
      HideSelection   =   0   'False
      Left            =   3750
      TabIndex        =   2
      Top             =   1470
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox txtGrid 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   1290
      Visible         =   0   'False
      Width           =   990
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   6645
      Left            =   135
      TabIndex        =   3
      Top             =   600
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   11721
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   13623520
      BackColorBkg    =   13623520
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
End
Attribute VB_Name = "frmPhysicalStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TAddMode As Boolean
'Grid Initializations
Private Const Col_SrNo As Byte = 0
Private Const Col_PartNo As Byte = 1
Private Const Col_PhyStk As Byte = 2
Private Const Col_TPStk As Byte = 3
Private Const Col_TBStk As Byte = 4
Private Const Col_PartName As Byte = 5
Dim Master As ADODB.Recordset
Private Sub Ini_Grid()
  With FGrid
        .Rows = 2
        .Cols = 6
        .left = Me.left + 60
        .width = 11700
        .top = 700
        .TextMatrix(0, Col_SrNo) = "S.No"
        .TextMatrix(1, Col_SrNo) = 1
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 800
        .Font.Size = 10
        .Font.Bold = True

        .TextMatrix(0, Col_PartNo) = "Part No"
        .ColAlignment(Col_PartNo) = flexAlignLeftCenter
        .ColWidth(Col_PartNo) = 2500
        .Font.Size = 10
        .Font.Bold = True
        
        .TextMatrix(0, Col_PhyStk) = "Phy.Stock"
        .ColAlignment(Col_PhyStk) = flexAlignLeftCenter
        .ColWidth(Col_PhyStk) = 1500
        .Font.Size = 10
        .Font.Bold = True

        
        .TextMatrix(0, Col_TPStk) = "TP Stock"
        .ColAlignment(Col_TPStk) = flexAlignLeftCenter
        .ColWidth(Col_TPStk) = 1500
        .Font.Size = 10
        .Font.Bold = True

        .TextMatrix(0, Col_TBStk) = "TB Stock"
        .ColAlignment(Col_TBStk) = flexAlignLeftCenter
        .ColWidth(Col_TBStk) = 1500
        .Font.Size = 10
        .Font.Bold = True
        
        .TextMatrix(0, Col_PartName) = "Part Name"
        .ColAlignment(Col_PartName) = flexAlignLeftCenter
        .ColWidth(Col_PartName) = 3500
        .Font.Size = 10
        .Font.Bold = True
    
End With
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = vbKeyS Then
        TopCtrl1_eSave
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        TopCtrl1_eEdit
    ElseIf KeyCode = vbKeyF10 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
Disp_Text False
Set Master = GCn.Execute("Select Part_No,Part_Name,(Cur_MRP_TBStk+Cur_TB_Stk) as TBStk,(Cur_MRP_TPStk+Cur_TP_Stk) as TPStk from Part where Div_Code='" & PubDivCode & "'")
MDIForm1.Picture1.Visible = True
MDIForm1.Label1.CAPTION = "Please Wait ! Filling Part Master's Data. It will take a little time....."
Call FillData
MDIForm1.Picture1.Visible = False
End Sub
Private Sub Disp_Text(Enb As Boolean)
TopCtrl1.tEdit = Not Enb
TopCtrl1.tExit = Not Enb
TopCtrl1.tPrn = Not Enb
TopCtrl1.tSave = Enb

TopCtrl1.tCancel = False
TopCtrl1.tRef = False
TopCtrl1.tAdd = False
TopCtrl1.tFirst = False
TopCtrl1.tNext = False
TopCtrl1.tPrev = False
TopCtrl1.tLast = False
TopCtrl1.tFind = False
TopCtrl1.tDel = False

TxtGrid(0).Visible = False
TxtSearch.Visible = False
FrameUpdate.Visible = False
End Sub

Private Sub TopCtrl1_eEdit()
Disp_Text SETS("EDIT", Me, Master)
FGrid.Row = 1
FGrid.Col = Col_PartNo
FGrid.SetFocus
End Sub

Private Sub TopCtrl1_eSave()
    Dim TotRec As Double, i As Long, Count As Integer
    Dim UpdtRec As Double
    On Error GoTo err
    
    TotRec = FGrid.Rows - 1
    For i = 1 To FGrid.Rows - 1
        If Val(FGrid.TextMatrix(i, Col_PhyStk)) <> 0 Then
            UpdtRec = UpdtRec + 1
        End If
    Next
    'If FGrid.Visible = True Then FGrid.Visible = False: If txtGrid(0).Visible = True Then txtGrid(0).Visible = False: If TxtSearch.Visible = True Then TxtSearch.Visible = False
    FrameUpdate.Visible = True
    lblTotal = TotRec
    lblupdate = UpdtRec
    lblTotal.Refresh: lblupdate.Refresh
    For i = 1 To TotRec
        If Val(FGrid.TextMatrix(i, Col_PhyStk)) <> 0 Then
            Count = Count + 1
            GCn.Execute "Update Part set PhyStk = " & Val(FGrid.TextMatrix(i, Col_PhyStk)) & " where Part_No='" & FGrid.TextMatrix(i, Col_PartNo) & "' and Div_Code='" & PubDivCode & "'"
            ProgressBar1.Value = Round((Count * 100) / UpdtRec, 2)
        End If
    Next
    
    Unload Me
    
err:
    CheckError
    
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop

TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then TxtGrid(0).TEXT = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case Col_PhyStk
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown)) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PhyStk - 1, 1, 1
                    FGrid.CellBackColor = vbWhite
                Else
                    TxtGrid(0).SetFocus
                End If
            End If
        Case Col_Part_No
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown)) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PhyStk - 1, 1, 2
                    FGrid.CellBackColor = vbWhite
                Else
                    TxtGrid(0).SetFocus
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = vbKeyEscape Then Exit Sub
CheckQuote KeyAscii
Select Case FGrid.Col
    Case Col_PhyStk
        NumPress TxtGrid(Index), KeyAscii, 6, 2
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case 0
    Select Case FGrid.Col
        Case Col_PhyStk
            FGrid.TextMatrix(FGrid.Row, Col_PhyStk) = Format(Val(TxtGrid(Index).TEXT), "0.00")
    End Select
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)

Exit Sub
ELoop:
    CheckError

End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case Col_PhyStk
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
End Select
    TxtGridLeave = True
    'Important at the time of validating  a control if you are making the visibility of
    'control false forcefully it will generate error
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.Col = Col_PhyStk Then
    FGrid_KeyPress vbKeyReturn
End If
TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
On Error GoTo ELoop
SetMaxLength
    Select Case FGrid.Col
        Case Col_PhyStk
            Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case Col_PartNo
            SelGridKeyPress TxtSearch, FGrid, Master, KeyAscii, "Part_No", CellBackColEnter, vbWhite
        Case Col_PartName
            SelGridKeyPress TxtSearch, FGrid, Master, KeyAscii, "Part_Name", CellBackColEnter, vbWhite
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtSearch_Click()
 FGrid.SetFocus: TxtSearch.TEXT = "": TxtSearch.Visible = False
End Sub
Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If NavigationKey(KeyCode) = True Then FGrid.SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then FGrid.Col = Col_PhyStk: FGrid.SetFocus: TxtSearch.Visible = False

End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
Select Case FGrid.Col
    Case Col_PartNo
        SelGridKeyPress TxtSearch, FGrid, Master, KeyAscii, "Part_No", CellBackColEnter, vbWhite: KeyAscii = 0
    Case Col_PartName
        SelGridKeyPress TxtSearch, FGrid, Master, KeyAscii, "Part_Name", CellBackColEnter, vbWhite: KeyAscii = 0
End Select
End Sub

Private Sub TxtSearch_LostFocus()
     FGrid.SetFocus: TxtSearch.Visible = False:: TxtSearch.TEXT = "": FGrid.CellBackColor = vbWhite
End Sub
Private Sub TxtGridValid_PNo()
'Called from TxtGrid_Validate & TxtGridLeave procedures
End Sub
Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
End Sub
Private Sub TopCtrl1_eExit()
Unload Me
End Sub
Private Sub SetMaxLength()
    Select Case FGrid.Col
        Case Col_PhyStk
            TxtGrid(0).MaxLength = 100
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Else
            TxtGrid(0).MaxLength = 0
    End Select
End Sub

Private Sub FGrid_GotFocus()
    TxtSearch.TEXT = ""
    TxtGrid(0).Visible = False
End Sub
Private Function FillData()
Master.MoveFirst
FGrid.Font.Size = 8
FGrid.Font.Bold = False
FGrid.ForeColor = vbRed
For i = 1 To Master.RecordCount
    FGrid.TextMatrix(i, Col_SrNo) = i
    FGrid.TextMatrix(i, Col_PartNo) = Master!Part_No
    FGrid.TextMatrix(i, Col_PartName) = Master!Part_Name
    FGrid.TextMatrix(i, Col_TPStk) = Master!TPStk
    FGrid.TextMatrix(i, Col_TBStk) = Master!TBStk
    Master.MoveNext: FGrid.Refresh: FGrid.Rows = FGrid.Rows + 1
Next
End Function
