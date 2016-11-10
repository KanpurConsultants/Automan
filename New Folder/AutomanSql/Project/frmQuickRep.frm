VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQuickRep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Report View"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11910
   Begin VB.Frame FrmCondition 
      BackColor       =   &H00808080&
      Height          =   2220
      Left            =   4695
      TabIndex        =   7
      Top             =   4710
      Width           =   4485
      Begin VB.CommandButton CmdShow 
         Caption         =   "Show"
         Height          =   450
         Left            =   1575
         TabIndex        =   14
         Top             =   1650
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   270
         Left            =   1890
         TabIndex        =   13
         Top             =   1200
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   476
         _Version        =   393216
         Format          =   51183617
         CurrentDate     =   38624
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1905
         TabIndex        =   12
         Top             =   735
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   503
         _Version        =   393216
         Format          =   51183617
         CurrentDate     =   38624
      End
      Begin VB.ComboBox CmbOrg 
         Height          =   315
         ItemData        =   "frmQuickRep.frx":0000
         Left            =   1890
         List            =   "frmQuickRep.frx":0016
         TabIndex        =   11
         Text            =   "CmbOrg"
         Top             =   285
         Width           =   2220
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Up To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Originate From "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   330
         Width           =   1410
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Condition"
      Height          =   390
      Left            =   3225
      TabIndex        =   6
      Top             =   6495
      Width           =   1485
   End
   Begin VB.CommandButton CmdFilter 
      Caption         =   "Filter"
      Height          =   390
      Index           =   0
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Apply Filter"
      Top             =   6495
      Width           =   975
   End
   Begin VB.CommandButton CmdFilter 
      Caption         =   "Remove Filter"
      Height          =   390
      Index           =   1
      Left            =   1020
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Remove Filter"
      Top             =   6495
      Width           =   1125
   End
   Begin VB.CommandButton CmdFilter 
      Caption         =   "Sort"
      Height          =   390
      Index           =   2
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sort Ascending"
      Top             =   6495
      Width           =   1020
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      HideSelection   =   0   'False
      Left            =   345
      TabIndex        =   1
      Top             =   675
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   6510
      Left            =   45
      TabIndex        =   0
      Top             =   -60
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   11483
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   128
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   16777088
      ForeColorSel    =   128
      BackColorBkg    =   16777215
      GridColor       =   64
      GridColorFixed  =   65280
      GridColorUnpopulated=   16711935
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   -30
      TabIndex        =   2
      Top             =   6480
      Width           =   11835
   End
   Begin VB.Menu popup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu FSR 
         Caption         =   "Filter Same Records"
         Index           =   0
      End
      Begin VB.Menu RF 
         Caption         =   "Remove Filter"
         Index           =   1
      End
      Begin VB.Menu RC 
         Caption         =   "Remove Column"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmQuickRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Master As ADODB.Recordset
Private Const CellBackColEnter1 = &HFFFFC0
Private Const CellBackColLeave1 = &HEDF7FE
Public ColumnWidth   '() As Integer

Private Sub CmdFilter_Click(Index As Integer)
Dim SortStr As String
Dim I As Integer
    Select Case Index
        Case 0
            Master.Filter = adFilterNone
            Master.Filter = Master.Fields(GridSel.Col).Name & " = '" & GridSel.TextMatrix(GridSel.Row, GridSel.Col) & "'"
            
            GridSel.Refresh
            Set GridSel.DataSource = Master
            GridSel.Col = 1
            GridSel.SetFocus
            
        Case 1
            Master.Filter = adFilterNone
            GridSel.Refresh
            GridSel.Col = 1
            GridSel.SetFocus
        Case 2
                If GridSel.Col = 0 Then Exit Sub
'                For I = GridSel.Col To 1 Step -1
'                    SortStr = SortStr + Master.Fields(I).Name + ","
'                Next
'                SortStr = Mid(SortStr, 1, Len(SortStr) - 1)
                GridSel.CellBackColor = CellBackColEnter1
'                Master.Sort = SortStr   'Master.Fields(GridSel.Col).Name
                Master.Sort = Master.Fields(GridSel.Col).Name
                GridSel.Refresh
    End Select
End Sub

Private Sub Command1_Click()
FrmCondition.Visible = True
End Sub
Private Sub CmdShow_Click()
Dim MyQry As String
FrmCondition.Visible = False
GridSel.Visible = True
Select Case CmbOrg
    Case "Counter Sale"
        MyQry = "Select V_Type,V_No,V_Date,Cash_Credit,Party_Name,Address,L_C,format(SprAmt_MRP_TB,'0.00') as SprAmt_MRP_TB,format(SprAmt_MRP_TP,'0.00') as SprAmt_MRP_TP,OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB,SprAmt_TP,OilAmt_TB,OilAmt_TP,D_Per_TB,D_Amt_TB,D_Per_TP,D_Amt_TP,Addition,Tax_Per,Tax_Amt,Tax_AmtMRP,Tax_Sur_Per,Tax_Sur_Amt,TaxSur_AmtMRP,TOT_Per,TOT_Amt,TOT_AmtMRP,Total_Amt,Rounded from SP_Sale where V_type in ('SYSIC','SYSIR')"
        MyQry = MyQry + " and V_Date >=" & ConvertDate(DTPicker1.Value) & " and V_Date <=" & ConvertDate(DTPicker2.Value) & ""
    Case "Workshop Sale"
        MyQry = "Select V_Type,V_No,V_Date,Cash_Credit,Party_Name,Address,L_C,SprAmt_MRP_TB,SprAmt_MRP_TP,OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB,SprAmt_TP,OilAmt_TB,OilAmt_TP,D_Per_TB,D_Amt_TB,D_Per_TP,D_Amt_TP,Addition,Tax_Per,Tax_Amt,Tax_AmtMRP,Tax_Sur_Per,Tax_Sur_Amt,TaxSur_AmtMRP,TOT_Per,TOT_Amt,TOT_AmtMRP,Total_Amt,Rounded from SP_Sale where V_type in ('W_SIC','W_SIR')"
        MyQry = MyQry + " and V_Date >=" & ConvertDate(DTPicker1.Value) & " and V_Date <=" & ConvertDate(DTPicker2.Value) & ""
    Case "History Card"
        MyQry = "Select * from HisCard"
    Case "Job Card"
        MyQry = "Select * from job_Card"
        MyQry = MyQry + " where Job_Date >=" & ConvertDate(DTPicker1.Value) & " and Job_Date <=" & ConvertDate(DTPicker2.Value) & ""
End Select
Set Master = GCn.Execute(MyQry)
MultiComp = False
Set GridSel.DataSource = Master
GridSel.ColWidth(0) = 0
If Master.RecordCount > 0 Then
    For aa = 1 To Master.Fields.Count - 1
        If Master.Fields(aa).Type = adNumeric Then
            GridSel.ColWidth(aa) = 1100
            GridSel.ColAlignment(aa) = vbRightJustify
        Else
            If Master.Fields(aa).ActualSize = 0 Then
                GridSel.ColWidth(aa) = Master.Fields(aa).DefinedSize * 500
            Else
                GridSel.ColWidth(aa) = Master.Fields(aa).ActualSize * 120
                GridSel.ColAlignment(aa) = flexAlignRightCenter
            End If
        End If
    Next
End If
If GridSel.Rows = 1 Then GridSel.AddItem ""
GridSel.Col = 1
GridSel_GotFocus
IniGrid Master
Me.CAPTION = CmbOrg.TEXT
'GridSel.Col = 1
'GridSel_EnterCell
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SortStr As String
Dim I As Integer
If KeyCode = vbKeyEscape And TxtSearch.Visible = False Then Unload Me
'If KeyCode = vbKeyS And Shift = 2 Then
'    For I = GridSel.Col To 1 Step -1
'        SortStr = SortStr + Master.Fields(I).Name + ","
'    Next
'    SortStr = Mid(SortStr, 1, Len(SortStr) - 1)
'    GridSel.CellBackColor = CellBackColEnter1
'    Master.Sort = SortStr   'Master.Fields(GridSel.Col).Name
'    GridSel.Refresh
'End If
End Sub

Private Sub Form_Load()

    Check = ""
    Me.width = 11940: Me.left = 0: Me.top = 0
'    Me.width = 8355: Me.top = 960: Me.left = 0
    Label1.width = Me.width
    FrmCondition.Visible = False
    GridSel.Visible = False

End Sub

Private Sub FSR_Click(Index As Integer)
Master.Filter = Master.Fields(GridSel.Col).Name & " =  '" & GridSel.TextMatrix(GridSel.Row, GridSel.Col) & "'"
GridSel.Refresh
GridSel.Col = 1
GridSel.SetFocus
End Sub
Private Sub GridSel_DblClick()
If GridSel.TextMatrix(1, 1) <> "" Then
'If Master.AbsolutePosition <> adPosBOF And Master.AbsolutePosition <> adPosEOF And Master.AbsolutePosition <> adPosUnknown Then
    SearchForm.SEARCHBACK GridSel.TextMatrix(GridSel.Row, 0)
    Check = GridSel.TextMatrix(GridSel.Row, 0)
'    Check = Master!SearchCode
Else
    Check = ""
End If
Unload Me
Exit Sub
errorbox:
    MsgBox err.Description, vbInformation
End Sub

Private Sub GridSel_GotFocus()
GridSel.CellBackColor = CellBackColEnter1
End Sub

Private Sub GridSel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then GridSel_DblClick
End Sub
Private Sub GridSel_KeyPress(KeyAscii As Integer)
FaSelGridKeyPress TxtSearch, GridSel, Master, KeyAscii, Master.Fields(GridSel.Col).Name, CellBackColEnter1, CellBackColLeave1
End Sub
Private Sub GridSel_LeaveCell()
GridSel.CellBackColor = CellBackColLeave1
End Sub
Private Sub GridSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu popup
End If
End Sub

Private Sub RC_Click(Index As Integer)
GridSel.ColWidth(GridSel.Col) = 0
End Sub

Private Sub RF_Click(Index As Integer)
Master.Filter = adFilterNone
GridSel.Refresh
GridSel.Col = 1
GridSel.SetFocus
End Sub
Private Sub TxtSearch_Click()
TxtSearch.TEXT = "": GridSel.SetFocus: TxtSearch.Visible = False
End Sub
Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If FaNavigationKey(KeyCode) = True Then GridSel.SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then GridSel.SetFocus: TxtSearch.Visible = False
End Sub
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
FaSelGridKeyPress TxtSearch, GridSel, Master, KeyAscii, Master.Fields(GridSel.Col).Name, CellBackColEnter1, CellBackColLeave1: KeyAscii = 0
End Sub
Private Sub TxtSearch_LostFocus()
    TxtSearch.TEXT = "": GridSel.SetFocus: TxtSearch.Visible = False
End Sub

Private Sub IniGrid(Master As ADODB.Recordset)
Dim aa As Integer
Dim ColCnt As Integer
GridSel.ColWidth(0) = 0
If Master.RecordCount > 0 Then
    For aa = 1 To Master.Fields.Count - 1
        If Master.Fields(aa).Type = adNumeric Then
            GridSel.ColWidth(aa) = 1100
        Else
            If Master.Fields(aa).ActualSize = 0 Then
                GridSel.ColWidth(aa) = Master.Fields(aa).DefinedSize * 50
            Else
                GridSel.ColWidth(aa) = Master.Fields(aa).ActualSize * 120
            End If
       End If
    Next
End If
'If UBound(ColumnWidth) = 0 Then Exit Sub

'ColCnt = UBound(ColumnWidth) + 1 'ColumnWidth(0)
'GridSel.Cols = ColCnt
'For I = 1 To ColCnt - 1
'    GridSel.ColWidth(I) = ColumnWidth(I)
'Next
End Sub
