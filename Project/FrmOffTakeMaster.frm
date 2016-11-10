VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmOffTakeMaster 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Offtake Incentive Master"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
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
   ScaleHeight     =   5745
   ScaleWidth      =   9480
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDF4B5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   3015
      TabIndex        =   5
      Top             =   2745
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   2685
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   7020
      TabIndex        =   11
      Top             =   2475
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   0
         TabIndex        =   12
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
      Height          =   210
      Index           =   2
      Left            =   900
      TabIndex        =   10
      Top             =   5565
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   1605
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1080
      Width           =   1605
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   4
      Left            =   1815
      MaxLength       =   20
      TabIndex        =   7
      Top             =   4995
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   5
      Left            =   900
      MaxLength       =   20
      TabIndex        =   8
      Top             =   5805
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   6
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   1605
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   7
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1560
      Width           =   1605
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   8
      Left            =   870
      MaxLength       =   20
      TabIndex        =   9
      Top             =   5325
      Visible         =   0   'False
      Width           =   1605
   End
   Begin MSDataGridLib.DataGrid DgOffTake 
      Height          =   2100
      Left            =   4380
      Negotiate       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4575
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3704
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
         DataField       =   "SchemeNo"
         Caption         =   "Scheme No"
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
            ColumnWidth     =   3314.835
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DgModelGrp 
      Height          =   2190
      Left            =   2340
      Negotiate       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4185
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3863
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
         Caption         =   "Model Group"
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
            ColumnWidth     =   3314.835
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2430
      Left            =   1695
      TabIndex        =   6
      Top             =   2130
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   4286
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   0
      Cols            =   5
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   192
      GridColorUnpopulated=   16761024
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   "SrNo."
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
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheme No"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1710
      TabIndex        =   24
      Top             =   615
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   -870
      TabIndex        =   23
      Top             =   5580
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effective From Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1710
      TabIndex        =   22
      Top             =   855
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Till Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1710
      TabIndex        =   21
      Top             =   1095
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Category"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   45
      TabIndex        =   20
      Top             =   5010
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   -870
      TabIndex        =   19
      Top             =   5820
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target Qty"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   1710
      TabIndex        =   18
      Top             =   1335
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incentive Amount"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   1710
      TabIndex        =   17
      Top             =   1575
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Subvention"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   -900
      TabIndex        =   16
      Top             =   5340
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "FrmOffTakeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSite As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsModelCategory As ADODB.Recordset
Dim RsModel As ADODB.Recordset
Dim RsModelGrp As ADODB.Recordset




Private Const SchemeNo As Byte = 0
Private Const FromDate As Byte = 1
Private Const ToDate As Byte = 3
Private Const ModelCategory As Byte = 4
Private Const Model As Byte = 5
Private Const Qty As Byte = 6
Private Const Amount As Byte = 7
Private Const SType As Byte = 2



Dim EditName As String
Dim EditDesc As String
Dim ListArray As Variant
Dim mListItem As ListItem


Private Const ModelGrp_Name As Byte = 1
Private Const ModelGrp_Code As Byte = 2

Dim ForeColorSelEnter$
Dim BackColorSelLeave$


Dim GridKey As Integer
Dim TAddMode As Boolean



Private Sub Ini_Grid()
    With FGrid
        .Cols = 3
        .RowHeightMin = PubGridRowHeight
        .height = .RowHeight(0) * 10
        
        .TextMatrix(0, 0) = "S.No."
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 500

        .TextMatrix(0, ModelGrp_Name) = "ModelGrp_Name"
        .ColAlignment(ModelGrp_Name) = flexAlignLeftCenter
        .ColWidth(ModelGrp_Name) = 4200
        
        .TextMatrix(0, ModelGrp_Code) = "ModelGrp_Code"
        .ColAlignment(ModelGrp_Code) = flexAlignLeftCenter
        .ColWidth(ModelGrp_Code) = 0
        
                
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    DgOffTake.left = Txt(SchemeNo).left: DgOffTake.top = Txt(SchemeNo).top + Txt(SchemeNo).height + 20
End Sub
Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
    FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case ModelGrp_Name
            FGrid = ""
            FGrid.TextMatrix(FGrid.Row, ModelGrp_Code) = ""
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case ModelGrp_Name
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)

Select Case FGrid.Col
    Case ModelGrp_Name
        Get_Text Me, FGrid, TxtGrid, 0, False, keyascii
End Select
If keyascii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.Rows > 2 Then
                FGrid.RemoveItem (FGrid.Row)
            Else
                FGrid.Rows = 1
                FGrid.AddItem FGrid.Rows
                FGrid.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid.Rows - 1
            FGrid.TextMatrix(I, 0) = I
         Next
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
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub



Private Sub ListView_Click()
On Error GoTo ELoop
    Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    Txt(Val(ListView.Tag)).SetFocus
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
    WinSetting Me, 6000, 8715: Ini_Grid
    Me.Icon = MDIForm1.Icon
    
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select Code as SearchCode,* from OffTake order by SchemeNo", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 Code as SearchCode,* from OffTake order by SchemeNo", GCn, adOpenDynamic, adLockOptimistic
    End If
   
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select Code, SchemeNo from OffTake order by SchemeNo", GCn, adOpenDynamic, adLockOptimistic
    Set DgOffTake.DataSource = RsSite
                        
    Set RsModelGrp = GCn.Execute("Select ModelGrp_Code As Code, ModelGrp_Name as Name From Model_Grp Order By ModelGrp_Name")
    Set DGModelGrp.DataSource = RsModelGrp
                
    Disp_Text SETS("INI", Me, Master)
    MoveRec
  Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsSite = Nothing
Set Master = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Txt(SchemeNo).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                GCn.Execute ("delete from OffTake where Code= '" & Master!SearchCode & "'")
                GCn.Execute ("delete from OffTake1 where Code= '" & Master!SearchCode & "'")
                GCn.CommitTrans
                Master.Requery
                Call MoveRec
                RsSite.Requery
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
    EditName = Txt(SchemeNo).TEXT
    EditDesc = Txt(FromDate).TEXT
    Txt(SchemeNo).SetFocus
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
    GSQL = "select Code as Searchcode,SchemeNo, FromDate, TODate  from OffTake S  order by SchemeNo"
    Set SearchForm = Me
    FIND.Show vbModal
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
        Set Master = GCn.Execute("select Code as SearchCode,* from OffTake Where Code = '" & MyValue & "' order by SchemeNo")
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

Private Sub TopCtrl1_ePrn()
'Dim I As Integer, mQRY$, mRepName$
'Dim Rst As ADODB.Recordset
'On Error GoTo ERRORHANDLER
'
'    mRepName = "OffTake"
'    mQRY = "SELECT Code, SchemeNo, FromDate, ToDate, MG.ModelGrp_Name, Qty, Amount  " & _
'           "From OffTake O  Left Join  Left Join Model_Grp MG On .ModelCategory=MG.ModelCat_Code Order by SchemeNo"
'
'    Set Rst = New Recordset
'    Rst.CursorLocation = adUseClient
'    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
'    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
'    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
'    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
'     rpt.Database.SetDataSource Rst
'     rpt.ReadRecords
'    Call Report_View(rpt, Me.CAPTION, , True)
'    Set Rst = Nothing
'Exit Sub
'ERRORHANDLER:
'      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eRef()
    RsSite.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mCnt As Integer
    Dim mMaxId As Long
    Dim mTrans As Boolean
    Dim ItemCode As Integer
    Dim Rst As ADODB.Recordset
'    On Error GoTo errlbl
   
     If IsValid(Txt(SchemeNo), "Scheme No") = False Then Exit Sub
     If IsValid(Txt(FromDate), "From Date") = False Then Exit Sub
     If IsValid(Txt(ToDate), "ToDate") = False Then Exit Sub
     
     

    If TopCtrl1.TopText2 = "Add" Or (TopCtrl1.TopText2 = "Edit" And UCase(Txt(SchemeNo).TEXT) <> UCase(EditName)) Then
       Set Rst = New ADODB.Recordset
       Set Rst = GCn.Execute("select * from OffTake where SchemeNo = '" & Txt(SchemeNo) & "' And FromDate=" & ConvertDate(Txt(FromDate)) & " and Todate= " & ConvertDate(Txt(ToDate)) & "")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate SchemeNo", vbInformation, "Validation Check": Txt(SchemeNo).SetFocus: Exit Sub
            End If
        Set Rst = Nothing
    End If
    
    
    mCnt = 0
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ModelGrp_Name) <> "" Then
            mCnt = mCnt + 1
        End If
    Next I
    If mCnt = 0 Then MsgBox "Please Select ModelGroup First": FGrid.SetFocus: Exit Sub
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        Txt(SchemeNo).Tag = GCn.Execute("Select " & vIsNull("Max(Code)", "0") & " +1 From OffTake ").Fields(0)
        GCn.Execute ("insert into OffTake(Code, SchemeNo, FromDate, ToDate,  Qty, Amount, U_Name,U_EntDt,U_AE) " & _
            " values(" & Val(Txt(SchemeNo).Tag) & ", '" & Txt(SchemeNo) & "' ," & ConvertDate(Txt(FromDate)) & "," & ConvertDate(Txt(ToDate)) & ", " & Val(Txt(Qty)) & ", " & Val(Txt(Amount)) & ", '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
            
        GCn.Execute "Delete From OffTake1 Where Code=" & Val(Txt(SchemeNo).Tag) & ""
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ModelGrp_Name) <> "" Then
                GCn.Execute "Insert Into Offtake1(Code, SrlNo, ModelGrp) Values(" & Val(Txt(SchemeNo).Tag) & ", " & I & ", '" & FGrid.TextMatrix(I, ModelGrp_Code) & "')"
            End If
        Next I
            
    Else
        GCn.Execute "update OffTake  set SchemeNo='" & Txt(SchemeNo) & "', FromDate=" & ConvertDate(Txt(FromDate)) & ", ToDate=" & ConvertDate(Txt(ToDate)) & ", Qty=" & Val(Txt(Qty)) & ", Amount=" & Val(Txt(Amount)) & ", U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "' Where Code= " & Val(Txt(SchemeNo).Tag) & ""
        
        GCn.Execute "Delete From OffTake1 Where Code=" & Val(Txt(SchemeNo).Tag) & ""
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ModelGrp_Name) <> "" Then
                GCn.Execute "Insert Into Offtake1(Code, SrlNo, ModelGrp) Values(" & Txt(SchemeNo).Tag & ", " & I & ", '" & FGrid.TextMatrix(I, ModelGrp_Code) & "')"
            End If
        Next I
        
    End If
    
    
    
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select Code as SearchCode,* from OffTake Where Code = '" & Val(Txt(SchemeNo).Tag) & "' order by SchemeNo")
    End If
    RsSite.Requery
    
    Master.FIND "searchcode = " & Val(Txt(SchemeNo).Tag) & " "
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
            
        Case SType
            ListArray = Array("HO", "Branch")
            Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
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
    Case SType
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 600
    Case SchemeNo
        DGridTxtKeyDown_Mast DgOffTake, Txt, Index, RsSite, KeyCode, False, 0
End Select
If DgOffTake.Visible = False And DGModelGrp.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If Index <> SchemeNo Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
    Select Case Index
        Case ModelCategory
            DGridTxtKeyPress Txt, Index, RsModelCategory, keyascii, "Name"
        Case Model
            DGridTxtKeyPress Txt, Index, RsModel, keyascii, "Name"
    End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case SchemeNo
        DGridTxtKeyUp_Mast Txt, Index, RsSite, KeyCode, "SchemeNo"
    Case SType
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
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

FGrid.Rows = 1
FGrid.AddItem ""
FGrid.FixedRows = 1
End Sub

Private Sub MoveRec()
Dim RsTemp As ADODB.Recordset
Dim I As Integer

On Error GoTo error1
If Master.RecordCount > 0 Then
    Txt(SchemeNo).Tag = Master!SearchCode
    Txt(SchemeNo) = Master!SchemeNo
    Txt(FromDate) = XNull(Master!FromDate)
    Txt(ToDate) = XNull(Master!ToDate)
    Txt(Qty) = Format(VNull(Master!Qty), "0.00")
    Txt(Amount) = Format(VNull(Master!Amount), "0.00")

    FGrid.Rows = 1
    I = 1
    Set RsTemp = GCn.Execute("Select O.Code, O.ModelGrp, M.ModelGrp_Name From OffTake1 O Left Join Model_Grp M On  M.ModelGrp_Code=O.ModelGrp  Where Code = " & Val(Txt(SchemeNo).Tag) & "  Order By SrlNo")
    If RsTemp.RecordCount > 0 Then
        Do Until RsTemp.EOF
            FGrid.AddItem ""
            
            FGrid.TextMatrix(I, 0) = I
            FGrid.TextMatrix(I, ModelGrp_Name) = XNull(RsTemp!ModelGrp_Name)
            FGrid.TextMatrix(I, ModelGrp_Code) = XNull(RsTemp!ModelGrp)
            
            I = I + 1
            RsTemp.MoveNext
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem ""
        FGrid.FixedRows = 1
    End If

End If


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
    If DgOffTake.Visible = True Then DgOffTake.Visible = False
    If DGModelGrp.Visible = True Then DGModelGrp.Visible = False
    
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
Select Case Index
    Case ModelCategory
        If RsModelCategory.RecordCount > 0 And RsModel.EOF = False And RsModel.BOF = False And Txt(Index) <> "" Then
            Txt(Index) = RsModelCategory!Name
            Txt(Index).Tag = RsModelCategory!Code
        Else
            Txt(Index) = ""
            Txt(Index).Tag = ""
        End If
    Case Model
        If RsModel.RecordCount > 0 And RsModel.EOF = False And RsModel.BOF = False And Txt(Index) <> "" Then
            Txt(Index) = RsModel!Name
            Txt(Index).Tag = RsModel!Code
        Else
            Txt(Index) = ""
            Txt(Index).Tag = ""
        End If
        
    Case Qty, Amount
        Txt(Index) = Format(Txt(Index), "0.00")
        
    Case FromDate, ToDate
        Txt(Index) = RetDate(Txt(Index))
    Case SType
            If Txt(Index).TEXT <> "" Then Txt(Index).TEXT = ListView.SelectedItem.TEXT
End Select
End Sub





Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Grid_Hide
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Select Case FGrid.Col
    Case ModelGrp_Name
        DGModelGrp.Move TxtGrid(0).left, TxtGrid(0).top + TxtGrid(0).height + 20
        If RsModelGrp.RecordCount = 0 Or (RsModelGrp.EOF = True Or RsModelGrp.BOF = True) Then Exit Sub
        If FGrid.TextMatrix(FGrid.Row, ModelGrp_Name) <> "" Then
            RsModelGrp.MoveFirst
            RsModelGrp.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, ModelGrp_Name) & "'"
            If RsModelGrp.EOF = True Then RsModelGrp.MoveFirst
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
    Exit Sub
End If
Select Case FGrid.Col
    Case ModelGrp_Name
        DGridTxtKeyDown DGModelGrp, TxtGrid, Index, RsModelGrp, KeyCode, True, 1
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, ModelGrp_Name, 1
            End If
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case FGrid.Col
    Case ModelGrp_Name
        If DGModelGrp.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsModelGrp, keyascii, "name"
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case ModelGrp_Name
        If KeyCode <> 13 And DGModelGrp.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsModelGrp, KeyCode, "name", True
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
Select Case FGrid.Col
    Case ModelGrp_Name
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        TxtGridValid_ModelGrp_Name
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

Private Function ChkDuplicate() As Boolean
Dim I As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte
Select Case FGrid.Col
    Case ModelGrp_Name
        Col1 = ModelGrp_Code
        Col2 = ModelGrp_Name
    End Select
    X = UCase(Trim(TxtGrid(0).TEXT))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, FGrid.Col))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
            TxtGrid(0).SetFocus
            Ctrl_GetFocus TxtGrid(0)
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function


Private Sub TxtGridValid_ModelGrp_Name()
If RsModelGrp.RecordCount = 0 Or (RsModelGrp.EOF = True Or RsModelGrp.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, ModelGrp_Code) = ""
    FGrid.TextMatrix(FGrid.Row, ModelGrp_Name) = ""
Else
    FGrid.TextMatrix(FGrid.Row, ModelGrp_Code) = IIf(IsNull(RsModelGrp!Code), "", RsModelGrp!Code)
    FGrid.TextMatrix(FGrid.Row, ModelGrp_Name) = IIf(IsNull(RsModelGrp!Name), "", RsModelGrp!Name)
End If
If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

