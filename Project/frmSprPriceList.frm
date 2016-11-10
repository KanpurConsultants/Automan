VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSprPriceList 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Spare Price List"
   ClientHeight    =   8115
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   1845
      TabIndex        =   18
      Top             =   6525
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   90
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   15
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3942
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
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
      HideSelection   =   0   'False
      Index           =   2
      Left            =   2325
      TabIndex        =   3
      Top             =   1470
      Width           =   1485
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
      HideSelection   =   0   'False
      Index           =   0
      Left            =   2325
      TabIndex        =   1
      Top             =   930
      Width           =   1485
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
      Left            =   8955
      TabIndex        =   11
      Top             =   1110
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox TxtGrid1 
      Alignment       =   1  'Right Justify
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
      Left            =   9810
      TabIndex        =   7
      Top             =   645
      Visible         =   0   'False
      Width           =   690
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
      HideSelection   =   0   'False
      Index           =   1
      Left            =   2325
      TabIndex        =   2
      Top             =   1200
      Width           =   1485
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
      Alignment       =   1  'Right Justify
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
      Left            =   1695
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   4650
      Left            =   180
      TabIndex        =   4
      Top             =   2475
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   8202
      _Version        =   393216
      BackColor       =   15525079
      Cols            =   8
      BackColorFixed  =   14940925
      ForeColorFixed  =   16576
      BackColorSel    =   16711680
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14865856
      GridColor       =   14940925
      GridColorFixed  =   12632319
      FocusRect       =   0
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1410
      Left            =   5565
      TabIndex        =   6
      Top             =   690
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   2487
      _Version        =   393216
      BackColor       =   15525079
      Cols            =   8
      BackColorFixed  =   14940925
      ForeColorFixed  =   16576
      BackColorSel    =   16711680
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14865856
      GridColor       =   14940925
      GridColorFixed  =   12632319
      FocusRect       =   0
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSDataGridLib.DataGrid DGDate 
      Height          =   2910
      Left            =   6180
      Negotiate       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Date Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Salect Date"
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
         EndProperty
      EndProperty
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   1
      Left            =   2040
      TabIndex        =   16
      Top             =   1455
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   2
      Left            =   720
      TabIndex        =   15
      Top             =   1485
      Width           =   570
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   0
      Left            =   2040
      TabIndex        =   14
      Top             =   1185
      Width           =   120
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   3
      Left            =   2040
      TabIndex        =   13
      Top             =   915
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   1
      Left            =   720
      TabIndex        =   12
      Top             =   945
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   37
      Left            =   720
      TabIndex        =   10
      Top             =   1215
      Width           =   1110
   End
   Begin VB.Label LBLListDesc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Spare Price List  General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Index           =   0
      Left            =   165
      TabIndex        =   9
      Top             =   2130
      Width           =   11460
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set Discount % For Party"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Index           =   9
      Left            =   5580
      TabIndex        =   8
      Top             =   390
      Width           =   6045
   End
End
Attribute VB_Name = "frmSprPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Master As ADODB.Recordset
Dim Master1 As ADODB.Recordset
Dim mAdd As Boolean, mEdit As Boolean, mDel As Boolean, mPrn As Boolean
Dim RsDate  As ADODB.Recordset
Dim AddMode As Byte
Dim PartyTypeCode As Byte
Private Const CellBackColLeave As String = &HECE4D7    '&HECE4D7   '&HEDF7FE
Private Const CellForeColLeave As String = &HFF00FF
Private Const CellBackColEnter As String = &HF0D5BF
Private Const GridBackColorBkg As String = &HE2D5C0
Private Const RateType As Byte = 0
Private Const VDate As Byte = 1
Private Const RefNo As Byte = 2
Private Const PartyType As Byte = 3

Private Const Col_PNo As Byte = 1
Private Const Col_PName As Byte = 2
Private Const Col_Unit As Byte = 3
Private Const Col_Grade = 4
Private Const Col_PDisc = 5
Private Const Col_ISalDisc = 6
Private Const Col_Origin As Byte = 7
Private Const Col_MRP As Byte = 8
Private Const Col_Taxable As Byte = 9
Private Const Col_TaxPaid As Byte = 10
Private Const Col_MRPEffectDt As Byte = 11
Private Const Col_TBEffectDt As Byte = 12

Private Const Col1_Desc As Byte = 1
Private Const Col1_DiscMRP As Byte = 2
Private Const Col1_DiscTaxable As Byte = 3
Private Const Col1_DiscTaxPaid As Byte = 4
Private Const Col1_Code As Byte = 5

Dim LastDate As String

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Dim GridKey As Integer

Private Sub ListView_Click()
Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
Txt(Val(ListView.Tag)).SetFocus
End Sub

Private Sub DgDate_Click()
DGDate.Visible = False
If RsDate.RecordCount > 0 Then
    Txt(VDate).TEXT = RsDate!Code
End If
Txt(VDate).SetFocus
End Sub
'Col_Pno,Col_Pname,Col_MRP,Col_Taxable,Col_TaxPaid,Col_DicsFac,Col_RefNo,Col_EffDate
'Col1_Desc,Col1_DiscMRP,Col1_DiscTaxable,Col1_DiscTaxPaid

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
'FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
    TopCtrl1.Tag = PubUParam
    WinSetting Me
    TopCtrl1.TopText2 = "Browse": TopCtrl1.TopText2.ForeColor = RGB(0, 0, 255)
    If InStr(PubUParam, "A") <> 0 Then mAdd = True Else mAdd = False
    If InStr(PubUParam, "E") <> 0 Then TopCtrl1.tEdit = True Else TopCtrl1.tEdit = False
    If InStr(PubUParam, "D") <> 0 Then mDel = True Else mDel = False
    If InStr(PubUParam, "P") <> 0 Then mPrn = True Else mPrn = False
    
    Set RsDate = New ADODB.Recordset
    RsDate.CursorLocation = adUseClient
    RsDate.Open "select distinct " & cCStr("Effect_Dt") & " as code from Part_PriceList", GCn, adOpenDynamic, adLockOptimistic
    Set DGDate.DataSource = RsDate
    Txt(RateType) = "Both"
    If PubBackEnd = "A" Then
        Set Master1 = GCn.Execute("select Description, format(MRP_Disc,'0.00') as MRPDisc,format(TB_Disc,'0.00') as TBDisc,format(TP_Disc,'0.00') as TPDisc,Party_Type from SubGroupType order by Party_Type")
    ElseIf PubBackEnd = "S" Then
        Set Master1 = GCn.Execute("select Description, MRP_Disc as MRPDisc,TB_Disc as TBDisc,TP_Disc as TPDisc,Party_Type from SubGroupType order by Party_Type")
    End If
    Set FGrid1.DataSource = Master1
    Ini_Grid
    Disp_Text False
  Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Master = Nothing
Set Master1 = Nothing
Set RsDate = Nothing
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
'Leave Cell-- > Enter Cell-- >KeyDown
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
            Case Col_MRP, Col_Taxable, Col_TaxPaid
                FGrid.CellForeColor = vbRed
                FGrid.TextMatrix(FGrid.Row, 0) = "¤"
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "0.00"
        End Select
End If
If KeyCode = vbKeyReturn Then
        Select Case FGrid.Col
            Case Col_MRP, Col_Taxable, Col_TaxPaid
                Call GridDblClick(Me, FGrid, TxtGrid, 0)
                TAddMode = False
        End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_DblClick()
If IsValid(Txt(RateType), "Rate Type") = False Then Exit Sub
If IsValid(Txt(VDate), "Effective Date") = False Then Exit Sub
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case Col_MRP, Col_Taxable, Col_TaxPaid
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
End Select
TAddMode = False
End Sub
Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    TxtSearch.TEXT = ""
    FGrid.CellBackColor = CellBackColEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
If IsValid(Txt(RateType), "Rate Type") = False Then Exit Sub
If IsValid(Txt(VDate), "Effective Date") = False Then Exit Sub

If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case Col_MRP, Col_Taxable, Col_TaxPaid
           Call Get_Text(Me, FGrid, TxtGrid, 0, True, keyascii)
        Case Col_PNo
            SelGridKeyPress TxtSearch, FGrid, Master, keyascii, "Code", CellBackColEnter, CellBackColLeave
        Case Col_PName
            SelGridKeyPress TxtSearch, FGrid, Master, keyascii, "Name", CellBackColEnter, CellBackColLeave
End Select
    If keyascii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
    FGrid1.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
    FGrid1.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
            Select Case FGrid1.Col
                Case Col1_DiscMRP, Col1_DiscTaxable, Col1_DiscTaxPaid
                    FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "0.00"
            End Select
End If
If KeyCode = vbKeyReturn Then
            Select Case FGrid1.Col
                Case Col1_DiscMRP, Col1_DiscTaxable, Col1_DiscTaxPaid
                    Call GridDblClick(Me, FGrid1, txtgrid1, 0)
                    TAddMode = False
            End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid1_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid1.Col
    Case Col1_DiscMRP, Col1_DiscTaxable, Col1_DiscTaxPaid
        Call GridDblClick(Me, FGrid1, txtgrid1, 0)
    Case Col1_Desc
         If Txt(VDate).TEXT = "" Then MsgBox "select date for price list", vbInformation: Exit Sub
         Call Fill_Grid(Val(FGrid1.TextMatrix(FGrid1.Row, Col1_Code)), FGrid1.TextMatrix(FGrid1.Row, Col1_Desc))
End Select
TAddMode = False
End Sub
Private Sub FGrid1_EnterCell()
FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.CellBackColor = CellBackColEnter
    txtgrid1(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid1_Validate(Cancel As Boolean)
    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid1_KeyPress(keyascii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid1.Col
        Case Col1_DiscMRP, Col1_DiscTaxable, Col1_DiscTaxPaid
           Call Get_Text(Me, FGrid1, txtgrid1, 0, True, keyascii)
    End Select
If keyascii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid1_Scroll()
txtgrid1(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid1_LeaveCell()
    FGrid1.CellBackColor = CellBackColLeave
End Sub



Private Sub TopCtrl1_eCancel()
    TopCtrl1.TopText2 = "Browse": TopCtrl1.TopText2.ForeColor = RGB(0, 0, 255)
    Disp_Text False
End Sub

Private Sub TopCtrl1_eEdit()
    TopCtrl1.TopText2 = "Edit": TopCtrl1.TopText2.ForeColor = RGB(255, 0, 0)
    Disp_Text True
    Txt(RateType).SetFocus
End Sub

Private Sub TopCtrl1_eExit()
Unload Me
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Long
Dim mTrans As Boolean
On Error GoTo errlbl
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    If txtgrid1(0).Visible = True Then
        If TxtGrid1Leave = False Then
            txtgrid1(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

GCn.BeginTrans
mTrans = True
'Part Price List Update
        '"Both", "MRP", "Taxable"
Select Case Txt(RateType)
    Case "Both"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" And FGrid.TextMatrix(I, Col_Origin) <> "$" Then
                GCn.Execute ("delete from part_pricelist where part_no = '" & FGrid.TextMatrix(I, Col_PNo) & "' and Div_Code='" & PubDivCode & "' and  Effect_dt = " & ConvertDate(Txt(VDate)) & "")
                GCn.Execute ("insert into part_pricelist (PART_NO,Div_Code,Effect_Dt,Site_Code,MRP,TB_SRate,TP_SRate,ref_no,U_Name,U_EntDt,U_AE)  " & _
                    "values('" & FGrid.TextMatrix(I, Col_PNo) & "','" & PubDivCode & "'," & ConvertDate(Txt(VDate)) & ",'" & PubSiteCode & "'," & Val(FGrid.TextMatrix(I, Col_MRP)) & "," & _
                    "" & Val(FGrid.TextMatrix(I, Col_Taxable)) & "," & Val(FGrid.TextMatrix(I, Col_TaxPaid)) & ",'" & Txt(RefNo) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
            ElseIf FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" And FGrid.TextMatrix(I, Col_Origin) = "$" Then
                GCn.Execute " update part_pricelist set MRP=" & Val(FGrid.TextMatrix(I, Col_MRP)) & ",TB_SRate=" & Val(FGrid.TextMatrix(I, Col_Taxable)) & ",TP_SRate=" & Val(FGrid.TextMatrix(I, Col_TaxPaid)) & " ,ref_no='" & Txt(RefNo) & "', " & _
                    "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where PART_NO= '" & FGrid.TextMatrix(I, Col_PNo) & "' and div_code=,'" & PubDivCode & "' and  Effect_dt = " & ConvertDate(Txt(VDate)) & ""
            End If
        Next
    Case "MRP"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" And FGrid.TextMatrix(I, Col_Origin) <> "$" Then
                GCn.Execute ("delete from part_pricelist where part_no = '" & FGrid.TextMatrix(I, Col_PNo) & "' and div_code='" & PubDivCode & "' and  Effect_dt = " & ConvertDate(Txt(VDate)) & "")
                GCn.Execute ("insert into part_pricelist (PART_NO,Div_Code,Effect_Dt,Site_Code,MRP,TB_SRate,TP_SRate,ref_no,U_Name,U_EntDt,U_AE)  " & _
                    "values('" & FGrid.TextMatrix(I, Col_PNo) & "','" & PubDivCode & "'," & ConvertDate(Txt(VDate)) & ",'" & PubSiteCode & "'," & Val(FGrid.TextMatrix(I, Col_MRP)) & "," & _
                    "0,0,'" & Txt(RefNo) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
            ElseIf FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" And FGrid.TextMatrix(I, Col_Origin) = "$" Then
                GCn.Execute " update part_pricelist set MRP=" & Val(FGrid.TextMatrix(I, Col_MRP)) & ",ref_no='" & Txt(RefNo) & "', " & _
                    "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where PART_NO= '" & FGrid.TextMatrix(I, Col_PNo) & "' and div_code='" & PubDivCode & "' and  Effect_dt = " & ConvertDate(Txt(VDate)) & ""
            End If
        Next
    Case "Taxable"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" And FGrid.TextMatrix(I, Col_Origin) <> "$" Then
                GCn.Execute ("delete from part_pricelist where part_no = '" & FGrid.TextMatrix(I, Col_PNo) & "' and Div_Code='" & PubDivCode & "' and  Effect_dt = " & ConvertDate(Txt(VDate)) & "")
                GCn.Execute ("insert into part_pricelist (PART_NO,Div_Code,Effect_Dt,Site_Code,MRP,TB_SRate,TP_SRate,ref_no,U_Name,U_EntDt,U_AE)  " & _
                    "values('" & FGrid.TextMatrix(I, Col_PNo) & "','" & PubDivCode & "'," & ConvertDate(Txt(VDate)) & ",'" & PubSiteCode & "',0," & _
                    "" & Val(FGrid.TextMatrix(I, Col_Taxable)) & "," & Val(FGrid.TextMatrix(I, Col_TaxPaid)) & ",'" & Txt(RefNo) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
            ElseIf FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" And FGrid.TextMatrix(I, Col_Origin) = "$" Then
                GCn.Execute " update part_pricelist set TB_SRate=" & Val(FGrid.TextMatrix(I, Col_Taxable)) & ",TP_SRate=" & Val(FGrid.TextMatrix(I, Col_TaxPaid)) & " ,ref_no='" & Txt(RefNo) & "', " & _
                    "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where PART_NO= '" & FGrid.TextMatrix(I, Col_PNo) & "' and Div_Code='" & PubDivCode & "' and  Effect_dt = " & ConvertDate(Txt(VDate)) & ""
            End If
        Next
    End Select
    ' Part Master Update
    Dim MRPEffectDt$, TBEffectDt$
    Select Case Txt(RateType)
    Case "Both"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" Then
                MRPEffectDt = IIf(FGrid.TextMatrix(I, Col_MRPEffectDt) = "", CDate(Txt(VDate)) - 1, FGrid.TextMatrix(I, Col_MRPEffectDt))
                TBEffectDt = IIf(FGrid.TextMatrix(I, Col_TBEffectDt) = "", CDate(Txt(VDate)) - 1, FGrid.TextMatrix(I, Col_TBEffectDt))
                If CDate(Txt(VDate)) > CDate(MRPEffectDt) Then
                    GCn.Execute " update part set MRP=" & Val(FGrid.TextMatrix(I, Col_MRP)) & ", " & _
                        "MRP_Effect_Dt=" & ConvertDate(Txt(VDate)) & "  , U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where PART_NO= '" & FGrid.TextMatrix(I, Col_PNo) & "' and Div_Code='" & PubDivCode & "'"
                End If
                If CDate(Txt(VDate)) > CDate(TBEffectDt) Then
                    GCn.Execute " update part set TB_SRate=" & Val(FGrid.TextMatrix(I, Col_Taxable)) & ",TP_SRate=" & Val(FGrid.TextMatrix(I, Col_TaxPaid)) & " , " & _
                        "TB_Effect_Dt =" & ConvertDate(Txt(VDate)) & ", U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where PART_NO= '" & FGrid.TextMatrix(I, Col_PNo) & "' and Div_Code='" & PubDivCode & "'"
                End If
            End If
        Next
    Case "MRP"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" Then
                MRPEffectDt = IIf(FGrid.TextMatrix(I, Col_MRPEffectDt) = "", CDate(Txt(VDate)) - 1, FGrid.TextMatrix(I, Col_MRPEffectDt))
                If CDate(Txt(VDate)) > CDate(MRPEffectDt) Then
                    GCn.Execute " update part set MRP=" & Val(FGrid.TextMatrix(I, Col_MRP)) & ", " & _
                        "MRP_Effect_Dt=" & ConvertDate(Txt(VDate)) & "  , U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where PART_NO= '" & FGrid.TextMatrix(I, Col_PNo) & "' and Div_Code='" & PubDivCode & "'"
                End If
            End If
        Next
    Case "Taxable"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, 0) = "¤" Then
                TBEffectDt = IIf(FGrid.TextMatrix(I, Col_TBEffectDt) = "", CDate(Txt(VDate)) - 1, FGrid.TextMatrix(I, Col_TBEffectDt))
                If CDate(Txt(VDate)) > CDate(MRPEffectDt) Then
                    GCn.Execute " update part set TB_SRate=" & Val(FGrid.TextMatrix(I, Col_Taxable)) & ",TP_SRate=" & Val(FGrid.TextMatrix(I, Col_TaxPaid)) & " , " & _
                        "TB_Effect_Dt =" & ConvertDate(Txt(VDate)) & ", U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where PART_NO= '" & FGrid.TextMatrix(I, Col_PNo) & "' and Div_Code='" & PubDivCode & "'"
                End If
            End If
        Next
    End Select

    ' SubGroupType Update
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, Col1_Desc) <> "" Then
            GCn.Execute "update SubGroupType set MRP_Disc = " & Val(FGrid1.TextMatrix(I, Col1_DiscMRP)) & ",TB_Disc=" & Val(FGrid1.TextMatrix(I, Col1_DiscTaxable)) & ",TP_Disc=" & Val(FGrid1.TextMatrix(I, Col1_DiscTaxPaid)) & " where Description= '" & FGrid1.TextMatrix(I, Col1_Desc) & "'"
        End If
    Next
GCn.CommitTrans
mTrans = False
RsDate.Requery
Fill_Grid 100, "General"
TopCtrl1.TopText2 = "Browse": TopCtrl1.TopText2.ForeColor = RGB(0, 0, 255)
Disp_Text False
Exit Sub
errlbl:
If mTrans = True Then
    GCn.RollbackTrans: CheckError
Else
    CheckError
End If

End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim RstPartyType As ADODB.Recordset
    Ctrl_GetFocus Txt(Index)
    TxtGrid(0).Visible = False
    txtgrid1(0).Visible = False
    Grid_Hide
    Select Case Index
    Case RateType
        ListArray = Array("Both", "MRP", "Taxable")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 3)
    End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case RateType
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 900
    Case VDate
        DGridTxtKeyDown_Mast DGDate, Txt, Index, RsDate, KeyCode, False, 0
End Select
If DGDate.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If Index <> RateType And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
 Call CheckQuote(keyascii)
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case RateType
        ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
    Case VDate
        If DGDate.Visible = True Then DGridTxtKeyUp_Mast Txt, Index, RsDate, KeyCode, "code"
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case RateType
        If Txt(Index).TEXT <> "" Then Txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case VDate
        If DGDate.Visible = True Then Exit Sub
        If IsValid(Txt(Index), "Rate Date") = False Then Cancel = True: Exit Sub
        Txt(Index).TEXT = RetDate(Txt(Index))
        If Txt(Index).TEXT <> LastDate Then
            Fill_Grid 100, "General"
        End If
End Select
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid(0).Visible = False
        Grid_Hide
        FGrid.SetFocus
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Col_MRP, Col_Taxable, Col_TaxPaid
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 11, , , True, True
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
 Call CheckQuote(keyascii)
Select Case FGrid.Col
    Case Col_MRP, Col_Taxable, Col_TaxPaid
        Call NumPress(TxtGrid(Index), keyascii, 8, 2)
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
    Case Col_MRP, Col_Taxable, Col_TaxPaid
        If Val(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) <> Val(TxtGrid(Index).TEXT) Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
            FGrid.CellForeColor = vbRed
            FGrid.TextMatrix(FGrid.Row, 0) = "¤"
        End If
End Select
TxtGridLeave = True
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

Private Sub TxtGrid1_GotFocus(Index As Integer)
    FGrid1.CellBackColor = CellBackColLeave
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtgrid1(0).TEXT = txtgrid1(0).Tag
        txtgrid1(0).Visible = False
        Grid_Hide
        FGrid1.SetFocus
        Exit Sub
    End If
    Select Case FGrid1.Col
        Case Col1_DiscMRP, Col1_DiscTaxable, Col1_DiscTaxPaid
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, Col1_DiscTaxPaid + 1, , , True, True
                End If
            End If
    End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, keyascii As Integer)
 Call CheckQuote(keyascii)
Select Case FGrid1.Col
    Case Col1_DiscMRP, Col1_DiscTaxable, Col1_DiscTaxPaid
        Call NumPress(txtgrid1(Index), keyascii, 3, 2)
End Select
End Sub


Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid1.Col
    Case Col1_DiscMRP, Col1_DiscTaxable, Col1_DiscTaxPaid
        FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = Format(Val(txtgrid1(Index).TEXT), "0.00")
End Select
TxtGrid1Leave = True
If ValidateCall = False Then
    txtgrid1(0).Visible = False
    FGrid1.SetFocus
End If
End Function

'******* Fuctions **********
Private Sub Ini_Grid()
DGDate.left = Txt(VDate).left: DGDate.top = Txt(VDate).top + Txt(VDate).height + 15
Dim I As Byte
    With FGrid
        .RowHeightMin = PubGridRowHeight
        .Cols = 13
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColWidth(0) = 400

        .TextMatrix(0, Col_PNo) = "Part No"
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 1300

        .TextMatrix(0, Col_PName) = "Part Name"
        .ColAlignment(Col_PName) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_PName) = flexAlignLeftCenter
        .ColWidth(Col_PName) = 2000
        
        .TextMatrix(0, Col_Unit) = "Unit"
        .ColAlignment(Col_Unit) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Unit) = flexAlignLeftCenter
        .ColWidth(Col_Unit) = 600

        .TextMatrix(0, Col_Grade) = "Grade"
        .ColAlignment(Col_Grade) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Grade) = flexAlignLeftCenter
        .ColWidth(Col_Grade) = 700
        .TextMatrix(0, Col_PDisc) = "PDis%"
        .ColAlignment(Col_PDisc) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_PDisc) = flexAlignLeftCenter
        .ColWidth(Col_PDisc) = 700
        .TextMatrix(0, Col_ISalDisc) = "ISalD%"
        .ColAlignment(Col_ISalDisc) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_ISalDisc) = flexAlignLeftCenter
        .ColWidth(Col_ISalDisc) = 700

        .TextMatrix(0, Col_Origin) = "PList"
        .ColAlignment(Col_Origin) = flexAlignCenterCenter
        .ColAlignmentFixed(Col_Origin) = flexAlignCenterCenter
        .ColWidth(Col_Origin) = 500

        .TextMatrix(0, Col_MRP) = "MRP Rate"
        .ColAlignment(Col_MRP) = flexAlignRightCenter
        .ColAlignmentFixed(Col_MRP) = flexAlignRightCenter
        .ColWidth(Col_MRP) = 1300

        .TextMatrix(0, Col_Taxable) = "Taxable Rate"
        .ColAlignment(Col_Taxable) = flexAlignRightCenter
        .ColAlignmentFixed(Col_Taxable) = flexAlignRightCenter
        .ColWidth(Col_Taxable) = 1300

        .TextMatrix(0, Col_TaxPaid) = "TaxPaid Rate"
        .ColAlignment(Col_TaxPaid) = flexAlignRightCenter
        .ColAlignmentFixed(Col_TaxPaid) = flexAlignRightCenter
        .ColWidth(Col_TaxPaid) = 1300
        
        .ColWidth(Col_MRPEffectDt) = 0
        .ColWidth(Col_TBEffectDt) = 0

    End With
    
    With FGrid1
        .RowHeightMin = PubGridRowHeight
        .Cols = 6

        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignmentFixed(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, Col1_Desc) = "Description"
        .ColAlignment(Col1_Desc) = flexAlignLeftCenter
        .ColAlignmentFixed(Col1_Desc) = flexAlignLeftCenter
        .ColWidth(Col1_Desc) = 1500

        .TextMatrix(0, Col1_DiscMRP) = "MRP %"
        .ColAlignment(Col1_DiscMRP) = flexAlignRightCenter
        .ColAlignmentFixed(Col1_DiscMRP) = flexAlignRightCenter
        .ColWidth(Col1_DiscMRP) = 1100

        .TextMatrix(0, Col1_DiscTaxable) = "Taxable %"
        .ColAlignment(Col1_DiscTaxable) = flexAlignRightCenter
        .ColAlignmentFixed(Col1_DiscTaxable) = flexAlignRightCenter
        .ColWidth(Col1_DiscTaxable) = 1200

        .TextMatrix(0, Col1_DiscTaxPaid) = "TaxPaid %"
        .ColAlignment(Col1_DiscTaxPaid) = flexAlignRightCenter
        .ColAlignmentFixed(Col1_DiscTaxPaid) = flexAlignRightCenter
        .ColWidth(Col1_DiscTaxPaid) = 1100
        
        .ColWidth(Col1_Code) = 0

    End With
Exit Sub
End Sub

Private Sub Grid_Hide()
    If DGDate.Visible = True Then DGDate.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub Fill_Grid(Index As Integer, GridCaption As String)
Dim Rst As ADODB.Recordset
Dim MRPPer As Double, TBPer As Double, TPPer As Double
If IsValid(Txt(RateType), "Rate Type") = False Then Exit Sub
If IsValid(Txt(VDate), "Effective Date") = False Then Exit Sub
LastDate = Txt(VDate)
Select Case Index
    Case 100
        LBLListDesc(0) = "Spare Price List (General)"
        If GCn.Execute("select count(*) from part_pricelist where Div_Code='" & PubDivCode & "' and Effect_dt = " & ConvertDate(Txt(VDate)) & "").Fields(0).Value > 0 Then
            GSQL = "SELECT P.Part_No as Code, P.Part_Name as Name,P.Unit, " _
                & "P.Part_Grade, format(PDisc.PurcDisc_Per,'0.00'),format(PDisc.SalDisc_Per,'0.00'),'$' as Origin, " _
                & "format(PL.MRP,'0.00') as MRP," _
                & "format(PL.TB_SRate,'0.00') as TB_SRate," _
                & "format(PL.TP_SRate,'0.00') as TP_SRate," _
                & "P.MRP_Effect_Dt,P.TB_Effect_Dt " _
                & "FROM (Part P LEFT JOIN Part_PriceList PL on P.PART_NO&P.Div_Code = PL.PART_NO&PL.Div_Code) " _
                & "Left Join Part_DiscFactor  PDisc on P.Part_Grade=Pdisc.DiscFac_Catg " _
                & "where PL.Div_Code='" & PubDivCode & "' and PL.Effect_dt = " & ConvertDate(Txt(VDate)) & " " _
                & "union " _
                & "Select P.Part_No AS Code,P.Part_Name as Name,P.Unit," _
                & "P.Part_Grade, format(PDisc.PurcDisc_Per,'0.00'),format(PDisc.SalDisc_Per,'0.00'),'' as Origin," _
                & "format(P.MRP,'0.00') as MRP,format(P.TB_SRate,'0.00') as TB_SRate,format(P.TP_SRate,'0.00') as TP_SRate,P.MRP_Effect_Dt,P.TB_Effect_Dt " _
                & "FROM (Part P LEFT JOIN Part_PriceList PL on P.PART_NO = PL.PART_NO) " _
                & "Left Join Part_DiscFactor  PDisc on P.Part_Grade=Pdisc.DiscFac_Catg " _
                & "where P.Div_Code='" & PubDivCode & _
                "' and P.Part_No not in (select Part_no from Part_PriceList where  Part_PriceList.Div_Code='" & PubDivCode & "' and Part_PriceList.Effect_dt = " & ConvertDate(Txt(VDate)) & ") " _
                & "Order By code,Name"
                
            Set Master = GCn.Execute(GSQL)
            Txt(RefNo).TEXT = GCn.Execute("select " & xIsNull("ref_no", "") & " from part_pricelist where Div_Code='" & PubDivCode & "' and Effect_dt = " & ConvertDate(Txt(VDate)) & "").Fields(0).Value
        Else
            If MsgBox("No Record Found Of Given Criteria.Do You Want To Create New Rate List ? ", vbYesNo + vbCritical + vbDefaultButton2, "No Matching Record!") = vbYes Then
                AddMode = 1
                If PubBackEnd = "A" Then
                    Set Master = GCn.Execute("Select P.Part_No AS Code,P.Part_Name as Name,P.Unit," _
                        & "P.Part_Grade, format(PDisc.PurcDisc_Per,'0.00'),format(PDisc.SalDisc_Per,'0.00'),'' as Origin, " _
                        & "format(P.MRP,'0.00') as MRP,format(P.TB_SRate,'0.00') as TB_SRate,format(P.TP_SRate,'0.00') as TP_SRate,P.MRP_Effect_Dt,P.TB_Effect_Dt " _
                        & "From Part P Left Join Part_DiscFactor PDisc on P.Part_Grade=Pdisc.DiscFac_Catg " _
                        & "where P.Div_Code='" & PubDivCode & "' Order By P.Part_No,P.Part_Name")
                ElseIf PubBackEnd = "S" Then
                    Set Master = GCn.Execute("Select P.Part_No AS Code,P.Part_Name as Name,P.Unit," _
                        & "P.Part_Grade, PDisc.PurcDisc_Per, PDisc.SalDisc_Per, '' as Origin, " _
                        & "P.MRP as MRP,P.TB_SRate as TB_SRate,P.TP_SRate as TP_SRate,P.MRP_Effect_Dt,P.TB_Effect_Dt " _
                        & "From Part P Left Join Part_DiscFactor PDisc on P.Part_Grade=Pdisc.DiscFac_Catg " _
                        & "where P.Div_Code='" & PubDivCode & "' Order By P.Part_No,P.Part_Name")
                End If
            Else
                Txt(VDate).TEXT = ""
                FGrid.Rows = 1
                FGrid.AddItem ""
                'Ini_Grid
                FGrid.FixedRows = 1
                Exit Sub
            End If
        End If
    Case Index
        Set Rst = GCn.Execute("select MRP_Disc,TB_Disc,tp_Disc from SubGroupType where Party_Type = " & Index & "")
        If Rst.RecordCount > 0 Then
            MRPPer = Rst!mrp_Disc
            TBPer = Rst!tb_Disc
            TPPer = Rst!tp_Disc
        Else
            MsgBox "No Record Found", vbInformation: Exit Sub
        End If
        LBLListDesc(0) = "Spare Price List " & GridCaption
        If GCn.Execute("select count(*) from part_pricelist where Effect_dt = " & ConvertDate(Txt(VDate)) & "").Fields(0).Value > 0 Then
            GSQL = "SELECT P.Part_No as Code, P.Part_Name as Name,P.Unit, " _
                & "P.Part_Grade, PDisc.PurcDisc_Per,PDisc.SalDisc_Per,'$' as Origin, " _
                & "format(PL.MRP - ((PL.MRP * " & MRPPer & ")/100),'0.00') as MRP," _
                & "format(PL.TB_SRate - ((PL.TB_SRate * " & TBPer & ")/100),'0.00') as TB_SRate," _
                & "format(PL.TP_SRate - ((PL.TP_SRate * " & TPPer & ")/100),'0.00') as TP_SRate," _
                & "P.MRP_Effect_Dt,P.TB_Effect_Dt " _
                & "FROM (Part P LEFT JOIN Part_PriceList PL on P.PART_NO&P.Div_Code = PL.PART_NO&PL.Div_Code) " _
                & "Left Join Part_DiscFactor PDisc on P.Part_Grade=Pdisc.DiscFac_Catg " _
                & "where PL.Div_Code='" & PubDivCode & "' and PL.Effect_dt = " & ConvertDate(Txt(VDate)) & " " _
                & "union " _
                & "Select P.Part_No AS Code,P.Part_Name as Name,P.Unit," _
                & "P.Part_Grade, PDisc.PurcDisc_Per,PDisc.SalDisc_Per,'' as Origin," _
                & "format(P.MRP,'0.00') as MRP,format(P.TB_SRate,'0.00') as TB_SRate,format(P.TP_SRate,'0.00') as TP_SRate,P.MRP_Effect_Dt,P.TB_Effect_Dt " _
                & "FROM (Part P LEFT JOIN Part_PriceList PL on P.PART_NO&P.Div_Code = PL.PART_NO&PL.Div_Code) " _
                & "Left Join Part_DiscFactor PDisc on P.Part_Grade=Pdisc.DiscFac_Catg " _
                & "where P.Part_No not in (select Part_no from Part_PriceList where Part_PriceList.Div_Code='" & PubDivCode & "' and Part_PriceList.Effect_dt = " & ConvertDate(Txt(VDate)) & ") " _
                & "Order By code,Name"
                
            Set Master = GCn.Execute(GSQL)
            Txt(RefNo).TEXT = GCn.Execute("select " & xIsNull("ref_no", "") & " from part_pricelist where Div_Code='" & PubDivCode & "' and Effect_dt = " & ConvertDate(Txt(VDate)) & "").Fields(0).Value
        Else
            Set Master = GCn.Execute("Select P.Part_No AS Code,P.Part_Name as Name,P.Unit," _
                & "P.Part_Grade, PDisc.PurcDisc_Per,PDisc.SalDisc_Per,'' as Origin, " _
                & "format(P.MRP - ((P.MRP * " & MRPPer & ")/100),'0.00') as MRP,format(P.TB_SRate - ((P.TB_SRate * " & TBPer & ")/100),'0.00') as TB_SRate,format(P.TP_SRate - ((P.TP_SRate * " & TPPer & ")/100),'0.00') as TP_SRate,P.MRP_Effect_Dt,P.TB_Effect_Dt " _
                & "From Part P Left Join Part_DiscFactor PDisc on P.Part_Grade=Pdisc.DiscFac_Catg " _
                & "where P.Div_Code='" & PubDivCode & "' Order By P.Part_No,P.Part_Name")
        End If
        Set Rst = Nothing
End Select
Set FGrid.DataSource = Master
Ini_Grid
End Sub

Private Sub TxtSearch_Click()
 FGrid.SetFocus: TxtSearch.TEXT = "": TxtSearch.Visible = False
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If NavigationKey(KeyCode) = True Then FGrid.SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then FGrid.SetFocus: TxtSearch.Visible = False
End Sub

Private Sub TxtSearch_KeyPress(keyascii As Integer)
Select Case FGrid.Col
    Case Col_PNo
        SelGridKeyPress TxtSearch, FGrid, Master, keyascii, "Code", CellBackColEnter, CellBackColLeave: keyascii = 0
    Case Col_PName
        SelGridKeyPress TxtSearch, FGrid, Master, keyascii, "NAME", CellBackColEnter, CellBackColLeave: keyascii = 0
End Select
End Sub

Private Sub TxtSearch_LostFocus()
 FGrid.SetFocus: TxtSearch.Visible = False:: TxtSearch.TEXT = ""
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
Next
TopCtrl1.tEdit = Not Enb
TopCtrl1.tExit = Not Enb
TopCtrl1.tPrn = Not Enb
TopCtrl1.tCancel = Enb
TopCtrl1.tSave = Enb

TopCtrl1.tRef = False
TopCtrl1.tAdd = False
TopCtrl1.tFirst = False
TopCtrl1.tNext = False
TopCtrl1.tPrev = False
TopCtrl1.tLast = False
TopCtrl1.tFind = False
TopCtrl1.tDel = False

TxtGrid(0).Visible = False
txtgrid1(0).Visible = False
TxtSearch.Visible = False
End Sub
