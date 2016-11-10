VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmOffTakeTgtForeCast 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Off-take Target Forecasting Entry"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
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
      Left            =   1965
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   6390
      Left            =   45
      TabIndex        =   1
      Top             =   555
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   11271
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   3
      Cols            =   30
      FixedRows       =   2
      FixedCols       =   3
      BackColorFixed  =   13623520
      ForeColorFixed  =   8388736
      BackColorSel    =   13300221
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   16761024
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "WWW"
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
      _Band(0).Cols   =   30
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmOffTakeTgtForeCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim Master As ADODB.Recordset
Dim ExitCtrl As Boolean

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_ModelCode As Byte = 1         ' Model Code
Private Const Col_ModelDesc As Byte = 2         ' Model Desc
Private Const Col_ForYear As Byte = 3           ' For Year
Private Const Col_ChasType As Byte = 4          ' Chassis Type
Private Const Col_ModelType As Byte = 5         ' Model Type
Private Const Col_Qty4 As Byte = 6              ' Qty Apr.
Private Const Col_Qty4Tel As Byte = 7           ' Telco Qty Apr.
Private Const Col_Qty5 As Byte = 8              ' Qty May.
Private Const Col_Qty5Tel As Byte = 9           ' Telco Qty May.
Private Const Col_Qty6 As Byte = 10             ' Qty Jun.
Private Const Col_Qty6Tel As Byte = 11          ' Telco Qty Jun.
Private Const Col_Qty7 As Byte = 12             ' Qty Jul.
Private Const Col_Qty7Tel As Byte = 13          ' Telco Qty Jul.
Private Const Col_Qty8 As Byte = 14             ' Qty Aug.
Private Const Col_Qty8Tel As Byte = 15          ' Telco Qty Aug.
Private Const Col_Qty9 As Byte = 16             ' Qty Sep.
Private Const Col_Qty9Tel As Byte = 17          ' Telco Qty Sep.
Private Const Col_Qty10 As Byte = 18            ' Qty Oct.
Private Const Col_Qty10Tel As Byte = 19         ' Telco Qty Oct.
Private Const Col_Qty11 As Byte = 20            ' Qty Nov.
Private Const Col_Qty11Tel As Byte = 21         ' Telco Qty Nov.
Private Const Col_Qty12 As Byte = 22            ' Qty Dec.
Private Const Col_Qty12Tel As Byte = 23         ' Telco Qty Dec.
Private Const Col_Qty1 As Byte = 24             ' Qty Jan.
Private Const Col_Qty1Tel As Byte = 25          ' Telco Qty Jan.
Private Const Col_Qty2 As Byte = 26             ' Qty Feb.
Private Const Col_Qty2Tel As Byte = 27          ' Telco Qty Feb.
Private Const Col_Qty3 As Byte = 28             ' Qty Mar.
Private Const Col_Qty3Tel As Byte = 29          ' Telco Qty Mar.

Private Sub Disp_Text(Enb As Boolean)
    
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error Resume Next
Dim I As Integer
    For I = 2 To FGrid.Rows - 1
        If MyValue = Trim(FGrid.TextMatrix(I, Col_ModelCode)) Then
            FGrid.Row = I
            FGrid.Col = 6
            TxtGrid(0).left = FGrid.CellLeft
            FGrid.SetFocus
            Exit For
        End If
    Next
End Sub
'* Used for intialize grid columns
Private Sub Grid_Ini()
' To Change
Dim TelcoHead As String, ColW As Integer, ColW1 As Integer
    ColW = 465: ColW1 = 485
    TelcoHead = "Telco"
    With FGrid
        .left = Me.left '+ 45
        .width = Me.width - 90
        .height = Me.height - 900
        .top = 400

        .RowHeightMin = PubGridRowHeight '220
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeRow(0) = True
        .Cols = 30
        .ColAlignmentFixed = flexAlignCenterCenter
        .TextMatrix(0, Col_SrNo) = "S.No"
        .TextMatrix(1, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignRightCenter
        .ColWidth(Col_SrNo) = 450

        .MergeCol(2) = True
        .TextMatrix(0, Col_ModelDesc) = "Model"
        .TextMatrix(1, Col_ModelDesc) = "Model"
        .ColAlignmentFixed(Col_ModelDesc) = flexAlignLeftCenter
        .ColWidth(Col_ModelDesc) = 3000

        .TextMatrix(1, Col_ModelCode) = "Model Code"
        .ColWidth(Col_ModelCode) = 0

        .TextMatrix(1, Col_ForYear) = "For Year"
        .ColWidth(Col_ForYear) = 0

        .TextMatrix(1, Col_ChasType) = "Chassis Type"
        .ColWidth(Col_ChasType) = 0

        .TextMatrix(0, Col_ModelType) = "Model Type"
        .ColWidth(Col_ModelType) = 0
        
        .MergeCol(Col_Qty4) = True
        .MergeCol(Col_Qty4Tel) = True
        .TextMatrix(0, Col_Qty4) = "April"
        .ColAlignment(Col_Qty4) = flexAlignRightCenter
        .ColWidth(Col_Qty4) = ColW

        .TextMatrix(0, Col_Qty4Tel) = .TextMatrix(0, Col_Qty4)
        .ColAlignment(Col_Qty4Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty4Tel) = ColW1
        
        .TextMatrix(1, Col_Qty4) = "Own"
        .ColWidth(Col_Qty4) = ColW

        .TextMatrix(1, Col_Qty4Tel) = TelcoHead
        .ColWidth(Col_Qty4Tel) = ColW1

        
        .MergeCol(Col_Qty5) = True
        .MergeCol(Col_Qty5Tel) = True
        .TextMatrix(0, Col_Qty5) = "May"
        .ColAlignment(Col_Qty5) = flexAlignRightCenter
        .ColWidth(Col_Qty5) = ColW

        .TextMatrix(0, Col_Qty5Tel) = .TextMatrix(0, Col_Qty5)
        .ColAlignment(Col_Qty5Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty5Tel) = ColW1

        .TextMatrix(1, Col_Qty5) = "Own"
        .ColWidth(Col_Qty5) = ColW

        .TextMatrix(1, Col_Qty5Tel) = TelcoHead
        .ColWidth(Col_Qty5Tel) = ColW1

        .MergeCol(Col_Qty6) = True
        .MergeCol(Col_Qty6Tel) = True
        .TextMatrix(0, Col_Qty6) = "June"
        .ColAlignment(Col_Qty6) = flexAlignRightCenter
        .ColWidth(Col_Qty6) = ColW

        .TextMatrix(0, Col_Qty6Tel) = .TextMatrix(0, Col_Qty6)
        .ColAlignment(Col_Qty6Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty6Tel) = ColW1

        .TextMatrix(1, Col_Qty6) = "Own"
        .ColWidth(Col_Qty6) = ColW

        .TextMatrix(1, Col_Qty6Tel) = TelcoHead
        .ColWidth(Col_Qty6Tel) = ColW1

        .MergeCol(Col_Qty7) = True
        .MergeCol(Col_Qty7Tel) = True
        .TextMatrix(0, Col_Qty7) = "July"
        .ColAlignment(Col_Qty7) = flexAlignRightCenter
        .ColWidth(Col_Qty6) = ColW

        .TextMatrix(0, Col_Qty7Tel) = .TextMatrix(0, Col_Qty7)
        .ColAlignment(Col_Qty7Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty7Tel) = ColW1

        .TextMatrix(1, Col_Qty7) = "Own"
        .ColWidth(Col_Qty7) = ColW

        .TextMatrix(1, Col_Qty7Tel) = TelcoHead
        .ColWidth(Col_Qty7Tel) = ColW1
        
        .MergeCol(Col_Qty8) = True
        .MergeCol(Col_Qty8Tel) = True
        .TextMatrix(0, Col_Qty8) = "August"
        .ColAlignment(Col_Qty8) = flexAlignRightCenter
        .ColWidth(Col_Qty8) = ColW

        .TextMatrix(0, Col_Qty8Tel) = .TextMatrix(0, Col_Qty8)
        .ColAlignment(Col_Qty8Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty8Tel) = ColW1

        .TextMatrix(1, Col_Qty8) = "Own"
        .ColWidth(Col_Qty8) = ColW

        .TextMatrix(1, Col_Qty8Tel) = TelcoHead
        .ColWidth(Col_Qty8Tel) = ColW1

        .MergeCol(Col_Qty9) = True
        .MergeCol(Col_Qty9Tel) = True
        .TextMatrix(0, Col_Qty9) = "September"
        .ColAlignment(Col_Qty9) = flexAlignRightCenter
        .ColWidth(Col_Qty9) = ColW

        .TextMatrix(0, Col_Qty9Tel) = .TextMatrix(0, Col_Qty9)
        .ColAlignment(Col_Qty9Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty9Tel) = ColW1

        .TextMatrix(1, Col_Qty9) = "Own"
        .ColWidth(Col_Qty9) = ColW

        .TextMatrix(1, Col_Qty9Tel) = TelcoHead
        .ColWidth(Col_Qty9Tel) = ColW1

        .MergeCol(Col_Qty10) = True
        .MergeCol(Col_Qty10Tel) = True
        .TextMatrix(0, Col_Qty10) = "October"
        .ColAlignment(Col_Qty10) = flexAlignRightCenter
        .ColWidth(Col_Qty10) = ColW

        .TextMatrix(0, Col_Qty10Tel) = .TextMatrix(0, Col_Qty10)
        .ColAlignment(Col_Qty10Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty10Tel) = ColW1

        .TextMatrix(1, Col_Qty10) = "Own"
        .ColWidth(Col_Qty10) = ColW

        .TextMatrix(1, Col_Qty10Tel) = TelcoHead
        .ColWidth(Col_Qty10Tel) = ColW1

        .MergeCol(Col_Qty11) = True
        .MergeCol(Col_Qty11Tel) = True
        .TextMatrix(0, Col_Qty11) = "November"
        .ColAlignment(Col_Qty11) = flexAlignRightCenter
        .ColWidth(Col_Qty11) = ColW

        .TextMatrix(0, Col_Qty11Tel) = .TextMatrix(0, Col_Qty11)
        .ColAlignment(Col_Qty11Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty11Tel) = ColW1

        .TextMatrix(1, Col_Qty11) = "Own"
        .ColWidth(Col_Qty11) = ColW

        .TextMatrix(1, Col_Qty11Tel) = TelcoHead
        .ColWidth(Col_Qty11Tel) = ColW1

        .MergeCol(Col_Qty12) = True
        .MergeCol(Col_Qty12Tel) = True
        .TextMatrix(0, Col_Qty12) = "December"
        .ColAlignment(Col_Qty12) = flexAlignRightCenter
        .ColWidth(Col_Qty12) = ColW

        .TextMatrix(0, Col_Qty12Tel) = .TextMatrix(0, Col_Qty12)
        .ColAlignment(Col_Qty12Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty12Tel) = ColW1

        .TextMatrix(1, Col_Qty12) = "Own"
        .ColWidth(Col_Qty12) = ColW

        .TextMatrix(1, Col_Qty12Tel) = TelcoHead
        .ColWidth(Col_Qty12Tel) = ColW1

        .MergeCol(Col_Qty1) = True
        .MergeCol(Col_Qty1Tel) = True
        .TextMatrix(0, Col_Qty1) = "January"
        .ColAlignment(Col_Qty1) = flexAlignRightCenter
        .ColWidth(Col_Qty1) = ColW

        .TextMatrix(0, Col_Qty1Tel) = .TextMatrix(0, Col_Qty1)
        .ColAlignment(Col_Qty1Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty1Tel) = ColW1

        .TextMatrix(1, Col_Qty1) = "Own"
        .ColWidth(Col_Qty1) = ColW

        .TextMatrix(1, Col_Qty1Tel) = TelcoHead
        .ColWidth(Col_Qty1Tel) = ColW1

        .MergeCol(Col_Qty2) = True
        .MergeCol(Col_Qty2Tel) = True
        .TextMatrix(0, Col_Qty2) = "February"
        .ColAlignment(Col_Qty2) = flexAlignRightCenter
        .ColWidth(Col_Qty2) = ColW

        .TextMatrix(0, Col_Qty2Tel) = .TextMatrix(0, Col_Qty2)
        .ColAlignment(Col_Qty2Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty2Tel) = ColW1

        .TextMatrix(1, Col_Qty2) = "Own"
        .ColWidth(Col_Qty2) = ColW

        .TextMatrix(1, Col_Qty2Tel) = TelcoHead
        .ColWidth(Col_Qty2Tel) = ColW1

        .MergeCol(Col_Qty3) = True
        .MergeCol(Col_Qty3Tel) = True
        .TextMatrix(0, Col_Qty3) = "March"
        .ColAlignment(Col_Qty3) = flexAlignRightCenter
        .ColWidth(Col_Qty3) = ColW

        .TextMatrix(0, Col_Qty3Tel) = .TextMatrix(0, Col_Qty3)
        .ColAlignment(Col_Qty3Tel) = flexAlignRightCenter
        .ColWidth(Col_Qty3Tel) = ColW1

        .TextMatrix(1, Col_Qty3) = "Own"
        .ColWidth(Col_Qty3) = ColW

        .TextMatrix(1, Col_Qty3Tel) = TelcoHead
        .ColWidth(Col_Qty3Tel) = ColW1

'        .Row = 0
'        .Col = Col_Qty4
'        .MergeCol(Col_Qty4) = True    ' Allow merge on Columns 0 thru 3
'        .MergeCells = flexMergeRestrictColumns
'        .Col = Col_Qty4Tel
'        .MergeCol(Col_Qty4Tel) = True
'        .MergeCells = flexMergeRestrictColumns
    End With
End Sub

Private Function TxtGridLeave() As Boolean
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = MakeBlank(Val(TxtGrid(0).TEXT))
    ExitCtrl = True
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function

Private Function MakeBlank(Temp As Double) As String
    MakeBlank = IIf(Temp = 0, "", Temp)
End Function

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
    CheckError
End Sub

Private Sub FillModel()
Dim Rst As ADODB.Recordset, I As Integer
Dim ForYear As String
    FGrid.Rows = 2
    ForYear = CStr(Year(PubStartDate)) & "-" & Right(CStr(Year(PubEndDate)), 2)
    Set Rst = GCn.Execute("Select Model.Model As ModelCode,Model.Model,Model.Chas_Type,Model.Model_Type " _
        & "From Model " _
        & "Order by Model.Model")
    If Rst.RecordCount > 0 Then
        I = 1
        Do Until Rst.EOF
                        '|0 Col_SrNo |1 Col_ModelCode         |2 Col_ModelDesc       |3 Col_ForYear         |4 Col_ChasType      |5 Col_ModelType
            FGrid.AddItem I & Chr(9) & Rst!ModelCode & Chr(9) & Rst!Model & Chr(9) & ForYear & Chr(9) & Rst!Chas_Type & Chr(9) & Rst!Model_Type
            Rst.MoveNext '0                    1                    2                       3                       4                        5
            I = I + 1
        Loop
        FGrid.FixedRows = 2
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 2
    End If
    '|0 Col_SrNo    |1 Col_ModelDesc       |2 Col_ForYear         |3 Col_ChasType          |4 Col_ModelType           |5 Col_Qty4            |6 Col_Qty4Tel           |7 Col_Qty5         |8 Col_Qty5Tel              |9 Col_Qty6           |10 Col_Qty6Tel          |11 Col_Qty7           |12 Col_Qty7Tel          |13 Col_Qty8           |14 Col_Qty8Tel          |15 Col_Qty9           |16 Col_Qty9Tel         |17 Col_Qty10          |18 Col_Qty10Tel         |19 Col_Qty11          |20 Col_Qty11Tel         |21 Col_Qty12          |22 Col_Qty12Tel          |23 Col_Qty1           |24 Col_Qty1Tel          |25 Col_Qty2           |26 Col_Qty2Tel          |27 Col_Qty3           |28 Col_Qty3Tel
    If Master.RecordCount > 0 Then
        Set Rst = GCn.Execute("Select Model.Model As ModelCode,V.* " _
            & "From (Model Left Join Veh_Forecast V on Model.Model=V.Model) " _
            & "Where V.For_Year='" & ForYear & "' " _
            & "Order by Model.Model")
        For I = 2 To FGrid.Rows - 1
            Rst.FIND "MODEL='" & FGrid.TextMatrix(I, Col_ModelCode) & "'"
            If Rst.EOF = False Then
                FGrid.TextMatrix(I, Col_Qty4) = IIf(Rst!Qty_04 = 0, "", Rst!Qty_04)
                FGrid.TextMatrix(I, Col_Qty4Tel) = IIf(Rst!TargQty_04 = 0, "", Rst!TargQty_04)
                FGrid.TextMatrix(I, Col_Qty5) = IIf(Rst!Qty_05 = 0, "", Rst!Qty_05)
                FGrid.TextMatrix(I, Col_Qty5Tel) = IIf(Rst!TargQty_05 = 0, "", Rst!TargQty_05)
                FGrid.TextMatrix(I, Col_Qty6) = IIf(Rst!Qty_06 = 0, "", Rst!Qty_06)
                FGrid.TextMatrix(I, Col_Qty6Tel) = IIf(Rst!TargQty_06 = 0, "", Rst!TargQty_06)
                FGrid.TextMatrix(I, Col_Qty7) = IIf(Rst!Qty_07 = 0, "", Rst!Qty_07)
                FGrid.TextMatrix(I, Col_Qty7Tel) = IIf(Rst!TargQty_07 = 0, "", Rst!TargQty_07)
                FGrid.TextMatrix(I, Col_Qty8) = IIf(Rst!Qty_08 = 0, "", Rst!Qty_08)
                FGrid.TextMatrix(I, Col_Qty8Tel) = IIf(Rst!TargQty_08 = 0, "", Rst!TargQty_08)
                FGrid.TextMatrix(I, Col_Qty9) = IIf(Rst!Qty_09 = 0, "", Rst!Qty_09)
                FGrid.TextMatrix(I, Col_Qty9Tel) = IIf(Rst!TargQty_09 = 0, "", Rst!TargQty_09)
                FGrid.TextMatrix(I, Col_Qty10) = IIf(Rst!Qty_10 = 0, "", Rst!Qty_10)
                FGrid.TextMatrix(I, Col_Qty10Tel) = IIf(Rst!TargQty_10 = 0, "", Rst!TargQty_10)
                FGrid.TextMatrix(I, Col_Qty11) = IIf(Rst!Qty_11 = 0, "", Rst!Qty_11)
                FGrid.TextMatrix(I, Col_Qty11Tel) = IIf(Rst!TargQty_11 = 0, "", Rst!TargQty_11)
                FGrid.TextMatrix(I, Col_Qty12) = IIf(Rst!Qty_12 = 0, "", Rst!Qty_12)
                FGrid.TextMatrix(I, Col_Qty12Tel) = IIf(Rst!TargQty_12 = 0, "", Rst!TargQty_12)
                FGrid.TextMatrix(I, Col_Qty1) = IIf(Rst!Qty_01 = 0, "", Rst!Qty_01)
                FGrid.TextMatrix(I, Col_Qty1Tel) = IIf(Rst!TargQty_01 = 0, "", Rst!TargQty_01)
                FGrid.TextMatrix(I, Col_Qty2) = IIf(Rst!Qty_02 = 0, "", Rst!Qty_02)
                FGrid.TextMatrix(I, Col_Qty2Tel) = IIf(Rst!TargQty_02 = 0, "", Rst!TargQty_02)
                FGrid.TextMatrix(I, Col_Qty3) = IIf(Rst!Qty_03 = 0, "", Rst!Qty_03)
                FGrid.TextMatrix(I, Col_Qty3Tel) = IIf(Rst!TargQty_03 = 0, "", Rst!TargQty_03)
            End If
        Next
    End If
    
'    FGrid.AddItem I & Chr(9) & Rst!Model_Desc & Chr(9) & Rst!For_Year & Chr(9) & Rst!Chas_Type & Chr(9) & Rst!Model_Type & Chr(9) & Rst!Qty_04 & Chr(9) & Rst!TargQty_04 & Chr(9) & Rst!Qty_05 & Chr(9) & Rst!TargQty_05 & Chr(9) & Rst!Qty_06 & Chr(9) & Rst!TargQty_06 & Chr(9) & Rst!Qty_07 & Chr(9) & Rst!TargQty_07 & Chr(9) & Rst!Qty_08 & Chr(9) & Rst!TargQty_08 & Chr(9) & Rst!Qty_09 & Chr(9) & Rst!TargQty_09 & Chr(9) & Rst!Qty_10 & Chr(9) & Rst!TargQty_10 & Chr(9) & Rst!Qty_11 & Chr(9) & Rst!TargQty_11 & Chr(9) & Rst!Qty_12 & Chr(9) & Rst!TargQty_12 & Chr(9) & Rst!Qty_01 & Chr(9) & Rst!TargQty_01 & Chr(9) & Rst!Qty_02 & Chr(9) & Rst!TargQty_02 & Chr(9) & Rst!Qty_03 & Chr(9) & Rst!TargQty_03 & Chr(9) & Rst!ModelCode
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim ForYear As String
    TopCtrl1.Tag = "EP": WinSetting Me: Grid_Ini

    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic

 Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(v.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    ForYear = CStr(Year(PubStartDate)) & "-" & Right(CStr(Year(PubEndDate)), 2)
    Set Master = GCn.Execute("Select " & cTrim("V.MODEL") & " As SearchCode,Model.Model,V.* " _
        & "From (Model Right Join Veh_Forecast V on Model.Model=V.Model) " _
        & "Where V.For_Year='" & ForYear & "' " & sitecond & " " _
        & "Order by V.For_Year,Model.Model")

    FillModel
    TopCtrl1.tAdd = False
    TopCtrl1.tDel = False
    TopCtrl1.tPrn = False
    TopCtrl1.tNext = False
    TopCtrl1.tPrev = False
    TopCtrl1.tFirst = False
    TopCtrl1.tLast = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    SETS "EDIT", Me, Master
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
End Sub

Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
End Sub

Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
End Sub

Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
Dim ForYear As String
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(Veh_Forecast.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    ForYear = CStr(Year(PubStartDate)) & "-" & Right(CStr(Year(PubEndDate)), 2)
    GSQL = "Select " & cTrim("V.MODEL") & " As SearchCode,M.Model FROM Veh_Forecast V Left Join Model M On V.Model = M.Model Where V.For_Year='" & ForYear & "' " & sitecond & " Order by V.For_Year,M.Model"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
    Master.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean
Dim ForYear As String
Dim Rst As ADODB.Recordset
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid_LostFocus 0
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If

    ForYear = CStr(Year(PubStartDate)) & "-" & Right(CStr(Year(PubEndDate)), 2)
    GCn.BeginTrans
        mTrans = True
'        If Master.RecordCount = 0 Then      ' When no record in the table
            GCn.Execute ("Delete  From Veh_Forecast Where FOR_YEAR='" & ForYear & "'")
            For I = 2 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, Col_ModelDesc) <> "" Then
                    GCn.Execute "Insert Into Veh_Forecast(" _
                    & "FOR_YEAR,MODEL,Site_Code,CHAS_TYPE,MODEL_TYPE," _
                    & "QTY_04,TargQty_04,QTY_05,TargQty_05," _
                    & "QTY_06,TargQty_06,QTY_07,TargQty_07," _
                    & "QTY_08,TargQty_08,QTY_09,TargQty_09," _
                    & "QTY_10,TargQty_10,QTY_11,TargQty_11," _
                    & "QTY_12,TargQty_12,QTY_01,TargQty_01," _
                    & "QTY_02,TargQty_02,QTY_03,TargQty_03," _
                    & "U_Name,U_EntDt,U_AE) " _
                    & "Values(" _
                    & "'" & ForYear & "','" & FGrid.TextMatrix(I, Col_ModelCode) & "','" & PubSiteCode & "','" & FGrid.TextMatrix(I, Col_ChasType) & "','" & FGrid.TextMatrix(I, Col_ModelType) & "'," _
                    & "" & Val(FGrid.TextMatrix(I, Col_Qty4)) & "," & Val(FGrid.TextMatrix(I, Col_Qty4Tel)) & "," & Val(FGrid.TextMatrix(I, Col_Qty5)) & "," & Val(FGrid.TextMatrix(I, Col_Qty5Tel)) & "," _
                    & "" & Val(FGrid.TextMatrix(I, Col_Qty6)) & "," & Val(FGrid.TextMatrix(I, Col_Qty6Tel)) & "," & Val(FGrid.TextMatrix(I, Col_Qty7)) & "," & Val(FGrid.TextMatrix(I, Col_Qty7Tel)) & "," _
                    & "" & Val(FGrid.TextMatrix(I, Col_Qty8)) & "," & Val(FGrid.TextMatrix(I, Col_Qty8Tel)) & "," & Val(FGrid.TextMatrix(I, Col_Qty9)) & "," & Val(FGrid.TextMatrix(I, Col_Qty9Tel)) & "," _
                    & "" & Val(FGrid.TextMatrix(I, Col_Qty10)) & "," & Val(FGrid.TextMatrix(I, Col_Qty10Tel)) & "," & Val(FGrid.TextMatrix(I, Col_Qty11)) & "," & Val(FGrid.TextMatrix(I, Col_Qty11Tel)) & "," _
                    & "" & Val(FGrid.TextMatrix(I, Col_Qty12)) & "," & Val(FGrid.TextMatrix(I, Col_Qty12Tel)) & "," & Val(FGrid.TextMatrix(I, Col_Qty1)) & "," & Val(FGrid.TextMatrix(I, Col_Qty1Tel)) & "," _
                    & "" & Val(FGrid.TextMatrix(I, Col_Qty2)) & "," & Val(FGrid.TextMatrix(I, Col_Qty2Tel)) & "," & Val(FGrid.TextMatrix(I, Col_Qty3)) & "," & Val(FGrid.TextMatrix(I, Col_Qty3Tel)) & "," _
                    & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
                End If
            Next
'        ElseIf Master.RecordCount  > 0 Then      ' When record Exists in the table
'            For I = 2 To FGrid.Rows - 1
'                If FGrid.TextMatrix(I, Col_ModelDesc) <> "" Then
'                    GCn.Execute "Update Veh_Forecast Set " _
'                    & "QTY_04=" & Val(FGrid.TextMatrix(I, Col_Qty4)) & ",TargQty_04=" & Val(FGrid.TextMatrix(I, Col_Qty4Tel)) & "," _
'                    & "QTY_05=" & Val(FGrid.TextMatrix(I, Col_Qty5)) & ",TargQty_05=" & Val(FGrid.TextMatrix(I, Col_Qty5Tel)) & "," _
'                    & "QTY_06=" & Val(FGrid.TextMatrix(I, Col_Qty6)) & ",TargQty_06=" & Val(FGrid.TextMatrix(I, Col_Qty6Tel)) & "," _
'                    & "QTY_07=" & Val(FGrid.TextMatrix(I, Col_Qty7)) & ",TargQty_07=" & Val(FGrid.TextMatrix(I, Col_Qty7Tel)) & "," _
'                    & "QTY_08=" & Val(FGrid.TextMatrix(I, Col_Qty8)) & ",TargQty_08=" & Val(FGrid.TextMatrix(I, Col_Qty8Tel)) & "," _
'                    & "QTY_09=" & Val(FGrid.TextMatrix(I, Col_Qty9)) & ",TargQty_09=" & Val(FGrid.TextMatrix(I, Col_Qty9Tel)) & "," _
'                    & "QTY_10=" & Val(FGrid.TextMatrix(I, Col_Qty10)) & ",TargQty_10=" & Val(FGrid.TextMatrix(I, Col_Qty10Tel)) & "," _
'                    & "QTY_11=" & Val(FGrid.TextMatrix(I, Col_Qty11)) & ",TargQty_11=" & Val(FGrid.TextMatrix(I, Col_Qty11Tel)) & "," _
'                    & "QTY_12=" & Val(FGrid.TextMatrix(I, Col_Qty12)) & ",TargQty_12=" & Val(FGrid.TextMatrix(I, Col_Qty12Tel)) & "," _
'                    & "QTY_01=" & Val(FGrid.TextMatrix(I, Col_Qty1)) & ",TargQty_01=" & Val(FGrid.TextMatrix(I, Col_Qty1Tel)) & "," _
'                    & "QTY_02=" & Val(FGrid.TextMatrix(I, Col_Qty2)) & ",TargQty_02=" & Val(FGrid.TextMatrix(I, Col_Qty2Tel)) & "," _
'                    & "QTY_03=" & Val(FGrid.TextMatrix(I, Col_Qty3)) & ",TargQty_03=" & Val(FGrid.TextMatrix(I, Col_Qty3Tel)) & "," _
'                    & "U_Name='" & pubUName & "',U_EntDt=#" & PubLoginDate & "#,U_AE='E' " _
'                    & "Where FOR_YEAR='" & ForYear & "' And MODEL='" & FGrid.TextMatrix(I, Col_ModelCode) & "'"
'                End If
'            Next
'        End If
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    SETS "INI", Me, Master
    TopCtrl1.tAdd = False
    TopCtrl1.tDel = False
    TopCtrl1.tPrn = False
    TopCtrl1.tNext = False
    TopCtrl1.tPrev = False
    TopCtrl1.tFirst = False
    TopCtrl1.tLast = False
Exit Sub
ELoop:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        SETS "INI", Me, Master
        TopCtrl1.tAdd = False
        TopCtrl1.tDel = False
        TopCtrl1.tPrn = False
        TopCtrl1.tNext = False
        TopCtrl1.tPrev = False
        TopCtrl1.tFirst = False
        TopCtrl1.tLast = False
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_GetFocus TxtGrid(Index)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid(0).Visible = False
        FGrid.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
        If TxtGridLeave = True Then
            If FGrid.Col = Col_Qty3Tel And KeyCode = vbKeyReturn Then
                If FGrid.Row <> FGrid.Rows - 1 Then
                    FGrid.Row = FGrid.Row + 1
                    FGrid.Col = Col_Qty4
                    TxtGrid(0).left = FGrid.CellLeft
                    FGrid.SetFocus
                End If
            Else
                GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 29
            End If
        Else
            TxtGrid_LostFocus 0
            TxtGrid(0).SetFocus
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    CheckQuote KeyAscii
    NumPress TxtGrid(Index), KeyAscii, 4, 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_validate TxtGrid(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = MakeBlank(Val(TxtGrid(Index).TEXT))
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
On Error GoTo ELoop
FGrid_KeyPress vbKeyReturn
'    If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
'    GridDblClick Me, FGrid, TxtGrid, 0
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_EnterCell()
    FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End If
'    If KeyCode = vbKeyReturn Then
'        GridDblClick Me, FGrid, TxtGrid, 0
'        TAddMode = False
'    End If
    If KeyCode = vbKeyDown And FGrid.Row = FGrid.Rows - 1 Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    If FGrid.Col <> 3 Then
        Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
    End If
    If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_RowColChange()
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
    End If
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub
