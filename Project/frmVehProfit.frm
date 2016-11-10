VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehProfit 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Profitability Report"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmVehProfit.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   11880
   Begin VB.CommandButton CmdPrn 
      Caption         =   "Print Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6870
      TabIndex        =   16
      Top             =   390
      Width           =   1785
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   5190
      Left            =   6630
      Negotiate       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3180
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   9155
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
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
      Caption         =   "Party Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Party Name"
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
            ColumnWidth     =   4424.882
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdFill 
      Caption         =   "Fill Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6855
      TabIndex        =   3
      Top             =   75
      Width           =   1785
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
      Left            =   3435
      MaxLength       =   25
      TabIndex        =   0
      Top             =   135
      Width           =   3360
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
      Index           =   2
      Left            =   5610
      MaxLength       =   15
      TabIndex        =   2
      Top             =   405
      Width           =   1185
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   5655
      TabIndex        =   4
      Top             =   4275
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
      Index           =   1
      Left            =   3435
      MaxLength       =   15
      TabIndex        =   1
      Top             =   405
      Width           =   1185
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   5340
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   9419
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   14
      BackColorFixed  =   12632319
      ForeColorFixed  =   16384
      BackColorSel    =   15196124
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   12632319
      GridColorFixed  =   32896
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   14
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6540
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
   End
   Begin VB.Label LblNetProfit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   9285
      TabIndex        =   13
      Top             =   6540
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Profit  -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8070
      TabIndex        =   12
      Top             =   6525
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   480
      Left            =   7890
      Shape           =   4  'Rounded Rectangle
      Top             =   6405
      Width           =   3660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   2
      Left            =   2205
      TabIndex        =   11
      Top             =   150
      Width           =   960
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   91
      Left            =   3270
      TabIndex        =   10
      Top             =   135
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   27
      Left            =   4725
      TabIndex        =   9
      Top             =   435
      Width           =   645
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   8
      Left            =   5490
      TabIndex        =   8
      Top             =   435
      Width           =   45
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
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   92
      Left            =   3270
      TabIndex        =   7
      Top             =   405
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   405
      Width           =   870
   End
End
Attribute VB_Name = "frmVehProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsParty As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim GridKey As Integer

Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const PName As Byte = 0
Private Const FromDate As Byte = 1
Private Const Uptodate As Byte = 2

' Col Declaration

Private Const SrNo As Byte = 0
Private Const CustomerName As Byte = 1
Private Const Model As Byte = 2
Private Const InvNo As Byte = 3
Private Const InvDate As Byte = 4
Private Const SalePrice As Byte = 5
Private Const SalesTax As Byte = 6
Private Const NetSalePrice  As Byte = 7
Private Const PurchPrice As Byte = 8
Private Const Offtake  As Byte = 9
Private Const GP  As Byte = 10
Private Const InsuComm  As Byte = 11
Private Const FinPayout  As Byte = 12
Private Const FinInc As Byte = 13
Private Const EBTA As Byte = 14
Private Const Retail As Byte = 15
Private Const SPInc As Byte = 16
Private Const Disc As Byte = 17
Private Const Brokrage As Byte = 18
Private Const Subvention As Byte = 19
Private Const NetProfit As Byte = 20
Private Const Chassis As Byte = 21
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim PlusVal As Double, MinusVal As Double
Private Sub CmdFill_Click()
Dim Rs As ADODB.Recordset, i As Double, Condstr As String
On Error GoTo error1
    If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
        Condstr = " and Veh_Order.DelCh_Dt  >= " & ConvertDate(Format(txt(FromDate), "dd/MMM/yyyy")) & " and Veh_Order.DelCh_Dt <= " & ConvertDate(Format(txt(Uptodate), "dd/MMM/yyyy")) & ""
    Else
        Condstr = " and Veh_Order.Inv_Date  >= " & ConvertDate(Format(txt(FromDate), "dd/MMM/yyyy")) & " and Veh_Order.Inv_Date <= " & ConvertDate(Format(txt(Uptodate), "dd/MMM/yyyy")) & ""
    End If
    If txt(PName) <> "All" Then
        Condstr = Condstr & " and Veh_Order.PartyCode='" & txt(PName).Tag & "'"
    End If
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Veh_Order.*, SubGroup.Name AS party,Veh_Stock.VRate as PurRate" & _
        " FROM (Veh_Order LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.Subcode) " & _
        " LEFT JOIN Veh_Stock ON Veh_Order.chassis = Veh_Stock.ChassisNo " & _
        " where left(Veh_Order.Inv_DocId,1)  = '" & PubDivCode & "' and len(Veh_Order.Inv_Docid) > 0" & Condstr & " order By Veh_Order.Inv_Date")
    FGrid.Rows = 1
    If Rs.RecordCount <= 0 Then MsgBox "************No Data To Display************", vbInformation: Exit Sub
    If Rs.RecordCount > 0 Then
        i = 1
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(i, 0) = i
                .TextMatrix(i, CustomerName) = Rs!Party
                .TextMatrix(i, Model) = Rs!Model
                .TextMatrix(i, InvNo) = XNull(Rs!Inv_No)
                If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
                    .TextMatrix(i, InvDate) = XNull(Rs!DelCh_DT)
                    .TextMatrix(i, SalePrice) = Format((Rs!SubTot) - Rs!Rebate, "0.00")
                Else
                    .TextMatrix(i, InvDate) = XNull(Rs!Inv_Date)
                    .TextMatrix(i, SalePrice) = Format(Rs!Net_Amount, "0.00")
                End If
                .TextMatrix(i, SalesTax) = Format(Rs!Tax_Amt + Rs!Tot_Amt, "0.00")
                .TextMatrix(i, NetSalePrice) = Format(.TextMatrix(i, SalePrice) - .TextMatrix(i, SalesTax), "0.00")
                .TextMatrix(i, PurchPrice) = Format(Rs!PurRate, "0.00")
                .TextMatrix(i, Offtake) = Format(Rs!Offtake, "0.00")
                .TextMatrix(i, GP) = Format(.TextMatrix(i, NetSalePrice) - Val(.TextMatrix(i, PurchPrice)) - .TextMatrix(i, Offtake), "0.00")
                .TextMatrix(i, InsuComm) = Format(Rs!InsComm, "0.00")
                .TextMatrix(i, FinPayout) = Format(Rs!FinPayout, "0.00")
                .TextMatrix(i, FinInc) = Format(Rs!FinInc, "0.00")
                .TextMatrix(i, EBTA) = Format(Rs!EBTA, "0.00")
                .TextMatrix(i, Retail) = Format(Rs!Retail, "0.00")
                .TextMatrix(i, SPInc) = Format(Rs!SPInc, "0.00")
                .TextMatrix(i, Disc) = Format(Rs!Rebate, "0.00")
                .TextMatrix(i, Brokrage) = Format(Rs!Brokrage, "0.00")
                .TextMatrix(i, Subvention) = Format(Rs!Subvention, "0.00")
                .TextMatrix(i, NetProfit) = "0.00"
                .TextMatrix(i, Chassis) = Rs!Chassis
           End With
            Amt_Cal i
            Rs.MoveNext
            i = i + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Set Rs = Nothing
    
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub
Private Sub CmdPrn_Click()
SpeedPrintDet
End Sub

Private Sub FGrid_RowColChange()
LblNetProfit = Format(FGrid.TextMatrix(FGrid.Row, NetProfit), "0.00")
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
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
'Dim i As Byte
   TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid

   Set RsParty = New ADODB.Recordset
   RsParty.CursorLocation = adUseClient
   RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME from SubGroup  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
   Set DGParty.DataSource = RsParty
   Ini_Grid
   TopCtrl1.TopText2 = "Edit"
   TopCtrl1.tAdd = False
   TopCtrl1.tDel = False
   TopCtrl1.tFirst = False
   TopCtrl1.tLast = False
   TopCtrl1.tNext = False
   TopCtrl1.tPrev = False
   txt(FromDate).TEXT = PubStartDate
   txt(Uptodate).TEXT = PubLoginDate
   
Exit Sub

ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        If MsgBox("Do you want to exit", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsParty = Nothing
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Select Case Index
    Case PName
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(PName) = "" Then Exit Sub
            If txt(PName) <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.FIND "name ='" & txt(PName) & "'"
            End If
End Select
txtgrid(0).Visible = False
    Ctrl_GetFocus txt(Index)
    Grid_Hide
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
Select Case Index
     Case PName
            DGridTxtKeyDown DGParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
End Select
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
If DGParty.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> Uptodate Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(keyascii)
Select Case Index
    Case PName
        If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, keyascii, "Name"
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Select Case Index
    Case FromDate
            txt(Index).TEXT = RetDate(txt(Index))
    Case Uptodate
            txt(Index).TEXT = RetDate(txt(Index))
    Case PName
        If txt(Index) = "" Then
            txt(Index) = "All"
        End If
End Select
Set Rst = Nothing
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(PName).TEXT = RsParty!Name
        txt(PName).Tag = RsParty!Code
    End If
    DGParty.Visible = False
    txt(PName).SetFocus
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case PurchPrice, Offtake, InsuComm, FinPayout, FinInc, EBTA, Retail, SPInc, Brokrage, Subvention
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Offtake, InsuComm, FinPayout, FinInc, EBTA, Retail, SPInc
            Call GridDblClick(Me, FGrid, txtgrid, 0)
            TAddMode = False
        Case PurchPrice
            If Val(FGrid.TextMatrix(FGrid.Row, PurchPrice)) = 0 Then
                Call GridDblClick(Me, FGrid, txtgrid, 0)
                TAddMode = False
            End If
        
    End Select
End If
KeyCode = 0
KeyCode = 0
End Sub

Private Sub FGrid_DblClick()
Select Case FGrid.Col
    Case Offtake, InsuComm, FinPayout, FinInc, EBTA, Retail, SPInc
        Call GridDblClick(Me, FGrid, txtgrid, 0)
    Case PurchPrice
        If Val(FGrid.TextMatrix(FGrid.Row, PurchPrice)) = 0 Then
            Call GridDblClick(Me, FGrid, txtgrid, 0)
        End If
End Select
TAddMode = False
End Sub
Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    txtgrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If FGrid.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                FGrid.RemoveItem (FGrid.Row)
        End If
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
   
FGrid.SetFocus
End If
Exit Sub
End Sub
Private Sub FGrid_KeyPress(keyascii As Integer)
Select Case FGrid.Col
    Case Offtake, InsuComm, FinPayout, FinInc, EBTA, Retail, SPInc
       Call Get_Text(Me, FGrid, txtgrid, 0, False, keyascii)
    Case PurchPrice
        If Val(FGrid.TextMatrix(FGrid.Row, PurchPrice)) = 0 Then
            Call Get_Text(Me, FGrid, txtgrid, 0, False, keyascii)
        End If
End Select
If keyascii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
txtgrid(0).Visible = False
Grid_Hide
End Sub
'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
For i = 0 To txt.Count - 1
    txt(i).TEXT = ""
Next i
End Sub

Private Sub MoveRec()
End Sub

'SrNo.|Customer Name | Model |Inv No. |Inv Date |Sale Price |Sales Tax|Net Sales Price|

Private Sub Ini_Grid()
    With FGrid
        .Cols = 22
        .left = Me.left '+45
        .top = 1000
        .RowHeightMin = PubGridRowHeight
        
        .TextMatrix(0, 0) = "SrlNo"
        .ColAlignment(0) = flexAlignCenterCenter
        .ColWidth(0) = 450
        
        .TextMatrix(0, CustomerName) = "Party Name"
        .ColAlignment(CustomerName) = flexAlignLeftCenter
        .ColWidth(CustomerName) = 2500

        .TextMatrix(0, Model) = "Model"
        .ColAlignment(Model) = flexAlignLeftCenter
        .ColWidth(Model) = 1500

        .TextMatrix(0, InvNo) = "Inv.No."
        .ColAlignment(InvNo) = flexAlignLeftCenter
        .ColWidth(InvDate) = 1500
        
        .TextMatrix(0, InvDate) = "Inv.Date"
        .ColAlignment(InvDate) = flexAlignRightCenter
        .ColWidth(InvDate) = 1500
        
        .TextMatrix(0, SalePrice) = "Sale Price"
        .ColAlignment(SalePrice) = flexAlignRightCenter
        .ColWidth(SalePrice) = 1500

        .TextMatrix(0, SalesTax) = "Sales Tax"
        .ColAlignment(SalesTax) = flexAlignRightCenter
        .ColWidth(SalesTax) = 1500
        
        .TextMatrix(0, NetSalePrice) = "Net Sale Price"
        .ColAlignment(NetSalePrice) = flexAlignRightCenter
        .ColWidth(NetSalePrice) = 1500

        .TextMatrix(0, PurchPrice) = "Purch.Price"
        .ColAlignment(PurchPrice) = flexAlignRightCenter
        .ColWidth(PurchPrice) = 1500
        
        .TextMatrix(0, Offtake) = "Offtake"
        .ColAlignment(Offtake) = flexAlignRightCenter
        .ColWidth(Offtake) = 1500

        .TextMatrix(0, GP) = "Gross Profit"
        .ColAlignment(GP) = flexAlignRightCenter
        .ColWidth(GP) = 1500

        .TextMatrix(0, InsuComm) = "Insu.Comm."
        .ColAlignment(InsuComm) = flexAlignRightCenter
        .ColWidth(InsuComm) = 1500

        .TextMatrix(0, FinPayout) = "Fin Payout"
        .ColAlignment(FinPayout) = flexAlignRightCenter
        .ColWidth(FinPayout) = 1500
        
        .TextMatrix(0, FinInc) = "Fin Inc."
        .ColAlignment(FinInc) = flexAlignRightCenter
        .ColWidth(FinInc) = 1500
        
        .TextMatrix(0, EBTA) = "EBTA"
        .ColAlignment(EBTA) = flexAlignRightCenter
        .ColWidth(EBTA) = 1500
        
        .TextMatrix(0, Retail) = "Retail"
        .ColAlignment(Retail) = flexAlignRightCenter
        .ColWidth(Retail) = 1500
        
        .TextMatrix(0, SPInc) = "SPInc"
        .ColAlignment(SPInc) = flexAlignRightCenter
        .ColWidth(SPInc) = 1500
        
        .TextMatrix(0, Disc) = "Disc."
        .ColAlignment(Disc) = flexAlignRightCenter
        .ColWidth(Disc) = 1500
        
        .TextMatrix(0, Brokrage) = "Brokrage"
        .ColAlignment(Brokrage) = flexAlignRightCenter
        .ColWidth(Brokrage) = 1500
        
        .TextMatrix(0, Subvention) = "Subvention"
        .ColAlignment(Subvention) = flexAlignRightCenter
        .ColWidth(Subvention) = 1500
        
        .TextMatrix(0, NetProfit) = "Net Profit"
        .ColAlignment(NetProfit) = flexAlignRightCenter
        .ColWidth(NetProfit) = 1500
        
        .TextMatrix(0, Chassis) = "Chassis"
        .ColAlignment(Chassis) = flexAlignRightCenter
        .ColWidth(Chassis) = 2000
        
        
End With
BackColorSelLeave = FGrid.BackColorSel
ForeColorSelEnter = FGrid.ForeColorSel
'DGParty.left = Me.left + 45: DGParty.top = FGrid.top + FGrid.height + 50
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim i As Integer
For i = 0 To txt.Count - 1
    txt(i).Enabled = Enb
    txt(i).ForeColor = CtrlFColOrg
Next
If TopCtrl1.TopText2 = "Edit" Then
    txt(Uptodate).Enabled = False
    txt(FromDate).Enabled = False
End If
txtDisabled_Color Me
txtgrid(0).BackColor = CtrlBCol
txtgrid(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    txtgrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtgrid(0) = txtgrid(0).Tag: Exit Sub
            Select Case FGrid.Col
                Case PurchPrice, Offtake, InsuComm, FinPayout, FinInc, EBTA, Retail, SPInc, Brokrage, Subvention
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, txtgrid, Index, KeyCode, TAddMode, 21
                        End If
                    End If
           End Select
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If keyascii = vbKeyEscape Then Exit Sub
Call CheckQuote(keyascii)
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case FGrid.Col
    Case PurchPrice, Offtake, InsuComm, FinPayout, FinInc, EBTA, Retail, SPInc, Brokrage, Subvention
        If KeyCode <> 13 Then TxtGrid_KeyDown Index, GridKey, 0
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    txtgrid(0).Visible = False
    Grid_Hide
End If
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub
Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim J As Integer
    Select Case FGrid.Col
        Case PurchPrice, Offtake, InsuComm, FinPayout, FinInc, EBTA, Retail, SPInc, Brokrage, Subvention
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(txtgrid(0).TEXT), "0.00")
            FGrid.TextMatrix(FGrid.Row, GP) = Format(Val(FGrid.TextMatrix(FGrid.Row, NetSalePrice)) - Val(FGrid.TextMatrix(FGrid.Row, PurchPrice)) - Val(FGrid.TextMatrix(FGrid.Row, Offtake)), "0.00")
            Amt_Cal FGrid.Row
            GCn.Execute "update Veh_Order set OffTake=" & Val(FGrid.TextMatrix(FGrid.Row, Offtake)) & ",InsComm=" & Val(FGrid.TextMatrix(FGrid.Row, InsuComm)) & ",FinPayOut=" & Val(FGrid.TextMatrix(FGrid.Row, FinPayout)) & ",FinInc=" & Val(FGrid.TextMatrix(FGrid.Row, FinInc)) & "," & _
                        "EBTA=" & Val(FGrid.TextMatrix(FGrid.Row, EBTA)) & ",Retail=" & Val(FGrid.TextMatrix(FGrid.Row, Retail)) & ",SPInc=" & Val(FGrid.TextMatrix(FGrid.Row, SPInc)) & ",Brokrage=" & Val(FGrid.TextMatrix(FGrid.Row, Brokrage)) & ",Subvention=" & Val(FGrid.TextMatrix(FGrid.Row, Subvention)) & "" & _
                        " where Veh_ORDER.Chassis='" & FGrid.TextMatrix(FGrid.Row, Chassis) & "'"
    End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        txtgrid(0).Visible = False
    End If
End Function
Private Sub SpeedPrintDet()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double, mCounter As Double
    
    Dim SalePrice1 As Double, SalesTax1 As Double, NetSPrice1 As Double, PurPrice1 As Double, Offtake1 As Double, GProfit1 As Double, InsComm1 As Double, FinPayout1 As Double
    Dim FinInc1 As Double, EBTA1 As Double, Retail1 As Double, SPInc1 As Double, Disc1 As Double, Brok1 As Double, Subvent1 As Double, NetProfit1 As Double
    
    Dim PageNo As Double
    Dim fob As New FileSystemObject
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    PageLength = PubPageLength
    PageWidth = 132
    PageNo = 1
    'Header printing
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
    If PubComp_Add2 <> "" Then
        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    
    Print #1, PRN_TIT("Vehicle Profitability Register", "C", PageWidth)
    mHeader = mHeader + 1
    Print #1, "Print Date : " & PubLoginDate
    mHeader = mHeader + 1
    Print #1, "From : " & txt(FromDate) & "  To : " & txt(Uptodate) & Space(40) & "Page No. " & PageNo
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth + 2), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("Sr.No", 5) & PSTR("Party Name", 25) & PSTR("Model", 15) & PSTR("Inv.No.", 8) & PSTR("Inv.Date", 15) & PSTR("Sale", 11, , AlignRight) & PSTR("Sales ", 11, , AlignRight) & PSTR("Net Sale", 12, , AlignRight) & PSTR("Purchase", 11, , AlignRight) & PSTR("Offtake", 9, , AlignRight) & PSTR("Gross", 11, , AlignRight) & PSTR("Insu.", 9, , AlignRight) & PSTR("Fin.", 9, , AlignRight) & PSTR("Fin.Inc.", 9, , AlignRight) & PSTR("EBTA", 9, , AlignRight) & PSTR("Retail", 9, , AlignRight) & PSTR("Special", 9, , AlignRight) & PSTR("Disc.", 9, , AlignRight) & PSTR("Brok.", 9, , AlignRight) & PSTR("Subvent.", 9, , AlignRight) & PSTR("Net", 12, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, mChr17 & Space(5) & Space(25) & Space(15) & Space(8) & Space(15) & PSTR("Price", 11, , AlignRight) & PSTR("Tax ", 11, , AlignRight) & PSTR("Price", 12, , AlignRight) & PSTR("Price", 11, , AlignRight) & Space(9) & PSTR("Profit", 11, , AlignRight) & PSTR("Comm.", 9, , AlignRight) & PSTR("Payout", 9, , AlignRight) & Space(9) & Space(9) & Space(9) & PSTR("Incent.", 9, , AlignRight) & Space(9) & Space(9) & Space(9) & PSTR("Profit", 12, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth + 2), " ", "-")
    mHeader = mHeader + 1
    mHeader = 1
    For Counter = 1 To FGrid.Rows - 1
        If mCounter <= 15 Then
            With FGrid
                Print #1, mChr17 & PSTR(.TextMatrix(Counter, 0), 5) & PSTR(.TextMatrix(Counter, CustomerName), 25) & PSTR(.TextMatrix(Counter, Model), 14) & Space(1) & PSTR(.TextMatrix(Counter, InvNo), 8) & PSTR(.TextMatrix(Counter, InvDate), 15) & PSTR(.TextMatrix(Counter, SalePrice), 11, , AlignRight) & PSTR(.TextMatrix(Counter, SalesTax), 11, , AlignRight) & PSTR(.TextMatrix(Counter, NetSalePrice), 12, , AlignRight) & PSTR(.TextMatrix(Counter, PurchPrice), 11, , AlignRight) & PSTR(.TextMatrix(Counter, Offtake), 9, , AlignRight) & PSTR(.TextMatrix(Counter, GP), 11, , AlignRight) & PSTR(.TextMatrix(Counter, InsuComm), 9, , AlignRight) & PSTR(.TextMatrix(Counter, FinPayout), 9, , AlignRight) & PSTR(.TextMatrix(Counter, FinInc), 9, , AlignRight) & _
                PSTR(.TextMatrix(Counter, EBTA), 9, , AlignRight) & PSTR(.TextMatrix(Counter, Retail), 9, , AlignRight) & PSTR(.TextMatrix(Counter, SPInc), 9, , AlignRight) & PSTR(.TextMatrix(Counter, Disc), 9, , AlignRight) & PSTR(.TextMatrix(Counter, Brokrage), 9, , AlignRight) & PSTR(.TextMatrix(Counter, Subvention), 9, , AlignRight) & PSTR(.TextMatrix(Counter, NetProfit), 12, , AlignRight) & mChr18
                mHeader = mHeader + 1
                
                SalePrice1 = Format(SalePrice1 + Val(.TextMatrix(Counter, SalePrice)), "0.00")
                SalesTax1 = Format(SalesTax1 + Val(.TextMatrix(Counter, SalesTax)), "0.00")
                NetSPrice1 = Format(NetSPrice1 + Val(.TextMatrix(Counter, NetSalePrice)), "0.00")
                PurPrice1 = Format(PurPrice1 + Val(.TextMatrix(Counter, PurchPrice)), "0.00")
                Offtake1 = Format(Offtake1 + Val(.TextMatrix(Counter, Offtake)), "0.00")
                GProfit1 = Format(GProfit1 + Val(.TextMatrix(Counter, GP)), "0.00")
                InsComm1 = Format(InsComm1 + Val(.TextMatrix(Counter, InsuComm)), "0.00")
                FinPayout1 = Format(FinPayout1 + Val(.TextMatrix(Counter, FinPayout)), "0.00")
                FinInc1 = Format(FinInc1 + Val(.TextMatrix(Counter, FinInc)), "0.00")
                EBTA1 = Format(EBTA1 + Val(.TextMatrix(Counter, EBTA)), "0.00")
                Retail1 = Format(Retail1 + Val(.TextMatrix(Counter, Retail)), "0.00")
                SPInc1 = Format(SPInc1 + Val(.TextMatrix(Counter, SPInc)), "0.00")
                Disc1 = Format(Disc1 + Val(.TextMatrix(Counter, Disc)), "0.00")
                Brok1 = Format(Brok1 + Val(.TextMatrix(Counter, Brokrage)), "0.00")
                Subvent1 = Format(Subvent1 + Val(.TextMatrix(Counter, Subvention)), "0.00")
                NetProfit1 = Format(NetProfit1 + Val(.TextMatrix(Counter, NetProfit)), "0.00")
            End With
        Else
                Print #1, Replace(Space(PageWidth + 2), " ", "-")
                mCounter = 0
                Print #1, Space(PageWidth / 2) & "Page :" & PageNo + 1
                PageNo = PageNo + 1
                Print #1, mEject
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT("Vehicle Profitability Register", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, "From : " & txt(FromDate) & "  To : " & txt(Uptodate) & Space(40) & "Page No. " & PageNo
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth + 2), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & PSTR("Sr.No", 5) & PSTR("Party Name", 25) & PSTR("Model", 15) & PSTR("Inv.No.", 8) & PSTR("Inv.Date", 15) & PSTR("Sale", 11, , AlignRight) & PSTR("Sales ", 11, , AlignRight) & PSTR("Net Sale", 12, , AlignRight) & PSTR("Purchase", 11, , AlignRight) & PSTR("Offtake", 9, , AlignRight) & PSTR("Gross", 11, , AlignRight) & PSTR("Insu.", 9, , AlignRight) & PSTR("Fin.", 9, , AlignRight) & PSTR("Fin.Inc.", 9, , AlignRight) & PSTR("EBTA", 9, , AlignRight) & PSTR("Retail", 9, , AlignRight) & PSTR("Special", 9, , AlignRight) & PSTR("Disc.", 9, , AlignRight) & PSTR("Brok.", 9, , AlignRight) & PSTR("Subvent.", 9, , AlignRight) & PSTR("Net", 12, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, mChr17 & Space(5) & Space(25) & Space(15) & Space(8) & Space(15) & PSTR("Price", 11, , AlignRight) & PSTR("Tax ", 11, , AlignRight) & PSTR("Price", 12, , AlignRight) & PSTR("Price", 11, , AlignRight) & Space(9) & PSTR("Profit", 11, , AlignRight) & PSTR("Comm.", 9, , AlignRight) & PSTR("Payout", 9, , AlignRight) & Space(9) & Space(9) & Space(9) & PSTR("Incent.", 9, , AlignRight) & Space(9) & Space(9) & Space(9) & PSTR("Profit", 12, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth + 2), " ", "-")
                mHeader = mHeader + 1
        End If
   
    Next
    Print #1, Replace(Space(PageWidth + 2), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & Space(5) & Space(25) & PSTR("Grand Total -->", 22) & Space(16) & PSTR(Format(SalePrice1, "0.00"), 11, , AlignRight) & PSTR(Format(SalesTax1, "0.00"), 11, , AlignRight) & PSTR(Format(NetSPrice1, "0.00"), 12, , AlignRight) & PSTR(Format(PurPrice1, "0.00"), 11, , AlignRight) & PSTR(Format(Offtake1, "0.00"), 9, , AlignRight) & PSTR(Format(GProfit1, "0.00"), 11, , AlignRight) & PSTR(Format(InsComm1, "0.00"), 9, , AlignRight) & PSTR(Format(FinPayout1, "0.00"), 9, , AlignRight) & PSTR(Format(FinInc1, "0.00"), 9, , AlignRight) & PSTR(Format(EBTA1, "0.00"), 9, , AlignRight) & PSTR(Format(Retail1, "0.00"), 9, , AlignRight) & PSTR(Format(SPInc1, "0.00"), 9, , AlignRight) & PSTR(Format(Disc1, "0.00"), 9, , AlignRight) & PSTR(Format(Brok1, "0.00"), 9, , AlignRight) & PSTR(Format(Subvent1, "0.00"), 9, , AlignRight) & PSTR(Format(NetProfit1, "0.00"), 12, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth + 2), " ", "-")
    mHeader = mHeader + 1
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt > Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
    End If
End Sub
Private Function Amt_Cal(mROW As Double)
    PlusVal = Val(FGrid.TextMatrix(mROW, GP)) + Val(FGrid.TextMatrix(mROW, InsuComm)) + Val(FGrid.TextMatrix(mROW, FinPayout)) + Val(FGrid.TextMatrix(mROW, FinInc)) + Val(FGrid.TextMatrix(mROW, EBTA)) + Val(FGrid.TextMatrix(mROW, Retail)) + Val(FGrid.TextMatrix(mROW, SPInc))
    MinusVal = Val(FGrid.TextMatrix(mROW, Brokrage)) + Val(FGrid.TextMatrix(mROW, Subvention))
    FGrid.TextMatrix(mROW, NetProfit) = Format(PlusVal - MinusVal, "0.00")
    LblNetProfit = Format(FGrid.TextMatrix(mROW, NetProfit), "0.00")
End Function
