VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form RepIncClaim 
   BackColor       =   &H00BBDBB3&
   Caption         =   "Offtake Incentive Claims"
   ClientHeight    =   7335
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11610
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1185
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "012345"
      Top             =   480
      Width           =   855
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   661
   End
   Begin VB.CommandButton CmdTrn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Display Purchases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   435
      Width           =   1995
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   5
      Left            =   1710
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "RepIncClaim.frx":0000
      Top             =   6135
      Width           =   9720
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1710
      MaxLength       =   40
      TabIndex        =   9
      Text            =   "0123456789012345678901234567890123456789"
      Top             =   6855
      Width           =   4725
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   8505
      MaxLength       =   25
      TabIndex        =   10
      Text            =   "0123456789012345678901234"
      Top             =   6855
      Width           =   2925
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6150
      MaxLength       =   70
      TabIndex        =   5
      Top             =   1290
      Width           =   5280
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   2
      Left            =   480
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "RepIncClaim.frx":0009
      Top             =   870
      Width           =   4305
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   4
      Left            =   1710
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "RepIncClaim.frx":0023
      Top             =   1560
      Width           =   9720
   End
   Begin VB.TextBox Txt 
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
      Left            =   4755
      MaxLength       =   12
      TabIndex        =   2
      Text            =   "01/Jan/2003"
      Top             =   480
      Width           =   1185
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
      Left            =   2520
      TabIndex        =   11
      Top             =   4815
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3075
      Left            =   90
      TabIndex        =   7
      Top             =   2295
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   5424
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   64
      Cols            =   9
      BackColorFixed  =   12640511
      ForeColorFixed  =   128
      BackColorSel    =   15718112
      ForeColorSel    =   8388608
      BackColorBkg    =   13623520
      GridColor       =   8438015
      GridColorFixed  =   16512
      GridColorUnpopulated=   16711935
      FocusRect       =   0
      GridLinesFixed  =   1
      BorderStyle     =   0
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblTotInc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Inc.Rs.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   9585
      TabIndex        =   24
      Top             =   930
      Width           =   1065
   End
   Begin VB.Label lblTgtInc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgt Incentive Rs.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   7230
      TabIndex        =   23
      Top             =   930
      Width           =   1395
   End
   Begin VB.Label lblInc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incentive Rs.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   5010
      TabIndex        =   22
      Top             =   930
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No.* :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   2
      Left            =   135
      TabIndex        =   21
      Top             =   495
      Width           =   975
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   8595
      TabIndex        =   20
      Top             =   450
      Width           =   660
   End
   Begin VB.Label LblSite 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   10335
      TabIndex        =   19
      Top             =   450
      Width           =   810
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   11475
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kind Attention :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   10
      Left            =   4890
      TabIndex        =   18
      Top             =   1290
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Footer Message :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   8
      Left            =   180
      TabIndex        =   17
      Top             =   6090
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By* :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   7
      Left            =   1320
      TabIndex        =   16
      Top             =   6855
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   6
      Left            =   7335
      TabIndex        =   15
      Top             =   6855
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Month/Date* :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   3
      Left            =   2760
      TabIndex        =   14
      Top             =   495
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To* :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   13
      Top             =   870
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Header Message* :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   12
      Top             =   1530
      Width           =   1590
   End
End
Attribute VB_Name = "RepIncClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TAddMode As Boolean
Dim GridKey As Integer
Dim Master As ADODB.Recordset
Dim mDocType$

'grid color scheme
Private Const BackColorSelEnter As String = &HF8D7FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const DispPurch As Byte = 0
Private Const PrintLet As Byte = 1
Private Const ExitLet As Byte = 2

Private Const Srlno As Byte = 0              ' Letter Serial No. for Uniqueness
Private Const Date1 As Byte = 1              ' Date To
Private Const AddressTo As Byte = 2          ' AddressTo
Private Const KindAttn As Byte = 3           ' Kind Attn
Private Const Header As Byte = 4             ' Header Text
Private Const Footer As Byte = 5             ' Footer Text
Private Const LetterBy As Byte = 6           ' Letter By
Private Const Designation As Byte = 7        ' Designation

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_Select As Byte = 1            ' Select
Private Const Col_Model As Byte = 2             ' Model
Private Const Col_ChassisNo As Byte = 3         ' ChassisNo
Private Const Col_EgineNo As Byte = 4           ' EngineNo
Private Const Col_TelcoBNo As Byte = 5          ' TelcoBNo
Private Const Col_TelcoDate As Byte = 6         ' Telco Date
Private Const Col_Incentive As Byte = 7         ' Incentive
Private Const Col_TargetInc As Byte = 8         ' Target Link Incentive
Private Const Col_TotInc As Byte = 9            ' Tot Incentive

Private Sub Disp_Text(Enb As Boolean)
    
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
    CmdTrn(0).Enabled = False

End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub BlankText()
Dim I As Byte
'* Used for clear all text boxes used in the form
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
    Next I
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
    With FGrid
        .left = Me.left '+ 45
        .width = Me.width - 150
        .top = 2295
        .height = FGrid.RowHeight(0) * 14
        .RowHeightMin = 0 'PubGridRowHeight
        .Cols = 11
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        
        .TextMatrix(0, Col_SrNo) = "S.No."
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 480
        
        .TextMatrix(0, Col_Select) = "Select"
        .ColAlignment(Col_Select) = flexAlignLeftCenter
        .ColWidth(Col_Select) = 0

        .TextMatrix(0, Col_Model) = "Model"
        .ColAlignment(Col_Model) = flexAlignLeftCenter
        .ColWidth(Col_Model) = 1755
        
        .TextMatrix(0, Col_ChassisNo) = "ChassisNo"
        .ColAlignment(Col_ChassisNo) = flexAlignLeftCenter
        .ColWidth(Col_ChassisNo) = 1725
        
        .TextMatrix(0, Col_EgineNo) = "EgineNo"
        .ColAlignment(Col_EgineNo) = flexAlignLeftCenter
        .ColWidth(Col_EgineNo) = 1800

        .TextMatrix(0, Col_TelcoBNo) = "Telco BNo."
        .ColAlignment(Col_TelcoBNo) = flexAlignLeftCenter
        .ColWidth(Col_TelcoBNo) = 1455

        .TextMatrix(0, Col_TelcoDate) = "Telco Date"
        .ColAlignment(Col_TelcoDate) = flexAlignLeftCenter
        .ColWidth(Col_TelcoDate) = 1335
        
        .TextMatrix(0, Col_Incentive) = "Incentive"
        .ColAlignment(Col_Incentive) = flexAlignRightCenter
        .ColWidth(Col_Incentive) = 900

        .TextMatrix(0, Col_TargetInc) = "TgtLinkInc"
        .ColAlignment(Col_TargetInc) = flexAlignRightCenter
        .ColWidth(Col_TargetInc) = 900
        
        .TextMatrix(0, Col_TotInc) = "Total Incen"
        .ColAlignment(Col_TotInc) = flexAlignRightCenter
        .ColWidth(Col_TotInc) = 1000
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
End Sub

Private Sub FillData(Optional MoveRecCall As Boolean)

Dim Rst As ADODB.Recordset, I As Integer, mTotInc As Double
On Error GoTo ELoop
    GSQL = "Select Model,ChassisNo,EngineNo,PBILL_NO,PBILL_DATE,OfftakeIncentive,TgtLinkIncentive " & _
        " From Veh_Stock Where "
    If MoveRecCall Then
        GSQL = GSQL & " OfftakeIncentiveSrlNo='" & Txt(Srlno) & "' Order By Model, ChassisNo"
    Else
        GSQL = GSQL & " " & cMID("Veh_Stock.Pur_DocId", "1", "1") & "='" & PubDivCode & "' and " & cCStr("Month(Pur_Vdate)") & " + " & cCStr("Year(Pur_Vdate)") & " ='" & Format(Txt(Date1), "MYYYY") & "' Order By Model, ChassisNo"
    End If
    Set Rst = GCn.Execute(GSQL)
    FGrid.Redraw = False
    FGrid.Rows = 1
    If Rst.RecordCount <= 0 Then
        Set Rst = Nothing
        MsgBox "No Records Found!", vbOKOnly, "Validation"
         GoTo lblExit
    End If
    I = 1
    Do Until Rst.EOF
        mTotInc = IIf(IsNull(Rst!OfftakeIncentive), 0, Rst!OfftakeIncentive)
        mTotInc = mTotInc + IIf(IsNull(Rst!TgtLinkIncentive), 0, Rst!TgtLinkIncentive)
        FGrid.AddItem ""
        With FGrid
            .TextMatrix(I, Col_SrNo) = I
            .TextMatrix(I, Col_Select) = "" 'Rst!FormNo
            .TextMatrix(I, Col_Model) = Rst!Model
            .TextMatrix(I, Col_ChassisNo) = Rst!ChassisNo
            .TextMatrix(I, Col_EgineNo) = Rst!EngineNo
            .TextMatrix(I, Col_TelcoBNo) = Rst!PBILL_NO
            .TextMatrix(I, Col_TelcoDate) = Rst!PBILL_DATE
            .TextMatrix(I, Col_Incentive) = IIf(IsNull(Rst!OfftakeIncentive) Or Rst!OfftakeIncentive <= 0, "", Format(Rst!OfftakeIncentive, "0.00"))
            .TextMatrix(I, Col_TargetInc) = IIf(IsNull(Rst!TgtLinkIncentive) Or Rst!TgtLinkIncentive <= 0, "", Format(Rst!TgtLinkIncentive, "0.00"))
            .TextMatrix(I, Col_TotInc) = IIf(IsNull(mTotInc) Or mTotInc <= 0, "", Format(mTotInc, "0.00"))
        End With
        Rst.MoveNext
        I = I + 1
    Loop
    AmtCal

lblExit:
    If FGrid.Rows = 1 Then FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    FGrid.Redraw = True
    Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub MoveRec()
Dim Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        Txt(Srlno) = Master!Srlno
        Txt(Date1) = Master!PurMonthDate
        Txt(AddressTo) = Master!AddressTo
        Txt(KindAttn) = Master!KindAttn
        Txt(Header) = Master!HeaderText
        Txt(Footer) = Master!FooterText
        Txt(LetterBy) = Master!LetterBy
        Txt(Designation) = Master!Designation
        LblDiv.CAPTION = "Division : " & mID(Master!Srlno, 1, 1)
        LblSite.CAPTION = "Site Code : " & Master!Site_Code
        CmdTrn(0).Enabled = False
        FillData True
    Else
        BlankText
        FGrid.Rows = 1
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
Set Rst = Nothing
Exit Sub
ELoop:
    FGrid.Redraw = True
    CheckError
End Sub

Private Sub CmdTrn_Click(Index As Integer)
Select Case Index
    Case DispPurch
        'Date checking
        If IsValid(Txt(Srlno), "SrlNo") = False Then Exit Sub
        If IsValid(Txt(Date1), "Purchase Month/Date") = False Then Exit Sub
        If CheckFinYear(CDate(Txt(Date1))) = False Then Txt(Date1).SetFocus: Exit Sub
        'eof date checking

        FillData
        Disp_Text True
        Txt(AddressTo).SetFocus
End Select
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
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        Txt(I).ForeColor = CtrlFColOrg
    Next
    mDocType = PubDivCode & "I"
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
'    Set Master = GCn.Execute("Select SrlNo As SearchCode, Veh_OfftakeIncentive.* , '" & Month(Veh_OfftakeIncentive.PurMonthDate) & Year(Veh_OfftakeIncentive.PurMonthDate) & "' AS MTHYEAR from Veh_OfftakeIncentive Where mid(Veh_OfftakeIncentive.srlNo,1,2)='CI' Order by PurMonthDate,SrlNo ")
     Set Master = GCn.Execute("Select SrlNo As SearchCode, Veh_OfftakeIncentive.* , (Month(PurMonthDate)+Year(PurMonthDate))  AS MTHYEAR from Veh_OfftakeIncentive Where " & cMID("Veh_OfftakeIncentive.srlNo", "1", "2") & "='PI' Order by PurMonthDate,SrlNo ")
      
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    
'    BlankText
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    Txt(Srlno) = PubDivCode & "I"
    Txt(Srlno).SelStart = Len(Txt(Srlno))
    CmdTrn(0).Enabled = True
    Txt(Srlno).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    Txt(Srlno).Enabled = False
    Txt(Date1).Enabled = False
    CmdTrn(0).Enabled = False
    Txt(AddressTo).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant, mTrans As Boolean
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
            mTrans = True
            GCn.Execute ("Update Veh_Stock Set OfftakeIncentiveSrlNo='', OfftakeIncentive=0,TgtLinkIncentive=0 " & _
                " where OfftakeIncentiveSrlNo='" & Txt(Srlno) & "'")
            GCn.Execute ("Delete from Veh_OfftakeIncentive where SrlNo='" & Txt(Srlno) & "'")
            GCn.CommitTrans
            mTrans = False
            Master.Requery
            If Master.RecordCount > 0 Then
                If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            End If
            BUTTONS True, Me, Master, 0
            MoveRec
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub

Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub

Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select SrlNo As SearchCode,SrlNo, PurMonthDate From Veh_OfftakeIncentive Where " & cMID("Veh_OfftakeIncentive.srlNo", "1", "2") & "='" & mDocType & "' Order by " & cCStr("Month(PurMonthDate)") & " + " & cCStr("Year(PurMonthDate)") & ""
    Set SearchForm = Me
    FIND.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
Dim Rst As ADODB.Recordset, mQRY$, mRepName$
Dim I As Integer, RstRep As ADODB.Recordset
On Error GoTo ERRORHANDLER
mRepName = "OffIncClaim"


        mQRY = "Select '" & Txt(Srlno) & "' as SrlNo,Model,ChassisNo,EngineNo,PBILL_NO,PBILL_DATE,OfftakeIncentive,TgtLinkIncentive,'" & Txt(AddressTo) & "' as AddressTo, '" & Txt(Header) & "' as HeaderText,'" & Txt(Footer) & "' as FooterText " & _
            " From Veh_Stock Where OfftakeIncentiveSrlNo='" & Txt(Srlno) & "' Order By Model, ChassisNo"
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Set Rst = Nothing: Exit Sub
                
        'Create temp table
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "SrlNo", adChar, 6, adFldIsNullable
            .Fields.Append "Model", adChar, 15, adFldIsNullable
            .Fields.Append "ChassisNo", adChar, 15, adFldIsNullable
            .Fields.Append "EngineNo", adChar, 20, adFldIsNullable
            .Fields.Append "PBill_No", adChar, 10, adFldIsNullable
            .Fields.Append "PBill_Date", adDate, 7, adFldIsNullable
            .Fields.Append "OfftakeIncentive", adDouble, 12, adFldIsNullable
            .Fields.Append "TgtLinkIncentive", adDouble, 12, adFldIsNullable
            .Fields.Append "AddressTo", adChar, 120, adFldIsNullable
            .Fields.Append "HeaderText", adChar, 255, adFldIsNullable
            .Fields.Append "FooterText", adChar, 255, adFldIsNullable
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
        End With
        'temp table created
        
        Do While Rst.EOF = False
            With RstRep
                .AddNew
                .Fields("SrlNo") = Rst!Srlno
                .Fields("Model") = Rst!Model
                .Fields("ChassisNo") = Rst!ChassisNo
                .Fields("EngineNo") = Rst!EngineNo
                .Fields("PBill_No") = Rst!PBILL_NO
                .Fields("PBill_Date") = Rst!PBILL_DATE
                .Fields("OfftakeIncentive") = Rst!OfftakeIncentive
                .Fields("TgtLinkIncentive") = Rst!TgtLinkIncentive
                .Fields("AddressTo") = Rst!AddressTo
                .Fields("HeaderText") = Rst!HeaderText
                .Fields("FooterText") = Rst!FooterText
                .Update
            End With
            Rst.MoveNext
        Loop
        
        CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".ttx", True
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("comp_name")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
                Case UCase("comp_add1")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
                Case UCase("comp_add2")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_Add2 & "'"
                Case UCase("comp_city")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
                Case UCase("KindAttn")
                    rpt.FormulaFields(I).TEXT = "'" & Txt(KindAttn) & "'"
                Case UCase("LetterBy")
                    rpt.FormulaFields(I).TEXT = "'" & Txt(LetterBy) & "'"
                Case UCase("Designation")
                    rpt.FormulaFields(I).TEXT = "'" & Txt(Designation) & "'"
            End Select
        Next
        Report_View rpt, Me.CAPTION, , False
Set Rst = Nothing
Set RstRep = Nothing
Exit Sub
ERRORHANDLER:
        CheckError

End Sub

Private Sub TopCtrl1_eRef()
On Error GoTo ELoop
    Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, mGridFilled As Boolean, mTotInc As Double
Dim Rst As ADODB.Recordset, DocIdHlp$, TmpStr$
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If

    If Len(Txt(Srlno)) <= 2 Then
        MsgBox "Serial No. is required", vbCritical, "Duplicate Serial No"
        Txt(Srlno).SetFocus: Exit Sub
    End If
    If IsValid(Txt(Date1), "Purchase Month/Date") = False Then Exit Sub
    If IsValid(Txt(AddressTo), "Address To") = False Then Exit Sub
    If IsValid(Txt(Header), "Header Message") = False Then Exit Sub
    If IsValid(Txt(LetterBy), "Letter By") = False Then Exit Sub
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_TotInc) <> "" Then
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Incentive Details", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_Incentive: FGrid.SetFocus: Exit Sub
    If TopCtrl1.TopText2 = "Add" Then
        If GCn.Execute("Select SrlNo From Veh_OfftakeIncentive Where SrlNo='" & Txt(Srlno) & "'").RecordCount > 0 Then
            MsgBox "Serial No. Already Exists", vbCritical, "Validation Error"
            Txt(Srlno).SetFocus
            Exit Sub
        End If
    End If
    
    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Add" Then
            GCn.Execute ("insert into Veh_OfftakeIncentive (SrlNo,PurMonthDate,AddressTo," & _
                " KindAttn,HeaderText,FooterText,LetterBy," & _
                " Designation,Site_Code,U_Name,U_EntDt,U_AE) " & _
                " Values('" & Txt(Srlno) & "', " & ConvertDate(Txt(Date1)) & ",'" & Txt(AddressTo) & _
                "','" & Txt(KindAttn) & "','" & Txt(Header) & "','" & Txt(Footer) & "','" & Txt(LetterBy) & _
                "','" & Txt(Designation) & "','" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
        Else
            GCn.Execute ("Update Veh_OfftakeIncentive set AddressTo='" & Txt(AddressTo) & _
                "',KindAttn='" & Txt(KindAttn) & "',HeaderText='" & Txt(Header) & "',FooterText='" & Txt(Footer) & _
                "',LetterBy='" & Txt(LetterBy) & "',Designation='" & Txt(Designation) & "',Site_Code='" & PubSiteCode & _
                "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & _
                " where SrlNo = '" & Txt(Srlno) & "'")
        End If
        'Update Veh_Stock
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_ChassisNo) <> "" Then
                mTotInc = Val(FGrid.TextMatrix(I, Col_Incentive)) + Val(FGrid.TextMatrix(I, Col_TargetInc))
                If mTotInc > 0 Then
                    GSQL = "update Veh_Stock set OfftakeIncentiveSrlNo='" & Txt(Srlno) & "', OfftakeIncentive=" & Val(FGrid.TextMatrix(I, Col_Incentive)) & ", TgtLinkIncentive= " & Val(FGrid.TextMatrix(I, Col_TargetInc)) & _
                        " where " & cMID("Veh_Stock.Pur_DocId", "1", "1") & "='" & PubDivCode & "' and " & cCStr("Month(Pur_Vdate)") & " + " & cCStr("Year(Pur_Vdate)") & "='" & Format(Txt(Date1), "MYYYY") & "' and ChassisNo='" & FGrid.TextMatrix(I, Col_ChassisNo) & "'"
                Else
                    GSQL = "update Veh_Stock set OfftakeIncentiveSrlNo='', OfftakeIncentive=0, TgtLinkIncentive= 0 " & _
                        " where OfftakeIncentiveSrlNo='" & Txt(Srlno) & "' and ChassisNo='" & FGrid.TextMatrix(I, Col_ChassisNo) & "'"
                End If
                GCn.Execute (GSQL)
            End If
        Next
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    Master.FIND "SearchCode = '" & Txt(Srlno) & "'"
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus Txt(Index)
TxtGrid(0).Visible = False
Select Case Index
    Case AddressTo
'        ListArray = Array("General", "Warranty")
'        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
    Case KindAttn
'        If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
'        If Txt(Index).Text <> RsJob!Name Then
'            RsJob.MoveFirst
'            RsJob.FIND "Name ='" & Txt(Index).Text & "'"
'        End If
    Case Header
'        If RsMech.RecordCount = 0 Or (RsMech.EOF = True Or RsMech.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
'        If Txt(Index).Text <> RsMech!Name Then
'            RsMech.MoveFirst
'            RsMech.FIND "Name ='" & Txt(Index).Text & "'"
'        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case Srlno
        'Div Code + Type Char Edit restricted
        KeyCode = RestrictCode(KeyCode, Txt(Index), Shift, True)
End Select
    If (Txt(Index).MultiLine And KeyCode = vbKeyTab) Or Txt(Index).MultiLine = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Designation Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
                Ctrl_DownKeyDown KeyCode, Shift
            End If
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
If keyascii = 39 Then keyascii = 0: Exit Sub
Select Case Index
    Case Srlno
        keyascii = RestrictCode(keyascii, Txt(Index), 0, True)
'    Case SrlNoTo
'        NumPress Txt(Index), KeyAscii, 6, 0
End Select
Exit Sub

ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
'    Select Case Index
'    Case DocType
'        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
'    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
'Validate- >LostFocus
On Error GoTo ELoop
Select Case Index
    Case Date1
        Txt(Index).TEXT = RetDate(Txt(Index))
        Master.FIND ("MTHYEAR = " & Month(Txt(Index).TEXT) + Year(Txt(Index).TEXT))
        If Not Master.EOF Then
              MsgBox "For this month claim is already exist!", vbOKOnly, "Validation"
              Txt(Date1).SetFocus
              Cancel = True
        End If
   
   Case Header
        FGrid.Col = Col_Incentive
        FGrid.Row = 1
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus TxtGrid(Index)
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        TxtGrid(0).Visible = False
        FGrid.SetFocus
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Col_Incentive, Col_TargetInc
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_TotInc
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
CheckQuote keyascii
Select Case FGrid.Col
    Case Col_Incentive, Col_TargetInc
        NumPress TxtGrid(Index), keyascii, 6, 2
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_Incentive, Col_TargetInc
        If TxtGrid(Index) <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(TxtGrid(Index).TEXT, "0.00")
        If Val(FGrid.TextMatrix(FGrid.Row, Col_Incentive)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TargetInc)) > 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_TotInc) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_Incentive)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TargetInc)), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, Col_TotInc) = ""
        End If
End Select
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
Dim j As Integer
Select Case FGrid.Col
    Case Col_Incentive, Col_TargetInc
        If TxtGrid(Index) <> "" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(TxtGrid(Index).TEXT, "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End If
        If Val(FGrid.TextMatrix(FGrid.Row, Col_Incentive)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TargetInc)) > 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_TotInc) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_Incentive)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TargetInc)), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, Col_TotInc) = ""
        End If
        AmtCal
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
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
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
    If KeyCode = vbKeyDelete And Shift = 0 Then 'Delete Key
        Select Case FGrid.Col
            Case Col_Incentive, Col_TargetInc
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                If Val(FGrid.TextMatrix(FGrid.Row, Col_Incentive)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TargetInc)) > 0 Then
                    FGrid.TextMatrix(FGrid.Row, Col_TotInc) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_Incentive)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TargetInc)), "0.00")
                Else
                    FGrid.TextMatrix(FGrid.Row, Col_TotInc) = ""
                End If
                AmtCal
        End Select
    ElseIf KeyCode = vbKeyReturn Then
        TAddMode = False
        FGrid_KeyPress KeyCode
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_Incentive, Col_TargetInc
        Get_Text Me, FGrid, TxtGrid, 0, True, keyascii
End Select
If keyascii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim I As Integer, mSrlNo As Integer
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid.Row >= 1 Then
            If MsgBox("Are You Sure To Delete Current Row?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid.Rows > 2 Then
                    FGrid.RemoveItem (FGrid.Row)
                Else
                    FGrid.Rows = 1
                    FGrid.AddItem FGrid.Rows
                    FGrid.FixedRows = 1
                End If
                For I = 1 To FGrid.Rows - 1
                    mSrlNo = mSrlNo + 1
                    FGrid.TextMatrix(I, Col_SrNo) = mSrlNo
                Next
                AmtCal
            End If
            FGrid.Redraw = True
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
End Sub
Private Sub AmtCal()
Dim I As Integer, mInc As Double, mTgtInc As Double
    For I = 1 To FGrid.Rows - 1
        mInc = Val(FGrid.TextMatrix(I, Col_Incentive))
        mTgtInc = Val(FGrid.TextMatrix(I, Col_TargetInc))
    Next
    lblInc.CAPTION = "Incentive Rs.:" & IIf(mInc <= 0, "", Format(mInc, "0.00"))
    lblTgtInc.CAPTION = "Tgt Incentive Rs.:" & IIf(mTgtInc <= 0, "", Format(mTgtInc, "0.00"))
    lblTotInc.CAPTION = "Total Inc.Rs.:" & IIf(mInc + mTgtInc <= 0, "", Format(mInc + mTgtInc, "0.00"))
    lblInc.Refresh
    lblTgtInc.Refresh
    lblTotInc.Refresh
End Sub
