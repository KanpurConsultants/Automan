VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form RepSubvention 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Subvention Report / Letter"
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
   LockControls    =   -1  'True
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6015
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "99.999"
      Top             =   480
      Width           =   630
   End
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
      Caption         =   "Display Sales"
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
      Left            =   6765
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   435
      Width           =   1635
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
      TabIndex        =   9
      Text            =   "RepSubvention.frx":0000
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
      TabIndex        =   10
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
      TabIndex        =   11
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
      TabIndex        =   6
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
      TabIndex        =   5
      Text            =   "RepSubvention.frx":0009
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
      TabIndex        =   7
      Text            =   "RepSubvention.frx":0023
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
      Left            =   3810
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
      TabIndex        =   12
      Top             =   4815
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3075
      Left            =   90
      TabIndex        =   8
      Top             =   2565
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
      AllowUserResizing=   1
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
   Begin VB.Label lblParty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer  :"
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
      Left            =   105
      TabIndex        =   26
      Top             =   2280
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telco % :"
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
      Index           =   1
      Left            =   5220
      TabIndex        =   25
      Top             =   495
      Width           =   750
   End
   Begin VB.Label lblTotInc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Telco Share Rs.:"
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
      Left            =   8610
      TabIndex        =   24
      Top             =   930
      Width           =   1815
   End
   Begin VB.Label lblInc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Finance Amount Rs.:"
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
      Left            =   4890
      TabIndex        =   23
      Top             =   930
      Width           =   2160
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
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   480
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
      TabIndex        =   20
      Top             =   480
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   6855
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Month/Date* :"
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
      Left            =   2220
      TabIndex        =   15
      Top             =   495
      Width           =   1515
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   1530
      Width           =   1590
   End
End
Attribute VB_Name = "RepSubvention"
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
Private Const TelcoShare As Byte = 8         ' Telco Share

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_Model As Byte = 1             ' Model
Private Const Col_ChassisNo As Byte = 2         ' ChassisNo
Private Const Col_EgineNo As Byte = 3           ' EngineNo
Private Const Col_InvNo As Byte = 4          ' Sale Invoice BNo
Private Const Col_InvDate As Byte = 5         ' Sale Invoice Date
Private Const Col_FinAmt As Byte = 6         ' Fin Amt
Private Const Col_TelcoShare As Byte = 7         ' TelcoShare
Private Const Col_Financier As Byte = 8            ' Financier
Private Const Col_Party As Byte = 9            ' Customer
Private Const Col_SalDocID As Byte = 10           ' Sale Doc Id

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
        .top = 2565
        .height = FGrid.RowHeight(0) * 13
        .RowHeightMin = 0 'PubGridRowHeight
        .Cols = 11
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        
        .TextMatrix(0, Col_SrNo) = "S.No."
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 480
        
        .TextMatrix(0, Col_Model) = "Model"
        .ColAlignment(Col_Model) = flexAlignLeftCenter
        .ColWidth(Col_Model) = 1650
        
        .TextMatrix(0, Col_ChassisNo) = "ChassisNo"
        .ColAlignment(Col_ChassisNo) = flexAlignLeftCenter
        .ColWidth(Col_ChassisNo) = 1605
        
        .TextMatrix(0, Col_EgineNo) = "EgineNo"
        .ColAlignment(Col_EgineNo) = flexAlignLeftCenter
        .ColWidth(Col_EgineNo) = 1800

        .TextMatrix(0, Col_InvNo) = "InvoiceNo."
        .ColAlignment(Col_InvNo) = flexAlignLeftCenter
        .ColWidth(Col_InvNo) = 1140

        .TextMatrix(0, Col_InvDate) = "InvDate"
        .ColAlignment(Col_InvDate) = flexAlignLeftCenter
        .ColWidth(Col_InvDate) = 1125
        
        .TextMatrix(0, Col_FinAmt) = "Fin.Amt."
        .ColAlignment(Col_FinAmt) = flexAlignRightCenter
        .ColWidth(Col_FinAmt) = 945

        .TextMatrix(0, Col_TelcoShare) = "Mfg.Share"
        .ColAlignment(Col_TelcoShare) = flexAlignRightCenter
        .ColWidth(Col_TelcoShare) = 900
        
        .TextMatrix(0, Col_Financier) = "Financier"
        .ColAlignment(Col_Financier) = flexAlignLeftCenter
        .ColWidth(Col_Financier) = 2625
        
        .TextMatrix(0, Col_Party) = "Name of Customer"
        .ColAlignment(Col_Party) = flexAlignLeftCenter
        .ColWidth(Col_Party) = 2625
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
End Sub

Private Sub FillData(Optional MoveRecCall As Boolean)

Dim Rst As ADODB.Recordset, I As Integer, mFinAmt As Double, mTelcoShare As Double
On Error GoTo ELoop
    GSQL = "Select VStk.Model,VStk.ChassisNo,VStk.EngineNo,VO.Inv_Prefix,VStk.Sal_DocId,VO.Inv_Date,VO.FIN_AMT,VStk.MfgShare,CF.FinName,SG.Name " & _
        " From ((Veh_Stock VStk left join Veh_Order VO on VStk.Sal_DocId=VO.Inv_DocId) " & _
        " left join  ContractFinance CF on VO.FB_CODE=CF.FinCode) " & _
        " left Join SubGroup SG on VO.PartyCode=SG.SubCode " & _
        " Where VStk.Sal_DocId<>'' and "
    If MoveRecCall Then
        GSQL = GSQL & " SubventionSrlNo='" & Txt(Srlno) & "' Order By VStk.Model,VStk.ChassisNo"
    Else
        GSQL = GSQL & " VO.Fund_Source<>2 and " & cMID("VStk.Sal_DocId", "1", "1") & "='" & PubDivCode & "' and " & cCStr("Month(VO.Inv_Date)") & " + " & cCStr("Year(VO.Inv_Date)") & "='" & Format(Txt(Date1), "MYYYY") & "' Order By VStk.Model,VStk.ChassisNo"
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
        mFinAmt = IIf(IsNull(Rst!Fin_Amt), 0, Rst!Fin_Amt)
        If mFinAmt > 0 Then
            If MoveRecCall Then
                mTelcoShare = Rst!MfgShare
            Else
                mTelcoShare = Round((mFinAmt * Val(Txt(TelcoShare))) / 100, 0)
            End If
        Else
            mTelcoShare = 0
        End If
        FGrid.AddItem ""
        With FGrid
            .TextMatrix(I, Col_SrNo) = I
            .TextMatrix(I, Col_Party) = Rst!Name
            .TextMatrix(I, Col_Model) = Rst!Model
            .TextMatrix(I, Col_ChassisNo) = Rst!ChassisNo
            .TextMatrix(I, Col_EgineNo) = Rst!EngineNo
            .TextMatrix(I, Col_InvNo) = Rst!Inv_Prefix & Replace(DeCodeDocID(Rst!Sal_Docid, Document_No), " ", "")
            .TextMatrix(I, Col_InvDate) = Rst!Inv_Date
            .TextMatrix(I, Col_FinAmt) = IIf(IsNull(Rst!Fin_Amt) Or Rst!Fin_Amt <= 0, "", Format(Rst!Fin_Amt, "0.00"))
            .TextMatrix(I, Col_TelcoShare) = IIf(mTelcoShare <= 0, "", Format(mTelcoShare, "0.00"))
            .TextMatrix(I, Col_Financier) = IIf(IsNull(Rst!FinName), "", Rst!FinName)
            .TextMatrix(I, Col_Party) = IIf(IsNull(Rst!Name), "", Rst!Name)
            .TextMatrix(I, Col_SalDocID) = IIf(IsNull(Rst!Sal_Docid), "", Rst!Sal_Docid)
        End With
        Rst.MoveNext
        I = I + 1
    Loop
    FGrid.Row = 1
    lblParty.CAPTION = "Customer : " & FGrid.TextMatrix(FGrid.Row, Col_Party)
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
        Txt(TelcoShare) = ""
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
        If IsValid(Txt(TelcoShare), "Telco Share %") = False Then Exit Sub
        'eof date checking

        FillData
        Disp_Text True
        Txt(AddressTo).SetFocus
End Select
End Sub

Private Sub FGrid_RowColChange()
    lblParty.CAPTION = "Customer : " & FGrid.TextMatrix(FGrid.Row, Col_Party)
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
    mDocType = PubDivCode & "S"
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    Set Master = GCn.Execute("Select SrlNo As SearchCode, Veh_OfftakeIncentive.*  from Veh_OfftakeIncentive Where " & cMID("Veh_OfftakeIncentive.srlNo", "1", "2") & "='" & mDocType & "'  Order by PurMonthDate,SrlNo")
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
    Txt(Srlno) = PubDivCode & "S"
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
    Txt(TelcoShare).Enabled = False
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
            
            GCn.Execute ("Update Veh_Stock Set OfftakeIncentiveSrlNo='',SubventionSrlNo='', OfftakeIncentive=0,TgtLinkIncentive=0 " & _
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
mRepName = "SubventionRep"

        mQRY = "Select '" & Txt(Srlno) & "' as SrlNo,VStk.Model,VStk.ChassisNo,VStk.EngineNo,VO.Inv_Prefix,VStk.Sal_DocId,VO.Inv_Date,VO.FIN_AMT,VStk.MfgShare,CF.FinName,SG.Name, " & _
        " '" & Txt(AddressTo) & "' as AddressTo, '" & Txt(Header) & "' as HeaderText,'" & Txt(Footer) & "' as FooterText " & _
        " From ((Veh_Stock VStk left join Veh_Order VO on VStk.Sal_DocId=VO.Inv_DocId) " & _
        " left join  ContractFinance CF on VO.FB_CODE=CF.FinCode) " & _
        " left Join SubGroup SG on VO.PartyCode=SG.SubCode " & _
        " Where VStk.SubventionSrlNo='" & Txt(Srlno) & "' Order By VStk.Model,VStk.ChassisNo"
        
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
            .Fields.Append "Inv_No", adChar, 15, adFldIsNullable
            .Fields.Append "Inv_Date", adDate, 7, adFldIsNullable
            .Fields.Append "Fin_Amt", adDouble, 12, adFldIsNullable
            .Fields.Append "MfgShare", adDouble, 12, adFldIsNullable
            .Fields.Append "FinName", adChar, 40, adFldIsNullable
            .Fields.Append "Name", adChar, 40, adFldIsNullable
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
                .Fields("Inv_No") = Rst!Inv_Prefix & Replace(DeCodeDocID(Rst!Sal_Docid, Document_No), " ", "")
                .Fields("Inv_Date") = Rst!Inv_Date
                .Fields("Fin_Amt") = Rst!Fin_Amt
                .Fields("MfgShare") = Rst!MfgShare
                .Fields("FinName") = Rst!FinName
                .Fields("Name") = Rst!Name
                .Fields("AddressTo") = Rst!AddressTo
                .Fields("HeaderText") = Rst!HeaderText
                .Fields("FooterText") = Rst!FooterText
                .Update
            End With
            Rst.MoveNext
        Loop
        
        CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".ttx", True
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
        rpt.Database.SetDataSource RstRep
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
        If FGrid.TextMatrix(I, Col_TelcoShare) <> "" Then
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Telco Share details", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_TelcoShare: FGrid.SetFocus: Exit Sub
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
                If Val(FGrid.TextMatrix(I, Col_TelcoShare)) > 0 Then
                    GSQL = "update Veh_Stock set SubventionSrlNo='" & Txt(Srlno) & "', MfgShare=" & Val(FGrid.TextMatrix(I, Col_TelcoShare)) & _
                        " where Veh_Stock.Sal_DocId='" & FGrid.TextMatrix(I, Col_SalDocID) & "' and ChassisNo='" & FGrid.TextMatrix(I, Col_ChassisNo) & "'"
                Else
                    GSQL = "update Veh_Stock set SubventionSrlNo='" & Txt(Srlno) & "', MfgShare=0 " & _
                        " where SubventionSrlNo='" & Txt(Srlno) & "' and ChassisNo='" & FGrid.TextMatrix(I, Col_ChassisNo) & "'"
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
    Case Header
        FGrid.Col = Col_TelcoShare
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
        Case Col_TelcoShare
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Financier
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
    Case Col_TelcoShare
        NumPress TxtGrid(Index), keyascii, 6, 2
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_TelcoShare
        If TxtGrid(Index) <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(TxtGrid(Index).TEXT, "0.00")
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
    Case Col_TelcoShare
        If TxtGrid(Index) <> "" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(TxtGrid(Index).TEXT, "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
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
            Case Col_TelcoShare
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
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
    Case Col_TelcoShare
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
        mInc = Val(FGrid.TextMatrix(I, Col_FinAmt))
        mTgtInc = Val(FGrid.TextMatrix(I, Col_TelcoShare))
    Next
    lblInc.CAPTION = "Total Finance Amount Rs.:" & IIf(mInc <= 0, "", Format(mInc, "0.00"))
    lblTotInc.CAPTION = "Total Telco Share Rs.:" & IIf(mTgtInc <= 0, "", Format(mTgtInc, "0.00"))
    lblInc.Refresh
    lblTotInc.Refresh
End Sub
