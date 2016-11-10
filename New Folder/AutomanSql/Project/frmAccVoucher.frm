VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmAccVoucher 
   BackColor       =   &H00D7C6C8&
   Caption         =   "Account Voucher Entry"
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
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Index           =   8
      Left            =   7410
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   600
      Width           =   3930
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   1845
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6390
      Visible         =   0   'False
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   3254
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Account Name"
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
         DataField       =   "Code"
         Caption         =   "Code"
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
      BeginProperty Column02 
         DataField       =   "Nature"
         Caption         =   "Nature"
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
      BeginProperty Column03 
         DataField       =   "curbal"
         Caption         =   "                Balance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
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
            ColumnWidth     =   4245.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1695.118
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   8775
      TabIndex        =   21
      Top             =   1215
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   90
         TabIndex        =   22
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
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4620
      MaxLength       =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6045
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1140
      Width           =   270
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   4080
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1410
      Width           =   2040
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   2
      Left            =   5025
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox TxtGrid1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
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
      Height          =   270
      Index           =   0
      Left            =   750
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2625
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txt 
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
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   870
      Width           =   2040
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3450
      MaxLength       =   15
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6045
      Width           =   1110
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   4275
      Left            =   15
      TabIndex        =   6
      Top             =   1740
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   7541
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16384
      Cols            =   6
      BackColorFixed  =   12632319
      ForeColorFixed  =   128
      BackColorSel    =   16777215
      ForeColorSel    =   16711680
      BackColorBkg    =   13623520
      GridColor       =   12632319
      GridColorFixed  =   8421504
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   7
      Left            =   6420
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   6045
      Width           =   5115
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   3
      Left            =   2790
      TabIndex        =   17
      Top             =   1425
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   2790
      TabIndex        =   16
      Top             =   1155
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   2790
      TabIndex        =   12
      Top             =   885
      Width           =   1170
   End
   Begin VB.Label lblPrefix 
      BackStyle       =   0  'Transparent
      Caption         =   "VPREFI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4350
      TabIndex        =   19
      Top             =   1155
      Width           =   720
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Doc Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   8130
      TabIndex        =   18
      Top             =   795
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   2250
      TabIndex        =   15
      Top             =   6060
      Width           =   1125
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   7215
      TabIndex        =   14
      Top             =   525
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblDocCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DocID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   7215
      TabIndex        =   13
      Top             =   795
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   660
      Left            =   7035
      Top             =   450
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   9585
      TabIndex        =   11
      Top             =   525
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblday 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6420
      TabIndex        =   10
      Top             =   1425
      Width           =   870
   End
End
Attribute VB_Name = "frmAccVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim ExitCtrl As Boolean
Dim GridKey As Integer
Dim TempBal As Double
Dim mVType As String
Dim ListArray As Variant
Dim mListItem As ListItem
Dim OldTrnType As String
Dim LVHeight As Integer
Private Const LV_VType As Byte = 1
Private Const LVCommnNarr As Byte = 3
Private Const LVSeparateNarr As Byte = 4

Private Const CellBackColLeave As String = &HD6EBE9     '&HEDF7FE
'Private Const CellForeColLeave As String = &HFF00FF
'Private Const CellBackColEnter As String = &HF0D5BF    '&HFFC0C0
'Private Const GridBackColorBkg As String = &HFFC0C0
Private Const GridBackColorBkg As String = &HD7C6C8
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$


Dim VoucherEditFlag As Boolean
Dim ForSiteCode As String
Dim SeparateNarr As Boolean
Dim CommanNarr As Boolean
Dim MyIndex As Byte
Dim Rst As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim RsParty As ADODB.Recordset
'Dim RsVoucher As ADODB.Recordset

Private Const VDesc As Byte = 1
Private Const VNo As Byte = 2
Private Const VAdd As Byte = 3
Private Const Vdate As Byte = 4
Private Const TotDr As Byte = 5
Private Const TotCr As Byte = 6
Private Const Narr As Byte = 7
Private Const NarrDisp As Byte = 8

'Fgrid1 Columns
Private Const C_AcName As Byte = 1
Private Const C_Debit As Byte = 2
Private Const C_Credit As Byte = 3
Private Const C_Narration As Byte = 4
Private Const C_ChqNo As Byte = 5
Private Const C_ChqDt As Byte = 6
Private Const C_ClgDt As Byte = 7
Private Const C_CurrBal As Byte = 8
Private Const C_AcCode As Byte = 9
Private Const C_Narr1 As Byte = 10
Private Const C_MainGrCode As Byte = 11
Private Const C_Nature As Byte = 12

Private Sub DGParty_Click()
On Error GoTo errorbox
If RsParty.RecordCount > 0 Then
    TxtGrid1(0).Tag = RsParty!Code
    TxtGrid1(0).Text = RsParty!Name
End If
TxtGrid1(0).SetFocus
DGParty.Visible = False
Exit Sub
errorbox:
    MsgBox err.Description, vbInformation
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
    '******************** Colour Setting
    Dim i As Integer
    Me.BackColor = &HBAD3C9
    TopCtrl1.Tag = PubUParam:   WinSetting Me:    Ini_Grid
    ForSiteCode = PubSiteCode
    For i = 1 To txt.Count - 1
        txt(i).BackColor = CtrlBColOrg
        txt(i).ForeColor = CtrlFColOrg
    Next
    '    Remark   :  Ini_Grid Procrdure :  .ColWidth(C_ClgDt) = 0
    '                                      .ColWidth(C_CurrBal) = 1125
    '                                      .ColWidth(C_AcName) = 2750
    '                                      .ColWidth(0) = 300
    '************************************
    Call BlankText
    lblPrefix.Caption = ""
    
    '' Pending Points
    ''  1. Current A/c balance - done lps
    ''  2. Adjustment Entry
    Set GRs = New Recordset
    Set GRs = GCnFa.Execute("Select V_Type as Code, Description as Name, Number_Method,Common_Narr,Separate_Narr from Voucher_Type where Category='FA' and TRIM(V_TYPE)<>'F_AO' order by description")
    Set mListItem = ListView_Items_RecordSet(ListView, txt, VDesc, GRs)
    Set GRs = Nothing
    ListView.ColumnHeaders(1).width = txt(VDesc).width - 90
    ListView.ColumnHeaders(2).width = 0
    If ListView.ListItems.Count > 10 Then
        LVHeight = 10 * 300
    Else
        LVHeight = ListView.ListItems.Count * 300
    End If
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select LM.docId as searchcode,lm.docid,LM.Site_Code,LM.V_Type,LM.V_Prefix,LM.V_No,LM.V_Date,LM.NARRATION,Common_Narr,Separate_Narr,VT.DESCRIPTION " & _
        "from LEDGERM as LM " & _
        "left Join Voucher_Type as VT on LM.V_TYPE=VT.V_TYPE " & _
        "where VT.Category='FA' and TRIM(VT.V_TYPE)<>'F_AO' " & _
        "Order by LM.V_DATE desc,LM.V_TYPE,LM.V_NO desc", GCnFa, adOpenDynamic, adLockOptimistic

    Set RsParty = New ADODB.Recordset
    With RsParty
        .CursorLocation = adUseClient
        .Open "Select subcode as Code, Name, subgroup.Nature, iif(curr_bal<>0, format(abs(curr_bal),'##,##,##0.00'),'')+' '+iif(curr_bal<0,'Cr',iif(curr_bal>0,'Dr','  ')) as curbal,curr_bal,Acgroup.maingrcode " & _
            "from subgroup left join acgroup on subgroup.groupcode=acgroup.groupcode " & _
            "where acgroup.aliasyn='N' order by Name", GCnFa, adOpenDynamic, adLockOptimistic
    End With
    Set DGParty.DataSource = RsParty
    
'    Set RsVoucher = New ADODB.Recordset
'    With RsVoucher
'        .CursorLocation = adUseClient
'        .Open "Select V_Type as Code, Description as Name, Number_Method,Common_Narr,Separate_Narr from Voucher_Type where Category='FA' and TRIM(V_TYPE)<>'F_AO' order by description", GCnFa, adOpenDynamic, adLockOptimistic
'    End With
'    Set DGVoucher.DataSource = RsVoucher
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
    
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsParty = Nothing
    Set mListItem = Nothing
End Sub

Private Sub ListView_Click()
    txt(Val(ListView.Tag)).Text = ListView.SelectedItem.Text
    txt(Val(ListView.Tag)).Tag = ListView.SelectedItem.SubItems(LV_VType)
    txt(Val(ListView.Tag)).SelStart = Len(txt(Val(ListView.Tag)).Text)
    ColScheme txt(Val(ListView.Tag)).Tag
    txt(Val(ListView.Tag)).SetFocus
    FrmList.Visible = False
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim i As Integer
    Disp_Text SETS("ADD", Me, Master)
    
    'Call BlankText
    
    txt(VNo).Text = ""
    txt(VNo).Tag = ""
    txt(VAdd).Text = ""
    txt(VAdd).Tag = ""
    txt(TotCr).Text = ""
    txt(TotCr).Tag = ""
    txt(TotDr).Text = ""
    txt(TotDr).Tag = ""
    txt(Narr).Text = ""
    txt(Narr).Tag = ""
    txt(NarrDisp).Text = ""
    
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    
    FGrid1.Col = 1
    txt(Vdate).Text = Format(IIf(txt(Vdate).Tag = "", date, txt(Vdate).Tag), "dd/MMM/yyyy")
    lblday = WeekdayName(Weekday(txt(Vdate).Text))
    
    If txt(VDesc).Text <> "" Then
        Set mListItem = ListView.FindItem(txt(VDesc).Text, 0, , 1)
        txt(VDesc).Tag = mListItem.SubItems(LV_VType)
        CommanNarr = IIf(mListItem.SubItems(LVCommnNarr) = "Y", True, False)
        SeparateNarr = IIf(mListItem.SubItems(LVSeparateNarr) = "Y", True, False)
        Set mListItem = Nothing
    Else
        txt(VDesc).Tag = ""
        txt(VDesc).Text = ""
        CommanNarr = True
        SeparateNarr = True
    End If
    
    txt(Narr).Enabled = CommanNarr
    txt(Narr).Locked = Not CommanNarr
    If Val(txt(VDesc).Tag) > 0 Then
        If GCnFa.Execute("Select VT.Number_Method From Voucher_Type VT  Where VT.V_Type='" & txt(VDesc).Tag & "'").Fields(0).Value = "Manual" Then
            txt(VNo).Text = GCnFa.Execute("select iif(isnull(max(v_no)),0,max(v_no))+1 from LedgerM where left(docid,1)='" & PubDivCode & "' and mid(docid,2,2)='" & PubSiteCode + ForSiteCode & "' and v_type='" & txt(VDesc).Tag & "'").Fields(0)
        End If
        lblDocId = AccGetDocID(txt(VDesc).Tag, txt(Vdate).Text, VoucherEditFlag, txt(VNo), lblPrefix, ForSiteCode)
    End If
    
    LblDiv.Caption = "Division : " & PubDivCode
    LblSite.Caption = "Site Code : " & PubSiteCode
    lblPrefix.Caption = ""
    txt(VDesc).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCnFa.BeginTrans
            GCn.BeginTrans
            Call Remove_Curbal
            GCnFa.Execute "Delete from ledgerm where Docid='" & lblDocId & "'"
            GCnFa.Execute "Delete from Ledger  where Docid='" & lblDocId & "'"
            GCn.CommitTrans
            GCnFa.CommitTrans
            
            Master.Requery
            Call UpdRequery
            
            If Master.RecordCount > 0 Then
                If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
                MoveRec
            Else
                Call BlankText
            End If
            BUTTONS True, Me, Master, 0
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
    Exit Sub
eloop1:
    GCnFa.RollbackTrans: GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message"
End Sub

Private Sub TopCtrl1_eEdit()
Dim i As Integer
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    
    txt(VDesc).Enabled = False
    txt(VNo).Enabled = False
    txt(VAdd).Enabled = False
    
    txt(Narr).Enabled = CommanNarr
    txt(Narr).Locked = Not CommanNarr
    
    
    txt(Vdate).SetFocus
    
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
'    If TopCtrl1.TopText2 = "Browse" Then Unload Me
    Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
    GSQL = "select LM.docId as searchcode,lm.docid,LM.V_Type,LM.V_No,LM.V_Date " & _
        "from LEDGERM as LM " & _
        "left Join Voucher_Type as VT on LM.V_TYPE=VT.V_TYPE " & _
        "where VT.Category='FA' and TRIM(VT.V_TYPE)<>'F_AO' " & _
        "Order by LM.V_DATE desc,LM.V_TYPE,LM.V_NO desc"
    Set SearchForm = Me
    FIND2.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='C11V_DCLVDCHL       1'") '" & MyValue & "'")
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
Dim i As Integer
On Error GoTo ErrorLoop
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        txt(Vdate).Tag = ""
        Call MoveRec
    Else
        Me.ActiveControl.SetFocus
    End If
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
On Error GoTo errlbl
    Dim mDrAmt As Double, mCrAmt As Double
    Dim i As Integer, SrNo As Integer, YN As Integer
    Dim mTrans As Boolean
    Dim AddFlg$, MyPrefix$, MyContra$, MyDrCode$, MyCrCode$
    Dim MyDCount As Integer, MyCCount As Integer
    
    mDrAmt = 0#:   mCrAmt = 0#
    i = 0: SrNo = 0: MyDCount = 0: MyCCount = 0: YN = 0
    
    If TxtGrid1(0).Visible = True Then
        If TxtGrid1Leave = False Then
            TxtGrid1_LostFocus 0
            TxtGrid1(0).SetFocus
            Exit Sub
        Else
            TxtGrid1(0).Visible = False
        End If
    End If

    Grid_Hide
    
    If IsValid(txt(VDesc), "Voucher Type") = False Then Exit Sub
    If IsValid(txt(VNo), "Voucher No.") = False Then Exit Sub
    If IsValid(txt(Vdate), "Voucher Date") = False Then Exit Sub
    
    If Val(txt(VNo).Text) = 0 Then
        MsgBox "Invalid Voucher No. ", vbInformation, "Validation"
        If txt(VNo).Enabled = True Then
            txt(VNo).SetFocus
        End If
        Exit Sub
    End If
    
'    If CheckFinYear(txt(vdate).Text) = False Then Exit Sub
    If Trim(lblDocId) = "" Then MsgBox "Invalid DocId", vbInformation, "Validation": Exit Sub
    If Val(txt(TotCr)) = 0 Then MsgBox "Credit value is Zero", vbInformation, "Validation": FGrid1.SetFocus: Exit Sub
    If Val(txt(TotDr)) = 0 Then MsgBox "Debit value is Zero", vbInformation, "Validation": FGrid1.SetFocus: Exit Sub
    
    '' checking of any row in fgrid without amount
    MyDCount = 0: MyCCount = 0: MyDrCode = "": MyCrCode = ""
    For i = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(i, C_AcCode) <> "" Then
            If Val(FGrid1.TextMatrix(i, C_Debit)) + Val(FGrid1.TextMatrix(i, C_Credit)) = 0 Then
                MsgBox "Please fill Amount in Row No. " & i, vbInformation, "Validation"
                FGrid1.Row = i: FGrid1.Col = C_Debit
                FGrid1.SetFocus
                Exit Sub
            End If
            If PubChqNoReq And UCase(FGrid1.TextMatrix(i, C_Nature)) = "BANK" Then
                If FGrid1.TextMatrix(i, C_ChqNo) = "" Then
                    If CheqNoReq Then
                        YN = MsgBox("Please Enter Chq/DD No.!" & vbCrLf & "Do you want to fill 'Chq/DD No.'?", vbYesNo, "Chq/DD No.Check")
                        If YN = vbYes Then
                            FGrid1.Row = i: FGrid1.Col = C_ChqNo
                            FGrid1.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Else
            If Val(FGrid1.TextMatrix(i, C_Debit)) + Val(FGrid1.TextMatrix(i, C_Credit)) <> 0 Then
                MsgBox "Please Fill A/c in Row No. " & i, vbInformation, "Validation"
                FGrid1.Row = i: FGrid1.Col = C_AcName
                FGrid1.SetFocus
                Exit Sub
            End If
        End If
        If Val(FGrid1.TextMatrix(i, C_Debit)) > 0 Then
            MyDrCode = FGrid1.TextMatrix(i, C_AcCode)
            MyDCount = MyDCount + 1
        ElseIf Val(FGrid1.TextMatrix(i, C_Credit)) > 0 Then
            MyCrCode = FGrid1.TextMatrix(i, C_AcCode)
            MyCCount = MyCCount + 1
        End If
    Next i
    
    If MyCCount = 1 And MyDCount > 1 Then
        MyContra = MyCrCode
    End If
    If MyDCount = 1 And MyCCount > 1 Then
        MyContra = MyDrCode
    End If
    'Dr / Cr checking
    For i = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(i, C_AcName) <> "" Then
            mDrAmt = mDrAmt + Val(FGrid1.TextMatrix(i, C_Debit))
            mCrAmt = mCrAmt + Val(FGrid1.TextMatrix(i, C_Credit))
        End If
    Next
    If Not Val(mDrAmt) = Val(mCrAmt) Then
        MsgBox "Total Debit/Credit Amount mismatch", vbInformation, "Validation"
        FGrid1.SetFocus
        Exit Sub
    End If
    '' checking for data in fgrid1
    If mDrAmt + mCrAmt = 0 Then MsgBox "No A/c Details Feeded ", vbInformation: Exit Sub
Mynxt:
    '' eof : checking of data in fgrid1
mDrAmt = 999999.99

    GCn.BeginTrans
    GCnFa.BeginTrans
    
    mTrans = True
    MyPrefix = Space(1 - Len(Trim(txt(VAdd).Text))) + lblPrefix
    
    Select Case TopCtrl1.TopText2
        Case "Add"
            AddFlg = "A"
            GSQL = "Select DocID From LEDGERM Where DocID='" & lblDocId & "'"
            If VoucherEditFlag Then  'Manual Numbering
                If GCnFa.Execute(GSQL).RecordCount > 0 Then
                    MsgBox "Voucher No. " & txt(VNo).Text & " Already Exists", vbCritical, "Validation Error"
                    txt(VNo).SetFocus
                    GoTo errlbl
                End If
            Else    'Automatic Numbering
                If GCnFa.Execute(GSQL).RecordCount > 0 Then
                    MsgBox "Voucher No. " & txt(VNo).Text & " Already Exists", vbCritical, "Validation Error"
                End If
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & txt(VDesc).Tag & "' And VP.Date_From<=#" & Format(txt(Vdate).Text, "dd/MMM/yyyy") & "# Order By VP.Date_From DESC", GCnFa, adOpenDynamic, adLockOptimistic
                If Val(Rst!Start_Srl_No) >= Val(txt(VNo).Text) Then
                    lblDocId = AccGetDocID(txt(VDesc).Tag, txt(Vdate).Text, VoucherEditFlag, txt(VNo), lblPrefix, ForSiteCode)
                End If
                If Rst.RecordCount > 0 Then
                    GSQL = "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_Type & "' and Date_From=#" & Format(Rst!Date_From, "dd/MMM/yyyy") & "#"
                    GCnFa.Execute GSQL
                End If
            End If
            GSQL = "insert into ledgerm(" _
                & "DocId ,Site_Code,V_Type,v_prefix,V_No," _
                & "V_Date,Narration,U_Name, U_EntDt, U_AE) " _
                & " values('" & lblDocId & "','" & PubSiteCode & "','" & txt(VDesc).Tag & "','" & MyPrefix & "'," & txt(VNo).Text & "," _
                & "" & ConvertDate(txt(Vdate).Text) & ",'" & IIf(CommanNarr, txt(Narr).Text, "") & "','" & pubUName & "',#" & PubServerDate & "#,'" & AddFlg & "')"
            GCnFa.Execute GSQL
        Case "Edit"
            AddFlg = "E"
            Call Remove_Curbal
            GCnFa.Execute "Delete from Ledger where Docid='" & lblDocId & "'"
            GSQL = "Update ledgerm set V_Date=" & ConvertDate(txt(Vdate).Text) & ",Narration='" & txt(Narr).Text & "',U_Name='" & pubUName & "', U_EntDt=#" & PubServerDate & "#, U_AE='" & AddFlg & "' where docid='" & lblDocId & "'"
            GCnFa.Execute GSQL
    End Select
    SrNo = 1
    For i = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(i, C_AcCode) <> "" Then
            If MyCCount + MyDCount = 2 Then    '' in case Only One Credit Entry & One Debit Entry
                If i = 1 Then MyContra = FGrid1.TextMatrix(2, C_AcCode)
                If i = 2 Then MyContra = FGrid1.TextMatrix(1, C_AcCode)
            End If

            GSQL = "insert into ledger(" _
                & "DocId,Site_Code,v_sNo,V_type,v_no," _
                & "v_date,subcode,contrasub,amtcr,amtdr,Chq_no," _
                & "chq_date,clg_Date,narration,U_Name, U_EntDt, U_AE)" _
                & " values(" _
                & "'" & lblDocId & "','" & PubSiteCode & "'," & SrNo & ",'" & txt(VDesc).Tag & "'," & txt(VNo).Text & "," _
                & "" & ConvertDate(txt(Vdate).Text) & ",'" & FGrid1.TextMatrix(i, C_AcCode) & "','" & MyContra & "'," & Val(FGrid1.TextMatrix(i, C_Credit)) & "," & Val(FGrid1.TextMatrix(i, C_Debit)) & ",'" & FGrid1.TextMatrix(i, C_ChqNo) & "'," _
                & "" & ConvertDate(FGrid1.TextMatrix(i, C_ChqDt)) & "," & ConvertDate(FGrid1.TextMatrix(i, C_ClgDt)) & ",'" & IIf(SeparateNarr, FGrid1.TextMatrix(i, C_Narr1), txt(Narr)) & "','" & pubUName & "',#" & PubServerDate & "#,'" & AddFlg & "')"
            GCnFa.Execute GSQL
            
            If Val(FGrid1.TextMatrix(i, C_Credit)) > 0 Then
                CalBalAcGroup "SubGroup", GCnFa, FGrid1.TextMatrix(i, C_MainGrCode), Val(FGrid1.TextMatrix(i, C_Credit)), "-"
                GCnFa.Execute ("Update Subgroup set curr_bal=Curr_bal-" & Val(FGrid1.TextMatrix(i, C_Credit)) & " where subcode='" & FGrid1.TextMatrix(i, C_AcCode) & "'")
                GCn.Execute ("Update Subgroup set curr_bal=Curr_bal-" & Val(FGrid1.TextMatrix(i, C_Credit)) & " where subcode='" & FGrid1.TextMatrix(i, C_AcCode) & "'")
            Else
                CalBalAcGroup "SubGroup", GCnFa, FGrid1.TextMatrix(i, C_MainGrCode), Val(FGrid1.TextMatrix(i, C_Debit)), "+"
                GCnFa.Execute ("Update Subgroup set curr_bal=Curr_bal+" & Val(FGrid1.TextMatrix(i, C_Debit)) & " where subcode='" & FGrid1.TextMatrix(i, C_AcCode) & "'")
                GCn.Execute ("Update Subgroup set curr_bal=Curr_bal+" & Val(FGrid1.TextMatrix(i, C_Debit)) & " where subcode='" & FGrid1.TextMatrix(i, C_AcCode) & "'")
            End If
            SrNo = SrNo + 1
        End If
    Next i
    GCnFa.CommitTrans
    GCn.CommitTrans
    mTrans = False
    
    Master.Requery
    Call UpdRequery
    Master.FIND "searchcode = '" & lblDocId & "'"
    If TopCtrl1.TopText2.Caption = "Add" Then txt(Vdate).Tag = txt(Vdate).Text: TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then GCnFa.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus txt(Index)
    TxtGrid1(0).Visible = False
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case VDesc
            OldTrnType = txt(VDesc).Text
            If txt(VDesc).Text <> "" Then
                Set mListItem = ListView.FindItem(txt(VDesc).Text, 0, , 1)
                If mListItem Is Nothing Then
                Else
                    mListItem.SELECTED = True
                End If
            End If
            mVType = txt(VDesc).Tag
        Case Narr 'Comman Narration
            Dim i As Integer
            For i = 1 To FGrid1.Rows - 1
                If FGrid1.TextMatrix(i, C_AcCode) <> "" Then
                    If PubChqNoReq And UCase(FGrid1.TextMatrix(i, C_Nature)) = "BANK" Then
                        If txt(Narr) = "" Then
                            txt(Narr) = "Chq.No.:"
                            txt(Narr).SelStart = Len(txt(Narr))
                            Exit For
                        End If
                    End If
                End If
            Next
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case Vdate
            If KeyCode = vbKeyDelete Then
                txt(Vdate).Text = ""
            End If
    End Select
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case VDesc
           ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).Height), txt(Index).width, LVHeight
    End Select
    
    
    If FrmList.Visible = False Then
        '' KEY DOWN and Enter Key
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Narr Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyTab) And Index = Narr Then
            If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If (txt(VDesc).Enabled = False And Index <> Vdate) Or (txt(VDesc).Enabled = True And Index <> VDesc) Then
                If Index = Narr Then Exit Sub
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index = Narr Then Exit Sub
            If Index <> Vdate Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
        Case VNo
            Call NumPress(txt(Index), KeyAscii, 8, 0)
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case VDesc
            If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case VDesc
            If txt(Index).Text <> "" Then txt(Index).Text = ListView.SelectedItem.Text
            If txt(Index).Text = "" Then
                MsgBox "No Voucher Type selected"
                txt(VNo).Text = ""
                txt(VAdd).Text = ""
                txt(VNo).Enabled = True
                txt(VAdd).Enabled = True
                CommanNarr = True
                SeparateNarr = True
            Else
                txt(VDesc).Tag = ListView.SelectedItem.SubItems(LV_VType)
                If GCnFa.Execute("Select VT.Number_Method From Voucher_Type VT  Where VT.V_Type='" & txt(VDesc).Tag & "'").Fields(0).Value = "Manual" Then
                    txt(VNo).Text = GCnFa.Execute("select iif(isnull(max(v_no)),0,max(v_no))+1 from LedgerM where left(docid,1)='" & PubDivCode & "' and mid(docid,2,2)='" & PubSiteCode + ForSiteCode & "' and v_type='" & txt(VDesc).Tag & "'").Fields(0)
                End If
                lblDocId = AccGetDocID(txt(VDesc).Tag, txt(Vdate).Text, VoucherEditFlag, txt(VNo), lblPrefix, ForSiteCode)
                CommanNarr = IIf(ListView.SelectedItem.SubItems(LVCommnNarr) = "Y", True, False)
                SeparateNarr = IIf(ListView.SelectedItem.SubItems(LVSeparateNarr) = "Y", True, False)
                
                If mVType <> txt(VDesc).Tag Then
                    mVType = txt(VDesc).Tag
                    ColScheme (txt(VDesc).Tag)
                End If
            End If
            txt(Narr).Enabled = CommanNarr
        Case VNo, VAdd
            If Val(txt(VNo).Text) = 0 Then MsgBox "Zero Voucher No. ", vbInformation, "Validation": Exit Sub
            If txt(VDesc).Text = "" Then Exit Sub
            txt(Index).Text = UCase(txt(Index).Text)
            lblDocId = AccGetDocID(txt(VDesc).Tag, txt(Vdate).Text, VoucherEditFlag, txt(VNo), lblPrefix, ForSiteCode)
            
            If VoucherEditFlag = True Then    ' Manual
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select Docid From ledgerm Where DocID='" & lblDocId & "'", GCnFa, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Voucher No. Not Allowed", vbInformation, "Validation"
                    If txt(Index).Enabled = True Then txt(Index).SetFocus
                End If
            End If
        Case Vdate
            If txt(Index) = "" Then
                txt(Index) = PubLoginDate
            End If
            txt(Index).Text = RetDate(txt(Index))
            lblday = WeekdayName(Weekday(txt(Index).Text))
            If CheckFinYear(CDate(txt(Index).Text)) = False Then txt(Vdate).SetFocus: Cancel = True
            If TopCtrl1.TopText2 = "Add" Then
                lblDocId = AccGetDocID(txt(VDesc).Tag, txt(Vdate).Text, VoucherEditFlag, txt(VNo), lblPrefix, ForSiteCode)
            End If
    End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
    For i = 1 To txt.Count
        txt(i).Text = ""
        txt(i).Tag = ""
    Next i
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    
    lblDocId.Caption = ""
    lblDocId.Refresh
    lblday.Caption = ""
    lblday.Refresh
End Sub

Private Sub MoveRec()
On Error GoTo error1
    If Master.RecordCount > 0 Then
        LblDiv.Caption = "Division : " & DeCodeDocID(Master!DocId, Division_Code)
        LblSite.Caption = "Site Code : " & DeCodeDocID(Master!DocId, Current_Site)
        lblDocId.Caption = Master!DocId
'F_AR   Receipts
'F_BP   Payments
'F_CRN  Credit Note
'F_DRN  Debit Note
'F_JV   Journal Voucher
        If txt(VDesc).Tag <> Master!V_Type Then
            ColScheme (Master!V_Type)
        End If
        txt(VDesc).Tag = XNull(Master!V_Type)
        txt(VDesc).Text = XNull(Master!Description)
        lblPrefix.Caption = Mid(Master!DocId, 10, 4)
        txt(VNo).Text = Master!V_No
        txt(VAdd).Text = Mid(Master!DocId, 9, 1)
        txt(Vdate).Text = Format(Master!V_DATE, "dd/MMM/yyyy")
        lblday = WeekdayName(Weekday(txt(Vdate).Text))
        CommanNarr = IIf(Master!Common_Narr = "Y", True, False)
        SeparateNarr = IIf(Master!Separate_Narr = "Y", True, False)
        
        txt(Narr).Text = XNull(Master!Narration)
        txt(Narr).Locked = True
        txt(Narr).Enabled = True
        Call Fill_Grid
    Else
        Call BlankText
    End If
    TempBal = 0
    Grid_Hide
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    With FGrid1
        .left = Me.left ' +45
        .width = Me.width - 120
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 13
    '                                      .ColWidth(C_ClgDt) = 0
    '                                      .ColWidth(C_CurrBal) = 1125
    '                                      .ColWidth(C_AcName) = 2750
    '                                      .ColWidth(0) = 300
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 300
        
        .TextMatrix(0, C_AcName) = "Account Description"
        .ColAlignment(C_AcName) = flexAlignLeftCenter
        .ColAlignmentFixed(C_AcName) = flexAlignLeftCenter
        .ColWidth(C_AcName) = 3150
        
        .TextMatrix(0, C_Debit) = "Debit"
        .ColAlignment(C_Debit) = flexAlignRightCenter
        .ColAlignmentFixed(C_Debit) = flexAlignRightCenter
        .ColWidth(C_Debit) = 1125
        
        .TextMatrix(0, C_Credit) = "Credit"
        .ColAlignment(C_Credit) = flexAlignRightCenter
        .ColAlignmentFixed(C_Credit) = flexAlignRightCenter
        .ColWidth(C_Credit) = 1125
        
        .TextMatrix(0, C_Narration) = "Narration"
        .ColAlignment(C_Narration) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Narration) = flexAlignLeftCenter
        .ColWidth(C_Narration) = 2625
        
        .TextMatrix(0, C_ChqNo) = "Cheque No."
        .ColAlignment(C_ChqNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ChqNo) = flexAlignLeftCenter
        .ColWidth(C_ChqNo) = 1035
        
        .TextMatrix(0, C_ChqDt) = "Cheque Dt."
        .ColAlignment(C_ChqDt) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ChqDt) = flexAlignLeftCenter
        .ColWidth(C_ChqDt) = 1035
        
        .TextMatrix(0, C_ClgDt) = "Clearing Dt."
        .ColAlignment(C_ClgDt) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ClgDt) = flexAlignLeftCenter
        .ColWidth(C_ClgDt) = 0
        
        .TextMatrix(0, C_CurrBal) = "Curr.Bal."
        .ColAlignment(C_CurrBal) = flexAlignRightCenter
        .ColAlignmentFixed(C_CurrBal) = flexAlignRightCenter
        .ColWidth(C_CurrBal) = 1170
        
        .TextMatrix(0, C_AcCode) = "AcCode"
        .ColAlignment(C_AcCode) = flexAlignLeftCenter
        .ColAlignmentFixed(C_AcCode) = flexAlignLeftCenter
        .ColWidth(C_AcCode) = 0
        
        .TextMatrix(0, C_Narr1) = "tnarr"
        .ColAlignment(C_Narr1) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Narr1) = flexAlignLeftCenter
        .ColWidth(C_Narr1) = 0
    
        .TextMatrix(0, C_MainGrCode) = "maingrcode"
        .ColAlignment(C_MainGrCode) = flexAlignLeftCenter
        .ColAlignmentFixed(C_MainGrCode) = flexAlignLeftCenter
        .ColWidth(C_MainGrCode) = 0
        .ColWidth(C_Nature) = 0
    End With
    BackColorSelLeave = FGrid1.BackColorSel
    ForeColorSelEnter = FGrid1.ForeColorSel
    With DGParty
        .width = 6750
        .left = Me.width - (DGParty.width + mRtScale)
        .top = mTopScale
        .Height = Me.Height - (mTopScale + mBotScale) ' 8135
        .Columns(3).width = 1860.095
    End With
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim i As Integer
    For i = 1 To txt.Count
        txt(i).Enabled = Enb
        txt(i).BackColor = CtrlBColOrg
        txt(i).ForeColor = CtrlFColOrg
    Next
    
    TxtGrid1(0).BackColor = CtrlBColOrg
    TxtGrid1(0).ForeColor = CtrlFColOrg
    TxtGrid1(0).Enabled = Enb
    txt(Narr).Enabled = True
    txt(Narr).Locked = True
    txt(NarrDisp).Enabled = False
    
    txt(TotCr).Enabled = True
    txt(TotDr).Enabled = True
    txt(TotCr).Locked = True
    txt(TotDr).Locked = True
End Sub

Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub UpdRequery()
    RsParty.Requery
End Sub

Private Sub FGrid1_Click()
    TxtGrid1(0).Visible = False
End Sub

Private Sub FGrid1_DblClick()
On Error GoTo ELoop
    FGrid1_KeyPress (vbKeyReturn)
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = BackColorSelEnter
    FGrid1.ForeColorSel = ForeColorSelEnter
    TxtGrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
        If CommanNarr Then
            SendKeys vbTab
        Else
            If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        FGrid1.SetFocus
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid1.Col
            Case C_Debit, C_Credit, C_ChqDt, C_ChqNo    ', C_ClgDt
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            Case C_Narration
                If SeparateNarr = False Then Exit Sub
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Narr1) = ""
                txt(NarrDisp) = ""
        End Select
    End If


'    If KeyCode = vbKeyReturn Then
'        Select Case FGrid1.Col
'            Case C_AcName, C_Debit, C_Credit, C_ChqDt, C_ChqNo, C_ClgDt
'                GridDblClick Me, FGrid1, TxtGrid1, 0
'            Case C_Narration
'                If SeparateNarr = False Then Exit Sub
'                GridDblClick Me, FGrid1, TxtGrid1, 0
'                TxtGrid1(0).Text = FGrid1.TextMatrix(FGrid1.Row, C_Narr1)
'        End Select
'        TAddMode = False
'    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
'Dim mNature$
SetMaxLength
    'If KeyAscii = 115 Then FGrid1_LeaveCell: Exit Sub
    Select Case FGrid1.Col
        Case C_AcName
            Get_Text Me, FGrid1, TxtGrid1, 0, False, KeyAscii
        Case C_Debit, C_Credit
            If FGrid1.TextMatrix(FGrid1.Row, C_AcName) <> "" Then
                Get_Text Me, FGrid1, TxtGrid1, 0, True, KeyAscii
            End If
        Case C_Narration
            If FGrid1.TextMatrix(FGrid1.Row, C_AcName) <> "" Then
                If SeparateNarr = False Then Exit Sub
                If FGrid1.TextMatrix(FGrid1.Row, C_Narration) = "" Then
                    If UCase(FGrid1.TextMatrix(FGrid1.Row, C_Nature)) = "BANK" Then
                        FGrid1.TextMatrix(FGrid1.Row, C_Narration) = "Chq.No.:"
                        FGrid1.TextMatrix(FGrid1.Row, C_Narr1) = "Chq.No.:"
                    End If
                End If
                Get_Text Me, FGrid1, TxtGrid1, 0, False, KeyAscii
    '            GridDblClick Me, FGrid1, TxtGrid1, 0
    '            TxtGrid1(0).Text = FGrid1.TextMatrix(FGrid1.Row, C_Narr1)
            End If
        Case C_ChqDt, C_ChqNo   ', C_ClgDt
            If FGrid1.TextMatrix(FGrid1.Row, C_AcName) <> "" Then
                Get_Text Me, FGrid1, TxtGrid1, 0, False, KeyAscii
            End If
        Case C_CurrBal  'skip Row/col
            If FGrid1.Row <> FGrid1.Rows - 1 Then
                FGrid1.Row = FGrid1.Row + 1
            Else
                If FGrid1.TextMatrix(FGrid1.Row, C_AcName) <> "" Then
                    If FGrid1.Row = FGrid1.Rows - 1 Then FGrid1.AddItem FGrid1.Rows - 1
                End If
            End If
            FGrid1.Col = C_AcName
            FGrid1.SetFocus
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
    If FGrid1.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid1.Row >= 1 Then
            If MsgBox("Are You Sure To Delete Entry ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid1.Rows > 2 Then
                    FGrid1.RemoveItem (FGrid1.Row)
                Else
                    FGrid1.Rows = 1
                    FGrid1.AddItem FGrid1.Rows
                    FGrid1.FixedRows = 1
                End If
                Call Calc_GridAmt
            End If
            For i = 1 To FGrid1.Rows - 1
                FGrid1.TextMatrix(i, 0) = i
            Next
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid1.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub FGrid1_LostFocus()
    FGrid1.BackColorSel = BackColorSelLeave
    FGrid1.ForeColorSel = FGrid1.ForeColor
End Sub

Private Sub FGrid1_RowColChange()
txt(NarrDisp) = FGrid1.TextMatrix(FGrid1.Row, C_Narration)
FGrid1.ToolTipText = FGrid1.TextMatrix(FGrid1.Row, C_AcName) & "  Current Balance  Rs. " & FGrid1.TextMatrix(FGrid1.Row, C_CurrBal)
If TopCtrl1.TopText2.Caption = "Browse" Then Exit Sub
End Sub

Private Sub FGrid1_Scroll()
    TxtGrid1(0).Visible = False
    DGParty.Visible = False
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
On Error GoTo ELoop
'If ExitCtrl = False Then Exit Sub
    Ctrl_GetFocus TxtGrid1(0)
    Grid_Hide
    If FGrid1.Col = C_Narration Then
        TxtGrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, C_Narr1)
    Else
        TxtGrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    End If
    Select Case FGrid1.Col
        Case C_AcName
            If RsParty.RecordCount = 0 Or RsParty.EOF = True Or RsParty.BOF = True Or TxtGrid1(Index).Text = "" Then Exit Sub
            RsParty.MoveFirst
            RsParty.FIND "name='" & FGrid1.TextMatrix(FGrid1.Row, C_AcName) & "'"
        Case C_Debit
            If Val(FGrid1.TextMatrix(FGrid1.Row, C_Credit)) <= 0 And Val(FGrid1.TextMatrix(FGrid1.Row, C_Debit)) = 0 Then
                If TempBal < 0 Then TxtGrid1(0).Text = Abs(TempBal)     ' - Val(FGrid1.TextMatrix(FGrid1.Row, C_Credit)) + Val(FGrid1.TextMatrix(FGrid1.Row, C_Debit)))
                'SendKeys "{Home}+{End}"
            End If
        Case C_Credit
            If Val(FGrid1.TextMatrix(FGrid1.Row, C_Debit)) <= 0 And Val(FGrid1.TextMatrix(FGrid1.Row, C_Credit)) = 0 Then
                If TempBal > 0 Then TxtGrid1(0).Text = Abs(TempBal)     '' + Val(FGrid1.TextMatrix(FGrid1.Row, C_Debit)) - Val(FGrid1.TextMatrix(FGrid1.Row, C_Credit)))
                'SendKeys "{Home}+{End}"
            End If
        Case C_Narration    'Special Case
            TxtGrid1(0).Height = FGrid1.RowHeight(0) * 3
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    
    If KeyCode = vbKeyEscape Then
        TxtGrid1(0).Text = TxtGrid1(0).Tag
        TxtGrid1_KeyUp Index, KeyCode, Shift
        FGrid1.SetFocus
        TxtGrid1(0).Visible = False
        Exit Sub
    End If
    
    Select Case FGrid1.Col
        Case C_AcName
            DGridTxtKeyDown DGParty, TxtGrid1, 0, RsParty, KeyCode, True, 1, frmSubGroup, "frmSubGroup"
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave = True Then
                    If Val(FGrid1.TextMatrix(FGrid1.Row, C_Credit)) > 0 Then
                        GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, C_ChqDt, , C_Credit
                    Else
                        GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, C_ChqDt, , C_Debit
                    End If
                Else
                    TxtGrid1_LostFocus 0
                    TxtGrid1(0).SetFocus
                End If
            End If
        Case C_Debit
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, C_ChqDt, , IIf(Val(TxtGrid1(0).Text) <> 0, IIf(SeparateNarr, C_Narration, C_ChqNo), C_Credit)
                Else
                    TxtGrid1_LostFocus 0
                    TxtGrid1(0).SetFocus
                End If
            End If
        Case C_Credit
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, C_ChqDt, , IIf(SeparateNarr, C_Narration, C_ChqNo)
                Else
                    TxtGrid1_LostFocus 0
                    TxtGrid1(0).SetFocus
                End If
            End If
        Case C_Narration
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, C_ChqDt
                Else
                    TxtGrid1_LostFocus 0
                    TxtGrid1(0).SetFocus
                End If
            End If
        Case C_ChqNo, C_ChqDt     ' , C_ClgDt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, C_ChqDt
                Else
                    TxtGrid1_LostFocus 0
                    TxtGrid1(0).SetFocus
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    CheckQuote KeyAscii
    Select Case FGrid1.Col
        Case C_Credit, C_Debit
            NumPress TxtGrid1(Index), KeyAscii, 10, 2
        Case C_AcName
            If DGParty.Visible = True Then DGridTxtKeyPress TxtGrid1, Index, RsParty, KeyAscii, "Name"
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case 0
        Select Case FGrid1.Col
            Case C_AcName
                If KeyCode <> 13 And DGParty.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid1, Index, RsParty, KeyCode, "Name", True
        End Select
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_validate TxtGrid1(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim i As Integer
Dim Mylen As Integer
    Select Case FGrid1.Col
        Case C_AcName
            If RsParty.RecordCount = 0 Or RsParty.EOF = True Or RsParty.BOF = True Or TxtGrid1(0).Text = "" Then
                FGrid1.TextMatrix(FGrid1.Row, C_AcCode) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_AcName) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_MainGrCode) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_CurrBal) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Debit) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Credit) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Narration) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Narr1) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Nature) = ""
                
                FGrid1.TextMatrix(FGrid1.Row, C_ChqDt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ChqNo) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ClgDt) = ""
            Else
                FGrid1.TextMatrix(FGrid1.Row, C_AcCode) = RsParty!Code
                FGrid1.TextMatrix(FGrid1.Row, C_AcName) = RsParty!Name
                FGrid1.TextMatrix(FGrid1.Row, C_MainGrCode) = RsParty!MainGrCode
                FGrid1.TextMatrix(FGrid1.Row, C_Nature) = IIf(IsNull(RsParty!Nature), "", RsParty!Nature)
                If IsNull(RsParty!Curr_Bal) Or RsParty!Curr_Bal = 0 Then
                    FGrid1.TextMatrix(FGrid1.Row, C_CurrBal) = ""
                Else
                    FGrid1.TextMatrix(FGrid1.Row, C_CurrBal) = Format(Abs(RsParty!Curr_Bal), "0.00") & " " & IIf(RsParty!Curr_Bal < 0, "Cr", "Dr")
                End If
                If Val(FGrid1.TextMatrix(FGrid1.Row, C_Credit)) + Val(FGrid1.TextMatrix(FGrid1.Row, C_Debit)) = 0 Then
                    If TempBal > 0 Then
                        FGrid1.TextMatrix(FGrid1.Row, C_Credit) = Format(Abs(TempBal), "0.00")
                    ElseIf TempBal < 0 Then
                        FGrid1.TextMatrix(FGrid1.Row, C_Debit) = Format(Abs(TempBal), "0.00")
                    End If
                    Call Calc_GridAmt
                End If
            End If
            If FGrid1.TextMatrix(FGrid1.Rows - 1, C_AcName) <> "" Then FGrid1.AddItem FGrid1.Rows
        
        Case C_Debit
            FGrid1.TextMatrix(FGrid1.Row, C_Debit) = Format(TxtGrid1(0).Text, "0.00")
            If Val(TxtGrid1(0).Text) > 0 Then
                FGrid1.TextMatrix(FGrid1.Row, C_Credit) = ""
            End If
            Call Calc_GridAmt
        Case C_Credit
            If Val(TxtGrid1(0).Text) > 0 Then
                FGrid1.TextMatrix(FGrid1.Row, C_Debit) = ""
            End If
            FGrid1.TextMatrix(FGrid1.Row, C_Credit) = Format(TxtGrid1(0).Text, "0.00")
            Call Calc_GridAmt
            
        Case C_Narration
            FGrid1.TextMatrix(FGrid1.Row, C_Narr1) = TxtGrid1(0).Text
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid1(0).Text
            txt(NarrDisp) = TxtGrid1(0).Text

        Case C_ChqNo
            If PubChqNoReq And UCase(FGrid1.TextMatrix(FGrid1.Row, C_Nature)) = "BANK" Then
                If TxtGrid1(0).Text = "" Then
                    MsgBox "Please Enter Chq/DD No.!", vbOKOnly, "Chq/DD No.Check"
                    TxtGrid1(0).Text = ""
'                    TxtGrid1Leave = False: Exit Function
                End If
            End If
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid1(0).Text
        Case C_ChqDt
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = RetDate(TxtGrid1(0))
            If FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) <> "" Then
                If CDate(FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)) > CDate(txt(Vdate)) Then
                    MsgBox "Cheque Date is greater than Voucher Date", vbOKOnly, "Cheque Date Validation"
                    FGrid1.SetFocus
                End If
            End If
        Case C_ClgDt
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = RetDate(TxtGrid1(0))
    End Select
NXT:
    TxtGrid1Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid1.SetFocus
    TxtGrid1(0).Visible = False
End If

End Function

Private Sub Fill_Grid()
Dim i As Integer
Dim Mylen As Integer
    FGrid1.Rows = 1
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    GSQL = "Select LG.*,Subgroup.Name,Subgroup.Nature,SubGroup.Curr_Bal,subgroup.GroupCode,Acgroup.MainGrCode from (Ledger as lg left join subgroup on lg.subcode=subgroup.subcode) left join acgroup on subgroup.GroupCode=Acgroup.GroupCode Where lg.DocId='" & Master!DocId & "'  and acgroup.aliasyn='N' order by lg.v_sno"
    Rst.Open GSQL, GCnFa, adOpenDynamic, adLockOptimistic

    i = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            With FGrid1
                .AddItem ""
                .TextMatrix(i, 0) = i
                .TextMatrix(i, C_AcCode) = Rst!SubCode
                .TextMatrix(i, C_AcName) = Rst!Name
                .TextMatrix(i, C_Nature) = IIf(IsNull(Rst!Nature), "", Rst!Nature)
                .TextMatrix(i, C_Credit) = IIf(Rst!AmtCr <> 0, Format(Rst!AmtCr, "0.00"), "")
                .TextMatrix(i, C_Debit) = IIf(Rst!AmtDr <> 0, Format(Rst!AmtDr, "0.00"), "")
                If SeparateNarr Then
                    Mylen = InStr(1, XNull(Rst!Narration), vbLf) - 2
                    If Mylen > 0 Then
                        .TextMatrix(i, C_Narration) = Mid(XNull(Rst!Narration), 1, Mylen)
                    Else
                        .TextMatrix(i, C_Narration) = XNull(Rst!Narration)
                    End If
                    .TextMatrix(i, C_Narr1) = XNull(Rst!Narration)
                End If
                .TextMatrix(i, C_ChqNo) = XNull(Rst!Chq_No)
                .TextMatrix(i, C_ChqDt) = XNull(Rst!Chq_Date)
                .TextMatrix(i, C_ClgDt) = XNull(Rst!clg_date)
'                .TextMatrix(i, C_CurrBal) = Abs(Rst!Curr_Bal) & " " & IIf(Rst!Curr_Bal < 0, "Cr", "Dr")
                If IsNull(Rst!Curr_Bal) Or Rst!Curr_Bal = 0 Then
                    FGrid1.TextMatrix(i, C_CurrBal) = ""
                Else
                    FGrid1.TextMatrix(i, C_CurrBal) = Format(Abs(Rst!Curr_Bal), "0.00") & " " & IIf(Rst!Curr_Bal < 0, "Cr", "Dr")
                End If
                .TextMatrix(i, C_MainGrCode) = Rst!MainGrCode
            End With
            i = i + 1
            Rst.MoveNext
        Loop
        FGrid1.AddItem FGrid1.Rows
        FGrid1.FixedRows = 1
    Else
        FGrid1.Rows = FGrid1.Rows
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
    End If
    Set Rst = Nothing
    FGrid1.Tag = FGrid1.Row
    FGrid1_RowColChange
    Call Calc_GridAmt
End Sub

Private Sub Calc_GridAmt()
Dim MyCr As Double, MyDr As Double, i As Integer
    MyCr = 0: MyDr = 0
    For i = 1 To FGrid1.Rows - 1
        If Val(FGrid1.TextMatrix(i, C_Debit)) > 0 Then
            MyDr = MyDr + Val(FGrid1.TextMatrix(i, C_Debit))
        End If
        If Val(FGrid1.TextMatrix(i, C_Credit)) > 0 Then
            MyCr = MyCr + Val(FGrid1.TextMatrix(i, C_Credit))
        End If
    Next i
    txt(TotCr).Text = Format(MyCr, "0.00")
    txt(TotDr).Text = Format(MyDr, "0.00")
    TempBal = Val(txt(TotDr)) - Val(txt(TotCr))
End Sub

Private Function AccGetDocID(ByVal Vtype As String, ByVal Vdate As String, ByRef VoucherEditFlag As Boolean, ByRef TxtSrlNo As Object, ByRef lblPrefix As Object, Optional ForSiteCode As String) As String
Dim Rst As ADODB.Recordset, VNo As Long
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select VT.Number_Method,VT.SerialNo_From_Table,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & Vtype & "' And VP.Date_From<=#" & Format(Vdate, "dd/MMM/yyyy") & "# Order By VP.Date_From DESC", GCnFa, adOpenStatic, adLockReadOnly
    lblPrefix = ""
    
    If Rst.RecordCount <= 0 Then
        If txt(VDesc).Text <> "" Then MsgBox "Voucher Numbring System not defined", vbInformation, "Validation": AccGetDocID = "": GoTo errlbl
        AccGetDocID = ""
        GoTo errlbl
    Else
        If Rst!Number_Method = "Manual" Then
            If Val(TxtSrlNo.Text) = 0 Then GoTo errlbl
            VoucherEditFlag = True
            TxtSrlNo.Enabled = True
            txt(VAdd).Enabled = True
            VNo = Val(TxtSrlNo)
        Else
            txt(VAdd).Text = ""
            VoucherEditFlag = False
            TxtSrlNo.Enabled = False
            txt(VAdd).Enabled = False
            VNo = Rst!Start_Srl_No + 1
        End If
    End If
    lblPrefix = Rst!Prefix
    TxtSrlNo = VNo
    AccGetDocID = PubDivCode + PubSiteCode + IIf(IsMissing(ForSiteCode), PubSiteCode, ForSiteCode) + Space(5 - Len(CStr(txt(VDesc).Tag))) + txt(VDesc).Tag + Space(1 - Len(txt(VAdd))) + txt(VAdd).Text + Space(4 - Len(CStr(Rst!Prefix))) + Rst!Prefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
errlbl:
    Set Rst = Nothing
End Function

Private Sub Remove_Curbal()
    Set Rst = GCnFa.Execute("Select ledger.*,subgroup.GroupCode,Acgroup.MainGrCode from (Ledger left join subgroup on ledger.subcode=subgroup.subcode) left join acgroup on subgroup.GroupCode=Acgroup.GroupCode where docid='" & lblDocId & "' and acgroup.aliasyn='N'")
    If Rst.RecordCount > 0 Then
        While Not Rst.EOF
            If Rst!AmtCr > 0 Then
                CalBalAcGroup "SubGroup", GCnFa, Rst!MainGrCode, Rst!AmtCr, "+"
                GCnFa.Execute ("Update Subgroup set curr_bal=Curr_bal+" & Rst!AmtCr & " where subcode='" & Rst!SubCode & "'")
                GCn.Execute ("Update Subgroup set curr_bal=Curr_bal+" & Rst!AmtCr & " where subcode='" & Rst!SubCode & "'")
            Else
                CalBalAcGroup "SubGroup", GCnFa, Rst!MainGrCode, Rst!AmtDr, "-"
                GCnFa.Execute ("Update Subgroup set curr_bal=Curr_bal-" & Rst!AmtDr & " where subcode='" & Rst!SubCode & "'")
                GCn.Execute ("Update Subgroup set curr_bal=Curr_bal-" & Rst!AmtDr & " where subcode='" & Rst!SubCode & "'")
            End If
            Rst.MoveNext
        Wend
    End If
    Set Rst = Nothing
End Sub

Private Sub ColScheme(Vtype As String)
    If Vtype = "F_AR" Then
        Me.BackColor = &HC8E8DA
    ElseIf Vtype = "F_BP" Then
        Me.BackColor = &HE8D8FE
    ElseIf Vtype = "F_CRN" Then
        Me.BackColor = &HB9D8EE
    ElseIf Vtype = "F_DRN" Then
        Me.BackColor = &HBAD3C9
    ElseIf Vtype = "F_JV" Then
        Me.BackColor = &HD7C6C8
    End If
End Sub

Private Sub SetMaxLength()
    Select Case FGrid1.Col
        Case C_AcName
            TxtGrid1(0).MaxLength = 40
            TxtGrid1(0).Alignment = 0
        Case C_Debit
            TxtGrid1(0).MaxLength = 13
            TxtGrid1(0).Alignment = 1
        Case C_Credit
            TxtGrid1(0).MaxLength = 13
            TxtGrid1(0).Alignment = 1
        Case C_Narration
            TxtGrid1(0).MaxLength = 255
            TxtGrid1(0).Height = FGrid1.RowHeight(0) * 3
            TxtGrid1(0).Alignment = 0
        Case C_ChqNo
            TxtGrid1(0).MaxLength = 15
            TxtGrid1(0).Alignment = 0
        Case C_ChqDt
            TxtGrid1(0).MaxLength = 12
            TxtGrid1(0).Alignment = 0
        Case Else
            TxtGrid1(0).Alignment = 0
            TxtGrid1(0).MaxLength = 0
'            TxtGrid1(0).Height = FGrid1.RowHeight(0)
    End Select
End Sub

Private Function CheqNoReq(Optional GridRow As Integer) As Boolean
Dim xI As Integer

    For xI = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(xI, C_AcCode) <> "" Then
            If FGrid1.TextMatrix(xI, C_Credit) <> "" Then
                If FGrid1.TextMatrix(xI, C_Nature) = "Cash" Then
                    CheqNoReq = False: Exit Function
                End If
            End If
        End If
    Next
    CheqNoReq = True
End Function
