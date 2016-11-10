VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmContract 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Contractor/OEM Master"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   11415
   Begin MSDataGridLib.DataGrid DGHelp 
      Height          =   3225
      Left            =   7785
      TabIndex        =   25
      Top             =   375
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5689
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   -2147483624
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   13504523
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "FinName"
         Caption         =   "Existing Contractor/OEM"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   4004.788
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   12
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "Yes"
      Top             =   3240
      Width           =   450
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   10
      Left            =   2220
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2970
      Width           =   4245
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   11
      Left            =   6510
      TabIndex        =   12
      Top             =   2970
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   9
      Left            =   2220
      MaxLength       =   15
      TabIndex        =   10
      Top             =   2700
      Width           =   3045
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2430
      Width           =   3045
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   7
      Left            =   6750
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1890
      Width           =   1335
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   6
      Left            =   2220
      MaxLength       =   50
      TabIndex        =   8
      Top             =   2160
      Width           =   5865
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   4
      Left            =   2220
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1620
      Width           =   5865
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   3
      Left            =   2220
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1350
      Width           =   5865
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2220
      MaxLength       =   6
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "OEM"
      Top             =   810
      Width           =   1830
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   5
      Left            =   2220
      MaxLength       =   25
      TabIndex        =   6
      Top             =   1890
      Width           =   3045
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2220
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1080
      Width           =   5865
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   3810
      Left            =   5730
      Negotiate       =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3870
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6720
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   -2147483624
      BorderStyle     =   0
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "City Name"
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
            ColumnWidth     =   2940.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGAc 
      Height          =   3810
      Left            =   615
      Negotiate       =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4005
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6720
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   -2147483624
      BorderStyle     =   0
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "A/c Name"
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
            ColumnWidth     =   3660.095
         EndProperty
      EndProperty
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contractor ( C ), OEM ( O ) , Insu.Authority ( I )"
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
      Height          =   225
      Index           =   4
      Left            =   4080
      TabIndex        =   28
      Top             =   825
      Width           =   3645
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active (Yes/No)"
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
      Height          =   225
      Index           =   10
      Left            =   735
      TabIndex        =   24
      Top             =   3255
      Width           =   1230
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger A/c"
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
      Height          =   225
      Index           =   9
      Left            =   735
      TabIndex        =   23
      Top             =   2985
      Width           =   870
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Height          =   225
      Index           =   8
      Left            =   735
      TabIndex        =   22
      Top             =   2715
      Width           =   285
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Height          =   225
      Index           =   7
      Left            =   735
      TabIndex        =   21
      Top             =   2445
      Width           =   540
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode"
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
      Height          =   225
      Index           =   6
      Left            =   6030
      TabIndex        =   20
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Height          =   225
      Index           =   5
      Left            =   735
      TabIndex        =   19
      Top             =   2175
      Width           =   1275
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   225
      Index           =   1
      Left            =   735
      TabIndex        =   18
      Top             =   1365
      Width           =   690
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Height          =   225
      Index           =   4
      Left            =   735
      TabIndex        =   17
      Top             =   555
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contractor/OEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   3
      Left            =   735
      TabIndex        =   16
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Height          =   225
      Index           =   2
      Left            =   735
      TabIndex        =   15
      Top             =   1905
      Width           =   300
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   735
      TabIndex        =   14
      Top             =   1095
      Width           =   510
   End
End
Attribute VB_Name = "frmContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstAc As ADODB.Recordset, mFlag As Byte
Dim RstCity As ADODB.Recordset
Private Const FinCode = 0, FinCatg = 1, FinName = 2, Add1 = 3, Add2 = 4, City = 5
Private Const ContactPerson = 6, PinCode = 7, Phone = 8, FAx = 9, AcName = 10, AcCode = 11, Ac_YN = 12

Private Sub DGAc_Click()
If RstAc.RecordCount > 0 Then
    txt(AcName) = RstAc!Name
    txt(AcCode) = RstAc!SubCode
End If
txt(AcName).SetFocus
DGAc.Visible = False
End Sub

Private Sub DGCity_Click()
If RstCity.RecordCount > 0 Then
    txt(City) = RstCity!Name
    txt(City).Tag = RstCity!Code
End If
txt(City).SetFocus
DgCity.Visible = False
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
WinSetting Me, 5600, 10380
TopCtrl1.Tag = PubUParam

DGHelp.left = Me.width - (DGHelp.width + mRtScale): DGHelp.top = mTopScale
DgCity.left = Me.width - (DgCity.width + mRtScale): DgCity.top = mTopScale
DGAc.left = Me.width - (DGAc.width + mRtScale): DGAc.top = mTopScale

Set RstMain = New ADODB.Recordset
RstMain.CursorLocation = adUseClient
If PubMoveRecYn Then
    RstMain.Open "Select ContractFinance.*,ContractFinance.FinCode as SearchCode,City.CityName,SUBGROUP.Name AS ACNAME " & _
        " From (ContractFinance Left Join SUBGROUP On ContractFinance.AcCode=SUBGROUP.SubCode) " & _
        " LEFT JOIN City ON ContractFinance.City = City.CityCode " & _
        " WHERE ContractFinance.FinCatg <> 0 Order by FinName", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 ContractFinance.*,ContractFinance.FinCode as SearchCode,City.CityName,SUBGROUP.Name AS ACNAME " & _
        " From (ContractFinance Left Join SUBGROUP On ContractFinance.AcCode=SUBGROUP.SubCode) " & _
        " LEFT JOIN City ON ContractFinance.City = City.CityCode " & _
        " WHERE ContractFinance.FinCatg <> 0 Order by FinName", GCn, adOpenDynamic, adLockOptimistic
End If

Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select FinCode,FinName FROM ContractFinance WHERE ContractFinance.FinCatg <> 0 Order by FinName", GCn, adOpenDynamic, adLockOptimistic
Set DGHelp.DataSource = RstHelp

Set RstAc = New ADODB.Recordset
RstAc.CursorLocation = adUseClient
RstAc.Open "Select SubCode,Name,NameHelp FROM SUBGROUP Order by NAME", GCn, adOpenDynamic, adLockOptimistic
Set DGAc.DataSource = RstAc

Set RstCity = New ADODB.Recordset
RstCity.CursorLocation = adUseClient
RstCity.Open "select citycode as code,cityname as name from city order by cityname,citycode", GCn, adOpenDynamic, adLockOptimistic
Set DgCity.DataSource = RstCity
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0: mFlag = 0
Exit Sub

ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstAc = Nothing
    Set RstCity = Nothing
End Sub

Private Sub Disp_Text(Enb As Boolean)
    txt(FinCode).Enabled = Enb
    txt(FinCatg).Enabled = Enb
    txt(FinName).Enabled = Enb
    txt(Add1).Enabled = Enb
    txt(Add2).Enabled = Enb
    txt(City).Enabled = Enb
    txt(ContactPerson).Enabled = Enb
    txt(PinCode).Enabled = Enb
    txt(Phone).Enabled = Enb
    txt(FAx).Enabled = Enb
    txt(AcName).Enabled = Enb
    txt(AcCode).Enabled = Enb
    txt(Ac_YN).Enabled = Enb
    If TopCtrl1.TopText2 = "Edit" Then
        txt(FinCode).Enabled = False
    End If

    txtDisabled_Color Me
End Sub

Private Sub MakeBlank()
    txt(FinCode) = ""
    txt(FinCatg) = "OEM"
    txt(FinName) = ""
    txt(Add1) = ""
    txt(Add2) = ""
    txt(City).Tag = ""
    txt(City) = ""
    txt(ContactPerson) = ""
    txt(PinCode) = ""
    txt(Phone) = ""
    txt(FAx) = ""
    txt(AcName) = ""
    txt(AcCode) = ""
    txt(Ac_YN) = "Yes"
End Sub

Private Sub MoveRec()
On Error GoTo Errloop
    Grid_Hide

If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    txt(FinCode) = RstMain!FinCode
    txt(FinCatg) = IIf(VNull(RstMain!FinCatg) = 2, "OEM", IIf(VNull(RstMain!FinCatg) = 3, "Insu.Authority", IIf(VNull(RstMain!FinCatg) = 1, "Contractor", "")))
    txt(FinName) = XNull(RstMain!FinName)
    txt(Add1) = XNull(RstMain!Add1)
    txt(Add2) = XNull(RstMain!Add2)
    txt(City).Tag = XNull(RstMain!City)
    txt(City) = XNull(RstMain!CityName)
    txt(ContactPerson) = XNull(RstMain!ContactPerson)
    txt(PinCode) = XNull(RstMain!PinCode)
    txt(Phone) = XNull(RstMain!Phone)
    txt(FAx) = XNull(RstMain!FAx)
    txt(AcName) = XNull(RstMain!AcName)
    txt(AcCode) = XNull(RstMain!AcCode)
    txt(Ac_YN) = IIf(RstMain!Ac_YN = "Y", "Yes", "No")
End If
'TopCtrl1.tPrn = False
TopCtrl1.tDel = False
Exit Sub
Errloop:        MsgBox err.Description
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
MakeBlank
If ADDFLAG <> 1 Then Disp_Text SETS("ADD", Me, RstMain)
ADDFLAG = 1
Txt_GotFocus FinCatg
txt(FinCatg).SelStart = Len(txt(FinCatg))
txt(FinCatg).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(FinName).Tag = txt(FinName)
    Txt_GotFocus FinName
    txt(FinName).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo Errloop
Dim transFalg As Byte
transFalg = 0
If RstMain.RecordCount > 0 Then
    If MsgBox("Are You Sure to Delete This Record", vbYesNo, "Confirmation") = vbYes Then
        GCn.BeginTrans
        transFalg = 1
        GCn.Execute ("Delete From ContractFinance Where FinCode=" & Chk_Text(PubSiteCode + Trim(txt(FinCode))))
        GCn.CommitTrans
        transFalg = 0
        RstMain.Requery
        RstHelp.Requery
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
    End If
Else
    MsgBox "There Is No Record To Delete.", vbInformation, "Information"
End If
Exit Sub
Errloop:    If transFalg = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbExclamation, " Deletion Error "
End Sub

Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, RstMain, 1
    MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, RstMain, 2
    MoveRec
End Sub

Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, RstMain, 3
    MoveRec
End Sub

Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, RstMain, 4
    MoveRec
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT FinCode as SearchCode,FinName as Name,FinCode as Code FROM ContractFinance WHERE ContractFinance.FinCatg <> 0 Order by FinName"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim rep As CrystalReport, Form1 As frmMastList
    Set Form1 = New frmMastList
    With Form1
        .g_FormID = 66
        .LblName.CAPTION = Me.CAPTION
        .CAPTION = Me.CAPTION
        .Show
    End With
    Set Form1 = Nothing
    Set rep = Nothing
End Sub

Private Sub TopCtrl1_eSave()
Dim transFlag As Byte, j As Integer, NewFinCode As Boolean, mFinBankCode2$
Dim CType As Integer
On Error GoTo Errloop
    Grid_Hide
    transFlag = 0
    If IsValid(txt(FinCatg), "Category") = False Then Txt_GotFocus FinCatg: Exit Sub
    If IsValid(txt(FinName), "Name") = False Then Txt_GotFocus FinName: Exit Sub
    
    If ADDFLAG = 1 Then
        'Auto Code generation
        j = 1
        Do Until NewFinCode
            mFinBankCode2 = PubSiteCode & Right("00000" & j, 5)
            If GCn.Execute("select FinCode from ContractFinance where FinCode='" & mFinBankCode2 & "'").RecordCount <= 0 Then
                txt(FinCode) = mFinBankCode2
                NewFinCode = Not NewFinCode
            End If
            j = j + 1
        Loop
    End If
        
    GCn.BeginTrans
    transFlag = 1
     
    If txt(FinCatg) = "Contractor" Then
        CType = 1
    ElseIf txt(FinCatg) = "Insu.Authority" Then
        CType = 3
    Else
        CType = 2
    End If
     
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("Insert Into ContractFinance (FinCode,Site_Code,FinCatg,FinName,Add1,Add2,City,ContactPerson,PinCode,Phone,Fax,AcCode,Ac_YN,U_Name,U_EntDt,U_AE)" & _
                    "Values('" & txt(FinCode) & "','" & PubSiteCode & "'," & CType & "," & Chk_Text(txt(FinName)) & "," & Chk_Text(txt(Add1)) & "," & Chk_Text(txt(Add2)) & ",'" & txt(City).Tag & "'," & Chk_Text(txt(ContactPerson)) & "," & Chk_Text(txt(PinCode)) & "," & Chk_Text(txt(Phone)) & "," & Chk_Text(txt(FAx)) & "," & Chk_Text(txt(AcCode)) & ",'" & left(txt(Ac_YN), 1) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Else
        GCn.Execute ("update ContractFinance set Site_Code='" & PubSiteCode & "',FinCatg=" & CType & ",FinName=" & Chk_Text(txt(FinName)) & "" & _
                    ",Add1=" & Chk_Text(txt(Add1)) & ",Add2=" & Chk_Text(txt(Add2)) & ",City='" & txt(City).Tag & "',ContactPerson=" & Chk_Text(txt(ContactPerson)) & ",PinCode=" & Chk_Text(txt(PinCode)) & ",Phone=" & Chk_Text(txt(Phone)) & ",Fax=" & Chk_Text(txt(FAx)) & ",AcCode=" & Chk_Text(txt(AcCode)) & ",Ac_YN='" & left(txt(Ac_YN), 1) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & " Where FINCODE='" & txt(FinCode) & "'")
    End If
    GCn.CommitTrans
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select ContractFinance.*,ContractFinance.FinCode as SearchCode,City.CityName,SUBGROUP.Name AS ACNAME " & _
            " From (ContractFinance Left Join SUBGROUP On ContractFinance.AcCode=SUBGROUP.SubCode) " & _
            " LEFT JOIN City ON ContractFinance.City = City.CityCode " & _
            " WHERE ContractFinance.FinCatg <> 0  And ContractFinance.FinCode = '" & txt(FinCode) & "'  Order by FinName")
    End If
    RstHelp.Requery
    RstMain.FIND ("FinCode='" & txt(FinCode) & "'")
    If ADDFLAG = 1 Then
        TopCtrl1_eAdd
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        ADDFLAG = 0
    End If
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
    If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eRef()
    RstHelp.Requery
    RstAc.Requery
    RstCity.Requery
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus txt(Index)
Grid_Hide
    Select Case Index
        Case FinName
            RstHelp.Sort = "FINNAME ASC"
        Case City
            RstCity.Sort = "NAME ASC"
            If RstCity.RecordCount = 0 Or (RstCity.EOF = True Or RstCity.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RstCity!Name Then
                RstCity.MoveFirst
                RstCity.FIND "name ='" & txt(Index) & "'"
            End If
        Case AcName
            RstAc.Sort = "NAME ASC"
            If RstAc.RecordCount = 0 Or (RstAc.EOF = True Or RstAc.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RstAc!Name Then
                RstAc.MoveFirst
                RstAc.FIND "name ='" & txt(Index) & "'"
            End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case FinName
        DGridTxtKeyDown_Mast DGHelp, txt, Index, RstHelp, KeyCode, False, 1
    Case City
        DGridTxtKeyDown DgCity, txt, Index, RstCity, KeyCode, False, 1, frmCity, "frmCity"
    Case AcName
        DGridTxtKeyDown DGAc, txt, Index, RstAc, KeyCode, False, 1
End Select
If DGHelp.Visible = False And DgCity.Visible = False And DGAc.Visible = False Then
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
        If Index = Ac_YN Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then TopCtrl1_eSave
        Else
            Ctrl_DownKeyDown KeyCode, Shift
        End If
    End If
    If Index <> FinName Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Select Case Index
    Case FinCatg
        If UCase(Chr(keyascii)) = "O" Then
            txt(FinCatg) = "OEM"
        ElseIf UCase(Chr(keyascii)) = "C" Then
            txt(FinCatg) = "Contractor"
        ElseIf UCase(Chr(keyascii)) = "I" Then
            txt(FinCatg) = "Insu.Authority"
        End If
        keyascii = 0
    Case City
        If DgCity.Visible Then DGridTxtKeyPress txt, Index, RstCity, keyascii, "Name"
    Case AcName
        If DGAc.Visible Then DGridTxtKeyPress txt, Index, RstAc, keyascii, "Name"
    Case Ac_YN
        If UCase(Chr(keyascii)) = "Y" Then
            txt(Index) = "Yes"
        ElseIf UCase(Chr(keyascii)) = "N" Then
            txt(Index) = "No"
        ElseIf keyascii = vbKeyBack Or keyascii = vbKeyDelete Then
            txt(Index) = ""
        End If
        keyascii = 0
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case FinName
        If DGHelp.Visible Then DGridTxtKeyUp_Mast txt, Index, RstHelp, KeyCode, "FinName"
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case FinName
        If TopCtrl1.TopText2 = "Edit" Then
            If UCase(txt(FinName)) = UCase(txt(FinName).Tag) Then: Exit Sub
        End If
        If GCn.Execute("SELECT FinName FROM ContractFinance WHERE FinName='" & txt(FinName) & "'").RecordCount > 0 Then
            MsgBox "Name already exists", vbCritical, "Validation Error"
            txt(Index) = txt(Index).Tag
            Cancel = True
            Exit Sub
        End If
    Case City
        If RstCity.RecordCount = 0 Or (RstCity.EOF = True Or RstCity.BOF = True) Or txt(Index) = "" Then
            txt(Index) = ""
            txt(Index).Tag = ""
        Else
            txt(Index) = RstCity!Name
            txt(Index).Tag = RstCity!Code
        End If
    Case AcName
        If RstAc.RecordCount = 0 Or (RstAc.EOF = True Or RstAc.BOF = True) Or txt(AcName) = "" Then
            txt(AcCode) = ""
            txt(AcName) = ""
        Else
            txt(AcCode) = RstAc!SubCode
            txt(AcName) = RstAc!Name
        End If
End Select
End Sub
Private Sub Grid_Hide()
    If DGHelp.Visible Then DGHelp.Visible = False
    If DgCity.Visible Then DgCity.Visible = False
    If DGAc.Visible Then DGAc.Visible = False
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select ContractFinance.*,ContractFinance.FinCode as SearchCode,City.CityName,SUBGROUP.Name AS ACNAME " & _
            " From (ContractFinance Left Join SUBGROUP On ContractFinance.AcCode=SUBGROUP.SubCode) " & _
            " LEFT JOIN City ON ContractFinance.City = City.CityCode " & _
            " WHERE ContractFinance.FinCatg <> 0  And ContractFinance.FinCode = '" & MyValue & "'  Order by FinName")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

