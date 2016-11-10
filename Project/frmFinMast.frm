VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmFinMast 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "Financier Master"
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
   LinkTopic       =   " "
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1785
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1500
      Width           =   705
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2205
      Left            =   12135
      Negotiate       =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   3889
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
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
         Weight          =   700
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGHlp 
      Height          =   5205
      Left            =   2730
      Negotiate       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6990
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   9181
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
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
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Name"
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
            DividerStyle    =   6
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   3600
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   1785
      MaxLength       =   35
      TabIndex        =   3
      Top             =   1230
      Width           =   4905
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   967
      Width           =   4905
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
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1785
      MaxLength       =   4
      TabIndex        =   1
      Top             =   690
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   930
      TabIndex        =   6
      Top             =   4170
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3765
      Left            =   75
      TabIndex        =   5
      Top             =   2700
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   6641
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   15
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   12243913
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   8421504
      GridColorFixed  =   32896
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmFinMast.frx":0000
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
      _Band(0).Cols   =   15
      _Band(0).GridLineWidthBand=   1
   End
   Begin MSDataGridLib.DataGrid DGFinGrp 
      Height          =   4935
      Left            =   5145
      Negotiate       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6975
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
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
         Caption         =   "Finance Group Name"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   2205
      Left            =   1245
      Negotiate       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6945
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   3889
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Inv. Prefix"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   15
      Top             =   1515
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FEE0FD&
      BackStyle       =   0  'Transparent
      Caption         =   "Branch Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   14
      Top             =   2460
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier Group ...."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   1245
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier Name ......"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   31
      Left            =   135
      TabIndex        =   9
      Top             =   975
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier Code ......"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   7
      Top             =   705
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmFinMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim RsFinGrp As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsHlp As ADODB.Recordset
Dim Master As ADODB.Recordset

Dim ForeColorSelEnter$
Private BackColorSelLeave$

Dim GridKey As Integer
Dim DocID As String

Private Const VehInvPrefix  As Byte = 0
Private Const FinCode       As Byte = 1
Private Const FinName       As Byte = 2
Private Const FinGrp        As Byte = 3
Private Const PayOutPer     As Byte = 4

'Col Declaration
'SrNo.|Financier Code|Contact Person|First Address|Second Address|City|Pin |Phone |Fax |ActiveYN |A/c Name |Citycode |AcCode |SrNo
' 0     1                 2              3               4        5    6     7     8     9         10         11        12    13
Private Const SrNo As Byte = 0
Private Const ConPer As Byte = 1
Private Const Add1 As Byte = 2
Private Const Add2 As Byte = 3
Private Const City As Byte = 4
Private Const Pin As Byte = 5
Private Const Phone As Byte = 6
Private Const FAx As Byte = 7
Private Const ActiveYN As Byte = 8
Private Const AcName  As Byte = 9
Private Const CityCode  As Byte = 10
Private Const AcCode  As Byte = 11
Private Const FinCode2 As Byte = 12
Private Const AddEdit As Byte = 13

Dim TAddMode As Boolean

Private Sub FGrid_LostFocus()
If TxtGrid(0).Visible = False Then
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End If
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
    WinSetting Me: Ini_Grid: TopCtrl1.Tag = PubUParam
    Ini_Grid

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
        Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If



    If PubMoveRecYn Then
        Master.Open "select FinBankCode as searchcode,FinBank.* from FinBank " & sitecond & " order by FinBankCode", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 FinBankCode as searchcode,FinBank.* from FinBank " & sitecond & " order by FinBankCode", GCn, adOpenDynamic, adLockOptimistic
    End If

    Set RsHlp = New ADODB.Recordset
    RsHlp.CursorLocation = adUseClient
    RsHlp.Open "select FinBank.FinBankCode as code,FinBank.FinBankname as name from FinBank order by FinBankCode", GCn, adOpenDynamic, adLockOptimistic
    Set DgHlp.DataSource = RsHlp
    RsHlp.Sort = "code"
    RsHlp.Sort = "name"

    Set RsFinGrp = New ADODB.Recordset
    RsFinGrp.CursorLocation = adUseClient
    RsFinGrp.Open "select FinGroup.FinGrpCode as code,FinGroup.FinGrpname as name from FinGroup order by FinGroup.FinGrpname", GCn, adOpenDynamic, adLockOptimistic
    Set DGFinGrp.DataSource = RsFinGrp
    
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "select citycode as code,cityname as name from city order by cityname,citycode", GCn, adOpenDynamic, adLockOptimistic
    Set DgCity.DataSource = RsCity

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME from SubGroup order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
Set RsFinGrp = Nothing
Set RsParty = Nothing
Set RsCity = Nothing
Set RsHlp = Nothing
Set Master = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    txt(FinName).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer
If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    GCn.BeginTrans
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, FinCode2) <> "" Then
            GCn.Execute ("delete from ContractFinance where FinBankCode = '" & Master!FinbankCode & "'")
        End If
    Next
    GCn.Execute ("delete from FinBank where FinBankCode = '" & Master!FinbankCode & "'")
    GCn.CommitTrans
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If err.NUMBER <> 0 Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    DocID = txt(FinName).TEXT
'    Txt(VehInvPrefix).SetFocus
    FGrid.AddItem FGrid.Rows
    FGrid.SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
'    Master.Cancel
    Unload Me
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
If MasterFormExit Then Unload Me: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
    RsFinGrp.Requery
    RsCity.Requery
    RsHlp.Requery
    RsParty.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Long, j As Integer, mFinBankCode$, mFinBankCode2$
    Dim mTrans As Boolean, mGridFilled As Boolean, NewFinCode As Boolean
    
   ' On Error GoTo errlbl
    
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    
    If IsValid(txt(FinName), "Financier Name") = False Then Exit Sub
    If IsValid(txt(FinGrp), "Financier Group") = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Add1) <> "" Then
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Branch Details", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Add1: FGrid.SetFocus: Exit Sub
    
RemoveTxtNull
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        'Auto Code generation
        mFinBankCode = PubSiteCode & VNull(GCn.Execute("Select Max(" & cVal(cMID("FinBankCode", "2", "4")) & ") + 1 From FinBank").Fields(0).Value)
        txt(FinCode) = mFinBankCode
'        I = 1
'        Do Until NewFinCode
'            mFinBankCode = PubSiteCode & Right("00" & I, 2)
'            If GCn.Execute("select FinBankCode from FinBank where FinBankCode='" & mFinBankCode & "'").RecordCount <= 0 Then
'                txt(FinCode) = mFinBankCode
'                NewFinCode = Not NewFinCode
'            End If
'            I = I + 1
'        Loop
    End If
    
    GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GCn.Execute ("insert into FinBank(Site_Code,FinbankCode,FinBankName,FinGrpCode, Inv_Prefix,U_Name,U_EntDt,U_AE) " & _
            " VALUES('" & PubSiteCode & "','" & txt(FinCode) & "','" & txt(FinName) & "','" & txt(FinGrp).Tag & "', '" & txt(VehInvPrefix) & "', '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
    Else
        GCn.Execute ("update FinBank set FinBankName = '" & txt(FinName) & "',FinGrpCode = '" & txt(FinGrp).Tag & "', " & _
                     "Inv_Prefix='" & txt(VehInvPrefix) & "',  " & _
                     "U_Name = '" & pubUName & "',U_EntDt = " & ConvertDate(PubServerDate) & ", U_AE = '" & left(TopCtrl1.TopText2, 1) & "' " & _
                     "Where FinBankCode = '" & txt(FinCode) & "'")
    End If
    For I = 1 To FGrid.Rows - 1
        GSQL = ""
        If FGrid.TextMatrix(I, AddEdit) = "E" Then
            'Case of Edit
            GSQL = "Update ContractFinance set UnderFinGrp='" & txt(FinGrp).Tag & "',FinName='" & txt(FinName) & "',Add1='" & FGrid.TextMatrix(I, Add1) & "',Add2='" & FGrid.TextMatrix(I, Add2) & _
                "',City='" & FGrid.TextMatrix(I, CityCode) & "',ContactPerson='" & FGrid.TextMatrix(I, ConPer) & "',PinCode='" & FGrid.TextMatrix(I, Pin) & "',Phone='" & FGrid.TextMatrix(I, Phone) & _
                "',Fax='" & FGrid.TextMatrix(I, FAx) & "',AcCode='" & FGrid.TextMatrix(I, AcCode) & "',Ac_YN=" & IIf(FGrid.TextMatrix(I, AcCode) = "", 0, 1) & ",U_Name='" & pubUName & _
                "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'  " & _
                " where FinCode = '" & FGrid.TextMatrix(I, FinCode2) & "'"
        ElseIf FGrid.TextMatrix(I, AddEdit) = "" Then
            If FGrid.TextMatrix(I, Add1) <> "" Or _
                FGrid.TextMatrix(I, Add2) <> "" Or _
                FGrid.TextMatrix(I, CityCode) <> "" Then
                
                mFinBankCode2 = PubSiteCode & Format(VNull(GCn.Execute("Select (Max(" & cVal(cTrim(cMID("FinCode", "3", "4"))) & ") + 1) From ContractFinance Where " & cTrim(cMID("FinCode", "3", "4")) & "<>'' ").Fields(0).Value), "00000")
                              
                GSQL = "Insert Into ContractFinance (FinCode,Site_Code,FinCatg,FinBankCode,UnderFinGrp,FinName,Add1,Add2,City,ContactPerson,PinCode,Phone,Fax,AcCode,Ac_YN,U_Name,U_EntDt,U_AE) " & _
                    " Values ('" & mFinBankCode2 & "','" & PubSiteCode & "', 0,'" & txt(FinCode) & "','" & txt(FinGrp).Tag & "','" & txt(FinName) & "','" & FGrid.TextMatrix(I, Add1) & _
                    "','" & FGrid.TextMatrix(I, Add2) & "','" & FGrid.TextMatrix(I, CityCode) & "','" & FGrid.TextMatrix(I, ConPer) & "','" & FGrid.TextMatrix(I, Pin) & "','" & FGrid.TextMatrix(I, Phone) & _
                    "','" & FGrid.TextMatrix(I, FAx) & "','" & FGrid.TextMatrix(I, AcCode) & "'," & IIf(FGrid.TextMatrix(I, AcCode) = "", 0, 1) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            End If
        End If
        If GSQL <> "" Then
            GCn.Execute GSQL
        End If
    Next
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select FinBankCode as searchcode,FinBank.* from FinBank Where  FinBankCode ='" & txt(FinCode).TEXT & "' order by FinBankCode")
    End If
    RsHlp.Requery
    Master.FIND "FinBankCode = '" & txt(FinCode).TEXT & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
    
errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
        Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If


    GSQL = "select FinBankCode as SearchCode,FinBank.FinBankCode as Code,FinBank.FinBankname as Name from FinBank " & sitecond & " order by FinBankCode,FinBankName"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    CheckError
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select FinBankCode as searchcode,FinBank.* from FinBank Where  FinBankCode ='" & MyValue & "' order by FinBankCode")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
TxtGrid(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case FinGrp
        If RsFinGrp.RecordCount = 0 Or (RsFinGrp.EOF = True Or RsFinGrp.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsFinGrp!Name Then
            RsFinGrp.MoveFirst
            RsFinGrp.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case FinCode
        If RsHlp.RecordCount = 0 Or (RsHlp.EOF = True Or RsHlp.BOF = True) Then Exit Sub
        RsHlp.Sort = "code"
    Case FinName
        If RsHlp.RecordCount = 0 Or (RsHlp.EOF = True Or RsHlp.BOF = True) Then Exit Sub
        RsHlp.Sort = "Name"
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case FinGrp
            DGridTxtKeyDown DGFinGrp, txt, Index, RsFinGrp, KeyCode, False, 1
    Case FinCode
          If DgHlp.Visible = False Then DGridColSwap DgHlp, 0
            DGridTxtKeyDown_Mast DgHlp, txt, Index, RsHlp, KeyCode, False, 0
    Case FinName
          If DgHlp.Visible = False Then DGridColSwap DgHlp, 1
            DGridTxtKeyDown_Mast DgHlp, txt, Index, RsHlp, KeyCode, False, 1
End Select
If DgHlp.Visible = False And DgCity.Visible = False And DGFinGrp.Visible = False And DGParty.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> FinCode Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> FinName Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case Index
    Case FinGrp
        If DGFinGrp.Visible = True Then DGridTxtKeyPress txt, Index, RsFinGrp, KeyAscii, "Name"
    Case PayOutPer
        NumPress txt(Index), KeyAscii, 2, 2
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case FinCode
        If DgHlp.Visible = True Then DGridTxtKeyUp_Mast txt, Index, RsHlp, KeyCode, "Code"
    Case FinName
        If DgHlp.Visible = True Then DGridTxtKeyUp_Mast txt, Index, RsHlp, KeyCode, "Name"
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer
Select Case Index
    Case FinCode
        If GCn.Execute("select count(*) from FinBank where FinBankCode='" & txt(Index).TEXT & "'").Fields(0) > 0 Then
            MsgBox "Duplicate Code  No.", vbCritical, "Validation Error"
            txt(Index).TEXT = ""
            Cancel = True
            Exit Sub
        End If
        
    Case FinName
        If DocID <> txt(FinName).TEXT Then
            If GCn.Execute("select count(*) from FinBank where FinBankname='" & txt(Index).TEXT & "'").Fields(0) > 0 Then
                MsgBox "Duplicate Name", vbCritical, "Validation Error"
                txt(Index).TEXT = ""
                Cancel = True
                Exit Sub
            End If
        End If
    Case FinGrp
        If IsValid(txt(Index), "Finance Group") = False Then Cancel = True: Exit Sub
        If RsFinGrp.RecordCount = 0 Or (RsFinGrp.EOF = True Or RsFinGrp.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsFinGrp!Name
            txt(Index).Tag = RsFinGrp!Code
        End If
End Select
End Sub

Private Sub DgHlp_Click()
    DgHlp.Visible = False
End Sub

Private Sub DGFinGrp_Click()
    If RsFinGrp.RecordCount > 0 Then
        txt(FinGrp).TEXT = RsFinGrp!Name
        txt(FinGrp).Tag = RsFinGrp!Code
    End If
    DGFinGrp.Visible = False
    txt(FinGrp).SetFocus
End Sub
Private Sub DGCity_Click()
    If RsCity.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsCity!Name
        FGrid.TextMatrix(FGrid.Row, City) = RsCity!Name
        FGrid.TextMatrix(FGrid.Row, CityCode) = RsCity!Code
    End If
    TxtGrid(0).SetFocus
    DgCity.Visible = False
End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsParty!Name
        FGrid.TextMatrix(FGrid.Row, AcName) = RsParty!Name
        FGrid.TextMatrix(FGrid.Row, AcCode) = RsParty!Code
    End If
    TxtGrid(0).SetFocus
    DGParty.Visible = False
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid_KeyPress (vbKeyReturn)
End Sub

Private Sub FGrid_GotFocus()
    If FGrid.BackColorSel = BackColorSelLeave Then FGrid.Col = 1
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case ConPer, Add1, Add2, City, Pin, FAx, Phone
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGridAddEditDel
        Case City
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, CityCode) = ""
            FGridAddEditDel
        Case AcName
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, AcCode) = ""
            FGridAddEditDel
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
    SetMaxLength
    Select Case FGrid.Col
        Case FinCode2
            If FGrid.TextMatrix(FGrid.Row, AddEdit) = "" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, Asc(UCase(Chr(KeyAscii))))
                If Len(TxtGrid(0).TEXT) <= 1 Then TxtGrid(0).TEXT = PubSiteCode & TxtGrid(0).TEXT
                TxtGrid(0).SelStart = Len(TxtGrid(0).TEXT)
            End If
        Case ConPer, Add1, Add2, City, Pin, FAx, Phone, City, ActiveYN, AcName
            Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub

Dim I As Integer, mSrlNo As Integer
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.TextMatrix(FGrid.Row, AddEdit) = "" Then
                FGrid.Redraw = False
                FGrid.RemoveItem (FGrid.Row)
                FGrid.Redraw = True
            Else
                MsgBox "Delete Denied !", vbOKOnly, "Validation": FGrid.Redraw = True: FGrid.SetFocus: Exit Sub
            End If
            For I = 1 To FGrid.Rows - 1
                If FGrid.RowHeight(I) > 0 Then
                    mSrlNo = mSrlNo + 1
                    FGrid.TextMatrix(I, SrNo) = mSrlNo
                End If
            Next
        End If
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.Row = 1
    FGrid.SetFocus
End If
Exit Sub

ELoop:
    CheckError
End Sub

Private Sub FGrid_Scroll()
Grid_Hide
TxtGrid(0).Visible = False
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub

Private Sub MoveRec()
Dim I As Integer
On Error GoTo error1
If Master.RecordCount > 0 Then
    Dim Rs As Recordset
    txt(FinCode).TEXT = Master!FinbankCode
    txt(FinName).TEXT = Master!finbankname
    txt(FinGrp).Tag = Master!Fingrpcode
    If txt(FinGrp).Tag <> "" Then
        txt(FinGrp).TEXT = GCn.Execute("select FinGrpName from fingroup where FinGrpcode = '" & txt(FinGrp).Tag & "'").Fields(0).Value
    Else
        txt(FinGrp).TEXT = ""
    End If
    'Txt(PayOutPer) = Format(IIf(IsNull(Master!PayOutPer), 0, Master!PayOutPer), "0.00")
    txt(VehInvPrefix) = IIf(IsNull(Master!Inv_Prefix), "", Master!Inv_Prefix)

    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT CF.FinCode, CF.FinName, CF.Add1, CF.Add2, CF.ContactPerson, CF.City, CF.PinCode, CF.Phone, CF.Fax, CF.Ac_YN, CF.AcCode, City.CityName, SubGroup.Name as Party " & _
            " FROM ((ContractFinance as CF " & _
            " LEFT JOIN FinBank ON CF.FinBankCode = FinBank.FinBankCode) " & _
            " LEFT JOIN SubGroup ON CF.AcCode = SubGroup.Subcode) " & _
            " LEFT JOIN City ON CF.City = City.CityCode " & _
            " where CF.FinBankCode = '" & Master!FinbankCode & "'")
    FGrid.Redraw = False
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        I = 1
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, SrNo) = I
                .TextMatrix(I, FinCode2) = Rs!FinCode
                .TextMatrix(I, ConPer) = Rs!ContactPerson
                .TextMatrix(I, Add1) = IIf(IsNull(Rs!Add1), "", Rs!Add1)
                .TextMatrix(I, Add2) = IIf(IsNull(Rs!Add2), "", Rs!Add2)
                .TextMatrix(I, City) = IIf(IsNull(Rs!CityName), "", Rs!CityName)
                .TextMatrix(I, Pin) = IIf(IsNull(Rs!PinCode), "", Rs!PinCode)
                .TextMatrix(I, Phone) = IIf(IsNull(Rs!Phone), "", Rs!Phone)
                .TextMatrix(I, FAx) = IIf(IsNull(Rs!FAx), "", Rs!FAx)
                .TextMatrix(I, ActiveYN) = IIf(Rs!Ac_YN = 1, "Yes", "No")
                .TextMatrix(I, AcName) = IIf(IsNull(Rs!Party), "", Rs!Party)
                .TextMatrix(I, CityCode) = IIf(IsNull(Rs!City), "", Rs!City)
                .TextMatrix(I, AcCode) = IIf(IsNull(Rs!AcCode), "", Rs!AcCode)
                .TextMatrix(I, AddEdit) = "N"
            End With
            Rs.MoveNext
            I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
        TopCtrl1.tDel = False
    End If
    FGrid.Redraw = True
    Set Rs = Nothing
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
Grid_Hide
TopCtrl1.tDel = False
TopCtrl1.tPrn = False
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte

DGFinGrp.left = 6700: DGFinGrp.top = mTopScale
DgHlp.left = 6700: DgHlp.top = mTopScale
DgCity.left = 6700: DgCity.top = mTopScale
DGParty.left = 6700: DGParty.top = mTopScale
'SrNo.|Financier Code|First Address 1|Second Address 2|Contact Person 3|City 4|Pin 5|Phone 6|Fax 7|ActiveYN 8|A/C Name 9|Citycode 10|Accode 11|Srno 12
' 0     1                 2            3                   4               5     6      7      8     9         10         12          13         14

With FGrid
    .left = Me.left ' + 45
    .width = Me.width - 200
    .top = 2700
    .Cols = 14
    .RowHeightMin = 0 'PubGridRowHeight
    .ColAlignmentFixed = flexAlignCenterCenter
    .ColAlignment = flexAlignLeftCenter

    .TextMatrix(0, 0) = "Srl"
    .ColAlignmentFixed(0) = flexAlignRightCenter
    .ColWidth(0) = 400
    
    .TextMatrix(0, FinCode2) = "Code"
    .ColWidth(FinCode2) = 0 '780
    
    .TextMatrix(0, ConPer) = "Contact Person"
    .ColAlignment(ConPer) = flexAlignLeftCenter
    .ColWidth(ConPer) = 2205
    
    .TextMatrix(0, Add1) = "Address-1"
    .ColWidth(Add1) = 1995
    
    .TextMatrix(0, Add2) = "Address-2"
    .ColWidth(Add2) = 1995
    
    .TextMatrix(0, City) = "City"
    .ColWidth(City) = 1095
    
    .TextMatrix(0, Pin) = "Pin"
    .ColWidth(Pin) = 645
    
    .TextMatrix(0, Phone) = "Phone"
    .ColWidth(Phone) = 2000
    
    .TextMatrix(0, FAx) = "Fax"
    .ColWidth(FAx) = 825
    
    .TextMatrix(0, ActiveYN) = "A/cYN"
    .ColWidth(ActiveYN) = 0 '540
    
    .TextMatrix(0, AcName) = "Ledger A/c Name"
    .ColAlignmentFixed(AcName) = flexAlignLeftCenter
    .ColWidth(AcName) = 2265
    
    .ColWidth(CityCode) = 0
    .ColWidth(AcCode) = 0
    .ColWidth(AddEdit) = 0
End With
BackColorSelLeave = FGrid.BackColor
ForeColorSelEnter = FGrid.ForeColorSel
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next
If TopCtrl1.TopText2 = "Edit" Then
     txt(FinCode).Enabled = False
End If

txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol

End Sub
 
Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
         Case City
            If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or FGrid.TextMatrix(FGrid.Row, City) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, City) <> RsCity!Name Then
                RsCity.MoveFirst
                RsCity.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, City) & "'"
            End If
        Case AcName
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or FGrid.TextMatrix(FGrid.Row, AcName) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, AcName) <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, AcName) & "'"
            End If
'       Case FinCode2
'            TxtGrid(0).MaxLength = 6
'       Case ConPer
'            TxtGrid(0).MaxLength = 50
'       Case Add1, Add2
'            TxtGrid(0).MaxLength = 40
'       Case Pin
'            TxtGrid(0).MaxLength = 6
'       Case Phone
'            TxtGrid(0).MaxLength = 20
'       Case FAx
'            TxtGrid(0).MaxLength = 15
End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    Grid_Hide
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Exit Sub
End If
Select Case FGrid.Col
    Case City   '3
        DGParty.left = 6645: DGParty.top = mTopScale
        DGridTxtKeyDown DgCity, TxtGrid, Index, RsCity, KeyCode, True, 1, frmCity, "frmCity"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
               GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 11
            End If
        End If
    Case AcName
        DGridTxtKeyDown DGParty, TxtGrid, Index, RsParty, KeyCode, True, 1, frmSubGroup, "frmSubGroup"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
               GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 11
            End If
        End If
    Case FinCode2
        If FGrid.TextMatrix(FGrid.Row, AddEdit) = "" Then
            KeyCode = RestrictKey(1, KeyCode, TxtGrid(Index), Shift)
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 11
                End If
            End If
        End If
    Case ConPer, Add1, Add2, Pin, FAx, Phone, ActiveYN
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 11
            End If
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case FGrid.Col   'Index
    Case FinCode2
        KeyAscii = RestrictKey(1, Asc(UCase(Chr(KeyAscii))), TxtGrid(Index), 0)
    Case City
       If DgCity.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsCity, KeyAscii, "Name"
    Case AcName
       If DGParty.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsParty, KeyAscii, "name"
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case City
        If KeyCode <> 13 And DgCity.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, 0, RsCity, KeyCode, "Name", True
    Case AcName
        If KeyCode <> 13 And DGParty.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsParty, KeyCode, "name", True
    Case ActiveYN
        If Len(TxtGrid(0)) = 0 Or UCase(mID(TxtGrid(0), 1, 1)) = "N" Then
            TxtGrid(0) = "No"
        ElseIf UCase(mID(TxtGrid(0), 1, 1)) = "Y" Then
            TxtGrid(0) = "Yes"
        Else
            TxtGrid(0) = "No"
        End If
    Case ConPer, Add1, Add2, Pin, FAx, Phone, ActiveYN
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
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
Dim I As Integer
Select Case FGrid.Col
    Case FinCode2
        For I = 1 To FGrid.Rows - 1
            If I = FGrid.Row Then GoTo nxt1
            If UCase(TxtGrid(0).TEXT) = UCase(FGrid.TextMatrix(I, FinCode2)) And FGrid.TextMatrix(I, FinCode2) <> "" Then
                MsgBox "Code already added on Serial No. " & I, vbInformation, "Grid Validation"
                TxtGridLeave = False: Exit Function
            End If
nxt1:
        Next
        If GCn.Execute("Select COUNT(*) From ContractFinance Where FINCODE='" & TxtGrid(0).TEXT & "'").Fields(0).Value > 0 Then
            MsgBox "Code Already Exists", vbInformation, "Validation"
            TxtGridLeave = False: Exit Function
        End If
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT: FGridAddEditDel
        If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
    Case AcName
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, AcName) = ""
            FGrid.TextMatrix(FGrid.Row, AcCode) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, AcName) = RsParty!Name
            FGrid.TextMatrix(FGrid.Row, AcCode) = RsParty!Code
        End If
        FGridAddEditDel
    Case City
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, City) = ""
            FGrid.TextMatrix(FGrid.Row, CityCode) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, City) = RsCity!Name
            FGrid.TextMatrix(FGrid.Row, CityCode) = RsCity!Code
        End If
        FGridAddEditDel
    Case ConPer, Add1, Add2, Pin, FAx, Phone, ActiveYN
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT: FGridAddEditDel
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

Private Sub Grid_Hide()
    If DgHlp.Visible = True Then DgHlp.Visible = False
    If DgCity.Visible = True Then DgCity.Visible = False
    If DgHlp.Visible = True Then DgHlp.Visible = False
    If DGFinGrp.Visible = True Then DGFinGrp.Visible = False
End Sub

Private Sub RemoveTxtNull()
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).TEXT = IIf(IsNull(txt(I).TEXT), "", txt(I).TEXT)
Next I
End Sub

Private Sub FGridAddEditDel()
If FGrid.TextMatrix(FGrid.Row, AddEdit) <> "" Then FGrid.TextMatrix(FGrid.Row, AddEdit) = "E"
End Sub

Private Sub SetMaxLength()
Select Case FGrid.Col   'Index
       Case FinCode2
            TxtGrid(0).MaxLength = 6
       Case ConPer
            TxtGrid(0).MaxLength = 50
       Case Add1, Add2
            TxtGrid(0).MaxLength = 40
       Case Pin
            TxtGrid(0).MaxLength = 6
       Case Phone
            TxtGrid(0).MaxLength = 20
       Case FAx
            TxtGrid(0).MaxLength = 15
        Case Else
            TxtGrid(0).MaxLength = 0
End Select

End Sub
