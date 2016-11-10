VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehRect 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Receipt Entry"
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.OptionButton OptTrfRect 
      BackColor       =   &H00BAD3C9&
      Caption         =   "Transfer Reciept"
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
      Left            =   3840
      TabIndex        =   19
      Top             =   1065
      Width           =   2115
   End
   Begin VB.OptionButton OptPurRect 
      BackColor       =   &H00BAD3C9&
      Caption         =   "Purchase Reciept"
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
      Left            =   1620
      TabIndex        =   18
      Top             =   1065
      Width           =   2115
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2865
      Left            =   -3030
      Negotiate       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6945
      Visible         =   0   'False
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5054
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
      Caption         =   "Model Help"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Model Code"
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
         DataField       =   "ModelGroup"
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
      BeginProperty Column02 
         DataField       =   "Colour"
         Caption         =   "Colour"
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
         DataField       =   "Name"
         Caption         =   "Model Name"
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
      BeginProperty Column04 
         DataField       =   "Chas_Type"
         Caption         =   "Chassis Type"
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
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4500.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1349.858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2130
      Left            =   -630
      Negotiate       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   3757
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
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   210
      Index           =   1
      Left            =   4215
      MaxLength       =   8
      TabIndex        =   2
      Top             =   645
      Width           =   1680
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
      Left            =   5670
      TabIndex        =   4
      Top             =   3510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   210
      Index           =   2
      Left            =   8010
      MaxLength       =   4
      TabIndex        =   3
      Top             =   645
      Width           =   675
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   210
      Index           =   0
      Left            =   1095
      MaxLength       =   12
      TabIndex        =   1
      Top             =   645
      Width           =   1665
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3270
      Left            =   15
      TabIndex        =   5
      Top             =   1590
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   5768
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   14
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15196124
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   32896
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   $"frmVehRect.frx":0000
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
      _Band(0).Cols   =   14
   End
   Begin MSDataGridLib.DataGrid DGCol 
      Height          =   2130
      Left            =   75
      Negotiate       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   3757
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
      Caption         =   "Colour Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Colors"
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
            ColumnWidth     =   4380.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGod 
      Height          =   2130
      Left            =   2610
      Negotiate       =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7140
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3757
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
      Caption         =   "Godown Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Godown Name"
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
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   17
      Top             =   1050
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipt Type"
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
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   16
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
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
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   12
      Top             =   645
      Width           =   840
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   195
      Index           =   92
      Left            =   3975
      TabIndex        =   11
      Top             =   645
      Width           =   75
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   195
      Index           =   8
      Left            =   7770
      TabIndex        =   9
      Top             =   645
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RSO Purchase Y/N"
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
      Height          =   195
      Index           =   27
      Left            =   6045
      TabIndex        =   8
      Top             =   645
      Width           =   1575
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   195
      Index           =   91
      Left            =   855
      TabIndex        =   7
      Top             =   645
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   6
      Top             =   645
      Width           =   405
   End
End
Attribute VB_Name = "frmVehRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGod As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim RsMod  As ADODB.Recordset
Dim RsCol  As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim GridKey As Integer

Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const VDate As Byte = 0
Private Const SerialNo As Byte = 1
Private Const RSO_WORK As Byte = 2
' Col Declaration

'Chassis_RctDocNo,RSO_WORK,MODEL ,TAX_YN,Colour_Code,ChassisNo, EngineNo,SDM_STM_NO,Srv_BookNo,Godown,PartyCode,Remarks

Private Const Model As Byte = 1
Private Const Taxable As Byte = 2
Private Const Colours As Byte = 3
Private Const ChassisNo As Byte = 4
Private Const EngineNo As Byte = 5
Private Const PurRate As Byte = 6
Private Const InDate As Byte = 7
Private Const SDM_STM_NO As Byte = 8
Private Const Srv_BookNo  As Byte = 9
Private Const Godown As Byte = 10
Private Const PartyName  As Byte = 11
Private Const Remarks  As Byte = 12
Private Const ColCode  As Byte = 13
Private Const PartyCode  As Byte = 14
Private Const God As Byte = 15
Private Const ChassisNoOld As Byte = 16


Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub DGCol_Click()
    DGCol.Visible = False
    If RsCol.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsCol!Name
         FGrid.TextMatrix(FGrid.Row, Colours) = RsCol!Name
         FGrid.TextMatrix(FGrid.Row, ColCode) = RsCol!Code
    End If
   TxtGrid(0).SetFocus
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
'Dim i As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
'    Master.Open "SELECT DISTINCT veh_stock.Chassis_RctDocNo AS searchcode, veh_stock.Chassis_RctDocNo, veh_stock.Chassis_RctDate, veh_stock.RSO_WORK FROM veh_stock where Chassis_RctSiteCode  = '" & PubSiteCode & "' and Chassis_RctDivCode  = '" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
   Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  left(veh_stock.pur_sitecode,1) ='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If

    If PubMoveRecYn Then
        Master.Open "SELECT DISTINCT veh_stock.Chassis_RctDocNo AS searchcode, veh_stock.Chassis_RctDocNo, veh_stock.Chassis_RctDate, veh_stock.RSO_WORK,veh_stock.RectType FROM veh_stock where Chassis_RctDivCode  = '" & PubDivCode & "' " & sitecond & " ", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "SELECT DISTINCT Top 1 veh_stock.Chassis_RctDocNo AS searchcode, veh_stock.Chassis_RctDocNo, veh_stock.Chassis_RctDate, veh_stock.RSO_WORK,veh_stock.RectType FROM veh_stock where Chassis_RctDivCode  = '" & PubDivCode & "' " & sitecond & " ", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    Set RsCol = New ADODB.Recordset
    RsCol.CursorLocation = adUseClient
    RsCol.Open "select Col_code as code,col_Desc  as name from colmast order by col_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGCol.DataSource = RsCol
    
    Set rsGod = New ADODB.Recordset
    rsGod.CursorLocation = adUseClient
    rsGod.Open "select god_code as code,god_name as name from godown where appli_for = 1 order by god_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGod.DataSource = rsGod
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME from SubGroup Where firmCode = '" & PubFirmCode & "' and Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME from SubGroup " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        "order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    If PubSiebelActiveYn = 1 Then
        Set RsMod = New ADODB.Recordset
        RsMod.CursorLocation = adUseClient
        RsMod.Open "select Model as code,ModelGrp_Name as ModelGroup,Col_desc  as Colour,Model_Desc as NAME, Chas_Type from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code where Div_Code='" & PubDivCode & "' order by Model", GCn, adOpenDynamic, adLockOptimistic
        Set DGMod.DataSource = RsMod
    Else
        Set RsMod = New ADODB.Recordset
        RsMod.CursorLocation = adUseClient
        RsMod.Open "select Model as code,Model_Desc as NAME, Chas_Type from model where (Div_Code='" & PubDivCode & "' or Div_Code='') order by model", GCn, adOpenDynamic, adLockOptimistic
        Set DGMod.DataSource = RsMod
        DGMod.Columns(1).width = 0
        DGMod.Columns(2).width = 0
    End If
   
    Call MoveRec
    
    Ini_Grid
    Disp_Text SETS("INI", Me, Master)
Exit Sub

ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsGod = Nothing
Set RsParty = Nothing
Set RsMod = Nothing
Set RsCol = Nothing
Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Txt(SerialNo).TEXT = IIf(GCn.Execute("select count(*) from veh_stock where Chassis_RctSiteCode  = '" & PubSiteCode & "' and Chassis_RctDivCode  = '" & PubDivCode & "'").Fields(0).Value > 0, GCn.Execute("select MAX(Chassis_RctDocNo) from veh_stock where Chassis_RctSiteCode  = '" & PubSiteCode & "' and Chassis_RctDivCode  = '" & PubDivCode & "'").Fields(0).Value + 1, 1)
    Txt(SerialNo).Enabled = False
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    Txt(VDate).SetFocus
    OptPurRect.Value = True
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer

If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
'**********modi shekhar
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ChassisNo) <> "" Then
            If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "' and Pur_DocId <> ''").Fields(0).Value > 0 Then
                MsgBox "Purchase bill has been made " & vbCrLf & "you can delete it from Purchase Bill", vbInformation, "Deletion Denied": FGrid.SetFocus: Exit Sub
            End If
        End If
    Next
'*****end modi
    GCn.BeginTrans
    For I = 1 To FGrid.Rows - 1
        If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(I, ChassisNo) & "'").Fields(0).Value = 0 Then
            GCn.Execute ("delete from Veh_stock where Chassis_RctSiteCode  = '" & PubSiteCode & "' and Chassis_RctDivCode  = '" & PubDivCode & "' and Chassis_RctDocNo = " & Master!Chassis_RctDocNo & " and Chassisno = '" & FGrid.TextMatrix(I, ChassisNo) & "'")
'            GCn.Execute "delete from hiscard Where chassis='" & FGrid.TextMatrix(i, ChassisNo) & "'"
        Else
            MsgBox "Chassis No " & FGrid.TextMatrix(I, ChassisNo) & " is Sold" & vbCrLf & "Deletion Denied", vbInformation, "Deletion Denied"
        End If
    Next
    GCn.CommitTrans
    Master.Requery
    Call MoveRec
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
 Dim I As Integer
    Disp_Text SETS("EDIT", Me, Master)
    FGrid.AddItem FGrid.Rows
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
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
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    rsGod.Requery
    RsParty.Requery
    RsCol.Requery
    RsMod.Requery
End Sub
Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim CardNo As String
   On Error GoTo errlbl

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(Txt(VDate), "Receipt Date") = False Then Exit Sub
    If IsValid(Txt(SerialNo), "Serial Number") = False Then Exit Sub
    If FGrid.Rows = 2 And FGrid.TextMatrix(1, Model) = "" Then MsgBox "Fill Transaction Data", vbInformation, "Required data": FGrid.Row = 1: FGrid.Col = Model: FGrid.SetFocus: Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Model) <> "" Then
            If FGrid.TextMatrix(I, Taxable) = "" Then MsgBox "Fill Taxable Yes/No in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Taxable: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, ChassisNo) = "" Then MsgBox "Fill Chassis No  in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = ChassisNo: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, EngineNo) = "" Then MsgBox "Fill Engine No  in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = EngineNo: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Colours) = "" Then MsgBox "Fill Colour  in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Colours: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Srv_BookNo) = "" Then MsgBox "Fill Service Book No. in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Srv_BookNo: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, PurRate)) <= 0 Then MsgBox "Purchase Rate in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = PurRate: FGrid.SetFocus: Exit Sub
            
            'Modi LPS 25-08-2003 with Vikash for Stock in Transit Report
            If FGrid.TextMatrix(I, InDate) = "" Then
                If MsgBox("In Date not feeded in Row No. " & I & vbCrLf & "Save Data ?", vbYesNo + vbCritical + vbDefaultButton2, "Validation") = vbNo Then
                    FGrid.Row = I: FGrid.Col = InDate: FGrid.SetFocus:  Exit Sub
                End If
            End If
            'eof modi
            If FGrid.TextMatrix(I, Godown) = "" Then MsgBox "Fill Godown in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Godown: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, PartyName) = "" Then MsgBox "Fill Party in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = PartyName: FGrid.SetFocus: Exit Sub
        End If
    Next
'Chassis_RctDocNo,MODEL ,TAX_YN,Colour_Code,ChassisNo, EngineNo,SDM_STM_NO,Srv_BookNo,Godown,PartyCode,Remarks
'**********modi shekhar
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ChassisNo) <> "" Then
            If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "' and Pur_DocId <> ''").Fields(0).Value > 0 Then
                MsgBox "Purchase bill has been made " & vbCrLf & "you can edit/delete it from Purchase Bill", vbInformation, "Editing Denied": FGrid.SetFocus: Exit Sub
            End If
            If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "' and (Sal_DocId = '' or len(Sal_DocId) = 0)").Fields(0).Value > 0 Then
                MsgBox "Vehicle is already in stock " & vbCrLf & " Saving aborted !", vbInformation, "Information": FGrid.SetFocus: Exit Sub
            End If
        End If
    Next
   
  '*****end modi
    GCn.BeginTrans
    mTrans = True
    If OptPurRect.Value = True Then
        GCn.Execute ("delete from Veh_stock where Chassis_RctDocNo=" & Val(Txt(SerialNo).TEXT) & "")
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Model) <> "" Then
                GCn.Execute ("insert into Veh_stock(Chassis_RctDivCode,Chassis_RctSiteCode,Chassis_RctDate,Chassis_RctDocNo,Chassis_RctSrl_No,RSO_WORK,MODEL ,TAX_YN,Colour_Code,ChassisNo, EngineNo,SDM_STM_NO,Srv_BookNo,Godown,PartyCode,Remarks, " & _
                " indate,Rate,U_Name, U_EntDt, U_AE,RectType ) " & _
                " values('" & PubDivCode & "','" & PubSiteCode & "'," & ConvertDate(Txt(VDate)) & "," & Val(Txt(SerialNo).TEXT) & "," & I & ", " & _
                " " & IIf(Txt(RSO_WORK) = "Yes", 1, 0) & ",'" & FGrid.TextMatrix(I, Model) & "'," & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ", '" & FGrid.TextMatrix(I, ColCode) & "', " & _
                " '" & FGrid.TextMatrix(I, ChassisNo) & "','" & FGrid.TextMatrix(I, EngineNo) & "','" & FGrid.TextMatrix(I, SDM_STM_NO) & "','" & FGrid.TextMatrix(I, Srv_BookNo) & "','" & FGrid.TextMatrix(I, God) & "','" & FGrid.TextMatrix(I, PartyCode) & "','" & FGrid.TextMatrix(I, Remarks) & "'," & ConvertDate(FGrid.TextMatrix(I, InDate)) & "," & Val(FGrid.TextMatrix(I, PurRate)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','P')")
            End If
        Next
    Else
        GCn.Execute ("Delete from Veh_stock where Chassis_RctDocNo=" & Val(Txt(SerialNo).TEXT) & "")
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Model) <> "" Then
                GCn.Execute ("insert into Veh_stock(Chassis_RctDivCode,Chassis_RctSiteCode,Chassis_RctDate,Chassis_RctDocNo,Chassis_RctSrl_No,RSO_WORK,MODEL ,TAX_YN,Colour_Code,ChassisNo, EngineNo,SDM_STM_NO,Srv_BookNo,Godown,PartyCode,Remarks, " & _
                " indate,Rate,U_Name, U_EntDt, U_AE,Trf_Date,TrfParty,RectType ) " & _
                " values('" & PubDivCode & "','" & PubSiteCode & "'," & ConvertDate(Txt(VDate)) & "," & Val(Txt(SerialNo).TEXT) & "," & I & ", " & _
                " " & IIf(Txt(RSO_WORK) = "Yes", 1, 0) & ",'" & FGrid.TextMatrix(I, Model) & "'," & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ", '" & FGrid.TextMatrix(I, ColCode) & "', " & _
                " '" & FGrid.TextMatrix(I, ChassisNo) & "','" & FGrid.TextMatrix(I, EngineNo) & "','" & FGrid.TextMatrix(I, SDM_STM_NO) & "','" & FGrid.TextMatrix(I, Srv_BookNo) & "','" & FGrid.TextMatrix(I, God) & "','" & FGrid.TextMatrix(I, PartyCode) & "','" & FGrid.TextMatrix(I, Remarks) & "'," & ConvertDate(FGrid.TextMatrix(I, InDate)) & "," & Val(FGrid.TextMatrix(I, PurRate)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A',Null,'" & FGrid.TextMatrix(I, PartyCode) & "','T')")
            End If
        Next
    End If
'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
'    For i = 1 To FGrid.Rows - 1
'        If FGrid.TextMatrix(i, ChassisNoOld) = "" Then
'
'            If FGrid.TextMatrix(i, Model) <> "" Then
'                Dim urs As Recordset
'                Set urs = GCn.Execute("select max(val(mid(cardno,2,6))),max(carddate) from hiscard")
'                CardNo = PubSiteCode + Right("000000" & IIf(IsNull((urs.Fields(0).Value)), 1, (urs.Fields(0).Value) + 1), 6)
'                Set urs = Nothing
''                If GCn.Execute("select max(val(Mid(cardno,2,len(cardno)-1))) from hiscard").RecordCount  > 0 Then
''                    CardNo = PubSiteCode + str(GCn.Execute("select max(val(Mid(cardno,2,len(cardno)-1)))+1 from hiscard").Fields(0))
''                End If
'                GCn.Execute "delete from hiscard Where chassis='" & FGrid.TextMatrix(i, ChassisNo) & "'"
'                GCn.Execute "insert into hiscard(cardno,Site_Code,Div_Code,carddate,model,chassis,engine,U_Name, U_EntDt, U_AE) " & _
'                "values('" & CardNo & "','" & PubSiteCode & "','" & PubDivCode & "'," & ConvertDate(FGrid.TextMatrix(i, InDate)) & ",'" & FGrid.TextMatrix(i, Model) & "','" & FGrid.TextMatrix(i, ChassisNo) & "','" & FGrid.TextMatrix(i, EngineNo) & "', " & _
'                "'" & pubUName & "',#" & PubServerDate & "#,'A')"
'            End If
'        Else
'            If FGrid.TextMatrix(i, Model) <> "" Then
'                GCn.Execute "update hiscard set carddate=" & ConvertDate(FGrid.TextMatrix(i, InDate)) & "," & _
'                "model='" & FGrid.TextMatrix(i, Model) & "',chassis='" & FGrid.TextMatrix(i, ChassisNo) & "', " & _
'                "engine='" & FGrid.TextMatrix(i, EngineNo) & "',U_Name='" & pubUName & "', U_EntDt=#" & PubServerDate & "#, U_AE='E' " & _
'                "Where chassis='" & FGrid.TextMatrix(i, ChassisNoOld) & "'"
'            End If
'        End If
'    Next
GCn.CommitTrans
mTrans = False

    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("SELECT DISTINCT veh_stock.Chassis_RctDocNo AS searchcode, veh_stock.Chassis_RctDocNo, veh_stock.Chassis_RctDate, veh_stock.RSO_WORK,veh_stock.RectType FROM veh_stock where Chassis_RctDivCode  = '" & PubDivCode & "' and veh_stock.Chassis_RctDocNo = " & Val(Txt(SerialNo)) & "  ")
    End If
    RsParty.Requery
    Master.FIND "Chassis_RctDocNo = " & Val(Txt(SerialNo)) & ""
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
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    Dim sitecond As String
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
       If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and left(veh_stock.pur_sitecode,1) ='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "SELECT DISTINCT Chassis_RctDocNo AS SearchCode, Chassis_RctDocNo, Chassis_RctDate, ChassisNo as Chassis_No,EngineNo as Engine_No, RSO_WORK FROM veh_stock where Chassis_RctDivCode  = '" & PubDivCode & "' " & sitecond & " order by ChassisNo"
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
        Master.FIND "searchcode=" & MyValue
    Else
        Set Master = GCn.Execute("SELECT DISTINCT veh_stock.Chassis_RctDocNo AS searchcode, veh_stock.Chassis_RctDocNo, veh_stock.Chassis_RctDate, veh_stock.RSO_WORK,veh_stock.RectType FROM veh_stock where Chassis_RctDivCode  = '" & PubDivCode & "' and veh_stock.Chassis_RctDocNo = " & MyValue & "  ")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub


Private Sub Txt_GotFocus(Index As Integer)
TxtGrid(0).Visible = False
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
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
If DGParty.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> VDate Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> RSO_WORK Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case Index
    Case SerialNo
        Call NumPress(Txt(Index), KeyAscii, 6, 0)
    Case RSO_WORK
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Select Case Index
    Case VDate
        If Len(Trim(Txt(VDate).TEXT)) = 0 Then
             Txt(VDate).TEXT = PubLoginDate
        Else
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
End Select
Set Rst = Nothing
End Sub

Private Sub DGGod_Click()
    DGGod.Visible = False
    If rsGod.RecordCount > 0 Then
        TxtGrid(0).TEXT = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
    End If
   TxtGrid(0).SetFocus
End Sub

Private Sub DGMod_Click()
DGMod.Visible = False
If RsMod.RecordCount > 0 Then
    TxtGrid(0).TEXT = RsMod!Code
    FGrid.TextMatrix(FGrid.Row, Model) = RsMod!Code
End If
TxtGrid(0).SetFocus
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsParty!Name
        FGrid.TextMatrix(FGrid.Row, PartyCode) = RsParty!Code
    End If
    DGParty.Visible = False
    TxtGrid(0).SetFocus
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
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
        Case Model
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case Godown
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, God) = ""
        Case Colours
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, ColCode) = ""
        Case PartyName
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, PartyCode) = ""
        Case ChassisNo, EngineNo, SDM_STM_NO, Srv_BookNo, Remarks, PurRate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Taxable
            FGrid.Col = FGrid.Col + 1
        Case Model, PartyName, Colours, ChassisNo, EngineNo, SDM_STM_NO, Srv_BookNo, Godown, Remarks, InDate
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case Model, Taxable, PartyName, Colours, ChassisNo, EngineNo, SDM_STM_NO, Srv_BookNo, Godown, Remarks, InDate
        Call GridDblClick(Me, FGrid, TxtGrid, 0)
End Select
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.Rows > 2 Then
                '**********modi shekhar
                If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "' and Pur_DocId <> ''").Fields(0).Value > 0 Then
                    MsgBox "Purchase bill has been made " & vbCrLf & "you can edit it from Purchase Bill", vbInformation, "Editing Denied": FGrid.SetFocus: Exit Sub
                End If
                '**********end modi
                If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                  MsgBox "Chassis Sold" & vbCrLf & "Deletion Denied", vbInformation, "Deletion Denied": FGrid.SetFocus: Exit Sub
                End If
                FGrid.RemoveItem (FGrid.Row)
            Else
            '**********modi shekhar
                If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "' and Pur_DocId <> ''").Fields(0).Value > 0 Then
                    MsgBox "Purchase bill has been made " & vbCrLf & "you can edit it from Purchase Bill", vbInformation, "Editing Denied": FGrid.SetFocus: Exit Sub
                End If
            '****end modi
                If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                    MsgBox "Chassis Sold" & vbCrLf & "Deletion Denied", vbInformation, "Deletion Denied": FGrid.SetFocus: Exit Sub
                End If
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
Exit Sub
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
Select Case FGrid.Col
    Case Model, PartyName, Colours, ChassisNo, EngineNo, SDM_STM_NO, Srv_BookNo, Godown, Remarks, InDate
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case PurRate
        Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
    Case Taxable
        If UCase(Chr(KeyAscii)) = "N" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No"
        ElseIf UCase(Chr(KeyAscii)) = "Y" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
        Else
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End If
        KeyAscii = 0
        FGrid.Col = FGrid.Col + 1
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
Next I
End Sub

Private Sub MoveRec()
Dim Rs As ADODB.Recordset, I As Integer
On Error GoTo error1
TopCtrl1.tPrn = False
If Master.RecordCount > 0 Then
    Txt(SerialNo).TEXT = Master!Chassis_RctDocNo
    Txt(VDate).TEXT = Master!Chassis_RctDate
    Txt(RSO_WORK) = IIf(Master!RSO_WORK = 1, "Yes", "No")
    If XNull(Master!RectType) = "P" Then
        OptPurRect.Value = True
    ElseIf XNull(Master!RectType) = "T" Then
        OptTrfRect.Value = True
    Else
        OptTrfRect.Value = False
        OptPurRect.Value = False
    End If
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Veh_Stock.*, ColMast.Col_Desc, SubGroup.Name AS party, Godown.God_Name " & _
    "FROM ((Veh_Stock LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code) LEFT JOIN SubGroup ON Veh_Stock.PartyCode = SubGroup.Subcode) LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code " & _
    "where Veh_Stock.Chassis_RctDivCode  = '" & PubDivCode & "' and Veh_Stock.Chassis_RctDocNo = " & Master!Chassis_RctDocNo & "")
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        I = 1
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = Rs!Chassis_RctSrl_No
                .TextMatrix(I, Model) = Rs!Model
                .TextMatrix(I, Taxable) = IIf(Rs!Tax_YN = 0, "No", "Yes")
                .TextMatrix(I, Colours) = XNull(Rs!Col_Desc)
                .TextMatrix(I, ChassisNo) = Rs!ChassisNo
                .TextMatrix(I, EngineNo) = Rs!EngineNo
                .TextMatrix(I, InDate) = XNull(Rs!InDate)
                .TextMatrix(I, SDM_STM_NO) = Rs!SDM_STM_NO
                .TextMatrix(I, Srv_BookNo) = Rs!Srv_BookNo
                .TextMatrix(I, Godown) = XNull(Rs!God_Name)
                .TextMatrix(I, PartyName) = XNull(Rs!Party)
                .TextMatrix(I, Remarks) = Rs!Remarks
                .TextMatrix(I, ColCode) = Rs!Colour_Code
                .TextMatrix(I, PartyCode) = Rs!PartyCode
                .TextMatrix(I, God) = Rs!Godown
                .TextMatrix(I, ChassisNoOld) = Rs!ChassisNo
                .TextMatrix(I, PurRate) = Format(VNull(Rs!Rate), "0.00")
            End With
            Rs.MoveNext
            I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Set Rs = Nothing
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
'Dim i As Byte
'SrNo.0|Model1|RSO/Work2|Tax3|Quantiy4|Rate5|Tax%6|TaxAmt7|Surch%8|SurchAmt9|Amount10

'Model 1| Taxable 2|Colour 3| Chassis No 4|Engine No 5|SDM/STM 6|Service Book No 7|Chassis Godown 8|Received from 8 |Remark 10
    With FGrid
        .Cols = 17
        .left = Me.left '+45
        .width = Me.width - 90
        .top = 1590
        .RowHeightMin = PubGridRowHeight
        
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, Model) = "Model"
        .ColAlignment(Model) = flexAlignLeftCenter
        .ColWidth(Model) = 1500

        .TextMatrix(0, Taxable) = "Tax"
        .ColAlignment(Taxable) = flexAlignLeftCenter
        .ColWidth(Taxable) = 360

        .TextMatrix(0, Colours) = "Colours"
        .ColAlignment(Colours) = flexAlignLeftCenter
        .ColWidth(Colours) = 1200

        .TextMatrix(0, ChassisNo) = "Chassis No"
        .ColAlignment(ChassisNo) = flexAlignLeftCenter
        .ColWidth(ChassisNo) = 1575

        .TextMatrix(0, EngineNo) = "Engine No"
        .ColAlignment(EngineNo) = flexAlignLeftCenter
        .ColWidth(EngineNo) = 1590
        
        .TextMatrix(0, InDate) = "InDate"
        .ColAlignment(InDate) = flexAlignLeftCenter
        .ColWidth(InDate) = 1080

        .TextMatrix(0, SDM_STM_NO) = "SDM/STM No"
        .ColAlignment(SDM_STM_NO) = flexAlignLeftCenter
        .ColWidth(SDM_STM_NO) = 1125

        .TextMatrix(0, Srv_BookNo) = "SrvBookNo"
        .ColAlignment(Srv_BookNo) = flexAlignLeftCenter
        .ColWidth(Srv_BookNo) = 930

        .TextMatrix(0, Godown) = "Godown"
        .ColAlignment(Godown) = flexAlignLeftCenter
        .ColWidth(Godown) = 1305

        .TextMatrix(0, PartyName) = "Party"
        .ColAlignment(PartyName) = flexAlignLeftCenter
        .ColWidth(PartyName) = 1770

        .TextMatrix(0, Remarks) = "Remarks"
        .ColAlignment(Remarks) = flexAlignLeftCenter
        .ColWidth(Remarks) = 1200
        
        .TextMatrix(0, PurRate) = "Pur.Rate"
        .ColAlignment(PurRate) = flexAlignLeftCenter
        .ColWidth(PurRate) = 1500
        
        .ColWidth(ColCode) = 0
        .ColWidth(PartyCode) = 0
        .ColWidth(God) = 0
        .ColWidth(ChassisNoOld) = 0
End With
BackColorSelLeave = FGrid.BackColorSel
ForeColorSelEnter = FGrid.ForeColorSel
DGParty.left = Me.left + 45: DGParty.top = FGrid.top + FGrid.height + 50: DGParty.height = Me.height - (DGParty.top + mBotScale)
DGMod.left = Me.left + 45: DGMod.top = FGrid.top + FGrid.height + 50: DGMod.height = Me.height - (DGMod.top + mBotScale)
DGCol.left = Me.left + 45: DGCol.top = FGrid.top + FGrid.height + 50: DGCol.height = Me.height - (DGCol.top + mBotScale)
DGGod.left = Me.left + 45: DGGod.top = FGrid.top + FGrid.height + 50: DGGod.height = Me.height - (DGGod.top + mBotScale)

End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    OptPurRect.Enabled = Enb
    OptTrfRect.Enabled = Enb
    Txt(I).ForeColor = CtrlFColOrg
Next
If TopCtrl1.TopText2 = "Edit" Then
    Txt(VDate).Enabled = False
    Txt(SerialNo).Enabled = False
End If
txtDisabled_Color Me
TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGMod.Visible = True Then DGMod.Visible = False
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
         Case Model
            TxtGrid(0).MaxLength = 15
            If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Model) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Model) <> RsMod!Code Then
                RsMod.MoveFirst
                RsMod.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, Model) & "'"
            End If
         Case Colours
            TxtGrid(0).MaxLength = 15
            If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Colours) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Colours) <> RsCol!Name Then
                RsCol.MoveFirst
                RsCol.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Colours) & "'"
            End If
        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Godown) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Godown) <> rsGod!Name Then
                rsGod.MoveFirst
                rsGod.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Godown) & "'"
            End If
            TxtGrid(0).MaxLength = 4
       
        Case PartyName
            TxtGrid(0).MaxLength = 40
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or FGrid.TextMatrix(FGrid.Row, PartyName) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, PartyName) <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, PartyName) & "'"
            End If
       Case InDate
            TxtGrid(0).MaxLength = 12
       Case ChassisNo
            TxtGrid(0).MaxLength = 20
       Case EngineNo
            TxtGrid(0).MaxLength = 25
       Case SDM_STM_NO
            TxtGrid(0).MaxLength = 15
       Case Srv_BookNo
            TxtGrid(0).MaxLength = 10
       Case Remarks
            TxtGrid(0).MaxLength = 40
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TxtGrid(0) = TxtGrid(0).Tag: Exit Sub
            Select Case FGrid.Col
                Case Model    '1
                    DGridTxtKeyDown DGMod, TxtGrid, Index, RsMod, KeyCode, True, 0, frmModel, "frmModel"
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 10
                        End If
                    End If
                Case PartyName
                    DGridTxtKeyDown DGParty, TxtGrid, Index, RsParty, KeyCode, True, 1, frmSubGroup, "frmSubGroup"
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                           GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 10
                        End If
                    End If
                Case Colours
                    DGridTxtKeyDown DGCol, TxtGrid, 0, RsCol, KeyCode, True, 1, frmColor, "frmColor"
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                           GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 10
                        End If
                    End If
                Case Godown
                    DGridTxtKeyDown DGGod, TxtGrid, 0, rsGod, KeyCode, True, 1, frmGodown, "frmGodown"
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                            GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, 18
                        End If
                    End If
                Case Taxable, ChassisNo, EngineNo, SDM_STM_NO, Srv_BookNo, Godown, Remarks, InDate
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 10
                        End If
                    End If
                Case PurRate
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 10
                        End If
                    End If
                                    
            End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case Model
        If DGMod.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsMod, KeyAscii, "CODE"
    Case PartyName
        If DGParty.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsParty, KeyAscii, "Name"
    Case Godown
        If DGGod.Visible = True Then DGridTxtKeyPress TxtGrid, Index, rsGod, KeyAscii, "Name"
    Case Colours
        If DGCol.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsCol, KeyAscii, "Name"
    Case Taxable
        If UCase(Chr(KeyAscii)) = "Y" Then
            TxtGrid(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            TxtGrid(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            TxtGrid(Index) = ""
        End If
        KeyAscii = 0
    Case PurRate
        NumPress TxtGrid(0), KeyAscii, 8, 2
End Select
End Sub


Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case Model
        If KeyCode <> 13 And DGMod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsMod, KeyCode, "code", True
    Case PartyName
        If KeyCode <> 13 And DGParty.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsParty, KeyCode, "name", True
    Case Godown
        If KeyCode <> 13 And DGGod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, 0, rsGod, KeyCode, "Name", True
    Case Colours
        If KeyCode <> 13 And DGCol.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsCol, KeyCode, "name", True
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
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
Dim j As Integer
Select Case FGrid.Col
        Case Model
            If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Model) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Model) = RsMod!Code
                FGrid.TextMatrix(FGrid.Row, ChassisNo) = RsMod!Chas_Type
            End If
            If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
        Case ChassisNo
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = UCase(TxtGrid(0).TEXT)
            '**********modi shekhar
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
                If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "' and Pur_DocId <> ''").Fields(0).Value > 0 Then
                    MsgBox "Purchase bill has been made " & vbCrLf & "you can edit/delete it from Purchase Bill", vbInformation, "Editing Denied": FGrid.SetFocus: TxtGridLeave = False: Exit Function
                End If
            End If
'**********modi end
            'MODISHEKHAR
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
                If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                  MsgBox "Chassis Sold" & vbCrLf & "Editing Denied", vbInformation, "Editing Denied": FGrid.SetFocus: TxtGridLeave = False
                  FGrid.TextMatrix(FGrid.Row, ChassisNo) = ""
                  Exit Function
                End If
            End If
            'END MODI
            
            If ChkDul_Chassis = True Then TxtGridLeave = False: Exit Function
        Case EngineNo
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = UCase(TxtGrid(0).TEXT)
        Case InDate
            TxtGrid(0).TEXT = RetDate(TxtGrid(0))
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
'        Case Taxable
'            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).Text
        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Godown) = ""
                FGrid.TextMatrix(FGrid.Row, God) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
                FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
            End If
        Case PartyName
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, PartyName) = ""
                FGrid.TextMatrix(FGrid.Row, PartyCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, PartyName) = RsParty!Name
                FGrid.TextMatrix(FGrid.Row, PartyCode) = RsParty!Code
            End If
        Case Colours
            If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Colours) = ""
                FGrid.TextMatrix(FGrid.Row, ColCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Colours) = RsCol!Name
                FGrid.TextMatrix(FGrid.Row, ColCode) = RsCol!Code
            End If
        Case SDM_STM_NO, Srv_BookNo, Godown, Remarks
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        Case PurRate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(TxtGrid(0).TEXT, "0.00")
    End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function

Private Function ChkDuplicate() As Boolean
Dim I As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte
    Select Case FGrid.Col
    Case Model
        Col2 = Model
        Col1 = ChassisNo
    Case ChassisNo
        Col1 = Model
        Col2 = ChassisNo
    End Select
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
            TxtGrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function
Private Function ChkDul_Chassis() As Boolean
Dim I As Integer
If OptPurRect.Value = True Then
    If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
          MsgBox "Purchase Reciept of This Chassis is Already Feeded. " & vbCrLf & " Entery aborted !", vbInformation, "Information": FGrid.SetFocus
          FGrid.TextMatrix(FGrid.Row, ChassisNo) = ""
          ChkDul_Chassis = True
          Exit Function
    End If
ElseIf OptTrfRect.Value = True Then
    If GCn.Execute("select count(*) from Veh_stock where ChassisNo='" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "' and (Sal_DocId = '' or len(Sal_DocId) = 0)").Fields(0).Value > 0 Then
          MsgBox "Vehicle is already in stock " & vbCrLf & " Entry aborted !", vbInformation, "Information": FGrid.SetFocus
          FGrid.TextMatrix(FGrid.Row, ChassisNo) = ""
          ChkDul_Chassis = True
          Exit Function
    End If
End If
For I = 1 To FGrid.Rows - 1
    If I <> FGrid.Row Then
        If FGrid.TextMatrix(I, ChassisNo) = TxtGrid(0).TEXT Then
            MsgBox "Same Chassis No already taken ", vbInformation, "Duplicate Chassis"
            ChkDul_Chassis = True
            Exit Function
        End If
    End If
Next
ChkDul_Chassis = False
End Function

