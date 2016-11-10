VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmVehStkTrn 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Stock Transfer"
   ClientHeight    =   7005
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11820
   Visible         =   0   'False
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
      Height          =   210
      Index           =   6
      Left            =   2040
      TabIndex        =   17
      Top             =   2670
      Width           =   4605
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
      Height          =   210
      Index           =   5
      Left            =   2055
      TabIndex        =   14
      Top             =   2430
      Width           =   4605
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
      Height          =   210
      Index           =   0
      Left            =   2055
      MaxLength       =   12
      TabIndex        =   5
      Top             =   2190
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
      Height          =   210
      Index           =   1
      Left            =   2055
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1230
      Width           =   4605
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
      Height          =   210
      Index           =   2
      Left            =   2055
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1470
      Width           =   4605
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
      Height          =   210
      Index           =   3
      Left            =   2055
      TabIndex        =   3
      Top             =   1710
      Width           =   4605
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
      Height          =   210
      Index           =   4
      Left            =   2055
      TabIndex        =   4
      Top             =   1950
      Width           =   4605
   End
   Begin MSDataGridLib.DataGrid DgGod 
      Height          =   4515
      Left            =   -1950
      Negotiate       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6390
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
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
         Caption         =   "Godown"
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
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   3315
      Left            =   810
      Negotiate       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5670
      Visible         =   0   'False
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   5847
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
      ColumnCount     =   2
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
            ColumnWidth     =   6075.213
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   3165
      Left            =   -1545
      Negotiate       =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6645
      Visible         =   0   'False
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   5583
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
      Caption         =   "Chassis Help"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Chassis No"
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
         DataField       =   "EngineNo"
         Caption         =   "Engine No"
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
         DataField       =   "VehSerialNo"
         Caption         =   "VehSerialNo"
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
         DataField       =   "Model"
         Caption         =   "Model"
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
         DataField       =   "God_Name"
         Caption         =   "Godown"
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
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2700.284
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgAgBooking 
      Height          =   2565
      Left            =   225
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   4524
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
      Caption         =   "Booking Help"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Booking No."
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
         DataField       =   "PartyName"
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
      BeginProperty Column02 
         DataField       =   "CityName"
         Caption         =   "City"
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
         DataField       =   "Site"
         Caption         =   "Site Name"
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
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3344.882
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   225
      TabIndex        =   18
      Top             =   2670
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Against Booking No."
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
      Height          =   255
      Index           =   1
      Left            =   225
      TabIndex        =   15
      Top             =   2445
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer  Date"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2205
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Name"
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
      Index           =   26
      Left            =   240
      TabIndex        =   9
      Top             =   1245
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
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
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1485
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Godown To "
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
      Index           =   45
      Left            =   240
      TabIndex        =   7
      Top             =   1965
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Godown From"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1725
      Width           =   1185
   End
End
Attribute VB_Name = "FrmVehStkTrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsChassis As ADODB.Recordset
Dim rsGod As ADODB.Recordset
Dim RsMod As ADODB.Recordset
Dim RsAgBooking As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim TrnDocId    As String * 21
Dim VSrNo       As String
Dim VEngNo      As String
Dim TrnSrlNo    As Byte
Private Const Model As Byte = 1
Private Const ChassisNo As Byte = 2
Private Const FromGod As Byte = 3
Private Const ToGod As Byte = 4
Private Const AgBooking As Byte = 5
Private Const TrnDt As Byte = 0

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
Dim mQry As String

Dim I As Byte
TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
    Dim sitecond As String
    
    sitecond = " Where V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = sitecond & " And  " & cMID("Veh_Transfer.Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
        
    If PubMoveRecYn Then
        Master.Open "select (docid + " & cCStr("srl_no") & ") as searchcode,Veh_Transfer.*, Veh_Order.Ord_No from Veh_Transfer Left Join Veh_order On Veh_Transfer.OrdDocID = Veh_Order.OrdDocID " & sitecond & " order by docid", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 (docid + " & cCStr("srl_no") & ") as searchcode,Veh_Transfer.*, Veh_Order.Ord_No from Veh_Transfer Left Join Join Veh_order On Veh_Transfer.OrdDocID = Veh_Order.OrdDocID   " & sitecond & " order by docid", GCn, adOpenDynamic, adLockOptimistic
    End If
   
    Set rsGod = New ADODB.Recordset
    rsGod.CursorLocation = adUseClient
    rsGod.Open "select god_code as code,god_Name as name from godown where Appli_For=1 order by god_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGod.DataSource = rsGod
  
    Set RsMod = New ADODB.Recordset
    RsMod.CursorLocation = adUseClient
    RsMod.Open "select Model as code,Model_Desc as NAME from model order by model", GCn, adOpenDynamic, adLockOptimistic
    Set DGMod.DataSource = RsMod
    
    Set RsChassis = New ADODB.Recordset
    RsChassis.CursorLocation = adUseClient
    'RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code,Veh_Stock.pur_docid,Veh_Stock.pur_srlno, Veh_Stock.EngineNo,Veh_Stock.VehSerialNo, Godown.God_Name,Veh_Stock.Godown " & _
    "FROM Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code  where Veh_Stock.pur_docid<> ''"), GCn, adOpenDynamic, adLockOptimistic
    '*********MODISHEKHAR 23Jan
    Set RsChassis = GCn.Execute("SELECT Veh_Stock.ChassisNo as code,Veh_Stock.pur_docid,Veh_Stock.pur_srlno, Veh_Stock.EngineNo,Veh_Stock.VehSerialNo, Godown.God_Name,Veh_Stock.Godown,Veh_Stock.Model " & _
    "FROM Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code where (Veh_Stock.Sal_DocId =  '' Or Veh_Stock.Sal_DocId is Null)")
    '*********ENDMODI
    Set DgChassis.DataSource = RsChassis
    
    mQry = "SELECT O.OrdDocId AS Code, Convert(VARCHAR,O.Ord_No) AS Name, O.Ord_Date, S.Name AS PartyName, C.CityName, Site.Site_Desc AS Site, IsNull(O.DelCh_DocId ,'') " & _
           "FROM Veh_Order O " & _
           "LEFT JOIN SubGroup S ON O.PartyCode = S.SubCode " & _
           "LEFT JOIN City C ON S.CityCode = C.CityCode " & _
           "LEFT JOIN Site  on Left(O.Ord_SiteCode,1) = Site.Site_Code " & _
           "WHERE IsNull(O.DelCh_DocId ,'')='' "
    Set RsAgBooking = GCn.Execute(mQry)
    Set DgAgBooking.DataSource = RsAgBooking
    
    MoveRec
    Disp_Text SETS("INI", Me, Master)
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsGod = Nothing
Set RsChassis = Nothing
Set RsMod = Nothing
Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(TrnDt) = PubLoginDate
    txt(ChassisNo).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                GCn.Execute ("delete from Veh_Transfer where  docid = '" & TrnDocId & "' and  Srl_No = " & TrnSrlNo & "")
                GCn.CommitTrans
                Master.Requery
                RsChassis.Requery
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
    Disp_Text SETS("EDIT", Me, Master)
    txt(ChassisNo).SetFocus
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
    Dim sitecond As String
    
    sitecond = " Where V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " And " & cMID("Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "select (docid + " & cCStr("srl_no") & ") as searchcode,Veh_Transfer.* from Veh_Transfer " & sitecond & " order by docid"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
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
        Set Master = GCn.Execute("select (docid + " & cCStr("srl_no") & ") as searchcode,Veh_Transfer.* from Veh_Transfer Where (docid + " & cCStr("srl_no") & ") = '" & MyValue & "' order by docid")
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
On Error GoTo ELoop
Dim RstRep As ADODB.Recordset, RstRep1 As ADODB.Recordset
Dim mQry As String, I As Integer, X11


    mQry = "SELECT vt.DocId,vt.V_Date ,vt.MODEL,vt.ChassisNo,vs.EngineNo,vt.GodownFrom " & _
           " ,vt.GodownTo,m.TYRES,m.RIMS,Godown.God_Name AS SiteGodown,ss.Name AS CustomerName,cf.FinName AS FinancerName ,vv.Rate, M.Sale_Rate as ModelMasterSaleRate,m.Model_Desc,vt.Narration FROM ((Veh_Transfer as vt LEFT JOIN Veh_Stock vs ON vt.ChassisNo=vs.ChassisNo)" & _
           " LEFT JOIN Model m ON vt.MODEL=m.MODEL) LEFT JOIN Godown  ON vt.godownto=Godown.God_Code LEFT JOIN Veh_Order vv ON vt.orddocid=vv.OrdDocId LEFT JOIN SubGroup ss ON vv.partycode=ss.subcode LEFT JOIN ContractFinance cf ON vv.FB_Code=cf.FInCode where vt.ChassisNo='" & txt(ChassisNo).TEXT & "'"
    
        
    Set RstRep = GCn.Execute(mQry)
        
       
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    X11 = CreateFieldDefFile(RstRep, PubRepoPath + "\VehStkTrn.ttx", True)
    Set rpt = rdApp.OpenReport(PubRepoPath + "\VehStkTrn.RPT")
    rpt.Database.SetDataSource RstRep
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
            Case UCase("title")
                rpt.FormulaFields(I).TEXT = "'" & "Vehicle Purchase Bill" & "'"
        End Select
    Next
    rpt.ReadRecords
    
    Call Report_View(rpt, Me.CAPTION, 0, True)
    
    Set RstRep = Nothing
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub


End Sub

Private Sub TopCtrl1_eRef()
    RsChassis.Requery
    RsMod.Requery
    rsGod.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim STR As String
On Error GoTo errlbl

    txt(TrnDt).TEXT = RetDate(txt(TrnDt))
    If IsValid(txt(Model), "Model") = False Then Exit Sub
    If IsValid(txt(ChassisNo), "Chassis") = False Then Exit Sub
    If IsValid(txt(FromGod), "Godown From") = False Then Exit Sub
    If IsValid(txt(ToGod), "To Godown") = False Then Exit Sub
    If IsValid(txt(TrnDt), "Transfer Date") = False Then Exit Sub
 
 If txt(FromGod) = txt(ToGod) Then MsgBox "Both Godown Can't Be Same", vbExclamation, "Validation Check": Exit Sub
 
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GCn.Execute ("delete from veh_transfer where  docid  = '" & TrnDocId & "' and srl_no = " & TrnSrlNo & "")
        GCn.Execute ("insert into veh_transfer( docid , srl_no,MODEL,ChassisNo,GodownFrom,GodownTo,V_Date,VehSerialNo,engineno, OrdDocID,U_Name, U_EntDt, U_AE,Narration ) " & _
        " values('" & TrnDocId & "'," & TrnSrlNo & ",'" & txt(Model).TEXT & "','" & txt(ChassisNo).TEXT & "' ," & _
        "'" & txt(FromGod).Tag & "','" & txt(ToGod).Tag & "'," & ConvertDate(txt(TrnDt).TEXT) & " , " & _
        "'" & VSrNo & "','" & VEngNo & "', '" & txt(AgBooking).Tag & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & txt(6).TEXT & "')")
        GCn.Execute "update veh_stock set godown = '" & txt(ToGod).Tag & "' where pur_docid = '" & TrnDocId & "' and pur_srlno = " & TrnSrlNo & ""
    Else
        GCn.Execute "update veh_transfer set MODEL='" & txt(Model).TEXT & "',ChassisNo='" & txt(ChassisNo).TEXT & "',GodownFrom='" & txt(FromGod).Tag & "',GodownTo='" & txt(ToGod).Tag & "' , " & _
        "v_date = " & ConvertDate(txt(TrnDt).TEXT) & " ,VehSerialNo='" & VSrNo & "',engineno='" & VEngNo & "', OrdDocID = '" & txt(AgBooking).Tag & "  ',U_Name ='" & pubUName & "' ,U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E' Where DocID='" & Master!DocID & "' And srl_no = " & Master!Srl_No & " "
        GCn.Execute "update veh_stock set godown = '" & txt(ToGod).Tag & "' where pur_docid = '" & TrnDocId & "' and pur_srlno = " & TrnSrlNo & ""
    End If
GCn.CommitTrans
mTrans = False
    STR = TrnDocId & Trim(TrnSrlNo)
    
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select (docid + " & cCStr("srl_no") & ") as searchcode,Veh_Transfer.* from Veh_Transfer Where (docid + " & cCStr("srl_no") & ") = '" & STR & "' order by docid")
    End If
    RsChassis.Requery
    
    Master.FIND "SearchCode = '" & STR & "'"
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
Ctrl_GetFocus txt(Index)
Select Case Index
    Case ChassisNo
'        RsChassis.Close
'        If Txt(Model) = "" Then MsgBox "Select Model First", vbInformation, "Validation": Txt(Model).SetFocus: Exit Sub
'        Set RsChassis = New ADODB.Recordset
'        RsChassis.CursorLocation = adUseClient
'        RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code,Veh_Stock.pur_docid,Veh_Stock.pur_srlno, Veh_Stock.EngineNo,Veh_Stock.VehSerialNo, Godown.God_Name,Veh_Stock.Godown " & _
'        "FROM Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code  where Veh_Stock.pur_docid<> ''and  Veh_Stock.MODEL  = '" & txt(Model).Text & "'"), GCn, adOpenDynamic, adLockOptimistic
            
        '***********MODISHEKHAR 23Jan
'        Set RsChassis = GCn.Execute("SELECT Veh_Stock.ChassisNo as code,Veh_Stock.pur_docid,Veh_Stock.pur_srlno, Veh_Stock.EngineNo,Veh_Stock.VehSerialNo, Godown.God_Name,Veh_Stock.Godown " & _
        "FROM Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code  where Veh_Stock.Sal_DocId = ''and  Veh_Stock.MODEL  = '" & Txt(Model).Text & "'")
        '**************ENDMODI
'        Set DgChassis.DataSource = RsChassis
    Case FromGod, ToGod
        If Index = FromGod Then DGGod.Tag = 1 Else DGGod.Tag = 2
        If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(FromGod).TEXT <> rsGod!Name Then
            rsGod.MoveFirst
            rsGod.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case AgBooking
        If RsAgBooking.RecordCount = 0 Or (RsAgBooking.EOF = True Or RsAgBooking.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsAgBooking!Name Then
            RsAgBooking.MoveFirst
            RsAgBooking.FIND "name ='" & txt(Index).TEXT & "'"
        End If
        
'    Case Model
'        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or Txt(Model).Text = "" Then Exit Sub
'        If Txt(Model).Text <> RsMod!Code Then
'            RsMod.MoveFirst
'            RsMod.FIND "code ='" & Txt(Model).Text & "'"
'        End If
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
    Case ChassisNo
        DGridTxtKeyDown DgChassis, txt, Index, RsChassis, KeyCode, False, 0
    Case FromGod, ToGod
        DGridTxtKeyDown DGGod, txt, Index, rsGod, KeyCode, False, 1, frmGodown, "frmGodown"
    Case AgBooking
        DGridTxtKeyDown DgAgBooking, txt, Index, RsAgBooking, KeyCode, False, 1, frmVehBook, "frmVehBook"
        
'    Case Model
'        DGridTxtKeyDown DGMod, Txt, Index, RsMod, KeyCode, False, 0, frmModel, "frmModel"
End Select
If DgChassis.Visible = False And DGMod.Visible = False And DgAgBooking.Visible = False And DGGod.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> TrnDt Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = AgBooking Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If Index <> Model Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case FromGod, ToGod
        If DGGod.Visible = True Then DGridTxtKeyPress txt, Index, rsGod, KeyAscii, "name"
    Case AgBooking
        If DgAgBooking.Visible = True Then DGridTxtKeyPress txt, Index, RsAgBooking, KeyAscii, "name"
        
    Case ChassisNo
        If DgChassis.Visible = True Then DGridTxtKeyPress txt, Index, RsChassis, KeyAscii, "code"
'    Case Model
'        If DGMod.Visible = True Then DGridTxtKeyPress Txt, Index, RsMod, KeyAscii, "code"
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
'    Case Model
'        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or Txt(Index).Text = "" Then
'            Txt(Index).Text = ""
'        Else
'            Txt(Index).Text = RsMod!Code
'        End If
    Case FromGod, ToGod
            If rsGod.RecordCount = 0 Or rsGod.EOF = True Or rsGod.BOF = True Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = rsGod!Name
                txt(Index).Tag = rsGod!Code
            End If
            
    Case AgBooking
            If RsAgBooking.RecordCount = 0 Or RsAgBooking.EOF = True Or RsAgBooking.BOF = True Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = RsAgBooking!Name
                txt(Index).Tag = RsAgBooking!Code
            End If
            
    Case ChassisNo
        If txt(1) = Empty And txt(2) = Empty And txt(3) = Empty And txt(4) = Empty Then
        Else
        If IsValid(txt(ChassisNo), "ChassisNo") = False Then Exit Sub
        End If
        If RsChassis.RecordCount = 0 Or (RsChassis.EOF = True Or RsChassis.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Model).TEXT = ""
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Model).TEXT = RsChassis!Model
            txt(Index).TEXT = RsChassis!Code
            TrnDocId = XNull(RsChassis!Pur_DocId)
            TrnSrlNo = VNull(RsChassis!Pur_SrlNo)
            txt(FromGod).Tag = IIf(IsNull(RsChassis!Godown), "", RsChassis!Godown)
            If txt(FromGod).Tag <> "" Then txt(FromGod).TEXT = GCn.Execute("select god_name from godown where god_code = '" & txt(FromGod).Tag & "'").Fields(0).Value
        End If
    Case TrnDt
        txt(Index).TEXT = RetDate(txt(Index))
End Select
End Sub


Private Sub DgChassis_Click()
    DgChassis.Visible = False
    If RsChassis.RecordCount > 0 Then
        txt(ChassisNo).TEXT = RsChassis!Code
    End If
    txt(ChassisNo).SetFocus
End Sub
Private Sub DGGod_Click()
    DGGod.Visible = False
    If DGGod.Tag = 1 Then
        If rsGod.RecordCount > 0 Then
            txt(FromGod).TEXT = rsGod!Name
            txt(FromGod).Tag = rsGod!Code
        txt(FromGod).SetFocus
        End If
    ElseIf DGGod.Tag = 2 Then
        If rsGod.RecordCount > 0 Then
            txt(ToGod).TEXT = rsGod!Name
            txt(ToGod).Tag = rsGod!Code
        txt(ToGod).SetFocus
        End If
    End If
End Sub
Private Sub DGMod_Click()
    DGMod.Visible = False
    If RsMod.RecordCount > 0 Then
        txt(Model).TEXT = RsMod!Code
    End If
    txt(Model).SetFocus
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub

Private Sub MoveRec()
On Error GoTo error1
' TopCtrl1.tPrn = False
If Master.RecordCount > 0 Then
    TrnDocId = Master!DocID
    TrnSrlNo = Master!Srl_No
    VSrNo = IIf(IsNull(Master!VehSerialNo), "", Master!VehSerialNo)
    VEngNo = IIf(IsNull(Master!EngineNo), "", Master!EngineNo)
    txt(TrnDt) = IIf(IsNull(Master!V_DATE), "", Master!V_DATE)
    txt(Model).TEXT = IIf(IsNull(Master!Model), "", Master!Model)
    txt(FromGod).Tag = IIf(IsNull(Master!GodownFrom), "", Master!GodownFrom)
    txt(ChassisNo).TEXT = Master!ChassisNo
    
    If txt(FromGod).Tag <> "" Then
        txt(FromGod).TEXT = GCn.Execute("select god_name from godown where god_code = '" & txt(FromGod).Tag & "'").Fields(0).Value
    Else
        txt(FromGod).TEXT = ""
    End If
    txt(ToGod).Tag = IIf(IsNull(Master!Godownto), "", Master!Godownto)
    If txt(ToGod).Tag <> "" Then
        txt(ToGod).TEXT = GCn.Execute("select god_name from godown where god_code = '" & txt(ToGod).Tag & "'").Fields(0).Value
    Else
        txt(ToGod).TEXT = ""
    End If
    txt(AgBooking).Tag = XNull(Master!OrdDocId)
    txt(AgBooking) = XNull(Master!Ord_No)
    txt(6).TEXT = XNull(Master!Narration)
Else
    Call BlankText
End If
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next
    txt(FromGod).Enabled = False
   txtDisabled_Color Me
End Sub
Private Sub Grid_Hide()
    If DgChassis.Visible = True Then DgChassis.Visible = False
    If DGMod.Visible = True Then DGMod.Visible = False
    If DGGod.Visible = True Then DGGod.Visible = False
End Sub


Private Sub Ini_Grid()
    DgChassis.left = Me.left + 45: DgChassis.top = 2745
    DGGod.left = 6660: DGGod.top = mTopScale
    DGMod.left = Me.left + 45: DGMod.top = 2745
End Sub

