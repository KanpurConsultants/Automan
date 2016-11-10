VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "topctl.ocx"
Begin VB.Form FaTDSChal 
   BackColor       =   &H00A2D1F4&
   Caption         =   "T.D.S.Challan Entry"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11415
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   9
      Left            =   5625
      MaxLength       =   40
      TabIndex        =   7
      Top             =   960
      Width           =   3555
   End
   Begin VB.CommandButton btsadd 
      DisabledPicture =   "FaTDSChal.frx":0000
      Height          =   330
      Left            =   9225
      Picture         =   "FaTDSChal.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Add T.D.S. Entries"
      Top             =   930
      Width           =   390
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   8
      Left            =   5625
      TabIndex        =   4
      Top             =   690
      Width           =   1080
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   7
      Left            =   7335
      MaxLength       =   20
      TabIndex        =   5
      Top             =   690
      Width           =   1230
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   9315
      MaxLength       =   21
      TabIndex        =   20
      Top             =   435
      Width           =   2040
   End
   Begin MSDataGridLib.DataGrid DGVType 
      Height          =   3330
      Left            =   10410
      Negotiate       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5130
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   13234931
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
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
         Caption         =   "Challan Type"
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
         DataField       =   "NCat"
         Caption         =   "NCat"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3555.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   5
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   0
      Top             =   420
      Width           =   2235
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2670
      Left            =   9990
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   4710
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   13234931
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   16
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4334.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   1200
      TabIndex        =   9
      Top             =   3885
      Width           =   1230
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   7335
      MaxLength       =   13
      TabIndex        =   2
      Top             =   420
      Width           =   1230
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   1275
      TabIndex        =   6
      Top             =   960
      Width           =   1125
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   3
      Top             =   690
      Width           =   3075
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   5625
      TabIndex        =   1
      Top             =   420
      Width           =   1080
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      tAdd            =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2535
      Left            =   15
      TabIndex        =   8
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   12648447
      Cols            =   16
      BackColorFixed  =   15718825
      ForeColorFixed  =   128
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   10670580
      GridColor       =   255
      GridColorFixed  =   32896
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   16
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name (TDS Deposited)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   9
      Left            =   3240
      TabIndex        =   25
      Top             =   975
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   8
      Left            =   4620
      TabIndex        =   23
      Top             =   705
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   7
      Left            =   6885
      TabIndex        =   22
      Top             =   705
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
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
      Index           =   3
      Left            =   8655
      TabIndex        =   21
      Top             =   435
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Challan Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   45
      TabIndex        =   18
      Top             =   435
      Width           =   1035
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VPrefix"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4875
      TabIndex        =   17
      Top             =   435
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   6
      Left            =   30
      TabIndex        =   15
      Top             =   3900
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   6885
      TabIndex        =   14
      Top             =   435
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month && Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   90
      TabIndex        =   13
      Top             =   975
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Bank A/C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   5
      Left            =   30
      TabIndex        =   12
      Top             =   705
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   4
      Left            =   3960
      TabIndex        =   11
      Top             =   435
      Width           =   900
   End
End
Attribute VB_Name = "FaTDSChal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BackColorSelLeave As String
Dim VNo As Long, NCat As String, GridKey As Integer, TAddMode As Boolean
Dim RsParty As ADODB.Recordset, Master As ADODB.Recordset, RsVType As ADODB.Recordset
Private Const ChalNo As Byte = 0, ChalDate As Byte = 1, BankCode As Byte = 2
Private Const MonthNo As Byte = 3, TDSAmt As Byte = 4, ChalType As Byte = 5
Private Const TxtDocID As Byte = 6, Chq_No As Byte = 8, Chq_Date As Byte = 7
Private Const FVSNo As Byte = 0, FTDSDocId As Byte = 1, FTDSVSno As Byte = 2
Private Const FV_Type As Byte = 3, FV_No As Byte = 4, FV_Sno As Byte = 5
Private Const FV_Date As Byte = 6, FACCode As Byte = 7, FACName As Byte = 8, BankName As Byte = 9
Private Const FAmt As Byte = 9, FTDS As Byte = 10, FTDSAmt As Byte = 11, FCertiNo As Byte = 12, FCertiDate As Byte = 13, FTDSCode As Byte = 14
Private PubDatamanFa As New DMFa.ClsFa

Private Sub TopCtrl1_ePrn()
On Error GoTo ERRORHANDLER
Dim mQRY As String, X11, RST1 As ADODB.Recordset, I As Integer
If Master.RecordCount <= 0 Then Exit Sub
    mQRY = "Select T.CHALTYPE,T.ChalNo,T.ChalDate,T.MonthNo,T.TDSAmt,SG.NAME as Bank,SGA.NAme as AcName,T1.Amt,T1.TDS,T1.TDSAmt as TDSAmt1,T1.CertiNo,T1.CertiDate ," & IIf(PubBackEnd = "A", "Mid(LT.DocID, 4, 5)", "SUBSTRING(LT.DocID, 4, 5)") & " AS LTV_TYPE," & IIf(PubBackEnd = "A", "Mid(LT.DocID, 14, 8)", "SUBSTRING(LT.DocID, 14, 8)") & " AS LTV_NO,LT.V_DATE,T.BankName FROM (((TdsChal as T Left Join TDSChal1 as T1 On T.DOCID=T1.DOCID )  Left Join SubGroup as SG On SG.SubCode=T.BankCode )  Left Join Subgroup as SGA On T1.AcCode=SGA.SubCode) LEFT JOIN LEDGERTDS LT ON (LT.TDSDOCID=T1.TDSDOCID) AND (LT.TDSV_SNO=T1.TDSVSNO) Where T.DOCID='" & Master!DocID & "' Order By T.DOCID,T1.VSno"
    If mQRY = "" Then Exit Sub
    Set RST1 = G_FaCn.Execute(mQRY)
'    X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaTDSChal.ttx", True)
    Set rpt = PubDatamanFa.FaTDSChalRpt
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("TITLE")
                rpt.FormulaFields(I).TEXT = "'TDS Challan'"
        End Select
    Next
    rpt.Database.SetDataSource RST1
    rpt.ReadRecords
    FaReport_View rpt, 0, Me.CAPTION, True
Set RST1 = Nothing
Exit Sub
ERRORHANDLER:    MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub Form_Activate()
    If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        TopCtrl1_eRef
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FaFormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    If PubSec = "SANJEEV" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_NAME='" & Me.CAPTION & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    ElseIf PubSec = "RAHUL" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_CODE='" & Me.Name & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    End If
    '''''''''''''
    PubDatamanFa.FaBackEnd = PubBackEnd
    PubDatamanFa.FaPubLoginDate = PubLoginDate
    PubDatamanFa.FaPubDivCode = PubDivCode
    PubDatamanFa.FaPubSiteCode = PubSiteCode
    PubDatamanFa.FaPubSiteCodeDisplay = PubSiteCodeDisplay
    PubDatamanFa.FaPubSiteName = PubSiteName
    PubDatamanFa.FapubUName = pubUName
    PubDatamanFa.FaDosPort = PubFaDosPort
    PubDatamanFa.FaRunPIF = PubRunPIF
    PubDatamanFa.FaPubSiteType = PubFaSiteType
    Set PubDatamanFa.SetG_FaCn = G_FaCn
    Set PubDatamanFa.SetG_CompCn = G_CompCn
    Set PubDatamanFa.SetrsUserPerm = rsUserPerm.Clone
    Set PubDatamanFa.SetMasterRst = FaMasterRst.Clone
    '''''''''''''
    Set RsVType = New ADODB.Recordset
    RsVType.CursorLocation = adUseClient
    RsVType.Open "Select V_Type As Code,Description As Name,NCat,ContraType From Voucher_Type Where Category in ('TDSCH') Order by Description", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGVType.DataSource = RsVType

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "Select SubCode As Code,Name From SubGroup Left Join City C on SubGroup.CityCode=C.CityCode Where Nature in ('Bank','Cash') Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "Select DocID as SearchCode,TDSChal.*,SUBGROUP.NAME AS BankAcName,VOUCHER_tYPE.Description From (TDSChal LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=TDSCHAL.BANKCODE) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=TDSCHAL.CHALTYPE Order By DocID", G_FaCn, adOpenDynamic, adLockOptimistic
    Disp_Text SETS("INI", Me, Master)
    btsadd.Visible = False
    Ini_Grid
    MoveRec
    Me.left = 0
    Me.top = 0
    Me.width = 11500
'    WinSetting Me
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsParty = Nothing
    Set RsVType = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub Ini_Grid()
    FGrid.left = 0
    FGrid.top = 1320
    FGrid.width = 11400
    DGParty.left = Txt(BankCode).left
    DGParty.top = Txt(BankCode).top + Txt(BankCode).height
    
    DGVType.left = Txt(ChalType).left
    DGVType.top = Txt(ChalType).top + Txt(ChalType).height
    
    DGParty.height = 5640
    With FGrid
        BackColorSelLeave = .BackColor
        .Cols = 15
        .ColWidth(FVSNo) = 250
        
        .TextMatrix(0, FTDSDocId) = "TDS DOCID"
        .ColWidth(FTDSDocId) = 0

        .TextMatrix(0, FTDSVSno) = "FTDSVSno"
        .ColWidth(FTDSVSno) = 0

        .TextMatrix(0, FV_Type) = "Vr.Type"
        .ColAlignmentFixed(FV_Type) = flexAlignLeftCenter
        .ColAlignment(FV_Type) = flexAlignLeftCenter
        .ColWidth(FV_Type) = 600

        .TextMatrix(0, FV_No) = "Vr.No."
        .ColAlignmentFixed(FV_No) = flexAlignRightCenter
        .ColAlignment(FV_No) = flexAlignRightCenter
        .ColWidth(FV_No) = 800

        .TextMatrix(0, FV_Sno) = "S.No"
        .ColAlignmentFixed(FV_Sno) = flexAlignRightCenter
        .ColAlignment(FV_Sno) = flexAlignRightCenter
        .ColWidth(FV_Sno) = 450

        .TextMatrix(0, FV_Date) = "Date"
        .ColAlignmentFixed(FV_Date) = flexAlignLeftCenter
        .ColAlignment(FV_Date) = flexAlignLeftCenter
        .ColWidth(FV_Date) = 950

        .ColWidth(FACCode) = 0
        
        .TextMatrix(0, FACName) = "A/C Name"
        .ColAlignmentFixed(FACName) = flexAlignLeftCenter
        .ColAlignment(FACName) = flexAlignLeftCenter
        .ColWidth(FACName) = 3500
        
        .ColWidth(FAmt) = 0
        .TextMatrix(0, FAmt) = "On Amount"
        .ColAlignmentFixed(FAmt) = flexAlignRightCenter
        .ColAlignment(FAmt) = flexAlignRightCenter
        .ColWidth(FAmt) = 1000

        .TextMatrix(0, FTDS) = "T.D.S. %"
        .ColAlignmentFixed(FTDS) = flexAlignRightCenter
        .ColAlignment(FTDS) = flexAlignRightCenter
        .ColWidth(FTDS) = 800

        .TextMatrix(0, FTDSAmt) = "T.D.S.Amt"
        .ColAlignmentFixed(FTDSAmt) = flexAlignRightCenter
        .ColAlignment(FTDSAmt) = flexAlignRightCenter
        .ColWidth(FTDSAmt) = 900

        .TextMatrix(0, FCertiNo) = "Certi.No."
        .ColAlignmentFixed(FCertiNo) = flexAlignLeftCenter
        .ColAlignment(FCertiNo) = flexAlignLeftCenter
        .ColWidth(FCertiNo) = 850
        
        .TextMatrix(0, FCertiDate) = "Certi.Date"
        .ColAlignmentFixed(FCertiDate) = flexAlignLeftCenter
        .ColAlignment(FCertiDate) = flexAlignLeftCenter
        .ColWidth(FCertiDate) = 950
        
        .ColWidth(FTDSCode) = 0
    End With
End Sub
Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
        Txt(I).Tag = ""
    Next I
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
    Txt(TDSAmt).Enabled = False
End Sub
Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGVType.Visible = True Then DGVType.Visible = False
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    Disp_Text SETS("ADD", Me, Master)
    BlankText
    btsadd.Visible = False
    Txt(ChalType) = G_FaCn.Execute("Select Description From Voucher_Type Where NCat='TDSCH'").Fields(0).Value
    Txt(ChalDate) = PubLoginDate
    Txt(ChalType).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
    End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM, Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
If Master.RecordCount > 0 Then
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, FCertiNo) <> "" Then
            MsgBox "Some Certificate(s) of this challan are already made"
            Exit Sub
        End If
    Next
    
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        G_FaCn.BeginTrans
        XBM = Master.Bookmark
        Set Rst = G_FaCn.Execute("SELECT * From TDSCHAL1 Where DocID='" & Txt(TxtDocID) & "'")
        Do Until Rst.EOF
            G_FaCn.Execute "UPDATE LEDGERTDS SET TDSPOST='' WHERE TDSDocId='" & Rst!TDSDocId & "' AND TDSV_Sno=" & Rst!TDSVSno
            Rst.MoveNext
        Loop
        G_FaCn.Execute ("Delete From TDSCHAL1 Where DocID='" & Txt(TxtDocID) & "'")
        G_FaCn.Execute ("Delete From TDSCHAL  Where DocID='" & Txt(TxtDocID) & "'")
        G_FaCn.Execute ("Delete From LEDGER   Where DocID='" & Txt(TxtDocID) & "'")
        G_FaCn.CommitTrans
        Master.Requery
        If Master.RecordCount >= XBM Then
            Master.Bookmark = XBM
        Else
            If Master.EOF = False Then Master.MoveLast
        End If
        MoveRec
        BUTTONS True, Me, Master, 0
    End If
End If
Set Rst = Nothing
Exit Sub
ELoop:      G_FaCn.RollbackTrans
            MsgBox err.Description, vbCritical, " Deletion Message"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    btsadd.Visible = True
    Txt(ChalType).Enabled = False
    Txt(ChalNo).Enabled = False
    Txt(ChalDate).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT DocID as SearchCode,CHALTYPE,ChalNo,ChalDate,MonthNo,Name As Bank,TDSAmt FROM TDSChal INNER JOIN SubGroup ON TDSChal.BankCode=SubGroup.SubCode ORDER BY CHALNO"
    Set SearchForm = Me
    FAFind.Show vbModal
Exit Sub
ELoop:  If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.Find ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
Exit Sub
ELoop:  If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eRef()
    RsParty.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim Rst As ADODB.Recordset, mTrans As Boolean, SearchCode As String, I As Integer
On Error GoTo ELoop
    If FaIsValid(Txt(ChalType), "Challan Type") = False Then Exit Sub
    If FaIsValid(Txt(ChalNo), "Challan No") = False Then Exit Sub
    If FaIsValid(Txt(ChalDate), "Challan Date") = False Then Exit Sub
    If FaIsValid(Txt(BankCode), "Bank") = False Then Exit Sub
    If FaIsValid(Txt(MonthNo), "Month & Year") = False Then Exit Sub
    If Validate = True Then Exit Sub
    If Trim(FGrid.TextMatrix(1, 1)) = "" Then
        MsgBox "Item Detail Required ": FGrid.Row = 1: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If G_FaCn.Execute("Select Count(*) From TDSChal Where DocID='" & Txt(TxtDocID) & "'").Fields(0) > 0 Then
            Txt(TxtDocID) = VoucherNo(G_FaCn)
            Exit Sub
        End If
    End If
    mTrans = True
    G_FaCn.BeginTrans
    Set Rst = G_FaCn.Execute("SELECT * From TDSCHAL1 Where DocID='" & Txt(TxtDocID) & "'")
    Do Until Rst.EOF
        G_FaCn.Execute "UPDATE LEDGERTDS SET TDSPOST='' WHERE TDSDocId='" & Rst!TDSDocId & "' AND TDSV_Sno=" & Rst!TDSVSno
        Rst.MoveNext
    Loop
    G_FaCn.Execute ("Delete From TDSCHAL1 Where  DocID='" & Txt(TxtDocID) & "'")
    G_FaCn.Execute ("Delete From TDSCHAL  Where  DocID='" & Txt(TxtDocID) & "'")
    G_FaCn.Execute ("Delete From LEDGER   Where  DocID='" & Txt(TxtDocID) & "'")
    G_FaCn.Execute ("Insert Into TDSCHAL (DOCID,ChalType,ChalNo,V_Prefix,ChalDate,BankCode,Site_Code,MonthNo,TDSAmt,u_name,U_ENTDt,U_AE,CHQ_NO,CHQ_DATE,BANKNAME) Values ('" & Txt(TxtDocID) & "','" & Txt(ChalType).Tag & "'," & Val(Txt(ChalNo)) & ",'" & LblVPrefix & "'," & FaConvertDate(Txt(ChalDate)) & ",'" & Txt(BankCode).Tag & "','" & PubSiteCode & "','" & Txt(MonthNo) & "'," & Val(Txt(TDSAmt)) & ",'" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'" & _
    IIf(TopCtrl1.TopText2.CAPTION = "Add", "A", "E") & "'," & FaChk_Text(Txt(Chq_No)) & "," & FaConvertDate(Txt(Chq_Date)) & "," & FaChk_Text(Txt(BankName)) & ")")
    For I = 1 To FGrid.Rows - 1
        G_FaCn.Execute "Insert Into TDSCHAL1 (DOCID,CHALTYPE,ChalNo,VSno,Site_Code,ChalDate,TDSDocId,TDSVSno,V_Date,ACCode,Amt,TDS,TDSAmt,CertiNo,CertiDate,TDSCODE,u_name,U_ENTDt,U_AE) Values ('" & Txt(TxtDocID) & "','" & Txt(ChalType).Tag & "'," & Val(Txt(ChalNo)) & "," & Val(FGrid.TextMatrix(I, FVSNo)) & ",'" & PubSiteCode & "'," & FaConvertDate(Txt(ChalDate)) & ",'" & FGrid.TextMatrix(I, FTDSDocId) & "'," & Val(FGrid.TextMatrix(I, FTDSVSno)) & "," & FaConvertDate(FGrid.TextMatrix(I, FV_Date)) & ",'" & _
        FGrid.TextMatrix(I, FACCode) & "'," & Val(FGrid.TextMatrix(I, FAmt)) & "," & Val(FGrid.TextMatrix(I, FTDS)) & "," & Val(FGrid.TextMatrix(I, FTDSAmt)) & ",'" & FGrid.TextMatrix(I, FCertiNo) & "'," & FaConvertDate(FGrid.TextMatrix(I, FCertiDate)) & ",'" & FGrid.TextMatrix(I, FTDSCode) & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'" & IIf(TopCtrl1.TopText2.CAPTION = "Add", "A", "E") & "')"
        G_FaCn.Execute "UPDATE LEDGERTDS SET TDSPOST='Y' WHERE TDSDocId='" & FGrid.TextMatrix(I, FTDSDocId) & "' AND TDSV_Sno=" & Val(FGrid.TextMatrix(I, FTDSVSno)) & ""
    Next
    I = 0
    Set Rst = G_FaCn.Execute("SELECT SUM(TDSAMT)AS TDSAMOUNT,TDSCODE FROM TDSCHAL1 Where DocID='" & Txt(TxtDocID) & "' GROUP BY TDSCODE")
    Do Until Rst.EOF
        I = I + 1
        G_FaCn.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Chq_No,Chq_Date,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & Txt(TxtDocID) & "'," & I & ",'" & Txt(ChalType).Tag & "'," & Val(Txt(ChalNo)) & ",'" & LblVPrefix & "','" & PubSiteCode & "'," & FaConvertDate(Txt(ChalDate)) & ",'" & Rst!TDSCODE & "',0," & Rst!TDSAMOUNT & ",'" & Txt(BankCode).Tag & "'," & FaChk_Text(Txt(Chq_No)) & "," & _
        FaConvertDate(Txt(Chq_Date)) & ",'" & "T.D.S. Challan " + Trim(Txt(MonthNo)) & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'" & IIf(TopCtrl1.TopText2.CAPTION = "Add", "A", "E") & "')")
        I = I + 1
        G_FaCn.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Chq_No,Chq_Date,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & Txt(TxtDocID) & "'," & I & ",'" & Txt(ChalType).Tag & "'," & Val(Txt(ChalNo)) & ",'" & LblVPrefix & "','" & PubSiteCode & "'," & FaConvertDate(Txt(ChalDate)) & ",'" & Txt(BankCode).Tag & "'," & Rst!TDSAMOUNT & ",0,'" & Rst!TDSCODE & "'," & FaChk_Text(Txt(Chq_No)) & "," & _
        FaConvertDate(Txt(Chq_Date)) & ",'" & "T.D.S. Challan " + Trim(Txt(MonthNo)) & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'" & IIf(TopCtrl1.TopText2.CAPTION = "Add", "A", "E") & "')")
        Rst.MoveNext
    Loop
    If TopCtrl1.TopText2.CAPTION = "Add" Then G_FaCn.Execute "UPDATE VOUCHER_Prefix SET Start_Srl_No=" & Val(Txt(ChalNo)) & " WHERE V_Type='" & Txt(ChalType).Tag & "' and Prefix='" & LblVPrefix & "'"
    G_FaCn.CommitTrans
    mTrans = False
    Master.Requery
    Master.Find "SearchCode ='" & Txt(TxtDocID) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
Set Rst = Nothing
Exit Sub
ELoop:      If mTrans = True Then G_FaCn.RollbackTrans
            If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
Exit Sub
End Sub
Private Sub DGParty_Click()
    DGParty.Visible = False
    If RsParty.RecordCount > 0 Then
        Txt(BankCode).Tag = RsParty!Code
        Txt(BankCode).TEXT = RsParty!Name
    End If
    Txt(BankCode).SetFocus
End Sub
Private Sub MoveRec()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, I As Integer
If Master.RecordCount > 0 Then
    btsadd.Visible = False
    LblVPrefix = Master!V_Prefix
    Txt(TxtDocID) = Master!DocID
    Txt(ChalType).Tag = Master!ChalType
    Txt(ChalType) = Master!Description
    Txt(ChalNo) = Master!ChalNo
    Txt(ChalDate) = Master!ChalDate
    Txt(BankCode).Tag = Master!BankCode
    Txt(BankCode).TEXT = Master!BankAcName
    Txt(MonthNo) = Master!MonthNo
    Txt(TDSAmt) = Master!TDSAmt
    Txt(BankName) = FaXNull(Master!BankName)
    If Not IsNull(Master!Chq_Date) Then
        Txt(Chq_Date) = Master!Chq_Date
    Else
        Txt(Chq_Date) = ""
    End If
    Txt(Chq_No) = FaXNull(Master!Chq_No)
    FGrid.Redraw = False
    FGrid.Rows = 1
    I = 1
    Set Rst = G_FaCn.Execute("Select TDSCHAL1.*,SUBGROUP.NAME AS ACNAME,LEDGERTDS.DOCID AS LDOCID,LEDGERTDS.V_SNO AS LV_SNO,LEDGERTDS.V_DATE AS LV_DATE " & _
    "From (TDSCHAL1 LEFT Join SUBGROUP on SUBGROUP.SUBCODE=TDSCHAL1.ACCODE) " & _
    " LEFT JOIN LEDGERTDS ON (TDSCHAL1.TDSVSno = LEDGERTDS.TDSV_SNo) AND (TDSCHAL1.TDSDocId = LEDGERTDS.TDSDocId) WHERE TDSCHAL1.DOCID='" & Master!DocID & "' ORDER BY VSNO")
    
    Do Until Rst.EOF
        FGrid.AddItem ""
        With FGrid
            .TextMatrix(I, FVSNo) = Rst!VSNO
            .TextMatrix(I, FTDSDocId) = FaXNull(Rst!TDSDocId)
            .TextMatrix(I, FTDSVSno) = FaVNull(Rst!TDSVSno)
            .TextMatrix(I, FV_Type) = MID(FaXNull(Rst!LDocID), 4, 5)
            .TextMatrix(I, FV_No) = MID(FaXNull(Rst!LDocID), 14, 8)
            .TextMatrix(I, FV_Sno) = FaVNull(Rst!LV_SNO)
            If Not IsNull(Rst!LV_Date) Then .TextMatrix(I, FV_Date) = Rst!LV_Date
            .TextMatrix(I, FACCode) = Rst!AcCode
            .TextMatrix(I, FACName) = Rst!AcName
            .TextMatrix(I, FAmt) = Format(Rst!Amt, "0.00")
            .TextMatrix(I, FTDS) = Format(Rst!TDS, "0.0000")
            .TextMatrix(I, FTDSAmt) = Format(Rst!TDSAmt, "0.00")
            .TextMatrix(I, FCertiNo) = FaXNull(Rst!CertiNo)
            .TextMatrix(I, FTDSCode) = FaXNull(Rst!TDSCODE)
            If Not IsNull(Rst!CertiDate) Then
                .TextMatrix(I, FCertiDate) = Rst!CertiDate
            Else
                .TextMatrix(I, FCertiDate) = ""
            End If
        End With
        I = I + 1
        Rst.MoveNext
    Loop
    FGrid.FixedRows = 1
    FGrid.Redraw = True
    If I = 1 Then
        FGrid.Rows = 1
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
Else
    BlankText
End If
Grid_Hide
Set Rst = Nothing
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FillDetail()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, I As Integer, mVdate1 As Date, mVdate2 As Date
If FaIsValid(Txt(MonthNo), "Month & Year") = False Then Exit Sub
FGrid.Redraw = False
FGrid.Rows = 1
I = 1
mVdate1 = Format(CDate("01/" + Trim(Txt(MonthNo))), "DD/MMM/YYYY")
mVdate2 = DateAdd("M", 1, mVdate1)
mVdate2 = mVdate2 - 1
Set Rst = G_FaCn.Execute("Select LEDGERtds.*,SUBGROUP.NAME AS ACNAME From LEDGERTDS Left Join SUBGROUP On SUBGROUP.SUBCODE=LEDGERTDS.TDSDRCODE Where (TDSPOST='' OR TDSPOST IS NULL) AND V_DATE BETWEEN " & FaConvertDate(mVdate1) & " AND " & FaConvertDate(mVdate2) & " Order by TDSDOCID,TDSV_SNO")
Do Until Rst.EOF
    FGrid.AddItem ""
    With FGrid
        .TextMatrix(I, FVSNo) = I
        .TextMatrix(I, FTDSDocId) = FaXNull(Rst!TDSDocId)
        .TextMatrix(I, FTDSVSno) = FaVNull(Rst!TDSV_Sno)
        .TextMatrix(I, FV_Type) = MID(Rst!DocID, 4, 5)
        .TextMatrix(I, FV_No) = MID(Rst!DocID, 14, 8)
        .TextMatrix(I, FV_Sno) = Rst!V_SNo
        .TextMatrix(I, FV_Date) = Rst!v_Date
        .TextMatrix(I, FACCode) = Rst!TDSDRCODE
        .TextMatrix(I, FACName) = Rst!AcName
        .TextMatrix(I, FAmt) = Format(Rst!ONAmt, "0.00")
        .TextMatrix(I, FTDS) = Format(Rst!TDS, "0.0000")
        .TextMatrix(I, FTDSAmt) = Format(Rst!TDSAmt, "0.00")
        .TextMatrix(I, FCertiNo) = ""
        .TextMatrix(I, FCertiDate) = ""
        .TextMatrix(I, FTDSCode) = Rst!TDSCODE
    End With
    Txt(TDSAmt) = Format(Val(Txt(TDSAmt)) + Rst!TDSAmt, "0.00")
    I = I + 1
    Rst.MoveNext
Loop
If I = 1 Then
    FGrid.Rows = 2
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
Else
    FGrid.FixedRows = 1
End If
FGrid.Redraw = True
Grid_Hide
Set Rst = Nothing
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub btsadd_Click()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, I As Integer, mVdate1 As Date, mVdate2 As Date, dFlag As Boolean, J As Integer
If FaIsValid(Txt(MonthNo), "Month & Year") = False Then Exit Sub
FGrid.Redraw = False
I = 0
J = 0
For J = 1 To FGrid.Rows - 1
    I = FGrid.TextMatrix(J, FVSNo)
Next
I = I + 1
J = 0
mVdate1 = Format(CDate("01/" + Trim(Txt(MonthNo))), "DD/MMM/YYYY")
mVdate2 = DateAdd("M", 1, mVdate1)
mVdate2 = mVdate2 - 1
Set Rst = G_FaCn.Execute("Select LEDGERtds.*,SUBGROUP.NAME AS ACNAME From LEDGERTDS Left Join SUBGROUP On SUBGROUP.SUBCODE=LEDGERTDS.TDSDRCODE Where (TDSPOST='' OR TDSPOST IS NULL) AND V_DATE BETWEEN " & FaConvertDate(mVdate1) & " AND " & FaConvertDate(mVdate2) & " Order by TDSDOCID,TDSV_SNO")
Do Until Rst.EOF
    dFlag = False
    For J = 1 To FGrid.Rows - 1
        If FaXNull(Rst!TDSDocId) = FGrid.TextMatrix(J, FTDSDocId) And FGrid.TextMatrix(J, FTDSVSno) = FaVNull(Rst!TDSV_Sno) Then dFlag = True
    Next
    If dFlag = False Then
        FGrid.AddItem ""
        With FGrid
            .TextMatrix(I, FVSNo) = I
            .TextMatrix(I, FTDSDocId) = FaXNull(Rst!TDSDocId)
            .TextMatrix(I, FTDSVSno) = FaVNull(Rst!TDSV_Sno)
            .TextMatrix(I, FV_Type) = MID(Rst!DocID, 4, 5)
            .TextMatrix(I, FV_No) = MID(Rst!DocID, 14, 8)
            .TextMatrix(I, FV_Sno) = Rst!V_SNo
            .TextMatrix(I, FV_Date) = Rst!v_Date
            .TextMatrix(I, FACCode) = Rst!TDSDRCODE
            .TextMatrix(I, FACName) = Rst!AcName
            .TextMatrix(I, FAmt) = Format(Rst!ONAmt, "0.00")
            .TextMatrix(I, FTDS) = Format(Rst!TDS, "0.0000")
            .TextMatrix(I, FTDSAmt) = Format(Rst!TDSAmt, "0.00")
            .TextMatrix(I, FCertiNo) = ""
            .TextMatrix(I, FCertiDate) = ""
            .TextMatrix(I, FTDSCode) = Rst!TDSCODE
        End With
        Txt(TDSAmt) = Format(Val(Txt(TDSAmt)) + Rst!TDSAmt, "0.00")
        I = I + 1
    End If
    Rst.MoveNext
Loop
If I = 1 Then
    FGrid.Rows = 2
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
Else
    FGrid.FixedRows = 1
End If
FGrid.Redraw = True
Grid_Hide
Set Rst = Nothing
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Txt_GotFocus(Index As Integer)
    FaCtrl_GetFocus Txt(Index)
    Grid_Hide
    Select Case Index
        Case ChalType
            If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).TEXT <> RsVType!Name Then
                RsVType.MoveFirst
                RsVType.Find "Name ='" & Txt(Index).TEXT & "'"
            End If
        Case BankCode
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).TEXT <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.Find "Name =" & FaChk_Text(Txt(Index).TEXT)
            End If
        Case ChalNo, ChalDate, MonthNo, Chq_Date, Chq_No
            SendKeys "{Home}+{End}"
    End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case ChalType
            FaDGridTxtKeyDown DGVType, Txt, Index, RsVType, KeyCode, False, 1
        Case BankCode
            FaDGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, False, 1
    End Select
    If DGVType.Visible = False And DGParty.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then FaCtrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then FaCtrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then FaCtrl_UpKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
'            SaveMsg Index
        End If
    End If
End Sub
Private Sub Txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    FaCheckQuote KeyAscii
    Select Case Index
       Case ChalType
            If DGVType.Visible = True Then
                DGVType.Tag = Index
                FaDGridTxtKeyPress Txt, Index, RsVType, KeyAscii, "Name"
            End If
        Case ChalNo
            FaNumPress Txt(Index), KeyAscii, 8, 0
        Case BankCode
            If DGParty.Visible = True Then FaDGridTxtKeyPress Txt, Index, RsParty, KeyAscii, "Name"
    End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    FaCtrl_validate Txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case ChalType
            If FaIsValid(Txt(ChalType), "Challan Type") = False Then Txt(ChalType).SetFocus: Cancel = True:   Exit Sub
            If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).TEXT = RsVType!Name
                Txt(Index).Tag = RsVType!Code
                If TopCtrl1.TopText2.CAPTION = "Add" Then
                    Txt(TxtDocID) = VoucherNo(G_FaCn)
                End If
            End If
        Case ChalNo
            If FaIsValid(Txt(ChalNo), "Challan No") = False Then Txt(ChalNo).SetFocus: Cancel = True: Exit Sub
            Txt(ChalNo) = FaValidate_Numeric(Txt(ChalNo))
            If TopCtrl1.TopText2.CAPTION = "Add" Then Txt(TxtDocID).TEXT = PubDivCode + PubSiteCode & PubSiteCode + Trim(Txt(ChalType).Tag) + Space(5 - Len(Trim(CStr(Txt(ChalType).Tag)))) + Trim(LblVPrefix) + Space(5 - Len(Trim(CStr(LblVPrefix)))) + Space(8 - Len(CStr(Txt(ChalNo)))) + CStr(Txt(ChalNo))
            If Validate = True Then Cancel = True: Exit Sub
        Case ChalDate, Chq_Date
            If Len(Trim(Txt(Index).TEXT)) = 0 Then
                 Txt(Index).TEXT = PubLoginDate
            Else
                Txt(Index).TEXT = PubDatamanFa.FaRetDateFunc(Txt(Index))
            End If
        Case MonthNo
            Txt(Index).TEXT = PubDatamanFa.FaRetMonthYearFunc(Txt(Index))
            If TopCtrl1.TopText2.CAPTION = "Add" Then FillDetail
        Case BankCode
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).TEXT = RsParty!Name
                Txt(Index).Tag = RsParty!Code
            End If
    End Select
End Sub
Private Function Validate() As Boolean
Dim I As Integer, J As Integer, X As String, Y As String, Count As Integer
If TopCtrl1.TopText2 = "Add" Then
    If G_FaCn.Execute("Select Count(*) From TDSCHAL Where DOCID='" & Txt(TxtDocID) & "'").Fields(0) > 0 Then
        MsgBox "Duplicate Challan No.", vbInformation, Me.CAPTION
        Validate = True
        Exit Function
    End If
End If
End Function
Private Function VoucherNo(Conn As ADODB.Connection) As String
Dim Rst As ADODB.Recordset, VouType As String, VNo As Long, vPrefix As String, DocID As String
    VouType = Txt(ChalType).Tag
    Set Rst = Conn.Execute("Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VouType & "' And VP.Date_From<=" & FaConvertDate(Txt(ChalDate)) & " Order By VP.Date_From DESC")
    If Rst.RecordCount > 0 Then
        If Rst!Number_Method = "Manual" Then
        Else
            vPrefix = Rst!prefix
            VNo = Rst!start_srl_no + 1
        End If
    End If
    DocID = PubDivCode + PubSiteCode & PubSiteCode + Trim(VouType) + Space(5 - Len(Trim(CStr(VouType)))) + Trim(vPrefix) + Space(5 - Len(Trim(CStr(vPrefix)))) + Space(8 - Len(CStr(VNo))) + CStr(VNo)
    LblVPrefix.CAPTION = vPrefix
    Txt(TxtDocID).TEXT = DocID
    Txt(ChalNo).TEXT = VNo
    VoucherNo = DocID
Set Rst = Nothing
End Function
Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.Rows > 2 Then
                FGrid.RemoveItem (FGrid.Row)
            Else
                FGrid.Rows = 1
                FGrid.AddItem FGrid.Rows
                FGrid.FixedRows = 1
            End If
            Txt(TDSAmt) = ""
            For I = 1 To FGrid.Rows - 1
                Txt(TDSAmt) = Val(Txt(TDSAmt)) + Val(FGrid.TextMatrix(I, FTDSAmt))
            Next
        End If
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
Exit Sub
End Sub


'TDSChal
'ChalNo,ChalDate,BankCode,MonthNo,TDSAmt,user_name,date_update,AD_ED

'TDSChal1
'ChalNo,VSno,ChalDate,TDSDocId,TDSVSno,V_Date,ACCode,Amt,TDS,TDSAmt,CertiNo,CertiDate
