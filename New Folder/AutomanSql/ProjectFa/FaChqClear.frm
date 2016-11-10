VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FaChqClear 
   BackColor       =   &H00CAF1FD&
   Caption         =   "Cheque/DD Clearing Entry"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11580
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11580
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   300
      Left            =   7695
      TabIndex        =   22
      Top             =   15
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   2655
      TabIndex        =   6
      Top             =   1740
      Width           =   1395
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   6135
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   -15
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         Sorted          =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3330
      Left            =   990
      Negotiate       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3045
      Visible         =   0   'False
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
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
         Caption         =   "Party A/c"
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
            ColumnWidth     =   4919.811
         EndProperty
      EndProperty
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
      Index           =   4
      Left            =   2655
      TabIndex        =   3
      Top             =   975
      Width           =   4935
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   661
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
      Index           =   1
      Left            =   2655
      TabIndex        =   2
      Top             =   720
      Width           =   4935
   End
   Begin MSDataGridLib.DataGrid DGBank 
      Height          =   3330
      Left            =   6780
      Negotiate       =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
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
         Caption         =   "Bank A/c"
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
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDCCFB&
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   2655
      TabIndex        =   5
      Top             =   1485
      Width           =   1395
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   2
      Left            =   2655
      TabIndex        =   4
      Top             =   1230
      Width           =   1395
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
      Index           =   0
      Left            =   2655
      TabIndex        =   1
      Top             =   465
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   4905
      Left            =   45
      TabIndex        =   7
      Top             =   2130
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8652
      _Version        =   393216
      BackColor       =   12648447
      Cols            =   16
      BackColorFixed  =   13352606
      ForeColorFixed  =   128
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   14875388
      GridColor       =   0
      GridColorFixed  =   32896
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
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
         Name            =   "Tahoma"
         Size            =   9
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
      Caption         =   "(Not Applicable in case of multiple Debit && Credit Vouchers)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Index           =   5
      Left            =   7620
      TabIndex        =   21
      Top             =   990
      Width           =   3780
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status(Clear./UnClea/All)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   2
      Left            =   450
      TabIndex        =   20
      Top             =   1755
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   435
      TabIndex        =   16
      Top             =   990
      Width           =   690
   End
   Begin VB.Label LblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   5505
      TabIndex        =   15
      Top             =   1500
      Width           =   45
   End
   Begin VB.Label LblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   5505
      TabIndex        =   14
      Top             =   1215
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   28
      Left            =   435
      TabIndex        =   13
      Top             =   735
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance As Per Book"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   435
      TabIndex        =   11
      Top             =   1500
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance As Per Bank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   4
      Left            =   435
      TabIndex        =   10
      Top             =   1245
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UpTo Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   3
      Left            =   435
      TabIndex        =   9
      Top             =   480
      Width           =   885
   End
End
Attribute VB_Name = "FaChqClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CellBackColLeave As String = &HC0FFFF
Dim RSBank As ADODB.Recordset, RsParty As ADODB.Recordset, Master As ADODB.Recordset, GridKey As Integer, TAddMode As Boolean, mOldName As String
Private BackColorSelLeave As String
Private Const FromDate As Byte = 0, BankAc As Byte = 1, BalBank As Byte = 2, BalBook As Byte = 3, PartyAc As Byte = 4, Status As Byte = 5
Private Const FSNo As Byte = 0, FDocID As Byte = 1, FVSNo As Byte = 2, FVType As Byte = 3
Private Const FVPrefix As Byte = 4, FVNo As Byte = 5, FVDate As Byte = 6, FPartyCode As Byte = 7
Private Const FPartyName As Byte = 8, FDrAmt As Byte = 9, FCrAmt As Byte = 10, FChqNo As Byte = 11
Private Const FChqDate As Byte = 12, FClgDate As Byte = 13, FNarration As Byte = 14, FxClgDate As Byte = 15
Private PubDatamanFa As New DMFa.ClsFa
Dim ListArray As Variant, mListItem As ListItem

Private Sub Ini_Grid()
    FGrid.left = 0: FGrid.top = 2130 ': FGrid.width = Me.width - 300
    With FGrid
        BackColorSelLeave = .BackColor
        .Cols = 16
        .ColWidth(FSNo) = 250                           ' marker
        .TextMatrix(0, FDocID) = "DocID"                ' DocID
        .ColAlignment(FDocID) = flexAlignLeftCenter
        .ColWidth(FDocID) = 0
        .TextMatrix(0, FVSNo) = "V.SNo"                ' V.Sno
        .ColAlignment(FVSNo) = flexAlignLeftCenter
        .ColWidth(FVSNo) = 0
        .TextMatrix(0, FVType) = "V.Type"               ' V.Type
        .ColAlignment(FVType) = flexAlignLeftCenter
        .ColWidth(FVType) = 700
        .TextMatrix(0, FVPrefix) = "V.Prefix"           ' V.Prefix
        .ColAlignment(FVPrefix) = flexAlignLeftCenter
        .ColWidth(FVPrefix) = 700
        .TextMatrix(0, FVNo) = "V.No"
        .ColAlignmentFixed(FVNo) = flexAlignRightCenter
        .ColAlignment(FVNo) = flexAlignRightCenter
        .ColWidth(FVNo) = 800
        .TextMatrix(0, FVDate) = "V.Date"           ' V.Date
        .ColAlignment(FVDate) = flexAlignLeftCenter
        .ColWidth(FVDate) = 900
        .TextMatrix(0, FPartyCode) = "Party Code"           ' party Code
        .ColAlignment(FPartyCode) = flexAlignLeftCenter
        .ColWidth(FPartyCode) = 0
        .TextMatrix(0, FPartyName) = "Ledger A/c"           ' Ledger A/c
        .ColAlignment(FPartyName) = flexAlignLeftCenter
        .ColWidth(FPartyName) = 2500
        .TextMatrix(0, FDrAmt) = "Debit"
        .ColAlignmentFixed(FDrAmt) = flexAlignRightCenter
        .ColAlignment(FDrAmt) = flexAlignRightCenter
        .ColWidth(FDrAmt) = 1000
        .TextMatrix(0, FCrAmt) = "Credit"
        .ColAlignmentFixed(FCrAmt) = flexAlignRightCenter
        .ColAlignment(FCrAmt) = flexAlignRightCenter
        .ColWidth(FCrAmt) = 1000
        .TextMatrix(0, FChqNo) = "Cheque No"
        .ColAlignmentFixed(FChqNo) = flexAlignLeftCenter
        .ColAlignment(FChqNo) = flexAlignLeftCenter
        .ColWidth(FChqNo) = 1100
        .TextMatrix(0, FChqDate) = "Cheque Date"
        .ColAlignment(FChqDate) = flexAlignLeftCenter
        .ColWidth(FChqDate) = 1200
        .TextMatrix(0, FClgDate) = "Clearing Date"
        .ColAlignment(FClgDate) = flexAlignLeftCenter
        .ColWidth(FClgDate) = 1300
        .TextMatrix(0, FNarration) = "Narration"
        .ColAlignment(FNarration) = flexAlignLeftCenter
        .ColWidth(FNarration) = 11000
        .ColWidth(FxClgDate) = 0
    End With
End Sub

Private Sub Command1_Click()
'    On Error GoTo ELoop
'Dim RstRep As ADODB.Recordset
'Dim I As Integer
'Dim X1 As String
'    Set RstRep = New Recordset
'    RstRep.CursorLocation = adUseClient
'
'    Set RstRep = TmpChqPrn(RstRep)
'    If FGrid.Rows > 1 Then
'            For I = 1 To FGrid.Rows - 1
'                  With RstRep
'                      .AddNew
'                      !FVType = XNull(FGrid.TextMatrix(I, FVType))
'                      !FVPrefix = FGrid.TextMatrix(I, FVPrefix)
'                      !FVDate = FGrid.TextMatrix(I, FVDate)
'                      !FChqDate = VNull(FGrid.TextMatrix(I, FChqDate))
'                      !FClgDate = VNull(FGrid.TextMatrix(I, FClgDate))
'                      !FPartyName = FGrid.TextMatrix(I, FPartyName)
'                      !FVNo = Val(FGrid.TextMatrix(I, FVNo))
'                      !FDrAmt = Val(FGrid.TextMatrix(I, FDrAmt))
'                      !FCrAmt = Val(FGrid.TextMatrix(I, FCrAmt))
'                      !FChqNo = FGrid.TextMatrix(I, FChqNo)
'                      !FNarration = FGrid.TextMatrix(I, FNarration)
'                      !BankName = Txt(BankAc)
'                      !Bal_As_Bank = Val(Txt(BalBank))
'                      !Bal_as_Book = Val(Txt(BalBook))
'                      !Status = Txt(Status)
'                  End With
'              Next
'    End If
'
' CreateFieldDefFile RstRep, PubRepoPath + "\BANKREL.TTX", True
'    Set rpt = rdApp.OpenReport(PubRepoPath + "\BANKREL.RPT")
'
'    rpt.Database.SetDataSource RstRep
'    rpt.ReadRecords
'    Call Report_View(rpt, Me.CAPTION, , True)
'    Set RstRep = Nothing
'
'
'
'
'
'Exit Sub
'ELoop:
'    CheckError
'
End Sub

Private Sub DGBank_Click()
    DGBank.Visible = False
    If RSBank.RecordCount > 0 Then
        txt(BankAc).Tag = RSBank!Code
        txt(BankAc).TEXT = RSBank!Name
    End If
    txt(BankAc).SetFocus
End Sub
Private Sub DGParty_Click()
    DGParty.Visible = False
    If RsParty.RecordCount > 0 Then
        txt(PartyAc).Tag = RsParty!Code
        txt(PartyAc).TEXT = RsParty!Name
    End If
    txt(PartyAc).SetFocus
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
    TopCtrl1.Tag = PubUParam
    Me.left = 0
    Me.top = 0
    Me.width = 11900
    Me.height = 7725
    FGrid.height = 5000
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
    Set RSBank = New ADODB.Recordset
    RSBank.CursorLocation = adUseClient
    RSBank.Open "Select SubCode As Code,Name From SubGroup Where Nature in('Bank') Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGBank.DataSource = RSBank
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "Select SubCode As Code,Name From SubGroup Where Nature <>'Bank' Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "Select SubCode As Code,Name From SubGroup Where Nature in('Bank') Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Disp_Text SETS("INI", Me, Master)
    Ini_Grid
    Me.TopCtrl1.TopText1.left = 5800
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_Unload (-1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RSBank = Nothing
    Set RsParty = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    Disp_Text SETS("ADD", Me, Master)
    BlankText
    txt(FromDate) = PubLoginDate
    txt(FromDate).SetFocus
    TopCtrl1.tPrn = True
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
'
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        FGrid.Rows = 1
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub

Private Sub TopCtrl1_ePrn()
'  On Error GoTo ELoop
Dim RstRep As ADODB.Recordset
Dim I As Integer
Dim X1 As String
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    
    Set RstRep = TmpChqPrn(RstRep)
    If FGrid.Rows > 1 Then
            For I = 1 To FGrid.Rows - 1
                  With RstRep
                      .AddNew
                      !FVType = XNull(FGrid.TextMatrix(I, FVType))
                      !FVPrefix = FGrid.TextMatrix(I, FVPrefix)
                      !FVDate = FGrid.TextMatrix(I, FVDate)
                      !FChqDate = VNull(FGrid.TextMatrix(I, FChqDate))
                      !FClgDate = VNull(FGrid.TextMatrix(I, FClgDate))
                      !FPartyName = FGrid.TextMatrix(I, FPartyName)
                      !FVNo = Val(FGrid.TextMatrix(I, FVNo))
                      !FDrAmt = Val(FGrid.TextMatrix(I, FDrAmt))
                      !FCrAmt = Val(FGrid.TextMatrix(I, FCrAmt))
                      !FChqNo = FGrid.TextMatrix(I, FChqNo)
                      !FNarration = FGrid.TextMatrix(I, FNarration)
                      !BankName = txt(BankAc)
                      !Bal_As_Bank = Val(txt(BalBank))
                      !Bal_as_Book = Val(txt(BalBook))
                      !Status = txt(Status)
                      !mVDate = txt(FromDate)
                  End With
              Next
    End If
    
 CreateFieldDefFile RstRep, PubRepoPath + "\BANKREL.TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\BANKREL.RPT")
   
    rpt.Database.SetDataSource RstRep
    rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , True)
    Set RstRep = Nothing
     
Exit Sub
ELoop:
    CheckError
    
End Sub

Private Sub TopCtrl1_eRef()
    RSBank.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim Rst As ADODB.Recordset, mTrans As Boolean, SearchCode As String, I As Integer, j As Integer
On Error GoTo ELoop
    If Validate = True Then Exit Sub
    mTrans = True
    G_FaCn.BeginTrans
    For I = 1 To FGrid.Rows - 1
        If Trim(FGrid.TextMatrix(I, FDocID)) <> "" Then
            If Not StrCmp(FGrid.TextMatrix(I, FClgDate), FGrid.TextMatrix(I, FxClgDate)) Then
                G_FaCn.Execute "Update Ledger Set Chq_No='" & Trim(FGrid.TextMatrix(I, FChqNo)) & "',Chq_Date=" & FaConvertDate(FGrid.TextMatrix(I, FChqDate)) & ",Clg_Date=" & FaConvertDate(FGrid.TextMatrix(I, FClgDate)) & " Where DocID='" & FGrid.TextMatrix(I, FDocID) & "' AND V_SNo=" & Val(FGrid.TextMatrix(I, FVSNo))
            End If
        End If
    Next
    G_FaCn.CommitTrans
    MoveRec
    mTrans = False
    Disp_Text SETS("INI", Me, Master)
Set Rst = Nothing
Exit Sub
ELoop:  If mTrans = True Then G_FaCn.RollbackTrans
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
FaCtrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case BankAc
        If RSBank.RecordCount = 0 Or (RSBank.EOF = True Or RSBank.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RSBank!Name Then
            RSBank.MoveFirst
            RSBank.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case PartyAc
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case Status
        ListArray = Array("All", "Cleared", "Un-Cleared")
        Set mListItem = FaListView_Items(ListView, txt, 1, ListArray, 3)
        ListView.Tag = Index
End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case BankAc
        FaDGridTxtKeyDown DGBank, txt, Index, RSBank, KeyCode, False, 1
    Case PartyAc
        FaDGridTxtKeyDown DGParty, txt, Index, RsParty, KeyCode, False, 1
    Case Status
        If KeyCode <> vbKeyEscape Then
            FaListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 750
        End If
End Select
If DGBank.Visible = False And DGParty.Visible = False And FrmList.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then FaCtrl_DownKeyDown KeyCode, Shift
    If Index <> FromDate Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then FaCtrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
FaCheckQuote keyascii
Select Case Index
    Case BankAc
        If DGBank.Visible = True Then FaDGridTxtKeyPress txt, Index, RSBank, keyascii, "Name"
    Case PartyAc
        If DGParty.Visible = True Then FaDGridTxtKeyPress txt, Index, RsParty, keyascii, "Name"
End Select
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case Status
            FaListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    FaCtrl_validate txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case Status
        If FrmList.Visible = True Then If txt(Index) <> "" Then txt(Index) = ListView.SelectedItem.TEXT
        If txt(BankAc).Tag <> "" Then
            MoveRec
            FGrid.Row = 1
            FGrid.SetFocus
            FGrid.Col = FClgDate
        End If
    Case BankAc
        If RSBank.RecordCount = 0 Or (RSBank.EOF = True Or RSBank.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RSBank!Name
            txt(Index).Tag = RSBank!Code
        End If
    Case PartyAc
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
        End If
    Case FromDate
        txt(Index).TEXT = RetDate(txt(Index))
        MoveRec
End Select
End Sub
Private Sub FGrid_Click()
    FGridClick
End Sub
Private Sub FGrid_DblClick()
    FGrid_KeyPress vbKeyReturn
    TAddMode = False
End Sub
Private Sub FGrid_EnterCell()
    FGrid.CellBackColor = CellBackColEnter
End Sub
Private Sub FGrid_GotFocus()
'    If FGrid.BackColorSel = BackColorSelLeave Then FGrid.Col = 1
    FGrid.BackColorSel = FaBackColorSelEnter
    txtgrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell--> Enter Cell-->KeyDown
    If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        FGrid.CellBackColor = CellBackColLeave
        KeyCode = 0
    ElseIf (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Val(FGrid.Tag) = FGrid.Rows - 1 Then
        FGrid.CellBackColor = CellBackColLeave
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid.Col
            Case FChqNo, FChqDate, FClgDate
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End Select
    End If
    KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
    If TopCtrl1.TopText2 = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case FClgDate  'FChqNo, FChqDate
            FaGet_Text Me, FGrid, txtgrid, 0, False, keyascii
    End Select
    If keyascii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_Scroll()
    Grid_Hide
    txtgrid(0).Visible = False
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_LostFocus()
    If txtgrid(0).Visible = False Then FGrid.BackColorSel = BackColorSelLeave
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    Select Case Index
        Case 0
            FGrid.CellBackColor = CellBackColLeave
            txtgrid(Index).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
            Select Case FGrid.Col
                Case FChqNo, FChqDate, FClgDate
                    If TAddMode = False Then SendKeys "{Home}+{End}"
            End Select
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 0
            If KeyCode = vbKeyEscape Then
                txtgrid(Index).TEXT = txtgrid(Index).Tag
                TxtGrid_KeyUp Index, KeyCode, Shift
                Grid_Hide
                FGrid.SetFocus
                txtgrid(Index).Visible = False
                Exit Sub
            End If
            Select Case FGrid.Col
                Case FChqNo, FChqDate
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave(Index) = True Then
                             FaGridTxtDown FGrid, txtgrid, Index, KeyCode, TAddMode, 14, , FGrid.Col
                        End If
                    End If
                Case FClgDate
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave(Index) = True Then
                            If FGrid.Row < FGrid.Rows - 1 Then
                               FGrid.Row = FGrid.Row + 1
                               FGrid.SetFocus
                            Else
                               FGrid.SetFocus
                            End If
                        End If
                    End If
            End Select
    End Select
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
FaCheckQuote keyascii
Select Case Index
    Case 0
        Select Case FGrid.Col   'Index
        End Select
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            Cancel = Not TxtGridLeave(Index)
    End Select
End Sub
Private Sub FGridClick()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub
Private Function TxtGridLeave(Optional Index As Integer) As Boolean
    Select Case Index
        Case 0
            Select Case FGrid.Col
                Case FChqNo
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txtgrid(Index)
                Case FChqDate, FClgDate
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(txtgrid(Index))
            End Select
    End Select
    TxtGridLeave = True
End Function
Private Sub MoveRec()
On Error GoTo ELoop
Dim Condstr As String
Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset, I As Integer
Dim TotAmtDr As Double, TotAmtCr As Double
    If txt(BankAc).Tag = "" Then Exit Sub
    FGrid.Redraw = False
    FGrid.Rows = 1
    I = 1
    If txt(Status) = "Cleared" Then
        Condstr = " and Clg_Date is not null "
    ElseIf txt(Status) = "Un-Cleared" Then
        Condstr = " and Clg_Date is null "
    End If
    Set Rst = G_FaCn.Execute("Select L.*,SG.Name As PartyName From VIEWLedger L Left Join SubGroup SG on L.PARTY1=SG.SubCode Where L.V_Date<=" & FaConvertDate(txt(FromDate)) & " And L.PARTY='" & txt(BankAc).Tag & "'" & IIf(txt(PartyAc).Tag <> "", " And L.Party1='" & txt(PartyAc).Tag & "'", "") & Condstr & " ORDER BY L.V_Date,L.V_TYPE,L.V_NO")
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, FSNo) = I
                .TextMatrix(I, FDocID) = Rst!DocID
                .TextMatrix(I, FVSNo) = Rst!V_SNo
                .TextMatrix(I, FVType) = Rst!V_Type
                .TextMatrix(I, FVPrefix) = FaXNull(Rst!V_ADD)
                .TextMatrix(I, FVNo) = Rst!V_NO
                .TextMatrix(I, FVDate) = Format(Rst!V_DATE, "dd/MMM/yyyy")
                .TextMatrix(I, FPartyCode) = Rst!Party
                .TextMatrix(I, FPartyName) = FaXNull(Rst!PartyName)
                .TextMatrix(I, FDrAmt) = Format(Rst!Debit, "0.00")
                .TextMatrix(I, FCrAmt) = Format(Rst!Credit, "0.00")
                .TextMatrix(I, FChqNo) = IIf(IsNull(Rst!Chq_No), "", Rst!Chq_No)
                .TextMatrix(I, FChqDate) = Format(Rst!Chq_Date, "dd/MMM/yyyy")
                .TextMatrix(I, FClgDate) = Format(Rst!Clg_Date, "dd/MMM/yyyy")
                .TextMatrix(I, FxClgDate) = .TextMatrix(I, FClgDate)
                .TextMatrix(I, FNarration) = FaXNull(Rst!mNarr) + " " + FaXNull(Rst!Narr)
                
            End With
            I = I + 1
            Rst.MoveNext
        Loop
        
        If PubBackEnd = "S" Then
            Set RST1 = G_FaCn.Execute("Select IsNull(Sum(L.AmtDr),0) As AmtDr From Ledger L Where L.Clg_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        ElseIf PubBackEnd = "A" Then
            Set RST1 = G_FaCn.Execute("Select IIF(IsNull(Sum(L.AmtDr)),0,Sum(L.AmtDr)) As AmtDr From Ledger L Where L.Clg_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        End If
        
        If RST1.RecordCount > 0 Then TotAmtDr = RST1!AmtDr Else TotAmtDr = 0
        
        If PubBackEnd = "S" Then
            Set RST1 = G_FaCn.Execute("Select IsNull(Sum(L.AmtCr),0) As AmtCr From Ledger L Where L.Clg_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        ElseIf PubBackEnd = "A" Then
            Set RST1 = G_FaCn.Execute("Select IIF(IsNull(Sum(L.AmtCr)),0,Sum(L.AmtCr)) As AmtCr From Ledger L Where L.Clg_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        End If
        
        If RST1.RecordCount > 0 Then TotAmtCr = RST1!AmtCr Else TotAmtCr = 0
        
        txt(BalBank) = Format(Abs(TotAmtDr - TotAmtCr), "0.00")
        LblType(0).CAPTION = IIf(TotAmtDr - TotAmtCr > 0, "Dr", "Cr")
        If PubBackEnd = "S" Then
            Set RST1 = G_FaCn.Execute("Select IsNull(Sum(L.AmtDr),0) As AmtDr From Ledger L Where L.V_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        ElseIf PubBackEnd = "A" Then
            Set RST1 = G_FaCn.Execute("Select IIF(IsNull(Sum(L.AmtDr)),0,Sum(L.AmtDr)) As AmtDr From Ledger L Where L.V_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        End If
        If RST1.RecordCount > 0 Then TotAmtDr = RST1!AmtDr Else TotAmtDr = 0
        If PubBackEnd = "S" Then
            Set RST1 = G_FaCn.Execute("Select IsNull(Sum(L.AmtCr),0) As AmtCr From Ledger L Where L.V_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        ElseIf PubBackEnd = "A" Then
            Set RST1 = G_FaCn.Execute("Select IIF(IsNull(Sum(L.AmtCr)),0,Sum(L.AmtCr)) As AmtCr From Ledger L Where L.V_Date<=" & FaConvertDate(txt(FromDate)) & " And L.SubCode='" & txt(BankAc).Tag & "'")
        End If
        If RST1.RecordCount > 0 Then TotAmtCr = RST1!AmtCr Else TotAmtCr = 0
        txt(BalBook) = Format(Abs(TotAmtDr - TotAmtCr), "0.00")
        LblType(1).CAPTION = IIf(TotAmtDr - TotAmtCr > 0, "Dr", "Cr")
        FGrid.FixedRows = 1
    End If
    FGrid.Redraw = True
    If I = 1 Then
        FGrid.Rows = 1
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Grid_Hide
    
Set Rst = Nothing: Set RST1 = Nothing
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Function Validate() As Boolean

End Function
Private Sub BlankText()
Dim I As Byte
    mOldName = ""
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
        txt(I).Tag = ""
    Next I
    LblType(0).CAPTION = ""
    LblType(1).CAPTION = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    txt(Status) = "All"
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
    txt(BalBank).Enabled = False
    txt(BalBook).Enabled = False
    TopCtrl1.tEdit = False
    TopCtrl1.tDel = False
    TopCtrl1.tFirst = False
    TopCtrl1.tNext = False
    TopCtrl1.tPrev = False
    TopCtrl1.tLast = False
    TopCtrl1.tFind = False
End Sub
Private Sub Grid_Hide()
    If DGBank.Visible = True Then DGBank.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
End Sub
Private Sub SaveMsg(Index As Integer)
    Grid_Hide
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
    Else
        Me.ActiveControl.SetFocus
    End If
End Sub
