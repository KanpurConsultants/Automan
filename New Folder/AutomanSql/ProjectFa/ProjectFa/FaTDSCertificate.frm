VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "topctl.ocx"
Begin VB.Form FaTDSCertificate 
   BackColor       =   &H00DFE7C0&
   Caption         =   "T.D.S.Certificate Entry"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9135
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
   ScaleWidth      =   9135
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame_DETAIL1 
      BackColor       =   &H00E7D1DD&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   0
      TabIndex        =   16
      Top             =   4440
      Width           =   9240
      Begin VB.TextBox TXT 
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
         Index           =   4
         Left            =   5655
         TabIndex        =   8
         Top             =   360
         Width           =   3525
      End
      Begin VB.TextBox TXT 
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
         Left            =   5655
         TabIndex        =   7
         Top             =   75
         Width           =   3525
      End
      Begin VB.TextBox TXT 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   75
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total TDS Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   75
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name of Sign.Auth."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3480
         TabIndex        =   18
         Top             =   75
         Width           =   2115
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   3480
         TabIndex        =   17
         Top             =   360
         Width           =   1125
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox TXT 
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
      Index           =   6
      Left            =   1395
      TabIndex        =   3
      Top             =   930
      Width           =   1080
   End
   Begin VB.TextBox TXT 
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
      Left            =   3375
      MaxLength       =   20
      TabIndex        =   4
      Top             =   930
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2670
      Left            =   9990
      Negotiate       =   -1  'True
      TabIndex        =   13
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
   Begin VB.TextBox TXT 
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
      Left            =   3375
      MaxLength       =   13
      TabIndex        =   1
      Top             =   420
      Width           =   1230
   End
   Begin VB.TextBox TXT 
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
      Left            =   1395
      MaxLength       =   50
      TabIndex        =   2
      Top             =   675
      Width           =   3210
   End
   Begin VB.TextBox TXT 
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
      Left            =   1395
      TabIndex        =   0
      Top             =   420
      Width           =   1080
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      tAdd            =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2535
      Left            =   15
      TabIndex        =   5
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
      BackColorBkg    =   14673856
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
      Caption         =   "From Date"
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
      Left            =   180
      TabIndex        =   15
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   2625
      TabIndex        =   14
      Top             =   945
      Width           =   675
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
      Left            =   2625
      TabIndex        =   12
      Top             =   435
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Name"
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
      Left            =   180
      TabIndex        =   11
      Top             =   690
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Certificate No."
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
      Left            =   180
      TabIndex        =   10
      Top             =   435
      Width           =   1170
   End
End
Attribute VB_Name = "FaTDSCertificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BackColorSelLeave As String
Dim VNo As Long, NCat As String, GridKey As Integer, TAddMode As Boolean
Dim RsParty As ADODB.Recordset, Master As ADODB.Recordset
Private Const CertiNo As Byte = 0, CertiDate As Byte = 1, Code As Byte = 2
Private Const FromDate As Byte = 6, ToDate As Byte = 7, TDSAmt As Byte = 3
Private Const FullName As Byte = 5, Desig As Byte = 4
Private Const FVSNO1 As Byte = 0, FChalType As Byte = 1, FChalNo As Byte = 2
Private Const FChalDate As Byte = 3, FAmt As Byte = 4, FTDS As Byte = 5
Private Const FTDSAmt As Byte = 6, FBankCode As Byte = 7, FBankName As Byte = 8
Private Const FDocId As Byte = 9, FVSNo As Byte = 10
Private PubDatamanFa As New DMFa.ClsFa

Private Sub TopCtrl1_ePrn()
On Error GoTo ERRORHANDLER
Dim mQRY As String, X11, RST1 As ADODB.Recordset, I As Integer
If Master.RecordCount <= 0 Then Exit Sub
    mQRY = "SELECT PARTY_LIST.ADD1,PARTY_LIST.ADD2,PARTY_LIST.CITY_NAME,PARTY_LIST.name AS Party,0 AS Interest,TDSCHAL1.Amt AS ONAmt,TDSCHAL1.TDS AS TDSPer,TDSCHAL1.TDSAmt,TDSCHAL1.ChalNo,TDSCHAL1.ChalDate,TDSCHAL1.CertiNo,TDSCHAL1.CertiDate,FullName,Desig,TDSCHAL.BANKNAME AS TDS_BankName From ((TDSCERTI LEFT JOIN TDSCHAL1 ON TDSCERTI.CERTINO=TDSCHAL1.CERTINO) LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=TDSCERTI.CODE) LEFT JOIN TDSCHAL ON TDSCHAL.DOCID=TDSCHAL1.DOCID WHERE TDSCERTI.CERTINO='" & Master!CertiNo & "'"
    If mQRY = "" Then Exit Sub
    Set RST1 = G_FaCn.Execute(mQRY)
'    X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaTDSCert.ttx", True)
    Set rpt = PubDatamanFa.FaTDSCertRpt
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
            Case UCase("Title")
                rpt.FormulaFields(I).TEXT = "'TDS Certificate'"
            Case UCase("FinYear")
                rpt.FormulaFields(I).TEXT = "'FOR THE PERIOD F/Y. " + CStr(Year(PubStartDate)) + "-" + CStr(Year(PubEndDate)) + "'"
            Case UCase("CerMth")
                rpt.FormulaFields(I).TEXT = "'Month Of " + PubDatamanFa.FaRetMonthYearFunc(Txt(CertiDate)) + "'"
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
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "Select SubCode As Code,Name From SubGroup Left Join City C on SubGroup.CityCode=C.CityCode Where Nature Not in ('Bank','Cash','Sale','Purchase','Customer') Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "SELECT CertiNo AS Search_Code,T.*,SG.Name AS PartyName From TDSCerti T Left Join SubGroup SG On SG.SubCode=T.Code Order By T.CertiNo", G_FaCn, adOpenDynamic, adLockOptimistic
    Disp_Text SETS("INI", Me, Master)
    Ini_Grid
    MoveRec
    Me.left = 0
    Me.top = 0
    Me.width = 9350
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsParty = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub Ini_Grid()
    FGrid.left = 0
    FGrid.top = 1320
    FGrid.width = 9200
    DGParty.left = Txt(Code).left
    DGParty.top = Txt(Code).top + Txt(Code).height
    DGParty.height = 5640
    Frame_DETAIL1.top = FGrid.top + FGrid.height + 50
    Frame_DETAIL1.width = FGrid.width
    With FGrid
        BackColorSelLeave = .BackColor
        .Cols = 11
        .ColWidth(FVSNO1) = 250
        
        .TextMatrix(0, FChalType) = "Chal.Type"
        .ColAlignmentFixed(FChalType) = flexAlignLeftCenter
        .ColAlignment(FChalType) = flexAlignLeftCenter
        .ColWidth(FChalType) = 700

        .TextMatrix(0, FChalNo) = "Chal.No."
        .ColAlignmentFixed(FChalNo) = flexAlignRightCenter
        .ColAlignment(FChalNo) = flexAlignRightCenter
        .ColWidth(FChalNo) = 800

        .TextMatrix(0, FChalDate) = "Date"
        .ColAlignmentFixed(FChalDate) = flexAlignLeftCenter
        .ColAlignment(FChalDate) = flexAlignLeftCenter
        .ColWidth(FChalDate) = 950

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

        .ColWidth(FBankCode) = 0
        
        .TextMatrix(0, FBankName) = "Cash/Bank A/C Name"
        .ColAlignmentFixed(FBankName) = flexAlignLeftCenter
        .ColAlignment(FBankName) = flexAlignLeftCenter
        .ColWidth(FBankName) = 3500
        
        .ColWidth(FDocId) = 0
        .ColWidth(FVSNo) = 0
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
End Sub
Private Sub TopCtrl1_eAdd()
Dim RST1 As ADODB.Recordset
On Error GoTo ELoop
    Disp_Text SETS("ADD", Me, Master)
    BlankText
    Set RST1 = G_FaCn.Execute("SELECT Max(CertiNo) as ID FROM TDSCerti")
    If RST1.RecordCount > 0 Then
        Txt(CertiNo) = FaVNull(RST1!ID) + 1
    Else
        Txt(CertiNo) = 1
    End If
    Txt(Code).Enabled = True
    Txt(CertiNo).SetFocus
    Set RST1 = Nothing
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
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        G_FaCn.BeginTrans
        XBM = Master.Bookmark
        G_FaCn.Execute "Update TDSCHAL1 Set CertiNo='',CertiDate=NULL WHERE CertiNo=" & FaChk_Text(Txt(CertiNo).TEXT)
        G_FaCn.Execute "Delete From TDSCerti Where CertiNo =" & FaChk_Text(Txt(CertiNo).TEXT)
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
    Txt(CertiNo).Enabled = False
    Txt(Code).Enabled = False
    Txt(CertiDate).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT CertiNo as SearchCode,CertiNo,CertiDate,Name,TDSAmt FROM TDSCerti Left Join SubGroup ON TDSCerti.Code=SubGroup.SubCode ORDER BY CertiNo"
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
    If FaIsValid(Txt(CertiNo), "Certificate No.") = False Then Exit Sub
    If FaIsValid(Txt(Code), "A/C Name") = False Then Exit Sub
    If FaIsValid(Txt(FromDate), "From Date") = False Then Exit Sub
    If FaIsValid(Txt(ToDate), "To Date") = False Then Exit Sub
    If Validate = True Then Exit Sub
    If Trim(FGrid.TextMatrix(1, 1)) = "" Then
        MsgBox "Item Detail Required ": FGrid.Row = 1: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If G_FaCn.Execute("Select Count(*) From TDSCerti Where CertiNo='" & Txt(CertiNo) & "'").Fields(0) > 0 Then
            MsgBox "Duplicate Certificate No."
            Exit Sub
        End If
    End If
    mTrans = True
    G_FaCn.BeginTrans
    G_FaCn.Execute "Update TDSChal1 Set CertiNo='',CertiDate=NULL WHERE CertiNo='" & Txt(CertiNo) & "'"
    G_FaCn.Execute "DELETE FROM TDSCerti WHERE CertiNo='" & Txt(CertiNo) & "'"
    G_FaCn.Execute "Insert Into TDSCerti (CertiNo,CertiDate,Code,Site_Code,FromDate,ToDate,TDSAmt,FullName,Desig,U_Name,U_EntDt,U_AE)  Values ('" & Txt(CertiNo) & "'," & FaConvertDate(Txt(CertiDate)) & ",'" & Txt(Code).Tag & "','" & PubSiteCode & "'," & FaConvertDate(Txt(FromDate)) & "," & FaConvertDate(Txt(ToDate)) & "," & Val(Txt(TDSAmt)) & ",'" & Txt(FullName) & "','" & Txt(Desig) & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'" & IIf(TopCtrl1.TopText2.CAPTION = "Add", "A", "E") & "')"
    For I = 1 To FGrid.Rows - 1
        G_FaCn.Execute "Update TDSCHAL1 Set CertiNo='" & Txt(CertiNo) & "',CertiDate=" & FaConvertDate(Txt(CertiDate)) & " WHERE DOCID='" & FGrid.TextMatrix(I, FDocId) & "' AND VSNO=" & Val(FGrid.TextMatrix(I, FVSNo))
    Next
    G_FaCn.CommitTrans
    mTrans = False
    Master.Requery
    Master.Find "CERTINO ='" & Txt(CertiNo) & "'"
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
        Txt(Code).Tag = RsParty!Code
        Txt(Code).TEXT = RsParty!Name
    End If
    Txt(Code).SetFocus
End Sub
Private Sub MoveRec()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, I As Integer
If Master.RecordCount > 0 Then
    Txt(CertiNo) = Master!CertiNo
    Txt(CertiDate) = Master!CertiDate
    Txt(Code).TEXT = Master!PartyName
    Txt(Code).Tag = Master!Code
    Txt(FromDate) = Master!FromDate
    Txt(ToDate) = Master!ToDate
    Txt(TDSAmt) = Master!TDSAmt
    Txt(FullName) = Master!FullName
    Txt(Desig) = Master!Desig
    FGrid.Redraw = False
    FGrid.Rows = 1
    I = 1
    Set Rst = G_FaCn.Execute("SELECT TDSChal1.*,SubGroup.Name AS BANBNAME,TDSCHAL.BANKCODE FROM (TDSChal LEFT JOIN TDSChal1 ON TDSChal.DocId=TDSChal1.DocId) LEFT JOIN SubGroup ON TDSChal.BankCode = SubGroup.SubCode WHERE TDSCHAL1.CERTINO='" & Master!CertiNo & "' ORDER BY TDSChal1.CHALDATE,TDSChal1.CHALTYPE,TDSChal1.CHALNO")
    Do Until Rst.EOF
        FGrid.AddItem ""
        With FGrid
            .TextMatrix(I, FVSNo) = ""
            .TextMatrix(I, FChalType) = FaXNull(Rst!ChalType)
            .TextMatrix(I, FChalNo) = FaVNull(Rst!ChalNo)
            .TextMatrix(I, FChalDate) = Rst!ChalDate
            .TextMatrix(I, FAmt) = Format(Rst!Amt, "0.00")
            .TextMatrix(I, FTDS) = Format(Rst!TDS, "0.0000")
            .TextMatrix(I, FTDSAmt) = Format(Rst!TDSAmt, "0.00")
            .TextMatrix(I, FBankCode) = Rst!BankCode
            .TextMatrix(I, FBankName) = Rst!BANBNAME
            .TextMatrix(I, FDocId) = FaXNull(Rst!DocID)
            .TextMatrix(I, FVSNo) = FaVNull(Rst!VSNO)
        End With
        Txt(TDSAmt) = Format(Val(Txt(TDSAmt)) + Rst!TDSAmt, "0.00")
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
Dim Rst As ADODB.Recordset, I As Integer, J As Integer, dFlag As Boolean
If FaIsValid(Txt(Code), "A/C Name") = False Then Exit Sub
If FaIsValid(Txt(FromDate), "From Date") = False Then Exit Sub
If FaIsValid(Txt(ToDate), "To Date") = False Then Exit Sub
FGrid.Redraw = False
FGrid.Rows = 1
I = 1
Set Rst = G_FaCn.Execute("SELECT TDSChal1.*,SubGroup.Name AS BANBNAME,TDSCHAL.BANKCODE FROM (TDSChal LEFT JOIN TDSChal1 ON TDSChal.DocId=TDSChal1.DocId) LEFT JOIN SubGroup ON TDSChal.BankCode=SubGroup.SubCode WHERE TDSCHAL1.CHALDATE BETWEEN " & FaConvertDate(Txt(FromDate)) & " AND " & FaConvertDate(Txt(ToDate)) & " AND TDSCHAL1.ACCODE='" & Txt(Code).Tag & "' AND (TDSCHAL1.CertiNo='' OR TDSCHAL1.CertiNo IS NULL) ORDER BY TDSCHAL1.CHALDATE,TDSCHAL1.CHALTYPE,TDSCHAL1.CHALNO")
Do Until Rst.EOF
    dFlag = False
    For J = 1 To FGrid.Rows - 1
        If FaXNull(Rst!DocID) = FGrid.TextMatrix(J, FDocId) And FaVNull(Rst!VSNO) = FGrid.TextMatrix(J, FVSNo) Then dFlag = True
    Next
    If dFlag = False Then
        FGrid.AddItem ""
        With FGrid
            .TextMatrix(I, FVSNo) = ""
            .TextMatrix(I, FChalType) = FaXNull(Rst!ChalType)
            .TextMatrix(I, FChalNo) = FaVNull(Rst!ChalNo)
            .TextMatrix(I, FChalDate) = Rst!ChalDate
            .TextMatrix(I, FAmt) = Format(Rst!Amt, "0.00")
            .TextMatrix(I, FTDS) = Format(Rst!TDS, "0.0000")
            .TextMatrix(I, FTDSAmt) = Format(Rst!TDSAmt, "0.00")
            .TextMatrix(I, FBankCode) = Rst!BankCode
            .TextMatrix(I, FBankName) = Rst!BANBNAME
            .TextMatrix(I, FDocId) = FaXNull(Rst!DocID)
            .TextMatrix(I, FVSNo) = FaVNull(Rst!VSNO)
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
        Case Code
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).TEXT <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.Find "Name =" & FaChk_Text(Txt(Index).TEXT)
            End If
        Case CertiNo, CertiDate, FromDate, ToDate, FullName, Desig
            SendKeys "{Home}+{End}"
    End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case Code
        FaDGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, False, 1
End Select
If DGParty.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then FaCtrl_DownKeyDown KeyCode, Shift
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then FaCtrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then FaCtrl_UpKeyDown KeyCode, Shift
    End If
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
'        SaveMsg Index
    End If
End If
End Sub
Private Sub Txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    FaCheckQuote KeyAscii
    Select Case Index
        Case CertiNo
            FaNumPress Txt(Index), KeyAscii, 8, 0
        Case Code
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
        Case CertiNo
            If FaIsValid(Txt(CertiNo), "Certificate No") = False Then Txt(CertiDate).SetFocus: Cancel = True: Exit Sub
            Txt(CertiNo) = FaValidate_Numeric(Txt(CertiNo))
            If Validate = True Then Cancel = True: Exit Sub
        Case FromDate, ToDate, CertiDate
            If Len(Trim(Txt(Index).TEXT)) = 0 Then
                 Txt(Index).TEXT = PubLoginDate
            Else
                Txt(Index).TEXT = PubDatamanFa.FaRetDateFunc(Txt(Index))
            End If
            If Index = ToDate Then FillDetail
        Case Code
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
    If G_FaCn.Execute("Select Count(*) From TDSCERTI Where CERTINO='" & Txt(CertiNo) & "'").Fields(0) > 0 Then
        MsgBox "Duplicate Certificate No.", vbInformation, Me.CAPTION
        Validate = True
        Exit Function
    End If
End If
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


'TDSCerti
'CertiNo,CertiDate,Site_Code,Code,FromDate,ToDate,TDSAmt,FullName,Desig,U_NAME,U_EntDt,U_AE
