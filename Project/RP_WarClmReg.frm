VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form RP_WarClmReg 
   Caption         =   "Warranty Claim Register"
   ClientHeight    =   7140
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11460.21
   ScaleMode       =   0  'User
   ScaleWidth      =   14775.47
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Division Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1845
      Left            =   2700
      TabIndex        =   24
      Top             =   3195
      Width           =   6195
      Begin VB.OptionButton OptDiv 
         Alignment       =   1  'Right Justify
         Caption         =   "All Divisions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   2
         Left            =   2100
         TabIndex        =   9
         Top             =   180
         Width           =   1440
      End
      Begin VB.OptionButton OptDiv 
         Alignment       =   1  'Right Justify
         Caption         =   "Selected Divisions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   10
         Top             =   180
         Width           =   2025
      End
      Begin VB.OptionButton OptDiv 
         Alignment       =   1  'Right Justify
         Caption         =   "Current Division"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Value           =   -1  'True
         Width           =   1725
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
         Height          =   1260
         Left            =   45
         TabIndex        =   11
         Top             =   555
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   2223
         _Version        =   393216
         BackColor       =   12243913
         Cols            =   3
         BackColorFixed  =   128
         ForeColorFixed  =   65535
         BackColorSel    =   16711680
         BackColorBkg    =   13623520
         GridColor       =   128
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800080&
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11475
      TabIndex        =   23
      Top             =   6585
      Width           =   11535
      Begin VB.CommandButton BtnPrint 
         BackColor       =   &H00D3BEC9&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3585
         MaskColor       =   &H00800080&
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print Reports"
         Top             =   15
         Width           =   2175
      End
      Begin VB.CommandButton BtnExit 
         BackColor       =   &H00D3BEC9&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5775
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exit Form"
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1800
      Index           =   0
      Left            =   2700
      TabIndex        =   19
      Top             =   1350
      Width           =   6195
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   4590
         TabIndex        =   2
         Top             =   255
         Width           =   390
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   255
         Width           =   390
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "With Variation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   825
         Width           =   1665
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reporting Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   645
         Index           =   1
         Left            =   30
         TabIndex        =   25
         Top             =   1155
         Width           =   6135
         Begin VB.OptionButton OptOrd 
            Alignment       =   1  'Right Justify
            Caption         =   "Claim No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   165
            TabIndex        =   6
            Top             =   225
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton OptOrd 
            Alignment       =   1  'Right Justify
            Caption         =   "Requisition Slip No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   240
            Index           =   2
            Left            =   3150
            TabIndex        =   7
            Top             =   225
            Width           =   2115
         End
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   4590
         TabIndex        =   4
         Top             =   525
         Width           =   1515
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   3
         Top             =   525
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   3
         Left            =   3180
         TabIndex        =   27
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Prefix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   26
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claim No Upto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   3180
         TabIndex        =   21
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claim No. From "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   540
         Width           =   1425
      End
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty Claim Register"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   3600
      TabIndex        =   22
      Top             =   375
      Width           =   4320
   End
   Begin VB.Shape Shape1 
      Height          =   360
      Left            =   60
      Top             =   7665
      Width           =   11775
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orientation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8265
      TabIndex        =   18
      Top             =   7740
      Width           =   960
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Portrait"
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
      Left            =   9765
      TabIndex        =   17
      Top             =   7755
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4695
      TabIndex        =   16
      Top             =   7740
      Width           =   990
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "80/132 Columns"
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
      Left            =   6135
      TabIndex        =   15
      Top             =   7755
      Width           =   1380
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Printer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2940
      TabIndex        =   14
      Top             =   7710
      Width           =   2595
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1410
      TabIndex        =   0
      Top             =   7740
      Width           =   1245
   End
End
Attribute VB_Name = "RP_WarClmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsType As Recordset
Public g_FormID As Byte


'Report Index
Private Const V_NotVerified As Byte = 1          'Claims Not Verified
Private Const V_NotDispatch As Byte = 2          'Claims Not Dispatch
Private Const V_Rejected As Byte = 3             'Claims Rejected
Private Const V_Outstanding As Byte = 4          'Claims Outstanding At Telco
Private Const V_OverAll As Byte = 5              'Claims OverAll
Private Const V_PartNotIssued As Byte = 6          'Parts Not issued but claimed
Private Const V_PartNotClaimed As Byte = 7          'parts claimed but not issued

'Object Index
'Text Box
Private Const YearPrefix As Byte = 1
Private Const ClaimType As Byte = 2
Private Const ClaimFrom As Byte = 3
Private Const ClaimUpto As Byte = 4

'Order By Option Button
Private Const ProwacOrder As Byte = 1
Private Const ReqOrder As Byte = 2

'Division Selection Option Button
Private Const CurrentDiv  As Byte = 1
Private Const AllDiv As Byte = 2
Private Const SelectedDiv As Byte = 3

Private Sub btnexit_Click()
    Set rsType = Nothing
    Unload Me
End Sub

Private Sub BTNPRINT_Click()
On Error GoTo lblErrorBox
Dim I As Integer
Dim Rst As Recordset, SqlQry$
Dim RepFileName$
Dim DivStr$, DivName$
Dim mReportCount As Integer
        
    If IsValid(Txt(ClaimFrom), "Claim No. From") = False Then Exit Sub
    If IsValid(Txt(ClaimUpto), "Claim No. Upto") = False Then Exit Sub
    If Txt(ClaimUpto) < Txt(ClaimFrom) Then
        MsgBox "Invalid Claim No.", vbInformation, "Report Validation"
        Txt(ClaimUpto).SetFocus
        Exit Sub
    End If
    If OptDiv(SelectedDiv).Value = True Then
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, 0) <> "" Then GoTo StartReport
        Next
        MsgBox "No Division is selected", vbInformation, "Report Validation"
        FGrid.Enabled = True
        FGrid.SetFocus
        Exit Sub
    End If
    
    
StartReport:
    If OptOrd(ProwacOrder).Value = True Then
        SqlQry = "select JC.Job_No,JW3.Year_Prefix+JW3.Claim_Type+JW3.Claim_No as Prowac,JW3.Year_Prefix+JW3.Claim_Type+JW3.Claim_No as OrderByField,JW3.Claim_Date,JW1.ZSR_PASS,JW1.DispatchNo,JW1.DispatchDate,Jw3.IPO_No," & _
                 "JW3.IPO_DATE,JW3.SRL_NO,JW3.PART_NO,PART.PART_NAME,JW3.ISS_QTY,JW3.CLAIM_QTY,(JW3.ISS_QTY-JW3.CLAIM_QTY) AS VERIATION,JW3.LABOUR_AMT,JW3.SPL_CHRG," & _
                 "JW3.NDP,(JW3.NDP*JW3.CLAIM_QTY) AS SPR_AMT,JW3.MISC_CHRG,(JW3.LABOUR_AMT+JW3.SPL_CHRG+(JW3.NDP*JW3.CLAIM_QTY)+JW3.MISC_CHRG) AS TOTAL_CLAIM," & _
                 "JW3.Claim_Rejected,JW3.QTY_PASS,JW3.LABOUR_PASS,JW3.SPL_PASS,JW3.SPR_PASS,JW3.MISC_PASS,JW3.CRNOTENO,JW3.CRNOTEDATE,JW3.REMARKS " & _
                 "FROM ((JOB_WARR3 AS JW3 LEFT JOIN JOB_WARR1 AS JW1 ON JW3.DIV_CODE+JW3.SITE_CODE+JW3.YEAR_PREFIX+JW3.CLAIM_TYPE+JW3.CLAIM_NO=" & _
                 "JW1.DIV_CODE+JW1.SITE_CODE+JW1.YEAR_PREFIX+JW1.CLAIM_TYPE+JW1.CLAIM_NO) LEFT JOIN JOB_CARD AS JC ON JW3.JOB_DOCID=JC.DOCID) LEFT JOIN PART ON JW3.PART_NO=PART.PART_NO and Part.Div_Code = JW3.Div_Code "
    Else
        SqlQry = "select JC.Job_No,JW3.Year_Prefix+JW3.Claim_Type+JW3.Claim_No as Prowac,JW3.Claim_Date,JW1.ZSR_PASS,JW1.DispatchNo,JW1.DispatchDate,Jw3.IPO_No,jw3.ipo_no as OrderByField," & _
                 "JW3.IPO_DATE,JW3.SRL_NO,JW3.PART_NO,PART.PART_NAME,JW3.ISS_QTY,JW3.CLAIM_QTY,(JW3.ISS_QTY-JW3.CLAIM_QTY) AS VERIATION,JW3.LABOUR_AMT,JW3.SPL_CHRG," & _
                 "JW3.NDP,(JW3.NDP*JW3.CLAIM_QTY) AS SPR_AMT,JW3.MISC_CHRG,(JW3.LABOUR_AMT+JW3.SPL_CHRG+(JW3.NDP*JW3.CLAIM_QTY)+JW3.MISC_CHRG) AS TOTAL_CLAIM," & _
                 "JW3.Claim_Rejected,JW3.QTY_PASS,JW3.LABOUR_PASS,JW3.SPL_PASS,JW3.SPR_PASS,JW3.MISC_PASS,JW3.CRNOTENO,JW3.CRNOTEDATE,JW3.REMARKS " & _
                 "FROM ((JOB_WARR3 AS JW3 LEFT JOIN JOB_WARR1 AS JW1 ON JW3.DIV_CODE+JW3.SITE_CODE+JW3.YEAR_PREFIX+JW3.CLAIM_TYPE+JW3.CLAIM_NO=" & _
                 "JW1.DIV_CODE+JW1.SITE_CODE+JW1.YEAR_PREFIX+JW1.CLAIM_TYPE+JW1.CLAIM_NO) LEFT JOIN JOB_CARD AS JC ON JW3.JOB_DOCID=JC.DOCID) LEFT JOIN PART ON JW3.PART_NO=PART.PART_NO and Part.Div_Code = JW3.Div_Code"
    End If
    Select Case g_FormID
        Case V_NotVerified
            SqlQry = SqlQry & " WHERE JW1.ZSR_PASS=0 AND "
        Case V_NotDispatch
            SqlQry = SqlQry & " WHERE JW1.DISPATCHNO='' AND "
        Case V_Rejected
            SqlQry = SqlQry & " WHERE JW3.CLAIM_REJECTED=1 AND "
        Case V_Outstanding
            SqlQry = SqlQry & " WHERE JW1.DISPATCHNO<>'' AND JW3.CRNOTENO='' AND JW3.CLAIM_REJECTED=0 AND "
        Case V_OverAll
            SqlQry = SqlQry & " WHERE "
        Case V_PartNotClaimed
            SqlQry = SqlQry & " WHERE JW3.ISS_QTY<>0 AND JW3.CLAIM_QTY=0 AND "
        Case V_PartNotIssued
            SqlQry = SqlQry & " WHERE JW3.ISS_QTY=0 AND JW3.CLAIM_QTY<>0 AND "
    End Select
    
    '' For Division
    If OptDiv(CurrentDiv).Value = False Then
        For I = 1 To FGrid.Rows - 1
            If (FGrid.TextMatrix(I, 0) <> "" And OptDiv(SelectedDiv).Value = True) Or OptDiv(AllDiv).Value = True Then
                If DivStr = "" Then
                    DivStr = "'" & FGrid.TextMatrix(I, 2) & "'"
                    DivName = FGrid.TextMatrix(I, 1)
                Else
                    DivStr = DivStr & "," & "'" & FGrid.TextMatrix(I, 2) & "'"
                    DivName = DivName & "," & FGrid.TextMatrix(I, 1)
                End If
            End If
        Next I
    Else
        DivStr = "'" & PubDivCode & "'"
        DivName = GCn.Execute("select div_name from division where div_code='" & PubDivCode & "'").Fields(0).Value
    End If
    If OptDiv(AllDiv).Value = True Then
        DivName = "For All Divisions"
    End If
    SqlQry = SqlQry & " jw3.div_code in (" & DivStr & ") and jw3.site_code='" & PubSiteCode & "' and jw3.year_prefix='" & Txt(YearPrefix) & "' and jw3.claim_Type='" & Txt(ClaimType) & "' and jw3.claim_no >='" & Txt(ClaimFrom) & "' and jw3.claim_no<='" & Txt(ClaimUpto) & "'"
    
    
    If OptOrd(ProwacOrder).Value = True Then
            SqlQry = SqlQry & " order by JW3.Year_Prefix+JW3.Claim_Type+JW3.Claim_No"
    Else
            SqlQry = SqlQry & " order by JW3.IPO_NO"
    End If
    
    
    RepFileName = "WarrantyClaim"

    Set Rst = GCn.Execute(SqlQry)
    
    If Rst.RecordCount > 0 Then
        CreateFieldDefFile Rst, PubRepoPath + "\" & RepFileName & ".TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & RepFileName & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        For mReportCount = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(mReportCount).FormulaFieldName)
                Case UCase("Variation")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & chk.Value & "'"
                Case UCase("YearPrefix")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & Txt(YearPrefix) & "'"
                Case UCase("ClaimType")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & Txt(ClaimType) & "'"
                Case UCase("ClaimUpto")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & Txt(ClaimUpto) & "'"
                Case UCase("ClaimFrom")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & Txt(ClaimFrom) & "'"
                Case UCase("Divisions")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & DivName & "'"
                Case UCase("RepOrder")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & IIf(OptOrd(ReqOrder).Value = True, "Order By Requisition Slip No.", "Order By Warranty Claim No") & "'"
            End Select
        Next
        Call Report_View(rpt, Me.CAPTION)
    Else
        MsgBox "No Records to Print", vbInformation, "Information"
        Exit Sub
    End If
    Set Rst = Nothing
    Exit Sub
lblErrorBox:
    Set Rst = Nothing
    ProcErrorMsg
End Sub

Private Sub chk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
End Sub

Private Sub FGrid_Click()
    FGrid.Col = 0
    FGrid.CellFontName = "WINGDINGS"
    FGrid.CellFontSize = 14
    FGrid.TextMatrix(FGrid.Row, 0) = IIf(FGrid.TextMatrix(FGrid.Row, 0) = "ü", " ", "ü")
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True: Exit Sub
    If KeyCode = vbKeySpace Then Call FGrid_Click
End Sub

Private Sub OptDiv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
End Sub

Private Sub Optord_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
End Sub

Private Sub Form_Load()
Dim I As Byte
Dim Rst As ADODB.Recordset
On Error GoTo lblErrorLoop
    Call WinSetting(Me)
    IniGrid
    Set Rst = GCn.Execute("select Div_Code,Div_Name from Division Order By Div_Name")
    FGrid.Rows = 1
    If Rst.RecordCount > 0 Then
        While Not Rst.EOF
            FGrid.AddItem ""
            I = FGrid.Rows - 1
            FGrid.TextMatrix(I, 1) = Rst!Div_Name
            FGrid.TextMatrix(I, 2) = Rst!Div_Code
            Rst.MoveNext
        Wend
    Else
        FGrid.AddItem ""
    End If
    FGrid.FixedRows = 1
    
    For I = 0 To Frame1.Count - 1
        Frame1(I).BackColor = Me.BackColor
    Next
    For I = 1 To OptOrd.Count
        OptOrd(I).BackColor = Me.BackColor
    Next
    
    For I = 1 To OptDiv.Count
        OptDiv(I).BackColor = Me.BackColor
    Next
    OptOrd(ProwacOrder).Value = True
    OptDiv(CurrentDiv).Value = True
    
    With Frame1(0)
        .left = 3458.498
        .width = 7935.331
        .top = 2166.847
        .height = 2889.129
    End With
    With Frame1(1)
        .left = 30
        .width = 6135
        .height = 645
        .top = 1155
    End With
    With Frame2
        .left = 3458.498
        .width = 7935.331
        .top = 5128.204
        .height = 2961.357
    End With
    FGrid.Enabled = False
    Exit Sub
lblErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub OptDiv_Click(Index As Integer)
    Select Case Index
        Case CurrentDiv
            FGrid.Enabled = False
        Case AllDiv
            FGrid.Enabled = False
        Case SelectedDiv
            FGrid.Enabled = True
    End Select
End Sub

Private Sub IniGrid()
    With FGrid
        .left = 10
        .width = Frame2.width - 20
        .top = 555
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 3

        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450
    
        .TextMatrix(0, 1) = "Division"
        .ColAlignment(1) = flexAlignLeftCenter
        .ColWidth(1) = 5300
        
        .ColWidth(2) = 0
    End With
End Sub


Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
End Sub

