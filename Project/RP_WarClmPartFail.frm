VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form RP_WarClmPartFail 
   Caption         =   "Warranty Parts Failure Summary"
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
      TabIndex        =   21
      Top             =   2385
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   1725
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
         Height          =   1260
         Left            =   45
         TabIndex        =   8
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
      TabIndex        =   20
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
         TabIndex        =   9
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
         TabIndex        =   10
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
      Height          =   990
      Index           =   0
      Left            =   2700
      TabIndex        =   16
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   18
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
         TabIndex        =   17
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
      Caption         =   "Warranty Parts Failure Summary"
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
      Left            =   2895
      TabIndex        =   19
      Top             =   375
      Width           =   5790
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
Attribute VB_Name = "RP_WarClmPartFail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsType As Recordset
Public g_FormID As Byte

'Report Index
Private Const V_PartFailure As Byte = 1          'Claims Summary

'Object Index
'Text Box
Private Const YearPrefix As Byte = 1
Private Const ClaimType As Byte = 2
Private Const ClaimFrom As Byte = 3
Private Const ClaimUpto As Byte = 4

'Option Buttons
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
Dim RepFileName$, TTXFileName$
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
    SqlQry = "select Jw3.Part_No,Part.Part_name,JW3.Claim_qty as QtyClaim," & _
             "" & cIIF("Jw3.crnoteno<>'' and jw3.claim_rejected=0", "JW3.Qty_PASS", "0") & " as QtyPass," & _
             "" & cIIF("Jw3.Claim_Rejected=1", "JW3.claim_qty", "0") & " as QtyReject," & _
             "" & cIIF("Jw3.crnoteno='' or jw3.crnoteno Is Null", "JW3.Claim_Qty", "0") & " as QtyOutstand " & _
             "FROM (JOB_WARR3 AS JW3 Left JOIN JOB_WARR1 AS JW1 ON JW3.DIV_CODE=JW1.DIV_CODE and " & _
             "JW3.SITE_CODE=JW1.SITE_CODE and JW3.YEAR_PREFIX=JW1.YEAR_PREFIX and  JW3.CLAIM_TYPE=JW1.CLAIM_TYPE and  JW3.CLAIM_NO=JW1.CLAIM_NO) " & _
             "Left Join Part on Jw3.Part_no=Part.Part_No and Part.Div_Code = JW3.Div_Code "
    Select Case g_FormID
        Case V_PartFailure
            SqlQry = SqlQry & "  where (jw1.DispatchNo<>'' and not isnull(jw1.DispatchNo)) and "
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

    TTXFileName = "WarrantyClaim"
    RepFileName = "WarrPartFail"
    Set Rst = GCn.Execute(SqlQry)
    
    If Rst.RecordCount > 0 Then
        CreateFieldDefFile Rst, PubRepoPath + "\" & TTXFileName & ".TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & RepFileName & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        For mReportCount = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(mReportCount).FormulaFieldName)
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
    
    For I = 1 To OptDiv.Count
        OptDiv(I).BackColor = Me.BackColor
    Next
    OptDiv(CurrentDiv).Value = True
    With Frame1(0)
        .left = 3458.498
        .width = 7935.331
        .top = 2166.847
        .height = 1589.021
    End With
    
    With Frame2
        .left = 3458.498
        .width = 7935.331
        .top = 3828.095
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

