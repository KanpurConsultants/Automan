VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSprStkUpdate 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Stock Opening Updation"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   11505
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   2
      Left            =   2850
      TabIndex        =   13
      Top             =   525
      Width           =   6870
   End
   Begin VB.CommandButton CmdSel 
      BackColor       =   &H00D3BEC9&
      Caption         =   "Select &Text File"
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
      Index           =   1
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit Form"
      Top             =   4905
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton CmdSel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Select E&xcel File"
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
      Index           =   0
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Form"
      Top             =   840
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   5055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.xls"
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C2D5B9&
      Caption         =   "UMRP Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   4500
      TabIndex        =   3
      Top             =   2010
      Width           =   2160
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   2850
      TabIndex        =   10
      Top             =   3450
      Visible         =   0   'False
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton BtnUpdate 
      BackColor       =   &H00D3BEC9&
      Caption         =   "&Update Price List"
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
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit Form"
      Top             =   3705
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit Form"
      Top             =   3705
      Width           =   2175
   End
   Begin VB.CommandButton BtnPrint 
      BackColor       =   &H00D3BEC9&
      Caption         =   "&Print Variation List"
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
      Left            =   5025
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Reports"
      Top             =   3705
      Width           =   2175
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   1
      Left            =   6450
      MaxLength       =   12
      TabIndex        =   4
      Top             =   2280
      Width           =   1605
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   0
      Left            =   6450
      TabIndex        =   2
      Top             =   1740
      Width           =   1605
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6.MRP_YN     --> Number (0->No/1->Yes)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   9
      Left            =   75
      TabIndex        =   20
      Top             =   2580
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5.Effect_Dt      --> Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   8
      Left            =   75
      TabIndex        =   19
      Top             =   2205
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4.Rate              --> Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   7
      Left            =   75
      TabIndex        =   18
      Top             =   1890
      Width           =   2205
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3.Disc_Code  --> 2 Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   6
      Left            =   75
      TabIndex        =   17
      Top             =   1590
      Width           =   2085
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.Part Name   --> 40 Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   5
      Left            =   75
      TabIndex        =   16
      Top             =   1245
      Width           =   2190
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.Part No.        --> 21 Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   4
      Left            =   75
      TabIndex        =   15
      Top             =   900
      Width           =   2175
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excel Sheet Col Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   105
      TabIndex        =   14
      Top             =   585
      Width           =   2085
   End
   Begin VB.Label LblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%%%%%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   0
      Left            =   8460
      TabIndex        =   12
      Top             =   3165
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Updation Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   2
      Left            =   2850
      TabIndex        =   11
      Top             =   3135
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   1
      Left            =   4530
      TabIndex        =   9
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price Effective Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   4
      Left            =   4530
      TabIndex        =   8
      Top             =   1740
      Width           =   1755
   End
End
Attribute VB_Name = "frmSprStkUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ADDFLAG As Byte, mFileName$
Dim RstNew As ADODB.Recordset, RstPart As ADODB.Recordset
Dim mFlag As Byte
Private Const Eff_Date = 0
Private Const RefNo = 1
Private Const FilePath = 2
Dim mGen_Per As Double, mLST_Per As Double, mLST_Sur As Double

Private Sub btnexit_Click()
    Unload Me
End Sub

Private Sub BTNPRINT_Click()
On Error GoTo lblErrorBox
Dim mDate As Double
Dim RepFileName$, TTXFileName$
Dim mOldRef$, mNewRef$
Dim mReportCount As Integer
Dim Rst As ADODB.Recordset

    If IsValid(txt(Eff_Date), "Effective Date") = False Then Exit Sub
    
    Set GRs = GCn.Execute("select Distinct Effect_Dt,ref_no From Part_PriceList where Div_Code='" & PubDivCode & "' and Effect_Dt<=" & ConvertDate(txt(Eff_Date)) & " and MRP_Yn=" & Check1.Value & " order By Effect_Dt desc")
    
    If GRs.RecordCount = 0 Then
        MsgBox "No Price List Found for mentioned Effective Date", vbInformation, "Validation"
        Exit Sub
    Else
        If GRs!Effect_Dt <> CDate(txt(Eff_Date).TEXT) Then
            txt(Eff_Date) = GRs!Effect_Dt
        End If
        mNewRef = GRs!Ref_No
    End If
    
    Set GRs = GCn.Execute("select Distinct Effect_Dt,Ref_no From Part_PriceList where Div_Code='" & PubDivCode & "' and Effect_Dt<" & ConvertDate(txt(Eff_Date)) & " and MRP_Yn=" & Check1.Value & " order By Effect_Dt desc")
    If GRs.RecordCount = 0 Then
        MsgBox "No Old Price List Found before mentioned Effective Date", vbInformation, "Validation"
        Exit Sub
    Else
        mDate = GRs!Effect_Dt
        mOldRef = GRs!Ref_No
    End If
    
    If MsgBox("Sure to Print Variation List", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
    
    If Check1.Value = 1 Then
        GSQL = "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.MRP as Rate, " & vIsNull("PO.MRP", "0") & " as ORate " & _
            "From (Part_PriceList PN Left Join Part_PriceList PO on PN.Part_No=PO.Part_No) " & _
            "Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code " & _
            "where PN.New_Part=0 and PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PO.Effect_Dt=" & ConvertDate(mDate) & " and PN.MRP_Yn=" & Check1.Value & " and PO.MRP_YN=" & Check1.Value & _
            " Union " & _
            "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.MRP as Rate,0 as ORate " & _
            "From Part_PriceList PN Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code " & _
            "where PN.New_Part=1 and PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PN.MRP_Yn=" & Check1.Value
    Else
        GSQL = "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.TB_SRate as Rate," & vIsNull("PO.TB_SRate", "0") & " as ORate From (Part_PriceList PN Left Join Part_PriceList PO on PN.Part_No=PO.Part_No) Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code where PN.New_Part=0 and  PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PO.Effect_Dt=" & ConvertDate(mDate) & " and PN.MRP_Yn=" & Check1.Value & " and PO.MRP_YN=" & Check1.Value & _
            " Union " & _
            "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.TB_SRate as Rate,0 as ORate From Part_PriceList PN Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code where PN.New_Part=1 and PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PN.MRP_Yn=" & Check1.Value
    End If

    TTXFileName = "PriceListDiff"
    RepFileName = "PriceListDiff"
    Set Rst = GCn.Execute(GSQL)
    
    If Rst.RecordCount > 0 Then
        CreateFieldDefFile Rst, PubRepoPath + "\" & TTXFileName & ".TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & RepFileName & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        For mReportCount = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(mReportCount).FormulaFieldName)
                Case UCase("NewDate")
                    rpt.FormulaFields(mReportCount).TEXT = "'New Date : " & txt(Eff_Date).TEXT & "'"
                Case UCase("OldDate")
                    rpt.FormulaFields(mReportCount).TEXT = "'Old Date : " & Format(mDate, "dd/MMM/yyyy") & "'"
                Case UCase("UMRP")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & IIf(Check1.Value = 0, "List Price Based Price List", "UMRP Price Based Price List") & "'"
                Case UCase("OldRef")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & mOldRef & "'"
                Case UCase("NewRef")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & mNewRef & "'"
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


Private Sub BtnUpdate_Click()
On Error GoTo ErrorTrap
Dim mTran As Boolean
    
    ''  Vishal Jain  26/Oct/2002
    ''  Changes in Structure Made :
    ''  Table : Part            : Field NewPart - > Default Value changed to 1 From 0
    ''          Part_PriceList  : New Field  :   MRP_YN
    ''                          : New Field  :   New_Part
    ''                          : Change Made in Primary Key
    ''  New Table : PartList_New
    
    If Trim(txt(FilePath)) = "" Then
        MsgBox "File not Selected", vbCritical, "Select Import File"
        CmdSel(0).SetFocus
        Exit Sub
    End If
    If IsValid(txt(Eff_Date), "Effective Date") = False Then Exit Sub
    If IsValid(txt(RefNo), "Ref. No.") = False Then Exit Sub
    If GCn.Execute("select Distinct Effect_Dt From Part_PriceList where Div_Code='" & PubDivCode & "' and Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and MRP_Yn=" & Check1.Value).RecordCount > 0 Then
        MsgBox "PriceList is already updated", vbInformation, "Validation"
        Exit Sub
    End If
    
    If MsgBox("Start Price List Updation", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
    
'    Set RstNew = GCn.Execute("select * From  PartList_New")
    Set RstNew = GCn.Execute("select * From [PartList_New$] IN '" & txt(FilePath) & "' 'EXCEL 8.0;' ORDER BY 1;")

    Set RstPart = GCn.Execute("select Part_No From Part where Div_Code='" & PubDivCode & "' Order By Part_No")
    
    If RstNew.RecordCount = 0 Then MsgBox "No Data found for Updation in New Price List Table", vbInformation, "Validation": Exit Sub
    
    GCn.BeginTrans
    mTran = True
    
    GCn.Execute ("Update part set New_Part=0")
    
    With ProgressBar1
        .Min = 0
        .Max = 100
        .Value = 0
        .Visible = True
    End With
    LblName(2).Visible = True
    LblName(0).Visible = True
    Do While Not RstNew.EOF
        ProgressBar1.Value = Round(RstNew.AbsolutePosition * 100 / RstNew.RecordCount, 0)
        ProgressBar1.Refresh
        LblName(0).CAPTION = ProgressBar1.Value & "%"
        LblName(0).Refresh
        If RstPart.RecordCount > 0 Then
            RstPart.MoveFirst
            RstPart.FIND ("Part_No='" & RstNew!Part_No & "'")
        End If
        If RstPart.EOF = True Then      '' not found
            '' Insert Record in Part
            GSQL = "insert into Part(" & _
                "Div_Code,Site_Code,Part_No,Part_Name,Local_Name,Part_NoHelp,Part_NameHelp," & _
                "Part_Grade,Security_Grade,Value_Method,Disc_Factor,MRP,MRP_Effect_Dt," & _
                "TB_SRate,TP_SRate,TB_Effect_Dt," & _
                "New_Part,U_Name, U_EntDt, U_AE) " & _
                " values(" & _
                "'" & PubDivCode & "','" & PubSiteCode & "'," & Chk_Text(RstNew!Part_No) & "," & Chk_Text(XNull(RstNew!Part_Name)) & "," & Chk_Text(XNull(RstNew!Part_Name)) & "," & Replace(Chk_Text(XNull(RstNew!Part_No)), " ", "") & "," & Replace(Chk_Text(XNull(RstNew!Part_Name)), " ", "") & _
                ",'S','A','FIFO','" & RstNew!Disc_code & "'," & IIf(Check1.Value = 1, RstNew!Rate, 0) & "," & IIf(Check1.Value = 1, ConvertDate(txt(Eff_Date)), "Null") & _
                "," & IIf(Check1.Value = 0, RstNew!Rate, 0) & "," & IIf(Check1.Value = 0, ConvertTPRate(RstNew!Rate), 0) & "," & IIf(Check1.Value = 0, ConvertDate(txt(Eff_Date)), "Null") & _
                ",1,'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        Else
            '' Update Record in Part
            GSQL = "Update Part set Disc_Factor='" & RstNew!Disc_code & "',"
            If Check1.Value = 1 Then 'MRP
                GSQL = GSQL & "MRP=" & RstNew!Rate & ",MRP_Effect_Dt=" & ConvertDate(txt(Eff_Date)) & ","
            Else
                GSQL = GSQL & "TB_SRate=" & RstNew!Rate & ",TB_Effect_Dt=" & ConvertDate(txt(Eff_Date)) & _
                    ",TP_SRate=" & ConvertTPRate(RstNew!Rate) & ""
            End If
            GSQL = GSQL & "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & _
                "where part_no='" & RstNew!Part_No & "' and Div_Code='" & PubDivCode & "'"
        End If
        GCn.Execute GSQL
        '' Insert Record in Part_PriceList
        GSQL = "insert into Part_PriceList(" _
                & "Div_Code,Site_Code,Part_No," _
                & "MRP_YN,Ref_No,MRP,New_Part," _
                & "Effect_Dt,TB_SRate,TP_SRate,Disc_Factor," _
                & " U_Name, U_EntDt, U_AE) " _
                & " values(" _
                & "'" & PubDivCode & "','" & PubSiteCode & "','" & RstNew!Part_No & "'," _
                & Check1.Value & ",'" & txt(RefNo).TEXT & "'," & IIf(Check1.Value = 1, RstNew!Rate, 0) & "," & IIf(RstPart.EOF = True, 1, 0) & "," _
                & ConvertDate(txt(Eff_Date)) & "," & IIf(Check1.Value = 0, RstNew!Rate, 0) & "," & IIf(Check1.Value = 0, ConvertTPRate(RstNew!Rate), 0) & ",'" & RstNew!Disc_code & "'," _
                & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        GCn.Execute GSQL
        RstNew.MoveNext
    Loop
    GCn.CommitTrans
    mTran = False

ErrorTrap:
    If mTran = True Then GCn.RollbackTrans
    If err.NUMBER <> 0 Then MsgBox err.Description, vbCritical, "Error Message"
    
End Sub

Private Sub CmdSel_Click(Index As Integer)
On Error GoTo ErrHandler
    mFileName = ""
  CommonDialog1.InitDir = Pub_DataPath
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  CommonDialog1.DialogTitle = "Select XLS File for Price List Updation"
  'CommonDialog1.
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "Excel Files (*.xls)|*.xls" '|Text Files (*.txt)|*.Txt"
  ' Specify default filter
  CommonDialog1.FilterIndex = 1
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  txt(FilePath) = CommonDialog1.FileName
  mFileName = CommonDialog1.FileTitle
  
ErrHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
'    FormKeyDown Me, KeyCode, Shift
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
Call WinSetting(Me)
    'Me.top = 0: Me.left = 0
    Set GRs = GCn.Execute("select GenSurChrgOnSpr,TF.Tax_Per,TF.Tax_Sur_Per From (Syctrl left join TaxForms as TF on Syctrl.LocalTaxFormSpr=TF.Form_Code)")
    If GRs.RecordCount > 0 Then
        mGen_Per = IIf(IsNull(GRs!GenSurChrgOnSpr), 0, GRs!GenSurChrgOnSpr)
        mLST_Per = IIf(IsNull(GRs!Tax_Per), 0, GRs!Tax_Per)
        mLST_Sur = IIf(IsNull(GRs!Tax_Sur_Per), 0, GRs!Tax_Sur_Per)
    Else
        mGen_Per = 0
        mLST_Per = 0
        mLST_Sur = 0
    End If
    Disp_Text True
    CtrlClckCol
    ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_Unload (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstNew = Nothing: Set RstPart = Nothing
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(Eff_Date).BackColor = CtrlBColOrg:
    txt(RefNo).BackColor = CtrlBColOrg:
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus txt(Index)
    txt(Index).Tag = txt(Index)
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> Eff_Date Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf Index <> RefNo Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Eff_Date
            txt(Eff_Date) = RetDate(txt(Index))
    End Select
Set Rst = Nothing
End Sub

Private Sub BlankText()
Dim i As Byte
For i = 0 To txt.Count - 1
    txt(i).TEXT = ""
Next i
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim i As Byte
For i = 0 To txt.Count - 1
    txt(i).Enabled = Enb
Next
End Sub

Private Function ConvertTPRate(ByVal mRate As Double) As Double
Dim xRate As Double
    mRate = mRate + Round(mRate * mGen_Per / 100, 2)
    xRate = Round(mRate * mLST_Per / 100, 2)
    mRate = mRate + xRate + Round(xRate * mLST_Sur / 100, 2)
    ConvertTPRate = mRate
End Function
