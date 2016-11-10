VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSprRep 
   Caption         =   "Master List"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10545.32
   ScaleMode       =   0  'User
   ScaleWidth      =   12143.17
   WindowState     =   2  'Maximized
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
      Height          =   705
      Index           =   1
      Left            =   1530
      TabIndex        =   15
      Top             =   2730
      Visible         =   0   'False
      Width           =   6180
      Begin VB.CheckBox Check1 
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   1
         Left            =   3435
         TabIndex        =   17
         Top             =   285
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   0
         Left            =   1605
         TabIndex        =   16
         Top             =   285
         Value           =   1  'Checked
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00800080&
      Height          =   510
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   9420
      TabIndex        =   13
      Top             =   6060
      Width           =   9480
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
         Height          =   435
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Exit Form"
         Top             =   0
         Width           =   2190
      End
      Begin VB.CommandButton BtnPrint 
         BackColor       =   &H00D3BEC9&
         Caption         =   "&Print"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3345
         MaskColor       =   &H00800080&
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print Reports"
         Top             =   -15
         Width           =   2190
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Period"
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
      Height          =   705
      Index           =   0
      Left            =   1530
      TabIndex        =   9
      Top             =   2010
      Width           =   6180
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   0
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58130433
         CurrentDate     =   37018
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   315
         Index           =   1
         Left            =   4590
         TabIndex        =   2
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58130433
         CurrentDate     =   37018
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   3570
         TabIndex        =   11
         Top             =   277
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   255
         TabIndex        =   10
         Top             =   277
         Width           =   945
      End
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAA"
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
      Left            =   -1140
      TabIndex        =   12
      Top             =   15
      Width           =   11760
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   1
      Top             =   7740
      Width           =   1245
   End
End
Attribute VB_Name = "FrmSprRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents FGrid1 As MSFlexGridLib.MSFlexGrid
Attribute FGrid1.VB_VarHelpID = -1
Dim rsType As Recordset
Public G_SprId             As Byte
Private Const MrRct$ = "SXGR"           'Material Receipt
Private Const MrTrf$ = "SXGRT"          'Material Rectipt  Transfer
Private Const SlCsh$ = "S_SIC"          'Cash Sale
Private Const SlCre$ = "S_SIR"          'Credit Sale
Private Const SlTrfCsh$ = "SXSRC"       'Cash Sale Return
Private Const SlTrfCre$ = "SXSRR"       'Credit Sale Return
Private Sub Form_Load()
Dim i As Byte
Dim Rst As ADODB.Recordset
On Error GoTo lblErrorLoop
    Call WinSetting(Me)
    For i = 0 To Frame1.count - 1
        Frame1(i).BackColor = Me.BackColor
    Next
    DTP1(0).Value = PubStartDate
    DTP1(1).Value = PubLoginDate
    With Frame1(0)
        .left = 2228.81
        .width = 7916.117
        .height = 1131.575
        .top = 3226.194
    End With
    If G_SprId = 4 Or G_SprId = 5 Then
        With Frame1(1)
            .height = 1131.575
            .left = 2228.81
            .top = 4381.845
            .width = 7916.117
            .Visible = True
        End With
    End If
    Set Rst = New ADODB.Recordset
    Exit Sub
lblErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub BTNPRINT_Click()
Dim i As Integer, RepForm As New frmRepForm, NotModify As Boolean
Dim Rst As Recordset, SqlQry$, SelStr$, ForItem$
Dim RepFileName$, RepTitle$
Dim Ac_Str As String, j As Integer
Ac_Str = ""
On Error GoTo lblErrorBox
    If DTP1(0).Value > DTP1(1).Value Then
        MsgBox " The Starting Date Should be less then End Date", vbInformation, "Date"
        Exit Sub
    End If
    If G_SprId = 4 Or G_SprId = 5 Then
        If Check1(0).Value = 0 And Check1(1).Value = 0 Then MsgBox "Please Select the Selection Cash or Credit": Check1(0).SetFocus: Exit Sub
    End If
    Select Case G_SprId
        Case 0, 1
            RepFileName = "SprMatReg"
            RepTitle = Me.CAPTION
            If G_SprId = 0 Then
                SelStr = "SP_Purch.V_Type='" & MrRct & " ' And SP_Purch.V_Date>= #" & Format(DTP1(0).Value, "dd/mmm/yyyy") & "#  and SP_Purch.V_Date<= #" & Format(DTP1(1).Value, "dd/mmm/yyyy") & "#"
            Else
                SelStr = "SP_Purch.V_Type='" & MrTrf & "' And SP_Purch.V_Date>= #" & Format(DTP1(0).Value, "dd/mmm/yyyy") & "#  and SP_Purch.V_Date<= #" & Format(DTP1(1).Value, "dd/mmm/yyyy") & "#"
            End If
            SqlQry = "SELECT SP_Purch.DocID,SP_Purch.Party_Name, SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date, SP_Purch.GR_RR_No, " & _
                "SP_Purch.GR_RR_Date, SP_Purch.Cash_Credit, SP_Purch.Tot_No_of_Items, SP_Purch.Tot_Doc_Qty, SP_Purch.Tot_Phy_Qty," & _
                "SP_Purch.Tot_Goods_Value, SP_Purch.NET_AMT, TaxForms.Form_Desc, SP_Stock.Part_No, SP_Stock.Qty_Doc, SP_Stock.Qty_Rec," & _
                "SP_Stock.Rate, SP_Purch.V_Type, SP_Purch.V_No, SP_Stock.Amount,SP_Purch.V_Date " & _
                "FROM (SP_Purch LEFT JOIN SP_Stock ON (SP_Purch.V_Type = SP_Stock.V_Type) AND (SP_Purch.V_No = SP_Stock.V_No)) " & _
                "LEFT JOIN TaxForms ON SP_Purch.Form_Code = TaxForms.Form_Code " & _
                "Where " & SelStr & ""
        Case 4, 5
            If Check1(0).Value = 1 And Check1(1).Value = 1 Then
                RepFileName = "SprSalRegAll"
            Else
                RepFileName = "SprSalReg"
            End If
            RepTitle = Me.CAPTION
            If G_SprId = 4 Then 'Sales Cash/Credit
                If Check1(0).Value = 1 And Check1(1).Value = 1 Then SelStr = "SP_Sale.V_Type In ('" & SlCsh & "','" & SlCre & "') And "
                If Check1(0).Value = 0 And Check1(1).Value = 1 Then SelStr = "SP_Sale.V_Type = '" & SlCre & "' And "
                If Check1(0).Value = 1 And Check1(1).Value = 0 Then SelStr = "SP_Sale.V_Type = '" & SlCsh & "' And "
                SelStr = SelStr + "SP_Sale.V_Date>= #" & Format(DTP1(0).Value, "dd/mmm/yyyy") & "#  and SP_Sale.V_Date<= #" & Format(DTP1(1).Value, "dd/mmm/yyyy") & "#"
            Else                'Sales Return Cash/Credit
                If Check1(0).Value = 1 And Check1(1).Value = 1 Then SelStr = "SP_Sale.V_Type In ('" & SlTrfCsh & "','" & SlTrfCre & "') And "
                If Check1(0).Value = 0 And Check1(1).Value = 1 Then SelStr = "SP_Sale.V_Type = '" & SlTrfCre & "' And "
                If Check1(0).Value = 1 And Check1(1).Value = 0 Then SelStr = "SP_Sale.V_Type = '" & SlTrfCsh & "' And "
                SelStr = SelStr + "SP_Sale.V_Date>= #" & Format(DTP1(0).Value, "dd/mmm/yyyy") & "#  and SP_Sale.V_Date<= #" & Format(DTP1(1).Value, "dd/mmm/yyyy") & "#"
            End If
            SqlQry = "SELECT SP_Sale.DocID, SP_Sale.V_Date, SP_Sale.V_Type, SP_Sale.V_No, " & _
                     "SP_Sale.Party_Name, SP_Sale.Cash_Credit, SP_Sale.SprAmt_MRP_TB, " & _
                     "SP_Sale.SprAmt_MRP_TP,SP_Sale.SprAmt_TB, SP_Sale.SprAmt_TP, SP_Sale.OilAmt_TB, " & _
                     "SP_Sale.OilAmt_TP, SP_Sale.D_Per_TB, SP_Sale.D_Amt_TB, SP_Sale.D_Per_TP,SP_Sale.D_Amt_TP,SP_Sale.Addition, " & _
                     "SP_Sale.Packing, SP_Sale.Gen_Sur_Per, SP_Sale.Gen_Sur_Amt, SP_Sale.Trans_Amt, SP_Sale.Tax_Per, " & _
                     "SP_Sale.Tax_Amt, SP_Sale.Tax_Sur_Per, SP_Sale.Tax_Sur_Amt,SP_Sale.TOT_Per, SP_Sale.TOT_Amt, " & _
                     "SP_Sale.Rounded, SP_Sale.Total_Amt FROM SP_Sale Where " & SelStr & ""
    End Select
    Set Rst = GCn.Execute(SqlQry)
    If Rst.BOF = False Or Rst.EOF = False Then
        If G_SprId = 4 Or G_SprId = 5 Then
            CreateFieldDefFile Rst, PubRepoPath + "\" & left(RepFileName, 9) & ".TTX", True
        Else
            CreateFieldDefFile Rst, PubRepoPath + "\" & RepFileName & ".TTX", True
        End If
        If Check1(0).Value = 1 And Check1(1).Value = 1 Then
        Else
            RepTitle = RepTitle + "(" + IIf(Check1(0).Value = 1, "Cash Sale ", "Credit Sale ") + ")"
        End If
        RepForm.Tag = RepTitle
        RepForm.CAPTION = "* " + RepTitle + " *"
        With RepForm.CrysReport1
            .Connect = ConnectStr
            .ReportFileName = PubRepoPath + "\" & RepFileName & ".RPT"
             Call Formula_Title(RepForm, RepTitle)
            .Formulas(4) = "DateBetween ='From :'+ '" & Format(DTP1(0).Value, "dd/mmm/yyyy") & "' + ' To ' + '" & Format(DTP1(1).Value, "dd/mmm/yyyy") & "'"
            If G_SprId = 4 Or G_SprId = 5 Then
                If Check1(0).Value = 1 And Check1(1).Value = 1 Then
                Else
                    .Formulas(5) = "CashCredit ="
                End If
            End If
            .SetTablePrivateData 0, 3, Rst
            .Action = 1
        End With
    Else
        MsgBox "No Records to Print", vbInformation, "Information"
        Exit Sub
    End If
    Set RepForm = Nothing
    Set Rst = Nothing
    Exit Sub
lblErrorBox:
    Set RepForm = Nothing
    Set Rst = Nothing
    ProcErrorMsg
End Sub
Private Sub btnexit_Click()
    Set rsType = Nothing
    Unload Me
End Sub
Private Sub FGrid1_Click()
    Call FGrid_Click(FGrid1)
End Sub
Private Sub Opt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
