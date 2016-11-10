VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateBalance 
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Transfer "
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Height          =   1140
      Left            =   345
      TabIndex        =   2
      Top             =   780
      Width           =   5580
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   795
         Visible         =   0   'False
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   4935
         TabIndex        =   5
         Top             =   750
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update FA Balances"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3165
      TabIndex        =   1
      Top             =   150
      Width           =   2385
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Stock Balances"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   165
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   660
      Left            =   330
      Top             =   30
      Width           =   5520
   End
End
Attribute VB_Name = "frmUpdateBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, J As Integer, NoUpto As Integer, mNo As Integer, SourcePath$, DocID$, mTrf As Boolean, Found As Boolean
Dim mRecQty As Double, mIssQty As Double, mStkVal As Double
Dim XRecNo As Double
Dim CODE1 As Long, mCode As String
Dim cr As Double, dr As Double, SR As Long
Dim mS_NO As Long, RST1 As ADODB.Recordset
Dim mDR_BALANCE As Double, mBALANCE As Double, mPartyCode As Long, mDocId As String
Dim mVNo As Long
Dim TRec1 As ADODB.Recordset, TRec2 As ADODB.Recordset, Temp06 As ADODB.Recordset
Dim RstPart As ADODB.Recordset
Dim RstDiv As ADODB.Recordset
Dim RstVeh_Stk As ADODB.Recordset
Dim RstVeh_Pur As ADODB.Recordset
Dim RstVeh_Stk1 As ADODB.Recordset
Dim RstVeh_Pur1 As ADODB.Recordset
Dim GCN1 As ADODB.Connection
Dim GFa_Cn As ADODB.Connection
Dim mRate As Double, mNarr$

Dim TRec1Qty As Double, TRec2Qty As Double
Dim RstStock As ADODB.Recordset, RstStock2 As ADODB.Recordset, RstStock3 As ADODB.Recordset
Dim mOP_TB_QTY As Double, mOP_TP_QTY As Double, mOP_TB_VAL As Double, mOP_TP_VAL As Double
Dim mPART_ADD As Boolean, TQty As Double
Dim mname$, mInv_No$, mInv_Date$
Dim mRec_TB_Qty As Double, mRec_TB_Val As Double, mRec_TP_Qty As Double, mRec_TP_Val As Double
Dim mIss_TB_Qty As Double, mIss_TB_Val As Double, mIss_TP_Qty As Double, mIss_TP_Val As Double
Dim xMOP_TBQty As Double, xMOP_TBVal As Double, xMOP_TPQty As Double, xMOP_TPVal As Double

Private Function UpdateBalance()

On Error GoTo DispErr
'UPDATING SP_STOCK TABLE
Dim K As Integer, MRPYN As Integer
Dim Condstr$
Dim RsSprStkOld As ADODB.Recordset
Dim PBIncrVal As Long
Dim Counter As Long
    SourcePath = Pub_DataPath & "\" & G_CompCn.Execute("Select OldPath from Company where CentralData_Path='" & PubCenDataPath & "'").Fields(0).Value
    Set GCN1 = New ADODB.Connection
             With GCN1
                    .CursorLocation = adUseClient
                    .Provider = "Microsoft.Jet.OLEDB.4.0"
                    .ConnectionString = "Data Source=" & SourcePath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
                    .Open
                    .BeginTrans
             End With
    GCn.BeginTrans
        GCn.Execute ("Delete from SP_Stock where V_Type='SXAO'")
        'CondStr = " And SS.DocId Not In (Select Distinct DocId From [" & Pub_DataPath & "\" & PubCenDataPath & "\Automan.MDB" & "].Sp_Stock) "
        'Condstr = " And SS.V_Date >= #" & DateAdd("D", -1, DateAdd("D", -365, PubStartDate)) & "# And SS.V_Date <= #" & DateAdd("D", -1, PubStartDate) & "# "
        Set RsSprStkOld = GCN1.Execute(" Select Left(SS.DocId,1) as Div_Code, SS.Part_No,Sum(SS.Qty_Rec)-Sum(SS.Qty_Iss)+Sum(SS.Qty_Ret) as mQty, " & _
                                                "1 as Tax_YN, 1 as MRP_YN, Max(SS.Rate) as mRate, mQty*mRate as Amount, Max(SS.V_Rate) as V_Rate " & _
                                       " From Sp_Stock SS Where Tax_YN=1 And MRP_YN=1 " & Condstr & " " & _
                                       " Group By SS.Part_No,SS.Tax_YN,SS.MRP_YN,Left(SS.DocId,1) " & _
                                       " Union All " & _
                                       " Select Left(SS.DocId,1) as Div_Code, SS.Part_No,Sum(SS.Qty_Rec)-Sum(SS.Qty_Iss)+Sum(SS.Qty_Ret) as mQty, " & _
                                                "1 as Tax_YN, 0 as MRP_YN, Max(SS.Rate) as mRate, mQty*mRate as Amount, Max(SS.V_Rate) as V_Rate " & _
                                       " From Sp_Stock SS Where Tax_YN=1 And MRP_YN=0  " & Condstr & "" & _
                                       " Group By SS.Part_No,SS.Tax_YN,SS.MRP_YN,Left(SS.DocId,1) " & _
                                       " Union All " & _
                                       " Select Left(SS.DocId,1) as Div_Code, SS.Part_No,Sum(SS.Qty_Rec)-Sum(SS.Qty_Iss)+Sum(SS.Qty_Ret) as mQty, " & _
                                                "0 as Tax_YN, 1 as MRP_YN, Max(SS.Rate) as mRate, mQty*mRate as Amount, Max(SS.V_Rate) as V_Rate " & _
                                       " From Sp_Stock SS Where Tax_YN=0 And MRP_YN=1 " & Condstr & "" & _
                                       " Group By SS.Part_No,SS.Tax_YN,SS.MRP_YN,Left(SS.DocId,1) " & _
                                       " Union All " & _
                                       " Select Left(SS.DocId,1) as Div_Code, SS.Part_No,Sum(SS.Qty_Rec)-Sum(SS.Qty_Iss)+Sum(SS.Qty_Ret) as mQty, " & _
                                                "0 as Tax_YN, 0 as MRP_YN, Max(SS.Rate) as mRate, mQty*mRate as Amount, Max(SS.V_Rate) as V_Rate " & _
                                       " From Sp_Stock SS Where Tax_YN=0 And MRP_YN=0 " & Condstr & "" & _
                                       " Group By SS.Part_No,SS.Tax_YN,SS.MRP_YN,Left(SS.DocId,1) ")
        If RsSprStkOld.RecordCount > 0 Then
            PBIncrVal = RsSprStkOld.RecordCount / 100
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            Do Until RsSprStkOld.EOF
                Counter = Counter + 1
                DocID = RsSprStkOld!Div_Code & PubSiteCode & PubSiteCode & " SXAO" & " SPOP" & Format(Counter, "00000000")
                GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                    "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,V_Rate,MRP_Rate," & _
                    "Site_Code,U_Name,U_EntDt,U_AE) " & _
                    "VALUES ('" & DocID & "',1,'SXAO'," & Format(Counter, "00000000") & ",#" & DateAdd("yyyy", -1, PubEndDate) & "#,'" & RsSprStkOld!Part_No & _
                    "','" & PubSprCounterGodown & "'," & RsSprStkOld!MRP_YN & "," & RsSprStkOld!Tax_YN & "," & RsSprStkOld!mQty & "," & RsSprStkOld!mRate & "," & Round(RsSprStkOld!Amount, 2) & "," & RsSprStkOld!V_Rate & ",0" & _
                    ",'" & PubSiteCode & "','" & pubUName & "',#" & PubServerDate & "#,'A')"
                If Counter Mod PBIncrVal = 0 And ProgressBar1.Value < 100 Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                End If
                RsSprStkOld.MoveNext
            Loop
            ProgressBar1.Visible = False
        End If
    GCn.CommitTrans
        
''''    Set RstDiv = GCN1.Execute("select Div_Code,Div_Name from Division")
''''    For K = 1 To RstDiv.RecordCount
''''        For MRPYN = 0 To 1
''''                Label1.CAPTION = "Please Wait ! Calculation in Progress...."
''''                Label1.Refresh
''''                Condstr = " where SPStk.V_Date <= " & ConvertDate(DateAdd("yyyy", -1, PubEndDate)) & ""
''''                CondDivCode = " and left(SPStk.DocID,1) in ('" & RstDiv!Div_Code & "')"
''''                CondDivCode1 = " and left(Stk.DocID,1) in ('" & RstDiv!Div_Code & "')"
''''
''''                Condstr = Condstr & CondDivCode
''''                CondPartNos = " and SPStk.Part_No in (select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(DateAdd("yyyy", -1, PubEndDate)) & CondDivCode1 & ")"
''''                CondPartNos1 = " Part.Part_No in ( select Distinct Stk.Part_No from SP_Stock as Stk where Stk.V_Date<=" & ConvertDate(DateAdd("yyyy", -1, PubEndDate)) & CondDivCode1 & ")"
''''                CondPartNosOpStk = ""
''''                Set Temp06 = New ADODB.Recordset
''''                Set Temp06 = TmpTemp06(Temp06)
''''                If MRPYN = 0 Then
''''                    CondStrMRP = " and SPStk.MRP_YN=0"
''''                ElseIf MRPYN = 1 Then
''''                    CondStrMRP = " and SPStk.MRP_YN=1"
''''                End If
''''
''''                Condstr = Condstr & CondStrMRP
''''                'For RstPart, SQL
''''                GSQL = "Select Distinct Part_No From SP_Stock SPStk where Part_No <> '' "
''''                GSQL = GSQL & CondDivCode & " Order By Part_No"
''''
''''                Set RstPart = GCN1.Execute(GSQL)
''''
'''''                RstPart.FIND ("Part_No='265135608201'")
''''
''''                mQRY = "select SPStk.Part_No,SPStk.V_DATE,SPStk.Qty_Rec as Qty,SPStk.V_Rate as Rate " & _
''''                    "From " & _
''''                    "SP_Stock as SPStk left Join [" & PubSFADataPath & "].Voucher_Type as VT on Vt.V_type=SPStk.V_type " & Condstr & CondPartNos
''''
''''                GSQL = mQRY & " and SpStk.Tax_YN=1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
''''                Set TRec1 = New Recordset
''''                With TRec1
''''                    .CursorLocation = adUseClient
''''                    .Open (GSQL), GCN1, adOpenDynamic, adLockOptimistic
''''                End With
''''                '******* Taxpaid Qty
''''                GSQL = mQRY & " and SpStk.Tax_YN<>1 and Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"
''''                Set TRec2 = New Recordset
''''                With TRec2
''''                    .CursorLocation = adUseClient
''''                    .Open (GSQL), GCN1, adOpenDynamic, adLockOptimistic
''''                End With
''''
''''                mQRY = "select SPStk.V_Type,SPStk.Part_No,SPStk.V_DATE,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate,Vt.StkTrn " & _
''''                        "From " & _
''''                        "SP_Stock as SPStk left Join [" & PubSFADataPath & "].Voucher_Type as VT on Vt.V_type=SPStk.V_type " & _
''''                        Condstr & CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
''''                GSQL = mQRY & " and SPStk.V_Type='SXAO' Order By SPStk.Part_No,SPStk.V_Date,mid(SPStk.DocID,4,5)"
''''                Set RstStock = GCN1.Execute(GSQL)
''''                '******* Taxable + Taxpaid Qty for With in Date Period Loop
''''                    GSQL = "select SPStk.V_Type,SPStk.V_DATE,SPStk.Part_No,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate,Vt.StkTrn,Vt.Description " & _
''''                    "From " & _
''''                    "SP_Stock as SPStk left Join [" & PubSFADataPath & "].Voucher_Type as VT on Vt.V_type=SPStk.V_type " & _
''''                    "where (SPStk.V_Date >= " & ConvertDate(DateAdd("yyyy", -1, PubStartDate)) & " And SPStk.v_date <= " & ConvertDate(DateAdd("yyyy", -1, PubEndDate)) & " ) " & _
''''                     CondDivCode & CondPartNosOpStk & CondStrMRP & CondPartNos
''''                     GSQL = GSQL & " Order By SPStk.Part_No,SPStk.V_Date,SPStk.Tax_YN,mid(SPStk.DocID,4,5)"
''''                     Set RstStock2 = GCN1.Execute(GSQL)
''''                '***********
''''                Dim counter1 As Double
''''                counter1 = 0
''''                Label1.CAPTION = " Calculating Stock For " & RstDiv!Div_Name
''''                Label1.Refresh
''''                Do While Not RstPart.EOF
''''                    counter1 = counter1 + 1
''''                    ProgressBar1.Visible = True: Label2.Visible = True
''''                    Label2.CAPTION = Format((counter1 / RstPart.RecordCount) * 100, "0") & "%"
''''                    ProgressBar1.Value = Format((counter1 / RstPart.RecordCount) * 100, "0")
''''                    Label2.Refresh
''''
''''                    '' For All Parts
''''                        NoUpto = 1
''''                        mNo = 1
''''
''''                    TRec1Qty = 0
''''                    TRec2Qty = 0
''''
''''                    mOP_TB_QTY = 0: mOP_TP_QTY = 0: mOP_TB_VAL = 0: mOP_TP_VAL = 0
''''                    mIss_TB_Qty = 0: mIss_TB_Val = 0: mIss_TP_Qty = 0: mIss_TP_Val = 0
''''                    mRec_TB_Qty = 0: mRec_TB_Val = 0: mRec_TP_Qty = 0: mRec_TP_Val = 0
''''
''''                    TRec1.Filter = ""
''''                    If TRec1.RecordCount > 0 Then    'Taxable Rect
''''                        TRec1.MoveFirst
''''                        TRec1.Filter = ("Part_No='" & RstPart!Part_No & "'")
''''                        If TRec1.EOF = False Then
''''                            TRec1Qty = TRec1!Qty
''''                        End If
''''                    End If
''''                    TRec2.Filter = ""
''''                    If TRec2.RecordCount > 0 Then    'Taxpaid Rect
''''                        TRec2.MoveFirst
''''                        TRec2.Filter = ("Part_No='" & RstPart!Part_No & "'")
''''                        If TRec2.EOF = False Then
''''                            TRec2Qty = TRec2!Qty
''''                        End If
''''                    End If
''''                    If RstStock.RecordCount > 0 Then
''''
''''                        RstStock.MoveFirst
''''                        RstStock.FIND ("Part_No='" & RstPart!Part_No & "'")
''''                        If RstStock.EOF = False Then
''''                            Do While RstStock!Part_No = RstPart!Part_No    'Opening Calculation
''''                                If RstStock!StkTrn = "-" Then
''''                                    If RstStock!Tax_YN = 1 Then     '' Taxable
''''                                        mRate = 0
''''                                        Call X_VAL11(TRec1, RstStock!Qty_Iss, mRate, "")
''''                                    Else
''''                                        mRate = 0
''''                                        Call X_VAL22(TRec2, RstStock!Qty_Iss, mRate, "")
''''                                    End If
''''                                ElseIf RstStock!StkTrn = "+" Then
''''                                    If RstStock!Tax_YN = 1 Then     '' Taxable
''''                                        mOP_TB_QTY = mOP_TB_QTY + RstStock!Qty_Rec
''''                                        mOP_TB_VAL = mOP_TB_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
''''                                    Else
''''                                        mOP_TP_QTY = mOP_TP_QTY + RstStock!Qty_Rec
''''                                        mOP_TP_VAL = mOP_TP_VAL + (RstStock!Qty_Rec * RstStock!V_Rate)
''''                                    End If
''''                                End If
''''                                RstStock.MoveNext
''''                                If RstStock.EOF Then
''''                                    Exit Do
''''                                ElseIf RstStock!Part_No <> RstPart!Part_No Then
''''                                    Exit Do
''''                                End If
''''                            Loop
''''
''''                        End If
''''                    End If
''''                    xMOP_TBQty = mOP_TB_QTY:        xMOP_TPQty = mOP_TP_QTY
''''                    xMOP_TBVal = mOP_TB_VAL:        xMOP_TPVal = mOP_TP_VAL
''''                    '**
''''                    mIss_TB_Qty = 0:                mIss_TB_Val = 0
''''                    mIss_TP_Qty = 0:                mIss_TP_Val = 0
''''                    '**
''''                    mTrf = False
''''
''''                    If RstStock2.RecordCount > 0 Then
''''                        RstStock2.MoveFirst
''''                        RstStock2.FIND ("Part_No='" & RstPart!Part_No & "'")
''''                        If RstStock2.EOF = False Then
''''                            Do While RstStock2!Part_No = RstPart!Part_No
''''                                mNarr = ""
''''                                If RstStock2!StkTrn = "-" Then
''''                                   If RstStock2!Tax_YN = 1 Then     '' Taxable
''''                                       mRate = 0
''''                                       Call X_VAL11(TRec1, RstStock2!Qty_Iss, mRate, mNarr)
''''                                   Else
''''                                       mRate = 0
''''                                       Call X_VAL22(TRec2, RstStock2!Qty_Iss, mRate, mNarr)
''''                                   End If
''''
''''                                ElseIf RstStock2!StkTrn = "+" Then
''''                                    If RstStock2!Tax_YN = 1 Then     '' Taxable
''''                                        mOP_TB_QTY = mOP_TB_QTY + RstStock2!Qty_Rec
''''                                        mOP_TB_VAL = mOP_TB_VAL + (RstStock2!Qty_Rec * RstStock2!V_Rate)
''''
''''                                        mRec_TB_Qty = mRec_TB_Qty + RstStock2!Qty_Rec
''''                                        mRec_TB_Val = mRec_TB_Val + (RstStock2!Qty_Rec * RstStock2!V_Rate)
''''                                    Else
''''                                        mOP_TP_QTY = mOP_TP_QTY + RstStock2!Qty_Rec
''''                                        mOP_TP_VAL = mOP_TP_VAL + (RstStock2!Qty_Rec * RstStock2!V_Rate)
''''                                        mRec_TP_Qty = mRec_TP_Qty + RstStock2!Qty_Rec
''''                                        mRec_TP_Val = mRec_TP_Val + (RstStock2!Qty_Rec * RstStock2!V_Rate)
''''                                    End If
''''                                End If
''''                                RstStock2.MoveNext
''''                                If RstStock2.EOF Then
''''                                    Exit Do
''''                                ElseIf RstStock2!Part_No <> RstPart!Part_No Then
''''                                    Exit Do
''''                                End If
''''                            Loop
''''                        End If
''''                    End If
''''                    If (xMOP_TBQty + mOP_TB_QTY) <> 0 Or (xMOP_TPQty + mOP_TP_QTY) <> 0 Then
''''                        If mOP_TB_QTY = 0 Then
''''                            mOP_TB_VAL = 0
''''                        ElseIf mOP_TB_QTY < 0 Then
''''                            If mOP_TB_VAL > 0 Then
''''                                mOP_TB_VAL = -1 * mOP_TB_VAL
''''                            End If
''''                        End If
''''                        If mOP_TP_QTY = 0 Then
''''                            mOP_TP_VAL = 0
''''                        ElseIf mOP_TP_QTY < 0 Then
''''                            If mOP_TP_VAL > 0 Then
''''                                mOP_TP_VAL = -1 * mOP_TP_VAL
''''                            End If
''''                        End If
''''
''''                        With Temp06
''''                            .AddNew
''''                            .Fields("Part_No") = RstPart!Part_No
''''
''''                            .Fields("TB_OQty") = xMOP_TBQty
''''                            .Fields("TB_OVal") = xMOP_TBVal
''''                            .Fields("TP_OQty") = xMOP_TPQty
''''                            .Fields("TP_OVal") = xMOP_TPVal
''''
''''                            .Fields("RE_TB") = mRec_TB_Qty
''''                            .Fields("RE_TBV") = mRec_TB_Val
''''                            .Fields("RE_TP") = mRec_TP_Qty
''''                            .Fields("RE_TPV") = mRec_TP_Val
''''
''''                            .Fields("IS_TB") = mIss_TB_Qty
''''                            .Fields("IS_TBV") = mIss_TB_Val
''''                            .Fields("IS_TP") = mIss_TP_Qty
''''                            .Fields("IS_TPV") = mIss_TP_Val
''''
''''                            .Fields("TB_BQty") = mOP_TB_QTY
''''                            .Fields("TB_BVal") = mOP_TB_VAL
''''                            .Fields("TP_BQty") = mOP_TP_QTY
''''                            .Fields("TP_BVal") = mOP_TP_VAL
''''
''''                            .Fields("Net_Qty") = mOP_TB_QTY + mOP_TP_QTY
''''                            .Fields("Net_Val") = mOP_TB_VAL + mOP_TP_VAL
''''                            .Update
''''                        End With
''''                    End If
''''                    RstPart.MoveNext
''''                Loop
''''                Dim Counter As Integer
''''
''''   '------------------------------------------------------------------------------
''''                    ProgressBar1.Visible = False: Label2.Visible = False
''''                    If Temp06.RecordCount > 0 Then
''''                    Temp06.MoveFirst
''''                        For I = 1 To Temp06.RecordCount
''''                            Label1.CAPTION = "Transfering Opening Stock For Part : " & Temp06!Part_No
''''                            Label1.Refresh
''''                            If MRPYN = 1 Then
''''                              If Temp06!TB_BQty <> 0 Then
''''                                     Counter = Counter + 1
''''                                       'SP_StockTB Update
''''                                        DocID = RstDiv!Div_Code & PubSiteCode & PubSiteCode & " SXAO" & " SPOP" & Format(Counter, "00000000")
''''                                        GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
''''                                            "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,V_Rate,MRP_Rate," & _
''''                                            "Site_Code,U_Name,U_EntDt,U_AE) " & _
''''                                            "VALUES ('" & DocID & "',1,'SXAO'," & Format(Counter, "00000000") & ",#" & DateAdd("yyyy", -1, PubEndDate) & "#,'" & Temp06!Part_No & _
''''                                            "','" & PubSprCounterGodown & "',1,1," & Temp06!TB_BQty & "," & (Temp06!TB_BVal / Temp06!TB_BQty) & "," & Round((Temp06!TB_BVal), 2) & "," & (Temp06!TB_BVal / Temp06!TB_BQty) & ",0" & _
''''                                            ",'" & PubSiteCode & "','" & pubUName & "',#" & PubServerDate & "#,'A')"
''''                                End If
''''                                If Temp06!TP_BQty <> 0 Then
''''                                     Counter = Counter + 1
''''                                       'SP_StockTB Update
''''                                        DocID = RstDiv!Div_Code & PubSiteCode & PubSiteCode & " SXAO" & " SPOP" & Format(Counter, "00000000")
''''                                        GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
''''                                            "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,V_Rate,MRP_Rate," & _
''''                                            "Site_Code,U_Name,U_EntDt,U_AE) " & _
''''                                            "VALUES ('" & DocID & "',1,'SXAO'," & Format(Counter, "00000000") & ",#" & DateAdd("yyyy", -1, PubEndDate) & "#,'" & Temp06!Part_No & _
''''                                            "','" & PubSprCounterGodown & "',1,0," & Temp06!TP_BQty & "," & (Temp06!TP_BVal / Temp06!TP_BQty) & "," & Round((Temp06!TP_BVal), 2) & "," & (Temp06!TP_BVal / Temp06!TP_BQty) & ",0" & _
''''                                            ",'" & PubSiteCode & "','" & pubUName & "',#" & PubServerDate & "#,'A')"
''''                                End If
''''                            Else
''''                                If Temp06!TB_BQty <> 0 Then
''''                                     Counter = Counter + 1
''''                                       'SP_StockTB Update
''''                                        DocID = RstDiv!Div_Code & PubSiteCode & PubSiteCode & " SXAO" & " SPOP" & Format(Counter, "00000000")
''''                                        GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
''''                                            "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,V_Rate,MRP_Rate," & _
''''                                            "Site_Code,U_Name,U_EntDt,U_AE) " & _
''''                                            "VALUES ('" & DocID & "',1,'SXAO'," & Format(Counter, "00000000") & ",#" & DateAdd("yyyy", -1, PubEndDate) & "#,'" & Temp06!Part_No & _
''''                                            "','" & PubSprCounterGodown & "',0,1," & Temp06!TB_BQty & "," & (Temp06!TB_BVal / Temp06!TB_BQty) & "," & Round((Temp06!TB_BVal), 2) & "," & (Temp06!TB_BVal / Temp06!TB_BQty) & ",0" & _
''''                                            ",'" & PubSiteCode & "','" & pubUName & "',#" & PubServerDate & "#,'A')"
''''                                End If
''''                                If Temp06!TP_BQty <> 0 Then
''''                                     Counter = Counter + 1
''''                                       'SP_StockTB Update
''''                                        DocID = RstDiv!Div_Code & PubSiteCode & PubSiteCode & " SXAO" & " SPOP" & Format(Counter, "00000000")
''''                                        GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
''''                                            "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,V_Rate,MRP_Rate," & _
''''                                            "Site_Code,U_Name,U_EntDt,U_AE) " & _
''''                                            "VALUES ('" & DocID & "',1,'SXAO'," & Format(Counter, "00000000") & ",#" & DateAdd("yyyy", -1, PubEndDate) & "#,'" & Temp06!Part_No & _
''''                                            "','" & PubSprCounterGodown & "',0,0," & Temp06!TP_BQty & "," & (Temp06!TP_BVal / Temp06!TP_BQty) & "," & Round((Temp06!TP_BVal), 2) & "," & (Temp06!TP_BVal / Temp06!TP_BQty) & ",0" & _
''''                                            ",'" & PubSiteCode & "','" & pubUName & "',#" & PubServerDate & "#,'A')"
''''                                End If
''''                            End If
''''
''''                            Temp06.MoveNext
''''                        Next
''''                    End If
''''                    Set Temp06 = Nothing
''''                    Set TRec1 = Nothing
''''                    Set TRec2 = Nothing
''''                    Set RstStock = Nothing
''''                    Set RstStock2 = Nothing
''''                    Set RstPart = Nothing
''''        Next
''''        RstDiv.MoveNext
''''    Next
 
    'UPDATING VEH_STOCK TABLE
        Found = False
        Set RstVeh_Stk = GCN1.Execute("select * from Veh_Stock where len(Sal_DocID) = 0")
        Set RstVeh_Pur = GCN1.Execute("select * from Veh_Purch1 where DocID in (select Pur_DocID from Veh_Stock where len(Sal_DocID) = 0)")

        GCn.Execute ("Delete from Veh_Stock where len(Sal_DocID) = 0 and Pur_VType='V_OST'")
        GCn.Execute ("Delete from Veh_Purch1 where DocID not in (Select Pur_DocID from Veh_Stock)")

        Set RstVeh_Stk1 = GCn.Execute("Select ChassisNo as Chassis  from Veh_Stock where  Pur_VType='V_OST'")

        For i = 1 To RstVeh_Stk.RecordCount
            DocID = left(RstVeh_Stk!Pur_DocId, 3) & "V_OST" & "VOSTK" & Right(RstVeh_Stk!Pur_DocId, 8)
            If RstVeh_Stk1.RecordCount > 0 Then
                RstVeh_Stk1.MoveFirst
            End If
            For J = 1 To RstVeh_Stk1.RecordCount
                If RstVeh_Stk!ChassisNo = RstVeh_Stk1!Chassis Then
                    Found = True
                End If
                RstVeh_Stk1.MoveNext
            Next
            If Found = False Then
                GCn.Execute ("insert into veh_stock " & _
                    "(Pur_DocId,Pur_SrlNo,Pur_DocIDHelp,Pur_SiteCode,Pur_VType,Pur_VNO, " & _
                    "Chassis_RctDocNo ,Pur_VDate, Mfg_Month, Mfg_Yr, RSO_WORK,InDate, " & _
                    "MODEL,Godown,ChassisNo,EngineNo,VehSerialNo, " & _
                    "Srv_BookNo,RATE,vrate,Colour_Code,TAX_YN,SDM_STM_NO, " & _
                    "PBILL_NO,PBILL_DATE,PartyCode, U_Name, U_EntDt,U_AE, " & _
                    "OfftakeIncentiveSrlNo,OfftakeIncentive,TgtLinkIncentive,SubventionSrlNo,MfgShare) " & _
                    "values('" & DocID & "','" & RstVeh_Stk!Pur_SrlNo & "','" & UCase(Trim(DocID)) & "','" & PubSiteCode & PubSiteCode & "','V_OST'," & RstVeh_Stk!Pur_VNo & ", " & _
                    "" & RstVeh_Stk!Chassis_RctDocNo & "," & ConvertDate(RstVeh_Stk!Pur_VDate) & ",'" & RstVeh_Stk!Mfg_Month & "','" & RstVeh_Stk!Mfg_Yr & "'," & RstVeh_Stk!RSO_WORK & "," & ConvertDate(RstVeh_Stk!InDate) & ", " & _
                    "'" & RstVeh_Stk!Model & "','" & RstVeh_Stk!Godown & "','" & RstVeh_Stk!ChassisNo & "','" & RstVeh_Stk!EngineNo & "','" & RstVeh_Stk!VehSerialNo & "' , " & _
                    "'" & RstVeh_Stk!Srv_BookNo & "'," & RstVeh_Stk!Rate & "," & RstVeh_Stk!vrate & ",'" & RstVeh_Stk!Colour_Code & "'," & RstVeh_Stk!Tax_YN & ",'" & RstVeh_Stk!SDM_STM_NO & "', " & _
                    "'" & RstVeh_Stk!PBILL_NO & "'," & ConvertDate(RstVeh_Stk!PBILL_DATE) & ",'" & RstVeh_Stk!PartyCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'E', " & _
                    "'" & RstVeh_Stk!OfftakeIncentiveSrlNo & "'," & RstVeh_Stk!OfftakeIncentive & "," & _
                    "" & RstVeh_Stk!TgtLinkIncentive & ",'" & RstVeh_Stk!SubventionSrlNo & _
                    "', " & RstVeh_Stk!MfgShare & ")")

                GCn.Execute ("Delete from Veh_Purch1 where DocId='" & RstVeh_Stk!Pur_DocId & "'")

                GCn.Execute ("insert into Veh_Purch1( " & _
                    "DocID,DocIDHelp,Site_Code,V_Type,V_NO,V_Date, " & _
                    "PARTYCODE,PBILL_NO,PBILL_DATE,OBNO, " & _
                    "OBDate,BMS_CATEGORY,RSO_WORK,RSO_Code,DueDate, " & _
                    "GATE,GATEDATE,Form_Code,AMOUNT,Addition,Deduction,Exsice, " & _
                    "Tax_Per,TaxSur_Per,Tax_Amt,TaxSur_Amt,Misc_Amt, " & _
                    "Tot_Amount, U_Name, U_EntDt, U_AE,AcPostByU_Name,AcPostByU_EntDt,DrAcCode) " & _
                    "values( '" & DocID & "','" & UCase(Trim(DocID)) & "','" & PubSiteCode & PubSiteCode & "','V_OST'," & RstVeh_Pur!V_NO & "," & ConvertDate(RstVeh_Pur!V_DATE) & _
                    ",'" & RstVeh_Pur!PartyCode & "','" & RstVeh_Pur!PBILL_NO & "'," & ConvertDate(RstVeh_Pur!PBILL_DATE) & ",'" & RstVeh_Pur!OBNO & "'," & ConvertDate(RstVeh_Pur!OBDate) & _
                    ",'" & RstVeh_Pur!BMS_CATEGORY & "'," & RstVeh_Pur!RSO_WORK & ",'" & RstVeh_Pur!RSO_Code & "'," & ConvertDate(RstVeh_Pur!DueDate) & _
                    " ,'" & RstVeh_Pur!GATE & "'," & ConvertDate(RstVeh_Pur!GATEDATE) & ",'" & RstVeh_Pur!Form_Code & "','" & RstVeh_Pur!Amount & "'," & RstVeh_Pur!Addition & "," & RstVeh_Pur!Deduction & _
                    " , " & RstVeh_Pur!exsice & "," & RstVeh_Pur!Tax_Per & "," & RstVeh_Pur!TaxSur_Per & "," & RstVeh_Pur!Tax_Amt & "," & RstVeh_Pur!TaxSur_Amt & "," & RstVeh_Pur!Misc_Amt & _
                    " , " & RstVeh_Pur!Tot_Amount & ",'" & pubUName & "',#" & PubServerDate & "#,'A','" & RstVeh_Pur!AcPostByU_Name & "',#" & RstVeh_Pur!AcPostByU_EntDt & "#,'" & RstVeh_Pur!DrAcCode & "')")

                GCn.Execute ("Delete from Veh_Purch2 where DocId='" & RstVeh_Stk!Pur_DocId & "'")

                Set RstVeh_Pur1 = GCN1.Execute("select * from Veh_Purch2 where DocID ='" & RstVeh_Stk!Pur_DocId & "'")

                If RstVeh_Pur1.RecordCount > 0 Then
                    For J = 0 To RstVeh_Pur1.RecordCount
                          GCn.Execute ("insert into veh_purch2(DocId,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE, U_Name, U_EntDt, U_AE) " & _
                                "values('" & DocID & "'," & RstVeh_Pur1!Srl_No & ",'" & PubSiteCode & PubSiteCode & "','V_OST','" & RstVeh_Pur1!V_NO & "', " & _
                                "'" & RstVeh_Pur1!Prod_Code & "','" & RstVeh_Pur1!Trn_Type & "'," & RstVeh_Pur1!Qty & "," & RstVeh_Pur1!Rate & ",'" & pubUName & "',#" & PubServerDate & "#,'A')")
                    Next
                End If

            End If
            RstVeh_Stk.MoveNext
            RstVeh_Pur.MoveNext
            Found = False
        Next
    GCN1.CommitTrans
    'GCn.CommitTrans
MsgBox "All Balances has been updated to the new Year.Please reload the Software", vbInformation + vbOKOnly, "Update Balance"
End
DispErr:
    GCn.RollbackTrans
    MsgBox err.Description
End Function

Private Sub X_Val1(ByRef Temp06 As ADODB.Recordset, ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        xRate = 0
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        Exit Sub
    End If
    If xQty = TRec1!Qty Then
        TRec1.Fields("QTY") = 0
        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        TRec1.MoveNext
    ElseIf xQty < TRec1!Qty Then
        TRec1.Fields("QTY") = TRec1!Qty - xQty
        TRec1.Update
        
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
    ElseIf xQty > TRec1!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
            If TRec1!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec1!Qty <= TQty Then
                TQty = TQty - TRec1!Qty
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = left(xNARR, 25)
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TRec1!Qty
                        .Fields("Tb_Val") = TRec1!Qty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TRec1.Fields("QTY") = 0
                    TRec1.Update
                End If
            Else
                TRec1.Fields("QTY") = TRec1!Qty - TQty
                TRec1.Update
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TQty
                        .Fields("Tb_Val") = TQty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = TQty
                    .Fields("Tb_Val") = TQty * xRate
                    .Fields("Tb_BQty") = mOP_TB_QTY
                    .Fields("Tb_BVal") = mOP_TB_VAL
                    
                    .Fields("Is_Tp") = 0
                    .Fields("Tp_Val") = 0
                    .Fields("Tp_BQty") = 0
                    .Fields("Tp_BVal") = 0
                    .Update
                End With
            
            End If
        Loop
    End If
End Sub

Private Sub X_Val2(ByRef Temp06 As ADODB.Recordset, ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        xRate = 0
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = PrinID(RstStock!DocID)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        Exit Sub
    End If
    
    If xQty = TRec2!Qty Then
        TRec2.Fields("QTY") = 0
        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        TRec2.MoveNext
    ElseIf xQty < TRec2!Qty Then
        TRec2.Fields("QTY") = TRec2!Qty - xQty
        TRec2.Update
        
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                .Update
            End With
        End If
    ElseIf xQty > TRec2!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
            If TRec2!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec2!Qty <= TQty Then
                TQty = TQty - TRec2!Qty
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = xNARR
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TRec2!Qty
                        .Fields("Tp_Val") = TRec2!Qty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TRec2.Fields("QTY") = 0
                    TRec2.Update
                End If
            Else
                TRec2.Fields("QTY") = TRec2!Qty - TQty
                TRec2.Update
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TQty
                        .Fields("Tp_Val") = TQty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = 0
                    .Fields("Tb_Val") = 0
                    .Fields("Tb_BQty") = 0
                    .Fields("Tb_BVal") = 0
                    
                    .Fields("Is_Tp") = TQty
                    .Fields("Tp_Val") = TQty * xRate
                    .Fields("Tp_BQty") = mOP_TP_QTY
                    .Fields("Tp_BVal") = mOP_TP_VAL
                    .Update
                End With
            End If
        Loop
    End If
End Sub
Private Sub X_VAL11(ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
On Error GoTo Errloop
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        If mOP_TB_VAL <> 0 And mOP_TB_QTY <> 0 Then
            xRate = Round(mOP_TB_VAL / mOP_TB_QTY, 3)
        Else
            xRate = 0
        End If
            mOP_TB_QTY = mOP_TB_QTY - xQty
            mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
            mIss_TB_Qty = mIss_TB_Qty + xQty
            mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
          Exit Sub
    End If
    If xQty = TRec1Qty Then
        TRec1Qty = 0
'        TRec1!Qty = 0
'        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
        TRec1.MoveNext
        If TRec1.EOF = False Then
            TRec1Qty = TRec1!Qty
        End If
'    ElseIf xQty < TRec1!Qty Then
    ElseIf xQty < TRec1Qty Then
        TRec1Qty = TRec1Qty - xQty
'        TRec1!Qty = TRec1!Qty - xQty
'        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
'    ElseIf xQty  > TRec1!Qty Then
    ElseIf xQty > TRec1Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
'            If TRec1!Qty <= TQty Then
            If TRec1Qty <= TQty Then
'                TQty = TQty - TRec1!Qty
                TQty = TQty - TRec1Qty
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TRec1Qty 'TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + (TRec1Qty) '(TRec1!Qty)
                mIss_TB_Val = mIss_TB_Val + (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                TRec1Qty = 0
'                TRec1!Qty = 0
'                TRec1.Update
            Else
                TRec1Qty = TRec1Qty - TQty
'                TRec1!Qty = TRec1!Qty - TQty
'                TRec1.Update
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
            End If
            If TRec1.EOF = False Then
                TRec1Qty = TRec1!Qty
            End If
        Loop
    End If
Errloop:
     If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub X_VAL22(ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        xRate = 0
        
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        Exit Sub
    End If
'    If xQty = TRec2!Qty Then
    If xQty = TRec2Qty Then
        TRec2Qty = 0
'        TRec2!Qty = 0
'        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        TRec2.MoveNext
        If TRec2.EOF = False Then
            TRec2Qty = TRec2!Qty
        End If
'    ElseIf xQty < TRec2!Qty Then
    ElseIf xQty < TRec2Qty Then
        TRec2Qty = TRec2Qty - xQty
'        TRec2!Qty = TRec2!Qty - xQty
'        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
'    ElseIf xQty  > TRec2!Qty Then
    ElseIf xQty > TRec2Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
'            If TRec2!Qty <= TQty Then
            If TRec2Qty <= TQty Then
                TQty = TQty - TRec2Qty 'TRec2!Qty
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TRec2Qty 'TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2Qty * xRate)   '(TRec2!Qty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + (TRec2Qty)     '(TRec2!Qty)
                mIss_TP_Val = mIss_TP_Val + (TRec2Qty * xRate) '(TRec2!Qty * xRate)
                TRec2Qty = 0
'                TRec2!Qty = 0
'                TRec2.Update
            Else
                TRec2Qty = TRec2Qty - TQty
'                TRec2!Qty = TRec2!Qty - TQty
'                TRec2.Update
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
            End If
            If TRec2.EOF = False Then
                TRec2Qty = TRec2!Qty
            End If
        Loop
    End If
End Sub
Private Sub Command1_Click()
Call UpdateBalance
End Sub

Private Sub Command2_Click()
'Updating Finencial Accounting......
    Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset, Rst2 As ADODB.Recordset, mTrans As Boolean, DataPath$, SourcePathFA$
    Dim NewSubCode$
    SourcePathFA = Pub_DataPath & "\" & G_CompCn.Execute("Select OldPathFA from Company where CentralData_Path='" & PubCenDataPath & "'").Fields(0).Value
    GCn.BeginTrans
    G_FaCn.BeginTrans
    Set GFa_Cn = New ADODB.Connection
    With GFa_Cn
         .CursorLocation = adUseClient
         .Provider = "Microsoft.Jet.OLEDB.4.0"
         .ConnectionString = "Data Source=" & SourcePathFA & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
         .Open
         .BeginTrans
    End With
    Set RST1 = GFa_Cn.Execute("select Nature as NAT,GroupCode As MCODE, " & cIIF("Len(SUBGROUP.NewYrSubCode) > 0", " SUBGROUP.NewYrSubCode", " SUBGROUP.SubCode") & " AS PARTY_CODE,SUBGROUP.NAME,SUM(AMTCR)-SUM(AMTDR) AS BALANCE FROM LEDGER LEFT JOIN SUBGROUP ON LEDGER.SUBCODE=SUBGROUP.SUBCODE WHERE GroupNature IN ('A','L') GROUP BY GroupCode,NAME,subgroup.SUBCODE,Nature,GroupCode,SUBGROUP.NewYrSubCode")
    Label1.CAPTION = "Please Wait ! Updating Last Year FA Balances.."
    Label1.Refresh
    G_FaCn.Execute "delete from ledger where V_Type='F_AO'"
    G_FaCn.Execute "delete from ledgerM where V_Type='F_AO'"
    mCode = ""
    NewSubCode = ""
    mVNo = 0
    Do Until RST1.EOF
        NewSubCode = RST1!Party_code
        '***********Nra modfication for parties not coming*************
        'Set Rst = G_FaCn.Execute("select SG.SubCode,SG.Name from SubGroup as SG where SubCode&Name ='" & Rst1!Party_code + Rst1!Name & "'")
        'Set Rst2 = G_FaCn.Execute("select SG.Name from SubGroup as SG where Name ='" & Rst1!Name & "'")
        'If Rst2.RecordCount = 0 Then
        '    If Rst.RecordCount = 0 Then
        '        NewSubCode = PubSiteCode & IIf(PubFirmCode = "", "0", PubFirmCode) & Format(G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value, "000000")
        '        G_FaCn.Execute ("INSERT INTO SUBGROUP SELECT '" & NewSubCode & "' as AcID,Site_Code,'" & NewSubCode & "' AS SubCode,'" & PubFirmCode & "' as FirmCode,NamePrefix,Name,NameBiLang,NameHelp,GroupCode,GroupNature,Nature,AliasYN,ConPrefix," & _
        '                "ConPrefixBiLang,ConPerson,ConPersonBiLang,ConSuffix,Add1,Add1BiLang,Add2,Add2BiLang,Add3,Add3BiLang,CityCode,PIN,Phone,Mobile,FAX,EMail,Curr_Bal,CSTNo,LSTNo,PANNo,ITWARD_NO," & _
        '                "TDS_Catg,ActiveYN,Govt_YN,CreditLimit,CreditDays,FPrefix,FName,TAdd1,TAdd2,TAdd3,TCityCode,TPIN,TPhone,L_C,FB_Code,Religion,Party_Type,Transporter,Remark,U_Name,U_EntDt,U_AE,AREA,AcCode,AreaCode,Category,RC_No,CostCenterAppl,PhoneO,Transport FROM [" & SourcePathFA & "].SUBGROUP WHERE SUBCODE = '" & Rst1!Party_code & "'")
        '        GFa_Cn.Execute ("Update subgroup set NewYrSubCode ='" & NewSubCode & "' where subcode='" & Rst1!Party_code & "'")
        '        G_FaCn.Execute ("Update SubGroupCounter set SubGroupAcCode = SubGroupAcCode + 1")
        '    End If
        'End If
        '**************************************************************
        If mCode <> RST1!mCode Then
            mVNo = mVNo + 1
            mDocId = FaSetW(PubDivCode, 1) + PubSiteCode + PubSiteCode + FaSetW("F_AO", 5) + FaSetW(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate)))), 5) + FaSetN(STR(mVNo), 8)
            mCode = RST1!mCode
            mS_NO = 1
            G_FaCn.Execute ("INSERT INTO LEDGERM (DocId,V_Type,v_Prefix,V_No,Site_Code,V_Date,Narration,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & ",'F_AO'," & FaChk_Text(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate))))) & "," & mVNo & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("yyyy", -1, PubEndDate)) & ",'Opening Balance'," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
        End If
        If RST1!Balance < 0 Then       'DEBIT
            G_FaCn.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & "," & mS_NO & ",'F_AO'," & mVNo & "," & FaChk_Text(Trim(STR(Year(DateAdd("yyyy", -1, PubEndDate))))) & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("yyyy", -1, PubEndDate)) & "," & FaChk_Text(NewSubCode) & ",0," & Abs(RST1!Balance) & "," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
            FaCalCurrBal G_FaCn, NewSubCode, Abs(RST1!Balance), 0
            FaCalCurrBal GCn, NewSubCode, Abs(RST1!Balance), 0
        ElseIf RST1!Balance > 0 Then     'CREDIT
            G_FaCn.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & "," & mS_NO & ",'F_AO'," & mVNo & "," & FaChk_Text(Trim(STR(Year(DateAdd("yyyy", -1, PubEndDate))))) & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("yyyy", -1, PubEndDate)) & "," & FaChk_Text(NewSubCode) & "," & Abs(RST1!Balance) & ",0," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
            FaCalCurrBal G_FaCn, NewSubCode, 0, Abs(RST1!Balance)
            FaCalCurrBal GCn, NewSubCode, 0, Abs(RST1!Balance)
        End If
        mS_NO = mS_NO + 1
        RST1.MoveNext
    Loop
    
    G_FaCn.Execute ("update SubGroup set Curr_Bal=0 ")
    G_FaCn.Execute ("update SubGroupAlias set Curr_Bal=0 ")
    GCn.Execute ("update SubGroup set Curr_Bal=0 ")
    
    GSQL = "SELECT Ledger.SubCode,SUM(AmtCr-AmtDr) as CBal " & _
            "FROM Ledger left join SubGroup SG on SG.SubCOde=Ledger.SubCode " & _
            "group by Ledger.subcode,Name"
    Set Rst = G_FaCn.Execute(GSQL)
    If Rst.RecordCount > 0 Then
        Do While Rst.EOF = False
            GCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            G_FaCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            G_FaCn.Execute ("Update SubGroupAlias set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            Rst.MoveNext
        Loop
    End If
    'DataPath = Pub_DataPath & "\" & PubCenDataPath & "\Automan.mdb"
    'Set Rst = G_FaCn.Execute("select SG.SubCode from SubGroup as SG where SubCode not in (Select SubCode from [" & DataPath & ";pwd=dtman].SubGroup where FirmCode='" & PubFirmCode & "')")
    'If Rst.RecordCount > 0 Then
    '   Do Until Rst.EOF
    G_FaCn.Execute ("Delete from subgroupAlias")
    G_FaCn.Execute ("INSERT INTO SUBGROUPAlias SELECT * FROM SUBGROUP")
    
    GCn.Execute ("Delete from subgroup")
    GCn.Execute ("INSERT INTO SUBGROUP SELECT * FROM [" & PubFADataPath & "].SUBGROUP")
    '        Rst.MoveNext
    '   Loop
    GCn.Execute ("Drop Table SubGroupAlias")
    GCn.Execute ("Select SubGroup.* into SubGroupAlias from SubGroup")
    'End If
    Set Rst = Nothing
    Set RST1 = Nothing
    Set Rst2 = Nothing
    Screen.MousePointer = vbDefault
    GCn.CommitTrans
    G_FaCn.CommitTrans
    GFa_Cn.CommitTrans
    MsgBox "All Balances has been updated to the new Year.Please reload the Software", vbInformation + vbOKOnly, "Update Ballence"
    End
End Sub

