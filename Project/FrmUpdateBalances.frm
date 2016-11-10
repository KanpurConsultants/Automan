VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmUpdateOpeningBalances 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Opening  Balances"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   4410
      TabIndex        =   3
      Top             =   1230
      Width           =   1140
   End
   Begin MSComctlLib.ProgressBar Prg 
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   1650
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection Criteria"
      Height          =   1575
      Left            =   75
      TabIndex        =   1
      Top             =   45
      Width           =   2055
      Begin VB.OptionButton OptVehicle 
         Caption         =   "Vehicle"
         Height          =   420
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptSpare 
         Caption         =   "Spare"
         Height          =   420
         Left            =   90
         TabIndex        =   5
         Top             =   645
         Width           =   1215
      End
      Begin VB.OptionButton OptFa 
         Caption         =   "Financial Account"
         Height          =   420
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Width           =   1785
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Note - This Utility Updates the Balances of Last year to Current Year So please take backup before proceeding."
      Height          =   810
      Left            =   2235
      TabIndex        =   0
      Top             =   150
      Width           =   3195
   End
End
Attribute VB_Name = "FrmUpdateOpeningBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsUpdated As Boolean


Private Function UpdateBalance()
'On Error GoTo DispErr
Dim K As Integer, MRPYN As Integer
Dim Condstr$, SourcePath$, DocID$, mQry$, sQry$, CondStr1$

Dim RsSprStkOld As ADODB.Recordset, RsTemp As ADODB.Recordset
Dim RstVeh_Stk As ADODB.Recordset
Dim RstVeh_Pur As ADODB.Recordset
Dim RstVeh_Stk1 As ADODB.Recordset
Dim RstVeh_Pur1 As ADODB.Recordset
Dim I As Long, j As Long
Dim GCN1 As ADODB.Connection
Dim Found As Boolean
Dim PBIncrVal As Long
Dim Counter As Long
Dim TempStock As Double, TempAmount As Double
Dim MyPart_No, MyTax_YN, MyMRP_YN, MyDiv_Code, MyEof As Byte
Dim GFa_Cn As New ADODB.Connection
Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset, Rst2 As ADODB.Recordset, mTrans As Boolean, DataPath$, SourcePathFA$
Dim NewSubCode$, mCode$, mDocId$
Dim mVNo As Long
Dim mS_NO As Long
Dim xPurDocId As String


Dim mQty As Double
Dim mFifoCost As Double
Dim mReqQty As Double
Dim mRate As Double
Dim mAmount As Double





'TAKE BACKUP
'If PubBackEnd = "A" Then MDIForm1.DataBackup


        
    
    Set GCN1 = New ADODB.Connection
    With GCN1
        .CursorLocation = adUseClient
        .Mode = adModeWrite
        If PubBackEnd = "A" Then
            If OptFa Then
                SourcePath = Pub_DataPath & "\" & G_CompCn.Execute("Select OldPathfa from Company where CentralData_Path='" & PubCenDataPath & "'").Fields(0).Value
            Else
                SourcePath = Pub_DataPath & "\" & G_CompCn.Execute("Select OldPath from Company where CentralData_Path='" & PubCenDataPath & "'").Fields(0).Value
            End If
                .Provider = "Microsoft.Jet.OLEDB.4.0"
               .ConnectionString = "Data Source=" & SourcePath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
           Else
               SourcePath = G_CompCn.Execute("Select OldPath from Company where CentralData_Path='" & PubCenDataPath & "'").Fields(0).Value
               .ConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=" & SourcePath & ";Data Source=" & PubServerName
           End If
           .Open
    End With
    





'UPDATING SP_STOCK TABLE
    If OptSpare Then
        GCN1.BeginTrans
        GCn.BeginTrans
        
            If PubBackEnd = "A" Then
                If StrCmp(left(PubComp_Name, 3), "jmk") Then
                    mQry = "SELECT Left(S.DocId,1) As Div_Code, S.Part_No, sum(S.Qty_Rec)-Sum(S.Qty_Iss)+Sum(S.Qty_Ret) As mQty, Tax_Yn,1 As Mrp_Yn, " & _
                           "Max(Rate) As mRate, Max(Amount) As Amount, Max(S.V_Rate) As V_Rate " & _
                           "From sp_stock S " & _
                           "WHERE iif(S.v_Date=" & ConvertDate(Format(DateAdd("YYYY", -2, PubEndDate), "dd/MMM/yyyy")) & ",S.V_Type='SXAO',IIF(S.V_Date>=" & ConvertDate(Format(DateAdd("YYYY", -1, PubStartDate), "dd/MMM/yyyy")) & ",S.V_Type<>'SXAO')) " & _
                           "Group By Left(S.DocId,1), Part_No, Tax_Yn " & _
                           "Having sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret)<>0 "
                Else
                    mQry = "SELECT Left(S.DocId,1) As Div_Code, S.Part_No, sum(S.Qty_Rec)-Sum(S.Qty_Iss)+Sum(S.Qty_Ret) As mQty, Tax_Yn,1 as Mrp_Yn, " & _
                           "Max(Rate) As mRate, Max(Amount) As Amount, Max(S.V_Rate) As V_Rate " & _
                           "From sp_stock S " & _
                           "WHERE iif(S.v_Date=" & ConvertDate(Format(DateAdd("YYYY", -2, PubEndDate), "dd/MMM/yyyy")) & ",S.V_Type='SXAO',IIF(S.V_Date>=" & ConvertDate(Format(DateAdd("YYYY", -1, PubStartDate), "dd/MMM/yyyy")) & " and S.V_Date<" & ConvertDate(PubStartDate) & ",S.V_Type<>'SXAO')) " & _
                           "Group By Left(S.DocId,1), Part_No, Tax_Yn " & _
                           "Having sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret)<>0 "
                End If
            Else
                mQry = "SELECT Left(S.DocId,1) As Div_Code,SubString(S.DocID,3,1) Site_Code, S.Part_No, sum(S.Qty_Rec)-Sum(S.Qty_Iss)+Sum(S.Qty_Ret) As mQty, Tax_Yn,1 as Mrp_Yn, " & _
                       "Max(Rate) As mRate, Max(Amount) As Amount, Max(S.V_Rate) As V_Rate " & _
                       "From sp_stock S With (NoLock) " & _
                       "WHERE (S.V_Type = (Case When S.v_Date=" & ConvertDate(Format(DateAdd("YYYY", -2, PubEndDate), "dd/MMM/yyyy")) & " Then 'SXAO' End) OR S.V_Type <> (Case When S.V_Date>=" & ConvertDate(Format(DateAdd("YYYY", -1, PubStartDate), "dd/MMM/yyyy")) & " And s.V_Date < " & ConvertDate(PubStartDate) & " Then 'SXAO' End)) " & _
                       "Group By Left(S.DocId,1), SubString(S.DocID,3,1), Part_No, Tax_Yn " & _
                       "Having sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret)<>0 "
            End If
            
            Set RsSprStkOld = New ADODB.Recordset
            If RsSprStkOld.State <> 0 Then RsSprStkOld.Close
            RsSprStkOld.CursorLocation = adUseClient
            RsSprStkOld.Open mQry, GCN1, adOpenDynamic, adLockBatchOptimistic
            
            GCn.Execute ("Delete From SP_Stock Where V_Type='SXAO'    ")
                                        
            With RsSprStkOld
                If .RecordCount > 0 Then
                    Prg.Value = 0
                    Label1.CAPTION = "Please Wait ! Updating Last Year Spare Stock Balances.."
                    Label1.Refresh
                    
                    Do While Not .EOF

                        Set RsTemp = GCN1.Execute("Select Part_No, V_Date, Qty_Rec, V_Rate as Rate " & _
                                              "From Sp_Stock S With (NoLock)" & _
                                              "Where S.Part_No='" & !Part_No & "' And S.Tax_Yn = " & !Tax_YN & " " & _
                                              "And Left(S.DocId,1)='" & !Div_Code & "' And SubString(S.DocID,3,1)='" & !Site_Code & "' And Qty_Rec>0 And S.V_Date<" & ConvertDate(PubStartDate) & " " & _
                                              "Order by V_Date Desc")
                                                                    
                        mQty = 0
                        mReqQty = 0
                        mFifoCost = 0
                        Debug.Print RsTemp.RecordCount
                        Do Until RsTemp.EOF
                            If mQty < VNull(!mQty) Then
                                mReqQty = IIf((mQty + VNull(RsTemp!Qty_Rec)) > VNull(!mQty), VNull(!mQty) - mQty, RsTemp!Qty_Rec)
                                mQty = mQty + VNull(RsTemp!Qty_Rec)
                                
                                mFifoCost = mFifoCost + (mReqQty * VNull(RsTemp!Rate))
                                RsTemp.MoveNext
                            Else
                                Exit Do
                            End If
                        Loop
                        
                        
                        mAmount = CDbl(mFifoCost)
                        mRate = mFifoCost / VNull(!mQty)
                        

                        Counter = Counter + 1
                        DocID = !Div_Code & !Site_Code & !Site_Code & " SXAO" & " SPOP" & Format(Counter, "00000000")
                        
                        GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                            "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,V_Rate,MRP_Rate," & _
                            "Site_Code,U_Name,U_EntDt,U_AE) " & _
                            "VALUES ('" & DocID & "',1,'SXAO'," & Format(Counter, "00000000") & "," & ConvertDate(DateAdd("yyyy", -1, PubEndDate)) & ",'" & !Part_No & _
                            "','" & PubSprCounterGodown & "'," & !MRP_YN & "," & !Tax_YN & "," & VNull(!mQty) & "," & mRate & "," & Round(mAmount, 2) & ", " & mRate & ", " & mRate & " " & _
                            ",'" & !Site_Code & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
                        .MoveNext
                        
                        If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                    Loop
                End If
            End With
        GCn.CommitTrans
        GCN1.CommitTrans
        
        MsgBox "All Spare Balances has been Updated to the new Year.", vbInformation + vbOKOnly, "Update Balance"
    End If
    
    
    
    
    
    
    
    
    If OptVehicle Then
        Found = False
        
        
        GCN1.BeginTrans
        GCn.BeginTrans
            Set RstVeh_Stk = GCN1.Execute("select * from Veh_Stock where len(Sal_DocID) = 0")
            Set RstVeh_Pur = GCN1.Execute("select * from Veh_Purch1 where DocID in (select Pur_DocID from Veh_Stock where len(Sal_DocID) = 0)")
    
            GCn.Execute ("Delete from Veh_Stock where (len(Sal_DocID) = 0 and Pur_VType='V_OST') Or Sal_VDate < '" & PubStartDate & "'")
            GCn.Execute ("Delete from Veh_Purch1 where DocID not in (Select Pur_DocID from Veh_Stock)")
            GCn.Execute ("Delete from Veh_Purch2 where DocID not in (Select Pur_DocID from Veh_Stock)")
    
            'Set RstVeh_Stk1 = GCn.Execute("Select ChassisNo as Chassis  from Veh_Stock where  Pur_VType='V_OST'")
            Set RstVeh_Stk1 = GCn.Execute("Select ChassisNo as Chassis, Pur_DocId, MFG_YR  from Veh_Stock ")
            Label1.CAPTION = "Please Wait ! Updating Last Year Vehicle Balances.."
            Label1.Refresh
    
            Prg = 0
            For I = 1 To RstVeh_Stk.RecordCount
                'If UCase(RstVeh_Stk.Fields("ChassisNo")) = "MAT448069D3P22024" Then MsgBox ""
                DocID = left(RstVeh_Stk!Pur_DocId, 3) & "V_OST" & " " & XNull(RstVeh_Stk!Mfg_Yr) & Right(RstVeh_Stk!Pur_DocId, 8)
                If RstVeh_Stk1.RecordCount > 0 Then
                    RstVeh_Stk1.MoveFirst
                End If
                For j = 1 To RstVeh_Stk1.RecordCount
                    If RstVeh_Stk!ChassisNo = RstVeh_Stk1!Chassis Then
                        Found = True
                        xPurDocId = RstVeh_Stk1!Pur_DocId
                    End If
                    RstVeh_Stk1.MoveNext
                Next
                    RstVeh_Pur.MoveFirst
                    RstVeh_Pur.FIND "DocId='" & RstVeh_Stk!Pur_DocId & "'"
                    If Found = False Then
                        GCn.Execute ("insert into veh_stock " & _
                            "(Pur_DocId,Pur_SrlNo,Pur_DocIDHelp,Pur_SiteCode,Pur_VType,Pur_VNO, " & _
                            "Chassis_RctDocNo ,Pur_VDate, Mfg_Month, Mfg_Yr, RSO_WORK,InDate, " & _
                            "MODEL,Godown,ChassisNo,EngineNo,VehSerialNo, " & _
                            "Srv_BookNo,RATE,vrate,Colour_Code,TAX_YN,SDM_STM_NO, " & _
                            "PBILL_NO,PBILL_DATE,PartyCode, U_Name, U_EntDt,U_AE, " & _
                            "OfftakeIncentiveSrlNo,OfftakeIncentive,TgtLinkIncentive,SubventionSrlNo,MfgShare) " & _
                            "values('" & DocID & "','" & RstVeh_Stk!Pur_SrlNo & "','" & UCase(Trim(DocID)) & "','" & PubSiteCode & PubSiteCode & "','V_OST'," & RstVeh_Stk!Pur_VNo & ", " & _
                            "'" & RstVeh_Stk!Chassis_RctDocNo & "'," & ConvertDate(RstVeh_Stk!Pur_VDate) & ",'" & RstVeh_Stk!Mfg_Month & "','" & RstVeh_Stk!Mfg_Yr & "','" & RstVeh_Stk!RSO_WORK & "'," & ConvertDate(RstVeh_Stk!InDate) & ", " & _
                            "'" & RstVeh_Stk!Model & "','" & RstVeh_Stk!Godown & "','" & RstVeh_Stk!ChassisNo & "','" & RstVeh_Stk!EngineNo & "','" & RstVeh_Stk!VehSerialNo & "' , " & _
                            "'" & RstVeh_Stk!Srv_BookNo & "'," & RstVeh_Stk!Rate & "," & RstVeh_Stk!vrate & ",'" & RstVeh_Stk!Colour_Code & "'," & RstVeh_Stk!Tax_YN & ",'" & RstVeh_Stk!SDM_STM_NO & "', " & _
                            "'" & RstVeh_Stk!PBILL_NO & "'," & ConvertDate(RstVeh_Stk!PBILL_DATE) & ",'" & RstVeh_Stk!PartyCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'E', " & _
                            "'" & RstVeh_Stk!OfftakeIncentiveSrlNo & "'," & RstVeh_Stk!OfftakeIncentive & "," & _
                            "" & RstVeh_Stk!TgtLinkIncentive & ",'" & RstVeh_Stk!SubventionSrlNo & _
                            "', " & RstVeh_Stk!MfgShare & ")")
        
                            GCn.Execute ("Delete from Veh_Purch1 where DocId='" & RstVeh_Stk!Pur_DocId & "'")
                            'GCn.Execute ("Delete from Veh_Purch1 where DocId='" & DocID & "'")
                            
                            GCn.Execute ("insert into Veh_Purch1( " & _
                                "DocID,DocIDHelp,Site_Code,V_Type,V_NO,V_Date, " & _
                                "PARTYCODE,PBILL_NO,PBILL_DATE,OBNO, " & _
                                "OBDate,BMS_CATEGORY,RSO_WORK,RSO_Code,DueDate, " & _
                                "GATE,GATEDATE,Form_Code,AMOUNT,Addition,Deduction,Exsice, " & _
                                "Tax_Per,TaxSur_Per,Tax_Amt,TaxSur_Amt,Misc_Amt, " & _
                                "Tot_Amount, U_Name, U_EntDt, U_AE,AcPostByU_Name,AcPostByU_EntDt,DrAcCode) " & _
                                "values( '" & DocID & "','" & UCase(Trim(DocID)) & "','" & PubSiteCode & PubSiteCode & "','V_OST'," & RstVeh_Pur!V_NO & "," & ConvertDate(RstVeh_Pur!V_DATE) & _
                                ",'" & RstVeh_Pur!PartyCode & "','" & RstVeh_Pur!PBILL_NO & "'," & ConvertDate(RstVeh_Pur!PBILL_DATE) & ",'" & RstVeh_Pur!OBNO & "'," & ConvertDate(RstVeh_Pur!OBDate) & _
                                ",'" & RstVeh_Pur!BMS_CATEGORY & "','" & RstVeh_Pur!RSO_WORK & "','" & RstVeh_Pur!RSO_Code & "'," & ConvertDate(RstVeh_Pur!DueDate) & _
                                " ,'" & RstVeh_Pur!GATE & "'," & ConvertDate(RstVeh_Pur!GATEDATE) & ",'" & RstVeh_Pur!Form_Code & "','" & RstVeh_Pur!Amount & "'," & RstVeh_Pur!Addition & "," & RstVeh_Pur!Deduction & _
                                " , " & RstVeh_Pur!exsice & "," & RstVeh_Pur!Tax_Per & "," & RstVeh_Pur!TaxSur_Per & "," & RstVeh_Pur!Tax_Amt & "," & RstVeh_Pur!TaxSur_Amt & "," & RstVeh_Pur!Misc_Amt & _
                                " , " & RstVeh_Pur!Tot_Amount & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & RstVeh_Pur!AcPostByU_Name & "'," & ConvertDate(RstVeh_Pur!AcPostByU_EntDt) & ",'" & RstVeh_Pur!DrAcCode & "')")
            
                            GCn.Execute ("Delete from Veh_Purch2 where DocId='" & RstVeh_Stk!Pur_DocId & "'")
                            'GCn.Execute ("Delete from Veh_Purch2 where DocId='" & DocID & "'")
                            
                            Set RstVeh_Pur1 = GCN1.Execute("select * from Veh_Purch2 where DocID ='" & RstVeh_Stk!Pur_DocId & "'")
            
                            If RstVeh_Pur1.RecordCount > 0 Then
                                Do Until RstVeh_Pur1.EOF
                                    GCn.Execute ("insert into veh_purch2(DocId,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE, U_Name, U_EntDt, U_AE) " & _
                                          "values('" & DocID & "'," & RstVeh_Pur1!Srl_No & ",'" & PubSiteCode & PubSiteCode & "','V_OST','" & RstVeh_Pur1!V_NO & "', " & _
                                          "'" & RstVeh_Pur1!Prod_Code & "','" & RstVeh_Pur1!Trn_Type & "'," & RstVeh_Pur1!Qty & "," & RstVeh_Pur1!Rate & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
                                    RstVeh_Pur1.MoveNext
                                Loop
                            End If
                    Else
                        If UCase(left(PubComp_Name, 5)) = "UJWAL" Or UCase(left(PubComp_Name, 4)) = "ENAR" Or UCase(left(PubComp_Name, 6)) = "J.M.A." Then
                            GCn.Execute "Update Veh_Stock Set Rate=" & VNull(RstVeh_Stk!Rate) & ", VRate=" & VNull(RstVeh_Stk!vrate) & " Where Pur_DocId='" & xPurDocId & "'"
                            
                            GCn.Execute ("Update Veh_Purch1 Set " & _
                                "AMOUNT=" & RstVeh_Pur!Amount & ",Addition=" & RstVeh_Pur!Addition & ",Deduction=" & RstVeh_Pur!Deduction & ", " & _
                                "Tax_Per=" & RstVeh_Pur!Tax_Per & ",TaxSur_Per=" & RstVeh_Pur!TaxSur_Per & ",Tax_Amt=" & RstVeh_Pur!Tax_Amt & ",TaxSur_Amt=" & RstVeh_Pur!TaxSur_Amt & ",Misc_Amt=" & RstVeh_Pur!Misc_Amt & ", " & _
                                "Tot_Amount=" & RstVeh_Pur!Tot_Amount & ", Exsice=" & VNull(RstVeh_Pur!exsice) & " " & _
                                "Where DocId='" & xPurDocId & "'")
                
                                GCn.Execute ("Delete from Veh_Purch2 where DocId='" & xPurDocId & "'")
                
                                Set RstVeh_Pur1 = GCN1.Execute("select * from Veh_Purch2 where DocID ='" & RstVeh_Stk!Pur_DocId & "'")
                
                                If RstVeh_Pur1.RecordCount > 0 Then
                                    Do Until RstVeh_Pur1.EOF
                                        GCn.Execute ("insert into veh_purch2(DocId,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE, U_Name, U_EntDt, U_AE) " & _
                                              "values('" & xPurDocId & "'," & RstVeh_Pur1!Srl_No & ",'" & PubSiteCode & PubSiteCode & "','V_OST','" & RstVeh_Pur1!V_NO & "', " & _
                                              "'" & RstVeh_Pur1!Prod_Code & "','" & RstVeh_Pur1!Trn_Type & "'," & RstVeh_Pur1!Qty & "," & RstVeh_Pur1!Rate & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
                                                
                                        RstVeh_Pur1.MoveNext
                                    Loop
                                End If
                        End If
                    End If
                    xPurDocId = ""
                RstVeh_Stk.MoveNext
                RstVeh_Pur.MoveNext
                Found = False
                
                If Round(Prg.Value) < 100 Then Prg.Value = (RstVeh_Stk.AbsolutePosition / RstVeh_Stk.RecordCount) * 100
            Next
        GCn.CommitTrans
        GCN1.CommitTrans
        
        MsgBox "All Vehicle Balances has been Updated to the new Year.", vbInformation + vbOKOnly, "Update Balance"
    End If
    
    
    
    
    
    
    
    
    If OptFa Then
            
        Set GFa_Cn = New ADODB.Connection
        With GFa_Cn
            .CursorLocation = adUseClient
            If PubBackEnd = "A" Then
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & SourcePath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=" & SourcePath & ";Data Source=" & PubServerName
            End If
            
            .Open
        End With
        
        GCn.BeginTrans
        G_FaCn.BeginTrans
                
            Set RST1 = GFa_Cn.Execute("Select Nature as NAT,GroupCode As MCODE, " & cIIF("Len(SUBGROUP.NewYrSubCode) > 0", "SUBGROUP.NewYrSubCode", "SUBGROUP.SubCode") & " AS PARTY_CODE,SUBGROUP.NAME,SUM(AMTCR)-SUM(AMTDR) AS BALANCE FROM LEDGER LEFT JOIN SUBGROUP ON LEDGER.SUBCODE=SUBGROUP.SUBCODE WHERE GroupNature IN ('A','L') And Ledger.V_Date < " & ConvertDate(PubStartDate) & "  GROUP BY GroupCode,NAME,subgroup.SUBCODE,Nature,GroupCode,SUBGROUP.NewYrSubCode ")
            Label1.CAPTION = "Please Wait ! Updating Last Year FA Balances.."
            Label1.Refresh
            G_FaCn.Execute "delete from ledger where V_Type='F_AO'"
            G_FaCn.Execute "delete from ledgerM where V_Type='F_AO'"
            mCode = ""
            NewSubCode = ""
            mVNo = 0
            Prg = 0
            Do Until RST1.EOF
                NewSubCode = RST1!Party_code
                If mCode <> RST1!mCode Then
                    mVNo = mVNo + 1
                    mDocId = FaSetW(PubDivCode, 1) + PubSiteCode + PubSiteCode + FaSetW("F_AO", 5) + FaSetW(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate)))), 5) + FaSetN(STR(mVNo), 8)
                    mCode = RST1!mCode
                    mS_NO = 1
                    G_FaCn.Execute ("INSERT INTO LEDGERM (DocId,V_Type,v_Prefix,V_No,Site_Code,V_Date,Narration,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & ",'F_AO'," & FaChk_Text(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate))))) & "," & mVNo & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("yyyy", -1, PubEndDate)) & ",'Opening Balance'," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
                End If
                If RST1!Balance < 0 Then       'DEBIT
                    G_FaCn.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & "," & mS_NO & ",'F_AO'," & mVNo & "," & FaChk_Text(Trim(STR(Year(DateAdd("yyyy", -1, PubEndDate))))) & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("yyyy", -1, PubEndDate)) & "," & FaChk_Text(NewSubCode) & ",0," & Abs(RST1!Balance) & "," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
                    'FaCalCurrBal G_FaCn, NewSubCode, Abs(RST1!Balance), 0
                    'FaCalCurrBal GCn, NewSubCode, Abs(RST1!Balance), 0
                ElseIf RST1!Balance > 0 Then     'CREDIT
                    G_FaCn.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & "," & mS_NO & ",'F_AO'," & mVNo & "," & FaChk_Text(Trim(STR(Year(DateAdd("yyyy", -1, PubEndDate))))) & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("yyyy", -1, PubEndDate)) & "," & FaChk_Text(NewSubCode) & "," & Abs(RST1!Balance) & ",0," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
                    'FaCalCurrBal G_FaCn, NewSubCode, 0, Abs(RST1!Balance)
                    'FaCalCurrBal GCn, NewSubCode, 0, Abs(RST1!Balance)
                End If
                mS_NO = mS_NO + 1
                RST1.MoveNext
                If Round(Prg.Value) < 100 Then Prg.Value = (RST1.AbsolutePosition / RST1.RecordCount) * 100
            Loop
        
        G_FaCn.CommitTrans
        GCn.CommitTrans
        
        
        G_FaCn.BeginTrans
        GCn.BeginTrans
        
        G_FaCn.Execute ("update SubGroup set Curr_Bal=0 ")
        If PubBackEnd = "A" Then GCn.Execute ("update SubGroup set Curr_Bal=0 ")
        
        
        GSQL = "SELECT Ledger.SubCode,SUM(AmtCr-AmtDr) as CBal " & _
                "FROM Ledger left join SubGroup SG on SG.SubCOde=Ledger.SubCode " & _
                "group by Ledger.subcode,Name"
        Set Rst = G_FaCn.Execute(GSQL)
        If Rst.RecordCount > 0 Then
            Prg = 0
            Do While Rst.EOF = False
                If PubBackEnd = "A" Then GCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
                G_FaCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
                Rst.MoveNext
                If Round(Prg.Value) < 100 Then Prg.Value = (Rst.AbsolutePosition / Rst.RecordCount) * 100
            Loop
        End If
        
        
        G_FaCn.Execute ("Delete from subgroupAlias")
        G_FaCn.Execute ("INSERT INTO SUBGROUPAlias SELECT * FROM SUBGROUP")
        If PubBackEnd = "A" Then
            GCn.Execute ("Delete from subgroup")
            GCn.Execute ("INSERT INTO SUBGROUP SELECT * FROM [" & PubFADataPath & "].SUBGROUP")
            GCn.Execute ("Drop Table SubGroupAlias")
            GCn.Execute ("Select SubGroup.* into SubGroupAlias from SubGroup")
        End If
    
    
        Set Rst = Nothing
        Set RST1 = Nothing
        Set Rst2 = Nothing
        Screen.MousePointer = vbDefault
        GCn.CommitTrans
        G_FaCn.CommitTrans

        MsgBox "All FA Balances has been updated to the new Year.", vbInformation + vbOKOnly, "Update Ballence"
                        
    End If

Exit Function
DispErr:
    GCn.RollbackTrans
    If OptFa Then G_FaCn.RollbackTrans
    MsgBox err.Description
End Function

Private Sub CmdStart_Click()
    IsUpdated = True
    UpdateBalance
    Label1.CAPTION = "Note: All Closing Balances from last year will be Set as Opening Balances in Current Year. Please Take Backup Before Proceeding."
    Label1.Refresh
End Sub

Private Sub Form_Load()
    OptVehicle = True
    Label1.CAPTION = "Note: All Closing Balances from last year will be Set as Opening Balances in Current Year. Please Take Backup Before Proceeding."
    Label1.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsUpdated Then
        MsgBox "Software is now Reloading"
        End
    End If
End Sub
