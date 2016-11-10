Attribute VB_Name = "TmpTable"
Option Explicit
Public Enum TmpTableDef
    SprRstTmp = 0
    LabRstTmp = 1
    VehHisRstTmp = 2
End Enum


Public Function TmpTemp06(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "Date", adDate, , adFldIsNullable
    .Fields.Append "V_No", adVarChar, 18, adFldIsNullable
    .Fields.Append "V_Type", adVarChar, 5, adFldIsNullable
    
    .Fields.Append "Inv_No", adVarChar, 18, adFldIsNullable
    .Fields.Append "Inv_Date", adVarChar, 35, adFldIsNullable
        
    .Fields.Append "Part_No", adVarChar, 22, adFldIsNullable
    .Fields.Append "Part_Name", adVarChar, 40, adFldIsNullable
    .Fields.Append "Party_Name", adVarChar, 50, adFldIsNullable
    .Fields.Append "Job_Age", adVarChar, 7, adFldIsNullable
    .Fields.Append "Rate", adDouble, 12, adFldIsNullable
    
    .Fields.Append "TB_OQty", adDouble, 12, adFldIsNullable
    .Fields.Append "TB_OVal", adDouble, 12, adFldIsNullable
    .Fields.Append "Re_TB", adDouble, 12, adFldIsNullable
    .Fields.Append "Re_TBV", adDouble, 12, adFldIsNullable
    .Fields.Append "Is_TB", adDouble, 12, adFldIsNullable
    .Fields.Append "Is_TBV", adDouble, 12, adFldIsNullable
    .Fields.Append "Tb_BQty", adDouble, 12, adFldIsNullable
    .Fields.Append "TB_BVal", adDouble, 12, adFldIsNullable
    .Fields.Append "TB_Val", adDouble, 12, adFldIsNullable
    
    .Fields.Append "TP_OQty", adDouble, 12, adFldIsNullable
    .Fields.Append "TP_OVal", adDouble, 12, adFldIsNullable
    .Fields.Append "Re_TP", adDouble, 12, adFldIsNullable
    .Fields.Append "Re_TPV", adDouble, 12, adFldIsNullable
    .Fields.Append "Is_TP", adDouble, 12, adFldIsNullable
    .Fields.Append "Is_TPV", adDouble, 12, adFldIsNullable
    .Fields.Append "TP_BQty", adDouble, 12, adFldIsNullable
    .Fields.Append "TP_BVal", adDouble, 12, adFldIsNullable
    .Fields.Append "TP_Val", adDouble, 12, adFldIsNullable
    
    .Fields.Append "Net_Qty", adDouble, 12, adFldIsNullable
    .Fields.Append "Net_Val", adDouble, 12, adFldIsNullable
    
    .Fields.Append "MovePer", adDouble, 12, adFldIsNullable
    .Fields.Append "mType", adVarChar, 12, adFldIsNullable
    
    .Fields.Append "Narr", adVarChar, 30, adFldIsNullable
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With
Set TmpTemp06 = Rst
End Function

Public Function TmpChqPrn(RstRep As ADODB.Recordset) As ADODB.Recordset
With RstRep
    .Fields.Append "FVType", adChar, 40, adFldIsNullable
    .Fields.Append "FVPrefix", adChar, 40, adFldIsNullable
    .Fields.Append "FVDate", adDate, , adFldIsNullable
    .Fields.Append "FChqDate", adDate, , adFldIsNullable
    .Fields.Append "FClgDate", adDate, , adFldIsNullable
    .Fields.Append "FPartyName", adChar, 40, adFldIsNullable
    .Fields.Append "FVNo", adDouble, 19, adFldIsNullable
    .Fields.Append "FDrAmt", adDouble, 19, adFldIsNullable
    .Fields.Append "FCrAmt", adDouble, 19, adFldIsNullable
    .Fields.Append "FChqNo", adChar, 40, adFldIsNullable
    .Fields.Append "FNarration", adChar, 255, adFldIsNullable
    .Fields.Append "BankName", adChar, 80, adFldIsNullable
    .Fields.Append "Bal_As_Bank", adDouble, 19, adFldIsNullable
    .Fields.Append "Bal_as_Book", adDouble, 19, adFldIsNullable
    .Fields.Append "Status", adChar, 40, adFldIsNullable
    .Fields.Append "MVDate", adDate, , adFldIsNullable
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TmpChqPrn = RstRep

    
        

End Function

Public Function TmpTRec1(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "Date", adDate, , adFldIsNullable
    .Fields.Append "Part_No", adVarChar, 22, adFldIsNullable
    .Fields.Append "Rate", adDouble, 12, adFldIsNullable
    .Fields.Append "Qty", adDouble, 12, adFldIsNullable
    .Fields.Append "Cost", adDouble, 12, adFldIsNullable
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With
Set TmpTRec1 = Rst
End Function

Public Sub CreaTabRstRep(RstRep As ADODB.Recordset, TableType As TmpTableDef)
Select Case TableType
    Case VehHisRstTmp
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "RegNo", adChar, 14, adFldIsNullable
            .Fields.Append "CustName", adChar, 40, adFldIsNullable
            .Fields.Append "Add1", adChar, 40, adFldIsNullable
            .Fields.Append "Add2", adChar, 40, adFldIsNullable
            .Fields.Append "Add3", adChar, 40, adFldIsNullable
            .Fields.Append "DOSale", adDate, 15, adFldIsNullable
            .Fields.Append "Chassis", adChar, 20, adFldIsNullable
            .Fields.Append "Engine", adChar, 25, adFldIsNullable
            .Fields.Append "GBNo", adChar, 20, adFldIsNullable
            .Fields.Append "RANo", adChar, 20, adFldIsNullable
            
            .Fields.Append "CardNo", adChar, 8, adFldIsNullable
            .Fields.Append "JobDocID", adChar, 21, adFldIsNullable
            .Fields.Append "JobNo", adChar, 13, adFldIsNullable
            .Fields.Append "Job_Date", adDate, 7, adFldIsNullable
            .Fields.Append "JobCloseDate", adDate, 7, adFldIsNullable
            .Fields.Append "AtKMsHrs", adDouble, 12, adFldIsNullable
            .Fields.Append "Serv_Type", adVarChar, 2, adFldIsNullable
            .Fields.Append "RecBy_MechCode", adVarChar, 4, adFldIsNullable
            .Fields.Append "RecBy_MechName", adVarChar, 40, adFldIsNullable
            .Fields.Append "NetLab_Amt", adDouble, 12, adFldIsNullable
            .Fields.Append "NetSpr_Amt", adDouble, 12, adFldIsNullable
            .Fields.Append "Lab_Code", adChar, 6, adFldIsNullable
            .Fields.Append "Lab_Done", adChar, 40, adFldIsNullable
            .Fields.Append "Prob_Code", adChar, 6, adFldIsNullable
            .Fields.Append "Prob_Reported", adChar, 40, adFldIsNullable
            .Fields.Append "DocIdReq", adChar, 21, adFldIsNullable
            .Fields.Append "Part_No", adVarChar, 22, adFldIsNullable
            .Fields.Append "Part_Name", adVarChar, 40, adFldIsNullable
            .Fields.Append "Purpose", adChar, 1, adFldIsNullable
            .Fields.Append "Rate", adDouble, 12, adFldIsNullable
            .Fields.Append "Qty", adDouble, 12, adFldIsNullable
            .Fields.Append "Amount", adDouble, 12, adFldIsNullable
            .Fields.Append "ObservBy_Super", adVarChar, 255, adFldIsNullable
            .Fields.Append "ActionBy_Super", adVarChar, 255, adFldIsNullable
            .Fields.Append "PhoneOff", adVarChar, 25, adFldIsNullable
            .Fields.Append "PhoneResi", adVarChar, 25, adFldIsNullable
            .Fields.Append "Mobile", adVarChar, 12, adFldIsNullable
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
        End With
    Case SprRstTmp
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "Srl", adInteger, 8, adFldIsNullable
            .Fields.Append "PartNo", adVarChar, 22, adFldIsNullable
            .Fields.Append "Descrip", adVarChar, 40, adFldIsNullable
            .Fields.Append "MRP_YN", adTinyInt, 1, adFldIsNullable
            .Fields.Append "Tax_YN", adTinyInt, 1, adFldIsNullable
            .Fields.Append "Qty", adDouble, 12, adFldIsNullable
            .Fields.Append "MrpRate", adDouble, 12, adFldIsNullable
            .Fields.Append "Rate", adDouble, 12, adFldIsNullable
            .Fields.Append "DiscPer", adDouble, 12, adFldIsNullable
            .Fields.Append "DiscAmt", adDouble, 12, adFldIsNullable
            .Fields.Append "TBAmt", adDouble, 12, adFldIsNullable
            .Fields.Append "TPAmt", adDouble, 12, adFldIsNullable
            .Fields.Append "LandAmt", adDouble, 12, adFldIsNullable
            .Fields.Append "Amt", adDouble, 12, adFldIsNullable
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
    Case LabRstTmp
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "Srl", adInteger, 4, adFldIsNullable
            .Fields.Append "Lab_Code", adVarChar, 6, adFldIsNullable
            .Fields.Append "LabName", adVarChar, 40, adFldIsNullable
            .Fields.Append "Hrs_Taken", adSingle, 6, adFldIsNullable
            .Fields.Append "Lab_Rate", adSingle, 6, adFldIsNullable
            .Fields.Append "Hrs_War", adSingle, 6, adFldIsNullable
            .Fields.Append "War_Lab_Rate", adSingle, 10, adFldIsNullable
            .Fields.Append "LabourAmt", adDouble, 12, adFldIsNullable
            .Fields.Append "MechName", adVarChar, 40, adFldIsNullable
            .Fields.Append "ContName", adVarChar, 40, adFldIsNullable
            .Fields.Append "ContractIssue_Date", adDate, 7, adFldIsNullable
            .Fields.Append "ContractRecd_Date", adDate, 7, adFldIsNullable
            .Fields.Append "ContractAmt", adSingle, 10, adFldIsNullable
            .Fields.Append "Contract_Remarks", adVarChar, 20, adFldIsNullable
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
End Select
End Sub

Public Function TMPCR(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_TYPE", adChar, 3
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "PARTY_CODE", adInteger, 11
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "Party_Bill_No", adChar, 15
    .Fields.Append "Dhara", adChar, 10
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TMPCR = Rst
End Function
Public Function TMPDR(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_TYPE", adChar, 3
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "PARTY_CODE", adInteger, 11
    .Fields.Append "DR", adDouble, 19, 5
    .Fields.Append "AMTADJ", adDouble, 19, 5
    .Fields.Append "Party_Bill_No", adChar, 15
    .Fields.Append "Dhara", adChar, 10
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TMPDR = Rst
End Function

Public Function TMPTYPE(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "START_DATE", adDate
    .Fields.Append "END_DATE", adDate
    .Fields.Append "TYPE", adChar, 15
    .Fields.Append "ACC", adInteger, 11
    .Fields.Append "BALANCE", adDouble, 19, 5
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TMPTYPE = Rst
End Function

Public Function CASHTMP1(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_NO", adChar, 18
    .Fields.Append "V_TYPE", adChar, 5
    .Fields.Append "V_SNO", adInteger, 11
    .Fields.Append "V_ADD", adChar, 5
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "ADJAMT", adDouble, 19, 5
    .Fields.Append "SUBCODE", adInteger, 11
    .Fields.Append "NAME", adChar, 40
    .Fields.Append "ADJQTY", adDouble, 19, 5
    .Fields.Append "VTYPE", adChar, 5
    .Fields.Append "VNO", adChar, 10
    .Fields.Append "VADD", adChar, 5
    .Fields.Append "VSNO", adInteger, 6
    .Fields.Append "VAL", adChar, 2
    .Fields.Append "NARRATION1", adChar, 255
    .Fields.Append "NARRATION2", adChar, 255
    .Fields.Append "NARRATION3", adChar, 255
    .Fields.Append "NARRATION4", adLongVarChar, 920
    .Fields.Append "NARRATION5", adLongVarChar, 920
    .Fields.Append "NAME1", adChar, 40
    .Fields.Append "ADDRESS1", adChar, 50
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set CASHTMP1 = Rst
End Function

Public Function ADTMP1(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_NO", adChar, 18 'Integer, 10
    .Fields.Append "V_TYPE", adChar, 5
    .Fields.Append "V_SNO", adInteger, 11
    .Fields.Append "V_ADD", adChar, 5
    .Fields.Append "VNO", adInteger, 10
    .Fields.Append "VTYPE", adChar, 5
    .Fields.Append "VSNO", adInteger, 11
    .Fields.Append "VADD", adChar, 5
    .Fields.Append "V_NO1", adInteger, 18
    .Fields.Append "V_TYPE1", adChar, 5
    .Fields.Append "V_SNO1", adInteger, 11
    .Fields.Append "V_ADD1", adChar, 5
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "ADJAMT", adDouble, 19, 5
    .Fields.Append "SUBCODE", adChar, 10
    .Fields.Append "NAME", adChar, 50
    .Fields.Append "ADJQTY", adDouble, 19, 5
    .Fields.Append "VAL", adChar, 2
    .Fields.Append "NAME1", adChar, 50
    .Fields.Append "ADDRESS1", adChar, 150
    .Fields.Append "NARRATION1", adChar, 150
    .Fields.Append "NARRATION2", adChar, 150
    .Fields.Append "NARRATION3", adLongVarChar, 920 '460
    .Fields.Append "SNAME", adChar, 15
    .Fields.Append "CITY_NAME", adChar, 50
    .Fields.Append "GRNAME", adChar, 50
    .Fields.Append "GRCODE", adChar, 4
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set ADTMP1 = Rst
End Function

Public Function ADTMP10(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "V_TYPE", adChar, 2
    .Fields.Append "V_SNO", adInteger, 11
    .Fields.Append "V_ADD", adChar, 1
    .Fields.Append "VDATE", adDate
    .Fields.Append "VNO", adInteger, 10
    .Fields.Append "VTYPE", adChar, 2
    .Fields.Append "VSNO", adInteger, 11
    .Fields.Append "VADD", adChar, 1
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "ADJAMT", adDouble, 19, 5
    .Fields.Append "SUBCODE", adInteger, 11
    .Fields.Append "NAME", adChar, 40
    .Fields.Append "ADJQTY", adDouble, 19, 5
    .Fields.Append "VAL", adChar, 2
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set ADTMP10 = Rst
End Function
Public Function AGETMP(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "ACC_NAME", adChar, 40
    .Fields.Append "DEBIT1", adDouble, 19, 5
    .Fields.Append "DEBIT2", adDouble, 19, 5
    .Fields.Append "DEBIT3", adDouble, 19, 5
    .Fields.Append "DEBIT4", adDouble, 19, 5
    .Fields.Append "DEBIT5", adDouble, 19, 5
    .Fields.Append "DEBIT6", adDouble, 19, 5
    .Fields.Append "DEBIT", adDouble, 19, 5
    .Fields.Append "TOTALDR", adDouble, 19, 5
    .Fields.Append "CREDIT", adDouble, 19, 5
    .Fields.Append "ANAME", adChar, 40
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set AGETMP = Rst
End Function
        
Public Function MIS(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "MONTH", adInteger, 10
    .Fields.Append "V_DATE", adDate
    .Fields.Append "SALE_PCS", adInteger, 10
    .Fields.Append "SALE_MTR", adDouble, 19, 5
    .Fields.Append "SALE_VALUE", adDouble, 19, 5
    .Fields.Append "SALE_RET_PCS", adInteger, 10
    .Fields.Append "SALE_RET_MTR", adDouble, 19, 5
    .Fields.Append "SALE_RET_VALUE", adDouble, 19, 5
    .Fields.Append "PURCH_PCS", adInteger, 10
    .Fields.Append "PURCH_MTR", adDouble, 19, 5
    .Fields.Append "PURCH_VALUE", adDouble, 19, 5
    .Fields.Append "PURCH_RET_PCS", adInteger, 10
    .Fields.Append "PURCH_RET_MTR", adDouble, 19, 5
    .Fields.Append "PURCH_RET_VALUE", adDouble, 19, 5
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set MIS = Rst
End Function
Public Function TEMP_PARTY_LEDGER(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_TYPE", adChar, 3
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "PARTY_CODE", adInteger, 10
    .Fields.Append "PARTY_NAME", adChar, 50
    .Fields.Append "ADD1", adChar, 40
    .Fields.Append "ADD2", adChar, 40
    .Fields.Append "ADD3", adChar, 40
    .Fields.Append "ADD4", adChar, 40
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "DR", adDouble, 19, 5
    .Fields.Append "BALANCE", adDouble, 19, 5
    .Fields.Append "BAL_TYPE", adChar, 3
    .Fields.Append "DR_STAT", adChar, 1
    .Fields.Append "interest", adDouble, 19, 5
    .Fields.Append "DR1", adDouble, 19, 5
    .Fields.Append "ADDR", adChar, 100
    .Fields.Append "Party_Bill_No", adChar, 15
    .Fields.Append "Party_Bill_Date", adDate
    .Fields.Append "DAYS", adInteger, 10
    .Fields.Append "AREACODE", adInteger, 10
    .Fields.Append "AREANAME", adChar, 25
    .Fields.Append "CITYCODE", adInteger, 10
    .Fields.Append "CITYNAME", adChar, 21
    .Fields.Append "BROKER_CODE", adInteger, 10
    .Fields.Append "BROKER_NAME", adChar, 50
    .Fields.Append "SL", adInteger
    .Fields.Append "Dhara", adChar, 10
    .Fields.Append "Dhara1", adChar, 10
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TEMP_PARTY_LEDGER = Rst
End Function

Public Function TEMP_PARTY_LEDGER1(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_TYPE", adChar, 3
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "PARTY_CODE", adInteger, 10
    .Fields.Append "PARTY_NAME", adChar, 50
    .Fields.Append "ADD1", adChar, 40
    .Fields.Append "ADD2", adChar, 40
    .Fields.Append "ADD3", adChar, 40
    .Fields.Append "ADD4", adChar, 40
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "DR", adDouble, 19, 5
    .Fields.Append "BALANCE", adDouble, 19, 5
    .Fields.Append "BAL_TYPE", adChar, 3
    .Fields.Append "DR_STAT", adChar, 1
    .Fields.Append "interest", adDouble, 19, 5
    .Fields.Append "Party_Bill_No", adChar, 15
    .Fields.Append "DAYS", adInteger, 10
    .Fields.Append "Dhara", adChar, 10
    .Fields.Append "Dhara1", adChar, 10
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TEMP_PARTY_LEDGER1 = Rst
End Function

Public Function TMP_PARTY_LEDGER(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "PARTY_CODE", adInteger, 10
    .Fields.Append "PARTY_NAME", adChar, 50
    .Fields.Append "ADD1", adChar, 40
    .Fields.Append "ADD2", adChar, 40
    .Fields.Append "ADD3", adChar, 40
    .Fields.Append "ADD4", adChar, 40
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_TYPE", adChar, 3
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "DR", adDouble, 19, 5
    .Fields.Append "V_DATE1", adDate
    .Fields.Append "V_TYPE1", adChar, 3
    .Fields.Append "V_NO1", adInteger, 18
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "DR_STAT", adChar, 1
    .Fields.Append "interest", adDouble, 19, 5
    .Fields.Append "interest1", adDouble, 19, 5
    .Fields.Append "Party_Bill_No", adChar, 15
    .Fields.Append "Party_Bill_No1", adChar, 15
    .Fields.Append "DAYS", adInteger, 10
    .Fields.Append "DAYS1", adInteger, 10
    .Fields.Append "AREACODE", adInteger, 10
    .Fields.Append "AREANAME", adChar, 25
    .Fields.Append "CITYCODE", adInteger, 10
    .Fields.Append "CITYNAME", adChar, 21
    .Fields.Append "BROKER_CODE", adInteger, 10
    .Fields.Append "BROKER_NAME", adChar, 50
    .Fields.Append "comp_name", adChar, 50
    .Fields.Append "Dhara", adChar, 10
    .Fields.Append "Dhara1", adChar, 10
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TMP_PARTY_LEDGER = Rst
End Function
Public Function TEMP_PCS(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "Sr_No", adInteger, 10
    .Fields.Append "Sr_No1", adInteger, 10
    .Fields.Append "Pcs", adDouble, 19, 5
    .Fields.Append "Meters", adDouble, 19, 5
    .Fields.Append "DISC_NAME1", adChar, 15
    .Fields.Append "DISC_CODE1", adInteger, 10
    .Fields.Append "DISC_RATE1", adDouble, 19, 5
    .Fields.Append "DISC_CODE2", adInteger, 10
    .Fields.Append "DISC_RATE2", adDouble, 19, 5
    .Fields.Append "DISC_NAME2", adChar, 15
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TEMP_PCS = Rst
End Function
    
Public Function TEMPSTAT(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "CR", adDouble, 19, 5
    .Fields.Append "CL_DR", adDouble, 19, 5
    .Fields.Append "GRP_CODE", adInteger, 10
    .Fields.Append "S_NAME", adChar, 40
    .Fields.Append "CL_CR", adDouble, 19, 5
    .Fields.Append "G_NAME", adChar, 40
    .Fields.Append "LAST_DR", adDouble, 19, 5
    .Fields.Append "LAST_CR", adDouble, 19, 5
    .Fields.Append "DRAMT", adDouble, 19, 5
    .Fields.Append "GROUP_CODE", adInteger, 10
    .Fields.Append "GR_CODE", adInteger, 10
    .Fields.Append "G_TYPE", adChar, 15
    .Fields.Append "V_TYPE", adChar, 2
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "NET_AMT", adDouble, 19, 5
    .Fields.Append "SALE_RETU", adDouble, 19, 5
    .Fields.Append "RECD", adDouble, 19, 5
    .Fields.Append "V_DATE", adDate
    .Fields.Append "CODE", adInteger, 10
    .Fields.Append "NAME", adChar, 40
    .Fields.Append "OP_DR", adDouble, 19, 5
    .Fields.Append "OP_CR", adDouble, 19, 5
    .Fields.Append "DR", adDouble, 19, 5
    .Fields.Append "OUR_GROUPCODE", adChar, 4
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TEMPSTAT = Rst
End Function
Public Function TM_LEDGER(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate
    .Fields.Append "V_TYPE", adChar, 2
    .Fields.Append "V_NO", adInteger, 18
    .Fields.Append "V_ADD", adChar, 1
    .Fields.Append "V_SNO", adInteger, 10
    .Fields.Append "SUBCODE", adInteger, 10
    .Fields.Append "NAME", adChar, 40
    .Fields.Append "aMOUNT", adDouble, 19, 5
    .Fields.Append "CONTRASUB", adInteger, 10
    .Fields.Append "VAL", adChar, 2
    .Fields.Append "CHQ_NO", adChar, 15
    .Fields.Append "CHQ_DATE", adDate
    .Fields.Append "CLG_DATE", adDate
    .Fields.Append "NARRATION", adChar, 255
    .Fields.Append "CONTRANAME", adChar, 40
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TM_LEDGER = Rst
End Function
    
Public Function TMP(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_NUM", adChar, 15
    .Fields.Append "V_DATE", adDate
    .Fields.Append "ACC_TYPE", adChar, 2
    .Fields.Append "ACC_CODE", adInteger, 10
    .Fields.Append "ACC_CODENAME", adChar, 40
    .Fields.Append "ACC", adInteger, 10
    .Fields.Append "ACC_NAME", adChar, 50
    .Fields.Append "NEGBAL", adDouble, 19, 5
    .Fields.Append "V_TYPE", adChar, 2
    .Fields.Append "NARRATION", adChar, 255
    .Fields.Append "CHQ_DATE", adDate
    .Fields.Append "CHQ_NO", adChar, 15
    .Fields.Append "CREDIT", adDouble, 19, 5
    .Fields.Append "DEBIT", adDouble, 19, 5
    .Fields.Append "INTRST", adDouble, 19, 5
    .Fields.Append "CLG_DATE", adDate
    .Fields.Append "BALNAME", adChar, 20
    .Fields.Append "START_DATE", adDate
    .Fields.Append "END_DATE", adDate
    .Fields.Append "ADD1", adChar, 40
    .Fields.Append "ADD2", adChar, 40
    .Fields.Append "CITY", adChar, 21
    .Fields.Append "PIN", adChar, 6
    .Fields.Append "PRICE", adDouble, 19, 5
    .Fields.Append "REPORT_NAME", adChar, 50
    .Fields.Append "OTHERS", adDouble, 19, 5
    .Fields.Append "NAMT", adDouble, 19, 5
    .Fields.Append "AGE", adDouble, 19, 5
    .Fields.Append "inv_NO", adChar, 15
    .Fields.Append "v_no", adInteger, 18
    .Fields.Append "V_SNO", adInteger, 10
    .Fields.Append "V_ADD", adChar, 1
    .Fields.Append "CONTRA_CODE", adInteger, 19, 5
    .Fields.Append "CONTRA_NAME", adChar, 40
    .Fields.Append "CONTRA_SUBCODE", adInteger, 10
    .Fields.Append "CONTRA_SUBNAME", adChar, 50
    .Fields.Append "RS_IN_WORDS", adChar, 100
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TMP = Rst
End Function

Public Function TMP1(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "ACC", adInteger, 10
    .Fields.Append "V_ADD", adChar, 1
    .Fields.Append "BALANCE", adDouble, 19, 5
    .Fields.Append "START_DATE", adDate
    .Fields.Append "END_DATE", adDate
    .Fields.Append "INT_RATE", adChar, 20
    .Fields.Append "TYPE", adChar, 3
    .Fields.Append "TOT_INT", adDouble, 19, 5
    .Fields.Append "BALANCE_TYPE", adChar, 40
    .Fields.Append "ADD1", adChar, 40
    .Fields.Append "ADD2", adChar, 40
    .Fields.Append "CITY", adChar, 21
    .Fields.Append "PIN", adChar, 6
    .Fields.Append "INTRST", adDouble, 19, 5
    .Fields.Append "BAL1", adDouble, 19, 5
    .Fields.Append "BAL_TYPE1", adChar, 40
    .Fields.Append "DATE1", adDate
    .Fields.Append "ITEM_CODE", adInteger, 10
    .Fields.Append "REPORT_NAME", adChar, 50
    .Fields.Append "EX_AMT", adDouble, 19, 5
    .Fields.Append "G_TOT", adDouble, 19, 5
    .Fields.Append "EX_RATE", adDouble, 19, 5
    .Fields.Append "TOTAL", adDouble, 19, 5
    .Fields.Append "DIS_AMT", adDouble, 19, 5
    .Fields.Append "STAX_AMT", adDouble, 19, 5
    .Fields.Append "SUR_AMT", adDouble, 19, 5
    .Fields.Append "VENDOR", adChar, 10
    .Fields.Append "ORDER_NO", adChar, 15
    .Fields.Append "PART_NO", adChar, 10
    .Fields.Append "CUST_NAME", adChar, 40
    .Fields.Append "DESP_NO", adInteger, 10
    .Fields.Append "DESP_DT", adDate
    .Fields.Append "TOT_WORDS", adChar, 100
    .Fields.Append "PLA", adInteger, 10
    .Fields.Append "RGA", adInteger, 10
    .Fields.Append "RGC", adInteger, 10
    .Fields.Append "INV_TIME", adChar, 8
    .Fields.Append "TOT_WRDS1", adChar, 100
    .Fields.Append "TAR_NAME", adChar, 50
    .Fields.Append "EX_ECC_NO", adChar, 25
    .Fields.Append "TAX_PER", adDouble, 19, 5
    .Fields.Append "REM1", adChar, 100
    .Fields.Append "REM2", adChar, 100
    .Fields.Append "AMD_DATE", adDate
    .Fields.Append "V_DATE", adDate
    .Fields.Append "EXMP_NOTIFY", adChar, 40
    .Fields.Append "D3", adInteger, 10
    .Fields.Append "AG_FORM", adChar, 40
    .Fields.Append "V_TYPE", adChar, 2
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set TMP1 = Rst
End Function

