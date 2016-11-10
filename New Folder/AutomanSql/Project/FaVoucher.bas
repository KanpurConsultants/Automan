Attribute VB_Name = "FaVoucher"
Option Explicit
Public Sub Voucher(VType As String, V_NO As Long, V_ADD As String, FGrid1 As MSHFlexGrid)
Dim RstVType As ADODB.Recordset
Set RstVType = G_FaCn.Execute("SELECT NCAT FROM VOUCHER_TYPE WHERE V_TYPE='" & VType & "'")
If RstVType.RecordCount > 0 Then
    FGrid1.Tag = "VOUCHER"
    Select Case RstVType!NCat
        Case "CNT", "JV", "PMT", "RCT"
            FaVrEnt.Show
            FaVrEnt.Tag = " "
            FaVrEnt.FindMove FGrid1.TextMatrix(FGrid1.Row, 12)
        Case "V_SB"
            frmVehSale.Show
            frmVehSale.SEARCHBACK FGrid1.TextMatrix(FGrid1.Row, 12)
        Case "G_ABR", "G_ACR", "G_BBP", "G_BCP", "G_CRN", "G_DRN", "G_TLR"
            frmCustRect.Show
            frmCustRect.SEARCHBACK FGrid1.TextMatrix(FGrid1.Row, 12)
        Case "OPBAL"
            frmSubGroup.Show
            frmSubGroup.SEARCHBACK FGrid1.TextMatrix(FGrid1.Row, 11)
    End Select
End If
Set RstVType = Nothing
End Sub
Public Sub OpeningDiff(VType As String, V_NO As Long, V_ADD As String, FGrid1 As MSHFlexGrid)
Dim RstVType As ADODB.Recordset
Set RstVType = G_FaCn.Execute("SELECT NCAT FROM VOUCHER_TYPE WHERE V_TYPE='" & VType & "'")
If RstVType.RecordCount > 0 Then
    FGrid1.Tag = "OPDIFF"
    Select Case RstVType!NCat
        Case "OPBAL"
            frmSubGroup.Show
            frmSubGroup.SEARCHBACK FGrid1.TextMatrix(FGrid1.Row, 12)
    End Select
End If
Set RstVType = Nothing
End Sub
Public Function VoucherTypeCheck(VType As String) As Boolean
VoucherTypeCheck = False
    If G_FaCn.Execute("SELECT COUNT(*) FROM Ledger WHERE V_Type=" & FaChk_Text(VType)).Fields(0) > 0 Then VoucherTypeCheck = True: Exit Function
    If G_FaCn.Execute("SELECT COUNT(*) FROM LedgerM WHERE V_Type=" & FaChk_Text(VType)).Fields(0) > 0 Then VoucherTypeCheck = True: Exit Function
End Function
Public Sub FaVoucherPrintingModule(ByRef frm As Object, VRep As Object, mDocId As String)
Dim Rs_in_Words As String, Rs_in_AMT As Double, RST1 As ADODB.Recordset, I As Integer, X11
On Error GoTo err1
If frm.Opt2(1).Value = True Then
    If FaIsValid(frm.DataCombo2(0), "From Voucher No. ") = False Then Exit Sub
    If FaIsValid(frm.DataCombo2(1), "To Voucher No. ") = False Then Exit Sub
    Set RST1 = G_FaCn.Execute("SELECT VOUCHER_TYPE.DESCRIPTION,NAME,LEDGER.V_Type,LEDGER.V_No,LEDGER.v_Prefix,LEDGER.V_Date,LEDGER.V_SNo,LEDGER.AmtCr,LEDGER.AmtDr,LEDGER.Chq_No,LEDGER.Chq_Date,LEDGER.Narration,LEDGERM.Narration AS NARRMAIN,LEDGERM.DocId,LEDGERM.v_Prefix, Ledger.AddBy  FROM ((LEDGER LEFT JOIN LEDGERM ON LEDGERM.DOCID=LEDGER.DOCID) LEFT JOIN SUBGROUP ON  SUBGROUP.SUBCODE = LEDGER.SUBCODE ) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_tYPE WHERE LEDGER.V_TYPE='" & frm.DataCombo3.BoundText & "' AND LEDGER.V_NO BETWEEN " & frm.DataCombo2(0) & " AND " & frm.DataCombo2(1) & " order by ledger.v_no,LEDGER.V_SNo")
ElseIf frm.Opt2(2).Value = True Then
    If FaIsValid(frm.Text2, "For Date ") = False Then Exit Sub
    Set RST1 = G_FaCn.Execute("SELECT VOUCHER_TYPE.DESCRIPTION,NAME,LEDGER.V_Type,LEDGER.V_No,LEDGER.v_Prefix,LEDGER.V_Date,LEDGER.V_SNo,LEDGER.AmtCr,LEDGER.AmtDr,LEDGER.Chq_No,LEDGER.Chq_Date,LEDGER.Narration,LEDGERM.Narration AS NARRMAIN,LEDGERM.DocId,LEDGERM.v_Prefix, Ledger.AddBy  FROM ((LEDGER LEFT JOIN LEDGERM ON LEDGERM.DOCID=LEDGER.DOCID)LEFT JOIN SUBGROUP ON  SUBGROUP.SUBCODE = LEDGER.SUBCODE ) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_tYPE WHERE VOUCHER_TYPE.Category='FA' AND LEDGER.V_date=" & FaConvertDate(frm.Text2) & " order by ledger.v_type,ledger.v_no,LEDGER.V_SNo")
Else
    'Set RST1 = G_FaCn.Execute("SELECT VOUCHER_TYPE.DESCRIPTION,NAME,LEDGER.V_Type,LEDGER.V_No,LEDGER.v_Prefix,LEDGER.V_Date,LEDGER.V_SNo,LEDGER.AmtCr,LEDGER.AmtDr,LEDGER.Chq_No,LEDGER.Chq_Date,LEDGER.Narration,LEDGERM.Narration AS NARRMAIN,LEDGERM.DocId,LEDGERM.v_Prefix FROM ((LEDGER LEFT JOIN LEDGERM ON LEDGERM.DOCID=LEDGER.DOCID)LEFT JOIN SUBGROUP ON  SUBGROUP.SUBCODE = LEDGER.SUBCODE ) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_tYPE WHERE LEDGER.V_TYPE='" & frm.TxtVtYpe(0).Tag & "' AND LEDGER.V_NO=" & frm.TxtVno(0) & " And Left(Ledger.DocId,1)='" & PubDivCode & "' order by ledger.v_no,LEDGER.V_SNo")
    Set RST1 = G_FaCn.Execute("SELECT VOUCHER_TYPE.DESCRIPTION,NAME,LEDGER.V_Type,LEDGER.V_No,LEDGER.v_Prefix,LEDGER.V_Date,LEDGER.V_SNo,LEDGER.AmtCr,LEDGER.AmtDr,LEDGER.Chq_No,LEDGER.Chq_Date,LEDGER.Narration,LEDGERM.Narration AS NARRMAIN,LEDGERM.DocId,LEDGERM.v_Prefix, Ledger.AddBy FROM ((LEDGER LEFT JOIN LEDGERM ON LEDGERM.DOCID=LEDGER.DOCID)LEFT JOIN SUBGROUP ON  SUBGROUP.SUBCODE = LEDGER.SUBCODE ) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_tYPE WHERE Ledger.DocId='" & mDocId & "' order by ledger.v_no,LEDGER.V_SNo")
End If
If RST1.RecordCount = 0 Then MsgBox "No record found to Print", vbInformation, frm.CAPTION: Exit Sub
X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaJVCHR.TTX", True)
Set rpt = VRep
rpt.Database.SetDataSource RST1
rpt.ReadRecords
FaReport_View rpt, 0, frm.CAPTION, True
Set RST1 = Nothing
Exit Sub
err1:    MsgBox err.Description, vbCritical, frm.CAPTION
End Sub
Public Sub FaRectPrintingModule(ByRef frm As Object, VRep As Object)
Dim Rs_in_Words As String, Rs_in_AMT As Double, RST1 As ADODB.Recordset, I As Integer, X11
On Error GoTo err1
'If frm.Opt2(1).Value = True Then
'    If FaIsValid(frm.DataCombo2(0), "From Voucher No. ") = False Then Exit Sub
'    If FaIsValid(frm.DataCombo2(1), "To Voucher No. ") = False Then Exit Sub
'    Set RST1 = G_FaCn.Execute("SELECT VOUCHER_TYPE.DESCRIPTION,NAME,LEDGER.V_Type,LEDGER.V_No,LEDGER.v_Prefix,LEDGER.V_Date,LEDGER.V_SNo,LEDGER.AmtCr,LEDGER.AmtDr,LEDGER.Chq_No,LEDGER.Chq_Date,LEDGER.Narration,LEDGERM.Narration AS NARRMAIN,LEDGERM.DocId,LEDGERM.v_Prefix FROM ((LEDGER LEFT JOIN LEDGERM ON LEDGERM.DOCID=LEDGER.DOCID) LEFT JOIN SUBGROUP ON  SUBGROUP.SUBCODE = LEDGER.SUBCODE ) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_tYPE WHERE LEDGER.V_TYPE='" & frm.DataCombo3.BoundText & "' AND LEDGER.V_NO BETWEEN " & frm.DataCombo2(0) & " AND " & frm.DataCombo2(1) & " order by ledger.v_no,LEDGER.V_SNo")
'ElseIf frm.Opt2(2).Value = True Then
'    If FaIsValid(frm.Text2, "For Date ") = False Then Exit Sub
'    Set RST1 = G_FaCn.Execute("SELECT VOUCHER_TYPE.DESCRIPTION,NAME,LEDGER.V_Type,LEDGER.V_No,LEDGER.v_Prefix,LEDGER.V_Date,LEDGER.V_SNo,LEDGER.AmtCr,LEDGER.AmtDr,LEDGER.Chq_No,LEDGER.Chq_Date,LEDGER.Narration,LEDGERM.Narration AS NARRMAIN,LEDGERM.DocId,LEDGERM.v_Prefix FROM ((LEDGER LEFT JOIN LEDGERM ON LEDGERM.DOCID=LEDGER.DOCID)LEFT JOIN SUBGROUP ON  SUBGROUP.SUBCODE = LEDGER.SUBCODE ) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_tYPE WHERE VOUCHER_TYPE.Category='FA' AND LEDGER.V_date=" & FaConvertDate(frm.Text2) & " order by ledger.v_type,ledger.v_no,LEDGER.V_SNo")
'Else
    Set RST1 = G_FaCn.Execute("SELECT VOUCHER_TYPE.DESCRIPTION,NAME,LEDGER.V_Type,LEDGER.V_No,LEDGER.v_Prefix,LEDGER.V_Date,LEDGER.V_SNo,LEDGER.AmtCr,LEDGER.AmtDr,LEDGER.Chq_No,LEDGER.Chq_Date,LEDGER.Narration,LEDGERM.Narration AS NARRMAIN,LEDGERM.DocId,LEDGERM.v_Prefix, Ledger.AddBy FROM ((LEDGER LEFT JOIN LEDGERM ON LEDGERM.DOCID=LEDGER.DOCID)LEFT JOIN SUBGROUP ON  SUBGROUP.SUBCODE = LEDGER.SUBCODE ) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_tYPE WHERE LEDGER.V_TYPE='" & frm.TxtVtYpe(0).Tag & "' AND LEDGER.V_NO=" & frm.TxtVno(0) & " order by ledger.v_no,LEDGER.V_SNo")
'End If
If RST1.RecordCount = 0 Then MsgBox "No record found to Print", vbInformation, frm.CAPTION: Exit Sub
X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaRect.TTX", True)
Set rpt = VRep
rpt.Database.SetDataSource RST1
rpt.ReadRecords
FaReport_View rpt, 0, frm.CAPTION, True
Set RST1 = Nothing
Exit Sub
err1:    MsgBox err.Description, vbCritical, frm.CAPTION
End Sub
Public Sub FaDeleteTrack(Conn As ADODB.Connection, mDocId As String)
'
End Sub
Public Sub FaClosingBalanceCalculation()
'Opening & Closing Stock Calculation
'
'UPDATE FAENVIRO'S OPSTOCK & CLSTOCK FIELDS
End Sub

'''''' FOR SHRI RADHEY
'''''Dim RstVType As ADODB.Recordset
'''''Set RstVType = G_FaCn.Execute("SELECT NCAT FROM VOUCHER_TYPE WHERE V_TYPE='" & Vtype & "'")
'''''If RstVType.RecordCount > 0 Then
'''''    FGrid1.Tag = "VOUCHER"
'''''    Select Case RstVType!NCAT
'''''        Case "CNT", "JV", "PMT", "RCT"
'''''            FrmVent.Show
'''''            FrmVent.Tag = " "
'''''            FrmVent.SearchBackMagic Vtype + Space(8 - Len(LTrim(RTrim(STR(V_NO))))) + LTrim(RTrim(STR(V_NO)))
'''''            FrmVent.CurrMode = "Display"
''''''        Case "OPBAL"
''''''            frmSubGroup.Show
''''''            frmSubGroup.SEARCHBACK FGrid1.TextMatrix(FGrid1.Row, 10)
'''''        Case "R1", "R2", "S2"         'SALE BILL ENTRY
'''''            SBILL.Show
'''''            SBILL.Tag = " "
'''''            SBILL.SEARCHBACK Vtype + Space(8 - Len(LTrim(RTrim(STR(V_NO))))) + LTrim(RTrim(STR(V_NO)))
'''''            SBILL.CurrMode = "Display"
'''''            SBILL.FLD_ENABLE (False)
'''''        Case "P1", "P2"                    'PURCHASE BILL ENTRY
'''''            PBILL.Show
'''''            PBILL.Tag = " "
'''''            PBILL.SEARCHBACK Vtype + Space(8 - Len(LTrim(RTrim(STR(V_NO))))) + LTrim(RTrim(STR(V_NO)))
'''''            PBILL.CurrMode = "Display"
'''''            PBILL.FLD_ENABLE (False)
'''''    End Select
'''''End If
