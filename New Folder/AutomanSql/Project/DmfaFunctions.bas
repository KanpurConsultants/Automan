Attribute VB_Name = "Module1"
Option Explicit

Public IsConsolidatedPosting As Boolean


Public Sub ApplyConsolidatedPosting(VDate As Date)
    If VDate < CDate("31/Mar/2011") Then
        IsConsolidatedPosting = True
    Else
        IsConsolidatedPosting = False
    End If
End Sub


Public Function LedTrial(FormName As Object, Optional FRow As Integer, Optional Fcol As Integer) As ADODB.Recordset
Dim RST1 As ADODB.Recordset, mDR As Double, mCR As Double, mQRY1 As String, mQRY2 As String
Dim RstLedTrial As ADODB.Recordset, mGroupname As String, mGroupDR As Double, mGroupCR As Double
Dim moreThanOne As Integer, I As Integer, RstEnviro As ADODB.Recordset, mQtySum As Double, mGQtySum As Double
Dim XSpace As String, xOpQry As String, mCondStrForSite As String
''''Dim ClosingPostFlag As Boolean, RstClosStock As ADODB.Recordset
Set RstLedTrial = New ADODB.Recordset
Set RstLedTrial = mGroupTrial(RstLedTrial)
Set LedTrial = RstLedTrial
Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
If FormName.Check4.Value = 1 Then
    XSpace = Space(5)
Else
    XSpace = ""
End If
FormName.Text1 = ""
mQRY1 = ""
mQRY2 = ""
FormName.Check1.Enabled = True
FormName.Check2.Enabled = False
FormName.Check3.Enabled = True
FormName.Check4.Enabled = True
FormName.Check5.Enabled = True
FormName.Check6.Enabled = True
FormName.CAPTION = "Trial Balance (Ledger)"
If PubShowSiteWiseReport = True Then FormName.CAPTION = FormName.CAPTION + " [" + PubSiteName + "]"
FormName.Text2 = "From Date : " + CStr(FormName.TXTS_DATE) + " To : " + CStr(FormName.TXTE_DATE)
If FormName.Check1.Value = 1 Then
    mQRY1 = " Where ((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(FormName.TXTE_DATE) & " AND GroupNature IN ('E','R')) OR (V_DATE<=" & FaConvertDate(FormName.TXTE_DATE) & " AND GroupNature NOT IN ('E','R')))"
Else
    mQRY1 = " Where ((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " AND GroupNature IN ('E','R')) OR (V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " AND GroupNature NOT IN ('E','R')))"
End If
If FormName.BtnSite.Visible = True Then
    If PubFaSiteType = 1 Then
        mCondStrForSite = " AND RIGHT(LEDGER.SITE_CODE,1) IN " & PubSiteCodeDisplay
    ElseIf PubFaSiteType = 2 Then
        mCondStrForSite = " AND LEDGER.SITE_CODE IN " & PubSiteCodeDisplay
    End If
Else
    mCondStrForSite = ""
End If
G_FaCn.CommandTimeout = 120
If FormName.Check3.Value = 0 Then mQRY2 = "HAVING ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0"
If PubBackEnd = "A" Then
    If FormName.Check4.Value = 0 Then
        If RstEnviro!ShowCityName = "Yes" Then
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,LEFT(ViewSubgroup.NAMEWITHCITY,50) AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,IIF(ISNULL(SUM(PQTY)),0,SUM(PQTY))-IIF(ISNULL(SUM(SQTY)),0,SUM(SQTY)) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE")
        Else
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,NAME AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,IIF(ISNULL(SUM(PQTY)),0,SUM(PQTY))-IIF(ISNULL(SUM(SQTY)),0,SUM(SQTY)) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.NAME,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.NAME,ViewSubgroup.SUBCODE")
        End If
    Else
        If RstEnviro!ShowCityName = "Yes" Then
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,LEFT(ViewSubgroup.NAMEWITHCITY,50) AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,IIF(ISNULL(SUM(PQTY)),0,SUM(PQTY))-IIF(ISNULL(SUM(SQTY)),0,SUM(SQTY)) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.GNAME,ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.GNAME,ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE")
        Else
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,NAME AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,IIF(ISNULL(SUM(PQTY)),0,SUM(PQTY))-IIF(ISNULL(SUM(SQTY)),0,SUM(SQTY)) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.GNAME,ViewSubgroup.NAME,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.GNAME,ViewSubgroup.NAME,ViewSubgroup.SUBCODE")
        End If
    End If
'''''    Set RstClosStock = G_FaCn.Execute("SELECT ROUND(SUM(AMTDR),2)-ROUND(SUM(AMTCR),2) AS BAL  FROM LEDGER LEFT JOIN VIEWSUBGROUP ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " " & mCondStrForSite & " AND VIEWSUBGROUP.GROUPCODE='000A'")
ElseIf PubBackEnd = "S" Then
    If FormName.Check4.Value = 0 Then
        If RstEnviro!ShowCityName = "Yes" Then
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,LEFT(ViewSubgroup.NAMEWITHCITY,50) AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,ISNULL(SUM(PQTY),0)-ISNULL(SUM(SQTY),0) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE")
        Else
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,NAME AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,ISNULL(SUM(PQTY),0)-ISNULL(SUM(SQTY),0) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.NAME,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.NAME,ViewSubgroup.SUBCODE")
        End If
    Else
        If RstEnviro!ShowCityName = "Yes" Then
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,LEFT(ViewSubgroup.NAMEWITHCITY,50) AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,ISNULL(SUM(PQTY),0)-ISNULL(SUM(SQTY),0) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.GNAME,ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.GNAME,ViewSubgroup.NAMEWITHCITY,ViewSubgroup.SUBCODE")
        Else
            Set RST1 = G_FaCn.Execute("SELECT ViewSubgroup.SUBCODE AS PARTY,NAME AS PARTY_NAME,MAX(GROUPCODE) AS GRPCode,MAX(GNAME) AS GRPNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,ISNULL(SUM(PQTY),0)-ISNULL(SUM(SQTY),0) AS QtyBal FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ViewSubgroup.GNAME,ViewSubgroup.NAME,ViewSubgroup.SUBCODE " & mQRY2 & " ORDER BY ViewSubgroup.GNAME,ViewSubgroup.NAME,ViewSubgroup.SUBCODE")
        End If
    End If
'''''    Set RstClosStock = G_FaCn.Execute("SELECT ROUND(SUM(AMTDR),2)-ROUND(SUM(AMTCR),2) AS BAL,ISNULL(SUM(PQTY),0)-ISNULL(SUM(SQTY),0) AS QtyBal FROM LEDGER LEFT JOIN VIEWSUBGROUP ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " " & mCondStrForSite & " AND VIEWSUBGROUP.GROUPCODE='000A'")
End If

G_FaCn.CommandTimeout = 30
mGroupCR = 0
mGroupDR = 0
mGQtySum = 0
moreThanOne = 0
'''''ClosingPostFlag = False
If RST1.RecordCount > 0 Then
    RST1.MoveFirst
    Do Until RST1.EOF
'''''        If ClosingPostFlag = False Then
'''''            If RST1!GRPCode = "000A" Then
'''''                ClosingPostFlag = True
'''''                If RstClosStock.RecordCount > 0 Then
'''''                    If FaSNull(RstClosStock!BAL) > 0 Then
'''''                        mCR = mCR + Abs(RstClosStock!BAL)
'''''                        mGroupCR = mGroupCR + Abs(RstClosStock!BAL)
'''''                        mQtySum = mQtySum + FaVNull(RstClosStock!QTYBAL)
'''''                        mGQtySum = mGQtySum + FaVNull(RstClosStock!QTYBAL)
'''''                        With RstLedTrial
'''''                            .AddNew
'''''                            .Fields("GRCODE") = "000A"
'''''                            .Fields("SUBCODE") = ""
'''''                            .Fields("ACName") = "Closing Stock (Assets)"
'''''                            .Fields("Credit") = FaSNull(Abs(RstClosStock!BAL))
'''''                            .Fields("QtyBal") = IIf(FaVNull(RstClosStock!QTYBAL) <> 0, Format(FaVNull(RstClosStock!QTYBAL), "0.000"), "")
'''''                            .Update
'''''                        End With
'''''                    ElseIf FaSNull(RstClosStock!BAL) < 0 Then
'''''                        mDR = mDR + Abs(RstClosStock!BAL)
'''''                        mGroupDR = mGroupDR + Abs(RstClosStock!BAL)
'''''                        mQtySum = mQtySum + FaVNull(RstClosStock!QTYBAL)
'''''                        mGQtySum = mGQtySum + FaVNull(RstClosStock!QTYBAL)
'''''                        With RstLedTrial
'''''                            .AddNew
'''''                            .Fields("GRCODE") = "000A"
'''''                            .Fields("SUBCODE") = ""
'''''                            .Fields("ACName") = "Closing Stock (Assets)"
'''''                            .Fields("Debit") = FaSNull(Abs(RstClosStock!BAL))
'''''                            .Fields("QtyBal") = IIf(FaVNull(RstClosStock!QTYBAL) <> 0, Format(FaVNull(RstClosStock!QTYBAL), "0.000"), "")
'''''                            .Update
'''''                        End With
'''''                    End If
'''''                End If
'''''            End If
'''''        End If

        If FormName.Check4.Value = 1 Then
            If mGroupname <> RST1!GRPNAME Then
                mGroupname = RST1!GRPNAME
                mGroupCR = 0
                mGroupDR = 0
                mGQtySum = 0
                moreThanOne = 0
                With RstLedTrial
                    .AddNew
                    .Fields("GRCODE") = RST1!GRPCode
                    .Fields("ACName") = Trim(RST1!GRPNAME)
                    .Fields("GroupHead") = "*"
                    .Update
                End With
            End If
        End If
        moreThanOne = moreThanOne + 1
        If FaSNull(RST1!Bal) > 0 Then
            If FormName.Check5.Value = 0 Then
                mCR = mCR + Abs(RST1!Bal)
                mGroupCR = mGroupCR + Abs(RST1!Bal)
                mQtySum = mQtySum + FaVNull(RST1!QTYBAL)
                mGQtySum = mGQtySum + FaVNull(RST1!QTYBAL)
                With RstLedTrial
                    .AddNew
                    .Fields("GRCODE") = RST1!GRPCode
                    .Fields("SUBCODE") = RST1!Party
                    .Fields("ACName") = left(XSpace + RST1!Party_Name, 50)
                    .Fields("Credit") = FaSNull(Abs(RST1!Bal))
                    .Fields("QtyBal") = IIf(FaVNull(RST1!QTYBAL) <> 0, Format(FaVNull(RST1!QTYBAL), "0.000"), "")
                    .Update
                End With
            End If
        ElseIf FaSNull(RST1!Bal) < 0 Then
            If FormName.Check6.Value = 0 Then
                mDR = mDR + Abs(RST1!Bal)
                mGroupDR = mGroupDR + Abs(RST1!Bal)
                mQtySum = mQtySum + FaVNull(RST1!QTYBAL)
                mGQtySum = mGQtySum + FaVNull(RST1!QTYBAL)
                With RstLedTrial
                    .AddNew
                    .Fields("GRCODE") = RST1!GRPCode
                    .Fields("SUBCODE") = RST1!Party
                    .Fields("ACName") = left(XSpace + RST1!Party_Name, 50)
                    .Fields("Debit") = FaSNull(Abs(RST1!Bal))
                    .Fields("QtyBal") = IIf(FaVNull(RST1!QTYBAL) <> 0, Format(FaVNull(RST1!QTYBAL), "0.000"), "")
                    .Update
                End With
            End If
        ElseIf FaSNull(RST1!Bal) = 0 Then
            With RstLedTrial
                .AddNew
                .Fields("GRCODE") = RST1!GRPCode
                .Fields("SUBCODE") = RST1!Party
                .Fields("ACName") = left(XSpace + RST1!Party_Name, 50)
                .Update
            End With
        End If
        RST1.MoveNext
        If FormName.Check4.Value = 1 Then
            If RST1.EOF = True Then
                If moreThanOne > 1 Then
                    If Abs(mGroupCR) <> 0 Or Abs(mGroupDR) Then
                        With RstLedTrial
                            .AddNew
                            .Fields("GroupHead") = "X"
                            .Fields("QtyBal") = IIf(Abs(mQtySum) <> 0, String(12, "~"), "")
                            .Fields("Credit") = IIf(Abs(mGroupCR) <> 0, String(14, "~"), "")
                            .Fields("Debit") = IIf(Abs(mGroupDR) <> 0, String(14, "~"), "")
                            .Update
                        End With
                        With RstLedTrial
                            .AddNew
                            .Fields("GroupHead") = "X"
                            .Fields("ACName") = Space(25) + "Group Total"
                            .Fields("QtyBal") = IIf(FaVNull(mGQtySum) > 0, Format(FaVNull(mGQtySum), "0.000"), "")
                            .Fields("Credit") = IIf(Abs(mGroupCR) <> 0, FaSNull(Abs(mGroupCR)), "")
                            .Fields("Debit") = IIf(Abs(mGroupDR) <> 0, FaSNull(Abs(mGroupDR)), "")
                            .Update
                        End With
                    End If
                Else
                    With RstLedTrial
                        .AddNew
                        .Update
                    End With
                End If
            ElseIf mGroupname <> RST1!GRPNAME Then
                If moreThanOne > 1 Then
                    If Abs(mGroupCR) <> 0 Or Abs(mGroupDR) <> 0 Then
                        With RstLedTrial
                            .AddNew
                            .Fields("GroupHead") = "X"
                            .Fields("QtyBal") = IIf(Abs(mQtySum) <> 0, String(12, "~"), "")
                            .Fields("Credit") = IIf(Abs(mGroupCR) <> 0, String(14, "~"), "")
                            .Fields("Debit") = IIf(Abs(mGroupDR) <> 0, String(14, "~"), "")
                            .Update
                        End With
                        With RstLedTrial
                            .AddNew
                            .Fields("GroupHead") = "X"
                            .Fields("ACName") = Space(25) + "Group Total"
                            .Fields("QtyBal") = IIf(FaVNull(mGQtySum) > 0, Format(FaVNull(mGQtySum), "0.000"), "")
                            .Fields("Credit") = IIf(Abs(mGroupCR) <> 0, FaSNull(Abs(mGroupCR)), "")
                            .Fields("Debit") = IIf(Abs(mGroupDR) <> 0, FaSNull(Abs(mGroupDR)), "")
                            .Update
                        End With
                    End If
                Else
                    With RstLedTrial
                        .AddNew
                        .Update
                    End With
                End If
            End If
        End If
    Loop
''    If FormName.Check1.Value = 1 Then
''        xOpQry = " Where ((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(FormName.TXTS_DATE) & " AND SUBGROUP.GroupNature IN ('E','R')) OR (V_DATE<" & FaConvertDate(FormName.TXTS_DATE) & " AND SUBGROUP.GroupNature NOT IN ('E','R'))  AND LEFT(ACGROUP.MAINGRCODE,3) NOT IN ('999'))"
''        Set RST1 = G_FaCn.Execute("SELECT ROUND(SUM(AMTCR),2) AS CRSUM,ROUND(SUM(AMTDR),2) AS DRSUM FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=SUBGROUP.SUBCODE " & xOpQry & " " & mCondStrForSite)
''
''        If RST1.RecordCount > 0 Then
''            If Round(Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM)), 2) <> 0 Then
''                If FaVNull(RST1!CRSUM) > FaVNull(RST1!DRSUM) Then
''                    With RstLedTrial
''                        .AddNew
''                        .Fields("ACName") = "# Difference in Opening Balance"
''                        .Fields("Debit") = FaSNull(Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM)))
''                        .Update
''                    End With
''                    mDR = mDR + Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM))
''                Else
''                    With RstLedTrial
''                        .AddNew
''                        .Fields("ACName") = "# Difference in Opening Balance"
''                        .Fields("Credit") = FaSNull(Abs(FaVNull(RST1!CRSUM) - FaVNull(RST1!DRSUM)))
''                        .Update
''                    End With
''                    mCR = mCR + Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM))
''                End If
''            End If
''        End If
''    End If
End If
If RstLedTrial.RecordCount > 0 Then
    Set FormName.FGrid1.DataSource = RstLedTrial
End If
With FormName.FGrid1
    .Tag = "LEDTRIAL"
    .Cols = 7
    .TextMatrix(0, 0) = "GrCode"
    .ColWidth(0) = 0
    .TextMatrix(0, 1) = "SubCode"
    .ColWidth(1) = 0
    .TextMatrix(0, 2) = "ACName"
    .ColAlignment(2) = flexAlignLeftCenter
    .ColWidth(2) = 4500
    .TextMatrix(0, 3) = "Qty."
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignmentFixed(3) = flexAlignRightCenter
    .ColWidth(3) = IIf(RstEnviro!ShowQty = "Yes", 1500, 0)
    .TextMatrix(0, 4) = "Debit"
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignmentFixed(4) = flexAlignRightCenter
    .ColWidth(4) = IIf(RstEnviro!ShowQty = "Yes", 1800, 2000)
    .TextMatrix(0, 5) = "Credit"
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignmentFixed(5) = flexAlignRightCenter
    .ColWidth(5) = IIf(RstEnviro!ShowQty = "Yes", 1800, 2000)
    .TextMatrix(0, 6) = ""
    .ColWidth(6) = 0
End With
FormName.FGrid2.Rows = 1
With FormName.FGrid2
    .Cols = 7
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColAlignment(2) = flexAlignLeftCenter
    .ColWidth(2) = 4500
    .ColWidth(3) = IIf(RstEnviro!ShowQty = "Yes", 1500, 0)
    .ColAlignment(4) = flexAlignRightCenter
    .ColWidth(4) = IIf(RstEnviro!ShowQty = "Yes", 1800, 2000)
    .ColAlignment(5) = flexAlignRightCenter
    .ColWidth(5) = IIf(RstEnviro!ShowQty = "Yes", 1800, 2000)
    .TextMatrix(0, 6) = ""
    .ColWidth(6) = 0
End With
FormName.FGrid1.Refresh
For I = 0 To FormName.FGrid1.Rows - 1
    If FormName.FGrid1.TextMatrix(I, 5) = "*" Then
        FormName.FGrid1.Row = I
        FormName.FGrid1.Col = 2
        FormName.FGrid1.CellFontBold = True
    End If
Next
FormName.FGrid2.TextMatrix(0, 0) = ""
FormName.FGrid2.TextMatrix(0, 1) = ""
FormName.FGrid2.TextMatrix(0, 2) = "Total " + IIf(mDR - mCR = 0, "", " {" + FaBNull(Abs(mDR - mCR)) + IIf(mDR - mCR = 0, "", IIf(mDR - mCR > 0, " Dr", " Cr")) + "} ")
If RstEnviro!ShowQty = "Yes" Then
    FormName.FGrid2.TextMatrix(0, 3) = Format(IIf(mQtySum = 0, "", mQtySum), "0.00")
Else
    FormName.FGrid2.TextMatrix(0, 4) = ""
End If
FormName.FGrid2.TextMatrix(0, 4) = Format(IIf(mDR = 0, "", mDR), "0.00")
FormName.FGrid2.TextMatrix(0, 5) = Format(IIf(mCR = 0, "", mCR), "0.00")
FormName.FGrid2.TextMatrix(0, 6) = ""
If FormName.FGrid1.Rows = 1 Then FormName.FGrid1.AddItem ""
FormName.FGrid1.Row = IIf(FRow <> 0 And FormName.FGrid1.Rows - 1 >= FRow, FRow, 1)
FormName.FGrid1.Col = IIf(Fcol <> 0 And FormName.FGrid1.Cols - 1 >= Fcol, Fcol, 2)
If FRow <> 0 And FormName.FGrid1.Rows - 1 Then FormName.FGrid1.TopRow = FRow
Set LedTrial = RstLedTrial
Set RST1 = Nothing
Set RstLedTrial = Nothing
Set RstEnviro = Nothing
'''''Set RstClosStock = Nothing
End Function


Public Function GroupTrial(FormName As Object, Optional FRow As Integer, Optional Fcol As Integer, Optional xSite As String) As ADODB.Recordset
Dim RST1 As ADODB.Recordset, mDR As Double, mCR As Double, mQRY1 As String, mQRY2 As String, mQRY3 As String
Dim RstGroupTrial As ADODB.Recordset, RstEnviro As ADODB.Recordset, mQtySum As Double, xOpQry As String, mCondStrForSite As String
'''''Dim RstClosStock As ADODB.Recordset, ClosingPostFlag As Boolean
Set RstGroupTrial = New ADODB.Recordset
Set RstGroupTrial = mGroupTrial(RstGroupTrial)
Set GroupTrial = RstGroupTrial
Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
'''Optional xSite As String  ''' as per LP sir / Gurdeep  --- > 210307
FormName.Text1 = ""
mQRY1 = ""
FormName.CAPTION = "Trial Balance (Group)"
If PubShowSiteWiseReport = True Then FormName.CAPTION = FormName.CAPTION + " [" + PubSiteName + "]"
FormName.Check1.Enabled = True
FormName.Check2.Enabled = True
FormName.Check3.Enabled = True
FormName.Check4.Enabled = False
FormName.Check5.Enabled = True
FormName.Check6.Enabled = True
FormName.Text2 = "From Date : " + CStr(FormName.TXTS_DATE) + " To : " + CStr(FormName.TXTE_DATE)
If FormName.Check1.Value = 1 Then
    mQRY1 = " Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature IN ('E','R') AND acgroup.SYSGROUP='Y' AND ACGROUP.AliasYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE<=" & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature NOT IN ('E','R') AND ACGROUP.SYSGROUP='Y' AND ACGROUP.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
Else
    mQRY1 = " Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature IN ('E','R') AND ACGROUP.SYSGROUP='Y' AND ACGROUP.ALIASYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature NOT IN ('E','R') AND ACGROUP.SYSGROUP='Y' AND ACGROUP.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
End If
If FormName.Check1.Value = 1 Then
    mQRY3 = " Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature IN ('E','R')  AND ACGROUP.AliasYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE<=" & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature NOT IN ('E','R') AND ACGROUP.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
Else
    mQRY3 = " Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature IN ('E','R') AND ACGROUP.ALIASYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " AND ACGROUP.GroupNature NOT IN ('E','R') AND ACGROUP.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
End If
If FormName.BtnSite.Visible = True Then
    If PubFaSiteType = 1 Then
        mCondStrForSite = " AND RIGHT(LEDGER.SITE_CODE,1) IN " & PubSiteCodeDisplay
    ElseIf PubFaSiteType = 2 Then
        '''mCondStrForSite = " AND LEDGER.SITE_CODE IN " & PubSiteCodeDisplay
        If xSite <> "" Then  '' 210307
            mCondStrForSite = " AND LEDGER.SITE_CODE IN (" & xSite & ") "
        Else
            mCondStrForSite = " AND LEDGER.SITE_CODE IN " & PubSiteCodeDisplay
        End If
        
    End If
Else
    mCondStrForSite = ""
End If
'''''ClosingPostFlag = False
'''''Set RstClosStock = G_FaCn.Execute("SELECT ROUND(SUM(AMTDR),2)-ROUND(SUM(AMTCR),2) AS BAL,ISNULL(SUM(PQTY),0)-ISNULL(SUM(SQTY),0) AS QtyBal FROM (LEDGER LEFT JOIN VIEWSUBGROUP ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE " & mQRY1 & " " & mCondStrForSite & " AND VIEWSUBGROUP.GROUPCODE='000A'")

If PubBackEnd = "A" Then
    Set RST1 = G_FaCn.Execute("SELECT MAX(ACGROUP.MAINGRCODE) AS MAINGRCD,MAX(ACGROUP.GROUPCODE) AS GRPCode,ACGROUP.GROUPNAME AS GRPNAME,ROUND(sum(AMTCr),2)-ROUND(SUM(AMTDr),2) As Bal,IIF(ISNULL(SUM(PQTY)),0,SUM(PQTY))-IIF(ISNULL(SUM(SQTY)),0,SUM(SQTY)) AS QtyBal FROM (ACGROUP INNER JOIN VIEWSUBGROUP ON ACGROUP.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE and  LEDGER.Site_CODE=VIEWSUBGROUP.Site_CODE " & mQRY1 & " " & mCondStrForSite & " GROUP BY ACGROUP.GROUPNAME,ACGROUP.MAINGRCODE HAVING LEN(ACGROUP.MAINGRCODE)=3 ORDER BY ACGROUP.GROUPNAME,ACGROUP.MAINGRCODE")
ElseIf PubBackEnd = "S" Then
    Set RST1 = G_FaCn.Execute("SELECT MAX(ACGROUP.MAINGRCODE) AS MAINGRCD,MAX(ACGROUP.GROUPCODE) AS GRPCode,ACGROUP.GROUPNAME AS GRPNAME,ROUND(sum(AMTCr),2)-ROUND(SUM(AMTDr),2) As Bal,ISNULL(SUM(PQTY),0)-ISNULL(SUM(SQTY),0) AS QtyBal FROM (ACGROUP INNER JOIN VIEWSUBGROUP ON ACGROUP.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE  " & mQRY1 & " " & mCondStrForSite & " GROUP BY ACGROUP.GROUPNAME,ACGROUP.MAINGRCODE HAVING LEN(ACGROUP.MAINGRCODE)=3 ORDER BY ACGROUP.GROUPNAME,ACGROUP.MAINGRCODE")
End If
If RST1.RecordCount > 0 Then
    RST1.MoveFirst
    Do Until RST1.EOF
'''''        If ClosingPostFlag = False Then
'''''            If RST1!GRPCode = "000A" Then
'''''                ClosingPostFlag = True
'''''                If RstClosStock.RecordCount > 0 Then
'''''                    If FaSNull(RstClosStock!BAL) > 0 Then
'''''                        mCR = mCR + Abs(RstClosStock!BAL)
'''''                        mQtySum = mQtySum + FaVNull(RstClosStock!QTYBAL)
'''''                        With RstGroupTrial
'''''                            .AddNew
'''''                            .Fields("GRCODE") = ""
'''''                            .Fields("ACName") = "Closing Stock (Assets)"
'''''                            .Fields("Credit") = FaSNull(Abs(RstClosStock!BAL))
'''''                            .Fields("QtyBal") = IIf(FaVNull(RstClosStock!QTYBAL) <> 0, Format(FaVNull(RstClosStock!QTYBAL), "0.000"), "")
'''''                            .Update
'''''                        End With
'''''                    ElseIf FaSNull(RstClosStock!BAL) < 0 Then
'''''                        mDR = mDR + Abs(RstClosStock!BAL)
'''''                        mQtySum = mQtySum + FaVNull(RstClosStock!QTYBAL)
'''''                        With RstGroupTrial
'''''                            .AddNew
'''''                            .Fields("GRCODE") = ""
'''''                            .Fields("ACName") = "Closing Stock (Assets)"
'''''                            .Fields("Debit") = FaSNull(Abs(RstClosStock!BAL))
'''''                            .Fields("QtyBal") = IIf(FaVNull(RstClosStock!QTYBAL) <> 0, Format(FaVNull(RstClosStock!QTYBAL), "0.000"), "")
'''''                            .Update
'''''                        End With
'''''                    End If
'''''                End If
'''''            End If
'''''        End If
    
        If RST1!Bal > 0 Then
            If FormName.Check5.Value = 0 Then
                mCR = mCR + Abs(RST1!Bal)
                mQtySum = mQtySum + FaVNull(RST1!QTYBAL)
                With RstGroupTrial
                    .AddNew
                    .Fields("GRCODE") = RST1!GRPCode
                    .Fields("ACName") = RST1!GRPNAME
                    .Fields("Credit") = FaSNull(Abs(RST1!Bal))
                    .Fields("QtyBal") = IIf(FaVNull(RST1!QTYBAL) <> 0, Format(FaVNull(RST1!QTYBAL), "0.000"), "")
                    .Update
                End With
                If FormName.Check2.Value = 1 Then ChakRam FormName, RstGroupTrial, "GROUPTRIAL", mQRY3, RST1!GRPCode, , RST1!MAINGRCD
            End If
        ElseIf RST1!Bal < 0 Then
            If FormName.Check6.Value = 0 Then
                mDR = mDR + Abs(RST1!Bal)
                mQtySum = mQtySum + FaVNull(RST1!QTYBAL)
                With RstGroupTrial
                    .AddNew
                    .Fields("GRCODE") = RST1!GRPCode
                    .Fields("ACName") = RST1!GRPNAME
                    .Fields("Debit") = FaSNull(Abs(RST1!Bal))
                    .Fields("QtyBal") = IIf(FaVNull(RST1!QTYBAL) <> 0, Format(FaVNull(RST1!QTYBAL), "0.000"), "")
                    .Update
                End With
                If FormName.Check2.Value = 1 Then ChakRam FormName, RstGroupTrial, "GROUPTRIAL", mQRY3, RST1!GRPCode, , RST1!MAINGRCD
            End If
        ElseIf RST1!Bal = 0 Then
            If FormName.Check3.Value = 1 Then
                With RstGroupTrial
                    .AddNew
                    .Fields("GRCODE") = RST1!GRPCode
                    .Fields("ACName") = RST1!GRPNAME
                    .Fields("Debit") = "0.00"
                    .Update
                End With
                If FormName.Check2.Value = 1 Then ChakRam FormName, RstGroupTrial, "GROUPTRIAL", mQRY3, RST1!GRPCode, , RST1!MAINGRCD
            End If
        End If
        RST1.MoveNext
    Loop
'''    If FormName.Check1.Value = 1 Then
'''        xOpQry = " Where ((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(FormName.TXTS_DATE) & " AND SUBGROUP.GroupNature IN ('E','R')) OR (V_DATE<" & FaConvertDate(FormName.TXTS_DATE) & " AND SUBGROUP.GroupNature NOT IN ('E','R')) AND LEFT(ACGROUP.MAINGRCODE,3) NOT IN ('999'))"
'''        Set RST1 = G_FaCn.Execute("SELECT ROUND(SUM(AMTCR),2) AS CRSUM,ROUND(SUM(AMTDR),2) AS DRSUM FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE AND LEDGER.Site_Code=SUBGROUP.Site_Code) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=SUBGROUP.SUBCODE  " & xOpQry & " " & mCondStrForSite)
'''        If RST1.RecordCount > 0 Then
'''            If Round(Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM)), 2) <> 0 Then
'''                If FaVNull(RST1!CRSUM) > FaVNull(RST1!DRSUM) Then
'''                    With RstGroupTrial
'''                        .AddNew
'''                        .Fields("ACName") = "# Difference in Opening Balance"
'''                        .Fields("Debit") = FaSNull(Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM)))
'''                        .Update
'''                    End With
'''                    mDR = mDR + Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM))
'''                Else
'''                    With RstGroupTrial
'''                        .AddNew
'''                        .Fields("ACName") = "# Difference in Opening Balance"
'''                        .Fields("Credit") = FaSNull(Abs(FaVNull(RST1!CRSUM) - FaVNull(RST1!DRSUM)))
'''                        .Update
'''                    End With
'''                    mCR = mCR + Abs(FaVNull(RST1!DRSUM) - FaVNull(RST1!CRSUM))
'''                End If
'''            End If
'''        End If
'''    End If
End If
If RstGroupTrial.RecordCount > 0 Then
    Set FormName.FGrid1.DataSource = RstGroupTrial
End If
With FormName.FGrid1
    .Tag = "GROUPTRIAL"
    .Cols = 6
    .TextMatrix(0, 0) = "GrCode"
    .ColWidth(0) = 0
    .TextMatrix(0, 1) = "SubCode"
    .ColWidth(1) = 0
    .TextMatrix(0, 2) = "ACName"
    .ColAlignment(2) = flexAlignLeftCenter
    .ColWidth(2) = 4500
    .TextMatrix(0, 3) = "Qty."
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignmentFixed(3) = flexAlignRightCenter
    If RstEnviro!ShowQty = "Yes" Then
        .ColWidth(3) = 1700
        .ColWidth(4) = 1700
        .ColWidth(5) = 1700
    Else
        .ColWidth(3) = 0
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
    End If
    .TextMatrix(0, 4) = "Debit"
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignmentFixed(4) = flexAlignRightCenter
    .TextMatrix(0, 5) = "Credit"
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignmentFixed(5) = flexAlignRightCenter
End With
FormName.FGrid2.Rows = 1
With FormName.FGrid2
    .Cols = 6
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColAlignment(2) = flexAlignLeftCenter
    .ColWidth(2) = 4500
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    If RstEnviro!ShowQty = "Yes" Then
        .ColWidth(3) = 1700
        .ColWidth(4) = 1700
        .ColWidth(5) = 1700
    Else
        .ColWidth(3) = 0
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
    End If
End With
FormName.FGrid2.TextMatrix(0, 0) = ""
FormName.FGrid2.TextMatrix(0, 1) = ""
FormName.FGrid2.TextMatrix(0, 2) = "Total " + IIf(mDR - mCR = 0, "", " {" + FaBNull(Abs(mDR - mCR)) + IIf(mDR - mCR = 0, "", IIf(mDR - mCR > 0, " Dr", " Cr")) + "} ")
FormName.FGrid2.TextMatrix(0, 3) = Format(IIf(mQtySum = 0, "", mQtySum), "0.00")
FormName.FGrid2.TextMatrix(0, 4) = Format(IIf(mDR = 0, "", mDR), "0.00")
FormName.FGrid2.TextMatrix(0, 5) = Format(IIf(mCR = 0, "", mCR), "0.00")
If FormName.FGrid1.Rows = 1 Then FormName.FGrid1.AddItem ""
FormName.FGrid1.Row = IIf(FRow <> 0 And FormName.FGrid1.Rows - 1 >= FRow, FRow, 1)
FormName.FGrid1.Col = IIf(Fcol <> 0 And FormName.FGrid1.Cols - 1 >= Fcol, Fcol, 2)
If FRow <> 0 And FormName.FGrid1.Rows - 1 Then FormName.FGrid1.TopRow = FRow
'''  adi  140207
Dim X As Integer, xxDr, xxCr
For X = 1 To FormName.FGrid1.Rows - 1
    If Val(FormName.FGrid1.TextMatrix(X, 4)) > 0 Then
        xxDr = xxDr + 1
    ElseIf Val(FormName.FGrid1.TextMatrix(X, 5)) > 0 Then
        xxCr = xxCr + 1
    End If
Next
Set GroupTrial = RstGroupTrial
Set RST1 = Nothing
Set RstGroupTrial = Nothing
Set RstEnviro = Nothing
'''''Set RstClosStock = Nothing
End Function

Public Function mGroupTrial(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "GRCODE", adVarChar, 8, adFldIsNullable
    .Fields.Append "SUBCODE", adVarChar, 8, adFldIsNullable
    .Fields.Append "ACName", adVarChar, 50, adFldIsNullable
    .Fields.Append "QtyBal", adVarChar, 12, adFldIsNullable
    .Fields.Append "Debit", adVarChar, 14, adFldIsNullable
    .Fields.Append "Credit", adVarChar, 14, adFldIsNullable
    .Fields.Append "GroupHead", adVarChar, 1, adFldIsNullable
    .Fields.Append "MainGrCode", adVarChar, 255, adFldIsNullable
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set mGroupTrial = Rst
End Function

Private Sub ChakRam(FormName As Object, RstToAdd As ADODB.Recordset, ModuleName As String, Optional mQRY1 As String, Optional mCode1 As String, Optional mCode2 As String, Optional MainGrCode1 As String, Optional MainGrCode2 As String, Optional mSno As Integer, Optional xVal)
Dim RstCheck21 As ADODB.Recordset, RstCheck22 As ADODB.Recordset, Rst As ADODB.Recordset
Dim RstEnviro As ADODB.Recordset, mClStock As Double, mCondStrForSite As String, mCondStrOpening As String
Dim mClosingStockFlag As Boolean, mCurrentStockFlag As Boolean
mClStock = 0
Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
If RstEnviro.RecordCount > 0 Then
    mClStock = FaVNull(RstEnviro!ClStockValue)
End If
If PubShowSiteWiseReport = True Then
    If PubFaSiteType = 1 Then
        mCondStrForSite = " AND RIGHT(LEDGER.SITE_CODE,1) IN " & PubSiteCodeDisplay
    ElseIf PubFaSiteType = 2 Then
        mCondStrForSite = " AND LEDGER.SITE_CODE IN " & PubSiteCodeDisplay
    End If
Else
    mCondStrForSite = ""
End If
mCondStrOpening = ""
If FormName.Check1.Value = 1 Then
    mCondStrOpening = " WHERE V_DATE<=" & FaConvertDate(FormName.TXTE_DATE)
Else
    mCondStrOpening = " WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE)
End If
mClosingStockFlag = False
mCurrentStockFlag = False
Select Case ModuleName
    Case "BalSheet", "VBalSheet"
        Set RstCheck21 = New ADODB.Recordset
        With RstCheck21
            .Fields.Append "TT", adInteger, , adFldIsNullable
            .Fields.Append "GrCode", adVarChar, 4, adFldIsNullable
            .Fields.Append "GroupName", adVarChar, 50, adFldIsNullable
            .Fields.Append "SubCode", adVarChar, 8, adFldIsNullable
            .Fields.Append "AcYNAME", adVarChar, 50, adFldIsNullable
            .Fields.Append "Bal", adDouble, , adFldIsNullable
            .Fields.Append "MAGRCODE", adVarChar, 255, adFldIsNullable
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        Set RstCheck22 = New ADODB.Recordset
        With RstCheck22
            .Fields.Append "TT", adInteger, , adFldIsNullable
            .Fields.Append "GrCode", adVarChar, 4, adFldIsNullable
            .Fields.Append "GroupName", adVarChar, 50, adFldIsNullable
            .Fields.Append "SubCode", adVarChar, 8, adFldIsNullable
            .Fields.Append "AcYNAME", adVarChar, 50, adFldIsNullable
            .Fields.Append "Bal", adDouble, , adFldIsNullable
            .Fields.Append "MAGRCODE", adVarChar, 255, adFldIsNullable
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        If RstEnviro!ShowCityName = "Yes" Then
            Set Rst = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE " & mCondStrOpening & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " " & mCondStrForSite & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
            "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mCondStrOpening & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
        Else
            Set Rst = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE " & mCondStrOpening & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " " & mCondStrForSite & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
            "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mCondStrOpening & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
        End If
        Do Until Rst.EOF
            If left(Rst!magrcode, 3) = "060" Then mCurrentStockFlag = True
            If Rst!magrcode <> "060001" Then
                With RstCheck21
                    .AddNew
                    .Fields("TT") = Rst!TT
                    .Fields("GroupName") = Rst!GroupName
                    .Fields("GrCode") = Rst!GrCode
                    .Fields("SubCode") = Rst!SubCode
                    .Fields("AcYNAME") = Rst!AcYNAME
                    .Fields("Bal") = Rst!Bal
                    .Fields("MAGRCODE") = Rst!magrcode
                    .Update
                End With
            ElseIf mClStock <> 0 Then
                With RstCheck21
                    .AddNew
                    .Fields("TT") = 1
                    .Fields("GroupName") = ""
                    .Fields("GrCode") = ""
                    .Fields("SubCode") = ""
                    .Fields("AcYNAME") = "Closing Stock"
                    .Fields("Bal") = Trim(mClStock)
                    .Fields("MAGRCODE") = ""
                    .Update
                End With
                mClosingStockFlag = True
            End If
            Rst.MoveNext
        Loop
        If mCurrentStockFlag = True And mClosingStockFlag = False And mClStock <> 0 Then
            With RstCheck21
                .AddNew
                .Fields("TT") = 1
                .Fields("GroupName") = ""
                .Fields("GrCode") = ""
                .Fields("SubCode") = ""
                .Fields("AcYNAME") = "Closing Stock"
                .Fields("Bal") = Trim(mClStock)
                .Fields("MAGRCODE") = ""
                .Update
            End With
            mCurrentStockFlag = False
            mClosingStockFlag = False
        End If
        If RstEnviro!ShowCityName = "Yes" Then
            Set Rst = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE " & mCondStrOpening & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
            "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mCondStrOpening & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
        Else
            Set Rst = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE " & mCondStrOpening & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
            "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mCondStrOpening & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
        End If
        Do Until Rst.EOF
            If left(Rst!magrcode, 3) = "060" Then mCurrentStockFlag = True
            If Rst!magrcode <> "060001" Then
                With RstCheck22
                    .AddNew
                    .Fields("TT") = Rst!TT
                    .Fields("GroupName") = Rst!GroupName
                    .Fields("GrCode") = Rst!GrCode
                    .Fields("SubCode") = Rst!SubCode
                    .Fields("AcYNAME") = Rst!AcYNAME
                    .Fields("Bal") = Rst!Bal
                    .Fields("MAGRCODE") = Rst!magrcode
                    .Update
                End With
            ElseIf mClStock <> 0 Then
                With RstCheck22
                    .AddNew
                    .Fields("TT") = 1
                    .Fields("GroupName") = ""
                    .Fields("GrCode") = ""
                    .Fields("SubCode") = ""
                    .Fields("AcYNAME") = "Closing Stock"
                    .Fields("Bal") = Trim(mClStock)
                    .Fields("MAGRCODE") = ""
                    .Update
                End With
                mClosingStockFlag = True
            End If
            Rst.MoveNext
        Loop
        If mCurrentStockFlag = True And mClosingStockFlag = False And mClStock <> 0 Then
            With RstCheck22
                .AddNew
                .Fields("TT") = 1
                .Fields("GroupName") = ""
                .Fields("GrCode") = ""
                .Fields("SubCode") = ""
                .Fields("AcYNAME") = "Closing Stock"
                .Fields("Bal") = Trim(mClStock)
                .Fields("MAGRCODE") = ""
                .Update
            End With
            mCurrentStockFlag = False
            mClosingStockFlag = False
        End If
    Case "ProfLoss"
        If FormName.Check1.Value = 1 Then
            If MainGrCode1 = "060001" Then
                If RstEnviro!ShowCityName = "Yes" Then
                    Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE  V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                Else
                    Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE  V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                End If
            Else
                If RstEnviro!ShowCityName = "Yes" Then
                    Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE  V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                Else
                    Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE  V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                End If
            End If
            If MainGrCode2 = "060001" Then
                If RstEnviro!ShowCityName = "Yes" Then
                    Set RstCheck22 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                Else
                    Set RstCheck22 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                End If
            Else
                If RstEnviro!ShowCityName = "Yes" Then
                    Set RstCheck22 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE  V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                Else
                    Set RstCheck22 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                    "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE  V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0")
                End If
            End If
        Else
            If RstEnviro!ShowCityName = "Yes" Then
                Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                "Union  SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & "  AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
                
                Set RstCheck22 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                "Union  SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
            Else
                Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                "Union  SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTDR),2)-ROUND(SUM(AMTCR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & "  AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+" & IIf(Len(MainGrCode1) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")

                Set RstCheck22 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode2) & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
                "Union  SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(FormName.TXTS_DATE) & " AND " & FaConvertDate(FormName.TXTE_DATE) & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode2 & "'))='" & MainGrCode2 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode2 & "')+" & IIf(Len(MainGrCode2) = 0, 0, 3) & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
            End If
        End If
    Case "GROUPTRIAL"
        If RstEnviro!ShowCityName = "Yes" Then
            Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,LEFT(MAX(ViewSubgroup.NAMEWITHCITY),50) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE " & mCondStrOpening & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " " & mCondStrForSite & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
            "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+3 GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
        Else
            Set RstCheck21 = G_FaCn.Execute("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS SubCode,MAX(ViewSubgroup.NAME) AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE " & mCondStrOpening & " AND ViewSubgroup.GROUPCODE=" & FaChk_Text(mCode1) & " " & mCondStrForSite & " AND ACGROUP.AliasYN='N' GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 " & _
            "Union SELECT 2 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS SUBCODE,ACGROUP.GROUPNAME AS AcYNAME,ROUND(sum(AMTCR),2)-ROUND(SUM(AMTDR),2) As Bal,MAX(ACGROUP.MAINGRCODE) AS MAGRCODE FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " " & mCondStrForSite & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & MainGrCode1 & "'))='" & MainGrCode1 & "' AND LEN(MAINGRCODE)=LEN('" & MainGrCode1 & "')+3 GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME HAVING ROUND(SUM(AMTCR),2)-ROUND(SUM(AMTDR),2)<>0 ")
        End If
    Case "LEDGER"
        Set RstCheck21 = G_FaCn.Execute("SELECT SUBGROUP.NAME,LEDGER.* FROM LEDGER LEFT JOIN SUBGROUP ON LEDGER.SUBCODE=SUBGROUP.SUBCODE WHERE DOCID=" & FaChk_Text(mQRY1) & " AND V_SNo<>" & mSno & " " & mCondStrForSite)
End Select
If RstCheck21.RecordCount > 0 Then RstCheck21.MoveFirst
Select Case ModuleName
    Case "BalSheet", "VBalSheet", "ProfLoss"
        If RstCheck22.RecordCount > 0 Then RstCheck22.MoveFirst
End Select
Select Case ModuleName
    Case "VBalSheet"
        Do Until RstCheck21.EOF
            With RstToAdd
                .AddNew
                .Fields("CodeCr") = RstCheck21!GrCode
                .Fields("SourceOfFund") = Space(2) + FaSetW(RstCheck21!AcYNAME, 15) + " (" + Trim(FaSNull(RstCheck21!Bal)) + ")"
                .Update
            End With
            RstCheck21.MoveNext
        Loop
        Do Until RstCheck22.EOF
            With RstToAdd
                .AddNew
                .Fields("CodeCr") = RstCheck22!GrCode
                .Fields("SourceOfFund") = Space(2) + FaSetW(RstCheck22!AcYNAME, 15) + " (" + Trim(FaSNull(RstCheck22!Bal)) + ")"
                .Update
            End With
            RstCheck22.MoveNext
        Loop
    Case "BalSheet", "ProfLoss"
        Do While True
            If Not RstCheck21.EOF And Not RstCheck22.EOF Then
                With RstToAdd
                    .AddNew
                    .Fields("CodeCr") = RstCheck21!GrCode
                    .Fields("SourceOfFund") = Space(2) + FaSetW(RstCheck21!AcYNAME, 15) + " (" + Trim(FaSNull(RstCheck21!Bal)) + ")"
                    .Fields("Seperator") = "|"
                    .Fields("CodeDr") = RstCheck22!GrCode
                    .Fields("ApplicationOfFund") = Space(2) + FaSetW(RstCheck22!AcYNAME, 15) + " (" + Trim(FaSNull(RstCheck22!Bal)) + ")"
                    .Update
                End With
                RstCheck21.MoveNext
                RstCheck22.MoveNext
            End If
            If Not RstCheck21.EOF And RstCheck22.EOF Then
                With RstToAdd
                    .AddNew
                    .Fields("CodeCr") = RstCheck21!GrCode
                    .Fields("SourceOfFund") = Space(2) + FaSetW(RstCheck21!AcYNAME, 15) + " (" + Trim(FaSNull(RstCheck21!Bal)) + ")"
                    .Fields("Seperator") = "|"
                    .Update
                End With
                RstCheck21.MoveNext
            End If
            If RstCheck21.EOF And Not RstCheck22.EOF Then
                With RstToAdd
                    .AddNew
                    .Fields("Seperator") = "|"
                    .Fields("CodeDr") = RstCheck22!GrCode
                    .Fields("ApplicationOfFund") = Space(2) + FaSetW(RstCheck22!AcYNAME, 15) + " (" + Trim(FaSNull(RstCheck22!Bal)) + ")"
                    .Update
                End With
                RstCheck22.MoveNext
            End If
            If RstCheck21.EOF And RstCheck22.EOF Then Exit Do
        Loop
    Case "GROUPTRIAL"
        Do Until RstCheck21.EOF
            With RstToAdd
                .AddNew
                .Fields("GRCODE") = RstCheck21!GrCode
                .Fields("ACName") = Space(8) + FaSetW(RstCheck21!AcYNAME, 15) + " (" + Trim(FaSNull(Abs(RstCheck21!Bal))) + " " + IIf(RstCheck21!Bal > 0, "Cr", "Dr") + ")"
                .Update
            End With
            RstCheck21.MoveNext
        Loop
    Case "LEDGER"
        Do Until RstCheck21.EOF
            With RstToAdd
                .AddNew
                .Fields("Sub") = "*"
                .Fields("VAL") = xVal
                .Fields("PDate") = Format(RstCheck21!V_DATE, "dd/MMM/yyyy")
                .Fields("ACName1") = mCode1
                .Fields("DocId") = RstCheck21!DocID
                If RstCheck21!AmtDr > 0 Then
                    .Fields("ACName") = Space(2) + FaSetW(RstCheck21!Name, 19) + " " + FaSetN(FaSNull(RstCheck21!AmtDr), 13) + " Dr"
                Else
                    .Fields("ACName") = Space(2) + FaSetW(RstCheck21!Name, 19) + " " + FaSetN(FaSNull(RstCheck21!AmtCr), 13) + " Cr"
                End If
                .Update
            End With
            RstCheck21.MoveNext
        Loop
End Select
Set RstCheck21 = Nothing
Set RstCheck22 = Nothing
Set Rst = Nothing
Set RstEnviro = Nothing
End Sub




