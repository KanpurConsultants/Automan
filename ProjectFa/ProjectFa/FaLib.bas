Attribute VB_Name = "FaLib"
Option Explicit
Public Const FaBackColorSelEnter As String = &HF8D7FD
Public Const FaCellBackColLeave1 As String = &HEDF7FE
Public Const FaCellBackColEnter1 As String = &HFFFFC0
'''''
Public Const PubShowSiteWiseReport As Boolean = True
'''''
Public MultiComp As Boolean, PubSec As String, PubBackEnd As String, PubFaSiteType As Byte
Public PubFaReportPath As String, PubFaDosPort As String, PubRunPIF As String
Public FaMasterRst As ADODB.Recordset, Dater As String, Pager As String, Titler As String, LineFil As String
Public PubSiteCodeDisplay As String, PubSiteName As String
Public PubSiteCodeWiseMasterRst As Boolean, PubSiteCodeWiseHelp As Boolean, PubSiteCodeWidth As Byte

Public PubSeparateVrNoForSite As Byte, PubSeparateLogSite As String

Declare Function CreateFieldDefFile Lib "p2smon.dll" (X As Object, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%) As Integer
Dim Unit, tens, tenth, WORDs, PLACE As String

'0-General (Without Site)/1-FromSite ForSite/2-2 Char Site
Public Sub FaCalCurrBal(CC As ADODB.Connection, AcCode As String, AmtDr As Double, AmtCr As Double)
Dim Rst1 As ADODB.Recordset, ControlStr As String, I As Integer, Length As Integer
CC.Execute "Update SubGroup Set Curr_Bal=Curr_Bal+" & AmtCr - AmtDr & " Where SubCode='" & AcCode & "'"
Set Rst1 = Nothing
End Sub
Public Sub FaFormIni(mFORM As Form, CtrlBColOrg As String, CtrlFColOrg As String)
On Error Resume Next
Dim I As Long
For I = 0 To mFORM.Count - 1
    If TypeOf mFORM.Controls(I) Is TextBox Or TypeOf mFORM.Controls(I) Is ComboBox Or TypeOf mFORM.Controls(I) Is DataCombo Then
        mFORM.Controls(I).BackColor = CtrlBColOrg
        mFORM.Controls(I).ForeColor = CtrlFColOrg
    End If
Next I
End Sub
Public Function FaXNull(temp As Variant) As String
    FaXNull = IIf(IsNull(temp), "", temp)
End Function
Public Function FaVNull(temp As Variant) As Variant
    FaVNull = IIf(IsNull(temp) Or temp = "", 0, temp)
End Function
Public Function FaSNull(temp As Variant) As Variant
    FaSNull = Format(IIf(IsNull(temp) Or temp = "", 0, temp), "0.00")
End Function
Public Function FaBNull(temp As Variant) As Variant
    If IsNull(temp) Or temp = "" Or temp = 0 Then
        FaBNull = ""
    Else
        FaBNull = Format(IIf(IsNull(temp) Or temp = "", 0, temp), "0.00")
    End If
End Function
Public Function FaValidate_Numeric(temp As Variant) As Double
    FaValidate_Numeric = IIf(Trim(temp) = "" Or IsNumeric(temp) = False, 0, Val(temp))
End Function
Public Function FaConvertDate(temp)
    If IsNull(temp) Or temp = "" Then
        FaConvertDate = "Null"
    Else
        If PubBackEnd = "A" Then
            FaConvertDate = "#" & Format(CDate(temp), "dd/MMM/yyyy") & "#"
        Else
            FaConvertDate = "'" & Format(CDate(temp), "dd/MMM/yyyy") & "'"
        End If
    End If
End Function
Public Function FaConvertDateTime(temp)
    If IsNull(temp) Or temp = "" Then
        FaConvertDateTime = "Null"
    Else
        If PubBackEnd = "A" Then
            FaConvertDateTime = "#" & Format(CDate(temp), "dd/MMM/yyyy hh:nn:ss") & "#"
        Else
            FaConvertDateTime = "'" & Format(CDate(temp), "dd/MMM/yyyy hh:nn:ss") & "'"
        End If
    End If
End Function
Public Function FaChk_Text(temp As Variant) As Variant
FaChk_Text = temp
If IsNull(FaChk_Text) Or FaChk_Text = Null Then
    FaChk_Text = "Null"
    Exit Function
End If
FaChk_Text = "'" & Replace(FaChk_Text, "'", "''") & "'"
End Function
Public Sub FaReport_View(mREPORT As CRAXDRT.Report, Index As Integer, CAPTION As String, EJECT_YN As Boolean)
Dim rpt_form As New FaRepView, mFILE_NAME As String * 16, connectionId
Dim fob As New Scripting.FileSystemObject, mSno As Byte, varTxtstrm As Scripting.TextStream
Dim mReportCount As Integer
On Error GoTo ERRORHANDLER
'   connectionId = mREPORT.SelectPrinter(VB.Printer.DriverName, VB.Printer.DeviceName, VB.Printer.Port)
    If Index = 1 Then
        Select Case mREPORT.PaperSize
            Case 268
                MsgBox "Please Set 132 Column Paper in Your Printer on 12 cpi & Put it On", vbInformation, "Paper Setting"
            Case 263
                MsgBox "Please Set 80 Column Paper in Your Printer on 12 cpi & Put it On", vbInformation, "Paper Setting"
            Case 9
                MsgBox "Please Set A4 Paper in Your Printer on 12 cpi & Put it On", vbInformation, "Paper Setting"
            Case Else
                MsgBox "Paper Size No. is " + STR(mREPORT.PaperSize), vbInformation, "Paper Setting"
        End Select
        mSno = 0
        If fob.FolderExists("C:\REPTMP") = False Then fob.CreateFolder ("C:\REPTMP")
        Do While True
            If mSno >= 100 Then
                For mSno = 1 To 100
                    mFILE_NAME = "C:\REPTMP\REP" + Trim(STR(mSno))
                    If fob.FileExists(Trim(mFILE_NAME) + ".TXT") Then fob.DeleteFile (Trim(mFILE_NAME) + ".TXT")
                Next
                mSno = 1
            End If
            mFILE_NAME = "C:\REPTMP\REP" + Trim(STR(mSno))
            If fob.FileExists(Trim(mFILE_NAME) + ".TXT") Then
                mSno = mSno + 1
            Else
                Exit Do
            End If
        Loop
        
        mREPORT.ExportOptions.DiskFileName = Trim(mFILE_NAME) + ".TXT"
        If EJECT_YN = True Then
            mREPORT.ExportOptions.FormatType = crEFTPaginatedText       '10
        Else
            mREPORT.ExportOptions.FormatType = crEFTText '8
        End If
        mREPORT.ExportOptions.NumberOfLinesPerPage = 60
        mREPORT.ExportOptions.DestinationType = 1
        For mReportCount = 1 To mREPORT.FormulaFields.Count
            Select Case UCase(mREPORT.FormulaFields(mReportCount).FormulaFieldName)
                Case UCase("comp_name")
                    mREPORT.FormulaFields(mReportCount).TEXT = "'" & Chr(27) + "W1" + Chr(27) + "W1" + Chr(27) + "E" + Chr(27) + "G" + PubComp_Name + Chr(27) + "H" + Chr(27) + "F" + Chr(27) + "W0" + Chr(27) + "W0" & "'"
                Case UCase("comp_add1")
                    mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_Add & "'"
                Case UCase("comp_pin")
                    mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_City & "'"
            End Select
        Next
        Call mREPORT.Export(False)
        
        Set varTxtstrm = fob.OpenTextFile(Trim(mFILE_NAME) + ".TXT", ForAppending)
        varTxtstrm.Write (Chr(12))
        varTxtstrm.Close
        Set varTxtstrm = fob.OpenTextFile("C:\REPTMP\REPPRINT.BAT", ForWriting, True)
        varTxtstrm.Write ("TYPE %1>" + PubFaDosPort)
        varTxtstrm.Close
        
        If PubRunPIF = "Y" Then
            If fob.FileExists("C:\REPTMP\REPPRINT.PIF") = False Then MsgBox "C:\REPTMP\REPPRINT.PIF File Not Exist": Exit Sub
            connectionId = Shell("C:\REPTMP\REPPRINT.PIF " + Trim(mFILE_NAME) + ".TXT", vbHide)
        Else
            If fob.FileExists("C:\REPTMP\REPPRINT.BAT") = False Then MsgBox "C:\REPTMP\REPPRINT.BAT File Not Exist": Exit Sub
            connectionId = Shell("C:\REPTMP\REPPRINT.BAT " + Trim(mFILE_NAME) + ".TXT", vbHide)
        End If

'''''
'    Dim mReportCount As Integer, XyzAbc As String
'    Dim mPAGE As Integer, mRow As Integer
'    mPAGE = 1
'    mRow = 1
'    If IsMissing(LinesPerPage) = False Or LinesPerPage = 30 Then
'        If fob.FileExists(Trim(mFILE_NAME) + ".TXT") Then
'            Set varTxtstrm = fob.OpenTextFile(Trim(mFILE_NAME) + ".TXT", ForReading)
'            Set varTxtstrmWithoutEject = fob.OpenTextFile(Trim(mFILE_NAME) + "Ej" + ".TXT", ForAppending, True)
'            Do Until varTxtstrm.AtEndOfStream
'                XyzAbc = varTxtstrm.ReadLine
'                mRow = mRow + 1
'                If mPAGE = 1 Then
'                    If mRow = 28 Or mRow = 29 Or mRow = 30 Or mRow = 31 Then
'                        If left(XyzAbc, 1) <> Chr(12) Then varTxtstrm.ReadLine
'                    End If
'                End If
'                If left(XyzAbc, 1) = Chr(12) Then
'                    If mPAGE = 1 Then
'                        '
'                    End If
'                    mPAGE = mPAGE + 1
'                    mRow = 1
'                    'varTxtstrmWithoutEject.Write XyzAbc + vbCrLf
'                Else
'                    varTxtstrmWithoutEject.Write XyzAbc + vbCrLf
'                End If
'            Loop
'        End If
'        mFILE_NAME = Trim(mFILE_NAME) + "Ej"
'    End If
'    varTxtstrm.Close
'    varTxtstrmWithoutEject.Close
'
'        Set varTxtstrm = fob.OpenTextFile(Trim(mFILE_NAME) + ".TXT", ForAppending)
'        varTxtstrm.Write (Chr(12))
'        varTxtstrm.Close
        '''''
    Else
        Call FaPrint_Form_Chk("* " + CAPTION + " * ")
        For mReportCount = 1 To mREPORT.FormulaFields.Count
            Select Case UCase(mREPORT.FormulaFields(mReportCount).FormulaFieldName)
                Case UCase("comp_name")
                    mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_Name & "'"
                Case UCase("comp_add1")
                    mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_Add & "'"
                Case UCase("comp_pin")
                    mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_City & "'"
            End Select
        Next
        If Index = 0 Then
            rpt_form.CAPTION = "* " + CAPTION + " *"
            rpt_form.Rep_Set = mREPORT
        Else
            mREPORT.PrintOut
        End If
    End If
Set rpt_form = Nothing
Set mREPORT = Nothing
Set connectionId = Nothing
Set fob = Nothing
Exit Sub
ERRORHANDLER:  MsgBox err.Description, vbCritical
End Sub

Public Sub FaSelGridKeyPress(txt As Object, FGrid As Object, Rst As ADODB.Recordset, ByVal KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As String, Optional CellBackColLeave As String)
Dim LPlace As Byte, FindStr$   ' As String
'    If FAFilterKeyCode(KeyAscii) = True Then Exit Sub
If FGrid.Rows < 1 Then Exit Sub
If Rst.RecordCount <= 0 Then txt.TEXT = "": Exit Sub
If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then Exit Sub
If KeyAscii = vbKeyBack Then
    If Len(txt.SelText) > 1 Then
        txt.SelLength = Len(txt.SelText) - 1
        FindStr = txt.SelText
    Else
        txt.TEXT = ""
        FGrid.SetFocus
        txt.Visible = False
        Exit Sub
    End If
Else
    FindStr = txt.SelText + Chr(KeyAscii)
End If
FindStr = Replace(FindStr, "'", "''")
Rst.MoveFirst
If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
    Rst.FIND "" & FindFldName & " >=" & Val(FindStr) & ""
Else    'character serach
    Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
End If
KeyAscii = 0
If Rst.AbsolutePosition <> adPosEOF And Rst.AbsolutePosition <> adPosBOF Then
    FGrid.CellBackColor = CellBackColLeave
    FGrid.Row = Rst.AbsolutePosition
    FGrid.CellBackColor = CellBackColEnter
    txt.TEXT = Rst.Fields(FindFldName).Value
    txt.SelLength = Len(FindStr)
    txt.left = FGrid.CellLeft + FGrid.left
    txt.top = FGrid.CellTop + FGrid.top
    If txt.Visible = False Then
        txt.Visible = True
        txt.ZOrder 0
        txt.SetFocus
        txt.BackColor = FGrid.CellBackColor
        txt.ForeColor = FGrid.CellForeColor
        txt.width = FGrid.CellWidth
        txt.height = FGrid.CellHeight
    End If
End If
End Sub
Public Function FaNavigationKey(KeyCode As Integer) As Boolean
FaNavigationKey = False
If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
    FaNavigationKey = True
End If
End Function
Public Sub FaPrint_Form_Chk(CAP_STRING As String)
Dim Z As Byte
For Z = 0 To Forms.Count - 1
    If UCase(Forms(Z).CAPTION) = UCase(CAP_STRING) Then
        Unload Forms(Z)
    End If
Next Z
End Sub
Public Function FaSetW(mSTRING As String, mLEN As Integer) As String
    mSTRING = mID(mSTRING, 1, mLEN)
    FaSetW = Trim(mSTRING) + Space(mLEN - Len(Trim(mSTRING)))
End Function
Public Function FaSetN(mSTRING As String, mLEN As Integer) As String
    mSTRING = mID(mSTRING, 1, mLEN)
    FaSetN = Space(mLEN - Len(Trim(mSTRING))) + Trim(mSTRING)
End Function
Public Sub AddNewFieldForDatamanFa(DataPath As String, Optional PassWord As String)
On Error GoTo errorbox
    Dim Cat As New ADOX.Catalog, PubDatamanFa As New DMFa.ClsFa, TinTin As Boolean
    If Not IsMissing(PassWord) Then
        If Trim(PassWord) <> "" Then
            Cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=" & PassWord & ";Data Source=" & DataPath
        Else
            Cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataPath
        End If
    Else
        Cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataPath
    End If
    PubDatamanFa.FaBackEnd = PubBackEnd
    TinTin = PubDatamanFa.DMFaStructureCheck(Cat)
    Set PubDatamanFa = Nothing
    Set Cat = Nothing
    Exit Sub
errorbox:       MsgBox err.Description, vbInformation
End Sub
Public Sub AddFieldTableFa(DB As DAO.Database, TableName As String, FieldName As String, FIELDTYPE As Variant, Optional FieldSize As Integer, Optional RequiredYesNo As Boolean, Optional AllowZero As Boolean, Optional DefValue As Variant)
Dim tmprs As DAO.Recordset
Dim N As Integer, TDF As TableDef, FLD As DAO.Field
    Set tmprs = DB.OpenRecordset("select * from " & TableName)
    For N = 0 To tmprs.Fields.Count - 1
        If UCase(tmprs.Fields(N).Name) = UCase(FieldName) Then
            GoTo Myexit
        End If
    Next
    Set tmprs = Nothing
    Set TDF = DB.TableDefs(TableName)
    Set FLD = TDF.CreateField(FieldName)
    FLD.Type = FIELDTYPE
    If FIELDTYPE = 10 Then      '' Text Field
        FLD.Size = FieldSize
        If Not IsMissing(AllowZero) Then FLD.AllowZeroLength = AllowZero
    End If
    If Not IsMissing(DefValue) Then FLD.DefaultValue = DefValue
    If Not IsMissing(RequiredYesNo) Then FLD.Required = RequiredYesNo
    TDF.Fields.Append FLD
Myexit:
    Set tmprs = Nothing
End Sub
Public Function FaIsValid(ctlName As Object, mfldname As String) As Boolean
    If Len(Trim(ctlName.TEXT)) = 0 Then
        MsgBox mfldname & " is a required field.", vbCritical + vbDefaultButton1, "Validation Error"
        ctlName.SetFocus
        FaIsValid = False
    Else
        FaIsValid = True
    End If
End Function
Public Function FaFilterString(STR As String) As String
Dim Str1$, LEN1%, X%, Str2$
    FaFilterString = Replace(STR, " ", "")
    LEN1 = Len(FaFilterString)
    X = 1
    While LEN1 > 0
        Str1 = mID(FaFilterString, X, 1)
        If (Str1 >= Chr(65) And Str1 <= Chr(90)) Or (Str1 >= Chr(97) And Str1 <= Chr(122)) Or (Str1 >= Chr(48) And Str1 <= Chr(57)) Then
            Str2 = Str2 & Str1
        End If
        X = X + 1
        LEN1 = LEN1 - 1
    Wend
    FaFilterString = UCase(Str2)
End Function
Public Sub FaNumDown(ByRef TEXT As Object, KeyCode As Integer, LeftPlace As Integer, RightPlace As Integer)
    If KeyCode = 46 Then
        If InStr(TEXT, "-") <> 0 And mID(TEXT, TEXT.SelStart + 1, 1) = "." And Len(TEXT) - 1 - RightPlace >= LeftPlace Then
            KeyCode = 0
        ElseIf InStr(TEXT, "-") = 0 And mID(TEXT, TEXT.SelStart + 1, 1) = "." And Len(TEXT) - RightPlace >= LeftPlace Then
            KeyCode = 0
        End If
    End If
End Sub
Public Sub FaNumPress(ByRef TEXT As Object, KeyAscii As Integer, LeftPlace As Integer, RightPlace As Integer)
On Error Resume Next
Dim myString As String
    If RightPlace = 0 Then myString = "0123456789-" & TEXT.Tag Else myString = "0123456789.-" & TEXT.Tag
    If KeyAscii > 26 Then
       If InStr(myString, Chr(KeyAscii)) = 0 Then KeyAscii = 0
       If (InStr(TEXT, "-") <> 0) And KeyAscii = 45 Then KeyAscii = 0
       If InStr(TEXT, ".") <> 0 Then
            If KeyAscii = 46 Then KeyAscii = 0
            If InStr(TEXT, "-") <> 0 Then
                If InStr(TEXT, ".") - 1 > LeftPlace And TEXT.SelStart < InStr(TEXT, ".") Then
                    KeyAscii = 0
                ElseIf Len(TEXT) >= InStr(TEXT, ".") + RightPlace And TEXT.SelStart >= InStr(TEXT, ".") Then
                    KeyAscii = 0
                End If
            Else
                If InStr(TEXT, ".") > LeftPlace And TEXT.SelStart < InStr(TEXT, ".") Then
                    KeyAscii = 0
                ElseIf Len(TEXT) >= InStr(TEXT, ".") + RightPlace And TEXT.SelStart >= InStr(TEXT, ".") Then
                    KeyAscii = 0
                End If
            End If
       Else
            If KeyAscii = 46 Then Exit Sub
            If InStr(TEXT, "-") <> 0 Then
                If Len(TEXT) - 1 >= LeftPlace Then KeyAscii = 0
            Else
                If Len(TEXT) >= LeftPlace And KeyAscii <> 45 Then KeyAscii = 0
            End If
       End If
    ElseIf KeyAscii = 8 And InStr(TEXT, "-") <> 0 And mID(TEXT, TEXT.SelStart, 1) = "." And mID(TEXT, TEXT.SelStart + 1, 1) <> "" And Len(TEXT) - 1 - RightPlace >= LeftPlace Then
        KeyAscii = 0
    ElseIf KeyAscii = 8 And InStr(TEXT, "-") = 0 And mID(TEXT, TEXT.SelStart, 1) = "." And mID(TEXT, TEXT.SelStart + 1, 1) <> "" And Len(TEXT) - RightPlace >= LeftPlace Then
        KeyAscii = 0
    End If
End Sub
Public Sub FaListViewReport_KeyDown(FrmList As Object, LV As Object, txt As Object, Index As Integer, KeyCode As Integer, Shift As Integer, left As Integer, top As Integer, width As Integer, Optional height As Integer)
If FaFilterKeyCode(KeyCode) = True Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If txt(Index).TEXT <> "" Then
            txt(Index).TEXT = LV.SelectedItem.TEXT
        End If
        FrmList.Visible = False
        Exit Sub
   End If
    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If FrmList.Visible = False Then Exit Sub
    Else
        If FrmList.Visible = False Then
            FrmList.left = left
            FrmList.top = top
            FrmList.width = width
            If IsMissing(height) Or height = 0 Then  'Updated by shekhar
                FrmList.height = LV.ListItems.Count * 270
            Else
                FrmList.height = height
            End If
            LV.left = 0
            LV.top = 0
            LV.width = width
            If IsMissing(height) Or height = 0 Then  'Updated by shekhar
                LV.height = LV.ListItems.Count * 270
            Else
                LV.height = height
            End If
            LV.ColumnHeaders(1).width = width
            LV.Tag = Index
            FrmList.Visible = True
            FrmList.ZOrder 0
        End If
    End If
    If FrmList.Visible = True Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            LV.Tag = Index
            If KeyCode = vbKeyUp And LV.SelectedItem.Index > 1 Then
                LV.ListItems(LV.SelectedItem.Index - 1).SELECTED = True
                txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.SelectedItem.Index < LV.ListItems.Count Then
                LV.ListItems(LV.SelectedItem.Index + 1).SELECTED = True
                txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.ListItems.Count = 1 Then
                txt(Index).TEXT = LV.SelectedItem.TEXT
            End If
        End If
    End If
End Sub
Public Sub FaGridDblClick(myForm As Form, FGrid As Object, txt As Object, Index As Integer)
On Error GoTo err
Dim I As Integer, g_Row As Integer, g_Col As Integer
If myForm.TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid.CellBackColor = CellBackColLeave
g_Row = FGrid.Row
g_Col = FGrid.Col
txt(Index).height = FGrid.CellHeight - 10
txt(Index).width = FGrid.CellWidth - 10
txt(Index).left = FGrid.CellLeft + FGrid.left
txt(Index).top = FGrid.CellTop + FGrid.top
txt(Index).TEXT = FGrid.TextMatrix(g_Row, g_Col)
txt(Index).Visible = True
txt(Index).ZOrder 0
txt(Index).Tag = FGrid.TextMatrix(g_Row, g_Col)
txt(Index).SetFocus
Exit Sub
err:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Public Sub FaGridTxtDown(FGrid As Object, txt As Object, Index As Integer, KeyCode As Integer, TAddMode As Boolean, MaxCol As Byte, Optional SkipCol As Byte, Optional MoveToCol As Byte, Optional DisableGridSrlNo As Boolean)
Dim GCol As Byte
If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    If TAddMode = True Then
        txt(Index).Visible = False
        Select Case KeyCode
            Case vbKeyUp
                If FGrid.Row > 1 Then
                    FGrid.Row = FGrid.Row - 1
                    FGrid.SetFocus
                Else
                    FGrid.SetFocus
                End If
            Case vbKeyDown
                If FGrid.Row < FGrid.Rows - 1 Then
                    FGrid.Row = FGrid.Row + 1
                    FGrid.SetFocus
                Else
                    FGrid.SetFocus
                 End If
            Case vbKeyLeft
                If FGrid.Col > 1 Then
                    FGrid.Col = FGrid.Col - 1
                    FGrid.SetFocus
                Else
                   FGrid.SetFocus
                End If
            Case vbKeyRight
                If FGrid.Col < MaxCol Then
                    FGrid.Col = FGrid.Col + 1
                    FGrid.SetFocus
                Else
                   FGrid.SetFocus
                End If
        End Select
    End If
ElseIf KeyCode = vbKeyReturn Then
    If IsMissing(MoveToCol) Or MoveToCol = 0 Then
        If FGrid.Col + IIf(IsMissing(SkipCol), 0, SkipCol) < MaxCol Then
            FGrid.Col = FGrid.Col + 1 + IIf(IsMissing(SkipCol), 0, SkipCol)
            FGrid.SetFocus
        Else
            If IsMissing(DisableGridSrlNo) Or (Not IsMissing(DisableGridSrlNo) And DisableGridSrlNo = False) Then
                If FGrid.Row = FGrid.Rows - 1 Then FGrid.AddItem FGrid.Rows
            Else
                If FGrid.Row = FGrid.Rows - 1 Then FGrid.AddItem
            End If
            FGrid.Row = FGrid.Row + 1
            For GCol = 1 To FGrid.Cols - 1
                If FGrid.ColWidth(GCol) <> 0 Then Exit For
            Next
            FGrid.Col = GCol
            FGrid.SetFocus
        End If
    Else
        If FGrid.ColWidth(MoveToCol) = 0 Then
            For GCol = MoveToCol To FGrid.Cols - 1
                If FGrid.ColWidth(GCol) <> 0 Then Exit For
            Next
        End If
        FGrid.Col = MoveToCol
        FGrid.SetFocus
    End If
End If
End Sub
Public Sub FaGet_Text(myForm As Form, FGrid As Object, txt As Object, Index As Integer, NumericColNature As Boolean, KeyAscii As Integer)
Dim I As Integer, J As Integer, g_Row As Integer, g_Col As Integer
If myForm.TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyAscii = vbKeyReturn Then
    FaGridDblClick myForm, FGrid, txt, Index
ElseIf KeyAscii = vbKeyDelete Then
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
ElseIf (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 8 Then
    FGrid.CellBackColor = CellBackColLeave
    g_Row = FGrid.Row
    g_Col = FGrid.Col
    txt(Index).height = FGrid.CellHeight - 10
    txt(Index).width = FGrid.CellWidth - 10
    txt(Index).left = FGrid.CellLeft + FGrid.left
    txt(Index).top = FGrid.CellTop + FGrid.top
    txt(Index).TEXT = ""
    txt(Index).Visible = True
    txt(Index).ZOrder 0
    txt(Index).Tag = FGrid.Col
    txt(Index).SetFocus
    If NumericColNature = True Then
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Then
            txt(Index).TEXT = Chr(KeyAscii)
        End If
        GoTo NXT
    End If
    If KeyAscii = vbKeyBack Then
        txt(Index).TEXT = ""
    Else
        txt(Index).TEXT = Chr(KeyAscii)
    End If
NXT:
    txt(Index).SelStart = 1
End If
Exit Sub
err:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub

Public Sub FaCtrl_validate(Ctrl As Object)
    Ctrl.BackColor = CtrlBColOrg
    Ctrl.ForeColor = CtrlFColOrg
End Sub
Public Sub FaCtrl_GetFocus(Ctrl As Object)
    Ctrl.BackColor = CtrlBCol
    Ctrl.ForeColor = CtrlFCol
End Sub
Public Sub FaCtrl_DownKeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 144 Then Exit Sub
If KeyCode = 13 Or KeyCode = 40 Then
    SendKeys vbTab
    KeyCode = 0
End If
End Sub
Public Sub FaCtrl_UpKeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 144 Then Exit Sub
If KeyCode = 38 Then     'keyup = 38
    SendKeys "+{Tab}"
    KeyCode = 0
    Shift = 0
End If
End Sub
Public Sub FaDGridTxtKeyDown_Mast(DGrid As Object, txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, GridText As Boolean, Optional HelpIndex As Integer)
Dim I As Integer
'If Rst.RecordCount > 0 Then
If FaFilterKeyCode(KeyCode) = True Then Exit Sub
    If KeyCode = vbKeyReturn Then
        DGrid.Visible = False
        Exit Sub
    End If
    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If DGrid.Visible = False Then Exit Sub
    Else
        If DGrid.Visible = False Then DGrid.Visible = True: DGrid.ZOrder 0
    End If
    If DGrid.Visible = True Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            Select Case KeyCode
                Case vbKeyUp
                    If Rst.RecordCount > 0 Then If Rst.BOF = False Then Rst.MovePrevious Else Rst.MoveFirst
                Case vbKeyDown
                    If Rst.RecordCount > 0 Then If Rst.EOF = False Then Rst.MoveNext Else Rst.MoveLast
                Case vbKeyPageUp '33
                    For I = 1 To 10
                        If Rst.AbsolutePosition > 1 Then Rst.MovePrevious
                    Next
                Case vbKeyPageDown '34
                    For I = 1 To 10
                        If Rst.AbsolutePosition < Rst.RecordCount Then Rst.MoveNext
                    Next
            End Select
'            If Rst.BOF = False And Rst.EOF = False Then
'                If GridText = True Then
'                Select Case HelpIndex
'                    Case 1
'                        txt(Index).Text = Rst!Code
'                    Case 2
'                        txt(Index).Text = Rst!Name
'                    Case 3
'                        txt(Index).Text = Rst!LName
'                End Select
'                Else
'                    txt(Index).Text = Rst!Name
'                End If
'            End If
      End If
      Exit Sub
  End If
'Else
'  If DGrid.Visible = False Then DGrid.Visible = True: DGrid.ZOrder 0
'End If
End Sub
Public Sub FaDGridTxtKeyUp_Mast(txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, FindFldName As String)
Dim STR$    ' As String
Dim LPlace As Byte
    If Rst.RecordCount <= 0 Then Exit Sub
    If KeyCode = 13 Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Exit Sub
    LPlace = txt(Index).SelStart
    STR = mID(txt(Index).TEXT, 1, LPlace)
    Rst.MoveFirst
    If Rst.Fields(FindFldName).Type = adInteger Then
        Rst.FIND "" & FindFldName & " >=" & Val(STR) & ""
    Else
        Rst.FIND "" & FindFldName & " >='" & STR & "'"
    End If
    If Rst.EOF = True Then Rst.MoveFirst
End Sub
Public Sub FaDGridTxtKeyUp(DGrid As Object, txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, FindFldName As String)
Dim LPlace As Byte, FindStr$
    If Rst.RecordCount <= 0 Then txt(Index).TEXT = "": Exit Sub
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Exit Sub
    LPlace = txt(Index).SelStart
    FindStr = txt(Index).TEXT
    Rst.MoveFirst
    If Rst.Fields(FindFldName).Type = adInteger Then
        Rst.FIND "" & FindFldName & " >=" & Val(FindStr) & ""
        If Rst.AbsolutePosition = adPosEOF Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & " >=" & Val(FindStr) & ""
            txt(Index).TEXT = FindStr
        ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & " >=" & Val(FindStr) & ""
            txt(Index).TEXT = FindStr
        End If
    Else
        Rst.FIND "" & FindFldName & " >='" & FindStr & "'"
        If Rst.AbsolutePosition = adPosEOF Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & " >='" & FindStr & "'"
            txt(Index).TEXT = FindStr
        ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & " >='" & FindStr & "'"
            txt(Index).TEXT = FindStr
        End If
    End If
    txt(Index).SelStart = Len(txt(Index).TEXT)
End Sub
Public Sub FaDGridTxtKeyDown(DGrid As Object, txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, GridText As Boolean, Optional HelpIndex As Integer)
Dim I As Integer, xString As String
On Error GoTo ELoop
If Rst.RecordCount > 0 Then
    If FaFilterKeyCode(KeyCode) = True Then Exit Sub
    If IsMissing(HelpIndex) Then HelpIndex = 1
    If KeyCode = vbKeyReturn Then
       If txt(Index).TEXT <> "" Then
            If Rst.BOF = False And Rst.EOF = False Then
                If GridText = True Then
                    txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                Else
                    txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                    txt(Index).Tag = Rst.Fields(0).Value
                End If
            End If
        End If
        DGrid.Visible = False
        Exit Sub
    End If
    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If DGrid.Visible = False Then Exit Sub
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
        txt(Index).SelStart = Len(txt(Index).TEXT)
        If DGrid.Visible = False Then DGrid.Visible = True:   DGrid.ZOrder 0
        If KeyCode <> vbKeyBack Then KeyCode = 0
    ElseIf KeyCode = vbKeyDelete Then
        txt(Index).TEXT = ""
    Else
        If DGrid.Visible = False Then
            txt(Index).SelStart = Len(txt(Index).TEXT)
            DGrid.Visible = True: DGrid.ZOrder 0
        End If
    End If
    If DGrid.Visible = True Then
        If Rst.RecordCount = 0 Then Exit Sub
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            Select Case KeyCode
                Case vbKeyUp
                    If Rst.AbsolutePosition > 1 Then
                        Rst.MovePrevious
                    Else
                        KeyCode = 0
                    End If
                Case vbKeyDown
                    If Rst.AbsolutePosition < Rst.RecordCount Then
                        If Rst.EOF = False Then Rst.MoveNext
                    End If
                Case vbKeyPageUp '33
                    For I = 1 To (Int(DGrid.height / DGrid.RowHeight) - 2)
                        If Rst.AbsolutePosition > 1 Then Rst.MovePrevious
                    Next
                Case vbKeyPageDown '34
                    For I = 1 To (Int(DGrid.height / DGrid.RowHeight) - 2)
                        If Rst.AbsolutePosition < Rst.RecordCount And Rst.EOF = False Then Rst.MoveNext
                    Next
            End Select
            If Rst.BOF = False And Rst.EOF = False Then
                If GridText = True Then
                    txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                Else
                    txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                    txt(Index).Tag = Rst.Fields(0).Value
                End If
                txt(Index).SelStart = Len(txt(Index))
            End If
        End If
        Exit Sub
    End If
Else
    If DGrid.Visible = False Then DGrid.Visible = True: DGrid.ZOrder 0
End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub

Public Sub FaListView_KeyUp(LV As Object, txt As Object, Index As Integer, KeyCode As Integer, xITEM As ListItem)
Dim STR As String, LPlace As Integer
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Exit Sub
LPlace = txt(Index).SelStart
STR = mID(txt(Index).TEXT, 1, LPlace)
If LV.Visible = True Then
    Set xITEM = LV.FindItem(STR, 0, , 1)
    If xITEM Is Nothing Then
        Exit Sub
    Else
        xITEM.EnsureVisible
        xITEM.SELECTED = True
    End If
End If
End Sub
Public Sub FaListView_KeyDown(FrmList As Object, LV As Object, txt As Object, Index As Integer, KeyCode As Integer, Shift As Integer, left As Integer, top As Integer, width As Integer, height As Integer)
If KeyCode = 144 Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If txt(Index).TEXT <> "" Then
            txt(Index).TEXT = LV.SelectedItem.TEXT
        End If
        FrmList.Visible = False
        Exit Sub
   End If
    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If FrmList.Visible = False Then Exit Sub
    Else
        If FrmList.Visible = False Then
            FrmList.left = left
            FrmList.top = top
            FrmList.width = width
            FrmList.height = height
            LV.left = 0
            LV.top = 0
            LV.width = width
            LV.height = height
            LV.ColumnHeaders(1).width = width
            LV.Tag = Index
            FrmList.Visible = True
            FrmList.ZOrder 0
        End If
    End If
    If FrmList.Visible = True Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            LV.Tag = Index
            If KeyCode = vbKeyUp And LV.SelectedItem.Index > 1 Then
                LV.ListItems(LV.SelectedItem.Index - 1).SELECTED = True
                txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.SelectedItem.Index < LV.ListItems.Count Then
                LV.ListItems(LV.SelectedItem.Index + 1).SELECTED = True
                txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.ListItems.Count = 1 Then
                txt(Index).TEXT = LV.SelectedItem.TEXT
            End If
        End If
    End If
End Sub
Public Function FaListView_Items(LV As Object, txt As Object, Index As Integer, list_item As Variant, cnt As Integer) As ListItem
Dim xName As ListItem, I As Integer
    LV.ListItems.Clear
    For I = 0 To cnt - 1
         Set xName = LV.ListItems.Add(I + 1, , list_item(I))
    Next
    Set xName = LV.FindItem(txt(Index), 0, , 1)
    If xName Is Nothing Then
        Exit Function
    Else
        xName.EnsureVisible
        xName.SELECTED = True
    End If
    Set FaListView_Items = xName
End Function
 Public Function FaListView_Items_RecordSet(LV As Object, txt As Object, Index As Integer, Rst As ADODB.Recordset) As ListItem
    Dim xName As ListItem
    Dim I As Long
    LV.ListItems.Clear
    If Rst.RecordCount <= 0 Then Exit Function
    Do Until Rst.EOF
         Set xName = LV.ListItems.Add(, , Rst.Fields("Name").Value)
         xName.SubItems(1) = Rst.Fields("code").Value
         Rst.MoveNext
    Loop
    Set xName = LV.FindItem(txt(Index), 0, , 1)
    If xName Is Nothing Then
        Exit Function
    Else
        xName.EnsureVisible
        xName.SELECTED = True
    End If
    Set FaListView_Items_RecordSet = xName
End Function
Public Function FaFilterKeyCode(KeyCode As Integer) As Boolean
'Alter =18, WindowsStartUp = 91, CapsLock=vbKeyCapital=20, Shift =16
If (KeyCode = vbKeyControl Or KeyCode = vbKeyShift _
    Or KeyCode = vbKeyNumlock Or KeyCode = vbKeyCapital _
    Or KeyCode = vbKeyScrollLock Or KeyCode = 18 Or KeyCode = 91) Then    'And Shift = 0
    FaFilterKeyCode = True
    Exit Function
End If
FaFilterKeyCode = False
End Function
Public Sub FaFormKeyDown(FrmName As Form, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Or KeyCode = vbKeyF4 Or (KeyCode = 70 And Shift = 2) Or (KeyCode = 80 And Shift = 2) Or (KeyCode = 83 And Shift = 2) Or KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or KeyCode = vbKeyF10 Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEnd Then
    FrmName.TopCtrl1.TopKey_Down KeyCode, Shift
End If
If KeyCode <> vbKeyF10 Then
    If FrmName.TopCtrl1.PrvKeyCode = vbKeyEscape Then
        FrmName.TopCtrl1.PrvKeyCode = 0
    Else
        FrmName.TopCtrl1.PrvKeyCode = KeyCode
    End If
End If
End Sub
Public Sub FaDGridTxtKeyPress(txt As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyCode As Integer, FindFldName As String, Optional KeyUpCall As Boolean)
Dim FindStr$    ' As String
Dim LPlace As Byte
    If Rst.RecordCount <= 0 Then txt(Index).TEXT = "": Exit Sub
    If KeyCode = 0 Then Exit Sub
    If KeyCode = 13 Or KeyCode = 8 Or KeyCode = 16 Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then: Exit Sub
    If IsMissing(KeyUpCall) Or KeyUpCall = False Then 'KeyPressCall
        If txt(Index).TEXT = "" Then
            FindStr = Chr(KeyCode)
        Else
            FindStr = txt(Index).TEXT + Chr(KeyCode)
        End If
        If FindStr = "" Or FindStr = "*" Or FindStr = "_" Then KeyCode = 0: Exit Sub
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
                FindStr = Val(FindStr)
                Rst.FIND "" & FindFldName & " like " & FindStr & "*"
           If Rst.AbsolutePosition = adPosEOF Then
                FindStr = txt(Index).TEXT   'left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                If FindStr <> "" Then Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
           ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                FindStr = txt(Index).TEXT   'left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                Rst.FIND "" & FindFldName & " like " & FindStr & "*"
            End If
        Else    'character serach
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
             If Rst.AbsolutePosition = adPosEOF Then
                FindStr = txt(Index).TEXT
                Rst.MoveFirst
                If FindStr <> "" Then Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
            ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                FindStr = txt(Index).TEXT
                Rst.MoveFirst
               Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
            End If
        End If
        If FindStr = txt(Index).TEXT + Chr(KeyCode) Then
            txt(Index).TEXT = txt(Index).TEXT + Chr(KeyCode)
        End If
        KeyCode = 0
    Else    'KeyUp Call Search as per Old Process
        LPlace = txt(Index).SelStart
        FindStr = txt(Index).TEXT
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then
        FindStr = Val(FindStr)
            Rst.FIND "" & FindFldName & " like " & FindStr & "*"
            If Rst.AbsolutePosition = adPosEOF Then
                FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                If FindStr <> "" Then Rst.FIND "" & FindFldName & " like " & FindStr & "*"
                txt(Index).TEXT = FindStr
            ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                If FindStr <> "" Then Rst.FIND "" & FindFldName & " like " & FindStr & "*"
                txt(Index).TEXT = FindStr
            End If
        Else
            If FindStr <> "" Then
                Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
                If Rst.AbsolutePosition = adPosEOF Then
                    FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                    Rst.MoveFirst
                    If FindStr <> "" Then Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
                    txt(Index).TEXT = FindStr
                ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                    FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                    Rst.MoveFirst
                    If FindStr <> "" Then Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
                    txt(Index).TEXT = FindStr
                End If
            End If
        End If
        txt(Index).SelStart = Len(txt(Index).TEXT)
        KeyCode = 0
    End If
    txt(Index).SelStart = Len(txt(Index))
End Sub
Public Function FaCurrBal(xSubCode As String) As String
Dim mAmtDr As Double, mAmtCr As Double
    If Len(xSubCode) <= 0 Then Exit Function
    If PubBackEnd = "A" Then
        mAmtCr = G_FaCn.Execute("SELECT iif(isnull(sum(ledger.Amtcr)),0,sum(ledger.Amtcr)) AS OpRec FROM ledger where subcode='" & xSubCode & "'").Fields(0)
        mAmtDr = G_FaCn.Execute("SELECT iif(isnull(sum(ledger.Amtdr)),0,sum(ledger.Amtdr)) AS OpIss FROM ledger where subcode='" & xSubCode & "'").Fields(0)
    ElseIf PubBackEnd = "S" Then
        mAmtCr = G_FaCn.Execute("SELECT isnull(sum(ledger.Amtcr),0) AS OpRec FROM ledger where subcode='" & xSubCode & "'").Fields(0)
        mAmtDr = G_FaCn.Execute("SELECT isnull(sum(ledger.Amtdr),0) AS OpIss FROM ledger where subcode='" & xSubCode & "'").Fields(0)
    End If
    FaCurrBal = Format(Abs(mAmtDr - mAmtCr), "0.00") + " " + IIf((mAmtDr - mAmtCr) > 0, "Dr", "Cr")
End Function
Public Sub FaIniCombo(sqlstr As String, DBCNAME As DataCombo, LSTFIELD As String, BNDCOLUMN As String)
On Error GoTo errorbox
    Set DBCNAME.RowSource = G_FaCn.Execute(sqlstr)
    DBCNAME.ListField = LSTFIELD
    DBCNAME.BoundColumn = BNDCOLUMN
    DBCNAME.Tag = sqlstr
    Exit Sub
errorbox:       If err.NUMBER > 0 Then MsgBox err.Description, vbInformation
End Sub
Public Sub FaRefrCombo(DBC As DataCombo)
    Dim BT
    BT = DBC.BoundText
    FaIniCombo DBC.Tag, DBC, DBC.ListField, DBC.BoundColumn
    DBC.BoundText = BT
End Sub
Public Sub FaRstBofEof(ByRef Rst As ADODB.Recordset)
    If Rst.RecordCount > 0 Then
        If Rst.BOF Then Rst.MoveFirst
        If Rst.EOF Then Rst.MoveLast
    End If
End Sub
Public Sub FaCheckQuote(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
End Sub
Public Sub FaWinSetting(ByRef FrmName As Form)
On Error Resume Next
With FrmName
    .Visible = False
    .height = 8000
    .width = 12000
    .top = 0
    .left = 0
    .TopBar1.BackColor = .BackColor
End With
End Sub
Public Function FaValidDate(FRMNAME1 As Form) As Integer
FaValidDate = 1
If FaValidDateChk(FRMNAME1.TXTS_DATE, "Starting Date") = False Then FaValidDate = 0: Exit Function
If FaValidDateChk(FRMNAME1.TXTE_DATE, "Ending Date") = False Then FaValidDate = 0: Exit Function
If DateDiff("d", FRMNAME1.TXTS_DATE, FRMNAME1.TXTE_DATE) < 0 Then
    MsgBox " Ending Date Less than Starting Date ", vbCritical
    FaValidDate = 0
End If
End Function
Public Function FaValidDateChk(TXT_DATE As Date, mfldname As String) As Boolean
FaValidDateChk = True
If DateDiff("D", PubStartDate, TXT_DATE) < 0 Then
    MsgBox mfldname + " is Before Financial Year ", vbCritical
    FaValidDateChk = False
ElseIf DateDiff("D", TXT_DATE, PubEndDate) < 0 Then
    MsgBox mfldname + " is After Financial Year ", vbCritical
    FaValidDateChk = False
End If
End Function
Public Function FaNToW(ByRef NN As Double, mmajor As String, mminor As String) As String
Dim ps As Long, nums As Long, NUMBER As Long, I As Integer, cn As Long
Dim X As Variant
    Unit = "One  Two  ThreeFour Five Six  Seveneightnine " '5
    tens = "Eleven   Twelve   Therteen Fourteen Fifteen  Sixteen  SeventeenEighteen Nineteen " '9
    tenth = "Ten    Twenty Thirty Fourty Fifty  Sixty  SeventyEighty Ninty  " '7
    PLACE = "Hundred ThousandLacs    Crore   " '8
    cn = 10000000
    WORDs = mmajor
    NUMBER = Int(NN)
    ps = ((NN) - Int(NN)) * 100
    For I = 1 To 5
        nums = Int(NUMBER / cn)
        NUMBER = NUMBER - (nums * cn)
        If I <> 3 Then cn = cn / 100 Else cn = cn / 10
        If nums > 0 Then X = FaCONVERTs(nums, I)
    Next I
    WORDs = WORDs + mminor + " "
    If ps > 0 Then X = FaCONVERTs(ps, I) Else WORDs = WORDs + " Zero "
    WORDs = WORDs + " Only"
    FaNToW = WORDs
End Function
Private Function FaCONVERTs(ByRef NUM As Long, ByRef Index As Integer) As Boolean
    If NUM Mod 10 = 0 Then
        WORDs = WORDs + Space(1) + RTrim(mID(tenth, (((NUM \ 10) - 1) * 7) + 1, 7))
    ElseIf NUM < 10 Then
        WORDs = WORDs + Space(1) + RTrim(mID(Unit, ((NUM - 1) * 5) + 1, 5))
    ElseIf NUM < 20 Then
        WORDs = WORDs + Space(1) + RTrim(mID(tens, (((NUM - 10) - 1) * 9) + 1, 9))
    Else
        WORDs = WORDs + Space(1) + RTrim(mID(tenth, (((NUM \ 10) - 1) * 7) + 1, 7))
        WORDs = WORDs + Space(1) + RTrim(mID(Unit, (((NUM Mod 10) - 1) * 5) + 1, 5))
    End If
    If Index < 5 Then WORDs = WORDs + Space(1) + RTrim(mID(PLACE, ((4 - Index) * 8) + 1, 8))
End Function
