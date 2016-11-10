Attribute VB_Name = "MainLib"
Option Explicit
Dim Unit, tens, tenth, WORDs, PLACE As String

Public Enum ObjTypeDef
    Recordset = 0
    HFlexGrid = 1
    TextBox = 2
End Enum

Public Enum TxtAlignDef
    AlignLeft = 0
    AlignRight = 1
End Enum

Public Enum ObjTypeDef2
    leftSide = 0
    RightSide = 1
    FormLeft = 2
End Enum

Public Enum ObjTypeDef1
    Division_Code = 1
    Current_Site = 2
    For_Site_Code = 3
    Document_Type = 4
    Document_Prefix = 5
    Document_No = 6
End Enum

Public Enum ObjTypeDefChas
    ChasType = 1
    MfgMonth = 2
    MfgYear = 3
    ChasSerialNo = 4
End Enum

Public Enum DataTypeDef
    CharacterType = 0
    NumericType = 1
    DateType = 2
End Enum
Public Enum ObjTypeDefLab
    NoLabour = 0
    WithLabour = 1
End Enum
Public Enum ObjTypePerm
    ad = 1
    Ed = 2
    De = 3
    pr = 4
End Enum

Public Type LedgRec
'    DocId   As String * 21
'    V_Type  As String * 5
'    VNo     As Single
'    v_SNo   As Byte
'    V_Date  As Date
    SubCode As String   ' * 8
    AmtDr   As Double
    AmtCr   As Double
    ContraSub As String ' * 8
    Narration As String
    Chq_No As String
    Chq_Date As String
    EmpDetailYn As String
    Clg_Date As String
'    Site_Code As String * 2
'    U_Name    As String * 10
'    U_EntDt   As Date
'    U_AE      As String * 1
End Type

Dim lngTemp
Dim lngST_StartTime


Public Function AcPostAuthorisation(TxtStr As String) As Boolean

If PubAcPostingByAllUser = False Then
    If TxtStr <> "" Then
        If pubUAcPosting <> "Y" Then
            MsgBox "Edit/Delete denied " & vbCrLf & "A/c Posting Authorisation not found!", vbCritical, "Validation"
            Exit Function
        End If
    End If
End If
AcPostAuthorisation = True
End Function


''Disable Testbox PopUp Menu
'Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Const GWL_WNDPROC = -4
'Public Const WM_RBUTTONUP = &H205
'Public lpPrevWndProc As Long
'Private lngHWnd As Long
''eof popup menu section

Public Function UpdVouSrlNo(FACn As ADODB.Connection, DocID As String, VDate As Date) As Boolean
Dim Rst As ADODB.Recordset, TEMPSQL$, DivBaseNumber As Boolean, FaVoucher As Boolean
Dim VType$, vPrefix$, VNo As Long
On Error GoTo lblExit
'made at Udaipur
    VType = Trim(DeCodeDocID(DocID, Document_Type))
    vPrefix = Trim(DeCodeDocID(DocID, Document_Prefix))
    VNo = Val(DeCodeDocID(DocID, Document_No))
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    'Set Rst = FACn.Execute("Select distinct switch(Category='FA',True,Category<>'FA',False) as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
    Set Rst = FACn.Execute("Select distinct " & cIIF("Category='FA'", cBoolean(True), cBoolean(False)) & " as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
    FaVoucher = Rst!FaVoucher
    DivBaseNumber = IIf(Rst!DivBaseNumber = 0, False, True)
    If VType = "V_SB" Then  'Vehicle Sale Bill
        TEMPSQL = "Select top 1 VT.Number_Method,VT.V_Type,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join VehBill_Counter VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "' and VP.Prefix='" & vPrefix & "'"
        If DivBaseNumber Then
            TEMPSQL = TEMPSQL & " and VP.Div_Code='" & PubDivCode & "'"
        End If
        TEMPSQL = TEMPSQL & "Order By VP.Div_Code,VP.Date_From DESC"
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open TEMPSQL, FACn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            If DivBaseNumber Then
                FACn.Execute "Update VehBill_Counter Set Start_Srl_No=" & VNo & " Where V_Type='" & Rst!V_Type & "' and Div_Code='" & PubDivCode & "' and Prefix='" & vPrefix & "' and Start_Srl_No<" & VNo & ""
            Else
                FACn.Execute "Update VehBill_Counter Set Start_Srl_No=" & VNo & " Where V_Type='" & Rst!V_Type & "' and Prefix='" & vPrefix & "' and Start_Srl_No<" & VNo & ""
            End If
        End If
    Else
        TEMPSQL = "Select top 1 VT.Number_Method,VP.V_Type,VP.Date_From, VP.Date_To,VP.Prefix,VP.Start_Srl_No, IsNull(SiteBaseNumber,'N') as SiteBaseNumber From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "'"
        If DivBaseNumber Then
            TEMPSQL = TEMPSQL & " and VP.Div_Code='" & PubDivCode & "'"
        End If
        TEMPSQL = TEMPSQL & " And VP.Date_To>=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " And VP.Date_From<=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " Order By VP.Div_Code,VP.Date_From DESC"
        
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open TEMPSQL, FACn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            If DivBaseNumber Then
                FACn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & VNo & " Where V_Type='" & Rst!V_Type & "'   and Prefix='" & vPrefix & "'   and Div_Code='" & PubDivCode & "' and Date_From=" & ConvertDate(Format(Rst!Date_From, "dd/MMM/yyyy")) & " and Date_to=" & ConvertDate(Format(Rst!Date_to, "dd/MMM/yyyy")) & " and Start_Srl_No<" & VNo & "  aND Site_Code = (Case When '" & Rst!SiteBaseNumber & "'='Y' Then '" & PubSiteCode & "' Else Site_Code End) "
            Else
                FACn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & VNo & " Where V_Type='" & Rst!V_Type & "'   and Prefix='" & vPrefix & "'   and Date_From=" & ConvertDate(Format(Rst!Date_From, "dd/MMM/yyyy")) & "  and Date_to=" & ConvertDate(Format(Rst!Date_to, "dd/MMM/yyyy")) & " and Start_Srl_No<" & VNo & "  aND Site_Code = (Case When '" & Rst!SiteBaseNumber & "'='Y' Then '" & PubSiteCode & "' Else Site_Code End) "
            End If
        End If
    End If
    Set Rst = Nothing
    UpdVouSrlNo = True
    Exit Function
lblExit:
    Set Rst = Nothing
    UpdVouSrlNo = False
    MsgBox "Error in Serial No. updation", vbCritical, "Document Serial No"
End Function

Public Sub BlankText(Frm As Form)
Dim objctrl As Control
For Each objctrl In Frm.Controls
    If TypeOf objctrl Is TextBox Then
        objctrl.TEXT = ""
    End If
Next objctrl
End Sub

Public Sub Btn_Move_Rst(ByRef FRMNAME1 As Form, Rst As ADODB.Recordset, Index As Integer)
On Error GoTo LOOP1
Call RST_BOF_EOF(Rst)
If Rst.RecordCount > 0 Then
    With Rst
        Select Case Index
            Case 0
                .MoveFirst
            Case 1
                .MovePrevious
            Case 2
                .MoveNext
            Case 3
                .MoveLast
        End Select
    End With
    Call RST_BOF_EOF(Rst)
End If
Exit Sub
LOOP1:  MsgBox err.Description, vbCritical
End Sub

'Public Function roff(digit As Double)  'TO ROUND OFF THE NUMBER UPTO TWO PRECISSION
'    roff = Val(Format(Val(digit), "###0.00"))
'End Function

Public Function IsNegative(ByRef ctlName As Object) As Boolean
    If Val(ctlName.TEXT) < 0 Then
        MsgBox "Entered value is Negative", vbInformation, "Validation Error"
        ctlName.SetFocus
        IsNegative = True
    Else
        IsNegative = False
    End If
End Function

Public Function IsNegative2(ByRef ctlName As Object) As Boolean
    If Val(ctlName.TEXT) <= 0 Then
        MsgBox "Entered value should be greater than zero", vbInformation, "Validation Error"
        ctlName.SetFocus
        IsNegative2 = True
    Else
        IsNegative2 = False
    End If
End Function

Public Function IsCancel(ByRef BrowserObj As Object) As Boolean
    Select Case left(BrowserObj, 1)
    Case "B"
        IsCancel = False
    Case Else
        If MsgBox("Do You Want to Cancel Changes?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then IsCancel = True Else IsCancel = False
    End Select
End Function

'Public Function ChkDate(temp As Variant) As Variant
'    If LTrim(RTrim(temp)) = "" Or IsNull(temp) Then
'        ChkDate = "Null"
'    Else
'        ChkDate = "#" & temp & "#"
'    End If
'End Function

'Public Function ConvertDate(temp)
'    If IsNull(temp) Or temp = "" Then
'        '31-01-2001
'        ConvertDate = "Null"
'    Else
'        ConvertDate = "#" & Format(CDate(temp), "dd/MMM/yyyy") & "#"
'    End If
'End Function

'Public Sub UserPermission(FormName As Form)
'    If UserType = "Y" Then
'        FormName.TOPBAR1.tDel = True
'        FormName.TOPBAR1.tEdit = True
'        FormName.TOPBAR1.tPrn = True
'    Else
'        FormName.TOPBAR1.tDel = False
'        FormName.TOPBAR1.tEdit = False
'        FormName.TOPBAR1.tPrn = False
'    End If
'End Sub

Public Function Chk_Text(Temp As Variant) As Variant
Chk_Text = Temp
If IsNull(Chk_Text) Or Chk_Text = Null Then
    Chk_Text = "Null"
    Exit Function
End If
Chk_Text = "'" & Replace(Chk_Text, "'", "''") & "'"
End Function

'Public Function DNull(temp As Variant) As String
'    DNull = IIf(IsNull(temp) Or temp = "", "Null", temp)
'End Function

'Public Function NumNull(ByVal MyNum As Variant) As Double
'    If IsNull(MyNum) Or MyNum = "" Then
'         NumNull = 0
'    Else
'         NumNull = Val(MyNum)   'MyNum
'    End If
'End Function

Public Function StrNull(ByVal MyStr As Variant) As String
    If IsNull(MyStr) Or MyStr = "" Then
        StrNull = ""
    Else
        StrNull = MyStr
    End If
End Function

'Public Function UpperIn(ByRef TxtBoxRef As TextBox, ByRef KeyAscii As Integer) As Boolean
'    'THIS FUNCTION RETURNS UPPER CASE CHAR,LOWER CASE CHAR ACCORDING SPECIFIED IN TAG
'    If TxtBoxRef.Tag = "L" Then KeyAscii = Asc(LCase(Chr(KeyAscii))) Else KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    UpperIn = True
'End Function

Public Sub WinSetting(ByRef FrmName As Form, Optional frmHeight As Integer, Optional frmWidth As Integer, Optional frmTop As Integer, Optional frmLeft As Integer)
If frmHeight = 0 Then frmHeight = 8000 '7635
If frmWidth = 0 Then frmWidth = 11940

With FrmName
'    .Visible = False
'    .BackColor = FrmMain.BackColor
    .height = frmHeight
    .width = frmWidth
    .top = frmTop
    .left = frmLeft
    .WindowState = 0
'    .MaxButton = False 'Read Only Property
'    .TOPBAR1.BackColor = .BackColor
End With
End Sub

Public Sub RST_BOF_EOF(ByRef Rst As ADODB.Recordset)
    If Rst.RecordCount > 0 Then
        If Rst.BOF Then Rst.MoveFirst
        If Rst.EOF Then Rst.MoveLast
    End If
End Sub

Public Function Set_FontName(Index As Integer) As String
Select Case Index
    Case 1
        Set_FontName = "MS Sans Serif"
    Case 2
        Set_FontName = "Hitarth Hin Jalak"
End Select
End Function

Public Function Set_FontSize(Index As Integer) As Integer
Select Case Index
    Case 1
        Set_FontSize = 8
    Case 2
        Set_FontSize = 14
End Select
End Function

Public Function Set_RepName(Index As Integer) As String
Select Case Index
    Case 1
        Set_RepName = "English"
    Case 2
        Set_RepName = "Hindi"
End Select
End Function
Public Function IsValid(ctlName As Object, mfldname As String, Optional ProceedFurther As Boolean) As Boolean
    If Len(Trim(ctlName.TEXT)) = 0 Then
        MsgBox mfldname & " is a required field.", vbInformation, "Validation Error"
        If ProceedFurther = False Then If ctlName.Enabled = True Then ctlName.SetFocus
        IsValid = False
    Else
        IsValid = True
    End If
End Function

Public Sub FormExit(FrmName As Form)
    If MsgBox("Exit Yes/No", vbYesNo, "Close Form") = vbYes Then
        Unload FrmName
    End If
End Sub
'This Function Removes all Character Except (A to Z , a to z , 0 to 9)
'  >=65 <=90     A To Z
'  >=97 <=122    a To z
'  >=48 <=57     0 To 9
Public Function FilterString(STR As String) As String
Dim Str1$, LEN1%, X%, Str2$
    FilterString = Replace(STR, " ", "")
    LEN1 = Len(FilterString)
    X = 1
    While LEN1 > 0
        Str1 = mID(FilterString, X, 1)
        If (Str1 >= Chr(65) And Str1 <= Chr(90)) Or (Str1 >= Chr(97) And Str1 <= Chr(122)) Or (Str1 >= Chr(48) And Str1 <= Chr(57)) Then
            Str2 = Str2 & Str1
        End If
        X = X + 1
        LEN1 = LEN1 - 1
    Wend
    FilterString = UCase(Str2)
End Function

'This Function is used to validate the number textbox whether the user has pasted some character value in the textbox
Public Function NumValidate(Txt As TextBox) As Double
    If Txt = "" Or Txt.Tag = "" Then Exit Function
    NumValidate = IIf(Trim(Txt) = "" Or IsNumeric(Txt) = False, Txt.Tag, Val(Txt))
End Function

'This Function is used for Returning Valid date
''Public Function RetDate(ByRef txt As Object) As String
''On Error GoTo err1
''If Len(Trim(txt)) = 0 Then
''    RetDate = ""
''    Exit Function
''End If
''Dim mDay As Long, mMonth As String, mYear As Long, Txt1 As String, Test As Long
''        mDay = 0: mMonth = "": mYear = 0
''        Txt1 = Trim(txt)
''        '''' FOR DAY
''        Test = InStr(1, Txt1, "/")
''        If Test = 0 Then Test = InStr(1, Txt1, "-")
''        If Test = 0 Then Test = InStr(1, Txt1, ".")
''        If Test <> 0 Then
''            If IsNumeric(mID(Txt1, 1, Test - 1)) Then
''                mDay = Val(mID(Txt1, 1, Test - 1))
''            Else
''                mMonth = mID(Txt1, 1, Test - 1)
''            End If
''        End If
''        If Test = 0 Then
''            If IsNumeric(Txt1) Then
''                mDay = Val(Txt1)
''            Else
''                mMonth = Txt1
''            End If
''            GoTo EXITFLAG
''        End If
''        ''''' FOR MONTH
''        If mMonth = "" Then
''            Txt1 = mID(Txt1, Test + 1)
''            Test = InStr(1, Txt1, "/")
''            If Test = 0 Then Test = InStr(1, Txt1, "-")
''            If Test = 0 Then Test = InStr(1, Txt1, ".")
''            If Test <> 0 Then mMonth = mID(Txt1, 1, Test - 1)
''            If Test = 0 Then
''                mMonth = Txt1
''                GoTo EXITFLAG
''            End If
''        End If
''        ''''FOR YEAR
''        mYear = Val(mID(Txt1, Test + 1))
''EXITFLAG:
''        If mYear = 0 Then mYear = Year(date)
''        If mYear > 1999 Then mYear = Right(STR(mYear), 2)
''        If mYear < 10 Then
''            mYear = Val(mID(CStr(Year(date)), 1, 4 - Len(Trim(CStr(mYear)))) + Trim(CStr(mYear)) - 10)
''        Else
''            mYear = Val(mID(CStr(Year(date)), 1, 4 - Len(Trim(CStr(mYear)))) + Trim(CStr(mYear)))
''        End If
''        If mDay < 0 Then mDay = 0
''        mMonth = mID(mMonth, 1, 3)
''        Select Case Trim(UCase(mMonth))
''            Case "1", "01", "J", "JA", "JAN"
''                mMonth = "Jan"
''            Case "2", "02", "F", "FE", "FEB"
''                mMonth = "Feb"
''            Case "3", "03", "M", "MA", "MAR"
''                mMonth = "Mar"
''            Case "4", "04", "A", "AP", "APR"
''                mMonth = "Apr"
''            Case "5", "05", "MAY"
''                mMonth = "May"
''            Case "6", "06", "JU", "JUN"
''                mMonth = "Jun"
''            Case "7", "07", "JUL"
''                mMonth = "Jul"
''            Case "8", "08", "AU", "AUG"
''                mMonth = "Aug"
''            Case "9", "09", "S", "SE", "SEP"
''                mMonth = "Sep"
''            Case "10", "O", "OC", "OCT"
''                mMonth = "Oct"
''            Case "11", "N", "NO", "NOV"
''                mMonth = "Nov"
''            Case "12", "D", "DE", "DEC"
''               mMonth = "Dec"
''            Case Else
''                mMonth = Format(date, "MMM")
''        End Select
''        Select Case Trim(UCase(mMonth))
''            Case "1", "01", "J", "JA", "JAN"
''                If mDay > 31 Then mDay = 0
''            Case "2", "02", "F", "FE", "FEB"
''                If mDay > IIf(mYear Mod 4 = 0, 29, 28) Then mDay = 0
''            Case "3", "03", "M", "MA", "MAR"
''                If mDay > 31 Then mDay = 0
''            Case "4", "04", "A", "AP", "APR"
''                If mDay > 30 Then mDay = 0
''            Case "5", "05", "MAY"
''                If mDay > 31 Then mDay = 0
''            Case "6", "06", "JU", "JUN"
''                If mDay > 30 Then mDay = 0
''            Case "7", "07", "JUL"
''                If mDay > 31 Then mDay = 0
''            Case "8", "08", "AU", "AUG"
''                If mDay > 31 Then mDay = 0
''            Case "9", "09", "S", "SE", "SEP"
''                If mDay > 30 Then mDay = 0
''            Case "10", "O", "OC", "OCT"
''                If mDay > 31 Then mDay = 0
''            Case "11", "N", "NO", "NOV"
''                If mDay > 30 Then mDay = 0
''            Case "12", "D", "DE", "DEC"
''                If mDay > 31 Then mDay = 0
''            Case Else
''                mDay = 0
''        End Select
''        If mDay = 0 Then mDay = Day(Now)
''        RetDate = Format(Trim(STR(mDay)), "00") + "/" + Trim(mMonth) + "/" + Trim(STR(mYear))
''        Exit Function
''err1:
''    ' For Overflow Check
''    If err.NUMBER = 6 Then Resume Next
''End Function

Public Function RetDate(ByRef Txt As String) As String
On Error GoTo err1
If Len(Trim(Txt)) = 0 Then
    RetDate = ""
    Exit Function
End If
Dim mDay As Long, mMonth As String, mYear As Long, Txt1 As String, Test As Long
        mDay = 0: mMonth = "": mYear = 0
        Txt1 = Trim(Txt)
        '''' FOR DAY
        Test = InStr(1, Txt1, "/")
        If Test = 0 Then Test = InStr(1, Txt1, "-")
        If Test = 0 Then Test = InStr(1, Txt1, ".")
        If Test <> 0 Then
            If IsNumeric(mID(Txt1, 1, Test - 1)) Then
                mDay = Val(mID(Txt1, 1, Test - 1))
            Else
                mMonth = mID(Txt1, 1, Test - 1)
            End If
        End If
        If Test = 0 Then
            If IsNumeric(Txt1) Then
                mDay = Val(Txt1)
            Else
                mMonth = Txt1
            End If
            GoTo EXITFLAG
        End If
        ''''' FOR MONTH
        If mMonth = "" Then
            Txt1 = mID(Txt1, Test + 1)
            Test = InStr(1, Txt1, "/")
            If Test = 0 Then Test = InStr(1, Txt1, "-")
            If Test = 0 Then Test = InStr(1, Txt1, ".")
            If Test <> 0 Then mMonth = mID(Txt1, 1, Test - 1)
            If Test = 0 Then
                mMonth = Txt1
                GoTo EXITFLAG
            End If
        End If
        ''''FOR YEAR
        mYear = Format(Val(mID(Txt1, Test + 1)), "00")
EXITFLAG:
        If Val(mYear) = 0 Then mYear = Year(date)
        If mYear > 1999 Then mYear = Right(STR(mYear), 2)
        If mYear < 10 Then
            mYear = Val(mID(CStr(Year(date)), 1, 4 - Len(Trim(CStr(mYear)))) + Trim(CStr(mYear)) - 10)
        Else
            mYear = Val(mID(CStr(Year(date)), 1, 4 - Len(Trim(CStr(mYear)))) + Trim(CStr(mYear)))
        End If
        If mDay < 0 Then mDay = 0
        mMonth = mID(mMonth, 1, 3)
        Select Case Trim(UCase(mMonth))
            Case "1", "01", "J", "JA", "JAN"
                mMonth = "Jan"
            Case "2", "02", "F", "FE", "FEB"
                mMonth = "Feb"
            Case "3", "03", "M", "MA", "MAR"
                mMonth = "Mar"
            Case "4", "04", "A", "AP", "APR"
                mMonth = "Apr"
            Case "5", "05", "MAY"
                mMonth = "May"
            Case "6", "06", "JU", "JUN"
                mMonth = "Jun"
            Case "7", "07", "JUL"
                mMonth = "Jul"
            Case "8", "08", "AU", "AUG"
                mMonth = "Aug"
            Case "9", "09", "S", "SE", "SEP"
                mMonth = "Sep"
            Case "10", "O", "OC", "OCT"
                mMonth = "Oct"
            Case "11", "N", "NO", "NOV"
                mMonth = "Nov"
            Case "12", "D", "DE", "DEC"
               mMonth = "Dec"
            Case Else
                mMonth = Format(date, "MMM")
        End Select
        Select Case Trim(UCase(mMonth))
            Case "1", "01", "J", "JA", "JAN"
                If mDay > 31 Then mDay = 0
            Case "2", "02", "F", "FE", "FEB"
                If mDay > IIf(mYear Mod 4 = 0, 29, 28) Then mDay = 0
            Case "3", "03", "M", "MA", "MAR"
                If mDay > 31 Then mDay = 0
            Case "4", "04", "A", "AP", "APR"
                If mDay > 30 Then mDay = 0
            Case "5", "05", "MAY"
                If mDay > 31 Then mDay = 0
            Case "6", "06", "JU", "JUN"
                If mDay > 30 Then mDay = 0
            Case "7", "07", "JUL"
                If mDay > 31 Then mDay = 0
            Case "8", "08", "AU", "AUG"
                If mDay > 31 Then mDay = 0
            Case "9", "09", "S", "SE", "SEP"
                If mDay > 30 Then mDay = 0
            Case "10", "O", "OC", "OCT"
                If mDay > 31 Then mDay = 0
            Case "11", "N", "NO", "NOV"
                If mDay > 30 Then mDay = 0
            Case "12", "D", "DE", "DEC"
                If mDay > 31 Then mDay = 0
            Case Else
                mDay = 0
        End Select
        If mDay = 0 Then mDay = Day(Now)
        RetDate = Format(Trim(STR(mDay)), "00") + "/" + Trim(mMonth) + "/" + Trim(STR(mYear))
        Exit Function
err1:
    ' For Overflow Check
    If err.NUMBER = 6 Then Resume Next
End Function


Public Function dmRoundOff(Amt As Double, Optional noPlace As Byte, Optional rType As String) As Double
'rTYPE      - S-Standard  U-Upper Value  L-Lower Value
'noPlace    - 0- >Rupees   1- >10 Paise    2- >25 Paise   3- >50 Paise, 4- >No RoundOff
Dim RoundType As String, DecimalPlace As Byte
RoundType = IIf(IsMissing(rType) Or rType = "", PubRoundOffType, rType)
DecimalPlace = IIf(IsMissing(noPlace), PubRoundOffPosition, noPlace)

If DecimalPlace = 4 Then dmRoundOff = Amt: Exit Function
'Public Function dmRoundOff(Amt As Double, afterRoundAmt As Double, roundOffAmt As Double, rTYPE As String, noPlace As Integer) As Double

Dim xAmt As Double
xAmt = Format(Amt - Int(Amt), "0.00")
Select Case DecimalPlace
    Case 0  'Rupee
        dmRoundOff = IIf(RoundType = "S", Format(Amt, "0"), IIf(RoundType = "L", Int(Amt), IIf(CDbl(Amt - Round(Amt, 0)) <= 0, Round(Amt, 0), Int(Amt) + 1)))
    Case 1  '10 Paise
        Select Case RoundType
            Case "S"
                dmRoundOff = Round(Amt, 1)
            Case "L"
                dmRoundOff = Int(Amt) + IIf(xAmt >= 0# And xAmt <= 0.09, 0, IIf(xAmt >= 0.1 And xAmt <= 0.19, 0.1, IIf(xAmt >= 0.2 And xAmt <= 0.29, 0.2, IIf(xAmt >= 0.3 And xAmt <= 0.39, 0.3, IIf(xAmt >= 0.4 And xAmt <= 0.49, 0.4, IIf(xAmt >= 0.5 And xAmt <= 0.59, 0.5, IIf(xAmt >= 0.6 And xAmt <= 0.69, 0.6, IIf(xAmt >= 0.7 And xAmt <= 0.79, 0.7, IIf(xAmt > 0.8 And xAmt <= 0.89, 0.8, IIf(xAmt > 0.9 And xAmt <= 0.99, 0.9, 1))))))))))
            Case "U"
                dmRoundOff = Int(Amt) + IIf(xAmt = 0, 0, IIf(xAmt >= 0.01 And xAmt <= 0.1, 0.1, IIf(xAmt >= 0.11 And xAmt <= 0.2, 0.2, IIf(xAmt >= 0.21 And xAmt <= 0.3, 0.3, IIf(xAmt >= 0.31 And xAmt <= 0.4, 0.4, IIf(xAmt >= 0.41 And xAmt <= 0.5, 0.5, IIf(xAmt >= 0.51 And xAmt <= 0.6, 0.6, IIf(xAmt >= 0.61 And xAmt <= 0.7, 0.7, IIf(xAmt >= 0.71 And xAmt <= 0.8, 0.8, IIf(xAmt > 0.81 And xAmt <= 0.9, 0.9, 1))))))))))
        End Select
    Case 2  '25 Paise
        Select Case RoundType
            Case "S"
                dmRoundOff = Int(Amt) + IIf(xAmt <= 0.12, 0, IIf(xAmt > 0.12 And xAmt <= 0.37, 0.25, IIf(xAmt > 0.37 And xAmt <= 0.62, 0.5, IIf(xAmt > 0.62 And xAmt <= 0.87, 0.75, 1))))
            Case "L"
                dmRoundOff = Int(Amt) + IIf(xAmt <= 0.24, 0, IIf(xAmt > 0.24 And xAmt <= 0.49, 0.25, IIf(xAmt > 0.49 And xAmt <= 0.74, 0.5, 0.75)))
            Case "U"
                dmRoundOff = Int(Amt) + IIf(xAmt = 0, 0, IIf(xAmt <= 0.25, 0.25, IIf(xAmt > 0.25 And xAmt <= 0.5, 0.5, IIf(xAmt > 0.5 And xAmt <= 0.75, 0.75, 1))))
        End Select
    Case 3  '50 Paise
        Select Case RoundType
            Case "S"
                dmRoundOff = Int(Amt) + IIf(xAmt <= 0.24, 0, IIf(xAmt >= 0.25 And xAmt <= 0.74, 0.5, 1))
            Case "L"
                dmRoundOff = Int(Amt) + IIf(xAmt <= 0.49, 0, 0.5)
            Case "U"
                dmRoundOff = Int(Amt) + IIf(xAmt = 0, 0, IIf(xAmt <= 0.5, 0.5, 1))
        End Select
End Select
dmRoundOff = Format(CDbl(dmRoundOff - Amt), "0.00")
End Function

Public Sub FGrid_Delete(ByRef FrmName As Form)
    If FrmName.FGrid1.Rows >= 1 Then
        If MsgBox("Are you sure to Delete this ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            If FrmName.FGrid1.Rows = 2 Then
                FrmName.FGrid1.Rows = 1
                FrmName.FGrid1.AddItem ""
            Else
                FrmName.FGrid1.RemoveItem (FrmName.FGrid1.Row)
            End If
            FrmName.FGrid1.Col = 1
            FrmName.FGrid1.Sort = 0
        End If
    Else
        MsgBox "No Entries To Delete", vbCritical
    End If
End Sub

Public Function XNull(Temp As Variant) As String
    XNull = Replace(IIf(IsNull(Temp), "", Temp), "'", "")
End Function
Public Function VNull(Temp As Variant) As Variant
    VNull = IIf(IsNull(Temp) Or Temp = "", 0, Temp)
End Function

Public Function eVal(Temp As Variant) As Variant
    eVal = IIf(IsNull(Temp) Or Temp = "", 0, Round(Val(Replace(Replace(UCase(XNull(Temp)), "RS.", ""), ",", "")), 2))
End Function

Public Function eValSql(mFieldName As Variant) As String
    eValSql = "Val(Replace(Replace(UCase(mFieldName), 'RS.', ''), ',', ''))"
End Function

Public Function Validate_Numeric(Temp As Variant) As Double
    Validate_Numeric = IIf(Trim(Temp) = "" Or IsNumeric(Temp) = False, 0, Val(Temp))
End Function

Public Sub NumDown(ByRef TEXT As Object, KeyCode As Integer, LeftPlace As Integer, RightPlace As Integer)
    If KeyCode = 46 Then
        If InStr(TEXT, "-") <> 0 And mID(TEXT, TEXT.SelStart + 1, 1) = "." And Len(TEXT) - 1 - RightPlace >= LeftPlace Then
            KeyCode = 0
        ElseIf InStr(TEXT, "-") = 0 And mID(TEXT, TEXT.SelStart + 1, 1) = "." And Len(TEXT) - RightPlace >= LeftPlace Then
            KeyCode = 0
        End If
    End If
End Sub

Public Sub NumPress(ByRef TEXT As Object, KeyAscii As Integer, LeftPlace As Integer, RightPlace As Integer, Optional AllowNegetive As Boolean = False)
Dim myString    As String

If AllowNegetive Then
    myString = "0123456789.-"
Else
    myString = "0123456789."
End If

If KeyAscii > 26 Then
   If InStr(myString, Chr(KeyAscii)) = 0 Then KeyAscii = 0
   If (InStr(TEXT, "-") <> 0) And KeyAscii = 45 Then KeyAscii = 0
   If InStr(TEXT, ".") <> 0 Then
        If KeyAscii = vbKeyDelete Then KeyAscii = 0
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
        If KeyAscii = vbKeyDelete Then Exit Sub
        If InStr(TEXT, "-") <> 0 Then
            If Len(TEXT) - 1 >= LeftPlace Then KeyAscii = 0
        Else
            If Len(TEXT) >= LeftPlace And KeyAscii <> 45 Then KeyAscii = 0
        End If
   End If
'ElseIf KeyAscii = vbkeyback And InStr(Text, "-") <> 0 And Mid(Text, Text.SelStart, 1) = "." And Mid(Text, Text.SelStart + 1, 1) <> "" And Len(Text) - 1 - RightPlace  >= LeftPlace Then
'    KeyAscii = 0
'ElseIf KeyAscii = vbkeyback And InStr(Text, "-") = 0 And Mid(Text, Text.SelStart, 1) = "." And Mid(Text, Text.SelStart + 1, 1) <> "" And Len(Text) - RightPlace  >= LeftPlace Then
'    KeyAscii = 0
'lps 19-05-2003 due on error resume next
ElseIf KeyAscii = vbKeyBack And InStr(TEXT, "-") <> 0 Then
    If mID(TEXT, TEXT.SelStart, 1) = "." And mID(TEXT, TEXT.SelStart + 1, 1) <> "" And Len(TEXT) - 1 - RightPlace >= LeftPlace Then
        KeyAscii = 0
    ElseIf mID(TEXT, TEXT.SelStart, 1) = "." And mID(TEXT, TEXT.SelStart + 1, 1) <> "" And Len(TEXT) - RightPlace >= LeftPlace Then
        KeyAscii = 0
    End If
End If

'Dim Mystring As String
'    If RightPlace = 0 Then Mystring = "0123456789" & Text.Tag Else Mystring = "0123456789." & Text.Tag
'    If KeyAscii  > 26 Then
'       If InStr(Mystring, Chr(KeyAscii)) = 0 Then KeyAscii = 0
'       If (InStr(Text, "-") <> 0) And KeyAscii = 45 Then KeyAscii = 0 'Restrict two - symbol
'       If InStr(Text, ".") <> 0 Then
'            If KeyAscii = vbkeydelete Then KeyAscii = 0
'            If InStr(Text, "-") <> 0 Then
'                If InStr(Text, ".") - 1  > LeftPlace And Text.SelStart < InStr(Text, ".") Then
'                    KeyAscii = 0
'                ElseIf Len(Text)  >= InStr(Text, ".") + RightPlace And Text.SelStart  >= InStr(Text, ".") Then
'                    KeyAscii = 0
'                End If
'            Else
'                If InStr(Text, ".")  > LeftPlace And Text.SelStart < InStr(Text, ".") Then
'                    KeyAscii = 0
'                ElseIf Len(Text)  >= InStr(Text, ".") + RightPlace And Text.SelStart  >= InStr(Text, ".") Then
'                    KeyAscii = 0
'                End If
'            End If
'       Else
'            If KeyAscii = vbkeydelete Then Exit Sub
'            If InStr(Text, "-") <> 0 Then
'                If Len(Text) - 1  >= LeftPlace Then KeyAscii = 0
'            Else
'                If Len(Text)  >= LeftPlace And KeyAscii <> 45 Then KeyAscii = 0
'            End If
'       End If
'    ElseIf KeyAscii = vbkeyback And InStr(Text, "-") <> 0 And Mid(Text, Text.SelStart, 1) = "." And Mid(Text, Text.SelStart + 1, 1) <> "" And Len(Text) - 1 - RightPlace  >= LeftPlace Then
'        KeyAscii = 0
'    ElseIf KeyAscii = vbkeyback And InStr(Text, "-") = 0 And Mid(Text, Text.SelStart, 1) = "." And Mid(Text, Text.SelStart + 1, 1) <> "" And Len(Text) - RightPlace  >= LeftPlace Then
'        KeyAscii = 0
'    End If
End Sub
Public Function ntow(ByVal NN As Double, mmajor As String, mminor As String) As String
Dim mDecimals As Long, nums As Long, mIntegers As Long, I As Integer, cn As Long
Dim X As Variant
mDecimals = 0
nums = 0
mIntegers = 0
I = 0
cn = 0
    Unit = "One  Two  ThreeFour Five Six  SevenEightNine " '5
    tens = "Eleven   Twelve   Therteen Fourteen Fifteen  Sixteen  SeventeenEighteen Nineteen " '9
    tenth = "Ten    Twenty Thirty Fourty Fifty  Sixty  SeventyEighty Ninty  " '7
    PLACE = "Hundred ThousandLacs    Crore   " '8
    cn = 10000000
    WORDs = mmajor
    
    mIntegers = Int(NN)
    mDecimals = ((NN) - Int(NN)) * 100
    For I = 1 To 5
        nums = Int(mIntegers / cn)
        mIntegers = mIntegers - (nums * cn)
        If I <> 3 Then cn = cn / 100 Else cn = cn / 10
        If nums > 0 Then X = CONVERTs(nums, I)
    Next I
    If mDecimals > 0 Then
'altered by shekhar pandey 10/1/2003
        WORDs = WORDs + " " + mminor
        X = CONVERTs(mDecimals, I)
'earlier coding
'        WORDs = WORDs + mminor + " "
    End If
    WORDs = WORDs + " Only"
    ntow = WORDs
End Function

Private Function CONVERTs(ByRef NUM As Long, ByRef Index As Integer) As Boolean
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

Public Function Amount_Fill(Amt As Double, AmtPrefix As String) As String
    Amount_Fill = Replace(Space(12 - Len(Format(Amt, "0.00"))), " ", AmtPrefix) & Format(Amt, "0.00")
End Function

Public Sub Report_View(mREPORT As CRAXDRT.Report, CAPTION As String, Optional Index As Integer, Optional FirstPrint As Boolean)
'Eject_No property withdrawn 07-03-02
 Dim mFILE_NAME As String * 16, connectionId
 Dim PageWidth As Byte
Dim fob As New Scripting.FileSystemObject, mSno As Byte, varTxtstrm As Scripting.TextStream
Dim mReportCount As Integer
Dim rpt_form As New REPVIEW
On Error Resume Next
If Index = 1 Then
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    PageWidth = 80
    mREPORT.ExportOptions.DiskFileName = "C:\RepPrint.Txt"
    mREPORT.ExportOptions.FormatType = 10   'Enables printer to eject page
'        mREPORT.ExportOptions.FormatType = 8   'Disables Printer to Page Eject
    mREPORT.ExportOptions.NumberOfLinesPerPage = PubPageLength
    mREPORT.ExportOptions.DestinationType = 1
    For mReportCount = 1 To mREPORT.FormulaFields.Count
        Select Case UCase(mREPORT.FormulaFields(mReportCount).FormulaFieldName)
            Case UCase("comp_name")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PRN_TIT(PubComp_Name, "A", PageWidth) & "'"
            Case UCase("comp_add1")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PRN_TIT(PubComp_Add, "C", PageWidth) & "'"
            Case UCase("comp_add2")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PRN_TIT(PubComp_Add2, "C", PageWidth) & "'"
            Case UCase("comp_city")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PRN_TIT(PubComp_Add2 + IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth) & "'"
            Case UCase("Title")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PRN_TIT(CAPTION, "A", PageWidth) & "'"
            Case UCase("SpeedPrint")
                mREPORT.FormulaFields(mReportCount).TEXT = "'1'"
        End Select
    Next
    Call mREPORT.Export(False)
    Set varTxtstrm = fob.OpenTextFile("C:\RepPrint.Txt", ForAppending)
    varTxtstrm.Write (Chr(12))
    varTxtstrm.Close
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
Else
    If FirstPrint = True Then
        rpt_form.CRViewer1.EnablePrintButton = True
'            rpt_form.CRViewer1.EnableExportButton = False
        rpt_form.CRViewer1.EnableRefreshButton = True
    Else
        rpt_form.CRViewer1.EnablePrintButton = True
        rpt_form.CRViewer1.EnableExportButton = True
        rpt_form.CRViewer1.EnableRefreshButton = True
    End If
    rpt_form.CRViewer1.EnableAnimationCtrl = True
    rpt_form.CRViewer1.EnableNavigationControls = True
    rpt_form.CRViewer1.EnablePopupMenu = True
    Dim I As Integer
    For I = 1 To mREPORT.FormulaFields.Count
        Select Case UCase(mREPORT.FormulaFields(I).FormulaFieldName)
            Case UCase("comp_name")
                mREPORT.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
            Case UCase("comp_add1")
                mREPORT.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
            Case UCase("comp_add2")
                mREPORT.FormulaFields(I).TEXT = "'" & PubComp_Add2 & "'"
            Case UCase("comp_city")
                mREPORT.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
            Case UCase("Title")
                mREPORT.FormulaFields(I).TEXT = "'" & CAPTION & "'"
        End Select
    Next
    rpt_form.Rep_Set = mREPORT
    rpt_form.CAPTION = "* " & CAPTION & " {" & PrinterDetail & "}" & " *"
End If
Set rpt_form = Nothing
Set mREPORT = Nothing
Set fob = Nothing
Exit Sub
ERRORHANDLER:  MsgBox err.Description, vbCritical
End Sub

''Public Function Killer(mDate As Date) As Boolean
''Killer = False
''If DateDiff("D", Format("10/OCT/2001", "dd/mmm/yyyy"), Format(mDate, "dd/mmm/yyyy")) >= 0 Then
''    G_CompCn.Execute "Update Company Set CentralData_Path =  " & cTrim("CentralData_Path") & " + '.' "
'''    MsgBox "Contact to Dataman"
''    Killer = True
''End If
''End Function

Public Function LedgerPost(EntryMode As String, ByRef RecAry() As LedgRec, xGCnFA As ADODB.Connection, DocID As String, Optional VDate As Date, Optional CommNarr As String) As Byte

Dim I As Byte, mVSNo As Byte, mDR As Double, mCR As Double, mSId As Double
Dim RST1 As ADODB.Recordset ', rst2 As ADODB.Recordset
Dim mSiteCode$, mVType$, mVNo$, mMainGrCode$, mForSiteCode$, MyPrefix$, mSepNarYN$

Dim TempCn As ADODB.Connection

Set TempCn = New ADODB.Connection
TempCn.ConnectionString = xGCnFA.ConnectionString & ";Password=" & PubDbPass
TempCn.CursorLocation = adUseClient
TempCn.Open

'If UCase(left(PubComp_Name, 3)) = "JMK" And ConvertDate(Vdate) <= "30/Sep/2005" Then
'    LedgerPost = 1
'    Exit Function
'End If
LedgerPost = 0
mDR = 0
mCR = 0
For I = 0 To UBound(RecAry)
    Debug.Print IIf(IsNull(RecAry(I).AmtDr), 0, RecAry(I).AmtDr) & " Dr"
    Debug.Print IIf(IsNull(RecAry(I).AmtCr), 0, RecAry(I).AmtCr) & " Cr"
    mDR = mDR + IIf(IsNull(RecAry(I).AmtDr), 0, RecAry(I).AmtDr)
    mCR = mCR + IIf(IsNull(RecAry(I).AmtCr), 0, RecAry(I).AmtCr)
Next
If Format(mDR, "0.00") <> Format(mCR, "0.00") Then
    Debug.Print ""
    Debug.Print Format(mDR, "0.00")
    Debug.Print Format(mCR, "0.00")
End If
If EntryMode <> "D" And Format(mDR, "0.00") <> Format(mCR, "0.00") Then LedgerPost = 4: Exit Function

'If EntryMode = "A" Then If xGCnFA.Execute("SELECT COUNT(*) FROM LEDGERM WHERE Comp_Code=" & Chk_Text(RecAry(1).comp_code) & " AND V_Type=" & Chk_Text(RecAry(1).V_Type) & " AND v_Prefix=" & Chk_Text(RecAry(1).v_Prefix) & " AND V_No=" & RecAry(1).V_NO & " AND Site_Code=" & Chk_Text(RecAry(1).Site_Code)).Fields(0)  > 0 Then LedgerPost = 2: Exit Function
'If EntryMode <> "A" Then  'Edit or Delete
If EntryMode <> "C" Then    ' Cancel bill
    LedgerUnPost xGCnFA, DocID
End If
'End If
'For Ledger Save
If (EntryMode <> "D" And EntryMode <> "B") Then
'    mSiteCode = DeCodeDocID(DocId, Current_Site) + DeCodeDocID(DocId, For_Site_Code)
'    mVType = DeCodeDocID(DocId, Document_Type)
'    mVNo = DeCodeDocID(DocId, Document_No)
'    mForSiteCode = DeCodeDocID(DocId, For_Site_Code)
'    MyPrefix = DeCodeDocID(DocId, Document_Prefix)
    
    mSiteCode = mID(DocID, 2, 2)
    mVType = Trim(mID(DocID, 4, 5))
    mVNo = Right(DocID, 8)
    mForSiteCode = mID(DocID, 3, 1)
    MyPrefix = Trim(mID(DocID, 9, 5))
    If EntryMode = "C" Then
        'VDate = PubLoginDate
    End If
    'For Common / Separate Narration
    If PubBackEnd = "A" Then
        Set RST1 = xGCnFA.Execute("Select Separate_Narr,Common_Narr from Voucher_Type where V_Type='" & mVType & "'")
    Else
        Set RST1 = xGCnFA.Execute("Select Separate_Narr,Common_Narr from Voucher_Type WITH (NOLOCK) where V_Type='" & mVType & "'")
    End If
    If RST1!Common_Narr <> "Y" Then
        CommNarr = ""
    End If
    mSepNarYN = RST1!Separate_Narr
    Set RST1 = Nothing
    'Separate_Narr
    'Common_Narr
    mVSNo = 0
        GSQL = "insert into ledgerm(" _
        & "DocId,Site_Code,V_Type,v_prefix,V_No," _
        & "V_Date,Narration,U_Name, U_EntDt, U_AE) " _
        & " values('" & DocID & "','" & PubSiteCode & mForSiteCode & "','" & mVType & "','" & MyPrefix & "'," & mVNo & "," _
        & "" & ConvertDate(VDate) & ",'','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(EntryMode, 1) & "')"
    xGCnFA.Execute GSQL
    For I = 0 To UBound(RecAry)
        If RecAry(I).AmtCr + RecAry(I).AmtDr <> 0 Then
            mVSNo = mVSNo + 1
            GSQL = "insert into ledger(" _
                & "DocId,Site_Code,V_SNo,V_Type, V_Prefix,V_No," _
                & "V_Date,SubCode,ContraSub, " _
                & "AmtDr,AmtCr,Narration," _
                & "Chq_No, Chq_Date, Clg_Date, EmpDetailYn, U_Name, U_EntDt, U_AE)" _
                & " values(" _
                & "'" & DocID & "','" & mSiteCode & "'," & mVSNo & ",'" & mVType & "', '" & MyPrefix & "'," & mVNo & _
                "," & ConvertDate(VDate) & ",'" & RecAry(I).SubCode & "', '" & RecAry(I).ContraSub & _
                "'," & RecAry(I).AmtDr & "," & RecAry(I).AmtCr & ",'" & SETW(RecAry(I).Narration, 250) & _
                "', '" & RecAry(I).Chq_No & "', " & ConvertDate(RecAry(I).Chq_Date) & ", " & ConvertDate(RecAry(I).Clg_Date) & ", '" & RecAry(I).EmpDetailYn & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(EntryMode, 1) & "')"
            xGCnFA.Execute GSQL
            'LedgerRef posting
            If PubBackEnd = "S" Then

                Set GCnTemp = New ADODB.Connection
                If PubDbUser <> "" Then
                    GCnTemp.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                Else
                    GCnTemp.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                End If
                
                'GCnTemp.ConnectionString = GCn.ConnectionString
                GCnTemp.CursorLocation = adUseClient
                GCnTemp.Open
                GCnTemp.BeginTrans
                    mSId = IIf(IsNull(GCnTemp.Execute("Select Max(ID) from LedgerRef").Fields(0).Value), 0, G_FaCn.Execute("select Max(ID) from LedgerRef").Fields(0).Value) + 1
                    GSQL = "INSERT INTO LedgerRef (ID,DOCID,V_SNO,DR,CR,SUBCODE,U_Name,U_EntDt,U_AE,AgRefType,AgRefNo,DueDate,V_dATE)" & _
                            "VALUES (" & mSId & ",'" & DocID & "'," & mVSNo & "," & RecAry(I).AmtDr & "," & RecAry(I).AmtCr & ",'" & RecAry(I).SubCode & "','" & pubUName & "'," & FaConvertDate(Now) & ",'E','New Ref','" & DocID & "'," & FaConvertDate(Now) & "," & ConvertDate(VDate) & ")"
                    GCnTemp.Execute GSQL
                GCnTemp.CommitTrans
                GCnTemp.Close
            Else
                mSId = IIf(IsNull(G_FaCn.Execute("Select Max(ID) from LedgerRef").Fields(0).Value), 0, G_FaCn.Execute("select Max(ID) from LedgerRef").Fields(0).Value) + 1
                GSQL = "INSERT INTO LedgerRef (ID,DOCID,V_SNO,DR,CR,SUBCODE,U_Name,U_EntDt,U_AE,AgRefType,AgRefNo,DueDate,V_dATE)" & _
                        "VALUES (" & mSId & ",'" & DocID & "'," & mVSNo & "," & RecAry(I).AmtDr & "," & RecAry(I).AmtCr & ",'" & RecAry(I).SubCode & "','" & pubUName & "'," & FaConvertDate(Now) & ",'E','New Ref','" & DocID & "'," & FaConvertDate(Now) & "," & ConvertDate(VDate) & ")"
                G_FaCn.Execute GSQL
            End If
            '******************
            If Not PubImportData Then
                If PubBackEnd = "A" Then
                    mMainGrCode = TempCn.Execute("Select Acgroup.maingrcode from subgroup left join acgroup on subgroup.groupcode=acgroup.groupcode where acgroup.aliasyn='N' and  SubGroup.SubCode='" & RecAry(I).SubCode & "'").Fields(0).Value
                Else
                    mMainGrCode = TempCn.Execute("Select Acgroup.maingrcode from subgroup With (NOLOCK) left join acgroup on subgroup.groupcode=acgroup.groupcode where acgroup.aliasyn='N' and  SubGroup.SubCode='" & RecAry(I).SubCode & "'").Fields(0).Value
                End If
                If RecAry(I).AmtCr <> 0 Then
                    If PubBackEnd = "A" Then
                        GCn.Execute ("Update SubGroup set curr_bal=Curr_bal+" & RecAry(I).AmtCr & " where SubCode='" & RecAry(I).SubCode & "'")
                    End If
    
                    TempCn.Execute ("Update SubGroup set curr_bal=Curr_bal+" & RecAry(I).AmtCr & " where SubCode='" & RecAry(I).SubCode & "'")
                    CalBalAcGroup "SubGroup", TempCn, mMainGrCode, RecAry(I).AmtCr, "+"
                ElseIf RecAry(I).AmtDr <> 0 Then
                    If PubBackEnd = "A" Then
                        GCn.Execute ("Update SubGroup set curr_bal=Curr_bal-" & RecAry(I).AmtDr & " where SubCode='" & RecAry(I).SubCode & "'")
                    End If
                    TempCn.Execute ("Update SubGroup set curr_bal=Curr_bal-" & RecAry(I).AmtDr & " where SubCode='" & RecAry(I).SubCode & "'")
                    CalBalAcGroup "SubGroup", TempCn, mMainGrCode, RecAry(I).AmtDr, "-"
                End If
            End If
        End If
    Next
End If
LedgerPost = 1
Exit Function

End Function

Public Sub LedgerUnPost(xGCnFA As ADODB.Connection, DocID As String)
'Updating balance for FAData only
Dim TempCn As ADODB.Connection

Set TempCn = New ADODB.Connection
TempCn.ConnectionString = xGCnFA.ConnectionString & ";Password=" & PubDbPass
TempCn.CursorLocation = adUseClient
TempCn.Open


Set GRs = New Recordset
If Not PubImportData Then
    If PubBackEnd <> "A" Then
        Set GRs = TempCn.Execute("Select L.SubCode,L.AmtDr,L.AmtCr,SubGroup.GroupCode,AcGroup.MainGrCode " & _
            "from (Ledger  as L WITH (NOLOCK) left join SubGroup WITH (NOLOCK) on L.SubCode=SubGroup.SubCode) " & _
            "left join AcGroup WITH (NOLOCK) on SubGroup.GroupCode=AcGroup.GroupCode  where docid='" & DocID & "'  and AcGroup.AliasYN='N'")
    Else
        Set GRs = TempCn.Execute("Select L.SubCode,L.AmtDr,L.AmtCr,SubGroup.GroupCode,AcGroup.MainGrCode " & _
            "from (Ledger  as L left join SubGroup on L.SubCode=SubGroup.SubCode) " & _
            "left join AcGroup on SubGroup.GroupCode=AcGroup.GroupCode  where docid='" & DocID & "'  and AcGroup.AliasYN='N'")
    End If
        If GRs.RecordCount > 0 Then
            While Not GRs.EOF
                If GRs!AmtCr > 0 Then
                    CalBalAcGroup "SubGroup", TempCn, GRs!MainGrCode, GRs!AmtCr, "-"
                    If PubBackEnd = "A" Then
                        TempCn.Execute ("Update SubGroup  set curr_bal=Curr_bal-" & GRs!AmtCr & " where SubCode='" & GRs!SubCode & "'")
                        GCn.Execute ("Update SubGroup  set curr_bal=Curr_bal-" & GRs!AmtCr & " where SubCode='" & GRs!SubCode & "'")
                    Else
                        TempCn.Execute ("Update SubGroup WITH (ROWLOCK) set curr_bal=Curr_bal-" & GRs!AmtCr & " where SubCode='" & GRs!SubCode & "'")
                    End If
                Else
                    CalBalAcGroup "SubGroup", TempCn, GRs!MainGrCode, GRs!AmtDr, "+"
                    If PubBackEnd = "A" Then
                        TempCn.Execute ("Update SubGroup set curr_bal=Curr_bal+" & GRs!AmtDr & " where SubCode='" & GRs!SubCode & "'")
                        GCn.Execute ("Update SubGroup set curr_bal=Curr_bal+" & GRs!AmtDr & " where SubCode='" & GRs!SubCode & "'")
                    Else
                        TempCn.Execute ("Update SubGroup set curr_bal=Curr_bal+" & GRs!AmtDr & " where SubCode='" & GRs!SubCode & "'")
                    End If
                End If
                GRs.MoveNext
            Wend
        End If
    Set GRs = Nothing
End If
TempCn.Execute ("Delete from LedgerAdj where DocId2='" & DocID & "'")
TempCn.Execute ("Delete from LedgerAdj where DocId1='" & DocID & "'")
TempCn.Execute ("Delete from Ledger where DocId='" & DocID & "'")
TempCn.Execute ("Delete from LedgerM where DocId='" & DocID & "'")

Set TempCn = Nothing
End Sub


'This Function is used to Maintain Current Balance of Group
Public Sub CalBalAcGroup(TableType As String, CC As ADODB.Connection, MainGrCode As String, Amt As Double, PlusMinus As String)
Dim ControlStr As String, I As Integer, Length As Integer
    ControlStr = ""
    If UCase(TableType) = UCase("AcGroup") Then
        Length = Len(MainGrCode) - 3
    ElseIf UCase(TableType) = UCase("SubGroup") Then
        Length = Len(MainGrCode)
    End If
    For I = Length To 3 Step -3
        If ControlStr = "" Then
            ControlStr = "'" & left(MainGrCode, I) & "'"
        Else
            ControlStr = ControlStr & ",'" & left(MainGrCode, I) & "'"
        End If
    Next
    If ControlStr <> "" Then
        CC.Execute ("Update AcGroup Set CurrentBalance=CurrentBalance " & PlusMinus & " " & Amt & " Where MainGrCode In(" & ControlStr & ")")
    End If
End Sub

'''This Function is used to Maintain Current Balance of SubGroup
'Public Sub CalBalanceSubGroup(CC As ADODB.Connection, SubCode As String, Amt As Double, PlusMinus As String, Optional NotUpdAcGrpBal As Boolean)
'Dim ControlStr$, i As Integer, Length As Integer, MainGrCode$
'CC.Execute ("Update SubGroup Set Curr_Bal=Curr_Bal " & PlusMinus & " " & Amt & " Where SubCode='" & SubCode & "'")
'
'If NotUpdAcGrpBal Then
'    If CC.Execute("Select Acgroup.maingrcode from subgroup left join acgroup on subgroup.groupcode=acgroup.groupcode where acgroup.aliasyn='N'").RecordCount  > 0 Then
'        MainGrCode = CC.Execute("Select Acgroup.maingrcode from subgroup left join acgroup on subgroup.groupcode=acgroup.groupcode where acgroup.aliasyn='N'").Fields(0).Value
'    End If
'End If
'If Len(MainGrCode)  > 0 Then
'    For i = Len(MainGrCode) To 3 Step -3
'        If ControlStr = "" Then
'            ControlStr = "'" & left(MainGrCode, i) & "'"
'        Else
'            ControlStr = ControlStr & ",'" & left(MainGrCode, i) & "'"
'        End If
'    Next
'    If ControlStr <> "" Then
'        CC.Execute ("Update AcGroup Set CurrentBalance=CurrentBalance " & PlusMinus & " " & Amt & " Where MainGrCode In(" & ControlStr & ")")
'    End If
'End If
'End Sub

''*****
'Public Sub Hook(hwnd As Long)
''Disable right button popup menu
'    lngHWnd = hwnd
'    lpPrevWndProc = SetWindowLong(lngHWnd, GWL_WNDPROC, AddressOf WindowProc)
'End Sub
'
'Public Sub UnHook()
''Disable right button popup menu
'Dim lngReturnValue As Long
'lngReturnValue = SetWindowLong(lngHWnd, GWL_WNDPROC, lpPrevWndProc)
'End Sub
'
'Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
''Disable right button popup menu
'Select Case uMsg
'    Case WM_RBUTTONUP
'        'Do nothing
'        'Or popup you own menu
'    Case Else
'        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
'End Select
'End Function


'Public Function IsSelected(lst As ListBox) As Boolean
'Dim i As Integer
'  For i = 0 To lst.ListCount - 1
'    If lst.Selected(i) = True Then
'      IsSelected = True
'       Exit Function
'    End If
'  Next
'  IsSelected = False
'End Function


'Public Sub REPORT_VIEW(mREPORT As Variant, CAPTION As String)
'Dim report_form As New RepView
'        mREPORT.FormulaFields(1).text = "'" & Pub_Comp_Name & "'"
'        mREPORT.FormulaFields(2).text = "'" & Pub_Comp_Add & "'"
'        mREPORT.FormulaFields(3).text = "'" & Pub_Comp_City & "'"
'        mREPORT.FormulaFields(4).text = "'" & CAPTION & "'"
''        report_form.CRViewer1.ReportSource = mREPORT
''        report_form.CRViewer1.ViewReport
'        report_form.Rep_Set = mREPORT
'        report_form.CAPTION = "* " + CAPTION + " *"
'    Set report_form = Nothing
'Set mREPORT = Nothing
'Exit Sub
'ERRORHANDLER:  MsgBox Err.Description, vbCritical
'End Sub


''****** to recieve value in dtpicker
'Public Sub DtpVal(ByVal DtpName As Object, ByVal FldName As Variant)
'If IsNull(FldName) Then
'    If DtpName.CheckBox = True Then
'        DtpName.Value = Null    'False
'    End If
'Else
'    DtpName.Value = FldName
'End If
'End Sub

'
'Public Function RemoveQuot(temp As Variant) As Variant
'Dim Mypos As Integer
'RemoveQuot = temp
'If IsNull(RemoveQuot) Or RemoveQuot = Null Then
'    RemoveQuot = "Null"
'    Exit Function
'End If
'Mypos = InStr(1, RemoveQuot, "'", 1)
'Do While Mypos <> 0
'    If Mypos <= 0 Then
'        RemoveQuot = temp
'    Else
'        RemoveQuot = (Left(RemoveQuot, Mypos) & "'" & Right(RemoveQuot, Len(RemoveQuot) - Mypos))
'    End If
'    Mypos = InStr(1 + Mypos, temp, "'", 1)
'Loop
'RemoveQuot = "'" & RemoveQuot & "'"
'End Function

'Public Function Code_Check(ByVal str As String) As String
'Dim temp As String
'Dim i As Integer
'temp = ""
'    For i = 1 To Len(str)
'        temp = Mid(str, i, 1)
'        If temp = "'" Then temp = ""
'        Code_Check = Code_Check & temp
'    Next
'End Function

'Public Sub Single_Quote(frm As Form)
'Dim objctrl As Control
'For Each objctrl In frm.Controls
'If TypeOf objctrl Is TextBox Then
'    objctrl = Code_Check(objctrl)
'End If
'Next objctrl
'End Sub

Public Sub CheckQuote(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
End Sub

Public Function FormChk(CapStr As String) As Boolean
Dim Z As Integer
For Z = 0 To Forms.Count - 1
    If UCase(Forms(Z).CAPTION) = UCase(CapStr) Then
      Forms(Z).ZOrder 0
      FormChk = True
      Exit Function
    End If
Next Z
FormChk = False
End Function

'Public Sub Form_Chk(frmName As Form)
'        If frmName.Visible = False Then
'            frmName.Show
'        Else
''            frmName.WindowState = 2
'            frmName.ZOrder 0
'        End If
'        Set frmName = Nothing
'End Sub
'Public Function MfgYearMonth(ChassisNo As String, FGrid As MSHFlexGrid, MfgMonth As Integer, MfgYr As Integer)
'Public Function MfgYearMonth(ChassisNo As String, ByRef MfgMonth As Variant, ByRef MfgYr As Variant)

Public Function DeCodeChassis(ChassisNo As String, Decode As ObjTypeDefChas) As Variant
Dim Mth$, yr$
If Len(ChassisNo) < 9 Then
    MsgBox "Incomplete Chassis No.", vbCritical, "Wrong Chassis No."
    DeCodeChassis = False
End If
    
    If Len(ChassisNo) >= 17 Then
        Select Case Decode
            Case 6
                DeCodeChassis = left(ChassisNo, 6)
            Case 2
                Mth = mID(ChassisNo, 12, 1)
                If GCn.Execute("Select Name from Chas_Mth where MONTH_CD='" & Mth & "'").RecordCount > 0 Then
                    DeCodeChassis = GCn.Execute("Select Name from Chas_Mth where MONTH_CD='" & Mth & "'").Fields(0).Value
                End If
            Case 3
                yr = mID(ChassisNo, 10, 1)
                Select Case yr
                    Case "9"
                        DeCodeChassis = "2009"
                    Case "0"
                        DeCodeChassis = "2010"
                    Case "B"
                        DeCodeChassis = "2011"
                    Case "C"
                        DeCodeChassis = "2012"
                    Case "D"
                        DeCodeChassis = "2013"
                    Case "E"
                        DeCodeChassis = "2014"
                    Case "F"
                        DeCodeChassis = "2015"
                    Case "G"
                        DeCodeChassis = "2016"
                    Case "H"
                        DeCodeChassis = "2017"
                    Case "I"
                        DeCodeChassis = "2018"
                End Select
            Case 4
                DeCodeChassis = Right(ChassisNo, 6)
        End Select
    Else
        Select Case Decode
            Case 6
                DeCodeChassis = left(ChassisNo, 6)
            Case 2
                Mth = mID(ChassisNo, 7, 1)
                If GCn.Execute("Select Name from Chas_Mth where MONTH_CD='" & Mth & "'").RecordCount > 0 Then
                    DeCodeChassis = GCn.Execute("Select Name from Chas_Mth where MONTH_CD='" & Mth & "'").Fields(0).Value
                End If
            Case 3
                yr = mID(ChassisNo, 7, 2)
                If GCn.Execute("Select Name from Chas_Yr where YEAR_CD='" & yr & "'").RecordCount > 0 Then
                    DeCodeChassis = GCn.Execute("Select Name from Chas_Yr where YEAR_CD='" & yr & "'").Fields(0).Value
                End If
            Case 4
                DeCodeChassis = Right(ChassisNo, 6)
        End Select
    End If
End Function

Public Sub CheckError()
If err.NUMBER <> 0 Then
    If err.NUMBER = ErrNoVouNumManu Then
        MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
    ElseIf err.NUMBER = ErrNoVouNumAuto Then
        MsgBox "Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
    Else
        'MsgBox "Message No." & err.NUMBER & vbCrLf & err.Description, vbInformation, "Validation"
        MsgBox err.Description
    End If
    
End If
End Sub

Public Sub GridDblClick(myForm As Form, FGrid As Object, Txt As Object, Index As Integer)
On Error GoTo err
Dim I As Integer
If myForm.TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.FocusRect = flexFocusNone Then
    'FGrid.CellBackColor = CellBackColLeave
Else
    FGrid.CellBackColor = CellBackColLeave
End If
Dim g_Row As Integer
Dim g_Col As Integer
g_Row = FGrid.Row
g_Col = FGrid.Col
    Txt(Index).height = FGrid.CellHeight - 10
    Txt(Index).width = FGrid.CellWidth - 10
    Txt(Index).left = FGrid.CellLeft + FGrid.left
    Txt(Index).top = FGrid.CellTop + FGrid.top
    Txt(Index).TEXT = FGrid.TextMatrix(g_Row, g_Col)
    Txt(Index).Visible = True
    Txt(Index).ZOrder 0
    Txt(Index).Tag = FGrid.TextMatrix(g_Row, g_Col)
    Txt(Index).SetFocus
Exit Sub
err:
    CheckError
End Sub

Public Sub GridTxtDown( _
    FGrid As Object, Txt As Object, Index As Integer, _
    KeyCode As Integer, TAddMode As Boolean, ByVal MaxCol As Byte, _
    Optional SkipCol As Byte, Optional MoveToCol As Byte, Optional DisableGridSrlNo As Boolean, Optional RestrictAddingNewRow As Boolean)
Dim GCol As Byte
If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    If TAddMode = True Then
        Txt(Index).Visible = False
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
'modified by lps +1 added during AcPostCtrl EP
'        If FGrid.Col + 1 + IIf(IsMissing(SkipCol), 0, SkipCol) < MaxCol Then
        If FGrid.Col + IIf(IsMissing(SkipCol), 0, SkipCol) <= MaxCol Then
            FGrid.Col = FGrid.Col + 1 + IIf(IsMissing(SkipCol), 0, SkipCol)
            If FGrid.ColWidth(FGrid.Col) <= 0 And FGrid.Col < MaxCol Then
                FGrid.Col = FGrid.Col + 1
            End If
            FGrid.SetFocus
        Else
            If RestrictAddingNewRow = False Then
'                If FGrid.Row = FGrid.Rows - 1 And FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "" Then Exit Sub
'by lps at udaipur
                If FGrid.Row = FGrid.Rows - 1 And FGrid.TextMatrix(FGrid.Row, 1) = "" Then Exit Sub
                If IsMissing(DisableGridSrlNo) Or (Not IsMissing(DisableGridSrlNo) And DisableGridSrlNo = False) Then
'                    If FGrid.Row = FGrid.Rows - 1 Then FGrid.AddItem FGrid.Rows
                    If FGrid.Row = FGrid.Rows - 1 Then FGrid.AddItem FGrid.Rows - (IIf(FGrid.FixedRows > 1, FGrid.FixedRows - 1, 0))
                Else
                    If FGrid.Row = FGrid.Rows - 1 Then FGrid.AddItem ""
                End If
            End If
            If FGrid.Row <> FGrid.Rows - 1 Then
                FGrid.Row = FGrid.Row + 1
                For GCol = 1 To FGrid.Cols - 1
                    If FGrid.ColWidth(GCol) <> 0 Then Exit For
                Next
                FGrid.Col = GCol
            End If
            FGrid.SetFocus
            If FGrid.FocusRect = flexFocusNone Then
'               FGrid.CellBackColor = CellBackColEnter
            Else
               FGrid.CellBackColor = CellBackColEnter
            End If

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
Txt(Index).left = FGrid.left + FGrid.CellLeft
If FGrid.FocusRect = flexFocusNone Then
'   FGrid.CellBackColor = CellBackColEnter
Else
    FGrid.CellBackColor = CellBackColEnter
End If
End Sub
Public Sub Get_Text(myForm As Form, FGrid As Object, Txt As Object, Index As Integer, _
        NumericColNature As Boolean, KeyAscii As Integer)
Dim I As Integer
Dim j As Integer
If myForm.TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If NumericColNature And UCase(Chr(KeyAscii)) = "-" Then KeyAscii = 0: Exit Sub

If KeyAscii = vbKeyReturn Then
    GridDblClick myForm, FGrid, Txt, Index
'ElseIf KeyAscii = vbKeyDelete Then
'    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
ElseIf (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = vbKeySpace Then
    If FGrid.FocusRect = flexFocusNone Then
'        FGrid.CellBackColor = CellBackColLeave
    Else
        FGrid.CellBackColor = CellBackColLeave
    End If
    Dim g_Row As Integer
    Dim g_Col As Integer
        g_Row = FGrid.Row
        g_Col = FGrid.Col
        Txt(Index).height = FGrid.CellHeight - 10
        Txt(Index).width = FGrid.CellWidth - 10
        Txt(Index).left = FGrid.CellLeft + FGrid.left
        Txt(Index).top = FGrid.CellTop + FGrid.top
        Txt(Index).TEXT = ""
        Txt(Index).Visible = True
        Txt(Index).ZOrder 0
        Txt(Index).Tag = FGrid.Col
        Txt(Index).SetFocus
        If NumericColNature = True Then
            If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Then
                Txt(Index).TEXT = Chr(KeyAscii)
            End If
            GoTo NXT
        End If
        If KeyAscii = vbKeyBack Then
            Txt(Index).TEXT = ""
        Else
           Txt(Index).TEXT = Chr(KeyAscii)
        End If

NXT:
        Txt(Index).SelStart = 1
End If
Exit Sub
err:
    CheckError
End Sub

Public Sub ListView_KeyUp(LV As Object, Txt As Object, Index As Integer, KeyCode As Integer, xITEM As ListItem)
Dim STR As String
Dim LPlace As Integer
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Exit Sub
LPlace = Txt(Index).SelStart
STR = mID(Txt(Index).TEXT, 1, LPlace)
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

Public Sub ListView_KeyDown(FrmList As Object, LV As Object, Txt As Object, Index As Integer, KeyCode As Integer, Shift As Integer, left As Integer, top As Integer, width As Integer, height As Integer)
If FilterKeyCode(KeyCode) = True Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Txt(Index).TEXT <> "" Then
            Txt(Index).TEXT = LV.SelectedItem.TEXT
        End If
        FrmList.Visible = False
        Exit Sub
   End If
'    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'LPs 10-12-2K2
    If KeyCode = vbKeyEscape Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
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
                Txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.SelectedItem.Index < LV.ListItems.Count Then
                LV.ListItems(LV.SelectedItem.Index + 1).SELECTED = True
                Txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.ListItems.Count = 1 Then
                Txt(Index).TEXT = LV.SelectedItem.TEXT
            End If
        End If
    End If
End Sub

Public Function ListView_Items(LV As Object, Txt As Object, Index As Integer, list_item As Variant, Cnt As Integer) As ListItem
    Dim xName As ListItem
    Dim I As Integer
    LV.ListItems.Clear
    For I = 0 To Cnt - 1
         Set xName = LV.ListItems.Add(I + 1, , list_item(I))
    Next
    Set xName = LV.FindItem(Txt(Index), 0, , 1)
    If xName Is Nothing Then
        Exit Function
    Else
        xName.EnsureVisible
        xName.SELECTED = True
    End If
    Set ListView_Items = xName
End Function
 
Public Function ListView_Items_RecordSet(LV As Object, Txt As Object, Index As Integer, Rst As ADODB.Recordset) As ListItem
    Dim xName As ListItem
    Dim I As Long
    LV.ListItems.Clear
    If Rst.RecordCount <= 0 Then Exit Function
    Do Until Rst.EOF
        Set xName = LV.ListItems.Add(, , Rst.Fields("Name").Value)
        If Rst.Fields.Count > 2 Then
            If Not IsNull(Rst.Fields("Code").Value) Then
                xName.SubItems(1) = CStr(Rst.Fields("code").Value)
            End If
            
            'Number_Method,Common_Narr,Separate_Narr
            If Not IsNull(Rst.Fields("Number_Method").Value) Then
                xName.SubItems(2) = CStr(Rst.Fields("Number_Method").Value)
            End If
            If Not IsNull(Rst.Fields("Common_Narr").Value) Then
                xName.SubItems(3) = CStr(Rst.Fields("Common_Narr").Value)
            End If
            If Not IsNull(Rst.Fields("Separate_Narr").Value) Then
                xName.SubItems(4) = CStr(Rst.Fields("Separate_Narr").Value)
            End If
        ElseIf Rst.Fields.Count > 1 Then
            If Not IsNull(Rst.Fields("Code").Value) Then
                xName.SubItems(1) = CStr(Rst.Fields("code").Value)
            End If
        End If
        Rst.MoveNext
    Loop
    Set xName = LV.FindItem(Txt(Index), 0, , 1)
    If xName Is Nothing Then
        Exit Function
    Else
        xName.EnsureVisible
        xName.SELECTED = True
    End If
    Set ListView_Items_RecordSet = xName
End Function

Public Sub DGridColSwap(DGrid As Object, ColSwapNo As Integer)
Dim ColWidth As Integer
Dim ColCaption As String
Dim ColFieldName As String
Dim Xcol As Integer

    If Val(DGrid.Tag) <> 0 Then
        Xcol = Val(DGrid.Tag)
        ColWidth = DGrid.Columns(0).width
        ColFieldName = DGrid.Columns(0).DataField
        ColCaption = DGrid.Columns(0).CAPTION

        DGrid.Columns(0).width = DGrid.Columns(Xcol).width
        DGrid.Columns(0).CAPTION = DGrid.Columns(Xcol).CAPTION
        DGrid.Columns(0).DataField = DGrid.Columns(Xcol).DataField

        DGrid.Columns(Xcol).width = ColWidth
        DGrid.Columns(Xcol).CAPTION = ColCaption
        DGrid.Columns(Xcol).DataField = ColFieldName
        DGrid.ReBind
    End If
    ColWidth = DGrid.Columns(0).width
    ColFieldName = DGrid.Columns(0).DataField
    ColCaption = DGrid.Columns(0).CAPTION

    DGrid.Columns(0).width = DGrid.Columns(ColSwapNo).width
    DGrid.Columns(0).CAPTION = DGrid.Columns(ColSwapNo).CAPTION
    DGrid.Columns(0).DataField = DGrid.Columns(ColSwapNo).DataField

    DGrid.Columns(ColSwapNo).width = ColWidth
    DGrid.Columns(ColSwapNo).CAPTION = ColCaption
    DGrid.Columns(ColSwapNo).DataField = ColFieldName
    DGrid.ReBind
    DGrid.Tag = ColSwapNo
End Sub

Public Sub DGridTxtKeyDown(DGrid As Object, Txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, GridText As Boolean, Optional HelpIndex As Integer, Optional MasterForm, Optional MasterFormName As String)   ' FrmCity,"FrmCity"
Dim I As Integer, Menucaption As String
If FilterKeyCode(KeyCode) = True Then Exit Sub
    If KeyCode = vbKeyInsert Then
'        If Not IsMissing(MasterFormName) Then
        If MasterFormName <> "" Then
            'Menucaption = G_CompCn.Execute("select Form_Code from User_Module where ucase(name)='" & UCase(Menucaption) & "'").Fields(0).Value
            'PubUParam = MDIForm1.Permission(Menucaption)
            'If left(PubUParam, 1) <> "A" Then MsgBox "Insert Permission Denied": Exit Sub
            MasterForm.MasterFormExit = True
            MasterForm.TopCtrl1_eAdd
            KeyCode = 0
            Exit Sub
        End If
    End If

If IsMissing(HelpIndex) Then HelpIndex = 1
    If KeyCode = vbKeyReturn Then
       If Txt(Index).TEXT <> "" Then
            If Rst.BOF = False And Rst.EOF = False Then
                If GridText = True Then
                    Txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                Else
                    Txt(Index).TEXT = XNull(Rst.Fields(HelpIndex).Value)
                    Txt(Index).Tag = Rst.Fields(0).Value
                End If
            End If
        End If
        DGrid.Visible = False
        Exit Sub
    End If
    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEscape Then
        If DGrid.Visible = False Then Exit Sub
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
        Txt(Index).SelStart = Len(Txt(Index).TEXT)
        If DGrid.Visible = False Then DGrid.Visible = True:   DGrid.ZOrder 0
        If KeyCode <> vbKeyBack Then KeyCode = 0
        
    ElseIf KeyCode = vbKeyDelete Then
        Txt(Index).TEXT = ""
    Else
        If DGrid.Visible = False Then
            Txt(Index).SelStart = Len(Txt(Index).TEXT)
             DGrid.Visible = True: DGrid.ZOrder 0
        End If
'0 1 2
    End If
    If DGrid.Visible = True Then
'        If Rst.RecordCount = 0 Then Exit Sub
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            Select Case KeyCode
                Case vbKeyUp
                    If Rst.AbsolutePosition > 1 Then
                        Rst.MovePrevious
                    Else
                        KeyCode = 0
                    End If
                Case vbKeyDown
                    If Rst.AbsolutePosition < Rst.RecordCount And Rst.EOF = False Then Rst.MoveNext
'09-02-02
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
                    Txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                Else
                    Txt(Index).TEXT = XNull(Rst.Fields(HelpIndex).Value)
                    Txt(Index).Tag = Rst.Fields(0).Value
                End If
                Txt(Index).SelStart = Len(Txt(Index))
            End If
      End If
      Exit Sub
  End If
'Else
'  If DGrid.Visible = False Then DGrid.Visible = True: DGrid.ZOrder 0
'End If
End Sub


Public Sub Formula_Title(ByRef FrmName As Form, RepCaption As String)
    FrmName.CrysReport1.Formulas(0) = "CompName = '" & PubComp_Name & "'"
    FrmName.CrysReport1.Formulas(1) = "CompAdd ='" & PubComp_Add & "'"
    FrmName.CrysReport1.Formulas(2) = "CompCity  = '" & PubComp_City & "'"
    FrmName.CrysReport1.Formulas(3) = "reptitle ='" & RepCaption & "'"
End Sub

'Public Sub Formula_TitleLang(ByRef FrmName As Form, ByVal Index As Integer)
'Select Case Index
'    Case 1
'        FrmName.Report1.Formulas(0) = "compname = '" & Pub_Comp_Name & "'"
'        FrmName.Report1.Formulas(1) = "COMPADD ='" & Pub_Comp_Add & "'"
'        FrmName.Report1.Formulas(2) = "compcity  = '" & Pub_Comp_City & "'"
'    Case 2
'        FrmName.Report1.Formulas(0) = "compname = '" & H_Pub_Comp_Name & "'"
'        FrmName.Report1.Formulas(1) = "COMPADD ='" & H_Pub_Comp_Add & "'"
'        FrmName.Report1.Formulas(2) = "compcity  = '" & H_Pub_Comp_City & "'"
'End Select
'End Sub

Public Sub ProcErrorMsg()
  If err.NUMBER <> 0 Then
        MsgBox err.Description, vbCritical
  End If
End Sub

Public Sub Ctrl_validate(Ctrl As Object)
    Ctrl.BackColor = CtrlBColOrg
    Ctrl.ForeColor = CtrlFColOrg
End Sub

Public Sub Ctrl_GetFocus(Ctrl As Object)
    Ctrl.BackColor = CtrlBCol
    Ctrl.ForeColor = CtrlFCol
    Ctrl.SelStart = Len(Ctrl)
End Sub

Public Sub Ctrl_DownKeyDown(KeyCode As Integer, Shift As Integer)
If FilterKeyCode(KeyCode) = True Then Exit Sub
If KeyCode = 13 Or KeyCode = 40 Then
    SendKeysA vbKeyTab, True
    KeyCode = 0
    Shift = 0
End If
End Sub

Public Sub Ctrl_UpKeyDown(KeyCode As Integer, Shift As Integer)
If FilterKeyCode(KeyCode) = True Then Exit Sub
If KeyCode = 38 Then     'keyup = 38
    SendKeys "+{Tab}"
    KeyCode = 0
    Shift = 0
End If
End Sub

Public Sub CheckSoftware(xForm As Object)
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        MsgBox "This program is already in the air...."
        App.Title = ""
        xForm.CAPTION = ""
        AppActivate SaveTitle$
        SendKeys "%{ENTER}", True
        End
    End If

End Sub

'use as in the first form_load event
'Call CheckSoftware(Form1)
Public Function IsSelected(LST As ListBox) As Boolean
Dim I As Integer
  For I = 0 To LST.ListCount - 1
    If LST.SELECTED(I) = True Then
      IsSelected = True
       Exit Function
    End If
  Next
  IsSelected = False
End Function

Public Function FilterKeyCode(KeyCode As Integer) As Boolean
'Alter =18, WindowsStartUp = 91, CapsLock=vbKeyCapital=20, Shift =16
If (KeyCode = vbKeyControl Or KeyCode = vbKeyShift _
    Or KeyCode = vbKeyNumlock Or KeyCode = vbKeyCapital _
    Or KeyCode = vbKeyScrollLock Or KeyCode = 18 Or KeyCode = 91) Then    'And Shift = 0
    FilterKeyCode = True
    Exit Function
End If
FilterKeyCode = False

End Function

'************* This Fuction Is Used For Help Of Master Entry
Public Sub DGridTxtKeyDown_Mast(DGrid As Object, Txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, GridText As Boolean, Optional HelpIndex As Integer)
Dim I As Integer
'If Rst.RecordCount  > 0 Then
If FilterKeyCode(KeyCode) = True Then Exit Sub
If IsMissing(HelpIndex) Then HelpIndex = 1
    If KeyCode = vbKeyReturn Then
        DGrid.Visible = False
        Exit Sub
    ElseIf KeyCode = vbKeyDelete Then
        Txt(Index).TEXT = ""
        Exit Sub
    End If
    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If DGrid.Visible = False Then Exit Sub
    Else
        If DGrid.Visible = False Then DGrid.Visible = True: DGrid.ZOrder 0
    End If
    If DGrid.Visible = True Then
        If Rst.RecordCount = 0 Then Exit Sub
        If Rst.EOF = True Or Rst.BOF = True Then Rst.MoveFirst
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            Select Case KeyCode
                Case vbKeyUp
                    If Rst.AbsolutePosition > 1 Then
                        Rst.MovePrevious
                    Else
                        KeyCode = 0
                    End If
                Case vbKeyDown
                    If Rst.AbsolutePosition < Rst.RecordCount Then Rst.MoveNext
                Case vbKeyPageUp '33
                    For I = 1 To 10
                        If Rst.AbsolutePosition > 1 Then Rst.MovePrevious
                    Next
                Case vbKeyPageDown '34
                    For I = 1 To 10
                        If Rst.AbsolutePosition < Rst.RecordCount Then Rst.MoveNext
                    Next
            End Select
            If Rst.BOF = False And Rst.EOF = False Then
                If GridText = True Then
                    Txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                Else
                    Txt(Index).TEXT = Rst.Fields(HelpIndex).Value
                    Txt(Index).Tag = Rst.Fields(0).Value
                End If
                Txt(Index).SelStart = Len(Txt(Index))
            End If
      End If
      Exit Sub
  End If
'Else
'  If DGrid.Visible = False Then DGrid.Visible = True: DGrid.ZOrder 0
'End If
End Sub

Public Sub DGridTxtKeyUp_Mast(Txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, FindFldName As String)
Dim STR$    ' As String
Dim LPlace As Byte
    If Rst.RecordCount <= 0 Then Exit Sub
    If KeyCode = 13 Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Exit Sub
    LPlace = Txt(Index).SelStart
    STR = mID(Txt(Index).TEXT, 1, LPlace)
    Rst.MoveFirst
    If Rst.Fields(FindFldName).Type = adInteger Then
        Rst.FIND "" & FindFldName & "  >=" & Val(STR) & ""
    Else
        Rst.FIND "" & FindFldName & "  >='" & STR & "'"
        'Rst.FIND "" & FindFldName & "  Like '" & STR & "'"
    End If
    If Rst.EOF = True Then Rst.MoveFirst
End Sub

Public Function CheckFinYear(Temp As Variant) As Boolean
    If IsNull(Temp) Or Temp = "" Then
        CheckFinYear = False
    Else
        If CDate(Format(Temp, "dd/MMM/yyyy")) < PubStartDate Or CDate(Format(Temp, "dd/MMM/yyyy")) > PubEndDate Then
            CheckFinYear = False
        Else
            CheckFinYear = True
        End If
    End If
    If CheckFinYear = False Then MsgBox "Entered Date is beyond Financial Year!", vbCritical, "Financial Year Validation"
End Function

' Used For Getting the Rate of Item Whether it is MRP/TB_Rate/TP_Rate
Public Function GetRate( _
        PartyType As Byte, FGrid As MSHierarchicalFlexGridLib.MSHFlexGrid, VDate As Date, _
        PartCode As String, ByVal Col_MRPYN As Byte, ByVal MRPRate As Double, _
        ByVal Col_TaxableYN As Byte, ByVal TBRate As Double, ByVal TPRate As Double, _
        ByVal Col_EffectiveDate As Byte, Col_MRPRate As Byte, Optional Purpose As String, Optional LastRate As Double) As Double
        
Dim Rst As ADODB.Recordset, MRPDisc As Single, TBDisc As Single, TPDisc As Single
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select MRP_Disc,TB_Disc,TP_Disc from SubGroupType " _
            & "where Party_Type =" & PartyType, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        MRPDisc = Rst!mrp_Disc
        TBDisc = Rst!tb_Disc
        TPDisc = Rst!tp_Disc
    End If
    If StrCmp(Purpose, "Warranty") Then
        GetRate = LastRate
        FGrid.TextMatrix(FGrid.Row, Col_MRPRate) = Format(LastRate, "0.00")
    Else
        If Trim(FGrid.TextMatrix(FGrid.Row, Col_MRPYN)) = "Yes" And _
            Trim(FGrid.TextMatrix(FGrid.Row, Col_TaxableYN)) = "Yes" Then
            If FGrid.TextMatrix(FGrid.Row, Col_EffectiveDate) = "" Then
                GetRate = MRPRate
            ElseIf CDate(Format(VDate, "dd/MMM/yyyy")) >= CDate(Format(Trim(FGrid.TextMatrix(FGrid.Row, Col_EffectiveDate)), "dd/MMM/yyyy")) Then
                GetRate = MRPRate
            Else
                If MRPDisc <> 0 Then
                    Set Rst = GCn.Execute("Select round(MRP-(MRP*" & MRPDisc & "/100),2) as MRP From Part_PriceList Where Part_No='" & PartCode & "' And Effect_Dt<=" & ConvertDate(VDate) & " Order By Effect_Dt Desc")
                Else
                    Set Rst = GCn.Execute("Select MRP as MRP From Part_PriceList Where Part_No='" & PartCode & "' And Effect_Dt<=" & ConvertDate(VDate) & " Order By Effect_Dt Desc")
                    Set Rst = GCn.Execute("Select MRP as MRP From Part Where Part_No='" & PartCode & "' ")
                End If
                If Rst.RecordCount > 0 Then
                    GetRate = Rst!MRP
                    FGrid.TextMatrix(FGrid.Row, Col_MRPRate) = Format(Rst!MRP, "0.00")
                End If
            End If
        Else
            If Trim(FGrid.TextMatrix(FGrid.Row, Col_TaxableYN)) = "Yes" Then
                If FGrid.TextMatrix(FGrid.Row, Col_EffectiveDate) = "" Then
                    GetRate = TBRate
                ElseIf CDate(Format(VDate, "dd/MMM/yyyy")) >= CDate(Format(Trim(FGrid.TextMatrix(FGrid.Row, Col_EffectiveDate)), "dd/MMM/yyyy")) Then
                    GetRate = TBRate
                Else
                    If TBDisc <> 0 Then
                        Set Rst = GCn.Execute("Select round(TB_SRate-(TB_SRate*" & TBDisc & "/100),2) as TB_SRate From Part_PriceList Where Part_No='" & PartCode & "' And Effect_Dt<=" & ConvertDate(VDate) & " Order By Effect_Dt Desc")
                    Else
                        Set Rst = GCn.Execute("Select TB_SRate as TB_SRate From Part_PriceList Where Part_No='" & PartCode & "' And Effect_Dt<=" & ConvertDate(VDate) & " Order By Effect_Dt Desc")
                    End If
                    If Rst.RecordCount > 0 Then
                        GetRate = Rst!TB_SRate
                    End If
                End If
            Else
                If FGrid.TextMatrix(FGrid.Row, Col_EffectiveDate) = "" Then
                    GetRate = TPRate
                ElseIf CDate(Format(VDate, "dd/MMM/yyyy")) >= CDate(Format(Trim(FGrid.TextMatrix(FGrid.Row, Col_EffectiveDate)), "dd/MMM/yyyy")) Then
                    GetRate = TPRate
                Else
                    If TPDisc <> 0 Then
                        Set Rst = GCn.Execute("Select round(TP_SRate-(TP_SRate*" & TPDisc & "/100),2) as TP_SRate From Part_PriceList Where Part_No='" & PartCode & "' And Effect_Dt<=" & ConvertDate(VDate) & " Order By Effect_Dt Desc")
                    Else
                        Set Rst = GCn.Execute("Select TP_SRate as TP_SRate From Part_PriceList Where Part_No='" & PartCode & "' And Effect_Dt<=" & ConvertDate(VDate) & " Order By Effect_Dt Desc")
                    End If
                    If Rst.RecordCount > 0 Then
                        GetRate = Rst!TP_SRate
                    End If
                End If
            End If
        End If
    End If
Set Rst = Nothing
End Function

Public Sub FormKeyDown(FrmName As Form, KeyCode As Integer, Shift As Integer, Optional ByVal MasterFormExit As Boolean)
If Shift = 2 And KeyCode = vbKeyR Then
    If frmCustRect.Visible = False Then
        frmCustRect.Show
    Else
        frmCustRect.ZOrder 0
    End If
End If
FrmName.TopCtrl1.PrvKeyCode = KeyCode 'modify shekhar for insert change
If Not IsMissing(MasterFormExit) Then
    If MasterFormExit <> False Then
        If (Shift = 2 And KeyCode = vbKeyS) Then
            FrmName.TopCtrl1.TopKey_Down KeyCode, Shift
'            Unload FrmName
        ElseIf KeyCode = vbKeyEscape Then
            Unload FrmName
        End If
        Exit Sub
    End If
End If
If KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Or KeyCode = vbKeyF4 Or _
    (KeyCode = vbKeyF And Shift = 2) Or (KeyCode = vbKeyP And Shift = 2) Or _
    (KeyCode = vbKeyS And Shift = 2) Or KeyCode = vbKeyEscape Or _
    KeyCode = vbKeyF5 Or KeyCode = vbKeyF10 Or KeyCode = vbKeyHome Or _
    KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEnd Then
    FrmName.TopCtrl1.TopKey_Down KeyCode, Shift
    
    If KeyCode = vbKeyS And Shift = 2 Then KeyCode = 0 'by lps 15-06-2002
End If
If KeyCode <> vbKeyF10 Then
    If FrmName.TopCtrl1.PrvKeyCode = vbKeyEscape Then
        FrmName.TopCtrl1.PrvKeyCode = 0
    Else
        FrmName.TopCtrl1.PrvKeyCode = KeyCode
    End If
End If
End Sub

Public Function GetDocIDmPrefix(FACn As ADODB.Connection, ByVal VType As String, ByVal VDate As String, _
    ByRef VoucherEditFlag As Boolean, ByRef TxtSrlNo As Object, _
    ByRef lblPrefix As Object, Optional ForSiteCode As String) As String
Dim Rst As ADODB.Recordset, VNo As Long, NotExists As Boolean
Dim TEMPSQL$, DivBaseNumber As Boolean, FaVoucher As Boolean


    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    If PubBackEnd = "A" Then
        Set Rst = FACn.Execute("Select distinct switch(Category='FA',True,Category<>'FA',False) as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
    Else
        Set Rst = FACn.Execute("Select distinct (Case When Category='FA' then 'True' else 'False' End ) as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
    End If
    If Rst.RecordCount <= 0 Then
        MsgBox "Please Add Record in Voucher Type Table in FA Data" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error"
        Set Rst = Nothing
        GetDocIDmPrefix = "": Exit Function
    End If

    FaVoucher = Rst!FaVoucher
    DivBaseNumber = IIf(Rst!DivBaseNumber = 0, False, True)
    If IsMissing(ForSiteCode) Then
        ForSiteCode = PubSiteCode
    ElseIf ForSiteCode = "" Then
        ForSiteCode = PubSiteCode
    End If
    
    If lblPrefix = "" Then
        lblPrefix = G_FaCn.Execute("Select Top 1 " & xIsNull("Prefix", "") & " From Voucher_Prefix Where V_Type = '" & VType & "'").Fields(0).Value
    End If
    
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    'No Division Base No. in FA /(divison base no introduced by lps at udaipur
    'Voucher No. From FA Data as per connection passed
    TEMPSQL = "Select V_Type from Voucher_Prefix VP Where VP.V_Type='" & VType & "' And VP.Date_From<=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " "
    If DivBaseNumber Then
        TEMPSQL = TEMPSQL + " and VP.Div_Code='" & PubDivCode & "'"
    End If
    TEMPSQL = TEMPSQL + " Order By VP.Date_From DESC"
    If FACn.Execute(TEMPSQL).RecordCount > 0 Then
        TEMPSQL = "Select Top 1 VT.Number_Method,VT.SerialNo_From_Table,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No+1 as Start_Srl_No " & _
            " From Voucher_Type VT " & _
            " Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type " & _
            " Where VP.V_Type='" & VType & "' And Prefix='" & lblPrefix & "' And VP.Date_From<=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " And IsNull(VP.Site_Code,'')=(Case When IsNull(VT.SiteBaseNumber,'')='Y' Then '" & PubSiteCode & "' Else IsNull(VP.Site_Code,'') End)  "
        If DivBaseNumber Then
            TEMPSQL = TEMPSQL + " and VP.Div_Code='" & PubDivCode & "'"
        End If
        TEMPSQL = TEMPSQL + " Order By VP.Div_Code,VP.Date_From DESC"
        Rst.Open TEMPSQL, FACn, adOpenStatic, adLockReadOnly
    Else
        'Applicable for No Records in Prefix Table & Manual Only
        MsgBox "Please Add Record in Voucher Prefix Table " & vbCrLf & " through Voucher Controls under Utility Menu" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error": GetDocIDmPrefix = "": Exit Function
        GetDocIDmPrefix = ""
        GoTo errlbl
    End If
    If Rst.RecordCount <= 0 Then
        MsgBox "Please Define Document Numbering System  " & vbCrLf & " in Voucher Controls under Utility Menu", vbCritical, "System Configuration"
        GetDocIDmPrefix = ""
        GoTo errlbl
    End If
    '*---------
    'lblPrefix = Rst!Prefix
    If Rst!Number_Method = "Manual" Then
        VoucherEditFlag = True
        TxtSrlNo.Enabled = True
        VNo = Val(TxtSrlNo)
    Else
        VoucherEditFlag = False
        TxtSrlNo.Enabled = False

    Select Case UCase(Rst!SerialNo_From_Table)
'        Case "DIVISION"     '' Serial No. Required from Division Table
'            If Rst!V_Type = "W_JC" Then     '' JobCard
'                GSQL = "select JobCard_SrlNo+1 from Division where Div_code='" & PubDivCode & "'"
'            ElseIf Rst!V_Type = "W_RG" Then     '' General Requisition
'                GSQL = "select IPO_Gen_SrlNo+1 from Division where Div_code='" & PubDivCode & "'"
'            ElseIf Rst!V_Type = "W_RW" Then     '' Warranty Requisition
'                GSQL = "select IPO_War_SrlNo+1 from Division where Div_code='" & PubDivCode & "'"
'            End If
'            VNo = GCn.Execute(GSQL).Fields(0).Value
        Case "SP_ORDCOUN"
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select Start_No+1 as Start_Srl_No,End_No from Sp_OrdCoun where ord_type='" & VType & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst!end_no > 0 And Rst!start_srl_no > Rst!end_no Then
                MsgBox "Alloted Serials for this Order Type is Complete", vbInformation, "Validation"
                GoTo errlbl
            End If
            VNo = Rst!start_srl_no
        Case Else
            VNo = Rst!start_srl_no
    End Select
    End If
    If VNo > 0 Then
        TxtSrlNo = VNo
    End If
    GetDocIDmPrefix = PubDivCode + PubSiteCode + ForSiteCode + Space(5 - Len(CStr(VType))) + VType + Space(5 - Len(CStr(Rst!Prefix))) + Rst!Prefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
errlbl:
    Set Rst = Nothing
End Function



Public Function GetDocID(FACn As ADODB.Connection, ByVal VType As String, ByVal VDate As String, _
    ByRef VoucherEditFlag As Boolean, ByRef TxtSrlNo As Object, _
    ByRef lblPrefix As Object, Optional ForSiteCode As String) As String
Dim Rst As ADODB.Recordset, VNo As Long, NotExists As Boolean
Dim TEMPSQL$, DivBaseNumber As Boolean, FaVoucher As Boolean
'10-03-03
'Voucher_Type & Voucher_Prefix shiifted to FAData only
'Change in connection CGN to FACn
'    If FACn.Execute("Select distinct Category,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & Vtype & "'").RecordCount <= 0 Then
'        MsgBox "Please Add Record in Voucher Type Table in FA Data" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error": GetDocID = "": Exit Function
'        GetDocID = ""
'        GoTo errlbl
'    Else
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        'Set Rst = FACn.Execute("Select distinct switch(Category='FA',True,Category<>'FA',False) as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
        Set Rst = FACn.Execute("Select distinct " & cIIF("Category='FA'", cBoolean(True), cBoolean(False)) & " as FAVoucher,DivBaseNumber from Voucher_Type VT Where VT.V_Type='" & VType & "'")
'    End If
    If Rst.RecordCount <= 0 Then
        MsgBox "Please Add Record in Voucher Type Table in FA Data" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error"
        Set Rst = Nothing
        GetDocID = "": Exit Function
    End If

    FaVoucher = Rst!FaVoucher
    DivBaseNumber = IIf(Rst!DivBaseNumber = 0, False, True)
    If IsMissing(ForSiteCode) Then
        ForSiteCode = PubSiteCode
    ElseIf ForSiteCode = "" Then
        ForSiteCode = PubSiteCode
    End If
    
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    'No Division Base No. in FA /(divison base no introduced by lps at udaipur
    'Voucher No. From FA Data as per connection passed
    TEMPSQL = "Select V_Type from Voucher_Prefix VP Where VP.V_Type='" & VType & "' And VP.Date_From<=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " "
    If DivBaseNumber Then
        TEMPSQL = TEMPSQL + " and VP.Div_Code='" & PubDivCode & "'"
    End If
    TEMPSQL = TEMPSQL + " Order By VP.Date_From DESC"
    If FACn.Execute(TEMPSQL).RecordCount > 0 Then
        TEMPSQL = "Select Top 1 VT.Number_Method,VT.SerialNo_From_Table,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No+1 as Start_Srl_No " & _
            " From Voucher_Type VT " & _
            " Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type " & _
            " Where VP.V_Type='" & VType & "' And VP.Date_To>=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " And VP.Date_From<=" & ConvertDate(Format(VDate, "dd/MMM/yyyy")) & " And IsNull(VP.Site_Code,'')=(Case When IsNull(VT.SiteBaseNumber,'')='Y' Then '" & PubSiteCode & "' Else IsNull(VP.Site_Code,'') End) "
        If DivBaseNumber Then
            TEMPSQL = TEMPSQL + " and VP.Div_Code='" & PubDivCode & "'"
        End If
        TEMPSQL = TEMPSQL + " Order By VP.Div_Code,VP.Date_From DESC"
        Rst.Open TEMPSQL, FACn, adOpenStatic, adLockReadOnly
    Else
        'Applicable for No Records in Prefix Table & Manual Only
        MsgBox "Please Add Record in Voucher Prefix Table " & vbCrLf & " through Voucher Controls under Utility Menu" & vbCrLf & "Document ID Creation failed!", vbCritical, "Fatal Error": GetDocID = "": Exit Function
        GetDocID = ""
        GoTo errlbl
    End If
    If Rst.RecordCount <= 0 Then
        MsgBox "Please Define Document Numbering System  " & vbCrLf & " in Voucher Controls under Utility Menu", vbCritical, "System Configuration"
        GetDocID = ""
        GoTo errlbl
    End If
    '*---------
    lblPrefix = Rst!Prefix
    If Rst!Number_Method = "Manual" Then
        VoucherEditFlag = True
        TxtSrlNo.Enabled = True
        VNo = Val(TxtSrlNo)
    Else
        VoucherEditFlag = False
        TxtSrlNo.Enabled = False

    Select Case UCase(Rst!SerialNo_From_Table)
'        Case "DIVISION"     '' Serial No. Required from Division Table
'            If Rst!V_Type = "W_JC" Then     '' JobCard
'                GSQL = "select JobCard_SrlNo+1 from Division where Div_code='" & PubDivCode & "'"
'            ElseIf Rst!V_Type = "W_RG" Then     '' General Requisition
'                GSQL = "select IPO_Gen_SrlNo+1 from Division where Div_code='" & PubDivCode & "'"
'            ElseIf Rst!V_Type = "W_RW" Then     '' Warranty Requisition
'                GSQL = "select IPO_War_SrlNo+1 from Division where Div_code='" & PubDivCode & "'"
'            End If
'            VNo = GCn.Execute(GSQL).Fields(0).Value
        Case "SP_ORDCOUN"
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select Start_No+1 as Start_Srl_No,End_No from Sp_OrdCoun where ord_type='" & VType & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst!end_no > 0 And Rst!start_srl_no > Rst!end_no Then
                MsgBox "Alloted Serials for this Order Type is Complete", vbInformation, "Validation"
                GoTo errlbl
            End If
            VNo = Rst!start_srl_no
        Case Else
            VNo = Rst!start_srl_no
    End Select
    End If
    If VNo > 0 Then
        TxtSrlNo = VNo
    End If
    GetDocID = PubDivCode + PubSiteCode + ForSiteCode + Space(5 - Len(CStr(VType))) + VType + Space(5 - Len(CStr(Rst!Prefix))) + Rst!Prefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
errlbl:
    Set Rst = Nothing
End Function

Public Function PubServerDate() As Date
    PubServerDate = Format(CDate(CStr(date) & " " & CStr(time)), "dd/MMM/yyyy hh:mm:ss")
End Function

Public Function SumTime(Obj As Object, Col_Val1 As Integer, ObjType As ObjTypeDef, AddDed As String, Optional Col_Val2 As Integer) As Single
Dim mHr As Single, mMin As Single
Dim I As Integer
If AddDed = "Add" Then
    If ObjType = 0 Then
        If Obj.RecordCount > 0 Then
            Do While Obj.EOF = False
                mHr = mHr + left(Obj.Fields(Col_Val1).Value, 2) + IIf(Col_Val2 = 0, 0, left(Obj.Fields(Col_Val2).Value, 2))
                mMin = mMin + Right(Obj.Fields(Col_Val1).Value, 2) + IIf(Col_Val2 = 0, 0, Right(Obj.Fields(Col_Val2).Value, 2))
            Obj.MoveNext
            Loop
        End If
    ElseIf ObjType = 1 Then
         For I = 1 To Obj.Rows - 1
                mHr = mHr + left(Obj.TextMatrix(I, Obj.Col1), 2) + IIf(Col_Val2 = 0, 0, left(Obj.TextMatrix(I, Obj.Col2), 2))
                mMin = mMin + Right(Obj.TextMatrix(I, Obj.Col1), 2) + IIf(Col_Val2 = 0, 0, Right(Obj.TextMatrix(I, Obj.Col2), 2))
         Next
    ElseIf ObjType = 2 Then
                mHr = left(Obj(Obj.Col1).TEXT, 2) + IIf(Col_Val2 = 0, 0, left(Obj(Obj.Col1).TEXT, 2))
                mMin = Right(Obj(Obj.Col1).TEXT, 2) + IIf(Col_Val2 = 0, 0, Right(Obj(Obj.Col1).TEXT, 2))
    End If
Else
    If ObjType = 0 Then
        If Obj.RecordCount > 0 Then
            Do While Obj.EOF = False
                mHr = mHr + left(Obj.Fields(Col_Val1).Value, 2) - IIf(Col_Val2 = 0, 0, left(Obj.Fields(Col_Val2).Value, 2))
                mMin = mMin + Right(Obj.Fields(Col_Val1).Value, 2) - IIf(Col_Val2 = 0, 0, Right(Obj.Fields(Col_Val2).Value, 2))
            Obj.MoveNext
            Loop
        End If
    ElseIf ObjType = 1 Then
        For I = 1 To Obj.Rows - 1
           mHr = mHr + left(Obj.TextMatrix(I, Obj.Col1), 2) - IIf(Col_Val2 = 0, 0, left(Obj.TextMatrix(I, Obj.Col2), 2))
           mMin = mMin + Right(Obj.TextMatrix(I, Obj.Col1), 2) - IIf(Col_Val2 = 0, 0, Right(Obj.TextMatrix(I, Obj.Col2), 2))
        Next
    ElseIf ObjType = 2 Then
            mHr = left(Obj(Obj.Col1).TEXT, 2) - IIf(Col_Val2 = 0, 0, left(Obj(Obj.Col1).TEXT, 2))
            mMin = Right(Obj(Obj.Col1).TEXT, 2) - IIf(Col_Val2 = 0, 0, Right(Obj(Obj.Col1).TEXT, 2))
    End If
End If
        SumTime = mHr + (mMin \ 60) + ((mMin Mod 60) / 100)
    

End Function

Public Sub CreateSprIndent(DocID As String, Doc_Type As String, Site_Code As String, IndDate As Date, PartyCode As String, Part_No As String)
'Insert Record in Indent Table
End Sub

Public Function CheckSprStock( _
    FGrid As MSHierarchicalFlexGridLib.MSHFlexGrid, FRow As Integer, _
    Col_MRPYN As Byte, Col_TaxableYN As Byte, _
    Col_Qty As Byte, Col_MRPStkTB As Byte, _
    Col_MRPStkTP As Byte, Col_TBQty As Byte, _
    Col_TPQty As Byte) As Boolean
    
Dim MsgSQL$
If UCase(left(PubComp_Name, 3)) = "JMK" Then
        If Trim(FGrid.TextMatrix(FRow, Col_TaxableYN)) = "Yes" Then
            If Val(FGrid.TextMatrix(FRow, Col_Qty)) > Val(FGrid.TextMatrix(FRow, Col_MRPStkTB)) + Val(FGrid.TextMatrix(FRow, Col_TBQty)) Then
                MsgSQL = "Taxable MRP Qty  > Taxable MRP Stock in Hand"
                GoTo lblExit
            End If
        Else
            If Val(FGrid.TextMatrix(FRow, Col_Qty)) > Val(FGrid.TextMatrix(FRow, Col_MRPStkTP)) + Val(FGrid.TextMatrix(FRow, Col_TPQty)) Then
                MsgSQL = "Taxpaid MRP Qty  > Taxpaid MRP Stock in Hand"
                GoTo lblExit
            End If
        End If
Else
    If Trim(FGrid.TextMatrix(FRow, Col_MRPYN)) = "Yes" Then
        If Trim(FGrid.TextMatrix(FRow, Col_TaxableYN)) = "Yes" Then
            If Val(FGrid.TextMatrix(FRow, Col_Qty)) > Val(FGrid.TextMatrix(FRow, Col_MRPStkTB)) Then
                MsgSQL = "Taxable MRP Qty  > Taxable MRP Stock in Hand"
                GoTo lblExit
            End If
        Else
            If Val(FGrid.TextMatrix(FRow, Col_Qty)) > Val(FGrid.TextMatrix(FRow, Col_MRPStkTP)) Then
                MsgSQL = "Taxpaid MRP Qty  > Taxpaid MRP Stock in Hand"
                GoTo lblExit
            End If
        End If
    Else
        If Trim(FGrid.TextMatrix(FRow, Col_TaxableYN)) = "Yes" Then
            If Val(FGrid.TextMatrix(FRow, Col_Qty)) > Val(FGrid.TextMatrix(FRow, Col_TBQty)) Then
                MsgSQL = "Taxable Qty  > Taxable Stock in Hand"
                GoTo lblExit
            End If
        Else
            If Val(FGrid.TextMatrix(FRow, Col_Qty)) > Val(FGrid.TextMatrix(FRow, Col_TPQty)) Then
                MsgSQL = "Taxpaid Qty  > Taxpaid Stock in Hand"
                GoTo lblExit
            End If
        End If
    End If
End If
lblExit:
    If MsgSQL = "" Then
        CheckSprStock = True
    Else
        MsgBox MsgSQL, vbInformation, "Validation"
        If PubSprIssOnNegStk = 1 Then
            CheckSprStock = True
        Else
            CheckSprStock = False
        End If
    End If
End Function

Public Function RetMonth(ByRef Txt As Object) As String
On Error GoTo err1
Dim mMonth As String
If Len(Trim(Txt)) = 0 Then
    RetMonth = ""
    Exit Function
End If
        Select Case mID(Trim(UCase(Txt)), 1, 3)
            Case "1", "01", "J", "JA", "JAN"
                mMonth = "January"
            Case "2", "02", "F", "FE", "FEB"
                mMonth = "February"
            Case "3", "03", "M", "MA", "MAR"
                mMonth = "March"
            Case "4", "04", "A", "AP", "APR"
                mMonth = "April"
            Case "5", "05", "MAY"
                mMonth = "May"
            Case "6", "06", "JU", "JUN"
                mMonth = "June"
            Case "7", "07", "JUL"
                mMonth = "July"
            Case "8", "08", "AU", "AUG"
                mMonth = "August"
            Case "9", "09", "S", "SE", "SEP"
                mMonth = "September"
            Case "10", "O", "OC", "OCT"
                mMonth = "October"
            Case "11", "N", "NO", "NOV"
                mMonth = "November"
            Case "12", "D", "DE", "DEC"
               mMonth = "December"
            Case Else
                mMonth = Format(date, "MMMM")
        End Select
        RetMonth = mMonth
        Exit Function
err1:
    ' For Overflow Check
    If err.NUMBER = 6 Then Resume Next
End Function

Public Sub txtDisabled_Color(Frm As Form)
Dim objctrl As Control
For Each objctrl In Frm.Controls
'If (TypeOf objctrl Is TextBox) Or (TypeOf objctrl Is MaskEdBox) Then
If (TypeOf objctrl Is TextBox) Then
    If objctrl.Enabled = False Then
        objctrl.BackColor = CtrlBColDisabled
    Else
        objctrl.BackColor = CtrlBColOrg
    End If
End If
Next objctrl
End Sub

Public Sub DGridTxtKeyPress(Txt As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyCode As Integer, FindFldName As String, Optional KeyUpCall As Boolean)
Dim FindStr$    ' As String
Dim LPlace As Byte
    If Rst.RecordCount <= 0 Or (Txt(Index) = "" And KeyCode = vbKeyDelete) Then Txt(Index).TEXT = "": Exit Sub
    If KeyCode = 13 Or KeyCode = 8 Or KeyCode = 16 Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then: Exit Sub

    If IsMissing(KeyUpCall) Or KeyUpCall = False Then 'KeyPressCall
        If Txt(Index).TEXT = "" Then
            FindStr = Chr(KeyCode)
        Else
            FindStr = Txt(Index).TEXT + Chr(KeyCode)
        End If
        'ModiShekhar23jan On Blank Text It is producing Problem at press of esc
        If FindStr = "" Then Exit Sub
        'EndModi23jan
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
            MsgBox "Please convert to String Search", vbOKOnly
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
            If Rst.AbsolutePosition = adPosEOF Then
                FindStr = Txt(Index).TEXT   'left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
            ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                FindStr = Txt(Index).TEXT   'left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
            End If
        Else    'character serach
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
            If Rst.AbsolutePosition = adPosEOF Then
                FindStr = Txt(Index).TEXT
                Rst.MoveFirst
                If FindStr <> "" Then Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
            ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                FindStr = Txt(Index).TEXT
                Rst.MoveFirst
                Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
            End If
        End If
        If FindStr = Txt(Index).TEXT + Chr(KeyCode) Then
            Txt(Index).TEXT = Txt(Index).TEXT + Chr(KeyCode)
        End If
        KeyCode = 0
    Else    'KeyUp Call Search as per Old Process
        LPlace = Txt(Index).SelStart
        FindStr = Txt(Index).TEXT
        'ModiShekhar23jan On Blank Text It is producing Problem at press of esc
        If FindStr = "" Then Exit Sub
        'EndModi23jan
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
            If Rst.AbsolutePosition = adPosEOF Then
                FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
                Txt(Index).TEXT = FindStr
            ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
                Txt(Index).TEXT = FindStr
            End If
        Else
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
            If Rst.AbsolutePosition = adPosEOF Then
                FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                If FindStr <> "" Then Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
                Txt(Index).TEXT = FindStr
            ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
                FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
                Rst.MoveFirst
                If FindStr <> "" Then Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
                Txt(Index).TEXT = FindStr
            End If
        End If
        Txt(Index).SelStart = Len(Txt(Index).TEXT)
        KeyCode = 0
    End If
    Txt(Index).SelStart = Len(Txt(Index))
End Sub

Public Sub DGridTxtKeyUp1(Txt As Object, Index As Integer, Rst As ADODB.Recordset, KeyCode As Integer, FindFldName As String)
Dim FindStr$    ' As String
Dim LPlace As Byte
    If Rst.RecordCount <= 0 Then Txt(Index).TEXT = "": Exit Sub
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then Exit Sub
    LPlace = Txt(Index).SelStart
    FindStr = Txt(Index).TEXT
'    FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
    Rst.MoveFirst
    If Rst.Fields(FindFldName).Type = adInteger Then
        Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
        If Rst.AbsolutePosition = adPosEOF Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
            Txt(Index).TEXT = FindStr
        ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
            Txt(Index).TEXT = FindStr
        End If
    Else
        Rst.FIND "" & FindFldName & "  >='" & FindStr & "'"
        If Rst.AbsolutePosition = adPosEOF Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & "  >='" & FindStr & "'"
            Txt(Index).TEXT = FindStr
        ElseIf (UCase(mID(Rst.Fields(FindFldName).Value, 1, Len(FindStr))) <> UCase(FindStr)) Then
            FindStr = left(FindStr, LPlace - 1) + Right(FindStr, Len(FindStr) - LPlace)
            Rst.MoveFirst
            Rst.FIND "" & FindFldName & "  >='" & FindStr & "'"
            Txt(Index).TEXT = FindStr
        End If
    End If
    Txt(Index).SelStart = Len(Txt(Index).TEXT)
End Sub

Public Function RetDGKeyAscii(ByRef GridVar As Boolean, KeyAscii As Integer) As Integer
    '' Purpose : To Prohibit Extra keystrokes from User during DataGrid Help System
    '' Method  :    Initialize one variable on form  as boolean
    ''              Initialize that variable on text got focus with false
    ''              Call This Function on KeyPress Event of that text box
    ''              Initialize that variable with false on keyup event of that text box
    If GridVar = False Then
        GridVar = True
    Else
        KeyAscii = 0
    End If
    RetDGKeyAscii = KeyAscii
End Function

Public Function RestrictCode(KeyCode, Txt As Object, ByRef Shift As Integer, Optional Lock2ndPlace As Boolean)
'Purpose : given code restrict entered Code edit, added as prefix in code generation

If Len(Txt) = 1 And Lock2ndPlace = False Then
    If (KeyCode = vbKeyBack Or KeyCode = vbKeyLeft Or KeyCode = vbKeyHome) _
        Or (Shift = 1 And (KeyCode = vbKeyLeft Or KeyCode = vbKeyHome) _
        Or KeyCode = vbKeyDelete Or KeyCode = vbKeyUp) Then
        RestrictCode = 0: Shift = 0
    Else
        RestrictCode = KeyCode
    End If
ElseIf Len(Txt) = 2 And Lock2ndPlace Then
    If (KeyCode = vbKeyBack Or KeyCode = vbKeyLeft Or KeyCode = vbKeyHome) _
        Or (Shift = 1 And (KeyCode = vbKeyLeft Or KeyCode = vbKeyHome) _
        Or KeyCode = vbKeyDelete Or KeyCode = vbKeyUp) Then
        RestrictCode = 0: Shift = 0
    Else
        RestrictCode = KeyCode
    End If
Else
    RestrictCode = KeyCode
End If
End Function

Public Function RestrictKey(LockLen As Byte, KeyCode, Txt As Object, ByRef Shift As Integer)
'Purpose : given code restrict Site Code edit, added as prefix in code generation
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    RestrictKey = KeyCode
ElseIf Len(Txt) = LockLen Then
    If (KeyCode = vbKeyBack Or KeyCode = vbKeyLeft Or KeyCode = vbKeyHome) _
        Or (Shift = 1 And (KeyCode = vbKeyLeft Or KeyCode = vbKeyHome) _
        Or KeyCode = vbKeyDelete Or KeyCode = vbKeyUp) Then
        RestrictKey = 0: Shift = 0
    Else
        RestrictKey = KeyCode
    End If
Else
    RestrictKey = KeyCode
End If
End Function

Public Function FxStatus(ByVal Code As Byte) As String
'Used in forms : frmWarrantyPCR, frmWarrantyWCD
Select Case Code
    Case 1
        FxStatus = "Drive Away"
    Case 2
        FxStatus = "Sold"
    Case Else
        FxStatus = "** Unknown **"
End Select
End Function

Public Function FxFailure(ByVal Code As Byte) As String
'Used in forms : frmWarrantyPCR, frmWarrantyWCD
Select Case Code
    Case 1
        FxFailure = "OE"
    Case 2
        FxFailure = "Repeat"
    Case 3
        FxFailure = "Spare Parts"
    Case Else
        FxFailure = "** Unknown **"
End Select
End Function

Public Function FxOperation(ByVal Code As Byte) As String
'Used in forms : frmWarrantyPCR, frmWarrantyWCD
Select Case Code
    Case 1
        FxOperation = "Drive Away"
    Case 2
        FxOperation = "Long Route"
    Case 3
        FxOperation = "City Route"
    Case 4
        FxOperation = "Construction"
    Case 5
        FxOperation = "Mining"
    Case 6
        FxOperation = "Forest"
    Case 7
        FxOperation = "Marine"
    Case 8
        FxOperation = "Others"
    Case Else
        FxOperation = "** Unknown **"
End Select
End Function

Public Function FxRoad(ByVal Code As Byte) As String
'Used in forms : frmWarrantyPCR, frmWarrantyWCD
Select Case Code
    Case 1
        FxRoad = "Plain Metalled"
    Case 2
        FxRoad = "Plain Kutcha"
    Case 3
        FxRoad = "Off Road"
    Case 4
        FxRoad = "Hilly Metalled"
    Case 5
        FxRoad = "Killy Kutcha"
    Case 6
        FxRoad = "Desert"
    Case 7
        FxRoad = "Others"
    Case Else
        FxRoad = "** Unknown **"
End Select
End Function

Public Function DeCodeDocID(ByVal DocID As String, Decode As ObjTypeDef1) As String
    Select Case Decode
        Case 1
            DeCodeDocID = mID(DocID, 1, 1)
        Case 2
            DeCodeDocID = mID(DocID, 2, 1)
        Case 3
            DeCodeDocID = mID(DocID, 3, 1)
        Case 4
            DeCodeDocID = mID(DocID, 4, 5)
        Case 5
            DeCodeDocID = mID(DocID, 9, 5)
        Case 6
            DeCodeDocID = Right(DocID, 8)
    End Select
End Function

Public Sub RDisp(AdoName As ADODB.Recordset, FrmName As Form)
FrmName.TopCtrl1.TopText1.Alignment = 1
FrmName.TopCtrl1.TopText1.ForeColor = &HFF00FF
FrmName.TopCtrl1.TopText1 = "# " & IIf(AdoName.AbsolutePosition = adPosUnknown, 0, AdoName.AbsolutePosition) & "/" & AdoName.RecordCount

End Sub


Public Function FxFormTrnType(TrnType As Variant) As Variant
If IsNumeric(TrnType) Then
    If TrnType = 0 Then
        FxFormTrnType = "NA"
    ElseIf TrnType = 1 Then
        FxFormTrnType = "Issue"
    ElseIf TrnType = 2 Then
        FxFormTrnType = "Receipt"
    ElseIf TrnType = 3 Then
        FxFormTrnType = "Both"
    Else
        FxFormTrnType = "InValid"
    End If
Else
    If TrnType = "NA" Then
        FxFormTrnType = 0
    ElseIf TrnType = "Issue" Then
        FxFormTrnType = 1
    ElseIf TrnType = "Receipt" Then
        FxFormTrnType = 2
    ElseIf TrnType = "Both" Then
        FxFormTrnType = 3
    Else
        FxFormTrnType = 4
    End If
End If
End Function

''''' ADDED BY SKG

Public Sub INI_COMBO(sqlstr As String, DBCNAME As DataCombo, LSTFIELD As String, BNDCOLUMN As String)
Set DBCNAME.RowSource = G_FaCn.Execute(sqlstr)
DBCNAME.ListField = LSTFIELD
DBCNAME.BoundColumn = BNDCOLUMN
DBCNAME.Tag = sqlstr
End Sub

Public Sub REFR_COMBO(DBC As DataCombo)
    Dim BT
    BT = DBC.BoundText
    Call INI_COMBO(DBC.Tag, DBC, DBC.ListField, DBC.BoundColumn)
    DBC.BoundText = BT
End Sub

Public Function VALID_DATE(FRMNAME1 As Form) As Integer
VALID_DATE = 1
If VALID_DATE_CHK(FRMNAME1.TXTS_DATE, "Starting Date") = False Then VALID_DATE = 0: Exit Function
If VALID_DATE_CHK(FRMNAME1.TXTE_DATE, "Ending Date") = False Then VALID_DATE = 0: Exit Function
If DateDiff("d", FRMNAME1.TXTS_DATE, FRMNAME1.TXTE_DATE) < 0 Then
    MsgBox " Ending Date Less than Starting Date ", vbCritical
    VALID_DATE = 0
End If
End Function

Public Function VALID_DATE_CHK(TXT_DATE As Date, mfldname As String) As Boolean
VALID_DATE_CHK = True
If DateDiff("D", PubStartDate, TXT_DATE) < 0 Then
    MsgBox mfldname + " is Before Financial Year ", vbCritical
    VALID_DATE_CHK = False
ElseIf DateDiff("D", TXT_DATE, PubEndDate) < 0 Then
    MsgBox mfldname + " is After Financial Year ", vbCritical
    VALID_DATE_CHK = False
End If
''''''''
End Function

Public Sub SprMrp(FGrid As MSHFlexGrid, mMRevDisTBPer As Double, mMRevDisTPPer As Double, _
                ByVal Col_PNo As Byte, ByVal Col_MRP As Byte, ByVal Col_Taxable As Byte, _
                ByVal Col_Qty As Byte, ByVal Col_Rate As Byte, ByVal Col_MRPRate As Byte, _
                ByVal Col_DiscAmt As Byte, _
                ByVal vDiscPerTB As Single, ByVal vDiscPerTP As Single, _
                ByVal vSTaxPer As Single, ByVal vTaxSurPer As Single, ByVal vTurnOverPer As Single)
Dim I As Integer
'MRP Purpose
Dim mAmount As Double, mDPAmt As Double, mGAmt As Double
Dim mNetTaxPer As Double, mRevTaxAmt As Double
mMRevDisTBPer = 0:  mMRevDisTPPer = 0

    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                mAmount = Val(FGrid.TextMatrix(I, Col_Qty)) * Val(FGrid.TextMatrix(I, Col_MRPRate))
                mAmount = mAmount - Val(FGrid.TextMatrix(I, Col_DiscAmt))
                If pubTOT_On = 0 Then
                   mNetTaxPer = 0
                Else
                    mNetTaxPer = vTurnOverPer
                End If
                
                'Taxable Calculation
                
                If Trim(FGrid.TextMatrix(I, Col_Taxable)) = "Yes" Then
                    mDPAmt = Round(mAmount * vDiscPerTB / 100, 2)
                    mGAmt = mAmount - mDPAmt
                    mNetTaxPer = mNetTaxPer + vSTaxPer + (vSTaxPer * vTaxSurPer / 100)
                    If pubTOT_On = 0 Then
                       mNetTaxPer = mNetTaxPer + vTurnOverPer + (mNetTaxPer * vTurnOverPer / 100)
                    End If
                    mRevTaxAmt = (mGAmt * mNetTaxPer) / (100 + mNetTaxPer)
                    mGAmt = mAmount - mRevTaxAmt
                    If mGAmt <> 0 Then
                        mMRevDisTBPer = vDiscPerTB '(mDPAmt * 100) / mGAmt
                    End If
                Else    'Taxpaid
                    mDPAmt = Round(mAmount * vDiscPerTP / 100, 2)
                    mGAmt = mAmount - mDPAmt
                    mRevTaxAmt = (mGAmt * mNetTaxPer) / (100 + mNetTaxPer)
                    mGAmt = mAmount - mRevTaxAmt
                    If mGAmt <> 0 Then
                        mMRevDisTPPer = (mDPAmt * 100) / mGAmt
                    End If
                End If
                
'                If Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
'                    FGrid.TextMatrix(I, Col_Rate) = Format(mGAmt / Val(FGrid.TextMatrix(I, Col_Qty)), "0.00")
'                End If
            End If
        End If
    Next
End Sub

Public Sub SprCalc(WithLab As ObjTypeDefLab, FGrid As MSHFlexGrid, ByVal mMRevDisTBPer, ByVal mMRevDisTPPer, _
        ByRef mTBDisAmtMRP, ByRef mTPDisAmtMRP, _
        ByVal Col_PNo As Byte, ByVal Col_MRP As Byte, ByVal Col_Taxable As Byte, _
        ByVal Col_Qty As Byte, ByVal Col_Rate As Byte, ByVal Col_ItemVal As Byte, _
        ByVal Col_PartGrade As Byte, ByVal Col_DiscAmt As Byte, _
        TxtIWDiscTotTB As TextBox, TxtIWDiscTotTP As TextBox, _
        TxtMRPAmtTB As TextBox, TxtMRPAmtTP As TextBox, _
        TxtSprAmtTB As TextBox, TxtSprAmtTP As TextBox, _
        TxtOilAmtTB As TextBox, TxtOilAmtTP As TextBox, _
        TxtDiscPerTB As TextBox, TxtDiscPerTP As TextBox, _
        TxtDiscAmtTB As TextBox, TxtDiscAmtTP As TextBox, _
        TxtSTotATB As TextBox, TxtSTotATP As TextBox, _
        TxtGenSurPer As TextBox, TxtGenSurAmt As TextBox, _
        TxtTransAmt As TextBox, TxtTaxableTot As TextBox, _
        TxtSTaxPer As TextBox, TxtSTaxAmt As TextBox, _
        TxtTaxSurPer As TextBox, TxtTaxSurAmt As TextBox, _
        TxtPackCrg As TextBox, TxtSTotB As TextBox, _
        TxtTurnOverPer As TextBox, TxtTurnOverAmt As TextBox, _
        TxtReSalTaxPer As TextBox, TxtReSalTaxAmt As TextBox, _
        TxtSROff As TextBox, TxtNetSprAmt As TextBox, _
        TxtNetAmt As TextBox, mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double, mMRPLubeTB As Double, mMRPLubeTP As Double, _
        Optional Col_Purpose As Byte, Optional JobCall As Boolean)
        ', _
        Optional TxtLabAmt As TextBox, Optional TxtLabDisc As TextBox, _
        Optional TxtServTaxPer As TextBox, Optional TxtServTaxAmt As TextBox, _
        Optional TxtLabRoff As TextBox, Optional TxtNetLabAmt As TextBox, Optional TxtOutSideLabAmt As TextBox)
        
'Used to Calculate Total Values

Dim I As Integer
Dim TotItDiscAmtTB As Double, TotItDiscAmtTP As Double
Dim TotMRPAmtTB As Double, TotMRPAmtTP As Double
Dim TotSprAmtTB As Double, TotSprAmtTP As Double
Dim TotOilAmtTB As Double, TotOilAmtTP As Double, mMRPVAL As Double
'****
Dim mDiscAmtTB As Double, mDiscAmtTP As Double
Dim mSTotATB As Double, mSTotATP As Double, mGenSurAmt As Double
'MRP Purpose
Dim mAmount As Double, mTBTot As Double, mTPTot As Double
Dim mTBTotM As Double, mTPTotM As Double, mTBTotML As Double, mTPTotML As Double
Dim mTBSpr As Double, mTBOil As Double, mTPSpr As Double, mTPOil As Double
Dim mGenSurBasAmt As Double, mTBTot1 As Double, mTPTot1 As Double, mSPRTot As Double
Dim mTBDisAmtMRPLube As Double, mTPDisAmtMRPLube As Double
Dim mTBDisAmtLube As Double, mTPDisAmtLube As Double, mTotalDisLube As Double, mTotalLube As Double

mMRPVAL = 0
mMRPReSales = 0
mMRPLubeTB = 0
mMRPLubeTP = 0
TotItDiscAmtTB = 0
TotItDiscAmtTP = 0
TotMRPAmtTB = 0: TotMRPAmtTP = 0
TotSprAmtTB = 0: TotSprAmtTP = 0
TotOilAmtTB = 0: TotOilAmtTP = 0

mDiscAmtTB = 0:   mDiscAmtTP = 0
mTBDisAmtMRP = 0:   mTPDisAmtMRP = 0
mTBDisAmtMRPLube = 0:   mTPDisAmtMRPLube = 0
mTBDisAmtLube = 0:   mTPDisAmtLube = 0
mSTotATB = 0:   mSTotATP = 0:  mGenSurAmt = 0

mAmount = 0: mTBTot = 0: mTPTot = 0
mTBTotM = 0: mTPTotM = 0: mTBTotML = 0: mTPTotML = 0
mTBSpr = 0: mTBOil = 0: mTPSpr = 0: mTPOil = 0
mGenSurBasAmt = 0: mTBTot1 = 0: mTPTot1 = 0: mSPRTot = 0

    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" And _
            (JobCall = False Or (JobCall And FGrid.TextMatrix(I, Col_Purpose) = "Charge")) Then
            If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                'Qty * RevRate
                mAmount = Round(Val(FGrid.TextMatrix(I, Col_Qty)) * Val(FGrid.TextMatrix(I, Col_Rate)), 2)
            Else
                mAmount = Val(FGrid.TextMatrix(I, Col_ItemVal)) 'Less Disc
            End If
            If Trim(FGrid.TextMatrix(I, Col_Taxable)) = "Yes" Then
                If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                    If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                        mTBTotML = mTBTotML + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                        mMRPLubeTB = mMRPLubeTB + Val(FGrid.TextMatrix(I, Col_ItemVal))
                    Else
                        mTBTotM = mTBTotM + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                    End If
                    TotMRPAmtTB = TotMRPAmtTB + Val(FGrid.TextMatrix(I, Col_ItemVal))
                Else
                    If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                        TotOilAmtTB = TotOilAmtTB + mAmount
                    Else
                        TotSprAmtTB = TotSprAmtTB + mAmount
                    End If
                End If
                mTBOil = mTBTotML + TotOilAmtTB ' mTBOil + mAmount
                mTBSpr = mTBTotM + TotSprAmtTB  ' mTBSpr + mAmount
                TotItDiscAmtTB = TotItDiscAmtTB + Val(FGrid.TextMatrix(I, Col_DiscAmt))
            Else
                If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                    If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                        mTPTotML = mTPTotML + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                        mMRPLubeTP = mMRPLubeTP + Val(FGrid.TextMatrix(I, Col_ItemVal))
                    Else
                        mTPTotM = mTPTotM + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                    End If
                    TotMRPAmtTP = TotMRPAmtTP + Val(FGrid.TextMatrix(I, Col_ItemVal))
                Else
                    If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                        TotOilAmtTP = TotOilAmtTP + mAmount
                    Else
                        TotSprAmtTP = TotSprAmtTP + mAmount
                    End If
                End If
                mTPOil = mTPTotML + TotOilAmtTP    'mTPOil + mAmount
                mTPSpr = mTPTotM + TotSprAmtTP  'mTPSpr + mAmount
                TotItDiscAmtTP = TotItDiscAmtTP + Val(FGrid.TextMatrix(I, Col_DiscAmt))
            End If
        End If
    Next
    mTBTot = mTBSpr + mTBOil    'Includes MRP Rev Amt
    mTPTot = mTPSpr + mTPOil    'Includes MRP Rev Amt
    
    TxtIWDiscTotTB = IIf(TotItDiscAmtTB <> 0, Format(TotItDiscAmtTB, "0.00"), "")
    TxtIWDiscTotTP = IIf(TotItDiscAmtTP <> 0, Format(TotItDiscAmtTP, "0.00"), "")
    
    TxtMRPAmtTB = IIf(TotMRPAmtTB <> 0, Format(TotMRPAmtTB, "0.00"), "")
    TxtMRPAmtTP = IIf(TotMRPAmtTP <> 0, Format(TotMRPAmtTP, "0.00"), "")

    TxtSprAmtTB = IIf(TotSprAmtTB <> 0, Format(TotSprAmtTB, "0.00"), "")
    TxtSprAmtTP = IIf(TotSprAmtTP <> 0, Format(TotSprAmtTP, "0.00"), "")

    TxtOilAmtTB = IIf(TotOilAmtTB <> 0, Format(TotOilAmtTB, "0.00"), "")
    TxtOilAmtTP = IIf(TotOilAmtTP <> 0, Format(TotOilAmtTP, "0.00"), "")
    'Apply conditional Disc on Lube
    mTBDisAmtMRPLube = 0
    mTPDisAmtMRPLube = 0
    mTBDisAmtLube = 0
    mTPDisAmtLube = 0
    If PubDiscOnLube = 1 Then   'Yes
        mTBDisAmtMRP = Round((mTBTotM + mTBTotML) * mMRevDisTBPer / 100, 2)
        mTPDisAmtMRP = Round((mTPTotM + mTPTotML) * mMRevDisTPPer / 100, 2)
        mTBDisAmtMRPLube = Round(mTBTotML * mMRevDisTBPer / 100, 2)
        mTPDisAmtMRPLube = Round(mTPTotML * mMRevDisTPPer / 100, 2)
        If Val(TxtDiscPerTB) <> 0 Then
            mDiscAmtTB = Round(mTBDisAmtMRP + (TotSprAmtTB + TotOilAmtTB) * Val(TxtDiscPerTB) / 100, 2)
            TxtDiscAmtTB = IIf(mDiscAmtTB <> 0, Format(mDiscAmtTB, "0.00"), "")
            mTBDisAmtLube = Round(TotOilAmtTB * Val(TxtDiscPerTB) / 100, 2)
        ElseIf Val(TxtDiscPerTB.Tag) = 0 Then
            TxtDiscAmtTB = ""
        End If
        If Val(TxtDiscPerTP) <> 0 Then
            mDiscAmtTP = Round(mTPDisAmtMRP + (TotSprAmtTP + TotOilAmtTP) * Val(TxtDiscPerTP) / 100, 2)
            TxtDiscAmtTP = IIf(mDiscAmtTP <> 0, Format(mDiscAmtTP, "0.00"), "")
            mTPDisAmtLube = Round(TotOilAmtTP * Val(TxtDiscPerTP) / 100, 2)
        ElseIf Val(TxtDiscPerTP.Tag) = 0 Then
            TxtDiscAmtTP = ""
        End If
    Else
        mTBDisAmtMRP = Round((mTBTotM) * mMRevDisTBPer / 100, 2)
        mTPDisAmtMRP = Round((mTPTotM) * mMRevDisTPPer / 100, 2)
        If Val(TxtDiscPerTB) <> 0 Then
            mDiscAmtTB = Round(mTBDisAmtMRP + (TotSprAmtTB) * Val(TxtDiscPerTB) / 100, 2)
            TxtDiscAmtTB = IIf(mDiscAmtTB <> 0, Format(mDiscAmtTB, "0.00"), "")
        ElseIf Val(TxtDiscPerTB.Tag) = 0 Then
            TxtDiscAmtTB = ""
        End If
        If Val(TxtDiscPerTP) <> 0 Then
            mDiscAmtTP = Round(mTPDisAmtMRP + (TotSprAmtTP) * Val(TxtDiscPerTP) / 100, 2)
            TxtDiscAmtTP = IIf(mDiscAmtTP <> 0, Format(mDiscAmtTP, "0.00"), "")
        ElseIf Val(TxtDiscPerTP.Tag) = 0 Then
            TxtDiscAmtTP = ""
        End If
    End If
    mSTotATB = Val(TxtMRPAmtTB) + Val(TxtSprAmtTB) + Val(TxtOilAmtTB) - Val(TxtDiscAmtTB)
    mSTotATP = Val(TxtMRPAmtTP) + Val(TxtSprAmtTP) + Val(TxtOilAmtTP) - Val(TxtDiscAmtTP)
    mMRPVAL = Val(TxtMRPAmtTB) + Val(TxtMRPAmtTP) - (mTPDisAmtMRP + mTBDisAmtMRP)
    TxtSTotATB = IIf(mSTotATB <> 0, Format(mSTotATB, "0.00"), "")
    TxtSTotATP = IIf(mSTotATP <> 0, Format(mSTotATP, "0.00"), "")
'   check values of mTBTot1 & Txt(STotAB)
    mTBTot1 = mTBTot - mDiscAmtTB
    mGenSurBasAmt = Round((mTBTot1 - mTBTotM - mTBTotML + mTBDisAmtMRP), 2)
    If Val(TxtGenSurPer) <> 0 Then
        mGenSurAmt = (mGenSurBasAmt * Val(TxtGenSurPer)) / 100
        TxtGenSurAmt = IIf(mGenSurAmt <> 0, Format(mGenSurAmt, "0.00"), "")
    ElseIf Val(TxtGenSurPer.Tag) <> 0 Then
        TxtGenSurAmt = ""
    End If
    If Val(TxtSTotATB) + Val(TxtGenSurAmt) + Val(TxtTransAmt) <> 0 Then
        TxtTaxableTot = Format(Val(TxtSTotATB) + Val(TxtGenSurAmt) + Val(TxtTransAmt), "0.00")
    Else
        TxtTaxableTot = ""
    End If
   
    If Val(TxtTaxableTot) <> 0 Then
        TxtSTaxAmt = Format((mGenSurBasAmt + Val(TxtGenSurAmt) + Val(TxtTransAmt)) * Val(TxtSTaxPer) / 100, "0.00")
    Else
        TxtSTaxAmt = ""
    End If
    If Val(TxtSTaxAmt) <> 0 Then
        TxtTaxSurAmt = Format((Val(TxtSTaxAmt) * Val(TxtTaxSurPer)) / 100, "0.00")
    Else
        TxtTaxSurAmt = ""
    End If
    
    'check values of mSprTot & Txt(STotB)
    mSPRTot = mTBTot1 + mTPTot1 + Val(TxtGenSurAmt) + Val(TxtTransAmt) + Val(TxtSTaxAmt) + Val(TxtTaxSurAmt) + Val(TxtPackCrg)
    TxtSTotB = Format(Val(TxtSTotATP) + Val(TxtTaxableTot) + Val(TxtSTaxAmt) + Val(TxtTaxSurAmt) + Val(TxtPackCrg), "0.00")
    
'   PubTOT_On   0-Sub Total (B) TB+TP, 1-Taxable+Taxpaid Total
 If UCase(left(PubComp_Name, 3)) <> "JMK" Then
    If Val(TxtTurnOverPer) <> 0 Then
        If PubSDTYN = 1 Then
            If PubTOTOnLube = 1 Then
               TxtTurnOverAmt = Format((mGenSurBasAmt + Val(TxtGenSurAmt) + Val(TxtTransAmt)) * Val(TxtTurnOverPer) / 100, "0.00")
            Else
                mTotalDisLube = mTBDisAmtLube
                mTotalLube = Val(TxtOilAmtTB) - mTotalDisLube
                TxtTurnOverAmt = Format(((mGenSurBasAmt + Val(TxtGenSurAmt) + Val(TxtTransAmt)) - mTotalLube) * Val(TxtTurnOverPer) / 100, "0.00")
            End If
        Else
            If PubTOTOnLube = 1 Then
                If pubTOT_On = 0 Then   '0-Sub Total (B)
                    TxtTurnOverAmt = Format((Val(TxtSTotB) - mMRPVAL) * Val(TxtTurnOverPer) / 100, "0.00")
                Else    '1-Taxable+Taxpaid Total
                    TxtTurnOverAmt = Format((Val(TxtSTotATP) + Val(TxtTaxableTot) - mMRPVAL) * Val(TxtTurnOverPer) / 100, "0.00")
                End If
            Else
                mTotalDisLube = mTBDisAmtLube + mTPDisAmtLube
                mTotalLube = (Val(TxtOilAmtTB) + Val(TxtOilAmtTP)) - mTotalDisLube
                If pubTOT_On = 0 Then   '0-Sub Total (B)
                    TxtTurnOverAmt = Format((Val(TxtSTotB) - mMRPVAL - mTotalLube) * Val(TxtTurnOverPer) / 100, "0.00")
                Else    '1-Taxable+Taxpaid Total
                    TxtTurnOverAmt = Format((Val(TxtSTotATP) + Val(TxtTaxableTot) - mMRPVAL - mTotalLube) * Val(TxtTurnOverPer) / 100, "0.00")
                End If
            End If
        End If
    ElseIf Val(TxtTurnOverPer.Tag) <> 0 Then
        TxtTurnOverAmt = ""
    End If
 End If
    
    'forward
    If Val(TxtReSalTaxPer) <> 0 Then
        TxtReSalTaxAmt = Format(Val(TxtSTotB) * Val(TxtReSalTaxPer) / 100, "0.00")
    ElseIf Val(TxtReSalTaxPer.Tag) <> 0 Then
        TxtReSalTaxAmt = ""
    End If
    TxtSROff = dmRoundOff(Val(TxtSTotB) + Val(TxtTurnOverAmt) + Val(TxtReSalTaxAmt))
    TxtNetSprAmt = Format(Val(TxtSTotB) + Val(TxtTurnOverAmt) + Val(TxtReSalTaxAmt) + Val(TxtSROff), "0.00")
    'MRP Tax Calculation
    'Apply conditional Tax on Lube
    Dim rRTax As Double, gTax As Double
    rRTax = Val(TxtSTaxPer) + Format((Val(TxtSTaxPer) * Val(TxtTaxSurPer) / 100), "0.00")
    If PubSDTYN = 1 Then
        If UCase(left(PubComp_Name, 3)) <> "JMK" Then
            rRTax = rRTax + Val(TxtTurnOverPer)
        End If
    Else
     If pubTOT_On = 0 Then
        rRTax = rRTax + Val(TxtTurnOverPer) + (rRTax * Val(TxtTurnOverPer) / 100)
     Else
        rRTax = rRTax + Val(TxtTurnOverPer)
     End If
    End If
    If PubTOTOnLube = 1 Then
       'mMRPTax = Round((mTBTotM + mTBTotML - mTBDisAmtMRP) * Val(TxtSTaxPer) / 100, 2)
        gTax = Round((TotMRPAmtTB + mTBTotML - mTBDisAmtMRP) * rRTax / (100 + rRTax), 2)
        mMRPTax = Round((TotMRPAmtTB + mTBTotML - (mTBDisAmtMRP + gTax)) * Val(TxtSTaxPer) / 100, 2)
        mMRPTaxSur = Round(mMRPTax * Val(TxtTaxSurPer) / 100, 2)
        If PubSDTYN = 1 Then
              mMRPTOT = Round((TotMRPAmtTB + mTBTotML - (mTBDisAmtMRP + gTax)) * Val(TxtTurnOverPer) / 100, 2)
        Else
            If pubTOT_On = 0 Then   '0-Sub Total (B)
                mMRPTOT = Round((mTBTotM + mTBTotML + mTPTotM + mTPTotML - mTBDisAmtMRP - mTPDisAmtMRP + mMRPTax + mMRPTaxSur) * Val(TxtTurnOverPer) / 100, 2)
            Else    '1-Taxable+Taxpaid Total
                mMRPTOT = Round((mTBTotM + mTBTotML + mTPTotM + mTPTotML - mTBDisAmtMRP - mTPDisAmtMRP) * Val(TxtTurnOverPer) / 100, 2)
            End If
        End If
    Else
        mTotalDisLube = mTBDisAmtMRPLube + mTPDisAmtMRPLube
        mTotalLube = (mTBTotML + mTPTotML) - mTotalDisLube
        '--
        'mMRPTax = Round((mTBTotM - mTBDisAmtMRP) * Val(TxtSTaxPer) / 100, 2)
        gTax = Round((TotMRPAmtTB - mTBDisAmtMRP) * rRTax / (100 + rRTax), 2)
        mMRPTax = Round((TotMRPAmtTB - (mTBDisAmtMRP + gTax)) * Val(TxtSTaxPer) / 100, 2)
        
'        mMRPTax = Round((TotMRPAmtTB - mTBDisAmtMRP) * Val(TxtSTaxPer) / (100 + Val(TxtSTaxPer)), 2)
        mMRPTaxSur = Round(mMRPTax * Val(TxtTaxSurPer) / 100, 2)
        If PubSDTYN = 1 Then
              mMRPTOT = Round((TotMRPAmtTB - (mTBDisAmtMRP + gTax)) * Val(TxtTurnOverPer) / 100, 2)
        Else
            If pubTOT_On = 0 Then   '0-Sub Total (B)
                mMRPTOT = Round((mTBTotM + mTPTotM - mTotalLube - mTBDisAmtMRP - mTPDisAmtMRP + mMRPTax + mMRPTaxSur) * Val(TxtTurnOverPer) / 100, 2)
            Else    '1-Taxable+Taxpaid Total
                mMRPTOT = Round((mTBTotM + mTPTotM - mTotalLube - mTBDisAmtMRP - mTPDisAmtMRP) * Val(TxtTurnOverPer) / 100, 2)
            End If
        End If
    End If
    
    'Enable / Disable Text Box if values zero
    DisableEnableFooter TxtMRPAmtTB, TxtMRPAmtTP, TxtSprAmtTB, TxtSprAmtTP, _
        TxtOilAmtTB, TxtOilAmtTP, TxtDiscPerTB, TxtDiscPerTP, _
        TxtDiscAmtTB, TxtDiscAmtTP, TxtSTotATB, TxtGenSurPer, TxtGenSurAmt, _
        TxtTaxableTot, TxtSTaxPer, TxtSTaxAmt, TxtTaxSurPer, TxtTaxSurAmt
    'EOF enable / disable section
    TxtNetAmt = Format(Val(TxtNetSprAmt), "0.00")
End Sub

Public Sub SprCalcVAT(WithLab As ObjTypeDefLab, FGrid As MSHFlexGrid, ByVal mMRevDisTBPer, ByVal mMRevDisTPPer, _
        ByRef mTBDisAmtMRP, ByRef mTPDisAmtMRP, _
        ByVal Col_PNo As Byte, ByVal Col_MRP As Byte, ByVal Col_Taxable As Byte, _
        ByVal Col_Qty As Byte, ByVal Col_Rate As Byte, ByVal Col_ItemVal As Byte, _
        ByVal Col_PartGrade As Byte, ByVal Col_DiscAmt As Byte, ByVal Col_TaxPer As Byte, ByVal Col_TaxAmt1 As Byte, _
        TxtIWDiscTotTB As TextBox, TxtIWDiscTotTP As TextBox, _
        TxtMRPAmtTB As TextBox, TxtMRPAmtTP As TextBox, _
        TxtSprAmtTB As TextBox, TxtSprAmtTP As TextBox, _
        TxtOilAmtTB As TextBox, TxtOilAmtTP As TextBox, _
        TxtDiscPerTB As TextBox, TxtDiscPerTP As TextBox, _
        TxtDiscAmtTB As TextBox, TxtDiscAmtTP As TextBox, _
        TxtSTotATB As TextBox, TxtSTotATP As TextBox, _
        TxtGenSurPer As TextBox, TxtGenSurAmt As TextBox, _
        TxtTransAmt As TextBox, TxtTaxableTot As TextBox, _
        TxtSTaxPer As TextBox, TxtSTaxAmt As TextBox, _
        TxtTaxSurPer As TextBox, TxtTaxSurAmt As TextBox, _
        TxtPackCrg As TextBox, TxtSTotB As TextBox, _
        TxtTurnOverPer As TextBox, TxtTurnOverAmt As TextBox, _
        TxtReSalTaxPer As TextBox, TxtReSalTaxAmt As TextBox, _
        TxtSROff As TextBox, TxtNetSprAmt As TextBox, _
        TxtNetAmt As TextBox, mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double, mMRPLubeTB As Double, mMRPLubeTP As Double, Col_SatPer As Byte, Col_SatAmt As Byte, TxtSatAmt As TextBox, _
        Optional Col_Purpose As Byte, Optional JobCall As Boolean)
        ', _
        Optional TxtLabAmt As TextBox, Optional TxtLabDisc As TextBox, _
        Optional TxtServTaxPer As TextBox, Optional TxtServTaxAmt As TextBox, _
        Optional TxtLabRoff As TextBox, Optional TxtNetLabAmt As TextBox, Optional TxtOutSideLabAmt As TextBox)
        
'Used to Calculate Total Values

Dim I As Integer
Dim TotItDiscAmtTB As Double, TotItDiscAmtTP As Double
Dim TotMRPAmtTB As Double, TotMRPAmtTP As Double
Dim TotSprAmtTB As Double, TotSprAmtTP As Double
Dim TotOilAmtTB As Double, TotOilAmtTP As Double, mMRPVAL As Double
'****
Dim mDiscAmtTB As Double, mDiscAmtTP As Double
Dim mSTotATB As Double, mSTotATP As Double, mGenSurAmt As Double
'MRP Purpose
Dim mAmount As Double, mTBTot As Double, mTPTot As Double
Dim mTBTotM As Double, mTPTotM As Double, mTBTotML As Double, mTPTotML As Double
Dim mTBSpr As Double, mTBOil As Double, mTPSpr As Double, mTPOil As Double
Dim mGenSurBasAmt As Double, mTBTot1 As Double, mTPTot1 As Double, mSPRTot As Double
Dim mTBDisAmtMRPLube As Double, mTPDisAmtMRPLube As Double
Dim mTBDisAmtLube As Double, mTPDisAmtLube As Double, mTotalDisLube As Double, mTotalLube As Double

mMRPVAL = 0
mMRPReSales = 0
mMRPLubeTB = 0
mMRPLubeTP = 0
TotItDiscAmtTB = 0
TotItDiscAmtTP = 0
TotMRPAmtTB = 0: TotMRPAmtTP = 0
TotSprAmtTB = 0: TotSprAmtTP = 0
TotOilAmtTB = 0: TotOilAmtTP = 0

mDiscAmtTB = 0:   mDiscAmtTP = 0

If Not PubSiebelActiveYn = 1 Or mMRevDisTBPer > 0 Then mTBDisAmtMRP = 0
If Not PubSiebelActiveYn = 1 Or mMRevDisTPPer > 0 Then mTPDisAmtMRP = 0

mTBDisAmtMRPLube = 0:   mTPDisAmtMRPLube = 0
mTBDisAmtLube = 0:   mTPDisAmtLube = 0
mSTotATB = 0:   mSTotATP = 0:  mGenSurAmt = 0

mAmount = 0: mTBTot = 0: mTPTot = 0
mTBTotM = 0: mTPTotM = 0: mTBTotML = 0: mTPTotML = 0
mTBSpr = 0: mTBOil = 0: mTPSpr = 0: mTPOil = 0
mGenSurBasAmt = 0: mTBTot1 = 0: mTPTot1 = 0: mSPRTot = 0
TxtSTaxAmt = "": TxtSatAmt = 0

    
    For I = 1 To FGrid.Rows - 1
        TxtSTaxPer = FGrid.TextMatrix(I, Col_TaxPer)
        
        If StrCmp(left(PubComp_Name, 4), "Enar") Then
            If FGrid.TextMatrix(I, Col_PNo) <> "" And _
                (JobCall = False Or (JobCall And (FGrid.TextMatrix(I, Col_Purpose) = "Charge" Or FGrid.TextMatrix(I, Col_Purpose) = "AMC"))) Then
                If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                    mAmount = Val(FGrid.TextMatrix(I, Col_ItemVal))
                Else
                    mAmount = Val(FGrid.TextMatrix(I, Col_ItemVal)) 'Less Disc
                End If
                If Trim(FGrid.TextMatrix(I, Col_Taxable)) = "Yes" Then
                    If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            mTBTotML = mTBTotML + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                            mMRPLubeTB = mMRPLubeTB + Val(FGrid.TextMatrix(I, Col_ItemVal))
                        Else
                            mTBTotM = mTBTotM + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                        End If
                        TotMRPAmtTB = TotMRPAmtTB + Val(FGrid.TextMatrix(I, Col_ItemVal))
                    Else
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            TotOilAmtTB = TotOilAmtTB + mAmount
                        Else
                            TotSprAmtTB = TotSprAmtTB + mAmount
                        End If
                    End If
                    mTBOil = mTBTotML + TotOilAmtTB ' mTBOil + mAmount
                    mTBSpr = mTBTotM + TotSprAmtTB  ' mTBSpr + mAmount
                    TotItDiscAmtTB = TotItDiscAmtTB + Val(FGrid.TextMatrix(I, Col_DiscAmt))
                Else
                    If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            mTPTotML = mTPTotML + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                            mMRPLubeTP = mMRPLubeTP + Val(FGrid.TextMatrix(I, Col_ItemVal))
                        Else
                            mTPTotM = mTPTotM + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                        End If
                        TotMRPAmtTP = TotMRPAmtTP + Val(FGrid.TextMatrix(I, Col_ItemVal))
                    Else
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            TotOilAmtTP = TotOilAmtTP + mAmount
                        Else
                            TotSprAmtTP = TotSprAmtTP + mAmount
                        End If
                    End If
                    mTPOil = mTPTotML + TotOilAmtTP    'mTPOil + mAmount
                    mTPSpr = mTPTotM + TotSprAmtTP  'mTPSpr + mAmount
                    TotItDiscAmtTP = TotItDiscAmtTP + Val(FGrid.TextMatrix(I, Col_DiscAmt))
                End If
            End If
        Else
            If FGrid.TextMatrix(I, Col_PNo) <> "" And _
                (JobCall = False Or (JobCall And FGrid.TextMatrix(I, Col_Purpose) = "Charge")) Then
                If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                    mAmount = Val(FGrid.TextMatrix(I, Col_ItemVal))
                Else
                    mAmount = Val(FGrid.TextMatrix(I, Col_ItemVal)) 'Less Disc
                End If
                If Trim(FGrid.TextMatrix(I, Col_Taxable)) = "Yes" Then
                    If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            mTBTotML = mTBTotML + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                            mMRPLubeTB = mMRPLubeTB + Val(FGrid.TextMatrix(I, Col_ItemVal))
                        Else
                            mTBTotM = mTBTotM + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                        End If
                        TotMRPAmtTB = TotMRPAmtTB + Val(FGrid.TextMatrix(I, Col_ItemVal))
                    Else
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            TotOilAmtTB = TotOilAmtTB + mAmount
                        Else
                            TotSprAmtTB = TotSprAmtTB + mAmount
                        End If
                    End If
                    mTBOil = mTBTotML + TotOilAmtTB ' mTBOil + mAmount
                    mTBSpr = mTBTotM + TotSprAmtTB  ' mTBSpr + mAmount
                    TotItDiscAmtTB = TotItDiscAmtTB + Val(FGrid.TextMatrix(I, Col_DiscAmt))
                Else
                    If Trim(FGrid.TextMatrix(I, Col_MRP)) = "Yes" Then
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            mTPTotML = mTPTotML + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                            mMRPLubeTP = mMRPLubeTP + Val(FGrid.TextMatrix(I, Col_ItemVal))
                        Else
                            mTPTotM = mTPTotM + mAmount 'Qty * RevRate (less forward tax,surcharge,disc)
                        End If
                        TotMRPAmtTP = TotMRPAmtTP + Val(FGrid.TextMatrix(I, Col_ItemVal))
                    Else
                        If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                            TotOilAmtTP = TotOilAmtTP + mAmount
                        Else
                            TotSprAmtTP = TotSprAmtTP + mAmount
                        End If
                    End If
                    mTPOil = mTPTotML + TotOilAmtTP    'mTPOil + mAmount
                    mTPSpr = mTPTotM + TotSprAmtTP  'mTPSpr + mAmount
                    TotItDiscAmtTP = TotItDiscAmtTP + Val(FGrid.TextMatrix(I, Col_DiscAmt))
                End If
            End If
        End If
    TxtSTaxAmt = Format(Val(TxtSTaxAmt) + Val(FGrid.TextMatrix(I, Col_TaxAmt1)), "0.00")
    TxtSatAmt = Format(Val(TxtSatAmt) + Val(FGrid.TextMatrix(I, Col_SatAmt)), "0.00")
    Next
    mTBTot = mTBSpr + mTBOil    'Includes MRP Rev Amt
    mTPTot = mTPSpr + mTPOil    'Includes MRP Rev Amt
    
    TxtIWDiscTotTB = IIf(TotItDiscAmtTB <> 0, Format(TotItDiscAmtTB, "0.00"), "")
    TxtIWDiscTotTP = IIf(TotItDiscAmtTP <> 0, Format(TotItDiscAmtTP, "0.00"), "")
    
    TxtMRPAmtTB = IIf(TotMRPAmtTB <> 0, Format(TotMRPAmtTB, "0.00"), "")
    TxtMRPAmtTP = IIf(TotMRPAmtTP <> 0, Format(TotMRPAmtTP, "0.00"), "")

    TxtSprAmtTB = IIf(TotSprAmtTB <> 0, Format(TotSprAmtTB, "0.00"), "")
    TxtSprAmtTP = IIf(TotSprAmtTP <> 0, Format(TotSprAmtTP, "0.00"), "")

    TxtOilAmtTB = IIf(TotOilAmtTB <> 0, Format(TotOilAmtTB, "0.00"), "")
    TxtOilAmtTP = IIf(TotOilAmtTP <> 0, Format(TotOilAmtTP, "0.00"), "")
    'Apply conditional Disc on Lube
    
    mTBDisAmtMRPLube = 0
    mTPDisAmtMRPLube = 0
    mTBDisAmtLube = 0
    mTPDisAmtLube = 0
    If PubDiscOnLube = 1 Then
        If Not PubSiebelActiveYn = 1 Or mMRevDisTBPer > 0 Then
            mTBDisAmtMRP = Round((mTBTotM + mTBTotML) * mMRevDisTBPer / 100, 2)
        End If
        If Not PubSiebelActiveYn = 1 Or mMRevDisTPPer > 0 Then
            mTPDisAmtMRP = Round((mTPTotM + mTPTotML) * mMRevDisTPPer / 100, 2)
        End If
        mTBDisAmtMRPLube = Round(mTBTotML * mMRevDisTBPer / 100, 2)
        mTPDisAmtMRPLube = Round(mTPTotML * mMRevDisTPPer / 100, 2)
        
        If Val(TxtDiscPerTB) <> 0 Then
            mDiscAmtTB = Round(mTBDisAmtMRP + (TotSprAmtTB + TotOilAmtTB) * Val(TxtDiscPerTB) / 100, 2)
            TxtDiscAmtTB = IIf(mDiscAmtTB <> 0, Format(mDiscAmtTB, "0.00"), "")
            mTBDisAmtLube = Round(TotOilAmtTB * Val(TxtDiscPerTB) / 100, 2)
        ElseIf Val(TxtDiscPerTB.Tag) = 0 Then
            If Not PubSiebelActiveYn = 1 Then
                TxtDiscAmtTB = ""
            End If
        End If
        If Val(TxtDiscPerTP) <> 0 Then
            mDiscAmtTP = Round(mTPDisAmtMRP + (TotSprAmtTP + TotOilAmtTP) * Val(TxtDiscPerTP) / 100, 2)
            TxtDiscAmtTP = IIf(mDiscAmtTP <> 0, Format(mDiscAmtTP, "0.00"), "")
            mTPDisAmtLube = Round(TotOilAmtTP * Val(TxtDiscPerTP) / 100, 2)
        ElseIf Val(TxtDiscPerTP.Tag) = 0 Then
            TxtDiscAmtTP = ""
        End If
    Else
        If Not PubSiebelActiveYn = 1 Or mMRevDisTBPer > 0 Then
            mTBDisAmtMRP = Round((mTBTotM) * mMRevDisTBPer / 100, 2)
        End If
        If Not PubSiebelActiveYn = 1 Or mMRevDisTPPer > 0 Then
            mTPDisAmtMRP = Round((mTPTotM) * mMRevDisTPPer / 100, 2)
        End If
        If Val(TxtDiscPerTB) <> 0 Then
            mDiscAmtTB = Round(mTBDisAmtMRP + (TotSprAmtTB) * Val(TxtDiscPerTB) / 100, 2)
            TxtDiscAmtTB = IIf(mDiscAmtTB <> 0, Format(mDiscAmtTB, "0.00"), "")
        ElseIf Val(TxtDiscPerTB.Tag) = 0 Then
            If Not PubSiebelActiveYn = 1 Then
                TxtDiscAmtTB = ""
            End If
        End If
        If Val(TxtDiscPerTP) <> 0 Then
            mDiscAmtTP = Round(mTPDisAmtMRP + (TotSprAmtTP) * Val(TxtDiscPerTP) / 100, 2)
            TxtDiscAmtTP = IIf(mDiscAmtTP <> 0, Format(mDiscAmtTP, "0.00"), "")
        ElseIf Val(TxtDiscPerTP.Tag) = 0 Then
            TxtDiscAmtTP = ""
        End If
    End If
    mSTotATB = Val(TxtMRPAmtTB) + Val(TxtSprAmtTB) + Val(TxtOilAmtTB) - Val(TxtDiscAmtTB)
    mSTotATP = Val(TxtMRPAmtTP) + Val(TxtSprAmtTP) + Val(TxtOilAmtTP) - Val(TxtDiscAmtTP)
    mMRPVAL = Val(TxtMRPAmtTB) + Val(TxtMRPAmtTP) - (mTPDisAmtMRP + mTBDisAmtMRP)
    TxtSTotATB = IIf(mSTotATB <> 0, Format(mSTotATB, "0.00"), "")
    TxtSTotATP = IIf(mSTotATP <> 0, Format(mSTotATP, "0.00"), "")
'   check values of mTBTot1 & Txt(STotAB)
    mTBTot1 = mTBTot - mDiscAmtTB
    mGenSurBasAmt = Round((mTBTot1 - mTBTotM - mTBTotML + mTBDisAmtMRP), 2)
    If Val(TxtGenSurPer) <> 0 Then
        mGenSurAmt = (mGenSurBasAmt * Val(TxtGenSurPer)) / 100
        TxtGenSurAmt = IIf(mGenSurAmt <> 0, Format(mGenSurAmt, "0.00"), "")
    ElseIf Val(TxtGenSurPer.Tag) <> 0 Then
        TxtGenSurAmt = ""
    End If
    If Val(TxtSTotATB) + Val(TxtGenSurAmt) + Val(TxtTransAmt) <> 0 Then
        TxtTaxableTot = Format(Val(TxtSTotATB) + Val(TxtGenSurAmt) + Val(TxtTransAmt), "0.00")
    Else
        TxtTaxableTot = ""
    End If
   
   ' If Val(TxtTaxableTot) <> 0 Then
    '    TxtSTaxAmt = Format((mGenSurBasAmt + Val(TxtGenSurAmt) + Val(TxtTransAmt)) * Val(TxtSTaxPer) / 100, "0.00")
    'Else
     '   TxtSTaxAmt = ""
    'End If
'    If Val(TxtSTaxAmt) <> 0 Then
'        TxtTaxSurAmt = Format((Val(TxtSTaxAmt) * Val(TxtTaxSurPer)) / 100, "0.00")
'    Else
'        TxtTaxSurAmt = ""
'    End If
    
    'check values of mSprTot & Txt(STotB)
    mSPRTot = mTBTot1 + mTPTot1 + Val(TxtGenSurAmt) + Val(TxtTransAmt) + Val(TxtSTaxAmt) + Val(TxtSatAmt) + Val(TxtTaxSurAmt) + Val(TxtPackCrg)
    TxtSTotB = Format(Val(TxtSTotATP) + Val(TxtTaxableTot) + Val(TxtSTaxAmt) + Val(TxtSatAmt) + Val(TxtTaxSurAmt) + Val(TxtPackCrg), "0.00")
    
'   PubTOT_On   0-Sub Total (B) TB+TP, 1-Taxable+Taxpaid Total
If UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
TxtTurnOverAmt = Format(Val(TxtPackCrg.TEXT) * Val(TxtTurnOverPer) / 100, "0.00")
Else
    If Val(TxtTurnOverPer) <> 0 Then
        If PubSDTYN = 1 Then
            If PubTOTOnLube = 1 Then
                TxtTurnOverAmt = Format((mGenSurBasAmt + Val(TxtGenSurAmt) + Val(TxtTransAmt)) * Val(TxtTurnOverPer) / 100, "0.00")
            Else
                mTotalDisLube = mTBDisAmtLube
                mTotalLube = Val(TxtOilAmtTB) - mTotalDisLube
                TxtTurnOverAmt = Format(((mGenSurBasAmt + Val(TxtGenSurAmt) + Val(TxtTransAmt)) - mTotalLube) * Val(TxtTurnOverPer) / 100, "0.00")
            End If
        Else
            If PubTOTOnLube = 1 Then
                If pubTOT_On = 0 Then   '0-Sub Total (B)
                    TxtTurnOverAmt = Format((Val(TxtSTotB) - mMRPVAL) * Val(TxtTurnOverPer) / 100, "0.00")
                Else    '1-Taxable+Taxpaid Total
                    TxtTurnOverAmt = Format((Val(TxtSTotATP) + Val(TxtTaxableTot) - mMRPVAL) * Val(TxtTurnOverPer) / 100, "0.00")
                End If
            Else
                mTotalDisLube = mTBDisAmtLube + mTPDisAmtLube
                mTotalLube = (Val(TxtOilAmtTB) + Val(TxtOilAmtTP)) - mTotalDisLube
                If pubTOT_On = 0 Then   '0-Sub Total (B)
                    TxtTurnOverAmt = Format((Val(TxtSTotB) - mMRPVAL - mTotalLube) * Val(TxtTurnOverPer) / 100, "0.00")
                Else    '1-Taxable+Taxpaid Total
                    TxtTurnOverAmt = Format((Val(TxtSTotATP) + Val(TxtTaxableTot) - mMRPVAL - mTotalLube) * Val(TxtTurnOverPer) / 100, "0.00")
                End If
            End If
        End If
    ElseIf Val(TxtTurnOverPer.Tag) <> 0 Then
        TxtTurnOverAmt = ""
    End If
End If
    'forward
    If Val(TxtReSalTaxPer) <> 0 Then
        TxtReSalTaxAmt = Format(Val(TxtSTotB) * Val(TxtReSalTaxPer) / 100, "0.00")
    ElseIf Val(TxtReSalTaxPer.Tag) <> 0 Then
        TxtReSalTaxAmt = ""
    End If
    TxtSROff = dmRoundOff(Val(TxtSTotB) + Val(TxtTurnOverAmt) + Val(TxtReSalTaxAmt))
    TxtNetSprAmt = Format(Val(TxtSTotB) + Val(TxtTurnOverAmt) + Val(TxtReSalTaxAmt) + Val(TxtSROff), "0.00")
    'MRP Tax Calculation
    'Apply conditional Tax on Lube
   
    Dim rRTax As Double, gTax As Double
    rRTax = Val(TxtSTaxPer) + Format((Val(TxtSTaxPer) * Val(TxtTaxSurPer) / 100), "0.00")
    If PubSDTYN = 1 Then
        rRTax = rRTax + Val(TxtTurnOverPer)
    Else
     If pubTOT_On = 0 Then
        rRTax = rRTax + Val(TxtTurnOverPer) + (rRTax * Val(TxtTurnOverPer) / 100)
     Else
        rRTax = rRTax + Val(TxtTurnOverPer)
     End If
    End If
    If PubTOTOnLube = 1 Then
       'mMRPTax = Round((mTBTotM + mTBTotML - mTBDisAmtMRP) * Val(TxtSTaxPer) / 100, 2)
        gTax = Round((TotMRPAmtTB + mTBTotML - mTBDisAmtMRP) * rRTax / (100 + rRTax), 2)
        mMRPTax = Round((TotMRPAmtTB + mTBTotML - (mTBDisAmtMRP + gTax)) * Val(TxtSTaxPer) / 100, 2)
        mMRPTaxSur = Round(mMRPTax * Val(TxtTaxSurPer) / 100, 2)
        If PubSDTYN = 1 Then
              mMRPTOT = Round((TotMRPAmtTB + mTBTotML - (mTBDisAmtMRP + gTax)) * Val(TxtTurnOverPer) / 100, 2)
        Else
            If pubTOT_On = 0 Then   '0-Sub Total (B)
                mMRPTOT = Round((mTBTotM + mTBTotML + mTPTotM + mTPTotML - mTBDisAmtMRP - mTPDisAmtMRP + mMRPTax + mMRPTaxSur) * Val(TxtTurnOverPer) / 100, 2)
            Else    '1-Taxable+Taxpaid Total
                mMRPTOT = Round((mTBTotM + mTBTotML + mTPTotM + mTPTotML - mTBDisAmtMRP - mTPDisAmtMRP) * Val(TxtTurnOverPer) / 100, 2)
            End If
        End If
    Else
        mTotalDisLube = mTBDisAmtMRPLube + mTPDisAmtMRPLube
        mTotalLube = (mTBTotML + mTPTotML) - mTotalDisLube
        '--
        'mMRPTax = Round((mTBTotM - mTBDisAmtMRP) * Val(TxtSTaxPer) / 100, 2)
        gTax = Round((TotMRPAmtTB - mTBDisAmtMRP) * rRTax / (100 + rRTax), 2)
        mMRPTax = Round((TotMRPAmtTB - (mTBDisAmtMRP + gTax)) * Val(TxtSTaxPer) / 100, 2)
        
'        mMRPTax = Round((TotMRPAmtTB - mTBDisAmtMRP) * Val(TxtSTaxPer) / (100 + Val(TxtSTaxPer)), 2)
        mMRPTaxSur = Round(mMRPTax * Val(TxtTaxSurPer) / 100, 2)
        If PubSDTYN = 1 Then
              mMRPTOT = Round((TotMRPAmtTB - (mTBDisAmtMRP + gTax)) * Val(TxtTurnOverPer) / 100, 2)
        Else
            If pubTOT_On = 0 Then   '0-Sub Total (B)
                mMRPTOT = Round((mTBTotM + mTPTotM - mTotalLube - mTBDisAmtMRP - mTPDisAmtMRP + mMRPTax + mMRPTaxSur) * Val(TxtTurnOverPer) / 100, 2)
            Else    '1-Taxable+Taxpaid Total
                mMRPTOT = Round((mTBTotM + mTPTotM - mTotalLube - mTBDisAmtMRP - mTPDisAmtMRP) * Val(TxtTurnOverPer) / 100, 2)
            End If
        End If
    End If
    
    'Enable / Disable Text Box if values zero
    DisableEnableFooter TxtMRPAmtTB, TxtMRPAmtTP, TxtSprAmtTB, TxtSprAmtTP, _
        TxtOilAmtTB, TxtOilAmtTP, TxtDiscPerTB, TxtDiscPerTP, _
        TxtDiscAmtTB, TxtDiscAmtTP, TxtSTotATB, TxtGenSurPer, TxtGenSurAmt, _
        TxtTaxableTot, TxtSTaxPer, TxtSTaxAmt, TxtTaxSurPer, TxtTaxSurAmt
    'EOF enable / disable section
    TxtNetAmt = Format(Val(TxtNetSprAmt), "0.00")
End Sub

Public Sub LabCalc(TxtLabAmtTB As TextBox, TxtLabAmtTP As TextBox, TxtLabDisc As TextBox, _
                TxtServTaxPer As TextBox, TxtServTaxAmt As TextBox, _
                TxtLabRoff As TextBox, TxtNetLabAmt As TextBox, _
                OutSideLab As TextBox, mLabDiscAmtTB As Single, TxteCessPer As Double, TxteCessAmt As Double, TxtFreeWarrLabAmt As TextBox, ServiceTaxPer_Saperate As Double, ServiceTaxAmt_Saperate As Double, HECessPer As Double, HECessAmt As Double)
Dim mTotLab As Double
Dim mDiscTB As Double, mDiscTP As Double, mTaxableLab As Double
Dim mLabDiscAfterTaxYn As Byte
    If Val(TxtLabDisc) > (Val(TxtLabAmtTB) + Val(TxtLabAmtTP)) Then  ' + Val(OutSideLab)) Then
        TxtLabDisc = Format(Val(TxtLabAmtTB) + Val(TxtLabAmtTP), "0.00")
    End If
    mLabDiscAfterTaxYn = VNull(GCn.Execute("Select LabDiscAfterTaxYn From Syctrl").Fields(0).Value)
        
    mDiscTB = 0
    mDiscTP = 0
    If Val(TxtLabAmtTB) <> 0 And Val(TxtLabDisc) > 0 Then
        mDiscTB = (Val(TxtLabDisc) * Val(TxtLabAmtTB)) / (Val(TxtLabAmtTB) + Val(TxtLabAmtTP))
    End If
    If Val(TxtLabAmtTP) <> 0 And Val(TxtLabDisc) > 0 Then
        mDiscTP = Val(TxtLabDisc) * Val(TxtLabAmtTP) / (Val(TxtLabAmtTB) + Val(TxtLabAmtTP))
    End If
    If mLabDiscAfterTaxYn = 1 Then
        mTotLab = (Val(TxtLabAmtTB) + Val(TxtLabAmtTP))
        mTaxableLab = Val(TxtLabAmtTB)
    Else
        mTotLab = (Val(TxtLabAmtTB) + Val(TxtLabAmtTP)) - Val(TxtLabDisc)
        mTaxableLab = Val(TxtLabAmtTB) - mDiscTB
    End If

    
    If Val(TxtServTaxPer) <> 0 Then
        If PubTaxOnFreeLabYn = 1 Then
            TxtServTaxAmt = Format(((mTaxableLab + Val(TxtFreeWarrLabAmt)) * Val(TxtServTaxPer)) / 100, "0.00")
            ServiceTaxAmt_Saperate = Format((mTaxableLab + Val(TxtFreeWarrLabAmt)) * ServiceTaxPer_Saperate / 100, "0.00")
        Else
            TxtServTaxAmt = Format((mTaxableLab * Val(TxtServTaxPer)) / 100, "0.00")
            ServiceTaxAmt_Saperate = Format(mTaxableLab * ServiceTaxPer_Saperate / 100, "0.00")
        End If
    End If

    TxteCessAmt = Format(ServiceTaxAmt_Saperate * TxteCessPer / 100, "0.00")
    HECessAmt = Format(ServiceTaxAmt_Saperate * HECessPer / 100, "0.00")

    If mLabDiscAfterTaxYn = 1 Then
        mTotLab = Format(mTotLab + Val(TxtServTaxAmt) - Val(TxtLabDisc), "0.00")
    Else
        mTotLab = Format(mTotLab + Val(TxtServTaxAmt), "0.00")
    End If
    TxtLabRoff = Format(dmRoundOff(mTotLab), "0.00")
    TxtNetLabAmt = Format(mTotLab + Val(TxtLabRoff), "0.00")
End Sub

Public Function CheckPerm(PermType As ObjTypePerm, FrmName As Form) As Boolean
CheckPerm = False
If PermType = ad Then
    If mID(FrmName.TopCtrl1.Tag, 1, 1) = "A" Then
        CheckPerm = True
    End If
ElseIf PermType = Ed Then
    If mID(FrmName.TopCtrl1.Tag, 2, 1) = "E" Then
        CheckPerm = True
    End If
ElseIf PermType = De Then
    If mID(FrmName.TopCtrl1.Tag, 3, 1) = "D" Then
        CheckPerm = True
    End If
ElseIf PermType = pr Then
    If mID(FrmName.TopCtrl1.Tag, 4, 1) = "P" Then
        CheckPerm = True
    End If
End If
End Function

'Update Stock Qty
Public Sub UpdStkGridToTable(mPartNo As String, mSign As String, mMRPYN As String, mTaxYN As String, mQty As Double)
Dim Cur_TB_STk As Integer
        If UCase(mMRPYN) = "NO" Then  'MRP=No
            If UCase(mTaxYN) = "NO" Then  'Tax=No
                GSQL = "Update Part set Cur_TP_STk=Cur_TP_STk " & mSign & mQty & ""
'                If mSign = "+" Then  'Add
'                    RsPart!Cur_TP_Stk = RsPart!Cur_TP_Stk + mQty
'                Else
'                    RsPart!Cur_TP_Stk = RsPart!Cur_TP_Stk - mQty
'                End If
            Else
                GSQL = "Update Part set Cur_TB_STk=Cur_TB_STk " & mSign & mQty & ""
'                If mSign = "+" Then  'Add
'                    RsPart!Cur_TB_Stk = RsPart!Cur_TB_Stk + mQty
'                Else
'                    RsPart!Cur_TB_Stk = RsPart!Cur_TB_Stk - mQty
'                End If
            End If
        Else    'MRP=Yes
            If UCase(mTaxYN) = "NO" Then  'Tax=No
                GSQL = "Update Part set Cur_MRP_TPSTk=Cur_MRP_TPSTk " & mSign & mQty & ""
'                If mSign = "+" Then  'Add
'                    RsPart!Cur_MRP_TPStk = RsPart!Cur_MRP_TPStk + mQty
'                Else
'                    RsPart!Cur_MRP_TPStk = RsPart!Cur_MRP_TPStk - mQty
'                End If
            Else
                GSQL = "Update Part set Cur_MRP_TBSTk=Cur_MRP_TBSTk " & mSign & mQty & ""
'                If mSign = "+" Then  'Add
'                    RsPart!Cur_MRP_TBStk = RsPart!Cur_MRP_TBStk + mQty
'                Else
'                    RsPart!Cur_MRP_TBStk = RsPart!Cur_MRP_TBStk - mQty
'                End If
            End If
        End If
        GSQL = GSQL & " Where Part_No='" & mPartNo & "' and Div_Code='" & PubDivCode & "'"
        GCn.Execute GSQL
End Sub

Public Sub UpdStkTableToTable(mDocId As String, mSign As String, mRectIss As String)
Dim Rst As ADODB.Recordset, I As Byte

If mRectIss = "R" Then
    GSQL = "Select Part_No,Tax_YN,MRP_YN,Qty_Rec-Qty_Ret as Qty From SP_Stock Where DocId='" & mDocId & "' Order by DocID,Srl_No"
Else
    GSQL = "Select Part_No,Tax_YN,MRP_YN,Qty_Iss-Qty_Ret as Qty From SP_Stock Where DocId='" & mDocId & "' Order by DocID,Srl_No"
End If

Set Rst = New ADODB.Recordset
Rst.CursorLocation = adUseClient
Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
If Rst.RecordCount > 0 Then
    While Not Rst.EOF
        If Rst!MRP_YN = 0 Then  'MRP=No
            If Rst!Tax_YN = 0 Then  'Tax=No
                GSQL = "Update Part set Cur_TP_STk=" & vIsNull("Cur_TP_STk", "0") & " " & mSign & Rst!Qty & ""
            Else
                GSQL = "Update Part set Cur_TB_STk=" & vIsNull("Cur_TB_STk", "0") & " " & mSign & Rst!Qty & ""
            End If
        Else    'MRP=Yes
            If Rst!Tax_YN = 0 Then  'Tax=No
                GSQL = "Update Part set Cur_MRP_TPSTk=" & vIsNull("Cur_MRP_TPSTk", "0") & " " & mSign & Rst!Qty & ""
            Else
                GSQL = "Update Part set Cur_MRP_TBSTk=" & vIsNull("Cur_MRP_TBSTk", "0") & " " & mSign & Rst!Qty & ""
            End If
        End If
        GCn.Execute GSQL & " Where Part_No='" & Rst!Part_No & "' and Div_Code='" & PubDivCode & "'"
        Rst.MoveNext
    Wend
End If
Set Rst = Nothing
End Sub

Public Sub PostLedg(A_E As String, FACn As ADODB.Connection, DocID As String, VDate As String, _
    DrAc As String, CrAc As String, AmtDr As Double, AmtCr As Double, Narr As String, _
    VSNo As Integer, Optional SiteStr As String, _
    Optional VType As String, Optional VNo As Long)
    
'   If A_E = "Add" Then
        FACn.Execute "insert into ledger(" _
            & "DocId,Site_Code,v_sNo,V_type,v_no," _
            & "v_date,subcode,contrasub,amtdr,amtcr,narration,U_Name, U_EntDt, U_AE)" _
            & " values(" _
            & "'" & DocID & "','" & PubSiteCode & SiteStr & "'," & VSNo & ",'" & VType & "'," & VNo & "," _
            & "" & ConvertDate(VDate) & ",'" & DrAc & "','" & CrAc & "'," & AmtDr & "," & AmtCr & ",'" & Narr & "'," _
            & "'" & pubUName & "',#" & PubServerDate & "#,'" & left(A_E, 1) & "')"
'   Else
'        FACn.Execute "update ledger set subcode='" & DrAc & "',contrasub='" & CrAc & "', " & _
'            "amtdr=" & AmtDr & ",amtcr=" & AmtCr & ",narration='" & Narr & "', " & _
'            "U_Name='" & pubUName & "', U_EntDt=#" & PubServerDate & "#, U_AE='" & Mid(TopCtrl1.TopText2, 1, 1) & "'  " & _
'            "where docid = '" & Docid & "' and V_SNo = " & VSNo & ""
'    End If
End Sub

Public Sub UnPostLedg(FACn As ADODB.Connection, DocID As String)
FACn.Execute "delete from ledger " _
            & " where docid = '" & DocID & "'"
End Sub

Public Sub ListViewReport_KeyDown(FrmList As Object, LV As Object, Txt As Object, Index As Integer, KeyCode As Integer, Shift As Integer, left As Integer, top As Integer, width As Integer, Optional height As Integer)
If FilterKeyCode(KeyCode) = True Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Txt(Index).TEXT <> "" Then
            If Not LV.SelectedItem Is Nothing Then
                Txt(Index).TEXT = LV.SelectedItem.TEXT
            End If
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
                Txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.SelectedItem.Index < LV.ListItems.Count Then
                LV.ListItems(LV.SelectedItem.Index + 1).SELECTED = True
                Txt(Index).TEXT = LV.SelectedItem.TEXT
            ElseIf KeyCode = vbKeyDown And LV.ListItems.Count = 1 Then
                Txt(Index).TEXT = LV.SelectedItem.TEXT
            End If
        End If
    End If
End Sub


Public Function GetMrpTBStk(PartNo As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = GCn.Execute("Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & "", "'SXAO'") & ") And  S.Part_No='" & PartNo & "' And Tax_Yn=1 And Mrp_Yn=1")
    If RsTemp.RecordCount > 0 Then
        GetMrpTBStk = VNull(RsTemp(0))
    Else
        GetMrpTBStk = 0
    End If
    Set RsTemp = Nothing
End Function

Public Function GetMrpTPStk(PartNo As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = GCn.Execute("Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & "", "'SXAO'") & ") And  S.Part_No='" & PartNo & "' And Tax_Yn=0 And Mrp_Yn=1")
    If RsTemp.RecordCount > 0 Then
        GetMrpTPStk = VNull(RsTemp(0))
    Else
        GetMrpTPStk = 0
    End If
    Set RsTemp = Nothing
End Function

Public Function GetTBStk(PartNo As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = GCn.Execute("Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & "", "'SXAO'") & ") And  S.Part_No='" & PartNo & "' And Tax_Yn=1 And Mrp_Yn=0")
    If RsTemp.RecordCount > 0 Then
        GetTBStk = VNull(RsTemp(0))
    Else
        GetTBStk = 0
    End If
    Set RsTemp = Nothing
End Function

Public Function GetTPStk(PartNo As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = GCn.Execute("Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & "", "'SXAO'") & ") And  S.Part_No='" & PartNo & "' And Tax_Yn=0 And Mrp_Yn=0")
    If RsTemp.RecordCount > 0 Then
        GetTPStk = VNull(RsTemp(0))
    Else
        GetTPStk = 0
    End If
    Set RsTemp = Nothing
End Function

Private Function GetStk(PartNo As String) As Double
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = GCn.Execute("Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & "", "'SXAO'") & ") And  S.Part_No='" & PartNo & "'")
    If RsTemp.RecordCount > 0 Then
        GetStk = VNull(RsTemp(0))
    Else
        GetStk = 0
    End If
    Set RsTemp = Nothing
End Function



Public Sub Fill_Data(PartyType As Byte, LblFrm As Object, FGrid As Object, _
    PNo As String, PName As String, LName As String, _
    Col_Unit As Byte, Col_MRP As Byte, Col_Taxable As Byte, _
    Col_MRPStkTB As Byte, Col_MRPStkTP As Byte, _
    Col_TBQty As Byte, Col_TPQty As Byte, _
    Col_MRPRate As Byte, Col_TBRate As Byte, _
    Col_TPRate As Byte, Col_Bin As Byte, _
    Col_HPRate As Byte, Col_LPRate As Byte, _
    Col_LastRate As Byte, Col_PartGrade As Byte, _
    Col_EffectDate As Byte, Col_DiscPer As Byte, CheckNegetiveStockSiteWise As Boolean, Optional PurTrn As Boolean)


'If Trim(FGrid.TextMatrix(FGrid.Row, Col_PNo)) <> "" Then
If PNo <> "" Then
    Dim RsTemp As ADODB.Recordset, RsStkSite As ADODB.Recordset, MRPDisc As Single, TBDisc As Single, TPDisc As Single
    Set RsTemp = New Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open "Select MRP_Disc,TB_Disc,TP_Disc from SubGroupType where Party_Type =" & PartyType, GCn, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        MRPDisc = RsTemp!mrp_Disc
        TBDisc = RsTemp!tb_Disc
        TPDisc = RsTemp!tp_Disc
    End If
    
    'GSQL = "Select P.MRP_Effect_Dt,P.TB_Effect_Dt,P.Part_Grade,P.Unit,P.Bin_Loca ," _
        & "val(format(P.MRP-(P.MRP* " & MRPDisc & "/100),'0.00')) as MRP," _
        & "val(format(P.TB_SRate-(P.TB_SRate* " & TBDisc & "/100),'0.00')) as TB_SRate," _
        & "val(format(P.TP_SRate-(P.TP_SRate*" & TPDisc & "/100),'0.00')) as TP_SRate," _
        & "P.High_Pur_Rate, P.Low_Pur_Rate, " _
        & "P.Disc_Factor,Part_DiscFactor.SalDisc_Per, Part_DiscFactor.PurcDisc_Per, " _
        & "P.Cur_MRP_TBStk, P.Cur_MRP_TPStk,P.Cur_TB_Stk ,P.Cur_TP_Stk " _
        & "From Part P " _
        & " Left Join Part_DiscFactor On P.Disc_Factor = Part_DiscFactor.DiscFac_Catg " _
        & " where P.Part_No='" & PNo & "' AND P.div_code ='" & PubDivCode & "'"
    GSQL = "Select P.MRP_Effect_Dt,P.TB_Effect_Dt,P.Part_Grade,P.Unit,P.Bin_Loca ," _
        & " " & VFormat(" p.MRP - (p.MRP * " & MRPDisc & " / 100) ", "0.00") & "  as MRP," _
        & " " & VFormat(" p.TB_SRate - (p.TB_SRate * " & TBDisc & " / 100) ", "0.00") & " as TB_SRate," _
        & " " & VFormat(" p.TP_SRate - (p.TP_SRate * " & TPDisc & " / 100) ", "0.00") & " as TP_SRate," _
        & "P.High_Pur_Rate, P.Low_Pur_Rate, " _
        & "P.Disc_Factor,Part_DiscFactor.SalDisc_Per, Part_DiscFactor.PurcDisc_Per, " _
        & "P.Cur_MRP_TBStk, P.Cur_MRP_TPStk,P.Cur_TB_Stk ,P.Cur_TP_Stk, P.PurRate, P.NDP " _
        & "From Part P " _
        & " Left Join Part_DiscFactor On P.Disc_Factor = Part_DiscFactor.DiscFac_Catg " _
        & " where P.Part_No='" & PNo & "' AND P.div_code ='" & PubDivCode & "'"
            
    Set RsTemp = New Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    
    RsPartSiteWise.Filter = adFilterNone
    RsPartSiteWise.Filter = "Code='" & PNo & "'"
    
    If RsTemp!Part_Grade = PubPartGrade_Lub Then
        FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No"
    End If
'   FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "No"
    FGrid.TextMatrix(FGrid.Row, Col_Unit) = IIf(IsNull(RsTemp!Unit), "", RsTemp!Unit)
    
    If CheckNegetiveStockSiteWise Then
        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB) = VNull(RsPartSiteWise!Cur_MRP_TbStk)
        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP) = VNull(RsPartSiteWise!Cur_MRP_TPStk)
        FGrid.TextMatrix(FGrid.Row, Col_TBQty) = VNull(RsPartSiteWise!Cur_TB_STk)
        FGrid.TextMatrix(FGrid.Row, Col_TPQty) = VNull(RsPartSiteWise!Cur_TP_Stk)
    Else
'        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB) = Format(IIf(IsNull(RsTemp!Cur_MRP_TbStk), 0, RsTemp!Cur_MRP_TbStk), "0.000")
'        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP) = Format(IIf(IsNull(RsTemp!Cur_MRP_TPStk), 0, RsTemp!Cur_MRP_TPStk), "0.000")
'        FGrid.TextMatrix(FGrid.Row, Col_TBQty) = Format(IIf(IsNull(RsTemp!Cur_TB_STk), 0, RsTemp!Cur_TB_STk), "0.000")
'        FGrid.TextMatrix(FGrid.Row, Col_TPQty) = Format(IIf(IsNull(RsTemp!Cur_TP_Stk), 0, RsTemp!Cur_TP_Stk), "0.000")
        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB) = GetMrpTBStk(PNo)
        FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP) = GetMrpTPStk(PNo)
        FGrid.TextMatrix(FGrid.Row, Col_TBQty) = GetTBStk(PNo)
        FGrid.TextMatrix(FGrid.Row, Col_TPQty) = GetTPStk(PNo)

    End If
    FGrid.TextMatrix(FGrid.Row, Col_MRPRate) = Format(IIf(IsNull(RsTemp!MRP), 0, RsTemp!MRP), "0.000")
    FGrid.TextMatrix(FGrid.Row, Col_TBRate) = Format(IIf(IsNull(RsTemp!TB_SRate), 0, RsTemp!TB_SRate), "0.000")
    FGrid.TextMatrix(FGrid.Row, Col_TPRate) = Format(IIf(IsNull(RsTemp!TP_SRate), 0, RsTemp!TP_SRate), "0.000")
    FGrid.TextMatrix(FGrid.Row, Col_Bin) = IIf(IsNull(RsTemp!Bin_Loca), "", RsTemp!Bin_Loca)
    FGrid.TextMatrix(FGrid.Row, Col_LastRate) = Format(RsTemp!NDP, "0.00")
    FGrid.TextMatrix(FGrid.Row, Col_HPRate) = Format(IIf(IsNull(RsTemp!high_pur_rate), 0, RsTemp!high_pur_rate), "0.00")
    FGrid.TextMatrix(FGrid.Row, Col_LPRate) = Format(IIf(IsNull(RsTemp!low_pur_rate), 0, RsTemp!low_pur_rate), "0.00")
    
    FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = RsTemp!Part_Grade
    FGrid.TextMatrix(FGrid.Row, Col_EffectDate) = Format(IIf(FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes", RsTemp!MRP_Effect_Dt, RsTemp!TB_Effect_Dt), "dd/MMM/yyyy")
    If PurTrn Then
        FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsTemp!PurcDisc_Per, "0.00")
    Else
        FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsTemp!SalDisc_Per, "0.00")
    End If
    Set RsTemp = Nothing
  
Else
    FGrid.TextMatrix(FGrid.Row, Col_MRP) = ""
    FGrid.TextMatrix(FGrid.Row, Col_Taxable) = ""
    FGrid.TextMatrix(FGrid.Row, Col_Unit) = ""
    FGrid.TextMatrix(FGrid.Row, Col_MRPRate) = ""
    FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB) = ""
    FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP) = ""
    FGrid.TextMatrix(FGrid.Row, Col_TBQty) = ""
    FGrid.TextMatrix(FGrid.Row, Col_TPQty) = ""
    FGrid.TextMatrix(FGrid.Row, Col_TBRate) = ""
    FGrid.TextMatrix(FGrid.Row, Col_TPRate) = ""
    FGrid.TextMatrix(FGrid.Row, Col_Bin) = ""
    FGrid.TextMatrix(FGrid.Row, Col_LastRate) = ""
    FGrid.TextMatrix(FGrid.Row, Col_HPRate) = ""
    FGrid.TextMatrix(FGrid.Row, Col_LPRate) = ""
    FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = ""
    FGrid.TextMatrix(FGrid.Row, Col_EffectDate) = ""
    FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = ""
End If

Fill_Frame LblFrm, FGrid, PNo, PName, LName, _
    Col_MRPStkTB, Col_MRPStkTP, _
    Col_TBQty, Col_TPQty, _
    Col_MRPRate, Col_TBRate, _
    Col_TPRate, Col_Bin, _
    Col_LastRate, Col_HPRate, Col_LPRate, CheckNegetiveStockSiteWise

End Sub

Public Sub Fill_Frame(LblFrm As Object, FGrid As Object, _
    PNo As String, PName As String, LName As String, _
    Col_MRPStkTB As Byte, Col_MRPStkTP As Byte, _
    Col_TBQty As Byte, Col_TPQty As Byte, _
    Col_MRPRate As Byte, Col_TBRate As Byte, _
    Col_TPRate As Byte, Col_Bin As Byte, _
    Col_LastRate As Byte, Col_HPRate As Byte, Col_LPRate As Byte, CheckNegetiveStockSiteWise As Boolean)

Dim CurStkQty As Double
CurStkQty = Val(FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP)) + _
         Val(FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB)) + _
         Val(FGrid.TextMatrix(FGrid.Row, Col_TBQty)) + _
         Val(FGrid.TextMatrix(FGrid.Row, Col_TPQty))

    LblFrm(0).CAPTION = PNo
    LblFrm(1).CAPTION = PName
    LblFrm(2).CAPTION = LName
    LblFrm(3).CAPTION = FGrid.TextMatrix(FGrid.Row, Col_Bin)
            
    If PNo <> "" Then
        RsPartSiteWise.Filter = adFilterNone
        RsPartSiteWise.Filter = "Code='" & PNo & "'"
    End If
        
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        LblFrm(4).CAPTION = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TBQty)), "0.000")
        LblFrm(5).CAPTION = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TPQty)), "0.000")
    Else
        If CheckNegetiveStockSiteWise Then
            LblFrm(4).CAPTION = Format(VNull(RsPartSiteWise!Cur_MRP_TbStk), "0.000")
            LblFrm(5).CAPTION = Format(VNull(RsPartSiteWise!Cur_MRP_TPStk), "0.000")
            LblFrm(6).CAPTION = Format(VNull(RsPartSiteWise!Cur_TB_STk), "0.000")
            LblFrm(7).CAPTION = Format(VNull(RsPartSiteWise!Cur_TP_Stk), "0.000")
        Else
            LblFrm(4).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_MRPStkTB), "0.000")
            LblFrm(5).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_MRPStkTP), "0.000")
            LblFrm(6).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_TBQty), "0.000")
            LblFrm(7).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_TPQty), "0.000")
        End If
    End If
    
    If CheckNegetiveStockSiteWise Then
        LblFrm(8).CAPTION = Format(VNull(RsPartSiteWise!CurrStk), "0.000")
    Else
        LblFrm(8).CAPTION = Format(CurStkQty, "0.000")
    End If
    
    LblFrm(9).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_MRPRate), "0.000")
    LblFrm(10).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_TBRate), "0.000")
    LblFrm(11).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_TPRate), "0.000")
    'LblFrm(12).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_HPRate), "0.00")
    'LblFrm(13).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_LPRate), "0.00")
    'LblFrm(14).CAPTION = Format(FGrid.TextMatrix(FGrid.Row, Col_LastRate), "0.00")
    If PNo <> "" Then
        LblFrm(14).CAPTION = Format(VNull(GCn.Execute("Select PurRate From Part where Part_No='" & PNo & "' and Div_Code='" & PubDivCode & "'").Fields(0).Value), "0.00")
        LblFrm(12).CAPTION = Format(VNull(GCn.Execute("Select Max(Rate) From SP_Stock where Part_No='" & PNo & "' and v_Type='SXGR'").Fields(0).Value), "0.00")
        LblFrm(13).CAPTION = Format(VNull(GCn.Execute("Select Min(Rate) From SP_Stock where Part_No='" & PNo & "' and v_Type='SXGR'").Fields(0).Value), "0.00")
    Else
        LblFrm(14).CAPTION = "0.00"
        LblFrm(12).CAPTION = "0.00"
        LblFrm(13).CAPTION = "0.00"
    End If

End Sub

Public Sub SelGridKeyPress(Txt As Object, FGrid As MSHierarchicalFlexGridLib.MSHFlexGrid, _
    Rst As ADODB.Recordset, ByVal KeyAscii As Integer, FindFldName As String, _
    Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants, Optional FixedRows As Integer)
Dim FindStr$    ' As String
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If FixedRows = 0 Then FixedRows = 1
    If FGrid.Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then Txt.TEXT = "": Exit Sub
    
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack Then
            If Len(Txt.SelText) > 1 Then
                Txt.SelLength = Len(Txt.SelText) - 1
                FindStr = Txt.SelText
            Else
                Txt.TEXT = ""
                FGrid.SetFocus
                Txt.Visible = False
                Exit Sub
            End If
        Else
            FindStr = Txt.SelText + Chr(KeyAscii)
        End If
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
        Else    'character serach
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
        End If
        KeyAscii = 0
       If Rst.AbsolutePosition <> adPosEOF And Rst.AbsolutePosition <> adPosBOF Then
            FGrid.CellBackColor = CellBackColLeave
            FGrid.Row = (FixedRows - 1) + Rst.AbsolutePosition
            FGrid.CellBackColor = CellBackColEnter
            Txt.TEXT = Rst.Fields(FindFldName).Value
            Txt.SelLength = Len(FindStr)
            Txt.left = FGrid.CellLeft + FGrid.left
            Txt.top = FGrid.CellTop + FGrid.top
            
            If Txt.Visible = False Then
                Txt.Visible = True: Txt.ZOrder 0: Txt.SetFocus: Txt.BackColor = FGrid.CellBackColor
                Txt.ForeColor = FGrid.CellForeColor: Txt.width = FGrid.CellWidth: Txt.height = FGrid.CellHeight
            End If
           
       End If
    End Sub

Public Function NavigationKey(KeyCode As Integer) As Boolean
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyUp _
    Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        NavigationKey = True
    End If
End Function
Public Function PrinterDetail() As String
'SprQuot,SprSaleOrd,SprSaleReg,SprSaleRet,SprPurOrd,SprMatReg,SprPurReg,SprPurRet,SprStkTrf,
'SprStkReg10,SprStkSumm,SprStkInHand,VehMoneyRect,WksEstimate,WksPerforma,WksSaleReg,
'WksReqReg,WksVehDiary,WksJobReg,SprIndent,SprDailySale,SprMonthSale,SprPartPur,SprPartSale,
'SprPurSum,SprSaleSum,SprPurForm,SprStkReOrd,SprstkBin,SprPartMovement,SprStkAgeing
  Dim PrnDet As String
  PrnDet = Printer.DeviceName
  Select Case rpt.PaperSize  ' Letter 8.5 * 11 = 1/fanfold 8.5 * 12 = 263/A4 = 9/15 * 12 = 268 /LandScape 2 /Portrait =1
    Case 1
        PrnDet = PrnDet & " Size 8.5 * 11"
    Case 9
        PrnDet = PrnDet & " A4"
    Case 263
        PrnDet = PrnDet & " Size 8.5 * 12"
    Case 268
        PrnDet = PrnDet & " Size 15 * 12"
    Case Else
        PrnDet = PrnDet & " Size UnDefined"
End Select
Select Case rpt.PaperOrientation
Case 1
    PrnDet = PrnDet & " (Portrait)"
Case 2
    PrnDet = PrnDet & " (Landscape)"
End Select
PrinterDetail = PrnDet
End Function

Public Function PRN_TIT(st1 As String, mFont As String, LNT As Byte, Optional UnderLine As Boolean) As String
Dim LEN1 As Integer, WDT
PRN_TIT = ""
st1 = Trim(st1)
LEN1 = Len(st1)
Select Case mFont
    Case "A"
        WDT = Int(LNT / 2)
        If UnderLine Then
            PRN_TIT = Chr(18) + Chr(14) + Chr(27) + "G" + mUnd + Space(Abs(WDT - LEN1) / 2) + st1 + mUnd1 + Chr(27) + "H"
        Else
            PRN_TIT = Chr(18) + Chr(14) + Chr(27) + "G" + Space(Abs(WDT - LEN1) / 2) + st1 + Chr(27) + "H"
        End If
    Case "B"
'        WDT = Int(LNT * 7 / 8)  ' Alignment purpose
        WDT = LNT
        If UnderLine Then
'            PRN_TIT = Chr(14) + Chr(15) + mUnd + Space((WDT - LEN1) / 2) + ST1 + mUnd1 + Chr(18)
            PRN_TIT = Space((WDT - LEN1) / 2) + mEmph + mUnd + st1 + mUnd1 + mEmph1
        Else
'            PRN_TIT = Chr(14) + Chr(15) + Space((WDT - LEN1) / 2) + ST1 + Chr(18)
           PRN_TIT = Space((WDT - LEN1) / 2) + mEmph + st1 + mEmph1
        End If
    Case "C"
        WDT = LNT
        If UnderLine Then
            PRN_TIT = Space((WDT - LEN1) / 2) + mChr18 + mUnd + st1 + mUnd1
        Else
            PRN_TIT = Space((WDT - LEN1) / 2) + mChr18 + st1
        End If
End Select
End Function

Public Function SETW(mSTRING As String, mLEN As Integer) As String
    mSTRING = mID(mSTRING, 1, mLEN)
    SETW = Trim(mSTRING) + Space(mLEN - Len(Trim(mSTRING)))
End Function

Public Function SETN(mSTRING As String, mLEN As Integer) As String
    mSTRING = mID(mSTRING, 1, mLEN)
    SETN = Space(mLEN - Len(Trim(mSTRING))) + Trim(mSTRING)
End Function

Public Function PSTR(xVal As Variant, xLen As Byte, Optional xDeci As Byte, Optional TxtAlign As TxtAlignDef, Optional PrintCharacter As String) As String
Dim TempStr$
'xValType = VarType(xVal)
'xDeci = IIf(xDeci, 0, xDeci)
'PSTR = IIf(xVal = 0, Space(xLen - 2 - xDeci) + "--" + Space(xDeci), IIf(xValType = "N", str(xVal, xLen, xDeci), left(xVal, xLen)))
PrintCharacter = IIf(PrintCharacter = "", "--", "  ")
If Len(xVal) > xLen Then xVal = mID(xVal, 1, xLen)
If VarType(xVal) = vbByte Or VarType(xVal) = vbDecimal Or VarType(xVal) = vbDouble Or VarType(xVal) = vbInteger Or VarType(xVal) = vbLong Or VarType(xVal) = vbSingle Then
    If TxtAlign = 0 Then
        If xVal <> 0 Then
            TempStr = Format(xVal, "0" & IIf(xDeci > 0, "." & Replace(Space(xDeci), " ", "0"), ""))
            If xLen > Len(TempStr) Then
                PSTR = Space(xLen - Len(TempStr)) + TempStr
            Else
                PSTR = TempStr
            End If
        Else
            PSTR = Space(xLen - (2 + xDeci)) + PrintCharacter + Space(xDeci)
        End If
    Else
        If xVal <> 0 Then
            TempStr = Format(xVal, "0" & IIf(xDeci > 0, "." & Replace(Space(xDeci), " ", "0"), ""))
            PSTR = TempStr + Space(IIf(xLen - Len(STR(TempStr)) > 0, xLen - Len(STR(TempStr)), 1))
        Else
            PSTR = Space(xDeci) + PrintCharacter + Space(xLen - (2 + xDeci))
        End If
    End If
Else
    If TxtAlign = 0 Then
        If xVal <> "" Then
            PSTR = LTrim(xVal + Space(xLen - Len(xVal)))
        Else
            PSTR = LTrim(PrintCharacter + Space(xLen - 2))
        End If
    Else
        If xVal <> "" Then
            PSTR = Space(xLen - Len(xVal)) + xVal
        Else
            PSTR = Space(xLen - 2) + PrintCharacter
        End If
    End If
End If
End Function

'Public Function Det_Tax(MRate As Double, mOth_Amt As Double, mTB_D_PER As Double, ByRef mAmount As Double) As String
'Dim mDAMT, mOTH_IMAMT
'mDAMT = Round((mAmount * mTB_D_PER) / 100, 2)
'mOTH_IAMT = Round((mOth_Amt * (mAmount - mDAMT)) / mTB_TOT1, 2)
'MRate = Round((mAmount + mOTH_IAMT) / Qty, 2)
'mAmount = mAmount + mOTH_IAMT
'Det_Tax = PSTR(MRate, 9, 2)
'End Function

Public Sub DisableEnableFooter( _
        TxtMRPAmtTB As TextBox, TxtMRPAmtTP As TextBox, _
        TxtSprAmtTB As TextBox, TxtSprAmtTP As TextBox, _
        TxtOilAmtTB As TextBox, TxtOilAmtTP As TextBox, _
        TxtDiscPerTB As TextBox, TxtDiscPerTP As TextBox, _
        TxtDiscAmtTB As TextBox, TxtDiscAmtTP As TextBox, _
        TxtSTotATB As TextBox, _
        TxtGenSurPer As TextBox, TxtGenSurAmt As TextBox, _
        TxtTaxableTot As TextBox, _
        TxtSTaxPer As TextBox, TxtSTaxAmt As TextBox, _
        TxtTaxSurPer As TextBox, TxtTaxSurAmt As TextBox)
    'Enable / Disable Textbox considering zero values
Dim EnableText As Boolean
    EnableText = IIf(Val(TxtMRPAmtTB) + Val(TxtSprAmtTB) + Val(TxtOilAmtTB) = 0, False, True)
    TxtDiscPerTB.Enabled = EnableText
    TxtDiscAmtTB.Enabled = EnableText
    
    EnableText = IIf(Val(TxtMRPAmtTP) + Val(TxtSprAmtTP) + Val(TxtOilAmtTP) = 0, False, True)
    TxtDiscPerTP.Enabled = EnableText
    TxtDiscAmtTP.Enabled = EnableText
    
    EnableText = IIf(Val(TxtSTotATB) = 0, False, True)
    TxtGenSurPer.Enabled = EnableText
    TxtGenSurAmt.Enabled = EnableText
    
    EnableText = IIf(Val(TxtTaxableTot) = 0, False, True)
    TxtSTaxPer.Enabled = EnableText
    TxtSTaxAmt.Enabled = EnableText
    TxtTaxSurPer.Enabled = EnableText
    TxtTaxSurAmt.Enabled = EnableText
End Sub

Public Function fxLastDay(ByVal mDate As Date) As Byte
    If Format(mDate, "MM") = "12" Then
        fxLastDay = Day(CDate("31/" & Format(mDate, "MM") & "/" & Format(mDate, "YYYY")))
    Else
        fxLastDay = Day(CDate("1/" & Val(Format(mDate, "MM")) + 1 & "/" & Format(mDate, "YYYY")) - 1)
    End If
End Function

Public Function Check_Entry(TableName As String, FieldName As String, FieldValue As String, FieldDataType As DataTypeDef) As Boolean
Dim GSQL As String
If FieldDataType = 0 Then
    GSQL = "Select  count(" & FieldName & ") from " & TableName & " where " & FieldName & " = '" & FieldValue & "'"
ElseIf FieldDataType = 1 Then
    GSQL = "Select  count(" & FieldName & ") from " & TableName & " where " & FieldName & " = " & FieldValue & ""
ElseIf FieldDataType = 2 Then
    GSQL = "Select  count(" & FieldName & ") from " & TableName & " where " & FieldName & " = #" & FieldValue & "#"
End If
If GCn.Execute(GSQL).Fields(0).Value > 0 Then
    Check_Entry = False
    MsgBox "Related Record Exist in Table  " & TableName & ", Entry Can't Be Deleted", vbInformation, "Validation Check": Exit Function
Else
    Check_Entry = True
End If
End Function

Public Function VehSRate(EffDate As Date, Model As String, TaxYN As String, RsoYn As String, Optional NDP As Double) As Double
Dim rsRate As Recordset, Margin As Double, TaxYes As Byte, RSO As Byte
TaxYes = IIf(left(TaxYN, 1) = "Y", 1, 0)
RSO = IIf(left(RsoYn, 1) = "Y", 1, 0)

    Set rsRate = New Recordset
    rsRate.Open "Select top 1 P_RATE,s_rate,INCI_CHRG,OCTROI,REG_TEMP,INS_TRN,TRANSPORT,MVT,REG_FEE,INS_FEE " & _
        "from Veh_Rate " & _
        "where model = '" & Model & "' and Effective_Date<=" & ConvertDate(EffDate) & _
        " and RSO_WORK=" & RSO & " and TAXABLE_YN =" & TaxYes & "", GCn, adOpenStatic, adLockReadOnly
    If rsRate.RecordCount > 0 Then
         If NDP = 0 Then
            NDP = IIf(IsNull(rsRate!p_rate), 0, rsRate!p_rate)
         End If
         Margin = IIf(IsNull(rsRate!S_Rate), 0, rsRate!S_Rate) - IIf(IsNull(rsRate!p_rate), 0, rsRate!p_rate)
         VehSRate = NDP + Margin
    Else
        
        
    End If
    Set rsRate = Nothing
End Function

Public Function PartyAdvance(OrdDocId As String, Optional InvDate As String) As Double
    If InvDate = "" Then
        GSQL = "Select sum(" & cIIF("DrCr ='C'", "Amount", "Amount*-1") & ") as AdvAmt from Rect where Ord_DocId = '" & OrdDocId & "' and V_Type not in('G_TLR')"
    Else
        GSQL = "Select sum(" & cIIF("DrCr ='C'", "Amount", "Amount*-1") & ") as AdvAmt from Rect where Ord_DocId = '" & OrdDocId & "' and V_Type not in('G_TLR') and V_Date<=" & ConvertDate(InvDate) & ""
    End If
    Dim RstTemp As ADODB.Recordset
    Set RstTemp = GCn.Execute(GSQL)
    If RstTemp.RecordCount > 0 Then
        PartyAdvance = IIf(IsNull(RstTemp!AdvAmt), 0, RstTemp!AdvAmt)
    Else
        PartyAdvance = 0
    End If
    Set RstTemp = Nothing
End Function

Public Function FxReligion(TrnType As Variant) As Variant
If IsNumeric(TrnType) Then
    If TrnType = 1 Then
        FxReligion = "Hindu"
    ElseIf TrnType = 2 Then
        FxReligion = "Muslim"
    ElseIf TrnType = 3 Then
        FxReligion = "Sikh"
    ElseIf TrnType = 3 Then
        FxReligion = "Christian"
    Else
        FxReligion = "N/A"
    End If
Else
    If TrnType = "Hindu" Then
        FxReligion = 1
    ElseIf TrnType = "Muslim" Then
        FxReligion = 2
    ElseIf TrnType = "Sikh" Then
        FxReligion = 3
    ElseIf TrnType = "Christian" Then
        FxReligion = 4
    Else
        FxReligion = 0
    End If
End If
End Function
Public Sub Report_DocHeader(mREPORT As Variant)
 Dim mReportCount As Integer

    For mReportCount = 1 To mREPORT.FormulaFields.Count
        Select Case UCase(mREPORT.FormulaFields(mReportCount).FormulaFieldName)
            Case UCase("comp_name")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_Name & "'"
            Case UCase("comp_add1")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_Add & "'"
            Case UCase("comp_add2")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_Add2 & "'"
            Case UCase("comp_city")
                mREPORT.FormulaFields(mReportCount).TEXT = "'" & PubComp_City & "'"
'            Case UCase("Title")
'                mReport.FormulaFields(mReportCount).Text = "'" & Caption & "'"
        End Select
    Next

End Sub
Public Function UserPermission(FormCode As String) As String
On Error GoTo err1
Dim Rst As ADODB.Recordset
If pubUName = "SA" Then
    UserPermission = "AEDP"
    'modishekhar
    Set Rst = G_CompCn.Execute("select DelApply from User_Module where form_code='" & FormCode & "'")
    If Rst.RecordCount > 0 And Rst.EOF = False And Rst.BOF = False Then
        If Rst!delapply = 1 Then
            UserPermission = Replace(UserPermission, "D", "*")
        End If
    End If
    Exit Function
End If
Dim TSQL As String
Dim Rs As ADODB.Recordset
'Form_Code + Param_Str + Comp_Code + Div_Code
    TSQL = "select Param_Str from user2 where Comp_Code='" & PubCenCompCode & _
            "' and Div_code = '" & PubDivCode & "' and user_name='" & pubUName & _
            "' and form_code= '" & FormCode & "'"

    Set Rs = New ADODB.Recordset
    Rs.Open TSQL, G_CompCn, adOpenStatic, adLockReadOnly
    If Not Rs.EOF Then
        UserPermission = Rs!param_str
        'modishekhar
        Set Rst = G_CompCn.Execute("select DelApply from User_Module where form_code= '" & FormCode & "'")
        If Rst.RecordCount > 0 And Rst.EOF = False And Rst.BOF = False Then
            If Rst!delapply = 1 Then
                UserPermission = Replace(UserPermission, "D", "*")
            End If
        End If
    Else
        UserPermission = ""
        MsgBox "UnAuthorised Access", vbInformation, "Access Denied"
    End If
    Set Rs = Nothing
    Set Rst = Nothing
Exit Function
err1:
    MsgBox err.Description
End Function

Public Function PrinID(DocID As String) As String
PrinID = left(DocID, 1) & mID(DocID, 3, 1) & "/" & left(Trim(mID(DocID, 4, 5)), 1) & Right(mID(DocID, 4, 5), Len(Trim(mID(DocID, 4, 5))) - 2) & "/" & Trim(Right(DocID, 8))
End Function

Public Function TOTCal()
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
    Rst.Open "Select TOT_YN,SDT_YN,TOT_On,TOT_Rate from Syctrl", GCn, adOpenDynamic, adLockOptimistic, adCmdText
    If (Rst!TOT_YN = 1 And Rst!TOT_On = 0) Or Rst!SDT_YN = 1 Then
        TOTCal = Rst!TOT_Rate
    End If
    Set Rst = Nothing
    
End Function
Public Function Serv_Tax()
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
    Rst.Open "Select Service_Tax from Syctrl", GCn, adOpenDynamic, adLockOptimistic, adCmdText
            Serv_Tax = Rst!Service_Tax
    
    Set Rst = Nothing
End Function
Public Function CalcHrs(Date1 As Date, Date2 As Date, Time1 As String, time2 As String, WrkHrs As Double)
    Dim TotalHrs As Double, TotalDays As Double, TotalMins As Double, Hr1 As Double, Hr2 As Double, Min1 As Double, Min2 As Double
    Hr1 = IIf(Right(Time1, 2) = "AM", Val(left(Time1, 2)), Val(left(Time1, 2)) + 12)
    Hr2 = IIf(Right(time2, 2) = "AM", Val(left(time2, 2)), Val(left(time2, 2)) + 12)
    Min1 = Val(mID(Time1, 4, 2))
    Min2 = Val(mID(time2, 4, 2))
    TotalDays = DateDiff("d", Date1, Date2)
    If TotalDays > 0 Then
        TotalHrs = TotalDays * WrkHrs
        TotalHrs = TotalHrs + (Hr2 - Hr1)
        TotalMins = Min2 - Min1
        If TotalMins < 0 Then
            TotalMins = 60 + TotalMins
            TotalHrs = TotalHrs - 1
        End If
    Else
        TotalHrs = TotalHrs + (Hr2 - Hr1)
        TotalMins = Min2 - Min1
        If TotalMins < 0 Then
            TotalMins = (60 + TotalMins)
            TotalHrs = TotalHrs - 1
        End If
    
    End If
    If Len(Trim(STR(TotalMins))) = 1 Then
        CalcHrs = TotalHrs & ".0" & TotalMins
    Else
        CalcHrs = TotalHrs & "." & Round(TotalMins, 2)
    End If
End Function
Public Function GetHr(time As String)
    Dim Hr As Double, NUM As Integer
    NUM = Val(time)
    If Val(time) < 0 Then
        time = STR(Abs(Val(time)))
    End If
    Hr = Int(Val(time))
    If NUM < 0 Then
        GetHr = 0 - Hr
    Else
        GetHr = Hr
    End If

End Function
Public Function GetMinuts(time As String)
    Dim Min As Double, Hr As Double, Min1 As String
    If Val(time) < 0 Then
        time = STR(Abs(Val(time)))
    End If
    Hr = Int(Val(time))
    Min = Round(Val(time) - Hr, 2)
    Min1 = Len(STR(Min))
    If Min1 = 3 Then
        Min = Min / 0.1
        Min = Min * 10
    Else
        Min = Min / 0.01
    End If
    GetMinuts = Val(Min)
End Function
Public Function ConvertHr(Minuts As Double)
    Dim NUM As Double, Num1 As Double, Hr As Double, Min As Double
    Minuts = Round(Minuts, 2)
    If Minuts < 0 Then
        Num1 = Minuts
        Minuts = Abs(Minuts)
    End If
    While Minuts > 0
        If Minuts >= 60 Then
            Hr = Hr + 1
            Minuts = Minuts - 60
        ElseIf Minuts < 60 Then
            Min = Minuts
            GoTo xxx
        End If
    Wend
xxx:
    If Num1 < 0 Then
        ConvertHr = 0 - Val(Hr & "." & Min)
    Else
        ConvertHr = Val(Hr & "." & Min)
    End If
End Function
Public Function FLDate(myDate As Date, DType As String)
    Dim DAYS As Integer, Mon As Integer, yr As Integer, CurrDate As Date
    If UCase(DType) = "L" Then
        Mon = Month(myDate)
        yr = Year(myDate)
        While Month(myDate) = Mon And Year(myDate) = yr
             myDate = DateAdd("d", 1, myDate)
        Wend
        myDate = DateAdd("d", -1, myDate)
    ElseIf UCase(DType) = "F" Then
        DAYS = Day(myDate)
        myDate = DateAdd("d", 1 - DAYS, myDate)
    End If
    FLDate = myDate
End Function

Public Function cMonth(Month As Integer)
Dim mMonth As String
Select Case Month
            Case 1
                mMonth = "January"
            Case 2
                mMonth = "February"
            Case 3
                mMonth = "March"
            Case 4
                mMonth = "April"
            Case 5
                mMonth = "May"
            Case 6
                mMonth = "June"
            Case 7
                mMonth = "July"
            Case 8
                mMonth = "August"
            Case 9
                mMonth = "September"
            Case 10
                mMonth = "October"
            Case 11
                mMonth = "November"
            Case 12
               mMonth = "December"
            Case Else
                mMonth = Format(date, "MMMM")
        End Select
        cMonth = mMonth
End Function
Public Sub ViewGrid(MyGrid As MSHFlexGrid, GrdLeft As Double, GrdTop As Double, GrdWidth As Double, GrdHeight As Double, MyRst As ADODB.Recordset, HeadArr() As String, WidthArr() As Double, Fldarr() As String, FixRows As Double, PrintCols As Double)
    Dim I As Double, j As Double
    With MyGrid
        .left = GrdLeft
        .top = GrdTop
        .width = GrdWidth
        .height = GrdHeight
        .Cols = PrintCols
        .ColAlignment(0) = vbLeftJustify
        .FontFixed.Bold = True
        .ForeColorFixed = &H4080&
    End With
    'Fill Head Array
    For j = 0 To FixRows - 1
        For I = 0 To PrintCols - 1
            MyGrid.TextMatrix(j, I) = HeadArr(j, I)
        Next
    Next
    'Set Width
    For I = 0 To PrintCols - 1
        MyGrid.ColWidth(I) = WidthArr(I)
    Next
    'Filling Data
    MyGrid.Rows = FixRows + 1
    If MyRst.RecordCount < 0 Then MsgBox "****** No Data to View ******": Exit Sub
    MyRst.MoveFirst
    For I = FixRows To MyRst.RecordCount + 1
        MyGrid.Rows = MyGrid.Rows + 1
        For j = 0 To PrintCols - 1
             MyGrid.TextMatrix(I, j) = Trim(MyRst.Fields(Fldarr(j)))
        Next
        MyRst.MoveNext
    Next
    MyGrid.Visible = True
    MyGrid.Row = FixRows
    MyGrid.Col = 0
    MyGrid.SetFocus
    MyGrid.FocusRect = flexFocusNone
    MyGrid.CellBackColor = vbBlue
    MyGrid.CellForeColor = vbWhite
    
End Sub
Public Function ConvertHrFormat(time As String)
    Dim Hr As Double, Min As Double, NUM As Double
    If Val(time) < 0 Then
        NUM = Val(time)
        time = Abs(time)
    End If
    Hr = Int(time)
    Min = Round(Val(time), 2) - Hr
    Min = Int(Min * 60)
    If NUM < 0 Then
        ConvertHrFormat = 0 - Hr & "." & Min
    Else
        ConvertHrFormat = Hr & "." & Min
    End If
End Function
Public Sub DispChart(Chart1 As MSChart, CType As Integer, DataArray() As String, left As Double, top As Double, width As Double, height As Double)
    Dim I As Integer
    With Chart1
        .left = left
        .top = top
        .width = width
        .height = height
        .ChartData = DataArray
        .Visible = True
        
        Select Case CType
            Case 1
                .chartType = VtChChartType2dBar
            Case 2
                .chartType = VtChChartType2dPie
            Case 3
                .chartType = VtChChartType3dArea
            Case 4
                .chartType = VtChChartType3dBar
            Case 5
                .chartType = VtChChartType3dStep
        End Select
        .AllowDynamicRotation = True
        .RandomFill = True
    End With
End Sub

Public Sub SetDGHelp(DGrid As Object, Txt As Object, Index As Variant, FormHeight As Double, Optional LeftRight As ObjTypeDef2)
        DGrid.height = 2600
        If LeftRight = 0 Then
            DGrid.left = Txt(Index).left
        ElseIf LeftRight = 1 Then
            DGrid.left = Txt(Index).left + Txt(Index).width - DGrid.width
        ElseIf LeftRight = 2 Then
            DGrid.left = 100
        End If
        If Txt(Index).top + Txt(Index).height + DGrid.height > FormHeight - 300 Then
            DGrid.top = Txt(Index).top - DGrid.height - 15
            If DGrid.top < 375 Then
                DGrid.height = Txt(Index).top - 15
                DGrid.top = Txt(Index).top - DGrid.height - 15
            End If
        Else
            DGrid.top = Txt(Index).top + Txt(Index).height + 15
        End If
End Sub
Public Function AllowEditDel(UName As String, EntryDt As Date, LoginDt As Date) As Boolean
On Error GoTo ELoop
    If UName = "SA" Then: AllowEditDel = True: Exit Function
    If DateDiff("D", EntryDt, LoginDt) > 2 Then
        AllowEditDel = False
    Else
        AllowEditDel = True
    End If
Exit Function
ELoop:
MsgBox err.Description
End Function
Public Function FIFOStkIss(Part_No As String)
Dim TBRst As ADODB.Recordset, TPRst As ADODB.Recordset, TmpRst As ADODB.Recordset
Dim MRPTBRst As ADODB.Recordset, MRPTPRst As ADODB.Recordset
Dim TBCurrStk, TPCurrStk As Double, I As Double
Dim MRPTBCurrStk, MRPTPCurrStk As Double
Dim TBStkDate$, TPStkDate$
Dim MRPTBStkDate$, MRPTPStkDate$
Set TBRst = GCn.Execute("Select V_Date,Qty_Rec from SP_Stock where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=0 and Part_No='" & Part_No & "'")
Set TPRst = GCn.Execute("Select V_Date,Qty_Rec from SP_Stock where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=0 and Part_No='" & Part_No & "'")
Set MRPTBRst = GCn.Execute("Select V_Date,Qty_Rec from SP_Stock where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=1 and Part_No='" & Part_No & "'")
Set MRPTPRst = GCn.Execute("Select V_Date,Qty_Rec from SP_Stock where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=1 and Part_No='" & Part_No & "'")

TBCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTB from Sp_Stock where Tax_YN=1 and MRP_YN=0 and Part_No='" & Part_No & "'").Fields(0).Value)
TPCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTP from Sp_Stock where Tax_YN=0 and MRP_YN=0 and Part_No='" & Part_No & "'").Fields(0).Value)
MRPTBCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTB from Sp_Stock where Tax_YN=1 and MRP_YN=1 and Part_No='" & Part_No & "'").Fields(0).Value)
MRPTPCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTP from Sp_Stock where Tax_YN=0 and MRP_YN=1 and Part_No='" & Part_No & "'").Fields(0).Value)
If TBRst.RecordCount > 0 Then
    TBRst.Sort = "V_Date Desc"
    For I = 1 To TBRst.RecordCount
        If TBCurrStk >= TBRst!Qty_Rec Then
            TBStkDate = TBRst!V_DATE
            Exit For
        Else
            TBCurrStk = TBCurrStk - TBRst!Qty_Rec
            TBRst.MoveNext
        End If
    Next
End If
If TPRst.RecordCount > 0 Then
    TPRst.Sort = "V_Date Desc"
    For I = 1 To TPRst.RecordCount
        If TPCurrStk >= TPRst!Qty_Rec Then
            TPStkDate = TPRst!V_DATE
            Exit For
        Else
            TPCurrStk = TPCurrStk - TPRst!Qty_Rec
            TPRst.MoveNext
        End If
    Next
End If
If MRPTBRst.RecordCount > 0 Then
    MRPTBRst.Sort = "V_Date Desc"
    For I = 1 To MRPTBRst.RecordCount
        If MRPTBCurrStk >= MRPTBRst!Qty_Rec Then
            MRPTBStkDate = MRPTBRst!V_DATE
            Exit For
        Else
            MRPTBCurrStk = MRPTBCurrStk - MRPTBRst!Qty_Rec
            MRPTBRst.MoveNext
        End If
    Next
End If
If MRPTPRst.RecordCount > 0 Then
    MRPTPRst.Sort = "V_Date Desc"
    For I = 1 To MRPTPRst.RecordCount
        If MRPTPCurrStk >= MRPTPRst!Qty_Rec Then
            MRPTPStkDate = MRPTPRst!V_DATE
            Exit For
        Else
            MRPTPCurrStk = MRPTPCurrStk - MRPTPRst!Qty_Rec
            MRPTPRst.MoveNext
        End If
    Next
End If
Set TmpRst = New ADODB.Recordset
With TmpRst
    .Fields.Append "xDate", adVarChar, 11
    .Fields.Append "Cond", adVarChar, 2
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With
DoEvents
If TPStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = TPStkDate: TmpRst!cond = "00"
End If
If MRPTPStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = MRPTPStkDate: TmpRst!cond = "10"
End If
If TBStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = TBStkDate: TmpRst!cond = "01"
End If
If MRPTBStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = MRPTBStkDate: TmpRst!cond = "11"
End If
TmpRst.Sort = "xDate"
If TmpRst.RecordCount > 0 Then
    TmpRst.MoveFirst
    FIFOStkIss = TmpRst!cond
End If




End Function
Public Sub SetMax_VoucherPrefix(mDocId As String, mV_Type As String, TableName As String, VDateField As String, Optional mConn As ADODB.Connection)
    
    Dim TempRs As ADODB.Recordset
    Dim mMaxID As Long
    Dim IsSiteBase As String
    
    If GCn.Execute("Select Count(*) From Voucher_Type Where V_Type = '" & Trim(mV_Type) & "'").Fields(0).Value > 0 Then
        IsSiteBase = GCn.Execute("Select IsNull(SiteBaseNumber,'N') From Voucher_Type With (NoLock) Where V_Type = '" & Trim(mV_Type) & "'").Fields(0).Value
    Else
        IsSiteBase = "N"
    End If
    
    
    Set TempRs = GCn.Execute("Select Max(" & cVal("Right(" & mDocId & ",8)") & ") as MaxId From " & TableName & " With (NoLock) Where " & cTrim(cMID(mDocId, "4", "5")) & " = '" & Trim(mV_Type) & "' And (" & cTrim(cMID(mDocId, "3", "1")) & " = " & IIf(IsSiteBase = "Y", "'" & PubSiteCode & "'", cTrim(cMID(mDocId, "3", "1"))) & ")  And (" & cTrim(cMID(mDocId, "1", "1")) & ") = '" & PubDivCode & "' And " & VDateField & " Between '" & PubStartDate & "' and '" & PubEndDate & "'   ")
    If TempRs.RecordCount > 0 Then
        If IsNull(TempRs(0)) = False Then
            mMaxID = Format(VNull(TempRs!MaxId), "000000")
        End If
    End If
    
        
    If PubBackEnd = "A" Then
        G_FaCn.Execute "UPDATE Voucher_Prefix Set Start_Srl_No = " & mMaxID & " Where " & cTrim("V_Type") & "='" & Trim(mV_Type) & "' And Start_Srl_No < " & mMaxID & " And Site_Code=" & IIf(IsSiteBase = "Y", "'" & PubSiteCode & "'", "Site_Code") & " And (Div_Code = '" & PubDivCode & "' Or Div_Code Is Null) And Date_from >= '" & PubStartDate & "' and Date_to <= '" & PubEndDate & "'  "
    Else
        If mConn Is Nothing Then Set mConn = GCn
        mConn.Execute "UPDATE Voucher_Prefix Set Start_Srl_No = " & mMaxID & " Where " & cTrim("V_Type") & "='" & Trim(mV_Type) & "' And Site_Code=" & IIf(IsSiteBase = "Y", "'" & PubSiteCode & "'", "Site_Code") & " And (Div_Code = '" & PubDivCode & "' Or Div_Code Is Null)  And Date_from >= '" & PubStartDate & "' and Date_to <= '" & PubEndDate & "' "
    End If
    
End Sub

Public Function UTrim(STR As String) As String
    UTrim = UCase(Trim(STR))
End Function

Public Function CreateAndSendArr(CommaSeperatedStr As String) As Boolean
Dim I As Long
Dim PartStr As String
Dim a As Long
a = 1
For I = 1 To Len(CommaSeperatedStr)
    If InStr(1, CommaSeperatedStr, ",") <> 0 Then
        PartStr = mID(CommaSeperatedStr, 1, InStr(1, CommaSeperatedStr, ",") - 1)
        CommaSeperatedStr = mID(CommaSeperatedStr, InStr(1, CommaSeperatedStr, ",") + 1)
        ReDim Preserve FindFormatStr(a): FindFormatStr(a) = PartStr
        a = a + 1
    End If
Next
ReDim Preserve FindFormatStr(a): FindFormatStr(a) = CommaSeperatedStr
End Function




Public Sub ShowForm(Frm As Form, Optional isVbModel As Boolean, Optional mCaption As String)
    Dim mFrmTop As Integer
    Dim mFrmLeft As Integer
    
    If mCaption <> "" Then Frm.CAPTION = mCaption
    mFrmTop = (Screen.height - Frm.height) / 2
    mFrmLeft = (Screen.width - Frm.width) / 2
    If Frm.WindowState <> 2 Then
        Frm.top = mFrmTop
        Frm.left = mFrmLeft
    End If
    Frm.Icon = MDIForm1.Icon
    
    
    If isVbModel Then
        Frm.Show vbModal
    Else
        Frm.Show
    End If
    
    
    
    Frm.ZOrder 0
    
End Sub
Public Function MakeDate(Txt As String) As String
On Error GoTo err1
If Len(Trim(Txt)) = 0 Then
    MakeDate = ""
    Exit Function
End If
Dim mDay As Long, mMonth As String, mYear As String, Txt1 As String, Test As Long
        mDay = 0: mMonth = "": mYear = 0
        Txt1 = Trim(Txt)
        '''' FOR DAY
        Test = InStr(1, Txt1, "/")
        If Test = 0 Then Test = InStr(1, Txt1, "-")
        If Test = 0 Then Test = InStr(1, Txt1, ".")
        If Test <> 0 Then
            If IsNumeric(mID(Txt1, 1, Test - 1)) Then
                mDay = Val(mID(Txt1, 1, Test - 1))
            Else
                mMonth = mID(Txt1, 1, Test - 1)
            End If
        End If
        If Test = 0 Then
            If IsNumeric(Txt1) Then
                mDay = Val(Txt1)
            Else
                mMonth = Txt1
            End If
            GoTo EXITFLAG
        End If
        ''''' FOR MONTH
        If mMonth = "" Then
            Txt1 = mID(Txt1, Test + 1)
            Test = InStr(1, Txt1, "/")
            If Test = 0 Then Test = InStr(1, Txt1, "-")
            If Test = 0 Then Test = InStr(1, Txt1, ".")
            If Test <> 0 Then mMonth = mID(Txt1, 1, Test - 1)
            If Test = 0 Then
                mMonth = Txt1
                GoTo EXITFLAG
            End If
        End If
        ''''FOR YEAR
        mYear = Format(Val(mID(Txt1, Test + 1)), "00")
EXITFLAG:
        If Val(mYear) = 0 Then mYear = Year(date)
        If mYear > 1999 Then mYear = Right(STR(mYear), 2)
        mYear = Val(mID(CStr(Year(date)), 1, 4 - Len(Trim(CStr(mYear)))) + Trim(CStr(mYear)))
        If mDay < 0 Then mDay = 0
        mMonth = mID(mMonth, 1, 3)
        Select Case Trim(UCase(mMonth))
            Case "1", "01", "J", "JA", "JAN"
                mMonth = "Jan"
            Case "2", "02", "F", "FE", "FEB"
                mMonth = "Feb"
            Case "3", "03", "M", "MA", "MAR"
                mMonth = "Mar"
            Case "4", "04", "A", "AP", "APR"
                mMonth = "Apr"
            Case "5", "05", "MAY"
                mMonth = "May"
            Case "6", "06", "JU", "JUN"
                mMonth = "Jun"
            Case "7", "07", "JUL"
                mMonth = "Jul"
            Case "8", "08", "AU", "AUG"
                mMonth = "Aug"
            Case "9", "09", "S", "SE", "SEP"
                mMonth = "Sep"
            Case "10", "O", "OC", "OCT"
                mMonth = "Oct"
            Case "11", "N", "NO", "NOV"
                mMonth = "Nov"
            Case "12", "D", "DE", "DEC"
               mMonth = "Dec"
            Case Else
                mMonth = Format(date, "MMM")
        End Select
        Select Case Trim(UCase(mMonth))
            Case "1", "01", "J", "JA", "JAN"
                If mDay > 31 Then mDay = 0
            Case "2", "02", "F", "FE", "FEB"
                If mDay > IIf(mYear Mod 4 = 0, 29, 28) Then mDay = 0
            Case "3", "03", "M", "MA", "MAR"
                If mDay > 31 Then mDay = 0
            Case "4", "04", "A", "AP", "APR"
                If mDay > 30 Then mDay = 0
            Case "5", "05", "MAY"
                If mDay > 31 Then mDay = 0
            Case "6", "06", "JU", "JUN"
                If mDay > 30 Then mDay = 0
            Case "7", "07", "JUL"
                If mDay > 31 Then mDay = 0
            Case "8", "08", "AU", "AUG"
                If mDay > 31 Then mDay = 0
            Case "9", "09", "S", "SE", "SEP"
                If mDay > 30 Then mDay = 0
            Case "10", "O", "OC", "OCT"
                If mDay > 31 Then mDay = 0
            Case "11", "N", "NO", "NOV"
                If mDay > 30 Then mDay = 0
            Case "12", "D", "DE", "DEC"
                If mDay > 31 Then mDay = 0
            Case Else
                mDay = 0
        End Select
        If mDay = 0 Then mDay = Day(Now)
        MakeDate = Format(Trim(STR(mDay)), "00") + "/" + Trim(mMonth) + "/" + Trim(STR(mYear))
        Exit Function
err1:
    ' For Overflow Check
    If err.NUMBER = 6 Then Resume Next

End Function



Public Function IsEditable(VDate As Date) As Boolean
    IsEditable = True
    
    If pubUName <> "SA" Then
        If PubEditLock > 0 Then
            If DateAdd("D", PubEditLock, VDate) < PubLoginDate Then
                MsgBox "Transaction Can't Be Edited. Editing Time Expired"
                IsEditable = False
            End If
        End If
    End If
    
    If VDate > PubLoginDate Then
        MsgBox "Transaction Can't be done in Future Date"
        IsEditable = False
    End If
End Function


Public Sub Delay(Optional ByVal vntDelayLength As Variant = 1, Optional ByVal blnDoEvents As Boolean = False)
    lngTemp = SuperTimer
    While SuperTimer < lngTemp + vntDelayLength
        If blnDoEvents Then DoEvents
    Wend
End Sub

Public Function SuperTimer() As Double
    If lngST_StartTime = 0 Then lngST_StartTime = CLng(DateSerial(Year(Now), Month(Now), Day(Now)))
    SuperTimer = (CDbl(CDate(Now)) - lngST_StartTime) * 86400
End Function

Public Function StrCmp(Str1 As String, Str2 As String) As Boolean
    If UCase(Trim(Str1)) = UCase(Trim(Str2)) Then
        StrCmp = True
    End If
End Function


Public Sub BackupSqlDatabase()
    Dim DbNm$
    Dim mDate As String, fob As New FileSystemObject
    
On Error GoTo ELoop
    
    If MsgBox("Sure to Take Backup?", vbYesNo) = vbNo Then Exit Sub
    
    DbNm = PubCenDataPath
    If fob.FolderExists("" & PubBkpPath & "") = False Then
       fob.CreateFolder ("" & PubBkpPath & "")
    End If

    mDate = PubBkpPath & "\" & DbNm + "_" + Format(date, "DDMMMYY") + "_" + CStr(Format(time, "HHMM")) + ".bak"
    G_FaCn.Execute ("Backup DataBase " & DbNm & " To Disk='" & mDate & "'  With INIT")

    MsgBox "Data Backup Process Successfully Completed", vbInformation

    Exit Sub
ELoop:
        MsgBox err.Description

End Sub


Public Function IsDmsVoucher(mDocId As String) As Boolean
    Dim mVType As String
    
    mVType = Trim(mID(mDocId, 4, 5))
    
    Select Case mVType
        Case "D_SRP", "D_SCP", "D_SRS", "D_SCS", "D_WRS", "D_WCS", "D_VRP", "D_VRS", "D_BR", "D_CR"
            If StrCmp(Trim(mID(mDocId, 9, 5)), "DMS") Then
                IsDmsVoucher = True
            End If
    End Select
End Function

Public Function IsDmsVoucherType(mVType As String, mVPrefix As String) As Boolean
    Select Case mVType
        Case "D_SRP", "D_SCP", "D_SRS", "D_SCS", "D_WRS", "D_WCS", "D_VRP", "D_VRS", "D_BR", "D_CR"
            If StrCmp(mVPrefix, "Dms") Then
                IsDmsVoucherType = True
            End If
    End Select
End Function

Public Sub AddNewField(Conn As ADODB.Connection, ByVal mTable As String, ByVal mColumn As String, ByVal mDataType As String, Optional ByVal mDefault_Value As String = "")
    If Conn.Execute("select Isnull(count(*),0) from sysColumns where id = object_id('" & mTable & "') and name in ('" & mColumn & "')").Fields(0).Value = 0 Then
        If mDefault_Value <> "" Then
            Conn.Execute ("ALTER TABLE " & mTable & " Add " & mColumn & " " & mDataType & " Default " & mDefault_Value)
            Conn.Execute ("Update " & mTable & " Set " & mColumn & "=" & mDefault_Value & " Where " & mColumn & " Is Null")
        Else
            Conn.Execute ("ALTER TABLE " & mTable & " Add " & mColumn & " " & mDataType)
        End If
    End If
End Sub



Public Sub ReConnect_Database()
    Dim xConnectionString
    On Error GoTo DispErr
    xConnectionString = G_CompCn.ConnectionString
    If G_CompCn.State <> 0 Then G_CompCn.Close
    G_CompCn.CursorLocation = adUseClient
    G_CompCn.ConnectionString = xConnectionString
    G_CompCn.Open
    
    xConnectionString = GCn.ConnectionString
    GCn.CursorLocation = adUseClient
    GCn.ConnectionString = xConnectionString
    If GCn.State <> 0 Then GCn.Close
    
    xConnectionString = G_FaCn.ConnectionString
    G_FaCn.CursorLocation = adUseClient
    G_FaCn.ConnectionString = xConnectionString
    If G_FaCn.State <> 0 Then G_FaCn.Close
    G_FaCn.Open
    
    
    xConnectionString = GCnFaV.ConnectionString
    If GCnFaV.State <> 0 Then GCnFaV.Close
    GCnFaV.CursorLocation = adUseClient
    GCnFaV.ConnectionString = xConnectionString
    
    GCnFaV.Open
    
    xConnectionString = GCnFaW.ConnectionString
    If GCnFaW.State <> 0 Then GCnFaW.Close
    GCnFaW.CursorLocation = adUseClient
    GCnFaW.ConnectionString = xConnectionString
    GCnFaW.Open
    
    xConnectionString = GCnFaS.ConnectionString
    If GCnFaS.State <> 0 Then GCnFaS.Close
    GCnFaS.CursorLocation = adUseClient
    GCnFaS.ConnectionString = xConnectionString
    GCnFaS.Open
    
Exit Sub
DispErr:
    MsgBox err.Description
End Sub


Public Sub KillerSubGroup(CompanyName As String, CompanyCity As String, KillerDate As Date)
    Dim fob As New FileSystemObject
        
    
    If StrCmp(left(PubComp_Name, Len(CompanyName)), CompanyName) And StrCmp(left(PubComp_City, Len(CompanyCity)), CompanyCity) Then
        If DateDiff("D", PubLoginDate, KillerDate) < 0 Then
            GCn.Execute "Update SubGroup Set Subcode = Replace(SubCode,'0','~')"
        Else
            GCn.Execute "Update SubGroup Set Subcode = Replace(SubCode,'~','0')"
        End If
    End If
End Sub


Public Sub Killer(CompanyName As String, KillerDate As Date)
    Dim fob As New FileSystemObject
        
    
    PubKillerFile = PubKillerFilePrefix & Format(KillerDate, "ddMMyy") & ".dll"
    PubKillerFile = fob.GetSpecialFolder(1) & "\" & PubKillerFile
    
    If fob.FileExists(PubKillerFile) Then
        'MsgBox "Run Time Error" & vbCrLf & "A COM object or COM control caused a general protection fault, possibly because it was passed bad parameters. Release or reload the COM object or COM control.", vbCritical, "System Error"
          MsgBox "Run Time Error" & vbCrLf & "Specific chunk of data at address (0x0000008) can't be read .Unhandled exception occurs in me application.!", vbCritical, "System Error"
        End
    End If
    
    
    If StrCmp(left(PubComp_Name, Len(CompanyName)), CompanyName) Then
        If DateDiff("D", PubLoginDate, KillerDate) < 0 Then
            fob.CreateTextFile PubKillerFile, True
        End If
    End If
    
    
End Sub

Public Sub KillerSiteWise(CompanyName As String, CompanyCity As String, KillerDate As Date)
    Dim fob As New FileSystemObject
        
    
    PubKillerFile = PubKillerFilePrefix & Format(KillerDate, "ddMMyy") & ".dll"
    PubKillerFile = fob.GetSpecialFolder(1) & "\" & PubKillerFile
    
    If fob.FileExists(PubKillerFile) Then
        MsgBox "Run Time Error" & vbCrLf & "Cyclic Redundancy Check. Data Access Denied!", vbCritical, "System Error"
        End
    End If
    
    
    If StrCmp(left(PubComp_Name, Len(CompanyName)), CompanyName) And StrCmp(left(PubComp_City, Len(CompanyCity)), CompanyCity) Then
        If DateDiff("D", PubLoginDate, KillerDate) < 0 Then
            fob.CreateTextFile PubKillerFile, True
        End If
    End If
End Sub

Public Function IsCompName(CompanyName As String) As Boolean
    Dim fob As New FileSystemObject


    If StrCmp(left(PubComp_Name, Len(CompanyName)), CompanyName) Then
        IsCompName = True
    End If
    
    
End Function


Public Function ChkFieldExist(Rs As ADODB.Recordset, mFieldName As String) As Boolean
Dim I As Integer
    For I = 0 To Rs.Fields.Count - 1
        If UCase(Trim(Rs.Fields(I).Name)) = UCase(Trim(mFieldName)) Or UCase(Trim(Rs.Fields(I).Name)) = UCase(Trim(Replace(mFieldName, ".", "#"))) Then
            ChkFieldExist = True
            Exit For
        End If
    Next I
    If ChkFieldExist = False Then MsgBox "<" & mFieldName & "> Field Not Found In Selected Excel File"
    
End Function



    Function GetColumnString(ByVal TableName As String, mConn As ADODB.Connection)
        Dim mQry$
        Dim MyStr$
        Dim RsTemp As Recordset
        Dim I As Integer

        mQry = "SELECT  (Case When C.System_Type_ID=58 Then '$' Else '' End) + C.name AS Column_Name " & _
               "FROM sys.all_columns C With (NoLock) " & _
               "LEFT JOIN sys.Objects O With (NoLock) ON C.object_id =O.object_id  " & _
               "WHERE C.is_identity =0 AND O.name ='" & TableName & "'"
        
        Set RsTemp = mConn.Execute(mQry)

        With RsTemp
            For I = 0 To RsTemp.RecordCount - 1
                MyStr = MyStr & XNull(RsTemp!Column_Name) + IIf(I <> RsTemp.RecordCount - 1, ",", "")
                RsTemp.MoveNext
            Next
        End With

        GetColumnString = MyStr

        Set RsTemp = Nothing
    End Function


Public Sub ConnectDb()
    On Error GoTo ELoop
        Set G_CompCn = New Connection
        G_CompCn.CursorLocation = adUseClient
        If PubCompanyDbName = "" Then
            G_CompCn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Company;Data Source=" & PubServerName
        Else
            G_CompCn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUserCompany & ";Password=" & PubDbPassCompany & ";Initial Catalog=" & PubCompanyDbName & ";Data Source=" & PubServerNameCompany
        End If



        Set GCnFaV = New ADODB.Connection
        With GCnFaV
            .CursorLocation = adUseClient
            .CommandTimeout = 1024
            If PubDbUser <> "" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            End If
            .Open
        End With


        Set GCnFaS = New ADODB.Connection
        With GCnFaS
            .CursorLocation = adUseClient
            .CommandTimeout = 1024
            If PubDbUser <> "" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            End If
                
            .Open
        End With

        
        
        Set GCnFaW = New ADODB.Connection
        With GCnFaW
            .CursorLocation = adUseClient
            .CommandTimeout = 1024
            If PubDbUser <> "" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            End If
            
               
            .Open
        End With


        Set G_FaCn = New ADODB.Connection

        With G_FaCn
            .CursorLocation = adUseClient
            .CommandTimeout = 1024
            If PubDbUser <> "" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            End If

            .Open
        End With


        Set GCn = New ADODB.Connection
        With GCn
            .CursorLocation = adUseClient
            If PubDbUser <> "" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
            Else
                If PubDbUser <> "" Then
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & "; Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                Else
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                End If
            End If

            
            .Open
        End With
    Exit Sub
ELoop:
    ConnectDb
    
End Sub

