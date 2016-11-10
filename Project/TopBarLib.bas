Attribute VB_Name = "TopBarLib"
Public Function BUTTONS(Enb As Boolean, mFORM As Form, mado As ADODB.Recordset, navSig As Integer) As Boolean
On Error GoTo BTNERR
Dim mPRV As Boolean, mNEXT As Boolean

    Select Case navSig
        Case 1
            If mado.BOF = False Then mado.MoveFirst
        Case 2
            If mado.BOF = False Then mado.MovePrevious
        Case 3
            If mado.EOF = False Then mado.MoveNext
        Case 4
            If mado.EOF = False Then mado.MoveLast
    End Select
    If (mado.RecordCount <= 0) Or (Not Enb) Then
        mFORM.TopCtrl1.tFirst = False
        mFORM.TopCtrl1.tPrev = False
        mFORM.TopCtrl1.tNext = False
        mFORM.TopCtrl1.tLast = False
        mFORM.TopCtrl1.tFind = False
        mFORM.TopCtrl1.tDel = False
        mFORM.TopCtrl1.tEdit = False
        mFORM.TopCtrl1.tPrn = False
    Else
        If mado.BOF Or mado.AbsolutePosition = 1 Then
            mPRV = False
        Else
            mPRV = True
        End If
        
        If mado.EOF Or mado.AbsolutePosition = mado.RecordCount Then
            mNEXT = False
        Else
            mNEXT = True
        End If
        mFORM.TopCtrl1.tFirst = mPRV
        mFORM.TopCtrl1.tPrev = mPRV
        mFORM.TopCtrl1.tNext = mNEXT
        mFORM.TopCtrl1.tLast = mNEXT
    End If
    If mado.AbsolutePosition > 0 Then
        mFORM.TopCtrl1.TopText1 = mado.AbsolutePosition & "/" & mado.RecordCount
    Else
        mFORM.TopCtrl1.TopText1 = "0/" & mado.RecordCount
    End If
    mFORM.TopCtrl1.TopText1.ForeColor = &HFF00FF
    BUTTONS = True
Exit Function
BTNERR:
    If err.NUMBER > 0 Then MsgBox err.Description
End Function
Public Function SETS(MSET As String, mFORM As Form, mado As ADODB.Recordset) As Boolean
Dim xi As Boolean
    Select Case MSET
    Case "ADD"
        xi = Disp(False, mFORM)
        xi = BUTTONS(False, mFORM, mado, 0)
        mFORM.TopCtrl1.TopText2 = "Add"
        mFORM.TopCtrl1.TopText2.ForeColor = RGB(255, 0, 0)
        SETS = True
    Case "EDIT"
        xi = Disp(False, mFORM)
        xi = BUTTONS(False, mFORM, mado, 0)
        mFORM.TopCtrl1.TopText2 = "Edit"
        mFORM.TopCtrl1.TopText2.ForeColor = RGB(255, 0, 0)
        SETS = True
        
    Case "INI"
        xi = Disp(True, mFORM)
        xi = BUTTONS(True, mFORM, mado, 0)
        mFORM.TopCtrl1.TopText2 = "Browse"
        mFORM.TopCtrl1.TopText2.ForeColor = RGB(0, 0, 255)
        SETS = False
        If mFORM.TopCtrl1.Visible = True Then
            mFORM.TopCtrl1.SetFocus
        End If
    End Select
    mFORM.TopCtrl1.TopText1.Alignment = 1
'    If mado.AbsolutePosition  > 0 Then
'        mFORM.TopCtrl1.TopText1 = mado.AbsolutePosition & "/" & mado.RecordCount
'    Else
'        mFORM.TopCtrl1.TopText1 = "0/" & mado.RecordCount
'    End If
'    mFORM.TopCtrl1.TopText1.ForeColor = &HFF00FF
End Function
Public Function Disp(ByVal Enb As Boolean, ByRef mFORM As Form) As Boolean
    If Enb = True Then
        If InStr(mFORM.TopCtrl1.Tag, "A") <> 0 Then mFORM.TopCtrl1.tAdd = Enb Else mFORM.TopCtrl1.tAdd = Not Enb
        If InStr(mFORM.TopCtrl1.Tag, "E") <> 0 Then mFORM.TopCtrl1.tEdit = Enb Else mFORM.TopCtrl1.tEdit = Not Enb
        If InStr(mFORM.TopCtrl1.Tag, "D") <> 0 Then mFORM.TopCtrl1.tDel = Enb Else mFORM.TopCtrl1.tDel = Not Enb
        If InStr(mFORM.TopCtrl1.Tag, "P") <> 0 Then mFORM.TopCtrl1.tPrn = Enb Else mFORM.TopCtrl1.tPrn = Not Enb
    Else
        mFORM.TopCtrl1.tAdd = Enb
        mFORM.TopCtrl1.tEdit = Enb
        mFORM.TopCtrl1.tDel = Enb
        mFORM.TopCtrl1.tPrn = Enb
    End If
    mFORM.TopCtrl1.tFind = Enb
    mFORM.TopCtrl1.tSave = Not Enb
    mFORM.TopCtrl1.tCancel = Not Enb
    Disp = True
End Function
