Attribute VB_Name = "AccessToSqlServer"
Option Explicit

Public Function vIsNull(FieldName As String, ValueIfNull) As String
    If PubBackEnd = "A" Then
        vIsNull = "IIF(IsNull(" & FieldName & ")," & ValueIfNull & "," & FieldName & ")"
    Else
        vIsNull = "IsNull(" & FieldName & "," & ValueIfNull & ")"
    End If
End Function


Public Function xIsNull(FieldName As String, ValueIfNull) As String
    If PubBackEnd = "A" Then
        xIsNull = "IIF(IsNull(" & FieldName & "),'" & ValueIfNull & "'," & FieldName & ")"
    Else
        xIsNull = "IsNull(" & FieldName & ",'" & ValueIfNull & "')"
    End If
End Function

Public Function ConvertDate(Temp)

    If IsNull(Temp) = False And Temp <> "" Then Temp = IIf(Year(Temp) >= 2050 Or Year(Temp) <= 1900, Format(Day(Temp) & "/" & Month(Temp) & "/" & Year(date), "dd/mmm/yyyy"), Temp)

    If PubBackEnd = "A" Then
        If IsNull(Temp) Or Temp = "" Or Temp = Null Then
        '31-01-2001
        ConvertDate = "Null"
        Else
        ConvertDate = "#" & Format(CDate(Temp), "dd/MMM/yyyy") & "#"
        End If
    ElseIf PubBackEnd = "S" Then
            If IsNull(Temp) Or Temp = "" Then
        '31-01-2001
            ConvertDate = "Null"
            Else
             ConvertDate = "'" & Format(CDate(Temp), "dd/MMM/yyyy") & "'"
            End If
    End If
End Function


Public Function ConvertDateTime(Temp)

    If IsNull(Temp) = False And Temp <> "" Then Temp = IIf(Year(Temp) >= 2050 Or Year(Temp) <= 1900, Format(Day(Temp) & "/" & Month(Temp) & "/" & Year(date), "dd/mmm/yyyy"), Temp)

    If PubBackEnd = "A" Then
        If IsNull(Temp) Or Temp = "" Or Temp = Null Then
            ConvertDateTime = "Null"
        Else
            ConvertDateTime = "#" & CDate(Temp) & "#"
        End If
    ElseIf PubBackEnd = "S" Then
        If IsNull(Temp) Or Temp = "" Then
        '31-01-2001
            ConvertDateTime = "Null"
        Else
             ConvertDateTime = "'" & CDate(Temp) & "'"
        End If
    End If
End Function


Public Function cBoolean(ByVal Temp As Boolean) As Variant
    If PubBackEnd = "A" Then
        If Temp = False Then
            cBoolean = False
        Else
            cBoolean = True
        End If
    ElseIf PubBackEnd = "S" Then
        If Temp = False Then
            cBoolean = 0
        Else
            cBoolean = 1
        End If
       End If
End Function

Public Function cTime(Temp)
    If PubBackEnd = "A" Then
    If IsNull(Temp) Or Temp = "" Or Temp = Null Then
        '31-01-2001
        cTime = "Null"
    Else
        cTime = "#" & Format(CDate(Temp), "Short Time") & "#"
    End If
    ElseIf PubBackEnd = "S" Then
    If IsNull(Temp) Or Temp = "" Or Temp = Null Then
        '31-01-2001
        cTime = "Null"
    Else
        cTime = "'" & Format(CDate(Temp), "Short Time") & "'"
    End If
    
    End If
End Function


Public Function cUCase(FieldName As String) As String
    If PubBackEnd = "A" Then
        cUCase = "UCase(" & FieldName & ")"
    ElseIf PubBackEnd = "S" Then
        cUCase = "Upper(" & FieldName & ")"
    End If
End Function

Public Function cMID(FieldName As String, StartFrom As String, LenStr As String) As String
    If PubBackEnd = "A" Then
        cMID = "MID(" & FieldName & "," & StartFrom & ", " & LenStr & ")"
    ElseIf PubBackEnd = "S" Then
        cMID = "SubString(" & FieldName & "," & StartFrom & ", " & LenStr & ")"
    End If
End Function

Public Function cVal(FieldName As String) As String
    If PubBackEnd = "A" Then
        cVal = "Val(" & FieldName & ")"
    ElseIf PubBackEnd = "S" Then
        'cVal = "Cast(" & FieldName & " As Numeric)"
        cVal = "Convert(Numeric," & FieldName & ") "
    End If
End Function

Public Function cCStr(FieldName As String, Optional Length As Integer) As String
    If PubBackEnd = "A" Then
        cCStr = "CStr(" & FieldName & ")"
    ElseIf PubBackEnd = "S" Then
        If Length > 0 Then
            cCStr = "Convert(nVarChar(" & Length & ")," & FieldName & ")"
        Else
            cCStr = "Convert(nVarChar," & FieldName & ")"
        End If
    End If
End Function

Public Function cIIF(mCondition As String, StrIfTrue As String, Optional StrIfFalse As String) As String
    If PubBackEnd = "A" Then
        If StrIfFalse = "" Then
            cIIF = "IIF(" & mCondition & "," & StrIfTrue & ")"
        Else
            cIIF = "IIF(" & mCondition & "," & StrIfTrue & ", " & StrIfFalse & ")"
        End If
    ElseIf PubBackEnd = "S" Then
        If StrIfFalse = "" Then
            cIIF = "(Case When " & mCondition & " Then " & StrIfTrue & "  End)"
        Else
            cIIF = "(Case When " & mCondition & " Then " & StrIfTrue & " Else " & StrIfFalse & " End)"
        End If
    End If
End Function


Public Function cTrim(FieldName As String) As String
    If PubBackEnd = "A" Then
        cTrim = "Trim(" & FieldName & ")"
    ElseIf PubBackEnd = "S" Then
        cTrim = "LTrim(RTrim(" & FieldName & "))"
    End If
End Function

Public Function VFormat(NumericValue As Variant, PrecisionFormat As String) As String
    If PubBackEnd = "A" Then
        VFormat = "Format(" & NumericValue & ", '" & PrecisionFormat & "' )"
    ElseIf PubBackEnd = "S" Then
        VFormat = NumericValue
    End If
End Function

Public Function FaTable(FaTableName As String) As String
    If PubBackEnd = "A" Then
        FaTable = "[" & PubVFADataPath & "]." & FaTableName
    Else
        FaTable = FaTableName
    End If
End Function

Public Function cDt(FieldName As String) As String
    If PubBackEnd = "A" Then
        cDt = "cstr(IIF(ISNULL(" & FieldName & "),''," & FieldName & "))"
    Else
        cDt = "Convert(nVarChar(15)," & FieldName & ",3)"
    End If
End Function

Public Function cMth(FieldName As String) As String
    If PubBackEnd = "A" Then
        cMth = "Month(" & FieldName & ")"
    ElseIf PubBackEnd = "S" Then
        cMth = "DatePart(Month," & FieldName & ")"
    End If
End Function






 
