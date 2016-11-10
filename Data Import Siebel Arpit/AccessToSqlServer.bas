Attribute VB_Name = "AccessToSqlServer"
Option Explicit
'Arpit --- Lib for Converting Automan (MsAccess) to Automan (SQL Server)

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

Public Function ConvertDate(temp)
    If PubBackEnd = "A" Then
        If IsNull(temp) Or temp = "" Or temp = Null Then
        '31-01-2001
        ConvertDate = "Null"
        Else
        ConvertDate = "#" & Format(CDate(temp), "dd/MMM/yyyy") & "#"
        End If
    ElseIf PubBackEnd = "S" Then
            If IsNull(temp) Or temp = "" Then
        '31-01-2001
            ConvertDate = "Null"
            Else
             ConvertDate = "'" & Format(CDate(temp), "dd/MMM/yyyy") & "'"
            End If
    End If
End Function
Public Function cBoolean(ByVal temp As Boolean) As Variant
    If PubBackEnd = "A" Then
        If temp = False Then
            cBoolean = False
        Else
            cBoolean = True
        End If
    ElseIf PubBackEnd = "S" Then
        If temp = False Then
            cBoolean = 0
        Else
            cBoolean = 1
        End If
       End If
End Function

Public Function cTime(temp)
    If PubBackEnd = "A" Then
    If IsNull(temp) Or temp = "" Or temp = Null Then
        '31-01-2001
        cTime = "Null"
    Else
        cTime = "#" & Format(CDate(temp), "Short Time") & "#"
    End If
    ElseIf PubBackEnd = "S" Then
    If IsNull(temp) Or temp = "" Or temp = Null Then
        '31-01-2001
        cTime = "Null"
    Else
        cTime = "'" & Format(CDate(temp), "Short Time") & "'"
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

Public Function cCStr(FieldName As String) As String
    If PubBackEnd = "A" Then
        cCStr = "CStr(" & FieldName & ")"
    ElseIf PubBackEnd = "S" Then
        cCStr = "Convert(nVarChar," & FieldName & ")"
    End If
End Function

Public Function cIIF(mCondition As String, StrIfTrue As String, StrIfFalse As String) As String
    If PubBackEnd = "A" Then
        cIIF = "IIF(" & mCondition & "," & StrIfTrue & ", " & StrIfFalse & ")"
    ElseIf PubBackEnd = "S" Then
        cIIF = "(Case When " & mCondition & " Then " & StrIfTrue & " Else " & StrIfFalse & " End)"
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
        FaTable = "[" & PubFADataPath & "]." & FaTableName
    Else
        FaTable = FaTableName
    End If
End Function

Public Function cDt(FieldName As String) As String
    If PubBackEnd = "A" Then
        cDt = "& FieldName &"
    Else
        cDt = "Convert(nVarChar," & FieldName & ",3)"
    End If
End Function



