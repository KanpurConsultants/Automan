Attribute VB_Name = "ModuleGrid"
Public Sub Text_FGrid_Get_Txt(TEXT As Object, FGrid As MSHFlexGrid, MaxCol As Byte)
Dim J As Byte
    For J = 1 To MaxCol
        FGrid.TextMatrix(FGrid.Row, J) = TEXT(J).TEXT
    Next
End Sub
Public Sub Text_FGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer, TEXT As Object, FGrid As MSHFlexGrid, MaxCol As Byte)
Dim mROW As Integer, mr As Byte
If Shift <> 2 And KeyCode = 37 And TEXT(FGrid.Col).SelStart <> 0 Then Exit Sub
If Shift <> 2 And KeyCode = 39 And TEXT(FGrid.Col).SelStart <> Len(TEXT(FGrid.Col).TEXT) Then Exit Sub
Text_FGrid_Get_Txt TEXT, FGrid, MaxCol
If Shift = 2 And KeyCode = vbKeyDelete Then         'CTL DELETE
    If FGrid.Rows >= 1 Then
        If FGrid.Rows = 2 Then
            For mr = 1 To MaxCol
                FGrid.TextMatrix(FGrid.Row, mr) = ""
            Next
        Else
            FGrid.RemoveItem (FGrid.Row)
        End If
    End If
ElseIf (Shift = 2 And KeyCode = vbKeyUp) Or KeyCode = 38 Then         'UP ARROW
    If FGrid.Row <> 1 Then If FGrid.Parent.beforeRowUpdate = True Then FGrid.Row = FGrid.Row - 1
ElseIf (Shift = 2 And KeyCode = vbKeyDown) Or KeyCode = 40 Then     'DOWN ARROW
    If FGrid.Row <> FGrid.Rows - 1 Then If FGrid.Parent.beforeRowUpdate = True Then FGrid.Row = FGrid.Row + 1
ElseIf (Shift = 2 And KeyCode = vbKeyLeft) Or KeyCode = 37 Then       'LEFT ARROW
    mr = FGrid.Col
    Do While True
        If FGrid.Row = 1 And FGrid.Col = 1 Then
            FGrid.Col = mr
            Exit Do
        End If
        If FGrid.Col = 1 Then
            If FGrid.Parent.beforeRowUpdate = False Then FGrid.Col = mr
            FGrid.Row = FGrid.Row - 1
            FGrid.Col = MaxCol
        Else
            FGrid.Col = FGrid.Col - 1
        End If
        If TEXT(FGrid.Col).GridVisible = 1 Then Exit Do
    Loop
ElseIf (Shift = 2 And KeyCode = vbKeyRight) Or KeyCode = 39 Then      'RIGHT ARROW
    mr = FGrid.Col
    Do While True
        If FGrid.Row = FGrid.Rows - 1 And (FGrid.Col = FGrid.Cols - 1) Then
            FGrid.Col = mr
            Exit Do
        End If
        If FGrid.Col = MaxCol Then
            If FGrid.Parent.beforeRowUpdate = False Then FGrid.Col = mr
            FGrid.Row = FGrid.Row + 1
            FGrid.Col = 1
        Else
            FGrid.Col = FGrid.Col + 1
        End If
        If TEXT(FGrid.Col).GridVisible = 1 Then Exit Do
    Loop
ElseIf KeyCode = vbKeyReturn Then        ' Enter
    mr = FGrid.Col
    Do While True
        If FGrid.Cols - 1 = FGrid.Col Or FGrid.Col = MaxCol Then
            mROW = Val(FGrid.TextMatrix(FGrid.Row, 1))
            If FGrid.Rows - 1 = FGrid.Row Then
                If FGrid.Parent.beforeRowUpdate = True Then
                    FGrid.AddItem ""
                    mROW = mROW + 1
                End If
            End If
            If FGrid.Parent.beforeRowUpdate = True Then
                FGrid.Row = FGrid.Row + 1
                FGrid.Col = 1
            End If
        Else
            FGrid.Col = FGrid.Col + 1
        End If
        If TEXT(FGrid.Col).GridVisible = 1 Then Exit Do
    Loop
End If
Call Text_FGrid_Set_Txt(TEXT, FGrid, MaxCol)
Call Text_FGrid_Set(FGrid.Col, TEXT, FGrid, MaxCol)
End Sub

Public Sub Text_FGrid_Set(mIndex As Integer, TEXT As Object, FGrid As MSHFlexGrid, MaxCol As Byte)
Dim J As Byte
    For J = 1 To MaxCol
        TEXT(J).TEXT = ""
        If TEXT(J).Visible = True Then TEXT(J).Visible = False
    Next
    TEXT(mIndex).ZOrder 0
    TEXT(mIndex).height = FGrid.CellHeight + 10
    TEXT(mIndex).width = FGrid.CellWidth + 10
    TEXT(mIndex).top = FGrid.top + FGrid.CellTop '- 20
    TEXT(mIndex).left = FGrid.left + FGrid.CellLeft '- 20
    Call Text_FGrid_Set_Txt(TEXT, FGrid, MaxCol)
    TEXT(mIndex).Visible = True
    TEXT(mIndex).SetFocus
End Sub
Private Sub Text_FGrid_Set_Txt(TEXT As Object, FGrid As MSHFlexGrid, MaxCol As Byte)
Dim J As Byte
For J = 1 To MaxCol
    TEXT(J).TEXT = FGrid.TextMatrix(FGrid.Row, J)
Next
End Sub

'Provided by shekher discussion required
Public Function FillGrid(FGrid As MSFlexGridLib.MSFlexGrid, SQL As String) As Boolean
Dim rs As Recordset
    Set rs = New Recordset
    rs.Open SQL, GCn, adOpenStatic, adLockReadOnly
    FGrid.Rows = 1
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
            FGrid.AddItem "" & Chr(9) & rs.Fields(0).Value & Chr(9) & rs.Fields(1).Value
            rs.MoveNext
        Loop
    End If
    FGrid.ColAlignment(1) = flexAlignLeftCenter
    rs.Close
    Set rs = Nothing
End Function

'Provided by shekher discussion required
Public Function FGrid_Click(FGrid As MSFlexGridLib.MSFlexGrid) As Boolean
    FGrid.Col = 0
    FGrid.CellFontName = "WINGDINGS"
    FGrid.CellFontSize = 14
    
    FGrid.TextMatrix(FGrid.Row, 0) = IIf(FGrid.TextMatrix(FGrid.Row, 0) = "ü", " ", "ü")
End Function

'Provided by shekher discussion required
Public Function Ini_Fgrid(FGrid As MSFlexGridLib.MSFlexGrid, mLeft As Long, mTop As Long, mHeight As Long, mWidth As Long, mColWidth As Integer) As Boolean
With FGrid
    .Cols = 3
    .Rows = 2
    .FixedCols = 1
    .Appearance = flexFlat
    .FixedRows = 1
    .GridColor = &H0&
    .GridColorFixed = &H0&
    .left = mLeft
    .top = mTop
    .height = mHeight
    .width = mWidth
'    .ColAlignment(2) = flexAlignLeftCenter
    .ColWidth(0) = 500
    .ColWidth(1) = 0
    .ColWidth(2) = mColWidth
    .ScrollBars = flexScrollBarVertical
    .BackColorBkg = MDIForm1.BackColor
    .ForeColorFixed = &H800080
    .BackColorFixed = MDIForm1.BackColor
    .Visible = True
    .ZOrder 0
End With
    
End Function



