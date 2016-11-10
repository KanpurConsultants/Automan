VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDataRecd 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Data Received"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8895
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "All                 Table Group                                      Table Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   10
      Top             =   1230
      Value           =   1  'Checked
      Width           =   8130
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3225
      TabIndex        =   1
      Top             =   480
      Width           =   1485
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1215
      TabIndex        =   0
      Top             =   480
      Width           =   1485
   End
   Begin VB.CommandButton CmdBackUp 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update Database"
      Height          =   375
      Left            =   4755
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel  &&  Exit"
      Height          =   375
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1560
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2610
      TabIndex        =   2
      Top             =   840
      Width           =   4395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   8493
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      BackColorFixed  =   16777152
      ForeColorFixed  =   192
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
      Enabled         =   0   'False
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   3
      Left            =   1410
      TabIndex        =   14
      Top             =   195
      Width           =   420
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Type:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   2
      Left            =   195
      TabIndex        =   13
      Top             =   195
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6750
      TabIndex        =   12
      Top             =   345
      Width           =   2070
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Give Source Database Path"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Index           =   1
      Left            =   195
      TabIndex        =   11
      Top             =   870
      Width           =   2160
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Index           =   7
      Left            =   2805
      TabIndex        =   8
      Top             =   525
      Width           =   315
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date From :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Index           =   0
      Left            =   195
      TabIndex        =   7
      Top             =   525
      Width           =   915
   End
   Begin VB.Label lblTabName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6750
      TabIndex        =   6
      Top             =   90
      Width           =   420
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Index           =   4
      Left            =   5550
      TabIndex        =   5
      Top             =   90
      Width           =   1065
   End
End
Attribute VB_Name = "frmDataRecd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DateFrom As Byte = 0
Private Const DateTo As Byte = 1
Private Const TxtBackUpPath As Byte = 2
Dim GCnBackUp As ADODB.Connection
Dim GCnBackUpFA As ADODB.Connection
Dim RstData As ADODB.Recordset
Private Sub Check1_Click(Index As Integer)
    If Check1(Index).Value = Unchecked Then
        GridSel(Index).Enabled = True
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 1: GridSel(Index).Col = 1
        End If
    Else
        GridSel(Index).Enabled = False
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 0: GridSel(Index).Col = 0
            GridSel(Index).RowSel = GridSel(Index).Rows - 1
        End If
    End If
End Sub
Private Sub CmdBackUp_Click()
If Txt(TxtBackUpPath) = "" Then MsgBox LBLCNT(1): Exit Sub
If IsValid(Txt(DateFrom), "Date From") = False Then Exit Sub
If IsValid(Txt(DateTo), "Date To") = False Then Exit Sub
If MsgBox("Are You Sure to Update Data? ", vbYesNo + vbCritical + vbDefaultButton2, "Data Transfer (Recd) !") <> vbYes Then Exit Sub
Dim mTrans As Boolean, RstDataTemp As ADODB.Recordset, RstDate As ADODB.Recordset, I As Integer
On Error Resume Next
Dim BackUpPath$, BackUpPathFa$, UserDate1$, UDate$, ClosedUEntDt$, Counter As Integer
Dim DB As DAO.Database, TableFound As Boolean
'Table_Name,TableDesc,DataGatg,PrimName
BackUpPath = Txt(TxtBackUpPath) & "\Automan.mdb"
BackUpPathFa = Txt(TxtBackUpPath) & "\FaData.mdb "
Set GCnBackUp = New ADODB.Connection
With GCnBackUp
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & BackUpPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    .Open
End With

Set GCnBackUpFA = New ADODB.Connection
With GCnBackUpFA
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & BackUpPathFa & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    .Open
End With
GCn.BeginTrans
G_FaCn.BeginTrans
mTrans = True
Do Until RstData.EOF
    TableFound = False
    
    Set DB = OpenDatabase(Txt(TxtBackUpPath) & "\Automan.mdb")
    For I = 0 To DB.TableDefs.Count - 1
        If RstData!Table_Name = DB.TableDefs(I).Name Then
            TableFound = True
        End If
    Next I
    Set DB = OpenDatabase(Txt(TxtBackUpPath) & "\Fadata.mdb")
    For I = 0 To DB.TableDefs.Count - 1
        If RstData!Table_Name = DB.TableDefs(I).Name Then
            TableFound = True
        End If
    Next I
    If TableFound = False Then GoTo NXT
    If Check1(0).Value = Unchecked Then
        For I = 0 To GridSel(0).Rows - 1
            If RstData!Table_Name = GridSel(0).TextMatrix(I, 1) And GridSel(0).TextMatrix(I, 0) <> "" Then
                GoTo Jump
            End If
        Next
        GoTo NXT
    End If
Jump:
    lblTabName = RstData!Table_Name
    UserDate1 = RstData!userdate
    lblTabName.Refresh
  If RstData!DataGatg = "D" Then
        If RstData!userdate <> "None" And RstData!userdate <> "" Then
            If RstData!Table_Name = "Job_Card" Then
                Set RstDataTemp = GCnBackUp.Execute("select DocID as Prim,U_EntDt as User_Date,ClosedU_EntDt from " & RstData!Table_Name & " where " & UserDate1 & "  >=" & ConvertDate(Txt(DateFrom)) & " and " & UserDate1 & " <=" & ConvertDate(Txt(DateTo)) & " order by " & RstData!primName & "")
            Else
                Set RstDataTemp = GCnBackUp.Execute("select " & RstData!primName & " as Prim," & RstData!userdate & " as User_Date from " & RstData!Table_Name & " where " & UserDate1 & "  >=" & ConvertDate(Txt(DateFrom)) & " and " & UserDate1 & " <=" & ConvertDate(Txt(DateTo)) & " order by " & RstData!primName & "")
            End If
        ElseIf RstData!userdate = "None" Then
             Set RstDataTemp = GCnBackUp.Execute("select " & RstData!primName & " as Prim  From " & RstData!Table_Name & " order by " & RstData!primName & "")
        End If
        Label1.Tag = CStr(RstDataTemp.RecordCount) + "#"
        Counter = 1
        If RstDataTemp.RecordCount > 0 Then RstDataTemp.MoveFirst
        Do Until RstDataTemp.EOF
            Select Case RstData!Table_Name
                Case "Job_Card"
                        If UCase(Trim(RstData!userdate)) <> "NONE" And UCase(Trim(RstData!userdate)) <> "" Then
                            Set RstDate = GCn.Execute("Select U_EntDt,ClosedU_EntDt from " & RstData!Table_Name & " where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                            If RstDate.RecordCount <> 0 Then
                               UDate = XNull(RstDate.Fields(0))
                               ClosedUEntDt = XNull(RstDate.Fields(1))
                            End If
                        End If
                        If (Trim(UDate) <> "" Or Trim(ClosedUEntDt) <> "") Then
                            If CDate(UDate) <= CDate(XNull(RstDataTemp!User_Date)) Then
                                If XNull(RstDataTemp!ClosedU_EntDt) <> "" Then
                                      If CDate(ClosedUEntDt) <= CDate(XNull(RstDataTemp!ClosedU_EntDt)) Then
                                            If RstData!PrimGatg = "T" Then
                                                GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                                                GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                                            Else
                                                GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = " & RstDataTemp!prim & "")
                                                GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = " & RstDataTemp!prim & "")
                                            End If
                                      End If
                                End If
                            End If
                        Else
                            If RstData!PrimGatg = "T" Then
                                GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                                GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                            Else
                                GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = " & RstDataTemp!prim & "")
                                GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = " & RstDataTemp!prim & "")
                            End If
                        End If
                Case Else
                        If UCase(Trim(RstData!userdate)) <> "NONE" And UCase(Trim(RstData!userdate)) <> "" Then
                            Set RstDate = GCn.Execute("Select " & RstData!userdate & " from " & RstData!Table_Name & " where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                            If RstDate.RecordCount <> 0 Then
                               UDate = XNull(RstDate.Fields(0))
                            End If
                        End If
                        If Trim(UDate) <> "" Then
                            If CDate(UDate) <= CDate(RstDataTemp!User_Date) Then
                                If RstData!PrimGatg = "T" Then
                                    GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                                    GCn.CommitTrans
                                    GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & "]." & RstData!Table_Name & "  where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                                    GCn.CommitTrans
                                Else
                                    GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = " & RstDataTemp!prim & "")
                                    GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & "]." & RstData!Table_Name & "  where " & RstData!primName & " = " & RstDataTemp!prim & "")
                                End If
                            End If
                        Else
                            If RstData!PrimGatg = "T" Then
                                GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                                GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                            Else
                                GCn.Execute ("delete  from " & RstData!Table_Name & " where " & RstData!primName & " = " & RstDataTemp!prim & "")
                                GCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPath & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = " & RstDataTemp!prim & "")
                            End If
                        End If
                End Select
                
xxx:
            UDate = ""
            ClosedUEntDt = ""
            Label1.CAPTION = Label1.Tag + CStr(Counter)
            Label1.Refresh
            Counter = Counter + 1
            
            RstDataTemp.MoveNext
        Loop
        
    ElseIf RstData!DataGatg = "A" Then
        If RstData!userdate <> "None" And RstData!userdate <> "" Then
            Set RstDataTemp = GCnBackUpFA.Execute("select " & RstData!primName & " as Prim," & RstData!userdate & " as User_Date from " & RstData!Table_Name & " where " & UserDate1 & "  >=" & ConvertDate(Txt(DateFrom)) & " and " & UserDate1 & " <=" & ConvertDate(Txt(DateTo)) & " order by " & RstData!primName & "")
        ElseIf RstData!userdate = "None" Then
            Set RstDataTemp = GCnBackUpFA.Execute("select " & RstData!primName & " as Prim  from " & RstData!Table_Name & " order by " & RstData!primName & "")
        End If
        Label1.Tag = CStr(RstDataTemp.RecordCount) + "#"
        Counter = 1
        '************** For Opening Transfer ********
        Dim rsdataTmp As ADODB.Recordset
        If RstData!Table_Name = "Ledger" Then
             Set rsdataTmp = GCnBackUpFA.Execute("select subcode from " & RstData!Table_Name & " where V_type IN('F_AO','F_OP')")
             If rsdataTmp.RecordCount > 0 Then
                For I = 1 To rsdataTmp.RecordCount
                    G_FaCn.Execute ("Delete from ledger where v_type in ('F_AO','F_OP') and subcode='" & rsdataTmp!SubCode & "'")
                Next
             End If
        End If
        Set rsdataTmp = Nothing
        '*********************************************
        Do Until RstDataTemp.EOF
       
                    If RstData!PrimGatg = "T" Then
                        
                        G_FaCn.Execute ("delete * from " & RstData!Table_Name & " where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                        G_FaCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPathFa & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = '" & RstDataTemp!prim & "'")
                    Else
                      
                        G_FaCn.Execute ("delete * from " & RstData!Table_Name & " where " & RstData!primName & " = " & RstDataTemp!prim & "")
                        G_FaCn.Execute ("Insert into " & RstData!Table_Name & " select * from [" & BackUpPathFa & ";pwd=dtman]." & RstData!Table_Name & "  where " & RstData!primName & " = " & RstDataTemp!prim & "")
                    End If
            Label1.CAPTION = Label1.Tag + CStr(Counter)
            Label1.Refresh
            Counter = Counter + 1
            RstDataTemp.MoveNext
        Loop
    End If
NXT:
    RstData.MoveNext
Loop

'Update Current Balance's
'Account Balance
    Dim Rst As ADODB.Recordset
    G_FaCn.Execute ("update SubGroup set Curr_Bal=0")
    GCn.Execute ("update SubGroup set Curr_Bal=0 ")
    GSQL = "SELECT Ledger.SubCode,SUM(AmtCr-AmtDr) as CBal FROM Ledger left join SubGroup SG on SG.SubCOde=Ledger.SubCode group by Ledger.subcode,Name"
    Set Rst = G_FaCn.Execute(GSQL)
    If Rst.RecordCount > 0 Then
        Do While Rst.EOF = False
            GCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            G_FaCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            Rst.MoveNext
        Loop
    End If
    Set Rst = Nothing
    'eof balance updatation
    GCn.CommitTrans
    G_FaCn.CommitTrans

    MsgBox "Transfer Complete", vbInformation, "Transfer Data"
    Set RstDataTemp = Nothing
    Set DB = Nothing

Unload Me
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans: G_FaCn.RollbackTrans
    Set RstDataTemp = Nothing
    Set DB = Nothing
    CheckError
End Sub
Private Sub CmdCancel_Click()
Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte, SiteType$
WinSetting Me, 6990, 9015
    Txt(DateFrom) = ""
    Txt(DateTo) = ""
    lblTabName = ""
    lblTotTable = ""
    lblTotTabData = ""
'    Set RstData = GCn.Execute("select * from TableGroup order by TableDesc")
    SiteType = GCn.Execute("Select SiteType from Site where Site_Code='" & PubSiteCode & "'").Fields(0).Value
    If SiteType = "H" Then  'HO
        LBLCNT(3).CAPTION = "From WorkShop"
        Set RstData = GCn.Execute("select * from TableGroupClient order by TableDesc")
    Else
        LBLCNT(3).CAPTION = "From HO"
        Set RstData = GCn.Execute("select * from TableGroupHO order by TableDesc")
    End If
    Set GridSel(0).DataSource = RstData
    GridSel(0).ColWidth(1) = 0
    GridSel(0).ColWidth(3) = 0
    GridSel(0).ColWidth(4) = 0
    GridSel(0).ColWidth(5) = 0
    GridSel(0).ColWidth(6) = 3000
    GridSel(0).ColWidth(7) = 0
    GridSel(0).ColWidth(2) = 3000
    Label1.CAPTION = ""
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set GCnBackUp = Nothing
Set GCnBackUpFA = Nothing
Set RstData = Nothing
End Sub

Private Sub GridSel_Click(Index As Integer)
    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "�", " ", "�")
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus Txt(Index)
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Exit Sub
End If
If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
    Call CheckQuote(keyascii)
End Sub
Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case DateFrom
        Txt(Index).TEXT = RetDate(Txt(Index))
        If Txt(DateFrom) <> "" Then
            If Txt(DateTo) <> "" Then
                If CDate(Txt(DateFrom)) > CDate(Txt(DateTo)) Then
                    MsgBox "Date to is less than Date from", vbOKOnly, "Validation"
                    Cancel = True: Exit Sub
                End If
            End If
        End If
    Case DateTo
        Txt(Index).TEXT = RetDate(Txt(Index))
        If Txt(DateFrom) <> "" Then
            If Txt(DateTo) <> "" Then
                If CDate(Txt(DateFrom)) > CDate(Txt(DateTo)) Then
                    MsgBox "Date to is less than Date from", vbOKOnly, "Validation"
                    Cancel = True: Exit Sub
                End If
            End If
        End If
End Select
End Sub

