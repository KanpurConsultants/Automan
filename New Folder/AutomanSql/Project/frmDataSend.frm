VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDataSend 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Data Send"
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
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel  &&  Exit"
      Height          =   375
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6090
      Width           =   1560
   End
   Begin VB.CommandButton CmdBackUp 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Start Transfer"
      Height          =   375
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6090
      Width           =   2175
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
      Left            =   1260
      TabIndex        =   0
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
      Index           =   1
      Left            =   3270
      TabIndex        =   1
      Top             =   480
      Width           =   1485
   End
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
      TabIndex        =   4
      Top             =   1230
      Value           =   1  'Checked
      Width           =   8130
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   5
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
      Caption         =   "Transfer Type : "
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
      Index           =   3
      Left            =   1650
      TabIndex        =   11
      Top             =   195
      Width           =   1275
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Type : "
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
      Index           =   2
      Left            =   225
      TabIndex        =   10
      Top             =   195
      Width           =   1275
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
      Left            =   5415
      TabIndex        =   9
      Top             =   120
      Width           =   1065
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
      Left            =   6600
      TabIndex        =   8
      Top             =   120
      Width           =   420
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
      Height          =   300
      Index           =   0
      Left            =   225
      TabIndex        =   7
      Top             =   480
      Width           =   915
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
      Height          =   300
      Index           =   7
      Left            =   2850
      TabIndex        =   6
      Top             =   480
      Width           =   315
   End
End
Attribute VB_Name = "frmDataSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const DateFrom As Byte = 0
Private Const DateTo As Byte = 1
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
If IsValid(Txt(DateFrom), "Date From") = False Then Exit Sub
If IsValid(Txt(DateTo), "Date To") = False Then Exit Sub
If MsgBox("Are You Sure ? ", vbYesNo + vbCritical + vbDefaultButton2, "Data Transfer (Send) !") <> vbYes Then Exit Sub

Dim mTrans As Boolean, RstDataTemp As ADODB.Recordset, RstTmp As ADODB.Recordset, I As Integer
On Error GoTo ELoop
Dim BackUpPath$, BackUpPathFa$
Dim DataPath$, DataPathFa$
Dim fob As New FileSystemObject
Dim DB As DAO.Database
'** New System
Dim wrkDefault As Workspace, NewDBName$
   

'Table_Name,TableDesc,DataGatg,PrimName
DataPath = Pub_DataPath & "\" & PubCenDataPath & "\Automan.mdb"
If MDIForm1.SBAR.Panels(2).TEXT = "Vehicle" And PubVFADataPath <> "" Then
    DataPathFa = PubVFADataPath
ElseIf MDIForm1.SBAR.Panels(2).TEXT = "Spare" And PubSFADataPath <> "" Then
    DataPathFa = PubSFADataPath
ElseIf MDIForm1.SBAR.Panels(2).TEXT = "WorkShop" And PubWFADataPath <> "" Then
    DataPathFa = PubWFADataPath
End If

BackUpPath = Pub_DataPath & "\Transfer\Automan.mdb"
BackUpPathFa = Pub_DataPath & "\Transfer\FaData.mdb"

'New System
    Set wrkDefault = DBEngine.Workspaces(0)
    
    ' Make sure there isn't already a file with the name of
    ' the new database.
    NewDBName = BackUpPath
    If Dir(NewDBName) <> "" Then Kill NewDBName  '"NewDB.mdb"
    ' Create a new encrypted database with the specified
    ' collating order.
    Set dbsNew = wrkDefault.CreateDatabase(NewDBName, _
        dbLangGeneral, dbEncrypt)
    Set dbsNew = Nothing
    'Create FAData
    NewDBName = BackUpPathFa
    If Dir(NewDBName) <> "" Then Kill NewDBName  '"NewDB.mdb"
    ' Create a new encrypted database with the specified
    ' collating order.
    Set dbsNew = wrkDefault.CreateDatabase(NewDBName, _
        dbLangGeneral, dbEncrypt)
    Set dbsNew = Nothing
'EOF New System

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
On Error Resume Next
Do Until RstData.EOF
    If Check1(0).Value = Unchecked Then
        For I = 0 To GridSel(0).Rows - 1
            If RstData!Table_Name = GridSel(0).TextMatrix(I, 1) And GridSel(0).TextMatrix(I, 0) <> "" Then
                GoTo xxx
            End If
        Next
        GoTo NXT
    End If
xxx:
    lblTabName = RstData!Table_Name
    lblTabName.Refresh
    If RstData!DataGatg = "D" Then
        If RstData!userdate <> "None" And RstData!userdate <> "" Then
            If RstData!Table_Name = "Job_Card" Then
                GCn.Execute ("Update " & RstData!Table_Name & " set " & RstData!userdate & " = format(" & RstData!userdate & ",'dd/mmm/yyyy')")
                GCn.Execute ("Select " & RstData!Table_Name & ".* into [" & BackUpPath & "]." & RstData!Table_Name & " from " & RstData!Table_Name & " where (" & RstData!userdate & "  >=" & ConvertDate(Txt(DateFrom)) & " and " & RstData!userdate & " <=" & ConvertDate(Txt(DateTo)) & ") or (ClosedU_EntDt  >=" & ConvertDate(Txt(DateFrom)) & " and ClosedU_EntDt <=" & ConvertDate(Txt(DateTo)) & ") ")
            Else
                GCn.Execute ("Update " & RstData!Table_Name & " set " & RstData!userdate & " = format(" & RstData!userdate & ",'dd/mmm/yyyy')")
                GCn.Execute ("Select " & RstData!Table_Name & ".* into [" & BackUpPath & "]." & RstData!Table_Name & " from " & RstData!Table_Name & " where " & RstData!userdate & "  >=" & ConvertDate(Txt(DateFrom)) & " and " & RstData!userdate & " <=" & ConvertDate(Txt(DateTo)) & " ")
            End If
        ElseIf RstData!userdate = "None" Then
            GCn.Execute ("Select " & RstData!Table_Name & ".* into [" & BackUpPath & "]." & RstData!Table_Name & " from " & RstData!Table_Name & "")
        End If
    ElseIf RstData!DataGatg = "A" Then
        If RstData!userdate <> "None" And RstData!userdate <> "" Then
            G_FaCn.Execute ("Update " & RstData!Table_Name & " set " & RstData!userdate & " = format(" & RstData!userdate & ",'dd/mmm/yyyy')")
            G_FaCn.Execute ("Select " & RstData!Table_Name & ".* into [" & BackUpPathFa & "]." & RstData!Table_Name & " from " & RstData!Table_Name & " where " & RstData!userdate & "  >=" & ConvertDate(Txt(DateFrom)) & " and " & RstData!userdate & " <=" & ConvertDate(Txt(DateTo)) & " ")
        ElseIf RstData!userdate = "None" Then
            G_FaCn.Execute ("Select " & RstData!Table_Name & ".* into [" & BackUpPathFa & "]." & RstData!Table_Name & " from " & RstData!Table_Name & "")
        End If
    End If
NXT:
    RstData.MoveNext
Loop
GCn.CommitTrans
G_FaCn.CommitTrans
mTrans = False
Set DB = Nothing
Set RstDataTemp = Nothing
MsgBox "Transfer Complete", vbInformation, "Transfer Data"
Unload Me
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans: G_FaCn.RollbackTrans
    Set DB = Nothing
    Set RstDataTemp = Nothing
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
    SiteType = GCn.Execute("Select SiteType from Site where Site_Code='" & PubSiteCode & "'").Fields(0).Value
    If SiteType = "H" Then  'HO
        LBLCNT(3).CAPTION = "HO to WorkShop"
        Set RstData = GCn.Execute("select * from TableGroupHO order by TableDesc")
    Else
        LBLCNT(3).CAPTION = "WorkShop to HO"
        Set RstData = GCn.Execute("select * from TableGroupClient order by TableDesc")
    End If
    Set GridSel(0).DataSource = RstData
    GridSel(0).ColWidth(1) = 0
    GridSel(0).ColWidth(3) = 0
    GridSel(0).ColWidth(4) = 0
    GridSel(0).ColWidth(5) = 0
    GridSel(0).ColWidth(6) = 3000
    GridSel(0).ColWidth(7) = 0
    GridSel(0).ColWidth(2) = 3000
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
    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü", " ", "ü")
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

