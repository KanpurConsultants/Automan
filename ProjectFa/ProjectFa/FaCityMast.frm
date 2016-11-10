VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form FaCityMast 
   BackColor       =   &H00CFE0E0&
   Caption         =   "City Master"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "FaCityMast.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   8880
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   661
   End
   Begin VB.Frame FrState 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2460
      Left            =   1155
      TabIndex        =   2
      Top             =   2445
      Visible         =   0   'False
      Width           =   4215
      Begin MSDataGridLib.DataGrid DgState 
         Height          =   2115
         Left            =   30
         TabIndex        =   4
         Top             =   330
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   3731
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "citycode"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cityname"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   5
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   134.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3479.811
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "List of City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   4140
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2805
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1515
      Width           =   4215
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   2460
      TabIndex        =   5
      Top             =   1515
      Width           =   195
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   1620
      TabIndex        =   1
      Top             =   1515
      Width           =   555
   End
End
Attribute VB_Name = "FaCityMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADDFLAG As Byte, Master As ADODB.Recordset, RstHelp As ADODB.Recordset, mFlag As Byte
Private Const StName As Byte = 0
Private IntStaId As String
Private PubDatamanFa As New DMFa.ClsFa

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errloop
    TopCtrl1.PrvKeyCode = KeyCode
    If KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Or KeyCode = vbKeyF4 Or (KeyCode = 70 And Shift = 2) Or (KeyCode = 80 And Shift = 2) Or (KeyCode = 83 And Shift = 2) Or KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or KeyCode = vbKeyF10 Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEnd Then
        TopCtrl1.TopKey_Down KeyCode, Shift
    End If
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Form_Load()
Dim I As Byte
On Error GoTo Errloop
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    If PubSec = "SANJEEV" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_NAME='" & Me.CAPTION & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    ElseIf PubSec = "RAHUL" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_CODE='" & Me.Name & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    End If
    Me.top = 0
    Me.left = 0
    Me.BackColor = FrmBackCol
    '''''''''''''
    PubDatamanFa.FaBackEnd = PubBackEnd
    PubDatamanFa.FaPubLoginDate = PubLoginDate
    PubDatamanFa.FaPubDivCode = PubDivCode
    PubDatamanFa.FaPubSiteCode = PubSiteCode
    PubDatamanFa.FaPubSiteCodeDisplay = PubSiteCodeDisplay
    PubDatamanFa.FaPubSiteName = PubSiteName
    PubDatamanFa.FapubUName = pubUName
    PubDatamanFa.FaDosPort = PubFaDosPort
    PubDatamanFa.FaRunPIF = PubRunPIF
    PubDatamanFa.FaPubSiteType = PubFaSiteType
    Set PubDatamanFa.SetG_FaCn = G_FaCn
    Set PubDatamanFa.SetG_CompCn = G_CompCn
    Set PubDatamanFa.SetrsUserPerm = rsUserPerm.Clone
    Set PubDatamanFa.SetMasterRst = FaMasterRst.Clone
    '''''''''''''
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    Set Master = G_FaCn.Execute("SELECT * FROM city ORDER BY cityname")
    Set RstHelp = G_FaCn.Execute("SELECT citycode,cityname FROM city ORDER BY cityname")
    With FrState
        .left = 2805
        .top = 1800
    End With
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    ADDFLAG = 0
    mFlag = 0
    DgState.Columns(0).Visible = False
    Set DgState.DataSource = RstHelp
    Me.TopCtrl1.TopText1.left = 5800
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RstHelp = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
End Sub
Private Sub MakeBlank()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
    Next
End Sub
Private Sub MoveRec()
On Error GoTo Errloop
    FaRstBofEof Master
    If Master.RecordCount <= 0 Then
        MakeBlank
    Else
        IntStaId = Master!CityCode
        Txt(StName).TEXT = Master!CityName
    End If
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo Errloop
    MakeBlank
    ADDFLAG = 1
    Disp_Text SETS("ADD", Me, Master)
    Txt(StName).SetFocus
    Exit Sub
Errloop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
    If Master.RecordCount > 0 Then
        ADDFLAG = 2
        Disp_Text SETS("EDIT", Me, Master)
        Txt(StName).SetFocus
    Else
        MsgBox "There Is No Record To Edit.", vbInformation, "Information"
    End If
    Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
Dim transFalg As Byte
transFalg = 0
On Error GoTo Errloop
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure to Delete This Record", vbYesNo, "Confirmation") = vbYes Then
            G_FaCn.BeginTrans
            transFalg = 1
            G_FaCn.Execute ("Delete From city Where citycode='" & IntStaId & "'")
            G_FaCn.CommitTrans
            transFalg = 0
            Master.Requery
            RstHelp.Requery
            Disp_Text SETS("INI", Me, Master)
            MoveRec
        End If
    Else
        MsgBox "There Is No Record To Delete.", vbInformation, "Information"
    End If
    Exit Sub
Errloop:
    If transFalg = 1 Then
        G_FaCn.RollbackTrans
        MsgBox err.Description, vbExclamation, " Deletion Error "
    End If
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select city.citycode As SearchCode,city.cityname FROM city Order by city.cityname"
    Set SearchForm = Me
    FAFind.Show vbModal
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.Find ("citycode='" & MyValue & "'")
    MoveRec
    Exit Sub
ErrorLoop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
Dim Rst1 As ADODB.Recordset, X11, I As Integer
On Error GoTo ERRORHANDLER
Set Rst1 = G_FaCn.Execute("select * FROM CITY order by CityName")
If Rst1.RecordCount = 0 Then MsgBox "No record Found to Print": Exit Sub
'X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaCityMast.ttx", True)
Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaCityMast.RPT")
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("Title")
            rpt.FormulaFields(I).TEXT = "'City List'"
    End Select
Next
rpt.Database.SetDataSource Rst1
rpt.ReadRecords
FaReport_View rpt, 0, Me.CAPTION, True
Set Rst1 = Nothing
Exit Sub
ERRORHANDLER:  MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte, MySql$
On Error GoTo Errloop
    Ctrl_BckColor
    transFlag = 0
    If FaIsValid(Txt(StName), "City Name") = False Then Txt_GotFocus StName: Exit Sub
    If ADDFLAG = 1 Then
        IntStaId = G_FaCn.Execute("Select iif(IsNull(Max(val(citycode))),1,Max(val(citycode))+1)AS MyCode From city").Fields(0).Value
    End If
    G_FaCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        MySql = "Insert Into city(citycode,cityname,U_Name,U_EntDt,U_AE) Values('" & IntStaId & "','" & Txt(StName).TEXT & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A' )"
        G_FaCn.Execute (MySql)
    Else
        MySql = "UPDATE city SET cityname='" & Txt(StName) & "',U_name='" & pubUName & "',U_EntDt=" & FaConvertDate(PubLoginDate) & " ,U_AE='E' WHERE citycode='" & IntStaId & "'"
        G_FaCn.Execute (MySql)
    End If
    G_FaCn.CommitTrans
    transFlag = 0
    Master.Requery
    RstHelp.Requery
    DgState.Refresh
    Master.Find ("citycode='" & IntStaId & "'")
    If ADDFLAG = 1 Then
        MakeBlank
        Txt_GotFocus StName
        Txt(StName).SetFocus
    Else
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        ADDFLAG = 0
        FrState.Visible = False
    End If
    Exit Sub
Errloop:        If transFlag = 1 Then G_FaCn.RollbackTrans
                MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
    Ctrl_BckColor
    ADDFLAG = 0
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    FrState.Visible = False
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
Errloop:    MsgBox err.Description, vbCritical, "Information"
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub citynameSearch()
    If RstHelp.RecordCount <= 0 Then Exit Sub
    RstHelp.MoveFirst
    RstHelp.Find "cityname>='" & Txt(StName) & "'"
End Sub
Private Sub Txt_Change(Index As Integer)
    If ADDFLAG <> 0 Then
        Select Case Index
            Case StName
                FrState.Visible = True
                FrState.top = Txt(Index).top + Txt(Index).height + 10
                FrState.left = Txt(Index).left
                FrState.ZOrder 0
                citynameSearch
        End Select
    End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
On Error GoTo Errloop
    mFlag = 0
    Call Ctrl_GetFocus(Index)
    If FrState.Visible = True Then FrState.Visible = False
    FaRstBofEof RstHelp
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    Select Case Index
        Case StName
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case StName
            mBookMark = RstHelp.Bookmark
            RstHelp.Bookmark = mBookMark
    End Select
    If Txt(Index) = "" Then Txt_Change Index
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_Click(Index As Integer)
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim MyResult As Boolean
On Error GoTo Errloop
Dim I As Integer
On Error GoTo Errloop
    Select Case Index
        Case StName
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstHelp.BOF Then RstHelp.MovePrevious
                Case vbKeyDown
                    If Not RstHelp.EOF Then RstHelp.MoveNext
                Case 33
                    For I = 1 To 9
                        If Not RstHelp.BOF Then RstHelp.MovePrevious
                    Next
                Case 34
                    For I = 1 To 9
                        If Not RstHelp.EOF Then RstHelp.MoveNext
                    Next
                Case 13
                    FrState.Visible = False
                    Txt_Validate Index, MyResult
                    If MyResult = True Then Txt(StName).SetFocus: Exit Sub
                    If MsgBox("Save Record?", vbYesNo, "Save Entry") = vbYes Then
                        TopCtrl1_eSave
                        Exit Sub
                    Else
                        Me.ActiveControl.SetFocus
                    End If
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    FaRstBofEof RstHelp
                    If Not RstHelp.BOF And Not RstHelp.EOF Then
                        Txt(StName).TEXT = RstHelp!CityName
                        Txt(StName).SelStart = 0
                    End If
            End Select
    End Select
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    citynameSearch
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    Call Ctrl_validate(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Lrs As ADODB.Recordset
On Error GoTo Errloop
    Select Case Index
        Case StName
            Set Lrs = CreateObject("ADODB.Recordset")
            With Lrs
                Lrs.ActiveConnection = G_FaCn
                Lrs.CursorType = adOpenStatic
                Lrs.CursorLocation = adUseClient
                Lrs.Open "Select * From city Where cityname='" & Txt(StName).TEXT & "'"
            End With
            If ADDFLAG = 1 Then
                If Not Lrs.EOF Then MsgBox "City Name Already Exists", vbInformation, "Validation":  Cancel = True:  Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Lrs.EOF Then
                    If Lrs!CityName <> Master!CityName Then MsgBox "City Name Already Exists", vbInformation, "Validation":  Cancel = True: Exit Sub
                End If
            End If
    End Select
    Set Lrs = Nothing
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Ctrl_GetFocus(Index As Integer)
    Txt(Index).BackColor = CtrlBCol
    Txt(Index).ForeColor = CtrlFCol
    Txt(Index).BorderStyle = 1
End Sub
Private Sub Ctrl_validate(Index As Integer)
    Txt(Index).BackColor = CtrlBColOrg
    Txt(Index).ForeColor = CtrlFColOrg
    Txt(Index).BorderStyle = 0
End Sub
Private Sub Ctrl_BckColor()
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).BorderStyle = 0
    Next
End Sub
