VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmCity 
   Caption         =   "City Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10320
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   661
   End
   Begin VB.Frame FrState 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   6030
      TabIndex        =   11
      Top             =   2895
      Visible         =   0   'False
      Width           =   4095
      Begin MSDataGridLib.DataGrid DGState 
         Height          =   3225
         Left            =   30
         TabIndex        =   12
         Top             =   345
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "StateCode"
            Caption         =   "Code"
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
            DataField       =   "StateName"
            Caption         =   "Name"
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
         BeginProperty Column02 
            DataField       =   "StateHelp"
            Caption         =   "StateHelp"
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   4050
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   6660
      MaxLength       =   50
      TabIndex        =   15
      Top             =   1035
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2355
      MaxLength       =   4
      TabIndex        =   1
      Top             =   555
      Width           =   1260
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   4
      Left            =   2355
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "Local"
      Top             =   1275
      Width           =   1260
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1035
      Width           =   4245
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   2355
      MaxLength       =   25
      TabIndex        =   2
      Top             =   795
      Width           =   4245
   End
   Begin VB.Frame FrCity 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   660
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   4980
      Begin MSDataGridLib.DataGrid DGCity 
         Height          =   3225
         Left            =   30
         TabIndex        =   14
         Top             =   345
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "CityCode"
            Caption         =   "Code"
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
            DataField       =   "CityName"
            Caption         =   "Name"
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
         BeginProperty Column02 
            DataField       =   "CityCode"
            Caption         =   "Code"
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   4935
      End
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(L)ocal/(C)entral"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3990
      TabIndex        =   16
      Top             =   1275
      Width           =   1440
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code*................"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   870
      TabIndex        =   10
      Top             =   555
      Width           =   1515
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local/Central........."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   870
      TabIndex        =   9
      Top             =   1275
      Width           =   1680
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State Name*.........."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   870
      TabIndex        =   8
      Top             =   1035
      Width           =   1710
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City Name*.............."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   870
      TabIndex        =   5
      Top             =   795
      Width           =   1845
   End
End
Attribute VB_Name = "frmCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Don't Change Tag Property of (Txt) Control as it is used in other activities
'FORM COLOR &H00C0FFFF&
Option Explicit
Public MasterFormExit As Boolean
'Private Const CtrlBColOrg = &HC2D5B9        'Orginal BackColour
'Private Const CtrlFColOrg = &H80000012      'Orginal ForeColour
'Private Const CtrlBCol = &H80000008         'Changed BackColour
'Private Const CtrlFCol = &H8000000E         'Changed ForeColour
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstState As ADODB.Recordset, mFlag As Byte
Private Const CityCode = 0, CityName = 1, StateCode = 2, StateName = 3, LocalCentral = 4

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
ELoop:
Exit Sub
MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Load()
TopCtrl1.Tag = PubUParam: WinSetting Me, 5640, 9615

Set RstMain = New ADODB.Recordset
'RstMain.Open "Select mid(citycode,2,len(citycode)) as searchcode,CITY.*,State.StateName From City Left Join State On CITY.StateCode=STATE.StateCode Order by CityName", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
       
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      RstMain.Open "Select CityCode as SearchCode,CityName From City where site_code='" & PubSiteCode & "' Order by CityName", GCn, adOpenDynamic, adLockOptimistic
    Else
       RstMain.Open "Select  CityCode as SearchCode,CityName From City Order by CityName", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    
Else
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
           RstMain.Open "Select Top 1 CityCode as SearchCode,CityName From City where site_code='" & PubSiteCode & "' Order by CityName", GCn, adOpenDynamic, adLockOptimistic
        Else
           RstMain.Open "Select Top 1 CityCode as SearchCode,CityName From City Order by CityName", GCn, adOpenDynamic, adLockOptimistic
       End If
End If

Set RstHelp = New ADODB.Recordset
'RstHelp.Open "Select mid(citycode,2,len(citycode)) as searchcode,citycode,cityname FROM CITY  Order by CityName", GCn, adOpenDynamic, adLockOptimistic
   
RstHelp.Open "Select CityCode as SearchCode, CityCode, CityName FROM CITY  Order by CityName", GCn, adOpenDynamic, adLockOptimistic


Set RstState = New ADODB.Recordset
RstState.Open "Select StateCode,StateHelp, StateName FROM STATE Order by STATENAME", GCn, adOpenDynamic, adLockOptimistic

CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
mFlag = 0
Set DGCity.DataSource = RstHelp
FrCity.Visible = False
Set DGState.DataSource = RstState
FrState.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstState = Nothing
End Sub
Private Sub CtrlClckCol()
    Txt(CityCode).BackColor = CtrlBColOrg:      Txt(CityCode).ForeColor = CtrlFColOrg
    Txt(CityName).BackColor = CtrlBColOrg:      Txt(CityName).ForeColor = CtrlFColOrg
    Txt(StateName).BackColor = CtrlBColOrg:     Txt(StateName).ForeColor = CtrlFColOrg
    Txt(LocalCentral).BackColor = CtrlBColOrg:  Txt(LocalCentral).ForeColor = CtrlFColOrg
End Sub
Private Sub Disp_Text(Enb As Boolean)
    Txt(CityCode).Enabled = Enb
    Txt(CityName).Enabled = Enb
    Txt(StateName).Enabled = Enb
    Txt(LocalCentral).Enabled = Enb
End Sub
Private Sub MakeBlank()
    Txt(CityCode) = ""
    Txt(CityName) = ""
    Txt(StateName) = ""
    Txt(LocalCentral) = "Local"
End Sub
Private Sub MoveRec()
On Error GoTo ErrLoop
Dim Rstmain1 As ADODB.Recordset
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    Set Rstmain1 = New Recordset
    Rstmain1.CursorLocation = adUseClient
    Rstmain1.Open "Select CITY.*,State.StateName From City Left Join State On CITY.StateCode=STATE.StateCode where City.CityCode='" & RstMain!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
'    Txt(CityCode) = Mid(XNull(RstMain!CityCode), 2, 4)
    Txt(CityCode) = Rstmain1!CityCode
    Txt(CityName) = XNull(Rstmain1!CityName)
    Txt(StateCode) = VNull(Rstmain1!StateCode)
    Txt(StateName) = XNull(Rstmain1!StateName)
    Txt(LocalCentral) = IIf(Rstmain1!LocalCentral = "L", "Local", "Central")
End If
Set Rstmain1 = Nothing
TopCtrl1.tDel = False
'TopCtrl1.tPrn = False
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
MakeBlank
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
'Txt(CityCode).Tag = Txt(CityCode)
Txt(CityCode) = PubSiteCode
Txt_GotFocus CityCode
Txt(CityCode).SelStart = Len(Txt(CityCode))
Txt(CityCode).SetFocus
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(CityCode).Enabled = False
    Txt(CityName).Tag = Txt(CityName)
    Txt_GotFocus CityName
    Txt(CityName).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_eDel()
Dim XBM
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                G_FaCn.BeginTrans
                XBM = RstMain.Bookmark
                GCn.Execute ("Delete From City Where CityCode='" & Txt(CityCode) & "'")
                G_FaCn.Execute ("Delete From City Where CityCode='" & Txt(CityCode) & "'")
                GCn.CommitTrans
                G_FaCn.CommitTrans
                RstMain.Requery
                RstHelp.Requery
                If RstMain.RecordCount >= XBM Then
                    RstMain.Bookmark = XBM
                Else
                    If RstMain.EOF = False Then RstMain.MoveLast
                End If
                Call MoveRec
                BUTTONS True, Me, RstMain, 0
            End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
       G_FaCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If

End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, RstMain, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, RstMain, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, RstMain, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, RstMain, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop

    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
     sitecond = "where LEFT(city.site_code,1)='" & PubSiteCode & "'"
    Else
    sitecond = ""
    End If
    GSQL = "SELECT CityCode as SearchCode, City.CityCode,City.CityName,State.StateName FROM city left join state on state.statecode=city.statecode " & sitecond & "order by city.citycode"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
Dim rep As CrystalReport, Form1 As frmMastList
    Set Form1 = New frmMastList
    With Form1
        .g_FormID = 1
        .LblName.CAPTION = Me.CAPTION
        .CAPTION = Me.CAPTION
        .Show
    End With
    Set Form1 = Nothing
    Set rep = Nothing
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo ErrLoop
    transFlag = 0
'    If IsValid(Txt(CityCode), "City Code") = False Then Txt_GotFocus CityCode: Exit Sub
    If Len(Trim(Txt(CityCode))) = 1 Then MsgBox "Code should be filled ", vbOKOnly, "Validation": Txt(CityCode).SetFocus: Exit Sub
    If IsValid(Txt(CityName), "City Name") = False Then Txt_GotFocus CityName: Exit Sub
    If IsValid(Txt(StateName), "State") = False Then Txt_GotFocus StateName: Exit Sub
'    If AddFlag = 1 Then If GCn.Execute("Select COUNT(*) From City Where mid(CityCode,2,len(CityCode))=" & Chk_Text(Trim(Txt(CityCode)))).Fields(0)  > 0 Then MsgBox "City Code Already Exists", vbInformation, "City Code Validation": Txt_GotFocus CityCode: Txt(CityCode).SetFocus: Exit Sub
    If TopCtrl1.TopText2 = "Add" Then
        If GCn.Execute("Select COUNT(*) From City Where CityCode='" & Txt(CityCode) & "'").Fields(0) > 0 Then MsgBox "City Code Already Exists", vbInformation, "City Code Validation": Txt_GotFocus CityCode: Txt(CityCode).SetFocus: Exit Sub
        If GCn.Execute("Select COUNT(*) From City Where CityName='" & Txt(CityName) & "'").Fields(0) > 0 Then MsgBox "City Code Already Exists", vbInformation, "City Name Validation": Txt_GotFocus CityName: Txt(CityName).SetFocus: Exit Sub
        
        If G_FaCn.Execute("Select COUNT(*) From City Where CityCode='" & Txt(CityCode) & "'").Fields(0) > 0 Then MsgBox "City Code Already Exists", vbInformation, "City Code Validation": Txt_GotFocus CityCode: Txt(CityCode).SetFocus: Exit Sub
        If G_FaCn.Execute("Select COUNT(*) From City Where CityName='" & Txt(CityName) & "'").Fields(0) > 0 Then MsgBox "City Code Already Exists", vbInformation, "City Name Validation": Txt_GotFocus CityName: Txt(CityName).SetFocus: Exit Sub
    End If
    GCn.BeginTrans
    G_FaCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        If PubBackEnd = "A" Then GCn.Execute ("Insert Into City(CityCode,Site_Code,CityName,StateCode,LocalCentral,CityHelp,U_Name,U_EntDt,U_AE) Values('" & Txt(CityCode) & "','" & PubSiteCode & "'," & Chk_Text(Txt(CityName)) & "," & Txt(StateCode) & ",'" & left(Txt(LocalCentral), 1) & "','" & FilterString(Txt(CityName)) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
        G_FaCn.Execute ("DELETE From City Where CityCode='" & Txt(CityCode) & "'")
        G_FaCn.Execute ("Insert Into City(CityCode,Site_Code,CityName,StateCode,LocalCentral,CityHelp,U_Name,U_EntDt,U_AE) Values('" & Txt(CityCode) & "','" & PubSiteCode & "'," & Chk_Text(Txt(CityName)) & "," & Txt(StateCode) & ",'" & left(Txt(LocalCentral), 1) & "','" & FilterString(Txt(CityName)) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
    Else
        If PubBackEnd = "A" Then GCn.Execute ("update  City set CityName=" & Chk_Text(Txt(CityName)) & ",StateCode=" & Txt(StateCode) & ",LocalCentral='" & left(Txt(LocalCentral), 1) & "',CityHelp='" & FilterString(Txt(CityName)) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "'" & " Where CityCode='" & Txt(CityCode) & "'")
        G_FaCn.Execute ("update  City set CityName=" & Chk_Text(Txt(CityName)) & ",StateCode=" & Txt(StateCode) & ",LocalCentral='" & left(Txt(LocalCentral), 1) & "',CityHelp='" & FilterString(Txt(CityName)) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "'" & " Where CityCode='" & Txt(CityCode) & "'")
    End If
    GCn.CommitTrans
    G_FaCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select CityCode as SearchCode,CityName From City Where CityCode = '" & Txt(CityCode) & "' Order by CityName")
    End If
    RstHelp.Requery
    RstMain.FIND ("SearchCode= '" & Txt(CityCode) & "'")
    If TopCtrl1.TopText2 = "Add" Then
        TopCtrl1_eAdd
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrCity.Visible = False
        FrState.Visible = False
    End If
Exit Sub
ErrLoop:    If transFlag = 1 Then GCn.RollbackTrans: G_FaCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        FrState.Visible = False
        FrCity.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
    RstState.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub cityCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "searchcode >=" & Chk_Text(XNull(Trim(Txt(CityCode))))
End Sub
Private Sub cityNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "CityName >=" & Chk_Text(XNull(Txt(CityName)))
End Sub
Private Sub stateNameSearch()
If RstState.RecordCount <= 0 Then Exit Sub
RstState.MoveFirst
RstState.FIND "STATEName >=" & Chk_Text(XNull(Txt(StateName)))
If Not RstState.EOF Then
    If mID(RstState!StateName, 1, Len(Trim(XNull(Txt(StateName))))) <> Trim(XNull(Txt(StateName))) Then
        stateNameExSearch
    End If
Else
    stateNameExSearch
End If
End Sub
Private Sub stateNameExSearch()
Dim tempRst As ADODB.Recordset
Set tempRst = RstState.Clone
tempRst.Sort = "StateHelp ASC"
tempRst.FIND "StateHelp >='" & FilterString(XNull(Txt(StateName))) & "'"
If Not tempRst.EOF Then
    RstState.MoveFirst
    RstState.FIND "STATENAME >='" & XNull(tempRst!StateName) & "'"
    'Txt(StateCode) = xnull(tempRst!StateCode): Txt(StateName) = xnull(tempRst!StateName)
Else
    'Txt(1) = "": Txt(0) = ""
End If
Set tempRst = Nothing
End Sub
Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case CityCode, CityName
            If FrState.Visible = True Then FrState.Visible = False
            FrCity.Visible = True
            FrCity.top = Txt(Index).top + Txt(Index).height + 10
            FrCity.left = Txt(Index).left
            FrCity.ZOrder 0
        Case StateName
            If FrCity.Visible = True Then FrCity.Visible = False
            FrState.Visible = True
            FrState.top = Txt(Index).top + Txt(Index).height + 10
            FrState.left = Txt(Index).left
            FrState.ZOrder 0
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
mFlag = 0
    If FrCity.Visible = True Then FrCity.Visible = False
    If FrState.Visible = True Then FrState.Visible = False
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    Select Case Index
        Case CityCode, CityName
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
        Case StateName
            If RstState.BOF Or RstState.EOF Then Exit Sub
    End Select
    DGCity.Columns(0).width = 1000.1: DGCity.Columns(1).width = 3435.024: DGCity.Columns(2).width = 1000.1
    Select Case Index
        Case CityCode
            DGCity.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "searchcode ASC"
            RstHelp.Bookmark = mBookMark
            cityCodeSearch
        Case CityName
            DGCity.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "CITYNAME ASC"
            RstHelp.Bookmark = mBookMark
            cityNameSearch
        Case StateName
            DGState.Columns(0).width = 0: DGState.Columns(2).width = 0
            mBookMark = RstState.Bookmark
            RstState.Sort = "STATENAME ASC"
            RstState.Bookmark = mBookMark
            stateNameSearch
    End Select
    If Txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
Select Case Index
    Case LocalCentral
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt_GotFocus Index: Txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_Click Index
                Txt_GotFocus Index
                Txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case StateName
        If FrState.Visible = True Then
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstState.BOF Then RstState.MovePrevious
                Case vbKeyDown
                    If Not RstState.EOF Then RstState.MoveNext
                Case 33
                    For I = 1 To 9
                        If Not RstState.BOF Then RstState.MovePrevious
                    Next
                Case 34
                    For I = 1 To 9
                        If Not RstState.EOF Then RstState.MoveNext
                    Next
                Case 13
                    SendKeysA vbKeyTab, True
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    RST_BOF_EOF RstState
                    If Not RstState.BOF And Not RstState.EOF Then
                        Txt(StateCode) = XNull(RstState!StateCode): Txt(StateName) = XNull(RstState!StateName)
                        Txt(StateName).SelStart = 0
                    End If
            End Select
        End If
End Select
Select Case Index
    Case CityCode
        'SiteCode Edit restricted
        KeyCode = RestrictKey(1, KeyCode, Txt(Index), Shift)
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case CityName
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case StateName
        If FrState.Visible = False Then
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case CityCode
        KeyAscii = RestrictKey(1, KeyAscii, Txt(Index), 0)
End Select

End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case CityCode
        cityCodeSearch
    Case CityName
        cityNameSearch
    Case StateName
        stateNameSearch
    Case LocalCentral
        If Len(Txt(LocalCentral)) = 0 Or UCase(mID(Txt(LocalCentral), 1, 1)) = "L" Then
            Txt(LocalCentral) = "Local"
        ElseIf UCase(mID(Txt(LocalCentral), 1, 1)) = "C" Then
            Txt(LocalCentral) = "Central"
        Else
            Txt(LocalCentral) = "Local"
        End If
End Select
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case CityCode
            If PubBackEnd = "A" Then
                Set Rst = GCn.Execute("SELECT * FROM CITY WHERE " & cMID("CITYCODE", "2", "Len(CityCode)") & " = " & Chk_Text(Txt(CityCode)))
            ElseIf PubBackEnd = "S" Then
                Set Rst = GCn.Execute("SELECT * FROM CITY WHERE SubString(CITYCODE, 2, Len(CityCode))= " & Chk_Text(Txt(CityCode)))
            End If
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "City Code Already Exists", vbInformation, "Validation": Txt(CityCode) = Txt(CityCode).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!CityCode <> RstMain!CityCode Then MsgBox "City Code Already Exists", vbInformation, "Validation": Txt(CityCode) = Txt(CityCode).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case CityName
            Set Rst = GCn.Execute("SELECT * FROM CITY WHERE CITYNAME=" & Chk_Text(Txt(CityName)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "City Name Already Exists", vbInformation, "Validation": Txt(CityName) = Txt(CityName).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!CityName <> RstMain!CityName Then MsgBox "City Name Already Exists", vbInformation, "Validation": Txt(CityName) = Txt(CityName).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case StateName
            If Not RstState.EOF And Not RstState.BOF Then
                Txt(StateCode) = XNull(RstState!StateCode): Txt(StateName) = XNull(RstState!StateName)
            Else
                Txt(StateCode) = "": Txt(StateName) = ""
            End If
    End Select
Set Rst = Nothing
End Sub
Private Sub DGState_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    Txt(StateCode) = DGState.Columns(0).TEXT: Txt(StateName) = DGState.Columns(1).TEXT
End If
End Sub
Private Sub DGState_GotFocus()
    mFlag = 1
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select CityCode as SearchCode,CityName From City Where CityCode = '" & MyValue & "' Order by CityName")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

