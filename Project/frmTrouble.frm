VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmTrouble 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Trouble Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmTrouble.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.Frame FrJob1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   3945
      TabIndex        =   14
      Top             =   3015
      Visible         =   0   'False
      Width           =   5220
      Begin MSDataGridLib.DataGrid DGJob1 
         Height          =   3285
         Left            =   15
         TabIndex        =   15
         Top             =   270
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   5794
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483648
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Code"
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
            DataField       =   "name"
            Caption         =   "Trouble"
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
            DataField       =   "Code"
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
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3089.764
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Troubles"
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
         Left            =   0
         TabIndex        =   16
         Top             =   -15
         Width           =   5175
      End
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   6870
      TabIndex        =   12
      Top             =   3105
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   15
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   16379351
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3942
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
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
      Height          =   210
      Index           =   3
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1410
      Width           =   2400
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
      Height          =   210
      Index           =   2
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1170
      Width           =   2400
   End
   Begin VB.Frame FrJob 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   1575
      TabIndex        =   7
      Top             =   3105
      Visible         =   0   'False
      Width           =   5220
      Begin MSDataGridLib.DataGrid DGJob 
         Height          =   3285
         Left            =   15
         TabIndex        =   8
         Top             =   270
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   5794
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483648
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Code"
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
            DataField       =   "name"
            Caption         =   "Trouble"
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
            DataField       =   "Code"
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
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3089.764
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Troubles"
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
         Left            =   0
         TabIndex        =   9
         Top             =   -15
         Width           =   5175
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
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
      Height          =   210
      Index           =   1
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   2
      Top             =   930
      Width           =   4680
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
      Height          =   210
      Index           =   0
      Left            =   2805
      MaxLength       =   6
      TabIndex        =   1
      Top             =   690
      Width           =   900
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trouble Type"
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
      Left            =   1185
      TabIndex        =   11
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trouble Related"
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
      Left            =   1185
      TabIndex        =   10
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trouble*"
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
      Left            =   1185
      TabIndex        =   6
      Top             =   960
      Width           =   750
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code*"
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
      Left            =   1185
      TabIndex        =   5
      Top             =   720
      Width           =   555
   End
End
Attribute VB_Name = "frmTrouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim ListArray As Variant
Dim mListItem As ListItem
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstHelp1 As ADODB.Recordset
Dim mFlag As Byte
Dim Troubletype(2, 7) As String
Private Const Code = 0, Desc = 1, Trelated = 2, Ttype = 3

Private Sub DGJob1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        FrJob1.Visible = False
        txt(Ttype).SetFocus
    End If
End Sub

Private Sub DGJob1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DGJob1.Col = 1
txt(Trelated) = DGJob1.TEXT
DGJob1.Col = 0
txt(Trelated).Tag = DGJob1.TEXT
End Sub

Private Sub Form_Activate()
Dim I As Integer
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
For I = 0 To 6
    With RstHelp1
        .AddNew
        !Code = Troubletype(0, I)
        !Name = Troubletype(1, I)
        .Update
    End With
Next
Set DGJob1.DataSource = RstHelp1
FrJob.Visible = False: FrJob1.Visible = False
End Sub

Private Sub Form_Load()
Me.top = 0: Me.left = 0
TopCtrl1.Tag = PubUParam ': TopCtrl1.TopText1 = "Trouble Master"   ': TopCtrl1.TopText1.Width = 1000
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select * From Trouble  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by Trouble_Code", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select Trouble_Code as SearchCode, Trouble.* From Trouble Order by Trouble_Code", GCn, adOpenDynamic, adLockOptimistic
Else
    Set RstMain = GCn.Execute("Select Top 1 Trouble_Code as SearchCode, Trouble.* From Trouble Order by Trouble_Code")
End If

Set RstHelp = New ADODB.Recordset
Set RstHelp1 = New ADODB.Recordset
'RstHelp.Open "Select TROUBLE_CODE AS CODE,TROUBLE_NAME as name FROM Trouble  where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by TROUBLE_code", GCn, adOpenDynamic, adLockOptimistic
RstHelp.Open "Select TROUBLE_CODE AS CODE,TROUBLE_NAME as name FROM Trouble Order by TROUBLE_code", GCn, adOpenDynamic, adLockOptimistic

Troubletype(0, 0) = "001": Troubletype(1, 0) = "Engine Related"
Troubletype(0, 1) = "002": Troubletype(1, 1) = "Gear-Box Related"
Troubletype(0, 2) = "003": Troubletype(1, 2) = "Suspension/Steering Related"
Troubletype(0, 3) = "004": Troubletype(1, 3) = "AC Related"
Troubletype(0, 4) = "005": Troubletype(1, 4) = "Electrical Related"
Troubletype(0, 5) = "006": Troubletype(1, 5) = "Body Shell Related"
Troubletype(0, 6) = "007": Troubletype(1, 6) = "Misclleneous"
Set RstHelp1 = New ADODB.Recordset
With RstHelp1
    .Fields.Append "Code", adVarChar, 3, adFldIsNullable
    .Fields.Append "Name", adVarChar, 40, adFldIsNullable
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With


Set DGJob.DataSource = RstHelp
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
CtrlClckCol
MoveRec
ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
txt(Code).Tag = txt(Code)
Txt_GotFocus Code
ADDFLAG = 1
txt(Code).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(Code).Enabled = False
    txt(Desc).Tag = txt(Desc)
    Txt_GotFocus Desc
    ADDFLAG = 2
    txt(Desc).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo Errloop
Dim transFalg As Byte
transFalg = 0
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            XBM = RstMain.Bookmark
            transFalg = 1
            GCn.Execute ("delete * from Trouble Where TROUBLE_Code= " & Chk_Text(Trim(txt(Code))))
            GCn.CommitTrans
            transFalg = 0
            RstMain.Requery
            RstHelp.Requery
            If RstMain.RecordCount >= XBM Then
                RstMain.Bookmark = XBM
            Else
                If RstMain.EOF = False Then RstMain.MoveLast
            End If
            Call MoveRec
        End If
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If

Exit Sub
Errloop:    If transFalg = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbExclamation, " Deletion Error "
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
    GSQL = "Select Trouble_Code as SearchCode, Trouble_Code as Code,Trouble_Name as Name FROM Trouble  Order by Trouble_Code"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("searchcode='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select Trouble_Code as SearchCode, Trouble.* From Trouble Where Trouble_Code = '" & MyValue & "' Order by Trouble_Code")
    End If
    
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "Trouble"
    mQRY = "SELECT * from Trouble"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
Dim RstHelp As ADODB.Recordset
rpt.Database.SetDataSource Rst
     rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , False)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo Errloop
    transFlag = 0
    If IsValid(txt(Code), "Code") = False Then Txt_GotFocus Code: Exit Sub
    If IsValid(txt(Desc), "Description") = False Then Txt_GotFocus Desc: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From Trouble Where TROUBLE_Code= " & Chk_Text(Trim(txt(Code))) & " AND SITE_CODE='" & PubSiteCode & "'").Fields(0) > 0 Then MsgBox "Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus Code: txt(Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        GCn.Execute ("DELETE From Trouble Where TROUBLE_Code= " & Chk_Text(Trim(txt(Code))) & " AND SITE_CODE='" & PubSiteCode & "'")
        GCn.Execute ("Insert Into Trouble(TROUBLE_Code,Div_Code,Site_Code,TROUBLE_NAME,U_Name,U_EntDt,U_AE,TRelated,TType) Values('" & txt(Code) & "','" & PubDivCode & "','" & PubSiteCode & "','" & txt(Desc) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "','" & txt(Trelated).Tag & "','" & txt(Ttype) & "')")
    ElseIf ADDFLAG = 2 Then
        GCn.Execute ("UPDATE Trouble SET Site_Code='" & PubSiteCode & "',TROUBLE_NAME=" & Chk_Text(txt(Desc)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "',TRelated='" & txt(Trelated).Tag & "',TType='" & txt(Ttype) & "' Where TROUBLE_Code= " & Chk_Text(Trim(txt(Code))) & "")
    End If

    GCn.CommitTrans
    'If MasterFormExit Then Unload Me: Exit Sub: Cancel = True
    
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select Trouble_Code as SearchCode, Trouble.* From Trouble Where Trouble_Code = " & Chk_Text(Trim(txt(Code))) & " Order by Trouble_Code")
    End If
    
    RstHelp.Requery
    RstMain.FIND ("TROUBLE_Code=" & Chk_Text(Trim(txt(Code))))
    
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus Code
        txt(Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrJob.Visible = False
        FrJob1.Visible = False
        FrmList.Visible = False
    End If
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        Me.ActiveControl.SetFocus
        CtrlClckCol
        FrJob.Visible = False
        FrJob1.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(Code).BackColor = CtrlBColOrg:      txt(Code).ForeColor = CtrlFColOrg
    txt(Desc).BackColor = CtrlBColOrg:      txt(Desc).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
Dim I As Integer, Trelate$
On Error GoTo Errloop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    txt(Code) = XNull(RstMain!trouble_code)
    txt(Desc) = XNull(RstMain!trouble_name)
    For I = 0 To 6
        If Trim(Troubletype(0, I)) = XNull(RstMain!Trelated) Then Trelate = Trim(Troubletype(1, I))
    Next
    txt(Trelated) = Trelate
    txt(Trelated).Tag = XNull(RstMain!Trelated)
    txt(Ttype) = XNull(RstMain!Ttype)
End If
TopCtrl1.tDel = False
Exit Sub
Errloop:        MsgBox err.Description
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub
Private Sub ColCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
    RstHelp.MoveFirst
    RstHelp.FIND "Code >=" & Chk_Text(Trim(txt(Code)))
End Sub
Private Sub ColNameSearch()
    If RstHelp.RecordCount <= 0 Then Exit Sub
        RstHelp.MoveFirst
        RstHelp.FIND "name >=" & Chk_Text(XNull(txt(Desc)))
End Sub
Private Sub ColRelateSearch()
If RstHelp1.RecordCount <= 0 Then Exit Sub
    RstHelp1.MoveFirst
    RstHelp1.FIND "Code >=" & Chk_Text(Trim(txt(Trelated)))
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Code, Desc
            If RstHelp.RecordCount = 0 Then Exit Sub
            If FrJob.Visible = True Then FrJob.Visible = False
            FrJob.Visible = True
            FrJob.top = txt(Index).top + txt(Index).height + 10
            FrJob.left = txt(Index).left
            FrJob.ZOrder 0
         Case Trelated
            If RstHelp1.RecordCount = 0 Then Exit Sub
            If FrJob1.Visible = True Then FrJob1.Visible = False
            If FrJob1.Visible = False Then FrJob1.Visible = True
            FrJob1.top = txt(Index).top + txt(Index).height + 10
            FrJob1.left = txt(Index).left
            FrJob1.ZOrder 0
            DGJob1.SetFocus
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
DGJob.Columns(0).width = 1000.1: DGJob.Columns(1).width = 3535.024: DGJob.Columns(2).width = 1000.1
Dim mBookMark
    Ctrl_GetFocus txt(Index)
mFlag = 0
    If FrJob.Visible = True Then FrJob.Visible = False
    If FrJob1.Visible = True Then FrJob1.Visible = False
    RST_BOF_EOF RstHelp
    txt(Index).Tag = txt(Index)
    Select Case Index
        Case Code, Desc
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case Code
            DGJob.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Code ASC"
            RstHelp.Bookmark = mBookMark
            ColCodeSearch
        Case Desc
            DGJob.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "NAME ASC"
            RstHelp.Bookmark = mBookMark
            ColNameSearch
        Case Trelated
            If FrJob1.Visible = True Then FrJob1.Visible = False
            
        Case Ttype
            ListArray = Array("Complaint", "Job")
            Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    End Select
    If txt(Index) = "" Then Txt_Change Index

End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
Select Case Index
    Case Ttype
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown And ListView.Visible = False Then
            FrJob.Visible = False: FrJob1.Visible = False: FrmList.Visible = False
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt_GotFocus Index: txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_GotFocus Index
                txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" And ListView.Visible = False Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
End Select
Select Case Index
     Case Ttype
        ListViewReport_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, txt(Index).top + txt(Index).height + 25, txt(Index).width, 2000
End Select
Select Case Index
    Case Code, Desc
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown And FrJob.Visible = False Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case Trelated
            DGJob1.Columns(0).width = 0
            RstHelp1.Sort = "NAME ASC"
            
        If KeyCode = 13 Or KeyCode = vbKeyTab Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case Ttype
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown And ListView.Visible = False Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
End Select
Select Case Index
    Case Code, Desc, Trelated
        If KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
End Select
If FrJob1.Visible = True Then FrJob1.Visible = False
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)

End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case Code
        ColCodeSearch
    Case Desc
        ColNameSearch
End Select

End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Code
            Set Rst = GCn.Execute("SELECT * FROM Trouble WHERE TROUBLE_Code=" & Chk_Text(Trim(txt(Code))))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox " Code Already Exists", vbInformation, "Validation": txt(Code) = txt(Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!trouble_code <> RstMain!trouble_code Then MsgBox "Code Already Exists", vbInformation, "Validation": txt(Code) = txt(Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Desc
            Set Rst = GCn.Execute("SELECT * FROM Trouble WHERE TROUBLE_NAME=" & Chk_Text(txt(Desc)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Readon Already Exists", vbInformation, "Validation": txt(Desc) = txt(Desc).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!trouble_name <> RstMain!trouble_name Then MsgBox "Description Already Exists", vbInformation, "Validation": txt(Desc) = txt(Desc).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
Next
End Sub
