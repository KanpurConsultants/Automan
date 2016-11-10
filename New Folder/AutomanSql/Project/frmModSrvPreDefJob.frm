VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmModSrvPreDefJob 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Model-wise Service-wise Pre-Defined Job Details"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11820
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   2
      Left            =   7965
      MaxLength       =   40
      TabIndex        =   10
      Top             =   825
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DGLab 
      Height          =   4605
      Left            =   4995
      Negotiate       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5445
      Visible         =   0   'False
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   8123
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "CODE"
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
         Caption         =   "Description"
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
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSer 
      Height          =   4440
      Left            =   1200
      Negotiate       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5940
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7832
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "code"
         Caption         =   "Ser. Code"
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
         DataField       =   "Name"
         Caption         =   "Service"
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
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            DividerStyle    =   3
            ColumnWidth     =   30.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4470.236
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGlabM 
      Height          =   4425
      Left            =   60
      Negotiate       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6180
      Visible         =   0   'False
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   7805
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "code"
         Caption         =   "Model"
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
         DataField       =   "ListName"
         Caption         =   "Model Description"
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
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4635.213
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   1755
      MaxLength       =   15
      TabIndex        =   1
      Top             =   885
      Width           =   3000
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
      Height          =   225
      Index           =   1
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1140
      Width           =   4515
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   3
      Left            =   6330
      MaxLength       =   15
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1125
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   4665
      Left            =   195
      TabIndex        =   9
      Top             =   1725
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   8229
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   7
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   33023
      GridColorUnpopulated=   12640511
      FocusRect       =   0
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"frmModSrvPreDefJob.frx":0000
      RowSizingMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service*"
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
      Height          =   240
      Index           =   0
      Left            =   540
      TabIndex        =   3
      Top             =   1155
      Width           =   750
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model*"
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
      Height          =   255
      Index           =   4
      Left            =   540
      TabIndex        =   6
      Top             =   885
      Width           =   600
   End
End
Attribute VB_Name = "frmModSrvPreDefJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstLab As ADODB.Recordset, RstSer As ADODB.Recordset
Dim mFlag As Byte
Private Const Model = 0
Private Const Serv = 1
Private Const ServC = 3
Dim GridKey As Integer
Dim Gtf As Boolean
' Col Declaration
Dim ExitCtrl As Boolean
Private Const Lab_Code As Byte = 1
Private Const LCode As Byte = 2
Private Const LGroup As Byte = 3
Private Const LType As Byte = 4
Private Const ChHrs As Byte = 5
Private Const ChRate As Byte = 6
Private Const WrHrs As Byte = 7
Dim TAddMode As Boolean

Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Sub DGLab_Click()
DGLab.Visible = False
Fill_Data
Txt(2).Visible = False
End Sub

Private Sub DGlabM_Click()
Txt(Model).TEXT = RstHelp!Name
Txt(Model).Tag = RstHelp!Name
Txt(Model).SetFocus
DGlabM.Visible = False
End Sub

Private Sub DGSer_Click()
Txt(Serv).TEXT = RstSer!Name
Txt(Serv).Tag = RstSer!Name
Txt(Serv).SetFocus
DGSer.Visible = False
End Sub

Private Sub FGrid_Click()
Txt(2).Visible = False
'DGlabM_Click
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
FGrid.BackColorSel = BackColorSelEnter
FGrid.ForeColorSel = ForeColorSelEnter
Grid_Hide

End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
Dim result As Boolean
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And FGrid.Row = (FGrid.Rows - (FGrid.Rows - 1)) And TopCtrl1.TopText2 = "Add" Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And FGrid.Row = FGrid.Rows - 1 Then
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
    Else
        FGrid.SetFocus
    End If
    Exit Sub
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case ChHrs, ChRate, WrHrs
            Call Get_Text(Me, FGrid, Txt, 2, True, 48)
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
SetMaxLength
Select Case FGrid.Col
    Case Lab_Code, LCode
       Call Get_Text(Me, FGrid, Txt, 2, False, keyascii)
    Case ChHrs, ChRate, WrHrs
       Call Get_Text(Me, FGrid, Txt, 2, True, keyascii)
End Select
If keyascii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.Rows > 2 Then
                FGrid.RemoveItem (FGrid.Row)
            Else
                FGrid.Rows = 1
                FGrid.AddItem ""
                FGrid.FixedRows = 1
            End If
         End If
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
Exit Sub
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
Txt(2).Visible = False
Grid_Hide
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
WinSetting Me:    Ini_Grid
TopCtrl1.Tag = PubUParam
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select (trim(max(model))+trim(MAX(Labour_CheckList.SERV_TYPE))) AS Searchcode,max(model) as model,MAX(Labour_CheckList.SERV_TYPE) AS SERV ,max(service_type.serv_desc) as s_desc From (Labour_CheckList LEFT JOIN SERVICE_TYPE ON SERVICE_TYPE.SERV_TYPE=LABOUR_CheckList.SERV_TYPE) where Labour_CheckList.SITE_CODE=" & Chk_Text(PubSiteCode) & "group by LABOUR_CHECKLIST.SERV_TYPE,MODEL order by model", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select (" & cTrim("max(model)") & "+" & cTrim("MAX(Labour_CheckList.SERV_TYPE)") & ") AS Searchcode,max(model) as model,MAX(Labour_CheckList.SERV_TYPE) AS SERV ,max(service_type.serv_desc) as s_desc From (Labour_CheckList LEFT JOIN SERVICE_TYPE ON SERVICE_TYPE.SERV_TYPE=LABOUR_CheckList.SERV_TYPE) Group by LABOUR_CHECKLIST.SERV_TYPE,MODEL order by model", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 (" & cTrim("max(model)") & "+" & cTrim("MAX(Labour_CheckList.SERV_TYPE)") & ") AS Searchcode,max(model) as model,MAX(Labour_CheckList.SERV_TYPE) AS SERV ,max(service_type.serv_desc) as s_desc From (Labour_CheckList LEFT JOIN SERVICE_TYPE ON SERVICE_TYPE.SERV_TYPE=LABOUR_CheckList.SERV_TYPE) Group by LABOUR_CHECKLIST.SERV_TYPE,MODEL order by model", GCn, adOpenDynamic, adLockOptimistic
End If
Set RstHelp = New ADODB.Recordset
'RstHelp.Open "Select MODEL as code,model as name ,model_desc as Listname FROM Model where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by MODEL", GCn, adOpenDynamic, adLockOptimistic
RstHelp.Open "Select MODEL as code,model as name ,model_desc as Listname FROM Model Order by MODEL", GCn, adOpenDynamic, adLockOptimistic
Set RstSer = New ADODB.Recordset
'RstSer.Open "Select SERV_TYPE as code,SERV_DESC as name  FROM SERVICE_TYPE where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by SERV_DESC", GCn, adOpenDynamic, adLockOptimistic
RstSer.Open "Select SERV_TYPE as code,SERV_DESC as name  FROM SERVICE_TYPE Order by SERV_DESC", GCn, adOpenDynamic, adLockOptimistic
Set DGlabM.DataSource = RstHelp
Set DGSer.DataSource = RstSer
Disp_Text SETS("INI", Me, RstMain)
CtrlClckCol
MoveRec
ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
Txt(Model).Tag = Txt(Model)
Txt_GotFocus Model
ADDFLAG = 1
FGrid.Rows = 1
FGrid.AddItem ""
FGrid.FixedRows = 1
Txt(Model).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
Dim Rs As Recordset
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(Model).Enabled = False
    Txt(Serv).Enabled = False
    ADDFLAG = 2
    FGrid.AddItem "" & STR(FGrid.Rows)
    FGrid.SetFocus
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
            GCn.Execute ("delete  from Labour_CHECKLIST where MODEL= '" & Trim(Txt(Model)) & "' and serv_type='" & Txt(ServC) & "' and site_code='" & PubSiteCode & "'")
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
'    GSQL = "Select (trim(max(model))+trim(MAX(Labour_CheckList.SERV_TYPE))) AS Searchcode,max(model) as model,max(service_type.serv_desc) as s_desc From (Labour_CheckList LEFT JOIN SERVICE_TYPE ON SERVICE_TYPE.SERV_TYPE=LABOUR_CheckList.SERV_TYPE) where Labour_CheckList.SITE_CODE=" & Chk_Text(PubSiteCode) & "group by LABOUR_CHECKLIST.SERV_TYPE,MODEL order by model"
    GSQL = "Select (" & cTrim("max(model)") & " + " & cTrim("MAX(Labour_CheckList.SERV_TYPE)") & ") AS Searchcode,max(model) as model,max(service_type.serv_desc) as s_desc From (Labour_CheckList LEFT JOIN SERVICE_TYPE ON SERVICE_TYPE.SERV_TYPE=LABOUR_CheckList.SERV_TYPE) Group by LABOUR_CHECKLIST.SERV_TYPE,MODEL order by model"
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
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select  (" & cTrim("max(model)") & "+" & cTrim("MAX(Labour_CheckList.SERV_TYPE)") & ") AS Searchcode,max(model) as model,MAX(Labour_CheckList.SERV_TYPE) AS SERV ,max(service_type.serv_desc) as s_desc From (Labour_CheckList LEFT JOIN SERVICE_TYPE ON SERVICE_TYPE.SERV_TYPE=LABOUR_CheckList.SERV_TYPE) Where (" & cTrim("max(model)") & "+" & cTrim("MAX(Labour_CheckList.SERV_TYPE)") & ") = '" & MyValue & "' Group by LABOUR_CHECKLIST.SERV_TYPE,MODEL order by model")
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

    mRepName = "ModSrvPreDefJob"
    mQRY = "SELECT '" & RstMain!Model & "','" & RstMain!Serv & "','" & RstMain!S_Desc & _
        "',LG.Lab_Group AS GCode, LG.LabGrp_Desc as LGroup," & _
        " LT.Lab_Type AS TypeCode,LT.Lab_Desc as TypeDesc," & _
        " LC.Lab_Code,L.Lab_Desc,L.U_Name,L.U_EntDt,L.U_AE " & _
        " FROM ((Labour_Checklist LC left join labour L on LC.lab_code=L.lab_code) " & _
        " left join labour_type LT on L.lab_type=LT.lab_type) " & _
        " left join labour_group LG on L.lab_group=LG.lab_group " & _
        " WHERE LC.MODEL='" & Txt(Model) & "' AND LC.SERV_TYPE='" & Txt(ServC) & _
        "' Order By L.lab_group,L.lab_type,LC.lab_code"
    
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
     rpt.Database.SetDataSource Rst
     rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , False)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, I As Integer

On Error GoTo Errloop
    If Txt(2).Visible = True Then
        If TxtGridLeave = False Then Txt_LostFocus 2: Exit Sub
    End If
    Grid_Hide
    If IsValid(Txt(Model), "Model Number") = False Then Exit Sub
    If IsValid(Txt(Serv), "Service Type") = False Then Exit Sub
    GCn.BeginTrans
    mTrans = True
    GCn.Execute ("DELETE From Labour_checklist Where MODEL=" & Chk_Text(Trim(Txt(Model))) & " and serv_type=" & Chk_Text(Trim(Txt(ServC))) & "")
    For I = 1 To FGrid.Rows - 1
        If Len(FGrid.TextMatrix(I, Lab_Code)) <> 0 Then
            GCn.Execute ("Insert Into Labour_checklist(MODEL,Site_Code,Lab_code,serv_type,U_Name,U_EntDt,U_AE) Values('" & Txt(Model) & "','" & PubSiteCode & "','" & FGrid.TextMatrix(I, Lab_Code) & "','" & Txt(ServC) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
        End If
    Next
    GCn.CommitTrans
    mTrans = False
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select (" & cTrim("max(model)") & "+" & cTrim("MAX(Labour_CheckList.SERV_TYPE)") & ") AS Searchcode,max(model) as model,MAX(Labour_CheckList.SERV_TYPE) AS SERV ,max(service_type.serv_desc) as s_desc From (Labour_CheckList LEFT JOIN SERVICE_TYPE ON SERVICE_TYPE.SERV_TYPE=LABOUR_CheckList.SERV_TYPE) Where (" & cTrim("max(model)") & "+" & cTrim("MAX(Labour_CheckList.SERV_TYPE)") & ") = '" & Trim(Txt(Model)) + Trim(Txt(ServC)) & "' Group by LABOUR_CHECKLIST.SERV_TYPE,MODEL order by model")
    End If
    RstHelp.Requery
    RstMain.FIND ("Searchcode='" & Trim(Txt(Model)) + Trim(Txt(ServC)) & "'")
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, RstMain)
    Call MoveRec

'    If TopCtrl1.TopText2.Caption = "Add" Then
'        Disp_Text SETS("INI", Me, RstMain)
'        FGrid.CellBackColor = CellBackColLeave
'    Else
'        Disp_Text SETS("INI", Me, RstMain)
'        MoveRec
        CtrlClckCol
'        AddFlag = 0
'        DGLabM.Visible = False
'    End If
Exit Sub
Errloop:    If mTrans Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        ADDFLAG = 0
        Grid_Hide
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
        CtrlClckCol
        DGLab.Visible = False
        DGlabM.Visible = False
        DGSer.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    Txt(Model).BackColor = CtrlBColOrg:      Txt(Model).ForeColor = CtrlFColOrg
    Txt(Serv).BackColor = CtrlBColOrg:      Txt(Serv).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
Dim Rs As Recordset
On Error GoTo Errloop
RST_BOF_EOF RstMain
Txt(2).Visible = False
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    Txt(Model) = XNull(RstMain!Model)
    Txt(ServC) = XNull(RstMain!Serv)
    Txt(Serv) = XNull(RstMain!S_Desc)
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT LABOUR_CHECKLIST.LAB_CODE,labour.lab_desc as ldesc,labour_type.lab_desc as ltype,LABOUR_TYPE.LAB_TYPE AS TCODE,labour_group.labgrp_desc as lgroup,LABOUR_GROUP.LAB_GROUP AS GCODE FROM ((LABOUR_CHECKLIST left join labour on LABOUR_CHECKLIST.lab_code=labour.lab_code) left join labour_type on labour.lab_type=labour_type.lab_type) left join labour_group on labour.lab_group=labour_group.lab_group WHERE LABOUR_CHECKLIST.MODEL='" & Txt(Model) & "' AND LABOUR_CHECKLIST.SERV_TYPE='" & Txt(ServC) & "'")
    If Rs.RecordCount > 0 Then
        FGrid.Rows = 1
        Do Until Rs.EOF
            FGrid.AddItem "" & FGrid.Rows & Chr(9) & Rs!Lab_Code & Chr(9) & Rs!lDesc & Chr(9) & Rs!LType & Chr(9) & Rs!LGroup & Chr(9) & Rs!tCode & Chr(9) & Rs!GCODE
            Rs.MoveNext
        Loop
        FGrid.FixedRows = 1
    End If
End If
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

Private Sub Txt_GotFocus(Index As Integer)
Dim TStr$, mROW As Integer
Grid_Hide
If Index = 2 Then
    If Txt(ServC) = "" Then Exit Sub
    Txt(Index).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case Lab_Code, LCode
            TStr = ""
            Do Until mROW = FGrid.Rows - 1
                mROW = mROW + 1
                If mROW <> FGrid.Row Then
                    TStr = TStr + "'" + FGrid.TextMatrix(mROW, Lab_Code) + "'" + ","
                End If
            Loop
             Set RstLab = GCn.Execute("SELECT LABOUR.LAB_CODE AS CODE,LABOUR.LAB_DESC AS NAME," & _
                "labour_type.lab_desc as ltype,LABOUR_TYPE.LAB_TYPE AS TCODE," & _
                "labour_group.labgrp_desc as lgroup,LABOUR_GROUP.LAB_GROUP AS GCODE " & _
                "FROM (labour left join labour_type on labour.lab_type=labour_type.lab_type) " & _
                "left join labour_group on labour.lab_group=labour_group.lab_group " & _
                "where lab_code  NOT in (" & TStr & ")")
            Set DGLab.DataSource = RstLab
            If FGrid.Col = Lab_Code Then
                RstLab.Sort = "CODE"
                RstLab.FIND "code  >='" & FGrid.TextMatrix(FGrid.Row, Lab_Code) & "'"
            Else
                RstLab.Sort = "name"
                RstLab.FIND "name  >='" & FGrid.TextMatrix(FGrid.Row, LCode) & "'"
            End If
            If RstLab.RecordCount > 0 Then
                If RstLab.EOF = True Then RstLab.MoveFirst
            End If
    End Select
End If
Ctrl_GetFocus Txt(Index)
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Dim I As Byte
Dim Txtdate As Boolean
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case Model
        If KeyCode <> vbKeyEscape Then
            DGridTxtKeyDown DGlabM, Txt, Model, RstHelp, KeyCode, False, 0, frmGodown, "frmGodown"
        End If
    Case Serv
        If KeyCode <> vbKeyEscape Then
            DGridTxtKeyDown DGSer, Txt, Serv, RstSer, KeyCode, False, 1, frmService, "frmService"
        End If
        If KeyCode = vbKeyUp And DGSer.Visible = False Then
            SendKeys "+{tab}"
        End If
    Case 2
        If KeyCode = vbKeyEscape Then
            FGrid.SetFocus
            Txt(Index).TEXT = Txt(Index).Tag
            'Txt_KeyUp Index, KeyCode, Shift
            Txt(Index).Visible = False
            DGLab.Visible = False
            Exit Sub
        End If
        Select Case FGrid.Col
            Case Lab_Code    '1
                If DGLab.Visible = False Then DGridColSwap DGLab, 0
                DGridTxtKeyDown DGLab, Txt, 2, RstLab, KeyCode, False, 0, frmLabDesc, "frmLabDesc"
                If DGLab.Visible = True Then Gtf = True
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, Txt, 2, KeyCode, TAddMode, Lab_Code ', , , , True
                    Else
                        Txt_LostFocus 0
                        Txt(Index).SetFocus
                    End If
                    Txt(2).Visible = False
                End If
            Case LCode
                If DGLab.Visible = False Then DGridColSwap DGLab, 1
                DGridTxtKeyDown DGLab, Txt, 2, RstLab, KeyCode, True, 1, frmLabDesc, "frmLabDesc"
                If DGLab.Visible = True Then Gtf = True
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, Txt, 2, KeyCode, TAddMode, LGroup, , Lab_Code
                    Else
                        Txt_LostFocus 0
                        Txt(Index).SetFocus
                        Txt(2).Visible = False
                    End If
                End If
            Case ChHrs, ChRate, WrHrs
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                     If TxtGridLeave = True Then
                          GridTxtDown FGrid, Txt, Index, KeyCode, TAddMode, WrHrs    ', 3
                     Else
                          Txt_LostFocus 0
                          Txt(0).SetFocus
                     End If
                 End If
        End Select
End Select
If Index <> 2 And DGlabM.Visible = False And DGSer.Visible = False Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case Index
    Case Model
        If DGlabM.Visible = True Then DGridTxtKeyPress Txt, Model, RstHelp, keyascii, "Code"
    Case Serv
        If DGSer.Visible = True Then DGridTxtKeyPress Txt, Serv, RstSer, keyascii, "NAME"
    Case 2
        Select Case Val(Txt(Index).Tag)
            Case ChHrs, ChRate, WrHrs
                Call NumPress(Txt(2), keyascii, 8, 2)
        End Select
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Serv
'           If DGSer.Visible = True Then DGridTxtKeyUp Txt, Serv, RstSer, KeyCode, "NAME"
           If Not RstSer.BOF And Not RstSer.EOF Then Txt(3).TEXT = RstSer.Fields(0)
    Case 2
        Select Case FGrid.Col
            Case Lab_Code
                If KeyCode <> 13 And DGLab.Visible = False Then Txt_KeyDown Index, GridKey, 0
'                If RstLab.RecordCount  > 0 Then DGridTxtKeyUp Txt, Index, RstLab, KeyCode, "CODE"
                If RstLab.RecordCount > 0 Then DGridTxtKeyPress Txt, Index, RstLab, KeyCode, "CODE", True
            Case LCode
                If KeyCode <> 13 And DGLab.Visible = False Then Txt_KeyDown Index, GridKey, 0
'                If RstLab.RecordCount  > 0 Then DGridTxtKeyUp Txt, Index, RstLab, KeyCode, "name"
                If RstLab.RecordCount > 0 Then DGridTxtKeyPress Txt, Index, RstLab, KeyCode, "name", True
            Case ChHrs
                FGrid.TextMatrix(FGrid.Row, ChHrs) = Format(Val(Txt(Index).TEXT), "0.00")
            Case ChRate
                FGrid.TextMatrix(FGrid.Row, ChRate) = Format(Val(Txt(Index).TEXT), "0.00")
            Case WrHrs
                FGrid.TextMatrix(FGrid.Row, WrHrs) = Format(Val(Txt(Index).TEXT), "0.00")
        End Select
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
'If Index = Model Then
'    DGlabM_Click
'    'DGlabM.Visible = False
'ElseIf Index = Serv Then
'    DGSer_Click
'End If
Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rs As Recordset, Rst As Recordset
Select Case Index
    Case Serv
        Set Rst = GCn.Execute("Select max(model),max(serv_type) as model From labour_CHECKLIST where MODEL='" & Txt(Model) & "' AND SERV_TYPE='" & Txt(ServC) & "' AND SITE_CODE=" & Chk_Text(PubSiteCode) & " group by serv_type,MODEL")
        If ADDFLAG = 1 Then
            If Not Rst.EOF Then MsgBox "For this Model and Service Labour Details Already Exists", vbInformation, "Validation": Txt(Model) = Txt(Model).Tag: Cancel = True: Exit Sub
        End If
        If RstHelp.RecordCount = 0 Then Exit Sub
'        If DGlabM.Visible = True Then
'            Txt(MODEL).Text = RstHelp!Name
'            Txt(MODEL).Tag = RstHelp!code
'        End If
        If ADDFLAG = 1 Then
            Set Rs = New Recordset
            Set Rs = GCn.Execute("SELECT LABOUR.LAB_CODE AS LCODE,LABOUR.LAB_DESC AS ldesc,labour_type.lab_desc as ltype,LABOUR_TYPE.LAB_TYPE AS TCODE,labour_group.labgrp_desc as lgroup,LABOUR_GROUP.LAB_GROUP AS GCODE FROM (labour left join labour_type on labour.lab_type=labour_type.lab_type) left join labour_group on labour.lab_group=labour_group.lab_group WHERE LABOUR.SITE_CODE='" & PubSiteCode & "'")
                If Rs.RecordCount > 0 Then
                    FGrid.Rows = 1
                    Do Until Rs.EOF
                        FGrid.AddItem FGrid.Rows & Chr(9) & Rs!LCode & Chr(9) & Rs!lDesc & Chr(9) & Rs!LType & Chr(9) & Rs!LGroup
                        Rs.MoveNext
                    Loop
                    FGrid.FixedRows = 1
                End If
        End If
'       FGrid.Row = 1
'       FGrid.Col = 5
   Case 2
        Select Case FGrid.Col
            Case WrHrs
                 If FGrid.Row < FGrid.Rows - 1 Then FGrid.Row = FGrid.Row + 1
                 FGrid.Col = 4
            Case Lab_Code
                If FGrid.TextMatrix(FGrid.Row, Lab_Code) <> Txt(2) Then Call Fill_Data
            Case LCode
                If FGrid.TextMatrix(FGrid.Row, LCode) <> Txt(2) Then Call Fill_Data
       End Select
End Select
End Sub

Private Sub DGLab_GotFocus()
    mFlag = 1
End Sub

Private Sub DGLabM_GotFocus()
    mFlag = 1
End Sub

Private Sub DGSER_GotFocus()
    mFlag = 1
End Sub
Private Sub BlankText()
Dim I As Byte
    Txt(0).TEXT = ""
    Txt(1).TEXT = ""
    Txt(2).TEXT = ""
    FGrid.Rows = 1
    FGrid.AddItem ""
    FGrid.FixedRows = 1
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    Txt(0).Enabled = Enb
    Txt(1).Enabled = Enb
    Txt(2).Enabled = Enb
'    FGrid.Enabled = Enb
End Sub

Private Sub Ini_Grid()

    With FGrid
        .width = Me.width - 120
        .left = 0
        .top = 1725
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 7
        .height = .RowHeight(0) * 20
        .AllowUserResizing = flexResizeNone
        .TextMatrix(0, 0) = ""
        .ColAlignmentFixed(0) = flexAlignRightCenter
        .ColWidth(0) = 400
        
        .TextMatrix(0, Lab_Code) = "Code"
        .ColAlignment(Lab_Code) = flexAlignLeftCenter
        .ColWidth(Lab_Code) = 1000
        
        .TextMatrix(0, LCode) = "Labour Description"
        .ColAlignment(LCode) = flexAlignLeftCenter
        .ColWidth(LCode) = 4250
        
        .TextMatrix(0, LType) = "Type"
        .ColAlignmentFixed(LType) = flexAlignCenterCenter
        .ColAlignment(LType) = flexAlignLeftCenter
        .ColWidth(LType) = 2000

        .TextMatrix(0, LGroup) = "Group"
        .ColAlignmentFixed(LGroup) = flexAlignCenterCenter
        .ColAlignment(LGroup) = flexAlignLeftCenter
        .ColWidth(LGroup) = 2000

'        .TextMatrix(0, ChHrs) = "ChrgHrs"
'        .ColAlignmentFixed(chHrs) = flexAlignCenterCenter
'        .ColAlignment(ChHrs) = flexAlignRightCenter
'        .ColWidth(ChHrs) = 600
'
'        .TextMatrix(0, ChAmt) = "ChrgAmt"
'        .ColAlignmentFixed(ChAmt) = flexAlignCenterCenter
'        .ColAlignment(ChAmt) = flexAlignRightCenter
'        .ColWidth(ChAmt) = 795
'
'        .TextMatrix(0, WrHrs) = "Wr.Hrs."
'        .ColAlignmentFixed(WrHrs) = flexAlignCenterCenter
'        .ColAlignment(WrHrs) = flexAlignRightCenter
'        .ColWidth(WrHrs) = 600

        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel

    DGSer.left = Me.width - (DGSer.width + mRtScale): DGSer.top = mTopScale: DGSer.height = FGrid.height + FGrid.top - mTopScale ' Me.Height - (DGSer.top + mBotScale)
'    dglab.width = 7000:
    DGLab.left = Me.width - (DGLab.width + mRtScale): DGLab.top = mTopScale: DGLab.height = Me.height - (DGLab.top + mBotScale)
    DGlabM.width = Me.width - 120: DGlabM.left = Me.left: DGlabM.top = FGrid.top: DGlabM.height = Me.height - (DGlabM.top + mBotScale)
    DGlabM.Columns(0).width = 2000
    DGlabM.Columns(1).width = 6500
End Sub

Private Function TxtGridLeave() As Boolean
Select Case FGrid.Col
    Case Lab_Code, LCode
        If RstLab.RecordCount = 0 Then
            TxtGridLeave = False: ExitCtrl = False: Gtf = False: DGLab.Visible = False: Exit Function
        End If
        If Gtf = True Then Fill_Data
        Gtf = False
End Select
FGrid.SetFocus
ExitCtrl = True
TxtGridLeave = True
Txt(2).Visible = False
End Function

Private Sub Fill_Data()
If RstLab.RecordCount = 0 Then Exit Sub
   FGrid.TextMatrix(FGrid.Row, Lab_Code) = RstLab!Code
   FGrid.TextMatrix(FGrid.Row, LCode) = RstLab!Name
   FGrid.TextMatrix(FGrid.Row, LType) = XNull(RstLab!LType)
   FGrid.TextMatrix(FGrid.Row, LGroup) = RstLab!LGroup
End Sub

Private Sub Grid_Hide()
If DGLab.Visible = True Then DGLab.Visible = False
If DGlabM.Visible = True Then DGlabM.Visible = False
If DGSer.Visible = True Then DGSer.Visible = False
End Sub

Private Sub SetMaxLength()
Select Case FGrid.Col
    Case Lab_Code
        Txt(2).MaxLength = 6
    Case LCode
        Txt(2).MaxLength = 40
    Case Else
        Txt(2).MaxLength = 0
End Select
End Sub
