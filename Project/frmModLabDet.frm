VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmModLabDet 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Model-wise Labour Details Master"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11700
   Begin MSDataGridLib.DataGrid DGlabM 
      Height          =   4410
      Left            =   3825
      Negotiate       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   7779
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
         Caption         =   "Vehicle_Type"
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
            DividerStyle    =   3
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   2
      Left            =   0
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   690
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
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
      Height          =   240
      Index           =   0
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   1
      Top             =   713
      Width           =   3000
   End
   Begin MSDataGridLib.DataGrid DGLab 
      Height          =   5190
      Left            =   3855
      Negotiate       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   9155
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4635.213
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   5250
      Left            =   165
      TabIndex        =   5
      Top             =   1350
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   9260
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   10
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   8421631
      FocusRect       =   0
      Appearance      =   0
      FormatString    =   " "
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
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
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmModLabDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstLab As ADODB.Recordset
Dim mFlag As Byte
Private Const VehType = 0
Dim GridKey As Integer
' Col Declaration
Dim ExitCtrl As Boolean
Private Const Lab_Code As Byte = 1
Private Const LCode As Byte = 2
Private Const LType As Byte = 3
Private Const LGroup As Byte = 4
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
txt(2).Visible = False
End Sub

Private Sub DGlabM_Click()
txt(VehType).TEXT = RstHelp!Name
txt(VehType).Tag = RstHelp!Name
DGlabM.Visible = False
End Sub

Private Sub FGrid_Click()
'DGlabM_Click
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
txt(2).Visible = False
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid_KeyPress vbKeyReturn
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
''Leave Cell-- > Enter Cell-- >KeyDown
'Dim result As Boolean
'If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) And TopCtrl1.TopText2 = "Add" Then
'    FGrid.CellBackColor = CellBackColLeave
'    SendKeys "+{Tab}"
'    KeyCode = 0
'End If
'    GridKey = KeyCode
'    FGrid.Tag = FGrid.Row
'Select Case FGrid.Col
'    Case LCode, LType, LGroup
'
'        If KeyCode = vbKeyDelete And Shift = 0 Then Exit Sub
'    Case ChHrs, ChRate, WrHrs
'        If KeyCode = vbKeyDelete And Shift = 0 Then Call Get_Text(Me, FGrid, txt, 2, True, 48)     'FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
'End Select
'    If KeyCode = vbKeyReturn Then
'        Select Case FGrid.Col
'            Case LCode, Lab_Code, ChHrs, ChRate, WrHrs
'                Call GridDblClick(Me, FGrid, txt, 2)
'                TAddMode = False
'        End Select
'    End If
'If KeyCode = vbKeyTab Then KeyCode = 0
'KeyCode = 0
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) And TopCtrl1.TopText2 = "Add" Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    'SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case LCode, Lab_Code
            If KeyCode = vbKeyDelete And Shift = 0 Then Exit Sub
        Case ChHrs, ChRate, WrHrs
            If KeyCode = vbKeyDelete And Shift = 0 Then Call Get_Text(Me, FGrid, txt, 2, True, 48)      'FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
SetMaxLength
Select Case FGrid.Col
    Case Lab_Code, LCode
       Call Get_Text(Me, FGrid, txt, 2, False, keyascii)
    Case ChHrs, ChRate, WrHrs
       Call Get_Text(Me, FGrid, txt, 2, True, keyascii)
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
                FGrid.AddItem FGrid.Rows
                FGrid.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid.Rows - 1
            FGrid.TextMatrix(I, 0) = I
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
    txt(2).Visible = False
    Grid_Hide
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
GSQL = "SELECT L.Lab_Code AS LCode " & _
    " FROM Labour L " & _
    " where L.Modelbased=1"
If GCn.Execute(GSQL).RecordCount <= 0 Then
    MsgBox "Please define Model-base Labour in Labour Description Master", vbInformation, "Validation"
    Unload Me
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
WinSetting Me: Ini_Grid
TopCtrl1.Tag = PubUParam
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select max(model) as searchcode,max(model) as model From labour_MODEL Group by MODEL", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select max(Vehicle_Type) as SearchCode,max(Vehicle_Type) as VehicleType From Labour_MODEL Group by Vehicle_Type", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 max(Vehicle_Type) as SearchCode,max(Vehicle_Type) as VehicleType From Labour_MODEL Group by Vehicle_Type", GCn, adOpenDynamic, adLockOptimistic
End If

Set RstHelp = New ADODB.Recordset
'RstHelp.Open "Select MODEL as code,model as name ,model_desc as Listname FROM Model Order by MODEL", GCn, adOpenDynamic, adLockOptimistic
RstHelp.Open "Select Vehicle_Type as Code,Vehicle_Type as Name,Vehicle_Type as ListName FROM Vehicle_Type Order by Vehicle_Type", GCn, adOpenDynamic, adLockOptimistic

Set RstLab = New ADODB.Recordset
RstLab.Open "SELECT LABOUR.LAB_CODE AS CODE,LABOUR.LAB_DESC AS NAME,labour_type.lab_desc as ltype,LABOUR_TYPE.LAB_TYPE AS TCODE,labour_group.labgrp_desc as lgroup,LABOUR_GROUP.LAB_GROUP AS GCODE FROM (labour left join labour_type on labour.lab_type=labour_type.lab_type) left join labour_group on labour.lab_group=labour_group.lab_group ", GCn, adOpenDynamic, adLockOptimistic
Set DGlabM.DataSource = RstHelp

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
txt(VehType).Tag = txt(VehType)
Txt_GotFocus VehType
ADDFLAG = 1
FGrid.Rows = 1
FGrid.AddItem ""
FGrid.FixedRows = 1
txt(VehType).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
Dim Rs As Recordset
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(VehType).Enabled = False
    ADDFLAG = 2
    FGrid.SetFocus
    FGrid.Col = ChHrs
'    FGrid_EnterCell
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
            GCn.Execute ("delete * from Labour_model where Vehicle_Type= '" & Trim(txt(VehType)) & "'")
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
    GSQL = "Select max(Vehicle_Type) as searchcode,max(Vehicle_Type) as VehicleType From labour_MODEL group by Vehicle_Type"
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
        Set RstMain = GCn.Execute("Select max(Vehicle_Type) as SearchCode,max(Vehicle_Type) as VehicleType From Labour_MODEL Where max(Vehicle_Type) = '" & MyValue & "' Group by Vehicle_Type")
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

    mRepName = "ModLabDet"
    mQRY = "SELECT '" & txt(VehType) & "',lg.labgrp_desc as LabGroupDesc,LG.LAB_GROUP AS LabGrCode," & _
        " LT.Lab_Type AS LabType, LT.Lab_Desc as LabTypeDesc," & _
        " LM.Lab_Code,L.Lab_Desc as LabDesc,LM.Lab_Rate,LM.Time_Req,LM.WTime_Req," & _
        " LM.U_Name,LM.U_EntDt,LM.U_AE " & _
        " FROM ((LABOUR_MODEL LM left join labour L on LM.lab_code=L.lab_code) " & _
        " left join Labour_Type LT on L.Lab_Type=LT.Lab_Type) " & _
        " left join Labour_Group LG on L.Lab_Group=LG.Lab_Group " & _
        " WHERE LM.Vehicle_Type='" & txt(VehType) & "'"
    
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
Dim transFlag As Byte
Dim I As Byte
On Error GoTo Errloop
    If txt(2).Visible = True Then
        If TxtGridLeave = False Then
            Txt_LostFocus 2
            Exit Sub
        End If
    End If
'    FGrid.CellBackColor = CellBackColLeave
    Grid_Hide
    transFlag = 0
    If IsValid(txt(VehType), "Vehicle Type") = False Then Exit Sub
    GCn.BeginTrans
    transFlag = 1
    GCn.Execute ("DELETE From Labour_MODEL Where Vehicle_Type='" & txt(VehType) & "'")
    For I = 1 To FGrid.Rows - 1
        If Len(FGrid.TextMatrix(I, Lab_Code)) = 0 Then
            FGrid.Row = I
            FGrid.RemoveItem (FGrid.Row)
        End If
    Next
    For I = 1 To FGrid.Rows - 1
        GCn.Execute ("Insert Into Labour_MODEL(Vehicle_Type,Site_Code,Lab_code,Lab_Rate,Time_Req,WTime_Req,U_Name,U_EntDt,U_AE) Values('" & txt(VehType) & "','" & PubSiteCode & "','" & FGrid.TextMatrix(I, Lab_Code) & "'," & Val(FGrid.TextMatrix(I, ChRate)) & "," & Val(FGrid.TextMatrix(I, ChHrs)) & "," & Val(FGrid.TextMatrix(I, WrHrs)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Next
    GCn.CommitTrans
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select max(Vehicle_Type) as SearchCode,max(Vehicle_Type) as VehicleType From Labour_MODEL Where Vehicle_Type = '" & txt(VehType) & "' Group by Vehicle_Type")
    End If
    RstHelp.Requery
    RstMain.FIND ("SearchCode='" & txt(VehType) & "'")
    Disp_Text SETS("INI", Me, RstMain)
    If ADDFLAG = 1 Then
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        DGlabM.Visible = False
    End If
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
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
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(VehType).BackColor = CtrlBColOrg:      txt(VehType).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
Dim Rs As Recordset
On Error GoTo Errloop
RST_BOF_EOF RstMain

If RstMain.RecordCount <= 0 Then
    BlankText
Else
    txt(VehType) = XNull(RstMain!VehicleType)
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT LM.LAB_CODE,LM.LAB_RATE,LM.TIME_REQ,LM.WTIME_REQ,labour.lab_desc as ldesc,lt.lab_desc as ltype,LT.LAB_TYPE AS TCODE,lg.labgrp_desc as lgroup,LG.LAB_GROUP AS GCODE FROM ((LABOUR_MODEL LM left join labour on LM.lab_code=labour.lab_code) left join labour_type LT on labour.lab_type=lt.lab_type) left join labour_group LG on labour.lab_group=lg.lab_group " & _
    "WHERE LM.Vehicle_Type='" & txt(VehType) & "'")
    If Rs.RecordCount > 0 Then
        FGrid.Rows = 1
        Do Until Rs.EOF
            FGrid.AddItem "" & FGrid.Rows & Chr(9) & Rs!Lab_Code & Chr(9) & Rs!lDesc & Chr(9) & Rs!LType & Chr(9) & Rs!LGroup & Chr(9) & IIf(Rs!TIME_REQ = 0 Or IsNull(Rs!TIME_REQ), "", Format(Rs!TIME_REQ, "0.00")) & Chr(9) & IIf(Rs!Lab_Rate = 0 Or IsNull(Rs!Lab_Rate), "", Format(Rs!Lab_Rate, "0.00")) & Chr(9) & IIf(Rs!WTime_Req = 0 Or IsNull(Rs!WTime_Req), "", Format(Rs!WTime_Req, "0.00")) & Chr(9) & Rs!tCode & Chr(9) & Rs!GCODE
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
If Index <> 2 Then Ctrl_GetFocus txt(Index): txt(2).Visible = False
If Index = 2 Then
    txt(Index).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case Lab_Code, LCode    'Lab_Code , Description
            TStr = ""
            Do Until mROW = FGrid.Rows - 1
                mROW = mROW + 1
                If mROW <> FGrid.Row Then
                    TStr = TStr + "'" + FGrid.TextMatrix(mROW, Lab_Code) + "'" + ","
                End If
            Loop
            
            GSQL = "SELECT L.LAB_CODE AS CODE,L.LAB_DESC AS NAME," & _
                " LT.lab_desc as ltype,LT.LAB_TYPE AS TCODE," & _
                " LG.labgrp_desc as lgroup,LG.LAB_GROUP AS GCODE " & _
                " FROM (labour L left join labour_type LT on L.lab_type=LT.lab_type) " & _
                " left join labour_group LG on L.lab_group=LG.lab_group " & _
                " where L.Modelbased=1"
            If TStr <> "" Then
                GSQL = GSQL & " and L.lab_code NOT in (" & TStr & ")"
            End If
            Set RstLab = GCn.Execute(GSQL)
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
    Case VehType
        If KeyCode <> vbKeyEscape Then
            DGridTxtKeyDown DGlabM, txt, VehType, RstHelp, KeyCode, False, 0
        End If
    Case 2
        If KeyCode = vbKeyEscape Then
            FGrid.SetFocus
            txt(Index).TEXT = txt(Index).Tag
            Txt_KeyUp Index, KeyCode, Shift
            txt(Index).Visible = False
            DGLab.Visible = False
            Exit Sub
        End If
        Select Case FGrid.Col
            Case Lab_Code    '1
                If DGLab.Visible = False Then DGridColSwap DGLab, 0
                DGridTxtKeyDown DGLab, txt, Index, RstLab, KeyCode, True, 0, frmLabDesc, "frmLabDesc"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, txt, Index, KeyCode, TAddMode, WrHrs, , 5
                    Else
                        Txt_LostFocus 2
                        txt(2).SetFocus
                    End If
                End If
                
            Case LCode
                If DGLab.Visible = False Then DGridColSwap DGLab, 1
                DGridTxtKeyDown DGLab, txt, Index, RstLab, KeyCode, True, 1, frmLabDesc, "frmLabDesc"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, txt, Index, KeyCode, TAddMode, WrHrs, , 5
                    Else
                        Txt_LostFocus 2
                        txt(2).SetFocus
                    End If
                End If
            Case ChHrs, ChRate, WrHrs
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                     If TxtGridLeave = True Then
                        GridTxtDown FGrid, txt, Index, KeyCode, TAddMode, WrHrs   ', 3
                     Else
                        Txt_LostFocus 0
                        txt(0).SetFocus
                     End If
                 End If
            End Select
End Select
'If FGrid.Col = 1 Then FGrid.Col = 5: FGrid_EnterCell
If Index <> 2 And DGlabM.Visible = False Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case Index
    Case VehType
        If DGlabM.Visible = True Then DGridTxtKeyPress txt, VehType, RstHelp, keyascii, "Code"
    Case 2
        Select Case FGrid.Col       'Val(Txt(Index).Tag)
            Case Lab_Code
                If keyascii <> 13 And DGLab.Visible = False Then Txt_KeyDown Index, GridKey, 0
                If RstLab.RecordCount > 0 Then DGridTxtKeyPress txt, Index, RstLab, keyascii, "CODE"
            Case LCode
                If keyascii <> 13 And DGLab.Visible = False Then Txt_KeyDown Index, GridKey, 0
                If RstLab.RecordCount > 0 Then DGridTxtKeyPress txt, Index, RstLab, keyascii, "name"
            Case ChHrs, WrHrs
                Call NumPress(txt(2), keyascii, 3, 2)
            Case ChRate
                Call NumPress(txt(2), keyascii, 5, 2)
        End Select
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
'    Case VehType
'           If DGLabM.Visible = True Then DGridTxtKeyUp Txt, VehType, RstHelp, KeyCode, "Code"
    Case 2
        Select Case FGrid.Col
            Case ChHrs, ChRate, WrHrs
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(Val(txt(Index)) = 0, "", Format(Val(txt(Index)), "0.00"))
        End Select
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
If Index = VehType Then
    DGlabM_Click
End If
If Index <> 2 Then Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rs As Recordset, Rst As Recordset
Select Case Index
    Case VehType
        If txt(VehType) = "" Then Exit Sub
        Set Rst = GCn.Execute("Select max(Vehicle_Type) as model From labour_MODEL where Vehicle_Type='" & txt(VehType) & "' group by Vehicle_Type")
        If ADDFLAG = 1 Then
            If Not Rst.EOF Then MsgBox "Vehicle Type Already Exists", vbInformation, "Validation": txt(VehType) = txt(VehType).Tag: Cancel = True: Exit Sub
        End If
        If RstHelp.RecordCount = 0 Then Exit Sub
        If DGlabM.Visible = True Then
            txt(VehType).TEXT = RstHelp!Name
            txt(VehType).Tag = RstHelp!Code
        End If
        GSQL = "SELECT L.Lab_Code AS LCode,L.Lab_Desc AS LDesc, " & _
            " LT.lab_desc as LType,LT.Lab_Type AS TCode," & _
            " LG.LabGrp_Desc as LGroup,LG.Lab_Group AS GCode," & _
            " L.Time_Req as TReq,L.Lab_Rate as LRate,L.WTime_Req as WTReq " & _
            " FROM (Labour L left join Labour_Type LT on L.Lab_Type=LT.Lab_Type) " & _
            " left join Labour_Group LG on L.Lab_Group=LG.Lab_Group " & _
            " where L.Modelbased=1"
        Set Rs = New Recordset
        Set Rs = GCn.Execute(GSQL)
        If Rs.RecordCount > 0 Then
            FGrid.Rows = 1
            Do Until Rs.EOF
                FGrid.AddItem Rs.AbsolutePosition & Chr(9) & Rs!LCode & Chr(9) & Rs!lDesc & Chr(9) & Rs!LType & Chr(9) & Rs!LGroup & Chr(9) & IIf(Rs!treq = 0 Or IsNull(Rs!treq), "", Format(Rs!treq, "0.00")) & Chr(9) & IIf(Rs!Lrate = 0 Or IsNull(Rs!Lrate), "", Format(Rs!Lrate, "0.00")) & Chr(9) & IIf(Rs!wtreq = 0 Or IsNull(Rs!wtreq), "", Format(Rs!wtreq, "0.00")) & Chr(9) & Rs!tCode & Chr(9) & Rs!GCODE
                Rs.MoveNext
            Loop
            FGrid.FixedRows = 1
        End If
        FGrid.Row = 1
        FGrid.Col = 5
'       FGrid_EnterCell
   Case 2
        Cancel = TxtGridLeave(Index)
End Select
End Sub

Private Sub DGLab_GotFocus()
    mFlag = 1
End Sub

Private Sub DGLabM_GotFocus()
    mFlag = 1
End Sub

Private Sub BlankText()
Dim I As Byte
    txt(0).TEXT = ""
    txt(2).TEXT = ""
    FGrid.Rows = 1
    FGrid.AddItem ""
    FGrid.FixedRows = 1
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    txt(0).Enabled = Enb
    txt(2).Enabled = Enb
'    FGrid.Enabled = Enb
    txt(2).BackColor = CtrlBCol
    txt(2).ForeColor = CtrlFCol
End Sub

Private Sub Ini_Grid()
Dim I As Byte
    With FGrid
        .left = Me.left '+ 45
        .width = Me.width - 120
        .RowHeightMin = PubGridRowHeight '220
        .height = .RowHeight(0) * 20
'        .top = 1575
        .Cols = 10
        .ColAlignmentFixed = flexAlignCenterCenter
        .TextMatrix(0, 0) = "S.No."
        .ColAlignment(0) = flexAlignRightCenter
        .ColWidth(0) = 510
        
        .TextMatrix(0, Lab_Code) = "Code"
        .ColAlignment(Lab_Code) = flexAlignLeftCenter
        .ColWidth(Lab_Code) = 915

        .TextMatrix(0, LCode) = "Description"
        .ColAlignment(LCode) = flexAlignLeftCenter
        .ColWidth(LCode) = 4260
        
        .TextMatrix(0, LType) = "Labour Type"
        .ColAlignment(LType) = flexAlignLeftCenter
        .ColWidth(LType) = 1710
        
        .TextMatrix(0, LGroup) = "Labour Group"
        .ColAlignment(LGroup) = flexAlignLeftCenter
        .ColWidth(LGroup) = 1710
        
        .TextMatrix(0, ChHrs) = "ChrgHr"
        .ColAlignment(ChHrs) = flexAlignRightCenter
        .ColWidth(ChHrs) = 720
        
        .TextMatrix(0, ChRate) = "ChrgRate"
        .ColAlignment(ChRate) = flexAlignRightCenter
        .ColWidth(ChRate) = 945
        
        .TextMatrix(0, WrHrs) = "WarrHr"
        .ColAlignment(WrHrs) = flexAlignRightCenter
        .ColWidth(WrHrs) = 720
        
'  |Code |Description    |Type |Grade |Chrg.Hrs. | Chrg. Rate   |Warr. Hrs.||
'        .TextMatrix(1, Model) = "Model Code"
'        .ColAlignment(Model) = flexAlignLeftCenter
'        .ColWidth(Model) = 1635
        .ColWidth(8) = 0
        .ColWidth(9) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
   
    'DGLabM.width = 7005:
    DGlabM.left = Me.width - (DGlabM.width + mRtScale): DGlabM.top = mTopScale: DGlabM.height = Me.height - (mTopScale + mBotScale)
    DGLab.left = Me.width - (DGLab.width + mRtScale): DGLab.top = mTopScale: DGLab.height = Me.height - (mTopScale + mBotScale)

'    DGlabM.left = 4650
'    DGlabM.top = mTopScale '390
'    DGlabM.Height = 4425
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case Lab_Code, LCode
        If RstLab.RecordCount = 0 Or (RstLab.EOF = True Or RstLab.BOF = True) Or txt(2).TEXT = "" Then
        Else
           FGrid.TextMatrix(FGrid.Row, Lab_Code) = RstLab!Code
           FGrid.TextMatrix(FGrid.Row, LCode) = RstLab!Name
           FGrid.TextMatrix(FGrid.Row, LType) = XNull(RstLab!LType)
           FGrid.TextMatrix(FGrid.Row, LGroup) = RstLab!LGroup
        End If
    Case WrHrs
        If FGrid.Row < FGrid.Rows - 1 Then FGrid.Row = FGrid.Row + 1
'        FGrid.Col = 4
    Case Lab_Code
        If FGrid.TextMatrix(FGrid.Row, Lab_Code) <> txt(2) Then Call Fill_Data
    Case LCode
        If FGrid.TextMatrix(FGrid.Row, LCode) <> txt(2) Then Call Fill_Data
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    txt(2).Visible = False
End If
End Function

Private Sub Fill_Data()
If RstLab.RecordCount = 0 Or (RstLab.EOF = True Or RstLab.BOF = True) Or txt(2).TEXT = "" Then Exit Sub
   FGrid.TextMatrix(FGrid.Row, Lab_Code) = RstLab!Code
   FGrid.TextMatrix(FGrid.Row, LCode) = RstLab!Name
   FGrid.TextMatrix(FGrid.Row, LType) = XNull(RstLab!LType)
   FGrid.TextMatrix(FGrid.Row, LGroup) = RstLab!LGroup
End Sub

Private Sub Grid_Hide()
If DGLab.Visible = True Then DGLab.Visible = False
If DGlabM.Visible = True Then DGlabM.Visible = False
End Sub

Private Sub SetMaxLength()
Select Case FGrid.Col   'Index
    Case Lab_Code
        txt(2).MaxLength = 6
    Case LCode
        txt(2).MaxLength = 40
    Case Else
        txt(2).MaxLength = 0
End Select
End Sub

