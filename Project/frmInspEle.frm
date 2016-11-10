VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmInspEle 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Inspection Element Master"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11460
   Begin MSDataGridLib.DataGrid DGLab 
      Height          =   3885
      Left            =   2520
      Negotiate       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   6853
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
         DataField       =   "Name"
         Caption         =   "Category"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4694.74
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3945
      Left            =   150
      TabIndex        =   7
      Top             =   1485
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   6959
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   8
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   8421504
      FocusRect       =   0
      GridLinesUnpopulated=   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"frmInspEle.frx":0000
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   3960
      TabIndex        =   4
      Top             =   5505
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         Sorted          =   -1  'True
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   1635
      MaxLength       =   20
      TabIndex        =   2
      Top             =   713
      Width           =   3000
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   661
      tAdd            =   0   'False
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   2
      Left            =   3735
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   4
      Left            =   420
      TabIndex        =   3
      Top             =   720
      Width           =   690
   End
End
Attribute VB_Name = "frmInspEle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstLab As ADODB.Recordset
Dim mFlag As Byte
Private Const POn = 0
Dim GridKey As Integer
' Col Declaration
Dim ExitCtrl As Boolean
Private Const Code As Byte = 1
Private Const InsEle As Byte = 2
Private Const TypeIns As Byte = 3
Private Const DefVal As Byte = 4
Private Const PIndex As Byte = 5
Private Const Inscat As Byte = 6
Private Const InsCatCode As Byte = 7
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Sub DGLab_Click()
If RstHelp.RecordCount > 0 Then
    Select Case FGrid.Col
        Case Inscat
            Txt(2).TEXT = RstHelp!Name
            FGrid.TextMatrix(FGrid.Row, InsCatCode) = RstHelp!Code
    End Select
End If
DGLab.Visible = False
Txt(2).SetFocus
End Sub

Private Sub FGrid_Click()
Txt(2).Visible = False
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
'Select Case FGrid.Col
'    Case Code, InsEle, TypeIns, DefVal, Inscat, PIndex
'        Call GridDblClick(Me, FGrid, Txt, 2)
'End Select
'TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
FGrid.BackColorSel = BackColorSelEnter
FGrid.ForeColorSel = ForeColorSelEnter
Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
Dim result As Boolean
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case Code, InsEle, TypeIns, DefVal, Inscat
            Call Get_Text(Me, FGrid, Txt, 2, False, 48)
        Case PIndex
            Call Get_Text(Me, FGrid, Txt, 2, True, 48)
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
SetMaxLength
Select Case FGrid.Col
    Case Code
        Call Get_Text(Me, FGrid, Txt, 2, False, KeyAscii)
    Case InsEle, TypeIns, Inscat
        If FGrid.TextMatrix(FGrid.Row, Code) <> "" Then
            Call Get_Text(Me, FGrid, Txt, 2, False, KeyAscii)
        End If
    Case DefVal
        If FGrid.TextMatrix(FGrid.Row, Code) <> "" Then
            If FGrid.TextMatrix(FGrid.Row, TypeIns) = "Numeric" Then
                Call Get_Text(Me, FGrid, Txt, 2, True, KeyAscii)
            Else
                Call Get_Text(Me, FGrid, Txt, 2, False, KeyAscii)
            End If
        End If
    Case PIndex
        If FGrid.TextMatrix(FGrid.Row, Code) <> "" Then
           Call Get_Text(Me, FGrid, Txt, 2, True, KeyAscii)
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = 46 And Shift = 2 Then
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
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
    Grid_Hide
    Txt(2).Visible = False
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
'Me.top = 1000: Me.left = 200: Me.width = 11600: Me.Height = Me.Height + 700
WinSetting Me: Ini_Grid
TopCtrl1.Tag = PubUParam    ': TopCtrl1.TopText1 = "Inspection Elements Master"   ': TopCtrl1.TopText1.Width = 1000
CtrlClckCol
Txt(POn) = "Inspection Sheet"
MoveRec
TopCtrl1.tSave = False
TopCtrl1.tCancel = False
TopCtrl1.tPrev = False: TopCtrl1.tFirst = False
FGrid.ColWidth(7) = 0
ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
 Set RstHelp = Nothing
End Sub

Private Sub ListView_Click()
Txt(2).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
Txt(2).SetFocus
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Txt(POn).Tag = Txt(POn)
Txt_GotFocus POn
ADDFLAG = 1
FGrid.Rows = 1
FGrid.AddItem ""
FGrid.FixedRows = 1
'Txt(POn).SetFocus
'FGrid.Enabled = True
Txt(2).Enabled = True
TopCtrl1.TopText2 = "Add"
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo Errloop
Dim Res As Integer
Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            GCn.Execute ("DELETE  inspection_element.*  from (inspection_element left join inspection_catg on inspection_element.inspection_catg=inspection_catg.insp_code) where inspection_catg.print_on ='" & left(Txt(POn), 1) & "'" & " AND inspection_catg.SITE_CODE='" & PubSiteCode & "'")
            GCn.CommitTrans
            FGrid.Rows = 1
            FGrid.AddItem ""
            FGrid.FixedRows = 1
        End If
Exit Sub
Errloop:    GCn.RollbackTrans
            MsgBox err.Description, vbExclamation, " Deletion Error "
End Sub

Private Sub TopCtrl1_eEdit()
Dim rs As Recordset
On Error GoTo Errloop
        TopCtrl1.TopText2 = "Edit"
        TopCtrl1.TopText2.ForeColor = RGB(255, 0, 0)
        TopCtrl1.tDel = False
        TopCtrl1.tEdit = False
        'Txt(POn).Enabled = False
        TopCtrl1.tSave = True
        TopCtrl1.tCancel = True
        TopCtrl1.tFirst = False
        TopCtrl1.tNext = False
        TopCtrl1.tLast = False
        TopCtrl1.tPrev = False

    ADDFLAG = 2
    Set RstHelp = New ADODB.Recordset
    RstHelp.Open "Select INSP_CODE as code,INSP_DESCRIPTION as name FROM INSPECTION_CATG where PRINT_ON='" & left(Txt(POn), 1) & "'" & "and SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by INSP_DESCRIPTION", GCn, adOpenDynamic, adLockOptimistic
    Set DGLab.DataSource = RstHelp
    Txt(2).Enabled = True
'    FGrid.Enabled = True
    FGrid.AddItem "" & FGrid.Rows
    FGrid.SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_eFind()
    MsgBox "find"
End Sub

Private Sub TopCtrl1_eFirst()
Txt(POn) = "Inspection Sheet"
MoveRec
TopCtrl1.tPrev = False: TopCtrl1.tFirst = False
TopCtrl1.tNext = True: TopCtrl1.tLast = True
End Sub

Private Sub TopCtrl1_eLast()
Txt(POn) = "None"
MoveRec
TopCtrl1.tNext = False: TopCtrl1.tLast = False
TopCtrl1.tPrev = True: TopCtrl1.tFirst = True
End Sub

Private Sub TopCtrl1_eNext()
If Txt(POn) = "Inspection Sheet" Then
    Txt(POn) = "Job Card"
ElseIf Txt(POn) = "Job Card" Then
    Txt(POn) = "None"
End If
MoveRec
If Txt(POn) = "None" Then TopCtrl1.tNext = False: TopCtrl1.tLast = False
TopCtrl1.tPrev = True: TopCtrl1.tFirst = True
End Sub

Private Sub TopCtrl1_ePrev()
If Txt(POn) = "None" Then
Txt(POn) = "Job Card"
ElseIf Txt(POn) = "Job Card" Then
Txt(POn) = "Inspection Sheet"
End If
MoveRec
If Txt(POn) = "Inspection Sheet" Then TopCtrl1.tPrev = False: TopCtrl1.tFirst = False
TopCtrl1.tNext = True: TopCtrl1.tLast = True
End Sub

Private Sub TopCtrl1_ePrn()
Dim i As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "InspSheet"
    mQRY = "select ic.Print_On, IC.Report_Index as ICRIndex, IC.Insp_Code, IC.Insp_Description," & _
        " IE.Report_Index as IERIndex, IE.InspElem_Code, IE.InspElem_Description, IE.Default_Value " & _
        " from (Inspection_Catg IC Left Join Inspection_Element IE on ic.Insp_Code=ie.Inspection_Catg) " & _
        " where ic.Print_On ='" & left(Txt(POn), 1) & _
        "'Order by IC.Report_Index,IC.Insp_Code,IE.Report_Index, IE.InspElem_Code"
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
Dim i As Byte
On Error GoTo Errloop
    
    transFlag = 0
    GCn.BeginTrans
    transFlag = 1
    GCn.Execute ("DELETE  inspection_element.*  from (inspection_element left join inspection_catg on inspection_element.inspection_catg=inspection_catg.insp_code) where inspection_catg.print_on ='" & left(Txt(POn), 1) & "'" & "")
    For i = 1 To FGrid.Rows - 1
        If Len(FGrid.TextMatrix(i, Code)) = 0 Or Len(FGrid.TextMatrix(i, Inscat)) = 0 Then
            FGrid.Row = i
            FGrid.RemoveItem (FGrid.Row)
        End If
    Next
    For i = 1 To FGrid.Rows - 1
         GCn.Execute ("Insert Into INSPECTION_ELEMENT(InspElem_Code,Inspection_Catg,Site_Code,InspElem_Description,Insp_ValueType,Default_Value,Report_Index,U_Name,U_EntDt,U_AE) Values('" & FGrid.TextMatrix(i, Code) & "','" & FGrid.TextMatrix(i, InsCatCode) & "','" & PubSiteCode & "','" & FGrid.TextMatrix(i, InsEle) & "','" & FGrid.TextMatrix(i, TypeIns) & "','" & FGrid.TextMatrix(i, DefVal) & "'," & Val(FGrid.TextMatrix(i, PIndex)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Next
    GCn.CommitTrans
        TopCtrl1.TopText2 = "Browse"
        TopCtrl1.TopText2.ForeColor = RGB(0, 0, 0)
        TopCtrl1.tEdit = True
        TopCtrl1.tDel = True
        TopCtrl1.tSave = False
        TopCtrl1.tCancel = False
        TopCtrl1.tFirst = True: TopCtrl1.tNext = True: TopCtrl1.tPrev = True: TopCtrl1.tLast = True
        If Txt(POn) = "Inspection Sheet" Then TopCtrl1.tPrev = False: TopCtrl1.tFirst = False
        If Txt(POn) = "None" Then TopCtrl1.tNext = False: TopCtrl1.tLast = False
        'Txt(POn).Enabled = True
        'Txt(POn).SetFocus
        transFlag = 0
        Txt(2).Visible = False
        FGrid.CellBackColor = CellBackColLeave
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
            If err.NUMBER = -2147467259 Then MsgBox "Inspection Elements,Type of Inspection,Default Value are required Fields", vbInformation, "Validation Error": Exit Sub
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        MoveRec
        CtrlClckCol
        DGLab.Visible = False
        TopCtrl1.TopText2 = "Browse"
        TopCtrl1.TopText2.ForeColor = RGB(0, 0, 0)
        TopCtrl1.tDel = True
        TopCtrl1.tEdit = True
        TopCtrl1.tSave = False
        TopCtrl1.tCancel = False
        TopCtrl1.tFirst = True: TopCtrl1.tNext = True: TopCtrl1.tPrev = True: TopCtrl1.tLast = True
        If Txt(POn) = "Inspection Sheet" Then TopCtrl1.tPrev = False: TopCtrl1.tFirst = False
        If Txt(POn) = "None" Then TopCtrl1.tNext = False: TopCtrl1.tLast = False
        'Txt(POn).Enabled = True
'        Txt(POn).SetFocus
'        FGrid.Enabled = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    Txt(POn).BackColor = CtrlBColOrg:      Txt(POn).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
Dim rs As Recordset
On Error GoTo Errloop
GSQL = "select ie.Report_Index as RIndex,ie.InspElem_Code,ie.InspElem_Description," & _
    " ie.Insp_ValueType as InsType,ie.Default_Value as DefVal,ic.Insp_Code,ic.Insp_Description " & _
    " from (inspection_element ie left join inspection_catg ic on ie.inspection_catg=ic.insp_code) " & _
    " where ic.print_on ='" & left(Txt(POn), 1) & _
    "' Order by IC.Report_Index,IC.Insp_Code,ie.Report_Index,ie.InspElem_Code"
Set rs = New Recordset
Set rs = GCn.Execute(GSQL)
    If rs.RecordCount > 0 Then
        FGrid.Rows = 1
        Do Until rs.EOF
            FGrid.AddItem "" & FGrid.Rows & Chr(9) & rs!inspelem_code & Chr(9) & rs!inspelem_description & Chr(9) & rs!INSTYPE & Chr(9) & rs!DefVal & Chr(9) & rs!rindex & Chr(9) & rs!Insp_description & Chr(9) & rs!insp_Code
            rs.MoveNext
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.Rows = 1
        FGrid.AddItem ""
        FGrid.FixedRows = 1
    End If
Grid_Hide
TopCtrl1.tFind = False
Exit Sub
Errloop:        MsgBox err.Description
End Sub

Private Sub TopCtrl1_eRef()
MoveRec
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Dim TStr$, II As Byte
Grid_Hide
Ctrl_GetFocus Txt(Index)
If Index = 2 Then
    Txt(Index).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
End If
Select Case Index
    Case POn
        ListArray = Array("Inspection Sheet", "Job Card", "None")
        Set mListItem = ListView_Items(ListView, Txt, POn, ListArray, 3)
    Case 2
        Select Case FGrid.Col
            Case TypeIns
                ListArray = Array("Boolean", "Numeric", "Character")
                Set mListItem = ListView_Items(ListView, Txt, 2, ListArray, 3)
        End Select
End Select

End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Dim i As Byte
Dim Txtdate As Boolean
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case POn
        If KeyCode <> vbKeyEscape Then
            ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 900
        End If
    Case 2
        If KeyCode = vbKeyEscape Then
            Txt(2).TEXT = Txt(2).Tag
            Txt_KeyUp 2, KeyCode, Shift
            FGrid.SetFocus
            Txt(2).Visible = False
            Exit Sub
        End If
        Select Case FGrid.Col
            Case Code
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                     If TxtGridLeave = True Then
                          GridTxtDown FGrid, Txt, Index, KeyCode, TAddMode, Inscat       ', 3
                     End If
                 End If
            Case TypeIns    '1
                ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, Txt(Index).top + FGrid.CellHeight, Txt(Index).width, 900
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, Txt, 2, KeyCode, TAddMode, Inscat
                    End If
                End If
            Case InsEle, DefVal
                If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
                     If TxtGridLeave = True Then
                          GridTxtDown FGrid, Txt, Index, KeyCode, TAddMode, Inscat      ', 3
                     End If
                 End If
            Case PIndex
                If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
                     If TxtGridLeave = True Then
                          GridTxtDown FGrid, Txt, Index, KeyCode, TAddMode, Inscat, , Inscat ', 3
                     End If
                 End If
            Case Inscat
                DGridTxtKeyDown DGLab, Txt, 2, RstHelp, KeyCode, False, 1, frmInspCat, "frmInspCat"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, Txt, 2, KeyCode, TAddMode, Inscat
                    End If
                End If
        End Select
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case 2
        Select Case FGrid.Col
            Case DefVal
                If FGrid.TextMatrix(FGrid.Row, TypeIns) = "Numeric" Then
                    NumPress Txt(2), KeyAscii, 8, 0
                End If
            Case PIndex
                NumPress Txt(2), KeyAscii, 2, 0
            Case Inscat
                If RstHelp.RecordCount > 0 Then DGridTxtKeyPress Txt, Index, RstHelp, KeyAscii, "name"
        End Select
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case POn
          ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
            If FrmList.Visible = False Then
                Call MoveRec
                If TopCtrl1.TopText2 = "Add" Or TopCtrl1.TopText2 = "Edit" Then
                    Set RstHelp = New ADODB.Recordset
                    RstHelp.Open "Select INSP_CODE as code,INSP_DESCRIPTION as name FROM INSPECTION_CATG where PRINT_ON='" & left(Txt(POn), 1) & "'" & "and SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by INSP_DESCRIPTION", GCn, adOpenDynamic, adLockOptimistic
                    Set DGLab.DataSource = RstHelp
                End If
            End If
    Case 2
        Select Case FGrid.Col
           Case TypeIns
                If KeyCode <> 13 And FrmList.Visible = False Then Txt_KeyDown 2, GridKey, 0
                If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
           Case Inscat
                If KeyCode <> 13 And DGLab.Visible = False Then Txt_KeyDown Index, GridKey, 0
            Case Code, InsEle
                 FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(Index).TEXT
            Case PIndex
                 FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Val(Txt(Index).TEXT)
            Case DefVal
                If FGrid.TextMatrix(FGrid.Row, TypeIns) = "Boolean" Then
                    If Len(Txt(Index)) = 0 Or UCase(mID(Txt(Index), 1, 1)) = "N" Then
                        Txt(Index) = "No"
                    ElseIf UCase(mID(Txt(Index), 1, 1)) = "Y" Then
                        Txt(Index) = "Yes"
                    Else
                        Txt(Index) = "No"
                    End If
                End If
                If FGrid.TextMatrix(FGrid.Row, TypeIns) = "Numeric" Then
                     FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Val(Txt(Index).TEXT)
                Else
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(Index).TEXT
                End If
        End Select
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
If Index = 2 And ExitCtrl = False Then Exit Sub
    If Index = POn Then
End If
Dim i As Integer
Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim rs As Recordset, Rst As Recordset
Select Case Index
    Case 2
        Select Case FGrid.Col
            Case Code
                If ChkDuplicate = False Then Cancel = True: Exit Sub
            Case TypeIns
                 If Txt(2).TEXT <> "" Then
                     FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(2).TEXT
                 End If
            Case Inscat
                If ChkDuplicate = False Then Cancel = True: Exit Sub
                If RstHelp.RecordCount = 0 Or (RstHelp.EOF = True Or RstHelp.BOF = True) Or Txt(2).TEXT = "" Then
                     FGrid.TextMatrix(FGrid.Row, InsCatCode) = ""
                     FGrid.TextMatrix(FGrid.Row, Inscat) = ""
                Else
                     FGrid.TextMatrix(FGrid.Row, InsCatCode) = RstHelp!Code
                     FGrid.TextMatrix(FGrid.Row, Inscat) = RstHelp!Name
                End If
        End Select
End Select
End Sub

Private Sub DGLab_GotFocus()
    mFlag = 1
End Sub

Private Sub BlankText()
Dim i As Byte
    Txt(2).TEXT = ""
    FGrid.Rows = 1
    FGrid.AddItem ""
    FGrid.FixedRows = 1
End Sub

Private Sub Ini_Grid()
Dim i As Byte
    With FGrid
        .width = 11820 'Me.width - 120
'        .left = 120 ' (Me.width - FGrid.width) / 2
        .RowHeightMin = PubGridRowHeight '220
        .height = .RowHeight(0) * 20
'        .top = 1575
        .Cols = 8
        
        .ColAlignmentFixed = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignRightCenter
        .ColWidth(0) = 510
        
        .TextMatrix(0, Code) = "Code"
        .ColAlignment(Code) = flexAlignLeftCenter
        .ColWidth(Code) = 630

        .TextMatrix(0, InsEle) = "Inspection Elements"
        .ColAlignment(InsEle) = flexAlignLeftCenter
        .ColWidth(InsEle) = 4050
        
        .TextMatrix(0, TypeIns) = "Type of Insp.Value"
        .ColAlignment(TypeIns) = flexAlignLeftCenter
        .ColWidth(TypeIns) = 1635
        
        .TextMatrix(0, DefVal) = "Default Value"
        .ColAlignment(DefVal) = flexAlignLeftCenter
        .ColWidth(DefVal) = 1575
        
        .TextMatrix(0, PIndex) = "Print Index"
        .ColAlignment(PIndex) = flexAlignRightCenter
        .ColWidth(PIndex) = 1020
        
        .TextMatrix(0, Inscat) = "Inspection Category"
        .ColAlignment(Inscat) = flexAlignLeftCenter
        .ColWidth(Inscat) = 1845
        
        .TextMatrix(0, InsCatCode) = "Inspection Cat Code"
        .ColAlignment(InsCatCode) = flexAlignRightCenter
        .ColWidth(InsCatCode) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    DGLab.left = Me.width - (DGLab.width + mRtScale): DGLab.top = mTopScale: DGLab.height = Me.height - (mTopScale + mBotScale)
End Sub

Private Function TxtGridLeave() As Boolean
Dim i As Integer, tst As String, coun As Integer
Select Case FGrid.Col
    Case Code
        If ChkDuplicate = False Then TxtGridLeave = False: ExitCtrl = False: Exit Function
    Case TypeIns
        If Txt(2).TEXT <> "" Then
            Txt(2).TEXT = ListView.SelectedItem.TEXT
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(2).TEXT
            If FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Boolean" Then
                FGrid.TextMatrix(FGrid.Row, DefVal) = "Yes"
            ElseIf FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Numeric" Then
                FGrid.TextMatrix(FGrid.Row, DefVal) = "0"
            ElseIf FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Character" Then
                FGrid.TextMatrix(FGrid.Row, DefVal) = ""
            End If
        End If
'        Case Inscat
'            If ChkDuplicate = False Then TxtGridLeave = False: ExitCtrl = False: Exit Function
'            If RstHelp.RecordCount = 0 Then
'                TxtGridLeave = False: ExitCtrl = False: DGLab.Visible = False: Exit Function
'            End If
'                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(2).Text
       Case Inscat
            If ChkDuplicate = False Then TxtGridLeave = False: ExitCtrl = False: Exit Function
           If RstHelp.RecordCount = 0 Or (RstHelp.EOF = True Or RstHelp.BOF = True) Or Txt(2).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, InsCatCode) = ""
                FGrid.TextMatrix(FGrid.Row, Inscat) = ""
           Else
                FGrid.TextMatrix(FGrid.Row, InsCatCode) = RstHelp!Code
                FGrid.TextMatrix(FGrid.Row, Inscat) = RstHelp!Name
           End If
End Select
ExitCtrl = True
TxtGridLeave = True
FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(2)
Txt(2).Visible = False
FGrid.SetFocus
End Function

Private Sub Grid_Hide()
    If DGLab.Visible = True Then DGLab.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Function ChkDuplicate() As Boolean
Dim i As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte
    Select Case FGrid.Col
    Case Code
        Col1 = Inscat
        Col2 = Code
    Case Inscat
        Col1 = Code
        Col2 = Inscat
    End Select
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(Txt(2).TEXT)))
    For i = 1 To FGrid.Rows - 1
        If i = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(i, Col1))) + CStr(Trim(FGrid.TextMatrix(i, Col2))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
            Txt(2).SetFocus
            Ctrl_GetFocus Txt(2)
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Sub SetMaxLength()
Select Case FGrid.Col   'Index
    Case Code
        Txt(2).MaxLength = 4
    Case InsEle
        Txt(2).MaxLength = 25
    Case DefVal
        Txt(2).MaxLength = 8
    Case Inscat
        Txt(2).MaxLength = 0
    Case PIndex
        Txt(2).MaxLength = 2
    Case Else
        Txt(2).MaxLength = 0
End Select
End Sub


