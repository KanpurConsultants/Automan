VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPartDiscFact 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Part Discount Factor Master"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   8805
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   661
   End
   Begin VB.Frame FrState 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   5805
      TabIndex        =   5
      Top             =   1290
      Visible         =   0   'False
      Width           =   4125
      Begin MSDataGridLib.DataGrid DGState 
         Height          =   3240
         Left            =   30
         TabIndex        =   7
         Top             =   330
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   5715
         _Version        =   393216
         AllowUpdate     =   0   'False
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
            DataField       =   "DiscFac_Catg"
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
            DataField       =   "PurcDisc_Per"
            Caption         =   "Purch. %"
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
            DataField       =   "SalDisc_Per"
            Caption         =   "Sale %"
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
            BeginProperty Column00 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Disc. Factors"
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
         TabIndex        =   6
         Top             =   30
         Width           =   4065
      End
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Height          =   195
      Index           =   3
      Left            =   2865
      TabIndex        =   3
      Top             =   1215
      Width           =   1350
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Height          =   195
      Index           =   2
      Left            =   2865
      TabIndex        =   2
      Top             =   990
      Width           =   1350
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
      Height          =   195
      Index           =   1
      Left            =   2865
      MaxLength       =   2
      TabIndex        =   1
      Top             =   765
      Width           =   1350
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.% (Sale)..............."
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
      Left            =   1005
      TabIndex        =   9
      Top             =   1215
      Width           =   2085
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.% (Purchase)......"
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
      Left            =   1005
      TabIndex        =   8
      Top             =   990
      Width           =   1950
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc.Factor Code*......."
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
      Left            =   1005
      TabIndex        =   4
      Top             =   765
      Width           =   1980
   End
End
Attribute VB_Name = "frmPartDiscFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public MasterFormExit As Boolean
'Private Const CtrlBColOrg = &HC2D5B9, CtrlFColOrg = &H80000012        'Orginal
'Private Const CtrlBCol = &H80000008, CtrlFCol = &H8000000E         'Changed
Private ADDFLAG As Byte, RstMain As ADODB.Recordset
Private Const DiscFac_Catg As Byte = 1, PurcDisc_Per As Byte = 2, SalDisc_Per As Byte = 3

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Load()
Me.left = 0: Me.top = 0
TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = "Part Discount Factor Master"
Set RstMain = New ADODB.Recordset
If PubMoveRecYn Then
    RstMain.Open "Select DiscFac_Catg AS SEaRCHCODE,Part_DiscFactor.* From Part_DiscFactor Order by DiscFac_Catg", GCn, adOpenDynamic, adLockOptimistic
Else
    Set RstMain = GCn.Execute("Select Top 1 DiscFac_Catg AS SEaRCHCODE,Part_DiscFactor.* From Part_DiscFactor Order by DiscFac_Catg")
End If
CtrlClckCol
Disp_Text SETS("INI", Me, RstMain)
MoveRec
Set DGState.DataSource = RstMain
FrState.Visible = False
ADDFLAG = 0
End Sub
Private Sub Form_Resize()
'    TopCtrl1.Width = Me.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing
End Sub
Private Sub CtrlClckCol()
    txt(DiscFac_Catg).BackColor = CtrlBColOrg:         txt(DiscFac_Catg).ForeColor = CtrlFColOrg
    txt(PurcDisc_Per).BackColor = CtrlBColOrg:         txt(PurcDisc_Per).ForeColor = CtrlFColOrg
    txt(SalDisc_Per).BackColor = CtrlBColOrg:         txt(SalDisc_Per).ForeColor = CtrlFColOrg
End Sub
Private Sub Disp_Text(Enb As Boolean)
    txt(DiscFac_Catg).Enabled = Enb
    txt(PurcDisc_Per).Enabled = Enb
    txt(SalDisc_Per).Enabled = Enb
End Sub
Private Sub MakeBlank()
    txt(DiscFac_Catg) = ""
    txt(PurcDisc_Per) = ""
    txt(SalDisc_Per) = ""
End Sub
Private Sub MoveRec()
On Error GoTo Errloop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    txt(DiscFac_Catg) = XNull(RstMain!DiscFac_Catg)
    txt(PurcDisc_Per) = VNull(RstMain!PurcDisc_Per)
    txt(SalDisc_Per) = VNull(RstMain!SalDisc_Per)
End If
Exit Sub
Errloop:        MsgBox err.Description
End Sub
Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
MakeBlank
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
txt(DiscFac_Catg).Tag = txt(DiscFac_Catg)
Txt_GotFocus DiscFac_Catg
txt(DiscFac_Catg).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(DiscFac_Catg).Enabled = False
    txt(PurcDisc_Per).Tag = txt(PurcDisc_Per)
    txt(PurcDisc_Per).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub

Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM
On Error GoTo eloop1
'Disc_Factor
    If GCn.Execute("Select Disc_Factor From Part Where Disc_Factor=" & Chk_Text(txt(DiscFac_Catg))).RecordCount > 0 Then
        MsgBox "Transaction in Part Master !", vbCritical, "Delete Denied"
        Exit Sub
    End If
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        XBM = RstMain.Bookmark
        GCn.Execute ("Delete From Part_DiscFactor Where DiscFac_Catg=" & Chk_Text(txt(DiscFac_Catg)))
        GCn.CommitTrans
        RstMain.Requery
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
    GSQL = "SELECT DiscFac_Catg as searchcode,DiscFac_Catg as Disc_Factor,PurcDisc_Per AS Disc_Purch,SalDisc_Per as Disc_Sale FROM PART_DISCFACTOR "
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
'prn
    MsgBox "print"
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo Errloop
    transFlag = 0
    If IsValid(txt(DiscFac_Catg), "Disc.Factor Code") = False Then Txt_GotFocus DiscFac_Catg: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
    GCn.Execute ("DELETE From Part_DiscFactor Where DiscFac_Catg=" & Chk_Text(txt(DiscFac_Catg)))
    GCn.Execute ("Insert Into Part_DiscFactor (DiscFac_Catg,Site_Code,PurcDisc_Per,SalDisc_Per,U_Name,U_EntDt,U_AE) Values(" & Chk_Text(txt(DiscFac_Catg)) & ",'" & PubSiteCode & "'," & VNull(txt(PurcDisc_Per)) & "," & VNull(txt(SalDisc_Per)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Else
    GCn.Execute ("update Part_DiscFactor set DiscFac_Catg=" & Chk_Text(txt(DiscFac_Catg)) & ",Site_Code='" & PubSiteCode & "',PurcDisc_Per=" & VNull(txt(PurcDisc_Per)) & ",SalDisc_Per=" & VNull(txt(SalDisc_Per)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "'" & " Where DiscFac_Catg=" & Chk_Text(txt(DiscFac_Catg)))
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select DiscFac_Catg AS SEaRCHCODE,Part_DiscFactor.* From Part_DiscFactor Where DiscFac_Catg = " & Chk_Text(txt(DiscFac_Catg)) & " Order by DiscFac_Catg")
    End If
    RstMain.FIND ("DiscFac_Catg=" & Chk_Text(txt(DiscFac_Catg)))
    If ADDFLAG = 1 Then
        MakeBlank
        Txt_GotFocus DiscFac_Catg
        txt(DiscFac_Catg).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        FrState.Visible = False
        ADDFLAG = 0
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
    CtrlClckCol
    FrState.Visible = False
End If
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub DiscFac_CatgNameSearch()
If RstMain.RecordCount <= 0 Then Exit Sub
RstMain.MoveFirst
RstMain.FIND "DiscFac_Catg >=" & Chk_Text(XNull(txt(DiscFac_Catg)))
End Sub
Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case DiscFac_Catg
            FrState.Visible = True
            FrState.top = txt(Index).top + txt(Index).height + 10
            FrState.left = txt(Index).left
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
    If FrState.Visible = True Then FrState.Visible = False
    RST_BOF_EOF RstMain
    txt(Index).Tag = txt(Index)
    Txt_Click Index
    If RstMain.BOF Or RstMain.EOF Then Exit Sub
    Select Case Index
        Case DiscFac_Catg
            mBookMark = RstMain.Bookmark
            RstMain.Sort = "DiscFac_Catg ASC"
            RstMain.Bookmark = mBookMark
            DiscFac_CatgNameSearch
    End Select
    If txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    txt(Index).ForeColor = CtrlFCol: txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Select Case Index
    Case PurcDisc_Per, SalDisc_Per
        NumDown txt(Index), KeyCode, 2, 2
End Select
Select Case Index
    Case DiscFac_Catg
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case PurcDisc_Per
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case SalDisc_Per
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_Click Index
                Txt_GotFocus Index
                txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
End Select
End Sub
Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Select Case Index
    Case PurcDisc_Per, SalDisc_Per
        NumPress txt(Index), keyascii, 2, 2
End Select
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case DiscFac_Catg
        DiscFac_CatgNameSearch
End Select
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case DiscFac_Catg
            Set Rst = GCn.Execute("SELECT * FROM Part_DiscFactor WHERE DiscFac_Catg=" & Chk_Text(txt(DiscFac_Catg)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Disc.Factor Code Already Exists", vbInformation, "Validation": txt(DiscFac_Catg) = txt(DiscFac_Catg).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!DiscFac_Catg <> txt(DiscFac_Catg).Tag Then MsgBox "Disc.Factor Code Already Exists", vbInformation, "Validation": txt(DiscFac_Catg) = txt(DiscFac_Catg).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select DiscFac_Catg AS SEaRCHCODE,Part_DiscFactor.* From Part_DiscFactor Where DiscFac_Catg = '" & MyValue & "' Order by DiscFac_Catg")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub


